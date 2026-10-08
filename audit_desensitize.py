#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=========================================================
 审计数据本地脱敏工具  audit_desensitize.py
=========================================================
【安全声明】本脚本 100% 本地运行，不联网、不上传任何数据。

【用途】
   第一步：对指定文件夹内的审计资料进行本地脱敏（Excel / CSV / Word / 文本）
   第二步：将脱敏后的文件交给 AI 技能做底稿分析、程序清单、询证函等工作

【基本用法】
   python audit_desensitize.py "D:\\审计资料"
   python audit_desensitize.py "D:\\审计资料" -o "D:\\审计资料_脱敏"
   python audit_desensitize.py "D:\\审计资料" --names 敏感词表.txt
   python audit_desensitize.py "D:\\审计资料" --amount-factor 0.87

【常用参数】
   --names FILE        自定义敏感词表（一行一个真实名称，优先精确替换）
   --amount-factor X   金额按比例缩放（如 0.87），保持数据间比例关系，不破坏分析结论
   --no-person         关闭人名识别（担心误伤时使用）
   --no-addr           关闭地址脱敏
   --inplace           原地覆盖（默认输出到新目录，更安全）

【依赖】
   必需：pandas、openpyxl
   可选：python-docx（处理 Word）、pdfplumber（处理文本版 PDF）
   安装：pip install pandas openpyxl python-docx pdfplumber
=========================================================
"""

import argparse
import json
import re
import sys
from datetime import datetime
from pathlib import Path

if sys.platform.startswith("win"):
    # Windows 下跟随控制台编码（GBK / UTF-8 均可正常显示中文），仅容错避免编码异常中断
    try:
        sys.stdout.reconfigure(errors="replace")
    except Exception:
        pass

# ===================== 依赖检查 =====================

def _try_import(name):
    try:
        return __import__(name)
    except ImportError:
        return None

pd = _try_import("pandas")
openpyxl = _try_import("openpyxl")
docx = _try_import("docx")
pdfplumber = _try_import("pdfplumber")

if pd is None or openpyxl is None:
    miss = [n for n, m in (("pandas", pd), ("openpyxl", openpyxl)) if m is None]
    print("[错误] 缺少必需依赖：%s" % "、".join(miss))
    print("请先执行：pip install pandas openpyxl")
    sys.exit(1)

# ===================== 识别规则 =====================

RE_IDCARD = re.compile(r"(?<!\d)(\d{17}[\dXx]|\d{15})(?!\d)")
RE_BANK = re.compile(r"(?<!\d)\d{16,19}(?!\d)")
RE_PHONE = re.compile(r"(?<!\d)1[3-9]\d{9}(?!\d)")
RE_EMAIL = re.compile(r"[\w.\-+]+@[\w\-]+(\.[\w\-]+)+")
RE_USCC = re.compile(r"\b[0-9A-HJ-NPQRTUWXY]{2}\d{6}[0-9A-HJ-NPQRTUWXY]{10}\b")
RE_ORG = re.compile(
    r"(?!完成|协助|涉及|包括|包含|关于|根据|按照|通过|结合|由于|并且|与|和|或|及|同|向|对|在|由|至|从)"
    r"(?:(?!年|月|日)[\u4e00-\u9fa5()（）·]){2,10}?"
    r"(?:股份有限公司|有限责任公司|有限公司|分公司|集团公司|总公司|集团|总厂|分厂|工厂|"
    r"发电厂|实业公司|工程公司|公司|厂|矿)"
)
# 地址：要求 省? + 市/盟/州 + 区/县/旗 + 路/街/道 + 门牌号 的完整结构，避免碎片化残留
RE_ADDR = re.compile(
    r"(?:[\u4e00-\u9fa5]{2,8}(?:省|自治区))?"
    r"[\u4e00-\u9fa5]{2,10}(?:市|盟|州)"
    r"[\u4e00-\u9fa5]{2,10}(?:区|县|旗)"
    r"[\u4e00-\u9fa5\w]{2,20}?(?:路|街|道|大道|巷|园区|工业园)"
    r"[\w\u4e00-\u9fa5\-]{0,15}"
)

# 说明：已剔除审计/财务文本中的高频干扰姓（计、元、明、成、余、平、费、万、支、项、
#       和、常、时、安、毛、乐、于 等），这些字在"审计""余额""成本""万元"中极易误判。
#       如确需脱敏这些姓氏的人名，请用 --names 敏感词表精确指定。
SURNAMES = (
    "赵钱孙李周吴郑王冯陈褚卫蒋沈韩杨朱秦尤许何吕施张孔曹严华金魏陶姜戚谢邹喻柏水窦章"
    "云苏潘葛奚范彭郎鲁韦昌马苗凤花方俞任袁柳鲍史唐廉岑薛雷贺倪汤滕殷罗毕郝"
    "傅皮卞齐康伍卜顾孟黄穆萧尹姚邵湛汪祁禹狄米贝臧伏戴宋茅庞熊纪舒"
    "屈祝董梁杜阮蓝闵席季麻强贾路娄危江童颜郭梅盛林刁钟徐邱骆高夏蔡田樊胡凌霍虞"
)

def _all_chinese(s):
    """是否全为汉字（人名候选必须全汉字，避免"于20""计。"这类误判）。"""
    return all("\u4e00" <= ch <= "\u9fa5" for ch in s)


# 人名误伤防护：候选词若含以下字符，判定为地名/机构/部门/职务，不按人名处理
PERSON_BLOCK_CHARS = set("省市区县旗州盟路街道巷号楼层室组科处部厂店司园队班乡镇村人责任单位")

# 人名误伤防护：候选词尾字若为高频虚词/量词/计量单位，不按人名处理
PERSON_TAIL_STOP = set(
    "由的了是在和与为对从到被把将于及或等中上下内外后前时年月日号元万亿额"
    "部区市省县路街道层室组科处队班厂店司园项次件种类数量价比率新乡镇村计审人任责"
    "合账款存付收支借贷结核盘点本利银税费挂转提缴欠"
)

# 人名第二字若为财务/审计高频字，判定为财务词而非人名：
# 例："金额""余额""应付""纪要""对方""账户"等（首字"金/余/应/纪/方"为百家姓，易误判）
FINANCE_SECOND_CHAR = set(
    "额账款存付收支借贷结核盘点本利价税费率比数量审批报销挂转提缴欠余项"
    "有已要需应未拟将可无"
)

# 列名含以下关键词 → 整列按人名/单位脱敏（准确率高）
PERSON_COL_HINT = ("经办", "负责", "签字", "姓名", "人员", "主管", "领导", "审批",
                   "复核", "制表", "核对", "联系人", "接收人", "移交人", "接任")
ORG_COL_HINT = ("单位", "客户", "供应商", "往来", "对方", "公司", "部门", "购货", "销货")
ADDR_COL_HINT = ("地址", "住所", "注册地", "经营地", "所在地")
AMOUNT_COL_HINT = ("金额", "余额", "收入", "成本", "费用", "单价", "合计", "应收", "应付",
                   "账面", "发生额", "本金", "利息", "税额", "价税", "结算")


# ===================== 脱敏引擎 =====================

class Desensitizer:
    """一致性脱敏：同一个真实值始终映射为同一个代号，保证脱敏后数据仍可分析。"""

    def __init__(self, custom_map=None, amount_factor=None,
                 do_person=True, do_addr=True):
        self.maps = {}        # {类别: {真实值: 代号}}
        self.counters = {}
        self.stats = {}
        self.amount_factor = amount_factor
        self.do_person = do_person
        self.do_addr = do_addr
        self.custom = custom_map or {}

    # ---- 内部工具 ----
    def _code(self, real, category, prefix):
        d = self.maps.setdefault(category, {})
        if real in d:
            return d[real]
        n = self.counters.get(category, 0) + 1
        self.counters[category] = n
        code = "%s%03d" % (prefix, n)
        d[real] = code
        self._bump(category)
        return code

    def _bump(self, key):
        self.stats[key] = self.stats.get(key, 0) + 1

    # ---- 结构化字段（直接遮盖，保留部分可读）----
    def _mask_idcard(self, v):
        self._bump("身份证号")
        return v[:6] + "*" * (len(v) - 10) + v[-4:] if len(v) >= 15 else "*" * len(v)

    def _mask_bank(self, v):
        self._bump("银行账号")
        return v[:4] + "*" * (len(v) - 8) + v[-4:]

    @staticmethod
    def _mask_phone(v):
        return v[:3] + "****" + v[-4:]

    @staticmethod
    def _mask_email(v):
        if "@" not in v:
            return v
        local, domain = v.rsplit("@", 1)
        return (local[0] if local else "*") + "***@" + domain

    @staticmethod
    def _mask_uscc(v):
        return v[:2] + "*" * (len(v) - 6) + v[-4:]

    @staticmethod
    def _mask_addr(m):
        """地址：保留省、市（便于地区维度分析），其后的区县路号全部遮盖。"""
        s = m.group()
        head = re.match(r"^((?:[\u4e00-\u9fa5]{2,8}(?:省|自治区))?[\u4e00-\u9fa5]{2,10}(?:市|盟))", s)
        return (head.group(1) if head else "") + "****"

    def _is_person(self, cand):
        """人名候选词双重过滤，宁可漏判也不误伤业务数据。"""
        if any(ch in PERSON_BLOCK_CHARS for ch in cand):
            return False
        if cand[-1] in PERSON_TAIL_STOP:
            return False
        # 第二字为财务/审计高频字 → 判定为财务词（"金额""余额""纪要""对方"等），非人名
        if len(cand) >= 2 and cand[1] in FINANCE_SECOND_CHAR:
            return False
        return True

    def _mask_person_text(self, s):
        """自由文本人名替换：逐字符扫描，无效候选只前进 1 位，
        避免"审计办公室张建国"这类与前一候选重叠的人名被漏判。
        """
        parts, i, n = [], 0, len(s)
        while i < n:
            ch = s[i]
            if ch in SURNAMES:
                hit = False
                for length in (3, 2):        # 优先匹配 3 字姓名，再试 2 字
                    cand = s[i:i + length]
                    if len(cand) == length and _all_chinese(cand) and self._is_person(cand):
                        parts.append(self._code(cand, "人员", "人员"))
                        i += length
                        hit = True
                        break
                if not hit:
                    parts.append(ch)
                    i += 1
            else:
                parts.append(ch)
                i += 1
        return "".join(parts)

    # ---- 主替换 ----
    def mask_text(self, text, col_hint="", is_header=False):
        """对单个字符串执行全部脱敏规则。
        col_hint：所在列的表头（用于列名感知）；is_header：是否为表头行（表头不做人名映射）。
        """
        if text is None:
            return text
        s = str(text)
        if not s.strip():
            return s

        # 1) 自定义敏感词表优先（精确、最可靠）
        for real, code in self.custom.items():
            if real and real in s:
                s = s.replace(real, code)
                self._bump("自定义词表")

        # 2) 列名感知：整列按人名/单位处理（高精度规则；表头自身不作为值处理）
        if col_hint and not is_header:
            # 地址列优先：走地址脱敏（保留市区），不按单位/人名整列替换
            if not any(k in col_hint for k in ADDR_COL_HINT):
                if any(k in col_hint for k in PERSON_COL_HINT) and 1 < len(s) <= 12:
                    return self._code(s.strip(), "人员", "人员")
                if any(k in col_hint for k in ORG_COL_HINT) and 1 < len(s) <= 40:
                    # 纯数字/账号类（银行账号、账号、金额）不按单位处理，交给结构化规则遮盖
                    if not re.match(r"^[\d\s\-\*\.]+$", s):
                        return self._code(s.strip(), "单位", "单位")

        # 3) 结构化敏感字段（身份证 / 银行账号 / 手机 / 邮箱 / 统一社会信用代码）
        s = RE_IDCARD.sub(lambda m: self._mask_idcard(m.group()), s)
        s = RE_BANK.sub(lambda m: self._mask_bank(m.group()), s)
        s = RE_PHONE.sub(lambda m: (self._bump("手机号"), self._mask_phone(m.group()))[1], s)
        s = RE_EMAIL.sub(lambda m: (self._bump("邮箱"), self._mask_email(m.group()))[1], s)
        s = RE_USCC.sub(lambda m: (self._bump("信用代码"), self._mask_uscc(m.group()))[1], s)

        # 4) 地址（先于单位/人名处理，保留到路名，门牌号遮盖）
        if self.do_addr:
            s = RE_ADDR.sub(lambda m: (self._bump("地址"), self._mask_addr(m))[1], s)

        # 5) 单位名称（一致性代号）
        s = RE_ORG.sub(lambda m: self._code(m.group(), "单位", "单位"), s)

        # 6) 人名（保守匹配 + 双重过滤；表头行跳过，避免"应收金额"之类的误伤）
        if self.do_person and not is_header:
            s = self._mask_person_text(s)
        return s

    def mask_number(self, value, col_hint=""):
        """金额按比例缩放（保持数据间比例关系，不破坏分析结论）。"""
        if self.amount_factor is None:
            return value
        if not isinstance(value, (int, float)) or isinstance(value, bool):
            return value
        if col_hint and not any(k in col_hint for k in AMOUNT_COL_HINT):
            return value
        self._bump("金额缩放")
        return round(value * self.amount_factor, 2)


# ===================== 文件处理 =====================

SUPPORTED = {".xlsx", ".xlsm", ".csv", ".txt", ".md", ".docx", ".pdf"}


def _col_hint(ws, col_idx):
    """取该列第 1 行作为表头提示（表头通常在首行，避免把数据行误当表头）。"""
    v = ws.cell(row=1, column=col_idx).value
    return v.strip() if isinstance(v, str) else ""


def process_xlsx(src, dst, d):
    wb = openpyxl.load_workbook(src)
    for ws in wb.worksheets:
        # 先收集表头，避免处理过程中被改动
        headers = {c: _col_hint(ws, c) for c in range(1, ws.max_column + 1)}
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None:
                    continue
                # 第 1 行视为表头：不做人名映射与列名映射，避免"应收金额"之类的列名被误改
                is_header = cell.row == 1
                if isinstance(cell.value, str):
                    if cell.value.strip():
                        cell.value = d.mask_text(cell.value, headers.get(cell.column, ""),
                                                 is_header=is_header)
                elif isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool):
                    cell.value = d.mask_number(cell.value, headers.get(cell.column, ""))
    wb.save(dst)


def process_csv(src, dst, d):
    df = pd.read_csv(src, dtype=str, encoding="utf-8-sig")
    if df.empty:
        df.to_csv(dst, index=False, encoding="utf-8-sig")
        return
    # 前两行作为表头提示
    heads = [str(df.columns[i]) + str(df.iloc[0, i] if len(df) > 0 else "")
             for i in range(len(df.columns))]
    for i, col in enumerate(df.columns):
        df[col] = df[col].apply(lambda v: d.mask_text(v, heads[i]))
    df.to_csv(dst, index=False, encoding="utf-8-sig")


def process_txt(src, dst, d):
    text = Path(src).read_text(encoding="utf-8", errors="ignore")
    Path(dst).write_text(d.mask_text(text), encoding="utf-8")


def process_docx(src, dst, d):
    if docx is None:
        raise RuntimeError("处理 .docx 需要 python-docx：pip install python-docx")
    document = docx.Document(src)
    for p in document.paragraphs:
        if p.text.strip():
            for run in p.runs:
                if run.text.strip():
                    run.text = d.mask_text(run.text)
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        if run.text.strip():
                            run.text = d.mask_text(run.text)
    document.save(dst)


def process_pdf(src, dst, d):
    if pdfplumber is None:
        raise RuntimeError("处理 .pdf 需要 pdfplumber：pip install pdfplumber")
    out = []
    with pdfplumber.open(src) as pdf:
        for page in pdf.pages:
            out.append(d.mask_text(page.extract_text() or ""))
    Path(dst).write_text("\n\n".join(out), encoding="utf-8")


HANDLERS = {
    ".xlsx": process_xlsx, ".xlsm": process_xlsx,
    ".csv": process_csv,
    ".txt": process_txt, ".md": process_txt,
    ".docx": process_docx,
    ".pdf": process_pdf,
}


# ===================== 映射表与报告 =====================

def save_mapping(d, out_dir, stamp):
    rows = []
    for cat, mapping in d.maps.items():
        for real, code in mapping.items():
            rows.append([cat, real, code])
    if not rows:
        return None
    # xlsx（本地留存，禁止外传）
    xlsx_path = out_dir / ("_脱敏映射表_%s.xlsx" % stamp)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "映射表"
    ws.append(["类别", "真实值（敏感）", "脱敏代号"])
    for r in rows:
        ws.append(r)
    wb.save(xlsx_path)
    # json 备份
    json_path = out_dir / ("_脱敏映射表_%s.json" % stamp)
    json_path.write_text(json.dumps(d.maps, ensure_ascii=False, indent=2), encoding="utf-8")
    return xlsx_path


def save_report(d, out_dir, stamp, files, skipped, elapsed):
    lines = [
        "审计数据脱敏报告",
        "=" * 50,
        "生成时间：%s" % datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "耗时：%.1f 秒" % elapsed,
        "",
        "一、处理情况",
        "  成功处理文件：%d 个" % len(files),
    ]
    for f in files:
        lines.append("    - %s" % f)
    if skipped:
        lines.append("")
        lines.append("  跳过/失败文件：%d 个" % len(skipped))
        for f, reason in skipped:
            lines.append("    - %s（%s）" % (f, reason))
    lines += ["", "二、脱敏统计（替换次数）"]
    if d.stats:
        for k, v in sorted(d.stats.items(), key=lambda x: -x[1]):
            lines.append("  %-14s %d 处" % (k, v))
    else:
        lines.append("  未发现需脱敏内容（请人工复核）")
    lines += [
        "",
        "三、重要提醒（务必阅读）",
        "  1. 映射表是还原数据的'钥匙'，【禁止】与脱敏文件一起外发或上传，请本地妥善保管。",
        "  2. 人名识别采用保守规则，可能存在误判或漏判，请人工抽查复核。",
        "  3. 图片、扫描件 PDF、图表中的文字无法自动脱敏，需人工处理。",
        "  4. 本脚本未联网，但请确认脱敏后的文件在交给 AI 前已人工确认无误。",
        "",
        "四、下一步",
        "  将脱敏后的文件交给 AI 审计助手，例如：",
        "    '审计助手：梳理底稿——（粘贴脱敏后的资料摘要）'",
        "    '审计助手：核对这份账龄分析表（附脱敏文件）'",
    ]
    path = out_dir / ("_脱敏报告_%s.txt" % stamp)
    path.write_text("\n".join(lines), encoding="utf-8")
    return path


# ===================== 主流程 =====================

def load_custom_names(path):
    """敏感词表：每行一个真实名称，按顺序替换为 单位9xx / 人员9xx。"""
    if not path:
        return {}
    p = Path(path)
    if not p.exists():
        print("[警告] 未找到敏感词表：%s" % path)
        return {}
    result = {}
    lines = [l.strip() for l in p.read_text(encoding="utf-8").splitlines() if l.strip()]
    for i, name in enumerate(lines, 1):
        code = ("单位9%02d" if any(k in name for k in ("公司", "厂", "单位", "集团", "矿", "店")) else "人员9%02d") % i
        result[name] = code
    return result


def main():
    ap = argparse.ArgumentParser(
        description="审计数据本地脱敏工具（本地运行、不联网）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="示例：python audit_desensitize.py \"D:\\审计资料\" -o \"D:\\审计资料_脱敏\"",
    )
    ap.add_argument("input", help="待脱敏的文件夹路径")
    ap.add_argument("-o", "--output", help="输出文件夹（默认：输入文件夹_脱敏）")
    ap.add_argument("--names", help="自定义敏感词表（txt，一行一个真实名称）")
    ap.add_argument("--amount-factor", type=float,
                    help="金额缩放因子，如 0.87（保持比例关系，隐藏真实金额）")
    ap.add_argument("--no-person", action="store_true", help="关闭人名识别")
    ap.add_argument("--no-addr", action="store_true", help="关闭地址脱敏")
    ap.add_argument("--inplace", action="store_true", help="原地脱敏覆盖原文件（谨慎）")
    args = ap.parse_args()

    src_dir = Path(args.input)
    if not src_dir.is_dir():
        print("[错误] 输入路径不是文件夹：%s" % src_dir)
        sys.exit(1)
    if args.inplace:
        out_dir = src_dir
    else:
        out_dir = Path(args.output) if args.output else Path(str(src_dir) + "_脱敏")
    out_dir.mkdir(parents=True, exist_ok=True)

    print("=" * 56)
    print(" 审计数据本地脱敏工具")
    print(" 输入：%s" % src_dir)
    print(" 输出：%s" % out_dir)
    print(" 模式：本地运行，全程不联网、不上传")
    print("=" * 56)

    custom = load_custom_names(args.names)
    d = Desensitizer(custom_map=custom,
                     amount_factor=args.amount_factor,
                     do_person=not args.no_person,
                     do_addr=not args.no_addr)

    started = datetime.now()
    done, skipped = [], []

    for src in sorted(src_dir.rglob("*")):
        if not src.is_file():
            continue
        if src.name.startswith("~$") or src.name.startswith("."):
            continue
        ext = src.suffix.lower()
        if ext not in SUPPORTED:
            skipped.append((src.name, "暂不支持的格式"))
            continue
        rel = src.relative_to(src_dir)
        dst = out_dir / rel
        dst.parent.mkdir(parents=True, exist_ok=True)
        if ext == ".pdf":
            dst = dst.with_suffix(".txt")
        try:
            HANDLERS[ext](src, dst, d)
            done.append(str(rel))
            print("  [OK] %s" % rel)
        except Exception as e:
            skipped.append((str(rel), str(e)[:60]))
            print("  [跳过] %s —— %s" % (rel, str(e)[:60]))

    elapsed = (datetime.now() - started).total_seconds()
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    print("")
    print("-" * 56)
    print("处理完成：成功 %d 个，跳过 %d 个，用时 %.1f 秒" % (len(done), len(skipped), elapsed))
    if d.stats:
        print("脱敏统计：" + "、".join("%s %d 处" % (k, v)
                                  for k, v in sorted(d.stats.items(), key=lambda x: -x[1])))
    mp = save_mapping(d, out_dir, stamp)
    rp = save_report(d, out_dir, stamp, done, skipped, elapsed)
    if mp:
        print("映射表（本地留存·禁止外传）：%s" % mp.name)
    print("脱敏报告：%s" % rp.name)
    print("-" * 56)
    print("提示：映射表是还原钥匙，切勿与脱敏文件一同发送。")
    print("下一步：把脱敏后的文件交给 AI 审计助手做底稿分析。")


if __name__ == "__main__":
    main()
