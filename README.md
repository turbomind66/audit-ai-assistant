# 审计智能助手 · audit-ai-assistant

> 一个通用的审计辅助技能：把高频审计作业任务标准化封装，一句话触发，输出结构化成果。适用于 WorkBuddy / CodeBuddy 等支持自定义技能的 AI 客户端。

## ✨ 功能

- **本地脱敏**：数据不出本机，对单位名、人名、身份证、银行账号、手机号、邮箱、统一社会信用代码进行脱敏，生成脱敏文件 + 本地保管的还原映射表。
- **四大审计任务模式**（一句话触发）：
  1. 审计底稿要点梳理
  2. 审计程序清单生成
  3. 往来款询证函生成
  4. 账龄 / 成本数据核对
- **内置规则**：自动脱敏、风险分级标签（🔴 高风险 / 🟡 底稿完善 / 🔵 建议程序）、人工复核声明。

## 📁 仓库结构

```
audit-ai-assistant/
├── README.md                  # 本说明
├── SKILL.md                   # 技能定义（触发词、任务模式）
├── 安装与使用说明.md           # 详细安装与使用指南
├── audit_desensitize.py       # 本地脱敏主程序
├── 一键脱敏.bat               # Windows 脱敏入口（双击 / 拖图标）
└── examples/                  # 演示与示例数据
    ├── AI赋能审计业务-成果简报.html
    ├── 演示报告_完整链路.html
    └── 示例数据/
        ├── 测试数据/          # 应收账款清理测试底稿等
        └── test/              # 应收账款明细示例
```

> 注：本仓库为**源码 + 文档 + 演示**，不内含离线 Python 运行时与离线依赖包；离线运行时以 [Release](https://github.com/turbomind66/audit-ai-assistant/releases) 形式分发（见下方「离线安装」）。

## 🚀 安装（三步）

1. **安装依赖**（任选其一）：
   - 已装 Python：`pip install pandas openpyxl`
   - 完全离线：自行放置便携 Python 运行时到本目录（参考技能打包版）
2. **拷贝技能**：把本仓库（除 `examples/` 外）拷到本地技能目录，例如：
   ```
   C:\Users\<你的用户>\.workbuddy\skills\audit-assistant\
   ```
   最终路径应为 `...\skills\audit-assistant\SKILL.md`
3. **重启** AI 客户端，让技能加载生效

## 📦 离线安装（免配置，推荐）

不想装 Python、不想联网？用离线运行时包，解压即用：

1. 到 [Releases](https://github.com/turbomind66/audit-ai-assistant/releases) 下载 `audit-ai-assistant-offline-v1.zip`（约 55 MB，已内含便携 Python 3.11 + pandas/openpyxl + 离线依赖）。
2. 解压，把里面的 `audit-ai-assistant-offline/` 整个文件夹拷到技能目录：
   ```
   C:\Users\<你的用户>\.workbuddy\skills\audit-assistant\
   ```
3. 重启 AI 客户端即可。双击 `一键脱敏.bat`（或把文件夹拖到它图标上）即脱敏，**全程离线、数据不出本机**。

> 离线包与「在线安装」二选一即可；离线包已做通用化处理，不含任何单位专属信息。

## 🤖 让 Agent 帮你装（最省事）

不想自己解压、找目录？直接把离线包交给 Agent：

1. 下载 [离线运行时包](https://github.com/turbomind66/audit-ai-assistant/releases/download/offline-v1/audit-ai-assistant-offline-v1.zip)，或在对话里直接把文件发给 Agent。
2. 对 Agent 说一句：
   > 安装这个技能
3. Agent 会自动解压并放入技能目录，装好后会告知你；**重启 AI 客户端**后生效。

> 适用于 WorkBuddy / CodeBuddy 等支持「让 Agent 操作本地文件」的客户端。交给 Agent 的是离线包，全程无需联网。

## 🧰 使用

**第一步 · 本地脱敏**（涉及真实数据必做，数据不出本机）

把文件夹拖到 `一键脱敏.bat` 图标上松手，或：

```
python audit_desensitize.py "D:\审计资料"
python audit_desensitize.py "D:\审计资料" --amount-factor 0.87   # 金额按比例缩放
```

产出 `_脱敏` 目录（可交给 AI 分析）、映射表（**本地保管，禁止外发**）、脱敏报告。

**第二步 · 调用技能**

```
审计助手：生成应收账款清理的程序清单
审计助手：梳理底稿——某单位应收账款清理，账龄 1 年以上 8 户……
审计助手：生成询证函，致 XX 公司，截止 2025-12-31 应收余额 XX 万元
审计助手：核对这份账龄分析表
```

## ⚠️ 注意事项

- 本技能定位为**辅助工具**，AI 输出均需审计人员复核确认后使用，不替代职业判断。
- 脱敏映射表是"还原钥匙"，**本地保管、禁止外发**。
- 对外演示 / 共享材料中的业务数据，请使用脱敏样例。

## 📄 许可

本仓库以 MIT 许可证开源，可自由用于学习与工作场景。
