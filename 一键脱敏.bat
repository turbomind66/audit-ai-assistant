@echo off
title 审计数据本地脱敏工具
cd /d "%~dp0"

echo ============================================================
echo          审 计 数 据 本 地 脱 敏 工 具
echo   全程本地运行，不联网、不上传，数据不出本机
echo ============================================================

rem ---- 定位 Python：优先使用本包自带的便携运行时（无需安装、无需联网）----
set "PY="
if exist "%~dp0python_runtime\python.exe" set "PY=%~dp0python_runtime\python.exe"
if not defined PY (
    if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" set "PY=%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
)
if not defined PY (
    where python >nul 2>nul
    if not errorlevel 1 set "PY=python"
)
if not defined PY (
    echo.
    echo [错误] 未找到 Python。请先安装 Python 3 并加入 PATH，
    echo         或确认本目录下的 python_runtime 文件夹完整未被删除。
    pause
    exit /b
)

rem 让 Python 控制台输出使用 GBK，避免中文乱码
set "PYTHONIOENCODING=gbk"

rem ---- 若用的是系统 Python（非内置运行时），检查并尽力离线补齐依赖 ----
if not "%PY%"=="%~dp0python_runtime\python.exe" (
    "%PY%" -c "import pandas, openpyxl" >nul 2>nul
    if errorlevel 1 (
        echo.
        echo [提示] 当前 Python 缺少 pandas / openpyxl 依赖。
        if exist "%~dp0wheels" (
            echo   正在用本包内置离线轮子安装（无需联网）...
            "%PY%" -m pip install --no-index --find-links "%~dp0wheels" pandas openpyxl >nul 2>nul
            if errorlevel 1 (
                echo   ? 离线安装失败，请联网执行： "%PY%" -m pip install pandas openpyxl
                pause
                exit /b
            )
            echo   ? 依赖已从离线轮子安装完成。
        ) else (
            echo   请执行一次（联网安装，约 1 分钟）：
            echo       "%PY%" -m pip install pandas openpyxl
            echo   安装完成后重新运行本脚本即可。
            pause
            exit /b
        )
    )
)

rem 方式一：把文件夹拖到本 .bat 文件图标上，路径通过第 1 个参数传入
set "SRC=%~1"
if defined SRC goto gotpath

echo.
echo ------------------------------------------------------------
echo   请粘贴或输入要脱敏的文件夹路径，然后按回车。
echo.
echo   小技巧：在资源管理器中选中文件夹，按住 Shift 键点右键，
echo           选择"复制为路径"，然后在本窗口右键粘贴即可。
echo.
echo   （也可以关掉本窗口，把文件夹直接拖到本 .bat 文件图标上）
echo ------------------------------------------------------------
echo.
set /p SRC=文件夹路径：

:gotpath
if not defined SRC (
    echo.
    echo 未输入路径，已取消。
    pause
    exit /b
)

rem 去掉路径两侧的引号（复制为路径时会自带引号）
set SRC=%SRC:"=%

rem 去掉路径末尾多余的反斜杠（拖入时可能携带）
if "%SRC:~-1%"=="\" set "SRC=%SRC:~0,-1%"

if not exist "%SRC%\" (
    echo.
    echo [错误] 路径不存在或不是文件夹：%SRC%
    echo.
    echo   请确认输入的是「文件夹」路径，且路径中不含中文引号。
    pause
    exit /b
)

echo.
echo   是否对金额按比例缩放以隐藏真实金额？
echo   （选 y 会乘一个固定系数，保持数据比例，分析结论不受影响）
set /p AMT=请输入 y 或 n（默认 n）：

set EXTRA=
if /i "%AMT%"=="y" set EXTRA=--amount-factor 0.87

echo.
echo 开始脱敏，请稍候...
echo ------------------------------------------------------------
"%PY%" "%~dp0audit_desensitize.py" "%SRC%" %EXTRA%
echo ------------------------------------------------------------
echo.
echo   处理完成！
echo   脱敏结果输出在：%SRC%_脱敏
echo.
echo   [重要] 输出目录里的"脱敏映射表"是还原数据的钥匙，
echo          请勿与脱敏文件一起外发或上传，本地妥善保管。
echo.
echo   下一步：把脱敏后的文件交给 AI 审计助手做底稿分析。
echo.
pause
