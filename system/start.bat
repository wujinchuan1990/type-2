@echo off
rem 型式评价报告工作台 Windows 启动脚本（对应 macOS 的 start.command）
setlocal
cd /d "%~dp0"

if exist ".venv\Scripts\python.exe" (
    set "PYTHON=.venv\Scripts\python.exe"
) else if exist "..\.venv\Scripts\python.exe" (
    set "PYTHON=..\.venv\Scripts\python.exe"
) else (
    echo 缺少 Python 环境。请按 README.md 安装依赖后重试。
    pause
    exit /b 1
)

"%PYTHON%" -m run
pause
