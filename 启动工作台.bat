@echo off
rem 型式评价报告工作台 启动脚本（免安装版）
rem 优先使用项目内置的便携 Python 运行时（runtime\python），无需在电脑上安装 Python。
setlocal
cd /d "%~dp0"

if exist "runtime\python\python.exe" (
    set "PYTHON=%~dp0runtime\python\python.exe"
) else if exist ".venv\Scripts\python.exe" (
    set "PYTHON=%~dp0.venv\Scripts\python.exe"
) else (
    echo 未找到 Python 运行时。请保留项目内的 runtime 文件夹，或按 README.md 安装依赖。
    pause
    exit /b 1
)

cd system
"%PYTHON%" -m run
pause
