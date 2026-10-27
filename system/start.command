#!/bin/bash
set -euo pipefail
cd "$(dirname "$0")"
if [ -x ".venv/bin/python" ]; then
    PYTHON=".venv/bin/python"
elif [ -x "../.venv/bin/python" ]; then
    PYTHON="../.venv/bin/python"
else
    echo "缺少 Python 环境。请按 README.md 安装依赖后重试。"
    read -r -p "按回车关闭..."
    exit 1
fi
"$PYTHON" -m run
