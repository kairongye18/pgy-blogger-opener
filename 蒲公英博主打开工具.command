#!/bin/zsh
set -e
cd "$(dirname "$0")"

if ! command -v python3 >/dev/null 2>&1; then
  osascript -e 'display dialog "未找到 Python 3，请先安装 Python 3。" buttons {"好"} default button "好" with icon stop'
  exit 1
fi

python3 -m pip install -r requirements.txt
python3 pgy_gui.py
