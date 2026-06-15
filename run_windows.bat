@echo off
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
  echo 未找到 Python。请先安装 Python 3，并勾选 Add python.exe to PATH。
  pause
  exit /b 1
)

python -m pip install -r requirements.txt

python pgy_gui.py
