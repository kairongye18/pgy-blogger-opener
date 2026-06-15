#!/usr/bin/env python3
from __future__ import annotations

import importlib.util
import subprocess
import sys
from pathlib import Path


def main() -> None:
    project_dir = Path(__file__).resolve().parent
    _ensure_dependency(project_dir)

    if str(project_dir) not in sys.path:
        sys.path.insert(0, str(project_dir))

    from pgy_gui import main as run_gui

    run_gui()


def _ensure_dependency(project_dir: Path) -> None:
    if importlib.util.find_spec("openpyxl") is not None:
        return

    requirements = project_dir / "requirements.txt"
    if not requirements.exists():
        _show_dependency_error("缺少 requirements.txt，无法自动安装依赖。")
        raise SystemExit(1)

    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", str(requirements)])
    except Exception as exc:  # noqa: BLE001 - show a friendly desktop error.
        _show_dependency_error(f"自动安装依赖失败：{exc}\n\n请先安装 Python 依赖：openpyxl")
        raise SystemExit(1) from exc


def _show_dependency_error(message: str) -> None:
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("蒲公英博主打开工具", message)
        root.destroy()
    except Exception:
        print(message, file=sys.stderr)


if __name__ == "__main__":
    main()
