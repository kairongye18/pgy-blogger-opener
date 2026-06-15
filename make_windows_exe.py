#!/usr/bin/env python3
from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path


APP_NAME = "蒲公英博主打开工具"
ENTRY_FILE = "蒲公英博主打开工具.py"


def pyinstaller_command(project_dir: str | Path, dist_dir: str | Path, work_dir: str | Path) -> list[str]:
    project_path = Path(project_dir).resolve()
    return [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--windowed",
        "--name",
        APP_NAME,
        "--distpath",
        str(Path(dist_dir).resolve()),
        "--workpath",
        str(Path(work_dir).resolve()),
        "--specpath",
        str((project_path / "build").resolve()),
        str(project_path / ENTRY_FILE),
    ]


def create_windows_exe(project_dir: str | Path) -> Path:
    project_path = Path(project_dir).resolve()
    entry_path = project_path / ENTRY_FILE
    if not entry_path.exists():
        raise FileNotFoundError(f"未找到入口文件：{entry_path}")

    build_root = project_path / "build" / "windows-exe"
    dist_dir = build_root / "dist"
    work_dir = build_root / "work"
    if build_root.exists():
        shutil.rmtree(build_root)

    subprocess.run(pyinstaller_command(project_path, dist_dir, work_dir), check=True)

    exe_path = dist_dir / f"{APP_NAME}.exe"
    if not exe_path.exists():
        raise FileNotFoundError(f"PyInstaller 未生成应用：{exe_path}")
    return exe_path


def main() -> int:
    exe_path = create_windows_exe(Path(__file__).resolve().parent)
    print(f"已生成：{exe_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
