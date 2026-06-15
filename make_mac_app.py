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


def create_app_bundle(project_dir: str | Path) -> Path:
    project_path = Path(project_dir).resolve()
    entry_path = project_path / ENTRY_FILE
    if not entry_path.exists():
        raise FileNotFoundError(f"未找到入口文件：{entry_path}")

    output_app = project_path / f"{APP_NAME}.app"
    build_root = project_path / "build" / "macos-app"
    dist_dir = build_root / "dist"
    work_dir = build_root / "work"

    if output_app.exists():
        shutil.rmtree(output_app)
    if build_root.exists():
        shutil.rmtree(build_root)

    command = pyinstaller_command(project_path, dist_dir, work_dir)
    subprocess.run(command, check=True)

    built_app = dist_dir / f"{APP_NAME}.app"
    if not built_app.exists():
        raise FileNotFoundError(f"PyInstaller 未生成应用：{built_app}")

    shutil.copytree(built_app, output_app, symlinks=True)
    return output_app


def main() -> int:
    app_path = create_app_bundle(Path(__file__).resolve().parent)
    print(f"已生成：{app_path}")
    print("之后双击这个 .app 就能打开图形界面。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
