import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from make_windows_exe import APP_NAME, create_windows_exe, pyinstaller_command  # noqa: E402


class MakeWindowsExeTests(unittest.TestCase):
    def test_pyinstaller_command_builds_single_windowed_exe(self):
        command = pyinstaller_command(ROOT, ROOT / "build" / "dist", ROOT / "build" / "work")

        self.assertEqual(command[:3], [sys.executable, "-m", "PyInstaller"])
        self.assertIn("--onefile", command)
        self.assertIn("--windowed", command)
        self.assertIn("--noconfirm", command)
        self.assertIn(APP_NAME, command)
        self.assertEqual(command[-1], str(ROOT / "蒲公英博主打开工具.py"))

    def test_create_windows_exe_returns_built_executable(self):
        with tempfile.TemporaryDirectory() as tmp:
            project_dir = Path(tmp)
            (project_dir / "蒲公英博主打开工具.py").write_text("print('ok')\n", encoding="utf-8")

            def fake_run(command, *, check):
                dist_dir = Path(command[command.index("--distpath") + 1])
                dist_dir.mkdir(parents=True)
                (dist_dir / f"{APP_NAME}.exe").write_bytes(b"MZ")

            with patch("make_windows_exe.subprocess.run", side_effect=fake_run):
                exe_path = create_windows_exe(project_dir)

            self.assertEqual(exe_path.name, f"{APP_NAME}.exe")
            self.assertTrue(exe_path.exists())


if __name__ == "__main__":
    unittest.main()
