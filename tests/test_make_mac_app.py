import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from make_mac_app import APP_NAME, create_app_bundle, pyinstaller_command  # noqa: E402


class MakeMacAppTests(unittest.TestCase):
    def test_pyinstaller_command_builds_windowed_app(self):
        command = pyinstaller_command(ROOT, ROOT / "build" / "dist", ROOT / "build" / "work")

        self.assertEqual(command[:3], [sys.executable, "-m", "PyInstaller"])
        self.assertIn("--windowed", command)
        self.assertIn("--noconfirm", command)
        self.assertIn("--name", command)
        self.assertIn(APP_NAME, command)
        self.assertEqual(command[-1], str(ROOT / "蒲公英博主打开工具.py"))

    def test_create_app_bundle_copies_pyinstaller_app_to_project_root(self):
        with tempfile.TemporaryDirectory() as tmp:
            project_dir = Path(tmp)
            (project_dir / "蒲公英博主打开工具.py").write_text("print('ok')\n", encoding="utf-8")

            def fake_run(command, *, check):
                dist_dir = Path(command[command.index("--distpath") + 1])
                app_contents = dist_dir / f"{APP_NAME}.app" / "Contents"
                app_contents.mkdir(parents=True)
                (app_contents / "Info.plist").write_text("<plist />", encoding="utf-8")

            with patch("make_mac_app.subprocess.run", side_effect=fake_run):
                app_path = create_app_bundle(project_dir)

            self.assertEqual(app_path, project_dir.resolve() / f"{APP_NAME}.app")
            self.assertTrue((app_path / "Contents" / "Info.plist").exists())


if __name__ == "__main__":
    unittest.main()
