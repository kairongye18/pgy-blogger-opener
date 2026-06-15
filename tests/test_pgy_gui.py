import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from pgy_gui import INITIAL_WINDOW_SIZE, MINIMUM_WINDOW_SIZE  # noqa: E402


class PgyGuiTests(unittest.TestCase):
    def test_initial_window_is_wide_enough_for_all_settings(self):
        self.assertEqual(INITIAL_WINDOW_SIZE, "1100x600")
        self.assertEqual(MINIMUM_WINDOW_SIZE, (1000, 540))


if __name__ == "__main__":
    unittest.main()
