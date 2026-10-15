"""Lightweight Qt layout checks; run with QT_QPA_PLATFORM=offscreen on CI."""

import importlib.util
import os
import unittest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

HAS_GUI = all(importlib.util.find_spec(name) for name in ("PyQt6", "selenium", "pandas"))


@unittest.skipUnless(HAS_GUI, "GUI dependencies are not installed")
class LayoutTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        from PyQt6.QtWidgets import QApplication
        from app import load_ui_fonts
        cls.application = QApplication.instance() or QApplication([])
        load_ui_fonts()

    def test_grid_and_live_rows(self):
        from app import MainWindow

        window = MainWindow()
        try:
            window.resize(1264, 900)
            window.show()
            self.application.processEvents()
            self.assertGreater(window.log_table.height(), 250)
            self.assertEqual(window.tabs.height(), 165)
            self.assertEqual(window.btn_crawl.text(), "② 크롤링 시작")
            self.assertEqual(window.btn_repair.text(), "누락분만 수집")

            window._set_stats(227, 175, 5, "12:00")
            self.assertIn("180", window.tile_total.value.text())
            self.assertEqual(window.tile_pending.text(), "대기 47")
            window._table_event(181, "수인07-157호", 0, "RUN")
            self.assertEqual(window.log_table.item(0, 3).text(), "수집 중…")
            self.assertEqual(window.log_table.item(0, 1).background().color().name(), "#0a0a0a")
            window._table_event(181, "수인07-157호", 74, "OK")
            self.assertEqual(window.log_table.rowCount(), 1)
            self.assertEqual(window.log_table.item(0, 2).text(), "74")
            window._table_checkpoint(175)
            self.assertEqual(window.log_table.rowCount(), 2)

            window.tabs.setCurrentIndex(1)
            self.application.processEvents()
            self.assertEqual(window.tabs.height(), 250)
            self.assertIsNotNone(window.merge_permit)
        finally:
            window.close()


if __name__ == "__main__":
    unittest.main()
