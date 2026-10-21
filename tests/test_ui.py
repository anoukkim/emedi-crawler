"""Lightweight Qt layout checks; run with QT_QPA_PLATFORM=offscreen on CI."""

import importlib.util
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

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
            self.assertEqual(window.tabs.height(), 252)
            self.assertIsNotNone(window.merge_permit)
        finally:
            window.close()

    def test_dark_palette_labels_and_tab_switch_during_crawl(self):
        from PyQt6.QtCore import Qt
        from PyQt6.QtWidgets import QLabel
        from app import MainWindow

        window = MainWindow()
        try:
            window.show()
            self.application.processEvents()
            self.assertIn("color:#0a0a0a", window.lbl_status.styleSheet())
            self.assertIn("color:#0a0a0a", window.lbl_pct.styleSheet())
            heading = next(label for label in window.findChildren(QLabel)
                           if label.text().startswith("크롤링 결과에 전자민원"))
            self.assertIn("color:#0a0a0a", heading.styleSheet())
            for col in (0, 1):
                self.assertTrue(window.log_table.horizontalHeaderItem(col).textAlignment()
                                & Qt.AlignmentFlag.AlignLeft)
            for col in (2, 3):
                self.assertTrue(window.log_table.horizontalHeaderItem(col).textAlignment()
                                & Qt.AlignmentFlag.AlignRight)

            window.tabs.setCurrentIndex(1)
            window._table_event(38, "수허21-250호", 74, "OK")
            with tempfile.TemporaryDirectory() as tmp, patch("app.LOG_FILE", Path(tmp) / "ui.log"):
                window._log("[38/227] 수허21-250호 OK (74개 항목)", "ok")
                self.assertEqual(window.log_table.rowCount(), 1)
                window._log_table_mode = "merge"
                window._log("병합 완료", "ok")
                self.assertEqual(window.log_table.rowCount(), 2)
        finally:
            window.close()

    def test_window_resizes(self):
        from PyQt6.QtCore import QPoint, Qt
        from PyQt6.QtWidgets import QSizeGrip
        from app import MainWindow

        window = MainWindow()
        try:
            window.show()
            for w, h in ((900, 640), (1400, 1000)):
                window.resize(w, h)
                self.application.processEvents()
                self.assertEqual((window.width(), window.height()), (w, h))
            self.assertTrue(window.findChildren(QSizeGrip))
            self.assertEqual(window._resize_edge(QPoint(2, 300)), Qt.Edge.LeftEdge)
            self.assertEqual(window._resize_edge(QPoint(1399, 999)), Qt.Edge.RightEdge | Qt.Edge.BottomEdge)
            self.assertEqual(window._resize_edge(QPoint(500, 300)), Qt.Edge(0))
        finally:
            window.close()


if __name__ == "__main__":
    unittest.main()
