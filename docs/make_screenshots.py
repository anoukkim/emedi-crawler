"""README 스크린샷 다시 만들기: python docs/make_screenshots.py

화면 없이(offscreen) 앱 창을 띄우고 예시 데이터를 채운 뒤 docs/*.png 로 저장한다.
크롬이나 사이트 접속은 필요 없다. 화면 구성이 바뀌면 다시 실행하세요.
"""

import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QT_SCALE_FACTOR", "1.5")   # 고해상도 화면에서도 선명하게

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from PyQt6.QtGui import QFont
from PyQt6.QtWidgets import QApplication

qt = QApplication(sys.argv)
qt.setStyle("Fusion")

import app as emedi  # noqa: E402  (QApplication 생성 후 import)

emedi.load_ui_fonts()
qt.setFont(QFont("IBM Plex Sans KR", 11))
emedi.LOG_FILE = Path(os.devnull)   # 예시 로그를 실제 로그 파일에 남기지 않음

OUT = ROOT / "docs"
SAVE_DIR = r"C:\Users\me\Desktop\허가정보"


def window(width=1264, height=1000):
    w = emedi.MainWindow()
    w.resize(width, height)
    w.show()
    qt.processEvents()
    w.crawl_output.setText(SAVE_DIR)
    w.merge_output.setText(SAVE_DIR)
    w.crawl_name.setText("eMedi_허가정보_20261104")
    w.merge_name.setText("eMedi_허가정보_최종_20261104")
    return w


def save(w, name):
    qt.processEvents()
    w.grab().save(str(OUT / name))
    print("saved", OUT / name, flush=True)
    w.close()


def crawl_screenshot():
    w = window()
    w._set_chrome_state(True)
    w.btn_open.setEnabled(True)
    w.btn_crawl.setText("수집 중...")
    w.btn_stop.setEnabled(True)
    w.chk_fast.setEnabled(False)
    w.lbl_crawl_hint.setText("검색 결과 목록이 보이는 상태에서 시작하세요.")
    w.recovery_file.setText(SAVE_DIR + r"\eMedi_허가정보_20261103_partial.json")
    rows = [
        (170, "수허17-119호", 42, "OK"), (171, "수인07-157호", 74, "OK"),
        (172, "제허21-250호", 68, "OK"), (173, "수신19-1022호", 39, "OK"),
        (174, "수허22-318호", 0, "FAIL"), (175, "제인18-2204호", 55, "OK"),
    ]
    for index, permit, count, status in rows:
        w._table_event(index, permit, 0, "RUN")
        w._table_event(index, permit, count, status)
    w._table_checkpoint(175)
    for index, permit, count in ((176, "수허20-447호", 61), (177, "제신23-0915호", 47),
                                 (178, "수인16-3301호", 72), (179, "제허24-061호", 58)):
        w._table_event(index, permit, 0, "RUN")
        w._table_event(index, permit, count, "OK")
    w._table_event(180, "수허23-1187호", 0, "RUN")
    w._set_stats(227, 179, 1, "08:40")
    w._set_progress(79)
    w.lbl_status.setText("수집 중 · 18/23 페이지")
    save(w, "screenshot-crawl.png")


def merge_screenshot():
    w = window(1264, 860)
    w.tabs.setCurrentIndex(1)
    w.merge_crawl.setText(SAVE_DIR + r"\eMedi_허가정보_20261104.xlsx")
    w.merge_permit.setText(SAVE_DIR + r"\품목허가(신고)사항_20261104.xlsx")
    w._log_table_mode = "merge"
    for msg, level in (
        ("크롤링 파일: 227행", "ok"),
        ("허가(신고)사항 파일: 231행 / 매칭 키: '허가번호'", "ok"),
        ("매칭 완료: 225/227건", "ok"),
        ("엑셀 저장 중... (227행 × 116열)", ""),
        ("최종 저장: " + SAVE_DIR + r"\eMedi_허가정보_최종_20261104.xlsx", "ok"),
    ):
        w._log(msg, level)
    w._set_progress(100)
    w.lbl_status.setText("완료")
    w._set_result_buttons("merge", SAVE_DIR + r"\eMedi_허가정보_최종_20261104.xlsx")
    save(w, "screenshot-merge.png")


if __name__ == "__main__":
    crawl_screenshot()
    merge_screenshot()
    os._exit(0)   # 종료 시 Qt 정리 순서 때문에 생기는 오류 방지
