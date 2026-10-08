"""
eMedi 허가정보 수집기 - 의료기기 허가정보(emedi.mfds.go.kr) 크롤링 · 병합
Dev: Yuri Kim | 2026-10

실행: python app.py
필요: pip install -r requirements.txt
"""

import sys, time, json, re, html, traceback, platform
from datetime import datetime
from pathlib import Path

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QLineEdit, QTextEdit, QProgressBar,
    QFileDialog, QFrame, QTabWidget, QMessageBox, QCheckBox
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QObject, QUrl
from PyQt6.QtGui import QFont, QDesktopServices, QPainter, QLinearGradient, QColor, QIcon, QPixmap

# ── Colors ────────────────────────────────────────────────
APP_NAME  = "eMedi 허가정보 수집기"
APP_DATE  = "2026-10"
PRIMARY   = "#4F46E5"   # indigo
PRIMARY_2 = "#7C3AED"   # violet (그라데이션 끝색)
SOFT      = "#EEF2FF"   # primary 연한 배경
ACCENT    = "#06B6D4"   # cyan
SUCCESS   = "#10B981"
DANGER    = "#EF4444"
BG        = "#F6F7FB"
CARD      = "#FFFFFF"
BORDER    = "#E5E7EB"
TEXT      = "#111827"
MUTED     = "#6B7280"
DARK_LOG  = "#0F172A"

BASE_URL   = "https://emedi.mfds.go.kr/search/data/MNU20237"

# ── 대기 설정 ─────────────────────────────────────────────
# 빠른 모드: 모달 내용이 나타나고 MODAL_SETTLE초 동안 변하지 않으면 바로 수집
MODAL_TIMEOUT = 20     # 모달 최대 대기 (초)
MODAL_SETTLE  = 1.0    # 내용이 이 시간 동안 그대로면 로딩 완료로 판단 (초)
PAGE_TIMEOUT  = 15     # 페이지 이동 최대 대기 (초)
POLL          = 0.25   # 확인 간격 (초)
POLITE_DELAY  = 0.3    # 항목 사이 쉬는 시간 (서버 부담 줄이기)
# 안전 모드 (기존 방식): 고정 대기
MODAL_WAIT = 10
RETRY_WAIT = 20
MAX_RETRY  = 2

LIST_COLS = [
    "순번","업체명","품목명","품목허가번호","품목등급","품목상태","취소취하일자",
    "품목분류번호","제품명","모델명","품목허가일자","품목일련번호","유효기간",
    "업허가번호","업종","업상태","업허가일자","요양급여코드",
    "추적관리대상여부","수출용에한함","위탁제조여부","유지관리용여부",
    "인체이식형여부","제조국가","제조자","제조의뢰국가명",
    "품목영문명","주소","UDI코드","포장수량",
]

MODAL_SEL = ".modal__frame.fit_modal_open"

CUR_PAGE_JS = """
var el = document.querySelector('.paging .num a[title="현재페이지"] b');
return el ? parseInt(el.innerText) : 0;
"""

FIRST_ROW_JS = """
var r = document.querySelector('table.scrollTable tbody tr');
return r ? r.innerText : '';
"""

CLOSE_MODAL_JS = """
var b = document.querySelector('#modalCloseBtn');
if (b) { try { b.click(); } catch(e) {} }
return !document.querySelector('""" + MODAL_SEL + """');
"""

PARSE_JS = """
var modal=document.querySelector('.modal__frame.fit_modal_open');
if(!modal)return{error:'no modal'};
var result={};
var sectionMap={'표준코드':'UDI','제품정보':'제품정보','수입업자':'수입업자','제조업자':'수입업자','기타정보':'기타정보','보험':'보험'};
function getPrefix(t){for(var k in sectionMap)if(t.indexOf(k)>=0)return sectionMap[k];return null;}
function cleanText(el){var c=el.cloneNode(true);c.querySelectorAll('span,img').forEach(function(s){s.remove();});return c.innerText.trim();}
modal.querySelectorAll('table').forEach(function(table){
    var caption=table.querySelector('caption');
    var title=caption?caption.innerText.trim():'';
    var prefix=getPrefix(title);if(!prefix)return;
    var colHeaders=[];
    table.querySelectorAll('tr').forEach(function(row){
        var ths=row.querySelectorAll('th'),tds=row.querySelectorAll('td');
        if(ths.length>0&&tds.length===0){colHeaders=Array.from(ths).map(function(th){return cleanText(th);});}
        else if(ths.length>0&&tds.length>0){for(var i=0;i<Math.min(ths.length,tds.length);i++){var k=cleanText(ths[i]),v=tds[i].innerText.trim();if(!k)continue;var key=prefix+'_'+k;
                if(key.indexOf('UDI_')===0){  // UDI 값엔 쉼표가 들어갈 수 있어 ; 로 구분
                    result[key]=result[key]?result[key]+'; '+v:v;
                }else{
                    result[key]=result[key]?result[key]+', '+v:v;
                }}}
        else if(tds.length>0&&colHeaders.length>0){colHeaders.forEach(function(h,i){if(!h)return;var v=i<tds.length?tds[i].innerText.trim():'';
                var key=prefix+'_'+h;
                if(key.indexOf('UDI_')===0){  // UDI 값엔 쉼표가 들어갈 수 있어 ; 로 구분
                    result[key]=result[key]?result[key]+'; '+v:v;
                }else{
                    result[key]=result[key]?result[key]+', '+v:v;
                }});}
    });
});
return result;
"""

def default_output_dir():
    """바탕화면 (OneDrive 바탕화면 포함), 없으면 홈 폴더"""
    home = Path.home()
    for p in (home / "Desktop", home / "OneDrive" / "Desktop", home / "OneDrive" / "바탕 화면"):
        if p.is_dir():
            return p
    return home


class Aborted(Exception):
    pass


# ═══════════════════════════════════════════════════════════
# Worker
# ═══════════════════════════════════════════════════════════
class CrawlWorker(QObject):
    log      = pyqtSignal(str, str)
    status   = pyqtSignal(str)
    progress = pyqtSignal(int)
    stats    = pyqtSignal(int, int, int, str)   # 전체, 완료, 실패, 남은 시간
    done     = pyqtSignal(bool, str)

    def __init__(self, driver, output_dir, file_name, fast=True):
        super().__init__()
        self.driver     = driver
        self.output_dir = Path(output_dir)
        self.file_name  = clean_file_name(file_name, default_crawl_name())
        self.fast       = fast
        self._abort     = False
        self._page      = None

    def abort(self):
        self._abort = True

    # ── 대기 헬퍼 ─────────────────────────────────────────
    def _sleep(self, sec):
        """중단 버튼에 바로 반응하는 sleep"""
        end = time.monotonic() + sec
        while True:
            if self._abort:
                raise Aborted()
            rem = end - time.monotonic()
            if rem <= 0:
                return
            time.sleep(min(POLL, rem))

    def _wait_until(self, fn, timeout):
        """fn()이 참이 될 때까지 대기. 성공 시 값, 시간 초과 시 None"""
        end = time.monotonic() + timeout
        while True:
            if self._abort:
                raise Aborted()
            try:
                v = fn()
                if v:
                    return v
            except Aborted:
                raise
            except Exception:
                pass
            if time.monotonic() >= end:
                return None
            time.sleep(POLL)

    def run(self):
        all_records, items, finished = [], [], False
        partial_path = self.output_dir / f"{self.file_name}_partial.json"
        try:
            self.output_dir.mkdir(parents=True, exist_ok=True)
            mode = "빠른 모드" if self.fast else "안전 모드 (고정 대기)"
            self.log.emit(f"▶  목록 수집 시작  ·  {mode}", "section")
            items = self._collect_all_list()
            self.log.emit(f"  총 {len(items)}개 허가증 확인", "ok")
            if not items:
                self.done.emit(False, "목록이 비어 있습니다. 검색 결과가 보이는지 확인하세요.")
                return

            self.log.emit(f"\n▶  상세 정보 수집 ({len(items)}건)", "section")
            t0 = time.monotonic()
            n_fail = 0
            self.stats.emit(len(items), 0, 0, "계산 중")

            for idx, item in enumerate(items, 1):
                if self._abort:
                    raise Aborted()
                permit_no = item.get("품목허가번호", f"row_{idx}")
                self.status.emit(f"[{idx}/{len(items)}]  {permit_no} 수집 중")
                self.progress.emit(int(idx / len(items) * 95))

                record = {k: v for k, v in item.items() if not k.startswith("_")}
                try:
                    self._goto_page(item["_page"])
                    record.update(self._fetch_modal(item))
                    self.log.emit(f"  [{idx:3d}/{len(items)}]  {permit_no}  ✓", "ok")
                except Aborted:
                    raise
                except Exception as e:
                    self.log.emit(f"  [{idx:3d}/{len(items)}]  {permit_no}  →  {e}", "err")
                    record["오류"] = str(e)
                finally:
                    self._close_modal()
                all_records.append(record)
                n_fail += "오류" in record
                per = (time.monotonic() - t0) / idx
                left = int(per * (len(items) - idx))
                eta = f"{left // 60}분" if left >= 60 else f"{left}초"
                self.stats.emit(len(items), idx - n_fail, n_fail, eta)

                if idx % 25 == 0:
                    self._save_json(all_records, partial_path)
                    self.log.emit(f"  ★ 중간 저장 ({idx}건)", "dim")

                if self.fast:
                    self._sleep(POLITE_DELAY)

            finished = True

        except Aborted:
            self.log.emit(f"\n⏹  중단됨 ({len(all_records)}건 수집 완료)", "warn")
        except Exception as e:
            self.log.emit(f"\n❌  오류: {e}", "err")
            self.log.emit(traceback.format_exc(), "dim")

        # ── 어떤 경우든 지금까지 모은 데이터는 저장 ──
        if not all_records:
            if not finished:
                self.done.emit(False, "저장할 데이터가 없습니다.")
            return
        try:
            self.log.emit("\n▶  엑셀 변환 중", "section")
            xlsx = self._to_excel(all_records)
            self._save_json(all_records, xlsx.with_suffix(".json"))
            if partial_path.exists():
                partial_path.unlink()
        except Exception as e:
            self.done.emit(False, f"저장 실패: {e}")
            return

        failed = sum(1 for r in all_records if "오류" in r)
        if failed:
            self.log.emit(f"  ⚠  실패 {failed}건 → 엑셀의 '오류' 열 확인", "warn")
        self.progress.emit(100)
        if finished:
            self.done.emit(True, str(xlsx))
        else:
            self.done.emit(False, f"중단됨 ({len(all_records)}건 저장됨)\n{xlsx}")

    # ── 페이지 이동 ───────────────────────────────────────
    def _current_page(self):
        try:
            return self.driver.execute_script(CUR_PAGE_JS) or 0
        except Exception:
            return 0

    def _goto_page(self, page):
        cur = self._current_page()
        # cur == 0: 페이지 표시를 못 찾음 (1페이지뿐인 경우 등) → 마지막으로 이동한 페이지 기준
        if cur == page or (cur == 0 and self._page == page):
            return
        before = self.driver.execute_script(FIRST_ROW_JS)
        self.driver.execute_script("getList(arguments[0])", page)
        ok = self._wait_until(
            lambda: self._current_page() in (page, 0)
                    and self.driver.execute_script(FIRST_ROW_JS) not in ("", before),
            PAGE_TIMEOUT)
        if not ok:
            raise Exception(f"{page}페이지 이동 실패 (시간 초과)")
        self._page = page
        self._sleep(0.3)  # 행 렌더링 마무리

    # ── 상세 모달 ─────────────────────────────────────────
    def _fetch_modal(self, item):
        args = [item["_meddevseq"], item["_typename"], item["_index"],
                item["_mdentpsno"], item["_isfirst"], "", item["_seqgroup"]]

        for attempt in range(1, MAX_RETRY + 2):
            if attempt > 1:
                self.log.emit(f"    재시도 {attempt-1}회...", "warn")

            self._close_modal()
            # 인자를 문자열로 이어붙이지 않고 그대로 전달 (따옴표 등 특수문자 안전)
            self.driver.execute_script("itemDetail.apply(null, arguments)", *args)

            if self.fast:
                detail = self._wait_modal_settled(MODAL_TIMEOUT if attempt == 1 else MODAL_TIMEOUT * 2)
            else:
                self._sleep(RETRY_WAIT if attempt > 1 else MODAL_WAIT)
                detail = self.driver.execute_script(PARSE_JS) or {}
                if detail.get("error"):
                    detail = None
            if detail:
                return detail
            self._close_modal()
            self._sleep(1)

        raise Exception(f"모달 로딩 실패 ({MAX_RETRY}회 재시도)")

    def _wait_modal_settled(self, timeout):
        """모달 내용이 나타나고 MODAL_SETTLE초 동안 바뀌지 않으면 반환"""
        end = time.monotonic() + timeout
        last, stable_since = None, None
        while time.monotonic() < end:
            if self._abort:
                raise Aborted()
            try:
                d = self.driver.execute_script(PARSE_JS) or {}
            except Exception:
                d = {}
            if d and not d.get("error"):
                if d == last:
                    if time.monotonic() - stable_since >= MODAL_SETTLE:
                        return d
                else:
                    last, stable_since = d, time.monotonic()
            time.sleep(POLL)
        return last  # 시간 초과: 마지막으로 본 내용이라도 반환 (없으면 None)

    def _close_modal(self):
        try:
            self._wait_until(lambda: self.driver.execute_script(CLOSE_MODAL_JS), 5)
        except Aborted:
            raise
        except Exception:
            pass

    # ── 목록 ──────────────────────────────────────────────
    def _collect_all_list(self):
        JS = """
        var rows=document.querySelectorAll('table.scrollTable tbody tr'),result=[];
        rows.forEach(function(row){
            var tds=row.querySelectorAll('td'),cells=Array.from(tds).map(function(td){return td.innerText.trim();});
            var link=row.querySelector('a[data-meddevseq]'),attrs={};
            if(link){attrs={meddevseq:link.getAttribute('data-meddevseq')||'',
            typename:link.getAttribute('data-typename')||'',
            mdentpsno:link.getAttribute('data-mdentpsno')||'',
            isfirst:link.getAttribute('data-isfirst')||'',
            seqgroup:link.getAttribute('data-seqgroup')||'',
            index:link.getAttribute('data-index')||''};}
            if(cells.length>0)result.push({cells:cells,attrs:attrs});
        });return result;
        """
        all_items = []
        last = self._get_last_page()
        self.log.emit(f"  총 {last or '?'}페이지", "dim")
        page = self._current_page() or 1
        self._page = page
        if page != 1:
            self._goto_page(1)
            page = 1
        while True:
            rows = self.driver.execute_script(JS)
            items = []
            for row in rows:
                cells, attrs = row.get("cells", []), row.get("attrs", {})
                if not cells or not any(cells): continue
                item = {col: (cells[i] if i < len(cells) else "") for i, col in enumerate(LIST_COLS)}
                for k in ["meddevseq","typename","mdentpsno","isfirst","seqgroup","index"]:
                    item[f"_{k}"] = attrs.get(k, "")
                item["_page"] = page  # 상세 수집 때 이 페이지로 이동
                items.append(item)
            if not items: break
            all_items.extend(items)
            self.log.emit(f"  페이지 {page:3d}: {len(items):2d}건  (누적 {len(all_items):4d}건)", "dim")
            if last and page >= last: break
            if not last and len(items) < 10: break
            try:
                self._goto_page(page + 1)
            except Aborted:
                raise
            except Exception:
                if last:
                    raise
                break  # 마지막 페이지를 모르는 경우: 더 이상 넘어가지 않으면 종료
            page += 1
        return all_items

    def _get_last_page(self):
        try:
            els = self.driver.find_elements(By.XPATH, "//a[.//span[contains(text(),'맨 마지막')]]")
            if not els:
                return None
            m = re.search(r"getList\((\d+)\)", els[0].get_attribute("onclick") or "")
            return int(m.group(1)) if m else None
        except Exception:
            return None

    def _to_excel(self, records):
        clean = [{k:v for k,v in r.items() if not k.startswith("_")} for r in records]
        df = pd.json_normalize(clean)

        # dict/list 타입 컬럼 → str 변환
        for col in df.columns:
            if df[col].apply(lambda x: isinstance(x, (dict, list))).any():
                df[col] = df[col].astype(str)

        df = df.fillna("")

        if "UDI_모델명" in df.columns:
            def uniq(val):
                if not val: return ""
                seen = []
                for m in str(val).split("; "):
                    m = m.strip()
                    if m and m not in seen: seen.append(m)
                return "|".join(seen)
            pos = df.columns.get_loc("UDI_모델명") + 1
            df.insert(pos, "모델명_고유", df["UDI_모델명"].apply(uniq))

        if len(df.columns) > 16384:
            self.log.emit(f"  ⚠  컬럼 {len(df.columns)}개 → 16384 초과분 잘라냄", "warn")
            df = df.iloc[:, :16384]

        path = unique_path(self.output_dir / f"{self.file_name}.xlsx")
        self.log.emit(f"  엑셀 저장 중... ({len(df)}행 × {len(df.columns)}열)", "dim")
        df.to_excel(path, index=False, engine="openpyxl")
        self.log.emit(f"  ✓  저장: {path}", "ok")
        return path

    def _save_json(self, records, path):
        tmp = Path(str(path) + ".tmp")
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(records, f, ensure_ascii=False, indent=2)
        tmp.replace(path)


def default_crawl_name():
    return f"eMedi_허가정보_{datetime.now():%Y%m%d}"


def default_merge_name():
    return f"eMedi_허가정보_최종_{datetime.now():%Y%m%d}"


def clean_file_name(name, fallback):
    """파일 이름에 쓸 수 없는 문자 제거, .xlsx 확장자는 떼기"""
    name = re.sub(r'[\\/:*?"<>|]', "_", (name or "").strip())
    if name.lower().endswith(".xlsx"):
        name = name[:-5]
    return name.strip(" .") or fallback


def unique_path(path):
    """같은 이름 파일이 있거나 엑셀에서 열려 있으면 _2, _3 ... 붙이기"""
    path = Path(path)
    if not path.exists():
        return path
    n = 2
    while True:
        p = path.with_name(f"{path.stem}_{n}{path.suffix}")
        if not p.exists():
            return p
        n += 1



def format_excel(ws, df, highlight_cols=None):
    """Excel 워크시트 포매팅"""
    from openpyxl.styles import (
        PatternFill, Font, Alignment, Border, Side, GradientFill
    )
    from openpyxl.utils import get_column_letter

    HEADER      = PRIMARY.lstrip("#")
    HEADER_HL   = ACCENT.lstrip("#")
    HIGHLIGHT   = "ECFEFF"   # 전자민원 컬럼 배경
    ROW_EVEN    = "F8FAFC"   # 짝수행
    ROW_ODD     = "FFFFFF"   # 홀수행
    BORDER_COL  = "E5E7EB"

    thin = Side(style="thin", color=BORDER_COL)
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # ── 헤더 스타일 ──
    for col_idx, col_name in enumerate(df.columns, 1):
        cell = ws.cell(row=1, column=col_idx)
        is_highlight = highlight_cols and col_name in highlight_cols

        cell.fill = PatternFill("solid", fgColor=HEADER_HL if is_highlight else HEADER)
        cell.font = Font(
            bold=True, color="FFFFFF",
            name="맑은 고딕", size=10
        )
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
        cell.border = border

    # ── 행 스타일 ──
    n_cols = len(df.columns)
    for row_idx in range(2, ws.max_row + 1):
        bg = ROW_EVEN if row_idx % 2 == 0 else ROW_ODD
        for col_idx in range(1, n_cols + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            col_name = df.columns[col_idx - 1]
            is_highlight = highlight_cols and col_name in highlight_cols

            cell.fill = PatternFill("solid", fgColor=HIGHLIGHT if is_highlight else bg)
            cell.font = Font(name="맑은 고딕", size=10)
            cell.alignment = Alignment(
                vertical="center", wrap_text=False,
                horizontal="center" if col_name in [
                    "순번","품목등급","품목상태","인체이식형여부","일회용",
                    "추적관리대상","수출용에한함","UDI_코드체계"
                ] else "left"
            )
            cell.border = border

    # ── 열 너비 자동 조정 ──
    for col_idx, col_name in enumerate(df.columns, 1):
        col_letter = get_column_letter(col_idx)
        header_len = len(str(col_name)) * 1.8
        # NaN/float 안전하게 str 변환 후 길이 계산
        sample = df.iloc[:50][col_name].fillna("").astype(str)
        lengths = sample.map(lambda x: len(x))
        max_data = lengths.max() if len(lengths) > 0 else 0
        if pd.isna(max_data):
            max_data = 0
        width = min(max(header_len, float(max_data) * 1.1, 8), 60)
        ws.column_dimensions[col_letter].width = width

    # ── 헤더 행 높이 ──
    ws.row_dimensions[1].height = 32

    # ── 틀 고정 (첫 행) ──
    ws.freeze_panes = "A2"

    # ── 자동 필터 ──
    ws.auto_filter.ref = ws.dimensions

class MergeWorker(QObject):
    log      = pyqtSignal(str, str)
    status   = pyqtSignal(str)
    progress = pyqtSignal(int)
    done     = pyqtSignal(bool, str)

    def __init__(self, crawl_file, permit_file, output_dir, file_name):
        super().__init__()
        self.crawl_file  = crawl_file
        self.permit_file = permit_file
        self.output_dir  = Path(output_dir)
        self.file_name   = clean_file_name(file_name, default_merge_name())

    def run(self):
        try:
            self.log.emit("▶  파일 로딩 중", "section")
            self.progress.emit(10)

            df = pd.read_excel(self.crawl_file, engine="openpyxl")
            self.log.emit(f"  크롤링 파일: {len(df)}행", "ok")
            if "품목허가번호" not in df.columns:
                raise Exception("크롤링 파일에 '품목허가번호' 열이 없습니다. 파일 1을 확인하세요.")
            self.progress.emit(30)

            # A~F 열을 한 번만 읽음. 허가번호 컬럼(두 번째)은 매칭 키, 나머지에 전자민원_ 접두사
            df2 = pd.read_excel(self.permit_file, engine="openpyxl", usecols="A:F")
            orig_cols = list(df2.columns)
            if len(orig_cols) < 2:
                raise Exception("허가(신고)사항 파일의 열이 2개 미만입니다. 파일 2를 확인하세요.")
            permit_no_col = orig_cols[1]
            merge_cols = {col: f"전자민원_{col}" for col in orig_cols if col != permit_no_col}
            df2 = df2.rename(columns=merge_cols)
            renamed_cols = list(merge_cols.values())
            self.log.emit(f"  허가(신고)사항 파일: {len(df2)}행  /  매칭 키: '{permit_no_col}'  /  추가 컬럼: {renamed_cols}", "ok")
            self.progress.emit(50)

            self.log.emit("▶  매칭 중", "section")
            df["_key"]  = df["품목허가번호"].astype(str).str.replace(r"\s", "", regex=True)
            df2["_key"] = df2[permit_no_col].astype(str).str.replace(r"\s", "", regex=True)
            df2 = df2.drop_duplicates(subset=["_key"], keep="first")

            df = df.merge(
                df2[["_key"] + renamed_cols],
                on="_key", how="left", indicator="_m"
            )
            matched = int((df["_m"] == "both").sum())
            df = df.drop(columns=["_key", "_m"])
            self.log.emit(f"  매칭 완료: {matched}/{len(df)}건", "ok")
            if matched == 0:
                self.log.emit("  ⚠  매칭된 건이 없습니다. 파일 2의 두 번째 열이 허가번호인지 확인하세요.", "warn")
            self.progress.emit(80)

            from openpyxl import load_workbook
            self.output_dir.mkdir(parents=True, exist_ok=True)
            out = unique_path(self.output_dir / f"{self.file_name}.xlsx")
            df_save = df.copy()
            for col in df_save.columns:
                if df_save[col].apply(lambda x: isinstance(x, (dict, list))).any():
                    df_save[col] = df_save[col].astype(str)
            df_save = df_save.fillna("")
            self.log.emit(f"  엑셀 저장 중... ({len(df_save)}행 × {len(df_save.columns)}열)", "dim")
            df_save.to_excel(out, index=False, engine="openpyxl")
            wb = load_workbook(out)
            format_excel(wb.active, df_save, highlight_cols=set(renamed_cols) if renamed_cols else None)
            wb.save(out)
            self.log.emit(f"  ✓  최종 저장: {out}", "ok")
            self.progress.emit(100)
            self.status.emit("완료 ✓")
            self.done.emit(True, str(out))

        except Exception as e:
            self.log.emit(f"❌  오류: {e}", "err")
            self.done.emit(False, str(e))


class ChromeOpener(QObject):
    """크롬 실행은 오래 걸릴 수 있으므로 별도 스레드에서. 결과는 시그널로 UI에 전달."""
    done = pyqtSignal(object, str)   # driver 또는 None, 오류 메시지

    def run(self):
        try:
            opts = Options()
            opts.add_argument("--window-size=1440,900")
            opts.add_argument("--lang=ko-KR")
            # Selenium 4.11+ 내장 Selenium Manager가 크롬 버전에 맞는 드라이버를 자동으로 받음
            driver = webdriver.Chrome(options=opts)
            driver.implicitly_wait(0)   # 대기는 코드에서 직접 처리 (암묵적 대기는 느려짐)
            driver.get(BASE_URL)
            self.done.emit(driver, "")
        except Exception as e:
            self.done.emit(None, str(e).splitlines()[0] if str(e) else repr(e))


# ═══════════════════════════════════════════════════════════
# Main Window
# ═══════════════════════════════════════════════════════════
class AppLogo(QWidget):
    """헤더 왼쪽 앱 아이콘 (그라데이션 둥근 사각형)"""
    def __init__(self, size=38):
        super().__init__()
        self.setFixedSize(size, size)

    def paintEvent(self, e):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        g = QLinearGradient(0, 0, self.width(), self.height())
        g.setColorAt(0, QColor(PRIMARY))
        g.setColorAt(1, QColor(PRIMARY_2))
        p.setBrush(g)
        p.setPen(Qt.PenStyle.NoPen)
        p.drawRoundedRect(self.rect(), 10, 10)
        p.setPen(QColor("white"))
        f = QFont(self.font())
        f.setPixelSize(int(self.height() * 0.5))
        f.setBold(True)
        p.setFont(f)
        p.drawText(self.rect(), Qt.AlignmentFlag.AlignCenter, "e")


class DropLineEdit(QLineEdit):
    """파일을 끌어다 놓으면 경로가 입력되는 입력칸"""
    def __init__(self, *a, **kw):
        super().__init__(*a, **kw)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()

    def dropEvent(self, e):
        urls = e.mimeData().urls()
        if urls:
            self.setText(urls[0].toLocalFile())


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.setMinimumSize(760, 860)
        self.resize(840, 940)
        self.driver   = None
        self.worker   = None
        self._threads = []  # 완료 전까지 참조 유지
        self._last_crawl_file = None
        self._last_merge_file = None
        self._build_ui()

    def _build_ui(self):
        central = QWidget()
        central.setObjectName("central")
        central.setStyleSheet(f"""
            #central {{ background:{BG}; }}
            QLabel {{ background:transparent; }}
            QToolTip {{ background:{TEXT}; color:white; border:none; padding:6px; border-radius:6px; }}
        """)
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(0,0,0,0)
        root.setSpacing(0)

        # ── 헤더 ──────────────────────────────────────────
        hdr = QFrame()
        hdr.setObjectName("hdr")
        hdr.setFixedHeight(72)
        hdr.setStyleSheet(f"#hdr {{ background:{CARD}; border-bottom:1px solid {BORDER}; }}")
        hl = QHBoxLayout(hdr)
        hl.setContentsMargins(28,0,28,0)
        hl.setSpacing(12)

        titles = QVBoxLayout()
        titles.setSpacing(1)
        titles.setContentsMargins(0,14,0,14)
        lbl_title = QLabel(APP_NAME)
        lbl_title.setStyleSheet(f"color:{TEXT}; font-size:16px; font-weight:800;")
        lbl_sub = QLabel("의료기기 허가정보 수집 · 병합  ·  emedi.mfds.go.kr")
        lbl_sub.setStyleSheet(f"color:{MUTED}; font-size:11px;")
        titles.addWidget(lbl_title)
        titles.addWidget(lbl_sub)

        self.lbl_chrome = QLabel()
        self.lbl_chrome.setFixedHeight(28)
        self._set_chrome_state(False)

        hl.addWidget(AppLogo())
        hl.addLayout(titles)
        hl.addStretch()
        hl.addWidget(self.lbl_chrome, 0, Qt.AlignmentFlag.AlignVCenter)
        root.addWidget(hdr)

        # ── 탭 (알약 모양) ────────────────────────────────
        self.tabs = QTabWidget()
        self.tabs.setDocumentMode(True)
        self.tabs.setStyleSheet(f"""
            QTabWidget::pane {{ border:none; background:{BG}; }}
            QTabWidget::tab-bar {{ left:24px; }}
            QTabBar {{ background:transparent; }}
            QTabBar::tab {{
                background:transparent; color:{MUTED};
                font-size:12px; font-weight:700;
                padding:9px 20px; margin:16px 6px 0 0;
                border:none; border-radius:16px;
            }}
            QTabBar::tab:selected {{ background:{SOFT}; color:{PRIMARY}; }}
            QTabBar::tab:hover:!selected {{ color:{TEXT}; background:#ECEEF3; }}
        """)
        self.tabs.addTab(self._build_tab1(), "①  크롤링")
        self.tabs.addTab(self._build_tab2(), "②  허가사항 병합")
        root.addWidget(self.tabs, 1)

        # ── 하단 공통: 진행 상황 + 로그 ─────────────────────
        bottom = QFrame()
        bottom.setObjectName("bottom")
        bottom.setStyleSheet(f"#bottom {{ background:{CARD}; border-top:1px solid {BORDER}; }}")
        bl = QVBoxLayout(bottom)
        bl.setContentsMargins(28,14,28,12)
        bl.setSpacing(8)

        row_prog = QHBoxLayout()
        self.lbl_status = QLabel("대기 중")
        self.lbl_status.setStyleSheet(f"color:{TEXT}; font-size:11px; font-weight:600;")
        self.lbl_pct = QLabel("0%")
        self.lbl_pct.setStyleSheet(f"color:{PRIMARY}; font-size:11px; font-weight:800;")
        row_prog.addWidget(self.lbl_status)
        row_prog.addStretch()
        row_prog.addWidget(self.lbl_pct)
        bl.addLayout(row_prog)

        self.progress = QProgressBar()
        self.progress.setFixedHeight(8)
        self.progress.setTextVisible(False)
        self.progress.setStyleSheet(f"""
            QProgressBar {{ background:#EEF0F4; border-radius:4px; border:none; }}
            QProgressBar::chunk {{ border-radius:4px;
                background:qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 {PRIMARY}, stop:1 {ACCENT}); }}
        """)
        bl.addWidget(self.progress)

        log_hdr = QHBoxLayout()
        lbl_log = QLabel("LOG")
        lbl_log.setStyleSheet(f"color:{MUTED}; font-size:10px; font-weight:800; letter-spacing:1px;")
        log_hdr.addWidget(lbl_log)
        log_hdr.addStretch()
        log_hdr.addWidget(self._link_btn("지우기", lambda: self.log_box.clear()))
        bl.addLayout(log_hdr)

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setMinimumHeight(150)
        self.log_box.setStyleSheet(f"""
            QTextEdit {{
                background:{DARK_LOG}; color:#94A3B8;
                font-family:Menlo,Consolas,monospace; font-size:11px;
                border-radius:10px; border:none; padding:10px;
            }}
            QScrollBar:vertical {{ background:transparent; width:8px; margin:4px 2px; }}
            QScrollBar::handle:vertical {{ background:#334155; border-radius:3px; min-height:24px; }}
            QScrollBar::add-line, QScrollBar::sub-line {{ height:0; }}
            QScrollBar::add-page, QScrollBar::sub-page {{ background:transparent; }}
        """)
        bl.addWidget(self.log_box)

        credit = QLabel(f"Dev. Yuri Kim  ·  {APP_DATE}")
        credit.setAlignment(Qt.AlignmentFlag.AlignRight)
        credit.setStyleSheet(f"color:{MUTED}; font-size:9px;")
        bl.addWidget(credit)

        root.addWidget(bottom)

    # ── 탭 1: 크롤링 ──────────────────────────────────────
    def _build_tab1(self):
        w = QWidget()
        layout = QVBoxLayout(w)
        layout.setContentsMargins(24,16,24,16)
        layout.setSpacing(14)

        # STEP 1
        card1, c1 = self._card("1", "사이트 열기",
            "크롬이 열리면 검색 조건을 입력하고 검색하세요.")
        self.btn_open = self._btn("크롬 열기", PRIMARY, self._open_site)
        c1.addWidget(self.btn_open)
        layout.addWidget(card1)

        # STEP 2
        card2, c2 = self._card("2", "크롤링",
            "검색 결과 목록이 보이면 시작하세요. "
            f"실패한 항목은 최대 {MAX_RETRY}회 재시도하고, 중단해도 모은 데이터는 저장됩니다.")
        self.crawl_output, self.crawl_name = self._save_fields(c2, default_crawl_name())

        self.chk_fast = QCheckBox("빠른 모드")
        self.chk_fast.setChecked(True)
        self.chk_fast.setToolTip("켜짐: 상세 정보가 뜨면 바로 수집 (권장)\n"
                                 "꺼짐: 항목마다 10초씩 고정 대기 (예전 방식, 느리지만 확실)")
        self.chk_fast.setStyleSheet(f"QCheckBox {{ color:{TEXT}; font-size:11px; font-weight:600; spacing:8px; }}")
        fast_hint = QLabel("끄면 예전처럼 항목당 10초 고정 대기")
        fast_hint.setStyleSheet(f"color:{MUTED}; font-size:10px;")
        row_fast = QHBoxLayout()
        row_fast.addWidget(self.chk_fast)
        row_fast.addWidget(fast_hint)
        row_fast.addStretch()
        c2.addLayout(row_fast)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)
        self.btn_crawl = self._btn("▶  크롤링 시작", SUCCESS, self._start_crawl)
        self.btn_crawl.setEnabled(False)
        self.btn_stop = self._btn("■  중단", DANGER, self._stop_crawl)
        self.btn_stop.setEnabled(False)
        self.btn_crawl_open = self._ghost_btn("파일 열기", lambda: self._open_path(self._last_crawl_file))
        self.btn_crawl_dir  = self._ghost_btn("폴더 열기",
            lambda: self._open_path(Path(self._last_crawl_file).parent if self._last_crawl_file else None))
        btn_row.addWidget(self.btn_crawl, 3)
        btn_row.addWidget(self.btn_stop, 2)
        btn_row.addWidget(self.btn_crawl_open, 2)
        btn_row.addWidget(self.btn_crawl_dir, 2)
        c2.addLayout(btn_row)
        self._set_result_buttons("crawl", None)

        tiles = QHBoxLayout()
        tiles.setSpacing(8)
        self.tile_total = self._tile("전체", "—", TEXT)
        self.tile_done  = self._tile("완료", "—", SUCCESS)
        self.tile_fail  = self._tile("실패", "—", DANGER)
        self.tile_eta   = self._tile("남은 시간", "—", PRIMARY)
        for t in (self.tile_total, self.tile_done, self.tile_fail, self.tile_eta):
            tiles.addWidget(t)
        c2.addLayout(tiles)
        layout.addWidget(card2)

        layout.addStretch()
        return w

    # ── 탭 2: 허가사항 병합 ───────────────────────────────
    def _build_tab2(self):
        w = QWidget()
        layout = QVBoxLayout(w)
        layout.setContentsMargins(24,16,24,16)
        layout.setSpacing(14)

        card, c = self._card("+", "크롤링 결과에 전자민원 허가사항 붙이기",
            "파일을 끌어다 놓거나 '선택'을 누르세요. 크롤링이 끝나면 파일 1은 자동으로 채워집니다.")

        c.addWidget(self._field_label("파일 1  ·  크롤링 결과 엑셀"))
        row1 = QHBoxLayout()
        self.merge_crawl = DropLineEdit()
        self.merge_crawl.setPlaceholderText("여기에 파일을 끌어다 놓기")
        self._style_input(self.merge_crawl)
        row1.addWidget(self.merge_crawl)
        row1.addWidget(self._small_btn("선택", lambda: self._pick_file(self.merge_crawl)))
        c.addLayout(row1)

        c.addWidget(self._field_label("파일 2  ·  품목 허가(신고)사항  (전자민원 → 나의민원 → 엑셀 다운로드)"))
        row2 = QHBoxLayout()
        self.merge_permit = DropLineEdit()
        self.merge_permit.setPlaceholderText("품목+허가(신고)사항.xlsx")
        self._style_input(self.merge_permit)
        row2.addWidget(self.merge_permit)
        row2.addWidget(self._small_btn("선택", lambda: self._pick_file(self.merge_permit)))
        c.addLayout(row2)

        self.merge_output, self.merge_name = self._save_fields(c, default_merge_name())

        c.addSpacing(4)
        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)
        self.btn_merge = self._btn("파일 병합 시작", PRIMARY_2, self._start_merge)
        self.btn_merge_open = self._ghost_btn("파일 열기", lambda: self._open_path(self._last_merge_file))
        self.btn_merge_dir  = self._ghost_btn("폴더 열기",
            lambda: self._open_path(Path(self._last_merge_file).parent if self._last_merge_file else None))
        btn_row.addWidget(self.btn_merge, 5)
        btn_row.addWidget(self.btn_merge_open, 2)
        btn_row.addWidget(self.btn_merge_dir, 2)
        c.addLayout(btn_row)
        self._set_result_buttons("merge", None)
        layout.addWidget(card)

        layout.addStretch()
        return w

    # ── UI 헬퍼 ───────────────────────────────────────────
    def _save_fields(self, layout, default_name):
        """저장 폴더 + 파일 이름 입력칸"""
        grid = QHBoxLayout()
        grid.setSpacing(10)

        col_dir = QVBoxLayout()
        col_dir.setSpacing(6)
        col_dir.addWidget(self._field_label("저장 폴더"))
        row_dir = QHBoxLayout()
        out = DropLineEdit(str(default_output_dir()))
        self._style_input(out)
        row_dir.addWidget(out)
        row_dir.addWidget(self._small_btn("변경", lambda: self._pick_dir(out)))
        col_dir.addLayout(row_dir)

        col_name = QVBoxLayout()
        col_name.setSpacing(6)
        col_name.addWidget(self._field_label("파일 이름"))
        row_name = QHBoxLayout()
        row_name.setSpacing(6)
        name = QLineEdit(default_name)
        name.setPlaceholderText(default_name)
        name.setToolTip("같은 이름의 파일이 있으면 _2, _3 … 이 붙어 저장됩니다.")
        self._style_input(name)
        ext = QLabel(".xlsx")
        ext.setStyleSheet(f"color:{MUTED}; font-size:11px; font-weight:600;")
        row_name.addWidget(name)
        row_name.addWidget(ext)
        col_name.addLayout(row_name)

        grid.addLayout(col_dir, 3)
        grid.addLayout(col_name, 2)
        layout.addLayout(grid)
        return out, name

    def _card(self, badge_text, title, desc):
        card = QFrame()
        card.setObjectName("card")
        card.setStyleSheet(f"#card {{ background:{CARD}; border:1px solid {BORDER}; border-radius:14px; }}")
        v = QVBoxLayout(card)
        v.setContentsMargins(22,18,22,20)
        v.setSpacing(10)

        row = QHBoxLayout()
        badge = QLabel(badge_text)
        badge.setFixedSize(28, 28)
        badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
        badge.setStyleSheet(f"background:{SOFT}; color:{PRIMARY}; font-size:13px; font-weight:800; border-radius:8px;")
        lbl = QLabel(title)
        lbl.setStyleSheet(f"color:{TEXT}; font-size:14px; font-weight:800;")
        row.addWidget(badge)
        row.addSpacing(6)
        row.addWidget(lbl)
        row.addStretch()
        v.addLayout(row)

        d = QLabel(desc)
        d.setStyleSheet(f"color:{MUTED}; font-size:11px;")
        d.setWordWrap(True)
        v.addWidget(d)
        return card, v

    def _field_label(self, text):
        l = QLabel(text)
        l.setStyleSheet(f"color:{TEXT}; font-size:10px; font-weight:700; padding-top:4px;")
        return l

    def _tile(self, label, value, color):
        t = QFrame()
        t.setObjectName("tile")
        t.setStyleSheet(f"#tile {{ background:{BG}; border:1px solid {BORDER}; border-radius:10px; }}")
        v = QVBoxLayout(t)
        v.setContentsMargins(14,10,14,10)
        v.setSpacing(2)
        cap = QLabel(label)
        cap.setStyleSheet(f"color:{MUTED}; font-size:10px; font-weight:700;")
        val = QLabel(value)
        val.setStyleSheet(f"color:{color}; font-size:20px; font-weight:800;")
        v.addWidget(cap)
        v.addWidget(val)
        t.value = val
        return t

    def _btn(self, text, color, slot):
        b = QPushButton(text)
        b.setFixedHeight(42)
        b.setCursor(Qt.CursorShape.PointingHandCursor)
        b.setStyleSheet(f"""
            QPushButton {{
                background:{color}; color:white;
                font-size:12px; font-weight:700;
                border-radius:10px; border:none; padding:0 18px;
            }}
            QPushButton:hover {{ background:{self._dk(color)}; }}
            QPushButton:pressed {{ background:{self._dk(self._dk(color))}; }}
            QPushButton:disabled {{ background:#EEF0F4; color:#A1A7B3; }}
        """)
        b.clicked.connect(slot)
        return b

    def _ghost_btn(self, text, slot):
        b = QPushButton(text)
        b.setFixedHeight(42)
        b.setCursor(Qt.CursorShape.PointingHandCursor)
        b.setStyleSheet(f"""
            QPushButton {{
                background:{CARD}; color:{TEXT};
                font-size:12px; font-weight:700;
                border-radius:10px; border:1px solid {BORDER}; padding:0 14px;
            }}
            QPushButton:hover {{ border-color:{PRIMARY}; color:{PRIMARY}; background:{SOFT}; }}
            QPushButton:disabled {{ color:#C0C4CC; background:{CARD}; border-color:#EEF0F4; }}
        """)
        b.clicked.connect(slot)
        return b

    def _small_btn(self, text, slot):
        b = QPushButton(text)
        b.setFixedHeight(36)
        b.setCursor(Qt.CursorShape.PointingHandCursor)
        b.setStyleSheet(f"""
            QPushButton {{
                background:{SOFT}; color:{PRIMARY};
                font-size:11px; font-weight:700;
                border-radius:8px; border:none; padding:0 14px;
            }}
            QPushButton:hover {{ background:#E0E7FF; }}
        """)
        b.clicked.connect(slot)
        return b

    def _link_btn(self, text, slot):
        b = QPushButton(text)
        b.setCursor(Qt.CursorShape.PointingHandCursor)
        b.setStyleSheet(f"""
            QPushButton {{ background:transparent; color:{MUTED}; border:none; font-size:10px; }}
            QPushButton:hover {{ color:{PRIMARY}; text-decoration:underline; }}
        """)
        b.clicked.connect(slot)
        return b

    def _style_input(self, edit):
        edit.setFixedHeight(36)
        edit.setStyleSheet(f"""
            QLineEdit {{
                background:{CARD}; border:1px solid {BORDER};
                border-radius:8px; padding:0 10px;
                font-size:11px; color:{TEXT};
            }}
            QLineEdit:focus {{ border:1.5px solid {PRIMARY}; }}
        """)

    def _set_chrome_state(self, connected):
        if connected:
            fg, bg, text = "#047857", "#D1FAE5", "크롬 연결됨"
        else:
            fg, bg, text = MUTED, "#F1F2F6", "크롬 대기"
        self.lbl_chrome.setText(f"●  {text}")
        self.lbl_chrome.setStyleSheet(f"""
            color:{fg}; background:{bg}; font-size:11px; font-weight:700;
            border-radius:14px; padding:0 14px;
        """)

    def _set_result_buttons(self, which, path):
        """결과 파일이 생기면 '파일 열기' / '폴더 열기' 활성화"""
        if which == "crawl":
            self._last_crawl_file = path
            btns = (self.btn_crawl_open, self.btn_crawl_dir)
        else:
            self._last_merge_file = path
            btns = (self.btn_merge_open, self.btn_merge_dir)
        for b in btns:
            b.setEnabled(bool(path))
        btns[0].setToolTip(str(path) if path else "작업이 끝나면 결과 파일을 열 수 있습니다.")

    def _open_path(self, path):
        if path and Path(path).exists():
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))
        elif path:
            self._log(f"파일을 찾을 수 없습니다: {path}", "err")

    def _set_stats(self, total, done, failed, eta):
        self.tile_total.value.setText(f"{total:,}")
        self.tile_done.value.setText(f"{done:,}")
        self.tile_fail.value.setText(f"{failed:,}")
        self.tile_eta.value.setText(eta)

    def _dk(self, h):
        r,g,b=int(h[1:3],16),int(h[3:5],16),int(h[5:7],16)
        return f"#{max(0,r-24):02x}{max(0,g-24):02x}{max(0,b-24):02x}"

    # ── 이벤트 ───────────────────────────────────────────
    def _pick_file(self, edit):
        p,_ = QFileDialog.getOpenFileName(self,"파일 선택",str(default_output_dir()),"Excel (*.xlsx *.xls)")
        if p: edit.setText(p)

    def _pick_dir(self, edit):
        p = QFileDialog.getExistingDirectory(self,"폴더 선택",edit.text() or str(Path.home()))
        if p: edit.setText(p)

    def _run_in_thread(self, worker):
        thread = QThread()
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.done.connect(thread.quit)
        thread.finished.connect(lambda: self._forget_thread(thread))
        self._threads.append((thread, worker))
        thread.start()
        return thread

    def _forget_thread(self, thread):
        self._threads = [(t, w) for t, w in self._threads if t is not thread]
        thread.deleteLater()

    def _driver_alive(self):
        if not self.driver:
            return False
        try:
            self.driver.title  # 창이 닫혔으면 예외
            return True
        except Exception:
            return False

    def _open_site(self):
        if self._driver_alive():
            # 이미 열려 있으면 새 크롬을 띄우지 않고 검색 페이지로 이동
            self.driver.get(BASE_URL)
            self._log("이미 열린 크롬에서 검색 페이지로 이동했습니다.", "dim")
            return
        self.driver = None
        self.btn_open.setEnabled(False)
        self._log("크롬 여는 중... (처음 실행 시 드라이버 다운로드로 조금 걸릴 수 있어요)", "dim")
        opener = ChromeOpener()
        opener.done.connect(self._chrome_opened)
        self._run_in_thread(opener)

    def _chrome_opened(self, driver, err):
        self.btn_open.setEnabled(True)
        if driver is None:
            self._log(f"❌  크롬을 열지 못했습니다: {err}", "err")
            QMessageBox.warning(self, "크롬 실행 실패",
                "크롬을 열지 못했습니다.\n\n"
                "· Google Chrome이 설치되어 있는지 확인하세요.\n"
                "· 인터넷 연결을 확인하세요 (처음 실행 시 드라이버를 내려받습니다).\n\n"
                f"상세: {err}")
            return
        self.driver = driver
        self._set_chrome_state(True)
        self._log("✓  사이트 열림. 검색 후 크롤링 시작 버튼을 누르세요.", "ok")
        self.btn_crawl.setEnabled(True)

    def _start_crawl(self):
        if not self._driver_alive():
            self.btn_crawl.setEnabled(False)
            self._set_chrome_state(False)
            self._log("크롬 창이 닫혀 있습니다. '크롬 열기'를 다시 누르고 검색하세요.", "err"); return
        self.btn_crawl.setEnabled(False)
        self.btn_crawl.setText("수집 중...")
        self.btn_stop.setEnabled(True)
        self.chk_fast.setEnabled(False)
        self._set_progress(0)
        self._set_stats(0, 0, 0, "—")
        self.tile_total.value.setText("목록…")

        worker = CrawlWorker(self.driver, self.crawl_output.text(), self.crawl_name.text(),
                             fast=self.chk_fast.isChecked())
        worker.log.connect(self._log)
        worker.status.connect(self.lbl_status.setText)
        worker.progress.connect(self._set_progress)
        worker.stats.connect(self._set_stats)
        worker.done.connect(self._crawl_done)
        self.worker = worker
        self._run_in_thread(worker)

    def _stop_crawl(self):
        if self.worker:
            self.worker.abort()
            self._log("⏹  중단 요청 중... 지금까지 모은 데이터를 저장합니다.", "warn")
            self.btn_stop.setEnabled(False)

    def _crawl_done(self, ok, msg):
        self.worker = None
        self.btn_crawl.setEnabled(True)
        self.btn_crawl.setText("▶  크롤링 시작")
        self.btn_stop.setEnabled(False)
        self.chk_fast.setEnabled(True)
        saved = msg.splitlines()[-1]
        if saved.endswith(".xlsx"):
            self._set_result_buttons("crawl", saved)
        if not self._driver_alive():
            self._set_chrome_state(False)
        if ok:
            self._log(f"\n🎉  크롤링 완료! → {msg}", "ok")
            self.lbl_status.setText("크롤링 완료 ✓")
            self.merge_crawl.setText(msg)   # 탭 2에 결과 파일 자동 입력
            self._ask_open("크롤링 완료", msg)
        else:
            self._log(f"\n⏹  {msg}", "warn")
            self.lbl_status.setText(msg.splitlines()[0])
            QMessageBox.information(self,"중단됨", msg)

    def _ask_open(self, title, path):
        r = QMessageBox.question(self, title, f"{title}!\n{path}\n\n파일을 지금 열까요?")
        if r == QMessageBox.StandardButton.Yes:
            self._open_path(path)

    def _start_merge(self):
        cf = self.merge_crawl.text().strip()
        pf = self.merge_permit.text().strip()
        od = self.merge_output.text().strip()
        if not cf or not pf:
            self._log("두 파일을 모두 선택하세요.", "err"); return
        for f in (cf, pf):
            if not Path(f).is_file():
                self._log(f"파일을 찾을 수 없습니다: {f}", "err"); return
        self.btn_merge.setEnabled(False)
        self.btn_merge.setText("병합 중...")
        self._set_progress(0)

        worker = MergeWorker(cf, pf, od, self.merge_name.text())
        worker.log.connect(self._log)
        worker.status.connect(self.lbl_status.setText)
        worker.progress.connect(self._set_progress)
        worker.done.connect(self._merge_done)
        self._run_in_thread(worker)

    def _merge_done(self, ok, msg):
        self.btn_merge.setEnabled(True)
        self.btn_merge.setText("파일 병합 시작")
        if ok:
            self._set_result_buttons("merge", msg)
            self._log(f"\n🎉  병합 완료! → {msg}", "ok")
            self._ask_open("병합 완료", msg)
        else:
            self._log(f"❌  {msg}", "err")
            QMessageBox.warning(self, "병합 실패", msg)

    def closeEvent(self, event):
        if self.worker:
            r = QMessageBox.question(self, "종료",
                "크롤링이 진행 중입니다. 중단하고 종료할까요?\n(지금까지 모은 데이터는 저장됩니다)")
            if r != QMessageBox.StandardButton.Yes:
                event.ignore(); return
            self.worker.abort()
        for t, _ in list(self._threads):
            t.quit(); t.wait(15000)
        if self.driver:
            try: self.driver.quit()   # 남아 있는 크롬/드라이버 프로세스 정리
            except Exception: pass
        event.accept()

    def _log(self, msg, level=""):
        colors = {"ok":"#5EEAD4","err":"#FCA5A5","warn":"#FCD34D",
                  "section":"#A5B4FC","dim":"#64748B","":"#94A3B8"}
        c = colors.get(level,"#94A3B8")
        safe = html.escape(msg).replace("\n", "<br>")
        self.log_box.append(
            f'<span style="color:{c};font-family:Menlo,Consolas,monospace;">{safe}</span>')
        self.log_box.verticalScrollBar().setValue(
            self.log_box.verticalScrollBar().maximum())
        try:
            with open(LOG_FILE, "a", encoding="utf-8") as f:
                f.write(f"{datetime.now():%Y-%m-%d %H:%M:%S}  {msg}\n")
        except Exception:
            pass

    def _set_progress(self, v):
        self.progress.setValue(v)
        self.lbl_pct.setText(f"{v}%")


# ═══════════════════════════════════════════════════════════
LOG_FILE = Path.home() / "emedi_crawler.log"


def _excepthook(etype, value, tb):
    """앱으로 실행할 땐 콘솔이 없으므로, 예상 못한 오류를 창과 로그 파일로 보여줌"""
    text = "".join(traceback.format_exception(etype, value, tb))
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"\n{datetime.now()}\n{text}\n")
    except Exception:
        pass
    sys.__excepthook__(etype, value, tb)
    if QApplication.instance():
        QMessageBox.critical(None, "오류", f"예상치 못한 오류가 발생했습니다.\n로그: {LOG_FILE}\n\n{value}")


if __name__ == "__main__":
    sys.excepthook = _excepthook
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setFont(QFont("Apple SD Gothic Neo" if platform.system()=="Darwin" else "맑은 고딕", 11))
    app.setApplicationName(APP_NAME)
    app.setWindowIcon(QIcon(AppLogo(128).grab()))
    win = MainWindow()
    win.show()
    sys.exit(app.exec())
