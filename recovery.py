"""Resume an eMedi crawl without requesting details already saved locally."""

import json
import re
from pathlib import Path


DETAIL_PREFIXES = ("제품정보_", "UDI_", "수입업자_", "기타정보_", "보험_")
PERMIT_SHEET = "허가증별"


def record_key(record):
    """Use MFDS's stable item serial, falling back to its permit number."""
    serial = str(record.get("품목일련번호") or "").strip()
    if serial:
        return ("serial", serial)
    permit = re.sub(r"\s+", "", str(record.get("품목허가번호") or ""))
    if permit:
        return ("permit", permit)
    return None


def detail_complete(record):
    """A list row or an empty/partially loaded modal is not a completed crawl."""
    if not isinstance(record, dict) or record.get("오류"):
        return False
    fields = [k for k, v in record.items()
              if k.startswith(DETAIL_PREFIXES) and str(v or "").strip()]
    return len(fields) >= 8 and sum(k.startswith("제품정보_") for k in fields) >= 2


def load_existing_records(path):
    source = Path(path)
    if not source.is_file():
        raise FileNotFoundError(f"기존 수집 파일을 찾을 수 없습니다: {source}")
    if source.suffix.lower() == ".json":
        with source.open(encoding="utf-8-sig") as handle:
            records = json.load(handle)
    elif source.suffix.lower() == ".xlsx":
        import pandas as pd
        # 모델명별로 늘린 시트는 같은 허가가 여러 행이므로, 허가증 1건당 1행인 시트를 읽는다
        with pd.ExcelFile(source, engine="openpyxl") as book:
            sheets = book.sheet_names
        sheet = PERMIT_SHEET if PERMIT_SHEET in sheets else 0
        records = pd.read_excel(source, sheet_name=sheet, engine="openpyxl", dtype=str).fillna("").to_dict("records")
    else:
        raise ValueError("기존 결과는 .json 또는 .xlsx 파일을 선택하세요.")
    if not isinstance(records, list) or not records or not all(isinstance(r, dict) and record_key(r) for r in records):
        raise ValueError("수집 결과 형식이 아닙니다. 허가번호/품목일련번호가 있는 JSON 또는 Excel을 선택하세요.")
    return records


def build_recovery_plan(items, existing, list_columns):
    """Keep saved details, add fresh list rows, and queue only incomplete records."""
    saved = {}
    for record in existing:
        key = record_key(record)
        if key in saved:
            raise ValueError(f"기존 파일에 동일 품목이 중복되어 있습니다: {key}")
        saved[key] = record
    if len({record_key(item) for item in items}) != len(items):
        raise ValueError("사이트 검색 목록에 품목일련번호/허가번호 중복이 있어 자동 복구할 수 없습니다.")
    fresh_keys = {record_key(item) for item in items}
    overlap = len(fresh_keys & saved.keys())
    if existing and overlap != len(saved):
        raise ValueError(f"기존 파일 {len(saved)}건 중 {len(saved)-overlap}건이 현재 검색 목록에 없습니다. "
                         "사이트 목록 로딩 및 검색 조건을 확인한 뒤 다시 시도하세요.")

    records, queue = [], []
    for i, item in enumerate(items):
        old = saved.get(record_key(item), {})
        record = dict(old)
        record.update({k: item.get(k, "") for k in list_columns})
        if not detail_complete(record):
            record["오류"] = record.get("오류") or "상세 정보 미수집"
            queue.append(i)
        records.append(record)

    return records, queue, overlap, 0
