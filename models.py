"""Split one permit record into one row per model name (UPN)."""


MODEL_COL = "모델명(UPN)"
MODEL_SEQ_COL = "모델순번"
UDI_ROWS_KEY = "_UDI_rows"   # PARSE_JS가 표준코드(UDI) 표를 행 단위로 담는 키
UDI_SKIP = ("연번",)         # 사이트 표의 연번은 모델순번으로 대체


def _split(value, sep):
    return [p.strip() for p in str(value or "").split(sep)]


def udi_rows(record):
    """표준코드(UDI) 표의 행 목록. 예전 결과(JSON/엑셀)는 '; '로 이어진 UDI_ 컬럼에서 복원한다."""
    rows = record.get(UDI_ROWS_KEY)
    if isinstance(rows, list) and rows:
        return [{k: str(v or "").strip() for k, v in r.items()} for r in rows if isinstance(r, dict)]

    cols = {k[4:]: _split(v, ";") for k, v in record.items()
            if k.startswith("UDI_") and str(v or "").strip()}
    names = cols.get("모델명")
    if not names or not any(names):
        return []
    # 컬럼마다 값 개수가 다르면 어느 값이 어느 모델 것인지 알 수 없으므로 모델명만 쓴다
    if any(len(v) != len(names) for v in cols.values()):
        return [{"모델명": n} for n in names if n]
    return [{h: cols[h][i] for h in cols} for i in range(len(names)) if names[i]]


def model_names(record):
    """UDI 표가 없을 때 쓰는 모델명 목록 (제품정보/목록의 모델명, ; 구분)."""
    for key in ("제품정보_모델명", "모델명"):
        names = [n for n in _split(record.get(key), ";") if n]
        if names:
            return names
    return []


def expand_by_model(record):
    """허가 1건 → 모델명(UPN)별 행 목록. 허가증 전체 항목은 모든 행에 그대로 복사한다."""
    base = {k: v for k, v in record.items() if not k.startswith("_")}
    rows, from_udi = udi_rows(record), True
    if not rows:
        rows, from_udi = [{"모델명": n} for n in model_names(record)], False
    if not rows:
        return [{MODEL_SEQ_COL: "", MODEL_COL: "", **base}]

    out = []
    for i, udi in enumerate(rows, 1):
        row = {MODEL_SEQ_COL: f"{i}/{len(rows)}", MODEL_COL: udi.get("모델명", ""), **base}
        if from_udi:   # 모델별 UDI 값으로 '; '로 이어진 값을 덮어쓴다
            for k, v in udi.items():
                if k and k not in UDI_SKIP:
                    row[f"UDI_{k}"] = v
            for k in UDI_SKIP:
                row.pop(f"UDI_{k}", None)
        out.append(row)
    return out


def expand_records(records):
    rows = []
    for record in records:
        rows.extend(expand_by_model(record))
    return rows
