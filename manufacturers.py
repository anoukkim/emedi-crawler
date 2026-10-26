"""Split licenses with several manufacturers into one 상호/주소 column pair each.

The search list's 제조자 column separates manufacturers with '^'. The detail
popup joins both 상호 and 주소 with ',' instead, and both can contain commas
of their own (e.g. 'FMD Co.,Ltd.'), so names come from the list column. In the
popup address, the join between manufacturers is ' , ' (space before the
comma), while commas inside an address have no space before them.
"""

import re


NAME_COL = "제조자"
ADDR_COL = "수입업자_제조자 주소"
POPUP_NAME_COL = "수입업자_제조자 상호"
COUNT_COL = "제조자수"
CHECK_COL = "제조자_확인필요"

_DERIVED = re.compile(r"^제조자\d+_(상호|주소)$")
_ADDR_SEP = re.compile(r"\s+,\s*")


def _clean(text):
    return re.sub(r"\s+", " ", str(text or "")).strip()


def split_manufacturers(record):
    """Return ([(상호, 주소), ...], problem). problem is '' when the split is reliable."""
    names = [_clean(n) for n in str(record.get(NAME_COL) or "").split("^")]
    names = [n for n in names if n]
    if not names and _clean(record.get(POPUP_NAME_COL)):
        names = [_clean(record.get(POPUP_NAME_COL))]

    address = _clean(record.get(ADDR_COL))
    if not address:
        return [(n, "") for n in names], ""
    if len(names) <= 1:
        return [(names[0] if names else "", address)], ""

    addresses = [a.strip() for a in _ADDR_SEP.split(address)]
    if len(addresses) != len(names):
        return [(n, "") for n in names], f"제조자 {len(names)}곳 / 주소 {len(addresses)}개"
    return list(zip(names, addresses)), ""


def add_manufacturer_columns(records):
    """Copy records with 제조자수, 제조자N_상호, 제조자N_주소 placed right after 제조자.

    Every record gets the same number of column pairs so the Excel columns stay
    together. Columns left over from an earlier run (recovered from Excel) are
    replaced.
    """
    splits = [split_manufacturers(r) for r in records]
    width = max((len(pairs) for pairs, _ in splits), default=0)

    out = []
    for record, (pairs, problem) in zip(records, splits):
        derived = {COUNT_COL: len(pairs) or ""}
        for i in range(width):
            name, addr = pairs[i] if i < len(pairs) else ("", "")
            derived[f"제조자{i+1}_상호"] = name
            derived[f"제조자{i+1}_주소"] = addr
        derived[CHECK_COL] = problem

        base = {k: v for k, v in record.items()
                if k not in (COUNT_COL, CHECK_COL) and not _DERIVED.match(k)}
        if NAME_COL not in base:
            out.append({**base, **derived})
            continue
        row = {}
        for k, v in base.items():
            row[k] = v
            if k == NAME_COL:
                row.update(derived)
        out.append(row)
    return out
