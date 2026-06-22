# -*- coding: utf-8 -*-
"""AutoFix: '###' 렌더링 제거를 위한 열 너비 보정.

전략(계획 + 실측 보정):
  - '###' 는 숫자/날짜 셀이 열 폭보다 길 때만 발생한다.
  - 일반 AutoFit 은 같은 열의 '긴 텍스트 헤더'(줄바꿈된 한글 제목 등)에 맞춰
    열을 과도하게 넓혀 표 전체 폭을 부풀린다(→ 배율이 과도하게 작아짐).
  - 그래서 헤더가 아닌 '숫자의 표시 서식 길이'에 맞춰서만 'grow-only'로 넓힌다.
    (열당 NumberFormat 1회 샘플 + grid 값으로 길이 계산 → COM 왕복 최소화)
  - 폭 상한(문자수)으로 클램프.
원본 값은 절대 바꾸지 않는다(사본에서 열 너비만 조정).
"""
from __future__ import annotations

import datetime


def _col_letter(idx: int) -> str:
    s = ""
    while idx > 0:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


def _as_2d(val):
    if val is None:
        return [[None]]
    if not isinstance(val, tuple):
        return [[val]]
    if len(val) == 0 or not isinstance(val[0], tuple):
        return [list(val)]
    return [list(r) for r in val]


def _is_numeric(v) -> bool:
    if isinstance(v, bool):
        return False
    if isinstance(v, (int, float)):
        return True
    if isinstance(v, datetime.datetime):
        return True
    mod = getattr(type(v), "__module__", "") or ""
    return "pywintypes" in mod


def _parse_format(fmt: str):
    """(decimals, thousands, is_general)."""
    if not fmt or fmt.lower() == "general":
        return 0, False, True
    thousands = "," in fmt.split(".")[0] if "." in fmt else ("," in fmt)
    decimals = 0
    if "." in fmt:
        after = fmt.split(".", 1)[1]
        for ch in after:
            if ch in "0#":
                decimals += 1
            else:
                break
    return decimals, thousands, False


def _disp_len(v, decimals, thousands, is_general) -> int:
    """숫자 v 가 표시될 때의 대략 문자 길이."""
    try:
        if isinstance(v, datetime.datetime):
            return 10  # YYYY-MM-DD 가정
        if is_general:
            s = "{:g}".format(v)
        elif thousands:
            s = "{:,.{d}f}".format(v, d=decimals)
        else:
            s = "{:.{d}f}".format(v, d=decimals)
        return len(s)
    except Exception:
        return len(str(v))


def fix_sheet(ws, area, max_col_width_chars: float) -> dict:
    result = {"widened_cols": [], "clamped_cols": []}
    if area is None:
        return result

    first_col = area.Column
    first_row = area.Row

    try:
        grid = _as_2d(area.Value)
    except Exception:
        return result

    # 열별 숫자 값 + 첫 숫자 위치 수집
    col_numeric = {}      # ci -> [values]
    col_first_row = {}    # ci -> grid row index of first numeric cell
    for ri, row in enumerate(grid):
        for ci, v in enumerate(row):
            if _is_numeric(v):
                col_numeric.setdefault(ci, []).append(v)
                if ci not in col_first_row:
                    col_first_row[ci] = ri

    for ci in sorted(col_numeric):
        c = first_col + ci
        col = ws.Columns(c)
        try:
            before = float(col.ColumnWidth)
        except Exception:
            continue

        # 해당 열의 숫자 서식 1회 샘플
        try:
            sample = ws.Cells(first_row + col_first_row[ci], c)
            fmt = sample.NumberFormat
        except Exception:
            fmt = "General"
        decimals, thousands, is_general = _parse_format(fmt)

        max_len = 0
        for v in col_numeric[ci]:
            ln = _disp_len(v, decimals, thousands, is_general)
            if ln > max_len:
                max_len = ln
        if max_len == 0:
            continue

        needed = max_len + 1.5  # 좌우 패딩(문자수)
        if needed <= before + 0.01:
            continue  # 이미 충분 → 건드리지 않음

        target = min(needed, max_col_width_chars)
        if target <= before + 0.01:
            continue
        try:
            col.ColumnWidth = target
        except Exception:
            continue
        result["widened_cols"].append(_col_letter(c))
        if needed > max_col_width_chars:
            result["clamped_cols"].append(_col_letter(c))

    return result
