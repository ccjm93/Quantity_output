# -*- coding: utf-8 -*-
"""AutoFix: '###' 렌더링 제거를 위한 열 너비/행 높이 보정.

전략(계획 확정): 열 너비 확대 우선. AutoFit 으로 내용에 맞춰 넓히되,
폭 폭주를 막기 위해 상한(문자수)으로 클램프한다. 폭 초과분은 PageLayout 의
배율 축소가 처리한다. 행 높이는 필요 시 보정.

원본 값은 절대 바꾸지 않는다(사본에서만, 열 너비/행 높이만 조정).
"""
from __future__ import annotations


def _col_letter(idx: int) -> str:
    s = ""
    while idx > 0:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


def fix_sheet(ws, area, max_col_width_chars: float) -> dict:
    """area(=ws.Range 객체)의 열들을 AutoFit + 클램프.

    Returns: {"widened_cols": [..], "clamped_cols": [..]}
    """
    result = {"widened_cols": [], "clamped_cols": []}
    if area is None:
        return result

    first_col = area.Column
    n_cols = area.Columns.Count
    first_row = area.Row
    n_rows = area.Rows.Count

    for i in range(n_cols):
        c = first_col + i
        col = ws.Columns(c)
        try:
            before = float(col.ColumnWidth)
        except Exception:
            continue
        try:
            col.AutoFit()
        except Exception:
            continue
        try:
            after = float(col.ColumnWidth)
        except Exception:
            after = before

        if after > before + 0.01:
            result["widened_cols"].append(_col_letter(c))
        # 상한 클램프
        if after > max_col_width_chars:
            try:
                col.ColumnWidth = max_col_width_chars
                result["clamped_cols"].append(_col_letter(c))
            except Exception:
                pass

    # 행 높이: AutoFit 으로 줄바꿈/넓어진 내용 수용 (숫자 ###엔 영향 적지만 안전)
    try:
        ws.Rows(f"{first_row}:{first_row + n_rows - 1}").AutoFit()
    except Exception:
        pass

    return result
