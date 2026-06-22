# -*- coding: utf-8 -*-
"""PageLayout: A4 기준 페이지 분할 + 배율.

핵심 원칙(계획 확정 + 실측 보정):
  - PageSetup.Zoom 에 '정수%'를 주면 FitToPages 가 비활성화되고
    수동 페이지 나누기(V/H PageBreaks)가 보존된다. → 전 과정 Zoom 정수 모드.
  - 폭(가로): 우선 표 전체 폭을 A4 1페이지에 맞춰 '배율 축소'한다(열이 잘리지 않음).
    축소해도 min_zoom 보다 더 줄여야 할 만큼 넓으면, 그때 비로소 '열 경계(세로줄)'
    기준으로 페이지를 분할한다.
  - 세로: 행을 여러 페이지로 흘리되, 병합셀이 페이지 경계에 걸치면 병합블록
    전체를 다음 페이지로 밀어낸다.
  - 내용 없는 후행(trailing) 빈 행/열은 페이지 계산에서 제외(빈 페이지 방지).
"""
from __future__ import annotations

from config import (
    A4_HEIGHT_PT, A4_WIDTH_PT, CM,
    XL_LANDSCAPE, XL_PAPER_A4, XL_PORTRAIT,
)


def _printable(orientation: str, cfg: dict):
    if orientation == "landscape":
        pw, ph = A4_HEIGHT_PT, A4_WIDTH_PT
    else:
        pw, ph = A4_WIDTH_PT, A4_HEIGHT_PT
    w = pw - (cfg["margin_left_cm"] + cfg["margin_right_cm"]) * CM
    h = ph - (cfg["margin_top_cm"] + cfg["margin_bottom_cm"]) * CM
    return w, h


def _as_2d(val):
    if val is None:
        return [[None]]
    if not isinstance(val, tuple):
        return [[val]]
    if len(val) == 0 or not isinstance(val[0], tuple):
        return [list(val)]
    return [list(r) for r in val]


def _trim_trailing_empty(ws, area):
    """후행 빈 행/열을 제거한 Range 반환(선두는 유지).

    주의: 도형/그림(Shapes)이 있는 시트는 셀 값이 비어 보여도 그림이 차지하므로
    값 기준 trim 이 본문(그림)을 잘라낼 수 있다. 도형이 하나라도 있으면 trim 생략.
    """
    try:
        if ws.Shapes.Count > 0:
            return area
    except Exception:
        return area
    try:
        grid = _as_2d(area.Value)
    except Exception:
        return area
    n_rows = len(grid)
    n_cols = max((len(r) for r in grid), default=0)
    if n_rows == 0 or n_cols == 0:
        return area

    def nonempty(v):
        return v not in (None, "")

    last_row = 0
    last_col = 0
    for ri in range(n_rows):
        for ci in range(len(grid[ri])):
            if nonempty(grid[ri][ci]):
                if ri + 1 > last_row:
                    last_row = ri + 1
                if ci + 1 > last_col:
                    last_col = ci + 1
    if last_row == 0 or last_col == 0:
        return area
    if last_row == n_rows and last_col == n_cols:
        return area
    fr, fc = area.Row, area.Column
    try:
        return ws.Range(ws.Cells(fr, fc),
                        ws.Cells(fr + last_row - 1, fc + last_col - 1))
    except Exception:
        return area


def _col_widths(ws, first_col, n_cols):
    return [_safe(lambda: float(ws.Columns(first_col + i).Width), 0.0)
            for i in range(n_cols)]


def _row_heights(ws, first_row, n_rows):
    return [_safe(lambda: float(ws.Rows(first_row + i).Height), 0.0)
            for i in range(n_rows)]


def _safe(fn, default):
    try:
        return fn()
    except Exception:
        return default


def _merge_start_crossing(ws, boundary_row, first_col, n_cols):
    """boundary_row 가 병합블록 내부를 가르면 그 블록들의 최상단 시작행 반환."""
    top = None
    for i in range(n_cols):
        c = first_col + i
        try:
            cell = ws.Cells(boundary_row, c)
            if not cell.MergeCells:
                continue
            ma = cell.MergeArea
            ma_row = ma.Row
            ma_end = ma.Row + ma.Rows.Count - 1
            if ma_row < boundary_row <= ma_end:
                if top is None or ma_row < top:
                    top = ma_row
        except Exception:
            continue
    return top


def apply(ws, area, cfg) -> dict:
    ps = ws.PageSetup

    area = _trim_trailing_empty(ws, area)
    first_col, n_cols = area.Column, area.Columns.Count
    first_row, n_rows = area.Row, area.Rows.Count

    try:
        ps.PrintArea = area.Address
    except Exception:
        pass
    ps.PaperSize = XL_PAPER_A4

    col_w = _col_widths(ws, first_col, n_cols)
    total_w = sum(col_w)
    min_zoom = int(cfg["min_zoom"])

    # --- 방향/배율 초기 추정: 폭을 1페이지에 맞추는 배율(안전계수 0.97) ---
    def fit_zoom(orient):
        pw, _ph = _printable(orient, cfg)
        if total_w <= 0:
            return 100
        return min(100, int(pw / total_w * 100 * 0.97))

    zp, zl = fit_zoom("portrait"), fit_zoom("landscape")
    orientation = "landscape" if zl > zp else "portrait"
    zoom = max(min_zoom, min(100, max(zp, zl)))

    ps.Orientation = XL_LANDSCAPE if orientation == "landscape" else XL_PORTRAIT

    # 주의: ResetAllPageBreaks() 는 Zoom 을 100 으로 되돌린다(실측 확인).
    # 따라서 반드시 Zoom 설정 '전에' 호출한다.
    try:
        ws.ResetAllPageBreaks()
    except Exception:
        pass
    # FitToPages 를 먼저 끄지 않으면 Zoom 정수 설정이 무시된다(실측 확인).
    try:
        ps.FitToPagesWide = False
        ps.FitToPagesTall = False
    except Exception:
        pass
    ps.Zoom = int(zoom)

    # --- 가로: Excel 실제 분할(VPageBreaks)을 피드백 삼아 배율을 낮춰 폭 맞춤 ---
    def vcount():
        try:
            return ws.VPageBreaks.Count
        except Exception:
            return 0

    for _ in range(20):
        if vcount() == 0 or zoom <= min_zoom:
            break
        zoom = max(min_zoom, zoom - 2)
        ps.Zoom = int(zoom)
    # min_zoom 에서도 폭이 넘치면 Excel 자동 V분할(열 경계)이 그대로 사용됨.

    # --- 세로: Excel 실제 H분할을 읽어, 병합블록을 가르면 위로 당겨 보호 ---
    def hbreak_rows():
        rows = []
        try:
            for i in range(ws.HPageBreaks.Count):
                rows.append(int(ws.HPageBreaks(i + 1).Location.Row))
        except Exception:
            pass
        return rows

    added_h = set()
    for _ in range(300):
        moved = False
        for r in hbreak_rows():
            top = _merge_start_crossing(ws, r, first_col, n_cols)
            if top is not None and first_row < top < r and top not in added_h:
                try:
                    ws.HPageBreaks.Add(Before=ws.Rows(top))
                    added_h.add(top)
                    moved = True
                    break
                except Exception:
                    pass
        if not moved:
            break

    return {
        "orientation": orientation,
        "zoom": int(zoom),
        "area": area.Address,
        "v_breaks": vcount(),
        "h_breaks": len(hbreak_rows()),
        "merge_protected": len(added_h),
    }
