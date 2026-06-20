# -*- coding: utf-8 -*-
"""PageLayout: A4 기준 페이지 분할 + 배율.

핵심 원칙(계획 확정):
  - PageSetup.Zoom 에 '정수%'를 주면 FitToPages 가 비활성화되고
    수동 페이지 나누기(V/H PageBreaks)가 보존된다. → 전 과정 Zoom 정수 모드.
  - 가로: 열 경계(세로줄) 기준으로만 페이지 분할(열이 중간에 잘리지 않음).
  - 세로: 병합셀이 페이지 경계에 걸치면 병합블록 전체를 다음 페이지로 밀어냄.
  - 폭 초과 시 배율 축소, 그래도 안 되면 가로방향 전환.
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


def _col_widths(ws, first_col, n_cols):
    out = []
    for i in range(n_cols):
        try:
            out.append(float(ws.Columns(first_col + i).Width))
        except Exception:
            out.append(0.0)
    return out


def _row_heights(ws, first_row, n_rows):
    out = []
    for i in range(n_rows):
        try:
            out.append(float(ws.Rows(first_row + i).Height))
        except Exception:
            out.append(0.0)
    return out


def _merge_start_crossing(ws, boundary_row, first_col, n_cols):
    """boundary_row 가 어떤 병합블록 내부를 가르면, 그 병합블록들의
    최상단 시작행을 반환. 가르는 병합이 없으면 None."""
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
    """area(ws.Range)에 A4 레이아웃 적용. Returns 진단 dict."""
    ps = ws.PageSetup
    first_col, n_cols = area.Column, area.Columns.Count
    first_row, n_rows = area.Row, area.Rows.Count

    # 인쇄영역 고정
    try:
        ps.PrintArea = area.Address
    except Exception:
        pass
    ps.PaperSize = XL_PAPER_A4

    col_w = _col_widths(ws, first_col, n_cols)
    widest_col = max(col_w) if col_w else 0.0

    # --- 방향/배율 결정 ---
    orientation = "portrait"
    pw, ph = _printable(orientation, cfg)
    zoom = 100
    if widest_col > pw:  # 한 열이 세로 폭 초과 → 가로 전환 시도
        orientation = "landscape"
        pw, ph = _printable(orientation, cfg)
    if widest_col > pw:  # 가로로도 초과 → 배율 축소
        zoom = max(cfg["min_zoom"], int(pw / widest_col * 100))

    ps.Orientation = XL_LANDSCAPE if orientation == "landscape" else XL_PORTRAIT
    ps.Zoom = int(zoom)  # 정수 → FitToPages 비활성, 수동 break 보존

    scale = zoom / 100.0
    eff_w = pw / scale
    eff_h = ph / scale

    # 기존 나누기 초기화
    try:
        ws.ResetAllPageBreaks()
    except Exception:
        pass

    # --- 가로(열 경계) 분할: 누적 폭이 eff_w 초과 직전에서 끊기 ---
    v_breaks = []
    acc = 0.0
    for i in range(n_cols):
        w = col_w[i]
        if acc > 0 and acc + w > eff_w:
            brk_col = first_col + i
            v_breaks.append(brk_col)
            acc = w
        else:
            acc += w
    for c in v_breaks:
        try:
            ws.VPageBreaks.Add(Before=ws.Columns(c))
        except Exception:
            pass

    # --- 세로(행) 분할: 누적 높이 초과 시 끊되, 병합블록 가르면 위로 ---
    row_h = _row_heights(ws, first_row, n_rows)
    h_breaks = []
    acc = 0.0
    i = 0
    while i < n_rows:
        h = row_h[i]
        cur_row = first_row + i
        if acc > 0 and acc + h > eff_h:
            # cur_row 앞에서 끊으려 함 → 병합블록이 가르는지 확인
            top = _merge_start_crossing(ws, cur_row, first_col, n_cols)
            brk_row = top if top is not None and top > first_row else cur_row
            h_breaks.append(brk_row)
            # 끊은 지점부터 누적 재시작
            i = brk_row - first_row
            acc = 0.0
            continue
        acc += h
        i += 1
    for r in sorted(set(h_breaks)):
        try:
            ws.HPageBreaks.Add(Before=ws.Rows(r))
        except Exception:
            pass

    return {
        "orientation": orientation,
        "zoom": int(zoom),
        "v_breaks": len(v_breaks),
        "h_breaks": len(set(h_breaks)),
    }
