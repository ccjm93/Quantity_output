# -*- coding: utf-8 -*-
"""detect: 열려 있는 워크시트에서 '셀 표시 오류'를 결정론적으로 탐지한다.

PDF 텍스트 레이어에는 '###'(열 폭 부족 표시)가 기록되지 않으므로 PDF 사후 검토로는
이 문제를 잡을 수 없다. 대신 워크북이 COM 으로 열려 있을 때 Excel 자신이 계산한
표시 텍스트(Range.Text)를 직접 읽어 탐지한다.

- overflow: 셀이 열 폭 부족 등으로 '##'(또는 '#######')로 표시됨 → Range.Text 가 '#' 2개 이상으로만 채워짐.
- ref_error: 수식 오류 셀 중 #REF!.
- error:    그 외 수식 오류 셀(#DIV/0!, #VALUE!, #NAME?, #N/A, #NUM!, #NULL!).

SpecialCells 로 '숫자 셀'/'오류 셀'만 골라 검사하므로 대형 시트도 비용이 제한적이다.
원본 셀 크기·서식은 절대 변경하지 않는다(읽기만 한다).
"""
from __future__ import annotations

import re

from config import (
    XL_CELLTYPE_CONSTANTS,
    XL_CELLTYPE_FORMULAS,
    XL_ERRORS,
)

# '#' 2개 이상으로만(공백 허용) 이루어진 표시 = 열 폭 부족 등 표시 오류
_HASH_ONLY = re.compile(r"^#{2,}\s*$")


def _iter_cells(rng):
    """SpecialCells 결과(여러 Area 가능)를 셀 단위로 순회. 실패는 무시."""
    try:
        areas = rng.Areas
    except Exception:
        return
    for area_index in range(1, areas.Count + 1):
        try:
            area = areas.Item(area_index)
            cells = area.Cells
            n = cells.Count
        except Exception:
            continue
        for ci in range(1, n + 1):
            try:
                yield cells.Item(ci)
            except Exception:
                continue


def _special(ws, ctype, value):
    """ws.UsedRange.SpecialCells(...) — 매칭 셀이 없으면 com_error → None."""
    try:
        if value is None:
            return ws.UsedRange.SpecialCells(ctype)
        return ws.UsedRange.SpecialCells(ctype, value)
    except Exception:
        return None


# 시트당 셀 스캔 예산(병적으로 큰 시트에서 무한정 느려지지 않도록 상한)
_MAX_CELLS_PER_SHEET = 30000


def scan_sheet_issues(ws) -> list:
    """Returns [{sheet, cell, kind, text}]  (kind: 'overflow' | 'ref_error' | 'error').

    주의: 이 탐지는 셀의 '표시 텍스트(Range.Text)'가 ###/오류일 때만 잡는다.
    그림/단면도 안의 숫자가 '출력 배율' 때문에 ###로 찌그러지는 경우는 셀 값 자체가
    정상이라 여기서 잡히지 않는다(그건 출력 설정 문제이며 PDF 출력물 검토 영역).
    """
    name = (ws.Name or "").strip()
    findings = []
    seen = set()  # (addr, kind) 중복 방지

    def _add(cell, kind, text):
        # late-bound COM 에서 Address 는 메서드가 아니라 문자열 속성으로 평가된다.
        # ("$B$12") → '$' 제거로 상대주소("B12") 화. 실패해도 finding 을 버리지 않는다.
        try:
            addr = str(cell.Address).replace("$", "")
        except Exception:
            addr = "?"
        key = (addr, kind)
        if key in seen:
            return
        seen.add(key)
        findings.append({"sheet": name, "cell": addr, "kind": kind, "text": text})

    budget = _MAX_CELLS_PER_SHEET

    # 1) 연속 '#': 상수/수식 셀의 표시 텍스트를 직접 검사
    for ctype in (XL_CELLTYPE_CONSTANTS, XL_CELLTYPE_FORMULAS):
        rng = _special(ws, ctype, None)
        if rng is None:
            continue
        for cell in _iter_cells(rng):
            budget -= 1
            if budget < 0:
                return findings
            try:
                t = (cell.Text or "").strip()
            except Exception:
                continue
            if _HASH_ONLY.match(t):
                _add(cell, "overflow", t)

    # 2) 수식 오류 셀 — 실제로 '#...'를 '표시'하는 셀만(병합 빈칸 중복 제거)
    rng = _special(ws, XL_CELLTYPE_FORMULAS, XL_ERRORS)
    if rng is not None:
        for cell in _iter_cells(rng):
            budget -= 1
            if budget < 0:
                break
            try:
                t = (cell.Text or "").strip()
            except Exception:
                continue
            if t.startswith("#"):
                _add(cell, "ref_error" if t.upper().startswith("#REF!") else "error", t)

    return findings
