# -*- coding: utf-8 -*-
"""prep: 워크북의 각 시트에 '페이지 나누어 미리보기'(PBP)를 활성화한 뒤 '원본에 저장'한다.

인쇄영역 정책: 원래 지정된 인쇄영역은 그대로 두고, 지정되지 않은 시트는
Excel 기본값(설정 안 함=자동 결정)으로 둔다. (set_pa_if_missing=True 면 UsedRange 로 설정)

핵심 제약: 셀 크기(너비/높이)·배율 등은 절대 변경하지 않는다.
숨김/제외 패턴 시트는 건너뛴다.
"""
from __future__ import annotations

from config import XL_SHEET_VISIBLE, XL_VIEW_PAGEBREAK_PREVIEW


def _excluded(ws, patterns) -> str | None:
    try:
        if int(ws.Visible) != XL_SHEET_VISIBLE:
            return "hidden"
    except Exception:
        pass
    name = (ws.Name or "").strip()
    for p in patterns:
        if p.search(name):
            return f"name:{p.pattern}"
    return None


def _has_content(ws) -> bool:
    try:
        ur = ws.UsedRange
        if ur.Rows.Count == 1 and ur.Columns.Count == 1:
            return ur.Value not in (None, "")
        return True
    except Exception:
        return True


def prep_workbook(app, wb, patterns, set_pa_if_missing=True) -> dict:
    """Returns {included:[names], excluded:[(name,reason)]}."""
    result = {"included": [], "excluded": []}
    changed = False
    for ws in wb.Worksheets:
        reason = _excluded(ws, patterns)
        if reason:
            result["excluded"].append((ws.Name, reason))
            continue
        if not _has_content(ws):
            result["excluded"].append((ws.Name, "empty"))
            continue
        try:
            ws.Activate()
            # 페이지 나누어 미리보기 활성화 (뷰는 시트별로 파일에 저장됨)
            try:
                app.ActiveWindow.View = XL_VIEW_PAGEBREAK_PREVIEW
                changed = True
            except Exception:
                pass
            # 인쇄영역 없으면 UsedRange 로 설정 (셀은 무편집)
            if set_pa_if_missing:
                pa = (ws.PageSetup.PrintArea or "").strip()
                if not pa:
                    ws.PageSetup.PrintArea = ws.UsedRange.Address
                    changed = True
            result["included"].append(ws.Name)
        except Exception as e:
            result["excluded"].append((ws.Name, f"error:{e}"))

    if changed:
        try:
            wb.Save()  # 원본에 저장 (PBP/인쇄영역 반영)
        except Exception as e:
            print(f"    [경고] 원본 저장 실패: {e}")
    return result
