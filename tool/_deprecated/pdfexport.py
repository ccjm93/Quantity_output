# -*- coding: utf-8 -*-
"""PdfExport: 포함 시트만 남기고 워크북당 1개 PDF로 내보내기.

신뢰성을 위해 '제외 시트는 임시 숨김 → 워크북 전체 export' 방식을 쓴다.
(다중 Select 보다 상태 변화가 단순하고 탭 순서가 보존됨)
"""
from __future__ import annotations

import os

from config import XL_SHEET_HIDDEN, XL_SHEET_VISIBLE, XL_TYPE_PDF


def export_workbook(wb, included_names: list[str], out_pdf: str) -> bool:
    """included_names 시트만 보이게 한 뒤 PDF 출력. 성공 여부 반환."""
    if not included_names:
        return False

    inc = set(included_names)
    # 가시성 설정: 포함=보이게, 그 외=숨김. 최소 1개는 보여야 함.
    visible_count = 0
    for ws in wb.Worksheets:
        try:
            if ws.Name in inc:
                ws.Visible = XL_SHEET_VISIBLE
                visible_count += 1
            else:
                ws.Visible = XL_SHEET_HIDDEN
        except Exception:
            pass
    if visible_count == 0:
        return False

    os.makedirs(os.path.dirname(out_pdf), exist_ok=True)
    try:
        wb.ExportAsFixedFormat(
            Type=XL_TYPE_PDF,
            Filename=out_pdf,
            Quality=0,                  # xlQualityStandard
            IncludeDocProperties=True,
            IgnorePrintAreas=False,
            OpenAfterPublish=False,
        )
    except Exception as e:
        print(f"    [오류] PDF export 실패: {e}")
        return False
    return os.path.isfile(out_pdf)
