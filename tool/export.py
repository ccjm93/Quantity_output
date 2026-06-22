# -*- coding: utf-8 -*-
"""export: 포함 시트를 '시트별 임시 PDF'로 내보내고 페이지수를 수집한다.

시트 단위로 내보내야 최종 병합 PDF에서 페이지↔시트 매핑(AI 위치 보고용)을
정확히 만들 수 있다.
"""
from __future__ import annotations

import os

import fitz  # pymupdf

from config import XL_TYPE_PDF


def export_sheets(wb, included_names, tmp_dir) -> list:
    """Returns [{sheet, pdf, pages}] (export 실패 시 해당 시트 제외)."""
    os.makedirs(tmp_dir, exist_ok=True)
    out = []
    for i, name in enumerate(included_names):
        try:
            ws = wb.Worksheets(name)
        except Exception:
            continue
        pdf = os.path.join(tmp_dir, f"s{i:03d}.pdf")
        try:
            ws.ExportAsFixedFormat(
                Type=XL_TYPE_PDF, Filename=os.path.abspath(pdf),
                IgnorePrintAreas=False, OpenAfterPublish=False)
        except Exception as e:
            print(f"    [경고] 시트 export 실패({name}): {e}")
            continue
        if not os.path.isfile(pdf):
            continue
        try:
            d = fitz.open(pdf)
            pages = d.page_count
            d.close()
        except Exception:
            pages = 0
        if pages <= 0:
            continue
        out.append({"sheet": name, "pdf": pdf, "pages": pages})
    return out
