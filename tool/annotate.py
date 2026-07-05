# -*- coding: utf-8 -*-
"""annotate: 검토 결과(findings)를 PDF 사본에 빨간 표시로 그려 '검토용 PDF'를 만든다.

- 텍스트 근거가 있는 오류(#REF!/수식 오류/연속 #)는 페이지에서 해당 문자열을 찾아
  그 위치마다 빨간 박스를 그린다(같은 문자열이 여러 곳이면 모두 표시).
- 하단 표 테두리 누락 의심은 검사 대상이었던 하단 영역을 점선 박스로,
  빈 페이지 의심은 페이지 전체 테두리로 표시한다(휴리스틱이라 셀 좌표가 없음).
- 원본 PDF는 그대로 두고 '<이름>(검토표시).pdf' 사본에만 그린다(납품본 보호).
"""
from __future__ import annotations

import os
import re

import fitz  # pymupdf

MARKED_SUFFIX = "(검토표시)"

_RED = (0.86, 0.08, 0.08)
_QUOTED = re.compile(r"'([^']+)'")   # detail 문자열 속 '...' 에서 표시 원문 추출
_MAX_HITS_PER_FINDING = 50           # 같은 문자열이 비정상적으로 많을 때 상한


def marked_pdf_path(pdf_path: str) -> str:
    base, ext = os.path.splitext(pdf_path)
    return f"{base}{MARKED_SUFFIX}{ext or '.pdf'}"


def _location_label(finding) -> str | None:
    """박스 안에 표기할 '폴더/파일명/시트명' 라벨 (없는 항목은 생략)."""
    parts = []
    file_rel = (finding.get("file") or "").replace("\\", "/").strip("/")
    if file_rel:
        parts.append(file_rel)  # relpath 라 폴더명이 이미 포함됨
    elif finding.get("folder"):
        parts.append(str(finding["folder"]).replace("\\", "/").strip("/"))
    if finding.get("sheet"):
        parts.append(str(finding["sheet"]))
    return " / ".join(parts) if parts else None


def _insert_label(page, x, y, text, font_path):
    """빨간 라벨 텍스트 삽입 (한글 폰트 필요 — 실패해도 박스는 유지)."""
    try:
        if font_path and os.path.isfile(font_path):
            page.insert_text((x, y), text, fontsize=9, color=_RED,
                             fontname="krfont", fontfile=font_path)
        else:
            page.insert_text((x, y), text, fontsize=9, color=_RED)
    except Exception:
        pass


def _search_text(finding) -> str | None:
    """finding 에서 페이지 내 검색에 쓸 표시 텍스트를 얻는다."""
    text = (finding.get("text") or "").strip()
    if text:
        return text
    quoted = _QUOTED.findall(finding.get("detail") or "")
    if quoted:
        return quoted[-1]
    types = " ".join(finding.get("types") or [])
    if "#REF" in types.upper():
        return "#REF!"
    return None


def _mark(page, finding, seen: set, font_path=None) -> int:
    """finding 하나를 페이지에 표시. 그린 박스 개수를 반환.

    seen: 같은 페이지에서 이미 표시한 (종류/검색어) 집합 — 같은 텍스트를 여러
    finding 이 가리킬 때 같은 자리에 박스를 겹쳐 그리지 않기 위한 중복 방지."""
    types = " ".join(finding.get("types") or [])
    rect = page.rect
    if "빈 페이지" in types:
        if "빈페이지" in seen:
            return 0
        seen.add("빈페이지")
        box = fitz.Rect(rect.x0 + 6, rect.y0 + 6, rect.x1 - 6, rect.y1 - 6)
        page.draw_rect(box, color=_RED, width=2)
        loc = _location_label(finding)
        if loc:
            _insert_label(page, box.x0 + 6, box.y0 + 14,
                          f"[빈 페이지 의심] {loc}", font_path)
        return 1
    if "테두리" in types:
        if "테두리" in seen:
            return 0
        seen.add("테두리")
        # review._bottom_border_finding 이 검사한 하단 영역과 동일한 범위
        box = fitz.Rect(rect.width * 0.07, rect.height * 0.72,
                        rect.width * 0.93, rect.height * 0.965)
        page.draw_rect(box, color=_RED, width=1.2, dashes="[4 3] 0")
        loc = _location_label(finding)
        if loc:
            _insert_label(page, box.x0 + 6, box.y1 - 6,
                          f"[하단 테두리 확인] {loc}", font_path)
        return 1
    text = _search_text(finding)
    if not text or ("text", text) in seen:
        return 0
    seen.add(("text", text))
    n = 0
    for r in page.search_for(text)[:_MAX_HITS_PER_FINDING]:
        page.draw_rect(fitz.Rect(r.x0 - 2, r.y0 - 2, r.x1 + 2, r.y1 + 2),
                       color=_RED, width=1.2)
        n += 1
    return n


def create_marked_pdf(src_pdf, findings, out_path, manifest=None,
                      font_path=None, log=print):
    """page 가 매핑된 findings 를 빨간 표시로 그린 사본을 저장. 대상이 없으면 None.

    manifest 가 있으면 Excel 단계 finding(source='excel')은 해당 시트가 차지하는
    '모든 페이지'에서 텍스트를 찾는다 — 셀 표시 오류의 page 는 시트 첫 페이지로만
    매핑되어 있어, 여러 장짜리 시트에서 셀이 뒷 페이지에 인쇄되는 경우를 놓치지
    않기 위함이다. (인쇄영역 밖 셀은 PDF 에 없으므로 표시되지 않는다.)"""
    sheet_pages = {}
    for m in manifest or []:
        if m.get("type") == "sheet":
            sheet_pages.setdefault((m.get("file"), m.get("sheet")), []).append(m["page"])

    per_page = {}
    for f in findings:
        pages = [f.get("page")]
        if f.get("source") == "excel":
            pages = sheet_pages.get((f.get("file"), f.get("sheet"))) or pages
        for p in pages:
            if isinstance(p, int) and p >= 1:
                per_page.setdefault(p, []).append(f)
    if not per_page:
        return None

    doc = fitz.open(src_pdf)
    try:
        boxes = 0
        marked_pages = 0
        for pno, items in per_page.items():
            if pno > doc.page_count:
                continue
            page = doc.load_page(pno - 1)
            seen = set()
            n = 0
            for f in items:
                n += _mark(page, f, seen, font_path)
            boxes += n
            if n:
                marked_pages += 1
        doc.save(out_path, deflate=True, garbage=3)
    finally:
        doc.close()
    log(f"검토 표시 PDF 저장({marked_pages}페이지, 빨간 박스 {boxes}개): {out_path}")
    return out_path
