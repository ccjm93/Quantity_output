# -*- coding: utf-8 -*-
"""review: 최종 병합 PDF 를 페이지별로 검토해 이상(## 이상, #REF!, 빈 페이지, 표 테두리 누락 의심)과
그 '위치'(폴더/파일/시트/페이지)를 보고한다.
"""
from __future__ import annotations

import re

import fitz  # pymupdf


_HASH_RUN = re.compile(r"#{2,}")
_REF_ERROR = re.compile(r"#REF!", re.IGNORECASE)


def _loc(manifest_by_page, page_no):
    return manifest_by_page.get(page_no, {})


def _text_findings(page_text, page_no, manifest_by_page):
    findings = []
    loc = _loc(manifest_by_page, page_no)
    if _REF_ERROR.search(page_text):
        findings.append({"page": page_no, "types": ["#REF! 오류"], "source": "program",
                         "detail": "PDF 텍스트에서 #REF! 표시가 발견되었습니다.",
                         **{k: loc.get(k) for k in ("folder", "file", "sheet")}})
    hash_runs = [m.group(0) for m in _HASH_RUN.finditer(page_text)
                 if not page_text[m.start():m.start() + 5].upper().startswith("#REF!")]
    if hash_runs:
        longest = max(hash_runs, key=len)
        findings.append({"page": page_no, "types": ["# 연속 표시"], "source": "program",
                         "detail": f"PDF 텍스트에서 연속 # 표시가 발견되었습니다: '{longest}'",
                         **{k: loc.get(k) for k in ("folder", "file", "sheet")}})
    return findings


def _is_dark(samples, idx):
    return samples[idx] < 90 and samples[idx + 1] < 90 and samples[idx + 2] < 90


def _blank_page_finding(pix, page_text, page_no, manifest_by_page):
    if page_text.strip():
        return None
    w, h, n = pix.width, pix.height, pix.n
    if n < 3:
        return None
    samples = pix.samples
    dark = 0
    total = 0
    step = 4
    for y in range(0, h, step):
        row_base = y * w * n
        for x in range(0, w, step):
            total += 1
            if _is_dark(samples, row_base + x * n):
                dark += 1
    if total == 0 or dark / total > 0.0005:
        return None
    loc = _loc(manifest_by_page, page_no)
    return {"page": page_no, "types": ["빈 페이지"], "source": "program",
            "detail": "PDF 페이지에 텍스트와 눈에 띄는 선/문자 요소가 거의 없어 빈 페이지로 의심됩니다.",
            **{k: loc.get(k) for k in ("folder", "file", "sheet")}}


def _groups(values):
    groups = []
    start = prev = None
    for value in values:
        if start is None:
            start = prev = value
        elif value <= prev + 3:
            prev = value
        else:
            groups.append((start, prev))
            start = prev = value
    if start is not None:
        groups.append((start, prev))
    return groups


def _bottom_border_finding(pix, page_no, manifest_by_page):
    """페이지 하단에서 표 세로선은 이어지지만 닫는 가로선이 약한 경우를 휴리스틱으로 탐지."""
    w, h, n = pix.width, pix.height, pix.n
    if n < 3 or w < 80 or h < 120:
        return None
    samples = pix.samples
    x0, x1 = int(w * 0.07), int(w * 0.93)
    y0, y1 = int(h * 0.72), int(h * 0.965)
    step = 2
    band_rows = max(1, (y1 - y0) // step)
    col_counts = {}

    for y in range(y0, y1, step):
        row_base = y * w * n
        for x in range(x0, x1, step):
            idx = row_base + x * n
            if _is_dark(samples, idx):
                col_counts[x] = col_counts.get(x, 0) + 1

    vertical_xs = sorted(x for x, count in col_counts.items() if count >= band_rows * 0.22)
    col_groups = [g for g in _groups(vertical_xs) if g[1] - g[0] <= 8]
    if len(col_groups) < 2:
        return None

    first_x, last_x = col_groups[0][0], col_groups[-1][1]
    if last_x - first_x < w * 0.25:
        return None

    vertical_bottom = 0
    for y in range(y0, y1, step):
        row_base = y * w * n
        for g0, g1 in col_groups:
            mid = (g0 + g1) // 2
            idx = row_base + mid * n
            if _is_dark(samples, idx):
                vertical_bottom = max(vertical_bottom, y)
                break
    if vertical_bottom < h * 0.84:
        return None

    search0 = max(y0, vertical_bottom - 10)
    search1 = min(y1, vertical_bottom + 10)
    span_samples = max(1, (last_x - first_x) // step)
    horizontal_found = False
    for y in range(search0, search1, step):
        row_base = y * w * n
        dark = 0
        for x in range(first_x, last_x, step):
            if _is_dark(samples, row_base + x * n):
                dark += 1
        if dark >= span_samples * 0.28:
            horizontal_found = True
            break
    if horizontal_found:
        return None

    loc = _loc(manifest_by_page, page_no)
    return {"page": page_no, "types": ["하단 표 테두리 누락 의심"], "source": "program",
            "detail": "페이지 하단 표 영역에서 세로 테두리는 보이나 닫는 가로 테두리가 약하거나 누락된 것으로 의심됩니다.",
            **{k: loc.get(k) for k in ("folder", "file", "sheet")}}


def _rule_scan(doc, manifest_by_page, cfg, log):
    findings = []
    dpi = int(cfg.get("rule_dpi", 100))
    for i in range(doc.page_count):
        page_no = i + 1
        if page_no % 20 == 0:
            log(f"  출력물 검토 진행 {page_no}/{doc.page_count}...")
        page = doc.load_page(i)
        page_text = page.get_text()
        # 렌더링은 페이지당 1회만 — 빈 페이지/하단 테두리 검사가 공유한다.
        pix = page.get_pixmap(dpi=dpi, alpha=False)
        blank = _blank_page_finding(pix, page_text, page_no, manifest_by_page)
        if blank:
            findings.append(blank)
            continue
        findings.extend(_text_findings(page_text, page_no, manifest_by_page))
        border = _bottom_border_finding(pix, page_no, manifest_by_page)
        if border:
            findings.append(border)
    log(f"  프로그램 검토 완료: 이상 의심 {len(findings)}건")
    return findings


def review(merged_pdf, manifest, cfg, log=print, extra_findings=None) -> dict:
    """extra_findings: Excel 단계에서 결정론적으로 탐지한 셀 표시 오류(source='excel').
    이미 page 가 매핑된 finding 리스트를 받아 출력물 검토 결과와 병합한다."""
    manifest_by_page = {m["page"]: m for m in manifest}
    doc = fitz.open(merged_pdf)
    try:
        log(f"프로그램 내부 오류 검토 시작 (총 {doc.page_count}페이지)")
        findings = _rule_scan(doc, manifest_by_page, cfg, log)
    finally:
        doc.close()

    extra = list(extra_findings or [])
    if extra:
        log(f"  Excel 단계 결정론 탐지 병합: {len(extra)}건")
    # Excel-단계 findings 를 앞에 두어(확실한 이상) 페이지 순으로 정렬
    merged = extra + findings
    merged.sort(key=lambda f: (f.get("page") or 0))
    return {"mode": "program", "reason": "프로그램 내부 로직", "total_pages": len(manifest_by_page),
            "issue_count": len(merged), "findings": merged}
