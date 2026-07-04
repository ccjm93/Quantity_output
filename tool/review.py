# -*- coding: utf-8 -*-
"""review: 최종 병합 PDF 를 페이지별로 검토해 이상(깨진 문자/숫자, 비정상 출력범위)과
그 '위치'(폴더/파일/시트/페이지)를 보고한다.

- AI(Gemini 멀티모달) 사용 가능 시: 각 페이지 이미지를 검토.
- 키 없거나 실패 시: 규칙 기반 폴백(### 텍스트 탐지 등).
"""
from __future__ import annotations

import json
import os

import fitz  # pymupdf


def is_available(no_ai: bool) -> tuple[bool, str]:
    if no_ai:
        return False, "--no-ai"
    if not os.environ.get("GEMINI_API_KEY"):
        return False, "GEMINI_API_KEY 없음"
    try:
        import google.genai  # noqa
    except Exception:
        return False, "google-genai 미설치"
    return True, "사용 가능"


def _loc(manifest_by_page, page_no):
    return manifest_by_page.get(page_no, {})


def _rule_scan(doc, manifest_by_page, log):
    findings = []
    for i in range(doc.page_count):
        txt = doc.load_page(i).get_text()
        if "###" in txt:
            loc = _loc(manifest_by_page, i + 1)
            findings.append({"page": i + 1, "types": ["###표시"], "source": "rule",
                             "detail": "셀 값이 ###로 표시됨(열 폭 부족 가능)",
                             **{k: loc.get(k) for k in ("folder", "file", "sheet")}})
    log(f"  규칙 검토 완료: 이상 의심 {len(findings)}건")
    return findings


def _ai_scan(doc, manifest_by_page, cfg, log):
    from google import genai
    from google.genai import types
    client = genai.Client(api_key=os.environ["GEMINI_API_KEY"])
    model = cfg.get("ai_model", "gemini-2.5-flash")
    dpi = int(cfg.get("ai_dpi", 200))
    prompt = (
        "이 이미지는 토목 수량산출서 PDF의 한 페이지입니다. 다음 이상만 점검하세요: "
        "(1) 숫자나 문자가 '###', 깨짐, 잘림 등으로 비정상 표시됨 "
        "(표/단면도/그림 안의 작은 셀이 '###'나 잘림으로 표시되는 경우도 반드시 포함), "
        "(2) 출력 범위 이상(표 일부가 잘림, 거의 빈 페이지, 내용이 인쇄영역을 벗어남). "
        "정상이면 반드시 {\"issue\": false} 만, 이상이면 "
        "{\"issue\": true, \"types\": [\"...\"], \"detail\": \"간단 설명\"} JSON 만 출력."
    )
    findings = []
    total = doc.page_count
    for i in range(total):
        if (i + 1) % 20 == 0:
            log(f"  AI 검토 진행 {i+1}/{total}...")
        try:
            pix = doc.load_page(i).get_pixmap(dpi=dpi)
            png = pix.tobytes("png")
            resp = client.models.generate_content(
                model=model,
                contents=[types.Part.from_bytes(data=png, mime_type="image/png"), prompt])
            t = (resp.text or "").strip()
            if t.startswith("```"):
                t = t.strip("`")
                if t.startswith("json"):
                    t = t[4:]
            data = json.loads(t)
        except Exception:
            continue
        if data.get("issue"):
            loc = _loc(manifest_by_page, i + 1)
            findings.append({"page": i + 1, "source": "ai",
                             "types": data.get("types", []),
                             "detail": data.get("detail", ""),
                             **{k: loc.get(k) for k in ("folder", "file", "sheet")}})
    log(f"  AI 검토 완료: 이상 {len(findings)}건")
    return findings


def review(merged_pdf, manifest, cfg, no_ai, log=print, extra_findings=None) -> dict:
    """extra_findings: Excel 단계에서 결정론적으로 탐지한 셀 표시 오류(source='excel').
    이미 page 가 매핑된 finding 리스트를 받아 AI/규칙 검토 결과와 병합한다."""
    ok, reason = is_available(no_ai)
    manifest_by_page = {m["page"]: m for m in manifest}
    doc = fitz.open(merged_pdf)
    try:
        if ok:
            log(f"AI 사후검토 시작 (총 {doc.page_count}페이지)")
            findings = _ai_scan(doc, manifest_by_page, cfg, log)
            mode = "ai"
        else:
            log(f"AI 미사용({reason}) → 규칙 기반 검토 (총 {doc.page_count}페이지)")
            findings = _rule_scan(doc, manifest_by_page, log)
            mode = "rule"
    finally:
        doc.close()

    extra = list(extra_findings or [])
    if extra:
        log(f"  Excel 단계 결정론 탐지 병합: {len(extra)}건")
    # Excel-단계 findings 를 앞에 두어(확실한 이상) 페이지 순으로 정렬
    merged = extra + findings
    merged.sort(key=lambda f: (f.get("page") or 0))
    return {"mode": mode, "reason": reason, "total_pages": len(manifest_by_page),
            "issue_count": len(merged), "findings": merged}
