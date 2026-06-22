# -*- coding: utf-8 -*-
"""AI 판단(선택 기능): Gemini 멀티모달로 애매한 시트의 출력영역 추론.

- 완전 선택사항. GEMINI_API_KEY 가 없거나 패키지가 없으면 None 을 반환하는
  팩토리로 동작하여 도구 전체는 규칙 기반으로 정상 동작한다.
- import 는 지연(lazy)으로 처리.
"""
from __future__ import annotations

import json
import os
import tempfile

from config import XL_TYPE_PDF


def is_available(no_ai: bool) -> tuple[bool, str]:
    """(사용가능여부, 사유)."""
    if no_ai:
        return False, "--no-ai 지정"
    if not os.environ.get("GEMINI_API_KEY"):
        return False, "GEMINI_API_KEY 환경변수 없음"
    try:
        import google.genai  # noqa: F401
    except Exception:
        return False, "google-genai 미설치"
    try:
        import fitz  # noqa: F401  (pymupdf)
    except Exception:
        return False, "pymupdf 미설치"
    return True, "사용 가능"


def create(cfg: dict, no_ai: bool):
    """가능하면 AIJudge 인스턴스, 아니면 None."""
    ok, _ = is_available(no_ai)
    if not ok:
        return None
    try:
        return _AIJudge(cfg)
    except Exception:
        return None


class _AIJudge:
    def __init__(self, cfg: dict):
        from google import genai
        self.cfg = cfg
        self.client = genai.Client(api_key=os.environ["GEMINI_API_KEY"])
        self.model = cfg.get("ai_model", "gemini-2.5-flash")

    def _capture_png(self, ws, area) -> str | None:
        """area 를 임시 PDF로 export 후 첫 페이지를 PNG로 렌더."""
        import fitz
        tmp_pdf = tempfile.mktemp(suffix=".pdf")
        tmp_png = tempfile.mktemp(suffix=".png")
        try:
            ws.PageSetup.PrintArea = area.Address
            ws.ExportAsFixedFormat(Type=XL_TYPE_PDF, Filename=tmp_pdf)
            doc = fitz.open(tmp_pdf)
            page = doc.load_page(0)
            pix = page.get_pixmap(dpi=150)
            pix.save(tmp_png)
            doc.close()
            return tmp_png
        except Exception:
            return None
        finally:
            if os.path.exists(tmp_pdf):
                try:
                    os.remove(tmp_pdf)
                except Exception:
                    pass

    def judge_sheet(self, ws, used_range) -> dict | None:
        """{area_addr, confidence} 또는 None."""
        png = self._capture_png(ws, used_range)
        if not png:
            return None
        try:
            from google.genai import types
            with open(png, "rb") as f:
                img_bytes = f.read()
            prompt = (
                "이 엑셀 시트는 토목 수량산출서의 한 시트입니다. "
                "인쇄(PDF 출력)에 포함해야 할 핵심 표 영역의 셀 범위만 알려주세요. "
                "주석/메모/임시 계산/안내문 등 본문과 무관한 영역은 제외합니다. "
                f"현재 사용 영역은 {used_range.Address} 입니다. "
                "반드시 JSON 으로만 답하세요: "
                '{"area_addr": "$A$1:$H$40", "confidence": 0.0~1.0}'
            )
            resp = self.client.models.generate_content(
                model=self.model,
                contents=[
                    types.Part.from_bytes(data=img_bytes, mime_type="image/png"),
                    prompt,
                ],
            )
            text = (resp.text or "").strip()
            if text.startswith("```"):
                text = text.strip("`")
                if text.startswith("json"):
                    text = text[4:]
            data = json.loads(text)
            return {"area_addr": data["area_addr"],
                    "confidence": float(data.get("confidence", 0.5))}
        except Exception:
            return None
        finally:
            if os.path.exists(png):
                try:
                    os.remove(png)
                except Exception:
                    pass
