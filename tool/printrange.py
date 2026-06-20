# -*- coding: utf-8 -*-
"""PrintRangeAI: 출력영역/포함 시트 결정.

규칙(필수 경로):
  - 숨김/매우숨김 시트 제외
  - 제외 시트명 패턴 매칭 시 제외
  - 기존 PrintArea 가 있으면 그대로 존중
  - 없으면 UsedRange 기반 추론(숨김 행/열은 Excel 인쇄 시 자동 제외됨)
AI(선택): PrintArea 없고 ai_judge 가 가능할 때만, 애매한 시트에 한해 호출.
실패/미사용 시 규칙 폴백.
"""
from __future__ import annotations


def _norm(s: str) -> str:
    return (s or "").strip()


def is_sheet_excluded(ws, compiled_patterns) -> str | None:
    """제외 사유 문자열 반환(제외 시) / None(포함)."""
    try:
        if int(ws.Visible) != -1:  # 숨김 또는 매우숨김
            return "hidden"
    except Exception:
        pass
    name = _norm(ws.Name)
    for pat in compiled_patterns:
        if pat.search(name):
            return f"name_pattern:{pat.pattern}"
    return None


def _used_range(ws):
    try:
        ur = ws.UsedRange
        if ur is None:
            return None
        if ur.Rows.Count == 1 and ur.Columns.Count == 1:
            # 빈 시트 가능성: 값 확인
            if ur.Value in (None, ""):
                return None
        return ur
    except Exception:
        return None


def decide_area(ws, ai_judge=None) -> dict:
    """포함 시트의 출력영역 결정.

    Returns dict: {area_obj, area_addr, reason, confidence}
    area_obj 가 None 이면 출력할 내용 없음(스킵).
    """
    ps = ws.PageSetup
    existing = ""
    try:
        existing = _norm(ps.PrintArea)
    except Exception:
        existing = ""

    if existing:
        try:
            area = ws.Range(existing)
            return {"area_obj": area, "area_addr": existing,
                    "reason": "rule:print_area", "confidence": 1.0}
        except Exception:
            pass  # 잘못된 PrintArea → 아래 추론으로

    ur = _used_range(ws)
    if ur is None:
        return {"area_obj": None, "area_addr": "", "reason": "rule:empty",
                "confidence": 1.0}

    # AI 사용 가능하면 애매 시트로 보고 판단 위임 (선택)
    if ai_judge is not None:
        try:
            verdict = ai_judge.judge_sheet(ws, ur)
            if verdict and verdict.get("area_addr"):
                area = ws.Range(verdict["area_addr"])
                return {"area_obj": area, "area_addr": verdict["area_addr"],
                        "reason": "ai", "confidence": verdict.get("confidence", 0.5)}
        except Exception as e:
            print(f"    [경고] AI 판단 실패({ws.Name}): {e} → 규칙 폴백")

    addr = ur.Address  # $A$1:$..
    return {"area_obj": ur, "area_addr": addr,
            "reason": "rule:usedrange", "confidence": 0.8}
