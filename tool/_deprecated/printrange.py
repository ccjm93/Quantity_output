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
                ok, cover = _ai_area_is_safe(ws, ur, area)
                if ok:
                    return {"area_obj": area, "area_addr": verdict["area_addr"],
                            "reason": "ai", "confidence": verdict.get("confidence", 0.5),
                            "ai_coverage": round(cover, 3)}
                else:
                    # AI 가 실제 데이터를 과도하게 잘라냄 → 거부, 규칙 폴백
                    print(f"    [경고] AI 영역이 데이터를 누락(coverage={cover:.2f}) "
                          f"({ws.Name}) → 규칙 폴백")
        except Exception as e:
            print(f"    [경고] AI 판단 실패({ws.Name}): {e} → 규칙 폴백")

    addr = ur.Address  # $A$1:$..
    return {"area_obj": ur, "area_addr": addr,
            "reason": "rule:usedrange", "confidence": 0.8}


# AI 영역이 실제 데이터의 이 비율 이상을 담아야 채택(미만이면 데이터 누락으로 간주)
AI_MIN_COVERAGE = 0.95


def _count_nonempty(rng) -> int:
    try:
        v = rng.Value
    except Exception:
        return 0
    if v is None:
        return 0
    if not isinstance(v, tuple):
        return 1 if v not in (None, "") else 0
    n = 0
    for row in v:
        if isinstance(row, tuple):
            for c in row:
                if c not in (None, ""):
                    n += 1
        elif row not in (None, ""):
            n += 1
    return n


def _ai_area_is_safe(ws, used_range, ai_area):
    """AI 영역이 UsedRange 내 비어있지 않은 셀의 대부분을 포함하는지 검사.
    Returns (안전여부, coverage비율). 데이터 누락 방지용 가드."""
    total = _count_nonempty(used_range)
    if total == 0:
        return True, 1.0
    inside = _count_nonempty(ai_area)
    cover = inside / total
    return cover >= AI_MIN_COVERAGE, cover
