# -*- coding: utf-8 -*-
"""수량산출서 출력 자동화 도구 - 설정/상수.

여기 값들은 일반 토목 수량산출서 관행을 기준으로 한 기본값이며,
프로젝트마다 다르면 CLI 인자나 settings.json 으로 덮어쓸 수 있다.
"""
from __future__ import annotations

import json
import os
import re

# ---- Excel COM 상수 (gencache 없이 직접 사용) ----
XL_TYPE_PDF = 0
XL_PAPER_A4 = 9
XL_PORTRAIT = 1
XL_LANDSCAPE = 2
XL_CALC_MANUAL = -4135
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3
XL_SHEET_VISIBLE = -1
XL_SHEET_HIDDEN = 0
XL_SHEET_VERYHIDDEN = 2

# ---- 용지/여백 (포인트, 1cm = 28.3465pt) ----
CM = 28.3465
A4_WIDTH_PT = 210 * 2.834645   # 595.28
A4_HEIGHT_PT = 297 * 2.834645  # 841.89

DEFAULTS = {
    # 출력물 루트 (입력 프로젝트 폴더 기준 상대). 절대경로면 그대로 사용.
    "output_dir_name": "_output",
    # 여백 (cm)
    "margin_left_cm": 1.0,
    "margin_right_cm": 1.0,
    "margin_top_cm": 1.2,
    "margin_bottom_cm": 1.2,
    # ### 보정: AutoFit 후 열 너비(문자수) 상한 (폭 폭주 방지)
    "max_col_width_chars": 45.0,
    # 페이지 배율 하한(%). 이보다 더 줄여야 하면 가로방향 전환 시도.
    "min_zoom": 40,
    # AI 호출당 처리 단위 등은 추후 확장
    "ai_model": "gemini-2.5-flash",
    # 처리 대상 확장자
    "extensions": [".xlsx", ".xlsm", ".xls"],
}

# 제외 시트명 패턴 (대소문자 무시, 부분일치/정규식 혼합).
# 일반 수량산출서 관행: 출력 금지 표시, 중복/스크래치 시트.
DEFAULT_EXCLUDE_SHEET_PATTERNS = [
    r"출력\s*하지\s*마세요",
    r"출력\s*x",
    r"출력\s*안",
    r"\(\s*2\s*\)\s*$",      # 이름 끝 "(2)"
    r"-\s*1\s*$",            # 이름 끝 "-1"
    r"^sheet\d+$",           # Sheet1, Sheet2 ...
    r"설계조건",
    r"^\d{3,}",             # "3333돌망태집계" 류 스크래치 접두
]


def load_settings(project_dir: str | None = None) -> dict:
    """기본값 + (있으면) settings.json 병합."""
    cfg = dict(DEFAULTS)
    cfg["exclude_sheet_patterns"] = list(DEFAULT_EXCLUDE_SHEET_PATTERNS)

    candidates = []
    if project_dir:
        candidates.append(os.path.join(project_dir, "qto_settings.json"))
    candidates.append(os.path.join(os.path.dirname(__file__), "qto_settings.json"))

    for path in candidates:
        if os.path.isfile(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    user = json.load(f)
                if "exclude_sheet_patterns" in user:
                    cfg["exclude_sheet_patterns"] = user.pop("exclude_sheet_patterns")
                cfg.update(user)
            except Exception as e:  # 설정 파일 오류는 치명적이지 않게
                print(f"[경고] 설정 파일 읽기 실패({path}): {e}")
            break
    return cfg


def compile_exclude_patterns(patterns: list[str]):
    return [re.compile(p, re.IGNORECASE) for p in patterns]
