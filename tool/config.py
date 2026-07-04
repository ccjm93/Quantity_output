# -*- coding: utf-8 -*-
"""수량산출서 출력 자동화 도구 - 설정/상수.

기본값은 일반 토목 수량산출서 관행 기준이며, 프로젝트별로 qto_settings.json 으로
덮어쓸 수 있다.
"""
from __future__ import annotations

import json
import os
import re

# ---- Excel COM 상수 (gencache 없이 직접 사용) ----
XL_TYPE_PDF = 0
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3
XL_CALC_MANUAL = -4135
XL_SHEET_VISIBLE = -1
XL_VIEW_PAGEBREAK_PREVIEW = 2   # ActiveWindow.View = xlPageBreakPreview

# ---- SpecialCells (셀 표시 오류 탐지용) ----
XL_CELLTYPE_CONSTANTS = 2       # xlCellTypeConstants
XL_CELLTYPE_FORMULAS = -4123    # xlCellTypeFormulas
XL_ERRORS = 16                  # xlErrors

# ---- 용지(간지 페이지용, 포인트) ----
A4_WIDTH_PT = 595.0
A4_HEIGHT_PT = 842.0

DEFAULTS = {
    # 산출물 폴더명 (입력 루트 아래에 생성, 모든 산출물이 여기에 모임)
    "output_dir_name": "_output",
    # 기본 통합 PDF 파일명 (산출물 폴더 안에 저장)
    "output_pdf_name": "수량산출서 output.pdf",
    # 원본 변경 전 자동 백업 (원본에 PBP 저장하므로 기본 ON)
    "backup": True,
    "backup_suffix": "_backup",
    # 인쇄영역 처리: 원래 지정된 인쇄영역은 그대로 두고,
    # 지정되지 않은 시트는 Excel 기본값(설정 안 함=자동 결정)으로 둔다.
    # True 로 바꾸면 인쇄영역 없는 시트를 UsedRange 로 설정한다(기본 False).
    "set_print_area_if_missing": False,
    # 간지 한글 폰트
    "divider_font": r"C:\Windows\Fonts\malgun.ttf",
    # 출력물 이미지 검토 해상도(dpi). 높을수록 하단 테두리 휴리스틱이 촘촘해짐.
    "rule_dpi": 100,
    # Excel 단계 셀 표시 오류(### 오버플로우 / 수식오류) 결정론적 탐지
    "detect_cell_issues": True,
    # 처리 대상 확장자
    "extensions": [".xlsx", ".xlsm", ".xls"],
}

# 제외 시트명 패턴 (대소문자 무시). 명시적 '출력 금지' 표지 + 중복/스크래치 시트.
DEFAULT_EXCLUDE_SHEET_PATTERNS = [
    r"출력\s*하지\s*\s*마",   # "출력하지마세요"
    r"출력\s*x",              # "출력X(...)"
    r"출력\s*안",             # "출력안함"
    r"\(\s*2\s*\)\s*$",       # 끝 "(2)"
    r"-\s*1\s*$",             # 끝 "-1"
    r"^sheet\d+$",            # Sheet1, Sheet2 ...
]


def load_settings(project_dir: str | None = None) -> dict:
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
            except Exception as e:
                print(f"[경고] 설정 파일 읽기 실패({path}): {e}")
            break
    return cfg


def compile_exclude_patterns(patterns: list[str]):
    return [re.compile(p, re.IGNORECASE) for p in patterns]
