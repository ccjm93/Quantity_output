# -*- coding: utf-8 -*-
"""Excel/COM 없이 실행 가능한 순수 로직 유닛테스트.

실행: python tool/test_units.py
(pywin32/pymupdf/openpyxl 은 import 용으로만 필요하며 Excel 실행은 필요 없다.)
"""
from __future__ import annotations

import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import detect  # noqa: E402
import review  # noqa: E402
import structure  # noqa: E402
import toolruntime  # noqa: E402
from config import compile_exclude_patterns, load_settings  # noqa: E402


class TestStructure(unittest.TestCase):
    def test_natural_key_order(self):
        names = ["10. 부대공", "2. 토공", "1. 공통공사"]
        self.assertEqual(sorted(names, key=structure.natural_key),
                         ["1. 공통공사", "2. 토공", "10. 부대공"])

    def test_skip_names(self):
        self.assertTrue(structure._is_skip("~$임시.xlsx", "_output", "_backup"))
        self.assertTrue(structure._is_skip("_output", "_output", "_backup"))
        self.assertTrue(structure._is_skip("루트_backup_20260101", "_output", "_backup"))
        self.assertFalse(structure._is_skip("2.0 토공.xlsx", "_output", "_backup"))

    def test_scan_skips_tool_artifacts(self):
        with tempfile.TemporaryDirectory() as d:
            open(os.path.join(d, "1. 공통.xlsx"), "w").close()
            open(os.path.join(d, "출력물 오류 검토결과.xlsx"), "w").close()
            events = structure.scan(d, [".xlsx"],
                                    skip_names={"출력물 오류 검토결과.xlsx"})
            self.assertEqual([e["name"] for e in events], ["1. 공통.xlsx"])

    def test_scan_orders_folders_and_files(self):
        with tempfile.TemporaryDirectory() as d:
            os.makedirs(os.path.join(d, "2. 취수틀"))
            os.makedirs(os.path.join(d, "10. 기타"))
            open(os.path.join(d, "1. 공통.xlsx"), "w").close()
            open(os.path.join(d, "2. 취수틀", "2.0 토공.xlsx"), "w").close()
            events = structure.scan(d, [".xlsx"])
            names = [e["name"] for e in events]
            self.assertEqual(names, ["1. 공통.xlsx", "2. 취수틀", "2.0 토공.xlsx", "10. 기타"])


class TestConfig(unittest.TestCase):
    def test_defaults_loaded(self):
        cfg = load_settings(None)
        self.assertEqual(cfg["output_dir_name"], "_output")
        self.assertFalse(cfg["set_print_area_if_missing"])
        self.assertIn("exclude_sheet_patterns", cfg)

    def test_exclude_patterns_match(self):
        pats = compile_exclude_patterns(load_settings(None)["exclude_sheet_patterns"])

        def excluded(name):
            return any(p.search(name) for p in pats)

        self.assertTrue(excluded("이후시트는 출력하지마세요"))
        self.assertTrue(excluded("토공집계 (2)"))
        self.assertTrue(excluded("Sheet1"))
        self.assertFalse(excluded("주요자재집계표"))


class TestDetectRegex(unittest.TestCase):
    def test_hash_only(self):
        self.assertTrue(detect._HASH_ONLY.match("####"))
        self.assertTrue(detect._HASH_ONLY.match("##  "))
        self.assertFalse(detect._HASH_ONLY.match("#"))
        self.assertFalse(detect._HASH_ONLY.match("#REF!"))
        self.assertFalse(detect._HASH_ONLY.match("가##"))


class TestReview(unittest.TestCase):
    def test_text_findings_ref_and_hash(self):
        found = review._text_findings("합계 #REF! 오류 ####", 3, {})
        types = [f["types"][0] for f in found]
        self.assertIn("#REF! 오류", types)
        self.assertIn("# 연속 표시", types)

    def test_groups(self):
        self.assertEqual(review._groups([1, 2, 3, 10, 11, 30]),
                         [(1, 3), (10, 11), (30, 30)])
        self.assertEqual(review._groups([]), [])


class TestToolruntime(unittest.TestCase):
    def test_normalize_pdf_path(self):
        self.assertTrue(toolruntime._normalize_pdf_path("out").endswith(".pdf"))
        self.assertTrue(toolruntime._normalize_pdf_path("out.PDF").endswith("out.PDF"))

    def test_korean_error_types(self):
        k = toolruntime._korean_error_types
        self.assertEqual(k(["#REF! 오류"]), "#REF! 오류")
        self.assertEqual(k(["# 연속 표시"]), "# 연속 표시")
        self.assertEqual(k(["수식 오류"]), "수식 오류")
        self.assertEqual(k(["빈 페이지"]), "빈 페이지 의심")
        self.assertEqual(k(["하단 표 테두리 누락 의심"]), "하단 표 테두리 누락 의심")
        self.assertEqual(k(["#REF! 오류", "#REF! 오류"]), "#REF! 오류")  # 중복 제거

    def test_assign_pages(self):
        manifest = [
            {"type": "file_divider", "page": 1, "file": "a.xlsx", "sheet": None},
            {"type": "sheet", "page": 2, "file": "a.xlsx", "sheet": "토공"},
            {"type": "sheet", "page": 3, "file": "a.xlsx", "sheet": "토공"},
        ]
        findings = [{"file": "a.xlsx", "sheet": "토공", "cell": "B2",
                     "kind": "overflow", "text": "###"}]
        toolruntime._assign_pages(findings, manifest)
        self.assertEqual(findings[0]["page"], 2)  # 시트 첫 페이지
        self.assertEqual(findings[0]["types"], ["# 연속 표시"])

    def test_find_review_pdf_prefers_configured_name(self):
        with tempfile.TemporaryDirectory() as d:
            open(os.path.join(d, "AAA 다른문서.pdf"), "w").close()
            open(os.path.join(d, "수량산출서 output.pdf"), "w").close()
            cfg = load_settings(None)
            pdf = toolruntime._find_review_pdf(d, cfg)
            self.assertEqual(os.path.basename(pdf), "수량산출서 output.pdf")

    def test_display_width(self):
        self.assertEqual(toolruntime._display_width("abc"), 3)
        self.assertEqual(toolruntime._display_width("가나"), 4)

    def test_save_review_xlsx_without_excel(self):
        rev = {"findings": [{"page": 5, "types": ["#REF! 오류"],
                             "detail": "시트!B2 → '#REF!'", "file": "a.xlsx",
                             "folder": "", "sheet": "토공"}]}
        with tempfile.TemporaryDirectory() as d:
            path = toolruntime.save_review_xlsx(rev, d, log=lambda *a: None)
            self.assertTrue(os.path.isfile(path))
            from openpyxl import load_workbook
            wb = load_workbook(path)
            ws = wb.active
            self.assertEqual(ws["C2"].value, "#REF! 오류")
            self.assertEqual(ws["B2"].value, "최종 PDF 5페이지")
            self.assertEqual(ws.freeze_panes, "A2")


if __name__ == "__main__":
    unittest.main(verbosity=2)
