# -*- coding: utf-8 -*-
"""merge: 간지(divider) 생성 + 시트 PDF들을 단일 PDF로 병합 + manifest 작성.

- 폴더/파일 진입 시 간지 페이지를 만들어 끼운다(제목=폴더명/파일명, 한글 폰트).
- 시트 PDF를 순서대로 이어붙이며, 페이지↔(폴더/파일/시트) manifest 를 기록한다.
"""
from __future__ import annotations

import os

import fitz  # pymupdf

from config import A4_HEIGHT_PT, A4_WIDTH_PT

_FONT_TAG = "krfont"


class Merger:
    def __init__(self, font_path: str, log=print):
        self.doc = fitz.open()
        self.log = log
        self.manifest = []  # [{page, type, folder, file, sheet}]
        self.font_path = font_path if (font_path and os.path.isfile(font_path)) else None
        self._font = None
        if self.font_path:
            try:
                self._font = fitz.Font(fontfile=self.font_path)
            except Exception:
                self._font = None

    # --- 내부: 텍스트 폭(가운데 정렬용) ---
    def _text_len(self, text, size):
        if self._font:
            try:
                return self._font.text_length(text, fontsize=size)
            except Exception:
                pass
        return len(text) * size * 0.5

    def _put(self, page, text, size, y, color=(0, 0, 0)):
        x = max(40, (A4_WIDTH_PT - self._text_len(text, size)) / 2)
        if self.font_path:
            page.insert_text((x, y), text, fontsize=size,
                             fontname=_FONT_TAG, fontfile=self.font_path, color=color)
        else:
            page.insert_text((x, y), text, fontsize=size, color=color)

    def _name_page(self, title):
        """간지: 군더더기 없이 이름만 가운데에 표시(스타일/라벨/부제 없음)."""
        page = self.doc.new_page(width=A4_WIDTH_PT, height=A4_HEIGHT_PT)
        size = 28
        while size > 12 and self._text_len(title, size) > A4_WIDTH_PT - 80:
            size -= 1
        self._put(page, title, size, A4_HEIGHT_PT / 2)
        return page

    def add_folder_divider(self, name, relpath, depth):
        self._name_page(name)  # 폴더명만
        self.manifest.append({"page": len(self.doc), "type": "folder_divider",
                              "folder": relpath, "file": None, "sheet": None})

    def add_file_divider(self, name, relpath):
        title = os.path.splitext(name)[0]  # 확장자 제거, 파일명만
        self._name_page(title)
        self.manifest.append({"page": len(self.doc), "type": "file_divider",
                              "folder": os.path.dirname(relpath), "file": relpath,
                              "sheet": None})

    def add_sheet(self, pdf_path, folder_rel, file_rel, sheet_name):
        try:
            src = fitz.open(pdf_path)
        except Exception as e:
            self.log(f"    [경고] 시트 PDF 열기 실패({sheet_name}): {e}")
            return
        start = len(self.doc)
        self.doc.insert_pdf(src)
        n = src.page_count
        src.close()
        for k in range(n):
            self.manifest.append({"page": start + k + 1, "type": "sheet",
                                  "folder": folder_rel, "file": file_rel,
                                  "sheet": sheet_name})

    def page_count(self):
        return len(self.doc)

    def save(self, out_path):
        os.makedirs(os.path.dirname(out_path), exist_ok=True)
        self.doc.save(out_path, deflate=True, garbage=3)
        self.doc.close()
