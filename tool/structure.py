# -*- coding: utf-8 -*-
"""구조 스캔: 루트 폴더의 폴더/파일을 '문서(자연) 순서'로 나열.

자연 정렬로 '2' < '10' 을 보장해 목차 순서를 유지한다.
폴더·파일을 같은 레벨에서 함께 정렬한 뒤 폴더는 그 위치에서 재귀한다.
반환: 순서대로의 이벤트 리스트
  {"type": "folder", "name", "relpath", "depth"}
  {"type": "file",   "name", "relpath", "abspath", "depth"}
"""
from __future__ import annotations

import os
import re

_num = re.compile(r"(\d+)")


def natural_key(s: str):
    return [int(t) if t.isdigit() else t.lower() for t in _num.split(s)]


def _is_skip(name: str, output_dir_name: str, backup_suffix: str) -> bool:
    if name.startswith("~$") or name.startswith("."):
        return True
    if name == output_dir_name:
        return True
    if backup_suffix and backup_suffix in name:
        return True
    return False


def scan(root: str, extensions, output_dir_name="_output", backup_suffix="_backup"):
    root = os.path.abspath(root)
    exts = tuple(e.lower() for e in extensions)
    events = []

    def walk(folder: str, depth: int):
        try:
            entries = os.listdir(folder)
        except Exception:
            return
        dirs, files = [], []
        for e in entries:
            full = os.path.join(folder, e)
            if _is_skip(e, output_dir_name, backup_suffix):
                continue
            if os.path.isdir(full):
                dirs.append(e)
            elif os.path.splitext(e)[1].lower() in exts:
                files.append(e)
        # 폴더+파일을 같은 레벨에서 함께 자연정렬
        merged = sorted([(d, True) for d in dirs] + [(f, False) for f in files],
                        key=lambda x: natural_key(x[0]))
        for name, is_dir in merged:
            full = os.path.join(folder, name)
            rel = os.path.relpath(full, root)
            if is_dir:
                events.append({"type": "folder", "name": name,
                               "relpath": rel, "depth": depth})
                walk(full, depth + 1)
            else:
                events.append({"type": "file", "name": name, "relpath": rel,
                               "abspath": full, "depth": depth})

    walk(root, 0)
    return events


def summarize(events) -> str:
    folders = sum(1 for e in events if e["type"] == "folder")
    files = sum(1 for e in events if e["type"] == "file")
    return f"폴더 {folders}개, 파일 {files}개"
