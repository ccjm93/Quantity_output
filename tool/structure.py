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


def _is_skip(name: str, output_dir_name: str, backup_suffix: str,
             skip_names=()) -> bool:
    if name.startswith("~$") or name.startswith("."):
        return True
    if name == output_dir_name:
        return True
    if backup_suffix and backup_suffix in name:
        return True
    if name in skip_names:  # 도구 자신이 만든 산출물(검토 결과표 등)은 입력으로 취급하지 않음
        return True
    return False


def scan(root: str, extensions, output_dir_name="_output", backup_suffix="_backup",
         skip_names=()):
    root = os.path.abspath(root)
    exts = tuple(e.lower() for e in extensions)

    def walk(folder: str, depth: int) -> list:
        try:
            entries = os.listdir(folder)
        except Exception:
            return []
        dirs, files = [], []
        for e in entries:
            full = os.path.join(folder, e)
            if _is_skip(e, output_dir_name, backup_suffix, skip_names):
                continue
            if os.path.isdir(full):
                dirs.append(e)
            elif os.path.splitext(e)[1].lower() in exts:
                files.append(e)
        # 폴더+파일을 같은 레벨에서 함께 자연정렬
        merged = sorted([(d, True) for d in dirs] + [(f, False) for f in files],
                        key=lambda x: natural_key(x[0]))
        events = []
        for name, is_dir in merged:
            full = os.path.join(folder, name)
            rel = os.path.relpath(full, root)
            if is_dir:
                children = walk(full, depth + 1)
                # 하위에 처리할 파일이 하나도 없는 폴더는 간지만 남으므로 제외
                if children:
                    events.append({"type": "folder", "name": name,
                                   "relpath": rel, "depth": depth})
                    events.extend(children)
            else:
                events.append({"type": "file", "name": name, "relpath": rel,
                               "abspath": full, "depth": depth})
        return events

    return walk(root, 0)


def summarize(events) -> str:
    folders = sum(1 for e in events if e["type"] == "folder")
    files = sum(1 for e in events if e["type"] == "file")
    return f"폴더 {folders}개, 파일 {files}개"
