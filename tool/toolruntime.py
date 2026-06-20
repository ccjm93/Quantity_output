# -*- coding: utf-8 -*-
"""수량산출서 출력 자동화 도구 - CLI / 오케스트레이터.

사용 예:
  python tool/toolruntime.py "01 수량_가야"          # 폴더 일괄
  python tool/toolruntime.py --file "....xlsx"        # 특정 파일
  python tool/toolruntime.py --help-apikey            # API 키 발급 안내
  python tool/toolruntime.py "폴더" --no-ai           # AI 끄고 규칙 기반

여러 프로젝트의 수량산출서에 범용으로 적용. 특정 폴더명에 의존하지 않는다.
"""
from __future__ import annotations

import argparse
import datetime as _dt
import json
import os
import shutil
import sys
import traceback

# 콘솔 인코딩 안정화(한글 출력 깨짐/오류 방지)
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

# 이 파일을 직접 실행할 때 동일 폴더의 모듈을 import 할 수 있게.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import ai_judge  # noqa: E402
import autofix  # noqa: E402
import pagelayout  # noqa: E402
import pdfexport  # noqa: E402
import printrange  # noqa: E402
from config import compile_exclude_patterns, load_settings  # noqa: E402
from excel_app import ExcelApp  # noqa: E402

APIKEY_HELP = """
============================================================
 Google AI Studio API 키 발급 방법 (선택사항)
============================================================
이 도구는 API 키가 없어도 정상 동작합니다.
AI 키를 등록하면 '애매한 시트의 출력영역 판단' 품질만 향상됩니다.

[1] 웹브라우저에서 아래 주소로 접속
    https://aistudio.google.com/app/apikey

[2] 구글 계정으로 로그인

[3] 'API 키 만들기 (Create API key)' 버튼 클릭

[4] 생성된 키 문자열을 복사

[5] 윈도우에 환경변수로 등록 (명령 프롬프트에서):
        setx GEMINI_API_KEY "여기에_복사한_키_붙여넣기"
    → 등록 후 '새' 명령 프롬프트/터미널을 열어야 적용됩니다.

[6] 확인:
        echo %GEMINI_API_KEY%

※ 키 없이 그냥 쓰셔도 됩니다. 그 경우 기존 인쇄설정/규칙 기준으로
  PDF가 생성됩니다.
============================================================
"""


def _is_temp(name: str) -> bool:
    return name.startswith("~$")


def _collect_files(input_path: str, exts: list[str]) -> list[str]:
    exts = [e.lower() for e in exts]
    if os.path.isfile(input_path):
        return [input_path]
    files = []
    for root, _dirs, names in os.walk(input_path):
        # 출력 폴더는 건너뜀
        if os.sep + "_output" in root + os.sep:
            continue
        for n in names:
            if _is_temp(n):
                continue
            if os.path.splitext(n)[1].lower() in exts:
                files.append(os.path.join(root, n))
    return sorted(files)


def _output_path(project_root: str, src: str, out_root: str) -> str:
    rel = os.path.relpath(src, project_root)
    return os.path.join(out_root, rel)


def process_workbook(app, src: str, copy_path: str, pdf_path: str,
                     cfg, patterns, judge) -> dict:
    """워크북 1개 처리. 리포트 dict 반환."""
    rep = {"source": src, "copy": copy_path, "pdf": pdf_path,
           "sheets": [], "included": [], "status": "ok", "error": None}

    os.makedirs(os.path.dirname(copy_path), exist_ok=True)
    shutil.copy2(src, copy_path)  # 원본 보존: 사본에서만 작업

    wb = app.Workbooks.Open(os.path.abspath(copy_path), UpdateLinks=0, ReadOnly=False)
    try:
        for ws in wb.Worksheets:
            entry = {"name": ws.Name}
            excl = printrange.is_sheet_excluded(ws, patterns)
            if excl:
                entry["included"] = False
                entry["reason"] = f"excluded:{excl}"
                rep["sheets"].append(entry)
                continue

            decision = printrange.decide_area(ws, ai_judge=judge)
            if decision["area_obj"] is None:
                entry["included"] = False
                entry["reason"] = decision["reason"]
                rep["sheets"].append(entry)
                continue

            area = decision["area_obj"]
            fixed = autofix.fix_sheet(ws, area, cfg["max_col_width_chars"])
            # AutoFit 으로 열 너비가 바뀌었으므로 영역 객체 재취득(주소 동일)
            area = ws.Range(decision["area_addr"])
            layout = pagelayout.apply(ws, area, cfg)

            entry.update({
                "included": True,
                "reason": decision["reason"],
                "area": decision["area_addr"],
                "confidence": decision["confidence"],
                "widened_cols": fixed["widened_cols"],
                "clamped_cols": fixed["clamped_cols"],
                "layout": layout,
            })
            rep["sheets"].append(entry)
            rep["included"].append(ws.Name)

        wb.Save()
        ok = pdfexport.export_workbook(wb, rep["included"], os.path.abspath(pdf_path))
        rep["status"] = "ok" if ok else "no_pdf"
    finally:
        try:
            wb.Close(SaveChanges=False)
        except Exception:
            pass
    return rep


def run(input_path: str, no_ai: bool, out_dir_arg: str | None) -> int:
    input_path = os.path.abspath(input_path)
    if not os.path.exists(input_path):
        print(f"[오류] 경로를 찾을 수 없습니다: {input_path}")
        return 2

    project_root = input_path if os.path.isdir(input_path) else os.path.dirname(input_path)
    cfg = load_settings(project_root if os.path.isdir(input_path) else None)
    patterns = compile_exclude_patterns(cfg["exclude_sheet_patterns"])

    out_root = out_dir_arg or os.path.join(project_root, cfg["output_dir_name"])
    out_root = os.path.abspath(out_root)

    ai_ok, ai_reason = ai_judge.is_available(no_ai)
    print(f"AI: {'사용 가능' if ai_ok else '미설정 → 규칙 기반으로 진행'} ({ai_reason})")

    files = _collect_files(input_path, cfg["extensions"])
    if not files:
        print("[경고] 처리할 엑셀 파일이 없습니다.")
        return 1
    print(f"대상 워크북: {len(files)}개")
    print(f"출력 위치: {out_root}\n")

    judge = ai_judge.create(cfg, no_ai)
    reports = []
    ok_count = 0

    with ExcelApp() as xl:
        app = xl.app
        for idx, src in enumerate(files, 1):
            name = os.path.basename(src)
            print(f"[{idx}/{len(files)}] {name}")
            dst = _output_path(project_root, src, out_root)
            copy_path = dst  # 사본 엑셀(미러)
            pdf_path = os.path.splitext(dst)[0] + ".pdf"
            try:
                rep = process_workbook(app, src, copy_path, pdf_path,
                                       cfg, patterns, judge)
                if rep["status"] == "ok":
                    ok_count += 1
                    print(f"    -> PDF 생성 ({len(rep['included'])}시트): {pdf_path}")
                else:
                    print(f"    -> 실패({rep['status']})")
                reports.append(rep)
            except Exception as e:
                print(f"    [오류] 건너뜀: {e}")
                reports.append({"source": src, "status": "error",
                                "error": str(e),
                                "trace": traceback.format_exc()})

    # 리포트 기록
    os.makedirs(out_root, exist_ok=True)
    stamp = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    report_path = os.path.join(out_root, f"report_{stamp}.json")
    with open(report_path, "w", encoding="utf-8") as f:
        json.dump({"input": input_path, "ai_available": ai_ok,
                   "ai_reason": ai_reason, "total": len(files),
                   "ok": ok_count, "workbooks": reports},
                  f, ensure_ascii=False, indent=2)

    print(f"\n완료: {ok_count}/{len(files)} 성공")
    print(f"리포트: {report_path}")
    return 0 if ok_count > 0 else 1


def main(argv=None) -> int:
    p = argparse.ArgumentParser(
        prog="수량산출서 출력 자동화 도구",
        description="엑셀 수량산출서를 자동 보정하여 워크북별 1개 PDF로 출력합니다.",
    )
    p.add_argument("input", nargs="?", help="처리할 프로젝트 폴더 또는 엑셀 파일 경로")
    p.add_argument("--file", dest="file", help="특정 엑셀 파일만 처리")
    p.add_argument("--out", dest="out", help="출력 폴더(기본: 입력폴더\\_output)")
    p.add_argument("--no-ai", action="store_true", help="AI 판단 비활성(규칙 기반만)")
    p.add_argument("--help-apikey", action="store_true",
                   help="Google AI Studio API 키 발급 방법 안내")
    args = p.parse_args(argv)

    if args.help_apikey:
        print(APIKEY_HELP)
        return 0

    target = args.file or args.input
    if not target:
        p.print_help()
        return 2

    return run(target, args.no_ai, args.out)


if __name__ == "__main__":
    raise SystemExit(main())
