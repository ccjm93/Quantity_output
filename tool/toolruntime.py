# -*- coding: utf-8 -*-
"""수량산출서 출력 자동화 도구 - CLI / 오케스트레이션 (단일 PDF 병합판).

흐름: 백업 → 구조 스캔 → 파일별(PBP 활성화+원본저장 → 시트별 임시 PDF)
      → 간지 포함 단일 PDF 병합 → 프로그램 내부 오류 검토 → 산출물 저장.

사용 예:
  python tool/toolruntime.py "01 수량_가야"
  python tool/toolruntime.py "루트" --no-backup       # 원본 백업 생략(주의)
"""
from __future__ import annotations

import argparse
import datetime as _dt
import json
import os
import shutil
import sys
import tempfile
import traceback

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import detect as detect_mod  # noqa: E402
import export as export_mod  # noqa: E402
import prep as prep_mod  # noqa: E402
import review as review_mod  # noqa: E402
import structure as structure_mod  # noqa: E402
from config import XL_OPENXML_WORKBOOK, compile_exclude_patterns, load_settings  # noqa: E402
from excel_app import ExcelApp  # noqa: E402
from merge import Merger  # noqa: E402

REVIEW_XLSX_NAME = "출력물 오류 검토결과.xlsx"


def _normalize_pdf_path(path):
    path = os.path.abspath(path)
    if os.path.splitext(path)[1].lower() != ".pdf":
        path += ".pdf"
    return path


def _assign_pages(cell_findings, manifest):
    """각 (file, sheet) 의 통합 PDF 첫 페이지를 구해 cell finding 에 page 를 부여한다."""
    first_page = {}
    for m in manifest:
        if m.get("type") == "sheet":
            key = (m.get("file"), m.get("sheet"))
            p = m.get("page")
            if key not in first_page or p < first_page[key]:
                first_page[key] = p
    for f in cell_findings:
        f["page"] = first_page.get((f.get("file"), f.get("sheet")))
        f["source"] = "excel"
        if f["kind"] == "ref_error":
            f["types"] = ["#REF! 오류"]
        elif f["kind"] == "error":
            f["types"] = ["수식 오류"]
        else:
            f["types"] = ["# 연속 표시"]
        f["detail"] = f"{f.get('sheet')}!{f.get('cell')} → '{f.get('text')}'"
    return cell_findings


def _original_location(finding):
    file_path = (finding.get("file") or "").replace("\\", "/").strip("/")
    folder = (finding.get("folder") or "").replace("\\", "/").strip("/")
    if file_path:
        return "/" + file_path
    if folder:
        return "/" + folder
    return ""


def _display_location(finding):
    page = finding.get("page")
    return f"최종 PDF {page}페이지" if page else ""


def _korean_error_types(types):
    if isinstance(types, str):
        types = [types]
    out = []
    for t in types or []:
        raw = str(t).strip()
        low = raw.lower()
        if "#REF" in raw.upper() or "ref" in low:
            out.append("#REF! 오류")
        elif "#" in raw or "overflow" in low or "열폭" in raw:
            out.append("# 연속 표시")
        elif "ref" in low or "div" in low or "value" in low or "수식" in raw:
            out.append("수식 오류")
        elif "테두리" in raw:
            out.append("하단 표 테두리 누락 의심")
        elif "잘림" in raw or "crop" in low or "cut" in low:
            out.append("출력 범위 잘림")
        elif "빈" in raw or "blank" in low:
            out.append("빈 페이지 의심")
        elif "깨짐" in raw or "broken" in low:
            out.append("문자 또는 숫자 깨짐")
        elif raw:
            out.append(raw)
    return ", ".join(dict.fromkeys(out))


def save_review_xlsx(review_result, out_dir, cfg, log=print):
    """프로그램 내부 오류 검토 결과를 사용자가 바로 열 수 있는 xlsx 표로 저장한다."""
    os.makedirs(out_dir, exist_ok=True)
    path = os.path.join(out_dir, REVIEW_XLSX_NAME)
    headers = ["검토방식", "출력물 위치", "오류형태", "오류내용", "원본위치"]
    rows = []
    method = "프로그램 내부 로직"
    for fnd in review_result.get("findings", []):
        rows.append([
            method,
            _display_location(fnd),
            _korean_error_types(fnd.get("types", [])),
            str(fnd.get("detail", "") or ""),
            _original_location(fnd),
        ])
    if not rows:
        rows.append([method, "", "이상 없음", "검토 결과 발견된 오류가 없습니다.", ""])

    with ExcelApp() as xl:
        app = xl.app
        wb = app.Workbooks.Add()
        try:
            ws = wb.Worksheets(1)
            ws.Name = "출력물 오류 검토결과"
            while wb.Worksheets.Count > 1:
                wb.Worksheets(wb.Worksheets.Count).Delete()
            for col, header in enumerate(headers, start=1):
                cell = ws.Cells(1, col)
                cell.Value = header
                cell.Font.Bold = True
                cell.Interior.Color = 0xD9EAF7
            for r, row in enumerate(rows, start=2):
                for c, value in enumerate(row, start=1):
                    ws.Cells(r, c).Value = value
            ws.Columns("A:E").EntireColumn.AutoFit()
            try:
                ws.Range("A1:E1").AutoFilter()
                ws.Activate()
                ws.Range("A2").Select()
                app.ActiveWindow.FreezePanes = True
            except Exception:
                pass
            if os.path.exists(path):
                try:
                    os.remove(path)
                except Exception:
                    pass
            wb.SaveAs(path, FileFormat=XL_OPENXML_WORKBOOK)
        finally:
            wb.Close(SaveChanges=False)
    log(f"출력물 오류 검토결과 Excel: {path}")
    return path


def backup_root(root, suffix, log):
    ts = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    dst = root.rstrip("\\/") + f"{suffix}_{ts}"
    log(f"원본 백업 생성 중: {dst}")
    shutil.copytree(root, dst,
                    ignore=shutil.ignore_patterns("_output", "*_backup_*", "~$*"))
    return dst


def run(root, no_backup=False, out_path=None, log=print) -> int:
    root = os.path.abspath(root)
    if not os.path.isdir(root):
        log(f"[오류] 폴더가 아닙니다: {root}")
        return 2

    cfg = load_settings(root)
    patterns = compile_exclude_patterns(cfg["exclude_sheet_patterns"])

    # 출력 위치: 사용자가 '출력 PDF'를 지정하면 그 폴더에 모든 산출물(PDF/manifest/리포트)을 둔다.
    if out_path:
        out_path = _normalize_pdf_path(out_path)
        out_root = os.path.dirname(out_path)
    else:
        out_root = root
        out_path = os.path.join(out_root, cfg.get("output_pdf_name", "수량산출서 output.pdf"))

    log("출력물 오류 검토: 프로그램 내부 로직 사용")

    # [0] 백업
    if cfg.get("backup", True) and not no_backup:
        try:
            backup_root(root, cfg.get("backup_suffix", "_backup"), log)
        except Exception as e:
            log(f"[오류] 백업 실패로 중단(원본 보호): {e}")
            return 3
    else:
        log("원본 백업 생략됨(주의: 원본이 변경됩니다).")

    # [1] 구조 스캔
    events = structure_mod.scan(root, cfg["extensions"],
                                cfg["output_dir_name"], cfg.get("backup_suffix", "_backup"))
    log("구조 파악: " + structure_mod.summarize(events))
    if not any(e["type"] == "file" for e in events):
        log("[경고] 처리할 엑셀 파일이 없습니다.")
        return 1

    merger = Merger(cfg.get("divider_font"))
    tmp_root = tempfile.mkdtemp(prefix="qto_")
    file_no = 0
    n_files = sum(1 for e in events if e["type"] == "file")
    report = {"root": root, "files": [], "errors": []}
    detect_on = cfg.get("detect_cell_issues", True)
    cell_findings = []  # [{file, sheet, cell, kind, text}] — Excel 단계 결정론 탐지

    # [2] COM 세션에서 파일별 처리
    with ExcelApp() as xl:
        app = xl.app
        for ev in events:
            if ev["type"] == "folder":
                merger.add_folder_divider(ev["name"], ev["relpath"], ev["depth"])
                continue
            file_no += 1
            log(f"[{file_no}/{n_files}] {ev['relpath']}")
            try:
                wb = app.Workbooks.Open(os.path.abspath(ev["abspath"]),
                                        UpdateLinks=0, ReadOnly=False)
            except Exception as e:
                log(f"    [오류] 열기 실패, 건너뜀: {e}")
                report["errors"].append({"file": ev["relpath"], "error": str(e)})
                continue
            try:
                pr = prep_mod.prep_workbook(app, wb, patterns,
                                            cfg.get("set_print_area_if_missing", True))
                # 셀 표시 오류(###/수식오류) 결정론적 탐지 (워크북 열린 상태에서)
                if detect_on:
                    for sname in pr["included"]:
                        try:
                            issues = detect_mod.scan_sheet_issues(wb.Worksheets(sname))
                        except Exception as e:
                            log(f"    [경고] 셀 탐지 실패({sname}): {e}")
                            continue
                        for it in issues:
                            it["file"] = ev["relpath"]
                            cell_findings.append(it)
                    n_ov = sum(1 for f in cell_findings
                               if f["file"] == ev["relpath"] and f["kind"] == "overflow")
                    n_ref = sum(1 for f in cell_findings
                                if f["file"] == ev["relpath"] and f["kind"] == "ref_error")
                    n_er = sum(1 for f in cell_findings
                               if f["file"] == ev["relpath"] and f["kind"] == "error")
                    if n_ov or n_ref or n_er:
                        log(f"    [탐지] 연속 # {n_ov}건, #REF! {n_ref}건, 수식오류 {n_er}건")
                tmp_dir = os.path.join(tmp_root, f"f{file_no:04d}")
                sheets = export_mod.export_sheets(wb, pr["included"], tmp_dir)
                if sheets:
                    merger.add_file_divider(ev["name"], ev["relpath"])
                    folder_rel = os.path.dirname(ev["relpath"])
                    for s in sheets:
                        merger.add_sheet(s["pdf"], folder_rel, ev["relpath"], s["sheet"])
                    log(f"    -> {len(sheets)}시트 추가 (제외 {len(pr['excluded'])})")
                else:
                    log(f"    -> 포함 시트 없음(제외 {len(pr['excluded'])})")
                report["files"].append({"file": ev["relpath"],
                                        "included": [s["sheet"] for s in sheets],
                                        "excluded": pr["excluded"]})
            except Exception as e:
                log(f"    [오류] 처리 실패, 건너뜀: {e}")
                report["errors"].append({"file": ev["relpath"],
                                         "error": str(e),
                                         "trace": traceback.format_exc()})
            finally:
                try:
                    wb.Close(SaveChanges=False)
                except Exception:
                    pass

    # [3] 단일 PDF 저장
    if merger.page_count() == 0:
        log("[경고] 생성할 페이지가 없습니다.")
        return 1
    log(f"단일 PDF 저장 중: {out_path} ({merger.page_count()}페이지)")
    manifest = list(merger.manifest)
    merger.save(out_path)

    os.makedirs(out_root, exist_ok=True)
    with open(os.path.join(out_root, "manifest.json"), "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)

    with open(os.path.join(out_root, "process_report.json"), "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)
    try:
        shutil.rmtree(tmp_root, ignore_errors=True)
    except Exception:
        pass

    # [3.5] Excel 단계 셀 표시 오류 — PDF 생성과 함께 항상 보고/저장
    _assign_pages(cell_findings, manifest)
    n_ov = sum(1 for f in cell_findings if f["kind"] == "overflow")
    n_ref = sum(1 for f in cell_findings if f["kind"] == "ref_error")
    n_er = sum(1 for f in cell_findings if f["kind"] == "error")
    log("")
    log("===== PDF 생성 완료 =====")
    log(f"단일 PDF: {out_path}")
    if cell_findings:
        stamp = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        ci_path = os.path.join(out_root, f"cell_issues_{stamp}.json")
        with open(ci_path, "w", encoding="utf-8") as f:
            json.dump({"overflow": n_ov, "ref_error": n_ref, "error": n_er,
                       "findings": cell_findings}, f, ensure_ascii=False, indent=2)
        log(f"셀 표시 오류: 연속 # {n_ov}건, #REF! {n_ref}건, 수식오류 {n_er}건  → {ci_path}")
        for fnd in cell_findings[:15]:
            log(f"  p.{fnd.get('page')} | {fnd.get('file')} > {fnd.get('sheet')}"
                f"!{fnd.get('cell')} | {fnd['types'][0]} '{fnd.get('text')}'")
    else:
        log("셀 표시 오류(연속 #/#REF!/수식오류): 발견 없음")
    log("  ※ 셀 값이 정상이어도 PDF 하단 표 테두리 누락 의심은 출력물 검토 단계에서 확인합니다.")

    # [4] 출력물 오류 검토 — 프로그램 내부 로직으로 수행
    rev = review_mod.review(out_path, manifest, cfg, log=log, extra_findings=cell_findings)
    stamp = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    rev_path = os.path.join(out_root, f"review_{stamp}.json")
    with open(rev_path, "w", encoding="utf-8") as f:
        json.dump(rev, f, ensure_ascii=False, indent=2)
    xlsx_path = save_review_xlsx(rev, out_root, cfg, log=log)
    log(f"이상 의심: {rev['issue_count']}건 ({rev['mode']})  → {rev_path}")
    log(f"검토 결과표: {xlsx_path}")
    for fnd in rev["findings"][:10]:
        log(f"  p.{fnd['page']} | {fnd.get('file')} > {fnd.get('sheet')} "
            f"| {','.join(fnd.get('types', []))} {fnd.get('detail','')}")
    return 0


def run_review_only(target, log=print) -> int:
    """이미 생성된 통합 PDF에 대해 프로그램 내부 오류 검토만 수행."""
    target = os.path.abspath(target)
    if os.path.isdir(target):
        cands = [f for f in os.listdir(target) if f.lower().endswith(".pdf")]
        merged = [f for f in cands if "통합" in f] or cands
        if not merged:
            log(f"[오류] 검토할 PDF를 찾을 수 없습니다: {target}")
            return 2
        pdf = os.path.join(target, sorted(merged)[0])
        out_dir = target
    else:
        pdf = target
        out_dir = os.path.dirname(pdf)
    if not os.path.isfile(pdf):
        log(f"[오류] PDF 없음: {pdf}")
        return 2

    cfg = load_settings(out_dir)
    manifest = []
    mpath = os.path.join(out_dir, "manifest.json")
    if os.path.isfile(mpath):
        try:
            manifest = json.load(open(mpath, encoding="utf-8"))
        except Exception:
            pass
    else:
        log("[안내] manifest.json 이 없어 위치(폴더/파일/시트) 매핑이 제한됩니다.")

    log(f"출력물 오류 검토 대상: {pdf}")
    rev = review_mod.review(pdf, manifest, cfg, log=log)
    stamp = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    rev_path = os.path.join(out_dir, f"review_{stamp}.json")
    with open(rev_path, "w", encoding="utf-8") as f:
        json.dump(rev, f, ensure_ascii=False, indent=2)
    xlsx_path = save_review_xlsx(rev, out_dir, cfg, log=log)
    log("")
    log("===== 검토 완료 =====")
    log(f"이상 의심: {rev['issue_count']}건 ({rev['mode']}) → {rev_path}")
    log(f"검토 결과표: {xlsx_path}")
    for fnd in rev["findings"][:15]:
        log(f"  p.{fnd['page']} | {fnd.get('file')} > {fnd.get('sheet')} "
            f"| {','.join(fnd.get('types', []))} {fnd.get('detail','')}")
    return 0


def main(argv=None) -> int:
    p = argparse.ArgumentParser(
        prog="수량산출서 출력 자동화 도구",
        description="루트 폴더의 모든 시트를 폴더·파일 순서로 단일 PDF(간지 포함)로 병합합니다.")
    p.add_argument("input", nargs="?", help="수량산출서 루트 폴더")
    p.add_argument("--out", help="출력 PDF 경로(기본: 루트\\수량산출서 output.pdf)")
    p.add_argument("--review", help="기존 통합 PDF(또는 출력 폴더)에 대해 오류 검토만 수행")
    p.add_argument("--no-backup", action="store_true", help="원본 백업 생략(주의)")
    args = p.parse_args(argv)

    if args.review:
        return run_review_only(args.review)
    if not args.input:
        p.print_help()
        return 2
    return run(args.input, no_backup=args.no_backup, out_path=args.out)


if __name__ == "__main__":
    raise SystemExit(main())
