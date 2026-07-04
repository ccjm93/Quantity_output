# AGENTS.md

This file provides guidance to Codex (Codex.ai/code) when working with code in this repository.

## What this repository is

This is **not a software codebase** — it is a set of Excel workbooks for a Korean civil engineering
quantity takeoff (수량 산출서) for the "가야" (Gaya) water intake project. The workbooks compute and
tabulate material/earthwork/structure quantities for a water intake structure (취수틀), a raw water
transmission pipeline (도수관로), and a water intake pump station (취수펌프장).

There is no build, lint, or test tooling. "Working" on this repo means reading, auditing, or
programmatically editing `.xlsx` / `.xlsm` / `.xls` files — there is no code to compile or run.

## Directory structure and numbering convention

Everything lives under `01 수량_가야/`. Subfolders follow the project's official table of contents
(see the `목차` sheet in `0. 목차간지/0.0 목차,간지_.xlsx`), and each file is prefixed with the section
number from that TOC:

- `0. 목차간지/` — cover page and table of contents (표지, 목차, 간지/section dividers)
- `1. 공통공사/` — common-work summary tables (총괄주요자재집계표, 공통공사집계표, 품질시험집계표, 폐기물집계표)
- `2. 취수틀 및 도수관로/` — intake structure + transmission pipeline (TOC sections "Ⅰ"):
  `1.0` 주요자재집계표, `2.0` 토공, `3.0` 관부설공, `5.0`/`5.1` 구조물공(취수틀), `6.x` 가시설공,
  `7.0` 추진공 (pipe-jacking/thrust boring), `8.0` 부대공
- `3. 취수펌프장/` — pump station (TOC sections "Ⅱ"): `1.0` 주요자재집계표, `2.x` 토공, `3.0` 구조물공,
  `4.0` 구내배관공, `5.0` 우배수공 (rainwater drainage), `6.0` 가시설공, `7.0` 포장공 (paving),
  `8.x` 부대공, `9.0` 부대시설철거공 (demolition)
- `4. 단위수량/` — unit-quantity reference workbooks shared across sections (manhole covers, gratings,
  handrails, stairs, ladders, pipe clamps, etc.) — TOC section "Ⅲ"

File names embed the same `N.N` numbering as the TOC entry they implement, plus a parenthetical
suffix naming the facility when a calc applies to one structure specifically, e.g.
`2.0 토공,유용,수량집계표(취수펌프장).xlsx`. When adding a new workbook, follow this same
`<section#> <title>(<facility>).<ext>` pattern and slot it into the matching folder/TOC section.

## File format meaning

- `.xlsx` — standard quantity tables/summaries, no macros.
- `.xlsm` — contains VBA macros (verified via `vba_archive` in openpyxl). Used for the more complex
  parametric calculators: `2.2 추진 계산식토공.xlsm`, `7.0 추진공.xlsm`, `8.1 승강용 가설계단(...).xlsm`,
  `4. 단위수량/5. 계단논슬립.xlsm`. Treat macros as load-bearing logic, not boilerplate — don't strip
  them when re-saving with a library that doesn't preserve VBA (use `keep_vba=True` in openpyxl).
- `.xls` — legacy binary format, mostly for gasiseolgong (가시설공, temporary works) and unit-quantity
  calc sheets. openpyxl cannot read these — use `xlrd` (read-only, `.xls` only) or `pandas.read_excel`
  with `engine="xlrd"`, or convert to `.xlsx` first if you need to write changes back.

## Sheet naming conventions (recurring across most workbooks)

- `주요자재(집계표)` — major-material summary/rollup sheet, typically the first sheet, aggregating
  numbers computed on the other sheets in the same file.
- `OOO집계(표)` — quantity summary/rollup for one work item (e.g. `토공집계표`, `구내배관집계표`).
- `OOO산근` / `산출근거` — calculation worksheet showing the derivation ("근거") behind a summary number;
  summary sheets pull from these via cell references, so don't edit a 산근 sheet without checking which
  집계 sheet reads from it.
- `OOO조서` — itemized schedule/log (e.g. `추진조서`, `맨홀조서`, `구내배관조서`) listing per-segment or
  per-item data that summary sheets aggregate.
- Sheets named like `Sheet1`, suffixed `(2)`/`-1`, or containing "출력하지마세요" / "이후시트는
  출력하지마세요" ("do not print this sheet") are either scratch/duplicate sheets or intentionally
  excluded from print output — don't treat them as the canonical source when several similarly-named
  sheets exist in one file.
- Large pipe-jacking workbooks (`7.0 추진공.xlsm`) enumerate per-shaft/per-segment sheets numbered
  `#1`, `#2`, ... — cross-check `추진조서` for which numbered sheet corresponds to which physical
  segment before editing one.

## Working with these files programmatically

Python with `openpyxl` (xlsx/xlsm, formulas via `data_only=False` or computed values via
`data_only=True`) and `pandas` are available in this environment. When writing scripts against these
workbooks:
- Always pass `keep_vba=True` when loading/saving `.xlsm` files to avoid silently dropping macros.
- Many sheets are linked by cell reference across sheets within the same workbook (집계 sheets sum from
  산근/조서 sheets) — `data_only=True` only returns the last-calculated cached value, not a live
  formula recompute, so re-open in Excel (or recalculate) after editing a 산근 sheet if you need the
  집계 totals to reflect the change.
- File/sheet names are in Korean and contain spaces, commas, and parentheses — quote paths and match
  sheet names exactly (including trailing spaces, e.g. `'주요자재집계표 '` in some files).
