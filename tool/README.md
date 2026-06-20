# 수량산출서 출력 자동화 도구

토목 **수량산출서** 엑셀 워크북을 자동으로 보정하여 **워크북별 1개 PDF**로 출력합니다.
여러 프로젝트에 범용으로 적용할 수 있습니다.

## 무엇을 해주나요

- `###`로 잘려 나오던 숫자 셀의 **열 너비를 자동으로 넓혀** 값이 온전히 보이게 합니다.
- 기존 인쇄영역(print_area)·숨김 행/열을 **존중**하고, "출력하지마세요"·중복 시트 등
  불필요한 시트는 **자동 제외**합니다.
- **A4 기준**으로 페이지를 나누되, **열이 페이지 중간에서 잘리지 않게**(세로줄 기준 분할)
  하고, **병합셀이 두 페이지로 쪼개지지 않게**(병합블록을 통째로 다음 페이지로) 합니다.
- A4 폭을 넘으면 **배율을 자동 축소**합니다.
- **원본 파일은 절대 수정하지 않습니다.** `_output` 폴더에 사본과 PDF가 생성됩니다.

## 사용 전 준비

1. **Microsoft Excel** 과 **Python** 이 설치되어 있어야 합니다.
2. 필요한 패키지 설치:
   ```
   pip install -r tool/requirements.txt
   ```
   (최소한 `pywin32` 만 있으면 동작합니다. 나머지는 AI 기능용 선택사항입니다.)

## 사용법

```bash
# 폴더 전체 일괄 처리
python tool/toolruntime.py "01 수량_가야"

# 특정 파일만
python tool/toolruntime.py --file "01 수량_가야/2. 취수틀 및 도수관로/3.0 관부설공.xlsx"

# 출력 폴더 지정
python tool/toolruntime.py "01 수량_가야" --out "D:\\PDF출력"

# AI 끄고 규칙 기반으로만
python tool/toolruntime.py "01 수량_가야" --no-ai

# Google AI Studio API 키 발급 방법 보기
python tool/toolruntime.py --help-apikey
```

윈도우에서는 `tool\run.bat` 을 더블클릭한 뒤 폴더 경로를 입력해도 됩니다.

## AI 기능 (선택사항)

AI는 **있으면 좋은 보조 기능**입니다. 없어도 모든 기능이 정상 동작합니다.
기존 인쇄영역이 없는 애매한 시트의 출력영역 판단 **품질만** 높여줍니다.

쓰려면 `GEMINI_API_KEY` 환경변수에 Google AI Studio 키를 등록하세요.
발급 방법은 `python tool/toolruntime.py --help-apikey` 로 확인할 수 있습니다.

## 결과물

- `_output/<원본과 같은 폴더구조>/<워크북이름>.pdf` — 최종 PDF
- `_output/<...>/<워크북이름>.xlsx` — 조정된 사본 (검토용)
- `_output/report_<날짜시각>.json` — 시트별 포함/제외 사유·조정 내역·오류 로그

## 설정 (선택)

`tool/qto_settings.json` 또는 프로젝트 폴더의 `qto_settings.json` 으로 기본값을 덮어쓸 수 있습니다.
예: 여백, 열 너비 상한, 최소 배율, 제외 시트명 패턴 등. (키 이름은 `config.py` 의 `DEFAULTS` 참고)
