# Quantity_output

토목 **수량산출서**(Excel) 루트 폴더의 모든 시트를 폴더·파일 순서대로
**간지 포함 단일 PDF**로 병합하고, 셀 표시 오류(`###`, `#REF!`, 수식 오류)와
출력물 이상(빈 페이지, 하단 표 테두리 누락 의심)을 자동 검토하는 도구입니다.

- 사용법·동작 순서·결과물 설명: [tool/README.md](tool/README.md)
- 요구사항: Windows + Microsoft Excel + Python (`pip install -r tool/requirements.txt`)

```bash
python tool/toolruntime.py "수량산출서_루트폴더"      # CLI
tool\run_gui.bat                                      # GUI
```
