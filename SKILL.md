---
name: hwpx
description: 파이썬으로 한글 문서(.hwpx)를 코드로 생성한다. 표·셀병합·배경색·글자모양·가로/세로 페이지 지원. 사용자가 "hwpx로 출력/저장/변환", "한글 파일/문서로 만들어", "hwpx 만들어줘", "시간표/회의록/표를 hwpx로", "한글(.hwpx) 생성" 등을 요청할 때 트리거. docx→hwpx 변환이 한컴 자동화에서 실패하는 환경을 전제로, 템플릿 기반 section0.xml 직접 생성 + 한글 baking 방식을 쓴다. (Windows + 한컴오피스 설치 필요)
---

# HWPX 자동 생성 스킬

파이썬으로 `.hwpx`(한글) 문서를 **직접 생성**한다. docx→hwpx **변환에 의존하지 않는다**
(한컴 설치본에 따라 자동화 Open이 외부형식에서 실패함). 대신 **템플릿의 header.xml을 재사용**하고
`section0.xml`만 생성→`zipfile` 재패키징→**한글 baking**으로 완성한다.

전체 배경·함정·원리는 동봉 `GUIDE.md` 또는 옵시디언 노트 `content/hwpxgen/` 참조.

## 번들 파일
- `hwpxgen.py` — 생성 엔진(단락/표/병합/색/페이지). **이것을 import해서 쓴다.**
- `hwpx_bake.py` — 생성한 hwpx를 한글로 열고 다시 저장(레이아웃 재계산) + 검증 PDF.
- `make_seed.py` — 기준 시드(_seed.hwpx)를 새로 만들 때(다른 PC/한글 버전).
- `_seed.hwpx` — 기준 템플릿(스타일·색 borderFill·글자 charPr 정의 보유).
- `GUIDE.md` — 상세 가이드(v6).

## 작업 절차

1. **준비물 확인**: Windows + 한컴오피스, 파이썬 `openpyxl`/`pywin32`/`pyhwpx`(보안 DLL용)/`PyMuPDF`(검증).
   `_seed.hwpx`가 없으면 `python make_seed.py`로 생성(한글 자동화).

2. **작업 폴더에 복사**: `hwpxgen.py`, `hwpx_bake.py`, `_seed.hwpx`를 대상 폴더로 복사한 뒤 사용.

3. **문서 생성 스크립트 작성** — `hwpxgen`의 `HwpxDoc`로 본문을 만든다.
   ```python
   from hwpxgen import *
   d = HwpxDoc("_seed.hwpx").page(landscape=False, margin_lr_mm=16)  # 가로면 landscape=True
   navy=d.fill("1F3864"); green=d.fill("CFE9DA")     # 셀 배경색
   red =d.char(CH_SMALL,"C0392C")                    # 글자색 변형
   cw=[int(d.content_w*0.2), d.content_w-int(d.content_w*0.2)]
   rows=[[ d.cell([("항목",CH_WHITE)],0,0,cw[0],bf=navy),
           d.cell([("값", CH_CELL )],1,0,cw[1]) ]]
   body = d._para([("문단",CH_NORMAL)]) + d.table(rows,cw)
   d.save("out.hwpx","제목", body=body)
   ```
   - 스타일 상수: `CH_TITLE/CH_H1/CH_H2/CH_NORMAL/CH_SMALL/CH_WHITE/CH_CELL`, `PA_CENTER/PA_LEFT`, `BF_DATA`.
   - 임의 크기 글자: `d.char_sz(base_id, height_pt, "RRGGBB", bold=True)`.
   - 셀 병합: `d.cell(..., colspan=N, rowspan=M)` + **가려지는 칸은 생략**, `col/row`는 실제 격자 인덱스.

4. **baking + 검증**:
   ```bash
   python hwpx_bake.py out.hwpx     # 한글로 열고 재저장(필수) + _verify_out.pdf 생성
   ```
   그 후 PyMuPDF로 PDF를 이미지 렌더해 눈으로 확인하고, **span size가 0.12pt면 charPr 버그**.

## 반드시 지킬 함정 (자세한 건 GUIDE.md)
- **baking 필수**: 빈 `linesegarray`는 한글이 "열 때" 계산 → 안 하면 **글자가 안 보임**.
- **charPr**: ratio/relSz=100, fontRef 유효값, 색은 `#RRGGBB`(# 필수). 엔진이 시드 로드 시 자동 교정.
- **페이지 방향**: width/height는 항상 세로기준, 방향은 `landscape` 속성(WIDELY=세로, NARROWLY=가로).
- **가로 인쇄폭**: 한글 실제 가로폭 ≈67176 HWPUNIT(≈237mm). 표가 넓으면 오른쪽 열이 잘림 → 너비 실측·축소.
- **재패키징**: 파이썬 `zipfile`만(쉘 zip 금지), 원본 `infolist()` 순서 유지.
- **header.xml은 추가만**(itemCnt 갱신), 기존 항목 수정 금지.

## 엑셀 데이터 → 표 예시
구조화 데이터(예: 시간표 xlsx)는 `openpyxl`로 읽어 격자/트랙을 계산한 뒤,
시작 셀에 `rowspan/colspan`을 주고 가려진 칸을 생략해 병합 표를 만든다.
(원본 프로젝트 `build_timetable_hwpx.py`/`build_minutes_hwpx.py` 패턴 참고.)
