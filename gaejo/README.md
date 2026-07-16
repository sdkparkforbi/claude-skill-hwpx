# gaejo — 관공서 개조식 문서 생성(HTML·PDF·HWPX·DOCX 4형식 일치)

하나의 HTML 마스터에서 **웹·PDF·HWPX(한글)·DOCX(워드)** 네 형식을 **같은 개조식(관공서) 서식**으로 뽑는다.
표지(상하 파란 굵은 선 + 가운데 정렬 + 독립 페이지), 번호 네모칸 섹션 헤더, `❍ / ❐ / - / *` 기호,
명사형 종결, 표·참고박스까지 네 형식에서 동일하게 렌더된다.

hwpx 스킬의 `hwpxgen`(순수 파이썬 XML 생성)과는 **다른 접근**이다:
마스터 HTML → 한글 COM으로 Open/SaveAs → HWPX/DOCX → XML 후처리(개요 스타일·표 여백·쪽 나눔).
서식이 복잡한 개조식 공문을 사람이 HTML로 편하게 쓰고, 네 형식을 한 번에 얻고 싶을 때 쓴다.

## 요구 사항
- Windows + 한컴오피스(한글) — HWPX/DOCX 생성에 필요
- Python: `pip install playwright pywin32 python-docx PyMuPDF` + `python -m playwright install chromium`
  - PDF는 playwright(크로미움)로 마스터를 인쇄 → 웹과 픽셀 동일(한컴 불필요)
  - HWPX/DOCX만 한컴 COM 사용

## 빠른 사용
```python
import sys; sys.path.insert(0, r"~/.claude/skills/hwpx/gaejo")   # 실제 경로로
import build_gaejo
# 마스터 HTML → out_base.pdf / .hwpx / .docx  (세 번째 인자 = 표지 다음에 올 첫 섹션 제목: 표지 독립 페이지)
build_gaejo.build("연구계획서.html", "dist/research-plan", "연구 배경 및 목적")
```
웹(HTML)은 마스터 자체를 그대로 배포한다(상단 nav는 `@media print`로 인쇄 시 숨김).

## 마스터 작성 규칙 (templates/ 참고)
- **섹션 헤더** = `<table class="sec"><tr><td class="num">1</td><td class="tit">제목</td></tr></table>`
- **개조식 문단** = `<p class="o"><span class="mk">❍</span> …</p>` (주항목), `p.dash`(-), `p.star`(*), `p.sq`(❐)
- **소분류** = `<div class="grp">❐ Main …</div>`
- **표** = `<table class="d"><caption>표 1. …</caption> …</table>`
- **참고박스** = `<div class="box"><div class="t">※ …</div> …</div>`
- **강조** = `<b>`, 밑줄강조 `<span class="u">`
- 색·글꼴은 `:root` 변수와 `body{font-family}`만 바꾸면 전체 반영
- 문체는 **명사형 종결**(~함/~됨/~예정임). 공문(요청서)은 정중형(~드립니다) 허용

## 템플릿
- `templates/research_plan.html` — 연구계획서(안): 배경·가설·근거·연구내용·방법·일정·참고문헌
- `templates/data_status.html` — 데이터 현황: 데이터 목록·출처·함정
- `templates/data_request.html` — 자료 협조 요청서: 수신/발신 박스 + 요청 목록 표
- `templates/official_form.html` — 범용 관공서 개조식 공문(빈 양식)

## 함수 구성
| 파일 | 역할 |
|---|---|
| `build_gaejo.py` | 오케스트레이터: 마스터 → PDF+HWPX+DOCX |
| `make_pdf.py` | 마스터를 playwright로 인쇄 → PDF(웹과 동일) |
| `make_gaejo_hwp.py` | 마스터 → 한글 가져오기용 HTML(표지 bar·섹션표·기호·색을 인라인화) |
| `hwp_export.py` | 한글 COM: HTML → HWPX/DOCX + 문단 간격 + 제목 개요 수준 + (옵션)쪽 나눔 |
| `hwpx_style.py` | HWPX XML 후처리: 개요 1/2/3 **스타일 주입**, 표 바깥 여백(`space_tables`), 문단 앞 쪽 나눔(`page_break_before`) |
| `docx_style.py` | DOCX에 Word 제목 스타일(Heading 1/2/3, 흑백 고딕 + 개요 수준) 부여 |
| `make_official_html.py` | (대안) 흑백 관공서 서식 변환기 — 서술형 문서를 단색 표·개요 스타일로 |

## 한글(HWPX) 특유의 함정과 대응 (이 모듈이 처리하는 것)
- 한글은 `<p>` 세로 여백·`text-align(div)`·`::before`·flex를 무시 → 기호는 실제 문자, 정렬은 `align`, 헤더는 표로.
- 표 셀의 흰 글씨·글자 크기를 무시 → 번호·제목을 `<span>`으로 감싸 **문자서식으로 강제**(흰 번호/큰 굵은 제목).
- `<caption>`을 문서 끝으로 밀어냄 → 캡션을 표 앞 굵은 문단으로 빼냄.
- div 배경/테두리 무시 → 박스는 1칸 표로.
- HTML 가져오기 시 개요 스타일을 안 만듦 → `hwpx_style`이 개요 1/2/3 스타일을 XML로 주입.
- 표 근처 API 쪽 나눔은 문서 손상 → `page_break_before`가 표 앵커 문단에 `pageBreak`를 XML로 부여.
- COM이 간헐적으로 멈춤 → 매 실행 전 `Hwp.exe` 종료 + `SetMessageBoxMode`로 대화상자 자동 처리.
