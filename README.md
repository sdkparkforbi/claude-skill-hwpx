# hwpx 스킬 — 설치 & 공유 가이드

파이썬으로 한글 문서(`.hwpx`)를 코드로 **생성·읽기·편집·변환**하는 Claude Code(클로드 코드) **스킬**입니다.
"hwpx로 만들어줘", "hwp 내용 읽어줘", "hwp를 hwpx로 바꿔줘", "양식에 내용 채워줘" 같은 요청에서 자동 활용됩니다.

---

## 받는 사람: 설치 방법

### 1) 사전 준비 (필수)
- **Windows** + **한컴오피스(한글)** 설치
- **Python** + 패키지:
  ```bash
  pip install openpyxl pywin32 pyhwpx PyMuPDF python-docx lxml olefile
  ```
  (`pyhwpx`는 한글 자동화 보안 DLL 제공용으로만 쓰며, 직접 import하지는 않습니다.)
  (`lxml`·`olefile`은 읽기/편집용, `python-docx`는 →docx 변환용입니다.)

> **읽기·편집·일부 변환(→md/txt/docx)은 한컴오피스 없이도 동작**합니다.
> 생성·baking·hwp→hwpx·→pdf 변환에만 한컴오피스가 필요합니다.

### 2) 스킬 설치 (git clone)
스킬 폴더 위치에 바로 내려받습니다.
```bash
# 개인용: 모든 프로젝트에서 사용
git clone https://github.com/sdkparkforbi/claude-skill-hwpx ~/.claude/skills/hwpx
```
- Windows PowerShell이면 `~` 대신 `$HOME` 또는 `C:\Users\<내계정>` 사용.
- **협업 프로젝트**라면 그 저장소의 `<프로젝트>\.claude\skills\hwpx\`에 두면 팀원 모두 자동 사용.
- 설치 후 **Claude Code 재시작**하면 인식됩니다. (`.claude\skills` 폴더가 없으면 만들면 됨)

### 3) 시드(_seed.hwpx) 재생성 — 한글 버전이 다르면 권장
`_seed.hwpx`는 만든 PC의 한글 버전 기준입니다. 글자가 깨지거나 색이 이상하면:
```bash
python make_seed.py        # 내 한글로 새 _seed.hwpx 생성
```

### 4) 확인
새 대화에서 **"hwpx로 표 하나 만들어줘"** 라고 하면 스킬이 동작합니다.

---

## 보내는 사람: 공유 방법

### 권장 — Git 저장소 링크
저장소 주소만 알려주면 됩니다. 받는 사람은 위 `git clone` 한 줄로 설치.
> 협업 프로젝트면 그 저장소의 `.claude/skills/hwpx/`에 커밋 → 팀원 모두 자동 사용.

### 여러 스킬을 함께 배포할 때 — 플러그인 마켓플레이스
스킬을 **플러그인**으로 묶어 깃 마켓플레이스로 배포. 받는 사람은 Claude Code에서:
```
/plugin marketplace add <github-저장소>
/plugin install hwpx
```

> 오프라인 등 git을 못 쓰는 예외 상황에서만, `hwpx` 폴더를 zip으로 압축해 전달하고
> 받는 쪽이 `~/.claude/skills/`에 풀어 넣어도 됩니다(결과 경로 `...\skills\hwpx\SKILL.md`).
> `_seed.hwpx`는 한글 버전이 다르면 받는 쪽에서 `make_seed.py`로 재생성 권장.

---

## 폴더 구성
```
hwpx/
├─ SKILL.md        ← 스킬 정의(트리거 설명) — Claude가 읽음
├─ README.md       ← (이 파일) 설치·공유 가이드
├─ GUIDE.md        ← HWPX 생성 상세 가이드(원리·함정)
├─ hwpxgen.py      ← 생성 엔진 (표·병합·색·이미지·머리말/꼬리말·쪽번호·가로/세로)
├─ hwpx_read.py    ← 읽기 엔진 (.hwp/.hwpx → 본문·표 추출 + 손상 검사 check_refs/check_tables, 한컴 불필요)
├─ hwpx_edit.py    ← 편집 엔진 (양식 치환·셀 채우기·표 꾸미기 + 복구 repair: 참조·표격자 손상 자동수정)
├─ hwpx_convert.py ← 변환 엔진 (hwp→hwpx·→pdf/docx/md/txt)
├─ hwpx_bake.py    ← 한글로 열고 다시 저장(레이아웃 baking) + 검증 PDF
├─ make_seed.py    ← 기준 시드(_seed.hwpx) 재생성
└─ _seed.hwpx      ← 스타일·색·글자모양 템플릿
```

## 기능별 한컴오피스 필요 여부
| 기능 | 모듈 | 한컴 필요 |
|------|------|:--------:|
| 생성 | `hwpxgen.py` (+`hwpx_bake.py`) | ✅ (baking) |
| 읽기·추출 | `hwpx_read.py` | ❌ |
| 교정추적(변경 내용) 분리 | `hwpx_read.py` (`read_changes`) | ❌ (.hwp는 변환만 ✅) |
| 편집(치환·행추가) | `hwpx_edit.py` (+baking 권장) | ❌ (baking만 ✅) |
| 변환 hwp→hwpx, →pdf | `hwpx_convert.py` | ✅ |
| 변환 →docx, →md, →txt | `hwpx_convert.py` | ❌ |
| 복구(안 열리는 hwpx: 참조 무결성·표 격자) | `hwpx_read.check_refs`/`check_tables` 진단 / `hwpx_edit.repair` 수정 | ❌ |

## 한계 / 주의
- **생성·baking·hwp→hwpx·→pdf는 Windows + 한컴오피스 전용** (COM 자동화). 읽기·편집·→docx/md/txt는 한컴 없이 동작.
- 생성·편집한 hwpx는 **baking 권장**(`hwpx_bake.py`) — 안 하면 글자가 안 보일 수 있음.
- 한컴 자동화에는 **docx SaveAs 필터가 없어** →docx는 `python-docx`로 구조 재구성한다(텍스트·표 보존, 정밀 서식 미보존).
- .hwp는 순수 파이썬에서 **본문·표 텍스트**까지 추출(셀 구조가 필요하면 `hwp_to_hwpx` 후 `read_hwpx`).
- **교정추적(변경 내용)**: `read_changes`로 삽입/삭제·변경 전후본 분리. 평문 추출 기본값은 깨끗한 최종본(`extract_text(..., revisions='final')`); 옛 `merge` 동작은 삽입+삭제가 겹쳐 번호가 깨져 보임. COM 변환이 종료마커를 누락하면(`orphans>0`) `accepted_text`(한컴 최종본)를 `final_text=`로 주입해 보정.
- 한글 버전 차이로 문제가 생기면 **시드 재생성**이 1차 해결책.
- 자세한 원리·함정은 `GUIDE.md` 참고.
