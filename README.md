# hwpx 스킬 — 설치 & 공유 가이드

파이썬으로 한글 문서(`.hwpx`)를 코드로 생성하는 Claude Code(클로드 코드) **스킬**입니다.
"hwpx로 만들어줘" 같은 요청에서 자동으로 활용됩니다.

---

## 받는 사람: 설치 방법

### 1) 사전 준비 (필수)
- **Windows** + **한컴오피스(한글)** 설치
- **Python** + 패키지:
  ```bash
  pip install openpyxl pywin32 pyhwpx PyMuPDF python-docx
  ```
  (`pyhwpx`는 한글 자동화 보안 DLL 제공용으로만 쓰며, 직접 import하지는 않습니다.)

### 2) 스킬 폴더 넣기
받은 `hwpx` 폴더를 아래 위치에 통째로 복사합니다.

| 범위 | 위치 |
|---|---|
| **개인용**(모든 프로젝트에서 사용) | `C:\Users\<사용자>\.claude\skills\hwpx\` |
| **프로젝트용**(그 저장소에서만) | `<프로젝트>\.claude\skills\hwpx\` |

복사 후 **Claude Code를 새로 시작**하면 스킬이 인식됩니다.

### 3) 시드(_seed.hwpx) 재생성 — 한글 버전이 다르면 권장
`_seed.hwpx`는 만든 PC의 한글 버전 기준입니다. 글자가 깨지거나 색이 이상하면:
```bash
python make_seed.py        # 내 한글로 새 _seed.hwpx 생성
```

### 4) 확인
새 대화에서 **"hwpx로 표 하나 만들어줘"** 라고 하면 스킬이 동작합니다.

---

## 보내는 사람: 공유 방법 3가지

### 방법 A — 폴더(zip)로 전달 (가장 간단)
1. `~/.claude/skills/hwpx` 폴더를 zip으로 압축해서 전달.
2. 받는 사람은 위 "설치 방법"대로 풀어 넣음.
> ⚠️ `_seed.hwpx`는 동봉해도 되지만, 한글 버전이 다르면 받는 쪽에서 `make_seed.py`로 재생성 권장.

### 방법 B — Git 저장소
1. `hwpx` 폴더를 깃 저장소로 push (예: `github.com/<나>/claude-skill-hwpx`).
2. 받는 사람:
   ```bash
   git clone https://github.com/<나>/claude-skill-hwpx ~/.claude/skills/hwpx
   ```
> 협업 프로젝트라면 그 저장소의 `.claude/skills/hwpx/`에 커밋하면, 팀원 모두 자동 사용.

### 방법 C — 플러그인 마켓플레이스 (여러 스킬을 배포할 때)
스킬을 **플러그인**으로 묶어 깃 기반 마켓플레이스로 배포할 수 있습니다.
받는 사람은 Claude Code에서:
```
/plugin marketplace add <github-저장소>
/plugin install hwpx
```
> 한 번에 여러 스킬·명령을 배포하고 버전 관리하려면 이 방식이 가장 깔끔합니다.

---

## 폴더 구성
```
hwpx/
├─ SKILL.md        ← 스킬 정의(트리거 설명) — Claude가 읽음
├─ README.md       ← (이 파일) 설치·공유 가이드
├─ GUIDE.md        ← HWPX 생성 상세 가이드(원리·함정)
├─ hwpxgen.py      ← 생성 엔진 (표·병합·색·가로/세로)
├─ hwpx_bake.py    ← 한글로 열고 다시 저장(레이아웃 baking) + 검증 PDF
├─ make_seed.py    ← 기준 시드(_seed.hwpx) 재생성
└─ _seed.hwpx      ← 스타일·색·글자모양 템플릿
```

## 한계 / 주의
- **Windows + 한컴오피스 전용** (한글 자동화 COM 사용). macOS/Linux 불가.
- 생성한 hwpx는 **baking 필수**(`hwpx_bake.py`) — 안 하면 글자가 안 보일 수 있음.
- 한글 버전 차이로 문제가 생기면 **시드 재생성**이 1차 해결책.
- 자세한 원리·함정은 `GUIDE.md` 참고.
