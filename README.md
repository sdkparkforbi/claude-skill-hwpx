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
