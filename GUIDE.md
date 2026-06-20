---
title: HWPX 자동 생성 가이드 (v7)
tags:
  - hwpx
  - automation
  - python
  - hancom
---

# HWPX 코드 생성 — 핵심 교훈 & 실전 가이드 (v7)

> **무엇인가** · 파이썬으로 한글 문서(`.hwpx`)를 **프로그래밍 방식으로 생성**하는 범용 가이드.
> 표·셀병합·색상·글자모양·가로/세로 페이지까지 코드로 찍어낸다.
>
> **이력** · v5(시행착오 반영판)에 **2026-06 실전 경험**을 더한 일반화판.
> 추가된 핵심: ① docx→hwpx **변환이 막히는 환경** 대응, ② **시드(템플릿) 역공학**,
> ③ **한글 자동화(COM) 셋업**, ④ **charPr 크기 0 버그**, ⑤ **linesegarray & baking**,
> ⑥ **페이지 방향·가로 인쇄폭 함정**.
>
> 재사용 엔진 코드 전체는 [[#12. 재사용 엔진 hwpxgen.py]] 참조.

---

## 0. 두 갈래: "변환" 대신 "생성"

HWPX를 얻는 길은 두 가지다.

1. **변환** — Word(docx)/RTF를 한글로 열어 hwpx로 저장.
2. **생성** — 파이썬으로 HWPX(zip+xml)를 직접 만든다.

> [!warning] 변환은 환경을 탄다
> 한컴오피스 설치본에 따라 **자동화(COM)에서 외부 형식 열기가 통째로 실패**한다.
> 실측: `hwp.Open(path, "OOXML")`(docx), `"MSWORD"`, `"RTF"` 모두 `False` 반환 ·
> 빈 문서가 저장됨. GUI로 열어도 "빈 문서"가 됐다 → **문서 변환 필터 미설치**로 추정.
> 네이티브 HWP 열기/저장은 정상. 즉 **변환에 의존하지 말 것.**

따라서 이 가이드는 **생성**을 기본으로 한다. 전략 한 줄 요약:

> 기준 **템플릿**의 `header.xml`은 재사용하고, `Contents/section0.xml`만 새로 써서
> `zipfile`로 재패키징한 뒤, **한글로 한 번 열었다 저장(baking)**해 레이아웃을 굳힌다.

---

## 1. HWPX 파일 구조 기초

HWPX는 **ZIP + XML**이다. 파이썬 `zipfile`로 읽고 쓴다.

| 경로 | 내용 |
|---|---|
| `Contents/section0.xml` | 본문 전체 (단락·표·텍스트) ← **여기를 생성/편집** |
| `Contents/header.xml` | 스타일·폰트·테두리(borderFill)·글자모양(charPr) 정의 ← **추가만, 수정 금지** |
| `Preview/PrvImage.png` | 미리보기 이미지 (한글이 저장 시 자동 갱신 = **실제 렌더 확인용**) |
| `Preview/PrvText.txt` | 미리보기 텍스트 (내용 검증용) |
| `mimetype` | 고정. **무압축으로 그대로 복사** |

> [!tip] header.xml은 건드리지 마라 — 단, '추가'는 안전
> 기존 항목을 **수정**하면 스타일·폰트가 깨진다. 그러나 `borderFill`/`charPr`를
> **고유 id로 append + itemCnt 증가**하는 것은 표준적이고 안전하다(본 가이드 방식).

---

## 2. 생성 파이프라인 (6단계)

```
① 시드(템플릿) 확보  →  ② section0.xml 생성  →  ③ header.xml에 색/글자 추가
  →  ④ zipfile 재패키징  →  ⑤ 한글 baking(열고 다시 저장)  →  ⑥ PDF 렌더 검증
```

> [!danger] zip 명령어로 재압축 금지
> 쉘 `zip`으로 다시 묶으면 flag_bits·파일 순서·Central Directory 오프셋이 변형되어
> 한글이 **손상**으로 인식한다. 반드시 **파이썬 `zipfile` + 원본 `infolist()` 순서 유지**.

---

## 3. 한글 자동화(COM) 셋업 — 재사용 핵심

생성 자체는 한글 없이 되지만, **시드 확보**와 **baking·검증**에는 한글 자동화가 필요하다.
다음 셋업이 없으면 줄줄이 막힌다.

### 3-1. 지연 바인딩으로 gencache 회피

`pyhwpx`/`win32com`의 조기 바인딩(gencache)은 아나콘다 설치 경로에 쓰기 권한이 없으면
`PermissionError: ...win32com\gen_py\...`로 죽는다. **지연(late) 바인딩**으로 우회한다.

```python
import win32com.client
hwp = win32com.client.dynamic.Dispatch("HWPFrame.HwpObject")  # gencache 안 씀
```

### 3-2. 보안 모듈 등록 (파일 접근 팝업 차단)

자동화로 파일을 열고 저장하면 한글이 "다른 프로그램이 파일 접근" 팝업을 띄워 멈춘다.
`pyhwpx`가 동봉한 `FilePathCheckerModule.dll`을 레지스트리 **두 곳**에 등록 후 `RegisterModule`.

```python
import winreg
DLL = r"...\site-packages\pyhwpx\FilePathCheckerModule.dll"
for path in (r"Software\HNC\HwpAutomation\Modules",
             r"Software\Hnc\HwpUserAction\Modules"):
    k = winreg.CreateKey(winreg.HKEY_CURRENT_USER, path)
    winreg.SetValueEx(k, "FilePathCheckerModule", 0, winreg.REG_SZ, DLL)
    winreg.CloseKey(k)

hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")  # 1번째 인자 주의!
hwp.SetMessageBoxMode(0xFFFFFF)   # 그 외 대화상자 자동 응답
```

> [!note] `RegisterModule`의 첫 인자는 모듈 **타입** `"FilePathCheckDLL"`, 둘째는 레지스트리 **값 이름** `"FilePathCheckerModule"`. 순서를 헷갈리면 보안창이 그대로 뜬다.

### 3-3. Open 형식 코드

| 형식 | 코드 |
|---|---|
| HWP/HWPX | `"HWP"` / `"HWPX"` |
| MS Word docx | `"OOXML"` (또는 `"MSWORD"`) |
| MS Word doc | `"DOCRTF"` |
| RTF | `"RTF"` |

(단 3절 첫머리 경고대로, 변환 필터가 없으면 이 코드들로도 열기 실패.)

---

## 4. 시드(템플릿) 역공학

기준 HWPX가 없으면 **한글로 시드를 만든 뒤** id를 역공학한다.

1. 한글 자동화로 `FileNew` → 필요한 **글자모양**(CharShape)으로 텍스트 몇 줄,
   **셀 배경**을 입힌 표 한 개를 *알려진 순서로* 넣고 `SaveAs *.hwpx`.
2. 압축을 풀어 매핑:
   - `section0.xml`에서 각 문단의 `charPrIDRef`·각 셀의 `borderFillIDRef`를 **삽입한 순서대로** 읽어 의미를 부여.
   - `header.xml`에서 그 id의 실제 속성(height/textColor/align/fill)을 확인.

예시(본 프로젝트 시드 매핑):

| 용도 | charPr | 용도 | paraPr |
|---|---|---|---|
| 제목 18pt 남색굵게 | 7 | 가운데정렬 | 20 |
| 제목1 14 / 제목2 12(파랑) | 8 / 9 | 왼쪽정렬 | 21 |
| 본문 11 / 작은 9.5(회색) | 10 / 11 | | |
| 흰색 10굵게 / 셀본문 10 | 12 / 13 | 데이터셀(테두리흰) borderFill | 3 |

---

## 5. ⚠️ charPr 함정 — 글자가 0.12pt로 찍힌다 (오늘의 핵심 버그)

**증상** · 표·색은 멀쩡한데 **모든 글자가 거의 안 보임**(PDF span size ≈ 0.12pt).

**원인** · 자동화 `CharShape`로 만든 charPr이 크기 관련 필드를 **0**으로 남긴다:

```xml
<!-- ❌ 깨진 charPr (ratio/relSz=0 → 글자 0%) -->
<hh:charPr id="7" height="1800" textColor="#1F3864" ...>
  <hh:fontRef hangul="0" .../>   <!-- 폰트 미지정 -->
  <hh:ratio   hangul="0" .../>   <!-- 장평 0% -->
  <hh:relSz   hangul="0" .../>   <!-- 상대크기 0% → 사실상 0pt -->
```

**해결** · 로드 시 `ratio`/`relSz`를 **100**, `fontRef`를 유효 폰트(예 `2`)로 교정.

```python
zero = 'hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"'
full = 'hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"'
header = header.replace(f'<hh:ratio {zero}/>', f'<hh:ratio {full}/>')
header = header.replace(f'<hh:relSz {zero}/>', f'<hh:relSz {full}/>')
header = header.replace(f'<hh:fontRef {zero}/>',
        '<hh:fontRef hangul="2" latin="2" hanja="2" japanese="2" other="2" symbol="2" user="2"/>')
```

> 색 변경은 `textColor="#RRGGBB"` — **`#` 필수**(빠지면 엉뚱한 색/청록으로 렌더).

---

## 6. borderFill 추가 (셀 배경색)

데이터셀(테두리 있는 흰 셀) borderFill을 복사해 `fillBrush`만 추가하고 고유 id 부여,
`<hh:borderFills itemCnt="N">`의 N을 증가시킨다.

```xml
<hh:borderFill id="{새 id}" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0">
  <hh:slash type="NONE".../><hh:backSlash type="NONE".../>
  <hh:leftBorder type="SOLID" width="0.12 mm" color="#000000"/>
  <hh:rightBorder .../><hh:topBorder .../><hh:bottomBorder .../>
  <hh:diagonal type="SOLID" width="0.1 mm" color="#000000"/>
  <hc:fillBrush><hc:winBrush faceColor="#CFE9DA" hatchColor="#000000" alpha="0"/></hc:fillBrush>
</hh:borderFill>
```

> `faceColor`에도 **`#` 필수**. 여러 색을 쓰면 색→id 매핑을 캐시해 중복 생성을 막는다.

---

## 7. ⚠️ linesegarray & baking — 텍스트가 안 보이는 진짜 이유

`<hp:linesegarray>`는 한글이 줄 배치 결과를 **캐시**해 두는 영역이다.

- 직접 `horzsize`를 박으면 여러 줄 본문이 한 줄로 **겹쳐** 검은 블록처럼 렌더된다.
- 그래서 생성 시에는 **빈 태그 `<hp:linesegarray/>`** 로 둔다(원칙).

> [!danger] 빈 linesegarray는 "한글이 열 때" 계산된다 — 헤드리스 PDF로는 안 보인다
> 빈 채로 두고 **자동화 Open → SaveAs PDF**를 하면 줄 배치가 비어 **텍스트가 통째로 안 나온다**.
> (모델·`PrvText`엔 글자가 있는데 화면엔 없음.)
>
> **해결 = baking**: 생성한 hwpx를 한글로 **열어 다시 `SaveAs *.hwpx`** 하면
> 한글이 `linesegarray`와 **행 높이를 재계산**해 채워 넣는다. 이 baked 파일이 최종본.

```python
hwp.Open(path, "HWPX", "")
hwp.SaveAs(path, "HWPX", "")   # ← baking: lineseg/행높이 재계산
hwp.SaveAs(pdf,  "PDF",  "")   # 검증용
```

---

## 8. ⚠️ 페이지 방향 & 가로 인쇄폭

> [!warning] `landscape` 속성 의미가 직관과 반대다
> 용지 `width`/`height`는 **항상 세로 기준**(210×297mm = `59528`×`84189` HWPUNIT).
> 방향은 `landscape` 속성으로만 결정:
> - **`landscape="WIDELY"` = 세로(portrait)**
> - **`landscape="NARROWLY"` = 가로(landscape)**
>
> 즉 가로 문서도 width/height는 세로값 그대로 두고 속성만 `NARROWLY`.

> [!warning] 가로 인쇄폭은 생각보다 좁다
> 가로에서 한글의 **실제 본문 가로폭**이 예상(297−여백)보다 작다(실측 ≈ `67176` HWPUNIT ≈ 237mm).
> 표가 그보다 넓으면 **오른쪽 열이 잘린다**(페이지 밖으로 사라짐).
> → 한글이 만든 가로 문서의 `lineseg horzsize`로 **실제 폭을 실측**해 표 너비를 맞춰라.

단위: `1mm ≈ 283.465 HWPUNIT`, `글자 height = pt × 100`.

---

## 9. 셀 병합 — cellSpan / cellAddr

HTML 표 모델과 같다. 병합 셀은 **시작 셀 하나만** 두고 `colSpan`/`rowSpan`을 주며,
**가려지는 칸의 `<hp:tc>`는 생략**한다.

```xml
<hp:tc ... borderFillIDRef="6">
  <hp:subList ...> ...단락... </hp:subList>
  <hp:cellAddr colAddr="2" rowAddr="1"/>   <!-- 실제 (열,행) 인덱스 -->
  <hp:cellSpan colSpan="1" rowSpan="3"/>   <!-- 세로 3칸 병합 -->
  <hp:cellSz   width="5538" height="..."/>
  <hp:cellMargin left="340" right="340" top="100" bottom="100"/>
</hp:tc>
```

- `cellAddr`는 **실제 격자 위치**(전부 0,0 금지 → 손상).
- 병합 셀의 `width`는 **걸친 열 너비의 합**, `height`는 걸친 행 높이의 합.

---

## 10. 표 구조 원칙 (손상 방지 체크)

| 항목 | 규칙 |
|---|---|
| 인라인 삽입 | `<hp:pos treatAsChar="1" ...>` (0이면 floating → 화면에서 사라짐) |
| 표 위치 | `horzRelTo="COLUMN"` (PARA 아님) |
| 표 래핑 | 표는 `hp:p > hp:run` 안에. **표 뒤 `<hp:t/>` 필수**(run 닫힘) |
| 행 수 | `rowCnt`/`colCnt`를 실제와 일치(행 추가 시 반드시 갱신) |
| 열너비 합 | `sum(colWidths) == hp:sz width`(표 너비) |
| id/zOrder | 원본 최댓값+순번으로 **고유** 부여(중복 시 손상) |
| linesegarray | 모두 **빈 태그** + baking |

---

## 11. 검증 파이프라인

```python
import fitz  # PyMuPDF
p = fitz.open("_verify.pdf")[0]
print(p.rect.width, p.rect.height)          # 방향 확인 (가로면 width>height)
sizes = {round(s['size'],1)
         for b in p.get_text('dict')['blocks']
         for l in b.get('lines',[]) for s in l['spans']}
print(sizes)                                # {9.5,10,11,18} 정상 / {0.12} = charPr 버그
p.get_pixmap(dpi=130).save("_check.png")    # 눈으로 확인
```

- `span size`로 **글자 크기 버그(0.12pt)** 를 잡는다.
- hwpx 안의 `Preview/PrvImage.png`(한글 자체 렌더)·`PrvText.txt`도 교차 검증에 유용.

---

## 12. 재사용 엔진 `hwpxgen.py`

위 교훈을 모두 담은 범용 엔진. 시드 1개만 있으면 단락·표·병합·색·가로/세로 문서를 찍어낸다.

```python
# 사용 예
from hwpxgen import *
d = HwpxDoc("_seed.hwpx").page(landscape=False, margin_lr_mm=16)
navy  = d.fill("1F3864")          # 셀 배경색 → borderFill id
green = d.fill("CFE9DA")
red   = d.char(CH_SMALL, "C0392C")  # 글자색 변형 charPr
cw = [int(d.content_w*0.2), d.content_w-int(d.content_w*0.2)]
rows = [
  [ d.cell([("항목", CH_WHITE)], 0,0, cw[0], bf=navy),
    d.cell([("값",   CH_CELL )], 1,0, cw[1]) ],
]
body = d._para([("문단 텍스트", CH_NORMAL)]) + d.table(rows, cw)
d.save("out.hwpx", "문서 제목", body=body)
# 이후 한글로 baking: Open → SaveAs(hwpx) → (선택) SaveAs(pdf)
```

<details>
<summary>hwpxgen.py 전체 코드</summary>

```python
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HWPX 생성 엔진 (범용).
- 기준 템플릿(seed) HWPX의 header.xml/section0.xml을 역공학한 결과를 바탕으로,
  파이썬 자료구조 → section0.xml 본문(단락/표/셀병합)을 생성하고,
  header.xml에는 필요한 borderFill(셀 배경색)·charPr(글자색)만 '추가'한 뒤,
  zipfile 모듈로 원본 순서를 유지하며 재패키징한다.
"""
import re, html, zipfile, os

HWPU_PER_MM = 283.465
def mm(v): return int(round(v*HWPU_PER_MM))
def esc(s): return html.escape(str(s), quote=False)

# 시드에서 확보한 스타일 ID (header.xml) — 자신의 시드에 맞게 교체
CH_TITLE='7'; CH_H1='8'; CH_H2='9'; CH_NORMAL='10'; CH_SMALL='11'; CH_WHITE='12'; CH_CELL='13'
PA_CENTER='20'; PA_LEFT='21'
BF_DATA='3'  # 테두리 있는 흰 셀(데이터)

class HwpxDoc:
    def __init__(self, template='_seed.hwpx'):
        self.template = template
        with zipfile.ZipFile(template) as z:
            self.header = z.read('Contents/header.xml').decode('utf-8')
            self.section = z.read('Contents/section0.xml').decode('utf-8')
        # 시드 CharShape 자동화가 남긴 깨진 charPr 교정 (ratio/relSz=0 → 100, fontRef=2)
        zero='hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"'
        full='hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"'
        self.header = self.header.replace(f'<hh:ratio {zero}/>', f'<hh:ratio {full}/>')
        self.header = self.header.replace(f'<hh:relSz {zero}/>', f'<hh:relSz {full}/>')
        self.header = self.header.replace(f'<hh:fontRef {zero}/>',
            '<hh:fontRef hangul="2" latin="2" hanja="2" japanese="2" other="2" symbol="2" user="2"/>')
        self._next_bf = max(int(x) for x in re.findall(r'<hh:borderFill id="(\d+)"', self.header))+1
        self._next_ch = max(int(x) for x in re.findall(r'<hh:charPr id="(\d+)"', self.header))+1
        ids=[int(x) for x in re.findall(r'<hp:tbl[^>]*\bid="(\d+)"', self.section)]
        zs =[int(x) for x in re.findall(r'<hp:tbl[^>]*zOrder="(\d+)"', self.section)]
        self._tbl_id=(max(ids) if ids else 0)+1000
        self._z=(max(zs) if zs else 0)+200
        self._added_bf=[]; self._added_ch=[]; self._color_bf={}; self._color_ch={}

    def fill(self, hexcolor):
        hexcolor=hexcolor.upper()
        if not hexcolor.startswith('#'): hexcolor='#'+hexcolor
        if hexcolor in self._color_bf: return self._color_bf[hexcolor]
        bid=self._next_bf; self._next_bf+=1
        xml=(f'<hh:borderFill id="{bid}" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0">'
             '<hh:slash type="NONE" Crooked="0" isCounter="0"/><hh:backSlash type="NONE" Crooked="0" isCounter="0"/>'
             '<hh:leftBorder type="SOLID" width="0.12 mm" color="#000000"/>'
             '<hh:rightBorder type="SOLID" width="0.12 mm" color="#000000"/>'
             '<hh:topBorder type="SOLID" width="0.12 mm" color="#000000"/>'
             '<hh:bottomBorder type="SOLID" width="0.12 mm" color="#000000"/>'
             '<hh:diagonal type="SOLID" width="0.1 mm" color="#000000"/>'
             f'<hc:fillBrush><hc:winBrush faceColor="{hexcolor}" hatchColor="#000000" alpha="0"/></hc:fillBrush>'
             '</hh:borderFill>')
        self._added_bf.append(xml); self._color_bf[hexcolor]=str(bid); return str(bid)

    def char(self, base_id, hexcolor, bold=None):
        return self._char(base_id, None, hexcolor, bold)
    def char_sz(self, base_id, height_pt, hexcolor, bold=None):
        return self._char(base_id, height_pt, hexcolor, bold)
    def _char(self, base_id, height_pt, hexcolor, bold):
        hexcolor=hexcolor.upper()
        if not hexcolor.startswith('#'): hexcolor='#'+hexcolor
        h=int(round(height_pt*100)) if height_pt else None
        key=(base_id,h,hexcolor,bold)
        if key in self._color_ch: return self._color_ch[key]
        base=re.search(r'<hh:charPr id="'+str(base_id)+r'".*?</hh:charPr>', self.header, re.DOTALL).group()
        cid=self._next_ch; self._next_ch+=1
        new=re.sub(r'id="'+str(base_id)+r'"', f'id="{cid}"', base, count=1)
        if h: new=re.sub(r'height="\d+"', f'height="{h}"', new, count=1)
        new=re.sub(r'textColor="[^"]*"', f'textColor="{hexcolor}"', new, count=1)
        if bold is True and '<hh:bold' not in new: new=new.replace('</hh:charPr>','<hh:bold/></hh:charPr>')
        if bold is False: new=new.replace('<hh:bold/>','')
        self._added_ch.append(new); self._color_ch[key]=str(cid); return str(cid)

    def page(self, landscape=False, margin_lr_mm=20, margin_tb_mm=15):
        W=mm(210); H=mm(297); lr=mm(margin_lr_mm); tb=mm(margin_tb_mm); hf=mm(10)
        self.content_w=(H if landscape else W)-2*lr
        attr="NARROWLY" if landscape else "WIDELY"  # NARROWLY=가로, WIDELY=세로
        self._pagepr=(f'<hp:pagePr landscape="{attr}" width="{W}" height="{H}" gutterType="LEFT_ONLY">'
            f'<hp:margin header="{hf}" footer="{hf}" gutter="0" left="{lr}" right="{lr}" top="{tb}" bottom="{tb}"/></hp:pagePr>')
        return self

    def _para(self, runs, paraPr=PA_LEFT):
        rxml=''.join(f'<hp:run charPrIDRef="{cp}"><hp:t>{esc(t)}</hp:t></hp:run>' for t,cp in runs)
        if not rxml: rxml=f'<hp:run charPrIDRef="{CH_NORMAL}"><hp:t></hp:t></hp:run>'
        return (f'<hp:p id="0" paraPrIDRef="{paraPr}" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">'
                f'{rxml}<hp:linesegarray/></hp:p>')

    def cell(self, lines, col, row, width, height=900, bf=BF_DATA, paraPr=PA_CENTER, colspan=1, rowspan=1):
        if isinstance(lines,str): lines=[(lines,CH_CELL)]
        inner=''.join(
            f'<hp:p id="0" paraPrIDRef="{paraPr}" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">'
            f'<hp:run charPrIDRef="{cp}"><hp:t>{esc(t)}</hp:t></hp:run><hp:linesegarray/></hp:p>'
            for t,cp in lines)
        return (f'<hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="{bf}">'
                f'<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" '
                f'linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">'
                f'{inner}</hp:subList>'
                f'<hp:cellAddr colAddr="{col}" rowAddr="{row}"/>'
                f'<hp:cellSpan colSpan="{colspan}" rowSpan="{rowspan}"/>'
                f'<hp:cellSz width="{width}" height="{height}"/>'
                f'<hp:cellMargin left="340" right="340" top="100" bottom="100"/></hp:tc>')

    def table(self, rows, col_widths):
        rc=len(rows); cc=len(col_widths); tw=sum(col_widths)
        tid=self._tbl_id; self._tbl_id+=1; z=self._z; self._z+=1
        trs=''.join('<hp:tr>'+''.join(r)+'</hp:tr>' for r in rows)
        tbl=(f'<hp:tbl id="{tid}" zOrder="{z}" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" '
             f'textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="NONE" repeatHeader="1" '
             f'rowCnt="{rc}" colCnt="{cc}" cellSpacing="0" borderFillIDRef="{BF_DATA}" noAdjust="0">'
             f'<hp:sz width="{tw}" widthRelTo="ABSOLUTE" height="1282" heightRelTo="ABSOLUTE" protect="0"/>'
             f'<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" '
             f'vertRelTo="PARA" horzRelTo="COLUMN" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/>'
             f'<hp:outMargin left="0" right="0" top="0" bottom="0"/>'
             f'<hp:inMargin left="340" right="340" top="100" bottom="100"/>{trs}</hp:tbl>')
        return (f'<hp:p id="0" paraPrIDRef="{PA_LEFT}" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">'
                f'<hp:run charPrIDRef="0">{tbl}<hp:t/></hp:run><hp:linesegarray/></hp:p>')

    def save(self, out_path, title, title_charPr=CH_TITLE, body=''):
        header=self.header
        if self._added_bf:
            m=re.search(r'<hh:borderFills itemCnt="(\d+)">', header)
            header=header[:m.start()]+f'<hh:borderFills itemCnt="{int(m.group(1))+len(self._added_bf)}">'+header[m.end():]
            header=header.replace('</hh:borderFills>', ''.join(self._added_bf)+'</hh:borderFills>',1)
        if self._added_ch:
            m=re.search(r'<hh:charProperties itemCnt="(\d+)">', header)
            header=header[:m.start()]+f'<hh:charProperties itemCnt="{int(m.group(1))+len(self._added_ch)}">'+header[m.end():]
            header=header.replace('</hh:charProperties>', ''.join(self._added_ch)+'</hh:charProperties>',1)
        head=self.section[: self.section.find('</hp:ctrl></hp:run>')+len('</hp:ctrl></hp:run>')]
        head=re.sub(r'<hp:pagePr\b.*?</hp:pagePr>', self._pagepr, head, count=1, flags=re.DOTALL)
        title_run=f'<hp:run charPrIDRef="{title_charPr}"><hp:t>{esc(title)}</hp:t></hp:run><hp:linesegarray/></hp:p>'
        section=head+title_run+body+'</hs:sec>'
        with zipfile.ZipFile(self.template) as zin, zipfile.ZipFile(out_path,'w',zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename=='Contents/header.xml': zout.writestr(item, header.encode('utf-8'))
                elif item.filename=='Contents/section0.xml': zout.writestr(item, section.encode('utf-8'))
                else: zout.writestr(item, zin.read(item.filename))
        return out_path
```

</details>

---

## 13. 최종 체크리스트

- [ ] HWPX = ZIP+XML. **파이썬 zipfile**로만 읽고 쓴다(쉘 zip 금지, 원본 순서 유지).
- [ ] `section0.xml`만 새로 쓰고 `header.xml`은 **추가만**(itemCnt 갱신).
- [ ] charPr **ratio/relSz=100, fontRef 유효값**, 색은 `#RRGGBB`.
- [ ] borderFill `faceColor="#..."`, id·itemCnt 정확.
- [ ] 표: `treatAsChar="1"`, `horzRelTo="COLUMN"`, 표 뒤 `<hp:t/>`, `rowCnt`/열너비합 일치.
- [ ] 병합: 가려진 tc 생략, `cellAddr` 실제값, span 셀 width/height=합.
- [ ] **참조 무결성**: 본문(section)이 쓰는 `paraPrIDRef/charPrIDRef/styleIDRef/borderFillIDRef`가 머리부(header) 정의 안에 모두 존재(`hwpx_read.check_refs`가 빈 dict).
- [ ] **표 격자**: 모든 셀에 `<hp:cellAddr>` 존재, 행별 열너비 일관, sz.height=행높이 합(`hwpx_read.check_tables`가 빈 list). 깨졌으면 `hwpx_edit.repair`.
- [ ] 모든 `linesegarray` 빈 태그 → **한글 baking(열고 다시 저장)** 으로 굳히기.
- [ ] 페이지: width/height는 세로기준 고정, 방향은 `landscape`(WIDELY=세로/NARROWLY=가로).
- [ ] 가로 인쇄폭 **실측**(≈67176)에 표 너비 맞추기.
- [ ] PDF 렌더 + **span size**로 글자크기(0.12pt 버그) 검증.

---

## 14. ⚠️ 교정추적(변경 내용) 읽기 — 삽입/삭제 분리

심사·검토용 .hwp는 **변경 내용 추적(track changes)** 이 켜진 채 온다. 이걸 모르고 평문만 뽑으면
삽입·삭제 텍스트가 **한 흐름으로 섞여** 나온다(예: 참고문헌 재번호 중 옛 `[7]`이 지워지고 새 `[34]`가
들어가면 `"[347]"`로 보임 — 원고 오류가 아니라 추적 흔적).

### 14-1. HWPX의 저장 구조
- `header.xml` : `<hh:trackChange type="Insert|Delete|CharShape|ParaShape" date=.. authorID=.. id=../>`,
  작성자 `<hh:trackChangeAuthor name=.. id=../>`. (`id`가 본문 마커의 `TcId`)
- `section*.xml` : 본문 인라인 마커 `<hp:insertBegin Id=.. TcId=../>…<hp:insertEnd …/>`,
  `<hp:deleteBegin …/>…<hp:deleteEnd …/>`. **Begin~End 사이의 `hp:t`** 가 그 구간의 삽입/삭제 텍스트.

### 14-2. 함정 — COM 변환이 종료마커를 누락
`hwp→hwpx`(한컴 COM) 변환이 **종료 마커를 1개씩 빠뜨리는** 사례가 실측됨
(`insertBegin` 135 / `insertEnd` 134, `deleteBegin` 57 / `deleteEnd` 56). 종료 없는 구간(**orphan**)을
그대로 두면 그 뒤 본문 전체가 삭제로 오분류되어 **"5-2~참고문헌이 통째로 사라짐"** 같은 파국이 난다.
- `read_changes`는 정상적으로 닫힌 구간만 신뢰하고, orphan은 기본적으로 **문단 끝에서 강제 종료(cap)** 해
  피해를 1문단으로 한정하며 `orphans` 개수를 보고한다.
- **권위 있는 보정**: 한컴 `GetTextFile('TEXT')` 은 보기 모드와 무관하게 **늘 '변경 적용된 최종본'** 을
  돌려준다(삭제는 빠지고 삽입은 포함). 자동화의 변경 **수락/거부 액션(RevisionApplyAll 등)은 no-op**.
  → `hwpx_convert.accepted_text(path)`로 최종본 평문을 얻어 `read_changes(path, final_text=...)`에
  주입하면, 최종본에 남은 텍스트의 **거짓 삭제를 걷어내고** orphan 삽입 구간까지 복원한다.
  (편집 전 '원본/거부본'은 한컴으로 못 만들므로 `read_changes(...)['original']`로만 얻는다.)

### 14-3. 사용
```python
import hwpx_read, hwpx_convert
# .hwp면 먼저 변환
hwpx_convert.hwp_to_hwpx("심사본.hwp", "심사본.hwpx")
# (1) 깨끗한 최종 본문만:
clean = hwpx_read.extract_text("심사본.hwpx")            # revisions='final'(기본)
# (2) 무엇이 바뀌었나:
ft = hwpx_convert.accepted_text("심사본.hwp")            # 권위 최종본(오라클)
ch = hwpx_read.read_changes("심사본.hwpx", final_text=ft)
for c in ch["changes"]:
    print(c["type"], "p%s" % c["para"], c["author"], c["text"][:80])
assert ch["final"] == ft                                 # 검증
```
CLI: `python hwpx_read.py 심사본.hwpx --changes`

---

## 15. 관공서 양식(.hwp) 채우기 — v8 실전 함정

받은 **공공기관 신청서/수요조사서 양식**을 `hwp_to_hwpx`로 올린 뒤 셀을 채울 때
두 가지 비자명한 함정이 있다(2026-06 RISE 지산학협력 수요조사서 작성 실전).

### 15-1. '빈' 셀에는 `<hp:t>`가 아예 없다
양식의 입력 칸은 비어 보여도 `<hp:run charPrIDRef="3">`까지는 존재하지만 그 안에
**텍스트 노드 `<hp:t>`가 없다**. 옛 `set_cell`은 `_own_ts(p)`가 `[]`이면 조용히
아무것도 안 써서 **글자가 안 들어간다**(에러도 안 남 → 더 위험).
→ v8에서 `set_cell`이 t가 없으면 **`<hp:t>`를 생성**하도록 수정. run조차 없으면
   `_default_charpr()`(문서 최다 사용 charPr)로 run까지 만든다.

진단 팁: 채운 뒤 반드시 `hwpx_read.iter_tables()`로 되읽어 값이 실제로 들어갔는지 확인.

### 15-2. 체크박스는 Wingdings 글리프다
한글 양식의 ☐/☑는 별도 컨트롤이 아니라 **라벨과 같은 run 텍스트 안의 글자**다.
- `` = 빈 칸(☐), `` = 체크(☑) — Wingdings 사상 코드(charPr는 보통 일반 폰트값)
- 예: `" AI, DX"` 한 run에 글리프+공백+라벨이 함께 들어있음.
- 토글 = 글리프 치환. `check_option(tbl, row, col, checked=True)`가
  해당 셀에서 ``↔``를 바꿔준다.
- 선택지가 양식에 없을 땐(예: '공동연구형'이 칸에 없음) 기존 항목을 빈칸으로 돌리고
  `set_cell`로 `" 기술이전형     공동연구형"`처럼 직접 써넣는다.

> ⚠️ **함정(실전 추가): 한컴 baking이 Wingdings 글리프를 지운다.**
> `check_option`으로 글리프를 토글한 뒤 `hwpx_bake.py`를 돌리면 **일부 셀에서
> 글리프가 통째로 사라진다**(같은 문서에서 어떤 셀은 살고 어떤 셀은 죽는 비결정적 동작 관측).
> private-use(U+F0xx) 문자라 한글이 재저장 시 폰트 사상에 따라 떨어뜨리는 것으로 보인다.
> → **가장 견고한 방법: 유니코드 박스 `☐`(U+2610)·`☑`(U+2611)를 `set_cell`로 셀 전체에
>    다시 써넣기.** 일반 폰트 문자라 baking·`to_docx`·HTML/PDF 변환에서 모두 안전하다.
> ```python
> ed.set_cell(0, row=2, col=6, text="☑ AI, DX")             # 분야 체크
> ed.set_cell(0, row=3, col=1, text="☐ 기술이전형   ☑ 공동연구형")
> ```
> `check_option`(글리프 토글)은 **baking 안 하는** 경우엔 원본 스타일을 보존하므로 유효하지만,
> baking이 필요하면 유니코드 `set_cell` 방식을 쓸 것.

### 15-3. 셀 주소(cellAddr)는 병합 반영 좌표다
`set_cell`/`check_option`의 `row,col`은 **시각적 격자가 아니라 `<hp:cellAddr>`의
colAddr/rowAddr**다. 병합(`cellSpan`)이 있으면 인덱스를 건너뛴다. 먼저 표를 덤프해
각 `tc`의 `cellAddr`(col,row)·`cellSpan`을 출력해 매핑한 뒤 채울 것.

### 15-4. 긴 텍스트 줄간격 붕괴 — stale `linesegarray`
'빈 칸'은 1줄짜리 `<hp:linesegarray>`(lineseg 1개)를 갖고 있다. 여기에 여러 줄로
wrap될 긴 본문을 넣으면, 한글은 **남아있는 lineseg 1개에 모든 줄을 겹쳐 그려서**
글자가 한 줄에 뭉개진다(=줄간격 붕괴). baking은 `linesegarray`가 **비어 있을 때만**
줄 수를 다시 계산하므로, stale lineseg가 남아 있으면 재계산하지 않는다.
→ v8 `set_cell`은 텍스트 교체 시 해당 단락들의 `<hp:lineseg>`를 모두 제거한다.
   이후 `hwpx_bake.py`가 올바른 줄 수를 다시 계산한다(baking 필수).
   (직접 XML을 만질 때도 긴 셀은 lineseg를 비우고 bake할 것.)

```python
import hwpx_edit
ed = hwpx_edit.HwpxEditor("양식.hwpx")
ed.set_cell(0, row=1, col=1, text="AI·DX 기반 한우 수요예측 시스템 개발")  # 빈 칸도 채움
ed.check_option(0, row=2, col=6, checked=True)     # 분야: AI,DX 체크
ed.save("작성본.hwpx")   # → python hwpx_bake.py 작성본.hwpx (레이아웃 재계산 필수)
```

---

## 16. ⚠️ "열 수 없는 문서" / 열 때 멈춤 — 참조·표격자 손상 (check_refs · check_tables / repair)

### 16-1. 증상과 원인
외부 도구/스크립트로 **조립·병합**된 hwpx가 한글에서 *"파일을 열 수 없습니다 / 손상된 문서"* 로
거부되는데, 막상 뜯어보면 **XML은 well-formed, 표 구조도 정상, mimetype도 정상**인 경우가 있다.
범인은 대개 **참조 무결성** — 본문 `section*.xml`이 머리부 `header.xml`에 **정의되지 않은
스타일/속성 ID**를 가리키는 것:

```
section0.xml : ... paraPrIDRef="21" ...      ← 21번 문단모양을 사용
header.xml   : <hh:paraProperties itemCnt="21"> ... <hh:paraPr id="20"> </hh:paraProperties>
               → 문단모양은 0~20번만 존재. 21번이 없음 → dangling 참조 → 한글이 열기 거부.
```

표 병합 시 다른 문서의 본문만 가져오고 그 문서가 쓰던 `paraPr/charPr/borderFill/style` 정의를
머리부에 함께 옮기지 않으면 이렇게 된다. (이 검사는 `hwpx_read.validate()`에 통합되어 함께 잡힌다.)

### 16-2. 검사·복구 (진단=hwpx_read, 수정=hwpx_edit)
```python
import hwpx_read, hwpx_edit
hwpx_read.check_refs("안열리는.hwpx")
#  → {} 면 정상. 깨졌으면 {'paraPrIDRef': {'defined_max':20, 'missing':[21]}}
#    (hwpx_read.validate(...) 결과의 errors에도 '참조 무결성: …' 로 포함된다)
hwpx_edit.repair("안열리는.hwpx")
#  → 머리부에서 같은 종류의 '가장 큰 기존 ID 정의를 복제'해 누락 id로 추가하고 itemCnt를 +1.
#    원본은 <원본>.bak.hwpx로 백업, mimetype 무압축으로 재패키징.
#    반환 {'fixed':True,'added':{'paraPr':[21]},'out':...,'backup':...}
#  세밀 제어: ed=hwpx_edit.HwpxEditor(p); ed.missing_refs(); ed.repair_refs(); ed.save(p)
```
CLI: `python hwpx_read.py 파일.hwpx --validate`(검사) / `python hwpx_edit.py 파일.hwpx --repair`(복구).

복제 전략은 "없는 서식을 기존과 동일한 안전한 서식으로 채워 **일단 열리게** 만드는" 것이라,
깨졌던 문단의 외형은 복제본을 따른다(원래 의도한 고유 서식이면 열린 뒤 한글에서 조정).

### 16-3. 표 격자 손상 — 열 때 무한 레이아웃 루프 (★실측 파국)
참조는 멀쩡한데 한글이 열다가 **CPU 100%로 멈추는**(파일 크기 변화 없이 수백 CPU초 소모) 경우.
원인은 **표 격자가 깨진 것**이며, 외부 도구가 표를 병합·생성하며 만든다. 세 갈래:

1. **셀 `<hp:cellAddr>`(행·열 좌표) 누락** ← 가장 치명적. HWPX에서 모든 `<hp:tc>`는
   `<hp:cellAddr colAddr rowAddr/>`가 필수인데 통째로 빠지면 한글이 셀을 격자에 못 놓아 무한 루프.
2. **ragged 열너비** — 행마다 열 너비가 달라(예: 6~8가지 패턴) 열 경계가 안 맞음(병합 아님).
3. **표 `sz.height` 가 행 높이 합과 불일치**(예: 선언 100인데 실제 4920).

```python
import hwpx_read, hwpx_edit
hwpx_read.check_tables("안열리는.hwpx")
#  → [] 면 정상. 깨졌으면 ['표#4: 셀 cellAddr 누락 48개...', '표#4: 행별 열너비 불일치...']
hwpx_edit.repair("안열리는.hwpx")          # 참조 + 표 격자(cellAddr 삽입·열너비 통일·높이 보정) 일괄 복구
#    ed.fix_table_grid() 로 표만 따로 고칠 수도 있음. 반환 tables={'celladdr':N,'colgrid':[..],'height':[..]}
```
복구는 cellAddr를 colSpan/rowSpan 반영해 좌표 계산·삽입하고, ragged 행은 '가장 흔한 열너비 패턴'으로
통일한다(열만 정렬, 텍스트 보존). **복구 후 반드시 baking** 으로 행 높이를 재계산시킨다.

> 진단 팁: 자동화 `Open`이 안 끝나면 `Get-Process Hwp`의 **CPU가 계속 오르는지** 본다 —
> 계속 오르면(>수백 초) 손상(루프), 금방 끝나면 정상. `check_tables`로 표를 먼저 점검하면 빠르다.

### 16-4. 검증의 함정 (이 환경)
한글 COM `Open`/baking으로 최종 확인하는 게 정석이나, **헤드리스/자동화 세션에서는
`HWPFrame.HwpObject.Open`이 모달 대화상자에서 멈추거나 손상 문서에서 무한 루프**할 수 있다.
- **모달 대기**(CPU 거의 0) → 환경 문제일 수 있음. `validate`가 통과하면 구조는 정상으로 본다.
- **CPU 계속 상승**(수백 초) → 진짜 손상(표 격자 루프 등). `check_tables`/`check_refs`로 원인을 찾는다.
baking이 안 끝나면 먼저 `validate(pre_bake=False)`가 빈 결과인지, `check_tables`가 `[]`인지로 판단.

---

> 본 문서는 2026-06 실전(경영학 시간표·회의록 HWPX 자동 생성, 심사용 논문 교정추적 분석,
> RISE 지산학협력 수요조사서 양식 자동 작성) 경험을 일반화한 것임.
> 재사용 스킬(`hwpx`)로도 패키징되어 있어 "hwpx로 출력" 요청 시 자동 활용 가능.
