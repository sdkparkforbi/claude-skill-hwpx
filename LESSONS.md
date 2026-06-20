# 실전 교훈 — 안 열리는 .hwpx 진단·복구 (2026-06)

외부 도구/스크립트로 **조립·병합**한 `.hwpx`가 한글에서 *"파일을 열 수 없습니다"* 로 거부되거나,
열다가 **무한 정지(CPU 100%)** 하는 사례를 추적·복구하며 얻은 교훈을 정리한다.
관련 코드는 `hwpx_read.check_refs`·`check_tables`·`validate`, `hwpx_edit.repair`·`fix_table_grid`.

---

## 1. 핵심 관찰: "XML도 정상, 표도 정상인데 안 열린다"

문제 파일은 다음을 모두 **통과**했다.
- ZIP 무결성(`unzip -t`), `mimetype` 무압축·선두
- 모든 XML well-formed(lxml 파싱 성공)
- 표 `rowCnt`/`colCnt`·colSpan 합 일치, instId 중복 없음
- PNG 시그니처 정상, spine/manifest 정상

그런데도 한글이 거부/정지했다. → **표면 검사로는 안 잡히는 더 깊은 무결성 문제**가 있다.

## 2. 실제 원인 (두 층)

### 2-1. 참조 무결성 — dangling IDRef
본문 `section*.xml`이 머리부 `header.xml`에 **정의되지 않은 스타일/속성 ID**를 참조.
- 예: `paraPrIDRef="21"` 인데 `header`의 `paraPr`은 `0~20`만 존재.
- 한글은 이 dangling 참조를 만나면 **통째로 열기를 거부**한다.
- 원인: 표/본문을 다른 문서에서 가져오며 그 문서가 쓰던 `paraPr/charPr/borderFill/style`
  정의를 머리부에 함께 옮기지 않음.
- 대상 attr: `charPrIDRef`, `paraPrIDRef`, `styleIDRef`, `borderFillIDRef`.

### 2-2. 표 격자 손상 — 열 때 무한 레이아웃 루프 ★가장 까다로움
참조가 멀쩡해도 한글이 **열다가 CPU 100%로 멈춘다**(파일 크기 변화 없이 수백 CPU초 소모).
표 격자가 깨진 것이며 세 갈래로 나타났다.

1. **셀 `<hp:cellAddr>`(행·열 좌표) 누락** ← 가장 치명적.
   HWPX에서 모든 `<hp:tc>`는 `<hp:cellAddr colAddr rowAddr/>`가 필수. 통째로 빠지면
   한글이 셀을 격자에 배치하지 못해 무한 루프. (실측: 병합 표 144개 셀에서 전부 누락.)
2. **ragged 열너비** — 행마다 열 너비가 달라(6~8가지 패턴) 열 경계가 안 맞음(병합 아님).
   각 행의 너비 합은 표 너비와 같지만, 열 경계가 행마다 어긋나 격자가 성립하지 않음.
3. **표 `sz.height` ↔ 행 높이 합 불일치** — 선언 높이 100인데 실제 행 합은 4920/9840.

## 3. 진단 방법론 (재현 가능한 절차)

1. **정적 검사 우선(한컴 불필요)**:
   `hwpx_read.validate(path, pre_bake=False)` → 참조·표격자 손상을 한 번에 보고.
   세부는 `check_refs`(빈 dict면 정상), `check_tables`(빈 list면 정상).
2. **열림 확정은 baking**: `python hwpx_bake.py file.hwpx` (한글로 Open→SaveAs→PDF).
   완료되면 = 열린다. 무한 정지면 = 손상.
3. **모달 대기 vs 무한 루프 구분** — 자동화 `Open`이 안 끝날 때 `Get-Process Hwp`의 **CPU**를 본다.
   - CPU 거의 0 → 모달 대화상자 대기(환경 문제일 수 있음).
   - CPU 계속 상승(수백 초) → 진짜 손상(레이아웃 루프).
4. **원인 격리 = 이분 탐색**: 의심 요소(이미지/표)를 제거한 사본을 baking.
   - 이미지 전체 제거 → 여전히 정지 ⇒ 이미지 무관
   - 표 전체 제거 → 정상 ⇒ 표가 원인 → 문제 표만 제거 → 정상 ⇒ 그 표 확정
   각 시도는 CPU 추세로 빠르게 판정(완료 전 kill).

## 4. 복구 (한컴 불필요)

`hwpx_edit.repair(path)` 한 번으로 두 층을 모두 고친다.
- **참조**: 머리부의 *같은 종류 최대 id 정의를 복제*해 누락 id로 추가, `itemCnt +1`.
- **표 격자**: `cellAddr`를 colSpan/rowSpan 반영해 좌표 계산·삽입,
  ragged 행을 *가장 흔한 열너비 패턴*으로 통일, `sz.height`를 행 높이 합으로 보정.
- 원본은 `<원본>.bak.hwpx`로 백업, `mimetype` 무압축으로 재패키징.
- **복구 후 반드시 `hwpx_bake.py`** 로 한 번 열어 행 높이·줄배치를 재계산시킨다.

```python
import hwpx_read, hwpx_edit
print(hwpx_read.validate("안열리는.hwpx", pre_bake=False))   # 진단
hwpx_edit.repair("안열리는.hwpx")                            # 일괄 복구(+.bak)
# 이후: python hwpx_bake.py 안열리는.hwpx
```

## 5. 함정 모음

- **`validate`만 믿지 말 것(과거형)** — 원래 `validate`는 표 구조·생성단계 규칙만 봐서
  참조/표격자 손상을 못 잡았다. 그래서 이 두 검사를 `validate`에 통합했다.
- **재패키징은 파이썬 `zipfile`만**, 원본 순서 유지, `mimetype`은 `ZIP_STORED`로 선두.
- **헤드리스 한글 COM `Open`은 보이지 않는 인스턴스**(MainWindowTitle 빈 값)라
  화면에 대화상자가 안 보일 수 있다. 멈춤 판단은 **CPU 추세**로.
- **빈 매크로 스텁(`Scripts/*`)은 범인이 아니었다** — 제거해도 무한 루프 지속.
- ragged 격자 통일은 "일단 열리게" 만드는 것이라, 헤더성 행의 열 폭이 바뀔 수 있다 →
  복구·baking 후 PDF로 표 모양을 눈으로 확인할 것.

## 6. 한 줄 요약

> 안 열리는 .hwpx의 90%는 **dangling 스타일 참조** 아니면 **표 격자 손상(특히 셀 cellAddr 누락)**.
> `validate`로 진단 → `repair`로 복구 → `bake`로 굳히면 대부분 열린다.
> 자동화에서 한글이 멈추면 **CPU 추세**로 모달 대기와 무한 루프를 구분하라.
