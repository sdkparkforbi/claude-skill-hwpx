# -*- coding: utf-8 -*-
"""한글(HWP)로 _hwp_*.html 을 열어 HWPX·DOCX 생성.
사용:  python hwp_export.py <_hwp_src.html> <출력이름(확장자 제외)>
※ HWPX는 반드시 HTML에서 직접 — DOCX 경유 시 표 테두리 소실.
"""
import os, sys, time, subprocess, re, html as _html
import win32com.client as win32

def kill_hwp():
    for exe in ("Hwp.exe", "HwpApp.exe"):
        subprocess.run(["taskkill", "/F", "/IM", exe], capture_output=True)
    time.sleep(1)

def _set_para_spacing(hwp, tail_pt=5.0, line_pct=160):
    """문서 전체를 선택해 문단 아래 간격(tail_pt)과 줄간격(line_pct%)을 지정.
    한글은 HTML의 <p>{margin} 을 무시하므로 문단이 붙어 나오는 문제를 여기서 보정한다.
    NextSpacing 단위는 실측상 200/pt (600→3pt)."""
    try:
        hwp.Run("SelectAll")
        hwp.HAction.GetDefault("ParagraphShape", hwp.HParameterSet.HParaShape.HSet)
        ps = hwp.HParameterSet.HParaShape
        ps.NextSpacing = int(tail_pt * 200)   # 문단 아래 간격 (실측 200/pt)
        try:
            ps.LineSpacingType = hwp.HwpLineSpacingType("Percent")
        except Exception:
            ps.LineSpacingType = 0
        ps.LineSpacing = line_pct                  # 줄 간격 %
        hwp.HAction.Execute("ParagraphShape", ps.HSet)
        hwp.Run("Cancel")   # 선택 해제
        print(f"  문단 간격 적용: 아래 {tail_pt}pt / 줄 {line_pct}%")
    except Exception as e:
        print("  문단 간격 적용 실패(계속):", str(e)[:70])

def _headings(src_html):
    """공식 HTML에서 (개요수준, 제목텍스트) 목록을 문서 순서대로 뽑는다. h2=0,h3=1,h4=2."""
    h = open(src_html, encoding="utf-8").read()
    m = re.search(r'(?s)<main[^>]*>(.*?)</main>', h)
    body = m.group(1) if m else h
    out = []
    for mm in re.finditer(r'(?s)<h([234])[^>]*>(.*?)</h\1>', body):
        lvl = int(mm.group(1)) - 2
        txt = re.sub(r'(?s)<[^>]+>', '', mm.group(2))      # 태그 제거
        txt = _html.unescape(txt)
        txt = re.sub(r'\s+', ' ', txt).strip()
        if txt:
            out.append((lvl, txt))
    return out

def _find(hwp, s):
    hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
    fr = hwp.HParameterSet.HFindReplace
    fr.FindString = s
    fr.Direction = 0           # 정방향
    fr.IgnoreMessage = 1
    try: fr.FindType = 1
    except Exception: pass
    return hwp.HAction.Execute("RepeatFind", fr.HSet)

# 개요수준별 (제목 위 간격, 제목 아래 간격) — 단위 실측 200/pt
_HSPACE = {0: (2800, 1000), 1: (1900, 800), 2: (1500, 600)}

def _apply_heading(hwp, level):
    hwp.HAction.GetDefault("ParagraphShape", hwp.HParameterSet.HParaShape.HSet)
    ps = hwp.HParameterSet.HParaShape
    prev, nxt = _HSPACE[level]
    ps.PrevSpacing = prev       # 제목 위 간격(새 제목 앞 공간)
    ps.NextSpacing = nxt        # 제목 아래 간격
    ps.Level = level            # 개요 수준(계층) — 자동번호는 켜지 않음
    ps.HeadingType = 0          # 자동 개요번호 OFF (수기번호와 중복 방지)
    ps.KeepWithNext = 1         # 다음 문단과 함께(제목만 페이지 끝에 남지 않게)
    return hwp.HAction.Execute("ParagraphShape", ps.HSet)

def _apply_heading_hierarchy(hwp, src_html):
    """제목마다 개요 수준 + 제목 위/아래 간격 + KeepWithNext 부여."""
    heads = _headings(src_html)
    hwp.Run("MoveDocBegin")
    ok = miss = 0
    for lvl, txt in heads:
        key = txt if len(txt) <= 40 else txt[:40]
        found = _find(hwp, key)
        if not found and len(key) > 16:      # 특수문자(→,— 등)로 실패 시 앞부분만
            found = _find(hwp, txt[:16])
        if found:
            _apply_heading(hwp, lvl); ok += 1
            hwp.Run("Cancel")               # 선택 해제 후 다음 제목 탐색
        else:
            miss += 1
    print(f"  제목 계층 적용: {ok}개 (미검출 {miss})")

def export(src_html, base, page_break_after=None):
    src = os.path.abspath(src_html)
    kill_hwp()
    hwp = win32.Dispatch("HWPFrame.HwpObject")  # 지연 바인딩(gencache 캐시 손상 회피)
    try:
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")  # 보안 팝업 우회
    except Exception as e:
        print("  (RegisterModule 실패, 계속):", str(e)[:50])
    try:
        hwp.SetMessageBoxMode(0x00020000)  # 모든 대화상자 자동 처리(찾기 끝 '처음부터?' 등 블로킹 방지)
    except Exception:
        pass
    hwp.XHwpWindows.Item(0).Visible = False
    ok = hwp.Open(src, "HTML", "")
    print("  Open(HTML):", ok)
    time.sleep(1.0)
    _set_para_spacing(hwp)  # 한글은 <p> margin 무시 → 문단 아래 간격을 API로 부여
    _apply_heading_hierarchy(hwp, src)  # 제목 개요 수준 + 제목 위 간격
    if page_break_after:  # 지정 문구 뒤에 페이지 나눔(표지 분리 등)
        try:
            hwp.Run("MoveDocBegin")
            if _find(hwp, page_break_after):
                hwp.Run("Cancel"); hwp.Run("MoveLineEnd"); hwp.Run("BreakPage")
                print("  페이지 나눔 삽입:", page_break_after[:16])
        except Exception as e:
            print("  페이지 나눔 실패:", str(e)[:50])
    out = {}
    for fmt, ext in (("HWPX", "hwpx"), ("OOXML", "docx")):  # 한글의 워드 저장 포맷명은 OOXML
        p = os.path.abspath(f"{base}.{ext}")
        if os.path.exists(p):
            try: os.remove(p)
            except OSError: pass
        try:
            r = hwp.SaveAs(p, fmt, "")
            time.sleep(0.6)
            sz = os.path.getsize(p) if os.path.exists(p) else 0
            out[ext] = sz
            print(f"  SaveAs {fmt:5s} -> {os.path.basename(p)}  {sz:,} B  (ret={r})")
        except Exception as e:
            print(f"  SaveAs {fmt} 실패:", str(e)[:70])
    try: hwp.Clear(1); hwp.Quit()
    except Exception: pass
    kill_hwp()
    return out

if __name__ == "__main__":
    export(sys.argv[1], sys.argv[2])
