# -*- coding: utf-8 -*-
"""HWPX에 명명 스타일(개요 1/2/3)을 주입하고 제목 문단을 연결한다.
한글은 HTML을 가져올 때 h2~h6에 스타일을 안 만들어(제목이 '바탕글'로 보임),
COM으로는 스타일 생성이 안 되므로 HWPX(zip+xml)를 직접 편집한다.
제목 문단의 paraPr에는 이미 개요 수준(level)이 들어 있어, 스타일 이름만 붙이면
스타일 창(F6)에 '개요 1/2/3'으로 표시되고 차례/개요 보기가 동작한다.
사용:  inject(hwpx_path, off_html)
"""
import zipfile, re, os, html as _html

def _headings(off_html):
    h = open(off_html, encoding="utf-8").read()
    m = re.search(r'(?s)<main[^>]*>(.*?)</main>', h)
    body = m.group(1) if m else h
    out = []
    for mm in re.finditer(r'(?s)<h([234])[^>]*>(.*?)</h\1>', body):
        lvl = int(mm.group(1)) - 2
        txt = re.sub(r'(?s)<[^>]+>', '', mm.group(2))
        txt = re.sub(r'\s+', ' ', _html.unescape(txt)).strip()
        if txt:
            out.append((lvl, txt))
    return out

_NAMES = {0: ("개요 1", "Outline 1"), 1: ("개요 2", "Outline 2"), 2: ("개요 3", "Outline 3")}

def page_break_before(hwpx_path, text):
    """지정 텍스트가 든 문단에 pageBreak='1'을 부여해 그 앞에서 쪽을 나눔(표 손상 없음)."""
    import zipfile, os as _os
    z = zipfile.ZipFile(hwpx_path); order = z.namelist()
    data = {n: z.read(n) for n in order}; z.close()
    sec = data["Contents/section0.xml"].decode("utf-8")
    i = sec.find(text)
    if i < 0:
        return False
    tbl = sec.rfind("<hp:tbl", 0, i)      # 텍스트를 담은 표
    anchor = tbl if tbl >= 0 else i       # 표를 품은 문단을 찾기 위한 기준점
    p = sec.rfind("<hp:p ", 0, anchor)    # 표(또는 텍스트)를 품은 최상위 문단
    if p < 0:
        return False
    end = sec.find(">", p)
    tag = sec[p:end + 1]
    if 'pageBreak=' in tag:
        newtag = re.sub(r'pageBreak="\d+"', 'pageBreak="1"', tag)
    else:
        newtag = tag[:-1] + ' pageBreak="1">'
    sec = sec[:p] + newtag + sec[end + 1:]
    data["Contents/section0.xml"] = sec.encode("utf-8")
    tmp = hwpx_path + ".tmp"
    zo = zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED)
    zo.writestr("mimetype", data["mimetype"], compress_type=zipfile.ZIP_STORED)
    for n in order:
        if n != "mimetype":
            zo.writestr(n, data[n])
    zo.close(); _os.replace(tmp, hwpx_path)
    return True

def space_tables(hwpx_path, top=300, bottom=600):
    """표(상자) 바깥 여백을 키워 다음 문단과 겹치지 않게 함. HWPX zip 직접 편집."""
    import zipfile
    z = zipfile.ZipFile(hwpx_path); order = z.namelist()
    data = {n: z.read(n) for n in order}; z.close()
    sec = data["Contents/section0.xml"].decode("utf-8")
    def _om(m):
        s = re.sub(r'top="\d+"', 'top="%d"' % top, m.group(0))
        return re.sub(r'bottom="\d+"', 'bottom="%d"' % bottom, s)
    sec = re.sub(r'<hp:outMargin\b[^>]*/>', _om, sec)
    data["Contents/section0.xml"] = sec.encode("utf-8")
    tmp = hwpx_path + ".tmp"
    zo = zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED)
    zo.writestr("mimetype", data["mimetype"], compress_type=zipfile.ZIP_STORED)
    for n in order:
        if n != "mimetype":
            zo.writestr(n, data[n])
    zo.close()
    import os as _os
    _os.replace(tmp, hwpx_path)

def _ptext(block):
    return re.sub(r'\s+', ' ', _html.unescape(''.join(re.findall(r'<hp:t>(.*?)</hp:t>', block, re.S)))).strip()

def inject(hwpx_path, off_html):
    heads = _headings(off_html)
    if not heads:
        return 0
    levels = sorted(set(l for l, _ in heads))

    z = zipfile.ZipFile(hwpx_path)
    order = z.namelist()
    data = {n: z.read(n) for n in order}
    z.close()

    hdr = data["Contents/header.xml"].decode("utf-8")
    sec = data["Contents/section0.xml"].decode("utf-8")

    m = re.search(r'<hh:styles itemCnt="(\d+)">', hdr)
    base = int(m.group(1))
    style_id = {lvl: base + i for i, lvl in enumerate(levels)}   # 레벨→새 스타일 id

    # (1) section0: 제목 문단을 순서대로 매칭해 styleIDRef 교체 + 대표 ref 수집
    reps = {}
    state = {"i": 0, "n": 0}
    def repl(mm):
        block = mm.group(0)
        if state["i"] >= len(heads):
            return block
        lvl, txt = heads[state["i"]]
        if _ptext(block) == txt:
            state["i"] += 1; state["n"] += 1
            ot = re.match(r'<hp:p\b[^>]*>', block).group(0)
            if lvl not in reps:
                ppr = re.search(r'paraPrIDRef="(\d+)"', ot)
                cpr = re.search(r'charPrIDRef="(\d+)"', block)
                reps[lvl] = (ppr.group(1) if ppr else "0", cpr.group(1) if cpr else "0")
            ot2 = re.sub(r'styleIDRef="\d+"', 'styleIDRef="%d"' % style_id[lvl], ot)
            return ot2 + block[len(ot):]
        return block
    sec = re.sub(r'<hp:p\b[^>]*>.*?</hp:p>', repl, sec, flags=re.S)

    # (1-b) 표(상자) 바깥 여백 확대 — 기본 top/bottom 141(≈0.5mm)은 다음 문단과 붙는다.
    #       위 300 / 아래 600 HWPUNIT(≈1.1mm/2.1mm)로 키워 겹침 방지. 좌우는 유지.
    def _outm(mm):
        return re.sub(r'top="\d+"', 'top="300"', re.sub(r'bottom="\d+"', 'bottom="600"', mm.group(0)))
    sec = re.sub(r'<hp:outMargin\b[^>]*/>', _outm, sec)

    # (2) header: 스타일 요소 추가 + itemCnt 증가
    new_styles = ""
    for lvl in levels:
        ppr, cpr = reps.get(lvl, ("0", "0"))
        nm, en = _NAMES[lvl]
        new_styles += ('<hh:style id="%d" type="PARA" name="%s" engName="%s" '
                       'paraPrIDRef="%s" charPrIDRef="%s" nextStyleIDRef="0" '
                       'langID="1042" lockForm="0"/>' % (style_id[lvl], nm, en, ppr, cpr))
    hdr = hdr.replace("</hh:styles>", new_styles + "</hh:styles>", 1)
    hdr = hdr.replace('<hh:styles itemCnt="%d">' % base,
                      '<hh:styles itemCnt="%d">' % (base + len(levels)), 1)

    data["Contents/header.xml"] = hdr.encode("utf-8")
    data["Contents/section0.xml"] = sec.encode("utf-8")

    # (3) 재패키징 (mimetype 은 맨 앞·무압축)
    tmp = hwpx_path + ".tmp"
    zo = zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED)
    if "mimetype" in data:
        zo.writestr("mimetype", data["mimetype"], compress_type=zipfile.ZIP_STORED)
    for n in order:
        if n == "mimetype":
            continue
        zo.writestr(n, data[n])
    zo.close()
    os.replace(tmp, hwpx_path)
    return state["n"]

if __name__ == "__main__":
    import sys
    print("스타일 적용:", inject(sys.argv[1], sys.argv[2]), "개 제목")
