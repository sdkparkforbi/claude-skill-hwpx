# -*- coding: utf-8 -*-
"""DOCX에 Word 제목 스타일(Heading 1/2/3)을 우리 서식(검정 고딕)으로 정의·적용.
한글 OOXML 내보내기는 제목에 스타일을 안 붙여(바탕글) 두므로, 여기서
  · Heading 1/2/3 스타일을 흑백 고딕 + 개요 수준(outlineLvl)으로 정의
  · 제목 문단(표 안 포함)을 문서 순서로 매칭해 해당 스타일로 지정
  · 제목 런에 색/글꼴/굵기를 직접 고정해 흑백 유지
사용:  inject(docx_path, off_html)
"""
import re, html as _html
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from docx.table import Table

GOTHIC = "맑은 고딕"
SPEC = {0: (14, 0), 1: (12, 1), 2: (11, 2)}          # level: (pt, outlineLvl)
NAME = {0: "Heading 1", 1: "Heading 2", 2: "Heading 3"}

def _headings(off_html):
    h = open(off_html, encoding="utf-8").read()
    m = re.search(r'(?s)<main[^>]*>(.*?)</main>', h)
    body = m.group(1) if m else h
    out = []
    for mm in re.finditer(r'(?s)<h([234])[^>]*>(.*?)</h\1>', body):
        txt = re.sub(r'\s+', ' ', _html.unescape(re.sub(r'(?s)<[^>]+>', '', mm.group(2)))).strip()
        if txt:
            out.append((int(mm.group(1)) - 2, txt))
    return out

def _norm(s):
    return re.sub(r'\s+', ' ', s).strip()

def _walk(parent, doc):
    for ch in parent.iterchildren():
        if ch.tag == qn('w:p'):
            yield Paragraph(ch, doc)
        elif ch.tag == qn('w:tbl'):
            for row in Table(ch, doc).rows:
                for cell in row.cells:
                    yield from _walk(cell._tc, doc)

def _ensure_style(doc, level):
    nm = NAME[level]; sz, lvl = SPEC[level]
    try:
        st = doc.styles[nm]
    except KeyError:
        st = doc.styles.add_style(nm, WD_STYLE_TYPE.PARAGRAPH)
    st.font.name = GOTHIC
    st.font.bold = True
    st.font.size = Pt(sz)
    st.font.color.rgb = RGBColor(0, 0, 0)
    # 동아시아(한글) 글꼴 지정
    rpr = st.element.get_or_add_rPr()
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts'); rpr.append(rfonts)
    rfonts.set(qn('w:eastAsia'), GOTHIC)
    # 문단: 개요 수준 + 위/아래 간격 + 다음과 함께
    pPr = st.element.get_or_add_pPr()
    ol = pPr.find(qn('w:outlineLvl'))
    if ol is None:
        ol = OxmlElement('w:outlineLvl'); pPr.append(ol)
    ol.set(qn('w:val'), str(lvl))
    pf = st.paragraph_format
    pf.space_before = Pt({0: 14, 1: 10, 2: 8}[level])
    pf.space_after = Pt({0: 5, 1: 4, 2: 3}[level])
    pf.keep_with_next = True
    return st

def inject(docx_path, off_html):
    heads = _headings(off_html)
    if not heads:
        return 0
    doc = Document(docx_path)
    styles = {lvl: _ensure_style(doc, lvl) for lvl in sorted(set(l for l, _ in heads))}
    i = n = 0
    for p in _walk(doc.element.body, doc):
        if i >= len(heads):
            break
        lvl, txt = heads[i]
        if _norm(p.text) == txt:
            p.style = styles[lvl]
            for r in p.runs:                 # 흑백 고정(스타일 기본색에 지지 않게)
                r.font.color.rgb = RGBColor(0, 0, 0)
                r.font.name = GOTHIC
                r.font.bold = True
                rpr = r._element.get_or_add_rPr()
                rf = rpr.find(qn('w:rFonts'))
                if rf is None:
                    rf = OxmlElement('w:rFonts'); rpr.append(rf)
                rf.set(qn('w:eastAsia'), GOTHIC)
            i += 1; n += 1
    doc.save(docx_path)
    return n

if __name__ == "__main__":
    import sys
    print("DOCX 제목 스타일 적용:", inject(sys.argv[1], sys.argv[2]), "개")
