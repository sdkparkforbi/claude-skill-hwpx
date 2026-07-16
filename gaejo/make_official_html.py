# -*- coding: utf-8 -*-
"""관공서 표준 서식용 HTML 생성 — 웹 원본의 '내용'만 가져와 정식 보고서 형태로 바꾼다.
  · 흑백(검정 본문), 컬러 제목·색 박스 제거
  · 본문 '바탕'(명조) 11pt / 제목 '맑은 고딕' 굵게 / 표는 검정 실선
  · 박스(.bg/.note/.plain/.prelim/.toc)는 연회색 음영의 실선 1칸 표로
  · 질문블록(.ans)은 굵은 질문줄 + 본문(검정)
  · 배지(.rfp)는 [대괄호] 굵은 검정 텍스트로
  ※ 제목 번호는 원문 그대로(1. / 1.1) 유지 — 대량 재번호는 오류 위험이 커 보류.
사용:  python make_official_html.py <src.html> <out.html>
"""
import re, sys, os

CSS = """
@page{size:A4;margin:20mm 20mm}
body{font-family:'바탕',serif;font-size:11pt;line-height:1.7;color:#000;width:170mm}
.doc,main{max-width:170mm;width:170mm;padding:0;margin:0;background:#fff}
h1{display:none}
p{margin:0 0 7pt;text-align:justify;color:#000}
strong,b{font-weight:bold;color:#000}
ul,ol{margin:3pt 0 6pt 0;padding-left:15pt} li{margin:2pt 0;color:#000}
hr{border:0;border-top:0.4pt solid #000;margin:12pt 0}
em{color:#333;font-style:normal;font-size:10pt}
code{font-family:Consolas,monospace;font-size:10pt}
a{color:#000;text-decoration:none}
"""

# 모든 CSS 변수를 흑백/회색으로 수렴
VARS = {'--ink':'#000000','--soft':'#333333','--muted':'#555555','--line':'#000000',
        '--accent':'#000000','--accent2':'#000000','--band':'#e6e6e6','--paper':'#ffffff',
        '--green':'#000000','--red':'#000000'}

HEAD = {  # 태그 : (font-family, 크기, 여백, 아래선)
    'h2': ("'맑은 고딕',sans-serif", '14pt', '18pt 0 8pt', 'border-bottom:1pt solid #000;padding-bottom:3pt;'),
    'h3': ("'맑은 고딕',sans-serif", '12pt', '13pt 0 5pt', ''),
    'h4': ("'맑은 고딕',sans-serif", '11pt', '9pt 0 3pt', ''),
    'h5': ("'맑은 고딕',sans-serif", '10.5pt', '7pt 0 2pt', ''),
}
P_CLS = {
    'cap':  "font-size:9pt;color:#444;margin:2pt 0 8pt",
    'ref':  "font-size:10pt;color:#000;text-indent:-11pt;padding-left:11pt;margin:3pt 0;text-align:left",
    'lead': "color:#000",
    'st':   "color:#000;font-size:11pt;margin:3pt 0 8pt;text-align:center",
    'kicker': "letter-spacing:.2em;color:#000;font-weight:bold;font-size:10pt;margin-bottom:6pt;text-align:center",
    'meta': "border-top:0.4pt solid #000;border-bottom:0.4pt solid #000;padding:7pt 0;color:#000;font-size:10pt;margin:8pt 0",
}

def _match_div(h, open_start):
    tag = re.compile(r'<div\b[^>]*>|</div>')
    depth, pos = 0, open_start
    while True:
        m = tag.search(h, pos)
        if not m:
            return None, None
        if m.group(0).startswith('</'):
            depth -= 1
            if depth == 0:
                return m.start(), m.end()
        else:
            depth += 1
        pos = m.end()

def _cell(inner):
    return ('<table border="1" cellspacing="0" cellpadding="0" width="100%" '
            'style="border-collapse:collapse;width:170mm;margin:8pt 0;border:0.4pt solid #000"><tr>'
            '<td style="background:#f4f4f4;border:0.4pt solid #000;padding:7pt 10pt;'
            'font-size:10pt;color:#000">' + inner + '</td></tr></table>')

SIMPLE = ('bg', 'note', 'plain', 'prelim', 'toc')

def _boxes_to_tables(h):
    # (1) .ans → 2행 표(질문 음영 + 답변)
    while True:
        m = re.search(r'<div class="ans">', h)
        if not m:
            break
        cs, ce = _match_div(h, m.start())
        if cs is None:
            break
        inner = h[m.end():cs]
        mq = re.search(r'(?s)<div class="q">(.*?)</div>', inner)
        q = mq.group(1) if mq else ''
        mb = re.search(r'<div class="b">', inner)
        b = ''
        if mb:
            bs, be = _match_div(inner, mb.start())
            b = inner[mb.end():bs] if bs is not None else ''
        tbl = ('<table border="1" cellspacing="0" cellpadding="0" width="100%" '
               'style="border-collapse:collapse;width:170mm;border:0.4pt solid #000;margin:9pt 0">'
               '<tr><td style="background:#e8e8e8;border:0.4pt solid #000;'
               'padding:6pt 9pt;font-weight:bold;font-size:11pt;color:#000">' + q + '</td></tr>'
               '<tr><td style="border:0.4pt solid #000;padding:7pt 9pt 8pt;font-size:10.5pt;color:#000">' + b + '</td></tr></table>')
        h = h[:m.start()] + tbl + h[ce:]
    # (2) 단순 박스 → 연회색 1칸 표
    for cls in SIMPLE:
        while True:
            m = re.search(r'<div class="' + cls + r'"( style="[^"]*")?>', h)
            if not m:
                break
            cs, ce = _match_div(h, m.start())
            if cs is None:
                break
            inner = h[m.end():cs]
            h = h[:m.start()] + _cell(inner) + h[ce:]
    return h

def build(src_html, out_html):
    h = open(src_html, encoding="utf-8").read()
    # 1) 외부 폰트·nav 제거
    h = re.sub(r'<link[^>]*fonts\.(googleapis|gstatic)[^>]*>', '', h)
    h = re.sub(r'(?s)<nav[^>]*class="topnav".*?</nav>', '', h)
    # 1-1) 앵커 링크 해제 — 한글은 <a>를 파란 하이퍼링크로 강제한다. 인쇄문서엔 불필요.
    h = re.sub(r'(?s)<a\b[^>]*>(.*?)</a>', r'\1', h)
    # 2) CSS 교체
    h = re.sub(r'(?s)<style>.*?</style>', '<style>' + CSS + '</style>', h)
    # 3) 표 → 검정 실선 인라인
    h = re.sub(r'<table>', '<table border="1" cellspacing="0" cellpadding="3" width="100%" '
               'style="border-collapse:collapse;width:170mm;table-layout:fixed;font-size:9.5pt;'
               'border:0.4pt solid #000;word-break:break-all;margin:8pt 0">', h)
    h = re.sub(r'<th(\s+class="l")?>', lambda m: '<th style="border:0.4pt solid #000;background:#e6e6e6;'
               'font-weight:bold;padding:4pt 5pt;color:#000;text-align:' + ('left' if m.group(1) else 'center') + '">', h)
    h = re.sub(r'<td(\s+class="l")?>', lambda m: '<td style="border:0.4pt solid #000;padding:4pt 5pt;color:#000;'
               'vertical-align:top;text-align:' + ('left' if m.group(1) else 'center') + '">', h)
    h = re.sub(r'<th style="width:[^"]*"[^>]*>', '<th style="border:0.4pt solid #000;background:#e6e6e6;'
               'font-weight:bold;padding:4pt 5pt;color:#000;text-align:center">', h)
    h = re.sub(r'<caption>', '<caption style="text-align:left;font-weight:bold;font-size:10pt;color:#000;margin-bottom:3pt">', h)
    # 4) 박스 → 표
    h = _boxes_to_tables(h)
    # 5) 배지 .rfp → [대괄호] 굵은 검정
    h = re.sub(r'(?s)<span class="rfp">(.*?)</span>',
               lambda m: '<span style="font-weight:bold;color:#000">[' + m.group(1).strip() + ']</span> ', h)
    # 6) p/div 클래스 인라인
    for cls, st in P_CLS.items():
        h = re.sub(r'<(p|div) class="' + cls + r'">', lambda m, s=st: '<' + m.group(1) + ' style="' + s + '">', h)
    # 7) 표지: 상단 굵은선 + 가운데 정렬(검정)
    h = re.sub(r'<section class="cover">', '<section style="border-top:2.5pt solid #000;padding:10pt 0 12pt;'
               'margin-bottom:10pt;text-align:center">', h)
    h = re.sub(r'<h1>', '<h1 style="display:block;font-family:\'맑은 고딕\',sans-serif;font-size:20pt;'
               'font-weight:bold;color:#000;margin:6pt 0 5pt;line-height:1.3;text-align:center">', h)
    # 8) 제목 흑백 인라인 강제 (한글 색 상속 차단)
    for tag, (ff, sz, mg, extra) in HEAD.items():
        st = ("font-family:" + ff + ";font-size:" + sz + ";font-weight:bold;color:#000;margin:" + mg + ";" + extra)
        h = re.sub(r'<' + tag + r'([^>]*)>', lambda m, s=st: '<' + tag + m.group(1) + ' style="' + s + '">', h)
    h = re.sub(r'<p>', '<p style="color:#000;margin:0 0 7pt;text-align:justify">', h)
    # 목록 들여쓰기 최소화 — 한글은 <ol>/<ul> 기본 들여쓰기가 과도하다.
    h = re.sub(r'<(ul|ol)>', r'<\1 style="margin:3pt 0 6pt 0;padding-left:16pt">', h)
    h = re.sub(r'<li>', '<li style="color:#000;margin:2pt 0">', h)
    # 9) 제목 안쪽을 <span>으로 감싸 문자색을 검정으로 못박음(한글이 확실히 따름)
    def _wrap(tag):
        def f(m):
            inner = m.group(2)
            if '<span style="color:' in inner[:30]:
                return m.group(0)
            return '<' + tag + m.group(1) + '><span style="color:#000">' + inner + '</span></' + tag + '>'
        return f
    for tag in ('h2', 'h3', 'h4', 'h5'):
        h = re.sub(r'<' + tag + r'([^>]*)>(.*?)</' + tag + r'>', _wrap(tag), h, flags=re.S)
    # 10) CSS 변수 해소 → 흑백
    for var, hexv in VARS.items():
        h = h.replace('var(' + var + ')', hexv)
    # 남은 파랑/컬러 인라인을 검정으로 (본문 안 색 span 등)
    h = re.sub(r'color:#2f5aa8', 'color:#000', h)
    h = re.sub(r'color:#8a6416', 'color:#000', h)
    h = re.sub(r'color:#2f7d32', 'color:#000', h)
    open(out_html, "w", encoding="utf-8").write(h)
    return out_html

if __name__ == "__main__":
    src, out = sys.argv[1], sys.argv[2]
    build(src, out)
    print("생성:", out, os.path.getsize(out), "bytes")
