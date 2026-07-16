# -*- coding: utf-8 -*-
"""개조식 마스터(연구계획서_개조식.html) → 한글 가져오기용 HTML.
한글은 클래스 CSS·flex·div border-top을 잘 못 읽으므로 스타일을 인라인으로 옮기고,
표지의 파란 선은 표(bar)로 재구성한다. 기호(❍/❐/-/*)는 이미 실제 문자라 그대로 유지.
사용:  build(master.html, out.html)
"""
import re

BASE = ("@page{size:A4;margin:20mm 20mm 18mm}"
        "body{font-family:'맑은 고딕',sans-serif;font-size:11pt;line-height:1.7;color:#1a1a1a}"
        ".doc{width:180mm}")

NAVY = "#1f4e9c"

# 클래스 → 인라인 style
PSTYLE = {
    'o':    "margin:0 0 2.6mm 9mm;text-indent:-6mm;color:#1a1a1a",
    'sq':   "margin:0 0 2.6mm 9mm;text-indent:-6mm;color:#1a1a1a",
    'dash': "margin:0 0 1.8mm 16mm;text-indent:-5mm;color:#333",
    'star': "margin:0 0 1.7mm 16mm;text-indent:-5mm;color:#333",
    'sub-lead': "margin:1mm 0 3.5mm;color:#444;font-size:10.5pt",
    'ref':  "margin:1.5mm 0 1.5mm 6mm;text-indent:-6mm;font-size:9.5pt;color:#444",
    'body': "margin:2.5mm 0",
    'close':"margin:6mm 0 2mm;font-size:10.5pt",
}

def _bar():
    return ('<table width="100%%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:180mm">'
            '<tr><td style="background:%s;height:3pt;line-height:3pt;font-size:1pt">&nbsp;</td></tr></table>' % NAVY).replace('%%','%')

def _cover():
    sp = lambda mm: '<p align="center" style="margin:0;font-size:%dpt;line-height:%dpt">&nbsp;</p>' % (mm, mm)
    c = 'align="center" style="text-align:center;'
    return (
        _bar()
        + sp(66)
        + '<p ' + c + 'margin:0 0 7mm;font-size:14pt;font-weight:bold;color:' + NAVY + '">주세 정책이 경제·물가에 미치는 영향 연구</p>'
        + _bar()
        + '<p ' + c + 'margin:8mm 0;font-size:30pt;font-weight:bold;letter-spacing:8pt;color:#111">연구계획서(안)</p>'
        + _bar()
        + sp(26)
        + '<p ' + c + 'margin:0;font-size:14pt;font-weight:bold">2026. 7.</p>'
        + sp(40)
        + '<p ' + c + 'margin:0;font-size:13pt;font-weight:bold;color:#333">한국주류산업협회</p>'
    )

def build(master, out):
    h = open(master, encoding="utf-8").read()
    h = re.sub(r'<link[^>]*fonts\.(googleapis|gstatic)[^>]*>', '', h)
    h = re.sub(r'(?s)<nav[^>]*class="topnav".*?</nav>', '', h)
    # 1) style 교체
    h = re.sub(r'(?s)<style>.*?</style>', '<style>' + BASE + '</style>', h)
    # 2) 표지 교체 (뒤에 페이지 나눔)
    h = re.sub(r'(?s)<section class="cover">.*?</section>', _cover(), h)
    # 3) 섹션 헤더 표 인라인
    h = re.sub(r'<table class="sec"><tr>',
               '<table cellspacing="0" cellpadding="0" width="100%" style="border-collapse:collapse;width:180mm;margin:9mm 0 4mm"><tr>', h)
    # 번호칸: 파란 배경 + 흰 글씨(문자서식으로 강제). 내용을 span으로 감쌈.
    h = re.sub(r'<td class="num">(.*?)</td>',
               lambda m: ('<td style="width:10mm;background:' + NAVY + ';text-align:center;vertical-align:middle;'
                          'border-bottom:1.6pt solid ' + NAVY + ';padding:2mm 0">'
                          '<span style="color:#ffffff;font-size:15pt;font-weight:bold">' + m.group(1) + '</span></td>'), h, flags=re.S)
    # 제목칸: 큰 굵은 글씨(문자서식으로 강제).
    h = re.sub(r'<td class="tit">(.*?)</td>',
               lambda m: ('<td style="padding:0 0 1.5mm 4mm;vertical-align:bottom;border-bottom:1.6pt solid ' + NAVY + '">'
                          '<span style="font-size:16pt;font-weight:bold;color:#111">' + m.group(1) + '</span></td>'), h, flags=re.S)
    # 4) 본문 표(class d)
    h = re.sub(r'<table class="d">',
               '<table border="1" cellspacing="0" cellpadding="0" width="100%" '
               'style="border-collapse:collapse;width:180mm;font-size:9.5pt;border:0.5pt solid #8a97ad;margin:4mm 0">', h)
    h = re.sub(r'<th class="l"[^>]*>', '<th style="border:0.5pt solid #8a97ad;background:#eaf0f8;font-weight:bold;'
               'text-align:left;padding:2mm 3mm;color:#173a75">', h)
    h = re.sub(r'<th style="width:[^"]*">', '<th style="border:0.5pt solid #8a97ad;background:#eaf0f8;font-weight:bold;'
               'text-align:center;padding:2mm 3mm;color:#173a75">', h)
    h = re.sub(r'<th>', '<th style="border:0.5pt solid #8a97ad;background:#eaf0f8;font-weight:bold;'
               'text-align:center;padding:2mm 3mm;color:#173a75">', h)
    h = re.sub(r'<td class="l">', '<td style="border:0.5pt solid #8a97ad;padding:2mm 3mm;text-align:left;vertical-align:middle">', h)
    h = re.sub(r'<td>', '<td style="border:0.5pt solid #8a97ad;padding:2mm 3mm;text-align:center;vertical-align:middle">', h)
    # 한글은 <caption>을 문서 끝으로 밀어내므로, 캡션을 표 앞 굵은 문단으로 빼냄
    h = re.sub(r'(?s)(<table\b[^>]*>)\s*<caption[^>]*>(.*?)</caption>',
               r'<p style="font-weight:bold;font-size:10.5pt;margin:4mm 0 1.5mm;color:#111">\2</p>\1', h)
    # 5) 개조식 문단·리드·참고 인라인
    for cls, st in PSTYLE.items():
        h = re.sub(r'<p class="%s">' % cls, '<p style="%s">' % st, h)
    # 6) 그룹 제목(통째 치환)·h3
    h = re.sub(r'(?s)<div class="grp">(.*?)</div>',
               r'<p style="margin:5.5mm 0 2.5mm;font-size:11.5pt;font-weight:bold;color:%s">\1</p>' % NAVY, h)
    h = re.sub(r'<h3>', '<p style="margin:6mm 0 2.5mm;font-size:12pt;font-weight:bold;color:#173a75">', h)
    h = re.sub(r'</h3>', '</p>', h)
    # 7) 마크 기호·강조
    h = re.sub(r'<span class="mk">', '<span style="color:%s;font-weight:bold">' % NAVY, h)
    h = re.sub(r'<span class="u">', '<span style="color:#173a75;font-weight:bold;text-decoration:underline">', h)
    # 8) 참고 박스 → 1칸 표
    def _box(m):
        inner = m.group(1)
        inner = re.sub(r'<div class="t">(.*?)</div>', r'<b style="color:#173a75">\1</b><br>', inner, flags=re.S)
        return ('<table cellspacing="0" cellpadding="0" width="100%" style="border-collapse:collapse;width:180mm;margin:4mm 0">'
                '<tr><td style="background:#f6f9ff;border-left:3pt solid ' + NAVY + ';border-top:0.5pt solid #c9d3e4;'
                'border-right:0.5pt solid #c9d3e4;border-bottom:0.5pt solid #c9d3e4;padding:3.5mm 5mm;'
                'font-size:10.5pt;color:#333">' + inner + '</td></tr></table>')
    h = re.sub(r'(?s)<div class="box">(.*?)</div>\s*(?=<!--|<table)', _box, h)
    # 수신/발신 박스(.addr) → 1칸 표
    def _addr(m):
        inner = re.sub(r'<div>(.*?)</div>', r'<p style="margin:1mm 0">\1</p>', m.group(1), flags=re.S)
        return ('<table cellspacing="0" cellpadding="0" width="100%" style="border-collapse:collapse;width:180mm;margin:3mm 0 6mm">'
                '<tr><td style="background:#f6f9ff;border:0.5pt solid #c9d3e4;padding:4mm 6mm;font-size:10.5pt">'
                + inner + '</td></tr></table>')
    h = re.sub(r'(?s)<div class="addr">(.*?)</div>\s*(?=<table class="sec"|<table cellspacing)', _addr, h)
    open(out, "w", encoding="utf-8").write(h)
    return out

if __name__ == "__main__":
    import sys
    build(sys.argv[1], sys.argv[2])
    print("생성:", sys.argv[2])
