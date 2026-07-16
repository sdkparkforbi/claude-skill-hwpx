# -*- coding: utf-8 -*-
"""웹 HTML 을 그대로 인쇄해 PDF 생성 (웹 화면과 픽셀 단위로 동일).
사용:  python make_pdf.py <web.html> <out.pdf>
"""
import os, sys
from playwright.sync_api import sync_playwright

def make_pdf(src_html, out_pdf):
    src = os.path.abspath(src_html)
    out = os.path.abspath(out_pdf)
    url = "file:///" + src.replace("\\", "/")
    with sync_playwright() as p:
        b = p.chromium.launch()
        pg = b.new_page()
        pg.goto(url, wait_until="networkidle")
        try:
            pg.evaluate("document.fonts.ready")
        except Exception:
            pass
        pg.emulate_media(media="print")
        pg.pdf(path=out, format="A4",
               margin={"top": "14mm", "bottom": "14mm", "left": "12mm", "right": "12mm"},
               print_background=True, prefer_css_page_size=True)
        b.close()
    print(f"  PDF -> {os.path.basename(out)}  {os.path.getsize(out):,} B")
    return out

if __name__ == "__main__":
    make_pdf(sys.argv[1], sys.argv[2])
