# -*- coding: utf-8 -*-
"""개조식 마스터 1종 → PDF + HWPX + DOCX 생성(네 형식 일치용).
사용:  python build_gaejo.py <master.html> <out_base> [표지뒤_쪽나눔_기준텍스트]
  PDF  : 마스터를 playwright로 인쇄(웹과 동일)
  HWPX : make_gaejo_hwp → 한글 → 표 여백 확대 → 표지 뒤 쪽 나눔
  DOCX : 한글 OOXML 내보내기
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import make_gaejo_hwp, make_pdf, hwp_export, hwpx_style

def build(master, out_base, break_text=None):
    d = os.path.dirname(out_base) or "."
    os.makedirs(d, exist_ok=True)
    hwp_html = os.path.join(d, "_hwp_" + os.path.basename(out_base) + ".html")
    print("[PDF]"); make_pdf.make_pdf(master, out_base + ".pdf")
    print("[HWP] transform"); make_gaejo_hwp.build(master, hwp_html)
    print("[HWP] export"); hwp_export.export(hwp_html, out_base)
    hwpx_style.space_tables(out_base + ".hwpx", top=300, bottom=750)
    if break_text:
        ok = hwpx_style.page_break_before(out_base + ".hwpx", break_text)
        print("  표지 쪽나눔:", ok)
    for ext in ("pdf", "hwpx", "docx"):
        p = out_base + "." + ext
        print("  %-5s %s" % (ext, ("%,d B" % os.path.getsize(p)) if os.path.exists(p) else "없음"))

if __name__ == "__main__":
    build(sys.argv[1], sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else None)
