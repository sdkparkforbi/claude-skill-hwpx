#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HWP / HWPX 읽기·추출 엔진.

생성 전용이던 스킬에 '읽기' 방향을 추가한다. 한컴오피스 없이 순수 파이썬으로
기존 문서의 본문·표·스타일을 구조화 추출한다.

지원 형식
  * .hwpx  : ZIP + Contents/section*.xml (lxml로 파싱) — 표 구조까지 완전 추출.
  * .hwp   : OLE 복합문서 + BodyText/Section* 레코드(zlib 압축) — 본문 텍스트 추출
             (표 셀 텍스트는 읽기 순서대로 포함). 구조가 더 필요하면 한컴으로
             hwpx 변환 후(4단계 hwpx_convert) 다시 읽는다.

주요 API
  extract_text(path)            -> str            형식 자동판별 후 평문 추출
  read_hwpx(path)               -> dict           {'format','blocks':[para|table],'text'}
  read_hwp(path)                -> dict           {'format','paragraphs','text'}
  iter_tables(doc)              -> list[list[list[str]]]   표만 골라 행렬로

CLI
  python hwpx_read.py 파일.hwp [파일2.hwpx ...] [--json] [-o out.txt]
"""
import sys, os, io, re, zlib, struct, json, zipfile

# ───────────────────────── 공통 판별 ─────────────────────────
def detect_format(path):
    """확장자가 아니라 매직으로 판별 (hwpx=ZIP 'PK', hwp=OLE \\xD0\\xCF)."""
    with open(path, 'rb') as f:
        head = f.read(8)
    if head[:2] == b'PK':
        return 'hwpx'
    if head[:4] == b'\xd0\xcf\x11\xe0':
        return 'hwp'
    # 확장자 폴백
    ext = os.path.splitext(path)[1].lower()
    return 'hwpx' if ext == '.hwpx' else 'hwp'

# ───────────────────────── HWP (바이너리) ─────────────────────────
# HWP5 PARA_TEXT 인라인 제어문자 분류 (스펙 기준 코드 단위 크기)
_CTRL_SIZE1 = {0, 10, 13, 24, 25, 26, 27, 28, 29, 30, 31}   # 문자형(1 unit)
_CTRL_SIZE8 = {1, 2, 3, 4, 5, 6, 7, 8, 9, 11, 12, 14, 15,
               16, 17, 18, 19, 20, 21, 22, 23}               # 인라인/확장(8 unit)
HWPTAG_PARA_TEXT = 67  # 0x43

def _hwp_is_compressed(ole):
    header = ole.openstream('FileHeader').read()
    # FileHeader 36번째 바이트 bit0 = 본문 압축 여부
    return bool(header[36] & 0x01)

def _decode_para_text(payload):
    """PARA_TEXT 레코드(UTF-16LE + 인라인 제어문자) → 단락 문자열."""
    n = len(payload) // 2
    u = struct.unpack('<%dH' % n, payload[:n * 2])
    out = []
    i = 0
    while i < n:
        c = u[i]
        if c in _CTRL_SIZE8:
            if c == 9:
                out.append('\t')      # 탭
            # 표/그리기 등 개체 제어문자는 건너뜀(텍스트는 별도 단락에 존재)
            i += 8
        elif c in _CTRL_SIZE1:
            if c in (10, 13):
                out.append('\n')      # 줄/단락 구분
            i += 1
        else:
            out.append(chr(c))
            i += 1
    return ''.join(out)

def read_hwp(path):
    import olefile
    ole = olefile.OleFileIO(path)
    try:
        compressed = _hwp_is_compressed(ole)
        # BodyText/Section0,1,... 순서대로
        secs = [e for e in ole.listdir() if len(e) == 2 and e[0] == 'BodyText'
                and e[1].lower().startswith('section')]
        secs.sort(key=lambda e: int(re.sub(r'\D', '', e[1]) or 0))
        paragraphs = []
        for entry in secs:
            data = ole.openstream(entry).read()
            if compressed:
                data = zlib.decompress(data, -15)
            i, ln = 0, len(data)
            while i + 4 <= ln:
                hdr = struct.unpack('<I', data[i:i + 4])[0]
                tag = hdr & 0x3ff
                size = (hdr >> 20) & 0xfff
                i += 4
                if size == 0xfff:
                    size = struct.unpack('<I', data[i:i + 4])[0]
                    i += 4
                chunk = data[i:i + size]
                i += size
                if tag == HWPTAG_PARA_TEXT:
                    txt = _decode_para_text(chunk).strip('\n')
                    if txt.strip():
                        paragraphs.append(txt)
    finally:
        ole.close()
    return {'format': 'hwp', 'paragraphs': paragraphs,
            'text': '\n'.join(paragraphs)}

# ───────────────────────── HWPX (ZIP+XML) ─────────────────────────
_HP = '{http://www.hancom.co.kr/hwpml/2011/paragraph}'

def _xml_text_of_para(p):
    """hp:p 안의 모든 hp:t 텍스트를 이어붙임."""
    parts = []
    for t in p.iter(_HP + 't'):
        parts.append(t.text or '')
        # 줄바꿈(hp:lineBreak) 등 자식 tail 처리
    return ''.join(parts)

def _parse_table(tbl):
    """hp:tbl → 행렬(list[list[str]]). 병합 칸은 비워둠."""
    rows = []
    for tr in tbl.findall(_HP + 'tr'):
        row = []
        for tc in tr.findall(_HP + 'tc'):
            cell_txt = []
            for p in tc.iter(_HP + 'p'):
                line = _xml_text_of_para(p)
                if line:
                    cell_txt.append(line)
            row.append('\n'.join(cell_txt))
        rows.append(row)
    return rows

def read_hwpx(path):
    from lxml import etree
    blocks = []
    with zipfile.ZipFile(path) as z:
        names = [n for n in z.namelist()
                 if re.match(r'Contents/section\d+\.xml$', n)]
        names.sort(key=lambda n: int(re.search(r'section(\d+)', n).group(1)))
        for name in names:
            root = etree.fromstring(z.read(name))
            # 문서 순서대로 단락 순회. 표는 단락 안에 인라인으로 들어있고, 셀 안에도
            # 다시 단락이 있으므로(.iter는 중첩까지 훑음) '셀 안쪽 단락'은 건너뛰고
            # 표 파서(_parse_table)에 맡긴다 → 셀 텍스트 중복 출력 방지.
            for p in root.iter(_HP + 'p'):
                if next(p.iterancestors(_HP + 'tc'), None) is not None:
                    continue  # 표 셀 내부 단락 → 표 블록에서 처리됨
                tbls = p.findall('.//' + _HP + 'tbl')
                if tbls:
                    for tbl in tbls:
                        blocks.append({'type': 'table',
                                       'rows': _parse_table(tbl)})
                else:
                    txt = _xml_text_of_para(p)
                    if txt.strip():
                        blocks.append({'type': 'para', 'text': txt})
    text_lines = []
    for b in blocks:
        if b['type'] == 'para':
            text_lines.append(b['text'])
        else:
            for r in b['rows']:
                text_lines.append(' | '.join(c.replace('\n', ' ') for c in r))
    return {'format': 'hwpx', 'blocks': blocks, 'text': '\n'.join(text_lines)}

# ───────────────────────── 디스패치 ─────────────────────────
def read(path):
    fmt = detect_format(path)
    return read_hwpx(path) if fmt == 'hwpx' else read_hwp(path)

def extract_text(path):
    return read(path)['text']

def iter_tables(doc):
    """read_hwpx 결과에서 표만 행렬 리스트로 추출."""
    return [b['rows'] for b in doc.get('blocks', []) if b.get('type') == 'table']

# ───────────────────────── CLI ─────────────────────────
def _main(argv):
    args = [a for a in argv if not a.startswith('-')]
    as_json = '--json' in argv
    out_path = None
    if '-o' in argv:
        out_path = argv[argv.index('-o') + 1]
        args = [a for a in args if a != out_path]
    if not args:
        print(__doc__); return 1
    chunks = []
    for f in args:
        doc = read(f)
        if as_json:
            chunks.append(json.dumps({'file': f, **doc}, ensure_ascii=False, indent=2))
        else:
            chunks.append('=' * 70 + '\n' + f + '  [%s]\n' % doc['format']
                          + '=' * 70 + '\n' + doc['text'])
    result = '\n\n'.join(chunks)
    if out_path:
        with io.open(out_path, 'w', encoding='utf-8') as fp:
            fp.write(result)
        print('[OK] %d파일 → %s (%d자)' % (len(args), out_path, len(result)))
    else:
        sys.stdout.buffer.write(result.encode('utf-8'))
        sys.stdout.buffer.write(b'\n')
    return 0

if __name__ == '__main__':
    sys.exit(_main(sys.argv[1:]))
