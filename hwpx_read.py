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
  extract_text(path, revisions='final') -> str   평문 추출. 교정추적 .hwpx는
                                                 'final'(적용)/'original'(거부)/'merge'(옛 동작)
  read_hwpx(path)               -> dict           {'format','blocks':[para|table],'text'}
  read_hwp(path)                -> dict           {'format','paragraphs','text'}
  read_changes(path)            -> dict           교정추적 분리: {has_changes,changes,
                                                 original,final,authors,orphans,...} (.hwpx 전용)
  iter_tables(doc)              -> list[list[list[str]]]   표만 골라 행렬로
  check_refs(path)              -> dict           참조 무결성(없는 스타일/속성 ID 참조)
                                                 검사. {} 면 정상. 복구는 hwpx_edit.repair()
  check_tables(path)            -> list[str]      표 격자 손상(cellAddr 누락·ragged 격자)
                                                 검사. [] 면 정상. 복구는 hwpx_edit.repair()
  validate(path, pre_bake=True) -> (ok, errors)   손상 사전검사(구조+참조+표격자 무결성 포함)

CLI
  python hwpx_read.py 파일.hwp [파일2.hwpx ...] [--json] [-o out.txt]
  python hwpx_read.py 파일.hwpx --changes      # 삽입/삭제 변경 목록
  python hwpx_read.py 파일.hwpx --validate     # baking 전 손상 검사(참조 무결성 포함)
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
_HH = '{http://www.hancom.co.kr/hwpml/2011/head}'

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

# ───────────────────── 변경 내용 추적 (track changes) ─────────────────────
# HWPX는 교정추적을 다음과 같이 저장한다.
#   · header.xml : <hh:trackChange type="Insert|Delete|CharShape|ParaShape"
#                    date=.. authorID=.. id=../>, <hh:trackChangeAuthor name=.. id=../>
#   · section*.xml : 본문에 인라인 제어 마커
#       <hp:insertBegin Id=.. TcId=../> … <hp:insertEnd Id=.. TcId=../>
#       <hp:deleteBegin Id=.. TcId=../> … <hp:deleteEnd Id=.. TcId=../>
#     사이의 hp:t 텍스트가 그 구간의 삽입/삭제 대상이다(TcId가 header 정의와 연결).
#
# ⚠ 함정(실측): .hwp → .hwpx 변환(hwpx_convert, 한컴 COM)이 **종료 마커를 누락**시키는
#   경우가 있다(insertBegin 135 / insertEnd 134 식으로 1개 어긋남). 종료 없는 구간(orphan)을
#   그대로 두면 그 뒤 본문 전체가 삭제/삽입으로 오분류되어 "참고문헌이 통째로 사라짐" 같은
#   파국적 오판이 난다. 아래 구현은 (1) 정상적으로 닫힌 구간만 신뢰하고 (2) orphan 구간은
#   해당 '문단 끝'에서 강제로 닫아 피해를 1문단으로 한정하며 (3) orphans 개수를 보고한다.
#   orphans>0이면 결과를 한컴 최종본과 대조 검증할 것(hwpx_convert.accepted_text 참고).
#
# ⚠ 평문 추출의 함정: extract_text(... revisions='merge')나 .hwp 바이너리 리더는 삽입+삭제
#   텍스트를 한 흐름으로 이어붙인다 → 예: 참고문헌 재번호 중 옛 번호[7]가 지워지고 새 번호[34]가
#   들어가면 "[347]"처럼 보인다(원고 오류가 아니라 추적 흔적). 깨끗한 본문은 revisions='final'.

def _read_trackchange_meta(z):
    """header.xml → (changes{id:{type,date,authorID}}, authors{id:name})."""
    from lxml import etree
    changes, authors = {}, {}
    try:
        root = etree.fromstring(z.read('Contents/header.xml'))
    except KeyError:
        return changes, authors
    for a in root.iter(_HH + 'trackChangeAuthor'):
        authors[a.get('id')] = a.get('name')
    for tc in root.iter(_HH + 'trackChange'):
        changes[tc.get('id')] = {'type': tc.get('type'),
                                 'date': tc.get('date'),
                                 'authorID': tc.get('authorID')}
    return changes, authors

def _merge_changes(items):
    """인접한 같은 종류·같은 TcId·같은 문단의 변경 텍스트만 하나로 합친다
    (문단을 넘어 합치지 않음 → 서로 다른 위치의 편집이 뭉치는 것 방지)."""
    merged = []
    for c in items:
        if (merged and merged[-1]['type'] == c['type']
                and merged[-1]['tc'] == c['tc']
                and merged[-1]['para'] == c['para']):
            merged[-1]['text'] += c['text']
        else:
            merged.append(dict(c))
    return merged

def read_changes(path, final_text=None):
    """HWPX의 교정추적(삽입/삭제)을 분리 추출한다(한컴 불필요).

    반환 dict:
      has_changes : bool                삽입/삭제 마커 존재 여부
      authors     : {id: name}          변경 작성자
      meta_counts : {type: n}           header 기준 변경 종류별 개수(CharShape 등 포함)
      changes     : [ {type:'insert'|'delete', text, tc, author, date, para} ]
      original    : str                 변경 거부본(=일반+삭제 텍스트)
      final       : str                 변경 적용본(=일반+삽입 텍스트)
      orphans     : int                 종료 마커 누락 구간 수(>0이면 검증 권장)

    final_text(선택): 한컴 GetTextFile로 얻은 '권위 있는 최종(적용)본' 평문
      (hwpx_convert.accepted_text). 주면 orphan으로 인한 오판을 보정한다 →
      · 최종본에 그대로 남아있는 텍스트는 '삭제'로 보지 않는다(거짓 삭제 제거).
      · 반환 final 필드를 이 권위 텍스트로 대체한다.
      (단 삽입은 종료마커가 누락된 경우 원본 부재로 완전 복원이 불가하여 best-effort.)

    .hwp는 미지원 → hwpx_convert.hwp_to_hwpx로 변환 후 사용.
    """
    from lxml import etree
    if detect_format(path) != 'hwpx':
        raise ValueError('read_changes는 .hwpx 전용입니다. '
                         'hwpx_convert.hwp_to_hwpx(path)로 변환 후 사용하세요.')
    with zipfile.ZipFile(path) as z:
        meta, authors = _read_trackchange_meta(z)
        names = sorted((n for n in z.namelist()
                        if re.match(r'Contents/section\d+\.xml$', n)),
                       key=lambda n: int(re.search(r'(\d+)', n).group(1)))
        tokens = []   # ('ib'|'ie'|'db'|'de', Id, TcId) | ('t', text) | ('p',)
        for name in names:
            ctx = etree.iterparse(io.BytesIO(z.read(name)), events=('start', 'end'))
            for ev, el in ctx:
                tag = el.tag.rsplit('}', 1)[-1]
                if ev == 'start':
                    if tag == 'insertBegin': tokens.append(('ib', el.get('Id'), el.get('TcId')))
                    elif tag == 'insertEnd': tokens.append(('ie', el.get('Id'), el.get('TcId')))
                    elif tag == 'deleteBegin': tokens.append(('db', el.get('Id'), el.get('TcId')))
                    elif tag == 'deleteEnd': tokens.append(('de', el.get('Id'), el.get('TcId')))
                elif tag == 't' and el.text:
                    tokens.append(('t', el.text))
                elif tag == 'p':
                    tokens.append(('p',))

    # pass1 — orphan(종료 마커 없는 Begin) 탐지. Id 재사용 대비 type별 스택 매칭.
    ins_stack, del_stack, orphan = [], [], set()
    for k, tok in enumerate(tokens):
        if tok[0] == 'ib': ins_stack.append((tok[1], k))
        elif tok[0] == 'db': del_stack.append((tok[1], k))
        elif tok[0] == 'ie':
            for j in range(len(ins_stack) - 1, -1, -1):
                if ins_stack[j][0] == tok[1]: ins_stack.pop(j); break
        elif tok[0] == 'de':
            for j in range(len(del_stack) - 1, -1, -1):
                if del_stack[j][0] == tok[1]: del_stack.pop(j); break
    for _id, k in ins_stack + del_stack:
        orphan.add(k)

    # pass2 — 분류. open_* = [Id, TcId, is_orphan].
    # orphan 처리: 오라클(final_text) 없으면 문단 끝에서 강제 종료(cap; 피해 1문단 한정).
    # 오라클 있으면 종료하지 않고(bleed) 흘려보낸 뒤, 최종본에 남은 텍스트의 거짓 삭제를
    # 뒤에서 걷어낸다 → 종료마커가 누락된 삽입 구간까지 복원된다.
    cap_orphans = (final_text is None)
    open_ins, open_del = [], []
    para = 1
    raw_changes, original, final = [], [], []
    for k, tok in enumerate(tokens):
        kind = tok[0]
        if kind == 'ib': open_ins.append([tok[1], tok[2], k in orphan])
        elif kind == 'db': open_del.append([tok[1], tok[2], k in orphan])
        elif kind == 'ie':
            for j in range(len(open_ins) - 1, -1, -1):
                if open_ins[j][0] == tok[1]: open_ins.pop(j); break
        elif kind == 'de':
            for j in range(len(open_del) - 1, -1, -1):
                if open_del[j][0] == tok[1]: open_del.pop(j); break
        elif kind == 'p':
            if cap_orphans:
                open_ins = [s for s in open_ins if not s[2]]   # orphan 종료(cap)
                open_del = [s for s in open_del if not s[2]]
            original.append('\n'); final.append('\n')
            para += 1
        elif kind == 't':
            txt = tok[1]
            if open_del:                       # 삭제(거부본에만 남음)
                tc = open_del[-1][1]; m = meta.get(tc, {})
                raw_changes.append({'type': 'delete', 'text': txt, 'tc': tc, 'para': para,
                                    'author': authors.get(m.get('authorID')), 'date': m.get('date')})
                original.append(txt)
            elif open_ins:                     # 삽입(적용본에만 남음)
                tc = open_ins[-1][1]; m = meta.get(tc, {})
                raw_changes.append({'type': 'insert', 'text': txt, 'tc': tc, 'para': para,
                                    'author': authors.get(m.get('authorID')), 'date': m.get('date')})
                final.append(txt)
            else:                              # 일반(양쪽 공통)
                original.append(txt); final.append(txt)

    meta_counts = {}
    for v in meta.values():
        meta_counts[v['type']] = meta_counts.get(v['type'], 0) + 1
    changes = _merge_changes(raw_changes)
    final_str = ''.join(final)
    if final_text is not None:
        # 권위 최종본으로 거짓 삭제 보정: 최종본에 남은 텍스트는 삭제가 아니다.
        cmp = re.sub(r'\s+', '', final_text)
        changes = [c for c in changes
                   if not (c['type'] == 'delete'
                           and re.sub(r'\s+', '', c['text']) in cmp)]
        final_str = final_text
    return {'has_changes': any(t[0] in ('ib', 'db') for t in tokens),
            'authors': authors, 'meta_counts': meta_counts,
            'changes': changes, 'orphans': len(orphan),
            'original': ''.join(original), 'final': final_str}

# ───────────────────────── 디스패치 ─────────────────────────
def read(path):
    fmt = detect_format(path)
    return read_hwpx(path) if fmt == 'hwpx' else read_hwp(path)

def extract_text(path, revisions='final'):
    """평문 추출. 교정추적이 있는 .hwpx는 revisions로 어느 판본을 뽑을지 정한다.
       'final'    : 변경 적용본(기본) — 깨끗한 현재 본문
       'original' : 변경 거부본 — 편집 전 원문
       'merge'    : 마커 무시·삽입+삭제 동시 출력(옛 동작; 번호가 겹쳐 보일 수 있음)
    .hwp(바이너리)는 항상 merge 동작."""
    if revisions != 'merge' and detect_format(path) == 'hwpx':
        ch = read_changes(path)
        if ch['has_changes']:
            return ch['original'] if revisions == 'original' else ch['final']
    return read(path)['text']

def iter_tables(doc):
    """read_hwpx 결과에서 표만 행렬 리스트로 추출."""
    return [b['rows'] for b in doc.get('blocks', []) if b.get('type') == 'table']

# ──────────────────── 참조 무결성 (dangling IDRef) ────────────────────
# 본문 참조 attr → 머리부(header) 항목 태그(hh:)
_REF_TO_ELEM = {'charPrIDRef': 'charPr', 'paraPrIDRef': 'paraPr',
                'styleIDRef': 'style', 'borderFillIDRef': 'borderFill'}

def check_refs(path):
    """본문(section)이 참조하는 스타일/속성 ID가 머리부(header)에 모두 정의돼
    있는지 검사한다. 외부/병합 hwpx가 'XML·표는 정상인데 한글이 열기를 거부'하는
    대표 원인 — dangling IDRef(예: section의 paraPrIDRef="21"인데 header엔 0~20만
    존재) — 를 잡는다(validate가 못 보던 층).
    반환: {refattr: {'defined_max':int|-1, 'missing':sorted[int]}} — 누락이 있는
    항목만(정상이면 빈 dict). 복구는 hwpx_edit.repair()."""
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        header = (z.read('Contents/header.xml').decode('utf-8', 'replace')
                  if 'Contents/header.xml' in names else '')
        sect = '\n'.join(z.read(n).decode('utf-8', 'replace') for n in names
                         if re.match(r'Contents/section\d+\.xml$', n))
    report = {}
    for attr, elem in _REF_TO_ELEM.items():
        defined = set(int(m) for m in
                      re.findall(r'<hh:%s\b[^>]*?\bid="(\d+)"' % elem, header))
        used = set(int(m) for m in re.findall(r'\b%s="(\d+)"' % attr, sect))
        missing = sorted(used - defined)
        if missing:
            report[attr] = {'defined_max': (max(defined) if defined else -1),
                            'missing': missing}
    return report

# ──────────────────── 표 격자 손상 (한글 열기 무한루프) ────────────────────
def check_tables(path):
    """표 격자 손상을 검사한다. 외부/병합 hwpx에서 흔하며, 한글이 열 때 표 레이아웃을
    무한 반복(파일이 '안 열림')하게 만든다. 잡는 항목:
      · 셀 <hp:cellAddr>(행·열 좌표) 누락        ← 가장 치명적
      · 행마다 열 너비가 다른 ragged 격자(병합 아님)
      · 표 sz.height 가 행 높이 합과 불일치
    반환: list[str] 문제 설명(정상이면 []). 복구는 hwpx_edit.repair()."""
    from lxml import etree
    from collections import Counter
    probs = []
    def simple_w(tr):
        ws = []
        for tc in tr.findall(_HP+'tc'):
            sp = tc.find(_HP+'cellSpan'); cs = tc.find(_HP+'cellSz')
            if cs is None:
                return None
            if sp is not None and sp.get('colSpan') and int(sp.get('colSpan')) != 1:
                return None
            ws.append(int(cs.get('width')))
        return tuple(ws)
    with zipfile.ZipFile(path) as z:
        names = sorted((n for n in z.namelist()
                        if re.match(r'Contents/section\d+\.xml$', n)),
                       key=lambda n: int(re.search(r'(\d+)', n).group(1)))
        for nm in names:
            root = etree.fromstring(z.read(nm))
            for ti, tbl in enumerate(root.iter(_HP+'tbl')):
                trs = tbl.findall(_HP+'tr')
                miss = sum(1 for tr in trs for tc in tr.findall(_HP+'tc')
                           if tc.find(_HP+'cellAddr') is None)
                if miss:
                    probs.append('표#%d: 셀 cellAddr 누락 %d개(한글 열기 무한루프)' % (ti, miss))
                pats = Counter(s for s in (simple_w(tr) for tr in trs) if s)
                if len(pats) > 1:
                    probs.append('표#%d: 행별 열너비 불일치(ragged 격자, %d패턴)' % (ti, len(pats)))
    return probs

# ───────────────────────── 사전 검증 ─────────────────────────
def validate(path, pre_bake=True):
    """한컴으로 열기 전, .hwpx의 알려진 손상 신호를 정적 검사한다.
    (HWPX 기술문서 v5의 교훈 반영) → (ok: bool, errors: list[str]) 반환.

    구조적 손상(항상 검사) — 한글이 손상으로 인식하는 신호:
      · 표 id 중복(객체 식별 실패)
      · 한 표 안에 cellAddr (0,0)이 둘 이상(셀 주소 미설정)
      · 첫 행의 셀너비 합 ≠ 표 너비(구조 불일치)
      · 참조 무결성: 본문이 머리부에 없는 paraPr/charPr/style/borderFill ID를
        참조(dangling IDRef) → 한글이 통째로 열기 거부. 복구는 hwpx_edit.repair().
      · 표 격자 손상: 셀 cellAddr(행·열 좌표) 누락·행별 열너비 불일치(ragged) →
        한글이 열 때 표 레이아웃 무한 루프(파일이 안 열림). 복구는 hwpx_edit.repair().

    생성단계 규칙(pre_bake=True일 때만) — '아직 baking 안 한 raw 생성물'에만 해당:
      · 비어있지 않은 linesegarray(빈 태그여야 baking 시 자동 계산 → 안 그러면 글자 겹침)
      · 표 pos treatAsChar="0"(이 스킬 생성물은 항상 인라인)

    ⚠ 한컴이 '저장(baking)'한 정상 파일은 linesegarray가 채워져 있고 floating 표도
       정상이다. 따라서 baking된 파일/일반 외부 파일을 검사할 땐 pre_bake=False로 호출해
       구조적 손상만 본다(그렇지 않으면 오탐).
    """
    from lxml import etree
    errors = []
    tbl_ids = []
    bad_lsa = bad_tac = 0
    with zipfile.ZipFile(path) as z:
        names = sorted((n for n in z.namelist()
                        if re.match(r'Contents/section\d+\.xml$', n)),
                       key=lambda n: int(re.search(r'(\d+)', n).group(1)))
        for name in names:
            root = etree.fromstring(z.read(name))
            for tbl in root.iter(_HP + 'tbl'):
                tid = tbl.get('id')
                if tid is not None:
                    tbl_ids.append(tid)
                pos = tbl.find(_HP + 'pos')
                if pos is not None and pos.get('treatAsChar') == '0':
                    bad_tac += 1
                # 이 표의 직접 행/셀만(중첩표 제외)
                trs = tbl.findall(_HP + 'tr')
                zeros = 0
                for tr in trs:
                    for tc in tr.findall(_HP + 'tc'):
                        ca = tc.find(_HP + 'cellAddr')
                        if ca is not None and ca.get('colAddr') == '0' and ca.get('rowAddr') == '0':
                            zeros += 1
                if zeros > 1:
                    errors.append('표 id=%s: cellAddr (0,0) 셀이 %d개(표당 1개여야 함)' % (tid, zeros))
                # 열너비 합 vs 표너비
                sz = tbl.find(_HP + 'sz')
                if sz is not None and sz.get('width') and trs:
                    tw = int(sz.get('width'))
                    first = trs[0].findall(_HP + 'tc')
                    csum = sum(int(tc.find(_HP + 'cellSz').get('width'))
                               for tc in first if tc.find(_HP + 'cellSz') is not None)
                    if csum and csum != tw:
                        errors.append('표 id=%s: 첫 행 셀너비 합 %d ≠ 표 너비 %d' % (tid, csum, tw))
            if pre_bake:
                for lsa in root.iter(_HP + 'linesegarray'):
                    if len(lsa):  # 자식 lineseg가 있으면 비어있지 않음
                        bad_lsa += 1
    dup = sorted({i for i in tbl_ids if tbl_ids.count(i) > 1})
    if dup:
        errors.insert(0, '표 id 중복: %s' % dup)
    if pre_bake and bad_lsa:
        errors.append('linesegarray에 내용 있음: %d개(raw 생성물은 빈 태그여야 함; baking된 파일이면 pre_bake=False)' % bad_lsa)
    if pre_bake and bad_tac:
        errors.append('treatAsChar="0" 표: %d개(이 스킬 생성물은 인라인; 외부 floating 표면 pre_bake=False)' % bad_tac)
    # 참조 무결성(항상 검사): 본문이 머리부에 없는 스타일/속성 ID를 참조하면 한글이 열기를 거부
    for attr, info in check_refs(path).items():
        errors.append('참조 무결성: %s가 머리부에 없는 id %s 참조(정의된 최대 id=%d) → hwpx_edit.repair로 복구'
                      % (attr, info['missing'], info['defined_max']))
    # 표 격자 손상(항상 검사): cellAddr 누락·ragged 격자 → 한글 열기 무한루프
    for p in check_tables(path):
        errors.append('%s → hwpx_edit.repair로 복구' % p)
    return (len(errors) == 0, errors)

# ───────────────────────── CLI ─────────────────────────
def _main(argv):
    args = [a for a in argv if not a.startswith('-')]
    as_json = '--json' in argv
    out_path = None
    if '-o' in argv:
        out_path = argv[argv.index('-o') + 1]
        args = [a for a in args if a != out_path]
    if not args:
        sys.stdout.buffer.write((__doc__ or '').encode('utf-8'))  # cp949 콘솔 크래시 방지
        return 1
    if '--changes' in argv:
        rc = 0
        for f in args:
            try:
                ch = read_changes(f)
            except ValueError as e:
                sys.stdout.buffer.write(('[SKIP] %s — %s\n' % (f, e)).encode('utf-8')); rc = 1; continue
            lines = ['=' * 70, '%s  [변경 추적]' % f, '=' * 70]
            if not ch['has_changes']:
                lines.append('교정추적 변경 없음.')
            else:
                who = ', '.join('%s(%s)' % (n, i) for i, n in ch['authors'].items())
                lines.append('작성자: %s | 변경종류: %s' % (who, ch['meta_counts']))
                if ch['orphans']:
                    lines.append('⚠ 종료 마커 누락 구간 %d개 — 결과를 한컴 최종본과 대조 권장' % ch['orphans'])
                ins = [c for c in ch['changes'] if c['type'] == 'insert']
                dele = [c for c in ch['changes'] if c['type'] == 'delete']
                lines.append('── 삽입 %d건 ──' % len(ins))
                for c in ins:
                    lines.append('  [+] (문단 %s) %s' % (c['para'], c['text'][:120].replace('\n', ' ')))
                lines.append('── 삭제 %d건 ──' % len(dele))
                for c in dele:
                    lines.append('  [-] (문단 %s) %s' % (c['para'], c['text'][:120].replace('\n', ' ')))
            sys.stdout.buffer.write(('\n'.join(lines) + '\n').encode('utf-8'))
        return rc
    if '--validate' in argv:
        rc = 0
        pre_bake = '--baked' not in argv  # baking된/외부 파일은 --baked로 구조검사만
        for f in args:
            ok, errs = validate(f, pre_bake=pre_bake)
            line = ('[OK] ' if ok else '[FAIL] ') + f
            sys.stdout.buffer.write((line + '\n').encode('utf-8'))
            for e in errs:
                sys.stdout.buffer.write(('   - ' + e + '\n').encode('utf-8'))
            rc = rc or (0 if ok else 1)
        return rc
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
