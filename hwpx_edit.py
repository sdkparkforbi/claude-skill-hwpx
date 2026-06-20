#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HWPX 편집 엔진 — 기존 .hwpx를 열어 부분 수정 후 재패키징.

생성(hwpxgen)·읽기(hwpx_read)에 이어 '편집' 방향. 제안서 양식처럼 골격은
그대로 두고 내용만 채워넣는 작업에 쓴다. 한컴 없이 순수 파이썬으로 동작하지만,
저장 후에는 hwpx_bake.py로 baking(레이아웃 재계산)하는 것을 권장한다.

핵심 함정: hwpx 본문의 한 단락 텍스트는 여러 <hp:t> run으로 쪼개져 있다.
플레이스홀더가 run 경계를 넘나들면 단순 치환이 실패하므로, '단락 단위로
run 텍스트를 이어붙여 치환'하는 2-pass 전략을 쓴다.
  1) run 하나 안에 전부 들어있으면 그 run만 치환(글자모양 보존).
  2) run 경계에 걸치면 단락의 첫 run에 합쳐 넣고 나머지 run은 비움.

주요 API
  ed = HwpxEditor("양식.hwpx")
  ed.replace("{{과제명}}", "생성형 AI 인재양성")     # 단일 치환
  ed.replace_map({"{{이름}}":"홍길동", "{{금액}}":"5,000,000"})
  n = ed.count("{{과제명}}")                          # 남은 개수 확인
  ed.append_table_row(table_index=0, cells=["항목","값",...])  # 표 끝에 행 추가
  ed.save("작성본.hwpx")

복구 — 한글이 "열 수 없는 문서"로 거부하거나 열 때 무한 루프하는 hwpx 고치기(한컴 불필요)
  import hwpx_edit
  hwpx_edit.repair("안열리는.hwpx")                   # 참조 무결성 + 표 격자 손상을 함께 복구(+.bak)
  #  ├ 참조: 없는 paraPr/charPr/style/borderFill ID → 머리부 정의 복제로 메움
  #  └ 표 : 셀 cellAddr(행·열 좌표) 누락·행별 열너비 불일치(ragged)·표 높이 불일치 보정
  # 직접 제어: ed.missing_refs()/ed.repair_refs() , ed.fix_table_grid() → ed.save()
  # (사전 점검: hwpx_read.check_refs / hwpx_read.check_tables / hwpx_read.validate 가 함께 잡는다)
  # ⚠ 복구 후 python hwpx_bake.py 로 한 번 열어 레이아웃을 굳히는 것을 권장

CLI
  python hwpx_edit.py 양식.hwpx -o 결과.hwpx --set "{{이름}}=홍길동" --set "{{금액}}=500"
  python hwpx_edit.py 양식.hwpx --map 치환표.json -o 결과.hwpx
  python hwpx_edit.py 안열리는.hwpx --repair          # 참조 무결성 복구(덮어쓰기+.bak)
"""
import sys, os, re, json, zipfile, io
from lxml import etree

_HP = '{http://www.hancom.co.kr/hwpml/2011/paragraph}'
HWPU_PER_MM = 283.465   # 1mm = 283.465 HWPUNIT (셀 여백 mm 변환용)

class HwpxEditor:
    def __init__(self, path):
        self.path = path
        self._order = []            # zip 내 파일 순서 보존
        self._raw = {}              # name -> bytes (수정 안 하는 파일)
        self._sections = {}         # name -> lxml root (수정 대상)
        with zipfile.ZipFile(path) as z:
            for item in z.infolist():
                data = z.read(item.filename)
                self._order.append(item.filename)
                if re.match(r'Contents/section\d+\.xml$', item.filename):
                    self._sections[item.filename] = etree.fromstring(data)
                else:
                    self._raw[item.filename] = data
        self._sec_names = sorted(self._sections,
                                 key=lambda n: int(re.search(r'(\d+)', n).group(1)))

    # ── 조회 ──
    def _iter_paras(self):
        for name in self._sec_names:
            for p in self._sections[name].iter(_HP + 'p'):
                yield p

    @staticmethod
    def _own_ts(p):
        """p에 '직접' 속한 hp:t만 반환. 중첩 표 셀 안쪽 run은 제외
        (그렇지 않으면 표를 감싼 단락이 모든 셀 텍스트까지 들고 와 중복/붕괴 발생)."""
        return [t for t in p.iter(_HP + 't')
                if next(t.iterancestors(_HP + 'p'), None) is p]

    def count(self, needle):
        c = 0
        for p in self._iter_paras():
            joined = ''.join((t.text or '') for t in self._own_ts(p))
            c += joined.count(needle)
        return c

    def text(self):
        """현재 상태의 평문(검증용)."""
        out = []
        for p in self._iter_paras():
            out.append(''.join((t.text or '') for t in self._own_ts(p)))
        return '\n'.join(x for x in out if x.strip())

    # ── 치환 ──
    def replace(self, old, new):
        """문서 전체에서 old→new. 치환 횟수 반환."""
        return self._replace_many({old: new})

    def replace_map(self, mapping):
        return self._replace_many(dict(mapping))

    def _replace_many(self, mapping):
        total = 0
        for p in self._iter_paras():
            ts = self._own_ts(p)   # 직접 run만 (셀 내부는 해당 셀 단락에서 별도 처리)
            if not ts:
                continue
            # pass 1: run 단위(글자모양 보존)
            for t in ts:
                if not t.text:
                    continue
                for old, new in mapping.items():
                    if old and old in t.text:
                        cnt = t.text.count(old)
                        if cnt:
                            t.text = t.text.replace(old, new)
                            total += cnt
            # pass 2: run 경계에 걸친 경우 단락 합치기
            joined = ''.join((t.text or '') for t in ts)
            need = {o: n for o, n in mapping.items() if o and o in joined}
            if need:
                for old, new in need.items():
                    cnt = joined.count(old)
                    if cnt:
                        joined = joined.replace(old, new)
                        total += cnt
                ts[0].text = joined
                for t in ts[1:]:
                    t.text = ''
        return total

    # ── 셀 내용 교체 ──
    def _tables(self):
        tbls = []
        for name in self._sec_names:
            tbls.extend(self._sections[name].iter(_HP + 'tbl'))
        return tbls

    def set_cell(self, table_index, row, col, text):
        """table_index번째 표에서 cellAddr=(col,row)인 셀의 텍스트를 교체.
        lxml 트리 조작이라 정규식 오프셋 밀림 문제가 없다(v5 사례⑤ 대안)."""
        tbls = self._tables()
        if table_index >= len(tbls):
            raise IndexError('표 인덱스 범위 초과: %d (총 %d개)' % (table_index, len(tbls)))
        tbl = tbls[table_index]
        target = None
        for tr in tbl.findall(_HP + 'tr'):           # 직접 행/셀만(중첩표 제외)
            for tc in tr.findall(_HP + 'tc'):
                ca = tc.find(_HP + 'cellAddr')
                if ca is not None and ca.get('colAddr') == str(col) \
                        and ca.get('rowAddr') == str(row):
                    target = tc; break
            if target is not None:
                break
        if target is None:
            raise ValueError('셀을 찾지 못함: 표 %d (col=%d, row=%d)' % (table_index, col, row))
        sub = target.find(_HP + 'subList')
        ps = sub.findall(_HP + 'p') if sub is not None else []
        if not ps:
            raise ValueError('셀에 단락이 없습니다')
        # 첫 단락 첫 run에 text, 나머지 run/단락은 비움(셀을 단일 값으로).
        # 함정(v8): 양식의 '빈' 셀은 <hp:run>은 있어도 <hp:t>가 없는 경우가 많다.
        # 이때 _own_ts(p)는 []라 그냥 두면 글자가 안 들어간다 → t를 생성해 넣는다.
        first_ts = self._own_ts(ps[0])
        if first_ts:
            first_ts[0].text = text
            for extra in first_ts[1:]:
                extra.text = ''
        else:
            runs = ps[0].findall(_HP + 'run')        # 단락의 직속 run만
            if not runs:                              # run조차 없으면 생성
                run = etree.SubElement(ps[0], _HP + 'run')
                run.set('charPrIDRef', self._default_charpr())
                runs = [run]
            t = etree.SubElement(runs[0], _HP + 't')  # 빈 run에 t 신설
            t.text = text
        for p in ps[1:]:
            for t in self._own_ts(p):
                t.text = ''
        # 함정(v8): 텍스트가 길어져 여러 줄로 wrap되어도, 원래 '빈 칸'이 갖고 있던
        # lineseg 1개가 남아 있으면 한글이 그 1줄에 모든 줄을 겹쳐 그린다(줄간격 붕괴).
        # linesegarray를 비워두면 baking 시 한글이 줄 수를 다시 계산한다.
        for p in ps:
            lsa = p.find(_HP + 'linesegarray')
            if lsa is not None:
                for seg in lsa.findall(_HP + 'lineseg'):
                    lsa.remove(seg)
        return True

    def _default_charpr(self):
        """문서에서 흔히 쓰인 charPrIDRef를 추정(없으면 '0')."""
        from collections import Counter
        c = Counter()
        for name in self._sec_names:
            for r in self._sections[name].iter(_HP + 'run'):
                ref = r.get('charPrIDRef')
                if ref is not None:
                    c[ref] += 1
        return c.most_common(1)[0][0] if c else '0'

    def check_option(self, table_index, row, col, *, checked=True,
                     box_empty='', box_checked=''):
        """양식 체크박스 토글. 한글 양식은 Wingdings 글리프
        (빈 칸)·(체크)를 라벨과 같은 run 안에 둔다.
        해당 셀에서 빈칸↔체크 글리프를 서로 바꿔 표시한다."""
        tbls = self._tables()
        tbl = tbls[table_index]
        old, new = (box_empty, box_checked) if checked else (box_checked, box_empty)
        n = 0
        for tr in tbl.findall(_HP + 'tr'):
            for tc in tr.findall(_HP + 'tc'):
                ca = tc.find(_HP + 'cellAddr')
                if ca is None or ca.get('colAddr') != str(col) \
                        or ca.get('rowAddr') != str(row):
                    continue
                for t in tc.iter(_HP + 't'):
                    if t.text and old in t.text:
                        t.text = t.text.replace(old, new)
                        n += 1
        return n

    # ── 표 행 추가 ──
    def append_table_row(self, table_index, cells):
        """table_index번째 표의 마지막 행을 복제해 cells 텍스트로 채워 추가."""
        tbls = self._tables()
        if table_index >= len(tbls):
            raise IndexError('표 인덱스 범위 초과: %d (총 %d개)' % (table_index, len(tbls)))
        import copy
        tbl = tbls[table_index]
        trs = tbl.findall(_HP + 'tr')
        if not trs:
            raise ValueError('표에 행이 없습니다')
        new_tr = copy.deepcopy(trs[-1])
        tcs = new_tr.findall(_HP + 'tc')
        for i, tc in enumerate(tcs):
            val = cells[i] if i < len(cells) else ''
            t_nodes = list(tc.iter(_HP + 't'))
            if t_nodes:
                t_nodes[0].text = val
                for extra in t_nodes[1:]:
                    extra.text = ''
            # cellAddr rowAddr 증가
            addr = tc.find(_HP + 'cellAddr')
            if addr is not None and addr.get('rowAddr') is not None:
                addr.set('rowAddr', str(int(addr.get('rowAddr')) + 1))
        tbl.append(new_tr)
        # rowCnt 갱신
        if tbl.get('rowCnt') is not None:
            tbl.set('rowCnt', str(int(tbl.get('rowCnt')) + 1))
        return True

    # ── 표 꾸미기: 셀 여백 / 글자객체 해제 / 표 늘이기 ──
    @staticmethod
    def _norm_margin(m):
        """여백 인자 → (left,right,top,bottom) HWPUNIT. int=네 변, (lr,tb), (l,r,t,b). 1mm≈283."""
        if isinstance(m, (int, float)):
            v = int(round(m)); return (v, v, v, v)
        m = list(m)
        if len(m) == 2:
            lr, tb = int(round(m[0])), int(round(m[1])); return (lr, lr, tb, tb)
        if len(m) == 4:
            return tuple(int(round(x)) for x in m)
        raise ValueError('margin은 int 또는 (lr,tb) 또는 (l,r,t,b)')

    def set_cell_margin(self, table_index, margin, only=None, unit='hwpu'):
        """표의 모든(또는 일부) 셀 안쪽 여백 지정. 기존 양식 셀이 답답할 때 숨통을 틔운다.
        margin: int/(lr,tb)/(l,r,t,b). unit='mm'이면 mm로 입력(예: 0.49 → 한컴 표시 0.49mm), 기본 'hwpu'(1mm≈283).
        only: [(row,col),...] 지정 시 해당 셀만."""
        if unit == 'mm':
            margin = (int(round(margin * HWPU_PER_MM)) if isinstance(margin, (int, float))
                      else tuple(int(round(x * HWPU_PER_MM)) for x in margin))
        l, r, t, b = self._norm_margin(margin)
        tbl = self._tables()[table_index]
        for tc in tbl.findall(_HP + 'tr/' + _HP + 'tc'):
            ca = tc.find(_HP + 'cellAddr')
            if only is not None:
                if ca is None:
                    continue
                if (int(ca.get('rowAddr')), int(ca.get('colAddr'))) not in only:
                    continue
            tc.set('hasMargin', '1')                       # 셀 자체 여백 사용(0이면 표 기본값 무시됨)
            cm = tc.find(_HP + 'cellMargin')
            if cm is None:
                cm = etree.SubElement(tc, _HP + 'cellMargin')  # 자식 끝(cellSz 뒤)에 추가
            cm.set('left', str(l)); cm.set('right', str(r)); cm.set('top', str(t)); cm.set('bottom', str(b))
        im = tbl.find(_HP + 'inMargin')                    # 표 기본 여백도 동기화
        if im is not None:
            im.set('left', str(l)); im.set('right', str(r)); im.set('top', str(t)); im.set('bottom', str(b))
        return True

    def set_table_props(self, table_index, treat_as_char=False, pageBreak='CELL', repeat_header=True):
        """한컴 [표/셀 속성 ▸ 표] 대화상자를 그대로 코드로. 인자=None이면 해당 항목 미변경.
          treat_as_char : '글자처럼 취급' (False=글자객체로 인식 안 함 → treatAsChar=0)
          pageBreak     : '쪽 경계에서' — 'CELL'(나눔, 긴 표가 쪽경계서 안 잘림)/'TABLE'(통째 이동)/'NONE'(나누지 않음)
          repeat_header : '제목 줄 자동 반복' (True → repeatHeader=1)
        """
        tbl = self._tables()[table_index]
        if pageBreak is not None:
            tbl.set('pageBreak', pageBreak)
        if repeat_header is not None:
            tbl.set('repeatHeader', '1' if repeat_header else '0')
        if treat_as_char is not None:
            pos = tbl.find(_HP + 'pos')
            if pos is not None:
                pos.set('treatAsChar', '1' if treat_as_char else '0')
        return True

    # 하위호환 별칭
    set_table_float = set_table_props

    def stretch_table(self, table_index, total_height_mm=None, row_height=None):
        """표를 세로로 늘임(끌어당기기). total_height_mm=표 전체 높이(mm)를 행수로 등분, 또는 row_height(HWPUNIT) 직접.
        각 행 cellSz height를 키워 baking 후에도 줄지 않게 한다."""
        tbl = self._tables()[table_index]
        trs = tbl.findall(_HP + 'tr')
        if not trs:
            return False
        if total_height_mm is not None:
            row_height = int(round(total_height_mm * 283.465 / len(trs)))
        if row_height is None:
            raise ValueError('total_height_mm 또는 row_height 필요')
        for tc in tbl.iter(_HP + 'tc'):
            sz = tc.find(_HP + 'cellSz')
            if sz is not None:
                sz.set('height', str(row_height))
        tsz = tbl.find(_HP + 'sz')
        if tsz is not None:
            tsz.set('height', str(row_height * len(trs)))
        return True

    def set_cell_valign(self, table_index, valign='CENTER', only=None):
        """셀 세로 정렬 — 한컴 [표/셀 속성 ▸ 셀 ▸ 세로 정렬]. valign: 'TOP'/'CENTER'/'BOTTOM'.
        subList의 vertAlign 속성을 바꾼다. only=[(row,col),…] 지정 시 해당 셀만."""
        v = str(valign).upper()
        if v not in ('TOP', 'CENTER', 'BOTTOM'):
            raise ValueError("valign은 'TOP'/'CENTER'/'BOTTOM'")
        tbl = self._tables()[table_index]
        for tc in tbl.findall(_HP + 'tr/' + _HP + 'tc'):
            ca = tc.find(_HP + 'cellAddr')
            if only is not None:
                if ca is None:
                    continue
                if (int(ca.get('rowAddr')), int(ca.get('colAddr'))) not in only:
                    continue
            sub = tc.find(_HP + 'subList')
            if sub is not None:
                sub.set('vertAlign', v)
        return True

    # ── 복구: 참조 무결성(없는 스타일/속성 ID 참조) ──
    # 본문 참조 attr → 머리부 항목 태그(hh:) / 그 항목이 담긴 목록 태그
    _REF_ELEM = {'charPrIDRef': 'charPr', 'paraPrIDRef': 'paraPr',
                 'styleIDRef': 'style', 'borderFillIDRef': 'borderFill'}
    _ELEM_LIST = {'charPr': 'charProperties', 'paraPr': 'paraProperties',
                  'style': 'styles', 'borderFill': 'borderFills'}

    @staticmethod
    def _find_def_block(header, elem, eid):
        """머리부에서 <hh:elem id="eid" …>…</hh:elem>(또는 self-closing) 블록 문자열."""
        m = re.search(r'<hh:%s\b[^>]*?\bid="%s"[^>]*/>' % (elem, eid), header)
        if m:
            return m.group(0)
        m = re.search(r'<hh:%s\b[^>]*?\bid="%s".*?</hh:%s>' % (elem, eid, elem),
                      header, re.S)
        return m.group(0) if m else None

    def missing_refs(self):
        """본문(section)이 쓰는 paraPr/charPr/style/borderFill ID 중 머리부(header)에
        정의가 없는 것을 찾는다. {elem: {'defined_max':int,'missing':[id…]}}(없으면 {})."""
        header = self._raw['Contents/header.xml'].decode('utf-8', 'replace')
        sect = '\n'.join(etree.tostring(self._sections[n], encoding='unicode')
                         for n in self._sec_names)
        report = {}
        for attr, elem in self._REF_ELEM.items():
            defined = set(int(m) for m in
                          re.findall(r'<hh:%s\b[^>]*?\bid="(\d+)"' % elem, header))
            used = set(int(m) for m in re.findall(r'\b%s="(\d+)"' % attr, sect))
            miss = sorted(used - defined)
            if miss:
                report[elem] = {'defined_max': (max(defined) if defined else -1),
                                'missing': miss}
        return report

    def repair_refs(self):
        """한글이 '열 수 없는 문서'로 거부하는 dangling 참조를 고친다 — 본문이 쓰는
        ID 중 머리부에 없는 것을, 머리부의 '같은 종류 최대 id 정의를 복제'해 채우고
        해당 목록 itemCnt를 +1 한다. 추가한 {elem:[id…]} 반환(저장은 save())."""
        header = self._raw['Contents/header.xml'].decode('utf-8')
        added = {}
        for elem, info in self.missing_refs().items():
            defined = set(int(m) for m in
                          re.findall(r'<hh:%s\b[^>]*?\bid="(\d+)"' % elem, header))
            if not defined:
                continue
            src = max(defined)
            block = self._find_def_block(header, elem, src)
            if not block:
                continue
            clones = [block.replace('id="%s"' % src, 'id="%s"' % mid, 1)
                      for mid in info['missing']]
            pos = header.index(block) + len(block)
            header = header[:pos] + ''.join(clones) + header[pos:]
            lst = self._ELEM_LIST[elem]
            header = re.sub(r'(<hh:%s\b[^>]*?\bitemCnt=")(\d+)(")' % lst,
                            lambda mm: mm.group(1) + str(int(mm.group(2)) + len(clones)) + mm.group(3),
                            header, count=1)
            added[elem] = info['missing']
        self._raw['Contents/header.xml'] = header.encode('utf-8')
        return added

    # ── 복구: 표 격자 손상(한글 열기 무한루프) ──
    @staticmethod
    def _row_simple_widths(tr):
        """행이 모두 colSpan==1 단순행이면 (cellSz width 튜플), 아니면 None."""
        ws = []
        for tc in tr.findall(_HP + 'tc'):
            sp = tc.find(_HP + 'cellSpan'); cs = tc.find(_HP + 'cellSz')
            if cs is None:
                return None
            if sp is not None and sp.get('colSpan') and int(sp.get('colSpan')) != 1:
                return None
            ws.append(int(cs.get('width')))
        return tuple(ws)

    @staticmethod
    def _row_height(tr):
        """rowSpan==1 셀들의 cellSz height 최댓값 = 그 행 높이."""
        cand = []
        for tc in tr.findall(_HP + 'tc'):
            cs = tc.find(_HP + 'cellSz'); sp = tc.find(_HP + 'cellSpan')
            if cs is None:
                continue
            rs = int(sp.get('rowSpan')) if (sp is not None and sp.get('rowSpan')) else 1
            if rs == 1:
                cand.append(int(cs.get('height')))
        return max(cand) if cand else 0

    def fix_table_grid(self):
        """한글이 '열 수 없는 문서'로 거부하거나 열 때 무한 레이아웃 루프에 빠지게 하는
        표 격자 손상을 고친다(주로 외부 도구가 표를 병합·생성하며 망가뜨린 경우):
          (1) 셀 <hp:cellAddr>(행·열 좌표) 누락 → colSpan/rowSpan 반영해 위치 계산·삽입  ★치명적
          (2) 행마다 열 너비가 다른 ragged 격자 → 가장 흔한 패턴으로 cellSz width 통일
          (3) 표 sz.height 가 행 높이 합과 불일치 → 합으로 보정
        반환 {'celladdr':n, 'colgrid':[tbl idx…], 'height':[tbl idx…]}.  저장은 save()."""
        from collections import Counter
        rep = {'celladdr': 0, 'colgrid': [], 'height': []}
        for ti, tbl in enumerate(self._tables()):
            trs = tbl.findall(_HP + 'tr')
            # (1) cellAddr 누락 보정 (+ 기존 것도 올바른 좌표로 정규화)
            rowspan_left = {}
            for ri, tr in enumerate(trs):
                col = 0
                for tc in tr.findall(_HP + 'tc'):
                    while rowspan_left.get(col, 0) > 0:
                        col += 1
                    sp = tc.find(_HP + 'cellSpan')
                    cspan = int(sp.get('colSpan')) if (sp is not None and sp.get('colSpan')) else 1
                    rspan = int(sp.get('rowSpan')) if (sp is not None and sp.get('rowSpan')) else 1
                    ca = tc.find(_HP + 'cellAddr')
                    if ca is None:
                        ca = etree.Element(_HP + 'cellAddr')
                        sub = tc.find(_HP + 'subList')
                        pos = list(tc).index(sub) + 1 if sub is not None else 0
                        tc.insert(pos, ca)
                        rep['celladdr'] += 1
                    ca.set('colAddr', str(col)); ca.set('rowAddr', str(ri))
                    if rspan > 1:
                        for c in range(col, col + cspan):
                            rowspan_left[c] = rspan
                    col += cspan
                for c in list(rowspan_left):
                    if rowspan_left[c] > 0:
                        rowspan_left[c] -= 1
            # (2) ragged 열너비 → 최빈 패턴으로 통일
            sw = [self._row_simple_widths(tr) for tr in trs]
            pats = Counter(s for s in sw if s)
            if len(pats) > 1:
                modal, _ = pats.most_common(1)[0]
                nc = len(modal); changed = 0
                for tr, s in zip(trs, sw):
                    if s and len(s) == nc and s != modal:
                        for tc, w in zip(tr.findall(_HP + 'tc'), modal):
                            tc.find(_HP + 'cellSz').set('width', str(w))
                        changed += 1
                if changed:
                    rep['colgrid'].append(ti)
            # (3) 표 높이 보정
            sz = tbl.find(_HP + 'sz')
            if sz is not None:
                tot = sum(self._row_height(tr) for tr in trs)
                if tot > 0 and tot != int(sz.get('height')):
                    sz.set('height', str(tot)); rep['height'].append(ti)
        return rep

    # ── 저장 ──
    def save(self, out_path):
        # 메모리(_raw + 수정된 sections) 기준으로 재패키징. mimetype은 무압축·원래 순서 유지.
        # (기존엔 비-섹션 파일을 디스크 원본에서 다시 읽어 header.xml 수정이 무시됐음 → _raw 사용으로 보완.)
        rendered = {name: etree.tostring(root, xml_declaration=True,
                                         encoding='UTF-8', standalone=True)
                    for name, root in self._sections.items()}
        with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for name in self._order:
                if name in rendered:
                    data = rendered[name]
                elif name in self._raw:
                    data = self._raw[name]
                else:
                    continue
                if name == 'mimetype':
                    zi = zipfile.ZipInfo('mimetype')
                    zi.compress_type = zipfile.ZIP_STORED
                    zout.writestr(zi, data)
                else:
                    zout.writestr(name, data)
        return out_path


# ── 복구 편의 함수 ──
def repair(path, out=None, backup=True):
    """안 열리는 hwpx를 한컴 없이 고친다. 두 층의 손상을 함께 처리:
      · 참조 무결성 — 본문이 머리부에 없는 paraPr/charPr/style/borderFill ID 참조(repair_refs)
      · 표 격자 — 셀 cellAddr 누락·행별 열너비 불일치·표 높이 불일치(fix_table_grid)
    둘 다 한글이 '열 수 없는 문서'로 거부하거나 열 때 무한 루프하게 만드는 대표 원인이다.
      out   : 출력 경로(None이면 원본 덮어쓰기)
      backup: 덮어쓸 때 원본을 '<원본>.bak.hwpx'로 1회 백업
    반환 {'fixed':bool,'added':{elem:[id…]},'tables':{celladdr,colgrid,height},'out','backup'}."""
    out = out or path
    ed = HwpxEditor(path)
    added = ed.repair_refs()
    tables = ed.fix_table_grid()
    fixed = bool(added) or bool(tables['celladdr'] or tables['colgrid'] or tables['height'])
    bak = None
    if fixed and backup and os.path.abspath(out) == os.path.abspath(path):
        bak = os.path.splitext(path)[0] + '.bak.hwpx'
        if not os.path.exists(bak):
            import shutil
            shutil.copy2(path, bak)
    ed.save(out)
    return {'fixed': fixed, 'added': added, 'tables': tables, 'out': out, 'backup': bak}


# ── CLI ──
def _main(argv):
    if not argv or argv[0].startswith('-'):
        print(__doc__); return 1
    src = argv[0]
    out = None
    mapping = {}
    do_repair = '--repair' in argv
    i = 1
    while i < len(argv):
        a = argv[i]
        if a == '-o':
            out = argv[i + 1]; i += 2
        elif a == '--set':
            k, _, v = argv[i + 1].partition('=')
            mapping[k] = v; i += 2
        elif a == '--map':
            with io.open(argv[i + 1], encoding='utf-8') as fp:
                mapping.update(json.load(fp))
            i += 2
        else:
            i += 1
    if do_repair:
        rep = repair(src, out=out)   # out=None이면 덮어쓰기(+.bak 백업)
        if rep['fixed']:
            parts = []
            if rep['added']:
                parts.append('참조 ' + ', '.join('%s=%s' % (k, v) for k, v in rep['added'].items()))
            t = rep['tables']
            if t['celladdr']:
                parts.append('cellAddr %d개' % t['celladdr'])
            if t['colgrid']:
                parts.append('열격자통일 표%s' % t['colgrid'])
            if t['height']:
                parts.append('표높이보정 표%s' % t['height'])
            msg = '[FIXED] %s  (%s)' % (rep['out'], ' / '.join(parts))
            if rep['backup']:
                msg += '  (백업: %s)' % os.path.basename(rep['backup'])
            msg += '  → baking 권장: python hwpx_bake.py "%s"' % rep['out']
        else:
            msg = '[OK] %s — 참조·표 격자 정상(복구 불필요)' % rep['out']
        sys.stdout.buffer.write((msg + '\n').encode('utf-8'))
        return 0
    if not out:
        out = os.path.splitext(src)[0] + '_edited.hwpx'
    ed = HwpxEditor(src)
    n = ed.replace_map(mapping) if mapping else 0
    ed.save(out)
    print('[OK] %d건 치환 → %s  (baking: python hwpx_bake.py "%s")' % (n, out, out))
    return 0

if __name__ == '__main__':
    sys.exit(_main(sys.argv[1:]))
