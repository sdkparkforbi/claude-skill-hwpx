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

CLI
  python hwpx_edit.py 양식.hwpx -o 결과.hwpx --set "{{이름}}=홍길동" --set "{{금액}}=500"
  python hwpx_edit.py 양식.hwpx --map 치환표.json -o 결과.hwpx
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

    # ── 저장 ──
    def save(self, out_path):
        rendered = {name: etree.tostring(root, xml_declaration=True,
                                         encoding='UTF-8', standalone=True)
                    for name, root in self._sections.items()}
        with zipfile.ZipFile(self.path) as zin, \
             zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                name = item.filename
                if name in rendered:
                    zout.writestr(item, rendered[name])
                else:
                    zout.writestr(item, zin.read(name))
        return out_path


# ── CLI ──
def _main(argv):
    if not argv or argv[0].startswith('-'):
        print(__doc__); return 1
    src = argv[0]
    out = None
    mapping = {}
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
    if not out:
        out = os.path.splitext(src)[0] + '_edited.hwpx'
    ed = HwpxEditor(src)
    n = ed.replace_map(mapping) if mapping else 0
    ed.save(out)
    print('[OK] %d건 치환 → %s  (baking: python hwpx_bake.py "%s")' % (n, out, out))
    return 0

if __name__ == '__main__':
    sys.exit(_main(sys.argv[1:]))
