#!/usr/bin/env python3
# 한글 네이티브 자동화로 '시드' HWPX 생성.
# 알려진 순서로 글자모양(charPr)·셀배경(borderFill)을 심어두어,
# 이후 section0.xml/header.xml 역공학으로 ID를 매핑하기 위함.
import os, sys, winreg, importlib.util, win32com.client

# pyhwpx 보안 DLL 자동 탐색 (import 없이 위치만)
_spec = importlib.util.find_spec("pyhwpx")
if _spec is None:
    sys.exit("[오류] pyhwpx가 없습니다. 'pip install pyhwpx' 후 다시 실행하세요.")
DLL = os.path.join(os.path.dirname(_spec.origin), "FilePathCheckerModule.dll")
for p in (r"Software\HNC\HwpAutomation\Modules", r"Software\Hnc\HwpUserAction\Modules"):
    k = winreg.CreateKey(winreg.HKEY_CURRENT_USER, p)
    winreg.SetValueEx(k, "FilePathCheckerModule", 0, winreg.REG_SZ, DLL); winreg.CloseKey(k)

hwp = win32com.client.dynamic.Dispatch("HWPFrame.HwpObject")
hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
hwp.SetMessageBoxMode(0xFFFFFF)
hwp.XHwpWindows.Item(0).Visible = False

def rgb(r,g,b): return hwp.RGBColor(r,g,b)

def set_char(height_pt, color, bold):
    hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
    cs = hwp.HParameterSet.HCharShape
    cs.Height = int(height_pt*100)
    cs.TextColor = rgb(*color)
    cs.Bold = 1 if bold else 0
    hwp.HAction.Execute("CharShape", cs.HSet)

def insert(text):
    p = hwp.HParameterSet.HInsertText
    hwp.HAction.GetDefault("InsertText", p.HSet); p.Text = text
    hwp.HAction.Execute("InsertText", p.HSet)

def enter():
    hwp.HAction.Run("BreakPara")

def align_center():
    hwp.HAction.Run("ParagraphShapeAlignCenter")
def align_left():
    hwp.HAction.Run("ParagraphShapeAlignLeft")

NAVY=(0x1F,0x38,0x64); BLUE=(0x2E,0x54,0x96); BLACK=(0x22,0x22,0x22)
GREY=(0x66,0x66,0x66); WHITE=(0xFF,0xFF,0xFF)
F_LABEL=(0xD9,0xE2,0xF3); F_GREEN=(0xCF,0xE9,0xDA); F_ATT=(0xF4,0xF6,0xFB); F_WHITE=(0xFF,0xFF,0xFF)

hwp.HAction.Run("FileNew")

# === 알려진 순서의 스타일 단락들 (charPr 캡처용) ===
# [P0] 제목 18 navy bold, 가운데
set_char(18, NAVY, True); align_center(); insert("SEED_TITLE"); enter()
# [P1] 제목1 14 navy bold, 왼쪽
align_left(); set_char(14, NAVY, True); insert("SEED_H1"); enter()
# [P2] 제목2 12 blue bold
set_char(12, BLUE, True); insert("SEED_H2"); enter()
# [P3] 본문 11 black
set_char(11, BLACK, False); insert("SEED_NORMAL"); enter()
# [P4] 작은글씨 9.5 grey
set_char(9.5, GREY, False); insert("SEED_SMALL"); enter()
# [P5] 흰색 10 bold (헤더셀 글자용)
set_char(10, WHITE, True); insert("SEED_WHITE"); enter()
# [P6] 셀본문 10 black
set_char(10, BLACK, False); insert("SEED_CELL"); enter()

# 본문 색 복구
set_char(11, BLACK, False)

# === 색상 캡처용 표: 1행 5열, 각 셀 배경 = label/green/attach/navy/white ===
pset = hwp.HParameterSet.HTableCreation
hwp.HAction.GetDefault("TableCreate", pset.HSet)
pset.Rows = 1; pset.Cols = 5
pset.WidthType = 2; pset.HeightType = 0
pset.WidthValue = hwp.MiliToHwpUnit(150)
# 열너비 배열
pset.CreateItemArray("ColWidth", 5)
for i in range(5): pset.ColWidth.SetItem(i, hwp.MiliToHwpUnit(30))
hwp.HAction.Execute("TableCreate", pset.HSet)

def cell_fill(color):
    p = hwp.HParameterSet.HCellBorderFill
    hwp.HAction.GetDefault("CellFill", p.HSet)
    p.FillAttr.type = hwp.BrushType("NullBrush|WinBrush")
    p.FillAttr.WinBrushFaceColor = rgb(*color)
    p.FillAttr.WinBrushHatchColor = rgb(0,0,0)
    p.FillAttr.WinBrushAlpha = 0
    hwp.HAction.Execute("CellFill", p.HSet)

# 커서는 표 첫 셀. 셀블록 선택 후 채우기, 다음 셀 이동.
fills = [F_LABEL, F_GREEN, F_ATT, NAVY, F_WHITE]
labels= ["LABEL","GREEN","ATTACH","NAVY","WHITE"]
for i,(f,lb) in enumerate(zip(fills,labels)):
    hwp.HAction.Run("TableCellBlock")   # 현재 셀 블록 선택
    cell_fill(f)
    hwp.HAction.Run("Cancel")           # 블록 해제
    insert(lb)                          # 셀에 텍스트
    if i < 4:
        hwp.HAction.Run("TableRightCell")

# 표 밖으로
hwp.HAction.Run("CloseEx")

hwp.SaveAs(os.path.abspath("_seed.hwpx"), "HWPX", "")
print("[OK] _seed.hwpx")
hwp.Quit()
print("done")
