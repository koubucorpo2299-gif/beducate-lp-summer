import datetime
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.formatting.rule import ColorScaleRule

wb = openpyxl.Workbook()

# ─────────────────────────────────────────
# 定数
# ─────────────────────────────────────────
DARK_BLUE   = "1F4E79"
MID_BLUE    = "2E75B6"
LIGHT_BLUE  = "DAEEF3"
NUM_WEEKS   = 16
WEEK_COL0   = 4          # 週カラムの開始列インデックス（1始まり）
HDR_ROW1    = 3          # 週番号行
HDR_ROW2    = 4          # 日付行
FIRST_DATA  = 5          # データ開始行

SUBJECTS = [
    ("英語", "BDD7EE", [
        "文法：時制", "文法：仮定法", "文法：関係詞",
        "文法：不定詞・動名詞", "文法：比較表現",
        "長文読解①", "長文読解②", "リスニング①", "英作文①", "英作文②",
    ]),
    ("数学", "FCE4D6", [
        "因数分解", "二次方程式", "二次関数", "三角比",
        "確率・場合の数", "数列", "微分", "積分", "ベクトル", "整数の性質",
    ]),
    ("国語", "E2EFDA", [
        "現代文：評論①", "現代文：評論②", "現代文：小説",
        "古文：文法基礎", "古文：読解①", "古文：読解②",
        "漢文：句法", "漢文：読解", "語彙・漢字",
    ]),
    ("理科", "EAD1DC", [
        "物理：力学①", "物理：力学②", "物理：電磁気",
        "化学：理論化学", "化学：有機化学", "化学：無機化学",
        "生物：細胞・遺伝", "地学：大気・気象",
    ]),
    ("社会", "FFF2CC", [
        "日本史：近代①", "日本史：近代②", "日本史：現代",
        "世界史：近代ヨーロッパ", "世界史：近代アジア",
        "地理：地形・気候", "地理：産業・人口", "政治経済：憲法",
    ]),
]

def thin_border(color="BFBFBF"):
    s = Side(style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def fill(hex_color):
    return PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")

# ─────────────────────────────────────────
# マスターシート
# ─────────────────────────────────────────
ws_m = wb.active
ws_m.title = "マスター"

ws_m.column_dimensions["A"].width = 10
ws_m.column_dimensions["B"].width = 28
ws_m.column_dimensions["C"].width = 25
ws_m.column_dimensions["D"].width = 20

# ヘッダー
hdr_fill = fill(DARK_BLUE)
hdr_font = Font(color="FFFFFF", bold=True, size=11)
for col, txt in enumerate(["教科", "単元名", "教材名（任意）", "備考（任意）"], 1):
    c = ws_m.cell(row=1, column=col, value=txt)
    c.fill = hdr_fill
    c.font = hdr_font
    c.alignment = Alignment(horizontal="center", vertical="center")
ws_m.row_dimensions[1].height = 22

# データ行と named range 記録
subject_ranges = {}
cur = 2
for name, color, units in SUBJECTS:
    start = cur
    sf = fill(color)
    for unit in units:
        for ci, val in enumerate([name, unit, "", ""], 1):
            c = ws_m.cell(row=cur, column=ci, value=val)
            c.fill = sf
            c.alignment = Alignment(vertical="center")
        ws_m.row_dimensions[cur].height = 18
        cur += 1
    subject_ranges[name] = (start, cur - 1)

# 罫線
brd = thin_border()
for row in ws_m.iter_rows(min_row=1, max_row=cur - 1, min_col=1, max_col=4):
    for c in row:
        c.border = brd

# Named ranges（教科名 → 単元リスト）
for name, (s, e) in subject_ranges.items():
    ref = f"マスター!$B${s}:$B${e}"
    wb.defined_names[name] = DefinedName(name, attr_text=ref)

# ─────────────────────────────────────────
# 学習計画シート
# ─────────────────────────────────────────
ws_p = wb.create_sheet("学習計画")
last_col_letter = get_column_letter(WEEK_COL0 + NUM_WEEKS - 1)

ws_p.column_dimensions["A"].width = 10
ws_p.column_dimensions["B"].width = 22
ws_p.column_dimensions["C"].width = 9
for i in range(NUM_WEEKS):
    ws_p.column_dimensions[get_column_letter(WEEK_COL0 + i)].width = 12

# ── 行1: タイトル ──
ws_p.merge_cells(f"A1:{last_col_letter}1")
t = ws_p["A1"]
t.value = "学習計画表"
t.font = Font(size=22, bold=True, color=DARK_BLUE)
t.alignment = Alignment(horizontal="center", vertical="center")
t.fill = fill(LIGHT_BLUE)
ws_p.row_dimensions[1].height = 42

# ── 行2: 生徒情報 ──
ws_p.row_dimensions[2].height = 26
for col_ltr, label in [("A", "生徒名"), ("C", "開始日"), ("F", "終了日")]:
    c = ws_p[f"{col_ltr}2"]
    c.value = label
    c.font = Font(bold=True, size=11)
    c.alignment = Alignment(horizontal="right", vertical="center")

input_fill = fill("FFFFD1")
ws_p["B2"].fill = input_fill                         # 生徒名 入力セル
ws_p["D2"].value = datetime.date(2025, 7, 1)
ws_p["D2"].number_format = "YYYY/MM/DD"
ws_p["D2"].fill = input_fill
ws_p["E2"].value = "〜"
ws_p["E2"].alignment = Alignment(horizontal="center", vertical="center")
ws_p["G2"].value = datetime.date(2025, 8, 31)
ws_p["G2"].number_format = "YYYY/MM/DD"
ws_p["G2"].fill = input_fill

# ── 行3〜4: 列ヘッダー ──
ws_p.row_dimensions[3].height = 22
ws_p.row_dimensions[4].height = 28

col_hdr_font = Font(color="FFFFFF", bold=True, size=10)

# A-C: 行3〜4 結合
for ci, label in [(1, "教科"), (2, "目標・ゴール"), (3, "進捗率")]:
    ws_p.merge_cells(start_row=3, start_column=ci, end_row=4, end_column=ci)
    c = ws_p.cell(row=3, column=ci, value=label)
    c.fill = fill(DARK_BLUE)
    c.font = col_hdr_font
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

# 週ヘッダー
for i in range(NUM_WEEKS):
    col = WEEK_COL0 + i
    # 第n週
    c3 = ws_p.cell(row=3, column=col, value=f"第{i+1}週")
    c3.fill = fill(DARK_BLUE)
    c3.font = col_hdr_font
    c3.alignment = Alignment(horizontal="center", vertical="center")
    # 日付範囲（開始日セルD2から自動計算）
    c4 = ws_p.cell(row=4, column=col)
    c4.value = f'=TEXT($D$2+{i*7},"M/D")&"～"&TEXT($D$2+{i*7+6},"M/D")'
    c4.fill = fill(MID_BLUE)
    c4.font = Font(color="FFFFFF", size=9, bold=True)
    c4.alignment = Alignment(horizontal="center", vertical="center")

# ── 行5〜: 教科データ ──
for si, (name, color, _) in enumerate(SUBJECTS):
    row = FIRST_DATA + si
    ws_p.row_dimensions[row].height = 24

    sf = fill(color)
    s_rng, e_rng = subject_ranges[name]

    # 教科
    c = ws_p.cell(row=row, column=1, value=name)
    c.fill = sf
    c.font = Font(bold=True, size=11)
    c.alignment = Alignment(horizontal="center", vertical="center")

    # 目標
    c = ws_p.cell(row=row, column=2)
    c.fill = sf
    c.alignment = Alignment(vertical="center", wrap_text=True)

    # 進捗率
    fw = get_column_letter(WEEK_COL0)
    lw = get_column_letter(WEEK_COL0 + NUM_WEEKS - 1)
    c = ws_p.cell(row=row, column=3)
    c.value = f"=COUNTA({fw}{row}:{lw}{row})/{NUM_WEEKS}"
    c.number_format = "0%"
    c.fill = sf
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.font = Font(bold=True, size=11)

    # 週セル + プルダウン
    dv = DataValidation(
        type="list",
        formula1=f"'マスター'!$B${s_rng}:$B${e_rng}",
        allow_blank=True,
        showDropDown=False,
        showErrorMessage=False,
    )
    ws_p.add_data_validation(dv)

    cell_fill = fill("FFFFFF")
    for i in range(NUM_WEEKS):
        col = WEEK_COL0 + i
        c = ws_p.cell(row=row, column=col)
        c.fill = cell_fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.font = Font(size=9)
        dv.add(c)

# 進捗率カラースケール（赤→黄→緑）
last_data = FIRST_DATA + len(SUBJECTS) - 1
prog_range = f"C{FIRST_DATA}:C{last_data}"
ws_p.conditional_formatting.add(prog_range, ColorScaleRule(
    start_type="num", start_value=0,   start_color="FF4444",
    mid_type="num",   mid_value=0.5,   mid_color="FFDD00",
    end_type="num",   end_value=1,     end_color="00BB44",
))

# 罫線（ヘッダー〜データ全体）
brd = thin_border()
for row in ws_p.iter_rows(min_row=3, max_row=last_data,
                           min_col=1, max_col=WEEK_COL0 + NUM_WEEKS - 1):
    for c in row:
        if not c.border or c.border == Border():
            c.border = brd
        else:
            # 既存ボーダーを保持しつつ薄枠を追加（省略可）
            c.border = brd

# ヘッダー行の下に太線
med = Side(style="medium", color="000000")
for ci in range(1, WEEK_COL0 + NUM_WEEKS):
    c = ws_p.cell(row=4, column=ci)
    c.border = Border(left=c.border.left, right=c.border.right,
                      top=c.border.top, bottom=med)

# ペイン固定（左3列＋ヘッダー2行を固定）
ws_p.freeze_panes = ws_p.cell(row=FIRST_DATA, column=WEEK_COL0)

# ─────────────────────────────────────────
# 保存
# ─────────────────────────────────────────
OUT = "/home/user/beducate-lp-summer/学習計画テンプレート.xlsx"
wb.save(OUT)
print(f"✅ 保存完了: {OUT}")
