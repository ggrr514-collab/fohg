# -*- coding: utf-8 -*-
"""教員用評価シート.xlsx を生成するスクリプト"""
import math
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

FONT_NAME = "游ゴシック"

NAVY = PatternFill("solid", fgColor="1F3864")
BLUE = PatternFill("solid", fgColor="D6E4F0")
GRAY = PatternFill("solid", fgColor="F2F2F2")
INPUT = PatternFill("solid", fgColor="FFFFF0")
PURPLE_F = PatternFill("solid", fgColor="E8DAEF")
GREEN_F = PatternFill("solid", fgColor="D5F5E3")
EXAMPLE_F = PatternFill("solid", fgColor="FFF9C4")

wrap = Alignment(wrap_text=True, vertical="top")
thin = Side(style="thin", color="BFBFBF")
box = Border(left=thin, right=thin, top=thin, bottom=thin)

UNIT_GOAL = ("自分の体験（具体）とそこから得た力や成長（抽象）とのつながりを整理して自己PR文を"
             "書くとともに、グループでの発表を聞き合う中で、聴き取った内容を具体的な観点から"
             "評価する力を身に付ける。")


def _display_width(ch):
    return 2 if ord(ch) > 255 else 1


def _text_width(s):
    return sum(_display_width(ch) for ch in s)


def _get_col_width(ws, col_idx, default=8.43):
    letter = get_column_letter(col_idx)
    dim = ws.column_dimensions.get(letter)
    if dim and dim.width:
        return dim.width
    return default


def estimate_required_height(text, width_units, font_size=10, line_factor=1.7, padding=10):
    """テキストの折り返し行数を見積もり、必要な行高さ（pt）を返す。日本語フォントは
    Calibri基準の列幅より収まりが悪いため、余裕を持たせて保守的な係数を使う。"""
    if not text:
        return 0
    chars_per_line_half = max(4, width_units / 2.5) * 2  # 半角換算の1行あたり文字数
    total_lines = 0
    for para in str(text).split("\n"):
        if para == "":
            total_lines += 1
        else:
            w = _text_width(para)
            total_lines += max(1, math.ceil(w / chars_per_line_half))
    return total_lines * font_size * line_factor + padding


def _bump_height(ws, r, height):
    current = ws.row_dimensions[r].height
    ws.row_dimensions[r].height = height if current is None else max(current, height)


def mg(ws, r1, c1, r2, c2, value, font=None, fill=None, align=None, border=None):
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    cell = ws.cell(row=r1, column=c1)
    cell.value = value
    if font:
        cell.font = font
    if fill:
        cell.fill = fill
    cell.alignment = align if align else wrap
    if border:
        for row in ws.iter_rows(min_row=r1, max_row=r2, min_col=c1, max_col=c2):
            for c in row:
                c.border = border
    if isinstance(value, str) and value and not value.startswith("="):
        width_units = sum(_get_col_width(ws, c) for c in range(c1, c2 + 1))
        font_size = font.size if font and font.size else 11
        needed = estimate_required_height(value, width_units, font_size=font_size)
        span_rows = r2 - r1 + 1
        needed_per_row = needed / span_rows
        for rr in range(r1, r2 + 1):
            current = ws.row_dimensions[rr].height
            if current is None or current < needed_per_row:
                ws.row_dimensions[rr].height = needed_per_row
    return cell


def autofit_row_heights(ws):
    """マージされていない単独セルについても、テキスト量から必要な行高さを見積もって
    不足分を引き上げる（mg()経由のセルは既に個別処理済み）。"""
    merged_cells = set()
    for merged_range in ws.merged_cells.ranges:
        for row in range(merged_range.min_row, merged_range.max_row + 1):
            for col in range(merged_range.min_col, merged_range.max_col + 1):
                merged_cells.add((row, col))

    for row in ws.iter_rows():
        for cell in row:
            if (cell.row, cell.column) in merged_cells:
                continue
            value = cell.value
            if not isinstance(value, str) or not value or value.startswith("="):
                continue
            font_size = cell.font.size if cell.font and cell.font.size else 11
            width_units = _get_col_width(ws, cell.column)
            needed = estimate_required_height(value, width_units, font_size=font_size)
            current = ws.row_dimensions[cell.row].height
            if current is None or current < needed:
                ws.row_dimensions[cell.row].height = needed


wb = Workbook()

# ============================================================
# Sheet1: 回答入力
# ============================================================
ws1 = wb.active
ws1.title = "回答入力"
ws1.sheet_view.showGridLines = False

HEADERS = ["クラス", "出席番号", "氏名",
           "発問①\n（体験）", "発問②\n（学び）", "自己PR文\n（作文）",
           "自己PR文\nAI評価(S〜D)", "自己PR文\nAIコメント", "自己PR文\n教師最終評価(S〜D)",
           "聞き取り評価①\n（発表者・内容）", "聞き取り評価②\n（発表者・内容）",
           "聞き取り評価③\n（発表者・内容）",
           "聞き取り評価\nAI評価(S〜D)", "聞き取り評価\nAIコメント", "聞き取り評価\n教師最終評価(S〜D)",
           "振り返り"]
WIDTHS = [8, 8, 10, 24, 24, 30, 10, 26, 12, 26, 26, 26, 10, 26, 12, 22]
TOTAL_COLS = len(HEADERS)

for i, w in enumerate(WIDTHS, start=1):
    ws1.column_dimensions[ws1.cell(row=1, column=i).column_letter].width = w

mg(ws1, 1, 1, 1, TOTAL_COLS, "自己PR文を書く ―回答入力・評価シート―",
   Font(name=FONT_NAME, size=14, bold=True, color="FFFFFF"), NAVY,
   Alignment(horizontal="center", vertical="center"))
_bump_height(ws1, 1, 26)

mg(ws1, 2, 1, 2, TOTAL_COLS, "単元目標：" + UNIT_GOAL,
   Font(name=FONT_NAME, size=10, bold=True), BLUE, wrap)
_bump_height(ws1, 2, 30)

for col, h in enumerate(HEADERS, start=1):
    c = ws1.cell(row=3, column=col, value=h)
    c.font = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
    c.fill = NAVY
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = box
_bump_height(ws1, 3, 40)

DATA_START, DATA_END = 4, 43
dv_grade = DataValidation(type="list", formula1='"S,A,B,C,D"', allow_blank=True)
ws1.add_data_validation(dv_grade)

AI_EVAL_COLS = (7, 13)      # G:自己PR文AI評価 M:聞き取り評価AI評価
TEACHER_COLS = (9, 15)      # I:自己PR文教師評価 O:聞き取り評価教師評価
INPUT_COLS = (4, 5, 6, 8, 10, 11, 12, 14, 16)

for r in range(DATA_START, DATA_END + 1):
    for col in range(1, TOTAL_COLS + 1):
        c = ws1.cell(row=r, column=col)
        c.border = box
        c.font = Font(name=FONT_NAME, size=10)
        c.alignment = wrap
        if col in AI_EVAL_COLS:
            c.fill = PURPLE_F
        elif col in TEACHER_COLS:
            c.fill = GREEN_F
        elif col in INPUT_COLS:
            c.fill = INPUT
    _bump_height(ws1, r, 32)
    for col in AI_EVAL_COLS + TEACHER_COLS:
        dv_grade.add(ws1.cell(row=r, column=col))

# 記入例（1行のみ、実データではないことが分かるよう色で明示）
example_row = [
    "3", "12", "山田 太郎",
    "文化祭のクラス合唱でパートリーダーを務めた。練習初日に声がまとまらず雰囲気が悪くなった。",
    "問題の本質を見極めて、小さく分解して向き合う力が身についたと感じている。",
    "私は文化祭のクラス合唱で、苦手な音楽のパートリーダーを務めました。練習初日、声がまとまら"
    "ずクラスの空気が重くなったとき、逃げずに一人ひとりの音を聞いて回り、苦手な部分だけを取り出"
    "して繰り返し練習する方法を提案しました。この経験から、私は「問題の本質を見極めて、小さく分"
    "解して向き合う力」が身についたと感じています。高校でも、難しい課題ほど一度立ち止まって整理"
    "してから取り組みたいです。",
    "S", "体験と学びの両方が具体的で、独自の言葉で表現されている。", "S",
    "鈴木さん：練習で声が小さかった人に個別に声をかけていたのが良かった。誰が苦手か具体的に"
    "分かって伝わった。",
    "佐藤さん：委員会活動で決め方に困ったとき、多数決ではなく一人ずつ意見を聞いたと言っていて、"
    "配慮の具体例が分かりやすかった。",
    "",
    "S", "話し手の言葉を引用しながら、2人分とも具体的な根拠を挙げて評価できている。", "S",
    "友達の発表を聞いて、自分にはなかった視点に気づけた。",
]
for col, v in enumerate(example_row, start=1):
    c = ws1.cell(row=DATA_START, column=col, value=v)
    c.fill = EXAMPLE_F
    c.font = Font(name=FONT_NAME, size=9, italic=True, color="7D6608")

ws1.freeze_panes = "A4"

autofit_row_heights(ws1)
print("Sheet1 done")

# ============================================================
# Sheet2: AI設定
# ============================================================
ws2 = wb.create_sheet("AI設定")
ws2.sheet_view.showGridLines = False
ws2.column_dimensions["A"].width = 22
ws2.column_dimensions["B"].width = 50

mg(ws2, 1, 1, 1, 2, "AI設定", Font(name=FONT_NAME, size=14, bold=True, color="FFFFFF"), NAVY,
   Alignment(horizontal="center", vertical="center"))
_bump_height(ws2, 1, 26)
_bump_height(ws2, 2, 8)

settings = [
    (3, "AIモデル", "gemini-2.0-flash"),
    (4, "Temperature", 0.1),
    (6, "評価対象①", "自己PR文（回答入力シートF列、書くこと）"),
    (7, "評価対象②", "聞き取り評価①〜③（回答入力シートJ〜L列、聞くこと）"),
    (8, "知識・技能", "3段階（A/B/C）"),
    (9, "思考・判断・表現（書くこと／聞くこと）", "各5段階（S/A/B/C/D）"),
    (11, "データ範囲", "回答入力シート 4行目〜43行目"),
    (12, "ClassroomフォルダID", "（ここにフォルダIDを入力）"),
]
for row, label, value in settings:
    c1 = ws2.cell(row=row, column=1, value=label)
    c1.font = Font(name=FONT_NAME, size=11, bold=True)
    c1.fill = GRAY
    c1.border = box
    c1.alignment = Alignment(vertical="center")
    c2 = ws2.cell(row=row, column=2, value=value)
    c2.font = Font(name=FONT_NAME, size=11)
    c2.fill = INPUT
    c2.border = box
    c2.alignment = Alignment(vertical="center", wrap_text=True)
    _bump_height(ws2, row, 20)

mg(ws2, 14, 1, 16, 2,
   "※ Gemini APIキーは、このシートには入力しません。GASの「プロジェクトの設定」→"
   "「スクリプト プロパティ」で GEMINI_API_KEY を設定してください。",
   Font(name=FONT_NAME, size=9.5, italic=True, color="C00000"), None, wrap)
for i in range(3):
    _bump_height(ws2, 14 + i, 18)

autofit_row_heights(ws2)
print("Sheet2 done")

wb.save("/home/user/fohg/materials/自己PR文を書く/教員用評価シート.xlsx")
print("SAVED")
