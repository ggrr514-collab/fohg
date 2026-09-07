# -*- coding: utf-8 -*-
"""生徒用ワークシート.xlsx を生成するスクリプト"""
import math
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

FONT_NAME = "游ゴシック"

# カラーパレット
NAVY = PatternFill("solid", fgColor="1F3864")
BLUE = PatternFill("solid", fgColor="D6E4F0")
LIGHT_BLUE = PatternFill("solid", fgColor="E8F0FE")
ORANGE_F = PatternFill("solid", fgColor="FFF2CC")
PURPLE_F = PatternFill("solid", fgColor="E8DAEF")
GREEN_F = PatternFill("solid", fgColor="D5F5E3")
GRAY = PatternFill("solid", fgColor="F2F2F2")
INPUT = PatternFill("solid", fgColor="FFFFF0")
AI_TIP = PatternFill("solid", fgColor="EBF5FB")
RED_FONT = Font(name=FONT_NAME, size=11, color="C00000", bold=True)

wrap = Alignment(wrap_text=True, vertical="top")
wrap_center = Alignment(wrap_text=True, vertical="center")
thin = Side(style="thin", color="BFBFBF")
box = Border(left=thin, right=thin, top=thin, bottom=thin)


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


def estimate_required_height(text, width_units, font_size=11, line_factor=1.7, padding=10):
    """テキストの折り返し行数を見積もり、必要な行高さ（pt）を返す。
    列幅の単位はExcelの列幅（おおむね半角文字数）に近似。日本語フォントは同じ列幅でも
    Calibri基準の文字数より収まりが悪いため、余裕を持たせて保守的な係数を使う。"""
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
    """シート全体を走査し、マージされていない単独セルについても、テキスト量から必要な
    行高さを見積もって不足分を引き上げる（mg()経由のセルは既に個別処理済み）。"""
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


def _bump_height(ws, r, height):
    current = ws.row_dimensions[r].height
    ws.row_dimensions[r].height = height if current is None else max(current, height)


def section_header(ws, r, text, fill, span=(1, 4), font_color="FFFFFF", size=12):
    mg(ws, r, span[0], r, span[1], text,
       Font(name=FONT_NAME, size=size, bold=True, color=font_color), fill,
       Alignment(vertical="center", horizontal="left", indent=1))
    _bump_height(ws, r, 24)
    return r + 1


def ai_tip_row(ws, r, text, span=(1, 3)):
    mg(ws, r, span[0], r, span[1], "🤖 " + text,
       Font(name=FONT_NAME, size=9, color="1A5276", italic=True), AI_TIP, wrap)
    lines = text.count("\n") + 1 + (len(text) // 40)
    _bump_height(ws, r, max(30, 14 * lines))
    return r + 1


def label_row(ws, r, text, span=(1, 3), height=20, bold=True, fill=GRAY, size=11):
    mg(ws, r, span[0], r, span[1], text,
       Font(name=FONT_NAME, size=size, bold=bold), fill, wrap)
    _bump_height(ws, r, height)
    return r + 1


def input_row(ws, r, span=(1, 3), height=90, counter_target=None):
    """counter_target=True で、字数の上限を示さない「現在字数のみ」のカウンターをD列に付ける。
    字数に制限はないため、目安の上限値は表示しない。"""
    mg(ws, r, span[0], r, span[1], None, Font(name=FONT_NAME, size=11), INPUT, wrap, box)
    ws.row_dimensions[r].height = height
    if counter_target:
        cell_ref = f"{get_column_letter(span[0])}{r}"
        d_cell = ws.cell(row=r, column=span[1] + 1)
        d_cell.value = f'=IF({cell_ref}="","0字",LEN({cell_ref})&"字")'
        d_cell.font = Font(name=FONT_NAME, size=9, color="808080")
        d_cell.alignment = Alignment(vertical="top", horizontal="left")
    return r + 1


UNIT_GOAL = ("自分の体験（具体）とそこから得た力や成長（抽象）とのつながりを整理して自己PR文を"
             "書くとともに、グループでの発表を聞き合う中で、聴き取った内容を具体的な観点から"
             "評価する力を身に付ける。")

wb = Workbook()

# ============================================================
# Sheet1: 単元ガイド
# ============================================================
ws1 = wb.active
ws1.title = "単元ガイド"
ws1.sheet_view.showGridLines = False
for col, w in zip("ABCDEF", [10, 10, 12, 14, 14, 14]):
    ws1.column_dimensions[col].width = w

mg(ws1, 1, 1, 1, 6, "自分の歩みを言葉にする ―自己PR文を書こう―",
   Font(name=FONT_NAME, size=16, bold=True, color="1F3864"), None,
   Alignment(horizontal="center", vertical="center"))
_bump_height(ws1, 1, 32)

# 基本情報欄
ws1["A2"] = 0
ws1["A2"].number_format = '0"組"'
ws1["A2"].fill = LIGHT_BLUE
ws1["A2"].font = Font(name=FONT_NAME, size=12)
ws1["A2"].alignment = Alignment(horizontal="center", vertical="center")
ws1["A2"].border = box

ws1["B2"] = 0
ws1["B2"].number_format = '0"番"'
ws1["B2"].fill = LIGHT_BLUE
ws1["B2"].font = Font(name=FONT_NAME, size=12)
ws1["B2"].alignment = Alignment(horizontal="center", vertical="center")
ws1["B2"].border = box

ws1["C2"] = "名前"
ws1["C2"].font = Font(name=FONT_NAME, size=11, bold=True)
ws1["C2"].alignment = Alignment(horizontal="center", vertical="center")

mg(ws1, 2, 4, 2, 6, None, Font(name=FONT_NAME, size=12), LIGHT_BLUE,
   Alignment(horizontal="left", vertical="center", indent=1), box)
_bump_height(ws1, 2, 24)

dv_class = DataValidation(type="whole", operator="between", formula1=1, formula2=20)
dv_number = DataValidation(type="whole", operator="between", formula1=1, formula2=50)
ws1.add_data_validation(dv_class)
ws1.add_data_validation(dv_number)
dv_class.add(ws1["A2"])
dv_number.add(ws1["B2"])

mg(ws1, 3, 1, 3, 6, "※ 組・番は半角数字で入力してね（例：2、15）", RED_FONT, None,
   Alignment(horizontal="left", vertical="center"))
_bump_height(ws1, 3, 18)

_bump_height(ws1, 4, 8)

r = section_header(ws1, 5, "■ 単元目標", NAVY, span=(1, 6))
mg(ws1, r, 1, r + 1, 6, UNIT_GOAL, Font(name=FONT_NAME, size=12, bold=True), BLUE, wrap)
_bump_height(ws1, r, 24)
_bump_height(ws1, r + 1, 24)
r += 2
_bump_height(ws1, r, 8)
r += 1

r = section_header(ws1, r, "■ この単元について", NAVY, span=(1, 6))
about_text = (
    "令和9年度埼玉県公立高等学校入学者選抜から、すべての受検生に「面接」が実施されることになりまし"
    "た。出願のときには「自己評価資料」を提出します。そこには、これまでの自分の体験を振り返り、"
    "力を注いだことや努力したこと、高校入学後や将来取り組んでみたいこと、自己PRなどを、自分の言葉"
    "で書きます。面接当日は、その内容をもとに「My Voice」として自分の言葉で語ります。\n"
    "県の資料には、「実績そのものではなく、そこに至るまでの過程（プロセス）や意欲、身に付いた力な"
    "どを多面的に評価する」「文章・文字の巧拙は評価の対象外」「内容に正解はない」と書かれています。"
    "立派な実績を並べる必要はありません。小さな出来事でも、そこから自分が何を感じ、何を学んだかを"
    "自分の言葉で伝えることが大切です。\n"
    "この単元では、第1時で自己PR文の書き方を学び、第2時ではグループの中で自己PR文をもとに発表し"
    "合い、友達の話を具体的な観点で聞き取り評価する活動を行います。話す力だけでなく、相手の話を"
    "正確に受け止めて評価する「聞く力」もあわせて育てます。"
)
mg(ws1, r, 1, r + 6, 6, about_text, Font(name=FONT_NAME, size=10.5), None, wrap)
for i in range(7):
    _bump_height(ws1, r + i, 24)
r += 7
_bump_height(ws1, r, 8)
r += 1

r = section_header(ws1, r, "■ 授業の流れ", NAVY, span=(1, 6))
schedule = [
    ("第1時", "① 自己PR文とは何かを知る　② 発問①：印象に残る体験を具体的に書く"
              "　③ 発問②：その体験から得た学びを自分の言葉でまとめる　④ AIと壁打ちして加筆する"
              "　⑤ 構想メモを作る　⑥「作文例」を参考に自己PR文（300〜500字）を書く"
              "　⑦ 必要に応じてAIに添削してもらい推敲する"),
    ("第2時", "① 面接の「My Voice」と、聞くことの大切さを知る　② 3〜4人グループで自己PR文を発表"
              "し合う　③ 聞き手は質問をし、発表者は答える　④ 聞き手は観点に沿って聞き取り評価を"
              "書く（発表そのものは評価しない）　⑤ 振り返りを書く"),
]
for title, detail in schedule:
    mg(ws1, r, 1, r, 2, title, Font(name=FONT_NAME, size=11, bold=True), GRAY,
       Alignment(horizontal="center", vertical="center"))
    mg(ws1, r, 3, r, 6, detail, Font(name=FONT_NAME, size=10.5), None, wrap)
    _bump_height(ws1, r, 62)
    r += 1
_bump_height(ws1, r, 8)
r += 1

r = section_header(ws1, r, "■ AIを使うときの約束", NAVY, span=(1, 6))
rules = (
    "・自分の体験や考えは、まず自分の力で書いてみよう。AIに最初から書いてもらわないこと。\n"
    "・AIには「書いてもらう」のではなく、「意見やヒントをもらう」ために使おう。\n"
    "・AIの言葉をそのまま貼り付けるのではなく、自分の言葉に直してから使おう。\n"
    "・自己評価資料には、生成AIの回答をそのまま転記・模倣してはいけません（自分で考え、記載すること）。"
)
mg(ws1, r, 1, r + 3, 6, rules, Font(name=FONT_NAME, size=10.5), None, wrap)
for i in range(4):
    _bump_height(ws1, r + i, 22)

autofit_row_heights(ws1)
print("Sheet1 done, last row:", r)

# ============================================================
# Sheet2: 評価基準
# ============================================================
ws2 = wb.create_sheet("評価基準")
ws2.sheet_view.showGridLines = False
for col, w in zip("ABCDE", [8, 30, 40, 40, 40]):
    ws2.column_dimensions[col].width = w

mg(ws2, 1, 1, 1, 5, "評価基準（ルーブリック）", Font(name=FONT_NAME, size=15, bold=True, color="1F3864"),
   None, Alignment(horizontal="center", vertical="center"))
_bump_height(ws2, 1, 28)
mg(ws2, 2, 1, 2, 5, UNIT_GOAL, Font(name=FONT_NAME, size=10.5, bold=True), BLUE, wrap)
_bump_height(ws2, 2, 30)
_bump_height(ws2, 3, 8)

r = 4
r = section_header(ws2, r, "■ 知識・技能（3段階）", NAVY, span=(1, 5))
mg(ws2, r, 1, r, 5,
   "評価規準：体験（具体）と、そこから得た学び（抽象）とを、適切な語句を選んで結びつけて表現している"
   "か。",
   Font(name=FONT_NAME, size=10.5, italic=True), None, wrap)
_bump_height(ws2, r, 22)
r += 1

chishiki = [
    ("A", "体験と学びを的確な語句で結びつけ、読み手に伝わるように表現している。", GREEN_F),
    ("B", "体験と学びを書いているが、語句の選択や結びつけ方がやや曖昧である。", ORANGE_F),
    ("C", "体験または学びのどちらか一方のみの記述にとどまる、あるいは両者のつながりが読み取れない。",
     GRAY),
]
for label, text, fill in chishiki:
    mg(ws2, r, 1, r, 1, label, Font(name=FONT_NAME, size=13, bold=True), fill,
       Alignment(horizontal="center", vertical="center"), box)
    mg(ws2, r, 2, r, 5, text, Font(name=FONT_NAME, size=10.5), None, wrap, box)
    _bump_height(ws2, r, 34)
    r += 1
_bump_height(ws2, r, 8)
r += 1

r = section_header(ws2, r, "■ 思考・判断・表現①（書くこと・5段階／第1時）", NAVY, span=(1, 5))
mg(ws2, r, 1, r, 5,
   "評価規準：目的（高校入試の自己評価資料）に応じて、社会生活の中から題材を決め、伝えたいことを"
   "明確にして書いているか。",
   Font(name=FONT_NAME, size=10.5, italic=True), None, wrap)
_bump_height(ws2, r, 22)
r += 1

shihan = [
    ("S", "具体的な体験と、そこから得た学びの両方が、独自性・説得力をもって明確に書かれている。",
     PatternFill("solid", fgColor="F9E79F")),
    ("A", "具体的な体験と、そこから得た学びの両方が明確に書かれている。", GREEN_F),
    ("B", "具体的な体験、学びのどちらか一方のみが書かれている。", ORANGE_F),
    ("C", "体験も学びも抽象的・一般的な表現にとどまっている。", GRAY),
    ("D", "題材が定まっておらず、自己PR文として内容が成立していない。",
     PatternFill("solid", fgColor="F2D7D5")),
]
for label, text, fill in shihan:
    mg(ws2, r, 1, r, 1, label, Font(name=FONT_NAME, size=13, bold=True), fill,
       Alignment(horizontal="center", vertical="center"), box)
    mg(ws2, r, 2, r, 5, text, Font(name=FONT_NAME, size=10.5), None, wrap, box)
    _bump_height(ws2, r, 34)
    r += 1
_bump_height(ws2, r, 8)
r += 1

mg(ws2, r, 1, r + 2, 5,
   "★ B と A の分かれ目\n"
   "必須条件①「具体的な体験」と必須条件②「そこから得た学び」の【両方】に触れていなければ、A以上には"
   "なりません。片方だけならB、両方とも曖昧・欠けていればC、題材が定まっていなければDです。",
   Font(name=FONT_NAME, size=11, bold=True, color="C00000"),
   PatternFill("solid", fgColor="FDEDEC"), wrap, box)
for i in range(3):
    _bump_height(ws2, r + i, 22)
r += 3
_bump_height(ws2, r, 10)
r += 1

r = section_header(ws2, r, "■ 思考・判断・表現②（聞くこと・5段階／第2時）", NAVY, span=(1, 5))
mg(ws2, r, 1, r, 5,
   "評価規準：話の展開を予測しながら聞き、聴き取った内容や表現の仕方を、観点に沿って具体的に評価して"
   "いるか。",
   Font(name=FONT_NAME, size=10.5, italic=True), None, wrap)
_bump_height(ws2, r, 22)
r += 1

mg(ws2, r, 1, r, 5,
   "評価の観点：①体験の具体性　②学びの独自性　③質問への応答からの理解の深まり",
   Font(name=FONT_NAME, size=10, italic=True, color="1F3864"), None, wrap)
_bump_height(ws2, r, 18)
r += 1

kiku = [
    ("S", "3つの観点すべてについて、話し手の言葉を具体的な根拠として挙げながら評価し、自分なりの"
          "気づきや考えも述べている。", PatternFill("solid", fgColor="F9E79F")),
    ("A", "3つの観点のうち少なくとも2つについて、話し手の言葉を具体的な根拠として挙げながら評価"
          "している。", GREEN_F),
    ("B", "観点には触れているが、根拠がやや一般的（「わかりやすかった」「よかった」など）にとど"
          "まっている。", ORANGE_F),
    ("C", "観点に基づかない感想（「面白かった」「すごいと思った」など）のみで、評価として成立して"
          "いない。", GRAY),
    ("D", "記述がない、または発表内容と無関係な記述にとどまっている。",
     PatternFill("solid", fgColor="F2D7D5")),
]
for label, text, fill in kiku:
    mg(ws2, r, 1, r, 1, label, Font(name=FONT_NAME, size=13, bold=True), fill,
       Alignment(horizontal="center", vertical="center"), box)
    mg(ws2, r, 2, r, 5, text, Font(name=FONT_NAME, size=10.5), None, wrap, box)
    _bump_height(ws2, r, 34)
    r += 1
_bump_height(ws2, r, 8)
r += 1

mg(ws2, r, 1, r + 2, 5,
   "★ B と A の分かれ目\n"
   "話し手の言葉を引用・要約するなど、具体的な根拠を伴って評価しているかどうかが分かれ目です。根拠"
   "のない感想はB以下、根拠を伴う具体的な評価はA以上とします。",
   Font(name=FONT_NAME, size=11, bold=True, color="C00000"),
   PatternFill("solid", fgColor="FDEDEC"), wrap, box)
for i in range(3):
    _bump_height(ws2, r + i, 22)

autofit_row_heights(ws2)
print("Sheet2 done")

# ============================================================
# Sheet3: 発問・回答
# ============================================================
ws3 = wb.create_sheet("発問・回答")
ws3.sheet_view.showGridLines = False
for col, w in zip("ABCD", [16, 16, 16, 18]):
    ws3.column_dimensions[col].width = w

mg(ws3, 1, 1, 1, 3, "発問・回答", Font(name=FONT_NAME, size=15, bold=True, color="1F3864"),
   None, Alignment(horizontal="center", vertical="center"))
_bump_height(ws3, 1, 28)
mg(ws3, 2, 1, 2, 3, UNIT_GOAL, Font(name=FONT_NAME, size=10, bold=True), BLUE, wrap)
_bump_height(ws3, 2, 30)
_bump_height(ws3, 3, 8)

r = 4
r = section_header(ws3, r, "■ 第1時：体験を掘り起こそう", BLUE, span=(1, 3), font_color="1F3864")
_bump_height(ws3, r, 6)
r += 1

# --- 発問① ---
Q1_TEXT = ("発問① 体験の具体化\n"
           "あなたがこれまでの中学校生活（部活動・委員会・学校行事・地域活動・家庭でのこと"
           "など、何でもよい）の中で、最も「力を注いだ」「努力した」と言える出来事を一つ選ぼう。その"
           "とき、どんな壁や難しさがあったか、いつ・どこで・誰と・何をしたのかが分かるように、具体的"
           "な場面を書こう。")
r = label_row(ws3, r, Q1_TEXT, height=110)
Q1_CELL_ROW = r
r = input_row(ws3, r, height=100, counter_target=True)
r = ai_tip_row(ws3, r,
               "まず自分の体験を思い出して書いてみよう。書けたら「AIプロンプト」シートの壁打ち①に"
               "この内容を貼り付け、AIから「具体的で伝わりやすいか」意見をもらおう（書き直しは自分で"
               "行うこと）。")
_bump_height(ws3, r, 6)
r += 1

# --- 発問② ---
Q2_TEXT = ("発問② 具体から抽象への言語化\n"
           "発問①で書いた体験を振り返り、その出来事を通して自分がどう変わったか、どんな力が身につい"
           "たと感じるかを、自分だけの言葉で一文にまとめよう。「成長した」「頑張った」のような一般的"
           "な言葉ではなく、自分にしか言えない表現を探して書こう。")
r = label_row(ws3, r, Q2_TEXT, height=90)
Q2_CELL_ROW = r
r = input_row(ws3, r, height=80, counter_target=True)
r = ai_tip_row(ws3, r,
               "発問①②の内容をまとめてAIプロンプトシートの壁打ち①に貼り付け、フィードバックをもら"
               "おう。AIの意見はあくまでヒント。書き直すかどうかは自分で判断しよう。")
_bump_height(ws3, r, 10)
r += 1

r = section_header(ws3, r, "■ 自己PR文を書こう（第1時）", ORANGE_F, span=(1, 3), font_color="7D6608")
_bump_height(ws3, r, 6)
r += 1

mg(ws3, r, 1, r, 3,
   "場面設定：あなたが志望する高校に提出する「自己評価資料」の自己PR欄に書くつもりで書こう。",
   RED_FONT, None, wrap)
_bump_height(ws3, r, 30)
r += 1

mg(ws3, r, 1, r, 3,
   "必須条件　① 具体的な体験（エピソード）が書かれていること　"
   "② その体験から得た学び・成長した力が、体験と結びつけて書かれていること　"
   "字数：300〜500字",
   Font(name=FONT_NAME, size=10.5, italic=True), None, wrap)
_bump_height(ws3, r, 34)
r += 1

r = label_row(ws3, r, "構想メモ（体験→学び→高校生活への意欲を簡単に整理しよう）", height=20)
r = input_row(ws3, r, height=70)
r = ai_tip_row(ws3, r,
               "構想メモができたら、AIプロンプトシートの壁打ち②に貼り付け、必須条件①②が満たされて"
               "いそうかアドバイスをもらおう。AIに代わりに書いてもらうのはNG。「作文例」シートも"
               "参考にしよう。")
_bump_height(ws3, r, 6)
r += 1

r = label_row(ws3, r, "自己PR文（300〜500字）", height=20, fill=ORANGE_F)
ESSAY_CELL_ROW = r
r = input_row(ws3, r, height=220, counter_target=True)
r = ai_tip_row(ws3, r,
               "書き終えたら、AIプロンプトシートの添削プロンプトに貼り付け、分かりやすさや具体性に"
               "ついてフィードバックをもらおう。AIに書き直してもらうのではなく、自分の言葉で推敲する"
               "こと。")
_bump_height(ws3, r, 10)
r += 1

# --- 第2時：グループで発表を聞き合おう ---
r = section_header(ws3, r, "■ 第2時：グループで発表を聞き合おう", PURPLE_F, span=(1, 3),
                    font_color="4A235A")
_bump_height(ws3, r, 6)
r += 1

mg(ws3, r, 1, r, 3,
   "3〜4人グループで、一人ずつ自己PR文を発表しよう。発表を聞いたら質問をし、答えてもらおう。"
   "あなたの発表そのものは評価されません。ここでは「聞き手」として、友達の発表をどれだけ具体的に"
   "受け止められたかを評価します。",
   Font(name=FONT_NAME, size=10, italic=True), None, wrap)
_bump_height(ws3, r, 44)
r += 1

mg(ws3, r, 1, r, 3,
   "聞き取り評価の観点　① 体験の具体性（いつ・どこで・何があったか）"
   "　② 学びの独自性（その人にしか言えない気づきか）"
   "　③ 応答からの理解の深まり（質問への答えで理解が深まったか）",
   RED_FONT, None, wrap)
_bump_height(ws3, r, 46)
r += 1
_bump_height(ws3, r, 6)
r += 1

LISTEN_ROWS = []
for i in range(1, 4):
    mg(ws3, r, 1, r, 1, f"発表者{['①','②','③'][i-1]}の名前", Font(name=FONT_NAME, size=10, bold=True),
       GRAY, Alignment(vertical="center"))
    mg(ws3, r, 2, r, 3, None, Font(name=FONT_NAME, size=11), INPUT, wrap, box)
    name_row = r
    _bump_height(ws3, r, 20)
    r += 1

    r = label_row(ws3, r,
                   "聞き取り評価（観点①②③と、質問への応答を踏まえて、具体的に書こう）", height=32)
    comment_row = r
    r = input_row(ws3, r, height=70)
    LISTEN_ROWS.append((name_row, comment_row))
    _bump_height(ws3, r, 8)
    r += 1

r = section_header(ws3, r, "■ 振り返り", GREEN_F, span=(1, 3), font_color="196F3D")
_bump_height(ws3, r, 6)
r += 1
r = label_row(ws3, r, "グループでの活動を通して、聞くことについて考えたこと・学んだことを書こう",
               height=32)
REFLECT_ROW = r
r = input_row(ws3, r, height=70)

autofit_row_heights(ws3)
print("Sheet3 done. Cell map:")
print("Q1_CELL =", f"A{Q1_CELL_ROW}")
print("Q2_CELL =", f"A{Q2_CELL_ROW}")
print("ESSAY_CELL =", f"A{ESSAY_CELL_ROW}")
for i, (name_row, comment_row) in enumerate(LISTEN_ROWS, start=1):
    print(f"LISTEN{i}_NAME_CELL =", f"B{name_row}")
    print(f"LISTEN{i}_COMMENT_CELL =", f"A{comment_row}")
print("REFLECT_CELL =", f"A{REFLECT_ROW}")

# ============================================================
# Sheet4: AIプロンプト
# ============================================================
ws4 = wb.create_sheet("AIプロンプト")
ws4.sheet_view.showGridLines = False
ws4.column_dimensions["A"].width = 95

mg(ws4, 1, 1, 1, 1, "AIプロンプト集（コピペして使おう）",
   Font(name=FONT_NAME, size=15, bold=True, color="1F3864"), None,
   Alignment(horizontal="center", vertical="center"))
_bump_height(ws4, 1, 28)
mg(ws4, 2, 1, 2, 1,
   "使い方：それぞれの見出しの下にあるプロンプトを丸ごとコピーし、AIチャットに貼り付けて、"
   "指示にある「★貼り付け箇所★」に自分が書いた文章を入れて使おう。",
   Font(name=FONT_NAME, size=10.5, italic=True), None, wrap)
_bump_height(ws4, 2, 34)
_bump_height(ws4, 3, 8)

r = 4
r = section_header(ws4, r, "① 壁打ち①（発問チェック）― 第1時で使う", BLUE, span=(1, 1),
                    font_color="1F3864")

prompt1 = (
    "あなたは中学3年生の作文相談にのる先生です。中学3年生にわかりやすい言葉で答えてください。\n"
    "以下は、高校入試の自己PR文を書くために、私が考えた「体験」と「そこから得た学び」です。\n\n"
    "【体験】\n★ここに発問①の答えを貼り付け★\n\n"
    "【そこから得た学び】\n★ここに発問②の答えを貼り付け★\n\n"
    "次の２点についてアドバイスをください。書き直した文章は書かず、ヒントだけを300字以内で"
    "教えてください。\n"
    "1. 体験の場面が具体的でわかりやすいか\n"
    "2. 学びの内容が自分の言葉で表現できているか"
)
mg(ws4, r, 1, r, 1, prompt1, Font(name=FONT_NAME, size=10.5), INPUT, wrap, box)
_bump_height(ws4, r, 220)
r += 1
_bump_height(ws4, r, 12)
r += 1

r = section_header(ws4, r, "② 壁打ち②（構想チェック）― 第2時の前半で使う", ORANGE_F, span=(1, 1),
                    font_color="7D6608")

prompt2 = (
    "あなたは中学3年生の作文相談にのる先生です。中学3年生にわかりやすい言葉で答えてください。\n"
    "私は、高校入試で提出する「自己評価資料」の自己PR欄（300〜500字）を書こうとしています。\n"
    "場面：志望する高校に提出する自己評価資料の自己PR欄\n"
    "必須条件：① 具体的な体験（エピソード）が書かれていること　② その体験から得た学び・成長した力"
    "が、体験と結びつけて書かれていること\n\n"
    "【私の構想メモ】\n★ここに構想メモを貼り付け★\n\n"
    "この構想メモが、①②の必須条件を満たせそうかチェックし、構成についてのアドバイスだけを300字"
    "以内で教えてください。代わりに文章を書くことはしないでください。"
)
mg(ws4, r, 1, r, 1, prompt2, Font(name=FONT_NAME, size=10.5), INPUT, wrap, box)
_bump_height(ws4, r, 220)
r += 1
_bump_height(ws4, r, 12)
r += 1

r = section_header(ws4, r, "③ 添削（まとめ活動アドバイス）― 第2時の後半で使う", PURPLE_F, span=(1, 1),
                    font_color="4A235A")

prompt3 = (
    "あなたは中学3年生の作文相談にのる先生です。中学3年生にわかりやすい言葉で答えてください。\n"
    "以下は、私が書いた高校入試の自己PR文（自己評価資料の自己PR欄、300〜500字）です。\n\n"
    "【私の自己PR文】\n★ここに自己PR文を貼り付け★\n\n"
    "次のチェックリストを確認し、書き直した文章は書かず、ヒントだけを400字以内で教えてください。\n"
    "□ 具体的な体験（エピソード）が書かれているか\n"
    "□ その体験から得た学び・成長した力が、体験と結びつけて書かれているか\n"
    "□ 300〜500字に収まっているか\n"
    "□ 読み手（高校の先生）に伝わる言葉になっているか"
)
mg(ws4, r, 1, r, 1, prompt3, Font(name=FONT_NAME, size=10.5), INPUT, wrap, box)
_bump_height(ws4, r, 230)

autofit_row_heights(ws4)
print("Sheet4 done")

# ============================================================
# Sheet5: 作文例
# ============================================================
ws5 = wb.create_sheet("作文例")
ws5.sheet_view.showGridLines = False
ws5.column_dimensions["A"].width = 90
ws5.column_dimensions["B"].width = 10

mg(ws5, 1, 1, 1, 1, "作文例（自己PR文の書き方を考えるヒント）",
   Font(name=FONT_NAME, size=15, bold=True, color="1F3864"), None,
   Alignment(horizontal="center", vertical="center"))
_bump_height(ws5, 1, 28)

mg(ws5, 2, 1, 2, 1,
   "これから紹介するのは「模範解答」ではなく、あくまで3つの「作文例」です。優劣はありません。"
   "それぞれ良いところが違うので、自分の体験に置きかえながら、どんな書き方が自分に合うか考えて"
   "みよう。",
   Font(name=FONT_NAME, size=10.5, italic=True), None, wrap)
_bump_height(ws5, 2, 40)
_bump_height(ws5, 3, 8)

r = 4

examples = [
    ("作文例①　文化祭のクラス合唱でパートリーダーを務めた話",
     "私は文化祭のクラス合唱で、大の苦手だった音楽のパートリーダーを務めました。立候補した理由は"
     "単純で、誰もやりたがらなかったからです。しかし、いざ始めてみると、練習初日から声がまとまら"
     "ず、教室の空気が重くなっていくのが分かりました。歌が苦手な人ほど声を出すのをためらい、私自"
     "身もどう声をかけていいか分からず、ただ焦るばかりでした。\n"
     "このままではいけないと思い、私は全員の音を一人ずつ聞いて回ることにしました。すると、音程が"
     "ずれている人の多くは、実は「どこの音を出せばいいか分からない」だけだと気づきました。そこで、"
     "苦手な部分だけを取り出し、フレーズを短く区切って繰り返し練習する方法を提案しました。少しずつ"
     "声がそろい始めると、練習に来る人も増えていきました。\n"
     "この経験を通して、私は「問題を大きいまま抱え込まず、小さく分解して一つずつ向き合う力」が身に"
     "ついたと感じています。高校でも、難しい課題に直面したときほど、この経験を思い出し、一度立ち"
     "止まって整理してから取り組んでいきたいです。",
     "・「声がまとまらず教室の空気が重くなった」など、その場にいたからこそ分かる具体的な場面が描か"
     "れている。\n"
     "・体験を「困難の発見→原因の分析→工夫の実行」という流れで描いており、結果だけでなく過程（プ"
     "ロセス）が伝わる。\n"
     "・「問題を大きいまま抱え込まず、小さく分解して一つずつ向き合う力」のように、一般的な言葉では"
     "なく自分にしか言えない表現で学びをまとめている。"),
    ("作文例②　誰にも読まれない学校新聞を作り続けた話",
     "私は生徒会の広報委員として、毎月学校新聞を作っています。正直に言うと、去年までは「誰も読ん"
     "でいないのでは」と感じていました。実際、配っても素通りされることが多く、部活動に取材を申し"
     "込んでも「忙しいから」と断られることもあり、やる気を失いかけたこともあります。\n"
     "それでも私は、記事の書き方を変えてみることにしました。行事の結果だけを伝えるのではなく、大"
     "会前の生徒の練習風景や、準備でがんばった裏側を、放課後に自分から足を運んで取材して書くよう"
     "にしたのです。最初は断られることもありましたが、根気強く声をかけ続けるうちに、「今度はうち"
     "の部活も取材してほしい」と頼まれることも増えていきました。すると、「あの記事、読んだよ」と"
     "声をかけてくれる同級生も少しずつ増えていきました。\n"
     "この経験から、私は「反応がなくても、伝え方を工夫し続ければ、必ず誰かに届く」ということを学"
     "びました。高校でも、すぐに結果が出ないことに直面したとき、あきらめずに工夫を重ねていきたい"
     "です。",
     "・「誰にも読まれていないかもしれない」という、あまり格好よくない本音から書き始めており、無理"
     "に実績を飾っていない。\n"
     "・「反応がなくても、伝え方を工夫し続ければ、必ず誰かに届く」という学びが、単なる「頑張った」"
     "で終わらず、自分の行動指針として言語化できている。"),
    ("作文例③　数学が苦手で、毎朝10分だけ計算問題を解くようにした話",
     "私は数学がとても苦手で、テストのたびに平均点を大きく下回っていました。塾に通う時間もなかっ"
     "たので、まずは朝、学校に行く前の10分間だけ、計算問題を解くことに決めました。最初は面倒に感"
     "じ、三日で挫折しそうになった日もありましたが、同じく数学が苦手な友人に頼んで、毎朝お互いに"
     "問題を出し合うようにしてから、少しずつ習慣として定着していきました。分からない問題があって"
     "も、その場ですぐ友人に聞けたことが、続けられた理由の一つだったと思います。\n"
     "半年続けた結果、点数が大きく上がったわけではありませんが、以前は「何が分からないかも分から"
     "ない」状態だったのが、「わからない部分がどこか」を自分で言えるようになりました。この経験か"
     "ら、私は「小さくても続けられる工夫を見つける力」が身についたと感じています。高校でも、苦手"
     "なことから逃げず、自分に合った続けられる方法を考えていきたいです。",
     "・「点数が大きく上がったわけではない」と正直に書いており、大きな成果がなくても、体験と学びが"
     "具体的であれば十分に説得力のある自己PRになることを示している。\n"
     "・「小さくても続けられる工夫を見つける力」という学びが、体験（10分間の積み重ね、友人との工"
     "夫）と具体的に結びついている。"),
]

for title, essay, points in examples:
    r = section_header(ws5, r, f"■ {title}", NAVY, span=(1, 1))
    mg(ws5, r, 1, r, 1, essay, Font(name=FONT_NAME, size=11), LIGHT_BLUE, wrap, box)
    _bump_height(ws5, r, 260)
    b_cell = ws5.cell(row=r, column=2)
    b_cell.value = f'=IF(A{r}="","0字",LEN(A{r})&"字")'
    b_cell.font = Font(name=FONT_NAME, size=9, color="808080")
    b_cell.alignment = Alignment(vertical="top", horizontal="left")
    r += 1
    r = section_header(ws5, r, "この作文例の良いところ", GREEN_F, span=(1, 1), font_color="196F3D",
                        size=10.5)
    mg(ws5, r, 1, r, 1, points, Font(name=FONT_NAME, size=10.5), None, wrap)
    _bump_height(ws5, r, 70)
    r += 1
    _bump_height(ws5, r, 14)
    r += 1

autofit_row_heights(ws5)
print("Sheet5 done")

wb.save("/home/user/fohg/materials/自己PR文を書く/生徒用ワークシート.xlsx")
print("SAVED")
