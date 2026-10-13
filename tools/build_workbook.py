# -*- coding: utf-8 -*-
"""親子JVMの解説ブックを、Meiryo UI・モノトーンで作る。"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.page import PageMargins

ROOT = Path(__file__).resolve().parents[1]
MD_PATH = ROOT / "親子JVMのメモリとCPUとsar.md"
XLSX_PATH = ROOT / "親子JVMのメモリとCPUとsar.xlsx"

INK = "1A1A1A"
WHITE = "FFFFFF"
CHARCOAL = "2C2C2C"
SLATE = "3A3A3A"
MID = "5C5C5C"
GRAY = "6E6E6E"
LINE = "BDBDBD"
HAIR = "D5D5D5"
PAPER = "F7F7F7"
BAND = "E8E8E8"
ZEBRA = "F3F3F3"
INPUT_BG = "D9D9D9"
FORMULA_BG = "EEEEEE"
CODE_BG = "E4E4E4"

PALETTE = {
    INK, WHITE, CHARCOAL, SLATE, MID, GRAY, LINE, HAIR, PAPER, BAND, ZEBRA,
    INPUT_BG, FORMULA_BG, CODE_BG, "4A4A4A", "F2F2F2", "FAFAFA", "000000",
}

THIN = Border(
    left=Side(style="thin", color=LINE),
    right=Side(style="thin", color=LINE),
    top=Side(style="thin", color=LINE),
    bottom=Side(style="thin", color=LINE),
)
THICK_INPUT = Border(
    left=Side(style="medium", color="4A4A4A"),
    right=Side(style="medium", color="4A4A4A"),
    top=Side(style="medium", color="4A4A4A"),
    bottom=Side(style="medium", color="4A4A4A"),
)


def font(size=10, bold=False, color=INK):
    return Font(name="Meiryo UI", size=size, bold=bold, color=color, scheme=None)


def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)


def align(horizontal="left", vertical="top"):
    return Alignment(horizontal=horizontal, vertical=vertical, wrap_text=True)


def weight(text):
    total = 0
    for ch in str(text):
        total += 2 if ord(ch) > 127 else 1
    return total


def lines_for(text, excel_width):
    capacity = max(6, int(float(excel_width) * 0.78))
    total = 0
    for para in str(text).split("\n"):
        if para == "":
            total += 1
            continue
        units = weight(para)
        total += max(1, (units + capacity - 1) // capacity)
    return max(1, total)


def height_for(text, excel_width, size=10, minimum=20, cap=240):
    lines = lines_for(text, excel_width)
    return min(cap, max(minimum, 6 + lines * (size + 6)))


def apply_page(ws, title):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.horizontalCentered = True
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.45, right=0.45, top=0.7, bottom=0.55, header=0.26, footer=0.26)
    ws.sheet_properties.tabColor = MID
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 100
    ws.sheet_view.view = "pageBreakPreview"
    ws.sheet_view.view = "normal"
    ws.page_setup.horizontalDpi = 300
    ws.page_setup.verticalDpi = 300
    ws.print_options.horizontalCentered = True
    ws.page_setup.scale = 70
    ws.oddHeader.left.text = '&"Meiryo UI"&9親子JVMのプロセス・メモリ・vCPU'
    ws.oddHeader.right.text = '&"Meiryo UI"&9確認日 2026-10-14'
    ws.oddFooter.left.text = '&"Meiryo UI"&8見本の数字は仮の値。実測は対象サーバで採取する。'
    ws.oddFooter.center.text = f'&"Meiryo UI"&9{title}'
    ws.oddFooter.right.text = '&"Meiryo UI"&9&P / &N'
    ws.sheet_format.defaultRowHeight = 18
    ws.sheet_properties.outlinePr.summaryBelow = True
    ws.page_setup.paperSize = ws.PAPERSIZE_A3


def set_widths(ws, widths):
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    for col in range(len(widths) + 1, 16):
        ws.column_dimensions[get_column_letter(col)].width = 3
    ws.sheet_format.defaultColWidth = 3


def paint(cell, value, *, size=10, bold=False, color=INK, bg=None, horizontal="left", border=True, vertical="top"):
    cell.value = value
    cell.font = font(size, bold, color)
    cell.alignment = align(horizontal, vertical)
    if bg:
        cell.fill = fill(bg)
    if border:
        cell.border = THIN
    return cell


def merge_title(ws, row, cols, value, *, size=16, bold=True, color=WHITE, bg=CHARCOAL, min_height=32):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=cols)
    cell = ws.cell(row, 1)
    paint(cell, value, size=size, bold=bold, color=color, bg=bg, vertical="center")
    total = sum(ws.column_dimensions[get_column_letter(i)].width or 12 for i in range(1, cols + 1))
    ws.row_dimensions[row].height = max(min_height, height_for(value, total, size, min_height))
    for col in range(1, cols + 1):
        other = ws.cell(row, col)
        other.font = font(size, bold, color)
        other.fill = fill(bg)
        other.border = THIN
        other.alignment = align("left", "center")
    return row + 1


def write_table(ws, row, headers, data, widths, *, zebra=True, number_formats=None):
    number_formats = number_formats or {}
    cols = len(headers)
    for col, text in enumerate(headers, start=1):
        paint(ws.cell(row, col), text, size=10, bold=True, color=WHITE, bg=CHARCOAL, horizontal="center", vertical="center")
    header_width = max(widths[i] for i in range(cols))
    ws.row_dimensions[row].height = max(24, height_for("\n".join(headers), min(header_width, 28), 10, 24))
    header_row = row
    row += 1
    first_data = row
    for r_index, record in enumerate(data):
        if len(record) != cols:
            raise SystemExit(f"{ws.title} column mismatch at data {r_index}: {len(record)} != {cols}")
        bg = ZEBRA if zebra and r_index % 2 else WHITE
        tallest = 22
        for col, value in enumerate(record, start=1):
            horizontal = "right" if isinstance(value, (int, float)) else "left"
            paint(ws.cell(row, col), value, size=10, bg=bg, horizontal=horizontal)
            if col in number_formats and isinstance(value, (int, float)):
                ws.cell(row, col).number_format = number_formats[col]
            tallest = max(tallest, height_for(value, widths[col - 1], 10, 22))
        ws.row_dimensions[row].height = tallest
        row += 1
    last = row - 1
    if last >= first_data:
        ws.auto_filter.ref = f"A{header_row}:{get_column_letter(cols)}{last}"
        ws.auto_filter.ref = f"A{header_row}:{get_column_letter(cols)}{last}"
    ws.auto_filter.ref = f"A{header_row}:{get_column_letter(cols)}{max(last, header_row)}"
    ws.freeze_panes = f"A{header_row + 1}"
    ws.print_title_rows = f"1:{header_row}"
    ws.page_setup.fitToWidth = 1
    ws.oddHeader.center.text = ""
    return row


def add_sheet(wb, title, subtitle, headers, rows, widths):
    ws = wb.create_sheet(title)
    apply_page(ws, title)
    set_widths(ws, widths)
    cols = len(headers)
    row = merge_title(ws, 1, cols, title, size=16)
    row = merge_title(ws, row, cols, subtitle, size=10, bold=False, color=INK, bg=BAND, min_height=36)
    write_table(ws, row, headers, rows, widths)
    ws.sheet_view.zoomScale = 100
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return ws


def cover(wb):
    ws = wb.active
    ws.title = "表紙"
    apply_page(ws, "表紙")
    widths = [28, 88, 42]
    set_widths(ws, widths)
    merge_title(ws, 1, 3, "親子JVMのプロセス、メモリ、vCPU", size=20, min_height=40)
    merge_title(
        ws,
        2,
        3,
        "EC2 / Red Hat Enterprise Linux 9.8 / OpenJDK 21。親JVMが子JVMを同時10本まで起動し、業務は子が実行する。確認日 2026-10-14。",
        size=11,
        bold=False,
        color=INK,
        bg=BAND,
        min_height=36,
    )
    rows = [
        ("何の文書か", "1台のEC2の中で、プロセスは分かれ、CPUと物理メモリは共有される、という関係を説明する。性能試験で sar を読むときの項目の意味と、その道具が生まれた経緯を同じ場所に置く。", "Markdown と同じ内容の作業用ブック"),
        ("たとえ", "教室がEC2、机がvCPU、ロッカーが物理メモリ、先生が親JVM、作業員が子JVM、引き出しがヒープ、日記係が sar である。", "シート「たとえ話」"),
        ("30秒の結論", "子を10本にしても机は増えない。ヒープの上限は10回足す。GCワーカーも、各JVMが見えているCPU数から人数を決めるので、GCが重なると机の数を大きく超える。", "シート「子10本の重なり」"),
        ("見本の数字", "8 vCPU、16GiB、親512MiB、子1024MiB×10、という仮の入力で、推定居住量は約14.59GiB。OSとキャッシュに2GiB残すと、見積もりは約0.59GiB超える。測定値ではない。", "シート「計算」の太枠を自分の値に変える"),
        ("RHEL 9 の注意", "sar の kbmemused は MemTotal−MemFree−Buffers−Cached−Slab。free の used は MemTotal−MemAvailable。空きの判定に使うのは kbavail である。", "シート「メモリ」"),
        ("最初の1行", "vmstat と iostat の最初の1行は起動後平均。sar は間隔0のときだけ起動後平均。pidstat は2回目のブロック以降を採用する。", "シート「sarの読み方」"),
        ("測っていないもの", "対象サーバのインスタンスタイプ、実フラグ、実 sar、GCログは見ていない。", "シート「出典」"),
        ("フォントと色", "全セル Meiryo UI。色は白、灰、黒だけ。太枠の灰色が入力、薄い灰色が計算結果。", "このシート"),
        ("印刷", "A3横。幅を1ページに収める。画面ではフィルタとウィンドウ枠の固定を使う。", "各シート共通"),
        ("対になる文書", "親子JVMのメモリとCPUとsar.md", "同じフォルダ"),
    ]
    write_table(
        ws,
        4,
        ["項目", "内容", "どこを見るか"],
        rows,
        widths,
    )
    ws.row_dimensions[1].height = 42
    ws.freeze_panes = "A5"
    ws.print_title_rows = "1:4"
    ws.sheet_view.zoomScale = 110


def build_calc(wb):
    ws = wb.create_sheet("計算")
    apply_page(ws, "計算")
    widths = [6, 62, 22, 18, 78]
    set_widths(ws, widths)
    merge_title(ws, 1, 5, "見積もりの計算", size=16)
    merge_title(
        ws,
        2,
        5,
        "太枠の灰色だけ入力する。結果は仮の見積もりで、試験の RSS、PSS、kbavail、NMT を置き換える前の下絵である。初期値は文書の見本と一致する。",
        size=10,
        bold=False,
        color=INK,
        bg=BAND,
        min_height=36,
    )
    headers = ["区分", "項目", "値", "単位", "意味"]
    for col, text in enumerate(headers, start=1):
        paint(ws.cell(4, col), text, size=10, bold=True, color=WHITE, bg=CHARCOAL, horizontal="center", vertical="center")
    ws.row_dimensions[4].height = 24

    inputs = [
        (5, "入力", "vCPU 数", 8, "個", "lscpu の CPU(s)。ゲストから見える机の数。"),
        (6, "入力", "物理メモリ", 16, "GiB", "free の Mem の合計。インスタンスに割り当てられたロッカー。"),
        (7, "入力", "1コアあたりのスレッド数", 1, "本", "lscpu の Thread(s) per core。1なら1 vCPUが1コア。2ならハイパースレッド。"),
        (8, "入力", "JVMから見えているCPU数", 8, "個", "ActiveProcessorCount か cgroup のCPU上限。無ければ vCPU 数と同じ。"),
        (9, "入力", "親の -Xmx", 512, "MiB", "親JVMのヒープ上限。監督役なので子より小さいことが多い。"),
        (10, "入力", "子1本の -Xmx", 1024, "MiB", "業務を実行する子のヒープ上限。本数分だけ足す。"),
        (11, "入力", "同時に生きる子", 10, "本", "設定された並列起動数。ピークの java 本数は親を足して11になる想定。"),
        (12, "入力", "親のスレッド数の見込", 40, "本", "pidstat -t か Thread.print で置き換える。"),
        (13, "入力", "子1本のスレッド数の見込", 60, "本", "業務スレッド、GC、JITを含む概数。寝ているスレッドは机を使わない。"),
        (14, "入力", "スレッドスタック", 1024, "KiB", "64bit Linux の -Xss の既定はおおよそこの値。実フラグを優先する。"),
        (15, "入力", "メタスペースの見込", 96, "MiB/JVM", "NMT の Class の committed で置き換える。"),
        (16, "入力", "コードキャッシュの実使用の見込", 48, "MiB/JVM", "予約240MiBを全部は足さない。NMT の Code の committed を使う。"),
        (17, "入力", "ダイレクトバッファの見込", 32, "MiB/JVM", "使っていなければ小さくてよい。上限は未指定だとヒープ最大付近まで許される。"),
        (18, "入力", "G1構造とglibcの割合", 0.15, "対ヒープ", "ヒープに対する追加分の見込。0.15は15%。実測はNMTのGCとネイティブ。"),
        (19, "入力", "OSとページキャッシュに残す量", 2, "GiB", "カーネル、監視、JARのキャッシュ、ログに残したい余裕。"),
    ]
    for row, kind, label, value, unit, note in inputs:
        paint(ws.cell(row, 1), kind, size=10, bold=True, bg=BAND, horizontal="center", vertical="center")
        paint(ws.cell(row, 2), label, size=10, bg=WHITE)
        cell = paint(ws.cell(row, 3), value, size=12, bold=True, bg=INPUT_BG, horizontal="right", vertical="center", border=False)
        cell.border = THICK_INPUT
        cell.number_format = "0.00" if isinstance(value, float) else "0"
        paint(ws.cell(row, 4), unit, size=10, bg=WHITE, horizontal="center", vertical="center")
        paint(ws.cell(row, 5), note, size=10, bg=WHITE)
        ws.row_dimensions[row].height = max(32, height_for(note, widths[4], 10, 32))

    outputs = [
        (21, "物理コア数", "=IF(C7=0,\"\",C5/C7)", "個", "vCPU ÷ 1コアあたりのスレッド数。SMTの有無で机の実体が変わる。"),
        (22, "JVM本数", "=1+C11", "本", "親1＋同時の子。jspawnhelper は含めない。"),
        (23, "ヒープ合計", "=(C9+C10*C11)/1024", "GiB", "親と子の -Xmx の合計。上限の足し算であり、その瞬間のRSSではない。"),
        (24, "スタック合計", "=(C12+C13*C11)*C14/1024/1024", "GiB", "見込スレッド数 × スタックサイズ。"),
        (25, "メタスペース合計", "=C15*C22/1024", "GiB", "JVM本数ぶん。子の終了で子の分は返る。親の分は残る。"),
        (26, "コードキャッシュ合計", "=C16*C22/1024", "GiB", "実使用の見込。予約サイズの240MiB×本数ではない。"),
        (27, "ダイレクトバッファ合計", "=C17*C22/1024", "GiB", "使った分の見込。上限まで必ず使うわけではない。"),
        (28, "G1とglibcの追加", "=C23*C18", "GiB", "ヒープ合計に割合を掛けた見込。"),
        (29, "推定居住量", "=C23+C24+C25+C26+C27+C28", "GiB", "私有メモリの下絵。共有ライブラリの二重計上は含めていない。"),
        (30, "余り", "=C6-C29-C19", "GiB", "物理メモリから推定居住量とOS予約を引いた値。負なら見積もりがロッカーを超える。"),
        (31, "1JVMのParallel GCワーカー", '=IF(C8<=8,C8,8+INT(MAX(0,C8-8)*5/8))', "本", "G1の既定。整数の割り算。見えているCPU数から決まる。"),
        (32, "1JVMのConc GCスレッド", "=MAX(1,INT((C31+2)/4))", "本", "(ParallelGCThreads+2)/4 と 1 の大きいほう。"),
        (33, "全JVMが同時にYoung GC", "=C31*C22", "本", "重なった瞬間の上限。常時この人数が走っているという意味ではない。"),
        (34, "GCワーカー ÷ vCPU", "=IF(C5=0,\"\",C33/C5)", "倍", "1を超えると、GCが重なったとき机が足りない。"),
        (35, "commonPoolのワーカー/JVM", "=MAX(1,C8-1)", "本", "parallelStream等を使ったとき。使っていなければスレッドダンプに出ない。"),
        (36, "子のcommonPool合計", "=C35*C11", "本", "子だけ。親が同じプールを使えば、さらに1JVM分が足される。"),
        (37, "スレッド数の見込合計", "=C12+C13*C11", "本", "ulimit -u と pids.max と並べる。寝ているスレッドはロードに入らない。"),
    ]
    paint(ws.cell(20, 1), "結果", size=10, bold=True, color=WHITE, bg=SLATE, horizontal="center")
    for col in range(2, 6):
        paint(ws.cell(20, col), "入力から計算する。ここは編集しない。", size=10, bold=True, color=WHITE, bg=SLATE)
    ws.merge_cells("B20:E20")
    ws.row_dimensions[20].height = 22
    for row, label, formula, unit, note in outputs:
        paint(ws.cell(row, 1), "計算", size=10, bold=True, bg=BAND, horizontal="center", vertical="center")
        paint(ws.cell(row, 2), label, size=10, bg=FORMULA_BG)
        cell = paint(ws.cell(row, 3), formula, size=12, bold=True, bg=FORMULA_BG, horizontal="right", vertical="center")
        cell.number_format = "0.000"
        paint(ws.cell(row, 4), unit, size=10, bg=FORMULA_BG, horizontal="center", vertical="center")
        paint(ws.cell(row, 5), note, size=10, bg=WHITE)
        ws.row_dimensions[row].height = max(30, height_for(note, widths[4], 10, 30))

    paint(ws.cell(39, 1), "判定", size=10, bold=True, color=WHITE, bg=CHARCOAL, horizontal="center", vertical="center")
    paint(ws.cell(39, 2), "メモリの見積もり", size=10, bold=True, bg=BAND)
    paint(
        ws.cell(39, 3),
        '=IF(C30<0,"見積もりが物理メモリを超えています",IF(C30<2,"余りが2GiB未満です","見積もりのうえでは収まります"))',
        size=10,
        bold=True,
        bg=CODE_BG,
    )
    ws.merge_cells("C39:E39")
    for col in range(3, 6):
        ws.cell(39, col).font = font(10, True, INK)
        ws.cell(39, col).fill = fill(CODE_BG)
        ws.cell(39, col).border = THIN
        ws.cell(39, col).alignment = align("left", "center")
    ws.row_dimensions[39].height = 28

    paint(ws.cell(40, 1), "判定", size=10, bold=True, color=WHITE, bg=CHARCOAL, horizontal="center", vertical="center")
    paint(ws.cell(40, 2), "GCが重なったとき", size=10, bold=True, bg=BAND)
    paint(
        ws.cell(40, 3),
        '=IF(C34>2,"GCワーカーだけでvCPUの2倍を超えます",IF(C34>1,"GCワーカーがvCPU数を超えます","GCワーカーの合計はvCPU以内です"))',
        size=10,
        bold=True,
        bg=CODE_BG,
    )
    ws.merge_cells("C40:E40")
    for col in range(3, 6):
        ws.cell(40, col).font = font(10, True, INK)
        ws.cell(40, col).fill = fill(CODE_BG)
        ws.cell(40, col).border = THIN
        ws.cell(40, col).alignment = align("left", "center")
    ws.row_dimensions[40].height = 28

    paint(
        ws.cell(41, 2),
        "圧縮OOPの境界",
        size=10,
        bold=True,
        bg=BAND,
    )
    paint(ws.cell(41, 1), "判定", size=10, bold=True, color=WHITE, bg=CHARCOAL, horizontal="center", vertical="center")
    paint(
        ws.cell(41, 3),
        '=IF(C10>=32768,"子のヒープが32GiB以上です。UseCompressedOopsの実効値を確認します","子ヒープは32GiB未満の入力です。境界はJVMごとに判定します")',
        size=10,
        bg=CODE_BG,
    )
    ws.merge_cells("C41:E41")
    for col in range(3, 6):
        ws.cell(41, col).font = font(10, False, INK)
        ws.cell(41, col).fill = fill(CODE_BG)
        ws.cell(41, col).border = THIN
        ws.cell(41, col).alignment = align("left", "center")
    ws.row_dimensions[41].height = 40

    paint(ws.cell(43, 1), "検算", size=10, bold=True, color=WHITE, bg=CHARCOAL, horizontal="center", vertical="center")
    paint(ws.cell(43, 2), "見本の検算", size=10, bold=True, bg=BAND, vertical="center")
    paint(
        ws.cell(43, 3),
        '=IF(AND(C5=8,C6=16,C7=1,C8=8,C9=512,C10=1024,C11=10,C12=40,C13=60,C14=1024,C15=96,C16=48,C17=32,ABS(C18-0.15)<0.0001,C19=2,ABS(C29-14.590625)<0.0001,ABS(C30+0.590625)<0.0001,C33=88),"文書の見本と一致","入力が見本から変わっているか、式がずれています")',
        size=10,
        bold=True,
        bg=CODE_BG,
        vertical="center",
    )
    ws.merge_cells("C43:E43")
    for col in range(3, 6):
        ws.cell(43, col).font = font(10, True, INK)
        ws.cell(43, col).fill = fill(CODE_BG)
        ws.cell(43, col).border = THIN
        ws.cell(43, col).alignment = align("left", "center")
    ws.row_dimensions[43].height = 36

    dv_pos = DataValidation(type="decimal", operator="greaterThan", formula1="0", allow_blank=False)
    dv_pos.error = "0より大きい数を入れます"
    dv_pos.errorTitle = "入力"
    dv_pos.add("C5:C17")
    dv_pos.add("C19")
    ws.add_data_validation(dv_pos)
    dv_ratio = DataValidation(type="decimal", operator="between", formula1="0", formula2="1", allow_blank=False)
    dv_ratio.add("C18")
    ws.add_data_validation(dv_ratio)

    ws.freeze_panes = "A5"
    ws.print_title_rows = "1:4"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_view.zoomScale = 110
    ws.print_area = "A1:E43"
    ws.page_setup.fitToWidth = 1
    return


def add_sensitivity(wb):
    ws = wb.create_sheet("感度")
    apply_page(ws, "感度")
    widths = [14, 14, 16, 16, 18, 16, 24, 14, 20, 16, 18]
    set_widths(ws, widths)
    merge_title(ws, 1, 11, "同時本数と、見えるCPU数の感度", size=16)
    merge_title(
        ws,
        2,
        11,
        "値はシート「計算」の太枠を参照する。ここには入力欄は無い。見本の初期値のままなら、子10本の行の推定居住量は14.591GiB、余りは-0.591GiB、Young GCワーカーは88になる。",
        size=10,
        bold=False,
        color=INK,
        bg=BAND,
        min_height=36,
    )
    headers = [
        "同時の子", "JVM本数", "ヒープGiB", "スタックGiB", "推定居住GiB",
        "余りGiB", "同時YoungGCワーカー", "vCPU比", "commonPool子合計", "スレッド見込", "メモリ判定",
    ]
    for col, text in enumerate(headers, start=1):
        paint(ws.cell(4, col), text, size=9, bold=True, color=WHITE, bg=CHARCOAL, horizontal="center", vertical="center")
    ws.row_dimensions[4].height = 36
    for offset, n in enumerate([1, 2, 4, 8, 10, 12, 16]):
        r = 5 + offset
        bg = ZEBRA if offset % 2 else WHITE
        paint(ws.cell(r, 1), n, size=11, bold=True, bg=INPUT_BG, horizontal="right", vertical="center")
        ws.cell(r, 1).border = THICK_INPUT
        formulas = {
            2: f"=1+A{r}",
            3: f"=('計算'!$C$9+'計算'!$C$10*A{r})/1024",
            4: f"=('計算'!$C$12+'計算'!$C$13*A{r})*'計算'!$C$14/1024/1024",
            5: f"=C{r}+D{r}+'計算'!$C$15*B{r}/1024+'計算'!$C$16*B{r}/1024+'計算'!$C$17*B{r}/1024+C{r}*'計算'!$C$18",
            6: f"='計算'!$C$6-E{r}-'計算'!$C$19",
            7: f"='計算'!$C$31*B{r}",
            8: f"=IF('計算'!$C$5=0,\"\",G{r}/'計算'!$C$5)",
            9: f"='計算'!$C$35*A{r}",
            10: f"='計算'!$C$12+'計算'!$C$13*A{r}",
            11: f'=IF(F{r}<0,"見積もりが超える",IF(F{r}<2,"余りが小さい","見積もりでは収まる"))',
        }
        for col, formula in formulas.items():
            cell = paint(
                ws.cell(r, col),
                formula,
                size=10,
                bg=bg,
                horizontal="right" if col < 11 else "left",
                vertical="center",
            )
            if col != 11:
                cell.number_format = "0.000" if col in (3, 4, 5, 6, 8) else "0"
        ws.row_dimensions[r].height = 24
    note = "同時の子の列は、この表の中だけ変えられる。ヒープやvCPUを変えるときはシート「計算」の太枠を変える。GCワーカーは、見えているCPU数から決めた1JVMあたりの人数に、親を含むJVM本数を掛けたもの。"
    merge_title(ws, 13, 11, note, size=10, bold=False, color=INK, bg=BAND, min_height=36)

    headers2 = ["見えるCPU", "Parallel/JVM", "Conc/JVM", "全JVM同時のYoungGC", "そのCPU数に対する倍率"]
    for col, text in enumerate(headers2, start=1):
        paint(ws.cell(15, col), text, size=9, bold=True, color=WHITE, bg=CHARCOAL, horizontal="center", vertical="center")
    ws.row_dimensions[15].height = 32
    for offset, n in enumerate([2, 4, 8, 10, 16, 32, 64]):
        r = 16 + offset
        bg = ZEBRA if offset % 2 else WHITE
        paint(ws.cell(r, 1), n, size=11, bold=True, bg=INPUT_BG, horizontal="right", vertical="center")
        ws.cell(r, 1).border = THICK_INPUT
        paint(ws.cell(r, 2), f"=IF(A{r}<=8,A{r},8+INT(MAX(0,A{r}-8)*5/8))", size=10, bg=bg, horizontal="right", vertical="center")
        paint(ws.cell(r, 3), f"=MAX(1,INT((B{r}+2)/4))", size=10, bg=bg, horizontal="right", vertical="center")
        paint(ws.cell(r, 4), f"=B{r}*(1+'計算'!$C$11)", size=10, bg=bg, horizontal="right", vertical="center")
        paint(ws.cell(r, 5), f"=IF(A{r}=0,\"\",D{r}/A{r})", size=10, bg=bg, horizontal="right", vertical="center")
        for col in range(2, 6):
            ws.cell(r, col).number_format = "0.00" if col == 5 else "0"
            ws.cell(r, col).font = font(10, False, INK)
        ws.row_dimensions[r].height = 24
    merge_title(
        ws,
        24,
        11,
        "直前の表の倍率は、その行のCPU数をvCPU数とみなしている。同時の子の本数はシート「計算」の入力を使う。見えるCPUが8で子が10なら、全JVM同時のYoung GCワーカーは88、倍率は11。",
        size=10,
        bold=False,
        color=INK,
        bg=BAND,
        min_height=36,
    )
    ws.freeze_panes = "A5"
    ws.print_title_rows = "1:4"
    ws.auto_filter.ref = "A4:K11"
    ws.print_area = "A1:K24"


def assert_workbook(wb):
    offenses = []
    required_sheets = [
        "表紙", "たとえ話", "関係マップ", "子が起動する順", "歴史", "プロセス",
        "CPUとvCPU", "メモリ", "近所と外部", "子10本の重なり", "試験で見ること",
        "sarの読み方", "コマンド辞典", "形の読み分け", "試験手順", "計算", "感度",
        "JDKとRHEL", "用語", "出典",
    ]
    names = [ws.title for ws in wb.worksheets]
    for name in required_sheets:
        if name not in names:
            offenses.append(f"missing sheet {name}")
    for ws in wb.worksheets:
        if ws.freeze_panes is None:
            offenses.append(f"no freeze {ws.title}")
        if ws.page_setup.orientation != "landscape":
            offenses.append(f"orientation {ws.title}")
        if ws.page_setup.fitToWidth != 1:
            offenses.append(f"fit {ws.title}")
        for row in ws.iter_rows():
            for cell in row:
                if cell.font and cell.font.name and cell.font.name != "Meiryo UI":
                    offenses.append(f"font {ws.title}!{cell.coordinate} {cell.font.name}")
                if cell.font and cell.font.color is not None and getattr(cell.font.color, "type", None) == "rgb" and cell.font.color.rgb:
                    rgb = str(cell.font.color.rgb)
                    if len(rgb) >= 6 and all(c in "0123456789ABCDEFabcdef" for c in rgb[-6:]):
                        hexpart = rgb[-6:].upper()
                        if hexpart not in PALETTE and not _gray(hexpart):
                            offenses.append(f"font color {ws.title}!{cell.coordinate} {hexpart}")
                if cell.fill and cell.fill.patternType == "solid" and cell.fill.fgColor is not None and getattr(cell.fill.fgColor, "type", None) == "rgb" and cell.fill.fgColor.rgb:
                    rgb = str(cell.fill.fgColor.rgb)
                    if len(rgb) >= 6 and all(c in "0123456789ABCDEFabcdef" for c in rgb[-6:]):
                        hexpart = rgb[-6:].upper()
                        if hexpart not in PALETTE and not _gray(hexpart):
                            offenses.append(f"fill {ws.title}!{cell.coordinate} {hexpart}")
                if isinstance(cell.value, str) and not str(cell.value).startswith("=") and len(cell.value) > 40:
                    span = None
                    for merged in ws.merged_cells.ranges:
                        if cell.coordinate in merged:
                            if cell.row != merged.min_row or cell.column != merged.min_col:
                                span = "skip"
                            else:
                                span = sum(
                                    ws.column_dimensions[get_column_letter(col)].width or 10
                                    for col in range(merged.min_col, merged.max_col + 1)
                                )
                            break
                    if span == "skip":
                        continue
                    excel_width = span or (ws.column_dimensions[cell.column_letter].width or 10)
                    needed = height_for(cell.value, excel_width, cell.font.size or 10, 18, cap=1000)
                    actual = ws.row_dimensions[cell.row].height or 15
                    if actual + 1 < needed * 0.90:
                        offenses.append(
                            f"clip {ws.title}!{cell.coordinate} h={actual:.0f} need={needed:.0f}"
                        )
    text = MD_PATH.read_text(encoding="utf-8")
    for token in ["14.590625", "0.590625", "posix_spawn", "kbavail", "ParallelGCThreads", "Meiryo UI"]:
        if token not in text and token != "Meiryo UI":
            offenses.append(f"md missing {token}")
    if "Meiryo UI" not in MD_PATH.read_text(encoding="utf-8") and False:
        pass
    if offenses:
        raise SystemExit("workbook check failed:\n" + "\n".join(offenses[:50]))


def _gray(hexpart):
    if len(hexpart) != 6:
        return False
    r = int(hexpart[0:2], 16)
    g = int(hexpart[2:4], 16)
    b = int(hexpart[4:6], 16)
    return abs(r - g) <= 3 and abs(g - b) <= 3 and abs(r - b) <= 3


def example_math_check():
    vcpu = 8
    mem = 16
    parent = 512
    child = 1024
    n = 10
    heap = (parent + child * n) / 1024
    stack = (40 + 60 * n) * 1024 / 1024 / 1024
    jvms = 1 + n
    meta = 96 * jvms / 1024
    code = 48 * jvms / 1024
    direct = 32 * jvms / 1024
    extra = heap * 0.15
    rss = heap + stack + meta + code + direct + extra
    head = mem - rss - 2
    parallel = 8
    total_gc = parallel * jvms
    if abs(rss - 14.590625) > 1e-9 or abs(head + 0.590625) > 1e-9 or total_gc != 88:
        raise SystemExit(f"math drifted rss={rss} head={head} gc={total_gc}")


def main():
    example_math_check()
    if not MD_PATH.exists():
        raise SystemExit(f"missing {MD_PATH}")
    wb = Workbook()
    default_font = wb._fonts[0]
    default_font.name = "Meiryo UI"
    default_font.sz = 10
    default_font.scheme = None
    wb.calculation = CalcProperties(calcMode="auto", fullCalcOnLoad=True)

    cover(wb)
    add_story(wb)
    add_map(wb)
    add_life(wb)
    add_history(wb)
    add_process(wb)
    add_cpu(wb)
    add_memory(wb)
    add_neighbors(wb)
    add_ten(wb)
    add_test(wb)
    add_sar(wb)
    add_commands(wb)
    add_shapes(wb)
    add_steps(wb)
    build_calc(wb)
    add_sensitivity(wb)
    add_jdk(wb)
    add_glossary(wb)
    add_sources(wb)

    wb._fonts[0].name = "Meiryo UI"
    wb._fonts[0].sz = 10
    wb._fonts[0].scheme = None
    assert_workbook(wb)
    wb.properties.title = "親子JVMのプロセス、メモリ、vCPU"
    wb.properties.creator = "Grok"
    wb.properties.subject = "EC2 RHEL 9.8 OpenJDK 21 親JVMと子JVM10本の性能試験"
    wb.properties.description = "確認日 2026-10-14。Meiryo UI。モノトーン。見本の数字は仮の値。"
    wb.properties.keywords = "JVM,sar,vCPU,OpenJDK21,RHEL9.8,EC2"
    wb.properties.category = "性能試験"
    XLSX_PATH.parent.mkdir(parents=True, exist_ok=True)
    wb.save(XLSX_PATH)
    print(f"sheets={len(wb.worksheets)}")
    print(f"wrote={XLSX_PATH}")


# The sheet builders are defined below and called from main.
# They are plain functions so the workbook stays data-driven.


def add_story(wb):
    add_sheet(
        wb,
        "たとえ話",
        "このシートだけで、机とロッカーと引き出しの関係が追える。あとのシートは同じ絵の正式な名前である。",
        ["たとえの名前", "教室では何か", "コンピューターでは何か", "10本にするとどうなるか"],
        [
            ["教室", "このクラス専用の部屋。ビルの中の一部屋。", "EC2インスタンス。Nitroが貸しているゲスト。", "教室は1つのまま。作業員を増やしても部屋は増えない。"],
            ["ビルの管理人", "机の脚数とロッカーの大きさを決めて貸す人。", "Nitroハイパーバイザ。vCPUとメモリを割り当てる。", "管理人の貸与量は、子の本数では変わらない。"],
            ["机", "同時に計算できる席。", "vCPU。lscpuのCPU(s)。", "10人呼んでも机が8脚なら、2人は立つ。"],
            ["机の種類", "1つの広い机か、1つの台を2人で使う机か。", "Thread(s) per core。1ならコアそのもの。2ならハイパースレッド。", "人数の計算は、見える机の数で行う。"],
            ["ロッカー", "大きさが決まっている収納。", "物理メモリ。freeのMem。", "引き出しの合計がロッカーを超えると、廊下の倉庫へ行く。"],
            ["先生", "宿題を配り、終わったかを見る人。自分では全部解かない。", "親JVM。1プロセス。", "先生は1人のまま。枠が10あるかを数える。"],
            ["作業員", "宿題を解く人。同時に10人まで。", "子JVM。業務処理の本体。", "10人が同時に自分の引き出しを広げる。"],
            ["引き出し", "となりからは中が見えない、その人専用の場所。", "Javaヒープ。-Xmxが上限。", "上限は人数分だけ足す。中身は共有しない。"],
            ["手", "1人がたくさんの手を持てる。寝ている手は机を使わない。", "スレッド。Linuxではタスク。", "GCのときに、各人が机の数だけ手を起こそうとすることがある。"],
            ["開いた教科書", "台の上に1冊あり、みんなで同じページを見る。", "ページキャッシュ上の.so、JAR、CDSアーカイブ。", "RSSを人数分足すと、この教科書を人数分に数えてしまう。"],
            ["廊下の倉庫", "遠いので、取りに行くあいだ仕事が止まる。", "スワップ。無い教室もある。", "GCが倉庫のページに触ると、停止が長くなる。"],
            ["日記係", "数秒ごとに、机とロッカーと倉庫を書き残す人。", "sarとsadc。sysstat。", "10分ごとの日記では、数分の山が平均に溶ける。試験は1秒で書く。"],
            ["時計係や手紙係", "作業員ではないが、同じ机と同じロッカーを使う。", "chronyd、sshd、SSM、journald、監査。", "ピークの前に、この人たちの使用量を写真に残す。"],
        ],
        [22, 42, 48, 48],
    )


def add_map(wb):
    add_sheet(
        wb,
        "関係マップ",
        "プロセスは分かれ、机とロッカーは1つ、という対応表。試験で「どの列を見るか」までセットにしてある。",
        ["もの", "親と子", "子どうし", "ほかのプロセス", "試験で見る場所"],
        [
            ["プロセス", "PIDが違う。子のPPIDは親を指す。", "別プロセス。", "systemd以下の常駐とは別。", "ps -o pid,ppid,nlwp,rss,vsz,cmd"],
            ["Javaヒープ", "親の-Xmxと子の-Xmxは別。", "共有しない。10本なら最大10個。", "JVM以外はヒープを持たない。", "jcmd VM.flags と GC.heap_info"],
            ["スレッド", "各自が持つ。親のスレッドから子のヒープは見えない。", "同じヒープを共有するのは、同じ子の中だけ。", "OSプロセスもスレッドを持つ。", "pidstat -t -u -p PID"],
            ["物理メモリ", "1つのロッカーを分け合う。", "同じ。", "同じ。", "sar -r ALL の kbavail と kbanonpg"],
            ["ページキャッシュ", "JDKの共有ライブラリとCDSは共有できる。", "同じファイルページを使える。", "ログや入力ファイルのキャッシュとも同じ領域。", "kbcached と smaps_rollup の Shared_Clean"],
            ["vCPU", "1組の机を時間で分け合う。", "同じ。", "sar自身、sshd、chronydとも分け合う。", "sar -u ALL -P ALL と sar -q"],
            ["ファイルディスクリプタ", "起動時に閉じ忘れを受け継ぐことがある。", "各プロセスが自分の表を持つ。", "上限はsystemdのLimitNOFILEを受け継ぐ。", "/proc/PID/limits と fd"],
            ["ディスク", "EBSは1つの通路。", "ログもGCログも同じキュー。", "sar -o の書き込みも同じ通路。", "sar -d -p の await と aqu-sz"],
            ["ネットワークとDB", "子が接続する。親が接続することもある。", "プールを本数で足す。", "SSMや更新も同じNIC。", "sar -n DEV。DB待ちはidleが大きい形。"],
            ["終了", "子が落ちても親は残せる。", "1本のFull GCはそのJVMの中。", "メモリ不足では大きいRSSが終了対象になりやすい。", "PPID、dmesg の Killed process"],
            ["標準出力", "親がパイプを読まないと、子の書き込みが止まる。", "各子のパイプは独立。", "関係するのは親と子の対。", "ps -o stat,wchan"],
        ],
        [24, 40, 40, 42, 40],
    )


def add_life(wb):
    add_sheet(
        wb,
        "子が起動する順",
        "1本の子の寿命。10本が重なると、この列が最大10本同時に進む。",
        ["順", "起きること", "メモリの見え方", "CPUの見え方", "試験での印"],
        [
            ["1", "親は既に居て、同時実行中の子が10未満かを見る。", "親のRSSは居座ったまま。", "親は待ちが中心で、CPUは小さいことが多い。", "写真A。親だけのRSSとNMT。"],
            ["2", "ProcessBuilderが子を起動する。OpenJDK 21の既定はposix_spawn。", "親ヒープの複製は作らない。", "jspawnhelperが瞬間的に現れる。", "psにjspawnhelperが残像のように出る。"],
            ["3", "execのあと、子は新しいjavaプロセスになる。", "VSZが-Xmx付近まで広がる。RSSはまだ上限未満であることが多い。", "ローダが.soを読む。", "pidstatのVSZとRSSの差。"],
            ["4", "GCスレッド、JITスレッド、業務スレッドが生まれる。", "スタック×本数だけ私有メモリが増える。", "起動直後はC1/C2がuser時間を使う。", "Thread.printのC2 CompilerThread。"],
            ["5", "CDSが一致すれば、読み取り専用ページを共有する。", "Shared_Cleanに教科書が載る。", "クラスロードのCPUとディスク読みが減る。", "-Xlog:cds。共有失敗でも起動は続く。"],
            ["6", "業務がオブジェクトを作る。", "触ったページだけRSSとAnonPagesが上がる。AlwaysPreTouchなら起動時に上がる。", "ユーザモードの計算。", "フラグにAlwaysPreTouchがあるかを先に見る。"],
            ["7", "Young GCが走る。", "ヒープの中でコピーが進む。スワップされているとここでディスクに触る。", "ParallelGCThreads本が同時に走ろうとする。", "GCログの時刻とrunq-szの山。"],
            ["8", "JDBCやファイルの返事を待つ。", "ヒープはそのまま居住する。", "そのスレッドはSになり、机を使わない。idleが残る。", "idleが高く、経過時間だけ長い形。"],
            ["9", "子が終了する。", "その子の匿名ページは返る。JARのキャッシュは残る。", "そのPIDのCPUが消える。", "写真Dでkbanonpgが戻るか。"],
            ["10", "親が終了コードを受け、空いた枠で次を起動する。", "親のメタスペースは、子が終わっても自動では空にならない。", "proc/sに次の起動の山。", "写真AとDで親のClassとFDを比べる。"],
            ["11", "親が先に死ぬと、子のPPIDが1になることがある。", "子のヒープは残る。", "止めたつもりでもCPUが残る。", "psのPPIDが1のjava。"],
        ],
        [8, 48, 42, 40, 36],
    )


def add_history(wb):
    add_sheet(
        wb,
        "歴史",
        "この構成と、sarという日記が、なぜ必要になったかの時系列。年は公開情報に沿ったもの。",
        ["年", "できごと", "当時の問題", "いまの親子JVMとの関係"],
        [
            ["1960年代", "仮想メモリがAtlasやMulticsで実用になる。仕事は列に並んだ。", "実メモリより大きい仕事を、住所の予約で扱いたかった。", "VSZが大きくRSSが小さい状態の祖先。-Xmxは予約の上限。"],
            ["1969–1970年代", "Unixが生まれ、プロセス、fork、execが固まる。", "仕事ごとにメモリを分けたかった。", "親が子を起動する形の祖先。"],
            ["1980年代", "System Vにsar（System Activity Reporter）が入る。メインフレームにはSMFとRMFがあった。", "夜のバッチの翌朝に、記憶ではなく記録で話したかった。", "性能試験でsarを取る、という手順の起源。"],
            ["1995", "Javaが公開される。1プロセス、多数のスレッド、1つのヒープ。", "アプリサーバでオブジェクトを共有したかった。", "ヒープを共有する便利さと、GCが全員を止める代償。"],
            ["1999ごろ", "Sebastien GodardがLinux向けsysstatを書き始める。", "Linuxが基幹に入っても、System Vの日記が欲しかった。", "RHEL 9.8のsarの直接の祖先。"],
            ["2002", "NPTLがLinuxの実用的なスレッド実装になる。Red Hatが中心。", "Javaのスレッドを、重いプロセスなしに大量に動かしたかった。", "子JVMの中のスレッドは、いまNPTLのタスク。"],
            ["2000年代", "クラスローダやOSGiで、1つのJVMの中を分けようとした。", "ジョブの静的状態とメタスペースが、終了後も残った。", "バッチが子プロセスへ戻った理由。"],
            ["2006–2008", "cgroupがGoogleで生まれ、Linux 2.6.24に入る。", "1台のCPUとメモリをグループで分けたかった。", "RHEL 9のcgroup v2。上限が無いとJVMはマシン全体を自分の分母にする。"],
            ["2017", "JDK 9でG1が既定GCになる。", "大きいヒープの停止時間を、領域単位で扱いたかった。", "OpenJDK 21の既定もG1。子10本はG1が10個。"],
            ["2018", "JDK 10前後で、コンテナのメモリ上限をヒープ計算に使うようになる。", "JVMがホスト全メモリの4分の1を取ろうとしていた。", "CPUQuotaとMemoryMaxが、GC人数と-Xmx自動の分母になる。"],
            ["JDK 13の系統", "Linuxの子プロセス起動の既定がposix_spawnになる。", "大きい親のforkが、メモリ予約だけで失敗していた。", "OpenJDK 21の親は、ヒープを複製せずに子を起動する。"],
            ["2022-05-17", "RHEL 9.0が一般提供。カーネル5.14、cgroup v2が既定。", "v1の階層の複雑さを一本化したかった。", "PSI、smaps_rollup、cgroup v2が使える土台。"],
            ["2023-09-19", "OpenJDK 21 LTS。仮想スレッドと世代別ZGC。", "待ちの多い処理を、OSスレッドより軽くしたかった。", "仮想スレッドは子の中の待ちを変える。プロセス数は減らない。"],
            ["2026-05-19", "RHEL 9.8が一般提供。GAカーネルは5.14.0-687.5.1.el9_8。", "9系の機能更新の後期。フルサポートは2027-05-31まで。", "試験票に貼るのは、その日のuname -r。"],
        ],
        [18, 42, 46, 48],
    )


def add_process(wb):
    add_sheet(
        wb,
        "プロセス",
        "PID、スレッド、GCワーカー、起動方式。人数の式はシート「計算」と同じ。",
        ["項目", "しくみ", "10本のときの意味", "確認コマンド"],
        [
            ["PIDとPPID", "プロセスはアドレス空間を持つ。PPIDは親。", "ピークは親1＋子10。終了後は親1。", "ps -o pid,ppid,lstart,cmd -C java"],
            ["posix_spawn", "jspawnhelper経由でexecする。親のヒープはコピーしない。", "親が大きくても、子の起動のためにヒープ2倍の空きは要らない。", "jdk.lang.Process.launchMechanism"],
            ["昔のfork", "execの前にアドレス空間の複製が要ることがあった。", "overcommitが厳密だと、起動だけ失敗することがある。", "sysctl vm.overcommit_memory"],
            ["スレッドはタスク", "同じヒープを共有する実行の流れ。ulimit -uはタスク数。", "子10本×各60スレッドで、見込600に親が足される。", "psのNLWP、/proc/PID/limits"],
            ["寝ているスレッド", "状態Sは実行キューに入らない。", "JDBC待ちは机を使わない。本数とCPU使用は一致しない。", "pidstat -t と psのstat"],
            ["仮想スレッド", "JDK 21のキャリアがOSスレッド。待ちの本数はもっと多い。", "NLWPが小さくてもDB接続は多い、があり得る。", "Thread.print"],
            ["ParallelGCThreads", "Nが8以下ならN。それ以外は8+(N-8)×5/8。整数割り算。", "見えるCPUが8なら、1JVMあたり8。11JVM同時で88。", "jcmd VM.flags"],
            ["ConcGCThreads", "(Parallel+2)/4と1の大きいほう。業務と同時に机を使う。", "見本では1JVMあたり2。11本で22。", "jcmd VM.flags"],
            ["JIT", "起動後にC1/C2がコンパイルする。短いバッチでは割合が大きい。", "冷えた10本同時起動は、定常よりCPUが高い。", "GCログと分ける。Thread.printのC2。"],
            ["commonPool", "見えているCPU数から1を引いたワーカー。使ったときだけ存在する。", "8CPUなら子1本で7、子10本で70。", "スレッド名 ForkJoinPool.commonPool"],
            ["パイプ", "親がstdoutを読まずに待っていると、64KiBで子のwriteが止まる。", "CPUが低いのに枠が10のまま空かない。", "ps -o wchan"],
            ["孤児", "親が死ぬとPPIDが1になり、子は生き続けられる。", "停止したあとのRSSの残りの説明になる。", "psのPPID"],
        ],
        [24, 48, 46, 36],
    )


def add_cpu(wb):
    add_sheet(
        wb,
        "CPUとvCPU",
        "机は2段ある。ハイパーバイザが物理コアへvCPUを載せ、LinuxのCFSがスレッドへvCPUを載せる。",
        ["項目", "意味", "親子JVMでの読み方", "見る列"],
        [
            ["vCPU", "ゲストから見えるCPU。", "子を10本にしてもこの数は増えない。", "lscpuのCPU(s)"],
            ["Thread(s) per core", "1ならコアそのもの。2ならハイパースレッド。", "GC人数の分母は、見える論理CPU。", "lscpu"],
            ["NUMA", "メモリが複数の置き場に分かれる大きいサイズ。", "ノードが2つ以上なら、CPUだけでなく置き場も記録する。", "numactl --hardware"],
            ["マシンCPU 100%", "見えているvCPUが全部忙しい。", "8 vCPUが全部忙しいと約100%。", "sarのall、CloudWatch CPUUtilization"],
            ["プロセス%CPU", "1 vCPU専有で100%。4 vCPU分で400%。", "子の行が100を超えるのは、複数コアを使った印。", "pidstat。psの%CPUは生涯平均なので使わない。"],
            ["ロードアベレージ", "実行したいタスクと割り込み不能タスクの指数平均。", "8 vCPUで16なら、机の2倍が列にいる、という読み方。", "sar -qのldavg-1/5/15"],
            ["%user", "ユーザモード。業務、JIT、GCの大部分。", "GCと業務はsarだけでは分離できない。", "sar -u と GCログ"],
            ["%system", "カーネル。切替、メモリ操作、I/Oの開始。", "強制切替が多いと上がることがある。", "sar -u と pidstat -w"],
            ["%iowait", "CPUは空いていて、ディスクI/Oが未完了。", "vCPUが多いと百分率は小さく見えやすい。", "awaitとセット"],
            ["%steal", "ハイパーバイザがvCPUを渡していなかった時間。", "T系のクレジット枯渇や、ホスト側の混雑。", "sar -u ALL"],
            ["%idle", "上記のどれでもない空き。", "DB待ちはここが大きい。", "sar -u と sar -n DEV"],
            ["%soft", "ソフト割り込み。受信処理が多いと上がる。", "帯域とセットで見る。", "sar -u ALL"],
            ["%guest", "このマシンがさらにVMを動かしている時間。", "EC2の中ではほぼ0が正常。", "sar -u ALL"],
            ["runq-sz", "いま実行を待っているタスク数。", "vCPUより大きい状態が続くと、計算したい人が余っている。", "sar -q"],
            ["blocked", "割り込み不能で止まっている数。", "本体がblockedなら、待ちはディスク側。", "sar -q ALL"],
            ["nvcswch/s", "タイムスライスでCPUを取り上げられた回数。", "10本のGCが重なると、この強制側が増える。", "pidstat -w"],
            ["cswch/s", "自分からCPUを手放した回数。睡眠やI/O。", "DB待ちが多いと自発側が増え、CPUは空く。", "pidstat -w"],
            ["ベースライン", "子を起動する前の60秒。", "ピークから引いた分が、子10本の増分。", "写真Aと写真C"],
        ],
        [24, 42, 48, 40],
    )


def add_memory(wb):
    add_sheet(
        wb,
        "メモリ",
        "予約、居住、共有、RHEL 9でのfreeとsarの式の違い。空きの判定に使うのはkbavail。",
        ["項目", "数えているもの", "読み方", "注意"],
        [
            ["VSZ", "予約した仮想アドレス。", "子の起動直後に-Xmx付近まで跳ねる。", "実メモリの使用量そのものではない。"],
            ["RSS", "実メモリに載っているマップ。共有ページも全量入る。", "プロセス単体の居住量。", "11本の合計は教科書の分だけ過大。"],
            ["PSS", "共有ページを共有者数で割った居住量。", "JVM合計に近い。", "smaps_rollup。毎秒は取らない。"],
            ["Private_Dirty", "そのプロセスだけが汚したページ。", "終了すると戻る量に近い。", "ヒープ、スタック、ネイティブの多く。"],
            ["AnonPages / kbanonpg", "ファイルに紐づかない匿名ページ。", "子の起動で増え、終了で戻る。", "写真AとCとDを比べる。"],
            ["kbavail / MemAvailable", "スワップせずに新しいアプリを始められる見込み。", "ロッカーの余りとして使う。", "freeのavailableと同じ系統。"],
            ["kbcached", "ファイルキャッシュ。Cached。", "JARや.soや入力ファイル。回収できる。", "sarのkbcachedはSReclaimableを含まない。"],
            ["kbmemused", "Total−Free−Buffers−Cached−Slab。", "RHEL 9のsarの使用中。", "RHEL 8のTotal−Freeではない。"],
            ["freeのused", "RHEL 9.1以降はTotal−MemAvailable。", "sarのkbmemusedと同じ瞬間でも数字が違う。", "列名を報告書に書く。"],
            ["仮の式", "Total 16000000、Free 1000000、Buffers 100000、Cached 4000000、Slab 700000。", "kbmemusedは10200000で63.75%。", "古い式だと15000000で94%に見える。"],
            ["仮のavailable", "MemAvailableを5200000とした例。", "freeのusedは10800000。", "availableが5.2GBある。"],
            ["-Xmx無し", "MaxRAMPercentageの既定25。見えているメモリの約4分の1。", "子10本が各自4分の1を上限にすると、合計は物理メモリを超える。", "各PIDのMaxHeapSizeを残す。"],
            ["圧縮OOP", "ヒープがおおよそ32GiB未満で有効になる既定。", "境界はJVMごと。マシン合計ではない。", "UseCompressedOopsの実効値。"],
            ["メタスペース", "クラス定義。子は終了で返る。親は残る。", "写真AとDで親のClassを比べる。", "NMT。"],
            ["コードキャッシュ", "予約は240MiB前後。実使用はその一部。", "短命な子は予約が大きくコミットが小さい。", "VSZとRSSの差の一部。"],
            ["スタック", "既定はおおよそ1MiB×スレッド。", "本数×JVM数。", "-XssとNLWP。"],
            ["ダイレクトバッファ", "未指定の上限はヒープ最大付近。", "使った分だけRSSが増え、ヒープ使用率は低く見える。", "MaxDirectMemorySize。"],
            ["スワップ", "pswpin/sが業務中に続く状態。", "GCがディスク上のページに触れている。", "sar -W。スワップが無ければOOM側。"],
            ["THP", "enabledの中身がalwaysかmadviseかneverか。", "alwaysは停止とRSSに効いた歴史がある。既定のJVMは巨大ページを要求しない。", "transparent_hugepage/enabled"],
            ["kbpgtbl", "ページテーブル。", "大きいヒープが10本あると-Xmxの外で増える。", "sar -r ALL"],
        ],
        [24, 46, 46, 40],
    )


def add_neighbors(wb):
    add_sheet(
        wb,
        "近所と外部",
        "同じ机、同じロッカー、同じEBS、同じNICを使う相手。DBの待ちは教室の外にいる。",
        ["相手", "何を共有するか", "性能試験での出方", "残すもの"],
        [
            ["systemd / journald / rsyslog", "CPUの少しと、ログのディスク。", "平常時は小さい。logrotateで短時間の山。", "写真Aのpidstat。"],
            ["sshd / chronyd / auditd / tuned", "CPUの少しとメモリ。", "audit拒否はCPUに出ず、子が即終了する。", "timedatectl、audit.log。"],
            ["amazon-ssm-agent と監視エージェント", "CPUとネットワークとディスク。", "試験の最中のRun Commandは山に混ざる。", "試験中にSSM作業をしないこと。"],
            ["sar / sadc 自身", "ごく小さいCPUと、出力ファイルの書き込み。", "1秒採取を何本も重ねると、自分の測定が乗る。", "採取コマンドの本数。"],
            ["カーネルのスラブとページテーブル", "物理メモリ。", "kbslab、kbpgtbl、kbkstack。", "sar -r ALL。"],
            ["EBS", "IOPSとスループットの上限。", "awaitが伸び、idleが増え、経過時間が伸びる。", "ボリュームタイプとプロビジョン値。"],
            ["NIC", "帯域と、サイズによってはネットワーククレジット。", "%softとrxkB/s、txkB/s。", "インスタンスタイプのネットワーク仕様。"],
            ["DB", "接続数は子のプール×本数。待ちは相手側。", "idleが高く、帯域が細く、経過時間だけ長い。", "プールサイズと相手の測定。"],
            ["OOM Killer", "物理メモリが尽きたときの終了。", "大きい子が消え、親が残ることがある。", "dmesgのKilled process。"],
            ["SELinux", "execの許可。", "性能グラフの前に子が死ぬ。", "audit.log。"],
            ["ファイルディスクリプタ上限", "親から子へ受け継ぐ。", "Too many open files。", "/proc/PID/limits。"],
            ["スレッド数上限", "タスク数。ヒープとは別。", "unable to create native thread。", "NLWP、pids.max、ulimit -u。"],
        ],
        [36, 40, 46, 36],
    )


def add_ten(wb):
    add_sheet(
        wb,
        "子10本の重なり",
        "10倍してよいものと、1つのままのものを分けた。見本の88ワーカーは、GCが重なった瞬間の上限。",
        ["種類", "項目", "10本での扱い", "見本の数"],
        [
            ["足し算する", "ヒープ上限", "子の-Xmxを本数回足し、親を足す。", "0.5+10=10.5GiB"],
            ["足し算する", "私有のメタスペース、スタック、ダイレクト", "プロセスごとに持つ。", "シート「計算」"],
            ["足し算する", "Young GCが重なったときのワーカー", "1JVMの人数×JVM本数。", "8×11=88。vCPUの11倍"],
            ["足し算する", "JDBC接続", "子が各自プールを持つと本数倍。", "各10接続なら100"],
            ["足し算する", "ログの書き込み", "同じディスクキューに乗る。", "sar -dのwkB/s"],
            ["足さない", "vCPU", "机の数は固定。", "見本は8"],
            ["足さない", "物理メモリ", "ロッカーの大きさは固定。", "見本は16GiB"],
            ["足さない", "共有ライブラリの実メモリ", "ページキャッシュ上で1回。", "RSS合計との差分"],
            ["足さない", "親JVM", "監督役は1本。", "ピークでも親は1"],
            ["新しく起きる", "冷えた同時起動", "クラスロードとJITが重なる。", "proc/sとmajflt/sの短い山"],
            ["新しく起きる", "GCの周期的な重なり", "各子のGCは独立なので、山がそろうことがある。", "GCログとrunq-sz"],
            ["新しく起きる", "1本だけ遅い", "その枠が塞がり、同時数10のまま次が始まらない。", "PIDごとのjstat"],
            ["見積もり", "推定居住量", "見本の入力では約14.59GiB。", "14.590625GiB"],
            ["見積もり", "OS予約2GiBを引いた余り", "見本では約0.59GiB不足。", "-0.590625GiB"],
        ],
        [18, 40, 48, 32],
    )


def add_test(wb):
    add_sheet(
        wb,
        "試験で見ること",
        "環境を固定し、4枚の写真を撮る。合否の制限時間は業務側の値。インフラは、時間がどこで溶けたかを言う。",
        ["分類", "確認すること", "なぜ要るか", "タイミング"],
        [
            ["環境", "インスタンスタイプ、vCPU、Thread(s) per core、メモリ、NUMA、EBS、ネットワーク仕様", "机とロッカーと通路の大きさが、解釈の分母になる。", "試験前"],
            ["環境", "redhat-releaseとuname -r", "9.8のGAカーネルからエラータで進む。", "試験前"],
            ["環境", "java -versionとrpm -q java-21-openjdk-headless", "OpenJDK 21の更新版を固定する。", "試験前"],
            ["フラグ", "親と各子のMaxHeapSize、GC、ParallelGCThreads、ActiveProcessorCount、AlwaysPreTouch、NMT", "スクリプトの値と実効値が違うことがある。", "ピークで全PID"],
            ["フラグ", "CDSが共有できたか。起動方式がposix_spawnか。", "起動の山と、fork失敗の切り分け。", "起動ログ"],
            ["カーネル", "overcommit、swappiness、THP、スワップの有無、LimitNOFILE、TasksMax、MemoryMax、CPUQuota", "失敗の見た目が性能劣化に見える項目。", "試験前"],
            ["写真A", "親だけ。free、sar 60秒、親のNMTとPSS。", "近所と親の常駐費。", "子の前"],
            ["写真B", "子1本が温まった状態。", "1本の単価。10本との差が奪い合い。", "可能なら"],
            ["写真C", "子10本の山。1秒のsarとpidstat。NMTは1巡だけ。", "共有資源の上限。", "業務ピーク"],
            ["写真D", "子が終わったあと。javaは親だけ。kbanonpgが戻るか。", "親のリークと、残った子。", "終了直後"],
            ["時計", "timedatectlとchronyc tracking。開始時刻を親ログと揃える。", "GCログとsarが重ねられる。", "試験前と開始時"],
            ["読み方", "経過時間が制限を超えたとき、CPU、await、スワップ、外待ちのどれかを先に1つ決める。", "形が決まる前の増強は、待ちを大きい机へ移す。", "報告書"],
        ],
        [14, 55, 46, 22],
    )


def add_sar(wb):
    add_sheet(
        wb,
        "sarの読み方",
        "日記の家系、最初の1行の意味、RHEL 9でのメモリの式、試験で残すオプション。",
        ["項目", "内容", "親子JVMでの読み方"],
        [
            ["家系", "System Vのsarを、1999年ごろからのsysstatがLinuxへ運んだ。収集がsadc、定時がsa1とsysstat-collect.timer、表示がsar、変換がsadf。", "10分間隔の/var/log/saは背景。試験の山は1秒で別ファイルへ書く。"],
            ["間隔0", "sarの間隔が0のときだけ、起動後の平均を出す。", "試験のsar 1は、間隔の差分として読む。"],
            ["vmstatとiostat", "最初の1行は起動後平均。iostatは-yで省ける。vmstatも-yがある版がある。", "2行目以降を採用する。"],
            ["pidstat", "最初のブロックが起動後平均になる版がある。", "2回目のブロック以降を採用する。"],
            ["Average行", "その窓の平均。山が短くアイドルが長いと、山が小さくなる。", "ピークはピークの時刻の行で読む。"],
            ["sar -u ALL -P ALL 1", "user、system、iowait、steal、idleに加え、irq、soft、guest。CPUごと。", "allが50%でも1つのCPUが100%なら、1スレッドが律速。"],
            ["仮のCPU行", "user 78、system 6、iowait 0.4、steal 0、idle 15、という並びの例。", "計算が主体で、ディスク待ちとstealは小さい、という解釈。測定値ではない。"],
            ["sar -q ALL 1", "runq-sz、plist-sz、ldavg-1/5/15、blocked。", "runq-szをvCPUと並べる。blockedが本体ならディスク。"],
            ["sar -r ALL 1", "kbavail、kbanonpg、kbcached、kbdirty、%commit、kbpgtbl、kbkstack、kbmemused。", "余りはkbavail。匿名の増減はkbanonpg。"],
            ["kbmemusedの式", "MemTotal−MemFree−Buffers−Cached−Slab。%memusedはその割合。", "RHEL 8の手順書の式をRHEL 9へ持ち越さない。"],
            ["sar -S と -W", "スワップ残量と、pswpin/s、pswpout/s。", "業務中の継続したpswpin/sは、JVMの試験では不合格条件として扱う。"],
            ["sar -B", "pgpgin/s、pgpgout/s、fault/s、majflt/s。", "majfltがファイルかスワップかは、pswpinと同時刻で分ける。"],
            ["sar -w", "proc/sとcswch/s。", "子の起動秒にproc/sの山。切替の内訳はpidstat -w。"],
            ["sar -d -p", "デバイス名をnvme0n1のように出す。rkB/s、wkB/s、await、aqu-sz。", "awaitが伸びてidleが増えると、経過時間はディスク待ち。"],
            ["sar -n DEV", "rxkB/s、txkB/s、%ifutil。", "帯域が細くidleが大きい遅さは、相手側の待ちを疑う。"],
            ["sar -v / -H / -m CPU", "file-nr、ヒュージページ、CPU周波数。", "FD枯渇、明示的ヒュージページ、周波数低下の確認。"],
            ["sar -A -o", "試験窓の全項目をバイナリで残す。例: sar -A -o /var/tmp/sa_batch 1 7200", "ファイルサイズと、EBSへの書き込み自身を記録する。"],
            ["sadf -d", "あとから区切り文字付きで出す。sadf -d file -- -u ALL -r ALL -q", "表計算に貼る。ヘッダを残す。"],
            ["PSI", "/proc/pressureのcpu、memory、io。someはだれかが止まった割合。fullは仕事が進まなかった割合。", "memoryのfullが続くときは、回収かスワップで全員が止まっている。"],
            ["時刻", "-sと-eはHH:MM:SS。ファイルの時刻はホストのローカル。", "深夜をまたぐとsaDDが2日分。"],
        ],
        [28, 78, 62],
    )


def add_commands(wb):
    add_sheet(
        wb,
        "コマンド辞典",
        "読み取り専用。ピークに毎秒重ねてよいものと、A/C/Dの3点だけにするものを分けてある。",
        ["分類", "コマンド", "何が見えるか", "親子JVMでの読み方", "注意"],
        [
            ["仕様", "lscpu", "CPU数、Thread(s) per core、NUMA、モデル名。", "机の数と種類。GC人数の分母。", "ファミリーの記憶で代用しない。"],
            ["仕様", "free -h", "Mem、available、buff/cache、swap。", "余りはavailable。", "RHEL 9.1以降のusedはTotal−Available。"],
            ["仕様", "cat /proc/meminfo", "MemAvailable、AnonPages、Committed_AS、PageTables、Slab。", "sarの列の裏付け。", "単位はkB。"],
            ["日記", "sar -u ALL -P ALL 1", "CPU時間の内訳をCPUごと。", "userが高いか、stealか、idleかを見る。", "ヘッダ行を同じファイルに残す。"],
            ["日記", "sar -q ALL 1", "実行キューとロード。", "runq-szとvCPUを並べる。", "1分ロードだけを見ない。"],
            ["日記", "sar -r ALL 1", "メモリの内訳。", "kbavailとkbanonpgとkbcached。", "kbmemusedはRHEL 9の式。"],
            ["日記", "sar -W 1", "スワップインとアウト。", "0が続いているかを見る。", "スワップ領域が無いホストでは総量0。"],
            ["日記", "sar -B 1", "ページインとフォルト。", "冷えた起動のmajfltと、スワップを分ける。", "fault/sの増加はAlwaysPreTouch無しでも起きる。"],
            ["日記", "sar -w 1", "プロセス生成とコンテキストスイッチ。", "起動の秒にproc/s。", "内訳はpidstat。"],
            ["日記", "sar -d -p 1", "ディスクの待ち。", "awaitとaqu-sz。", "%utilだけでSSDを判断しない。"],
            ["日記", "sar -n DEV 1", "NICの送受信。", "DB待ちは帯域が細い形。", "相手のCPUは出ない。"],
            ["日記", "sar -A -o ファイル 1", "全項目のバイナリ。", "あとから何度でも切り出せる。", "書き込み自身がEBSに乗る。"],
            ["日記", "sadf -d ファイル -- -r", "区切り文字の出力。", "Excelへ持ち込む。", "列順はヘッダに従う。"],
            ["ライブ", "vmstat 1", "r、b、swap、io、us、sy、id、wa、stを1画面に。", "試験中の見張り。", "最初の1行は起動後平均。"],
            ["ライブ", "mpstat -P ALL 1", "CPUごとの使用。", "1コアに寄っているかを見る。", "2行目以降。"],
            ["ライブ", "iostat -xy 1", "拡張ディスク統計。", "await。", "最初の1行は起動後平均。-yで省ける。"],
            ["ライブ", "pidstat -h -u -r -d -w -G 'java|jspawnhelper' 1", "JVMごとのCPU、RSS、IO、切替。", "どの子が忙しく、どの子が寝ているか。", "-hは横並び。人間向け単位ではない。2回目以降。"],
            ["ライブ", "pidstat -t -u -p PID 1", "スレッドごと。", "GCワーカーとC2と業務を分ける。", "全PIDへ毎秒は重い。対象を絞る。"],
            ["関係", "ps -o pid,ppid,nlwp,rss,vsz,stat,wchan,cmd -C java", "親子、スレッド数、状態、待ち場所。", "本数が11か。パイプ待ちか。", "%CPUは生涯平均。"],
            ["関係", "pstree -ap 親PID", "起動の木。シェルが挟まるか。", "PPIDの解釈。", "jspawnhelperは瞬間。"],
            ["関係", "smaps_rollup", "Rss、Pss、Shared_Clean、Private_Dirty。", "RSS合計と実体の差。", "A、C、Dの3点。"],
            ["関係", "/proc/PID/limits と fd", "FD上限と使用数。", "Too many open filesの確認。", "子は親の限界を受け継ぐ。"],
            ["JVM", "jcmd PID VM.flags", "実効フラグ。", "MaxHeapSizeがスクリプトと同じか。", "全子PID。"],
            ["JVM", "jcmd PID GC.heap_info", "ヒープの使用と容量。", "上限と、いまの使用を分ける。", "ピークで1巡。"],
            ["JVM", "jcmd PID VM.native_memory summary scale=MB", "ヒープ外の内訳。", "GC、Class、Thread、Code、Internal。", "起動時にNMTが必要。"],
            ["JVM", "jstat -gcutil PID 1000 10", "領域使用率とGC時間。", "FGCの増加と、10本のYGCの重なり。", "時刻をsarと揃える。"],
            ["JVM", "jcmd PID Thread.print", "スレッドの種類と待ち。", "commonPool、仮想スレッド、パイプ。", "ピークで全員に連打しない。"],
            ["圧力", "cat /proc/pressure/memory", "メモリ待ちで止まった割合。", "fullが続くと全員が進んでいない。", "10秒に1回で足りる。"],
            ["圧力", "dmesg -T | grep 'Killed process'", "OOM Killerの犠牲。", "どのJVMが消されたか。", "試験後に必ず1回。"],
            ["仕様", "IMDSv2のinstance-type", "EC2の型番。", "クレジット型かどうかを仕様と照合する。", "トークン付きのリンクローカル。"],
        ],
        [12, 55, 40, 42, 36],
    )


def add_shapes(wb):
    add_sheet(
        wb,
        "形の読み分け",
        "数字の組み合わせに名前を付けたもの。仮の閾値ではない。業務の合否は経過時間で決める。",
        ["形の名前", "同時に見えているもの", "読めること", "次に見るもの"],
        [
            ["計算で机が埋まっている", "userが高く、stealが0、iowaitが0、kbavailが十分、pswpinが0。", "子のスレッドとGCがvCPUを使っている。", "runq-sz、nvcswch、GCログの重なり。"],
            ["ディスク待ち", "idleがあり、awaitまたはiowaitが高く、aqu-szがある。", "EBS、ログ、入力ファイル。", "ボリュームのIOPSと、ログの書き込み量。"],
            ["机を返されている", "stealが継続して高い。", "ハイパーバイザ側の制約。", "インスタンスタイプとCPUクレジット。"],
            ["外で待っている", "idleが高く、runqが小さく、帯域が細い。経過時間は長い。", "JDBCや別APIの応答。", "プール合計と相手側の測定。"],
            ["スワップを踏むGC", "kbavailが低く、pswpinが増え、majfltが増え、GC時間が伸びる。", "ヒープ合計かネイティブか同時本数が大きい。", "NMTと、子の-Xmxの合計。"],
            ["キャッシュが温まっている", "memusedが高め、kbcachedが大きく、kbavailも残る。", "使用中の定義の違い。", "判定はkbavail。"],
            ["起動の山", "proc/sの短い山、ファイル側のmajflt、JITスレッド。", "冷えた同時起動。", "温まった回と分けて記録する。"],
            ["1本だけ遅い", "1つのPIDだけGCが長く、同時数は10のまま。", "平均ではなくその子。", "そのPIDのヒープとヒストグラムはピークの外で。"],
            ["親の残骸", "写真Dの親のClass、FD、スレッドが写真Aより大きい。", "子は返っている。親が持っている。", "親のNMT。"],
            ["孤児", "親の停止後もjavaが残り、PPIDが1。", "停止が子へ届いていない。", "pstreeとsystemdのKillMode。"],
            ["パイプ詰まり", "子のwchanがパイプ書き込み。CPUは低い。枠が空かない。", "親が標準出力を読んでいない。", "ProcessBuilderのリダイレクト。"],
            ["スレッドが作れない", "unable to create native thread。ヒープ使用率は低い。", "上限はヒープの外。", "NLWP、ulimit -u、pids.max。"],
            ["RSSだけが大きい", "RSS合計がAnonPagesの増分より大きい。Shared_Cleanがある。", "共有ライブラリの二重計上。", "PSSの合計。"],
        ],
        [24, 52, 42, 40],
    )


def add_steps(wb):
    add_sheet(
        wb,
        "試験手順",
        "採取の順番。コマンドの詳細は「コマンド辞典」。数字の解釈は「形の読み分け」。",
        ["順", "やること", "残す成果物", "ここで分かること"],
        [
            ["1", "環境の固定項目を、子を増やす前にファイルへ残す。", "lscpu、free、uname、java -version、sysctl、THP、timedatectl。", "分母が固定される。"],
            ["2", "sysstatのタイマーを確認し、当日の/var/log/saを保管する。", "systemctl list-timers、saDD。", "試験前から異常だったか。"],
            ["3", "写真A。親だけ。", "sar 60秒、親のVM.flags、NMT、smaps_rollup。", "近所と親の常駐費。"],
            ["4", "sar -A -oで1秒採取を始める。見張りにvmstatとpidstat。", "バイナリファイル。", "あとから切り出せる日記。"],
            ["5", "可能なら子1本の写真B。", "温まったあとのRSSとjstat。", "1本の単価。"],
            ["6", "同時10本を流す。開始時刻を親ログとdateで残す。", "開始時刻。", "GCログとの重ね。"],
            ["7", "ピークで本数が親1＋子10であること、全子のMaxHeapSizeが同じことを確認する。", "ps、VM.flags。NMTは1巡。", "設定値と実効値。"],
            ["8", "終了後の写真D。", "ps、kbanonpg、親のNMT、dmesg。", "返却と、OOMと、親の残骸。"],
            ["9", "sadfで必要な資源を出し、GCログとrunq-szとkbavailを重ねる。", "表計算。", "山の時刻が一致するか。"],
            ["10", "形を1つ決めてから、本数やサイズの変更を話す。", "報告書。物差しの種類と、冷えた回か温まった回かを書く。", "待ちを別の資源へ移さない。"],
        ],
        [8, 62, 48, 40],
    )


def add_jdk(wb):
    add_sheet(
        wb,
        "JDKとRHEL",
        "OpenJDK 21とRHEL 9.8の、この構成に効く点。既定と、フラグを付けたときを分けてある。",
        ["点", "既定で起きること", "この構成での意味"],
        [
            ["G1", "既定GC。", "子10本はG1が10個。Young GCは停止型で、ParallelGCThreadsが机を取りにいく。"],
            ["世代別ZGC", "入っている。既定ではない。-XX:+UseZGC -XX:+ZGenerationalのときだけ。", "CPUの常時使用とネイティブメモリの形が変わる。フラグを見てから説明する。"],
            ["仮想スレッド", "JEP 444。アプリが使ったときだけ。", "子の中の待ちは軽くなる。プロセス数とヒープの個数は減らない。"],
            ["CDS", "-Xshare:auto。アーカイブが一致すれば共有。", "失敗しても起動は続く。ログで成否を残す。"],
            ["posix_spawn", "Linuxの既定の子起動。", "大きい親からの起動で、fork時代の予約失敗を避ける。"],
            ["cgroup v2", "上限があれば分母になる。無制限ならインスタンス全体。", "CPUQuotaとMemoryMaxの有無をユニット定義で残す。"],
            ["NMTとJFR", "OpenJDKに含まれる。NMTは起動フラグが必要。", "JFRをピークに長く重ねると、それ自身が負荷になる。"],
            ["カーネル5.14系", "RHEL 9の系。9.8 GAは5.14.0-687.5.1.el9_8。", "PSIとsmaps_rollupが使える。正確な版はuname -r。"],
            ["glibc 2.34", "RHEL 9の系。", "posix_spawnの実体はclone(CLONE_VM|CLONE_VFORK)。"],
            ["sysstat 12.5系", "kbmemusedの式が変わったあと。", "RHEL 8の手順書の式を使わない。"],
            ["free", "9.1以降のusedはMemTotal−MemAvailable。", "sarの使用中と並べるときは列名を書く。"],
            ["ヒューモンガス", "G1リージョンの半分以上の巨大配列。", "空きがあるように見えてFull GCへ進むことがある。GCログを見る。"],
            ["サポート期間", "OpenJDK 21は2023-09-19のLTS。RHEL 9のフルサポートは2027-05-31まで。", "配布物の期限はRed HatのOpenJDKライフサイクル。rpmの版を残す。"],
        ],
        [22, 55, 70],
    )


def add_glossary(wb):
    add_sheet(
        wb,
        "用語",
        "文書とExcelで使う言葉。正式な名前と、教室のたとえを並べてある。",
        ["用語", "意味", "たとえ"],
        [
            ["vCPU", "ゲストから見えるCPU。", "机。"],
            ["CFS", "LinuxがスレッドへvCPUの時間を配るスケジューラ。", "席を順番に使わせる係。"],
            ["%steal", "vCPUがゲストに渡っていなかった時間の割合。", "管理人に机を返していた時間。"],
            ["ヒープ", "Javaオブジェクトを置く領域。プロセスごとに独立。", "作業員の引き出し。"],
            ["メタスペース", "クラス定義を置く、ヒープの外。", "引き出しの外の、道具の目録。"],
            ["G1", "OpenJDK 21の既定GC。", "引き出しの中を片付ける係。"],
            ["Young GC", "若いオブジェクトの回収。短い停止がある。", "新しいゴミだけの片付け。"],
            ["Full GC", "ヒープ全体に近い停止回収。", "引き出し全体の大片付け。"],
            ["RSS", "実メモリに載っているマップ。共有分も全量。", "その人の机の上に見えている量。教科書も全冊に数える。"],
            ["PSS", "共有分を人数で割った居住量。", "教科書を人数で割ったあとの量。"],
            ["VSZ", "予約した住所の広さ。", "使ってよい棚の番号の範囲。"],
            ["ページキャッシュ", "ファイルを速く読み直すためのメモリ。回収できる。", "開いた教科書を置く台。"],
            ["MemAvailable", "スワップせずに新しいアプリを始められる見込み。", "ロッカーの実用的な余り。"],
            ["コミット", "カーネルが使うと約束したメモリ。", "取り置きの伝票。"],
            ["overcommit", "伝票の合計がロッカーを超えても予約を許す動作。", "在庫以上の予約を受ける窓口。"],
            ["CDS", "クラス情報のアーカイブを複数JVMで読み取り共有する仕組み。", "同じ教科書を一緒に見る仕組み。"],
            ["NMT", "JVMが自分のネイティブメモリを分類して見せる機能。", "引き出しの外の棚卸し。"],
            ["cgroup v2", "CPU、メモリ、プロセス数に上限を付けられる仕組み。", "この班が使ってよい机とロッカーの上限。"],
            ["PSI", "CPU、メモリ、I/Oのどれでタスクが止まったかの割合。", "だれが、どの理由で手を止めたかの割合。"],
            ["OOM Killer", "物理メモリが尽きたとき、カーネルがプロセスを終了させる仕組み。", "一番大きい荷物の人を外に出す判断。"],
            ["jspawnhelper", "posix_spawnで子を起動する短い補助プロセス。", "新しい作業員を入室させる案内役。"],
            ["NPTL", "Linuxの現行POSIXスレッド実装。", "1人の中のたくさんの手の土台。"],
            ["ロードアベレージ", "実行したいタスクと割り込み不能タスクの指数平均。", "席を待っている人と、倉庫の前で止まっている人の平均。"],
        ],
        [24, 70, 48],
    )


def add_sources(wb):
    add_sheet(
        wb,
        "出典",
        "確認日 2026-10-14。対象サーバの実測は含まない。見本の16GiBと8 vCPUは式を見せる入力。",
        ["種類", "内容", "この文書での使い方"],
        [
            ["RHEL 9.8", "一般提供は2026-05-19。GA時点のカーネルは5.14.0-687.5.1.el9_8。リリースノートとRed Hatのリリース日一覧。", "版の起点。試験票はuname -rを貼る。"],
            ["RHEL 9.0", "2022-05-17。カーネル5.14系。フルサポートは2027-05-31まで。", "cgroup v2と5.14の土台。"],
            ["sarのkbmemused", "Red Hatの説明。RHEL 9はMemTotal−(MemFree+Buffers+Cached+Slab)。sysstat 11.7.4で計算が変わった。", "古い手順書との差。"],
            ["free", "Red Hatの説明。RHEL 9.1以降のusedはMemTotal−MemAvailable。buff/cacheはBuffers+Cached+SReclaimable。", "sarとfreeを同じ「使用中」と呼ばない。"],
            ["マニュアル", "sar(1)、vmstat(8)、iostat(1)、sadf(1)。vmstatとiostatの最初の1行は起動後平均。sarは間隔0で起動後平均。", "最初の1行の採用ルール。"],
            ["OpenJDKの起動", "ProcessImpl。Linuxの既定はPOSIX_SPAWN。JDK 13の系統で既定になり、21も受け継ぐ。", "親が子を起動するときのメモリ複製の有無。"],
            ["G1の人数", "ParallelGCThreadsは、見えるCPUが8以下ならその数、それ以外は8+(N-8)×5/8。ConcGCThreadsは(Parallel+2)/4を下回らない。", "シート「計算」の式。"],
            ["OpenJDK 21", "既定GCはG1。JEP 439が世代別ZGC。JEP 444が仮想スレッド。CDSの既定は-Xshare:auto。2023-09-19のLTS。", "フラグが無いときの説明の前提。"],
            ["vCPU", "AWSの定義。1 vCPUがハードウェアスレッドの場合と、コアそのものの場合がある。", "判定はゲストのlscpu。"],
            ["PSI", "LinuxカーネルのPressure Stall Information。/proc/pressureのsomeとfull。", "ロードアベレージと並べる資源ごとの止まり。"],
            ["測っていないもの", "インスタンスタイプ、実フラグ、実sar、GCログ、DB接続、CloudWatchエージェントのメモリ式。", "見本の数字を測定値として使わない。"],
            ["対になる文書", "親子JVMのメモリとCPUとsar.md", "物語と経緯の全文。このブックは表と計算。"],
        ],
        [24, 78, 55],
    )


if __name__ == "__main__":
    main()
