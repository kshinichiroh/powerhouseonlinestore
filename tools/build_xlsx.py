# -*- coding: utf-8 -*-
import csv, re, sys, unicodedata
from collections import OrderedDict, defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
from bracket_data import TOURNAMENT, VENUES, DATES, CODES_M, CODES_W, BRACKET_M, BRACKET_W

STAMP = "2026-09-15"
H_FILL = PatternFill("solid", fgColor="1F3864")
H_FONT = Font(color="FFFFFF", bold=True, size=11)
T_FONT = Font(bold=True, size=14)
SUB = Font(italic=True, size=9, color="555555")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def norm(s):
    s = unicodedata.normalize('NFKC', s)
    return s.replace(' ', '').replace('　', '').replace('_', '').lower().replace('2026', '')


def header(ws, row, cols):
    for i, c in enumerate(cols, 1):
        cell = ws.cell(row=row, column=i, value=c)
        cell.fill, cell.font, cell.border = H_FILL, H_FONT, BOX
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = ws.cell(row=row + 1, column=1)


def widths(ws, ws_widths):
    for i, w in enumerate(ws_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def title(ws, text, note=None):
    ws["A1"] = text
    ws["A1"].font = T_FONT
    if note:
        ws["A2"] = note
        ws["A2"].font = SUB


# ---------------------------------------------------------------- source data
def find(name):
    """members_master.csv はリポジトリに置かないため、カレント優先で探す。"""
    for d in ('.', os.path.join(ROOT, 'data'), HERE):
        p = os.path.join(d, name)
        if os.path.exists(p):
            return p
    raise SystemExit(
        f'{name} が見つかりません。data/README.md の手順で配置してください。')


members = list(csv.DictReader(open(find('members_master.csv'), encoding='utf-8')))
material = list(csv.DictReader(open(find('material_status.csv'), encoding='utf-8')))
# 素材台帳のチーム名がトーナメント表／JBA登録名と異なるもの（同一チーム）
ALIAS = {
    ('男子', 'U15 ドリームメーカーズ'): "U15Dream MakeR's",
    ('女子', 'PENGUINS U15'): 'PENGUINS',
}
MAT = {}
for d in material:
    MAT[(d['男女'], norm(d['チーム名']))] = d
for (g, real), alias in ALIAS.items():
    if (g, norm(alias)) in MAT:
        MAT[(g, norm(real))] = MAT[(g, norm(alias))]

WITHDRAWN = {
    ('男子', 'BAMBOO Shoot U15'): '棄権（メンバー表・素材とも掲載対象外）',
    ('男子', 'A.C.T. BASKETBALL CLUB 2026'): '棄権（男子のみ。女子は掲載対象）',
}

PRINT_NAME = {}   # 誌面に印字する名称（JBA登録名＝メンバー表の表記）
for r in members:
    PRINT_NAME[(r['男女'], norm(r['チーム名']))] = r['チーム名']

STAFF_FIELDS = [
    "ヘッドコーチ（HC）",
    "アシスタントコーチ（AC）①",
    "アシスタントコーチ（AC）②",
    "チーム責任者",
    "トレーナー",
    "チームスタッフ",
    "帯同審判①",
    "帯同審判②",
    "帯同MC",
]

STAFF_NOTE = "※スタッフは大会エントリーシートから後日差し込みます（2026-09-15時点で未収録）。"

RULES = [
    ("1", "チーム掲載名",
     "誌面に印字するチーム名はJBA帳票（2026年度チームメンバー一覧表）の登録名を正とする。"
     "フォルダ名・ファイル名の表記が異なる場合もJBA登録名を優先する。"),
    ("1-2", "チーム掲載名の例外（トーナメント表のみ）",
     "組合せページ（トーナメント表）だけは、既に公表済みのトーナメント表PDFの表記のまま印字する。"
     "該当は3か所のみ。男子59「ゴッド ドア」／女子8「ゴッド ドア」／女子32「HyogoForce」。"
     "理由は2つ。①トーナメント表は既にウェブ等で一般に公表済みであり、いま表記を変えると"
     "公表済みのものと誌面が食い違って混乱を招くため。②差はスペースの有無だけで、"
     "別チームと誤認される恐れのない軽微なものであるため。"
     "その結果、1冊の中で組合せページとチーム紹介ページの表記が2通りになるが、これは意図的である。"
     "統一しないでください。この例外はトーナメント表に限る。"
     "チーム紹介ページ・メンバー表・その他すべての資料ではJBA登録名を用いる（ルール1のとおり）。"
     "新たに表記の違いが見つかった場合はこの例外の対象ではなく、JBA登録名に直す。"),
    ("2", "チーム紹介文",
     "チーム紹介ページに紹介文は入れない。掲載はチーム名・集合写真（またはロゴ）・選手名簿のみ。"
     "広告枠内の文言はその広告デザインの一部としてそのまま掲載する。"),
    ("3", "集合写真がないチーム",
     "集合写真の提出は任意。写真がないチームは未提出ではなく「写真なしで確定」。"
     "ロゴがある場合はロゴを、ロゴもない場合はチーム名と名簿のみで組む。追加提出は発生しない。"),
    ("4", "スタッフ情報",
     "チーム紹介ページにはスタッフを掲載する（2025年度の誌面と同じ）。掲載項目は "
     "ヘッドコーチ／アシスタントコーチ／チーム責任者／トレーナー／チームスタッフ／"
     "帯同審判①②／帯同MC。出所は大会エントリーシートで、本ファイルには未収録のため後日差し込む。"
     "電話番号・メールアドレス・住所・JBAメンバーIDは昨年度も誌面になく、入稿データに含めない。"),
    ("5", "選手名簿の項目と順序",
     "No.／名前／学年／身長／学校名 の5項目・この順序で統一（2025年度パンフレットに準拠）。"
     "背番号の飛び番は原票どおり保持する。小学生の学年は「小5」「小6」と原票表記のまま。"),
    ("6", "素材の受付終了",
     "チーム素材は2026-09-11の第19便をもって受付終了。以降の追加・差し替えは受け付けない。"
     "旧版が過去便のフォルダに残っていても、差替指示のあるものは新版のみを使用する。"),
    ("7", "広告の重複配置",
     "複数チーム共通の広告は1点のみ配置し、チームごとに重複配置しない。"
     "共通広告に含まれるチーム数は台帳「チーム別照合」タブに記載。"),
    ("8", "データ未着の問い合わせ",
     "未着と思われる素材は、まず台帳「JWC2026_入稿一覧・差替履歴」の"
     "「チーム別照合」タブと本ファイルの素材ステータスを確認する。"
     "「広告不要」「写真なし」「チーム名のみ」は確定状態であり、未提出ではない。"),
    ("9", "似た名称のチーム",
     "WEST RIVER（男子・女子）と RIVER WEST（男子）は別チーム。"
     "同様に Three B / Three B U14、えびす / えびすSECOND / えびすTHIRD、"
     "Kobe Center Circle / Kobe Center Circle 2、EPIC BASKETBALL CLUB U15 / U14 などはすべて別チーム。"),
    ("10", "トーナメント表の開催日・会場",
     "トーナメント表の試合コードは「日付＋会場記号＋コート＋開始時刻」。"
     "例：18相A15:00＝10月18日(日)・相生市民体育館A面・15:00開始。凡例は本ファイルの会場コード表を参照。"),
]


def rules_sheet(wb, name="99_制作ルール"):
    ws = wb.create_sheet(name)
    title(ws, "JWC2026 パンフレット制作ルール（兵庫県バスケットボール協会U15部会 広報部 ⇔ ダイコロ株式会社）",
          f"確認日 {STAMP}／このシートの内容は毎回同じ運用ルールです。判断に迷う点はまずここをご確認ください。")
    header(ws, 4, ["No.", "項目", "ルール"])
    r = 5
    for no, item, body in RULES:
        ws.cell(row=r, column=1, value=no).border = BOX
        c = ws.cell(row=r, column=2, value=item); c.border = BOX; c.font = Font(bold=True)
        c.alignment = Alignment(vertical="top")
        c = ws.cell(row=r, column=3, value=body); c.border = BOX
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 42
        r += 1
    widths(ws, [6, 22, 104])
    return ws


# ============================================================ WORKBOOK 1: 大会・トーナメント
def build_tournament():
    wb = Workbook(); wb.remove(wb.active)

    ws = wb.create_sheet("00_はじめに")
    title(ws, "JWC2026 兵庫県予選大会 トーナメント表 データ版")
    lines = [
        ("", ""),
        ("作成", "一般財団法人兵庫県バスケットボール協会 U15部会 広報部"),
        ("作成日", STAMP),
        ("目的", "トーナメント表PDFの内容を、組版用にテキストデータ化したものです。"),
        ("", ""),
        ("収録内容", ""),
        ("  01_大会概要", "大会名称・主催・大会方式・抽選会"),
        ("  02_日程・会場", "全9日程の開催日・曜日・会場・使用コート（大会実施要項 第6項）"),
        ("  03_会場コード凡例", "トーナメント表の会場記号（総・相・北・グ・姫・大）の正式名称"),
        ("  04_試合コード一覧", "トーナメント表に記載された試合コードを日付・会場・コート・開始時刻に分解した一覧"),
        ("  05_男子_組合せ", "組合せ番号1〜77とチーム名"),
        ("  06_女子_組合せ", "組合せ番号1〜50とチーム名"),
        ("  07_表記ゆれ", "組合せページとチーム紹介ページで表記が異なる2チーム（意図的・統一不要）"),
        ("  99_制作ルール", "パンフレット制作の共通ルール"),
        ("", ""),
        ("ご質問への回答", "前回同様、今回のトーナメント表にも開催日・会場・開始時刻が入っています。"),
        ("", "トーナメント表の各試合に付いている「18相A15:00」のような記号が、"),
        ("", "「10月18日(日)・相生市民体育館A面・15:00開始」を表します。03・04のシートをご覧ください。"),
        ("", ""),
        ("正となるデータ", "トーナメント表の線・組合せの形は、共有フォルダ「04_組み合わせ等_9月11日-12日以降」の"),
        ("", "公式最新版PDF（男女各1点）が正です。本ファイルはその文字情報をデータ化したものです。"),
    ]
    r = 3
    for a, b in lines:
        ws.cell(row=r, column=1, value=a).font = Font(bold=True) if a and not a.startswith("  ") else Font()
        ws.cell(row=r, column=2, value=b)
        r += 1
    widths(ws, [22, 108])

    ws = wb.create_sheet("01_大会概要")
    title(ws, "大会概要", "出典：2026年度_JWC2026_大会実施要項_正式版_2026-08-18")
    header(ws, 4, ["項目", "内容"])
    r = 5
    for k, v in TOURNAMENT.items():
        ws.cell(row=r, column=1, value=k).border = BOX
        c = ws.cell(row=r, column=2, value=v); c.border = BOX
        c.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1
    ws.cell(row=r, column=1, value="出場チーム数").border = BOX
    ws.cell(row=r, column=2, value="127チーム（男子77・女子50）※うち男子2チーム棄権").border = BOX
    r += 1
    ws.cell(row=r, column=1, value="チーム紹介掲載数").border = BOX
    ws.cell(row=r, column=2, value="125チーム（男子75・女子50）／選手1,781名").border = BOX
    widths(ws, [18, 100])

    ws = wb.create_sheet("02_日程・会場")
    title(ws, "開催日・会場一覧", "出典：大会実施要項 第6項／トーナメント表の日付コードと対応")
    header(ws, 4, ["日付コード", "開催日", "曜日", "会場", "使用コート", "主なラウンド"])
    r = 5
    vmap = {v[0]: v for v in VENUES}
    for code, date, dow, vs, rounds in DATES:
        for v in vs:
            ws.cell(row=r, column=1, value=code).border = BOX
            ws.cell(row=r, column=2, value=date).border = BOX
            ws.cell(row=r, column=3, value=dow).border = BOX
            ws.cell(row=r, column=4, value=vmap[v][1]).border = BOX
            ws.cell(row=r, column=5, value=vmap[v][2]).border = BOX
            ws.cell(row=r, column=6, value=rounds).border = BOX
            r += 1
    widths(ws, [12, 14, 8, 44, 24, 24])

    ws = wb.create_sheet("03_会場コード凡例")
    title(ws, "会場コード凡例", "トーナメント表の試合コードは「日付＋会場記号＋コート＋開始時刻」の順です。")
    header(ws, 4, ["会場記号", "会場正式名称", "コート記号"])
    r = 5
    for code, name, courts in VENUES:
        ws.cell(row=r, column=1, value=code).border = BOX
        ws.cell(row=r, column=2, value=name).border = BOX
        ws.cell(row=r, column=3, value=courts).border = BOX
        r += 1
    r += 2
    ws.cell(row=r, column=1, value="読み方の例").font = Font(bold=True)
    for ex in ["18相A15:00 ＝ 10月18日(日)・相生市民体育館・A面・15:00開始",
               "25グC12:30 ＝ 10月25日(日)・グリーンアリーナ神戸・C面・12:30開始",
               "15総A11:00 ＝ 11月15日(日)・Life Partner Arena（兵庫県立総合体育館）・A面・11:00開始"]:
        r += 1
        ws.cell(row=r, column=1, value=ex)
    widths(ws, [12, 48, 26])

    ws = wb.create_sheet("04_試合コード一覧")
    title(ws, "試合コード一覧（トーナメント表PDF記載分）",
          "PDFに印字されている試合コードを分解した一覧です。試合の組合せ（線）は公式PDFが正です。")
    header(ws, 4, ["男女", "試合コード", "開催日", "曜日", "会場", "コート", "開始時刻"])
    dmap = {d[0]: d for d in DATES}
    vmap = {v[0]: v for v in VENUES}
    rowsx = []
    for gender, codes in (("男子", CODES_M), ("女子", CODES_W)):
        for tok in codes.split():
            m = re.match(r'(\d+)(総|相|北|グ|姫|大)([A-D])(\d{1,2}):(\d{2})', tok)
            d, v, ct, hh, mm = m.groups()
            rowsx.append([gender, tok, dmap[d][1], dmap[d][2], vmap[v][1], ct + "面",
                          f"{int(hh):02d}:{mm}"])
    rowsx.sort(key=lambda x: (x[2], x[4], x[5], x[6], x[0]))
    r = 5
    for x in rowsx:
        for i, v in enumerate(x, 1):
            ws.cell(row=r, column=i, value=v).border = BOX
        r += 1
    widths(ws, [8, 16, 14, 8, 44, 10, 12])

    for gender, bracket, sheet in (("男子", BRACKET_M, "05_男子_組合せ"),
                                   ("女子", BRACKET_W, "06_女子_組合せ")):
        ws = wb.create_sheet(sheet)
        title(ws, f"{gender} 組合せ番号とチーム名（全{len(bracket)}チーム）",
              "2つの列は、どちらが正しいかではなく、どのページで使うかの違いです。"
              "組合せページ（トーナメント表）はB列のまま印字してください。"
              "チーム紹介ページ・メンバー表はC列（JBA登録名）で印字してください。"
              "B列とC列が異なる3か所は意図的なものです。統一しないでください（2026-09-15 主催者決定）。")
        header(ws, 4, ["組合せ番号", "組合せページに印字する表記（トーナメント表PDFのまま）",
                       "チーム紹介ページに印字する表記（JBA登録名）",
                       "集合写真", "ロゴ", "広告", "備考"])
        r = 5
        for i, b in enumerate(bracket, 1):
            key = (gender, norm(b))
            printed = PRINT_NAME.get(key, b)
            mat = MAT.get(key)
            wd = WITHDRAWN.get((gender, b))
            if wd:
                photo = logo = ad = "―"
            elif mat:
                photo = mat['写真']
                logo = mat['ロゴ']
                ad = mat['広告区分'] or "―"
            else:
                photo = logo = "なし"
                ad = "広告なし"
            note = wd or ("" if mat or wd else "素材の提出なし。チーム名と名簿のみ掲載。")
            vals = [i, b, printed if not wd else b, photo, logo, ad, note]
            for j, v in enumerate(vals, 1):
                c = ws.cell(row=r, column=j, value=v); c.border = BOX
                if j == 7:
                    c.alignment = Alignment(wrap_text=True, vertical="top")
            r += 1
        widths(ws, [11, 40, 40, 10, 10, 16, 40])

    ws = wb.create_sheet("07_表記ゆれ")
    title(ws, "トーナメント表PDFとJBA登録名が異なる箇所（統一しないでください）",
          "組合せページはトーナメント表PDFの表記のまま、チーム紹介ページはJBA登録名で印字してください。"
          "組合せは既に公表済みのため、意図的に両方を残します（2026-09-15 主催者決定）。")
    header(ws, 4, ["男女", "組合せ番号", "組合せページの表記（トーナメント表PDF）",
                   "チーム紹介ページの表記（JBA登録名）", "備考"])
    r = 5
    diffs = []
    for gender, bracket in (("男子", BRACKET_M), ("女子", BRACKET_W)):
        for i, b in enumerate(bracket, 1):
            p = PRINT_NAME.get((gender, norm(b)))
            if p and p != b:
                diffs.append([gender, i, b, p, "同一チーム。スペースの有無のみの違い。組合せページは公表済みのため変更せず、この表記のまま印字してください。統一は不要です。"])
    for gender, b, memo in (("男子", "BAMBOO Shoot U15", "男子は棄権。女子 BAMBOO Shoot U15 は掲載対象。"),
                            ("男子", "A.C.T. BASKETBALL CLUB 2026", "男子は棄権。女子 A.C.T. BASKETBALL CLUB 2026 は掲載対象。")):
        idx = BRACKET_M.index(b) + 1
        diffs.append([gender, idx, b, "（掲載対象外）", memo])
    for x in diffs:
        for i, v in enumerate(x, 1):
            c = ws.cell(row=r, column=i, value=v); c.border = BOX
            if i == 5:
                c.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1
    r += 2
    ws.cell(row=r, column=1, value="別チームのためご注意ください").font = Font(bold=True, color="C00000")
    for t in ["WEST RIVER（男子・女子）と RIVER WEST（男子）は別チームです。素材もそれぞれ別に提出されています。",
              "Three B と Three B U14、えびす／えびすSECOND／えびすTHIRD、",
              "Kobe Center Circle と Kobe Center Circle 2、EPIC BASKETBALL CLUB U15 と U14、",
              "神戸ストークスU15 と 神戸ストークスU14、TURKEYS と Turkeys 2nd もそれぞれ別チームです。"]:
        r += 1
        ws.cell(row=r, column=1, value=t)
    widths(ws, [8, 11, 40, 40, 52])

    rules_sheet(wb)
    out = f"JWC2026_トーナメント表_データ版_{STAMP.replace('-','')}.xlsx"
    wb.save(out)
    return out


# ============================================================ WORKBOOK 2: メンバー表チーム別
def safe_sheet(name, used):
    s = re.sub(r'[:\\/?*\[\]]', '-', name)
    s = s[:31].rstrip()
    base, n = s, 2
    while s in used:
        suf = f"~{n}"
        s = base[:31 - len(suf)] + suf
        n += 1
    used.add(s)
    return s


def build_members():
    wb = Workbook(); wb.remove(wb.active)
    byteam = OrderedDict()
    for r in members:
        byteam.setdefault((int(r['掲載順']), r['男女'], int(r['整理番号']), r['チーム名']), []).append(r)

    idx = wb.create_sheet("00_目次")
    title(idx, "JWC2026 チーム紹介メンバー表（チーム別シート版）",
          f"作成日 {STAMP}／参加対象125チーム・選手1,781名／チーム名・氏名はJBAダウンロード正本と全件照合済み")
    idx["A3"] = ("【ご注意】本ファイルには選手のみを収録しています。"
                 "スタッフ（HC／AC／チーム責任者／トレーナー／チームスタッフ／帯同審判①②／帯同MC）は"
                 "大会エントリーシートから後日お渡しします。各チームのシートにスタッフ欄を空欄で用意しています。")
    idx["A3"].font = Font(bold=True, color="C00000", size=10)
    header(idx, 4, ["掲載順", "男女", "整理番号", "チーム名", "選手数", "シート名",
                    "集合写真", "ロゴ", "広告", "備考"])

    all_ws = wb.create_sheet("01_全体データ")
    title(all_ws, "全体データ（1行1名）", "チーム別シートと同じ内容を1枚にまとめたものです。検索用にお使いください。")
    header(all_ws, 4, ["掲載順", "男女", "整理番号", "チーム名", "No.", "名前", "学年", "身長", "学校名"])
    ar = 5
    for r in members:
        for i, k in enumerate(['掲載順', '男女', '整理番号', 'チーム名', 'No.', '名前', '学年', '身長', '学校名'], 1):
            c = all_ws.cell(row=ar, column=i, value=r[k]); c.border = BOX
        ar += 1
    widths(all_ws, [8, 7, 9, 38, 7, 18, 7, 8, 30])
    all_ws.auto_filter.ref = f"A4:I{ar-1}"

    used = set(["00_目次", "01_全体データ"])
    ir = 5
    for (order, gender, seq, team), plist in byteam.items():
        pre = "M" if gender == "男子" else "W"
        sheet_name = safe_sheet(f"{pre}{seq:02d}_{team}", used)
        ws = wb.create_sheet(sheet_name)
        ws["A1"] = team
        ws["A1"].font = Font(bold=True, size=16)
        ws["A2"] = f"{gender}／整理番号 {seq}／掲載順 {order}／選手 {len(plist)}名"
        ws["A2"].font = SUB
        mat = MAT.get((gender, norm(team)))
        if mat:
            ws["A3"] = f"素材：集合写真 {mat['写真']}／ロゴ {mat['ロゴ']}／広告 {mat['広告区分'] or '―'}"
        else:
            ws["A3"] = "素材：集合写真 なし／ロゴ なし／広告 なし（チーム名と名簿のみ掲載）"
        ws["A3"].font = SUB

        # スタッフ欄（昨年度の誌面と同じ項目。値は後日エントリーシートから差し込む）
        ws["A5"] = "チームスタッフ"
        ws["A5"].font = Font(bold=True, size=12)
        ws["C5"] = STAFF_NOTE
        ws["C5"].font = Font(size=9, color="C00000")
        header(ws, 6, ["区分", "氏名"])
        rr = 7
        for f in STAFF_FIELDS:
            c = ws.cell(row=rr, column=1, value=f); c.border = BOX
            c = ws.cell(row=rr, column=2, value=None); c.border = BOX
            rr += 1

        rr += 1
        ws.cell(row=rr, column=1, value="選手名簿").font = Font(bold=True, size=12)
        rr += 1
        header(ws, rr, ["No.", "名前", "学年", "身長", "学校名"])
        rr += 1
        for p in plist:
            for i, k in enumerate(['No.', '名前', '学年', '身長', '学校名'], 1):
                c = ws.cell(row=rr, column=i, value=p[k]); c.border = BOX
                if i in (1, 3, 4):
                    c.alignment = Alignment(horizontal="center")
            rr += 1
        widths(ws, [22, 22, 8, 9, 34])

        if mat:
            photo, logo, ad = mat['写真'], mat['ロゴ'], (mat['広告区分'] or "―")
            note = ""
        else:
            photo = logo = "なし"; ad = "なし"
            note = "素材の提出なし。チーム名と名簿のみ掲載。"
        vals = [order, gender, seq, team, len(plist), sheet_name, photo, logo, ad, note]
        for i, v in enumerate(vals, 1):
            c = idx.cell(row=ir, column=i, value=v); c.border = BOX
            if i == 10:
                c.alignment = Alignment(wrap_text=True, vertical="top")
        ir += 1

    widths(idx, [8, 7, 9, 38, 8, 34, 10, 9, 15, 34])
    idx.auto_filter.ref = f"A4:J{ir-1}"

    ws = wb.create_sheet("98_掲載対象外")
    title(ws, "棄権・掲載対象外のチーム", "入稿データには含めていません。誌面にも掲載しないでください。")
    header(ws, 4, ["男女", "整理番号", "チーム名", "状態", "取扱い"])
    excl = [("男子", "62", "BAMBOO Shoot U15", "棄権", "入稿対象外（女子 BAMBOO Shoot U15 は掲載対象）"),
            ("男子", "68", "A.C.T. BASKETBALL CLUB 2026", "棄権", "入稿対象外（女子 A.C.T. BASKETBALL CLUB 2026 は掲載対象）"),
            ("女子", "14", "佐用オレンジスターズ", "対象外", "参加対象から除外"),
            ("女子", "50", "BELUGA", "棄権", "入稿対象外（男子 BELUGA は掲載対象）"),
            ("女子", "53", "KC", "対象外", "参加対象から除外"),
            ("女子", "―", "ZERO", "申込取消", "入稿対象外（男子 ZERO は掲載対象）"),
            ("男女", "―", "御影ストーンズ", "参加不可", "2026-09-05に事前申込遅延で参加不可が確定。掲載しない。")]
    r = 5
    for x in excl:
        for i, v in enumerate(x, 1):
            c = ws.cell(row=r, column=i, value=v); c.border = BOX
            if i == 5:
                c.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1
    widths(ws, [8, 10, 34, 12, 60])

    rules_sheet(wb)
    out = f"JWC2026_チーム紹介メンバー表_チーム別シート版_{STAMP.replace('-','')}.xlsx"
    wb.save(out)
    return out


print(build_tournament())
print(build_members())
