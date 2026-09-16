# -*- coding: utf-8 -*-
"""JWC2026 兵庫県予選 トーナメント表（公式最新版 2026-09-14）からの抽出データ.

出典:
  男子: JWC2026_男子トーナメント表_公式最新版_20260914_1629.pdf (Drive 1cTq7CaAVouMxUXvT-Esb49kAkiYV4NVn)
  女子: JWC2026_女子トーナメント表_公式最新版_20260914_1631.pdf (Drive 1eI-P8LlI2zj0LhQH9J2Do7wXcLBix7gA)
  日程・会場: 2026年度_JWC2026_大会実施要項_正式版_2026-08-18.pdf (Drive 1gUwmFwPHYeGVErd4_Oi455ky5rPCP8Hj) 第6項
"""

TOURNAMENT = {
    "大会名称": "Jr.ウインターカップ2026-2027 2026年度第7回全国U15バスケットボール選手権 兵庫県予選大会（代表決定戦）",
    "主催": "公益財団法人日本バスケットボール協会／一般財団法人兵庫県バスケットボール協会／姫路バスケットボール協会",
    "共催": "相生市バスケットボール協会",
    "主管": "一般財団法人兵庫県バスケットボール協会U15部会",
    "大会方式": "トーナメント戦によるノックアウト方式",
    "抽選会": "2026年9月11日(金) 主催者による責任抽選（アプリ抽選）",
}

# 会場コード（トーナメント表の凡例）→ 正式名称
VENUES = [
    ("総", "Life Partner Arena（兵庫県立総合体育館）", "総A・総B・総C・総D"),
    ("相", "相生市民体育館", "相A・相B"),
    ("北", "北神戸田園スポーツ公園体育館", "北A・北B"),
    ("グ", "グリーンアリーナ神戸", "グA・グB・グC・グD"),
    ("姫", "姫路市立総合スポーツ会館", "姫A・姫B・姫C"),
    ("大", "大和工業アリーナ姫路（ひめじスーパーアリーナ）", "大A・大B・大C・大D"),
]

# 開催日コード（トーナメント表の先頭数字）→ 日付・曜日・会場
DATES = [
    ("17", "2026-10-17", "土", ["総", "相"], "1回戦ほか"),
    ("18", "2026-10-18", "日", ["北", "相"], "1回戦ほか"),
    ("24", "2026-10-24", "土", ["グ"], "女子 2回戦ほか"),
    ("25", "2026-10-25", "日", ["グ"], "男子 2回戦ほか"),
    ("1",  "2026-11-01", "日", ["姫"], "3回戦ほか（男子）"),
    ("3",  "2026-11-03", "火・祝", ["大"], "3回戦ほか"),
    ("7",  "2026-11-07", "土", ["相"], "準々決勝ほか"),
    ("14", "2026-11-14", "土", ["総"], "準決勝ほか"),
    ("15", "2026-11-15", "日", ["総"], "決勝・3位決定戦ほか"),
]

# トーナメント表に記載されている試合コード（日付コード＋会場コード＋コート＋開始時刻）
CODES_M = """15総A11:00 14総B15:00 14総D15:00 17総A14:30 25グA9:30 25グA12:30 17総C16:00
18相A15:00 18北B14:00 17総B14:30 18相B15:00 7相A11:30 15総C14:00 1姫A10:00 1姫A13:00
15総C11:00 3大A13:30 18相A9:00 18北A11:00 18相B12:00 18相A13:30 18相B13:30 18北A12:30
18北B11:00 18相B9:00 25グB9:30 25グD14:00 25グC14:00 18北A14:00 18相A10:30 14総A12:00
14総C12:00 1姫B10:00 1姫C10:00 1姫A11:30 1姫B11:30 25グD15:30 25グB12:30 17相A13:30
17相B13:30 17相B15:00 17相A15:00 1姫B13:00 18相A12:00 1姫C13:00 1姫A14:30 25グA14:00
1姫C14:30 3大B13:30 17相A16:30 17総C14:30 25グC9:30 3大B15:00 18相B10:30 17総D14:30
25グD9:30 25グD11:00 25グA11:00 17総B16:00 15総A14:00 14総A15:00 14総C15:00 18北A9:30
18北B12:30 3大A15:00 3大C13:30 18北B9:30 25グB11:00 14総B12:00 7相A13:00 3大C15:00
18北B15:30 17総A16:00 17相B16:30 18北A15:30 14総D12:00 25グB14:00 25グD12:30 25グC12:30"""

CODES_W = """15総A9:30 24グA11:00 24グA14:00 14総B13:30 14総D13:30 17総A10:00 17相A12:00
15総C12:30 15総C9:30 3大A10:30 3大A12:00 17総B10:00 17相B12:00 24グB14:00 24グB11:00
17総C11:30 14総A10:30 14総C10:30 1姫C11:30 17相B10:30 24グC12:30 17総D10:00 3大B10:30
3大C12:00 17総C10:00 17総A11:30 24グD14:00 24グD12:30 15総A12:30 14総A13:30 14総C13:30
24グD11:00 24グC14:00 24グA15:30 17総D11:30 17相A9:00 3大B12:00 17総A13:00 24グC11:00
17相B9:00 24グB15:30 14総D10:30 14総B10:30 17総B13:00 24グC15:30 24グB12:30 17相A10:30
17総C13:00 7相A10:00 3大C10:30 17総B11:30 17総D13:00 24グD15:30 24グA12:30"""

# 組合せ番号 → チーム名（トーナメント表の記載どおり）
BRACKET_M = [
    "センターサークル", "MIRRORS", "VLakers Basketball Club U14 男子", "SAKURA PRESS 男子",
    "COULEUR", "ZENITH", "甲南中学校", "B-LION",
    "Falcons U15 basketball club team", "ARMS", "G.Dark Horse", "BAYCROWN JUNIOR",
    "BAMBOO Shoot U15", "神戸HOPES", "WEST RIVER", "YFT PRESTO",
    "B.P.F ACADEMY", "Dpro Laluz", "神戸ストークスU15", "BRAVE BIRDS",
    "ZERO", "ICE", "EPIC BASKETBALL CLUB U14", "中崎B.B.C. U15",
    "HDC Academy Cranes", "えびす", "尼崎市立大庄中学校", "えびすTHIRD",
    "U15 ドリームメーカーズ", "神戸ブライアンツ", "All Blacks", "U14 ゴッドドア",
    "Turkeys 2nd", "VLakers Basketball Club U15 男子", "V-WAVE", "Westrick",
    "VEARTH", "HYOGO KOREA", "報徳学園中学校", "NorthWave",
    "池田バスケットボールクラブ", "BRAVEBIRDS U14", "DIVE basketball academy", "DRAGON basketball",
    "WINGS", "明石市立野々池中学校", "KARTER", "Grow wings",
    "EPIC BASKETBALL CLUB U15", "西神戸ユニバース", "BCJ Academy", "DunkGo Club",
    "UNICORN BASKETBALL CLUB", "BRUINS ashiya", "Improve Basketball Club", "SKYBLUEZ BASKETBALL CLUB",
    "神戸ストークスU14", "Three B", "ゴッド ドア", "Three B U14",
    "えびすSECOND", "SAMURAI", "FightingArts", "TURKEYS",
    "RIVER WEST", "Kobe Center Circle", "RED PIECE", "GOAT",
    "cantera karter", "Kobe Center Circle 2", "Wild Wolves", "team DENY",
    "尼崎市立塚口中学校", "BELUGA", "A.C.T. BASKETBALL CLUB 2026", "DREAMSEEKER",
    "明石市立望海中学校",
]

BRACKET_W = [
    "All BLUE", "BLACK PANTHERS_U14", "BRUINS kobe U15", "SAKURA PRESS 女子",
    "TURKEYS", "PENGUINS U15", "Dpro Leonora", "ゴッド ドア",
    "VODKA Basketball Club", "池田バスケットボールクラブ", "WEST RIVER", "ROSELLE",
    "GLANZ", "EPIC BASKETBALL CLUB U15", "ALL BLUE GAO", "VOLTAGE",
    "COULEUR", "KARTER", "WINGS", "HDC Academy Cranes",
    "KOBE RISING PHOENIX BASKETBALL SCHOOL", "VLakers Basketball Club U15 女子", "CLUTCH", "BAMBOO Shoot U15",
    "百合学院中学校", "NorthWave", "明石市立二見中学校", "DIVE basketball academy",
    "DRAGON basketball", "HyogoForce U13", "Wild wolves", "HyogoForce",
    "西神戸ユニバース", "南あわじ市立三原中学校", "A.C.T. BASKETBALL CLUB 2026", "MONKEYS basketball club",
    "BLACK PANTHERS", "BENCHERS", "EPIC BASKETBALL CLUB U14", "CORES",
    "BE BEST", "B.Victoire", "DREAMSEEKER", "V-WAVE",
    "DunkGo Club", "RC", "ICE", "VLakers Basketball Club U14 女子",
    "センターサークル", "高砂市立鹿島中学校",
]
