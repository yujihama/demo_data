# -*- coding: utf-8 -*-
"""
連結決算ダミーデータの単一情報源（Single Source of Truth）。

デモA株式会社グループ FY2025（2025/4/1〜2026/3/31）の連結決算業務データを
生成するための会社別財務データ・グループ内取引・連結消去仕訳を定義し、
貸借一致等の検算を行う。

単位：千円（Demo-A (Thailand) は 千THB）。
既存エビデンス（4.evidence/FCRP/ConsolidationSystem_*.csv 等）と整合するよう
以下をアンカーとして設計している。
  - 会社コード: DA-TB / DA-LOG / DAT / DATR（S05アップロードログと同一）
  - 期末 内部債権債務消去額 324,500千円
  - 内部取引高消去 年間 9,880,000千円（四半期 1,820,000+650,000 ×4）
  - 未実現利益消去（棚卸資産）42,800千円
  - 非支配株主損益 年間 34,080千円（四半期 8,520千円）
  - 資本金消去 300,000千円
  - Q1タイバーツ換算レート 4.05円/THB（バリデーション往復ログ）
"""

# ---------------------------------------------------------------------------
# 為替レート（円/THB）
# ---------------------------------------------------------------------------
FX = {
    "CR_open": 4.00,   # 2025/3/31 期末日レート（前期末）
    "CR_close": 4.20,  # 2026/3/31 期末日レート
    "AR": 4.10,        # FY2025 期中平均レート
    "HR_capital": 3.00,  # 資本金取得時レート（2012年設立時）
    "quarterly": [
        # (四半期, 期中平均AR, 期末CR)
        ("Q1", 4.05, 4.02),
        ("Q2", 4.08, 4.08),
        ("Q3", 4.12, 4.15),
        ("Q4", 4.15, 4.20),
    ],
    "monthly_ttm": [  # 月末TTM（参考）
        ("2025-04", 4.02), ("2025-05", 4.04), ("2025-06", 4.02),
        ("2025-07", 4.06), ("2025-08", 4.08), ("2025-09", 4.08),
        ("2025-10", 4.10), ("2025-11", 4.12), ("2025-12", 4.15),
        ("2026-01", 4.16), ("2026-02", 4.18), ("2026-03", 4.20),
    ],
    "opening_CTA": 120000.0,  # 前期末 為替換算調整勘定（千円）
}

# ---------------------------------------------------------------------------
# 連結勘定科目マスタ（連結パッケージ共通科目）
# ---------------------------------------------------------------------------
# (code, name, section)  section: CA流動資産 FA固定資産 CL流動負債 NCL固定負債 EQ純資産
BS_COA = [
    ("A100", "現金及び預金", "CA"),
    ("A110", "受取手形及び売掛金", "CA"),
    ("A115", "貸倒引当金", "CA"),          # 控除項目（負値）
    ("A120", "棚卸資産", "CA"),
    ("A130", "前払費用・その他流動資産", "CA"),
    ("A140", "未収入金", "CA"),
    ("A200", "有形固定資産（純額）", "FA"),
    ("A210", "無形固定資産", "FA"),
    ("A215", "のれん", "FA"),               # 連結のみ
    ("A220", "投資有価証券", "FA"),
    ("A230", "関係会社株式", "FA"),
    ("A240", "関係会社長期貸付金", "FA"),
    ("A250", "繰延税金資産", "FA"),
    ("A260", "その他投資等", "FA"),
    ("L100", "支払手形及び買掛金", "CL"),
    ("L110", "短期借入金", "CL"),
    ("L120", "リース債務", "CL"),
    ("L130", "未払金・未払費用", "CL"),
    ("L140", "未払法人税等", "CL"),
    ("L150", "賞与引当金", "CL"),
    ("L200", "長期借入金", "NCL"),
    ("L210", "退職給付に係る負債", "NCL"),
    ("L220", "その他固定負債", "NCL"),
    ("E100", "資本金", "EQ"),
    ("E110", "資本剰余金", "EQ"),
    ("E120", "利益剰余金", "EQ"),
    ("E130", "自己株式", "EQ"),
    ("E140", "為替換算調整勘定", "EQ"),      # 連結・在外子会社のみ
    ("E150", "非支配株主持分", "EQ"),        # 連結のみ
]
ASSET_CODES = [c for c, _, s in BS_COA if s in ("CA", "FA")]
LIAB_EQ_CODES = [c for c, _, s in BS_COA if s in ("CL", "NCL", "EQ")]
BS_NAME = {c: n for c, n, _ in BS_COA}

# PL科目（収益はプラス、費用はプラスで持ち、計算時に符号処理）
PL_COA = [
    ("P100", "売上高", "rev"),
    ("P200", "売上原価", "exp"),
    ("P300", "販売費及び一般管理費", "exp"),
    ("P400a", "受取利息", "rev"),
    ("P400b", "受取配当金", "rev"),
    ("P400c", "為替差益", "rev"),
    ("P400d", "その他営業外収益", "rev"),
    ("P410a", "支払利息", "exp"),
    ("P410b", "その他営業外費用", "exp"),
    ("P500", "特別利益", "rev"),
    ("P510", "特別損失", "exp"),
    ("P600", "法人税等", "exp"),
]
PL_NAME = {c: n for c, n, _ in PL_COA}
PL_SIGN = {c: (1 if k == "rev" else -1) for c, _, k in PL_COA}


def net_income(pl):
    return sum(pl.get(c, 0) * PL_SIGN[c] for c, _, _ in PL_COA)


def operating_income(pl):
    return pl.get("P100", 0) - pl.get("P200", 0) - pl.get("P300", 0)


# ---------------------------------------------------------------------------
# 会社別データ
# ---------------------------------------------------------------------------
COMPANIES = {
    # ============================ 親会社 =================================
    "DA-HQ": {
        "name": "デモA株式会社",
        "name_en": "Demo-A Inc.",
        "role": "親会社",
        "currency": "JPY",
        "ownership": None,
        "location": "神奈川県横浜市港北区",
        "established": "1998-04-01",
        "fye": "3月31日",
        "business": "精密機器部品の製造・販売",
        "employees": 580,
        "preparer": "経理部 主任 西村 拓海",
        "reviewer": "経理部 課長 高橋 美咲 (ACC002)",
        "approver": "経理部 部長 佐藤 一郎 (ACC001)",
        "pl": {
            "P100": 33380000, "P200": 26900000, "P300": 4480000,
            "P400a": 12000, "P400b": 175000, "P400c": 0, "P400d": 25000,
            "P410a": 45000, "P410b": 35000,
            "P500": 15000, "P510": 55000, "P600": 645000,
        },
        "bs_close": {
            "A100": 2100000, "A110": 8200000, "A115": -41000, "A120": 3100000,
            "A130": 250000, "A140": 0,
            "A200": 9800000, "A210": 450000, "A220": 600000, "A230": 780000,
            "A240": 210000, "A250": 320000, "A260": 180000,
            "L100": 5400000, "L110": 1000000, "L120": 0, "L130": 900000,
            "L140": 320000, "L150": 520000,
            "L200": 2200000, "L210": 1450000, "L220": 559000,
            "E100": 1200000, "E110": 1400000, "E120": 11050000, "E130": -50000,
        },
        "bs_open": {
            "A100": 1507250, "A110": 7850000, "A115": -39250, "A120": 2980000,
            "A130": 240000, "A140": 0,
            "A200": 9950000, "A210": 470000, "A220": 570000, "A230": 780000,
            "A240": 210000, "A250": 300000, "A260": 175000,
            "L100": 5150000, "L110": 1200000, "L120": 0, "L130": 850000,
            "L140": 300000, "L150": 500000,
            "L200": 2400000, "L210": 1400000, "L220": 540000,
            "E100": 1200000, "E110": 1400000, "E120": 10103000, "E130": -50000,
        },
        "dividend_paid": 500000,   # 外部株主への配当
        "inventory_detail": {"原材料": 900000, "仕掛品": 750000, "製品": 1450000},
        "inventory_detail_open": {"原材料": 870000, "仕掛品": 730000, "製品": 1380000},
        "fixed_assets": {  # 有形固定資産（純額ベース）
            "open": 9950000, "capex": 800000, "dep": 895000, "disposal": 55000,
        },
        "intangibles": {"open": 470000, "capex": 60000, "amort": 80000},
        "sga_detail": [
            ("人件費（給料・賞与・法定福利）", 2050000),
            ("運搬費（うちデモA物流委託 1,140,000）", 1310000),
            ("減価償却費", 180000),
            ("賃借料", 190000),
            ("研究開発費", 380000),
            ("貸倒引当金繰入額", 2500),
            ("その他", 367500),
        ],
        "tax_rate": 0.306,
        "notes_dep_in_cogs": 715000,  # 売上原価に含まれる減価償却費
    },
    # ========================== デモA東北 ================================
    "DA-TB": {
        "name": "デモA東北株式会社",
        "name_en": "Demo-A Tohoku Co., Ltd.",
        "role": "製造子会社",
        "currency": "JPY",
        "ownership": 1.00,
        "location": "宮城県仙台市",
        "established": "2005-10-01（設立時より100%出資）",
        "fye": "3月31日",
        "business": "精密金属部品の製造（切削加工）",
        "employees": 120,
        "preparer": "経理課 課長 大橋 直樹",
        "reviewer": "管理部長 米田 康夫",
        "approver": "代表取締役社長 石井 泰三",
        "submitted": "2026-04-08 14:00",
        "version": "1.0",
        "pl": {
            "P100": 6000000, "P200": 5050000, "P300": 550000,
            "P400a": 2000, "P400b": 0, "P400c": 0, "P400d": 13000,
            "P410a": 8000, "P410b": 17000,
            "P500": 0, "P510": 12000, "P600": 116000,
        },
        "bs_close": {
            "A100": 620000, "A110": 1180000, "A115": -5900, "A120": 520000,
            "A130": 80000, "A140": 0,
            "A200": 1850000, "A210": 60000, "A260": 120000,
            "L100": 780000, "L110": 150000, "L120": 0, "L130": 140000,
            "L140": 60000, "L150": 95000,
            "L200": 250000, "L210": 260000, "L220": 69100,
            "E100": 90000, "E110": 10000, "E120": 2520000,
        },
        "bs_open": {
            "A100": 527475, "A110": 1095000, "A115": -5475, "A120": 490000,
            "A130": 72000, "A140": 0,
            "A200": 1905000, "A210": 66000, "A260": 118000,
            "L100": 742000, "L110": 180000, "L120": 0, "L130": 128000,
            "L140": 58000, "L150": 90000,
            "L200": 280000, "L210": 248000, "L220": 64000,
            "E100": 90000, "E110": 10000, "E120": 2378000,
        },
        "dividend_paid": 120000,  # 全額親会社へ（2025/6/25支払）
        "inventory_detail": {"原材料": 160000, "仕掛品": 150000, "製品": 210000},
        "inventory_detail_open": {"原材料": 150000, "仕掛品": 140000, "製品": 200000},
        "fixed_assets": {"open": 1905000, "capex": 130000, "dep": 173000, "disposal": 12000},
        "intangibles": {"open": 66000, "capex": 6000, "amort": 12000},
        "cogs_detail": [  # 製造原価報告書
            ("材料費（うちグループ有償支給材 620,000）", 2620000),
            ("労務費", 1280000),
            ("経費（うち減価償却費 153,000）", 1180000),
            ("当期総製造費用", 5080000),
            ("仕掛品・製品棚卸高増減", -30000),
            ("売上原価", 5050000),
        ],
        "sga_detail": [
            ("人件費", 280000), ("運搬費", 90000), ("減価償却費", 20000),
            ("その他", 160000),
        ],
        "tax_rate": 0.306,
        "special_loss_note": "固定資産除却損 12,000千円（旧切削ライン設備、2025年12月除却。Q3パッケージで科目訂正済）",
    },
    # ======================== デモA物流サービス ==========================
    "DA-LOG": {
        "name": "デモA物流サービス株式会社",
        "name_en": "Demo-A Logistics Service Co., Ltd.",
        "role": "物流子会社",
        "currency": "JPY",
        "ownership": 1.00,
        "location": "神奈川県横浜市",
        "established": "2010-04-01（設立時より100%出資）",
        "fye": "3月31日",
        "business": "グループ製品の保管・輸配送",
        "employees": 40,
        "preparer": "総務経理課 主任 藤本 さくら",
        "reviewer": "総務経理課 課長 内山 誠",
        "approver": "代表取締役社長 村井 里佳",
        "submitted": "2026-04-08 17:00",
        "version": "1.0",
        "pl": {
            "P100": 1500000, "P200": 1230000, "P300": 185000,
            "P400a": 500, "P400b": 0, "P400c": 0, "P400d": 2500,
            "P410a": 2000, "P410b": 1000,
            "P500": 0, "P510": 0, "P600": 25000,
        },
        "bs_close": {
            "A100": 210000, "A110": 262000, "A115": -1310, "A120": 0,
            "A130": 20000, "A140": 15000,
            "A200": 500000, "A210": 25000, "A260": 45000,
            "L100": 148000, "L110": 0, "L120": 130000, "L130": 62000,
            "L140": 15000, "L150": 35000,
            "L200": 0, "L210": 0, "L220": 55690,
            "E100": 30000, "E110": 0, "E120": 600000,
        },
        "bs_open": {
            "A100": 150200, "A110": 240000, "A115": -1200, "A120": 0,
            "A130": 18000, "A140": 13000,
            "A200": 520000, "A210": 27000, "A260": 44000,
            "L100": 140000, "L110": 0, "L120": 145000, "L130": 58000,
            "L140": 14000, "L150": 32000,
            "L200": 0, "L210": 0, "L220": 52000,
            "E100": 30000, "E110": 0, "E120": 540000,
        },
        "dividend_paid": 0,
        "inventory_detail": {},
        "inventory_detail_open": {},
        "fixed_assets": {"open": 520000, "capex": 45000, "dep": 65000, "disposal": 0},
        "intangibles": {"open": 27000, "capex": 2000, "amort": 4000},
        "sga_detail": [
            ("人件費", 95000), ("賃借料", 38000), ("その他", 52000),
        ],
        "tax_rate": 0.306,
        "tax_note": "税負担率29.4%（住民税均等割・軽減税率の影響）",
    },
    # ======================== Demo-A (Thailand) ==========================
    "DAT": {
        "name": "Demo-A (Thailand) Co., Ltd.",
        "name_en": "Demo-A (Thailand) Co., Ltd.",
        "role": "海外製造子会社",
        "currency": "THB",
        "ownership": 1.00,
        "location": "タイ王国 サムットプラーカーン県（バンコク近郊）",
        "established": "2012-07-01（設立時より100%出資、資本金THB 30,000千 @3.00円）",
        "fye": "3月31日（親会社と統一）",
        "business": "精密金属部品の製造（プレス加工）・ASEAN域内販売",
        "employees": 80,
        "preparer": "Accounting Manager Somchai Prasert",
        "reviewer": "Finance Director 川島 亮（出向者）",
        "approver": "Managing Director 池谷 修平",
        "submitted": "2026-04-08 17:00",
        "version": "1.0",
        "pl": {  # 千THB
            "P100": 610000, "P200": 470000, "P300": 76000,
            "P400a": 0, "P400b": 0, "P400c": 2600, "P400d": 0,
            "P410a": 1200, "P410b": 800,
            "P500": 0, "P510": 0, "P600": 12920,
        },
        "bs_close": {  # 千THB
            "A100": 69500, "A110": 152000, "A115": -1500, "A120": 98000,
            "A130": 12000, "A140": 0,
            "A200": 285000, "A210": 0, "A260": 15000,
            "L100": 96000, "L110": 0, "L120": 0, "L130": 22000,
            "L140": 8000, "L150": 0,
            "L200": 50000, "L210": 0, "L220": 14000,
            "E100": 30000, "E110": 0, "E120": 410000,
        },
        "bs_open": {  # 千THB
            "A100": 25700, "A110": 138000, "A115": -1380, "A120": 90000,
            "A130": 11000, "A140": 0,
            "A200": 292000, "A210": 0, "A260": 14000,
            "L100": 88000, "L110": 0, "L120": 0, "L130": 20000,
            "L140": 7500, "L150": 0,
            "L200": 52500, "L210": 0, "L220": 13000,
            "E100": 30000, "E110": 0, "E120": 358320,
        },
        "dividend_paid": 0,
        "inventory_detail": {"原材料 Raw materials": 45000, "仕掛品 WIP": 18000, "製品 Finished goods": 35000},
        "inventory_detail_open": {"原材料 Raw materials": 42000, "仕掛品 WIP": 16000, "製品 Finished goods": 32000},
        "fixed_assets": {"open": 292000, "capex": 28000, "dep": 35000, "disposal": 0},
        "intangibles": {"open": 0, "capex": 0, "amort": 0},
        "cogs_detail": [
            ("材料費 Materials（うち親会社仕入 THB185,366 / JPY 760,000）", 268000),
            ("労務費 Labor", 108000),
            ("経費 Overheads（うち減価償却費 THB33,000）", 96000),
            ("当期総製造費用 Total mfg. cost", 472000),
            ("仕掛品・製品棚卸高増減 Inventory change", -2000),
            ("売上原価 COGS", 470000),
        ],
        "sga_detail": [
            ("人件費 Personnel", 32000), ("輸出諸掛 Export expenses", 18000),
            ("その他 Others", 26000),
        ],
        "tax_rate": 0.20,
        "loan_from_parent": {  # 親会社からの長期借入（JPY建て）
            "jpy": 210000, "rate": "年2.35%（固定）", "maturity": "2028-03-31",
            "interest_jpy": 4920, "interest_thb": 1200,
        },
        "notes": "洪水リスクに備えBCP保険付保。JPY建て借入210,000千円は期末レート4.20で換算（THB50,000千）。",
    },
    # ====================== デモAトレーディング ==========================
    "DATR": {
        "name": "株式会社デモAトレーディング",
        "name_en": "Demo-A Trading Co., Ltd.",
        "role": "商社子会社",
        "currency": "JPY",
        "ownership": 0.80,
        "location": "東京都中央区",
        "established": "2019-04-01 株式取得（80%、取得原価560,000千円）",
        "fye": "3月31日",
        "business": "精密機器部品・関連商材の国内販売",
        "employees": 30,
        "preparer": "経理グループ リーダー 後藤 千夏",
        "reviewer": "管理部長 三宅 宏",
        "approver": "代表取締役社長 榊原 元",
        "submitted": "2026-04-08 11:00",
        "version": "1.0",
        "pl": {
            "P100": 4500000, "P200": 3830000, "P300": 420000,
            "P400a": 1000, "P400b": 0, "P400c": 0, "P400d": 3000,
            "P410a": 5000, "P410b": 3000,
            "P500": 0, "P510": 0, "P600": 75600,
        },
        "bs_close": {
            "A100": 380000, "A110": 980000, "A115": -12900, "A120": 310000,
            "A130": 30000, "A140": 0,
            "A200": 90000, "A210": 0, "A260": 65000,
            "L100": 890000, "L110": 0, "L120": 0, "L130": 95000,
            "L140": 40000, "L150": 45000,
            "L200": 0, "L210": 0, "L220": 62100,
            "E100": 90000, "E110": 10000, "E120": 610000,
        },
        "bs_open": {
            "A100": 279175, "A110": 915000, "A115": -9575, "A120": 285000,
            "A130": 27000, "A140": 0,
            "A200": 96000, "A210": 0, "A260": 63000,
            "L100": 842000, "L110": 0, "L120": 0, "L130": 88000,
            "L140": 36000, "L150": 42000,
            "L200": 0, "L210": 0, "L220": 58000,
            "E100": 90000, "E110": 10000, "E120": 489600,
        },
        "dividend_paid": 50000,  # 2025/6/20支払（親会社40,000 / 非支配株主10,000）
        "inventory_detail": {"商品": 310000},
        "inventory_detail_open": {"商品": 285000},
        "fixed_assets": {"open": 96000, "capex": 8000, "dep": 14000, "disposal": 0},
        "intangibles": {"open": 0, "capex": 0, "amort": 0},
        "sga_detail": [
            ("人件費", 210000), ("運搬費（うちデモA物流委託 100,000）", 118000),
            ("貸倒引当金繰入額（個別引当：サンプル顧客R社）", 8000),
            ("その他", 84000),
        ],
        "tax_rate": 0.306,
        "baddebt_note": "サンプル顧客R社（民事再生手続開始 2025/11）に対する債権16,000千円に50%個別引当。",
    },
}

SUBS = ["DA-TB", "DA-LOG", "DAT", "DATR"]
ALL = ["DA-HQ"] + SUBS

# ---------------------------------------------------------------------------
# グループ内取引（年間取引高・期末/期首残高）  単位:千円（JPY建て取引額）
# ---------------------------------------------------------------------------
# (売り手, 買い手, 取引内容, 年間取引高, 買い手側計上区分)
IC_TRANSACTIONS = [
    ("DA-TB", "DA-HQ", "製品部品の販売（切削加工品）", 4000000, "売上原価(材料費)"),
    ("DA-HQ", "DATR", "製品（商品）の供給", 2780000, "売上原価(商品仕入)"),
    ("DA-HQ", "DAT", "原材料の輸出販売（JPY建て）", 760000, "売上原価(材料費)"),
    ("DA-HQ", "DA-TB", "原材料の有償支給（原価売買・利益率0%）", 620000, "売上原価(材料費)"),
    ("DAT", "DA-HQ", "製品部品の輸出販売（JPY建て）", 480000, "売上原価(材料費)"),
    ("DA-LOG", "DA-HQ", "物流業務受託（保管・輸配送）", 1140000, "販管費(運搬費)"),
    ("DA-LOG", "DATR", "物流業務受託（輸配送）", 100000, "販管費(運搬費)"),
]
# 検算: 合計 9,880,000 = S05四半期消去 (1,820,000+650,000)*4

# (売り手=債権者, 買い手=債務者, 期末残高, 期首残高)  ※JPY建て
IC_BALANCES = [
    ("DA-TB", "DA-HQ", 132000, 118000),
    ("DA-HQ", "DATR", 95500, 82000),    # 期末は未達6,200を含む（DATR帳簿上は89,300）
    ("DA-LOG", "DA-HQ", 27700, 25600),
    ("DAT", "DA-HQ", 42000, 38000),     # JPY建て請求（DAT帳簿 THB10,000 @4.20 / 9,500 @4.00）
    ("DA-HQ", "DAT", 27300, 26000),     # JPY建て請求（DAT帳簿 THB6,500 @4.20 / 6,500 @4.00）
]
IC_AR_TOTAL_CLOSE = sum(x[2] for x in IC_BALANCES)   # = 324,500
IC_AR_TOTAL_OPEN = sum(x[3] for x in IC_BALANCES)    # = 289,600

# 未達取引（期末照合差異）
TRANSIT = {
    "seller": "DA-HQ", "buyer": "DATR", "amount": 6200,
    "desc": "2026/3/30出荷（出荷番号SH-260330-018）・2026/4/1 DATR着荷の未達商品",
    "buyer_booked_ap": 89300,
}

# グループ内資金取引
IC_LOAN = {
    "lender": "DA-HQ", "borrower": "DAT", "jpy": 210000,
    "interest": 4920, "note": "長期貸付（JPY建て・年2.35%・2028/3/31返済期限）",
}

# 未実現利益（棚卸資産）
UNREALIZED = {
    "close": [
        # (売り手, 保有者, 期末在庫額, 売り手利益率, 未実現利益)
        ("DA-TB", "DA-HQ", 280000, 0.11, 30800),
        ("DA-HQ", "DATR", 60000, 0.20, 12000),  # 手許53,800 + 未達6,200
    ],
    "open": [
        ("DA-TB", "DA-HQ", 260000, 0.11, 28600),
        ("DA-HQ", "DATR", 50000, 0.20, 10000),
    ],
}
UNREALIZED_CLOSE = sum(x[4] for x in UNREALIZED["close"])  # 42,800
UNREALIZED_OPEN = sum(x[4] for x in UNREALIZED["open"])    # 38,600
URP_TAX_RATE = 0.306
URP_DTA_CLOSE = 13097  # 42,800×30.6%（千円未満四捨五入）
URP_DTA_OPEN = 11812   # 38,600×30.6%

# のれん（DATR 80%取得 2019/4/1）
GOODWILL = {
    "cost": 160000, "life": 10, "annual_amort": 16000,
    "accum_open": 96000, "accum_close": 112000,
    "nbv_open": 64000, "nbv_close": 48000,
    "acq_equity": {"E100": 90000, "E110": 10000, "E120": 400000},  # 取得時DATR純資産500,000
    "acq_cost": 560000, "acq_nci": 100000,
}

# 非支配株主持分（DATR 20%）
NCI = {
    "open": 117920,      # 20% × 期首純資産589,600
    "pl": 34080,         # 20% × 当期純利益170,400
    "dividend": 10000,   # 20% × 配当50,000
    "close": 142000,     # 20% × 期末純資産710,000
}

# 配当消去
DIVIDENDS = {
    "DA-TB": {"total": 120000, "to_parent": 120000, "to_nci": 0, "date": "2025-06-25"},
    "DATR": {"total": 50000, "to_parent": 40000, "to_nci": 10000, "date": "2025-06-20"},
}
DIV_TO_PARENT = 160000

# ---------------------------------------------------------------------------
# 換算（DAT: THB→JPY）
# ---------------------------------------------------------------------------

def translate_dat():
    """DATのBS/PLを円換算し、CTA（為替換算調整勘定）をプラグ計算する。"""
    c = COMPANIES["DAT"]
    ar, cr_c, cr_o, hr = FX["AR"], FX["CR_close"], FX["CR_open"], FX["HR_capital"]

    pl_jpy = {k: round(v * ar) for k, v in c["pl"].items()}
    ni_thb = net_income(c["pl"])
    ni_jpy = round(ni_thb * ar)

    def _trans_bs(bs, cr, re_hist, cta):
        out = {}
        for code in bs:
            if code == "E100":
                out[code] = round(bs[code] * hr)
            elif code == "E120":
                out[code] = re_hist
            else:
                out[code] = round(bs[code] * cr)
        out["E140"] = cta
        return out

    # 期首: CTAは既知（FX["opening_CTA"]）→ 期首利益剰余金（換算後）を逆算
    eq_open_jpy_at_cr = round((c["bs_open"]["E100"] + c["bs_open"]["E120"]) * cr_o)
    re_hist_open = eq_open_jpy_at_cr - round(c["bs_open"]["E100"] * hr) - round(FX["opening_CTA"])
    bs_open_jpy = _trans_bs(c["bs_open"], cr_o, re_hist_open, round(FX["opening_CTA"]))

    # 期末: 利益剰余金 = 期首換算後 + 当期純利益(AR換算) → CTAをプラグ
    re_hist_close = re_hist_open + ni_jpy
    eq_close_jpy_at_cr = round((c["bs_close"]["E100"] + c["bs_close"]["E120"]) * cr_c)
    cta_close = eq_close_jpy_at_cr - round(c["bs_close"]["E100"] * hr) - re_hist_close
    bs_close_jpy = _trans_bs(c["bs_close"], cr_c, re_hist_close, cta_close)

    return {
        "pl": pl_jpy, "bs_open": bs_open_jpy, "bs_close": bs_close_jpy,
        "ni_jpy": ni_jpy, "re_hist_open": re_hist_open, "re_hist_close": re_hist_close,
        "cta_open": round(FX["opening_CTA"]), "cta_close": cta_close,
        "cta_delta": cta_close - round(FX["opening_CTA"]),
    }


# ---------------------------------------------------------------------------
# 検算
# ---------------------------------------------------------------------------

def validate():
    errors = []
    for code, c in COMPANIES.items():
        for key in ("bs_close", "bs_open"):
            bs = c[key]
            a = sum(bs.get(k, 0) for k in ASSET_CODES)
            le = sum(bs.get(k, 0) for k in LIAB_EQ_CODES)
            if a != le:
                errors.append(f"{code} {key}: 資産{a} ≠ 負債純資産{le} (差{a-le})")
        ni = net_income(c["pl"])
        re_expected = c["bs_open"]["E120"] + ni - c["dividend_paid"]
        if re_expected != c["bs_close"]["E120"]:
            errors.append(f"{code}: 利益剰余金ロール不一致 期待{re_expected} 実際{c['bs_close']['E120']} (NI={ni})")
        fa = c["fixed_assets"]
        if fa["open"] + fa["capex"] - fa["dep"] - fa["disposal"] != c["bs_close"]["A200"]:
            errors.append(f"{code}: 有形固定資産ロール不一致")
        ia = c["intangibles"]
        if ia["open"] + ia["capex"] - ia["amort"] != c["bs_close"].get("A210", 0):
            errors.append(f"{code}: 無形固定資産ロール不一致")
        inv = sum(c["inventory_detail"].values())
        if inv != c["bs_close"].get("A120", 0):
            errors.append(f"{code}: 棚卸資産内訳不一致 {inv} vs {c['bs_close'].get('A120',0)}")

    if sum(x[3] for x in IC_TRANSACTIONS) != 9880000:
        errors.append("内部取引高合計が9,880,000千円でない")
    if IC_AR_TOTAL_CLOSE != 324500:
        errors.append("期末内部債権債務が324,500千円でない")
    if UNREALIZED_CLOSE != 42800:
        errors.append("期末未実現利益が42,800千円でない")
    if NCI["open"] + NCI["pl"] - NCI["dividend"] != NCI["close"]:
        errors.append("非支配株主持分ロール不一致")

    # 売上のIC整合: 各社売上に占めるIC売上が取引高表と一致するか（設計値のみ確認）
    ic_sales = {}
    for s, b, _, amt, _ in IC_TRANSACTIONS:
        ic_sales[s] = ic_sales.get(s, 0) + amt
    # DATはTHB建て帳簿のため、JPY換算後売上と比較（117,073千THB×4.10≒480,000）
    assert abs(round(117073 * FX["AR"]) - 480000) < 1000

    t = translate_dat()
    for key in ("bs_open", "bs_close"):
        bs = t[key]
        a = sum(bs.get(k, 0) for k in ASSET_CODES)
        le = sum(bs.get(k, 0) for k in LIAB_EQ_CODES)
        if a != le:
            errors.append(f"DAT換算後 {key}: 貸借不一致 差{a-le}")
    return errors, t


# ---------------------------------------------------------------------------
# 連結消去仕訳（年次・累計ベース、単位:千円）
# ---------------------------------------------------------------------------
# 各仕訳: dict(no, category, auto, lines=[(dr_cr, code, desc, amount)])
# code は BS/PL科目コード。"SS_DIV" は連結株主資本等変動計算書上の剰余金の配当調整。

def build_journal():
    J = []
    J.append({
        "no": "CNS-FY25-001", "category": "開始仕訳（資本連結・過年度連結修正）", "auto": "AUTO",
        "approved": ("高橋 美咲 (ACC002) 2026-04-08", "佐藤 一郎 (ACC001) 2026-04-14"),
        "lines": [
            ("Dr", "E100", "子会社資本金の相殺（4社合計）", 300000),
            ("Dr", "E110", "子会社資本剰余金の相殺（DA-TB/DATR）", 20000),
            ("Dr", "E120", "支配獲得時利益剰余金（DATR）", 400000),
            ("Dr", "E120", "のれん過年度償却累計（FY2019-24）", 96000),
            ("Dr", "E120", "非支配株主持分への振替（過年度累計）", 17920),
            ("Dr", "A215", "のれん期首残高", 64000),
            ("Cr", "A230", "関係会社株式の相殺", 780000),
            ("Cr", "E150", "非支配株主持分期首残高", 117920),
            ("Dr", "E120", "期首棚卸資産未実現利益（実現）", 38600),
            ("Cr", "P200", "期首未実現利益の当期実現", 38600),
            ("Dr", "A250", "期首未実現利益に係る繰延税金資産", 11812),
            ("Cr", "E120", "期首未実現利益税効果", 11812),
        ],
    })
    J.append({
        "no": "CNS-FY25-002", "category": "のれん当期償却", "auto": "AUTO",
        "lines": [
            ("Dr", "P300", "のれん償却額（10年定額・7年目）", 16000),
            ("Cr", "A215", "のれん", 16000),
        ],
    })
    J.append({
        "no": "CNS-FY25-003", "category": "内部取引高消去", "auto": "AUTO",
        "lines": [
            ("Dr", "P100", "グループ内売上高（製品・材料 5社間）", 8640000),
            ("Cr", "P200", "グループ内仕入高（売上原価）", 8640000),
            ("Dr", "P100", "グループ内売上高（物流役務 DA-LOG）", 1240000),
            ("Cr", "P300", "グループ内物流費（販管費 運搬費）", 1240000),
        ],
    })
    J.append({
        "no": "CNS-FY25-004", "category": "内部債権債務消去（未達整理含む）", "auto": "AUTO",
        "lines": [
            ("Dr", "A120", "未達商品の計上（DA-HQ→DATR 3/30出荷分）", 6200),
            ("Cr", "L100", "未達に係る買掛金の計上（DATR）", 6200),
            ("Dr", "L100", "グループ内買掛金の相殺（未達整理後）", 324500),
            ("Cr", "A110", "グループ内売掛金の相殺", 324500),
        ],
    })
    J.append({
        "no": "CNS-FY25-005", "category": "資金取引消去（貸付金・利息・配当）", "auto": "AUTO",
        "lines": [
            ("Dr", "L200", "DAT長期借入金（JPY建て210,000）", 210000),
            ("Cr", "A240", "関係会社長期貸付金", 210000),
            ("Dr", "P400a", "受取利息（DA-HQ←DAT）", 4920),
            ("Cr", "P410a", "支払利息（DAT）", 4920),
            ("Dr", "P400b", "受取配当金（DA-TB 120,000・DATR 40,000）", 160000),
            ("Dr", "E150", "非支配株主への配当（DATR）", 10000),
            ("Cr", "SS_DIV", "剰余金の配当（子会社支払分の戻し）", 170000),
        ],
    })
    J.append({
        "no": "CNS-FY25-006", "category": "非支配株主持分振替", "auto": "AUTO",
        "lines": [
            ("Dr", "PNCI", "非支配株主に帰属する当期純利益（DATR 20%）", 34080),
            ("Cr", "E150", "非支配株主持分", 34080),
        ],
    })
    J.append({
        "no": "CNS-FY25-007", "category": "未実現利益消去（棚卸資産）・税効果", "auto": "MANUAL",
        "note": "S07 未実現利益計算シート（EUC）に基づき手動起票",
        "lines": [
            ("Dr", "P200", "期末棚卸資産未実現利益の消去", 42800),
            ("Cr", "A120", "棚卸資産", 42800),
            ("Dr", "A250", "未実現利益に係る繰延税金資産（当期増加）", 1285),
            ("Cr", "P600", "法人税等調整額", 1285),
        ],
    })
    return J


# ---------------------------------------------------------------------------
# 連結精算表・連結財務諸表の計算
# ---------------------------------------------------------------------------

def consolidate():
    """期末連結BS・当期連結PL・期首連結BSを計算して返す。"""
    errors, t = validate()
    if errors:
        raise AssertionError("検算エラー:\n" + "\n".join(errors))

    cols = {}
    for code in ALL:
        c = COMPANIES[code]
        if code == "DAT":
            cols[code] = {"pl": t["pl"], "bs": t["bs_close"], "bs_open": t["bs_open"]}
        else:
            cols[code] = {"pl": dict(c["pl"]), "bs": dict(c["bs_close"]), "bs_open": dict(c["bs_open"])}

    def sum_col(key, code_):
        return sum(cols[co][key].get(code_, 0) for co in ALL)

    simple_bs = {code_: sum_col("bs", code_) for code_, _, _ in BS_COA}
    simple_bs_open = {code_: sum_col("bs_open", code_) for code_, _, _ in BS_COA}
    simple_pl = {code_: sum_col("pl", code_) for code_, _, _ in PL_COA}

    # --- 期末BS・当期PLへの消去仕訳適用 ---
    journal = build_journal()
    adj_bs = {}
    adj_pl = {}
    pnci = 0
    ss_div = 0
    for j in journal:
        for drcr, code_, desc, amt in j["lines"]:
            sgn = 1 if drcr == "Dr" else -1
            if code_ == "PNCI":
                pnci += sgn * amt
            elif code_ == "SS_DIV":
                ss_div += -sgn * amt  # Cr側で正=配当の戻し
            elif code_.startswith(("A", "L", "E")):
                # 資産はDrで増加、負債純資産はCrで増加
                if code_.startswith("A"):
                    adj_bs[code_] = adj_bs.get(code_, 0) + sgn * amt
                else:
                    adj_bs[code_] = adj_bs.get(code_, 0) - sgn * amt
            else:  # PL
                # 収益はCrで増加(消去はDr=減少)、費用はDrで増加(消去はCr=減少)
                if PL_SIGN[code_] == 1:
                    adj_pl[code_] = adj_pl.get(code_, 0) - sgn * amt
                else:
                    adj_pl[code_] = adj_pl.get(code_, 0) + sgn * amt

    cons_pl = {c_: simple_pl.get(c_, 0) + adj_pl.get(c_, 0) for c_, _, _ in PL_COA}
    cons_ni = net_income(cons_pl)
    ni_parent = cons_ni - pnci

    cons_bs = {c_: simple_bs.get(c_, 0) + adj_bs.get(c_, 0) for c_, _, _ in BS_COA}
    # PL消去のRE影響と配当戻しをREに反映
    pl_elim_effect = net_income(adj_pl)   # 消去仕訳による純利益への影響
    cons_bs["E120"] = cons_bs["E120"] + pl_elim_effect + ss_div - pnci

    # --- 期首連結BS（CF比較用）: 開始状態の消去を適用 ---
    ob = dict(simple_bs_open)
    ob["E100"] -= 300000
    ob["E110"] -= 20000
    ob["E120"] -= (400000 + 96000 + 17920)          # 資本連結
    ob["E120"] -= UNREALIZED_OPEN                    # 期首未実現
    ob["E120"] += URP_DTA_OPEN                       # 期首税効果
    ob["A215"] = GOODWILL["nbv_open"]
    ob["A230"] -= 780000
    ob["E150"] = NCI["open"]
    ob["A110"] -= IC_AR_TOTAL_OPEN
    ob["L100"] -= IC_AR_TOTAL_OPEN
    ob["A240"] -= IC_LOAN["jpy"]
    ob["L200"] -= IC_LOAN["jpy"]
    ob["A120"] -= UNREALIZED_OPEN
    ob["A250"] += URP_DTA_OPEN

    # 貸借一致検証
    for label, bs in (("連結期末BS", cons_bs), ("連結期首BS", ob)):
        a = sum(bs.get(k, 0) for k in ASSET_CODES)
        le = sum(bs.get(k, 0) for k in LIAB_EQ_CODES)
        if a != le:
            raise AssertionError(f"{label} 貸借不一致: 資産{a} 負債純資産{le} 差{a-le}")

    return {
        "cols": cols, "simple_bs": simple_bs, "simple_pl": simple_pl,
        "simple_bs_open": simple_bs_open,
        "adj_bs": adj_bs, "adj_pl": adj_pl, "pnci": pnci, "ss_div": ss_div,
        "cons_bs": cons_bs, "cons_pl": cons_pl, "cons_bs_open": ob,
        "cons_ni": cons_ni, "ni_parent": ni_parent,
        "translate": t, "journal": journal,
    }


if __name__ == "__main__":
    r = consolidate()
    print("検算OK")
    print("連結売上高:", f"{r['cons_pl']['P100']:,}千円")
    op = r["cons_pl"]["P100"] - r["cons_pl"]["P200"] - r["cons_pl"]["P300"]
    print("連結営業利益:", f"{op:,}千円")
    print("連結当期純利益:", f"{r['cons_ni']:,}千円")
    print("親会社株主帰属:", f"{r['ni_parent']:,}千円")
    print("非支配株主持分期末:", f"{r['cons_bs']['E150']:,}千円")
    print("為替換算調整勘定期末:", f"{r['cons_bs']['E140']:,}千円")
    a = sum(r["cons_bs"].get(k, 0) for k in ASSET_CODES)
    print("連結総資産:", f"{a:,}千円")
