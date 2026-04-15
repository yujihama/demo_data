"""
Phase 4/5/6: PLC-I, ITGC, ITAC, ELC, FCRP 拡張

新方針（RAWデータのみ）適用。個別書類は必要分のみPDF化。
"""
import random
import sys
import calendar
from pathlib import Path
from datetime import date, datetime, timedelta

sys.path.insert(0, str(Path(__file__).parent))
from pdf_util import JPPDF
from sample_gen_util import (
    CUSTOMERS, VENDORS, PRODUCTS, RAW_MATERIALS,
    generate_systematic_samples, create_sample_list_excel, write_raw_csv
)

BASE_I = Path(r"C:\Users\nyham\work\demo_data\4.evidence\PLC-I")
BASE_AC = Path(r"C:\Users\nyham\work\demo_data\4.evidence\ITGC\AC_アクセス管理")
BASE_CM = Path(r"C:\Users\nyham\work\demo_data\4.evidence\ITGC\CM_変更管理")
BASE_OM = Path(r"C:\Users\nyham\work\demo_data\4.evidence\ITGC\OM_運用管理")
BASE_EM = Path(r"C:\Users\nyham\work\demo_data\4.evidence\ITGC\EM_外部委託管理")
BASE_ITAC = Path(r"C:\Users\nyham\work\demo_data\4.evidence\ITAC")
BASE_ELC = Path(r"C:\Users\nyham\work\demo_data\4.evidence\ELC")
BASE_FCRP = Path(r"C:\Users\nyham\work\demo_data\4.evidence\FCRP")


# ============================================================
# PLC-I-001 実地棚卸（RAW SAP MI07出力）
# ============================================================
def gen_plc_i_001():
    random.seed(1001)
    # SAP MI07 (棚卸差異レポート) を倉庫別に生成
    warehouses = [("WH-A", "本社倉庫A"), ("WH-B", "本社倉庫B"), ("WH-T", "東北工場倉庫")]
    rows = []
    item_no = 1
    for wh_code, wh_name in warehouses:
        count = random.randint(15, 25)
        for _ in range(count):
            pcode, pname, cost, _ = random.choice(PRODUCTS)
            book_qty = random.randint(50, 2500)
            actual_qty = book_qty
            if random.random() < 0.08:
                actual_qty = book_qty + random.randint(-5, 5)
            diff = actual_qty - book_qty
            diff_amt = diff * cost
            rows.append([item_no, wh_code, wh_name, pcode, pname,
                         book_qty, actual_qty, diff, cost, diff_amt])
            item_no += 1

    write_raw_csv(
        BASE_I / "PLC-I-001_RAW_SAP_MI07_棚卸差異一覧_2025年9月.csv",
        ["# SAP S/4HANA - Transaction MI07",
         "# Report:   Physical Inventory Count Difference List",
         "# Period:   2025-09-26 to 2025-09-28 (semi-annual count)",
         "# Export:   2025-09-29 17:00:22 JST",
         "# Warehouses: WH-A (本社倉庫A) / WH-B (本社倉庫B) / WH-T (東北工場倉庫)"],
        "№,倉庫コード,倉庫名,製品コード,製品名,帳簿数量,実地数量,差異数量,標準原価,差異金額",
        rows,
        footer_lines=[f"# Records: {item_no - 1}"]
    )
    print("Created: PLC-I-001_RAW_SAP_MI07_棚卸差異一覧_2025年9月.csv")


# ============================================================
# PLC-I-002 棚卸差異調整 (24件全数)
# ============================================================
def gen_plc_i_002():
    random.seed(2002)
    # 24件の差異調整仕訳（MIGO移動タイプ701/702）
    rows = []
    warehouses = ["WH-A", "WH-B", "WH-T"]
    # サンプル15は ¥850,000 の重要不備ケース（存在するが原因分析書が欠落）
    for i in range(1, 25):
        pcode, pname, cost, _ = random.choice(PRODUCTS)
        if i == 15:
            diff_qty = 68  # +68個
            diff_amt = 850_000
            warehouse = "WH-B"
            move_type = "701"  # 在庫増加
        else:
            diff_qty = random.choice([-5, -3, -2, -1, 1, 2, 3, 5])
            diff_amt = diff_qty * cost
            warehouse = random.choice(warehouses)
            move_type = "701" if diff_qty > 0 else "702"
        diff_date = date(2025, 9, random.randint(28, 30))
        rows.append([i, f"INV-ADJ-2025-09-{i:03d}",
                     diff_date.strftime("%Y-%m-%d"),
                     warehouse, pcode, pname, diff_qty, diff_amt,
                     move_type,
                     f"JV-202509-ADJ-{i:04d}",
                     "ACC004 中村 真理"])

    write_raw_csv(
        BASE_I / "PLC-I-002_全24件RAW_SAP_MIGO_棚卸差異調整仕訳.csv",
        ["# SAP S/4HANA - Transaction MIGO (Goods Movement)",
         "# Report:   Inventory Adjustment Entries (Movement Type 701/702)",
         "# Period:   2025-09-28 to 2025-09-30",
         "# Export:   2025-10-01 09:00:15 JST"],
        "№,差異調整番号,計上日,倉庫,製品コード,製品名,差異数量,差異金額,移動タイプ,仕訳番号,計上者",
        rows,
        footer_lines=["# Records: 24 (exhaustive)"]
    )
    print("Created: PLC-I-002_全24件RAW_SAP_MIGO_棚卸差異調整仕訳.csv")


# ============================================================
# PLC-I-004 原価差異分析 (3件サンプル、RAWのみ)
# ============================================================
def gen_plc_i_004():
    random.seed(4004)
    # 月次原価差異レポートRAW (代表3ヶ月)
    rows = []
    for month_label, period_start, period_end in [
        ("2025-07", date(2025, 7, 1), date(2025, 7, 31)),
        ("2025-10", date(2025, 10, 1), date(2025, 10, 31)),
        ("2026-01", date(2026, 1, 1), date(2026, 1, 31)),
    ]:
        for category in [("材料費直接", 125_800_000),
                         ("労務費直接", 48_500_000),
                         ("製造間接費", 62_000_000),
                         ("外注加工費", 18_500_000),
                         ("能率差異", 0),
                         ("配賦差異", 0)]:
            std = category[1]
            # ランダム差異
            variance_pct = random.uniform(-3, 8) / 100
            actual = int(std * (1 + variance_pct)) if std > 0 else random.randint(-500_000, 2_000_000)
            rows.append([month_label, category[0], std,
                         actual, actual - std,
                         f"{(actual - std) / std * 100:.2f}%" if std > 0 else "N/A"])

    write_raw_csv(
        BASE_I / "PLC-I-004_3ヶ月サンプルRAW_SAP_CO88_原価差異計算結果.csv",
        ["# SAP Controlling (CO) - Transaction CO88",
         "# Report:   Variance Calculation Results (Settlement)",
         "# Sampling: 3 representative months out of FY2025 (12 months population)",
         "# Export:   2026-02-14 11:00:08 JST"],
        "対象月,費目分類,標準原価(円),実際原価(円),差異(円),差異率",
        rows,
        footer_lines=["# Records: 18 (6 categories x 3 months)"]
    )
    print("Created: PLC-I-004_3ヶ月サンプルRAW_SAP_CO88_原価差異計算結果.csv")


# ============================================================
# PLC-I-005 滞留在庫評価損 (全4回四半期)
# ============================================================
def gen_plc_i_005():
    random.seed(5005)
    rows = []
    for q_label, base_date in [
        ("Q1", date(2025, 6, 30)), ("Q2", date(2025, 9, 30)),
        ("Q3", date(2025, 12, 31)), ("Q4", date(2026, 3, 31)),
    ]:
        # 滞留在庫アイテム (5-8件/四半期)
        item_count = random.randint(5, 8)
        for _ in range(item_count):
            pcode, pname, cost, _ = random.choice(PRODUCTS)
            qty = random.randint(20, 2000)
            book_value = qty * cost
            # 最終出庫日: 12-30ヶ月前
            days_ago = random.randint(365, 900)
            last_issue = base_date - timedelta(days=days_ago)
            aging_months = days_ago // 30
            if aging_months >= 24:
                rate = 100
            elif aging_months >= 18:
                rate = 80
            else:
                rate = 50
            impairment = int(book_value * rate / 100)
            rows.append([q_label, base_date.strftime("%Y-%m-%d"),
                         pcode, pname, qty, book_value,
                         last_issue.strftime("%Y-%m-%d"),
                         f"{aging_months}ヶ月",
                         f"{rate}%", impairment])

    write_raw_csv(
        BASE_I / "PLC-I-005_全4四半期RAW_SAP_MB52_滞留在庫評価損計算結果.csv",
        ["# SAP S/4HANA - Transaction MB52 (Inventory Status) + Aging Analysis",
         "# Report:   Slow-moving Inventory Valuation Impairment",
         "# Aging rules: 12-18 months = 50% / 18-24 months = 80% / over 24 months = 100%",
         "# Period:   FY2025 Q1 to Q4 (exhaustive)",
         "# Export:   2026-04-05 10:00:00 JST"],
        "四半期,基準日,製品コード,製品名,在庫数,帳簿残高,最終出庫日,滞留期間,評価率,評価損額",
        rows,
        footer_lines=["# Records: FY2025 全4四半期（全数）"]
    )
    print("Created: PLC-I-005_全4四半期RAW_SAP_MB52_滞留在庫評価損計算結果.csv")


# ============================================================
# PLC-I-007 月次原価計算締め (12件全数RAW)
# ============================================================
def gen_plc_i_007():
    rows = []
    for month_offset in range(12):
        y = 2025 if month_offset < 9 else 2026
        m = 4 + month_offset if month_offset < 9 else month_offset - 8
        last_day = calendar.monthrange(y, m)[1]
        close_date = date(y, m, last_day)
        # Close date ~ 5営業日後
        exec_date = close_date + timedelta(days=5)
        rows.append([f"{y}-{m:02d}", close_date.strftime("%Y-%m-%d"),
                     exec_date.strftime("%Y-%m-%d"),
                     "CLOSED", "SAP_BATCH",
                     f"SAP_MM_CLOSE_{y}{m:02d}",
                     "完了"])

    write_raw_csv(
        BASE_I / "PLC-I-007_全12ヶ月RAW_SAP_MMPV_原価計算締めログ.csv",
        ["# SAP S/4HANA - Transaction MMPV (Period Close)",
         "# Report:   Monthly Period Close Execution Log",
         "# Period:   FY2025 (12 months exhaustive)",
         "# Export:   2026-04-10 09:00:00 JST"],
        "対象月,月末日,締め実行日,ステータス,実行ユーザ,バッチID,結果",
        rows,
        footer_lines=["# Records: 12 (exhaustive)"]
    )
    print("Created: PLC-I-007_全12ヶ月RAW_SAP_MMPV_原価計算締めログ.csv")


# ============================================================
# ITGC-AC-001 25件のユーザ登録申請 + RAW
# ============================================================
def gen_itgc_ac_001():
    random.seed(10001)
    # 25件の登録ケース（母集団28件）
    depts = ["営業本部", "購買部", "経理部", "製造本部", "情報システム部", "品質保証部"]
    roles_options = ["SD_USER", "MM_USER,PO_CREATE", "FI_USER",
                     "PP_SUP", "HR_USER", "BASIS"]
    samples = []
    for i in range(1, 26):
        month = random.choice([4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3])
        y = 2026 if month <= 3 else 2025
        day = random.randint(1, 25)
        apply_date = date(y, month, day)
        dept = random.choice(depts)
        role = random.choice(roles_options)
        user_id = f"U-2025-{i * 37:04d}"
        samples.append({
            "no": i, "app_no": f"USER-REG-2025-{i:04d}",
            "date": apply_date, "user_id": user_id,
            "dept": dept, "role": role,
        })

    create_sample_list_excel(
        BASE_AC / "ITGC-AC-001_監査対象25件サンプルリスト.xlsx",
        "【ITGC-AC-001】監査対象25件サンプルリスト（新規ユーザ登録承認）",
        "（RAWデータをナビゲートするための取引リスト）",
        [("母集団", "FY2025 ユーザ新規登録 28件（SAP SU01）"),
         ("抽出方法", "系統抽出 / 25件（ほぼ全数）"),
         ("抽出日時", "2026-02-15 09:00 JST"),
         ("関連RAWデータ", "ITGC-AC-001_25件対応_RAW_*.csv / 代表5件の申請書PDF")],
        ["サンプル\n№", "申請番号", "申請日", "ユーザID", "所属部門", "付与ロール"],
        [[s["no"], s["app_no"], s["date"], s["user_id"],
          s["dept"], s["role"]] for s in samples],
        col_widths=[6, 18, 11, 14, 18, 22],
        col_center=(0, 1, 3, 4), col_date=(2,),
    )
    print("Created: ITGC-AC-001_監査対象25件サンプルリスト.xlsx")

    # SAP SU01変更履歴 RAW
    rows = []
    for s in samples:
        reg_ts = datetime.combine(s["date"], datetime.min.time()) + timedelta(
            hours=random.randint(10, 16))
        rows.append([reg_ts.strftime("%Y-%m-%d %H:%M:%S"), s["no"],
                     s["app_no"], s["user_id"], "CREATE",
                     s["dept"], s["role"],
                     "IT004 西田 徹", "承認済"])

    write_raw_csv(
        BASE_AC / "ITGC-AC-001_25件対応_RAW_SAP_SU01_ユーザ作成履歴.csv",
        ["# SAP S/4HANA - Transaction SU01",
         "# Report:   User Master Change History (Table USR02 / USH02)",
         "# Export:   2026-02-15 09:20:30 JST",
         "# Filter:   25 samples under audit IA-REQ-2026-IT001"],
        "タイムスタンプ,サンプル№,申請番号,ユーザID,アクション,部門,付与ロール,実行ユーザ,ステータス",
        rows,
        footer_lines=["# Records: 25"]
    )
    print("Created: ITGC-AC-001_25件対応_RAW_SAP_SU01_ユーザ作成履歴.csv")

    # ワークフロー承認履歴
    wf_rows = []
    for s in samples:
        start = datetime.combine(s["date"], datetime.min.time()) + timedelta(hours=9)
        wf_no = f"WF-IT-2025-{s['no'] * 13:04d}"
        wf_rows.append([start.strftime("%Y-%m-%d %H:%M:%S"), wf_no, s["no"],
                        s["app_no"], "申請者(所属部門)", "起票"])
        wf_rows.append([(start + timedelta(hours=2)).strftime("%Y-%m-%d %H:%M:%S"),
                        wf_no, s["no"], s["app_no"],
                        "部門長", "承認"])
        wf_rows.append([(start + timedelta(hours=4)).strftime("%Y-%m-%d %H:%M:%S"),
                        wf_no, s["no"], s["app_no"],
                        "加藤 洋子 (IT003)", "承認"])

    write_raw_csv(
        BASE_AC / "ITGC-AC-001_25件対応_RAW_ワークフロー承認履歴.csv",
        ["# Workflow System - User Registration Approval History",
         "# Export:   2026-02-15 09:25:12 JST"],
        "タイムスタンプ,ワークフロー番号,サンプル№,申請番号,アクター,アクション",
        wf_rows,
        footer_lines=["# Records: 75"]
    )
    print("Created: ITGC-AC-001_25件対応_RAW_ワークフロー承認履歴.csv")

    # 代表5件の申請書PDF
    for s in samples[:5]:
        pdf = JPPDF()
        pdf.add_page()
        pdf.h1("SAPユーザ登録申請書")
        pdf.set_font("YuGoth", "", 10)
        pdf.cell(0, 5, f"申請番号: {s['app_no']} / 申請日: {s['date'].strftime('%Y年%m月%d日')}",
                 align="R", new_x="LMARGIN", new_y="NEXT")
        pdf.ln(5)

        pdf.kv("申請部門", s["dept"])
        pdf.kv("申請者", "申請者（部門）")
        pdf.kv("登録対象者", f"{s['user_id']} (新入社員/異動)")
        pdf.kv("申請理由", "業務遂行に必要なアクセス権付与")
        pdf.ln(5)

        pdf.h2("1. 付与希望ロール")
        pdf.table_header(["ロール名", "内容", "業務上の必要性"], [30, 60, 90])
        for r in s["role"].split(","):
            pdf.table_row([r, "業務機能", "日常業務のため"],
                          [30, 60, 90])
        pdf.ln(3)

        pdf.h2("2. 職務分掌(SoD)チェック")
        pdf.body("付与予定ロールの組合せについてSoD違反なし。")
        pdf.ln(5)

        pdf.h3("■ 承認経路")
        pdf.set_font("YuGoth", "B", 10)
        pdf.cell(45, 7, "役割", border=1, align="C", fill=True)
        pdf.cell(55, 7, "氏名", border=1, align="C", fill=True)
        pdf.cell(40, 7, "承認日時", border=1, align="C", fill=True)
        pdf.cell(30, 7, "承認印", border=1, align="C", fill=True,
                 new_x="LMARGIN", new_y="NEXT")
        approvers = [("申請部門長", "部門長", (s["date"] + timedelta(hours=2)).strftime("%Y/%m/%d")),
                     ("情シス部アプリリーダー", "加藤 洋子 (IT003)",
                      (s["date"] + timedelta(hours=4)).strftime("%Y/%m/%d"))]
        pdf.set_font("YuGoth", "", 10)
        for role, name, dt in approvers:
            pdf.cell(45, 14, role, border=1, align="C")
            pdf.cell(55, 14, name, border=1, align="C")
            pdf.cell(40, 14, dt, border=1, align="C")
            x_stamp = pdf.get_x()
            y_stamp = pdf.get_y()
            pdf.cell(30, 14, "", border=1, new_x="LMARGIN", new_y="NEXT")
            pdf.stamp("承認", x=x_stamp + 15, y=y_stamp + 7)

        out = BASE_AC / f"ITGC-AC-001_25件対応_ユーザ登録申請書_サンプル{s['no']:02d}_{s['app_no']}.pdf"
        pdf.output(str(out))
    print("Created: 5 user registration PDFs for ITGC-AC-001")


# ============================================================
# ITGC-CM-001 25件の変更申請 + RAW
# ============================================================
def gen_itgc_cm_001():
    random.seed(11001)
    changes = [
        "販売価格マスタ連携IF修正", "ワークフロー承認ルーティング変更",
        "標準原価計算バッチ修正", "勘定科目マスタ追加",
        "仕入先マスタ項目拡張", "セキュリティパッチ適用",
        "バックアップバッチ改善", "売上レポート機能追加",
        "購買申請画面の改善", "連結仕訳バリデーション強化",
    ]
    samples = []
    for i in range(1, 26):
        month = random.choice([4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3])
        y = 2026 if month <= 3 else 2025
        day = random.randint(1, 25)
        app_date = date(y, month, day)
        prod_date = app_date + timedelta(days=random.randint(10, 30))
        samples.append({
            "no": i, "rel_no": f"REL-2025-{i * 2:03d}",
            "app_date": app_date, "prod_date": prod_date,
            "change": random.choice(changes),
        })

    create_sample_list_excel(
        BASE_CM / "ITGC-CM-001_監査対象25件サンプルリスト.xlsx",
        "【ITGC-CM-001】監査対象25件サンプルリスト（変更申請・承認）",
        "（RAWデータをナビゲートするための取引リスト）",
        [("母集団", "FY2025 変更申請 42件"),
         ("抽出方法", "系統抽出"),
         ("抽出日時", "2026-02-18 10:00 JST"),
         ("関連RAWデータ", "ITGC-CM-001_25件対応_RAW_*.csv / 代表3件の変更申請書PDF")],
        ["サンプル\n№", "REL番号", "申請日", "本番移送日", "変更内容"],
        [[s["no"], s["rel_no"], s["app_date"], s["prod_date"], s["change"]]
         for s in samples],
        col_widths=[6, 14, 11, 11, 40],
        col_center=(0, 1), col_date=(2, 3),
    )
    print("Created: ITGC-CM-001_監査対象25件サンプルリスト.xlsx")

    # 変更管理台帳 RAW
    rows = []
    for s in samples:
        rows.append([s["no"], s["rel_no"], s["app_date"].strftime("%Y-%m-%d"),
                     "加藤 洋子 (IT003)", s["change"], "低リスク",
                     "UAT合格",
                     s["prod_date"].strftime("%Y-%m-%d"),
                     "加藤 洋子 (IT003)", "業務部門長",
                     "完了"])

    write_raw_csv(
        BASE_CM / "ITGC-CM-001_25件対応_RAW_変更管理台帳.csv",
        ["# ITGC Change Management System / Change Register",
         "# Export:   2026-02-18 10:15:08 JST"],
        "サンプル№,REL番号,申請日,申請者,変更内容,リスク評価,テスト結果,本番移送日,承認者(IT),承認者(業務),ステータス",
        rows,
        footer_lines=["# Records: 25"]
    )
    print("Created: ITGC-CM-001_25件対応_RAW_変更管理台帳.csv")

    # 代表3件の変更申請書PDF
    for s in samples[:3]:
        pdf = JPPDF()
        pdf.add_page()
        pdf.h1("SAP変更申請書")
        pdf.set_font("YuGoth", "", 10)
        pdf.cell(0, 5, f"REL番号: {s['rel_no']} / 申請日: {s['app_date'].strftime('%Y年%m月%d日')}",
                 align="R", new_x="LMARGIN", new_y="NEXT")
        pdf.ln(5)

        pdf.kv("件名", s["change"], key_w=30)
        pdf.kv("申請者", "加藤 洋子（情シス部アプリリーダー IT003）", key_w=30)
        pdf.kv("リスクレベル", "低（局所的な機能修正）", key_w=30)
        pdf.kv("本番移送予定", s["prod_date"].strftime("%Y年%m月%d日"), key_w=30)
        pdf.ln(5)

        pdf.h2("1. 変更理由・概要")
        pdf.body(f"{s['change']}に関する変更。"
                 "ユーザ要望およびシステム改善による対応。")
        pdf.ln(5)

        pdf.h2("2. テスト計画・結果")
        pdf.table_header(["フェーズ", "期間", "結果"], [50, 45, 45])
        pdf.table_row(["単体テスト", "開発環境", "合格"], [50, 45, 45])
        pdf.table_row(["結合テスト", "テスト環境", "合格"], [50, 45, 45], fill=True)
        pdf.table_row(["UAT", "テスト環境", "合格"], [50, 45, 45])
        pdf.ln(5)

        pdf.h3("■ 承認経路")
        pdf.set_font("YuGoth", "B", 10)
        pdf.cell(50, 7, "役割", border=1, align="C", fill=True)
        pdf.cell(55, 7, "氏名", border=1, align="C", fill=True)
        pdf.cell(40, 7, "承認日", border=1, align="C", fill=True)
        pdf.cell(30, 7, "承認印", border=1, align="C", fill=True,
                 new_x="LMARGIN", new_y="NEXT")
        approvers = [
            ("情シス部アプリリーダー", "加藤 洋子 (IT003)",
             s["app_date"].strftime("%Y/%m/%d")),
            ("業務部門責任者", "対応部門長",
             (s["app_date"] + timedelta(days=1)).strftime("%Y/%m/%d")),
            ("情シス部長", "岡田 宏 (IT001)",
             (s["app_date"] + timedelta(days=2)).strftime("%Y/%m/%d")),
        ]
        pdf.set_font("YuGoth", "", 10)
        for role, name, dt in approvers:
            pdf.cell(50, 14, role, border=1, align="C")
            pdf.cell(55, 14, name, border=1, align="C")
            pdf.cell(40, 14, dt, border=1, align="C")
            x_stamp = pdf.get_x()
            y_stamp = pdf.get_y()
            pdf.cell(30, 14, "", border=1, new_x="LMARGIN", new_y="NEXT")
            pdf.stamp("承認", x=x_stamp + 15, y=y_stamp + 7)

        out = BASE_CM / f"ITGC-CM-001_25件対応_変更申請書_サンプル{s['no']:02d}_{s['rel_no']}.pdf"
        pdf.output(str(out))
    print("Created: 3 change request PDFs for ITGC-CM-001")


# ============================================================
# ITGC-CM-002 UATテスト結果 (25件RAW)
# ============================================================
def gen_itgc_cm_002():
    random.seed(12001)
    rows = []
    for i in range(1, 26):
        rel_no = f"REL-2025-{i * 2:03d}"
        test_date = date(2025, random.randint(4, 12), random.randint(1, 28))
        case_count = random.randint(4, 12)
        pass_count = case_count - (1 if i == 18 else 0)
        rows.append([i, rel_no, test_date.strftime("%Y-%m-%d"),
                     case_count, pass_count, case_count - pass_count,
                     "合格" if pass_count == case_count else "一部不合格→修正後再テスト合格"])

    write_raw_csv(
        BASE_CM / "ITGC-CM-002_25件対応_RAW_UATテスト結果ログ.csv",
        ["# UAT Test Management System",
         "# Export:   2026-02-18 13:00:00 JST"],
        "サンプル№,REL番号,UAT実施日,総ケース数,合格数,不合格数,結果",
        rows,
        footer_lines=["# Records: 25"]
    )
    print("Created: ITGC-CM-002_25件対応_RAW_UATテスト結果ログ.csv")


# ============================================================
# ITGC-CM-003 本番移送履歴 (25件RAW)
# ============================================================
def gen_itgc_cm_003():
    random.seed(13001)
    rows = []
    for i in range(1, 26):
        rel_no = f"REL-2025-{i * 2:03d}"
        mv_date = date(2025, random.randint(4, 12), random.randint(1, 28))
        mv_ts = datetime.combine(mv_date, datetime.min.time()) + timedelta(
            hours=random.randint(2, 5))
        tr_no = f"XXXK{random.randint(900000, 999999)}"
        rows.append([mv_ts.strftime("%Y-%m-%d %H:%M:%S"), i, tr_no,
                     rel_no, "IT003 加藤 洋子", "DEV", "PRD",
                     "ABAP + Function Module", "成功"])

    write_raw_csv(
        BASE_CM / "ITGC-CM-003_25件対応_RAW_SAP_STMS_本番移送履歴.csv",
        ["# SAP S/4HANA - Transaction STMS",
         "# Report:   Transport Management System History",
         "# Export:   2026-02-18 14:00:00 JST"],
        "タイムスタンプ,サンプル№,TR番号,REL番号,移送者,移送元,移送先,対象オブジェクト,結果",
        rows,
        footer_lines=["# Records: 25"]
    )
    print("Created: ITGC-CM-003_25件対応_RAW_SAP_STMS_本番移送履歴.csv")


# ============================================================
# ITGC-OM-001 バックアップログ (日次RAW 25件)
# ============================================================
def gen_itgc_om_001():
    random.seed(14001)
    # 25日分のバックアップ実行ログ
    dates = generate_systematic_samples(25, seed=14001)
    rows = []
    for i, d in enumerate(dates, 1):
        start = datetime.combine(d, datetime.min.time()) + timedelta(
            hours=1, minutes=random.randint(0, 5))
        dur_min = random.randint(75, 130)
        end = start + timedelta(minutes=dur_min)
        size_gb = round(random.uniform(1800, 2100), 1)
        storage = "TAPE+S3" if d.weekday() == 6 else "TAPE"
        rows.append([start.strftime("%Y-%m-%d %H:%M:%S"),
                     end.strftime("%Y-%m-%d %H:%M:%S"),
                     i, "FULL_BACKUP", size_gb, storage,
                     f"BACKUP_JOB_{d.strftime('%Y%m%d')}",
                     "SUCCESS", "IT002 吉田 雅彦"])

    write_raw_csv(
        BASE_OM / "ITGC-OM-001_25件対応_RAW_SAP_DB13_バックアップログ.csv",
        ["# SAP S/4HANA - Transaction DB13 / Database Backup Log",
         "# Export:   2026-02-19 08:00:00 JST"],
        "開始タイムスタンプ,終了タイムスタンプ,サンプル№,バックアップ種別,サイズ(GB),保管先,ジョブ番号,結果,監視担当",
        rows,
        footer_lines=["# Records: 25"]
    )
    print("Created: ITGC-OM-001_25件対応_RAW_SAP_DB13_バックアップログ.csv")


# ============================================================
# ITGC-OM-002 障害管理 (18件全数RAW)
# ============================================================
def gen_itgc_om_002():
    random.seed(15001)
    severity = ["HIGH", "MEDIUM", "LOW"]
    incidents_desc = [
        "夜間バッチ遅延", "ユーザ接続エラー", "WMS-SAP連携停止",
        "印刷キュー詰まり", "HANA DBメモリ不足警告", "ネットワーク一時切断",
        "ログインタイムアウト", "レポート出力失敗"
    ]
    rows = []
    for i in range(1, 19):
        month = random.choice([4, 5, 6, 7, 8, 9, 10, 11])
        day = random.randint(1, 28)
        incident_ts = datetime(2025, month, day, random.randint(0, 23),
                               random.randint(0, 59))
        sev = random.choices(severity, weights=[1, 5, 12])[0]
        dur_min = {"HIGH": 180, "MEDIUM": 60, "LOW": 20}[sev]
        resolve_ts = incident_ts + timedelta(minutes=dur_min + random.randint(-20, 20))
        desc = random.choice(incidents_desc)
        rows.append([i, incident_ts.strftime("%Y-%m-%d %H:%M:%S"),
                     resolve_ts.strftime("%Y-%m-%d %H:%M:%S"),
                     sev, desc, "IT002 吉田 雅彦",
                     "CLOSED", "再発防止策適用済"])

    write_raw_csv(
        BASE_OM / "ITGC-OM-002_全18件RAW_監視ツール障害検知ログ.csv",
        ["# Zabbix-compatible Monitoring Tool / Incident Log",
         "# Period:   FY2025 (exhaustive 18 incidents)",
         "# Export:   2026-02-19 09:00:00 JST"],
        "№,発生タイムスタンプ,解消タイムスタンプ,重大度,事象概要,対応者,ステータス,備考",
        rows,
        footer_lines=["# Records: 18 (exhaustive)"]
    )
    print("Created: ITGC-OM-002_全18件RAW_監視ツール障害検知ログ.csv")


# ============================================================
# ITGC-EM-001 委託先管理 (SOC1+委託先一覧、RAWとしてSOC1原本のみ残す)
# ============================================================
def gen_itgc_em_001():
    # SOC1レポートは外部委託先から受領する業務原本として再生成
    pdf = JPPDF()
    pdf.add_page()
    pdf.h1("SOC1 Type II レポート（抜粋）")
    pdf.set_font("YuGoth", "", 10)
    pdf.cell(0, 5, "Report Period: 2024-04-01 to 2025-03-31", align="R",
             new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 5, "Service Organization: 外部委託先SIer-A", align="R",
             new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 5, "Issued: 2025-05-20", align="R",
             new_x="LMARGIN", new_y="NEXT")
    pdf.ln(5)

    pdf.h2("Management's Assertion")
    pdf.body("Management of 外部委託先SIer-A asserts that the description fairly presents "
             "the service organization's system, and the controls were suitably designed "
             "and operating effectively throughout the reporting period.")
    pdf.ln(3)

    pdf.h2("Independent Service Auditor's Report")
    pdf.body("In our opinion, the description fairly presents the system that was designed "
             "and implemented throughout the specified period. "
             "The controls were suitably designed and operated effectively to provide "
             "reasonable assurance of the control objectives.")
    pdf.ln(5)

    pdf.h2("Control Objectives and Related Controls")
    pdf.table_header(["Control Objective", "Tested Controls", "Result"],
                     [70, 60, 40])
    pdf.table_row(["Logical Access Management", "User provisioning, access review", "Effective"],
                  [70, 60, 40])
    pdf.table_row(["Change Management", "Change approval, testing", "Effective"],
                  [70, 60, 40], fill=True)
    pdf.table_row(["Backup & Recovery", "Daily backup, quarterly DR test", "Effective"],
                  [70, 60, 40])
    pdf.table_row(["Incident Management", "Timely notification, escalation",
                   "Improvement Required"], [70, 60, 40], fill=True)
    pdf.table_row(["Physical Security", "Data center access", "Effective"],
                  [70, 60, 40])
    pdf.ln(5)

    pdf.h2("Complementary User Entity Controls (CUECs)")
    pdf.body("Service organization assumes that user entities will implement complementary "
             "controls, including periodic review of service organization reports, "
             "monitoring of service level agreements, and verification of data integrity.")

    pdf.output(str(BASE_EM / "ITGC-EM-001_RAW_SOC1レポート_SIerA_FY2024.pdf"))
    print("Created: ITGC-EM-001_RAW_SOC1レポート_SIerA_FY2024.pdf")

    # B社のSOC1も
    pdf = JPPDF()
    pdf.add_page()
    pdf.h1("SOC1 Type II Report (Excerpt)")
    pdf.set_font("YuGoth", "", 10)
    pdf.cell(0, 5, "Report Period: 2024-04-01 to 2025-03-31", align="R",
             new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 5, "Service Organization: 外部委託先B社", align="R",
             new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 5, "Issued: 2025-06-15", align="R",
             new_x="LMARGIN", new_y="NEXT")
    pdf.ln(5)
    pdf.h2("Management's Assertion")
    pdf.body("Management of 外部委託先B社 asserts that controls were suitably designed and operated "
             "throughout the reporting period.")
    pdf.h2("Control Objectives")
    pdf.table_header(["Control Objective", "Result"], [100, 60])
    pdf.table_row(["IT Infrastructure Operations", "Effective"], [100, 60])
    pdf.table_row(["Server/Network Maintenance", "Effective"], [100, 60], fill=True)
    pdf.table_row(["Security Patching", "Effective"], [100, 60])
    pdf.table_row(["Monitoring & Alerting", "Effective"], [100, 60], fill=True)

    pdf.output(str(BASE_EM / "ITGC-EM-001_RAW_SOC1レポート_B社_FY2024.pdf"))
    print("Created: ITGC-EM-001_RAW_SOC1レポート_B社_FY2024.pdf")


# ============================================================
# FCRP-001 月次決算締めジョブログ (全12回RAW)
# ============================================================
def gen_fcrp_001():
    rows = []
    for month_offset in range(12):
        y = 2025 if month_offset < 9 else 2026
        m = 4 + month_offset if month_offset < 9 else month_offset - 8
        last_day = calendar.monthrange(y, m)[1]
        close_date = date(y, m, last_day)
        exec_ts = datetime.combine(close_date + timedelta(days=5),
                                   datetime.min.time()) + timedelta(hours=18)
        rows.append([f"{y}-{m:02d}", close_date.strftime("%Y-%m-%d"),
                     exec_ts.strftime("%Y-%m-%d %H:%M:%S"),
                     "PERIOD_CLOSE_COMPLETE",
                     "SAP_BATCH",
                     f"FCLOSE_{y}{m:02d}",
                     "完了"])

    write_raw_csv(
        BASE_FCRP / "FCRP-001_全12ヶ月RAW_SAP_FB50_月次決算ジョブログ.csv",
        ["# SAP FI - Period Close Execution Log",
         "# Period:   FY2025 (12 months exhaustive)",
         "# Export:   2026-04-10 09:00:00 JST"],
        "対象月,月末日,締め実行タイムスタンプ,ステータス,実行ユーザ,ジョブID,結果",
        rows,
        footer_lines=["# Records: 12 (exhaustive)"]
    )
    print("Created: FCRP-001_全12ヶ月RAW_SAP_FB50_月次決算ジョブログ.csv")


# ============================================================
# FCRP-002 連結パッケージ受信ログ (全4件四半期RAW)
# ============================================================
def gen_fcrp_002():
    random.seed(16001)
    rows = []
    subs = [
        ("TP-TB", "テクノプレシジョン東北"),
        ("TP-LOG", "TP物流サービス"),
        ("TPT", "TechnoPrecision Thailand"),
        ("TPTR", "TPトレーディング"),
    ]
    for q_idx, base in enumerate([date(2025, 6, 30), date(2025, 9, 30),
                                   date(2025, 12, 31), date(2026, 3, 31)], 1):
        for sub_code, sub_name in subs:
            upload_ts = datetime.combine(base + timedelta(days=8),
                                          datetime.min.time()) + timedelta(hours=random.randint(9, 17))
            err_count = random.choices([0, 1, 2], weights=[60, 25, 15])[0]
            rows.append([f"Q{q_idx}", base.strftime("%Y-%m-%d"),
                         upload_ts.strftime("%Y-%m-%d %H:%M:%S"),
                         sub_code, sub_name, err_count,
                         "PASS" if err_count == 0 else "PASS_AFTER_CORRECTION"])

    write_raw_csv(
        BASE_FCRP / "FCRP-002_全4四半期RAW_連結システムS05_パッケージ受信ログ.csv",
        ["# Consolidation System S05 - Package Upload Log",
         "# Period:   FY2025 Q1-Q4 (exhaustive)",
         "# Export:   2026-04-10 10:00:00 JST"],
        "四半期,基準日,アップロードタイムスタンプ,子会社コード,子会社名,バリデーションエラー数,結果",
        rows,
        footer_lines=["# Records: 16 (4 quarters x 4 subsidiaries)"]
    )
    print("Created: FCRP-002_全4四半期RAW_連結システムS05_パッケージ受信ログ.csv")


# ============================================================
# FCRP-003 SAP売掛金&計算シート (全4四半期RAW)
# ============================================================
def gen_fcrp_003():
    random.seed(17001)
    rows = []
    for q_label, base in [("Q1", date(2025, 6, 30)), ("Q2", date(2025, 9, 30)),
                           ("Q3", date(2025, 12, 31)), ("Q4", date(2026, 3, 31))]:
        # 一般債権
        total_ar = random.randint(3_000_000_000, 4_000_000_000)
        general_reserve_rate = 0.18 / 100
        general_reserve = int(total_ar * general_reserve_rate)
        rows.append([q_label, base.strftime("%Y-%m-%d"), "一般債権",
                     total_ar, f"{general_reserve_rate * 100:.2f}%",
                     general_reserve, "AUTO_AGING"])
        # 個別評価（3社）
        individual_cases = [("C-10007", "サンプル顧客G社"),
                            ("C-10017", "サンプル顧客N社"),
                            ("C-10023", "サンプル顧客R社")]
        for cid, cname in individual_cases:
            debt = random.randint(3_000_000, 7_000_000)
            recovery_rate = random.choice([0.4, 0.5, 0.6, 0.8])
            reserve = int(debt * (1 - recovery_rate))
            rows.append([q_label, base.strftime("%Y-%m-%d"),
                         f"個別({cid} {cname})",
                         debt, f"{(1 - recovery_rate) * 100:.0f}%",
                         reserve, "MANUAL_EVAL"])

    write_raw_csv(
        BASE_FCRP / "FCRP-003_全4四半期RAW_SAP_FB10N_貸倒引当金算定データ.csv",
        ["# SAP FI - AR Aging (FB10N) + Impairment Estimation Input",
         "# Period:   FY2025 Q1-Q4 (exhaustive)",
         "# Export:   2026-04-10 11:00:00 JST"],
        "四半期,基準日,区分,対象金額(円),引当率,引当額(円),評価方式",
        rows,
        footer_lines=["# Records: 16 (4 quarters x 4 entries)"]
    )
    print("Created: FCRP-003_全4四半期RAW_SAP_FB10N_貸倒引当金算定データ.csv")


# ============================================================
# FCRP-004 連結仕訳一覧RAW
# ============================================================
def gen_fcrp_004():
    random.seed(18001)
    rows = []
    entry_types = [
        ("投資と資本の相殺", "資本金", "関係会社株式", 300_000_000),
        ("内部取引消去", "売上高", "売上原価", 1_820_000_000),
        ("内部取引消去", "売上高", "売上原価", 650_000_000),
        ("内部取引消去", "売掛金", "買掛金", 324_500_000),
        ("少数株主損益", "少数株主損益", "利益剰余金", 8_520_000),
        ("内部利益消去(在庫)", "売上原価", "棚卸資産", 42_800_000),
    ]
    for q_idx, base in enumerate([date(2025, 6, 30), date(2025, 9, 30),
                                   date(2025, 12, 31), date(2026, 3, 31)], 1):
        for j, (etype, dr, cr, amt) in enumerate(entry_types, 1):
            rows.append([f"CNS-Q{q_idx}-{j:03d}", f"Q{q_idx}",
                         base.strftime("%Y-%m-%d"),
                         etype, dr, cr, amt,
                         "AUTO" if etype != "内部利益消去(在庫)" else "MANUAL",
                         "PENDING_REVIEW" if q_idx == 4 else "REVIEWED"])

    write_raw_csv(
        BASE_FCRP / "FCRP-004_全4四半期RAW_連結システムS05_連結仕訳一覧.csv",
        ["# Consolidation System S05 - Consolidation Entries",
         "# Period:   FY2025 Q1-Q4 (exhaustive)",
         "# Export:   2026-04-10 12:00:00 JST"],
        "仕訳№,四半期,基準日,区分,借方科目,貸方科目,金額,起票区分,ステータス",
        rows,
        footer_lines=["# Records: 24 (4 quarters x 6 entries)"]
    )
    print("Created: FCRP-004_全4四半期RAW_連結システムS05_連結仕訳一覧.csv")


# ============================================================
# FCRP-005 開示書類XBRL検証ログ (全4四半期RAW)
# ============================================================
def gen_fcrp_005():
    rows = []
    for q_idx, (y, m) in enumerate([(2025, 8), (2025, 11), (2026, 2), (2026, 5)], 1):
        submit_date = date(y, m, 15 if m != 5 else 10)
        rows.append([f"Q{q_idx}", submit_date.strftime("%Y-%m-%d"),
                     f"EDINET-XBRL-{y}{m:02d}",
                     "PASS", "0 errors / 0 warnings",
                     "経営企画部", submit_date.strftime("%Y-%m-%d")])

    write_raw_csv(
        BASE_FCRP / "FCRP-005_全4四半期RAW_開示システムS06_XBRL検証ログ.csv",
        ["# Disclosure System S06 - XBRL Validation Log",
         "# Period:   FY2025 Q1-Q4 (exhaustive)",
         "# Export:   2026-05-15 10:00:00 JST"],
        "四半期,提出日,書類ID,検証結果,エラー/警告件数,提出部署,EDINET登録日",
        rows,
        footer_lines=["# Records: 4 (exhaustive)"]
    )
    print("Created: FCRP-005_全4四半期RAW_開示システムS06_XBRL検証ログ.csv")


if __name__ == "__main__":
    # PLC-I
    gen_plc_i_001()
    gen_plc_i_002()
    gen_plc_i_004()
    gen_plc_i_005()
    gen_plc_i_007()

    # ITGC
    gen_itgc_ac_001()
    gen_itgc_cm_001()
    gen_itgc_cm_002()
    gen_itgc_cm_003()
    gen_itgc_om_001()
    gen_itgc_om_002()
    gen_itgc_em_001()

    # FCRP
    gen_fcrp_001()
    gen_fcrp_002()
    gen_fcrp_003()
    gen_fcrp_004()
    gen_fcrp_005()
    print("\nAll remaining evidence expansion completed.")
