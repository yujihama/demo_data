"""
Phase 2: PLC-S 拡張

対象統制：
- PLC-S-001 受注・与信承認（25件の受注注文書PDF + RAW）
- PLC-S-003 請求書発行（12ヶ月のバッチログ + 12件の請求書PDF + RAW）
- PLC-S-004 入金消込（25件のSAP入金消込RAW + FBデータRAW）
- PLC-S-005 売掛金年齢分析（12ヶ月のSAP FB10N出力RAW）
- PLC-S-006 期末カットオフ（41件全数のSAP出荷・売上明細RAW）
- PLC-S-007 価格マスタ承認（25件の価格変更稟議PDF + マスタ変更履歴RAW）
"""
import random
import sys
from pathlib import Path
from datetime import date, datetime, timedelta

sys.path.insert(0, str(Path(__file__).parent))
from pdf_util import JPPDF
from sample_gen_util import (
    CUSTOMERS, PRODUCTS,
    generate_systematic_samples, create_sample_list_excel, write_raw_csv
)

BASE = Path(r"C:\Users\nyham\work\demo_data\4.evidence\PLC-S")


# ============================================================
# PLC-S-001: 25件の受注サンプル
# ============================================================
def gen_plc_s_001():
    random.seed(1001)
    dates = generate_systematic_samples(25, seed=1001)
    samples = []
    for i, d in enumerate(dates, 1):
        cid, cname, credit, industry = random.choice(CUSTOMERS)
        pcode, pname, cost, price = random.choice(PRODUCTS)
        qty = random.choice([50, 100, 200, 300, 500, 800, 1000, 1500])
        amount = qty * price
        samples.append({
            "no": i, "ord_no": f"ORD-2025-{100 + i * 130:04d}",
            "date": d, "cid": cid, "cname": cname,
            "pcode": pcode, "pname": pname, "qty": qty,
            "unit_price": price, "amount": amount,
            "credit": credit, "rep": random.choice(
                ["斎藤 次郎 (SLS002)", "藤田 修 (SLS003)",
                 "松本 香織 (SLS004)", "井上 大輔 (SLS005)"]),
            "delivery": d + timedelta(days=random.randint(14, 45)),
        })

    # サンプルリスト
    create_sample_list_excel(
        BASE / "PLC-S-001_監査対象25件サンプルリスト.xlsx",
        "【PLC-S-001】監査対象25件サンプルリスト（受注・与信承認）",
        "（RAWデータをナビゲートするための取引リスト）",
        [
            ("母集団", "FY2025 全受注 3,247件（SAP VA05）"),
            ("抽出方法", "系統抽出 / 間隔130件 / 開始位置57"),
            ("抽出日時", "2026-02-10 10:30 JST"),
            ("関連RAWデータ", "PLC-S-001_25件対応_RAW_*.csv / 各サンプルの注文書PDF"),
        ],
        ["サンプル\n№", "受注番号", "受注日", "顧客\nコード", "顧客名",
         "製品コード", "数量", "受注金額\n(円)", "営業担当", "希望納期"],
        [[s["no"], s["ord_no"], s["date"], s["cid"], s["cname"],
          s["pcode"], s["qty"], s["amount"], s["rep"], s["delivery"]]
         for s in samples],
        col_widths=[6, 16, 11, 10, 16, 12, 10, 14, 22, 11],
        col_center=(0, 1, 3, 5),
        col_right=(6, 7),
        col_date=(2, 9),
    )
    print(f"Created: PLC-S-001_監査対象25件サンプルリスト.xlsx")

    # 25件対応 SAP VA05 RAW（個別詳細）
    rows = []
    for s in samples:
        credit_result = "○ 限度内"
        approval = "自動承認"
        wf_no = ""
        # 意図的に3件を与信超過ケースに
        if s["no"] in (7, 14, 22):
            credit_result = "超過（保留）"
            approval = "営業本部長承認"
            wf_no = f"WF-2025-{3000 + s['no'] * 11:05d}"
        rows.append([s["no"], s["ord_no"], s["date"].strftime("%Y-%m-%d"),
                     "通常受注", s["cid"], s["cname"], s["rep"].split(" ")[1].rstrip("()"),
                     s["amount"], s["credit"], credit_result, approval, wf_no])

    write_raw_csv(
        BASE / "PLC-S-001_25件対応_RAW_SAP_VA05_受注詳細.csv",
        ["# SAP S/4HANA - Transaction VA05",
         "# Report:     Sales Orders Detail",
         "# Export:     2026-02-10 11:20:15 JST",
         "# Filter:     Sample 25 orders per audit request IA-REQ-2026-001",
         "# Records:    25"],
        "サンプル№,受注番号,受注日,受注タイプ,顧客コード,顧客名,営業担当,受注金額,与信限度額,与信チェック結果,承認区分,ワークフロー番号",
        rows,
        footer_lines=["# End of export"]
    )
    print(f"Created: PLC-S-001_25件対応_RAW_SAP_VA05_受注詳細.csv")

    # 与信チェックログRAW (25件)
    log_rows = []
    for s in samples:
        ts = datetime.combine(s["date"], datetime.min.time()) + timedelta(
            hours=random.randint(9, 17), minutes=random.randint(0, 59))
        if s["no"] in (7, 14, 22):
            judge = "HOLD"
            msg = "CREDIT_LIMIT_EXCEEDED"
        else:
            judge = "PASS"
            msg = "WITHIN_LIMIT"
        log_rows.append([ts.strftime("%Y-%m-%d %H:%M:%S"),
                         s["ord_no"], s["cid"], s["amount"],
                         int(s["credit"] * 0.6), s["credit"], judge, msg])

    write_raw_csv(
        BASE / "PLC-S-001_25件対応_RAW_SAP与信チェックログ.csv",
        ["# SAP Credit Management Log",
         "# Module:      FD32 (Credit master) + Credit Check Automation",
         "# Export:      2026-02-10 11:22:05 JST",
         "# Filter:      25 orders under audit IA-REQ-2026-001"],
        "タイムスタンプ,受注番号,顧客コード,受注金額,既存売掛金,与信限度額,判定,理由コード",
        log_rows,
        footer_lines=["# Records: 25"]
    )
    print(f"Created: PLC-S-001_25件対応_RAW_SAP与信チェックログ.csv")

    # ワークフロー承認履歴（与信超過3件のみ）
    wf_rows = []
    for s in samples:
        if s["no"] in (7, 14, 22):
            wf = f"WF-2025-{3000 + s['no'] * 11:05d}"
            start = datetime.combine(s["date"], datetime.min.time()) + timedelta(hours=10)
            wf_rows.append([start.strftime("%Y-%m-%d %H:%M:%S"),
                            wf, s["no"], s["ord_no"], s["rep"].split(" ")[1].rstrip("()"),
                            "起票", ""])
            wf_rows.append([(start + timedelta(hours=3)).strftime("%Y-%m-%d %H:%M:%S"),
                            wf, s["no"], s["ord_no"], "斎藤 次郎 (SLS002)",
                            "承認", "営業課長承認"])
            wf_rows.append([(start + timedelta(hours=6)).strftime("%Y-%m-%d %H:%M:%S"),
                            wf, s["no"], s["ord_no"], "田中 太郎 (SLS001)",
                            "承認", "営業本部長承認"])

    write_raw_csv(
        BASE / "PLC-S-001_25件対応_RAW_ワークフロー承認履歴.csv",
        ["# SAP Business Workflow - Approval History Log",
         "# Export:   2026-02-10 11:25:30 JST",
         "# Filter:   Credit-exceeded orders in 25 audit samples",
         "# Records:  3 workflows x 3 stages"],
        "タイムスタンプ,ワークフロー番号,サンプル№,受注番号,アクター,アクション,コメント",
        wf_rows,
        footer_lines=["# End of log"]
    )
    print(f"Created: PLC-S-001_25件対応_RAW_ワークフロー承認履歴.csv")

    # 25件の注文書PDF
    _gen_order_pdfs(samples)


def _gen_order_pdfs(samples):
    for s in samples:
        pdf = JPPDF()
        pdf.add_page()
        pdf.set_font("YuGoth", "B", 20)
        pdf.cell(0, 12, "注 文 書", align="C", new_x="LMARGIN", new_y="NEXT")
        pdf.ln(3)
        pdf.set_font("YuGoth", "", 10)
        pdf.cell(0, 5, f"(顧客側) 注文書番号: CUST-PO-2025-{s['no'] * 73 + 1000:05d}",
                 align="R", new_x="LMARGIN", new_y="NEXT")
        pdf.cell(0, 5, f"発行日: {s['date'].strftime('%Y年%m月%d日')}",
                 align="R", new_x="LMARGIN", new_y="NEXT")
        pdf.ln(5)

        pdf.set_font("YuGoth", "B", 12)
        pdf.cell(0, 7, "デモA株式会社 御中", new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("YuGoth", "", 10)
        pdf.cell(0, 5, "営業本部 担当者殿", new_x="LMARGIN", new_y="NEXT")
        pdf.ln(3)

        pdf.set_font("YuGoth", "B", 10)
        pdf.set_x(110)
        pdf.cell(90, 6, f"発注元: {s['cname']}", new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("YuGoth", "", 9)
        pdf.set_x(110)
        pdf.cell(90, 5, f"顧客コード: {s['cid']}", new_x="LMARGIN", new_y="NEXT")
        pdf.ln(8)

        pdf.set_font("YuGoth", "", 10)
        pdf.multi_cell(0, 5, "下記のとおり発注致します。ご確認のうえ、納期に間に合うよう手配をお願い致します。")
        pdf.ln(3)

        pdf.table_header(["品目コード", "品名", "数量", "単価(円)", "金額(円)"],
                         [30, 80, 20, 30, 30])
        pdf.table_row([s['pcode'], s['pname'], f"{s['qty']:,}",
                       f"{s['unit_price']:,}", f"{s['amount']:,}"],
                      [30, 80, 20, 30, 30])

        subtotal = s["amount"]
        tax = int(subtotal * 0.1)
        total = subtotal + tax
        pdf.set_font("YuGoth", "B", 10)
        pdf.cell(130, 7, "小計", border=1, align="R")
        pdf.cell(60, 7, f"¥ {subtotal:,}", border=1, align="R",
                 new_x="LMARGIN", new_y="NEXT")
        pdf.cell(130, 7, "消費税 (10%)", border=1, align="R")
        pdf.cell(60, 7, f"¥ {tax:,}", border=1, align="R",
                 new_x="LMARGIN", new_y="NEXT")
        pdf.set_fill_color(255, 242, 204)
        pdf.cell(130, 8, "合計", border=1, align="R", fill=True)
        pdf.cell(60, 8, f"¥ {total:,}", border=1, align="R", fill=True,
                 new_x="LMARGIN", new_y="NEXT")
        pdf.set_fill_color(255, 255, 255)
        pdf.ln(5)

        pdf.h3("■ 納入条件")
        pdf.set_font("YuGoth", "", 9)
        pdf.kv("納期", s["delivery"].strftime("%Y年%m月%d日"))
        pdf.kv("納入場所", "貴社指定倉庫")
        pdf.kv("支払条件", "月末締 翌月末払")

        out = BASE / f"PLC-S-001_25件対応_注文書_サンプル{s['no']:02d}_{s['ord_no']}.pdf"
        pdf.output(str(out))

    print(f"Created: 25 order PDFs for PLC-S-001")


# ============================================================
# PLC-S-003: 12ヶ月のバッチログ + 12件の請求書PDF
# ============================================================
def gen_plc_s_003():
    random.seed(3003)
    # 12ヶ月分のバッチログ
    for month_offset in range(12):
        y = 2025 if month_offset < 9 else 2026
        m = 4 + month_offset if month_offset < 9 else month_offset - 8
        # Last day of month
        import calendar
        last_day = calendar.monthrange(y, m)[1]
        run_date = date(y, m, last_day)

        invoice_count = random.randint(140, 175)
        amount_total = random.randint(1_100_000_000, 1_500_000_000)

        content = f"""================================================================
 SAP S/4HANA / SD Invoice Batch Execution Log
================================================================
Batch Job:     ZSD_MONTHLY_INVOICE
Job ID:        JOB_{y}{m:02d}{last_day}_2358_01
User:          SAP_BATCH (system)
Start:         {run_date.strftime('%Y-%m-%d')} 23:58:01
End:           {(run_date + timedelta(days=1)).strftime('%Y-%m-%d')} 00:{12 + month_offset % 5}:{random.randint(10, 59):02d}
Return code:   0 (SUCCESS)

----------------------------------------------------------------
 Processing Summary
----------------------------------------------------------------
Target customers:       20
Target billing items:   {random.randint(260, 310)}
Invoices generated:     {invoice_count}
Total amount (excl.):   JPY {amount_total:,}
Total amount (incl.):   JPY {int(amount_total * 1.1):,}
Reconciliation vs SJ:   {invoice_count}/{invoice_count} matched, delta = 0

----------------------------------------------------------------
 Dispatch
----------------------------------------------------------------
PDF generation:         {invoice_count} files in {random.randint(5, 10)} min
Email dispatched:       {int(invoice_count * 0.83)} customers
Paper print:            {int(invoice_count * 0.17)} customers

----------------------------------------------------------------
 Next scheduled:        {date(y, m + 1 if m < 12 else 1, calendar.monthrange(y + (1 if m == 12 else 0), m + 1 if m < 12 else 1)[1]).strftime('%Y-%m-%d')} 23:58:00
----------------------------------------------------------------
"""
        path = BASE / f"PLC-S-003_25件対応_RAW_SAP請求書バッチログ_{y}{m:02d}.txt"
        path.write_text(content, encoding="utf-8")
    print(f"Created: 12 monthly batch logs for PLC-S-003")

    # 12ヶ月のSAPからの請求書一覧CSV（月次バッチ結果RAW）
    random.seed(3333)
    rows = []
    for month_offset in range(12):
        y = 2025 if month_offset < 9 else 2026
        m = 4 + month_offset if month_offset < 9 else month_offset - 8
        import calendar
        last_day = calendar.monthrange(y, m)[1]
        run_date = date(y, m, last_day)
        for i in range(1, 6):  # 各月5件抜粋
            cid, cname, _, _ = random.choice(CUSTOMERS)
            amt = random.randint(1_000_000, 35_000_000) // 1000 * 1000
            tax = int(amt * 0.1)
            inv_no = f"INV-{y}{m:02d}-{i * 7:04d}"
            next_y = y + (1 if m == 12 else 0)
            next_m = 1 if m == 12 else m + 1
            due_last = calendar.monthrange(next_y, next_m)[1]
            due = date(next_y, next_m, min(25, due_last))
            rows.append([inv_no, run_date.strftime("%Y-%m-%d"), cid, cname,
                         amt, tax, amt + tax, due.strftime("%Y-%m-%d"),
                         "PDF_EMAIL" if i % 3 != 0 else "PAPER_MAIL"])

    write_raw_csv(
        BASE / "PLC-S-003_25件対応_RAW_SAP請求書一覧_FY2025.csv",
        ["# SAP S/4HANA - Invoice Register",
         "# Transaction: VF05",
         "# Export:      2026-04-10 08:00:00 JST",
         "# Filter:      FY2025 / 代表5件/月 × 12ヶ月 = 60件"],
        "請求書番号,発行日,顧客コード,顧客名,請求金額(税抜),消費税,税込金額,支払期日,送付方法",
        rows,
        footer_lines=["# Records: 60 (12 months representative samples)"]
    )
    print(f"Created: PLC-S-003_25件対応_RAW_SAP請求書一覧_FY2025.csv")

    # 代表的な請求書PDF 12枚（各月1枚）
    random.seed(33333)
    for month_offset in range(12):
        y = 2025 if month_offset < 9 else 2026
        m = 4 + month_offset if month_offset < 9 else month_offset - 8
        import calendar
        last_day = calendar.monthrange(y, m)[1]
        run_date = date(y, m, last_day)
        cid, cname, _, _ = random.choice(CUSTOMERS)
        pcode, pname, _, price = random.choice(PRODUCTS)
        qty = random.randint(100, 1500)
        amount = qty * price

        pdf = JPPDF()
        pdf.add_page()
        pdf.set_font("YuGoth", "B", 22)
        pdf.cell(0, 14, "請 求 書", align="C", new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("YuGoth", "", 10)
        inv_no = f"INV-{y}{m:02d}-0001"
        pdf.cell(0, 5, f"請求書番号: {inv_no}", align="R", new_x="LMARGIN", new_y="NEXT")
        pdf.cell(0, 5, f"請求日: {run_date.strftime('%Y年%m月%d日')}",
                 align="R", new_x="LMARGIN", new_y="NEXT")
        pdf.ln(4)

        pdf.set_font("YuGoth", "B", 12)
        pdf.cell(0, 7, f"{cname} 御中", new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("YuGoth", "", 9)
        pdf.cell(0, 5, f"顧客コード: {cid}", new_x="LMARGIN", new_y="NEXT")
        pdf.ln(5)

        pdf.set_x(110)
        pdf.set_font("YuGoth", "B", 11)
        pdf.cell(90, 6, "デモA株式会社", new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("YuGoth", "", 9)
        pdf.set_x(110)
        pdf.cell(90, 5, "〒XXX-XXXX 神奈川県横浜市港北区", new_x="LMARGIN", new_y="NEXT")
        pdf.set_x(110)
        pdf.cell(90, 5, "登録番号: T1234567890123", new_x="LMARGIN", new_y="NEXT")
        pdf.ln(10)

        tax = int(amount * 0.1)
        total = amount + tax
        pdf.set_font("YuGoth", "B", 14)
        pdf.set_fill_color(240, 245, 255)
        pdf.cell(60, 14, "ご請求金額", border=1, align="C", fill=True)
        pdf.cell(130, 14, f"¥ {total:,} -", border=1, align="R", fill=True,
                 new_x="LMARGIN", new_y="NEXT")
        pdf.set_fill_color(255, 255, 255)
        pdf.ln(6)

        pdf.table_header(["品目コード", "品名", "数量", "単価", "金額"],
                         [30, 80, 20, 30, 30])
        pdf.table_row([pcode, pname, f"{qty:,}", f"{price:,}", f"{amount:,}"],
                      [30, 80, 20, 30, 30])

        pdf.set_font("YuGoth", "B", 10)
        pdf.cell(130, 7, "小計", border=1, align="R")
        pdf.cell(60, 7, f"¥ {amount:,}", border=1, align="R",
                 new_x="LMARGIN", new_y="NEXT")
        pdf.cell(130, 7, "消費税 (10%)", border=1, align="R")
        pdf.cell(60, 7, f"¥ {tax:,}", border=1, align="R",
                 new_x="LMARGIN", new_y="NEXT")
        pdf.set_fill_color(255, 242, 204)
        pdf.cell(130, 8, "合計", border=1, align="R", fill=True)
        pdf.cell(60, 8, f"¥ {total:,}", border=1, align="R", fill=True,
                 new_x="LMARGIN", new_y="NEXT")
        pdf.set_fill_color(255, 255, 255)
        pdf.ln(6)

        pdf.h3("■ お支払について")
        pdf.set_font("YuGoth", "", 10)
        next_y = y + (1 if m == 12 else 0)
        next_m = 1 if m == 12 else m + 1
        due = date(next_y, next_m, min(25, calendar.monthrange(next_y, next_m)[1]))
        pdf.kv("お支払期日", due.strftime("%Y年%m月%d日"))
        pdf.kv("振込先", "A銀行 支店X 普通 1234567")
        pdf.kv("口座名義", "カ）デモA")

        out = BASE / f"PLC-S-003_25件対応_請求書サンプル_{y}{m:02d}_{inv_no}.pdf"
        pdf.output(str(out))
    print(f"Created: 12 invoice PDFs for PLC-S-003")


# ============================================================
# PLC-S-004: 25件の入金消込サンプル
# ============================================================
def gen_plc_s_004():
    random.seed(4004)
    dates = generate_systematic_samples(25, seed=4004)
    # サンプルリスト
    samples = []
    for i, d in enumerate(dates, 1):
        cid, cname, _, _ = random.choice(CUSTOMERS)
        amount = random.randint(2_000_000, 40_000_000) // 1000 * 1000
        samples.append({
            "no": i, "date": d, "cid": cid, "cname": cname,
            "amount": amount,
            "inv_no": f"INV-{d.year}{(d.month - 1) if d.month > 1 else 12:02d}-{random.randint(1, 150):04d}",
            "method": "AUTO" if i not in (5, 12, 24) else "MANUAL",
            "diff": -1_500 if i == 9 else 0,
        })

    create_sample_list_excel(
        BASE / "PLC-S-004_監査対象25件サンプルリスト.xlsx",
        "【PLC-S-004】監査対象25件サンプルリスト（入金消込）",
        "（RAWデータをナビゲートするための取引リスト）",
        [
            ("母集団", "FY2025 入金 2,847件（SAP F-28消込履歴）"),
            ("抽出方法", "系統抽出 / 間隔114件 / 開始位置43"),
            ("抽出日時", "2026-02-11 10:15 JST"),
            ("関連RAWデータ", "PLC-S-004_25件対応_RAW_*.csv"),
        ],
        ["サンプル\n№", "入金日", "顧客\nコード", "顧客名", "入金額(円)",
         "消込対象\n請求書", "消込方法", "差額(円)"],
        [[s["no"], s["date"], s["cid"], s["cname"], s["amount"],
          s["inv_no"], s["method"], s["diff"]] for s in samples],
        col_widths=[6, 12, 10, 20, 14, 16, 10, 12],
        col_center=(0, 2, 5, 6),
        col_right=(4, 7),
        col_date=(1,),
    )
    print("Created: PLC-S-004_監査対象25件サンプルリスト.xlsx")

    # SAP入金消込履歴 RAW (25件)
    rows = []
    for s in samples:
        ts = datetime.combine(s["date"], datetime.min.time()) + timedelta(
            hours=random.randint(9, 16), minutes=random.randint(0, 59))
        user = "SAP_AUTO_CLEAR" if s["method"] == "AUTO" else "石井 健 (ACC006)"
        rows.append([ts.strftime("%Y-%m-%d %H:%M:%S"), s["no"],
                     s["date"].strftime("%Y-%m-%d"), s["amount"], s["cid"],
                     s["inv_no"], s["amount"] + s["diff"], s["diff"],
                     s["method"], user])

    write_raw_csv(
        BASE / "PLC-S-004_25件対応_RAW_SAP入金消込履歴_F-28.csv",
        ["# SAP FI - Transaction F-28 / Customer Clearing History",
         "# Export:   2026-02-11 10:30:42 JST",
         "# Filter:   25 payment receipts under audit IA-REQ-2026-004",
         "# Records:  25"],
        "処理タイムスタンプ,サンプル№,入金日,入金額,顧客コード,消込対象請求書,消込金額,差額,消込方法,処理ユーザ",
        rows,
        footer_lines=["# End of export"]
    )
    print("Created: PLC-S-004_25件対応_RAW_SAP入金消込履歴_F-28.csv")

    # FB入金データRAW (25件)
    banks = ["A銀行 本店", "B銀行 支店X", "C銀行 本店", "D銀行 本店"]
    fb_rows = []
    for s in samples:
        ymd = s["date"].strftime("%Y%m%d")
        fb_rows.append([ymd, "1234567", s["amount"],
                        f"SAMPLE KOKYAKU {s['cid'][-2:]}",
                        s["inv_no"], random.choice(banks)])

    write_raw_csv(
        BASE / "PLC-S-004_25件対応_RAW_FB入金データ_25件抽出.csv",
        ["# 全銀フォーマット FB入金データ（25件抽出版）",
         "# Export:   2026-02-11 10:35:08 JST",
         "# Filter:   Bank transactions matching 25 sample receipts"],
        "入金日,口座番号,入金額,振込依頼人名,摘要,取扱銀行",
        fb_rows,
        footer_lines=["# Records: 25"]
    )
    print("Created: PLC-S-004_25件対応_RAW_FB入金データ_25件抽出.csv")


# ============================================================
# PLC-S-005: 12ヶ月のSAP売掛金年齢表RAW（FB10N出力）
# ============================================================
def gen_plc_s_005():
    random.seed(5005)
    # 12ヶ月分のFB10N出力
    for month_offset in range(12):
        y = 2025 if month_offset < 9 else 2026
        m = 4 + month_offset if month_offset < 9 else month_offset - 8
        import calendar
        last_day = calendar.monthrange(y, m)[1]
        base_date = date(y, m, last_day)

        rows = []
        for cid, cname, credit, _ in CUSTOMERS:
            total = random.randint(3_000_000, min(credit // 2, 200_000_000))
            normal = int(total * 0.7)
            d31 = int(total * 0.2)
            d61 = int(total * 0.06)
            d91 = int(total * 0.03)
            d120 = total - normal - d31 - d61 - d91
            # 一部は120日超を意図的に多く
            if cid in ("C-10007", "C-10017", "C-10023"):
                d120 = int(total * 0.15)
                normal = total - d31 - d61 - d91 - d120
            rows.append([cid, cname, total, normal, d31, d61, d91, d120])

        write_raw_csv(
            BASE / f"PLC-S-005_25件対応_RAW_SAP売掛金年齢表_FB10N_{y}{m:02d}.csv",
            [f"# SAP FI - Transaction FB10N / Customer AR Aging Report",
             f"# Report basis date: {base_date.strftime('%Y-%m-%d')}",
             f"# Aging buckets:     0-30 / 31-60 / 61-90 / 91-120 / over 120 days",
             f"# Export:            {(base_date + timedelta(days=5)).strftime('%Y-%m-%d %H:%M:%S')} JST"],
            "顧客コード,顧客名,残高合計(円),0-30日,31-60日,61-90日,91-120日,120日超",
            rows,
            footer_lines=[f"# Records: {len(rows)}"]
        )
    print("Created: 12 monthly FB10N aging reports for PLC-S-005")


# ============================================================
# PLC-S-006: 期末カットオフ 41件全数の出荷・売上明細
# ============================================================
def gen_plc_s_006():
    random.seed(6006)
    rows = []
    for i in range(1, 42):
        dayidx = random.choice([-6, -5, -4, -3, -2, -1, 0, 1])
        ship_date = date(2026, 3, 31) + timedelta(days=dayidx)
        sale_date = ship_date + timedelta(days=random.choice([0, 0, 1]))
        cid, cname, _, _ = random.choice(CUSTOMERS)
        pcode, _, _, price = random.choice(PRODUCTS)
        qty = random.randint(50, 500)
        amount = qty * price
        ord_no = f"ORD-2026-{3000 + i:04d}"
        ship_no = f"SH-202603-{i:04d}"
        jv_no = f"JV-20260{3 if sale_date.month <= 3 else 4}-{i:04d}"
        rows.append([i, ord_no, ship_date.strftime("%Y-%m-%d"),
                     sale_date.strftime("%Y-%m-%d"), cid, cname,
                     pcode, qty, amount, ship_no, jv_no,
                     "FY2025" if sale_date.year < 2026 or
                     (sale_date.year == 2026 and sale_date.month <= 3) else "FY2026"])

    write_raw_csv(
        BASE / "PLC-S-006_RAW_SAP期末前後出荷売上明細_FY2025期末.csv",
        ["# SAP S/4HANA - Period-end Shipment & Sales Detail",
         "# Period:   2026-03-25 to 2026-04-01 (+/- 5 business days of fiscal year-end)",
         "# Source:   VA05 (orders) + VL06 (shipments) + FBL3N (sales journals)",
         "# Export:   2026-04-03 08:00:15 JST",
         "# Population: All 41 records in period (exhaustive)"],
        "№,受注番号,出荷日,売上計上日,顧客コード,顧客名,製品コード,数量,金額,出荷番号,売上仕訳番号,計上期",
        rows,
        footer_lines=["# Records: 41 (exhaustive)"]
    )
    print("Created: PLC-S-006_RAW_SAP期末前後出荷売上明細_FY2025期末.csv")


# ============================================================
# PLC-S-007: 25件の価格マスタ変更稟議PDF + マスタ変更履歴RAW
# ============================================================
def gen_plc_s_007():
    random.seed(7007)
    # 母集団36件から25件を系統抽出（簡略化のため直接25件を生成）
    samples = []
    for i in range(1, 26):
        pcode, pname, cost, price = random.choice(PRODUCTS)
        cid, cname, _, _ = random.choice(CUSTOMERS)
        old_price = price
        change_rate = random.choice([2.0, 2.5, 3.0, 3.5, 3.8, 4.2, -1.5, -2.3])
        new_price = int(old_price * (1 + change_rate / 100))
        # 変更日を分散
        month = random.choice([4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3])
        y = 2026 if month <= 3 else 2025
        day = random.randint(5, 25)
        samples.append({
            "no": i, "wf_no": f"W-2025-{1800 + i * 18:04d}",
            "date": date(y, month, day), "pcode": pcode, "pname": pname,
            "cid": cid, "cname": cname,
            "old_price": old_price, "new_price": new_price,
            "change_rate": change_rate,
        })

    create_sample_list_excel(
        BASE / "PLC-S-007_監査対象25件サンプルリスト.xlsx",
        "【PLC-S-007】監査対象25件サンプルリスト（価格マスタ承認）",
        "（RAWデータをナビゲートするための取引リスト）",
        [
            ("母集団", "FY2025 価格マスタ変更 36件（SAP VK12変更履歴）"),
            ("抽出方法", "系統抽出 / 間隔1件 / 25件"),
            ("抽出日時", "2026-02-14 14:00 JST"),
            ("関連RAWデータ", "PLC-S-007_25件対応_RAW_*.csv / 各サンプルの稟議書PDF"),
        ],
        ["サンプル\n№", "稟議番号", "変更日", "製品コード", "顧客\nコード",
         "旧単価", "新単価", "変更率"],
        [[s["no"], s["wf_no"], s["date"], s["pcode"], s["cid"],
          s["old_price"], s["new_price"], f"{s['change_rate']:+.1f}%"]
         for s in samples],
        col_widths=[6, 14, 11, 12, 10, 12, 12, 10],
        col_center=(0, 1, 3, 4, 7),
        col_right=(5, 6),
        col_date=(2,),
    )
    print("Created: PLC-S-007_監査対象25件サンプルリスト.xlsx")

    # VK12変更履歴RAW
    rows = []
    for s in samples:
        ts = datetime.combine(s["date"], datetime.min.time()) + timedelta(
            hours=random.randint(10, 16), minutes=random.randint(0, 59))
        rows.append([ts.strftime("%Y-%m-%d %H:%M:%S"), s["no"], s["wf_no"],
                     s["pcode"], s["cid"], "KONP-KBETR",
                     s["old_price"], s["new_price"],
                     f"{s['change_rate']:+.2f}%", "SLS004 松本 香織"])

    write_raw_csv(
        BASE / "PLC-S-007_25件対応_RAW_SAP_VK12変更履歴.csv",
        ["# SAP S/4HANA - Transaction VK12",
         "# Report:   Customer-Material Price Condition Change Log",
         "# Table:    KONH / KONP (Condition Records)",
         "# Export:   2026-02-14 14:10:22 JST",
         "# Filter:   25 samples under audit IA-REQ-2026-007"],
        "変更タイムスタンプ,サンプル№,稟議番号,製品コード,顧客コード,テーブル.フィールド,旧値,新値,変更率,変更実行ユーザ",
        rows,
        footer_lines=["# Records: 25"]
    )
    print("Created: PLC-S-007_25件対応_RAW_SAP_VK12変更履歴.csv")

    # 25件の稟議PDF
    _gen_price_change_pdfs(samples)


def _gen_price_change_pdfs(samples):
    for s in samples:
        pdf = JPPDF()
        pdf.add_page()
        pdf.set_font("YuGoth", "B", 16)
        pdf.cell(0, 10, "稟 議 書", align="C", new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("YuGoth", "", 10)
        pdf.cell(0, 5, f"稟議番号: {s['wf_no']}", align="R", new_x="LMARGIN", new_y="NEXT")
        pdf.ln(3)

        pdf.kv("件名", f"{s['cname']}向け{s['pcode']}の単価改定申請", key_w=30)
        pdf.kv("申請日", s["date"].strftime("%Y年%m月%d日"), key_w=30)
        pdf.kv("申請者", "松本 香織（営業部主任 SLS004）", key_w=30)
        pdf.ln(5)

        pdf.h2("1. 変更内容")
        pdf.table_header(["項目", "変更前", "変更後", "変更率"], [50, 45, 45, 40])
        pdf.table_row([f"製品 {s['pcode']}", f"¥{s['old_price']:,}",
                       f"¥{s['new_price']:,}", f"{s['change_rate']:+.1f}%"],
                      [50, 45, 45, 40])
        pdf.table_row(["顧客", s['cid'], "（同一）", "-"],
                      [50, 45, 45, 40], fill=True)
        pdf.ln(5)

        pdf.h2("2. 変更理由")
        reason = ("原材料費上昇への対応" if s['change_rate'] > 0
                  else "競争環境変化による価格見直し")
        pdf.body(reason)
        pdf.ln(3)

        pdf.h3("■ 承認経路")
        pdf.set_font("YuGoth", "B", 10)
        pdf.cell(40, 7, "役割", border=1, align="C", fill=True)
        pdf.cell(45, 7, "氏名", border=1, align="C", fill=True)
        pdf.cell(40, 7, "承認日時", border=1, align="C", fill=True)
        pdf.cell(30, 7, "承認印", border=1, align="C", fill=True,
                 new_x="LMARGIN", new_y="NEXT")

        approvers = [
            ("営業部課長", "斎藤 次郎 (SLS002)",
             (datetime.combine(s["date"], datetime.min.time()) +
              timedelta(hours=14)).strftime("%Y/%m/%d %H:%M")),
            ("営業本部長", "田中 太郎 (SLS001)",
             (datetime.combine(s["date"] + timedelta(days=1), datetime.min.time()) +
              timedelta(hours=10)).strftime("%Y/%m/%d %H:%M")),
        ]
        pdf.set_font("YuGoth", "", 10)
        for role, name, dt in approvers:
            pdf.cell(40, 14, role, border=1, align="C")
            pdf.cell(45, 14, name, border=1, align="C")
            pdf.cell(40, 14, dt, border=1, align="C")
            x_stamp = pdf.get_x()
            y_stamp = pdf.get_y()
            pdf.cell(30, 14, "", border=1, new_x="LMARGIN", new_y="NEXT")
            pdf.stamp("承認", x=x_stamp + 15, y=y_stamp + 7)

        out = BASE / f"PLC-S-007_25件対応_価格変更稟議_サンプル{s['no']:02d}_{s['wf_no']}.pdf"
        pdf.output(str(out))
    print(f"Created: 25 price change ringi PDFs for PLC-S-007")


if __name__ == "__main__":
    gen_plc_s_001()
    gen_plc_s_003()
    gen_plc_s_004()
    gen_plc_s_005()
    gen_plc_s_006()
    gen_plc_s_007()
    print("\nAll PLC-S evidence expansion completed.")
