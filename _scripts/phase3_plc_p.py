"""
Phase 3: PLC-P 拡張

対象統制：
- PLC-P-001 購買依頼承認 (25件RAW + 依頼一覧)
- PLC-P-002 発注承認 (25件の発注書PDF + RAW、うち不備3件維持)
- PLC-P-003 検収 (25件の検収報告書PDF + RAW)
- PLC-P-004 3-wayマッチング (25件RAW)
- PLC-P-005 仕入先マスタ管理 (25件変更履歴RAW、代表5件申請書PDF)
- PLC-P-006 支払承認 (12件月次RAW)
- PLC-P-007 期末未払計上 (87件全数RAW)
"""
import random
import sys
from pathlib import Path
from datetime import date, datetime, timedelta

sys.path.insert(0, str(Path(__file__).parent))
from pdf_util import JPPDF
from sample_gen_util import (
    VENDORS, RAW_MATERIALS,
    generate_systematic_samples, create_sample_list_excel, write_raw_csv
)

BASE = Path(r"C:\Users\nyham\work\demo_data\4.evidence\PLC-P")


APPROVERS = [
    ("清水 智明 (PUR003)", 1_000_000),
    ("林 真由美 (PUR002)", 5_000_000),
    ("木村 浩二 (PUR001)", 20_000_000),
    ("渡辺 正博 (CFO001)", 100_000_000),
    ("山本 健一 (CEO001)", 999_999_999),
]


def select_approver(amount, deficient=False):
    if deficient is True:
        return "山田 純一 (PUR004)", 500_000  # 権限外
    for name, limit in APPROVERS:
        if amount <= limit:
            return name, limit
    return APPROVERS[-1]


# ============================================================
# PLC-P-001: 購買依頼
# ============================================================
def gen_plc_p_001():
    random.seed(1001)
    dates = generate_systematic_samples(25, seed=1001)
    samples = []
    depts = [("製造本部", "森 和雄 (MFG001)"), ("製造本部", "池田 昌夫 (MFG002)"),
             ("技術本部", "山田 技術長"), ("情報システム部", "岡田 宏 (IT001)"),
             ("品質保証部", "品証部長")]
    items_desc = ["SUS304鋼材 φ20×1000mm", "特殊合金材 A種 100kg",
                  "プレス金型用材料", "切削油 ML-2", "ベアリング部品一式",
                  "精密測定器具", "安全手袋 1000組", "冷却液 20L"]
    for i, d in enumerate(dates, 1):
        dept, person = random.choice(depts)
        amount = random.choice([
            random.randint(50_000, 300_000),
            random.randint(300_000, 2_000_000),
            random.randint(2_000_000, 10_000_000),
        ])
        samples.append({
            "no": i, "pr_no": f"PR-2025-{100 + i * 15:04d}",
            "date": d, "dept": dept, "person": person,
            "item": random.choice(items_desc),
            "qty": random.randint(1, 50), "amount": amount,
            "budget": f"BGT-{dept[:2]}-{random.randint(1, 20):02d}",
        })

    create_sample_list_excel(
        BASE / "PLC-P-001_監査対象25件サンプルリスト.xlsx",
        "【PLC-P-001】監査対象25件サンプルリスト（購買依頼承認）",
        "（RAWデータをナビゲートするための取引リスト）",
        [("母集団", "FY2025 購買依頼 4,128件（SAP ME5A）"),
         ("抽出方法", "系統抽出 / 間隔165件 / 開始位置32"),
         ("抽出日時", "2026-02-12 09:30 JST"),
         ("関連RAWデータ", "PLC-P-001_25件対応_RAW_*.csv")],
        ["サンプル\n№", "依頼番号", "依頼日", "起案部門", "起案者",
         "品名", "数量", "予算額(円)", "予算コード"],
        [[s["no"], s["pr_no"], s["date"], s["dept"], s["person"],
          s["item"], s["qty"], s["amount"], s["budget"]] for s in samples],
        col_widths=[6, 14, 11, 14, 20, 24, 8, 14, 14],
        col_center=(0, 1, 3, 6, 8), col_right=(7,), col_date=(2,),
    )
    print("Created: PLC-P-001_監査対象25件サンプルリスト.xlsx")

    # SAP ME5A RAW
    rows = []
    for s in samples:
        appr_date = s["date"] + timedelta(days=random.randint(0, 2))
        rows.append([s["no"], s["pr_no"], s["date"].strftime("%Y-%m-%d"),
                     s["dept"], s["person"], s["item"], s["qty"],
                     s["amount"], s["budget"],
                     appr_date.strftime("%Y-%m-%d"),
                     "部門長承認済", "NEW"])

    write_raw_csv(
        BASE / "PLC-P-001_25件対応_RAW_SAP_ME5A_購買依頼詳細.csv",
        ["# SAP S/4HANA - Transaction ME5A",
         "# Report:   Purchase Requisition List",
         "# Export:   2026-02-12 09:45:22 JST",
         "# Filter:   25 samples under audit IA-REQ-2026-P001"],
        "サンプル№,依頼番号,依頼日,起案部門,起案者,品名,数量,予算額,予算コード,承認日,承認ステータス,ステータス",
        rows,
        footer_lines=["# Records: 25"]
    )
    print("Created: PLC-P-001_25件対応_RAW_SAP_ME5A_購買依頼詳細.csv")


# ============================================================
# PLC-P-002: 25件の発注書PDF（不備3件含む）
# ============================================================
def gen_plc_p_002():
    random.seed(2002)
    # 既存の発注書PDFを削除して作り直し
    for p in BASE.glob("PLC-P-002_発注書_*.pdf"):
        p.unlink()
    for p in BASE.glob("PLC-P-002_SAP_ME2N_*.xlsx"):
        # 母集団ファイルは新しく作る
        p.unlink()

    samples = []
    # 22件の通常ケース
    for i in range(1, 23):
        month = random.choice([4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3])
        y = 2026 if month <= 3 else 2025
        day = random.randint(1, 28)
        po_date = date(y, month, day)
        vid, vname, vcat = random.choice(VENDORS)
        amount = random.choice([
            random.randint(100_000, 500_000),
            random.randint(500_000, 5_000_000),
            random.randint(5_000_000, 30_000_000),
        ])
        approver, limit = select_approver(amount)
        rawm, rawm_name, rawm_price = random.choice(RAW_MATERIALS)
        qty = max(1, amount // rawm_price)
        po_no = f"PO-2025-{2000 + i * 130:04d}"
        samples.append({
            "no": i, "po_no": po_no, "date": po_date,
            "vid": vid, "vname": vname, "vcat": vcat,
            "rawm": rawm, "rawm_name": rawm_name, "rawm_price": rawm_price,
            "qty": qty, "amount": qty * rawm_price,
            "approver": approver, "limit": limit,
            "deficient": False, "remark": "",
        })

    # 3件の不備ケース
    deficient_cases = [
        {"no": 23, "po_no": "PO-2025-0234", "date": date(2025, 9, 12),
         "vid": "V-20002", "amount": 680_000, "limit": 500_000,
         "approver": "清水 智明 (PUR003)",
         "remark": "※権限外承認：PUR003はPO_APPROVE権限なし"},
        {"no": 24, "po_no": "PO-2025-0789", "date": date(2025, 10, 3),
         "vid": "V-20008", "amount": 1_250_000, "limit": 500_000,
         "approver": "山田 純一 (PUR004)",
         "remark": "※上限50万円超：PUR004が上限超え承認"},
        {"no": 25, "po_no": "PO-2025-1456", "date": date(2025, 11, 8),
         "vid": "V-20004", "amount": 7_850_000, "limit": 5_000_000,
         "approver": "林 真由美 (PUR002)",
         "remark": "※課長上限500万円超の¥7.85M承認"},
    ]
    for d in deficient_cases:
        vid = d["vid"]
        vname = next(v[1] for v in VENDORS if v[0] == vid)
        vcat = next(v[2] for v in VENDORS if v[0] == vid)
        samples.append({
            **d, "vname": vname, "vcat": vcat,
            "rawm": "RAW-001", "rawm_name": "SUS304鋼材 φ30×3000mm",
            "rawm_price": 28500,
            "qty": max(1, d["amount"] // 28500),
            "deficient": True,
        })

    create_sample_list_excel(
        BASE / "PLC-P-002_監査対象25件サンプルリスト.xlsx",
        "【PLC-P-002】監査対象25件サンプルリスト（発注承認・金額別）",
        "（RAWデータをナビゲートするための取引リスト / 不備3件を含む）",
        [("母集団", "FY2025 発注 3,874件（SAP ME2N）"),
         ("抽出方法", "系統抽出 / 間隔155件 / 開始位置89"),
         ("抽出日時", "2026-02-12 14:20 JST"),
         ("関連RAWデータ", "PLC-P-002_25件対応_RAW_*.csv / 各サンプル発注書PDF")],
        ["サンプル\n№", "発注番号", "発注日", "仕入先\nコード", "仕入先名",
         "発注金額\n(円)", "承認者", "承認上限\n(円)", "不備"],
        [[s["no"], s["po_no"], s["date"], s["vid"], s["vname"],
          s["amount"], s["approver"], s["limit"],
          "⚠" if s["deficient"] else ""] for s in samples],
        col_widths=[6, 16, 11, 10, 18, 14, 22, 12, 6],
        col_center=(0, 1, 3, 6, 8), col_right=(5, 7), col_date=(2,),
    )
    print("Created: PLC-P-002_監査対象25件サンプルリスト.xlsx")

    # SAP ME2N RAW
    rows = []
    for s in samples:
        route = "担当→課長→部長" if s["amount"] > 5_000_000 else "担当→課長"
        status = "承認済" if not s["deficient"] else "承認済(要検討)"
        rows.append([s["no"], s["po_no"], s["date"].strftime("%Y-%m-%d"),
                     s["vid"], s["vname"], s["vcat"], s["amount"],
                     s["approver"], s["limit"], route, status, s["remark"]])

    write_raw_csv(
        BASE / "PLC-P-002_25件対応_RAW_SAP_ME2N_発注詳細.csv",
        ["# SAP S/4HANA - Transaction ME2N",
         "# Report:   Purchase Order List by Vendor",
         "# Export:   2026-02-12 14:45:08 JST",
         "# Filter:   25 samples under audit IA-REQ-2026-P002"],
        "サンプル№,発注番号,発注日,仕入先コード,仕入先名,品目分類,発注金額,承認者,承認者上限,承認ルート,ステータス,備考",
        rows,
        footer_lines=["# Records: 25 (including 3 deficient cases)"]
    )
    print("Created: PLC-P-002_25件対応_RAW_SAP_ME2N_発注詳細.csv")

    # ワークフロー履歴RAW
    wf_rows = []
    for s in samples:
        wf_no = f"WF-2025-{4000 + s['no'] * 23:05d}"
        start = datetime.combine(s["date"], datetime.min.time()) + timedelta(hours=10)
        requester = "清水 智明 (PUR003)"
        if s["deficient"] and s["no"] == 24:
            requester = "山田 純一 (PUR004)"
        wf_rows.append([start.strftime("%Y-%m-%d %H:%M:%S"), wf_no, s["no"],
                        s["po_no"], requester, "起票", ""])
        wf_rows.append([(start + timedelta(hours=3)).strftime("%Y-%m-%d %H:%M:%S"),
                        wf_no, s["no"], s["po_no"], s["approver"],
                        "承認", f"金額{s['amount']:,}円 / 上限{s['limit']:,}円"])

    write_raw_csv(
        BASE / "PLC-P-002_25件対応_RAW_SAPワークフロー承認履歴.csv",
        ["# SAP Business Workflow - Purchase Order Approval History",
         "# Export:   2026-02-12 14:50:12 JST",
         "# Filter:   25 PO approval workflows"],
        "タイムスタンプ,ワークフロー番号,サンプル№,発注番号,アクター,アクション,コメント",
        wf_rows,
        footer_lines=["# Records: 50 (25 start + 25 approval)"]
    )
    print("Created: PLC-P-002_25件対応_RAW_SAPワークフロー承認履歴.csv")

    # 25件の発注書PDF
    _gen_po_pdfs(samples)


def _gen_po_pdfs(samples):
    for s in samples:
        pdf = JPPDF()
        pdf.add_page()
        pdf.set_font("YuGoth", "B", 20)
        pdf.cell(0, 12, "発 注 書", align="C", new_x="LMARGIN", new_y="NEXT")
        pdf.ln(3)
        pdf.set_font("YuGoth", "", 10)
        pdf.cell(0, 5, f"発注番号: {s['po_no']}", align="R", new_x="LMARGIN", new_y="NEXT")
        pdf.cell(0, 5, f"発注日: {s['date'].strftime('%Y年%m月%d日')}",
                 align="R", new_x="LMARGIN", new_y="NEXT")
        pdf.ln(5)

        pdf.set_font("YuGoth", "B", 11)
        pdf.cell(0, 6, "デモA株式会社 購買部",
                 new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("YuGoth", "", 9)
        pdf.cell(0, 5, "〒XXX-XXXX 神奈川県横浜市港北区",
                 new_x="LMARGIN", new_y="NEXT")
        pdf.ln(5)

        pdf.set_font("YuGoth", "B", 12)
        pdf.cell(0, 7, f"{s['vname']} 御中", new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("YuGoth", "", 9)
        pdf.cell(0, 5, f"仕入先コード: {s['vid']} / 品目分類: {s['vcat']}",
                 new_x="LMARGIN", new_y="NEXT")
        pdf.ln(5)

        pdf.set_font("YuGoth", "", 10)
        pdf.multi_cell(0, 5, "下記のとおり発注申し上げます。納期厳守のうえ、ご納入ください。")
        pdf.ln(3)

        pdf.table_header(["品目コード", "品名", "数量", "単価(円)", "金額(円)"],
                         [30, 80, 20, 30, 30])
        pdf.table_row([s["rawm"], s["rawm_name"], f"{s['qty']:,}",
                       f"{s['rawm_price']:,}", f"{s['amount']:,}"],
                      [30, 80, 20, 30, 30])

        tax = int(s["amount"] * 0.1)
        total = s["amount"] + tax
        pdf.set_font("YuGoth", "B", 10)
        pdf.cell(130, 7, "小計", border=1, align="R")
        pdf.cell(60, 7, f"¥ {s['amount']:,}", border=1, align="R",
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

        pdf.h3("■ 社内承認記録")
        pdf.set_font("YuGoth", "B", 10)
        pdf.cell(50, 7, "承認者役割", border=1, align="C", fill=True)
        pdf.cell(60, 7, "氏名", border=1, align="C", fill=True)
        pdf.cell(40, 7, "承認日", border=1, align="C", fill=True)
        pdf.cell(30, 7, "承認印", border=1, align="C", fill=True,
                 new_x="LMARGIN", new_y="NEXT")

        role = ("代表取締役" if s["limit"] > 100_000_000 else
                "管理本部長（CFO）" if s["limit"] > 20_000_000 else
                "購買部長" if s["limit"] > 5_000_000 else
                "購買部課長" if s["limit"] > 1_000_000 else
                "購買部担当（※権限外）" if s["deficient"] else "購買部担当")

        pdf.set_font("YuGoth", "", 10)
        pdf.cell(50, 14, role, border=1, align="C")
        pdf.cell(60, 14, s["approver"], border=1, align="C")
        pdf.cell(40, 14, s["date"].strftime("%Y/%m/%d"), border=1, align="C")
        x_stamp = pdf.get_x()
        y_stamp = pdf.get_y()
        pdf.cell(30, 14, "", border=1, new_x="LMARGIN", new_y="NEXT")
        if s["deficient"]:
            pdf.set_text_color(200, 30, 30)
            pdf.set_draw_color(200, 30, 30)
            pdf.circle(x_stamp + 15, y_stamp + 7, 8)
            pdf.set_font("YuGoth", "B", 8)
            pdf.text(x_stamp + 9, y_stamp + 8, "要検討")
            pdf.set_text_color(0, 0, 0)
            pdf.set_draw_color(0, 0, 0)
        else:
            pdf.stamp("承認", x=x_stamp + 15, y=y_stamp + 7)

        suffix = "_不備" if s["deficient"] else ""
        out = BASE / f"PLC-P-002_25件対応_発注書_サンプル{s['no']:02d}_{s['po_no']}{suffix}.pdf"
        pdf.output(str(out))
    print(f"Created: 25 PO PDFs for PLC-P-002")


# ============================================================
# PLC-P-003: 25件の検収報告書PDF + RAW
# ============================================================
def gen_plc_p_003():
    random.seed(3003)
    # 古いPDFを削除
    for p in BASE.glob("PLC-P-003_検収*.pdf"):
        p.unlink()

    dates = generate_systematic_samples(25, seed=3003)
    samples = []
    for i, d in enumerate(dates, 1):
        vid, vname, vcat = random.choice(VENDORS)
        rawm, rawm_name, rawm_price = random.choice(RAW_MATERIALS)
        qty_po = random.randint(10, 200)
        qty_received = qty_po if i not in (7, 17) else qty_po - random.randint(2, 5)
        amount = qty_received * rawm_price
        po_no = f"PO-2025-{random.randint(100, 3500):04d}"
        rec_no = f"REC-2025-{5000 + i * 73:04d}"
        samples.append({
            "no": i, "rec_no": rec_no, "po_no": po_no, "date": d,
            "vid": vid, "vname": vname, "vcat": vcat,
            "rawm": rawm, "rawm_name": rawm_name, "rawm_price": rawm_price,
            "qty_po": qty_po, "qty_received": qty_received,
            "amount": amount,
            "has_diff": qty_po != qty_received,
        })

    create_sample_list_excel(
        BASE / "PLC-P-003_監査対象25件サンプルリスト.xlsx",
        "【PLC-P-003】監査対象25件サンプルリスト（検収）",
        "（RAWデータをナビゲートするための取引リスト）",
        [("母集団", "FY2025 検収 3,856件（SAP MIGO）"),
         ("抽出方法", "系統抽出 / 間隔154件 / 開始位置28"),
         ("抽出日時", "2026-02-13 08:30 JST"),
         ("関連RAWデータ", "PLC-P-003_25件対応_RAW_*.csv / 各サンプル検収報告書PDF")],
        ["サンプル\n№", "検収番号", "検収日", "発注番号", "仕入先\nコード",
         "仕入先名", "品目コード", "発注数量", "受領数量", "検収金額\n(円)"],
        [[s["no"], s["rec_no"], s["date"], s["po_no"], s["vid"], s["vname"],
          s["rawm"], s["qty_po"], s["qty_received"], s["amount"]] for s in samples],
        col_widths=[6, 14, 11, 14, 10, 18, 10, 10, 10, 14],
        col_center=(0, 1, 3, 4, 6, 7, 8), col_right=(9,), col_date=(2,),
    )
    print("Created: PLC-P-003_監査対象25件サンプルリスト.xlsx")

    # SAP MIGO RAW
    rows = []
    for s in samples:
        ts = datetime.combine(s["date"], datetime.min.time()) + timedelta(
            hours=random.randint(9, 17), minutes=random.randint(0, 59))
        result = "合格" if not s["has_diff"] else "数量差異"
        rows.append([s["no"], s["rec_no"], ts.strftime("%Y-%m-%d %H:%M:%S"),
                     s["po_no"], s["vid"], s["rawm"], s["qty_po"],
                     s["qty_received"], s["amount"], result,
                     "WHS001 橋本 明"])

    write_raw_csv(
        BASE / "PLC-P-003_25件対応_RAW_SAP_MIGO_検収詳細.csv",
        ["# SAP S/4HANA - Transaction MIGO",
         "# Report:   Goods Receipt History",
         "# Table:    MSEG (material documents)",
         "# Export:   2026-02-13 08:45:20 JST"],
        "サンプル№,検収番号,検収タイムスタンプ,発注番号,仕入先コード,品目コード,発注数量,受領数量,検収金額,判定,検収担当",
        rows,
        footer_lines=["# Records: 25"]
    )
    print("Created: PLC-P-003_25件対応_RAW_SAP_MIGO_検収詳細.csv")

    # 25件の検収報告書PDF
    _gen_grn_pdfs(samples)


def _gen_grn_pdfs(samples):
    for s in samples:
        pdf = JPPDF()
        pdf.add_page()
        pdf.set_font("YuGoth", "B", 20)
        pdf.cell(0, 12, "検 収 報 告 書", align="C", new_x="LMARGIN", new_y="NEXT")
        pdf.ln(3)
        pdf.set_font("YuGoth", "", 10)
        pdf.cell(0, 5, f"検収番号: {s['rec_no']}", align="R",
                 new_x="LMARGIN", new_y="NEXT")
        pdf.cell(0, 5, f"検収日: {s['date'].strftime('%Y年%m月%d日')}",
                 align="R", new_x="LMARGIN", new_y="NEXT")
        pdf.ln(5)

        pdf.kv("発注番号", s["po_no"], key_w=30)
        pdf.kv("仕入先", f"{s['vid']} {s['vname']}", key_w=30)
        pdf.kv("検収担当", "橋本 明（倉庫課長 WHS001）", key_w=30)
        pdf.kv("品質保証確認", "品質保証部 検査員", key_w=30)
        pdf.ln(5)

        pdf.h2("検収明細")
        pdf.table_header(["品目コード", "品名", "発注数量", "受領数量", "判定"],
                         [30, 80, 25, 25, 30])
        judgment = "合格" if not s["has_diff"] else "数量差異"
        pdf.table_row([s["rawm"], s["rawm_name"], f"{s['qty_po']}",
                       f"{s['qty_received']}", judgment],
                      [30, 80, 25, 25, 30])
        pdf.ln(5)

        if s["has_diff"]:
            pdf.h2("差異内容")
            pdf.set_font("YuGoth", "", 10)
            diff = s["qty_po"] - s["qty_received"]
            pdf.multi_cell(0, 5,
                           f"発注 {s['qty_po']}個に対し受領 {s['qty_received']}個（差 -{diff}個）。"
                           f"仕入先に追加納入を依頼し、別途検収予定。")
            pdf.ln(3)

        pdf.h3("■ 検収判定")
        pdf.set_font("YuGoth", "B", 12)
        if not s["has_diff"]:
            pdf.set_fill_color(220, 240, 220)
            pdf.cell(0, 10, "検収合格 / SAPに登録済", align="C", fill=True,
                     new_x="LMARGIN", new_y="NEXT")
        else:
            pdf.set_fill_color(255, 235, 200)
            pdf.cell(0, 10, f"一部検収 / 受領数量分のみSAP登録", align="C", fill=True,
                     new_x="LMARGIN", new_y="NEXT")
        pdf.set_fill_color(255, 255, 255)
        pdf.ln(5)

        pdf.set_font("YuGoth", "B", 10)
        pdf.cell(45, 7, "検収担当者", border=1, align="C", fill=True)
        pdf.cell(45, 7, "倉庫課長", border=1, align="C", fill=True)
        pdf.cell(45, 7, "品質保証部", border=1, align="C", fill=True)
        pdf.cell(45, 7, "購買部確認", border=1, align="C", fill=True,
                 new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("YuGoth", "", 10)
        pdf.cell(45, 16, "", border=1)
        pdf.cell(45, 16, "", border=1)
        pdf.cell(45, 16, "", border=1)
        pdf.cell(45, 16, "", border=1, new_x="LMARGIN", new_y="NEXT")
        y = pdf.get_y() - 12
        pdf.stamp("検収", x=22, y=y)
        pdf.stamp("確認", x=67, y=y)
        pdf.stamp("合格", x=112, y=y)
        pdf.stamp("確認", x=157, y=y)

        out = BASE / f"PLC-P-003_25件対応_検収報告書_サンプル{s['no']:02d}_{s['rec_no']}.pdf"
        pdf.output(str(out))
    print(f"Created: 25 GRN PDFs for PLC-P-003")


# ============================================================
# PLC-P-004: 3-wayマッチング RAW
# ============================================================
def gen_plc_p_004():
    random.seed(4004)
    dates = generate_systematic_samples(25, seed=4004)
    rows = []
    for i, d in enumerate(dates, 1):
        vid, vname, _ = random.choice(VENDORS)
        po_amount = random.randint(200_000, 10_000_000)
        rec_amount = po_amount
        inv_amount = po_amount
        if i == 19:
            inv_amount = po_amount - 3_000
            result = "ALLOWED_TOLERANCE"
        elif i == 25:
            inv_amount = po_amount + 150_000
            result = "EXCEPTION_OVER_TOLERANCE"
        else:
            result = "OK"
        ts = datetime.combine(d, datetime.min.time()) + timedelta(hours=2)
        rows.append([ts.strftime("%Y-%m-%d %H:%M:%S"), i,
                     f"INV-V-{d.strftime('%Y%m')}-{i * 7:04d}",
                     d.strftime("%Y-%m-%d"),
                     f"PO-2025-{random.randint(100, 3000):04d}",
                     f"REC-2025-{random.randint(1000, 9000):04d}",
                     vid, po_amount, rec_amount, inv_amount,
                     inv_amount - po_amount, result])

    write_raw_csv(
        BASE / "PLC-P-004_25件対応_RAW_SAP_MIRO_3wayマッチング結果.csv",
        ["# SAP S/4HANA - Transaction MIRO (Invoice Verification)",
         "# Report:   3-way Matching Result Log",
         "# Tolerance: +/- JPY 10,000 OR +/- 5.0%",
         "# Export:   2026-02-13 10:20:30 JST"],
        "実行タイムスタンプ,サンプル№,請求書番号,請求日,発注番号,検収番号,仕入先コード,PO金額,検収金額,請求金額,差異,判定",
        rows,
        footer_lines=["# Records: 25"]
    )
    print("Created: PLC-P-004_25件対応_RAW_SAP_MIRO_3wayマッチング結果.csv")

    create_sample_list_excel(
        BASE / "PLC-P-004_監査対象25件サンプルリスト.xlsx",
        "【PLC-P-004】監査対象25件サンプルリスト（3-wayマッチング）",
        "（RAWデータをナビゲートするための取引リスト）",
        [("母集団", "FY2025 3-wayマッチ 3,812件（SAP MIRO）"),
         ("抽出方法", "系統抽出 / 間隔152件"),
         ("抽出日時", "2026-02-13 10:15 JST"),
         ("関連RAWデータ", "PLC-P-004_25件対応_RAW_*.csv")],
        ["サンプル\n№", "請求書番号", "請求日", "発注番号", "仕入先\nコード",
         "PO金額(円)", "検収金額(円)", "請求金額(円)", "差異(円)", "判定"],
        [[r[1], r[2], r[3], r[4], r[6], r[7], r[8], r[9], r[10], r[11]] for r in rows],
        col_widths=[6, 16, 11, 14, 10, 14, 14, 14, 12, 22],
        col_center=(0, 1, 3, 4, 9), col_right=(5, 6, 7, 8), col_date=(2,),
    )
    print("Created: PLC-P-004_監査対象25件サンプルリスト.xlsx")


# ============================================================
# PLC-P-005: 25件の仕入先マスタ変更 RAW + 代表5件の申請書PDF
# ============================================================
def gen_plc_p_005():
    random.seed(5005)
    # 既存の申請書PDFを削除
    for p in BASE.glob("PLC-P-005_仕入先マスタ登録申請書*.pdf"):
        p.unlink()

    dates = generate_systematic_samples(25, seed=5005)
    samples = []
    change_types = ["新規登録", "住所変更", "銀行口座変更", "担当者変更", "支払条件変更"]
    for i, d in enumerate(dates, 1):
        vid, vname, vcat = random.choice(VENDORS)
        ctype = random.choice(change_types)
        app_no = f"VEND-CHG-2025-{i:04d}"
        samples.append({
            "no": i, "app_no": app_no, "date": d,
            "vid": vid, "vname": vname, "vcat": vcat,
            "change_type": ctype,
        })

    create_sample_list_excel(
        BASE / "PLC-P-005_監査対象25件サンプルリスト.xlsx",
        "【PLC-P-005】監査対象25件サンプルリスト（仕入先マスタ管理）",
        "（RAWデータをナビゲートするための取引リスト）",
        [("母集団", "FY2025 マスタ変更 48件（SAP XK01/XK02変更履歴）"),
         ("抽出方法", "系統抽出 / 間隔2件 / 25件"),
         ("抽出日時", "2026-02-13 14:00 JST"),
         ("関連RAWデータ", "PLC-P-005_25件対応_RAW_*.csv / 代表5件の申請書PDF")],
        ["サンプル\n№", "申請番号", "変更日", "仕入先\nコード", "仕入先名", "変更種別"],
        [[s["no"], s["app_no"], s["date"], s["vid"], s["vname"], s["change_type"]]
         for s in samples],
        col_widths=[6, 16, 11, 10, 20, 16],
        col_center=(0, 1, 3, 5), col_date=(2,),
    )
    print("Created: PLC-P-005_監査対象25件サンプルリスト.xlsx")

    # RAW変更履歴
    rows = []
    for s in samples:
        ts = datetime.combine(s["date"], datetime.min.time()) + timedelta(
            hours=random.randint(10, 16))
        rows.append([ts.strftime("%Y-%m-%d %H:%M:%S"), s["no"], s["app_no"],
                     s["vid"], s["vname"], s["change_type"],
                     "PUR003 清水 智明", "PUR001 木村 浩二"])

    write_raw_csv(
        BASE / "PLC-P-005_25件対応_RAW_SAP_XK01_XK02_仕入先マスタ変更履歴.csv",
        ["# SAP S/4HANA - Transaction XK01 (Create) / XK02 (Change)",
         "# Report:   Vendor Master Change History (Table CDHDR/CDPOS)",
         "# Export:   2026-02-13 14:15:08 JST"],
        "変更タイムスタンプ,サンプル№,申請番号,仕入先コード,仕入先名,変更種別,申請者,承認者",
        rows,
        footer_lines=["# Records: 25"]
    )
    print("Created: PLC-P-005_25件対応_RAW_SAP_XK01_XK02_仕入先マスタ変更履歴.csv")

    # 代表5件の申請書PDF
    for s in samples[:5]:
        pdf = JPPDF()
        pdf.add_page()
        pdf.h1("仕入先マスタ変更申請書")
        pdf.set_font("YuGoth", "", 10)
        pdf.cell(0, 5, f"申請番号: {s['app_no']} / 申請日: {s['date'].strftime('%Y年%m月%d日')}",
                 align="R", new_x="LMARGIN", new_y="NEXT")
        pdf.ln(5)

        pdf.kv("申請者", "清水 智明（購買部主任 PUR003）")
        pdf.kv("変更種別", s["change_type"])
        pdf.kv("対象仕入先", f"{s['vid']} {s['vname']}")
        pdf.kv("品目分類", s["vcat"])
        pdf.ln(5)

        pdf.h2("変更内容")
        pdf.body(f"{s['change_type']}に関する変更申請。詳細は添付の変更前後対比表を参照。")
        pdf.ln(3)

        if s["change_type"] == "新規登録":
            pdf.h2("反社会的勢力チェック")
            pdf.kv("チェック実施日", s["date"].strftime("%Y/%m/%d"))
            pdf.kv("チェック担当", "総務部 前田 美香 (GA001)")
            pdf.kv("チェック結果", "○ 問題なし")
            pdf.ln(5)

        pdf.h3("■ 承認")
        pdf.set_font("YuGoth", "B", 10)
        pdf.cell(60, 7, "役割", border=1, align="C", fill=True)
        pdf.cell(60, 7, "氏名", border=1, align="C", fill=True)
        pdf.cell(40, 7, "日付", border=1, align="C", fill=True)
        pdf.cell(30, 7, "承認印", border=1, align="C", fill=True,
                 new_x="LMARGIN", new_y="NEXT")
        approvers = [
            ("購買部課長", "林 真由美 (PUR002)",
             (s["date"] + timedelta(days=1)).strftime("%Y/%m/%d")),
            ("購買部長", "木村 浩二 (PUR001)",
             (s["date"] + timedelta(days=2)).strftime("%Y/%m/%d")),
        ]
        pdf.set_font("YuGoth", "", 10)
        for role, name, dt in approvers:
            pdf.cell(60, 14, role, border=1, align="C")
            pdf.cell(60, 14, name, border=1, align="C")
            pdf.cell(40, 14, dt, border=1, align="C")
            x_stamp = pdf.get_x()
            y_stamp = pdf.get_y()
            pdf.cell(30, 14, "", border=1, new_x="LMARGIN", new_y="NEXT")
            pdf.stamp("承認", x=x_stamp + 15, y=y_stamp + 7)

        out = BASE / f"PLC-P-005_25件対応_仕入先マスタ変更申請書_サンプル{s['no']:02d}_{s['app_no']}.pdf"
        pdf.output(str(out))
    print("Created: 5 vendor master change PDFs for PLC-P-005")


# ============================================================
# PLC-P-006: 12件の月次支払RAW
# ============================================================
def gen_plc_p_006():
    random.seed(6006)
    # 12ヶ月分の支払実行RAW
    import calendar
    rows = []
    for month_offset in range(12):
        y = 2025 if month_offset < 9 else 2026
        m = 4 + month_offset if month_offset < 9 else month_offset - 8
        pay_date = date(y, m, calendar.monthrange(y, m)[1])
        total = random.randint(300_000_000, 800_000_000)
        vendor_count = random.randint(18, 24)
        rows.append([pay_date.strftime("%Y-%m-%d"), month_offset + 1,
                     f"JOB_FB_{y}{m:02d}", total, vendor_count,
                     "小川 由紀 (ACC005)", "正常終了"])

    write_raw_csv(
        BASE / "PLC-P-006_全12ヶ月RAW_SAP_F110_支払実行バッチ.csv",
        ["# SAP S/4HANA - Transaction F110",
         "# Report:   Automatic Payment Program Execution Log",
         "# Period:   FY2025 (12 months exhaustive)",
         "# Export:   2026-04-10 08:00:00 JST"],
        "支払実行日,月次№,ジョブ番号,総支払額(円),支払先件数,実行ユーザ,実行結果",
        rows,
        footer_lines=["# Records: 12 months (exhaustive)"]
    )
    print("Created: PLC-P-006_全12ヶ月RAW_SAP_F110_支払実行バッチ.csv")


# ============================================================
# PLC-P-007: 期末未払計上RAW (全87件)
# ============================================================
def gen_plc_p_007():
    random.seed(7007)
    rows = []
    for i in range(1, 88):
        rec_date = date(2026, 3, random.randint(20, 31))
        vid, vname, _ = random.choice(VENDORS)
        amount = random.randint(500_000, 15_000_000)
        rows.append([i, f"REC-2026-{5000 + i:04d}",
                     rec_date.strftime("%Y-%m-%d"), vid, vname,
                     amount, "ARRIVED" if random.random() > 0.25 else "PENDING",
                     amount,
                     f"JV-ACC-202603-{i:04d}"])

    write_raw_csv(
        BASE / "PLC-P-007_全87件RAW_SAP期末未払計上明細.csv",
        ["# SAP S/4HANA - FY2025 Year-end Accrual Detail",
         "# Source:   MIGO (receipts) + MIRO (invoices) + ACC-JV (manual accruals)",
         "# Period:   2026-03-20 to 2026-03-31 (receipts without invoice)",
         "# Export:   2026-04-05 14:00:15 JST"],
        "№,検収番号,検収日,仕入先コード,仕入先名,検収金額,請求書到着状況,未払計上額,計上仕訳番号",
        rows,
        footer_lines=["# Records: 87 (exhaustive)"]
    )
    print("Created: PLC-P-007_全87件RAW_SAP期末未払計上明細.csv")


if __name__ == "__main__":
    gen_plc_p_001()
    gen_plc_p_002()
    gen_plc_p_003()
    gen_plc_p_004()
    gen_plc_p_005()
    gen_plc_p_006()
    gen_plc_p_007()
    print("\nAll PLC-P evidence expansion completed.")
