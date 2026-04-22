"""
抜粋型ドキュメントの修正
1. 職務権限規程R18: 完全版PDFを0.profile/に配置、抜粋PDFを削除
2. SOC1レポート: 抜粋表記を削除、完全な目次構成に拡充
3. ファイル名の「抜粋」「サンプル」を適切に
"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
from pdf_util import JPPDF

BASE = Path(r"C:\Users\nyham\work\demo_data")


# ============================================================
# 1. 完全版 職務権限規程 R18
# ============================================================
def gen_full_r18():
    pdf = JPPDF()
    pdf.add_page()

    # 表紙
    pdf.set_font("YuGoth", "B", 24)
    pdf.ln(30)
    pdf.cell(0, 16, "職 務 権 限 規 程", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("YuGoth", "B", 14)
    pdf.cell(0, 10, "（規程番号 R18）", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(20)

    pdf.set_font("YuGoth", "", 12)
    pdf.cell(0, 8, "デモA株式会社", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(80)

    pdf.cell(0, 6, "制定： 1998年4月1日", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 6, "最終改訂： 2025年4月1日（第15回改訂）", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 6, "主管部門： 総務部", align="C", new_x="LMARGIN", new_y="NEXT")

    # 目次
    pdf.add_page()
    pdf.set_font("YuGoth", "B", 16)
    pdf.cell(0, 10, "目 次", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(5)

    toc = [
        ("第1章", "総則", 3),
        ("  第1条", "目的", 3),
        ("  第2条", "適用範囲", 3),
        ("  第3条", "定義", 3),
        ("  第4条", "権限行使の原則", 4),
        ("第2章", "経営に関する権限", 4),
        ("  第5条", "取締役会の決議事項", 4),
        ("  第6条", "代表取締役の権限", 5),
        ("  第7条", "管理本部長(CFO)の権限", 5),
        ("第3章", "業務分野別の承認権限", 6),
        ("  第8条", "販売関連の承認権限", 6),
        ("  第9条", "購買関連の承認権限", 7),
        ("  第10条", "人事関連の承認権限", 8),
        ("  第11条", "IT関連の承認権限", 9),
        ("  第12条", "財務・会計関連の承認権限", 9),
        ("第4章", "例外手続", 10),
        ("  第13条", "緊急時の権限代行", 10),
        ("  第14条", "権限逸脱発生時の手続", 10),
        ("第5章", "職務分掌の原則", 11),
        ("  第15条", "職務分掌の基本原則", 11),
        ("  第16条", "併任禁止の職務", 11),
        ("附則", "", 11),
    ]
    pdf.set_font("YuGoth", "", 11)
    for art, title, page in toc:
        if art.startswith("  "):
            pdf.set_x(20)
        else:
            pdf.set_x(15)
            pdf.set_font("YuGoth", "B", 11)
        label = f"{art} {title}" if title else art
        pdf.cell(130, 7, label)
        pdf.cell(20, 7, f"...  {page}", align="R", new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("YuGoth", "", 11)

    # 第1章 総則
    pdf.add_page()
    pdf.h1("第1章 総則")
    pdf.h2("第1条 (目的)")
    pdf.body("本規程は、デモA株式会社（以下「当社」という）における"
             "業務執行上の職務権限及び責任の所在を明確にし、"
             "業務の効率的かつ適正な執行を図ることを目的とする。")

    pdf.h2("第2条 (適用範囲)")
    pdf.body("本規程は当社のすべての役員及び従業員に適用される。"
             "子会社については、本規程の趣旨に従い、各社が別途定める。")

    pdf.h2("第3条 (定義)")
    pdf.body("本規程において使用する用語の定義は次のとおりとする。\n"
             "(1) 「業務執行」とは、当社の事業に関する一切の行為をいう。\n"
             "(2) 「決裁」とは、権限を有する者が承認することをいう。\n"
             "(3) 「稟議」とは、ワークフローシステム（S04）を通じて行う決裁手続をいう。\n"
             "(4) 「承認金額」とは、消費税を含まない取引金額（税抜）をいう。")

    pdf.h2("第4条 (権限行使の原則)")
    pdf.body("職務権限の行使にあたっては、次の原則を遵守しなければならない。\n"
             "(1) 権限の範囲内で行使すること\n"
             "(2) 業務の必要性に基づくこと\n"
             "(3) 関連規程及び法令を遵守すること\n"
             "(4) 職務分掌を守ること（第15条参照）")

    # 第2章 経営に関する権限
    pdf.add_page()
    pdf.h1("第2章 経営に関する権限")
    pdf.h2("第5条 (取締役会の決議事項)")
    pdf.body("取締役会は次の事項を決議する。\n"
             "(1) 中長期経営計画及び年度予算の承認\n"
             "(2) 100億円超の設備投資・M&A\n"
             "(3) 10億円超の借入・社債発行\n"
             "(4) 定款変更及び株主総会議案\n"
             "(5) 取締役及び執行役員の選任・解任\n"
             "(6) 規程の制定・改廃（重要なもの）\n"
             "(7) 開示すべき重要な不備の認定")

    pdf.h2("第6条 (代表取締役の権限)")
    pdf.body("代表取締役は取締役会決議事項以外の経営上重要な事項について決裁する。"
             "特に次の事項について個別承認権を有する。\n"
             "(1) 1件¥1億円超の設備投資\n"
             "(2) 1件¥1億円超の契約締結\n"
             "(3) 重要人事（部長以上）\n"
             "(4) 規程の制定・改廃（軽微なもの）")

    pdf.add_page()
    pdf.h2("第7条 (管理本部長(CFO)の権限)")
    pdf.body("管理本部長は経理・財務・人事・総務・法務に関する業務を統括し、"
             "次の事項について承認権を有する。\n"
             "(1) 1件¥1億円以下の財務取引\n"
             "(2) 1件¥1億円以下の発注承認\n"
             "(3) 見積りの評価・承認\n"
             "(4) 月次・四半期・年次決算承認\n"
             "(5) 税務申告の承認")

    # 第3章 業務分野別
    pdf.h1("第3章 業務分野別の承認権限")

    pdf.h2("第8条 (販売関連の承認権限)")
    pdf.body("販売関連の承認権限は次のとおり定める。")
    pdf.set_font("YuGoth", "B", 10)
    pdf.ln(2)
    pdf.table_header(["業務", "承認者", "承認上限"], [60, 65, 45])
    pdf.table_row(["通常受注（与信枠内）", "自動承認（SAP）", "－"], [60, 65, 45])
    pdf.table_row(["与信限度超過の受注", "営業本部長", "超過全件対象"], [60, 65, 45], fill=True)
    pdf.table_row(["新規顧客登録", "営業本部長+CFO", "全件対象"], [60, 65, 45])
    pdf.table_row(["与信限度引上", "CFO", "年次見直し"], [60, 65, 45], fill=True)
    pdf.table_row(["¥100M超の個別受注", "代表取締役", "全件対象"], [60, 65, 45])
    pdf.table_row(["顧客別価格変更（既存）", "営業本部長", "全件対象"], [60, 65, 45], fill=True)
    pdf.table_row(["新規品目の初期価格設定", "営業本部長+CFO", "全件対象"], [60, 65, 45])
    pdf.table_row(["返品・値引き（¥5M超）", "営業本部長", "全件対象"], [60, 65, 45], fill=True)
    pdf.table_row(["回収不能認定", "CFO", "全件対象"], [60, 65, 45])

    pdf.add_page()
    pdf.h2("第9条 (購買関連の承認権限)")
    pdf.body("購買関連の承認権限は次のとおり定める。金額区分に応じて承認者が自動判定され、"
             "SAPワークフロー（S04）経由で承認手続を行う。")
    pdf.set_font("YuGoth", "B", 10)
    pdf.ln(2)
    pdf.table_header(["金額区分", "承認者", "SAPロール"], [55, 55, 60])
    pdf.table_row(["〜¥500,000", "購買部担当（主任）", "PO_CREATE のみ"],
                  [55, 55, 60])
    pdf.table_row(["〜¥5,000,000", "購買部課長", "PO_APPROVE"],
                  [55, 55, 60], fill=True)
    pdf.table_row(["〜¥20,000,000", "購買部長", "PO_APPROVE"],
                  [55, 55, 60])
    pdf.table_row(["〜¥100,000,000", "管理本部長（CFO）", "PO_APPROVE"],
                  [55, 55, 60], fill=True)
    pdf.table_row(["¥100,000,000超", "代表取締役", "PO_APPROVE"],
                  [55, 55, 60])
    pdf.ln(3)
    pdf.body("その他の購買関連権限：\n"
             "(1) 新規仕入先登録：購買部長（反社チェック・信用調査完了後）\n"
             "(2) 仕入先評価：購買部長が年1回実施\n"
             "(3) 継続契約解除：購買部長\n"
             "(4) 外注加工契約締結：購買部長（¥50M超はCFO追加承認）")

    pdf.add_page()
    pdf.h2("第10条 (人事関連の承認権限)")
    pdf.set_font("YuGoth", "B", 10)
    pdf.table_header(["業務", "承認者"], [80, 80])
    pdf.table_row(["採用（一般社員）", "人事部長"], [80, 80])
    pdf.table_row(["採用（課長以上）", "社長"], [80, 80], fill=True)
    pdf.table_row(["異動・昇格", "人事部長→関連部門長→社長"], [80, 80])
    pdf.table_row(["給与改定（年次）", "人事部長→CFO→社長"], [80, 80], fill=True)
    pdf.table_row(["賞与支給", "人事部長→CFO→社長"], [80, 80])
    pdf.table_row(["退職手続", "人事部長"], [80, 80], fill=True)
    pdf.table_row(["懲戒処分", "人事部長→コンプラ委員会→社長"], [80, 80])

    pdf.add_page()
    pdf.h2("第11条 (IT関連の承認権限)")
    pdf.set_font("YuGoth", "B", 10)
    pdf.table_header(["業務", "承認者"], [80, 80])
    pdf.table_row(["ユーザID新規登録", "所属部門長 + 情シス部アプリリーダー"], [80, 80])
    pdf.table_row(["特権ID付与", "情シス部長 + CFO"], [80, 80], fill=True)
    pdf.table_row(["ユーザID削除（退職）", "人事部 → 情シス部担当"], [80, 80])
    pdf.table_row(["プログラム変更", "情シス部アプリリーダー + 業務部門長"],
                  [80, 80], fill=True)
    pdf.table_row(["緊急変更", "情シス部長（事後承認）"], [80, 80])
    pdf.table_row(["本番移送", "情シス部アプリリーダー（専任者）"],
                  [80, 80], fill=True)
    pdf.table_row(["IT投資計画", "情シス部長 → CFO → 取締役会"], [80, 80])

    pdf.h2("第12条 (財務・会計関連の承認権限)")
    pdf.set_font("YuGoth", "B", 10)
    pdf.table_header(["業務", "承認者"], [80, 80])
    pdf.table_row(["月次決算承認", "経理部長"], [80, 80])
    pdf.table_row(["四半期決算承認", "経理部長 → CFO → 取締役会"],
                  [80, 80], fill=True)
    pdf.table_row(["年次決算承認", "CFO → 取締役会 → 監査等委員会"], [80, 80])
    pdf.table_row(["連結仕訳（非定型）", "経理部長 → CFO"],
                  [80, 80], fill=True)
    pdf.table_row(["会計上の見積（引当金等）", "CFO → 監査等委員会"], [80, 80])

    # 第4章 例外手続
    pdf.add_page()
    pdf.h1("第4章 例外手続")
    pdf.h2("第13条 (緊急時の権限代行)")
    pdf.body("通常の承認者が不在等により承認が困難な場合、以下の代行者による承認を認める。\n"
             "(1) 営業本部長不在時：営業副本部長または経営企画部長\n"
             "(2) 購買部長不在時：管理本部長（CFO）\n"
             "(3) 代表取締役不在時：取締役会議長\n"
             "代行承認は、通常承認者の復帰後に事後報告を行う。")

    pdf.h2("第14条 (権限逸脱発生時の手続)")
    pdf.body("承認権限を超過した承認が発覚した場合、次の手続をとる。\n"
             "(1) 権限者による事後承認の取得\n"
             "(2) 内部監査室への報告\n"
             "(3) 再発防止策の策定と実施\n"
             "(4) 監査等委員会への報告（重要な場合）")

    # 第5章 職務分掌
    pdf.h1("第5章 職務分掌の原則")
    pdf.h2("第15条 (職務分掌の基本原則)")
    pdf.body("権限の乱用及び誤謬を防止するため、次の職務は同一人物が兼任してはならない。")

    pdf.h2("第16条 (併任禁止の職務)")
    pdf.set_font("YuGoth", "B", 10)
    pdf.table_header(["業務", "禁止される兼任"], [80, 80])
    pdf.table_row(["発注", "発注作成 ⇔ 発注承認"], [80, 80])
    pdf.table_row(["検収", "発注承認 ⇔ 検収"], [80, 80], fill=True)
    pdf.table_row(["買掛・支払", "検収 ⇔ 買掛計上 ⇔ 支払実行"], [80, 80])
    pdf.table_row(["受注・売上", "受注登録 ⇔ 与信限度設定"], [80, 80], fill=True)
    pdf.table_row(["入金・消込", "現金受領 ⇔ 消込記帳 ⇔ 照合"], [80, 80])
    pdf.table_row(["ユーザ管理", "ユーザ作成 ⇔ ユーザ申請承認"],
                  [80, 80], fill=True)
    pdf.table_row(["変更管理", "変更開発 ⇔ 本番移送"], [80, 80])

    # 附則
    pdf.ln(8)
    pdf.h1("附 則")
    pdf.body("1. 本規程は1998年4月1日より施行する。\n"
             "2. 本規程の改廃は取締役会の決議による。\n"
             "3. 本規程に定めのない事項については、関連規程又は総務部長の判断による。\n\n"
             "改訂履歴：\n"
             "第1回改訂　1999年10月1日\n"
             "（中略）\n"
             "第14回改訂　2024年4月1日\n"
             "第15回改訂　2025年4月1日（金額区分の見直し、職務分掌の明文化）")

    pdf.ln(15)
    pdf.set_font("YuGoth", "", 10)
    pdf.cell(0, 6, "承認： 2025年4月1日施行", align="R", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 6, "代表取締役社長 山本 健一 [印]", align="R", new_x="LMARGIN", new_y="NEXT")

    out = BASE / "0.profile" / "規程_職務権限規程_R18.pdf"
    pdf.output(str(out))
    print(f"Created: {out.relative_to(BASE)}")


# ============================================================
# 2. 削除対象：抜粋PDFと「抜粋」「サンプル」付きファイル
# ============================================================
def cleanup_excerpts():
    targets = [
        "4.evidence/PLC-S/PLC-S-001_販売関連承認権限一覧_職務権限規程R18抜粋.pdf",
        "4.evidence/PLC-P/PLC-P-002_購買関連承認権限一覧_職務権限規程R18抜粋.pdf",
    ]
    for t in targets:
        p = BASE / t
        if p.exists():
            p.unlink()
            print(f"Deleted: {t}")


# ============================================================
# 3. ファイル名の改善 (「抜粋」「サンプル」→ 意味のある名前)
# ============================================================
def rename_files():
    renames = [
        # (old, new)
        ("4.evidence/PLC-P/PLC-P-002_SAPワークフロー承認履歴ログ_FY2025抜粋.csv",
         "4.evidence/PLC-P/PLC-P-002_SAPワークフロー承認履歴ログ_FY2025.csv"),
        ("4.evidence/PLC-I/PLC-I-004_3ヶ月サンプルRAW_SAP_CO88_原価差異計算結果.csv",
         "4.evidence/PLC-I/PLC-I-004_RAW_SAP_CO88_原価差異計算結果_2025年7月_10月_2026年1月.csv"),
        ("4.evidence/PLC-I/PLC-I-006_WMS-ERP在庫照合レポート_202511月次サンプル.csv",
         "4.evidence/PLC-I/PLC-I-006_WMS-ERP在庫照合レポート_202511_日次照合30日分.csv"),
    ]
    for old, new in renames:
        p_old = BASE / old
        p_new = BASE / new
        if p_old.exists():
            p_old.rename(p_new)
            print(f"Renamed: {Path(old).name}\n      -> {Path(new).name}")


# ============================================================
# 4. SOC1レポートを「抜粋」→「要約/目次版」に拡充
# ============================================================
def regenerate_soc1():
    BASE_EM = BASE / "4.evidence" / "ITGC" / "EM_外部委託管理"

    # 既存削除
    for p in BASE_EM.glob("ITGC-EM-001_RAW_SOC1レポート_*.pdf"):
        p.unlink()

    # SIer-A 完全版
    pdf = JPPDF()
    pdf.add_page()

    # 表紙
    pdf.set_font("YuGoth", "B", 22)
    pdf.ln(25)
    pdf.cell(0, 12, "SOC 1 Type II Report", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(3)
    pdf.set_font("YuGoth", "", 12)
    pdf.cell(0, 8,
             "Report on Management's Description of a Service Organization's System",
             align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 8,
             "and the Suitability of the Design and Operating Effectiveness of Controls",
             align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(20)
    pdf.set_font("YuGoth", "B", 14)
    pdf.cell(0, 10, "Service Organization: 外部委託先SIer-A",
             align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(15)
    pdf.set_font("YuGoth", "", 11)
    pdf.cell(0, 6, "Report Period: April 1, 2024 - March 31, 2025",
             align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 6, "Issued: May 20, 2025", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(50)
    pdf.cell(0, 6, "Prepared in accordance with SSAE No. 18 (AT-C 320)",
             align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 6, "By: Independent Service Auditor XYZ CPA Firm",
             align="C", new_x="LMARGIN", new_y="NEXT")

    # 目次
    pdf.add_page()
    pdf.set_font("YuGoth", "B", 16)
    pdf.cell(0, 10, "Table of Contents", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(5)
    toc = [
        ("Section I", "Independent Service Auditor's Report", 3),
        ("Section II", "Management's Assertion", 5),
        ("Section III", "Description of the Service Organization's System", 7),
        ("  III-1", "Company Overview", 7),
        ("  III-2", "Scope of Services", 8),
        ("  III-3", "System Components", 9),
        ("  III-4", "Control Environment", 10),
        ("Section IV", "Control Objectives, Related Controls, and Test Results", 12),
        ("  IV-1", "Logical Access", 12),
        ("  IV-2", "Change Management", 14),
        ("  IV-3", "Backup and Recovery", 16),
        ("  IV-4", "Incident Management", 18),
        ("  IV-5", "Physical Security", 20),
        ("Section V", "Complementary User Entity Controls (CUECs)", 22),
        ("Section VI", "Management Response to Deviations", 24),
    ]
    pdf.set_font("YuGoth", "", 11)
    for sec, title, page in toc:
        pdf.set_x(15 if not sec.startswith(" ") else 25)
        pdf.cell(130, 7, f"{sec} {title}")
        pdf.cell(20, 7, f"...  {page}", align="R", new_x="LMARGIN", new_y="NEXT")

    # Section I: Independent Service Auditor's Report
    pdf.add_page()
    pdf.h1("Section I. Independent Service Auditor's Report")
    pdf.set_font("YuGoth", "", 10)
    pdf.body("To the Management of 外部委託先SIer-A and other specified parties:")
    pdf.ln(3)
    pdf.body("Scope: We have examined 外部委託先SIer-A's description of its "
             "SAP ERP Operations and Development Services (the System) throughout "
             "the period April 1, 2024 to March 31, 2025, and the suitability "
             "of the design and operating effectiveness of the controls stated "
             "in the description to achieve the related control objectives.")
    pdf.ln(3)
    pdf.body("Opinion: In our opinion, in all material respects, "
             "(1) the description fairly presents the System; "
             "(2) the controls were suitably designed; "
             "(3) the controls operated effectively throughout the period; "
             "except for the matter noted in Section VI regarding Incident Management.")
    pdf.ln(3)
    pdf.body("Restriction on Use: This report is intended solely for the information "
             "and use of management of 外部委託先SIer-A, user entities of the System, "
             "and their auditors.")
    pdf.ln(10)
    pdf.set_font("YuGoth", "", 10)
    pdf.cell(0, 6, "XYZ CPA Firm", align="R", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 6, "Tokyo, Japan / May 20, 2025", align="R", new_x="LMARGIN", new_y="NEXT")

    # Section II: Management's Assertion
    pdf.add_page()
    pdf.h1("Section II. Management's Assertion")
    pdf.body("We, the management of 外部委託先SIer-A, have prepared the accompanying "
             "description of the SAP ERP Operations and Development Services System "
             "(the System) throughout the period April 1, 2024 to March 31, 2025. "
             "We confirm that:")
    pdf.ln(3)
    pdf.body("a. The description fairly presents the System.\n"
             "b. The controls related to the control objectives stated in the "
             "description were suitably designed.\n"
             "c. The controls operated effectively throughout the period.\n"
             "d. Complementary user entity controls (CUECs) are described in Section V.")
    pdf.ln(10)
    pdf.cell(0, 6, "外部委託先SIer-A 代表取締役 [Chief Executive Officer]",
             align="R", new_x="LMARGIN", new_y="NEXT")

    # Section III Overview
    pdf.add_page()
    pdf.h1("Section III. Description of the Service Organization's System")
    pdf.h2("III-1. Company Overview")
    pdf.body("外部委託先SIer-A is a system integration services provider specializing "
             "in SAP ERP implementation, operations, and enhancement services for "
             "manufacturing and service industries in Japan. Established in 1985, "
             "the company operates from headquarters in Tokyo with delivery centers "
             "in Osaka and Fukuoka. As of December 2024, the company employs approximately "
             "1,500 consultants and engineers.")

    pdf.h2("III-2. Scope of Services Provided to デモA株式会社")
    pdf.body("外部委託先SIer-A provides the following services to the user entity:\n"
             "(1) SAP S/4HANA application support (L2/L3 support)\n"
             "(2) Custom development and enhancements (ABAP programming)\n"
             "(3) Change management and deployment coordination\n"
             "(4) Testing support (UAT coordination)\n"
             "(5) Periodic SAP upgrade management")
    pdf.body("Note: Service scope does NOT include: infrastructure management, database administration, "
             "and backup operations (handled by 外部委託先B社 separately).")

    pdf.h2("III-3. System Components")
    pdf.body("The System includes:\n"
             "- SAP development environment (DEV client 100)\n"
             "- SAP quality assurance environment (QAS client 200)\n"
             "- Change request management tool (ServiceNow-equivalent)\n"
             "- Source code repository (Git-based)\n"
             "- Incident management system")

    pdf.h2("III-4. Control Environment")
    pdf.body("外部委託先SIer-A maintains the following control environment elements:\n"
             "- ISO 27001:2022 certification (latest recertification: 2024-09)\n"
             "- Annual employee code of conduct training\n"
             "- Segregation of duties between development and production access\n"
             "- Background check for all consultants handling client systems")

    # Section IV: Control Objectives
    pdf.add_page()
    pdf.h1("Section IV. Control Objectives, Related Controls, and Test Results")
    pdf.h2("IV-1. Logical Access")
    pdf.body("Control Objective: Controls provide reasonable assurance that logical access to "
             "the System is restricted to authorized individuals.")
    pdf.ln(2)
    pdf.table_header(["Control Activity", "Test Performed", "Result"], [70, 60, 40])
    pdf.table_row(["New user access requires manager approval",
                   "Inspected 25 new user requests", "No exceptions"], [70, 60, 40])
    pdf.table_row(["Quarterly access reviews performed",
                   "Inspected 4 quarterly reviews", "No exceptions"], [70, 60, 40], fill=True)
    pdf.table_row(["Terminated user access removed within 1 day",
                   "Tested 10 terminations", "No exceptions"], [70, 60, 40])

    pdf.h2("IV-2. Change Management")
    pdf.body("Control Objective: Controls provide reasonable assurance that changes are "
             "authorized, tested, and approved before production implementation.")
    pdf.ln(2)
    pdf.table_header(["Control Activity", "Test Performed", "Result"], [70, 60, 40])
    pdf.table_row(["All changes require written approval",
                   "Inspected 30 change requests", "No exceptions"], [70, 60, 40])
    pdf.table_row(["Changes tested in QAS before PRD",
                   "Inspected 30 test records", "No exceptions"], [70, 60, 40], fill=True)

    pdf.add_page()
    pdf.h2("IV-3. Backup and Recovery (Reference only - performed by B社)")
    pdf.body("This control area is primarily the responsibility of 外部委託先B社 (the infrastructure "
             "provider). 外部委託先SIer-A coordinates restore testing with B社.")

    pdf.h2("IV-4. Incident Management")
    pdf.body("Control Objective: Controls provide reasonable assurance that incidents are "
             "identified, tracked, and resolved timely.")
    pdf.ln(2)
    pdf.table_header(["Control Activity", "Test Performed", "Result"], [70, 60, 40])
    pdf.table_row(["All incidents logged within 1 hour",
                   "Inspected 45 incidents", "3 exceptions"], [70, 60, 40])
    pdf.table_row(["Root cause analysis for severity 1",
                   "Inspected 5 severity-1 incidents", "No exceptions"], [70, 60, 40], fill=True)
    pdf.ln(3)
    pdf.body("Exception note: 3 incidents were logged between 1-3 hours after identification "
             "(threshold = 1 hour). Management response in Section VI.")

    pdf.h2("IV-5. Physical Security")
    pdf.body("Control Objective: Access to development/QAS facilities is restricted "
             "to authorized personnel.")
    pdf.ln(2)
    pdf.table_header(["Control Activity", "Test Performed", "Result"], [70, 60, 40])
    pdf.table_row(["Badge access to development areas",
                   "Inspected access logs (3 months)", "No exceptions"], [70, 60, 40])

    # Section V: CUECs
    pdf.add_page()
    pdf.h1("Section V. Complementary User Entity Controls (CUECs)")
    pdf.body("外部委託先SIer-A assumes that user entities (including デモA株式会社) "
             "will implement the following controls:\n\n"
             "CUEC-1: User entity management will review and authorize all change requests "
             "before SIer-A implements them.\n\n"
             "CUEC-2: User entity will perform User Acceptance Testing (UAT) for all changes.\n\n"
             "CUEC-3: User entity will maintain its own access review process and notify "
             "SIer-A promptly of user access changes.\n\n"
             "CUEC-4: User entity will monitor incident resolution and escalate as needed.")

    # Section VI: Management Response
    pdf.add_page()
    pdf.h1("Section VI. Management Response to Deviations")
    pdf.body("Regarding the 3 exceptions noted in Section IV-4 (Incident Management):\n\n"
             "Root cause: The incident management tool was upgraded in September 2024, "
             "and the SLA timer configuration was incorrectly set during the upgrade.\n\n"
             "Corrective action taken:\n"
             "- Timer configuration reviewed and corrected on September 15, 2024\n"
             "- Additional training provided to L1 support staff\n"
             "- SLA monitoring dashboard enhanced\n\n"
             "Future state: No further exceptions have been noted since the correction. "
             "A special focus review will be included in the FY2025 SOC 1 engagement.")

    pdf.output(str(BASE_EM / "ITGC-EM-001_RAW_SOC1_TypeII_SIerA_FY2024.pdf"))
    print("Created: ITGC-EM-001_RAW_SOC1_TypeII_SIerA_FY2024.pdf")

    # B社も同様に拡充
    pdf = JPPDF()
    pdf.add_page()
    pdf.set_font("YuGoth", "B", 22)
    pdf.ln(25)
    pdf.cell(0, 12, "SOC 1 Type II Report", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(3)
    pdf.set_font("YuGoth", "", 12)
    pdf.cell(0, 8,
             "Report on Management's Description of a Service Organization's System",
             align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 8,
             "and the Suitability of the Design and Operating Effectiveness of Controls",
             align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(20)
    pdf.set_font("YuGoth", "B", 14)
    pdf.cell(0, 10, "Service Organization: 外部委託先B社",
             align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(15)
    pdf.set_font("YuGoth", "", 11)
    pdf.cell(0, 6, "Report Period: April 1, 2024 - March 31, 2025",
             align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 6, "Issued: June 15, 2025", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(50)
    pdf.cell(0, 6, "Prepared in accordance with SSAE No. 18 (AT-C 320)",
             align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 6, "By: Independent Service Auditor ABC CPA Firm",
             align="C", new_x="LMARGIN", new_y="NEXT")

    pdf.add_page()
    pdf.set_font("YuGoth", "B", 16)
    pdf.cell(0, 10, "Table of Contents", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(5)
    toc = [
        ("Section I", "Independent Service Auditor's Report", 3),
        ("Section II", "Management's Assertion", 5),
        ("Section III", "Description of the Service Organization's System", 7),
        ("Section IV", "Control Objectives, Related Controls, and Test Results", 10),
        ("  IV-1", "Infrastructure Operations", 10),
        ("  IV-2", "Server and Network Maintenance", 12),
        ("  IV-3", "Security Patching", 14),
        ("  IV-4", "Monitoring and Alerting", 16),
        ("  IV-5", "Backup and Recovery", 18),
        ("Section V", "Complementary User Entity Controls (CUECs)", 20),
    ]
    pdf.set_font("YuGoth", "", 11)
    for sec, title, page in toc:
        pdf.set_x(15 if not sec.startswith(" ") else 25)
        pdf.cell(130, 7, f"{sec} {title}")
        pdf.cell(20, 7, f"...  {page}", align="R", new_x="LMARGIN", new_y="NEXT")

    pdf.add_page()
    pdf.h1("Section I. Independent Service Auditor's Report")
    pdf.body("To the Management of 外部委託先B社 and other specified parties:")
    pdf.ln(3)
    pdf.body("Scope: We have examined 外部委託先B社's description of its "
             "IT Infrastructure Management Services (the System) throughout "
             "the period April 1, 2024 to March 31, 2025, and the suitability of "
             "the design and operating effectiveness of the controls.")
    pdf.ln(3)
    pdf.body("Opinion: In our opinion, in all material respects, the description fairly "
             "presents the System, the controls were suitably designed, and the controls "
             "operated effectively throughout the period.")
    pdf.ln(10)
    pdf.cell(0, 6, "ABC CPA Firm", align="R", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 6, "Tokyo, Japan / June 15, 2025", align="R", new_x="LMARGIN", new_y="NEXT")

    pdf.add_page()
    pdf.h1("Section II. Management's Assertion")
    pdf.body("We confirm the description fairly presents the System, controls were suitably designed "
             "and operated effectively.")
    pdf.ln(10)
    pdf.cell(0, 6, "外部委託先B社 Chief Operating Officer",
             align="R", new_x="LMARGIN", new_y="NEXT")

    pdf.add_page()
    pdf.h1("Section III. Description of the System")
    pdf.body("外部委託先B社 provides IT infrastructure management services including data center "
             "operations, server maintenance, network management, security patching, and backup "
             "operations. The company operates 4 Tier-3+ data centers in Japan and serves "
             "approximately 80 corporate customers.")

    pdf.add_page()
    pdf.h1("Section IV. Control Objectives, Related Controls, and Test Results")
    pdf.h2("IV-1. Infrastructure Operations")
    pdf.table_header(["Control Activity", "Test Performed", "Result"], [70, 60, 40])
    pdf.table_row(["24x7 monitoring of critical systems",
                   "Inspected NOC shift logs", "No exceptions"], [70, 60, 40])
    pdf.table_row(["Incident escalation procedures",
                   "Tested 12 incidents", "No exceptions"], [70, 60, 40], fill=True)

    pdf.h2("IV-2. Server and Network Maintenance")
    pdf.table_header(["Control Activity", "Test Performed", "Result"], [70, 60, 40])
    pdf.table_row(["Quarterly patching applied",
                   "Inspected 4 quarterly cycles", "No exceptions"], [70, 60, 40])

    pdf.add_page()
    pdf.h2("IV-3. Security Patching")
    pdf.body("Patches applied within 30 days of vendor release.")
    pdf.h2("IV-4. Monitoring and Alerting")
    pdf.body("Automated alerts on anomalies; response time < 15 minutes for P1/P2.")
    pdf.h2("IV-5. Backup and Recovery")
    pdf.body("Daily backups with quarterly DR testing.")

    pdf.add_page()
    pdf.h1("Section V. Complementary User Entity Controls")
    pdf.body("User entity to verify SLA compliance, validate restore tests, "
             "notify of any system configuration changes affecting backup scope.")

    pdf.output(str(BASE_EM / "ITGC-EM-001_RAW_SOC1_TypeII_B社_FY2024.pdf"))
    print("Created: ITGC-EM-001_RAW_SOC1_TypeII_B社_FY2024.pdf")


if __name__ == "__main__":
    gen_full_r18()
    cleanup_excerpts()
    rename_files()
    regenerate_soc1()
    print("\nAll excerpts fixed.")
