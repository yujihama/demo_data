"""
全統制の「_25件対応_」等の監査センタ的ファイル名を実データ風に統一する
"""
import re
import shutil
from pathlib import Path

BASE = Path(r"C:\Users\nyham\work\demo_data\4.evidence")


# ============================================================
# パターンベース改名（同シリーズの連番ファイル用）
# ============================================================
PATTERN_RULES = [
    # ---- PLC-S ----
    # 注文書 × 25 (PLC-S-001)
    (r"^PLC-S-001_25件対応_注文書_サンプル\d+_(ORD-\d{4}-\d+)(?:_.*?)?\.pdf$",
     r"注文書_\1.pdf"),
    # 請求書 × 12 (PLC-S-003)
    (r"^PLC-S-003_25件対応_請求書サンプル_\d{6}_(INV-\d{6}-\d+)\.pdf$",
     r"請求書_\1.pdf"),
    # SAP請求書バッチログ × 12 (PLC-S-003)
    (r"^PLC-S-003_25件対応_RAW_SAP請求書バッチログ_(\d{6})\.txt$",
     r"SAP_InvoiceBatch_Log_\1.txt"),
    # 月次FB10N × 12 (PLC-S-005)
    (r"^PLC-S-005_25件対応_RAW_SAP売掛金年齢表_FB10N_(\d{6})\.csv$",
     r"SAP_FB10N_AR_Aging_\1.csv"),
    # 価格変更稟議 × 25 (PLC-S-007)
    (r"^PLC-S-007_25件対応_価格変更稟議_サンプル\d+_(W-\d{4}-\d+)\.pdf$",
     r"価格変更稟議_\1.pdf"),
    # ---- PLC-P ----
    # 発注書 × 25 (PLC-P-002)
    (r"^PLC-P-002_25件対応_発注書_サンプル\d+_(PO-\d{4}-\d+)(?:_不備)?\.pdf$",
     r"発注書_\1.pdf"),
    # 検収報告書 × 25 (PLC-P-003)
    (r"^PLC-P-003_25件対応_検収報告書_サンプル\d+_(REC-\d{4}-\d+)\.pdf$",
     r"検収報告書_\1.pdf"),
    # 仕入先マスタ変更申請書 × 5 (PLC-P-005)
    (r"^PLC-P-005_25件対応_仕入先マスタ変更申請書_サンプル\d+_(VEND-CHG-\d{4}-\d+)\.pdf$",
     r"仕入先マスタ変更申請書_\1.pdf"),
    # ---- ITGC ----
    # ユーザ登録申請書 × 5 (ITGC-AC-001)
    (r"^ITGC-AC-001_25件対応_ユーザ登録申請書_サンプル\d+_(USER-REG-\d{4}-\d+)\.pdf$",
     r"ユーザ登録申請書_\1.pdf"),
    # 変更申請書 × 3 (ITGC-CM-001)
    (r"^ITGC-CM-001_25件対応_変更申請書_サンプル\d+_(REL-\d{4}-\d+)\.pdf$",
     r"変更申請書_\1.pdf"),
]


# ============================================================
# 一対一改名（個別ファイル）
# ============================================================
EXPLICIT_RENAMES = {
    # ---- PLC-S ----
    "PLC-S-001_25件対応_RAW_SAP_VA05_受注詳細.csv": "SAP_VA05_SalesOrderDetail_FY2025.csv",
    "PLC-S-001_25件対応_RAW_SAP与信チェックログ.csv": "SAP_CreditCheck_Log_FY2025.csv",
    "PLC-S-001_25件対応_RAW_ワークフロー承認履歴.csv": "Workflow_CreditException_ApprovalHistory_FY2025.csv",
    "PLC-S-001_SAP_VA05_受注伝票一覧_FY2025.xlsx": "SAP_VA05_SalesOrderList_FY2025.xlsx",
    "PLC-S-001_与信限度マスタ_SAP_FD32スナップショット.xlsx": "SAP_FD32_CreditLimitMaster_20260210.xlsx",
    "PLC-S-001_SAP与信超過アラート画面_C-10007.png": "SAP_VA01_CreditExceedAlert_Screen.png",
    "PLC-S-001_SAP受注登録画面_ORD-2025-1420.png": "SAP_VA03_SalesOrder_Screen_ORD-2025-1420.png",
    "PLC-S-001_ワークフロー承認_与信超過サンプル.png": "Workflow_CreditException_Approval_Screen.png",

    # PLC-S-002 (既に改名済みだが、まだ残っているものもある)
    "PLC-S-002_25件対応_RAW_WMS出荷実績エクスポート.csv": "WMS_ShipmentActual_Export_FY2025Samples.csv",
    "PLC-S-002_25件対応_RAW_SAP売上計上仕訳_FBL3N.csv": "SAP_FBL3N_SalesJournal_Detail_FY2025Samples.csv",
    "PLC-S-002_25件対応_RAW_SAPマッチングバッチログ.csv": "SAP_ShipSalesMatch_BatchLog_FY2025Samples.csv",
    "PLC-S-002_25件対応_RAW_例外サンプル14_SAPバッチログ_SH-202510-0068.txt":
        "SAP_BatchLog_ZSD_SHIP_SALES_MATCH_SH-202510-0068.txt",
    "PLC-S-002_25件対応_RAW_例外サンプル9_SAP_VA02変更履歴_SH-202508-0197.txt":
        "SAP_VA02_ChangeHistory_ORD-2025-0989.txt",
    "PLC-S-002_監査対象25件サンプルリスト.xlsx":
        "SalesShipMatch_SampleTransactionList_FY2025.xlsx",
    "PLC-S-002_SAP未マッチ明細リスト_202511.csv": "SAP_ZSD_UNMATCH_List_202511.csv",
    "PLC-S-002_SAP売上計上明細_202511.csv": "SAP_FBL3N_SalesJournal_Detail_202511.csv",
    "PLC-S-002_WMS出荷実績エクスポート_202511.csv": "WMS_ShipmentActual_Export_202511.csv",

    # PLC-S-003
    "PLC-S-003_月次請求書発行一覧_202511.xlsx": "SAP_VF05_InvoiceList_202511.xlsx",  # 存在しないが一応定義
    "PLC-S-003_25件対応_RAW_SAP請求書一覧_FY2025.csv": "SAP_VF05_InvoiceRegister_FY2025.csv",
    "PLC-S-003_SAP請求書バッチ実行ログ_202511.txt": "SAP_InvoiceBatch_Log_202511.txt",  # 存在しないが一応
    "PLC-S-003_請求書_INV-202511-0234.pdf": "請求書_INV-202511-0234.pdf",

    # PLC-S-004
    "PLC-S-004_監査対象25件サンプルリスト.xlsx":
        "Payment_Clearing_SampleTransactionList_FY2025.xlsx",
    "PLC-S-004_25件対応_RAW_SAP入金消込履歴_F-28.csv":
        "SAP_F-28_PaymentClearing_History_FY2025Samples.csv",
    "PLC-S-004_25件対応_RAW_FB入金データ_25件抽出.csv":
        "Bank_FB_PaymentReceipt_Data_FY2025Samples.csv",
    "PLC-S-004_FB入金データ_202511.csv": "Bank_FB_PaymentReceipt_Data_202511.csv",
    "PLC-S-004_SAP入金消込画面.png": "SAP_F-28_PaymentClearing_Screen.png",
    "PLC-S-004_入金消込リスト_202511.xlsx": "SAP_F-28_ClearingList_202511.xlsx",  # 存在しない

    # PLC-S-005
    # 個別月次FB10N (パターンで処理済)

    # PLC-S-006
    "PLC-S-006_RAW_SAP期末前後出荷売上明細_FY2025期末.csv":
        "SAP_YearEndCutoff_ShipSales_Detail_FY2025.csv",

    # PLC-S-007
    "PLC-S-007_監査対象25件サンプルリスト.xlsx":
        "PriceMaster_Change_SampleTransactionList_FY2025.xlsx",
    "PLC-S-007_25件対応_RAW_SAP_VK12変更履歴.csv":
        "SAP_VK12_PriceCondition_ChangeHistory_FY2025.csv",

    # ---- PLC-P ----
    "PLC-P-001_SAP購買依頼一覧_202511.xlsx": "SAP_ME5A_PurchaseRequisitionList_202511.xlsx",
    "PLC-P-001_監査対象25件サンプルリスト.xlsx":
        "PurchaseRequisition_SampleTransactionList_FY2025.xlsx",
    "PLC-P-001_25件対応_RAW_SAP_ME5A_購買依頼詳細.csv":
        "SAP_ME5A_PurchaseRequisition_Detail_FY2025.csv",

    "PLC-P-002_SAP_ME2N_発注伝票一覧_FY2025.xlsx": "SAP_ME2N_PurchaseOrderList_FY2025.xlsx",
    "PLC-P-002_監査対象25件サンプルリスト.xlsx":
        "PurchaseOrderApproval_SampleTransactionList_FY2025.xlsx",
    "PLC-P-002_25件対応_RAW_SAP_ME2N_発注詳細.csv":
        "SAP_ME2N_PurchaseOrder_Detail_FY2025Samples.csv",
    "PLC-P-002_25件対応_RAW_SAPワークフロー承認履歴.csv":
        "SAP_WF_PurchaseOrder_ApprovalHistory_FY2025Samples.csv",
    "PLC-P-002_SAPワークフロー承認履歴ログ_FY2025.csv":
        "SAP_WF_PurchaseOrder_ApprovalHistory_FY2025.csv",
    "PLC-P-002_SAP_WF_PurchaseOrder_ApprovalHistory_FY2025.csv":  # 万一重複時
        "SAP_WF_PurchaseOrder_ApprovalHistory_FY2025.csv",
    "PLC-P-002_SAP発注画面_不備ケースPO-2025-1456.png":
        "SAP_ME23N_Screen_PO-2025-1456.png",
    "PLC-P-002_SAP発注登録画面_PO-2025-2560.png":
        "SAP_ME23N_Screen_PO-2025-2560.png",
    "PLC-P-002_ワークフロー承認画面_通常案件.png": "Workflow_PurchaseOrder_Approval_Screen.png",

    "PLC-P-003_監査対象25件サンプルリスト.xlsx":
        "GoodsReceipt_SampleTransactionList_FY2025.xlsx",
    "PLC-P-003_25件対応_RAW_SAP_MIGO_検収詳細.csv":
        "SAP_MIGO_GoodsReceipt_Detail_FY2025Samples.csv",

    "PLC-P-004_監査対象25件サンプルリスト.xlsx":
        "ThreeWayMatch_SampleTransactionList_FY2025.xlsx",
    "PLC-P-004_25件対応_RAW_SAP_MIRO_3wayマッチング結果.csv":
        "SAP_MIRO_InvoiceVerification_FY2025Samples.csv",

    "PLC-P-005_監査対象25件サンプルリスト.xlsx":
        "VendorMaster_Change_SampleTransactionList_FY2025.xlsx",
    "PLC-P-005_25件対応_RAW_SAP_XK01_XK02_仕入先マスタ変更履歴.csv":
        "SAP_XK01_XK02_VendorMaster_ChangeHistory_FY2025.csv",

    "PLC-P-006_全12ヶ月RAW_SAP_F110_支払実行バッチ.csv":
        "SAP_F110_AutomaticPayment_RunLog_FY2025.csv",

    "PLC-P-007_全87件RAW_SAP期末未払計上明細.csv":
        "SAP_YearEndAccrual_Detail_20260331.csv",

    # ---- PLC-I ----
    "PLC-I-001_RAW_SAP_MI07_棚卸差異一覧_2025年9月.csv":
        "SAP_MI07_InventoryCountDifference_202509.csv",
    "PLC-I-001_棚卸写真_本社倉庫A_区画A-3.jpg": "棚卸写真_本社倉庫A_区画A-3.jpg",
    "PLC-I-001_棚卸写真_本社倉庫A_区画A-7_立会.jpg": "棚卸写真_本社倉庫A_区画A-7.jpg",
    "PLC-I-001_棚卸写真_本社倉庫B_区画B-3_差異発生区画.jpg": "棚卸写真_本社倉庫B_区画B-3.jpg",
    "PLC-I-001_棚卸写真_東北工場倉庫_区画T-1.jpg": "棚卸写真_東北工場倉庫_区画T-1.jpg",
    "PLC-I-001_SAP在庫数量一覧_MB52.png": "SAP_MB52_InventoryList_Screen.png",

    "PLC-I-002_全24件RAW_SAP_MIGO_棚卸差異調整仕訳.csv":
        "SAP_MIGO_InventoryAdjustment_Entries_202509.csv",

    "PLC-I-003_標準原価更新稟議_W-2025-0089.pdf": "標準原価更新稟議_W-2025-0089.pdf",

    "PLC-I-004_RAW_SAP_CO88_原価差異計算結果_2025年7月_10月_2026年1月.csv":
        "SAP_CO88_VarianceSettlement_FY2025_Quarterly.csv",

    "PLC-I-005_全4四半期RAW_SAP_MB52_滞留在庫評価損計算結果.csv":
        "SAP_MB52_SlowMovingImpairment_FY2025_Quarterly.csv",

    "PLC-I-006_WMS-ERP在庫照合レポート_202511_日次照合30日分.csv":
        "WMS_SAP_InventoryReconciliation_202511.csv",

    "PLC-I-007_全12ヶ月RAW_SAP_MMPV_原価計算締めログ.csv":
        "SAP_MMPV_PeriodClose_Log_FY2025.csv",

    # ---- ITGC ----
    "ITGC-AC-001_ユーザ登録申請書_USER-REG-2025-0087.pdf":
        "ユーザ登録申請書_USER-REG-2025-0087.pdf",
    "ITGC-AC-001_監査対象25件サンプルリスト.xlsx":
        "UserRegistration_SampleTransactionList_FY2025.xlsx",
    "ITGC-AC-001_25件対応_RAW_SAP_SU01_ユーザ作成履歴.csv":
        "SAP_SU01_UserMaster_ChangeHistory_FY2025.csv",
    "ITGC-AC-001_25件対応_RAW_ワークフロー承認履歴.csv":
        "Workflow_UserRegistration_ApprovalHistory_FY2025.csv",
    "ITGC-AC-001_SAP_SU01_ユーザ作成画面.png": "SAP_SU01_UserCreate_Screen.png",
    "ITGC-AC-001_SAPアクセス権マトリクス.png": "SAP_AccessRights_Matrix.png",

    "ITGC-AC-002_SAP_SUIM_有効ユーザ一覧_Q3棚卸用.xlsx":
        "SAP_SUIM_ActiveUserList_2025Q3.xlsx",

    "ITGC-AC-003_SAP_SM19_SM20_退職者ログインログ抽出.csv":
        "SAP_SM20_SecurityAuditLog_RetiredUsers.csv",

    "ITGC-AC-004_特権ID操作ログ_202511.csv":
        "SAP_SM20_PrivilegedUser_OperationLog_202511.csv",

    "ITGC-CM-001_監査対象25件サンプルリスト.xlsx":
        "ChangeManagement_SampleSubmissionList_FY2025.xlsx",
    "ITGC-CM-001_変更管理一覧_FY2025.xlsx":
        "ChangeManagement_Register_FY2025.xlsx",
    "ITGC-CM-001_変更申請書_REL-2025-023.pdf": "変更申請書_REL-2025-023.pdf",
    "ITGC-CM-001_25件対応_RAW_変更管理台帳.csv":
        "ChangeManagement_Register_Detailed_FY2025.csv",

    "ITGC-CM-003_25件対応_RAW_SAP_STMS_本番移送履歴.csv":
        "SAP_STMS_ProductionTransport_History_FY2025.csv",
    "ITGC-CM-003_SAP_STMS_本番移送記録_FY2025Q2-Q3.csv":
        "SAP_STMS_ProductionTransport_History_FY2025Q2-Q3.csv",

    "ITGC-OM-001_25件対応_RAW_SAP_DB13_バックアップログ.csv":
        "SAP_DB13_DatabaseBackup_Log_FY2025.csv",
    "ITGC-OM-001_DRリストアテスト報告書_2025Q3.pdf":
        "DR_RestoreTest_Report_2025Q3.pdf",

    "ITGC-OM-002_全18件RAW_監視ツール障害検知ログ.csv":
        "Zabbix_IncidentDetection_Log_FY2025.csv",

    "ITGC-EM-001_RAW_SOC1_TypeII_SIerA_FY2024.pdf":
        "SOC1_TypeII_Report_SIerA_FY2024.pdf",
    "ITGC-EM-001_RAW_SOC1_TypeII_B社_FY2024.pdf":
        "SOC1_TypeII_Report_B社_FY2024.pdf",

    # ---- ITAC ----
    "ITAC-001_SAP与信限度自動チェック設定画面_OVAK.png":
        "SAP_OVAK_CreditCheckConfig_Screen.png",
    "ITAC-002_SAP3wayマッチング設定画面_OMRK.png":
        "SAP_OMRK_InvoiceMatchingConfig_Screen.png",
    "ITAC-002_3wayマッチング結果ログ_202511.csv":
        "SAP_MIRO_3WayMatch_ResultLog_202511.csv",
    "ITAC-003_SAP減価償却実行画面_AFAB.png": "SAP_AFAB_DepreciationRun_Screen.png",

    # ---- ELC ----
    "ELC-001_取締役会議事録_第245回_2025年9月.pdf":
        "取締役会議事録_第245回_2025年9月.pdf",
    "ELC-002_RAW_HRシステム_倫理綱領受領確認ログ.csv":
        "HRIS_CodeOfEthics_AcknowledgmentLog_FY2025.csv",
    "ELC-004_全社リスクアセスメント結果_2025年度.xlsx":
        "全社リスクアセスメント結果_2025年度.xlsx",
    "ELC-008_内部通報受付台帳_FY2025.xlsx": "内部通報受付台帳_FY2025.xlsx",
    "ELC-010_2025年度内部監査計画書.pdf": "2025年度内部監査計画書.pdf",

    # ---- FCRP ----
    "FCRP-001_全12ヶ月RAW_SAP_FB50_月次決算ジョブログ.csv":
        "SAP_PeriodClose_JobLog_FY2025.csv",
    "FCRP-002_全4四半期RAW_連結システムS05_パッケージ受信ログ.csv":
        "ConsolidationSystem_PackageUpload_Log_FY2025.csv",
    "FCRP-003_全4四半期RAW_SAP_FB10N_貸倒引当金算定データ.csv":
        "SAP_FB10N_BadDebtImpairment_CalcData_FY2025.csv",
    "FCRP-004_全4四半期RAW_連結システムS05_連結仕訳一覧.csv":
        "ConsolidationSystem_Entries_FY2025.csv",
    "FCRP-005_全4四半期RAW_開示システムS06_XBRL検証ログ.csv":
        "DisclosureSystem_XBRL_ValidationLog_FY2025.csv",
}


def apply_renames():
    renamed_count = 0
    for rcm_dir in ["ELC", "PLC-S", "PLC-P", "PLC-I", "ITGC", "ITAC", "FCRP"]:
        folder = BASE / rcm_dir
        if not folder.exists():
            continue
        files = sorted(folder.iterdir())
        for f in files:
            if not f.is_file():
                continue
            old_name = f.name
            new_name = None

            # 1. パターン規則
            for pattern, replacement in PATTERN_RULES:
                if re.match(pattern, old_name):
                    new_name = re.sub(pattern, replacement, old_name)
                    break

            # 2. 明示的規則
            if not new_name and old_name in EXPLICIT_RENAMES:
                new_name = EXPLICIT_RENAMES[old_name]

            # 変更なし
            if not new_name or new_name == old_name:
                continue

            # 実行
            dst = folder / new_name
            if dst.exists() and dst != f:
                # 同じ内容の可能性があるのでold削除
                print(f"  Duplicate target, removing old: {rcm_dir}/{old_name}")
                f.unlink()
                continue
            f.rename(dst)
            print(f"  {rcm_dir}/{old_name}")
            print(f"    -> {new_name}")
            renamed_count += 1

    print(f"\n  Renamed {renamed_count} files")


if __name__ == "__main__":
    print("=== Rename all audit-centric filenames to realistic data names ===")
    apply_renames()
    print("\nDone.")
