"""
改名後のファイルを元に新マッピングCSVを生成
新方針：ファイル名に統制IDが含まれない実データ風命名のため、
明示的な統制ID↔ファイル名マッピングを構築
"""
import csv
from pathlib import Path

BASE = Path(r"C:\Users\nyham\work\demo_data")
EVIDENCE = BASE / "4.evidence"
OUTPUT_CSV = BASE / "2.RCM" / "統制_エビデンスマッピング.csv"


# 統制 ID → 区分フォルダ
CONTROL_TO_DIR = {}
for cid in [f"ELC-{i:03d}" for i in range(1, 13)]:
    CONTROL_TO_DIR[cid] = "ELC"
for cid in [f"PLC-S-{i:03d}" for i in range(1, 8)]:
    CONTROL_TO_DIR[cid] = "PLC-S"
for cid in [f"PLC-P-{i:03d}" for i in range(1, 8)]:
    CONTROL_TO_DIR[cid] = "PLC-P"
for cid in [f"PLC-I-{i:03d}" for i in range(1, 8)]:
    CONTROL_TO_DIR[cid] = "PLC-I"
for cid in [f"ITGC-AC-{i:03d}" for i in range(1, 5)]:
    CONTROL_TO_DIR[cid] = "ITGC"
for cid in [f"ITGC-CM-{i:03d}" for i in range(1, 4)]:
    CONTROL_TO_DIR[cid] = "ITGC"
for cid in [f"ITGC-OM-{i:03d}" for i in range(1, 3)]:
    CONTROL_TO_DIR[cid] = "ITGC"
CONTROL_TO_DIR["ITGC-EM-001"] = "ITGC"
for cid in [f"ITAC-{i:03d}" for i in range(1, 6)]:
    CONTROL_TO_DIR[cid] = "ITAC"
for cid in [f"FCRP-{i:03d}" for i in range(1, 6)]:
    CONTROL_TO_DIR[cid] = "FCRP"


# 統制 → エビデンスファイル（実データ風命名）
CONTROL_EVIDENCE = {
    # ===== ELC =====
    "ELC-001": ["取締役会議事録_第245回_2025年9月.pdf"],
    "ELC-002": ["HRIS_CodeOfEthics_AcknowledgmentLog_FY2025.csv"],
    "ELC-003": [
        "規程_職務権限規程_R18.pdf",
        "employees.xlsx",
        "user_roles_matrix.xlsx",
        "company_profile.md",
    ],
    "ELC-004": ["全社リスクアセスメント結果_2025年度.xlsx"],
    "ELC-005": ["全社リスクアセスメント結果_2025年度.xlsx"],
    "ELC-006": ["company_profile.md"],
    "ELC-007": [
        "user_roles_matrix.xlsx",
        "発注書_PO-2025-0234.pdf",
        "発注書_PO-2025-0789.pdf",
        "発注書_PO-2025-1456.pdf",
    ],
    "ELC-008": ["内部通報受付台帳_FY2025.xlsx"],
    "ELC-009": ["SAP_PeriodClose_JobLog_FY2025.csv"],
    "ELC-010": ["2025年度内部監査計画書.pdf"],
    "ELC-011": [
        "取締役会議事録_第245回_2025年9月.pdf",
        "内部通報受付台帳_FY2025.xlsx",
    ],
    "ELC-012": [
        "company_profile.md",
        "SOC1_TypeII_Report_SIerA_FY2024.pdf",
    ],

    # ===== PLC-S =====
    "PLC-S-001": [
        "SAP_VA05_SalesOrderList_FY2025.xlsx",
        "SAP_VA05_SalesOrderDetail_FY2025.csv",
        "SAP_FD32_CreditLimitMaster_20260210.xlsx",
        "SAP_CreditCheck_Log_FY2025.csv",
        "Workflow_CreditException_ApprovalHistory_FY2025.csv",
        "SAP_VA03_SalesOrder_Screen_ORD-2025-1420.png",
        "SAP_VA01_CreditExceedAlert_Screen.png",
        "Workflow_CreditException_Approval_Screen.png",
        "規程_職務権限規程_R18.pdf",
        # 25件の注文書は後でスキャンで追加
    ],
    "PLC-S-002": [
        "WMS_ShipmentActual_Export_202511.csv",
        "SAP_FBL3N_SalesJournal_Detail_202511.csv",
        "SAP_ZSD_UNMATCH_List_202511.csv",
        "WMS_ShipmentActual_Export_FY2025Samples.csv",
        "SAP_FBL3N_SalesJournal_Detail_FY2025Samples.csv",
        "SAP_ShipSalesMatch_BatchLog_FY2025Samples.csv",
        "SAP_BatchLog_ZSD_SHIP_SALES_MATCH_SH-202510-0068.txt",
        "SAP_VA02_ChangeHistory_ORD-2025-0989.txt",
        "SalesShipMatch_SampleTransactionList_FY2025.xlsx",
    ],
    "PLC-S-003": [
        "SAP_VF05_InvoiceRegister_FY2025.csv",
        # 月次請求書バッチログ × 12 (SAP_InvoiceBatch_Log_YYYYMM.txt)
        # 月次請求書PDF × 12 (請求書_INV-YYYYMM-XXXX.pdf)
    ],
    "PLC-S-004": [
        "Payment_Clearing_SampleTransactionList_FY2025.xlsx",
        "SAP_F-28_PaymentClearing_History_FY2025Samples.csv",
        "Bank_FB_PaymentReceipt_Data_FY2025Samples.csv",
        "Bank_FB_PaymentReceipt_Data_202511.csv",
        "SAP_F-28_PaymentClearing_Screen.png",
    ],
    "PLC-S-005": [
        # 月次 FB10N × 12 (SAP_FB10N_AR_Aging_YYYYMM.csv)
    ],
    "PLC-S-006": ["SAP_YearEndCutoff_ShipSales_Detail_FY2025.csv"],
    "PLC-S-007": [
        "PriceMaster_Change_SampleTransactionList_FY2025.xlsx",
        "SAP_VK12_PriceCondition_ChangeHistory_FY2025.csv",
        # 価格変更稟議 × 25 (価格変更稟議_W-YYYY-XXXX.pdf)
    ],

    # ===== PLC-P =====
    "PLC-P-001": [
        "SAP_ME5A_PurchaseRequisitionList_202511.xlsx",
        "PurchaseRequisition_SampleTransactionList_FY2025.xlsx",
        "SAP_ME5A_PurchaseRequisition_Detail_FY2025.csv",
    ],
    "PLC-P-002": [
        "PurchaseOrderApproval_SampleTransactionList_FY2025.xlsx",
        "SAP_ME2N_PurchaseOrder_Detail_FY2025Samples.csv",
        "SAP_WF_PurchaseOrder_ApprovalHistory_FY2025Samples.csv",
        "SAP_WF_PurchaseOrder_ApprovalHistory_FY2025.csv",
        "SAP_ME23N_Screen_PO-2025-2560.png",
        "SAP_ME23N_Screen_PO-2025-1456.png",
        "Workflow_PurchaseOrder_Approval_Screen.png",
        "規程_職務権限規程_R18.pdf",
        # 発注書 × 25 (発注書_PO-YYYY-XXXX.pdf)
    ],
    "PLC-P-003": [
        "GoodsReceipt_SampleTransactionList_FY2025.xlsx",
        "SAP_MIGO_GoodsReceipt_Detail_FY2025Samples.csv",
        # 検収報告書 × 25 (検収報告書_REC-YYYY-XXXX.pdf)
    ],
    "PLC-P-004": [
        "ThreeWayMatch_SampleTransactionList_FY2025.xlsx",
        "SAP_MIRO_InvoiceVerification_FY2025Samples.csv",
    ],
    "PLC-P-005": [
        "VendorMaster_Change_SampleTransactionList_FY2025.xlsx",
        "SAP_XK01_XK02_VendorMaster_ChangeHistory_FY2025.csv",
        # 代表5件の仕入先マスタ変更申請書
    ],
    "PLC-P-006": ["SAP_F110_AutomaticPayment_RunLog_FY2025.csv"],
    "PLC-P-007": ["SAP_YearEndAccrual_Detail_20260331.csv"],

    # ===== PLC-I =====
    "PLC-I-001": [
        "SAP_MI07_InventoryCountDifference_202509.csv",
        "SAP_MB52_InventoryList_Screen.png",
        "棚卸写真_本社倉庫A_区画A-3.jpg",
        "棚卸写真_本社倉庫A_区画A-7.jpg",
        "棚卸写真_本社倉庫B_区画B-3.jpg",
        "棚卸写真_東北工場倉庫_区画T-1.jpg",
    ],
    "PLC-I-002": ["SAP_MIGO_InventoryAdjustment_Entries_202509.csv"],
    "PLC-I-003": ["標準原価更新稟議_W-2025-0089.pdf"],
    "PLC-I-004": ["SAP_CO88_VarianceSettlement_FY2025_Quarterly.csv"],
    "PLC-I-005": ["SAP_MB52_SlowMovingImpairment_FY2025_Quarterly.csv"],
    "PLC-I-006": ["WMS_SAP_InventoryReconciliation_202511.csv"],
    "PLC-I-007": ["SAP_MMPV_PeriodClose_Log_FY2025.csv"],

    # ===== ITGC =====
    "ITGC-AC-001": [
        "UserRegistration_SampleTransactionList_FY2025.xlsx",
        "SAP_SU01_UserMaster_ChangeHistory_FY2025.csv",
        "Workflow_UserRegistration_ApprovalHistory_FY2025.csv",
        "SAP_SU01_UserCreate_Screen.png",
        "SAP_AccessRights_Matrix.png",
        "ユーザ登録申請書_USER-REG-2025-0087.pdf",
        "規程_職務権限規程_R18.pdf",
        # ユーザ登録申請書 × 5
    ],
    "ITGC-AC-002": [
        "SAP_SUIM_ActiveUserList_2025Q1.xlsx",
        "SAP_SUIM_ActiveUserList_2025Q2.xlsx",
        "SAP_SUIM_ActiveUserList_2025Q3.xlsx",
        "SAP_SUIM_ActiveUserList_2025Q4.xlsx",
    ],
    "ITGC-AC-003": ["SAP_SM20_SecurityAuditLog_RetiredUsers.csv"],
    "ITGC-AC-004": ["SAP_SM20_PrivilegedUser_OperationLog_202511.csv"],
    "ITGC-CM-001": [
        "ChangeManagement_SampleSubmissionList_FY2025.xlsx",
        "ChangeManagement_Register_FY2025.xlsx",
        "ChangeManagement_Register_Detailed_FY2025.csv",
        "変更申請書_REL-2025-023.pdf",
        # 変更申請書 × 3 追加
    ],
    "ITGC-CM-002": [
        "Xray_TestExecution_History_FY2025.csv",
        # UATテスト結果_REL-2025-XXX.xlsx × 25
    ],
    "ITGC-CM-003": [
        "SAP_STMS_ProductionTransport_History_FY2025.csv",
        "SAP_STMS_ProductionTransport_History_FY2025Q2-Q3.csv",
    ],
    "ITGC-OM-001": [
        "SAP_DB13_DatabaseBackup_Log_FY2025.csv",
        "DR_RestoreTest_Report_2025Q3.pdf",
    ],
    "ITGC-OM-002": ["Zabbix_IncidentDetection_Log_FY2025.csv"],
    "ITGC-EM-001": [
        "SOC1_TypeII_Report_SIerA_FY2024.pdf",
        "SOC1_TypeII_Report_B社_FY2024.pdf",
    ],

    # ===== ITAC =====
    "ITAC-001": [
        "SAP_OVAK_CreditCheckConfig_Screen.png",
    ],
    "ITAC-002": [
        "SAP_OMRK_InvoiceMatchingConfig_Screen.png",
        "SAP_MIRO_3WayMatch_ResultLog_202511.csv",
    ],
    "ITAC-003": ["SAP_AFAB_DepreciationRun_Screen.png"],
    "ITAC-004": [
        "SAP_WF_PurchaseOrder_ApprovalHistory_FY2025.csv",
        "SAP_WF_PurchaseOrder_ApprovalHistory_FY2025Samples.csv",
        "ChangeManagement_Register_Detailed_FY2025.csv",
    ],
    "ITAC-005": ["ConsolidationSystem_PackageUpload_Log_FY2025.csv"],

    # ===== FCRP =====
    "FCRP-001": ["SAP_PeriodClose_JobLog_FY2025.csv"],
    "FCRP-002": ["ConsolidationSystem_PackageUpload_Log_FY2025.csv"],
    "FCRP-003": ["SAP_FB10N_BadDebtImpairment_CalcData_FY2025.csv"],
    "FCRP-004": ["ConsolidationSystem_Entries_FY2025.csv"],
    "FCRP-005": ["DisclosureSystem_XBRL_ValidationLog_FY2025.csv"],
}


# パターンベースで追加取得するファイル（連番系）
PATTERN_EVIDENCE = {
    # 統制ID, ファイル名パターン（glob）
    "PLC-S-001": ["注文書_ORD-*.pdf"],
    "PLC-S-003": [
        "SAP_InvoiceBatch_Log_*.txt",
        "請求書_INV-*.pdf",
    ],
    "PLC-S-005": ["SAP_FB10N_AR_Aging_*.csv"],
    "PLC-S-007": ["価格変更稟議_W-*.pdf"],
    "PLC-P-002": ["発注書_PO-*.pdf"],
    "PLC-P-003": ["検収報告書_REC-*.pdf"],
    "PLC-P-005": ["仕入先マスタ変更申請書_VEND-CHG-*.pdf"],
    "ITGC-AC-001": ["ユーザ登録申請書_USER-REG-*.pdf"],
    "ITGC-CM-001": ["変更申請書_REL-*.pdf"],
    "ITGC-CM-002": ["UATテスト結果_REL-*.xlsx"],
}


def resolve_files():
    """各統制のエビデンスを実際のファイルから解決し、cross-ref もコピー"""
    import shutil
    mappings = set()

    for cid, files in CONTROL_EVIDENCE.items():
        rcm = CONTROL_TO_DIR.get(cid)
        target_dir = EVIDENCE / rcm
        for filename in files:
            # ファイルが存在すれば追加
            target = target_dir / filename
            if target.exists():
                mappings.add((cid, filename))
                continue
            # cross-ref候補: 0.profile/ や 1.master_data/ からコピー
            candidates = [
                BASE / "0.profile" / filename,
                BASE / "1.master_data" / filename,
            ]
            copied = False
            for candidate in candidates:
                if candidate.exists():
                    shutil.copy2(candidate, target)
                    mappings.add((cid, filename))
                    copied = True
                    break
            if not copied:
                # 他区分フォルダから探してコピー
                for other_rcm in ["ELC", "PLC-S", "PLC-P", "PLC-I", "ITGC", "ITAC", "FCRP"]:
                    src = EVIDENCE / other_rcm / filename
                    if src.exists() and src != target:
                        shutil.copy2(src, target)
                        mappings.add((cid, filename))
                        copied = True
                        break
            if not copied:
                print(f"  NOT FOUND: {cid} / {filename}")

    # パターンベース
    for cid, patterns in PATTERN_EVIDENCE.items():
        rcm = CONTROL_TO_DIR.get(cid)
        folder = EVIDENCE / rcm
        for pat in patterns:
            for f in folder.glob(pat):
                if f.is_file():
                    mappings.add((cid, f.name))

    return sorted(mappings, key=lambda x: (x[0], x[1]))


def write_csv(mappings):
    with open(OUTPUT_CSV, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f, quoting=csv.QUOTE_ALL)
        writer.writerow(["統制ID", "Filename"])
        for row in mappings:
            writer.writerow(row)

    # 検証
    missing = []
    for cid, fname in mappings:
        rcm = CONTROL_TO_DIR.get(cid)
        if not (EVIDENCE / rcm / fname).exists():
            missing.append((cid, fname, rcm))

    print(f"\n  Total mappings: {len(mappings)}")
    print(f"  Unique filenames: {len(set(m[1] for m in mappings))}")
    covered = set(m[0] for m in mappings)
    print(f"  Covered controls: {len(covered)}/{len(CONTROL_TO_DIR)}")
    uncovered = [c for c in CONTROL_TO_DIR if c not in covered]
    if uncovered:
        print(f"  Uncovered: {uncovered}")
    if missing:
        print(f"  WARNING: {len(missing)} missing files!")
        for m in missing[:10]:
            print(f"    - {m[0]} / {m[1]} (expected in {m[2]}/)")
    else:
        print("  [OK] All CSV entries map to existing files")


if __name__ == "__main__":
    mappings = resolve_files()
    write_csv(mappings)
