"""
Phase 1: 全統制に対して新方針を適用するための削除・簡素化

削除対象：統制実施者（経理部・倉庫課・情シス部等）の確認・レビュー・承認記録
残す：RAWデータ、業務上自然発生する原本書類（議事録・計画書・契約書・稟議書・申請書等）
"""
from pathlib import Path

BASE = Path(r"C:\Users\nyham\work\demo_data\4.evidence")

# 削除対象ファイル
DELETE_TARGETS = [
    # ===== PLC-S =====
    # 既存の経理部確認記録系
    "PLC-S/PLC-S-003_月次請求書発行一覧_202511.xlsx",
    "PLC-S/PLC-S-004_入金消込リスト_202511.xlsx",
    "PLC-S/PLC-S-005_売掛金年齢表_202511.xlsx",
    "PLC-S/PLC-S-005_売掛金年齢表_経理部長承認PDF_低解像度.pdf",
    "PLC-S/PLC-S-006_期末カットオフテスト.xlsx",
    "PLC-S/PLC-S-007_価格変更履歴レポート_Q3.xlsx",
    "PLC-S/PLC-S_月次売上会議_議事録_202511.md",

    # ===== PLC-P =====
    "PLC-P/PLC-P-004_3wayマッチング結果_202511.xlsx",
    "PLC-P/PLC-P-006_支払予定一覧_202511.xlsx",
    "PLC-P/PLC-P-007_期末未払計上リスト.xlsx",

    # ===== PLC-I =====
    "PLC-I/PLC-I-001_実地棚卸計画書_2025下期.pdf",  # 計画書は経理部長承認あり
    "PLC-I/PLC-I-001_実地棚卸報告書_2025年9月.xlsx",
    "PLC-I/PLC-I-002_棚卸差異分析書_INV-DIFF-2025-09-012.pdf",
    "PLC-I/PLC-I-004_原価差異分析表_202511.xlsx",
    "PLC-I/PLC-I-005_滞留在庫評価損計算_2025年12月末.xlsx",
    "PLC-I/PLC-I-007_月次原価計算締めチェックリスト_202511.xlsx",

    # ===== ITGC =====
    "ITGC/AC_アクセス管理/ITGC-AC-003_退職者アカウント停止記録_FY2025.xlsx",
    "ITGC/OM_運用管理/ITGC-OM-001_バックアップ実施記録_202511.xlsx",
    "ITGC/OM_運用管理/ITGC-OM-002_障害管理台帳_FY2025.xlsx",
    "ITGC/EM_外部委託管理/ITGC-EM-001_SOC1レポート評価レビューシート_2025.pdf",
    "ITGC/EM_外部委託管理/ITGC-EM-001_IT外部委託先一覧_FY2025.xlsx",

    # ===== ITAC =====
    "ITAC/ITAC-001_与信限度自動チェック_動作検証.xlsx",
    "ITAC/ITAC-003_減価償却手計算検証.xlsx",

    # ===== ELC =====
    "ELC/ELC-002_倫理綱領受領確認書提出状況_2025年度.xlsx",

    # ===== FCRP =====
    "FCRP/FCRP-001_月次決算チェックリスト_202511.xlsx",
    "FCRP/FCRP-002_連結パッケージ受領管理_2025Q3.xlsx",
    "FCRP/FCRP-003_貸倒引当金計算シート_2025年12月末.xlsx",
    "FCRP/FCRP-005_開示書類レビューシート_2026年3月期Q3.pdf",
]


def main():
    deleted = 0
    not_found = 0
    for rel in DELETE_TARGETS:
        p = BASE / rel
        if p.exists():
            p.unlink()
            print(f"Deleted: {rel}")
            deleted += 1
        else:
            print(f"Not found (skip): {rel}")
            not_found += 1
    print(f"\nTotal deleted: {deleted} / Not found: {not_found}")


if __name__ == "__main__":
    main()
