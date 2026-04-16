"""
AIエージェント評価用の再整理
- 4.evidence/{区分}/ 配下をフラット化
- クロス参照ファイルを該当区分フォルダにコピー
- マッピングCSVを (統制ID, Filename) 2列に整理
"""
import shutil
import csv
from pathlib import Path

BASE = Path(r"C:\Users\nyham\work\demo_data")
EVIDENCE = BASE / "4.evidence"
OUTPUT_CSV = BASE / "2.RCM" / "統制_エビデンスマッピング.csv"


# ============================================================
# 統制一覧（統制ID → 区分フォルダ名）
# ============================================================
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


# ============================================================
# クロス参照定義（現時点のファイル状態を反映）
# ============================================================
CROSS_REF = {
    "ELC-003": [
        "0.profile/規程_職務権限規程_R18.pdf",
        "1.master_data/employees.xlsx",
        "1.master_data/user_roles_matrix.xlsx",
        "0.profile/company_profile.md",
    ],
    "ELC-005": [
        # 全社リスクアセスメントに不正リスクも含まれる
        "4.evidence/ELC/ELC-004_全社リスクアセスメント結果_2025年度.xlsx",
    ],
    "ELC-006": ["0.profile/company_profile.md"],
    "ELC-011": [
        # 監査等委員会のモニタリング記録は取締役会議事録・内部通報台帳等で確認
        "4.evidence/ELC/ELC-001_取締役会議事録_第245回_2025年9月.pdf",
        "4.evidence/ELC/ELC-008_内部通報受付台帳_FY2025.xlsx",
    ],
    "ELC-007": [
        "1.master_data/user_roles_matrix.xlsx",
        "4.evidence/PLC-P/PLC-P-002_25件対応_発注書_サンプル23_PO-2025-0234_不備.pdf",
        "4.evidence/PLC-P/PLC-P-002_25件対応_発注書_サンプル24_PO-2025-0789_不備.pdf",
        "4.evidence/PLC-P/PLC-P-002_25件対応_発注書_サンプル25_PO-2025-1456_不備.pdf",
    ],
    "ELC-009": [
        "4.evidence/FCRP/FCRP-001_全12ヶ月RAW_SAP_FB50_月次決算ジョブログ.csv",
    ],
    "ELC-012": [
        "0.profile/company_profile.md",
        "4.evidence/ITGC/EM_外部委託管理/ITGC-EM-001_RAW_SOC1_TypeII_SIerA_FY2024.pdf",
    ],
    "PLC-S-001": ["0.profile/規程_職務権限規程_R18.pdf"],
    "PLC-P-002": ["0.profile/規程_職務権限規程_R18.pdf"],
    "ITGC-AC-001": ["0.profile/規程_職務権限規程_R18.pdf"],
    "ITAC-004": [
        "4.evidence/PLC-P/PLC-P-002_SAPワークフロー承認履歴ログ_FY2025.csv",
        "4.evidence/PLC-P/PLC-P-002_25件対応_RAW_SAPワークフロー承認履歴.csv",
        "4.evidence/ITGC/CM_変更管理/ITGC-CM-001_25件対応_RAW_変更管理台帳.csv",
    ],
    "ITAC-005": [
        "4.evidence/FCRP/FCRP-002_全4四半期RAW_連結システムS05_パッケージ受信ログ.csv",
    ],
}


# ============================================================
# ファイル名→統制ID マッピング（直接エビデンス判定用）
# ============================================================
def parse_control_id_from_filename(filename):
    """ファイル名から統制IDを抽出"""
    for cid in sorted(CONTROL_TO_DIR.keys(), key=len, reverse=True):
        if filename.startswith(cid + "_") or filename.startswith(cid + " ") or filename.startswith(cid + "."):
            return cid
    return None


# ============================================================
# Step 1: ITGC サブフォルダをフラット化
# ============================================================
def flatten_itgc():
    itgc_root = EVIDENCE / "ITGC"
    subdirs = ["AC_アクセス管理", "CM_変更管理", "OM_運用管理", "EM_外部委託管理"]
    moved = 0
    for sub in subdirs:
        sub_path = itgc_root / sub
        if sub_path.exists():
            for f in sub_path.iterdir():
                if f.is_file():
                    dest = itgc_root / f.name
                    if dest.exists():
                        print(f"  WARNING: target exists: {dest.name}")
                    shutil.move(str(f), str(dest))
                    moved += 1
            sub_path.rmdir()
            print(f"  Removed dir: ITGC/{sub}")
    print(f"  Moved {moved} files to ITGC/")


# ============================================================
# Step 2: クロス参照ファイルを各区分フォルダにコピー
# ============================================================
def copy_cross_references():
    copied = 0
    failed = 0
    for cid, refs in CROSS_REF.items():
        rcm = CONTROL_TO_DIR.get(cid)
        if not rcm:
            print(f"  WARNING: unknown control: {cid}")
            continue
        target_dir = EVIDENCE / rcm
        for ref in refs:
            src = BASE / ref
            # フラット化後のパスに自動補正
            if "ITGC/AC_" in ref or "ITGC/CM_" in ref or "ITGC/OM_" in ref or "ITGC/EM_" in ref:
                src = BASE / ref.split("/", 3)[0] / ref.split("/", 3)[1] / ref.split("/", 3)[3]
            if not src.exists():
                print(f"  ERROR: not found: {src}")
                failed += 1
                continue
            dst = target_dir / src.name
            if not dst.exists():
                shutil.copy2(src, dst)
                print(f"  Copied: {src.name} -> {rcm}/")
                copied += 1
    print(f"  Copied {copied} files, {failed} failed")


# ============================================================
# Step 3: CSV 生成 (統制ID, Filename のみ)
# ============================================================
def build_mapping_csv():
    mappings = set()  # (control_id, filename)

    # 直接エビデンス: 各区分フォルダのファイルをスキャン
    for rcm_dir in ["ELC", "PLC-S", "PLC-P", "PLC-I", "ITGC", "ITAC", "FCRP"]:
        folder = EVIDENCE / rcm_dir
        for f in folder.iterdir():
            if f.is_file():
                cid = parse_control_id_from_filename(f.name)
                if cid:
                    mappings.add((cid, f.name))

    # クロス参照: コピー済のファイル名を追加
    for cid, refs in CROSS_REF.items():
        rcm = CONTROL_TO_DIR.get(cid)
        if not rcm:
            continue
        for ref in refs:
            filename = Path(ref).name
            if (EVIDENCE / rcm / filename).exists():
                mappings.add((cid, filename))

    # ソート
    def sort_key(row):
        cid, fname = row
        # Extract control type order
        if cid.startswith("ELC"):
            ord_ = (1, cid)
        elif cid.startswith("PLC-S"):
            ord_ = (2, cid)
        elif cid.startswith("PLC-P"):
            ord_ = (3, cid)
        elif cid.startswith("PLC-I"):
            ord_ = (4, cid)
        elif cid.startswith("ITGC"):
            ord_ = (5, cid)
        elif cid.startswith("ITAC"):
            ord_ = (6, cid)
        elif cid.startswith("FCRP"):
            ord_ = (7, cid)
        else:
            ord_ = (99, cid)
        return (ord_, fname)

    mappings_list = sorted(mappings, key=sort_key)

    # CSV 出力 (ロック中の場合はtempに書いてから置換)
    import time
    temp_csv = OUTPUT_CSV.parent / (OUTPUT_CSV.stem + "_new.csv")
    with open(temp_csv, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f, quoting=csv.QUOTE_ALL)
        writer.writerow(["統制ID", "Filename"])
        for row in mappings_list:
            writer.writerow(row)
    # リネーム試行
    try:
        if OUTPUT_CSV.exists():
            OUTPUT_CSV.unlink()
        temp_csv.rename(OUTPUT_CSV)
    except PermissionError:
        print(f"  WARNING: couldn't overwrite, saved to {temp_csv.name}")

    # 検証: 各ファイルが実在するか
    missing = []
    for cid, fname in mappings_list:
        rcm = CONTROL_TO_DIR.get(cid)
        if not (EVIDENCE / rcm / fname).exists():
            missing.append((cid, fname, rcm))

    print(f"\n  Total mappings: {len(mappings_list)}")
    print(f"  Unique filenames: {len(set(m[1] for m in mappings_list))}")
    print(f"  Covered controls: {len(set(m[0] for m in mappings_list))}/{len(CONTROL_TO_DIR)}")
    if missing:
        print(f"  WARNING: {len(missing)} missing files!")
        for m in missing[:5]:
            print(f"    - {m[0]} / {m[1]} (expected in {m[2]}/)")

    # カバレッジチェック：評価対象になっていない統制を特定
    covered = set(m[0] for m in mappings_list)
    uncovered = [c for c in CONTROL_TO_DIR.keys() if c not in covered]
    if uncovered:
        print(f"  Uncovered controls: {uncovered}")


if __name__ == "__main__":
    print("=== Step 1: Flatten ITGC subdirectories ===")
    flatten_itgc()
    print("\n=== Step 2: Copy cross-referenced files ===")
    copy_cross_references()
    print("\n=== Step 3: Generate mapping CSV ===")
    build_mapping_csv()
    print("\nReorganization completed.")
