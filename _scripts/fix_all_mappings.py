"""全統制カテゴリのEvidence_Mapping CSVを整合化＆スキーマ変更

Step1: PLC-S, FCRP の未マップファイルを追加
Step2: 全CSVを新スキーマに変換
  - col1: key (旧: 統制ID)
  - col2: sample_no (新規, 全行=1)
  - col3: filename (旧: Filename)
"""
import csv
import os
import re
import sys
import io
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

ROOT = Path(r"C:\Users\nyham\work\demo_data")

CATEGORIES = [
    ('PLC-S', 'PLC-S'),
    ('PLC-P', 'PLC-P'),
    ('PLC-I', 'PLC-I'),
    ('ITGC', 'ITGC'),
    ('ITAC', 'ITAC'),
    ('ELC', 'ELC'),
    ('FCRP', 'FCRP'),
]

# Control order for sorting (canonical)
CONTROL_ORDERS = {
    'PLC-S': ['PLC-S-001', 'PLC-S-002', 'PLC-S-003', 'PLC-S-004', 'PLC-S-005', 'PLC-S-006', 'PLC-S-007'],
    'PLC-P': ['PLC-P-001', 'PLC-P-002', 'PLC-P-003', 'PLC-P-004', 'PLC-P-005', 'PLC-P-006', 'PLC-P-007'],
    'PLC-I': ['PLC-I-001', 'PLC-I-002', 'PLC-I-003', 'PLC-I-004', 'PLC-I-005', 'PLC-I-006', 'PLC-I-007'],
    'ITGC': ['ITGC-AC-001', 'ITGC-AC-002', 'ITGC-AC-003', 'ITGC-AC-004',
             'ITGC-CM-001', 'ITGC-CM-002', 'ITGC-CM-003',
             'ITGC-EM-001', 'ITGC-OM-001', 'ITGC-OM-002'],
    'ITAC': ['ITAC-001', 'ITAC-002', 'ITAC-003', 'ITAC-004'],
    'ELC': ['ELC-001', 'ELC-002', 'ELC-003', 'ELC-004', 'ELC-005', 'ELC-006', 'ELC-007', 'ELC-008'],
    'FCRP': ['FCRP-001', 'FCRP-002', 'FCRP-003', 'FCRP-004', 'FCRP-005', 'FCRP-006', 'FCRP-007'],
}


def infer_control_id_from_filename(filename, category):
    """PLC-S-001_xxx.csv のようなファイル名から統制IDを抽出"""
    # Pattern: PLC-S-NNN_ or ITGC-XX-NNN_ or FCRP-NNN_ etc.
    m = re.match(r'^(PLC-[SPI]-\d{3}|ITGC-(?:AC|CM|OM|EM)-\d{3}|ITAC-\d{3}|ELC-\d{3}|FCRP-\d{3})_', filename)
    if m:
        return m.group(1)
    return None


def step1_add_missing_plcs_fcrp():
    """PLC-S/FCRP の未マップファイルを追加"""
    for cat_name, dir_name in [('PLC-S', 'PLC-S'), ('FCRP', 'FCRP')]:
        csv_path = ROOT / "2.RCM" / f"Evidence_Mapping_{cat_name}.csv"
        dir_path = ROOT / "4.evidence" / dir_name

        # Read current mapping
        current = []
        with open(csv_path, encoding='utf-8-sig') as f:
            reader = csv.reader(f)
            header = next(reader)
            for row in reader:
                if len(row) >= 2 and row[0]:
                    current.append((row[0], row[1]))

        mapped_files = {f for _, f in current}
        actual = set(os.listdir(dir_path))

        added = 0
        for f in sorted(actual - mapped_files):
            cid = infer_control_id_from_filename(f, cat_name)
            if cid:
                current.append((cid, f))
                added += 1
                print(f'  [+] {cat_name}: added ({cid}, {f})')

        # Sort
        order = CONTROL_ORDERS[cat_name]
        def sk(item):
            cid, fn = item
            return (order.index(cid) if cid in order else 99, fn)
        current.sort(key=sk)

        # Write back in OLD schema (Step2 will convert)
        with open(csv_path, 'w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f, quoting=csv.QUOTE_ALL)
            writer.writerow(['統制ID', 'Filename'])
            for cid, fn in current:
                writer.writerow([cid, fn])

        print(f'[{cat_name}] Step1: added {added} files / total {len(current)}')


def step2_convert_schema():
    """全CSVを新スキーマに変換: key, sample_no, filename"""
    for cat_name, dir_name in CATEGORIES:
        csv_path = ROOT / "2.RCM" / f"Evidence_Mapping_{cat_name}.csv"

        rows = []
        with open(csv_path, encoding='utf-8-sig') as f:
            reader = csv.reader(f)
            next(reader)  # skip old header
            for row in reader:
                if len(row) >= 2 and row[0]:
                    rows.append((row[0], row[1]))

        # Sort by control order
        order = CONTROL_ORDERS[cat_name]
        def sk(item):
            cid, fn = item
            return (order.index(cid) if cid in order else 99, fn)
        rows.sort(key=sk)

        # Write new schema
        with open(csv_path, 'w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f, quoting=csv.QUOTE_ALL)
            writer.writerow(['key', 'sample_no', 'filename'])
            for cid, fn in rows:
                writer.writerow([cid, '1', fn])

        print(f'[{cat_name}] Step2: converted {len(rows)} rows to new schema (key, sample_no, filename)')


def verify():
    """Final verification"""
    print("\n=== FINAL VERIFICATION ===")
    for cat_name, dir_name in CATEGORIES:
        csv_path = ROOT / "2.RCM" / f"Evidence_Mapping_{cat_name}.csv"
        dir_path = ROOT / "4.evidence" / dir_name

        mapped = set()
        with open(csv_path, encoding='utf-8-sig') as f:
            reader = csv.reader(f)
            hdr = next(reader)
            for row in reader:
                if len(row) >= 3:
                    mapped.add(row[2])

        actual = set(os.listdir(dir_path))
        diff_a = actual - mapped
        diff_m = mapped - actual
        status = 'OK' if not diff_a and not diff_m else 'DRIFT'
        print(f'[{status}] {cat_name}: header={hdr}, mapped={len(mapped)}, actual={len(actual)}, unmapped={len(diff_a)}, stale={len(diff_m)}')
        if diff_a:
            for f in sorted(diff_a)[:5]: print(f'  + unmapped: {f}')
        if diff_m:
            for f in sorted(diff_m)[:5]: print(f'  - stale: {f}')


if __name__ == '__main__':
    print("=== Step 1: Add unmapped PLC-S/FCRP files ===")
    step1_add_missing_plcs_fcrp()
    print()
    print("=== Step 2: Convert schema to (key, sample_no, filename) ===")
    step2_convert_schema()
    verify()
