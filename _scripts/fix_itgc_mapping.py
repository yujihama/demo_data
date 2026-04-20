"""Evidence_Mapping_ITGC.csv を整列・欠落補完・削除済み除去する"""
import csv
import os
from pathlib import Path

ROOT = Path(r"C:\Users\nyham\work\demo_data")
MAPPING = ROOT / "2.RCM" / "Evidence_Mapping_ITGC.csv"
EVID_DIR = ROOT / "4.evidence" / "ITGC"

# Read current mapping (strip BOM if present)
current = []  # [(control_id, filename)]
with open(MAPPING, encoding='utf-8-sig') as f:
    reader = csv.reader(f)
    header = next(reader)
    for row in reader:
        if len(row) >= 2 and row[0]:
            current.append((row[0], row[1]))

# Actual files
actual_files = set(os.listdir(EVID_DIR))

# Remove stale entries (file not present)
current = [(c, f) for c, f in current if f in actual_files]

# Determine control ID for new files
def infer_control(filename):
    if filename.startswith('変更申請書_REL'):
        return 'ITGC-CM-001'
    return None

# Add new files
mapped_files = {f for _, f in current}
for f in sorted(actual_files):
    if f in mapped_files:
        continue
    cid = infer_control(f)
    if cid:
        current.append((cid, f))

# Sort: group by control ID (in canonical order), then by filename
control_order = [
    'ITGC-AC-001', 'ITGC-AC-002', 'ITGC-AC-003', 'ITGC-AC-004',
    'ITGC-CM-001', 'ITGC-CM-002', 'ITGC-CM-003',
    'ITGC-EM-001', 'ITGC-OM-001', 'ITGC-OM-002',
]

def sort_key(item):
    cid, fn = item
    return (control_order.index(cid) if cid in control_order else 99, fn)

current.sort(key=sort_key)

# Write with UTF-8 BOM (for Excel Japanese compatibility)
with open(MAPPING, 'w', encoding='utf-8-sig', newline='') as f:
    writer = csv.writer(f, quoting=csv.QUOTE_ALL)
    writer.writerow(['統制ID', 'Filename'])
    for cid, fn in current:
        writer.writerow([cid, fn])

print(f"Total mapping entries: {len(current)}")

# Summary per control
from collections import Counter
counts = Counter(c for c, _ in current)
for cid in control_order:
    print(f"  {cid}: {counts.get(cid, 0)}")
