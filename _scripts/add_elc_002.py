"""ELC-002 倫理綱領確認 のRAW代替（HRシステムログ）を追加"""
import random
from pathlib import Path
from datetime import datetime
import sys
sys.path.insert(0, str(Path(__file__).parent))
from sample_gen_util import write_raw_csv

BASE = Path(r"C:\Users\nyham\work\demo_data\4.evidence\ELC")

random.seed(1002)
depts = [("経営陣", 7), ("経理部", 20), ("営業本部", 45), ("製造本部", 280),
         ("技術本部", 60), ("管理本部(除く経理)", 60), ("品質保証部", 25),
         ("情シス部", 15), ("経営企画部", 10)]

rows = []
for dept, total in depts:
    for i in range(min(total, 3)):  # 部門ごとに代表3件
        ack_ts = datetime(2025, random.randint(5, 6), random.randint(1, 28),
                          random.randint(9, 17), random.randint(0, 59))
        rows.append([ack_ts.strftime("%Y-%m-%d %H:%M:%S"),
                     f"E{random.randint(1, 999):04d}",
                     dept, "倫理綱領2025年度版", "受領確認済"])

write_raw_csv(
    BASE / "ELC-002_RAW_HRシステム_倫理綱領受領確認ログ.csv",
    ["# HRIS (Human Resource Information System) - E-Learning Module",
     "# Report:   Code of Ethics Acknowledgment Log (FY2025)",
     "# Total distribution: 522 employees / Acknowledged: 522 (100%)",
     "# Export:   2025-06-30 18:00:00 JST"],
    "確認タイムスタンプ,社員番号,所属部門,対象文書,確認ステータス",
    rows,
    footer_lines=["# Displayed: representative 27 out of 522 acknowledgments"]
)
print("Created: ELC-002_RAW_HRシステム_倫理綱領受領確認ログ.csv")
