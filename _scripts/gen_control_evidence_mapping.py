"""
RCM統制 × エビデンス マッピングCSV生成
4.evidence/ 配下の全ファイルを統制ID別に対応付け
"""
import csv
from pathlib import Path

BASE = Path(r"C:\Users\nyham\work\demo_data")
EVIDENCE_DIR = BASE / "4.evidence"
OUTPUT = BASE / "2.RCM" / "統制_エビデンスマッピング.csv"


# 統制マスタ（53統制すべて）
CONTROLS = [
    # (統制ID, 統制名, RCM区分, キー統制, 評価結果サマリ)
    ("ELC-001", "取締役会の機能", "ELC", "Y", "有効"),
    ("ELC-002", "倫理綱領の浸透", "ELC", "N", "有効"),
    ("ELC-003", "職務権限と組織体制", "ELC", "Y", "有効"),
    ("ELC-004", "全社リスク評価の実施", "ELC", "Y", "有効"),
    ("ELC-005", "不正リスク評価", "ELC", "Y", "有効"),
    ("ELC-006", "規程・マニュアル整備", "ELC", "N", "有効"),
    ("ELC-007", "職務の分離", "ELC", "Y", "軽微な不備(PLC-P-002関連)"),
    ("ELC-008", "内部通報制度", "ELC", "Y", "有効"),
    ("ELC-009", "決算情報の伝達", "ELC", "N", "有効"),
    ("ELC-010", "内部監査の実施", "ELC", "Y", "有効"),
    ("ELC-011", "監査等委員会のモニタリング", "ELC", "Y", "有効"),
    ("ELC-012", "IT戦略と情報セキュリティ方針", "ELC", "Y", "有効"),
    ("PLC-S-001", "受注・与信承認", "PLC-S", "Y", "有効(軽微例外1件)"),
    ("PLC-S-002", "出荷-売上マッチング", "PLC-S", "Y", "有効"),
    ("PLC-S-003", "請求書発行", "PLC-S", "Y", "有効"),
    ("PLC-S-004", "入金消込", "PLC-S", "Y", "有効"),
    ("PLC-S-005", "売掛金年齢分析", "PLC-S", "N", "判断保留(HOLD-2026-001)"),
    ("PLC-S-006", "期末カットオフ", "PLC-S", "Y", "有効"),
    ("PLC-S-007", "価格マスタ承認", "PLC-S", "N", "有効"),
    ("PLC-P-001", "購買依頼承認", "PLC-P", "N", "有効"),
    ("PLC-P-002", "発注承認(金額別)", "PLC-P", "Y", "不備(DEF-2026-002 重要な不備の可能性)"),
    ("PLC-P-003", "検収", "PLC-P", "Y", "有効"),
    ("PLC-P-004", "3-wayマッチング", "PLC-P", "Y", "有効"),
    ("PLC-P-005", "仕入先マスタ管理", "PLC-P", "N", "有効"),
    ("PLC-P-006", "支払承認", "PLC-P", "Y", "有効"),
    ("PLC-P-007", "期末未払計上", "PLC-P", "Y", "有効"),
    ("PLC-I-001", "実地棚卸", "PLC-I", "Y", "有効"),
    ("PLC-I-002", "棚卸差異調整", "PLC-I", "Y", "不備(DEF-2026-003 軽微な不備)"),
    ("PLC-I-003", "標準原価更新承認", "PLC-I", "Y", "有効"),
    ("PLC-I-004", "原価差異分析", "PLC-I", "N", "有効"),
    ("PLC-I-005", "滞留在庫評価損", "PLC-I", "Y", "有効"),
    ("PLC-I-006", "WMS-ERP在庫一致", "PLC-I", "N", "有効"),
    ("PLC-I-007", "原価計算月次締め", "PLC-I", "Y", "有効"),
    ("ITGC-AC-001", "新規ユーザ登録承認", "ITGC", "Y", "有効"),
    ("ITGC-AC-002", "アクセス権定期棚卸", "ITGC", "Y", "判断保留(HOLD-2026-002)"),
    ("ITGC-AC-003", "退職者アカウント停止", "ITGC", "Y", "不備(DEF-2026-001 重要な不備)"),
    ("ITGC-AC-004", "特権ID管理", "ITGC", "Y", "有効"),
    ("ITGC-CM-001", "変更申請・承認", "ITGC", "Y", "有効"),
    ("ITGC-CM-002", "テスト実施", "ITGC", "Y", "有効(軽微例外許容)"),
    ("ITGC-CM-003", "本番移送", "ITGC", "Y", "有効"),
    ("ITGC-OM-001", "バックアップ", "ITGC", "Y", "有効"),
    ("ITGC-OM-002", "障害管理", "ITGC", "N", "有効"),
    ("ITGC-EM-001", "委託先管理", "ITGC", "Y", "有効"),
    ("ITAC-001", "与信限度自動チェック", "ITAC", "Y", "有効"),
    ("ITAC-002", "3-way自動マッチング", "ITAC", "Y", "有効"),
    ("ITAC-003", "減価償却自動計算", "ITAC", "Y", "有効"),
    ("ITAC-004", "承認ルーティング判定", "ITAC", "Y", "有効"),
    ("ITAC-005", "連結パッケージ検証", "ITAC", "Y", "有効"),
    ("FCRP-001", "月次決算チェックリスト", "FCRP", "Y", "有効"),
    ("FCRP-002", "連結パッケージ検証", "FCRP", "Y", "有効"),
    ("FCRP-003", "会計上の見積レビュー", "FCRP", "Y", "判断保留(HOLD-2026-003)"),
    ("FCRP-004", "連結仕訳承認", "FCRP", "Y", "有効"),
    ("FCRP-005", "開示書類レビュー", "FCRP", "Y", "有効"),
]

CONTROL_DICT = {c[0]: c for c in CONTROLS}


# エビデンスファイルの用途定義（主要ファイル）
PURPOSE = {
    # ELC
    "ELC-001_取締役会議事録_第245回_2025年9月.pdf": "取締役会の開催状況・議事内容",
    "ELC-002_倫理綱領受領確認書提出状況_2025年度.xlsx": "倫理綱領受領確認書の提出状況（全部門）",
    "ELC-004_全社リスクアセスメント結果_2025年度.xlsx": "年次リスクアセスメントの結果",
    "ELC-008_内部通報受付台帳_FY2025.xlsx": "内部通報受付・調査完了記録",
    "ELC-010_2025年度内部監査計画書.pdf": "年次内部監査計画の策定記録",
    # PLC-S-001
    "PLC-S-001_SAP_VA05_受注伝票一覧_FY2025.xlsx": "受注承認統制の母集団（FY2025全受注）",
    "PLC-S-001_与信限度マスタ_SAP_FD32スナップショット.xlsx": "与信限度マスタ（評価基準日時点）",
    "PLC-S-001_販売関連承認権限一覧_職務権限規程R18抜粋.pdf": "承認権限の規程根拠",
    "PLC-S-001_SAP与信チェックログ_202511.csv": "与信チェックの自動実行ログ",
    "PLC-S-001_注文書_ORD-2025-1420_サンプル顧客B社.pdf": "個別取引の受注原本（通常ケース）",
    "PLC-S-001_注文書_ORD-2025-0412_サンプル顧客H社.pdf": "個別取引の受注原本（複数品目）",
    "PLC-S-001_注文書_ORD-2025-1876_サンプル顧客L社.pdf": "個別取引の受注原本（承認遅延例外ケース）",
    "PLC-S-001_SAP受注登録画面_ORD-2025-1420.png": "SAP VA03 受注登録画面（通常ケース）",
    "PLC-S-001_SAP与信超過アラート画面_C-10007.png": "SAP与信超過時のアラート画面",
    "PLC-S-001_ワークフロー承認_与信超過サンプル.png": "与信超過時のワークフロー承認画面",
    # PLC-S-002
    "PLC-S-002_WMS出荷実績エクスポート_202511.csv": "WMSからの出荷実績データ（生データ）",
    "PLC-S-002_SAP売上計上明細_202511.csv": "SAPからの売上計上仕訳明細（生データ）",
    "PLC-S-002_出荷売上マッチング照合レポート_202511.xlsx": "経理部が作成した突合照合記録",
    "PLC-S-002_SAP未マッチ明細リスト_202511.csv": "未マッチ明細の検出ログ",
    # PLC-S-003
    "PLC-S-003_SAP請求書バッチ実行ログ_202511.txt": "月次請求書自動発行バッチのログ",
    "PLC-S-003_月次請求書発行一覧_202511.xlsx": "経理部による突合チェック記録",
    "PLC-S-003_請求書_INV-202511-0234.pdf": "個別請求書原本",
    # PLC-S-004
    "PLC-S-004_FB入金データ_202511.csv": "銀行FBデータ（全銀形式）",
    "PLC-S-004_入金消込リスト_202511.xlsx": "経理部作成の入金消込記録",
    "PLC-S-004_SAP入金消込画面.png": "SAP F-28 入金消込画面",
    # PLC-S-005【判断保留】
    "PLC-S-005_売掛金年齢表_202511.xlsx": "経理部作成の売掛金年齢分析",
    "PLC-S-005_売掛金年齢表_経理部長承認PDF_低解像度.pdf": "【判断保留】承認印が低解像度で判読不能",
    # PLC-S-006
    "PLC-S-006_期末カットオフテスト.xlsx": "期末前後5営業日の出荷全数検証",
    # PLC-S-007
    "PLC-S-007_価格マスタ変更稟議_W-2025-1876.pdf": "価格変更稟議書（承認印付）",
    "PLC-S-007_価格変更履歴レポート_Q3.xlsx": "経理部による四半期レビュー記録",
    # PLC-S共通
    "PLC-S_月次売上会議_議事録_202511.md": "営業本部の月次会議議事録（PLC-S全般のモニタリング）",
    # PLC-P-001
    "PLC-P-001_SAP購買依頼一覧_202511.xlsx": "購買依頼の月次実績（承認状況含む）",
    # PLC-P-002【真の不備】
    "PLC-P-002_SAP_ME2N_発注伝票一覧_FY2025.xlsx": "発注承認統制の母集団（不備3件含む）",
    "PLC-P-002_購買関連承認権限一覧_職務権限規程R18抜粋.pdf": "購買承認権限の規程根拠",
    "PLC-P-002_発注書_PO-2025-2560_通常承認.pdf": "個別発注書原本（通常・課長承認）",
    "PLC-P-002_発注書_PO-2025-3072_通常承認.pdf": "個別発注書原本（通常・部長承認）",
    "PLC-P-002_発注書_PO-2025-0234_不備ケース1_権限外承認.pdf": "【不備】PUR003が権限なしで承認",
    "PLC-P-002_発注書_PO-2025-0789_不備ケース2_担当者承認.pdf": "【不備】PUR004が上限50万円超を承認",
    "PLC-P-002_発注書_PO-2025-1456_不備ケース3_課長上限超過.pdf": "【不備】課長上限500万円超の¥7.85M承認",
    "PLC-P-002_SAPワークフロー承認履歴ログ_FY2025抜粋.csv": "承認ワークフロー履歴（不備ケース識別済）",
    "PLC-P-002_SAP発注登録画面_PO-2025-2560.png": "SAP ME23N 発注登録画面（通常）",
    "PLC-P-002_SAP発注画面_不備ケースPO-2025-1456.png": "SAP発注画面（不備ケースの権限超過警告）",
    "PLC-P-002_ワークフロー承認画面_通常案件.png": "ワークフロー承認画面（通常案件）",
    # PLC-P-003
    "PLC-P-003_検収報告書_REC-2025-5678.pdf": "個別検収報告書原本",
    "PLC-P-003_検収差異報告書_DIF-2025-0019.pdf": "差異発生時の報告書（例外処理）",
    # PLC-P-004
    "PLC-P-004_3wayマッチング結果_202511.xlsx": "3-wayマッチング月次結果記録",
    # PLC-P-005
    "PLC-P-005_仕入先マスタ登録申請書_V-20029.pdf": "新規仕入先登録申請（反社チェック含む）",
    # PLC-P-006
    "PLC-P-006_支払予定一覧_202511.xlsx": "月次支払予定一覧（経理部長承認）",
    # PLC-P-007
    "PLC-P-007_期末未払計上リスト.xlsx": "期末未払計上の全数検証記録",
    # PLC-I-001
    "PLC-I-001_実地棚卸計画書_2025下期.pdf": "半期棚卸の計画書",
    "PLC-I-001_実地棚卸報告書_2025年9月.xlsx": "棚卸報告書（倉庫別サマリ・差異明細）",
    "PLC-I-001_棚卸写真_本社倉庫A_区画A-3.jpg": "棚卸実施時の現場写真（通常区画）",
    "PLC-I-001_棚卸写真_本社倉庫B_区画B-3_差異発生区画.jpg": "棚卸写真（差異発生区画）",
    "PLC-I-001_棚卸写真_東北工場倉庫_区画T-1.jpg": "棚卸写真（東北工場）",
    "PLC-I-001_棚卸写真_本社倉庫A_区画A-7_立会.jpg": "棚卸写真（経理部立会）",
    "PLC-I-001_SAP在庫数量一覧_MB52.png": "SAP MB52 在庫数量照会画面",
    # PLC-I-002【真の不備】
    "PLC-I-002_棚卸差異分析書_INV-DIFF-2025-09-012.pdf": "差異分析書（実施済の正常ケース）。なお¥850,000差異は分析書未作成（不備）",
    # PLC-I-003
    "PLC-I-003_標準原価更新稟議_W-2025-0089.pdf": "期首標準原価更新の承認稟議",
    # PLC-I-004
    "PLC-I-004_原価差異分析表_202511.xlsx": "月次原価差異分析",
    # PLC-I-005
    "PLC-I-005_滞留在庫評価損計算_2025年12月末.xlsx": "四半期末の滞留在庫評価損計算",
    # PLC-I-006
    "PLC-I-006_WMS-ERP在庫照合レポート_202511月次サンプル.csv": "WMS-ERP日次照合レポート",
    # PLC-I-007
    "PLC-I-007_月次原価計算締めチェックリスト_202511.xlsx": "月次締め45項目のチェックリスト",
    # ITGC AC
    "ITGC-AC-001_ユーザ登録申請書_USER-REG-2025-0087.pdf": "新規ユーザ登録申請（承認付）",
    "ITGC-AC-001_SAP_SU01_ユーザ作成画面.png": "SAP SU01 ユーザ作成画面",
    "ITGC-AC-001_SAPアクセス権マトリクス.png": "アクセス権マトリクス（SoD違反1件を視覚化）",
    "ITGC-AC-002_SAP_SUIM_有効ユーザ一覧_Q3棚卸用.xlsx": "SUIMユーザ一覧＋棚卸結果（抽出条件不明→判断保留）",
    "ITGC-AC-003_退職者アカウント停止記録_FY2025.xlsx": "退職者停止記録（不備2件を記録）",
    "ITGC-AC-003_SAP_SM19_SM20_退職者ログインログ抽出.csv": "退職者の停止遅延期間中のログイン履歴（ゼロ確認）",
    "ITGC-AC-004_特権ID操作ログ_202511.csv": "月次特権ID操作ログレビュー",
    # ITGC CM
    "ITGC-CM-001_変更管理一覧_FY2025.xlsx": "FY2025変更申請全42件の一覧",
    "ITGC-CM-001_変更申請書_REL-2025-023.pdf": "個別変更申請書（承認付）",
    "ITGC-CM-002_UATテスト結果_REL-2025-023.xlsx": "UATテストケース実施結果",
    "ITGC-CM-003_SAP_STMS_本番移送記録_FY2025Q2-Q3.csv": "SAP STMS本番移送履歴",
    # ITGC OM
    "ITGC-OM-001_バックアップ実施記録_202511.xlsx": "日次バックアップ実施記録",
    "ITGC-OM-001_DRリストアテスト報告書_2025Q3.pdf": "四半期DRテスト結果",
    "ITGC-OM-002_障害管理台帳_FY2025.xlsx": "障害発生記録と対応",
    # ITGC EM
    "ITGC-EM-001_SOC1レポート評価レビューシート_2025.pdf": "外部委託先のSOC1レポート評価",
    "ITGC-EM-001_IT外部委託先一覧_FY2025.xlsx": "IT外部委託先の管理台帳",
    # ITAC
    "ITAC-001_SAP与信限度自動チェック設定画面_OVAK.png": "SAP OVAK 与信自動チェック設定",
    "ITAC-001_与信限度自動チェック_動作検証.xlsx": "ITAC-001動作検証テスト結果",
    "ITAC-002_SAP3wayマッチング設定画面_OMRK.png": "SAP OMRK 3-wayマッチング設定",
    "ITAC-002_3wayマッチング結果ログ_202511.csv": "月次3-wayマッチング実行ログ",
    "ITAC-003_SAP減価償却実行画面_AFAB.png": "SAP AFAB 月次減価償却実行画面",
    "ITAC-003_減価償却手計算検証.xlsx": "減価償却の再計算検証",
    # FCRP
    "FCRP-001_月次決算チェックリスト_202511.xlsx": "月次決算45項目のチェックリスト",
    "FCRP-002_連結パッケージ受領管理_2025Q3.xlsx": "四半期連結パッケージ受領・検証管理",
    "FCRP-003_貸倒引当金計算シート_2025年12月末.xlsx": "貸倒引当金計算（根拠資料不足→判断保留）",
    "FCRP-004_連結仕訳一覧_2025Q3.xlsx": "連結仕訳の一覧と承認",
    "FCRP-005_開示書類レビューシート_2026年3月期Q3.pdf": "四半期開示書類の3段階レビュー",
}


# 直接エビデンスがない統制の関連参照（相互参照）
CROSS_REF = {
    "ELC-003": ["1.master_data/employees.xlsx (承認権限金額列)",
                "1.master_data/user_roles_matrix.xlsx (SAPロール×ユーザ)",
                "0.profile/company_profile.md (組織図・職務権限規程R18)"],
    "ELC-005": ["4.evidence/ELC/ELC-004_全社リスクアセスメント結果_2025年度.xlsx (不正リスクファクター含む)"],
    "ELC-006": ["0.profile/company_profile.md (第6章 主要規程体系 R01-R27)"],
    "ELC-007": ["1.master_data/user_roles_matrix.xlsx (SoD違反PUR004を含むロールマトリクス)",
                "4.evidence/ITGC/AC_アクセス管理/ITGC-AC-001_SAPアクセス権マトリクス.png",
                "4.evidence/PLC-P/PLC-P-002_* (発注承認不備ケース)"],
    "ELC-009": ["4.evidence/FCRP/FCRP-001_月次決算チェックリスト_202511.xlsx",
                "4.evidence/PLC-S/PLC-S_月次売上会議_議事録_202511.md"],
    "ELC-011": ["4.evidence/ELC/ELC-001_取締役会議事録_第245回_2025年9月.pdf (監査等委員会陪席含む)",
                "4.evidence/ELC/ELC-008_内部通報受付台帳_FY2025.xlsx (監査等委員会報告)"],
    "ELC-012": ["0.profile/company_profile.md (第5章 ITシステム構成、R21 情報セキュリティ基本方針)",
                "4.evidence/ITGC/EM_外部委託管理/ITGC-EM-001_SOC1レポート評価レビューシート_2025.pdf"],
    "ITAC-004": ["4.evidence/PLC-P/PLC-P-002_SAPワークフロー承認履歴ログ_FY2025抜粋.csv (ルーティング実行履歴)",
                 "4.evidence/PLC-S/PLC-S-007_価格マスタ変更稟議_W-2025-1876.pdf (ワークフロー承認実例)",
                 "4.evidence/ITGC/CM_変更管理/ITGC-CM-001_変更申請書_REL-2025-023.pdf (変更承認ルート)"],
    "ITAC-005": ["4.evidence/FCRP/FCRP-002_連結パッケージ受領管理_2025Q3.xlsx (バリデーションエラー検出実績)"],
}


def parse_control_id(filename):
    """ファイル名から統制IDを抽出"""
    # ELC-001, PLC-S-001, PLC-P-001, PLC-I-001, ITGC-AC-001, ITAC-001, FCRP-001 など
    # PLC-S 共通ファイル（"PLC-S_" プレフィックス）は PLC-S 全体に関連
    if filename.startswith("PLC-S_月次売上会議"):
        return "PLC-S-004"  # 最も議事録と関連の深い統制として入金消込にマッピング（正確には全体だが）

    # 通常の統制IDパターン
    parts = filename.split("_", 1)[0]  # "ELC-001" など
    if parts in CONTROL_DICT:
        return parts
    return None


def get_file_format(filename):
    ext = Path(filename).suffix.upper().lstrip(".")
    return ext if ext else "UNKNOWN"


def main():
    # 4.evidence/ 配下のすべてのファイルをスキャン
    all_files = []
    for f in EVIDENCE_DIR.rglob("*"):
        if f.is_file():
            rel_path = f.relative_to(BASE).as_posix()
            all_files.append((f.name, rel_path))
    all_files.sort()

    # 統制IDごとにマッピング
    mapping_rows = []
    orphan_files = []

    for fname, rel_path in all_files:
        cid = parse_control_id(fname)
        if cid and cid in CONTROL_DICT:
            ctrl = CONTROL_DICT[cid]
            purpose = PURPOSE.get(fname, "(用途記述なし)")
            mapping_rows.append({
                "統制ID": cid,
                "統制名": ctrl[1],
                "RCM区分": ctrl[2],
                "キー統制": ctrl[3],
                "評価結果": ctrl[4],
                "エビデンス種別": "直接エビデンス",
                "ファイル名": fname,
                "ファイルパス": rel_path,
                "形式": get_file_format(fname),
                "用途・内容": purpose,
            })
        else:
            orphan_files.append((fname, rel_path))

    # PLC-S 共通の議事録を PLC-S-001 と 005 にも複製追加（明示的な相互参照）
    shared_ms = "PLC-S_月次売上会議_議事録_202511.md"
    for row in all_files:
        if row[0] == shared_ms:
            for cid_extra in ["PLC-S-005"]:
                ctrl = CONTROL_DICT[cid_extra]
                mapping_rows.append({
                    "統制ID": cid_extra,
                    "統制名": ctrl[1],
                    "RCM区分": ctrl[2],
                    "キー統制": ctrl[3],
                    "評価結果": ctrl[4],
                    "エビデンス種別": "関連エビデンス",
                    "ファイル名": row[0],
                    "ファイルパス": row[1],
                    "形式": get_file_format(row[0]),
                    "用途・内容": "営業本部月次会議での売掛金年齢分析レビュー記録",
                })

    # 直接エビデンスがない統制について、相互参照を追加
    covered_ctrls = set(r["統制ID"] for r in mapping_rows)
    for cid, refs in CROSS_REF.items():
        if cid not in covered_ctrls or cid in ("ELC-007", "ELC-009", "ELC-011", "ELC-012", "ITAC-004", "ITAC-005"):
            # 直接エビデンスがない、または補足参照が必要な統制
            ctrl = CONTROL_DICT[cid]
            for ref in refs:
                mapping_rows.append({
                    "統制ID": cid,
                    "統制名": ctrl[1],
                    "RCM区分": ctrl[2],
                    "キー統制": ctrl[3],
                    "評価結果": ctrl[4],
                    "エビデンス種別": "相互参照",
                    "ファイル名": ref.split("/")[-1].split(" ")[0],
                    "ファイルパス": ref,
                    "形式": get_file_format(ref.split(" ")[0]) or "参照",
                    "用途・内容": "直接エビデンスがないため、他統制または会社プロファイルから相互参照",
                })

    # 統制ID順にソート
    mapping_rows.sort(key=lambda r: (r["RCM区分"], r["統制ID"], r["エビデンス種別"], r["ファイル名"]))

    # CSV出力
    fieldnames = ["統制ID", "統制名", "RCM区分", "キー統制", "評価結果",
                  "エビデンス種別", "ファイル名", "ファイルパス", "形式", "用途・内容"]

    with open(OUTPUT, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, quoting=csv.QUOTE_ALL)
        writer.writeheader()
        for row in mapping_rows:
            writer.writerow(row)

    print(f"Created: {OUTPUT.relative_to(BASE)}")
    print(f"  Total rows: {len(mapping_rows)}")
    print(f"  Direct evidence: {sum(1 for r in mapping_rows if r['エビデンス種別'] == '直接エビデンス')}")
    print(f"  Cross-reference: {sum(1 for r in mapping_rows if r['エビデンス種別'] == '相互参照')}")
    print(f"  Related: {sum(1 for r in mapping_rows if r['エビデンス種別'] == '関連エビデンス')}")

    # 統制別のカバレッジ確認
    covered = set(r["統制ID"] for r in mapping_rows)
    uncovered = [c[0] for c in CONTROLS if c[0] not in covered]
    print(f"\n  Covered controls: {len(covered)}/53")
    if uncovered:
        print(f"  Uncovered: {uncovered}")

    if orphan_files:
        print(f"\n  Orphan files (not mapped): {len(orphan_files)}")
        for fn, rp in orphan_files:
            print(f"    - {rp}")


if __name__ == "__main__":
    main()
