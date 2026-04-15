"""
PLC-S（販売プロセス）のPDF形式エビデンス生成
- 注文書 (顧客からの注文書を受領・スキャン風)
- 請求書 (自社発行)
- 価格マスタ変更稟議書
- 売掛金年齢表の「低解像度スキャンPDF」（判断保留ケース用）
"""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).parent))
from pdf_util import JPPDF
from PIL import Image, ImageDraw, ImageFilter, ImageFont
from datetime import date

BASE = Path(r"C:\Users\nyham\work\demo_data\4.evidence\PLC-S")


# ============================================================
# 1. 注文書PDF（顧客→自社の受領スキャン風）
# ============================================================
def gen_order_pdf(order_no, order_date, customer_code, customer_name,
                  customer_addr, items, delivery_date, rep_name,
                  output_name):
    pdf = JPPDF()
    pdf.add_page()

    # 文書タイトル
    pdf.set_font("YuGoth", "B", 20)
    pdf.cell(0, 12, "注 文 書", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(3)
    pdf.set_font("YuGoth", "", 10)
    pdf.cell(0, 5, f"注文番号: {order_no}", align="R", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 5, f"発行日: {order_date.strftime('%Y年%m月%d日')}",
             align="R", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(5)

    # 宛名 (自社)
    pdf.set_font("YuGoth", "B", 12)
    pdf.cell(0, 7, "株式会社テクノプレシジョン 御中", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("YuGoth", "", 10)
    pdf.cell(0, 5, "営業本部 担当者殿", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(3)

    # 発注元 (顧客情報 右寄せ)
    pdf.set_font("YuGoth", "B", 10)
    pdf.set_x(110)
    pdf.cell(90, 6, f"発注元: {customer_name}", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("YuGoth", "", 9)
    pdf.set_x(110)
    pdf.cell(90, 5, f"顧客コード: {customer_code}", new_x="LMARGIN", new_y="NEXT")
    pdf.set_x(110)
    pdf.cell(90, 5, customer_addr, new_x="LMARGIN", new_y="NEXT")
    pdf.ln(8)

    # 挨拶文
    pdf.set_font("YuGoth", "", 10)
    pdf.multi_cell(0, 5, "下記のとおり発注致します。ご確認のうえ、納期に間に合うよう手配をお願い致します。")
    pdf.ln(3)

    # 明細テーブル
    pdf.table_header(["品目コード", "品名", "数量", "単価(円)", "金額(円)"],
                     [30, 80, 20, 30, 30])
    subtotal = 0
    for code, name, qty, unit in items:
        amount = qty * unit
        subtotal += amount
        pdf.table_row([code, name, f"{qty:,}", f"{unit:,}", f"{amount:,}"],
                      [30, 80, 20, 30, 30], align="L")

    # 小計・消費税・合計
    pdf.set_font("YuGoth", "B", 10)
    pdf.cell(130, 7, "小計", border=1, align="R")
    pdf.cell(60, 7, f"¥ {subtotal:,}", border=1, align="R", new_x="LMARGIN", new_y="NEXT")
    tax = int(subtotal * 0.1)
    pdf.cell(130, 7, "消費税 (10%)", border=1, align="R")
    pdf.cell(60, 7, f"¥ {tax:,}", border=1, align="R", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(130, 8, "合計", border=1, align="R", fill=True)
    pdf.set_fill_color(255, 242, 204)
    total = subtotal + tax
    pdf.cell(60, 8, f"¥ {total:,}", border=1, align="R", fill=True,
             new_x="LMARGIN", new_y="NEXT")
    pdf.set_fill_color(255, 255, 255)
    pdf.ln(5)

    # 納入条件
    pdf.h3("■ 納入条件")
    pdf.set_font("YuGoth", "", 9)
    pdf.kv("納期", delivery_date.strftime("%Y年%m月%d日"))
    pdf.kv("納入場所", "貴社指定倉庫")
    pdf.kv("支払条件", "月末締 翌月末払")
    pdf.ln(8)

    # 受領スタンプと担当者
    y_stamp = pdf.get_y()
    pdf.set_xy(10, y_stamp)
    pdf.stamp("受領", x=30, y=y_stamp + 10)
    pdf.set_font("YuGoth", "", 9)
    pdf.set_xy(50, y_stamp + 5)
    pdf.cell(0, 5, f"受領日: {order_date.strftime('%Y/%m/%d')}",
             new_x="LMARGIN", new_y="NEXT")
    pdf.set_x(50)
    pdf.cell(0, 5, f"受領担当: {rep_name}", new_x="LMARGIN", new_y="NEXT")
    pdf.set_x(50)
    pdf.cell(0, 5, "SAP VA01に受注登録済", new_x="LMARGIN", new_y="NEXT")

    pdf.output(str(BASE / output_name))
    print(f"Created: {output_name}")


# ============================================================
# 2. 請求書PDF
# ============================================================
def gen_invoice_pdf(invoice_no, invoice_date, customer_code, customer_name,
                    items, due_date, output_name):
    pdf = JPPDF()
    pdf.add_page()

    pdf.set_font("YuGoth", "B", 22)
    pdf.cell(0, 14, "請 求 書", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("YuGoth", "", 10)
    pdf.cell(0, 5, f"請求書番号: {invoice_no}", align="R", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 5, f"請求日: {invoice_date.strftime('%Y年%m月%d日')}",
             align="R", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(4)

    # 請求先
    pdf.set_font("YuGoth", "B", 12)
    pdf.cell(0, 7, f"{customer_name} 御中", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("YuGoth", "", 9)
    pdf.cell(0, 5, f"顧客コード: {customer_code}", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(5)

    # 発行元（自社）
    pdf.set_x(110)
    pdf.set_font("YuGoth", "B", 11)
    pdf.cell(90, 6, "株式会社テクノプレシジョン", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("YuGoth", "", 9)
    pdf.set_x(110)
    pdf.cell(90, 5, "〒222-0033 神奈川県横浜市港北区", new_x="LMARGIN", new_y="NEXT")
    pdf.set_x(110)
    pdf.cell(90, 5, "新横浜1-1-1", new_x="LMARGIN", new_y="NEXT")
    pdf.set_x(110)
    pdf.cell(90, 5, "TEL: 045-XXX-XXXX", new_x="LMARGIN", new_y="NEXT")
    pdf.set_x(110)
    pdf.cell(90, 5, "登録番号: T1234567890123", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(10)

    # 合計金額（枠付き）
    subtotal = sum(qty * unit for _, _, qty, unit in items)
    tax = int(subtotal * 0.1)
    total = subtotal + tax
    pdf.set_font("YuGoth", "B", 14)
    pdf.set_fill_color(240, 245, 255)
    pdf.cell(60, 14, "ご請求金額", border=1, align="C", fill=True)
    pdf.cell(130, 14, f"¥ {total:,} -", border=1, align="R", fill=True,
             new_x="LMARGIN", new_y="NEXT")
    pdf.set_fill_color(255, 255, 255)
    pdf.ln(6)

    # 明細
    pdf.table_header(["品目コード", "品名", "数量", "単価", "金額"],
                     [30, 80, 20, 30, 30])
    for code, name, qty, unit in items:
        pdf.table_row([code, name, f"{qty:,}", f"{unit:,}", f"{qty * unit:,}"],
                      [30, 80, 20, 30, 30])

    pdf.set_font("YuGoth", "B", 10)
    pdf.cell(130, 7, "小計", border=1, align="R")
    pdf.cell(60, 7, f"¥ {subtotal:,}", border=1, align="R",
             new_x="LMARGIN", new_y="NEXT")
    pdf.cell(130, 7, "消費税 (10%)", border=1, align="R")
    pdf.cell(60, 7, f"¥ {tax:,}", border=1, align="R",
             new_x="LMARGIN", new_y="NEXT")
    pdf.set_fill_color(255, 242, 204)
    pdf.cell(130, 8, "合計", border=1, align="R", fill=True)
    pdf.cell(60, 8, f"¥ {total:,}", border=1, align="R", fill=True,
             new_x="LMARGIN", new_y="NEXT")
    pdf.set_fill_color(255, 255, 255)
    pdf.ln(6)

    # お支払情報
    pdf.h3("■ お支払について")
    pdf.set_font("YuGoth", "", 10)
    pdf.kv("お支払期日", due_date.strftime("%Y年%m月%d日"))
    pdf.kv("お支払方法", "銀行振込")
    pdf.kv("振込先", "みずほ銀行 新横浜支店 普通 1234567")
    pdf.kv("口座名義", "カ）テクノプレシジョン")
    pdf.ln(6)

    # 社印
    y_stamp = pdf.get_y()
    pdf.stamp("会社印", x=170, y=y_stamp + 10)

    pdf.output(str(BASE / output_name))
    print(f"Created: {output_name}")


# ============================================================
# 3. 稟議書PDF（価格マスタ変更）
# ============================================================
def gen_ringi_pdf(ringi_no, apply_date, applicant, subject, product_code,
                  customer_code, old_price, new_price, reason,
                  approvals, output_name):
    pdf = JPPDF()
    pdf.add_page()

    # タイトル
    pdf.set_font("YuGoth", "B", 16)
    pdf.cell(0, 10, "稟 議 書", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("YuGoth", "", 10)
    pdf.cell(0, 5, f"稟議番号: {ringi_no}", align="R", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(3)

    # 申請情報
    pdf.kv("件名", subject, key_w=30)
    pdf.kv("申請日", apply_date.strftime("%Y年%m月%d日"), key_w=30)
    pdf.kv("申請者", applicant, key_w=30)
    pdf.kv("承認種別", "価格マスタ変更（顧客別単価）", key_w=30)
    pdf.ln(5)

    # 変更内容
    pdf.h2("1. 変更内容")
    pdf.set_fill_color(230, 237, 244)
    pdf.table_header(["項目", "変更前", "変更後", "変更率"],
                     [50, 45, 45, 40])
    change_rate = (new_price - old_price) / old_price * 100
    pdf.table_row([f"製品 {product_code}", f"¥{old_price:,}", f"¥{new_price:,}",
                   f"{change_rate:+.1f}%"], [50, 45, 45, 40])
    pdf.table_row(["顧客", customer_code, "（同一）", "-"], [50, 45, 45, 40], fill=True)
    pdf.ln(5)

    # 変更理由
    pdf.h2("2. 変更理由")
    pdf.body(reason, size=10)
    pdf.ln(3)

    # 適用開始日
    pdf.h2("3. 適用開始日")
    pdf.body("稟議承認後、次期受注分より適用")
    pdf.ln(5)

    # 承認経路
    pdf.h2("4. 承認経路")
    pdf.set_font("YuGoth", "B", 10)
    pdf.cell(40, 7, "役割", border=1, align="C", fill=True)
    pdf.cell(45, 7, "氏名", border=1, align="C", fill=True)
    pdf.cell(40, 7, "承認日時", border=1, align="C", fill=True)
    pdf.cell(30, 7, "承認印", border=1, align="C", fill=True, new_x="LMARGIN", new_y="NEXT")

    for role, name, dt in approvals:
        pdf.set_font("YuGoth", "", 10)
        pdf.cell(40, 14, role, border=1, align="C")
        pdf.cell(45, 14, name, border=1, align="C")
        pdf.cell(40, 14, dt, border=1, align="C")
        # 承認スタンプ枠
        x_stamp = pdf.get_x()
        y_stamp = pdf.get_y()
        pdf.cell(30, 14, "", border=1, new_x="LMARGIN", new_y="NEXT")
        pdf.stamp("承認", x=x_stamp + 15, y=y_stamp + 7)

    pdf.ln(5)

    # 添付
    pdf.set_font("YuGoth", "", 9)
    pdf.multi_cell(0, 5, "【添付資料】原材料費上昇根拠資料（V-20001からの値上げ通知）、過去3年の販売単価推移表")

    pdf.output(str(BASE / output_name))
    print(f"Created: {output_name}")


# ============================================================
# 4. 売掛金年齢表の「低解像度スキャンPDF」（判断保留ケース用）
# ============================================================
def gen_lowres_aging_pdf():
    """
    意図的に低解像度のスキャン画像を含む PDF を作成。
    承認印が判読不能な状態を再現。
    """
    # まず、きれいに年齢表のイメージを作る
    img = Image.new("RGB", (1600, 2200), (255, 255, 255))
    d = ImageDraw.Draw(img)
    font_path = "C:/Windows/Fonts/YuGothM.ttc"
    font_bold = "C:/Windows/Fonts/YuGothB.ttc"
    fh1 = ImageFont.truetype(font_bold, 48)
    fh2 = ImageFont.truetype(font_bold, 28)
    fb = ImageFont.truetype(font_path, 20)
    fs = ImageFont.truetype(font_path, 16)

    d.text((60, 50), "2025年11月末 売掛金年齢分析表", font=fh1, fill=(20, 20, 60))
    d.text((60, 130), "基準日: 2025/11/30  /  作成: 経理部課長 高橋 美咲", font=fb, fill=(60, 60, 60))

    # テーブル
    headers = ["顧客コード", "顧客名", "残高合計(円)", "0-30日", "31-60日", "61-90日", "91-120日", "120日超"]
    col_x = [60, 220, 500, 720, 900, 1080, 1260, 1440]
    col_w = [160, 280, 220, 180, 180, 180, 180, 160]
    y0 = 200
    d.rectangle([60, y0, 60 + sum(col_w), y0 + 50], fill=(31, 78, 120))
    for i, (h, x, w) in enumerate(zip(headers, col_x, col_w)):
        d.text((x + 10, y0 + 12), h, font=fh2, fill=(255, 255, 255))

    samples = [
        ("C-10001", "トヨタエンジニアリング", "128,540,000", "90,000,000", "26,000,000", "8,500,000", "3,000,000", "1,040,000"),
        ("C-10002", "本田技研部品", "87,320,000", "61,200,000", "17,400,000", "5,200,000", "2,800,000", "720,000"),
        ("C-10003", "日産精密パーツ", "56,780,000", "39,700,000", "11,300,000", "3,400,000", "1,700,000", "680,000"),
        ("C-10007", "三菱自動車部品販売", "23,450,000", "12,000,000", "4,500,000", "2,100,000", "1,300,000", "3,550,000"),
        ("C-10011", "東京エレクトロン", "156,890,000", "109,800,000", "31,400,000", "9,400,000", "4,700,000", "1,590,000"),
        ("C-10017", "日立ハイテク部品", "42,180,000", "22,000,000", "8,500,000", "3,200,000", "1,900,000", "6,580,000"),
        ("C-10023", "丸紅情報システムズ", "36,720,000", "19,500,000", "7,300,000", "2,800,000", "1,600,000", "5,520,000"),
    ]
    y = y0 + 50
    for r_idx, row in enumerate(samples):
        bg = (255, 255, 255) if r_idx % 2 == 0 else (240, 245, 252)
        d.rectangle([60, y, 60 + sum(col_w), y + 45], fill=bg, outline=(200, 200, 200))
        for i, v in enumerate(row):
            d.text((col_x[i] + 10, y + 12), v, font=fb, fill=(20, 20, 20))
        y += 45

    # 承認欄
    y += 60
    d.text((60, y), "■ 承認記録", font=fh2, fill=(20, 20, 60))
    y += 50
    d.text((60, y), "作成: 高橋 美咲（経理部課長）", font=fb, fill=(40, 40, 40))

    # 承認印を描画
    d.ellipse([520, y - 5, 620, y + 45], outline=(200, 30, 30), width=3)
    d.text((540, y + 8), "高橋", font=fh2, fill=(200, 30, 30))
    d.text((720, y + 10), "2025/12/08", font=fb, fill=(40, 40, 40))

    y += 70
    d.text((60, y), "承認: 佐藤 一郎（経理部長）", font=fb, fill=(40, 40, 40))
    # 経理部長印 ← ここが意図的に「にじんだ・不鮮明」に
    d.ellipse([520, y - 5, 620, y + 45], outline=(200, 30, 30), width=3)
    d.text((548, y + 8), "佐藤", font=fh2, fill=(200, 30, 30))
    d.text((720, y + 10), "2025/12/??", font=fb, fill=(40, 40, 40))  # 日付も不鮮明

    # 意図的な劣化処理（承認印領域）
    # 承認印部分だけを切り出してぼかし、戻す
    box = (500, y - 20, 900, y + 60)
    crop = img.crop(box)
    crop = crop.filter(ImageFilter.GaussianBlur(radius=4))
    # さらにノイズ（低解像度風に ダウンサンプル→アップサンプル）
    small = crop.resize((crop.width // 6, crop.height // 6))
    crop = small.resize(crop.size, Image.NEAREST)
    img.paste(crop, box)

    # 右上にスキャン日スタンプ（白黒FAXぽい）
    d.rectangle([1340, 40, 1550, 120], outline=(50, 50, 50), width=2)
    d.text((1360, 55), "SCAN", font=fh2, fill=(50, 50, 50))
    d.text((1360, 85), "2025/12/10", font=fb, fill=(50, 50, 50))

    # 画像全体をPDFに埋め込む
    img_path = BASE / "_temp_aging.png"
    img.save(img_path, "PNG")

    pdf = JPPDF(orientation="P", format="A4")
    pdf.add_page()
    pdf.set_font("YuGoth", "B", 14)
    pdf.cell(0, 8, "売掛金年齢分析表 【経理部長承認済】", align="C",
             new_x="LMARGIN", new_y="NEXT")
    pdf.ln(3)
    pdf.image(str(img_path), x=10, y=20, w=190)
    pdf.set_y(240)
    pdf.set_font("YuGoth", "", 8)
    pdf.cell(0, 5, "※ 本書類は原本をPDF化したものです",
             align="L", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 5, "保管: 経理部 / 文書番号: AR-AGE-202511-001",
             align="L", new_x="LMARGIN", new_y="NEXT")

    pdf.output(str(BASE / "PLC-S-005_売掛金年齢表_経理部長承認PDF_低解像度.pdf"))
    img_path.unlink()
    print("Created: PLC-S-005_売掛金年齢表_経理部長承認PDF_低解像度.pdf")


# ============================================================
# Main
# ============================================================
if __name__ == "__main__":
    # 注文書3件（リアリティのため異なる顧客・製品）
    gen_order_pdf(
        "PO-HONDA-2025-8843",  # 顧客側の注文書番号
        date(2025, 11, 10),
        "C-10002",
        "本田技研部品株式会社",
        "〒350-1305 埼玉県狭山市下奥富1-1",
        [("P-30006", "トランスミッションシャフト", 1000, 12500)],
        date(2025, 12, 15),
        "松本 香織",
        "PLC-S-001_注文書_ORD-2025-1420_HONDA.pdf",
    )

    gen_order_pdf(
        "TEL-EPP-2025-0934",
        date(2025, 10, 25),
        "C-10011",
        "東京エレクトロン購買部",
        "〒107-6325 東京都港区赤坂5-3-1",
        [
            ("P-30011", "ウェハー搬送ロボット用シャフト A", 80, 18500),
            ("P-30014", "エッチング装置チャンバ部品", 40, 38500),
        ],
        date(2025, 11, 30),
        "藤田 修",
        "PLC-S-001_注文書_ORD-2025-0412_TEL.pdf",
    )

    # 承認1日遅れの例外ケース
    gen_order_pdf(
        "NIKO-2025-4521",
        date(2025, 12, 3),
        "C-10015",
        "株式会社ニコン精機",
        "〒140-8601 東京都品川区西大井1-6-3",
        [("P-30015", "ウェハーチャックベース", 1500, 12800)],
        date(2026, 1, 20),
        "松本 香織",
        "PLC-S-001_注文書_ORD-2025-1876_NIKON_承認遅延サンプル14.pdf",
    )

    # 請求書
    gen_invoice_pdf(
        "INV-202511-0234",
        date(2025, 11, 30),
        "C-10002",
        "本田技研部品株式会社",
        [("P-30006", "トランスミッションシャフト", 1000, 12500)],
        date(2025, 12, 31),
        "PLC-S-003_請求書_INV-202511-0234.pdf",
    )

    # 稟議書 (価格マスタ変更)
    gen_ringi_pdf(
        "W-2025-1876",
        date(2025, 10, 15),
        "松本 香織（営業部主任）",
        "顧客C-10011向け製品P-30011の単価改定申請",
        "P-30011",
        "C-10011",
        18500,
        19200,
        "原材料費上昇（V-20008 日立金属ファインテック社からの特殊合金材の価格改定通知）に対応するため、"
        "顧客東京エレクトロン購買部への納入単価を+3.8%（¥700）上げる旨を交渉・合意した。"
        "年間販売数量約4,800本、年間売上増加額は ¥3,360,000 の見込み。",
        [
            ("営業部課長", "斎藤 次郎 (SLS002)", "2025/10/15 14:30"),
            ("営業本部長", "田中 太郎 (SLS001)", "2025/10/16 10:15"),
        ],
        "PLC-S-007_価格マスタ変更稟議_W-2025-1876.pdf",
    )

    # 低解像度スキャン PDF（判断保留ケース）
    gen_lowres_aging_pdf()
