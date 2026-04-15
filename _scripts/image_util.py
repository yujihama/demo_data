"""
SAP画面・ワークフロー等のスクリーンショット風画像生成ユーティリティ
Pillow + 日本語フォント使用
"""
from PIL import Image, ImageDraw, ImageFont
from pathlib import Path

FONT_PATH = "C:/Windows/Fonts/YuGothM.ttc"
FONT_BOLD = "C:/Windows/Fonts/YuGothB.ttc"


def _font(size, bold=False):
    return ImageFont.truetype(FONT_BOLD if bold else FONT_PATH, size)


def sap_screenshot(
    title,
    trx_code,
    fields,  # [(label, value), ...]
    grid_headers=None,
    grid_rows=None,
    status_bar="レコードが保存されました",
    output_path=None,
    size=(1024, 640),
):
    """SAP風画面のスクリーンショットを生成"""
    img = Image.new("RGB", size, (240, 244, 250))
    d = ImageDraw.Draw(img)

    # タイトルバー（SAP青）
    d.rectangle([0, 0, size[0], 30], fill=(11, 68, 122))
    d.text((10, 6), f"SAP S/4HANA  -  {title}", font=_font(12, bold=True), fill=(255, 255, 255))
    d.text((size[0] - 120, 6), "_  □  X", font=_font(12, bold=True), fill=(255, 255, 255))

    # メニューバー
    d.rectangle([0, 30, size[0], 58], fill=(225, 230, 240))
    d.text((10, 35), "ファイル(F)   編集(E)   表示(V)   移動(G)   ヘルプ(H)",
           font=_font(11), fill=(30, 30, 30))

    # トランザクションコード枠
    d.rectangle([10, 68, 120, 92], outline=(100, 100, 100), width=1, fill=(255, 255, 255))
    d.text((16, 72), trx_code, font=_font(13, bold=True), fill=(0, 0, 120))

    # ツールバー
    d.rectangle([130, 68, size[0] - 10, 92], outline=(200, 200, 200), fill=(245, 245, 245))
    d.text((140, 72), "[保存] [戻る] [中止] [終了] [印刷] [検索] [登録] [変更] [照会]",
           font=_font(10), fill=(50, 50, 50))

    # フィールド領域
    y = 110
    d.rectangle([10, y, size[0] - 10, y + 30 + len(fields) * 26],
                outline=(180, 180, 180), fill=(255, 255, 255))
    y += 10
    for label, value in fields:
        d.text((20, y), label, font=_font(11, bold=True), fill=(60, 60, 60))
        d.rectangle([180, y - 2, 550, y + 18], outline=(200, 200, 200), fill=(255, 255, 220))
        d.text((185, y), str(value), font=_font(11), fill=(20, 20, 20))
        y += 26
    y += 20

    # グリッド領域
    if grid_headers and grid_rows:
        col_width = (size[0] - 40) // len(grid_headers)
        # グリッドヘッダ
        d.rectangle([10, y, size[0] - 10, y + 24], fill=(70, 100, 140))
        for i, h in enumerate(grid_headers):
            d.text((20 + i * col_width, y + 5), h, font=_font(10, bold=True), fill=(255, 255, 255))
        y += 24
        # 行
        for r_idx, row in enumerate(grid_rows):
            bg = (255, 255, 255) if r_idx % 2 == 0 else (245, 248, 252)
            d.rectangle([10, y, size[0] - 10, y + 22], fill=bg, outline=(220, 220, 220))
            for i, v in enumerate(row):
                d.text((20 + i * col_width, y + 4), str(v), font=_font(10), fill=(20, 20, 20))
            y += 22

    # ステータスバー
    d.rectangle([0, size[1] - 22, size[0], size[1]], fill=(60, 100, 60))
    d.text((10, size[1] - 18), "✓ " + status_bar, font=_font(10, bold=True), fill=(255, 255, 255))

    if output_path:
        img.save(output_path, "PNG")
    return img


def workflow_screenshot(
    wf_id,
    title,
    applicant_name,
    approval_chain,  # [(name, role, date, status), ...]
    amount=None,
    comments=None,
    output_path=None,
    size=(900, 680),
):
    """ワークフロー承認画面風のスクリーンショット"""
    img = Image.new("RGB", size, (250, 250, 253))
    d = ImageDraw.Draw(img)

    # ヘッダ
    d.rectangle([0, 0, size[0], 60], fill=(40, 80, 130))
    d.text((20, 10), "稟議ワークフロー管理システム", font=_font(16, bold=True), fill=(255, 255, 255))
    d.text((20, 34), f"申請書ID: {wf_id}", font=_font(11), fill=(220, 230, 245))

    # メイン情報
    y = 80
    d.rectangle([20, y, size[0] - 20, y + 130], outline=(180, 180, 200), fill=(255, 255, 255))
    d.text((30, y + 10), title, font=_font(14, bold=True), fill=(30, 30, 60))

    d.text((30, y + 40), "申請者:", font=_font(11, bold=True), fill=(80, 80, 80))
    d.text((110, y + 40), applicant_name, font=_font(11), fill=(20, 20, 20))

    if amount is not None:
        d.text((30, y + 62), "申請金額:", font=_font(11, bold=True), fill=(80, 80, 80))
        d.text((110, y + 62), f"¥ {amount:,}", font=_font(13, bold=True), fill=(180, 40, 40))

    d.text((30, y + 84), "申請日:", font=_font(11, bold=True), fill=(80, 80, 80))
    d.text((110, y + 84), approval_chain[0][2] if approval_chain else "-",
           font=_font(11), fill=(20, 20, 20))

    if comments:
        d.text((30, y + 106), "申請理由:", font=_font(11, bold=True), fill=(80, 80, 80))
        d.text((110, y + 106), comments[:60], font=_font(10), fill=(40, 40, 40))

    # 承認チェーン
    y += 150
    d.text((20, y), "■ 承認経路", font=_font(12, bold=True), fill=(30, 30, 60))
    y += 25

    for i, (name, role, date, status) in enumerate(approval_chain):
        x = 30 + i * 175
        # ボックス
        color = (120, 180, 120) if status == "承認" else (220, 220, 120) if status == "保留" else (200, 120, 120) if status == "却下" else (180, 180, 180)
        d.rectangle([x, y, x + 155, y + 100], outline=color, width=2, fill=(250, 250, 250))
        d.text((x + 5, y + 5), role, font=_font(10, bold=True), fill=(60, 60, 60))
        d.text((x + 5, y + 25), name, font=_font(11, bold=True), fill=(20, 20, 20))
        d.text((x + 5, y + 50), date, font=_font(9), fill=(100, 100, 100))

        # 承認スタンプ
        if status == "承認":
            d.ellipse([x + 95, y + 60, x + 145, y + 95], outline=(200, 30, 30), width=2)
            d.text((x + 108, y + 72), "承認", font=_font(11, bold=True), fill=(200, 30, 30))
        elif status == "保留":
            d.text((x + 100, y + 72), "(保留中)", font=_font(10, bold=True), fill=(180, 140, 20))
        elif status == "却下":
            d.text((x + 100, y + 72), "(却下)", font=_font(10, bold=True), fill=(200, 30, 30))

        # 矢印
        if i < len(approval_chain) - 1:
            d.line([x + 160, y + 50, x + 170, y + 50], fill=(100, 100, 100), width=2)
            d.polygon([(x + 170, y + 45), (x + 175, y + 50), (x + 170, y + 55)],
                      fill=(100, 100, 100))

    # フッタ
    d.rectangle([0, size[1] - 30, size[0], size[1]], fill=(230, 230, 240))
    d.text((20, size[1] - 22), f"印刷日時: 2026/02/10 11:23:45 / User: IA001 長谷川 剛",
           font=_font(9), fill=(80, 80, 80))

    if output_path:
        img.save(output_path, "PNG")
    return img


def table_image(
    title,
    headers,
    rows,
    widths=None,
    output_path=None,
    caption=None,
    header_color=(31, 78, 120),
    size=None,
):
    """汎用のテーブル画像（アクセスマトリクス等用）"""
    font_title = _font(16, bold=True)
    font_h = _font(10, bold=True)
    font_b = _font(10)
    row_h = 24
    pad = 20

    if widths is None:
        widths = [max(80, len(str(h)) * 12) for h in headers]
    total_w = sum(widths) + pad * 2
    total_h = pad * 2 + 40 + len(rows) * row_h + (50 if caption else 10) + 30

    if size:
        total_w, total_h = size
    img = Image.new("RGB", (total_w, total_h), (255, 255, 255))
    d = ImageDraw.Draw(img)

    # タイトル
    d.text((pad, pad), title, font=font_title, fill=(20, 40, 80))
    y = pad + 35

    # ヘッダ
    x = pad
    for h, w in zip(headers, widths):
        d.rectangle([x, y, x + w, y + row_h], fill=header_color, outline=(0, 0, 0))
        d.text((x + 5, y + 4), str(h), font=font_h, fill=(255, 255, 255))
        x += w
    y += row_h

    # 行
    for r_idx, row in enumerate(rows):
        x = pad
        bg = (255, 255, 255) if r_idx % 2 == 0 else (240, 245, 252)
        for v, w in zip(row, widths):
            d.rectangle([x, y, x + w, y + row_h], fill=bg, outline=(200, 200, 200))
            vstr = str(v)
            align_center = vstr in ("○", "●", "×", "", "-")
            tx = x + (w // 2 - 6 if align_center else 5)
            d.text((tx, y + 4), vstr, font=font_b, fill=(20, 20, 20))
            x += w
        y += row_h

    if caption:
        d.text((pad, y + 10), caption, font=font_b, fill=(100, 100, 100))

    if output_path:
        img.save(output_path, "PNG")
    return img


def warehouse_photo(
    warehouse_name,
    section,
    date,
    inspector,
    output_path,
    size=(800, 600),
    scene_type="rack",
):
    """倉庫の「撮影された写真」風画像（模擬）"""
    img = Image.new("RGB", size, (70, 75, 90))
    d = ImageDraw.Draw(img)

    # 床
    d.polygon([(0, 400), (size[0], 400), (size[0], size[1]), (0, size[1])], fill=(100, 90, 75))
    for i in range(5):
        y = 400 + i * 40
        d.line([(0, y), (size[0], y)], fill=(85, 75, 62), width=1)

    # 壁
    d.rectangle([0, 0, size[0], 400], fill=(175, 180, 190))

    if scene_type == "rack":
        # 棚
        rack_color = (90, 105, 130)
        for col in range(4):
            x = 80 + col * 170
            # 縦柱
            d.rectangle([x - 5, 120, x + 5, 400], fill=rack_color)
            d.rectangle([x + 140, 120, x + 150, 400], fill=rack_color)
            # 棚板
            for row in range(3):
                sy = 160 + row * 80
                d.rectangle([x - 5, sy, x + 150, sy + 6], fill=rack_color)
                # 箱
                for b in range(2):
                    bx = x + 10 + b * 65
                    bc_hue = (140 + b * 25, 110 + row * 10, 80)
                    d.rectangle([bx, sy - 40, bx + 55, sy], fill=bc_hue, outline=(50, 50, 50))
                    d.text((bx + 8, sy - 30), f"P-3{col:02d}{row}{b}", font=_font(8, bold=True),
                           fill=(255, 255, 255))

        # ラベル
        for col in range(4):
            x = 80 + col * 170
            d.rectangle([x + 30, 110, x + 130, 130], fill=(255, 230, 100), outline=(0, 0, 0))
            d.text((x + 50, 113), f"区画 {section}-{col + 1}", font=_font(10, bold=True), fill=(0, 0, 0))

    elif scene_type == "floor":
        # パレット在庫風
        for r in range(3):
            for c in range(5):
                x = 50 + c * 140
                y = 200 + r * 70
                d.rectangle([x, y, x + 120, y + 40], fill=(160, 100, 50), outline=(80, 50, 30))
                d.rectangle([x + 5, y - 30, x + 115, y], fill=(200, 170, 110), outline=(100, 80, 50))
                d.text((x + 20, y - 25), f"P-{r}{c}", font=_font(11, bold=True), fill=(40, 30, 20))

    # 白いバンド（写真メタデータ）
    d.rectangle([0, size[1] - 50, size[0], size[1]], fill=(255, 255, 255))
    d.text((10, size[1] - 45), f"倉庫: {warehouse_name}  /  区画: {section}",
           font=_font(12, bold=True), fill=(20, 20, 20))
    d.text((10, size[1] - 25), f"撮影日時: {date}  /  撮影者: {inspector}",
           font=_font(11), fill=(60, 60, 60))
    d.text((size[0] - 140, size[1] - 25), "[棚卸立会写真]", font=_font(10, bold=True), fill=(180, 30, 30))

    # 右上に日付印
    d.rectangle([size[0] - 100, 10, size[0] - 10, 55], fill=(255, 230, 220), outline=(200, 30, 30), width=2)
    d.text((size[0] - 90, 15), date.split()[0] if " " in date else date,
           font=_font(11, bold=True), fill=(200, 30, 30))

    img.save(output_path, "JPEG", quality=80)
    return img


def test_images():
    """画像生成テスト"""
    out_dir = Path(r"C:\Users\nyham\work\demo_data\_scripts")

    sap_screenshot(
        "受注登録 - 照会",
        "VA01",
        [
            ("受注伝票番号", "ORD-2025-1420"),
            ("受注日", "2025/11/10"),
            ("販売先", "C-10002 本田技研部品株式会社"),
            ("販売担当", "SLS004 松本 香織"),
            ("与信限度額", "300,000,000 JPY"),
            ("当該受注金額", "12,500,000 JPY"),
            ("現在与信残額", "245,820,000 JPY (余裕あり)"),
        ],
        grid_headers=["品目", "製品コード", "数量", "単価", "金額"],
        grid_rows=[
            ["トランスミッションシャフト", "P-30006", "1,000", "12,500", "12,500,000"],
        ],
        status_bar="受注伝票 ORD-2025-1420 が保存されました。",
        output_path=str(out_dir / "_sap_test.png"),
    )

    workflow_screenshot(
        "W-2025-1876",
        "価格マスタ変更申請",
        "松本 香織（営業部主任）",
        [
            ("松本 香織", "申請者（営業部主任）", "2025/10/15 09:12", "申請"),
            ("斎藤 次郎", "課長（営業部）", "2025/10/15 14:30", "承認"),
            ("田中 太郎", "本部長（営業本部）", "2025/10/16 10:15", "承認"),
        ],
        amount=None,
        comments="顧客C-10011向け製品 P-30011 の価格改定（原材料費上昇のため）",
        output_path=str(out_dir / "_wf_test.png"),
    )

    table_image(
        "SAPアクセス権 ユーザ×ロール マトリクス（販売領域抜粋）",
        ["ユーザID", "氏名", "SD_MGR", "SD_SUP", "SD_USER", "PO_CREATE", "PO_APPROVE"],
        [
            ["SLS001", "田中 太郎", "●", "", "", "", ""],
            ["SLS002", "斎藤 次郎", "", "●", "", "", ""],
            ["SLS004", "松本 香織", "", "", "●", "", ""],
            ["PUR003", "清水 智明", "", "", "", "●", ""],
            ["PUR004", "山田 純一", "", "", "", "●", "● ⚠"],
        ],
        widths=[80, 100, 70, 70, 70, 90, 100],
        caption="⚠ PUR004 にPO_CREATEとPO_APPROVEが両方付与されている（SoD違反）",
        output_path=str(out_dir / "_matrix_test.png"),
    )

    warehouse_photo(
        "本社倉庫A",
        "A-3",
        "2025/09/26 14:32",
        "橋本 明（倉庫課長）",
        str(out_dir / "_warehouse_test.jpg"),
        scene_type="rack",
    )

    print("OK: 4 test images generated")


if __name__ == "__main__":
    test_images()
