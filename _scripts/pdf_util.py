"""
日本語PDF生成ユーティリティ
fpdf2 + Yu Gothic フォントを使用
"""
from fpdf import FPDF
from pathlib import Path

FONT_PATH = "C:/Windows/Fonts/YuGothM.ttc"
FONT_BOLD_PATH = "C:/Windows/Fonts/YuGothB.ttc"


class JPPDF(FPDF):
    """日本語対応PDFクラス"""
    def __init__(self, orientation="P", unit="mm", format="A4"):
        super().__init__(orientation=orientation, unit=unit, format=format)
        # フォントを追加
        self.add_font("YuGoth", style="", fname=FONT_PATH)
        self.add_font("YuGoth", style="B", fname=FONT_BOLD_PATH)
        self.set_auto_page_break(auto=True, margin=15)

    def h1(self, text):
        self.set_font("YuGoth", "B", 16)
        self.cell(0, 10, text, new_x="LMARGIN", new_y="NEXT")
        self.ln(2)

    def h2(self, text):
        self.set_font("YuGoth", "B", 13)
        self.set_fill_color(230, 237, 244)
        self.cell(0, 8, "  " + text, new_x="LMARGIN", new_y="NEXT", fill=True)
        self.ln(2)

    def h3(self, text):
        self.set_font("YuGoth", "B", 11)
        self.cell(0, 7, text, new_x="LMARGIN", new_y="NEXT")

    def body(self, text, size=10):
        self.set_font("YuGoth", "", size)
        self.multi_cell(0, 5.5, text)
        self.ln(1)

    def kv(self, key, value, key_w=40, font_size=10):
        """キー：値 の横並び"""
        self.set_font("YuGoth", "B", font_size)
        self.cell(key_w, 6, key, border=1, fill=True)
        self.set_font("YuGoth", "", font_size)
        self.cell(0, 6, " " + str(value), border=1, new_x="LMARGIN", new_y="NEXT")

    def table_header(self, headers, widths, fill_color=(31, 78, 120), text_color=(255, 255, 255)):
        self.set_font("YuGoth", "B", 9)
        self.set_fill_color(*fill_color)
        self.set_text_color(*text_color)
        for h, w in zip(headers, widths):
            self.cell(w, 8, h, border=1, align="C", fill=True)
        self.ln()
        self.set_text_color(0, 0, 0)

    def table_row(self, values, widths, height=6, align="L", fill=False, fill_color=(255, 242, 204)):
        self.set_font("YuGoth", "", 9)
        if fill:
            self.set_fill_color(*fill_color)
        for v, w in zip(values, widths):
            self.cell(w, height, str(v), border=1, align=align, fill=fill)
        self.ln()

    def stamp(self, text, x=None, y=None):
        """承認印風の丸スタンプ"""
        cx = x if x is not None else self.get_x() + 10
        cy = y if y is not None else self.get_y() + 5
        self.set_draw_color(200, 30, 30)
        self.set_line_width(0.5)
        self.circle(cx, cy, 8)
        self.set_text_color(200, 30, 30)
        self.set_font("YuGoth", "B", 9)
        self.text(cx - 5, cy + 2, text)
        self.set_text_color(0, 0, 0)
        self.set_draw_color(0, 0, 0)

    def signature_block(self, title, name, date_str, x=None, y=None):
        """承認ブロック"""
        if x is None:
            x = self.get_x()
        if y is None:
            y = self.get_y()
        self.rect(x, y, 40, 18)
        self.set_xy(x, y)
        self.set_font("YuGoth", "B", 8)
        self.cell(40, 5, title, align="C")
        self.set_xy(x, y + 5)
        self.set_font("YuGoth", "", 8)
        self.cell(40, 5, name, align="C")
        self.set_xy(x, y + 10)
        self.cell(40, 4, date_str, align="C")


def test_pdf():
    """テストPDF生成"""
    pdf = JPPDF()
    pdf.add_page()
    pdf.h1("日本語PDF生成テスト")
    pdf.h2("1. 企業情報")
    pdf.kv("会社名", "デモA株式会社")
    pdf.kv("所在地", "神奈川県横浜市港北区")
    pdf.kv("資本金", "1,200,000,000円")
    pdf.kv("従業員数", "連結820名（単体580名）")

    pdf.h2("2. 本文")
    pdf.body("このテストは日本語フォントの表示を確認するためのものです。"
             "漢字・ひらがな・カタカナ・英数字・記号（※・●・○）"
             "がすべて正しく表示されることを確認してください。")

    pdf.h2("3. テーブル")
    pdf.table_header(["項目", "内容", "数量"], [40, 80, 30])
    pdf.table_row(["切削加工品A", "エンジン部品", "1,000個"], [40, 80, 30])
    pdf.table_row(["プレス加工品B", "ブラケット", "5,000個"], [40, 80, 30], fill=True)

    pdf.ln(5)
    pdf.h2("4. 承認印")
    pdf.stamp("承認")

    out = Path(r"C:\Users\nyham\work\demo_data\_scripts\_pdf_test.pdf")
    pdf.output(str(out))
    print(f"OK: {out}")


if __name__ == "__main__":
    test_pdf()
