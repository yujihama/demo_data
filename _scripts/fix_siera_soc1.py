"""Fix SIer-A SOC1 to clean (unqualified) opinion - aligns with 運用評価=有効/例外0"""
from fpdf import FPDF
from pathlib import Path
import warnings
warnings.filterwarnings('ignore', category=DeprecationWarning)

ROOT = Path(r"C:\Users\nyham\work\demo_data\4.evidence\ITGC")
FONT_REG = r"C:\Windows\Fonts\YuGothM.ttc"
FONT_BLD = r"C:\Windows\Fonts\YuGothB.ttc"

pdf_path = ROOT / "SOC1_TypeII_Report_SIerA_FY2024.pdf"

class MyPDF(FPDF):
    pass

pdf = MyPDF(orientation='P', unit='mm', format='A4')
pdf.set_margins(left=15, top=15, right=15)
pdf.set_auto_page_break(auto=True, margin=15)
pdf.add_font('YG', '', FONT_REG, uni=True)
pdf.add_font('YGB', '', FONT_BLD, uni=True)

def p(text, size=10, bold=False):
    pdf.set_x(pdf.l_margin)
    pdf.set_font('YGB' if bold else 'YG', '', size)
    pdf.multi_cell(w=180, h=6, text=text)

def h(text, size=14):
    pdf.set_x(pdf.l_margin)
    pdf.set_font('YGB', '', size)
    pdf.cell(180, 10, text, ln=1)

def space(lines=1):
    pdf.ln(lines*2)

# Cover
pdf.add_page()
pdf.set_font('YGB', '', 20)
pdf.ln(40)
pdf.cell(180, 12, 'SOC 1 Type II Report', ln=1, align='C')
space(2)
pdf.set_font('YG', '', 11)
pdf.multi_cell(180, 7, "Report on Management's Description of a Service Organization's System and the Suitability of the Design and Operating Effectiveness of Controls", align='C')
pdf.ln(15)
pdf.set_font('YGB', '', 14)
pdf.cell(180, 10, 'Service Organization: 外部委託先SIer-A', ln=1, align='C')
pdf.ln(10)
pdf.set_font('YG', '', 11)
pdf.cell(180, 7, 'Report Period: April 1, 2024 - March 31, 2025', ln=1, align='C')
pdf.cell(180, 7, 'Issued: May 20, 2025', ln=1, align='C')
pdf.ln(40)
pdf.cell(180, 7, 'Prepared in accordance with SSAE No. 18 (AT-C 320)', ln=1, align='C')
pdf.cell(180, 7, 'By: Independent Service Auditor XYZ CPA Firm', ln=1, align='C')

# TOC
pdf.add_page()
h('Table of Contents', 16)
space()
pdf.set_font('YG', '', 11)
toc = [
    ("Section I   Independent Service Auditor's Report", "3"),
    ("Section II  Management's Assertion", "5"),
    ("Section III Description of the Service Organization's System", "7"),
    ("Section IV  Control Objectives, Related Controls, and Test Results", "12"),
    ("Section V   Complementary User Entity Controls (CUECs)", "22"),
]
for entry, page in toc:
    pdf.set_x(pdf.l_margin)
    pdf.cell(140, 7, entry)
    pdf.cell(40, 7, f'... {page}', ln=1, align='R')

# Section I - CLEAN OPINION
pdf.add_page()
h("Section I. Independent Service Auditor's Report")
space()
p("To the Management of 外部委託先SIer-A and other specified parties:")
space()
p("Scope: We have examined 外部委託先SIer-A's description of its SAP ERP Operations and Development Services (the System) throughout the period April 1, 2024 to March 31, 2025, and the suitability of the design and operating effectiveness of the controls stated in the description to achieve the related control objectives.")
space()
p("Opinion: In our opinion, in all material respects, (1) the description fairly presents the System; (2) the controls were suitably designed; and (3) the controls operated effectively throughout the period to provide reasonable assurance that the control objectives were achieved.")
space()
p("Restriction on Use: This report is intended solely for the information and use of management of 外部委託先SIer-A, user entities of the System, and their auditors.")
pdf.ln(15)
pdf.set_x(pdf.l_margin)
pdf.cell(180, 7, 'XYZ CPA Firm', ln=1, align='R')
pdf.cell(180, 7, 'Tokyo, Japan / May 20, 2025', ln=1, align='R')

# Section II
pdf.add_page()
h("Section II. Management's Assertion")
space()
p("We, the management of 外部委託先SIer-A, have prepared the accompanying description of the SAP ERP Operations and Development Services System (the System) throughout the period April 1, 2024 to March 31, 2025. We confirm that:")
space()
p("a. The description fairly presents the System.")
p("b. The controls related to the control objectives stated in the description were suitably designed.")
p("c. The controls operated effectively throughout the period.")
p("d. Complementary user entity controls (CUECs) are described in Section V.")
pdf.ln(20)
pdf.set_x(pdf.l_margin)
pdf.cell(180, 7, '外部委託先SIer-A 代表取締役 [Chief Executive Officer]', ln=1, align='R')

# Section III
pdf.add_page()
h("Section III. Description of the Service Organization's System")
space()
h("III-1. Company Overview", 12)
p("外部委託先SIer-A is a system integration services provider specializing in SAP ERP implementation, operations, and enhancement services for manufacturing and service industries in Japan. Established in 1985, the company operates from headquarters in Tokyo with delivery centers in Osaka and Fukuoka. As of December 2024, the company employs approximately 1,500 consultants and engineers.")
space()
h("III-2. Scope of Services Provided to デモA株式会社", 12)
p("外部委託先SIer-A provides the following services to the user entity:")
p("(1) SAP S/4HANA application support (L2/L3 support)")
p("(2) Custom development and enhancements (ABAP programming)")
p("(3) Change management and deployment coordination")
p("(4) Testing support (UAT coordination)")
p("(5) Periodic SAP upgrade management")
space()
p("Note: Service scope does NOT include: infrastructure management, database administration, and backup operations (handled by 外部委託先B社 separately).")
space()
h("III-3. System Components", 12)
p("The System includes: SAP development environment (DEV client 100), SAP quality assurance environment (QAS client 200), Change request management tool (ServiceNow-equivalent), Source code repository (Git-based), Incident management system.")
space()
h("III-4. Control Environment", 12)
p("外部委託先SIer-A maintains the following control environment elements: ISO 27001:2022 certification (latest recertification 2024-09); Annual employee code of conduct training; Segregation of duties between development and production access; Background check for all consultants handling client systems.")

# Section IV
pdf.add_page()
h("Section IV. Control Objectives, Related Controls, and Test Results")
space()

def iv_table(title, objective, rows):
    pdf.set_x(pdf.l_margin)
    pdf.set_font('YGB', '', 12)
    pdf.cell(180, 8, title, ln=1)
    pdf.set_x(pdf.l_margin)
    pdf.set_font('YG', '', 10)
    pdf.multi_cell(180, 6, f'Control Objective: {objective}')
    pdf.ln(2)
    pdf.set_x(pdf.l_margin)
    pdf.set_fill_color(48, 84, 150)
    pdf.set_text_color(255, 255, 255)
    pdf.cell(70, 8, 'Control Activity', border=1, align='C', fill=True)
    pdf.cell(70, 8, 'Test Performed', border=1, align='C', fill=True)
    pdf.cell(40, 8, 'Result', border=1, align='C', fill=True)
    pdf.ln()
    pdf.set_text_color(0, 0, 0)
    for activity, test, result in rows:
        pdf.set_x(pdf.l_margin)
        pdf.cell(70, 8, activity, border=1)
        pdf.cell(70, 8, test, border=1)
        pdf.cell(40, 8, result, border=1)
        pdf.ln()
    pdf.ln(5)

iv_table("IV-1. Logical Access",
    "Controls provide reasonable assurance that logical access is restricted to authorized individuals.",
    [
        ('New user access requires approval', 'Inspected 25 new user requests', 'No exceptions'),
        ('Quarterly access reviews performed', 'Inspected 4 quarterly reviews', 'No exceptions'),
        ('Terminated users access removed', 'Tested 10 terminations', 'No exceptions'),
    ])

iv_table("IV-2. Change Management",
    "Controls provide reasonable assurance that changes are authorized, tested, and approved.",
    [
        ('All changes require written approval', 'Inspected 30 change requests', 'No exceptions'),
        ('Changes tested in QAS before PRD', 'Inspected 30 test records', 'No exceptions'),
    ])

pdf.set_x(pdf.l_margin)
pdf.set_font('YGB', '', 12)
pdf.cell(180, 8, "IV-3. Backup and Recovery (Reference only - performed by B社)", ln=1)
pdf.set_x(pdf.l_margin)
pdf.set_font('YG', '', 10)
pdf.multi_cell(180, 6, "This control area is primarily the responsibility of 外部委託先B社 (the infrastructure provider). 外部委託先SIer-A coordinates restore testing with B社.")
pdf.ln(5)

iv_table("IV-4. Incident Management",
    "Controls provide reasonable assurance that incidents are identified, tracked, and resolved timely.",
    [
        ('All incidents logged within 1 hour', 'Inspected 45 incidents', 'No exceptions'),
        ('Root cause analysis for severity 1', 'Inspected 5 severity-1 incidents', 'No exceptions'),
    ])

iv_table("IV-5. Physical Security",
    "Access to development/QAS facilities is restricted to authorized personnel.",
    [
        ('Badge access to development areas', 'Inspected access logs (3 months)', 'No exceptions'),
    ])

# Section V
pdf.add_page()
h("Section V. Complementary User Entity Controls (CUECs)")
space()
p("外部委託先SIer-A assumes that user entities (including デモA株式会社) will implement the following controls:")
space()
p("CUEC-1: User entity management will review and authorize all change requests before SIer-A implements them.")
space()
p("CUEC-2: User entity will perform User Acceptance Testing (UAT) for all changes.")
space()
p("CUEC-3: User entity will maintain its own access review process and notify SIer-A promptly of user access changes.")
space()
p("CUEC-4: User entity will monitor incident resolution and escalate as needed.")

pdf.output(str(pdf_path))
print(f"[Fixed] SOC1 SIer-A: regenerated with unqualified (clean) opinion, Section VI removed")
