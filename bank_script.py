import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import LETTER
from reportlab.lib import colors
from PyPDF2 import PdfReader, PdfWriter


# ----------------------------
# 1. Configuration
# ----------------------------
TEMPLATE_PDF = "BrightDesk/bank/td_statement.pdf"
OUTPUT_PDF = "/Users/muj.mohmand/projects/brightdesk/generated_statements/td_chequing_statement.pdf"
OVERLAY_PDF = "overlay.pdf"
EXCEL_FILE = "/Users/muj.mohmand/projects/brightdesk/BrightDesk/realistic_revenue_ledger.xlsx"
FONT_NAME = "Helvetica"
FONT_SIZE = 8.5


# Column positions on the PDF (adjust as needed)sstart_x_payee = 50
start_x_payee = 70
start_x_withdrawal = 305
start_x_deposit = 405
start_x_date = 415
start_x_balance = 535
start_y = 485
line_height = 10

# ----------------------------
# 2. Load Excel Transactions
# ----------------------------
df = pd.read_excel(EXCEL_FILE)
df.columns = df.columns.str.strip()  # remove leading/trailing spaces
df['Date'] = pd.to_datetime(df['Date'])
df = df.sort_values('Date')

# ----------------------------
# 3. Create TD-style Overlay PDF
# ----------------------------
c = canvas.Canvas(OVERLAY_PDF, pagesize=LETTER)
c.setFont(FONT_NAME, FONT_SIZE)
y = start_y

for idx, row in df.iterrows():
    # --- Draw alternating background row ---
    if idx % 2 == 0:  # even row → light grey
        c.setFillColorRGB(237/255, 237/255, 237/255)
        c.rect(40, y - (line_height - 2), 500, line_height, fill=1, stroke=0)
    # odd row = white background (skip drawing)

    # Reset text color to black for writing
    c.setFillColorRGB(0, 0, 0)

    # Draw transaction data in new order
    c.drawString(start_x_payee, y, str(row['Payee']))

    withdrawal = f"{row['Credit']:.2f}" if pd.notnull(row['Credit']) and row['Credit'] > 0 else ""
    deposit = f"{row['Debit']:.2f}" if pd.notnull(row['Debit']) and row['Debit'] > 0 else ""

    c.drawRightString(start_x_withdrawal, y, withdrawal)
    c.drawRightString(start_x_deposit, y, deposit)

    day = f"{row['Date'].day:02d}"              # zero-padded (01, 02, …, 31)
    month = row['Date'].strftime('%b').upper()  # e.g., MAR
    date_str = f"{month}{day}"                  # e.g., MAR01
    c.drawString(start_x_date, y, date_str)

    c.drawRightString(start_x_balance, y, f"{row['Closing Balance']:.2f}" if pd.notnull(row['Closing Balance']) else "")

    # Move down
    y -= line_height
    if y < 50:
        c.showPage()
        c.setFont(FONT_NAME, FONT_SIZE)
        y = start_y
c.save()
print(f"TD-style overlay PDF generated: {OVERLAY_PDF}")

# ----------------------------
# 4. Merge Overlay with Template PDF
# ----------------------------
template = PdfReader(TEMPLATE_PDF)
overlay = PdfReader(OVERLAY_PDF)
writer = PdfWriter()

# Merge each page
for page_num in range(len(template.pages)):
    page = template.pages[page_num]
    
    # If overlay has multiple pages, cycle through them
    overlay_page_num = min(page_num, len(overlay.pages) - 1)
    page.merge_page(overlay.pages[overlay_page_num])
    
    writer.add_page(page)

with open(OUTPUT_PDF, "wb") as f:
    writer.write(f)

print(f"TD-style statement generated: {OUTPUT_PDF}")
