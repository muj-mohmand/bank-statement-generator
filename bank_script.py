import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import LETTER
from reportlab.lib import colors
from PyPDF2 import PdfReader, PdfWriter

# Get page dimensions
PAGE_WIDTH, PAGE_HEIGHT = LETTER
print(f"Page dimensions: Width = {PAGE_WIDTH} points, Height = {PAGE_HEIGHT} points")
print(f"Page dimensions: Width = {PAGE_WIDTH/72:.2f} inches, Height = {PAGE_HEIGHT/72:.2f} inches")


# ----------------------------
# 1. Configuration
# ----------------------------
TEMPLATE_PDF = "../bank/td_statement_edited_V4.pdf"
OUTPUT_PDF = "../bank/generated_statements/td_chequing_statement.pdf"
OVERLAY_PDF = "overlay.pdf"
EXCEL_FILE = "../../BrightDesk_Consulting_Ledger_Mar2022_to_Aug2025_v4.xlsx"
FONT_NAME = "Helvetica"
FONT_SIZE = 8.5
FONT_SIZE_LARGE = 10



# Column positions on the PDF (adjust as needed)sstart_x_payee = 50
start_x_payee = 70
start_x_withdrawal = 305
start_x_deposit = 405
start_x_date = 415
start_x_balance = 535
start_y = 485
line_height = 2.8 * 2.83465  # Convert 3.44mm to points (exact height)
start_x_total_withdrawal = 305
start_x_total_deposit = 404
start_y_total_withdrawal_and_deposit = 793 - 602
start_x_number_of_transactions, start_y_number_of_transactions = 240, PAGE_HEIGHT - 642
start_x_transactions_fees, start_y_transactions_fees = 290, PAGE_HEIGHT - 642
start_x_waived_fees, start_y_waived_fees = 428, PAGE_HEIGHT - 642

# Transaction fees configuration
FEE_PER_TRANSACTION = 1.25
GREY_ROW_COORDS = {
    'x1': 66,
    'y1': 298,
    'x2': 547.5,
    'y2': 307
}

# ----------------------------
# 2. Load Excel Transactions
# ----------------------------
df = pd.read_excel(EXCEL_FILE, sheet_name='chequing_savings_credit')
print(df.head())

# Filter for Chequing Account only
df = df[df['Account'] == 'Chequing Account']
print("\nFiltered for Chequing Account only:")
print(df.head())

df.columns = df.columns.str.strip()  # remove leading/trailing spaces
df['Date'] = pd.to_datetime(df['Date'])
df = df.sort_values('Date')

# Get unique months from the data
df['YearMonth'] = df['Date'].dt.to_period('M')
unique_months = df['YearMonth'].unique()
print(f"\nUnique months found: {unique_months}")


# ----------------------------
# 3. Create Monthly Statements
# ----------------------------
for month in unique_months:
    # Filter data for current month
    month_data = df[df['YearMonth'] == month]
    
    if len(month_data) == 0:
        continue
    
    # Create month-specific filenames
    month_str = month.strftime('%Y_%m')
    month_name = month.strftime('%B %Y')
    overlay_pdf = f"overlay_{month_str}.pdf"
    output_pdf = f"../bank/generated_statements/td_chequing_statement_{month_str}.pdf"
    
    print(f"\nProcessing {month_name} with {len(month_data)} transactions...")
    
    # Create TD-style Overlay PDF for this month
    c = canvas.Canvas(overlay_pdf, pagesize=LETTER)
    c.setFont(FONT_NAME, FONT_SIZE)
    y = start_y
    row_counter = 0
    
    # Add starting balance row as first row
    if len(month_data) > 0:
        first_row = month_data.iloc[0]
        starting_balance = first_row.get('Beginning Balance', 0)
        
        # Reset text color to black for writing
        c.setFillColorRGB(0, 0, 0)
        
        # Draw starting balance data
        c.drawString(start_x_payee, y, "STARTING BALANCE")
        c.drawRightString(start_x_withdrawal, y, "")  # No withdrawal
        c.drawRightString(start_x_deposit, y, "")     # No deposit
        
        # Use the date from first transaction
        day = f"{first_row['Date'].day:02d}"
        month_abbr = first_row['Date'].strftime('%b').upper()
        date_str = f"{month_abbr}{day}"
        c.drawString(start_x_date, y, date_str)
        
        # Format starting balance with thousands separators
        balance = f"{starting_balance:,.2f}" if pd.notnull(starting_balance) else "0.00"
        c.drawRightString(start_x_balance, y, balance)
        
        # Move to next row
        y -= line_height
        
        # Start new page after 38 rows
        if row_counter >= 38:
            c.showPage()
            c.setFont(FONT_NAME, FONT_SIZE)
            y = start_y
            row_counter = 0  # Reset counter for new page to start with white
    
    #total of withdrawal and deposit
    total_withdrawal = month_data['Credit'].sum()
    total_deposit = month_data['Debit'].sum()
    c.drawRightString(start_x_withdrawal, start_y_total_withdrawal_and_deposit, f"{total_withdrawal:,.2f}")
    c.drawRightString(start_x_deposit, start_y_total_withdrawal_and_deposit, f"{total_deposit:,.2f}")
    
    # Calculate and display transaction fees
    number_of_transactions = len(month_data) - 1
    total_transaction_fees = number_of_transactions * FEE_PER_TRANSACTION
    transaction_fees_text = f"{number_of_transactions} X{FEE_PER_TRANSACTION}=${total_transaction_fees:.2f}"
    c.drawRightString(start_x_transactions_fees, start_y_transactions_fees, transaction_fees_text)
    c.drawRightString(start_x_waived_fees, start_y_waived_fees, f"${total_transaction_fees:.2f}")
    
    # Add statement date range at the top
    # Get the month from the first transaction and create full month boundaries
    month_start = month_data['Date'].min().replace(day=1)  # First day of the month
    month_end = (month_start + pd.DateOffset(months=1) - pd.DateOffset(days=1))  # Last day of the month
    
    # Format dates as MMM DD/YY
    first_date_str = month_start.strftime('%b %d/%y').upper()
    last_date_str = month_end.strftime('%b %d/%y').upper()
    
    # Create date range string
    date_range = f"{first_date_str} - {last_date_str}"
    
    # Draw the date range at x=532, y=250
    c.setFont(FONT_NAME, FONT_SIZE_LARGE)
    c.drawRightString(532, PAGE_HEIGHT - 250, date_range)
    c.setFont(FONT_NAME, FONT_SIZE)

    # Now process the actual transactions
    for _, row in month_data.iterrows():
        # Draw alternating background: row 0=grey, row 1=white, row 2=grey, etc.
        # Since starting balance was row 0, first transaction is row 1 (should be white)
        # So we want odd rows (1, 3, 5, etc.) to be white, even rows (0, 2, 4, etc.) to be grey
        if row_counter % 2 == 0:  # Even rows (0, 2, 4, etc.) get grey background

            c.setFillColorRGB(237/255, 237/255, 237/255)
            c.rect(
                GREY_ROW_COORDS['x1'],
                y - (GREY_ROW_COORDS['y2'] - GREY_ROW_COORDS['y1']),
                GREY_ROW_COORDS['x2'] - GREY_ROW_COORDS['x1'],
                GREY_ROW_COORDS['y2'] - GREY_ROW_COORDS['y1'],
                fill=1,
                stroke=0
            )
            
            # Draw vertical lines at specified x-coordinates
            rect_height = GREY_ROW_COORDS['y2'] - GREY_ROW_COORDS['y1']
            line_y_start = y - rect_height
            line_y_end = y
            
            # Set line color to black for the vertical lines
            c.setStrokeColorRGB(0, 0, 0)
            c.setLineWidth(0.5)  # Set line width
            
            # Draw vertical lines at x = 206, 311, 410, 452
            vertical_lines = [206, 311, 410, 452]
            for x_pos in vertical_lines:
                c.line(x_pos, line_y_start, x_pos, line_y_end)
            
            print(f"Row {row_counter}: Drawing grey background {str(row['Payee'])} at y={y}")
        
        # Reset text color to black for writing
        c.setFillColorRGB(0, 0, 0)
        
        # Draw transaction data in new order with adjusted y position
        c.drawString(start_x_payee, y, str(row['Payee']))

        # Format numbers with thousands separators and 2 decimal places
        withdrawal = f"{row['Credit']:,.2f}" if pd.notnull(row['Credit']) and row['Credit'] > 0 else ""
        deposit = f"{row['Debit']:,.2f}" if pd.notnull(row['Debit']) and row['Debit'] > 0 else ""

        c.drawRightString(start_x_withdrawal, y, withdrawal)
        c.drawRightString(start_x_deposit, y, deposit)

        day = f"{row['Date'].day:02d}"
        month_abbr = row['Date'].strftime('%b').upper()
        date_str = f"{month_abbr}{day}"
        c.drawString(start_x_date, y, date_str)

        balance = f"{row['Closing Balance']:,.2f}" if pd.notnull(row['Closing Balance']) else ""
        c.drawRightString(start_x_balance, y, balance)

        # Move to next row
        y -= line_height
        row_counter += 1
        
        # Start new page after 38 rows
        if row_counter >= 38:
            c.showPage()
            c.setFont(FONT_NAME, FONT_SIZE)
            y = start_y
            row_counter = 0  # Reset counter for new page to start with white
    
    c.save()
    print(f"TD-style overlay PDF generated: {overlay_pdf}")

    # ----------------------------
    # 4. Merge Overlay with Template PDF
    # ----------------------------
    template = PdfReader(TEMPLATE_PDF)
    overlay = PdfReader(overlay_pdf)
    writer = PdfWriter()

    # Merge each page
    for page_num in range(len(template.pages)):
        page = template.pages[page_num]
        
        # If overlay has multiple pages, cycle through them
        overlay_page_num = min(page_num, len(overlay.pages) - 1)
        page.merge_page(overlay.pages[overlay_page_num])
        
        writer.add_page(page)

    with open(output_pdf, "wb") as f:
        writer.write(f)

    print(f"TD-style statement generated for {month_name}: {output_pdf}")
