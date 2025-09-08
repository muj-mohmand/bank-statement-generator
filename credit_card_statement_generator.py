#!/usr/bin/env python3
"""
Credit Card Statement Generator
Generates PDF statements from transaction data.
"""

import pandas as pd
import sys
import random
from datetime import datetime, timedelta
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import LETTER, letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle
from PyPDF2 import PdfReader, PdfWriter
from copy import deepcopy

# Configuration
INPUT_FILE = "../../BrightDesk_Consulting_Ledger_Mar2022_to_Aug2025_v7.xlsx"
OUTPUT_DIR = "credit_card_statements/"
TEMPLATE_PDF = "template/TD_GREEN_VISA_template_edited.pdf"

#company details
COMPANY_NAME = "BrightDesk Consulting"
STREET_ADDRESS = "22 WELLINGTON ST E"
CITY_PROVINCE_POSTAL_CODE = "TORONTO ON M3C 2Z4"
CREDIT_CARD_NUMBER_HIDDEN = "5213 03XX XXXX 1234"
CREDIT_CARD_NUMBER_VISIBLE = "5213 0300 0000 1234"


def load_data(file_path):
    """Load transaction data from Excel file."""
    return pd.read_excel(file_path, sheet_name='credit_card')

def create_transaction_table(data):
    """Create a transaction table for the statement."""
    col_widths = [95-47, 139-95, 309-139, 346-309]  # Column widths in points
    transaction_rows = []
    
    for _, row in data.iterrows():
        # Format dates
        date_str = row['Date'].strftime('%b %d')
        posting_date_str = row['Posting Date'].strftime('%b %d')
        
        # Format amount with proper sign and currency
        amount = row['Amount']
        if amount < 0:
            amount_str = f"-${abs(amount):,.2f}"
        else:
            amount_str = f"${amount:,.2f}"
        
        # Create table row
        table_row = [
            Paragraph(date_str, style=None),
            Paragraph(posting_date_str, style=None),
            Paragraph(str(row['Activity Description']), style=None),
            Paragraph(amount_str, style=None)
        ]
        transaction_rows.append(table_row)
    
    # Create the table
    table = Table(transaction_rows, colWidths=col_widths)
    
    # Apply styles
    table.setStyle(TableStyle([
        ("BOX", (0,0), (-1,-1), 0, colors.white),   # No full border
        ("GRID", (0,0), (-1,-1), 0, colors.white),  # No grid
        ("LINEABOVE", (0,0), (-1,0), 1, colors.black),  # Only top border (solid)
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("ALIGN", (0,0), (-1,-1), "LEFT"),
        ("FONTSIZE", (0,0), (-1,-1), 8),  # Set font size
        ("LEFTPADDING", (0,0), (-1,-1), 3),
        ("RIGHTPADDING", (0,0), (-1,-1), 3),
        ("TOPPADDING", (0,0), (-1,-1), 2),
        ("BOTTOMPADDING", (0,0), (-1,-1), 2),
    ]))
    
    return table

def generate_statement(data, statement_month, statement_year):
    """Generate PDF statement from transaction data with multi-page support."""
    # Create month-specific filenames
    month_str = f"{statement_year}_{statement_month:02d}"
    overlay_pdf = f"overlay_credit_card_{month_str}.pdf"
    output_pdf = f"{OUTPUT_DIR}credit_card_statement_{month_str}.pdf"
    
    # Pagination configuration
    PAGE_1_MAX_TRANSACTIONS = 14  # First page (template page 1)
    PAGE_3_MAX_TRANSACTIONS = 23  # Subsequent pages (template page 3)
    
    # Calculate how many pages we need
    total_transactions = len(data)
    if total_transactions <= PAGE_1_MAX_TRANSACTIONS:
        # All transactions fit on page 1
        overlay_pages_needed = 1
        page_transactions = [total_transactions]
    else:
        # Need page 1 + additional page 3 copies
        remaining_after_page1 = total_transactions - PAGE_1_MAX_TRANSACTIONS
        additional_pages = (remaining_after_page1 + PAGE_3_MAX_TRANSACTIONS - 1) // PAGE_3_MAX_TRANSACTIONS  # Ceiling division
        overlay_pages_needed = 1 + additional_pages
        
        # Calculate transactions per page
        page_transactions = [PAGE_1_MAX_TRANSACTIONS]
        remaining = remaining_after_page1
        for i in range(additional_pages):
            transactions_this_page = min(PAGE_3_MAX_TRANSACTIONS, remaining)
            page_transactions.append(transactions_this_page)
            remaining -= transactions_this_page
    
    print(f"Total transactions: {total_transactions}")
    print(f"Pages needed: {overlay_pages_needed}")
    print(f"Transactions per page: {page_transactions}")
    
    # Create overlay PDF with multiple pages
    c = canvas.Canvas(overlay_pdf, pagesize=letter)
    
    # Page positioning constants
    PAGE_1_START_Y = letter[1] - 221  # First page Y position
    PAGE_3_START_Y = letter[1] - 190  # Page 3 Y position (adjust as needed)
    LINE_HEIGHT = 22
    START_X = 47
    COL_WIDTHS = [95-47, 139-95, 309-139, 346-309]
    
    # Beginning Balance (only on first page)
    end_x_beginning_balance = 346
    start_y_beginning_balance = letter[1] - 201
    # c.setFont("Times-Bold", 10)
    # c.drawRightString(end_x_beginning_balance, start_y_beginning_balance, f"${data['Beginning Balance'].iloc[0]:,.2f}")
    
    # Process transactions page by page
    transaction_index = 0
    
    for page_num in range(overlay_pages_needed):
        if page_num > 0:
            c.showPage()  # Start new page for additional pages
        
        # Determine Y starting position based on page type
        if page_num == 0:
            y_position = PAGE_1_START_Y  # First page
            # Add Beginning Balance only on first page
            if page_num == 0:
                c.setFont("Helvetica-Bold", 10)
                c.drawRightString(end_x_beginning_balance, start_y_beginning_balance, f"${data['Beginning Balance'].iloc[0]:,.2f}")
                c.drawRightString(577, letter[1] - 438, f"${data['Beginning Balance'].iloc[0]:,.2f}")
                c.setFontSize(12)
                c.drawRightString(308, letter[1] - 597, f"${data['Closing Balance'].iloc[-1]:,.2f}")
                c.setFontSize(12)
                c.drawRightString(577, letter[1] - 522, f"${data['Closing Balance'].iloc[-1]:,.2f}")
                c.setFont("Times-Bold", 8)
                c.drawString(138, letter[1] - 135, f"{data.attrs['statement_date'].strftime('%B %d, %Y')}")
                c.setFont("Times-Roman", 8)
                c.drawString(140, letter[1] - 149, f"{data.attrs['previous_statement_date'].strftime('%B %d, %Y')}")
                c.drawString(131, letter[1] - 161, f"{data.attrs['statement_period_start'].strftime('%B %d, %Y')} - {data.attrs['statement_period_end'].strftime('%B %d, %Y')}")
        else:
            y_position = PAGE_3_START_Y  # Page 3 template
        
        # Get transactions for this page
        transactions_this_page = page_transactions[page_num]
        page_data = data.iloc[transaction_index:transaction_index + transactions_this_page]
        
        print(f"Page {page_num + 1}: Processing {len(page_data)} transactions")
        
        # Draw transactions for this page
        row_count = 0
        total_payments_credits = 0.0
        total_purchases_and_charges = 0.0
        for _, row in page_data.iterrows():
            row_count += 1
            
            # Format dates
            date_str = row['Date'].strftime('%b %d').upper()
            posting_date_str = row['Posting Date'].strftime('%b %d').upper()
            
            # Format amount with proper sign and currency
            amount = row['Amount']
            if amount < 0:
                amount_str = f"-${abs(amount):,.2f}"
                total_payments_credits += abs(amount)
            else:
                amount_str = f"${amount:,.2f}"
                total_purchases_and_charges += amount
            
            # Set font
            c.setFont("Times-Roman", 8)
            
            # Draw dashed line above the row
            if row_count > 1:
                c.setDash([1, 1])
                c.setStrokeColorRGB(139/255, 0, 0)  # Convert to 0-1 scale for RGB
                c.setLineWidth(0.3)
                c.line(START_X, y_position + 10, START_X + sum(COL_WIDTHS), y_position + 10)
            c.setDash([])  # Reset to solid line
            
            # Draw each column
            x_pos = START_X
            c.drawString(x_pos, y_position, date_str)
            x_pos += COL_WIDTHS[0]
            c.drawString(x_pos, y_position, posting_date_str)
            x_pos += COL_WIDTHS[1]
            c.drawString(x_pos, y_position, str(row['Activity Description']))
            x_pos += COL_WIDTHS[2]
            c.drawRightString(x_pos + COL_WIDTHS[3], y_position, amount_str)
            
            # Check if this is the last row on this page
            if row_count == len(page_data):
                # Draw solid bottom border for last row
                c.setDash([])  # Solid line
                c.setStrokeColorRGB(0, 0, 0)  # Black color
                if page_num == overlay_pages_needed - 1:  # Last page
                    c.setLineWidth(1.0)  # Bold line for final page
                else:
                    c.setLineWidth(1.0)  # Normal line for continuation pages
                c.line(START_X, y_position - 5, START_X + sum(COL_WIDTHS), y_position - 5)
                
                # Add "Continued" or final balance information
                if page_num == overlay_pages_needed - 1:  # Last page
                    # Add TOTAL NEW BALANCE
                    balance_y = y_position - 15  # 4 pts below last row
                    activity_x = START_X + COL_WIDTHS[0] + COL_WIDTHS[1]  # Activity description x-axis
                    amount_x = START_X + sum(COL_WIDTHS)  # Amount x-axis
                    
                    c.setFont("Times-Bold", 10)
                    c.drawString(activity_x, balance_y, "TOTAL NEW BALANCE")
                    c.drawRightString(amount_x, balance_y, f"${data['Closing Balance'].iloc[-1]:,.2f}")
                else:
                    # Add "Continued" text
                    continued_y = y_position - 15  # 4 pts below last row
                    continued_x = START_X + sum(COL_WIDTHS)  # Right aligned at end of row width
                    
                    c.setFont("Times-Roman", 8)
                    c.drawRightString(continued_x, continued_y, "Continued")
            
            # Move to next row
            y_position -= LINE_HEIGHT
        #update payements and credits on first page
        if page_num == 0:
            c.setFont("Helvetica-Bold", 10)
            c.drawRightString(577, letter[1] - 452, "${:,.2f}".format(total_payments_credits))
            c.setFont("Helvetica", 8)
            c.drawRightString(577, letter[1]-464, "${:,.2f}".format(total_purchases_and_charges))
            c.setFont("Helvetica-Bold", 8)
            c.drawRightString(577, letter[1]- 510, "${:,.2f}".format(total_purchases_and_charges))
        else:
            c.setFont("Times-Bold", 8)
            c.drawString(138, letter[1] - 135, f"{data.attrs['statement_date'].strftime('%B %d, %Y')}")
            c.setFont("Times-Roman", 8)
            c.drawString(140, letter[1] - 149, f"{data.attrs['previous_statement_date'].strftime('%B %d, %Y')}")


        # Update transaction index for next page
        transaction_index += transactions_this_page
    
    c.save()
    print(f"Overlay PDF created: {overlay_pdf}")
    
    # Merge overlay with template PDF
    try:
        template = PdfReader(TEMPLATE_PDF)
        overlay = PdfReader(overlay_pdf)
        writer = PdfWriter()

        # Determine final document structure
        template_page_1 = template.pages[0]  # First page of template
        template_page_2 = template.pages[1] if len(template.pages) > 1 else None  # Second page of template
        template_page_3 = template.pages[2] if len(template.pages) > 2 else None  # Third page of template
        
        # Add first page (always page 1 of template with first overlay)
        page_1 = template_page_1
        if len(overlay.pages) > 0:
            page_1.merge_page(overlay.pages[0])
        writer.add_page(page_1)
        
        # Add second page of template (unchanged)
        if template_page_2:
            writer.add_page(template_page_2)
        
        # Add remaining pages (page 3 template with additional overlays)
        for overlay_page_num in range(1, len(overlay.pages)):
            if template_page_3:
                # Create a copy of template page 3
                page_3_copy = deepcopy(template_page_3)
                # Merge with corresponding overlay page
                page_3_copy.merge_page(overlay.pages[overlay_page_num])
                writer.add_page(page_3_copy)

        # Create output directory if it doesn't exist
        import os
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        
        with open(output_pdf, "wb") as f:
            writer.write(f)

        print(f"Credit card statement generated: {output_pdf}")
        
    except Exception as e:
        print(f"Error merging PDFs: {e}")
        print(f"Make sure {TEMPLATE_PDF} exists in the current directory")

def generate_monthly_statements(data):
    """
    Generate credit card statements for all months that have transaction data.
    
    Args:
        data: Raw transaction data from Excel
    """
    # Convert Date column to datetime if not already
    data['Date'] = pd.to_datetime(data['Date'])
    
    # Get the range of months that have data
    min_date = data['Date'].min()
    max_date = data['Date'].max()
    
    print(f"Data range: {min_date.strftime('%B %d, %Y')} to {max_date.strftime('%B %d, %Y')}")
    
    # Determine the first statement month
    # If first transaction is March 3, 2022, first statement should be March 2022 (Feb 26 - Mar 25)
    first_statement_month = min_date.month
    first_statement_year = min_date.year
    
    # Determine the last statement month
    # If last transaction is August 15, 2025, last statement should be August 2025 (Jul 26 - Aug 25)
    last_statement_month = max_date.month
    last_statement_year = max_date.year
    
    print(f"Generating statements from {first_statement_month}/{first_statement_year} to {last_statement_month}/{last_statement_year}")
    
    # Generate statements for each month
    current_month = first_statement_month
    current_year = first_statement_year
    
    while (current_year < last_statement_year) or (current_year == last_statement_year and current_month <= last_statement_month):
        # Generate statement for this month
        statement_data = generate_transaction_data(data, current_month, current_year)
        
        if len(statement_data) > 0:
            print(f"\nGenerating statement for {datetime(current_year, current_month, 1).strftime('%B %Y')}")
            generate_statement(statement_data, current_month, current_year)
        else:
            print(f"\nNo transactions found for {datetime(current_year, current_month, 1).strftime('%B %Y')}")
        
        # Move to next month
        if current_month == 12:
            current_month = 1
            current_year += 1
        else:
            current_month += 1

def generate_transaction_data(data, statement_month, statement_year):
    """
    Process transaction data from Excel file and format for credit card statement.
    
    Args:
        data: DataFrame with columns: Date, Contact, Description, Reference, Payee, 
              Beginning Balance, Debit, Credit, Closing Balance, Account Code, 
              Account, Account Type, Related account
        statement_month: Month number (1-12) for the statement
        statement_year: Year for the statement
    
    Returns:
        DataFrame with processed transaction data for the statement period
    """
    # Clean column names (remove leading/trailing spaces)
    data.columns = data.columns.str.strip()
    
    # Convert Date column to datetime
    data['Date'] = pd.to_datetime(data['Date'])
    
    # Calculate statement period: 26th of previous month to 25th of current month
    # For example: statement_month=3, statement_year=2022 means Feb 26 - Mar 25, 2022
    
    # Calculate period start (26th of previous month)
    if statement_month == 1:
        period_start = datetime(statement_year - 1, 12, 26)
    else:
        period_start = datetime(statement_year, statement_month - 1, 26)
    
    # Calculate period end (25th of current month)
    period_end = datetime(statement_year, statement_month, 25)
    
    # Previous statement date (25th of previous month)
    if statement_month == 1:
        prev_statement_date = datetime(statement_year - 1, 12, 25)
    else:
        prev_statement_date = datetime(statement_year, statement_month - 1, 25)
    
    print(f"Statement Period: {period_start.strftime('%B %d, %Y')} to {period_end.strftime('%B %d, %Y')}")
    print(f"Previous Statement Date: {prev_statement_date.strftime('%B %d, %Y')}")
    print(f"Statement Date: {period_end.strftime('%B %d, %Y')}")
    
    # Filter transactions within the statement period
    data = data[(data['Date'] >= period_start) & (data['Date'] <= period_end)].copy()
    

    print(f"Found {len(data)} transactions in statement period")
    
    # Create posting date (0-3 days after transaction date, randomized)
    def generate_posting_date(transaction_date):
        days_delay = random.randint(0, 3)
        return transaction_date + timedelta(days=days_delay)
    
    data['Posting Date'] = data['Date'].apply(generate_posting_date)
    
    # Create activity description from Payee column
    data['Activity Description'] = data['Payee'].astype(str)
    
    # Determine amount from Debit or Credit column
    def get_transaction_amount(row):
        if pd.notnull(row['Debit']) and row['Debit'] > 0:
            # Debit amounts are made negative (refund/decrease to credit card balance)
            return -float(row['Debit'])
        elif pd.notnull(row['Credit']) and row['Credit'] > 0:
            # Credit amounts are positive (payment/increase to credit card balance)
            return float(row['Credit'])
        else:
            return 0.0
    
    data['Amount'] = data.apply(get_transaction_amount, axis=1)
    
    # Create reference number if not provided
    def generate_reference(row):
        if pd.notnull(row['Reference']) and str(row['Reference']).strip():
            return str(row['Reference']).strip()
        else:
            # Generate a reference number based on date and row index
            date_str = row['Date'].strftime('%Y%m%d')
            return f"REF{date_str}{random.randint(1000, 9999)}"
    
    data['Reference Number'] = data.apply(generate_reference, axis=1)

    
    # Select and rename columns for final output
    processed_data = data[['Date', 'Posting Date', 'Activity Description', 'Reference Number', 'Amount', 'Beginning Balance', 'Closing Balance']].copy()
    
    # Sort by date
    processed_data = processed_data.sort_values('Date')
    
    # Add statement period information to the returned data
    processed_data.attrs = {
        'statement_period_start': period_start,
        'statement_period_end': period_end,
        'previous_statement_date': prev_statement_date,
        'statement_date': period_end
    }
    
    return processed_data

def main():
    """Main function."""
    try:
        # Load raw data from Excel file
        raw_data = load_data(INPUT_FILE)
        print(f"Loaded {len(raw_data)} transactions from Excel file")
        
        # Convert Date column to datetime to analyze data range
        raw_data['Date'] = pd.to_datetime(raw_data['Date'])
        min_date = raw_data['Date'].min()
        max_date = raw_data['Date'].max()
        
        print(f"Transaction data range: {min_date.strftime('%B %d, %Y')} to {max_date.strftime('%B %d, %Y')}")
        
        # Example: Generate statement for March 2022 (Feb 26 - Mar 25, 2022)
        print("\n=== Example: Generating March 2022 Statement ===")
        march_2022_data = generate_transaction_data(raw_data, 3, 2022)
        print(f"Processed {len(march_2022_data)} transactions for March 2022")
        
        if len(march_2022_data) > 0:
            # Display sample of processed data
            print("\nSample of processed transaction data:")
            print(march_2022_data.head())
            
            # Generate statement
            generate_statement(march_2022_data, 3, 2022)
            print("March 2022 statement generated successfully!")
        else:
            print("No transactions found for March 2022 statement period")
        
        # Generate statements for all months in data
        print("\n=== Generating All Monthly Statements ===")
        generate_monthly_statements(raw_data)
        
    except Exception as e:
        print(f"Error: {e}")
        return 1
    return 0

if __name__ == "__main__":
    sys.exit(main())