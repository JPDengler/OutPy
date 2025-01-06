import win32com.client
import os
import openpyxl
from openpyxl.utils import get_column_letter

# Path to the Excel file
EXCEL_FILE = "Shift_Reports.xlsx"

# Filters
FILTER_KEYWORDS = ["Shift Report", "Maintenance shift report", "mech shift report"]

# Lines to remove
LINES_TO_REMOVE = [
    "*",
    "<",
    "Confidentiality Warning",
    "AUTOMATION TECHNICIAN",
    "Joseph Dengler",
]

# Keywords indicating irrelevant sections to remove
SECTION_KEYWORDS = ["Sent:", "Subject:", "From:", "To:", "Date:", "SHIFT REPORT"]

def clean_body(body):
    """Clean the email body by removing unwanted lines and irrelevant sections."""
    lines = body.splitlines()
    cleaned_lines = []
    for line in lines:
        stripped_line = line.strip()
        # Skip lines that contain unwanted keywords or start with unwanted patterns
        if not stripped_line or any(
            keyword in stripped_line for keyword in LINES_TO_REMOVE + SECTION_KEYWORDS
        ):
            continue
        cleaned_lines.append(stripped_line)
    return "\n".join(cleaned_lines)

def adjust_excel_formatting(sheet):
    """Set cell formatting to wrap text, adjust row heights, and sort by date."""
    # Auto-adjust cell formatting
    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
        for cell in row:
            cell.alignment = openpyxl.styles.Alignment(wrap_text=True)
        # Adjust the row height based on the longest text in the row
        max_length = max(len(str(cell.value) if cell.value else "") for cell in row)
        sheet.row_dimensions[row[0].row].height = min(15 + (max_length // 40) * 15, 300)
    
    # Sort the sheet by Received Time (Column C), descending order (newest first)
    rows = list(sheet.iter_rows(min_row=2, max_row=sheet.max_row, values_only=True))
    rows_sorted = sorted(rows, key=lambda x: x[2], reverse=True)  # Sort by date (column 3)
    
    # Clear existing rows and write the sorted rows
    for row_index, row in enumerate(rows_sorted, start=2):
        for col_index, value in enumerate(row, start=1):
            sheet.cell(row=row_index, column=col_index, value=value)

def scrape_outlook():
    try:
        # Connect to Outlook
        print("Connecting to Outlook...")
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

        # Access the root folder (adjust as needed)
        root_folder = outlook.Folders("Joseph.Dengler@bridor.com")
        inbox = root_folder.Folders("Inbox")

        print(f"Accessed folder: {inbox.Name}")
        messages = inbox.Items

        # Load or create the Excel workbook
        if os.path.exists(EXCEL_FILE):
            workbook = openpyxl.load_workbook(EXCEL_FILE)
            print(f"Loaded existing Excel file: {EXCEL_FILE}")
        else:
            workbook = openpyxl.Workbook()
            print(f"Created new Excel workbook: {EXCEL_FILE}")

        sheet = workbook.active
        if sheet.max_row == 1:  # Add headers if the sheet is empty
            sheet.append(["Subject", "Sender", "Received Time", "Body"])

        # Collect existing data to prevent duplicates
        existing_data = {
            (sheet.cell(row=i, column=1).value, sheet.cell(row=i, column=3).value)
            for i in range(2, sheet.max_row + 1)
        }

        # Process emails and populate the Excel file
        for message in messages:
            try:
                subject = message.Subject
                sender = message.SenderName
                received_time = message.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S")
                body = clean_body(message.Body)  # Clean the body of the email

                # Filter emails by keywords
                if not any(keyword.lower() in subject.lower() for keyword in FILTER_KEYWORDS):
                    continue

                # Skip duplicates
                if (subject, received_time) in existing_data:
                    print(f"Skipping duplicate email: {subject}")
                    continue

                # Append email data to the Excel sheet
                sheet.append([subject, sender, received_time, body])
                print(f"Added email: {subject}")

            except Exception as e:
                print(f"Error processing email: {e}")

        # Adjust Excel formatting (wrap text and sort by date)
        adjust_excel_formatting(sheet)

        # Save the updated Excel workbook
        workbook.save(EXCEL_FILE)
        print(f"Excel file updated: {EXCEL_FILE}")

    except Exception as e:
        print(f"Error accessing Outlook: {e}")

if __name__ == "__main__":
    scrape_outlook()
