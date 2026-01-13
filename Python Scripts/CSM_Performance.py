import os
import time
import logging
import win32com.client
import getpass
import shutil
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
from datetime import datetime

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

while True:
    try:
        username = getpass.getuser()

        # Define paths
        save_folder = f"C:\\Users\\{username}\\MTN Group\\CIM Reports Hub - Documents\\"
        download_folder = f"C:\\Users\\{username}\\Downloads\\"

        # Ensure Outlook is running
        logging.info("Checking if Outlook is running...")
        outlook_running = any("OUTLOOK.EXE" in proc for proc in os.popen('tasklist').read().splitlines())

        if not outlook_running:
            logging.info("Outlook is not running. Launching Outlook...")
            os.startfile("outlook")
            time.sleep(15)  # Wait for Outlook to fully open

        # Connect to Outlook
        logging.info("Connecting to Outlook inbox...")
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)

        # Find latest email with subject containing 'Cases_KE'
        attachment_file = None
        for message in messages:
            if "Cases_KE" in message.Subject:
                if message.Attachments.Count > 0:
                    attachment = message.Attachments.Item(1)
                    attachment_file = os.path.join(save_folder, attachment.FileName)
                    logging.info(f"Saving attachment to: {attachment_file}")
                    attachment.SaveAsFile(attachment_file)
                    shutil.copy(attachment_file, os.path.join(download_folder, attachment.FileName))
                break

        if not attachment_file or not os.path.exists(attachment_file):
            logging.error("No matching email with attachment found.")
            raise FileNotFoundError("Attachment not found.")

        # Load workbook
        logging.info(f"Loading workbook: {attachment_file}")
        wb = openpyxl.load_workbook(attachment_file)
        ws = wb.active

        # Find last row in column A
        last_row = ws.max_row
        logging.info(f"Last row in column A: {last_row}")

        # Helper to insert column
        def insert_column(col_idx, header):
            logging.info(f"Inserting column {get_column_letter(col_idx)} with header '{header}'")
            ws.insert_cols(col_idx)
            col_letter = get_column_letter(col_idx)
            ws[f"{col_letter}1"] = header
            for row in range(2, last_row + 1):
                ws[f"{col_letter}{row}"].number_format = "General"

        # Insert columns
        insert_column(8, "OpCo")  # Column H
        for row in range(2, last_row + 1):
            g_val = ws[f"G{row}"].value
            ws[f"H{row}"] = g_val.split("_")[0] if g_val and "_" in g_val else g_val

        insert_column(17, "MTTRespond")  # Column Q
        insert_column(18, "SLA Compliance")  # Column R
        insert_column(20, "MTTUpdate")  # Column T
        insert_column(21, "Aging(Days)")  # Column U
        insert_column(22, "Week")  # Column V
        # Insert MTTR(Mins) - Column AD
        insert_column(30, "MTTR(Mins)")
        for row in range(2, last_row + 1):
            ws[f"AD{row}"] = f'=IF(ISBLANK(AC{row}), "Not resolved", INT((AC{row}-O{row})*1440))'

        logging.info("Populating formulas and calculated values...")
        for row in range(2, last_row + 1):
            n = ws[f"N{row}"].value
        
            r = ws[f"S{row}"].value
            p = ws[f"P{row}"].value
            j = ws[f"J{row}"].value

            # MTTRespond
            # if r is None:
            if r is None:
                ws[f"Q{row}"] = "No response"
            elif n and r:
                ws[f"Q{row}"] = int((r - n).total_seconds() / 60)

            # SLA Compliance
            q_val = ws[f"Q{row}"].value
            if isinstance(q_val, int) and q_val > 15 or q_val == "No response":
                ws[f"R{row}"] = "Breached"
            else:
                ws[f"R{row}"] = "Not Breached"

            # MTTR
            if p is None:
                ws[f"T{row}"] = "Not Resolved"
            elif n and p:
                ws[f"T{row}"] = int((p - n).total_seconds() / 60)

            # Aging
            if n and j not in ["Closed", "Resolved"]:
                ws[f"U{row}"] = (datetime.now() - n).days

            # Week
            if n:
                ws[f"V{row}"] = f"Week{n.isocalendar()[1]}"

        # Replace values in column M
        logging.info("Replacing values in column M...")
        for row in range(2, last_row + 1):
            m_cell = ws[f"M{row}"]
            val = m_cell.value
            if val:
                replacements = {
                    "Servicedesk Kenya.Bayobab Africa": "Fixed",
                    ".NET Resolver L1.Bayobab Africa": "Fixed",
                    ".NET Resolver L2 IP.Bayobab Africa": "Fixed",
                    "Servicedesk Ghana.Bayobab Africa": "Mobility"
                }
                for old, new in replacements.items():
                    if old in val:
                        m_cell.value = val.replace(old, new)

        # Wrap text toggle
        logging.info("Applying wrap text formatting...")
        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = Alignment(wrap_text=False)

        # Save workbook
        output_file = os.path.join(save_folder, "CSM_Performance.xlsx")
        logging.info(f"Saving workbook as: {output_file}")
        wb.save(output_file)
        logging.info("Workbook saved successfully.")

        # Sleep for 1 hour and 50 minutes
        logging.info("Sleeping for 1 hour and 50 minutes before starting...")
        time.sleep(6600)

    
    except Exception as e:
            logging.error(f"An error occurred: {e}")
