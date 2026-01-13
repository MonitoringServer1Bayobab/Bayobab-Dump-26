import win32com.client
from collections import defaultdict
import datetime
import pandas as pd
import os
import time
from openpyxl import load_workbook
import sys

file_path = r"C:\Users\monitoring.server1\MTN Group\CIM Reports Hub - Documents\Subsea Emails.xlsx"
target_email = "Fixed-CSC@bayobab.africa"

senders = [
    {"email": "TsarLO@telkom.co.za", "name": "Eassy"},
    {"email": "EassyNOC@telkom.co.za", "name": "Eassy"},
    {"email": "snoc@seacom.com", "name": "Seacom"},
    {"email": "noc@2afgera.net", "name": "2Africa"},
    {"email": "technic-cable@djibtel.dj", "name": "AAE-1"},
    {"email": "Ace_snoc1@orange-sonatel.com", "name": "ACE"},
    {"email": "bnoc.ace@easymail.orange.com", "name": "ACE"},
    {"email": "Remedy@mainOne.net", "name": "Mainone"},
    {"email": "l2technicalsupport@equinix.com", "name": "Mainone"},
    {"email": "l2technicalsupport@equinix.com", "name": "Mainone"},
    {"email": "EIG-NOAF@europeindiagateway.com", "name": "EIG"},
    {"email": "eig-noaf@airtel.com", "name": "EIG"},
    {"email": "eig-noaf@airtel.com", "name": "EIG"},
    {"email": "PNOC.Equiano@asn.com", "name": "Equiano"},
    {"email": "peace-pm@peacecable.com", "name": "Peace"},
    {"email": "ongak@moyanetworks.com", "name": "Peace"},
    {"email": "WSMC@stc.com.sa", "name": "SMW4"},
    {"email": "Inoc@center3.com", "name": "SMW4"},
    {"email": "servicesupport@center3.com", "name": "SMW4"},
    {"email": "csm_haramous@intnet.dj", "name": "SMW5"},
    {"email": "tmglobalsoc_international@tm.com.my", "name": "SMW5"},
    # {"email": "pnocwacs@infraco.co.za", "name": "WACS"},
    {"email": "na.wacs@tatacommunications.com", "name": "WACS"},
    {"email": "pnocwacs_infraco@tatacommunications.com", "name": "WACS"},
    {"email": "glo1.noc@gloworld.com", "name": "GLO"},
    {"email": "datasupport@gloworld.com", "name": "GLO"},
    {"email": "customercare@gloworld.com", "name": "GLO"},
    {"email": "customercare@eacables.com", "name": "IMEWE"},
    {"email": "carriersupport@eand.com", "name": "IMEWE"},
    {"email": "carrier@camtel.cm", "name": "NCSCS"},
]

# ====== CONNECT TO OUTLOOK ======
while True:
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Find target inbox
    target_inbox = None
    for account in outlook.Folders:
        if account.Name.lower() == target_email.lower():
            target_inbox = account.Folders("Inbox")
            break

    # ====== LOAD EXISTING EMAILS ======
    existing_subject_date_sender = set()
    print("existing_subject_date_sender", existing_subject_date_sender)

    if os.path.exists(file_path):
        print("Line 65 if block")

        existing_data = pd.read_excel(file_path, sheet_name=None)
        for sheet in existing_data.values():
            if all(col in sheet.columns for col in ["Subject", "Received Date", "Sender"]):
                for _, row in sheet.iterrows():
                    existing_subject_date_sender.add((
                        row["Subject"],
                        row["Received Date"],
                        row["Sender"],
                    ))
    # ====== PROCESS EACH SENDER ======
    email_data = []
    email_counts = defaultdict(int)


    if target_inbox:
        def search_folders(folder, sender_email, sender_name):
            print("Sender email", sender_email)
            filter_query = f"[SenderEmailAddress] = '{sender_email}'"
            emails = folder.Items.Restrict(filter_query)
            for email in emails:
                print("Email in for loop", email, "Received Time:", email.ReceivedTime)
                if not email.ReceivedTime:
                    continue

                received_date = email.ReceivedTime
                received_date_str = received_date.strftime("%Y-%m-%d %H:%M:%S")
                month_year = received_date.strftime("%B %Y")
                email_counts[month_year] += 1

                subject_content = email.Subject.strip()
                unique_key = (subject_content, received_date_str, sender_name)

                if unique_key not in existing_subject_date_sender:
                    email_data.append({
                        "Month-Year": month_year,
                        "Received Date": received_date_str,
                        "Subject": subject_content,
                        "Sender": sender_name,
                        "Email": sender_email,
                    })
                    existing_subject_date_sender.add(unique_key)

            for subfolder in folder.Folders:
                search_folders(subfolder, sender_email, sender_name)

        for sender in senders:
            # search_folders(target_inbox, sender["email"], sender["name"])
            search_folders(target_inbox, sender["email"], sender["name"])

        # ====== OUTPUT RESULTS ======
        if email_data:
            df = pd.DataFrame(email_data)

            if os.path.exists(file_path):
                with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                    existing_df = pd.read_excel(file_path, sheet_name="Sheet1")
                    updated_df = pd.concat([existing_df, df], ignore_index=True)
                    updated_df = updated_df.drop_duplicates(subset=["Subject", "Received Date", "Sender"])
                    updated_df.to_excel(writer, sheet_name="Sheet1", index=False)
                    wb = writer.book
                    wb.save(file_path)
            else:
                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name="Sheet1", index=False)

            print(f"{len(df)} new email(s) appended to {file_path}")
            print("Last Update has run at: ", datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        else:
            print("No new unique emails found.")
    else:
        print(f"Mailbox '{target_email}' not found.")

    # Optional delay (e.g., for scheduled execution)
    time.sleep(14400)

