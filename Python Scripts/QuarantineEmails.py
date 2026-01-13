import win32com.client
from collections import defaultdict
import datetime
import pandas as pd
import os
import time
from openpyxl import load_workbook

# Connect to Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Loop through all accounts to find the target mailbox
target_email = "Fixed-CSC@bayobab.africa"
target_inbox = None
sender_email = 'quarantine@messaging.microsoft.com'

# file_path = os.path.join(os.getcwd(), "counts.xlsx")
file_path = r"C:\Users\monitoring.server1\MTN Group\CIM Reports Hub - Documents\Quarantine.xlsx"

# file_path ="Count.xlsx"
# sheet_name = "Page 1"

while True:
    for account in outlook.Folders:
        if account.Name.lower() == target_email.lower():  
            target_inbox = account.Folders("Inbox")
            break

    existing_subject_date_pairs = set()

    if os.path.exists(file_path):
        existing_data = pd.read_excel(file_path, sheet_name=None)
        for sheet in existing_data.values():
            if "Subject" in sheet.columns and "Date Info" in sheet.columns:
                for _, row in sheet.iterrows():
                    existing_subject_date_pairs.add((row["Subject"], row["Date Info"]))

    if target_inbox:
        filter_query = f"[SenderEmailAddress] = '{sender_email}'"
        all_emails = []
        def search_folders(folder):
            # Restrict to the filter query
            emails = folder.Items.Restrict(filter_query)
            all_emails.extend(emails)
            
            # Search subfolders recursively
            for subfolder in folder.Folders:
                search_folders(subfolder)

        search_folders(target_inbox)

        # Get count of filtered emails
    #     email_count = filtered_emails.Count
    #     print(f"Emails from Quarantine in {target_email}'s inbox: {email_count}")
    # else:
    #     print(f"Mailbox '{target_email}' not found in Outlook.")
        email_counts = defaultdict(int)

        email_data = []

            # Process emails
        for email in all_emails:
            print("Ïn for loop one")
            if email.ReceivedTime:  # Ensure the email has a received time
                received_date = email.ReceivedTime
                month_year = received_date.strftime("%B %Y")  # Format: "January 2025"
                email_counts[month_year] += 1

                body = email.Body
                subject_content = ""
                date_content = ""

                for line in body.splitlines():
                    if line.startswith("Subject:"):
                        subject_content = line.replace("Subject:", "").strip()
                    elif line.startswith("Date:"):
                        date_content = line.replace("Date:", "").strip()

                if (subject_content, date_content) not in existing_subject_date_pairs:
                    email_data.append({
                        "Month-Year": month_year,
                        "Received Date": received_date.strftime("%Y-%m-%d %H:%M:%S"),
                        "Subject": subject_content,
                        "Date Info": date_content,
                        "Week Num": None
                    })
                    existing_subject_date_pairs.add((subject_content, date_content))  # Mark as added


        # Print sorted results (latest month first)
        print(f"Emails from {sender_email} in {target_email}'s inbox, grouped by month-year:\n")
        for month_year, count in sorted(email_counts.items(), key=lambda x: datetime.datetime.strptime(x[0], "%B %Y"), reverse=True):
            print(f"{month_year}: {count} emails")

        data = {'Month-Year': list(email_counts.keys()), 'Email Count': list(email_counts.values())}

        if email_data:
            df = pd.DataFrame(email_data)

            if os.path.exists(file_path):
                with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                    existing_df = pd.read_excel(file_path, sheet_name="Sheet1")

                    # Append new rows at the bottom
                    updated_df = pd.concat([existing_df, df], ignore_index=True)

                    # Write back to "Sheet1" (without overwriting older data)
                    updated_df.to_excel(writer, sheet_name="Sheet1", index=False)

                    # Apply Week Number formula to new rows only
                    wb = writer.book
                    ws = writer.sheets["Sheet1"]

                    received_date_col = updated_df.columns.get_loc("Received Date") + 1
                    week_num_col = updated_df.columns.get_loc("Week Num") + 1
                    year_col = updated_df.columns.get_loc("Year") + 1
                    day_of_week_col = updated_df.columns.get_loc("Day of the Week") + 1

                    # Apply formula only for the new rows
                    start_row = len(existing_df) + 2  # New rows start after existing ones
                    end_row = len(updated_df) + 1  # Last row index

                    wb = writer.book
                    ws = writer.sheets["Sheet1"]


                    for row_num in range(start_row, end_row + 1):
                        ws.cell(row=row_num, column=week_num_col).value = f"=WEEKNUM(INDIRECT(ADDRESS({row_num},{received_date_col})), 1)"
                        ws.cell(row=row_num, column=year_col).value = f"=YEAR(INDIRECT(ADDRESS({row_num},{received_date_col})))"
                        ws.cell(row=row_num, column=day_of_week_col).value = f"=TEXT(INDIRECT(ADDRESS({row_num},{received_date_col})), \"ddd\")"

                    # Save and close the file so Excel can compute formulas
                    wb.save(file_path)
                    # wb.close()
                wb = load_workbook(file_path, data_only=True)  # 'data_only=True' ensures formulas return values
                ws = wb["Sheet1"]


                    # Convert formulas to values for new rows only
                for row_num in range(start_row, end_row + 1):
                    received_date = ws.cell(row=row_num, column=received_date_col).value
                    if received_date:
                        # ws.cell(row=row_num, column=week_num_col).value = received_date.isocalendar()[1]  # Week number
                        # ws.cell(row=row_num, column=year_col).value = received_date.year  # Year
                        # ws.cell(row=row_num, column=day_of_week_col).value = received_date.strftime("%a")  # Day of the week
                        received_date = datetime.datetime.strptime(received_date, "%Y-%m-%d %H:%M:%S")

                        # Replace formula results with actual values
                        ws.cell(row=row_num, column=week_num_col).value = received_date.isocalendar()[1]  # Week number
                        ws.cell(row=row_num, column=year_col).value = received_date.year  # Year
                        ws.cell(row=row_num, column=day_of_week_col).value = received_date.strftime("%a")

                # Save changes with values instead of formulas
                wb.save(file_path)
                wb.close()

                print(f"{len(df)} new email(s) processed and saved with values only in {file_path}")

            else:
                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name="Sheet1", index=False)
            
            print(f"{len(email_data)} new email(s) have been appended to the bottom of Sheet1 in {file_path}")
        else:
            print("No new unique emails found.")


    else:
        print(f"Mailbox '{target_email}' not found in Outlook.")

    print("Last Update has run at: ", datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    print("Sleeping for 4 hours before the next run...")
    time.sleep(14400)
    # time.sleep(21600)
    # time.sleep(100)