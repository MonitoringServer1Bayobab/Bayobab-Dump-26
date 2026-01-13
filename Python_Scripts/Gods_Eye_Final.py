import pandas as pd
from openpyxl import load_workbook, Workbook
import mysql.connector
import getpass
import mysql.connector
from datetime import datetime
from difflib import SequenceMatcher
import re
import openpyxl
import pandas as pd
from datetime import datetime, timedelta
import signal
import sys
from datetime import datetime
from openpyxl import load_workbook
import psutil
import win32com.client
import getpass
import os
import traceback
import logging
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import win32com.client
import openpyxl
import os
import shutil
import time

#Get logged in username
username = getpass.getuser()
print("Current User Logged In: ", username)

# Establish a connection to the MySQL database
cnx = mysql.connector.connect(
host='127.0.0.1',
user='vicky',
password='ILoveBlack@2022',
database='CSCMailsBackup'
)

cursor = cnx.cursor()


# Load the Excel file into a pandas DataFrame
#file_path = "C:\\Users\\victor.mwangi\\Documents\\Projects Files\\SD-L1 Data Integrity.xlsx"   # Replace 'your_excel_file.xlsx' with the path to your Excel file
 # Construct the file path
file_path = f"C:\\Users\\{username}\\Data_Integrity_Master.xlsx"

L1_data = []
SD_Data = []

#***************************************************************This Block Contains Code that verifies the integrity of Our Datacource and Does the necessarry ***********************************

def recover_From_SharePoint():
    source_file =f"C:\\Users\\{username}\\MTN Group\\CIM Reports Hub - Documents\\Data_Integrity_Master.xlsx"
    downloader_file =f"C:\\Users\\{username}\\MTN Group\\CIM Reports Hub - Documents\\Automation_Scripts\\Data_Integrity_SP_Recovery.xlsm"
 

    if not is_excel_file_corrupt(source_file):
        print("Sharepoint File Is Ok --- Recovering From It")
            # Open the Excel file with the default application
        print("Waiting for File to Download ")
        os.startfile(downloader_file)
        time.sleep(5)
        print("File Downloaded Already \n")
        
    else:
        print("Sharepoint File Is Corrupt --- creating a new Data Integrity File ")
        create_excel_file()
      

def upload_To_SharePoint():
    source_file = f"C:\\Users\\{username}\\Data_Integrity_Master.xlsx"
    destination_directory = f"C:\\Users\\{username}\\MTN Group\\CIM Reports Hub - Documents\\Data_Integrity_Master.xlsx"

    if not is_excel_file_corrupt(source_file):
        try:
            # Copy the file to the destination directory
            shutil.copy(source_file, destination_directory)
            print(f"File copied successfully to: {destination_directory}")
        except Exception as e:
            print(f"An error occurred: {e}")
    else:
        print("The File Has Corrupted")
        recover_Data_Source()

def create_Master_Copy():
# Construct the file path
    file_path = f"C:\\Users\\{username}\\Data_Integrity_Master.xlsx"
    if not is_excel_file_corrupt(file_path) :
        
        # Split the file path into directory and filename
        directory, filename = os.path.split(file_path)
        # Split the filename into name and extension
        name, extension = os.path.splitext(filename)
        # Construct the new filename with "Copy" suffix
        new_filename = f"{name}_Copy{extension}"
        # Construct the new file path
        new_file_path = os.path.join(directory, new_filename)
        try:
            # Copy the file to the new file path
            shutil.copy(file_path, new_file_path)
            print(f"File copied successfully to: {new_file_path}")
        except Exception as e:
            print(f"An error occurred: {e}")

def recover_Data_Source():
 # Define file paths
    file_path = f"C:\\Users\\{username}\\Data_Integrity_Master.xlsx"
    backup_file_path = f"C:\\Users\\{username}\\Data_Integrity_Master_Copy.xlsx"

    # Check if the backup file exists
    if os.path.exists(backup_file_path):
        # Construct the original file path by removing "_Copy" from the backup file path
        original_file_path = backup_file_path.replace("_Copy", "")

        try:
            # Remove the original file (if it exists)
            if os.path.exists(original_file_path):
                os.remove(original_file_path)
                print(f"Original file deleted: {original_file_path}")

            # Rename the backup file to remove "_Copy"
            os.rename(backup_file_path, original_file_path)
            print(f"Copied file renamed to: {original_file_path}")
        except Exception as e:
            print(f"An error occurred: {e}")
    else:
        print("Backup file does not exist.")

def verify_Data_File_Integrity():
    # Construct file paths
    file_path = f"C:\\Users\\{username}\\Data_Integrity_Master.xlsx"
    backup_file_path = f"C:\\Users\\{username}\\Data_Integrity_Master_Copy.xlsx"
   
    # Check if the file exists
    if not os.path.exists(file_path):
        print(f"The file '{file_path}' does not exist. Attempting to recover from SharePoint.")
     
        recover_From_SharePoint()
        return  # Exit the function after attempting recovery
       
    # Proceed with integrity checks if the file exists
    if is_excel_file_corrupt(file_path):
        print(f"The file '{file_path}' is corrupt.")
       
        # Verify if the local backup file is not corrupt
        if is_excel_file_corrupt(backup_file_path):
            print("Backup is Corrupt: Attempting recovery from SharePoint")
        else:
            print("Backup File is Ok, Attempting To Recover")
            recover_Data_Source()
    else:
        # If the main data file is not corrupt
        print(f"The file '{file_path}' is not corrupt.")
        print(f"Proceeding to create a copy of '{file_path}' \n")
        create_Master_Copy()

def is_excel_file_corrupt(file_path):
    try:
        # Create a new instance of Excel
        xl = win32com.client.Dispatch("Excel.Application")
        # Attempt to open the workbook
        xl.Workbooks.Open(file_path)
        # If opening succeeds, assume the file is not corrupt
        return False
    except Exception as e:
        # If an exception occurs, consider the file corrupt
        print(f"An error occurred: {e}")
        return True
    finally:
        # Close Excel application
        xl.Quit()

# Example usage:

def create_excel_file():
   
    # Construct the file path
    file_path = f"C:\\Users\\{username}\\Data_Integrity_Master.xlsx"

    # Create a new workbook
    workbook = openpyxl.Workbook()
    
    # Get the default sheet and rename it to "Sheet1"
    sheet1 = workbook.active
    sheet1.title = "Sheet1"

    # Define column headers for Sheet1
    sheet1_headers = [
        "No", "Incident No", "Opened By", "NCE Fault Occur Time", "SNOW Fault Occur Time",
        "Ticket Create Time", "Email Send Time", "Time Difference", "Occur Time Integrity",
        "NCE MTTN", "SNOW MTTN", "SNOW MTTC", "NCE MTTC", "MTTN SLA Compliance", 
        "Notification Status", "Time Of Day", "Type Of Day", "Assigned To", 
        "Resolved By", "Resolver Assignment Group", "Severity", "NCE Fault Clear Time", 
        "SNOW Fault Clear Time", "Repair Time Difference", "Repair Time Integrity", 
        "Overall Integrity", "Resolution Status", "MTTC_Difference", "MTTN_Difference", 
        "Responsible Party", "Category", "Type of Day", "Week", "State", "Fault"
    ]

    # Write column headers to Sheet1
    sheet1.append(sheet1_headers)

    # Create a new sheet named "Sheet2"
    sheet2 = workbook.create_sheet(title="Sheet2")

    # Define column headers for Sheet2
    sheet2_headers = [
        "No", "Incident No", "Fault Notification Time", "Fault Occur Time", 
        "Ticket Create Time", "First Reply Time", "SNOW MTTC", "OWA MTTC", 
        "MTTA", "Customer", "Category", "Data Integrity", "Time Difference", 
        "Severity", "Opened By", "Resolved By", "Notification Status", 
        "Assigned To", "Resolver Assignment Group", "Time of Day", 
        "Type of Day", "MTTA Compliance", "Day", "Week", "State", "Fault"
    ]

    # Write column headers to Sheet2
    sheet2.append(sheet2_headers)

    # Save the workbook
    workbook.save(file_path)
    print(f"Excel file created successfully at: {file_path}")


#***************************************************************This Block Contains Code that verifies the integrity of Our Datacource and Does the necessarry ***********************************



###########################################################################################################################################################################

customerEmailsDB = [

['Afrix','afr-ix.com','afr-ix','WHS',''],
['Amazon','amazon','amazon','WHS',''],
['Arelion','arelion.com','arelion','WHS',''],
['Avelacom','avelacom.com','avelacom','WHS',''],
['Bringcom','bringcom.com','bringcom','WHS',''],
['China Unicom','chinaunicom.cn','chinaunicom','WHS',''],
['CMC','cmcnetworks.net','cmc','WHS',''],
['CMI','cmi.chinamobile.com','cmi','WHS',''],
['Cogent','cogentco.com','cogent','WHS',''],
['CTG','chinatelecomglobal.com','CTHK NOC','WHS',''],
['Dolphin','dolphintelecom.net','dolphin','WHS',''],
['Expresso','expressotelecom.com','expresso','WHS',''],
['GABS','gabsgroup.com','gabs','WHS',''],
['GTT','gtt.net','gtt','WHS',''],
['GVA','gva.africa','canalbox.net','WHS',''],
['Isocell','isoceltelecom.com ','isocel','WHS',''],
['Livecom','livecom.hk','livecom','WHS',''],
['Meta','meta.com','meta','WHS',''],
['Moya','moyanetworks.com','moyanetworks','WHS',''],
['NTT','nttglobal.net','ntt','WHS',''],
['Orange','orange.com','orange','WHS',''],
['Paratus','paratus.africa','.ao','WHS',''],
['PCCW','pccwglobal.com','pccw','WHS',''],
['Space X','spacex.com','spacex','WHS',''],
['Sparkle','tisparkle.com','sparkle','WHS',''],
['Telefonica','telefonica.com','telefonica','WHS',''],
['Tata','tatacommunications.com','tata','WHS',''],
['Togo Tel','togocom.tg','togo','WHS',''],
['Togocom','togocom.tg','togo','WHS',''],
['Vodacom','vodacom','vcontractor.co.za','WHS',''],
['Wacren','wacren.net','wacren','WHS',''],
['Zayo','zayo.com','zayo','WHS',''],
['STC','stc.com.sa','stc','WHS',''],
['Simbanet','simbanet.co.tz','simbanet','WHS','']
]

actual_Category = {
        "Musa Otieno":"GC_Proactive",
        "Livingstone Lugaka":"GC_Proactive",
        "Gilbert Sang":"GC_Proactive",
        "Faith Nashipai":"GC_Proactive",
        "Emmy Jepchirchir":"GC_Proactive",
        "Kenneth Mwenda":"GC_Proactive",
        "Levin Mbaru":"GC_Proactive",
        "Lydia Kosgei":"GC_Proactive",
        "Timothy Mavuta":"GC_Proactive",
        "Sammy Chepsol":"GC_Proactive",
        "Lorine Aguti":"GC_Proactive",
        "Wesley Tonui":"GC_Reactive",
        "Mary Kataka":"GC_Reactive",
        "Josphat Kimani":"GC_Reactive",
        "Cyrus Soi":"GC_Reactive",
        "Felix Nandwa":"GC_Reactive",
        "Cecilia  Wanjiru":"GC_Reactive",
        "Evanson Mugo":"GC_Reactive",
        "Filbert Ogallo":"GC_Reactive",
        "Vikram  Sakhwal":"GC_Reactive",
        "Daniel Odhiambo":"GC_Proactive",
        "Lilian Kahoi":"GC_Reactive",
        "Michael Sabala":"GC_Reactive",
        "Daniel Alubbe":"GC_Reactive",
        "Victor Mwangi":"GC_Reactive",

            }

def close_excel_file(file_path):
    try:
        xl = win32com.client.Dispatch("Excel.Application")
        wb = xl.Workbooks.Open(file_path)
        
        # Undo all filters
        for sheet in wb.Sheets:
            sheet.AutoFilterMode = False
        
        # Unhide all rows and columns
        for sheet in wb.Sheets:
            sheet.Rows.Hidden = False
            sheet.Columns.Hidden = False
        
        # Close the workbook without saving changes
        wb.Close(SaveChanges=False)
        
        return "File closed successfully."
    except Exception as e:
        return f"Error: {e}"


def on_interrupt(signal, frame):
    # Handle the interrupt here
    print("Interrupt signal received:", signal)
    print("Closing Connection")

    # Close the cursor and database connection
    append_data_to_excel_v2(file_path, "Sheet2", L1_data)
    append_data_to_excel_v2(file_path, "Sheet1", SD_Data)

    upload_To_SharePoint()
    #cursor.close()
    #cnx.close()
    sys.exit(0)


def updateDBRecords():
        
    # Define the table name
    table_name = 'snow_data'

    # Truncate the table to remove all existing records
    truncate_query = f"TRUNCATE TABLE {table_name}"
    cursor.execute(truncate_query)

    # Read the Excel file into a DataFrame
    excel_file = pd.read_excel(f'C:\\Users\\{username}\\MTN Group\\CIM Reports Hub - Documents\\SD_L1_Performance_MasterFile.xlsx')  # Replace 'your_excel_file.xlsx' with your actual Excel file

    # Define the column names for the MySQL table and enclose them in backticks
    columns = ', '.join([f"`{col}`" for col in excel_file.columns])

    # Iterate through the DataFrame, replace NaN values with 'NULL', and insert records into the table
    for index, row in excel_file.iterrows():
        # Replace NaN values with 'NULL'
        row = row.apply(lambda x: 'NULL' if pd.isna(x) else x)

        # Print the current row before inserting
        print("Inserting row:", index)

        # Generate the SQL INSERT statement with explicit column names
        sql_insert = f"INSERT INTO {table_name} ({columns}) VALUES ({', '.join(['%s']*len(row))})"
        
        # Execute the SQL statement with the values from the current row
        cursor.execute(sql_insert, tuple(row))

    # Commit the changes and close the database connection
    cnx.commit()

def append_data_to_excel(file_path, sheet_name, data):
    print(file_path)
    print(sheet_name)

    # Load the workbook
    workbook = openpyxl.load_workbook(file_path)

    # Select the sheet
    sheet = workbook[sheet_name]

    # Find the last used row
    last_row = sheet.max_row

    # Append the data to the next row
    next_row = last_row + 1
    for i, value in enumerate(data, start=1):
        sheet.cell(row=next_row, column=i).value = value

    # Save the workbook
    workbook.save(file_path)
    print("Working Well")

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl import Workbook

def append_data_to_excel_v2(file_path, sheet_name, data_records):
    number_of_records = len(data_records)

    """
    Append multiple data records to an existing Excel file using openpyxl.

    Parameters:
        file_path (str): The path to the Excel file.
        sheet_name (str): The name of the sheet in the Excel file.
        data_records (list of iterables): List of data records to be appended.
    """
    try:
        # Load the original workbook
        workbook = load_workbook(file_path)

        # Select the sheet
        sheet = workbook[sheet_name]

        # Find the last used row
        last_row = sheet.max_row

        # Append the data to the next row
        for data in data_records:
            sheet.append(data)

        # Save the workbook (overwrite existing file)
        workbook.save(file_path)

        print(f"Inserted {number_of_records} records in {sheet_name} at {file_path}")
    except Exception as e:
        print(f"Error: {e}")


def exrtractAddresses(allEmails):
    
    withoutSDMails = []

    withoutSDMails = allEmails[5].split(';') #Put all emails in a list    


    # Strip any leading or trailing whitespace from each email address
    withoutSDMails = [email.strip() for email in withoutSDMails]
    
    # Email addresses to remove
    emails_to_remove = ["SD-F@bayobab.africa", "Fixed-CSC@bayobab.africa" , "L1-F@bayobab.africa"]

    # Create a new list without the emails_to_remove
    withoutSDMails = [email for email in withoutSDMails if email not in emails_to_remove]


    return withoutSDMails

def getFirstResponse(ticketNo):
    # Define your SQL query
    sql_query = f"SELECT * FROM Mails WHERE IncidentNo = '{ticketNo}' and MailCategory = 'Sent' order by MailDate asc limit 1"
    #print(sql_query)
    # Execute the SQL query
    cursor.execute(sql_query)
    # Fetch All row returned by the query
    allEmails = cursor.fetchone()
    return allEmails

def SNOW_Data(ticketNo):
    # Define your SQL query

    sql_query = f"SELECT * FROM snow_data WHERE Number = '{ticketNo}'"
    #print(sql_query)
    # Execute the SQL query
    cursor.execute(sql_query)
    # Fetch All row returned by the query
    ticketData = cursor.fetchall()
    return ticketData


def filter_emails(email_list):
    blocked_domains = ['@mtn', '@bayobab.africa']
    filtered_emails = [email for email in email_list if not any(email.endswith(domain) for domain in blocked_domains)]
    return filtered_emails

def process_email_list(email_list):
    flag = 1  # Assume initially that all emails contain at least one of the keywords

    for email in email_list:
        if "mtn" not in email.lower() and "bayobab" not in email.lower():
            flag = 0  # Set flag to 0 if an email doesn't contain either keyword
            break  # No need to continue checking if one email fails the condition

    if flag == 1:
        return email_list  # Return the original email list if the flag is set to 1
    elif flag == 0:
        # Remove emails containing "huawei" or "hpartners" and return the filtered list
        filtered_emails = [email for email in email_list if "mtn" not in email.lower() and "bayobab" not in email.lower()]
        return filtered_emails
    else:
        raise ValueError("Invalid flag value. Flag should be either 0 or 1.")

def getCustomerDomain(email_address):
       
        domainList = []   
     # Loop through each email address in the list
        for email in email_address:
            
            #print(email)   
            #print(emailAdresses)

            # Split the email address at the "@" symbol
            parts = email.split("@")
            #print(parts)
            
            # Check if there are two parts (username and domain)
            if len(parts) == 2:
                # Add the domain to the extracted_domain_list
                domain = parts[1].lower()
                domainList.append(domain)

            # Remove duplicates by converting the list to a set and back to a list
            unique_domain_list = list(set(domainList))
            return unique_domain_list

def getCustomerName(lists, keyword):
    #print(len(lists))
    for sublist in lists:        
        domain = sublist[1]
        alias = sublist[2]

        if domain in keyword or alias in keyword:
            return sublist[0]

def getOPCOname(ticket_Email_Data):

    withoutSDMails = []

    withoutSDMails = ticket_Email_Data[6].split(';') #Put all emails in a list    

    # Strip any leading or trailing whitespace from each email address
    withoutSDMails = [email.strip() for email in withoutSDMails]
    
    # Email addresses to remove
    emails_to_remove = ["Bayobab FixedConnectivity Service Desk [ Bayobab ]", "Bayobab FixedConnectivity Customer Success Center [ Bayobab ]" , "Bayobab FixedConnectivity L1 Escalation [ Bayobab ]"]

    # Create a new list without the emails_to_remove
    withoutSDMails = [email for email in withoutSDMails if email not in emails_to_remove]

    country_list = []

    for email in withoutSDMails:
        
        if "gict" not in email.lower():
            # Extract text inside square brackets            
            start_index = email.find('[')
            end_index = email.find(']')
            if start_index != -1 and end_index != -1:
                country = email[start_index + 1:end_index].strip()
                country_list.append(country)
        else:
            return "GICT"

    #For the sake of getting Customer name accurately, this is an experimental feature to remove anything [Bayobab] from thenew list.
    # this will only maintain MTN xx and remove anything bayobab only
    for item in country_list:
        if item.lower() == "bayobab":
            country_list.remove('Bayobab')


    country_counts = {}
    for country in country_list:
        country_counts[country] = country_counts.get(country, 0) + 1

    if not country_counts:
        return "3rd Party Proactive"  # Return None if no countries are found in the list

    max_country = max(country_counts, key=country_counts.get)
    return max_country   

def extract_subject(input_string):
    # Use regular expression to match 'Re:', everything after the last '||', and any trailing spaces
    result = re.sub(r'Re:|(\s*\|\|[^|]*)$', '', input_string).strip()

    return result

def plan_B(customerDomain, ticketNo, first_reply_time, first_reply_Subject):
    #print("Plan B")

    # Define the SQL query with placeholders
    sql_query = "SELECT * FROM Mails WHERE NOT (MailSubject LIKE %s AND EmailAddress NOT LIKE %s) " \
                "AND MailCategory = 'Inbox' AND IncidentNo NOT LIKE %s " \
                "AND EmailAddress LIKE %s AND EmailAddress LIKE %s AND MailDate < %s " \
                "ORDER BY MailDate DESC LIMIT 1"

    # Define the values to be substituted into the query
    values = ('%INC%', '%gict%', '%' + ticketNo + '%', '%' + customerDomain + '%', '%' + customerDomain + '%', first_reply_time)

    # Execute the SQL query with parameterized values
    cursor.execute(sql_query, values)

    # Fetch the first row returned by the query
    notification_Email = cursor.fetchone()

    if notification_Email is not None:
        notification_email_subject = str(notification_Email[1])

        similarity_score = getSimilarityScore(first_reply_Subject, notification_email_subject)        

        if similarity_score >= 5:            
             
             return notification_Email
           
        else:
            return None
    else:
        return None


def plan_A(first_reply_Subject, ticketNo, first_reply_time):
    customer_notification = extract_subject(first_reply_Subject)
    #print(customer_notification)
    
    sql_query = """
        SELECT *
        FROM Mails
        WHERE MailCategory = 'Inbox'
            AND MailSubject NOT LIKE %s
            AND MailDate < %s
            AND MailSubject LIKE %s
        ORDER BY MailDate ASC
    """

    cursor.execute(sql_query, ('%' + ticketNo + '%', first_reply_time, '%' + customer_notification + '%'))
    notification_Email = cursor.fetchall()

    #print(notification_Email)

    if notification_Email is not None:
        if len(notification_Email) > 1:
            last_email = len(notification_Email) - 1
            notification_Email = notification_Email[last_email]
           

        elif len(notification_Email) != 0:
            notification_Email = notification_Email[0]       

        else:
            #print("Recheck Ticket")
            #print(len(notification_Email))
            return None

        notification_email_subject = str(notification_Email[1])        

        similarity_score = getSimilarityScore(first_reply_Subject, notification_email_subject)
        #print(similarity_score)

        if similarity_score >= 5:
            return notification_Email
        else:
            return None

    else:
        return None


def plan_C(customerDomain, ticketNo, first_reply_time, short_description):
    # Define the SQL query with placeholders
    sql_query = "SELECT * FROM Mails WHERE MailCategory = 'Inbox' " \
                "AND MailSubject NOT LIKE %s AND MailDate < %s " \
                "AND MailSubject LIKE %s AND EmailAddress LIKE %s " \
                "AND EmailAddress LIKE %s ORDER BY MailDate DESC LIMIT 1"

    # Define the values to be substituted into the query
    values = ('%' + ticketNo + '%', first_reply_time, '%' + short_description + '%', '%' + customerDomain + '%', '%' + customerDomain + '%')

    # Execute the query with parameterized values
    cursor.execute(sql_query, values)

    # Fetch the first row returned by the query
    notification_Email = cursor.fetchone()
    
    return notification_Email
 
def getNotificationEmail(customerDomain , ticketNo , first_reply_time , first_reply_Subject , short_description):

    notification_email = plan_A(first_reply_Subject , ticketNo , first_reply_time)
    

    if notification_email is None:
        notification_email = plan_B(customerDomain , ticketNo , first_reply_time , first_reply_Subject)               
        
        if notification_email  is None:
            notification_email = plan_C(customerDomain , ticketNo , first_reply_time , short_description)
             
            if notification_email  is None:       
                return None
            
    return notification_email        


    
    #Execute Plan A first
      

def getTimeDifference(sendTime, responsereceiveTime):
    try:
        # Check if either timestamp is 'NULL'
        if sendTime == 'NULL' or responsereceiveTime == 'NULL':
            # Handle the case where time data is 'NULL'
            return None

        # Parse the timestamps into datetime objects
        dt1 = datetime.strptime(sendTime, '%Y-%m-%d %H:%M:%S')
        dt2 = datetime.strptime(responsereceiveTime, '%Y-%m-%d %H:%M:%S')

        # Calculate the time difference in minutes
        diff_minutes = round((dt2 - dt1).total_seconds() / 60)
        return diff_minutes
    except ValueError as e:
        # Handle the case where the time data is not in the expected format
        print(f"Error parsing datetime: {e}")
        return None


def getSimilarityScore(first_Response_Subject, notificatioN_email_subject):
    score = SequenceMatcher(None, first_Response_Subject, notificatioN_email_subject).ratio()
    score = int(score * 10)
    return score

def check_Data_Integrity(fault_notificaton_time, fault_occur_time):
    time_Difference = getTimeDifference(fault_occur_time, fault_notificaton_time)

    if time_Difference is not None:
        if -1 <= time_Difference <= 1:
            return ["Ok", time_Difference]
        else:
            return ["Fail", time_Difference]
    else:
        # Handle the case where time_Difference is None (e.g., 'NULL' values in the timestamps)
        return ["Error", None]  # You may choose to handle this case differently based on your requirements

def check_Repair_Time_Data_Integrity(time1, time2):
    time_Difference = getTimeDifference(time1, time2)

    if time_Difference is not None:
        if time_Difference >= -2 :
            return ["Ok", time_Difference]
        else:
            return ["Fail", time_Difference]
    else:
        # Handle the case where time_Difference is None (e.g., 'NULL' values in the timestamps)
        return ["Error", None]  # You may choose to handle this case differently based on your requirements

def extract_incident_numbers_and_latest_timestamp(excel_file_path , update_status):
    """
    Extracts incident numbers and the latest timestamp from two sheets in an Excel file.

    Parameters:
    excel_file_path (str): Path to the Excel file.

    Returns:
    tuple: A tuple containing a list of incident numbers and the latest timestamp.
    """
    try:
        incident_numbers = []
        # Read the first sheet
        df1 = pd.read_excel(excel_file_path)

        # Read the second sheet
        df2 = pd.read_excel(excel_file_path, sheet_name=1)  # Assuming the second sheet

        # Check if the number of rows in either DataFrame is 1
        if len(df1) <1 or len(df2) < 1:
            latest_timestamp = "2022-09-01 00:00:00"
            earliest_timestamp = "2022-09-01 00:00:00"
        else:
            # Get the latest timestamp from each DataFrame before filtering
            latest_timestamp_df1 = df1['Ticket Create Time'].max()
            latest_timestamp_df2 = df2['Ticket Create Time'].max()

            latest_timestamp_df1 = df1['Ticket Create Time'].min()
            latest_timestamp_df2 = df2['Ticket Create Time'].min()

            # Pick the latest of the two timestamps
            latest_timestamp = max(latest_timestamp_df1, latest_timestamp_df2)
            # Pick the latest of the two timestamps
            earliest_timestamp = min(latest_timestamp_df1, latest_timestamp_df2)

      
        if update_status == "Yes":
            # Filter rows where the value in column 'State' is not 'Closed' for both sheets
            df1 = df1[df1['State'] != 'Closed']
            df2 = df2[df2['State'] != 'Closed']

            # Extract the 'Incident No.' column from both filtered DataFrames and combine them into a single list
            incident_numbers = df1['Incident No'].tolist() + df2['Incident No'].tolist()

            # Open the Excel file for editing
            workbook = load_workbook(excel_file_path)

            # Get the sheets
            sheet1 = workbook['Sheet1']
            sheet2 = workbook['Sheet2']

            # Iterate over rows in each sheet and remove rows where 'State' is not 'Closed'
            rows_to_delete_sheet1 = []
            for row in sheet1.iter_rows(min_row=2, max_row=sheet1.max_row, min_col=1, max_col=sheet1.max_column):
                if row[25].value != 'Closed':
                    rows_to_delete_sheet1.append(row)
            for row in rows_to_delete_sheet1:
                sheet1.delete_rows(row[0].row)

            rows_to_delete_sheet2 = []
            for row in sheet2.iter_rows(min_row=2, max_row=sheet2.max_row, min_col=1, max_col=sheet2.max_column):
                if row[34].value != 'Closed':
                    rows_to_delete_sheet2.append(row)
            for row in rows_to_delete_sheet2:
                sheet2.delete_rows(row[0].row)

            # Save the changes to the Excel file
            workbook.save(excel_file_path)

        return incident_numbers, latest_timestamp , earliest_timestamp

    except Exception as e:
        print(f"Error: {e}")
        return None, None

def getTicketList(file_path, update_status):

    #first get all tickets that needs to be updated in the existing excel record. This are all tickets that arent closed

    incident_numbers, latest_timestamp , earliest_timestamp = extract_incident_numbers_and_latest_timestamp(file_path , update_status )
    print("Records To Be Updated: ", len(incident_numbers))
    print("Latest Timestamp:", latest_timestamp)
    print("Last record Timestamp:", earliest_timestamp)

    

    # now get from SNOW all tickets that have been added since latest timestamp    
    sql_query = f"SELECT Number FROM snow_data WHERE Created > '{latest_timestamp}' order by Created desc"
    sql_query2 = f"SELECT Number FROM snow_data WHERE Created < '{earliest_timestamp}' order by Created desc" #Get even other tickets , to ensure all records are updated
    
    # Execute the SQL query
    cursor.execute(sql_query)
    # Fetch All row returned by the query
    ticketData = cursor.fetchall()
     # Convert fetched data into a list of incident numbers
    new_Tickets = [row[0] for row in ticketData]

        # Execute the SQL query
    cursor.execute(sql_query2)
    # Fetch All row returned by the query
    ticketData = cursor.fetchall()
     # Convert fetched data into a list of incident numbers
    older_Tickets = [row[0] for row in ticketData]

    print("New added Tickets: " , len(new_Tickets))    
    print("Older Tickets to Process: " , len(older_Tickets)) 

    ticket_List = new_Tickets + incident_numbers + older_Tickets
    print("Total Tickets To Process: " , len(ticket_List))   

    return ticket_List


def get_MTTA_SLA(MTTA):

    try:
        if MTTA >= 0 and MTTA <=15:
            return "Within SLA"
        else:
            return "Breached"
    except TypeError as e:
        return "Breached"


def getFirstNotification(ticketNo):
    pass

def convert_To_EAT(timestamp_str):
    # Convert the input timestamp string to a datetime object
    dt_obj = datetime.strptime(timestamp_str, '%Y-%m-%d %H:%M:%S')

    # Add a second to the datetime object
    dt_obj_with_second = dt_obj + timedelta(hours=3)

    # Convert the updated datetime object back to a timestamp string
    updated_timestamp_str = dt_obj_with_second.strftime('%Y-%m-%d %H:%M:%S')

    return updated_timestamp_str

def getNCE_TicketData(ticketNo):
    # Define your SQL query

    sql_query = f"SELECT * FROM NCE_data WHERE Number = '{ticketNo}'"
    #print(sql_query)
    # Execute the SQL query
    cursor.execute(sql_query)
    # Fetch All row returned by the query
    ticketData = cursor.fetchone()

    if ticketData is not None:
        nce_fault_occur_time = str(ticketData[1])        
        nce_fault_occur_time = convert_To_EAT(nce_fault_occur_time)

        nce_fault_clear_time = str(ticketData[2])
        nce_fault_clear_time = convert_To_EAT(nce_fault_clear_time)

        return nce_fault_occur_time , nce_fault_clear_time
    else:
        return None

def getOverallIntegrity(repair_integrity_status, integrity_status):
    if repair_integrity_status == "Ok" and integrity_status == "Ok":
        return "Ok"
    else:
        return "Fail"

# Function to get the value associated with a key
def get_from_dictionary(key):  
    return actual_Category.get(key, "Key not found")

sampleTickets = [
'INC3900470',
'INC3900405',
]


#######################################################################################################################################################################3


def process_Ticket(ticketNo):
        result = []
        i = 0    
        k = 0
        y = 0
        try:
            k = k + 1  
            customer_category = ""
            fault = None                

            ticketData = SNOW_Data(ticketNo) #Get The ticket data in SNOW to determine the Category           
            #print(ticketData) 

            if ticketData is not None and not len(ticketData) == 0 : #if Ticket has data from SNOW 
                #print(ticketNo , "SNOW Ticket Data Found " )
                
                try:
                    
                    short_description = str(ticketData[0][39])
                    fault_occur_time = str(ticketData[0][7])
                    ticket_create_time = str(ticketData[0][9])
                    time_to_ticket_SNOW = int(float(ticketData[0][11]))
                    severity = str(ticketData[0][23])
                    opened_by = str(ticketData[0][48])
                    resolved_by = str(ticketData[0][54])
                    assigned_to = str(ticketData[0][49])
                    resolver_assignment_group = str(ticketData[0][36])
                    time_of_day = str(ticketData[0][34])
                    type_of_day = str(ticketData[0][63])
                    snow_fault_clear_time = ticketData[0][13]            
                    responsible_party = str(ticketData[0][2])
                    SNOW_mttc = str(ticketData[0][11])
                    SNOW_mttc = int(float(SNOW_mttc))
                    week = str(ticketData[0][28])
                    day = str(ticketData[0][26])
                    state = str(ticketData[0][22])                    
                    GC_Category = get_from_dictionary(opened_by)
                    #print(ticketNo , GC_Category)

                    #print(GC_Category , "\n")                 
                

                except Exception as e:
                    # Handle the ValueError (could not convert string to float)
                    print(f"Error converting string to float: {e}")
                    # You may choose to set a default value for time_to_ticket_SNOW or handle it in another way
                    short_description = str(ticketData[0][39])
                    fault_occur_time = str(ticketData[0][7])
                    ticket_create_time = str(ticketData[0][9])
                    time_to_ticket_SNOW = 0.0
                    severity = str(ticketData[0][23])
                    opened_by = str(ticketData[0][48])
                    resolved_by = str(ticketData[0][54])
                    assigned_to = str(ticketData[0][49])
                    resolver_assignment_group = str(ticketData[0][36])
                    time_of_day = str(ticketData[0][34])
                    type_of_day = str(ticketData[0][63])
                    snow_fault_clear_time = ticketData[0][13]            
                    responsible_party = str(ticketData[0][2])
                    SNOW_mttc = str(ticketData[0][11])
                    SNOW_mttc = 0.0
                    week = str(ticketData[0][28])
                    day = str(ticketData[0][26])
                    state = str(ticketData[0][22])   

                    GC_Category = get_from_dictionary(opened_by)

                    time_to_ticket_SNOW = 0.0  # Default value or any other appropriate value
                                   


                if GC_Category == "GC_Reactive": #If its a reactive ticket, then proceeed
                    #print(ticketNo , "Its a Reactive ")
                    
                    category = "SD"
                    SD_temp_list = []                
                    
                    i = i + 1  
                
                    ticket_Email_Data = getFirstResponse(ticketNo) #get the first sent email
                    #Extract the emails being responded to which are our customer
                    str_emails_address = ticket_Email_Data
                    
                    if str_emails_address is not None:   # This means that the ticket was raised via email
                        #print(ticketNo , "Raised in email")
                    

                        email_address = exrtractAddresses(str_emails_address) #strip Fixed csc email from the email list to remain with customer emails only
                    

                        filtered_email_list = process_email_list(email_address)                                     
                        
                        customerDomain = getCustomerDomain(filtered_email_list)
                        
                        customerDomain = customerDomain[0]

                        #print(customerDomain)
                        customerName = getCustomerName(customerEmailsDB , customerDomain)                    
                
                        
                        if customerName is None: # This Implies that the Customer is an OPCO
                            customerName = getOPCOname(ticket_Email_Data)

                        if customerName == "GICT":
                            customer_category = "GEBU"
                        elif "MTN" in customerName or "BAYOBAB" in customerName:
                            customer_category = "OPCO"
                        else:
                            customer_category = "WHS"

                        first_reply_time = str(ticket_Email_Data[2])
                        first_reply_Subject = str(ticket_Email_Data[1])        

                        #Get the Email That Customer Sent
                        fault_notification_Email = getNotificationEmail(customerDomain , ticketNo , first_reply_time , first_reply_Subject , short_description)                       
                        

                        if fault_notification_Email is not None: # Check if there was any email received before that, if there was, then proceed
                        #print(fault_notification_Email)

                            #from the notificationEmail, extract the fault received Time
                            fault_notificaton_time = str(fault_notification_Email[2])
                            fault = fault_notification_Email[1]
                            MTTA = getTimeDifference(fault_notificaton_time , first_reply_time)
                            MTTA_SLA = get_MTTA_SLA(MTTA)

                            MTTC_Outlook = getTimeDifference(fault_notificaton_time , ticket_create_time)
                            Data_Integrity = check_Data_Integrity(fault_notificaton_time , fault_occur_time)
                        

                            result = [[i, ticketNo , fault_notificaton_time , fault_occur_time , ticket_create_time , first_reply_time, time_to_ticket_SNOW , MTTC_Outlook, MTTA , customerName ,customer_category , Data_Integrity[0] , Data_Integrity[1] ,severity , opened_by , resolved_by , "Raised",assigned_to , resolver_assignment_group , time_of_day , type_of_day, MTTA_SLA , day, week, state, fault], category]
                            #print (k, ticketNo , state ," - SD ", fault_notificaton_time , fault_occur_time , ticket_create_time , first_reply_time, time_to_ticket_SNOW , MTTC_Outlook, MTTA , customerName ,customer_category , Data_Integrity[0] , Data_Integrity[1] ,severity , opened_by , resolved_by , "Raised",assigned_to , resolver_assignment_group , time_of_day , type_of_day, MTTA_SLA)
                            
                            #print(i, ticketNo , fault_notificaton_time , fault_occur_time , ticket_create_time , first_reply_time, time_to_ticket_SNOW , MTTC_Outlook, MTTA , customerName ,customer_category , Data_Integrity[0] , Data_Integrity[1] ,severity , opened_by , resolved_by , "Raised",assigned_to , resolver_assignment_group , time_of_day , type_of_day , MTTA_SLA)
                            #SD_Data.append(SD_temp_list)
                                        
                        else: #No Inbox Email was found , Maybe this is a proactive notification to customer or third party
                            #print(ticketNo , "No Inbox Email, THis was a proactive Ticket")
                            #print(k , ticketNo , state ," - SD ",  " --> " , "No Inbox Email" , fault_occur_time , ticket_create_time , "No Inbox Email", time_to_ticket_SNOW )
                            result = [[i, ticketNo , "None" , fault_occur_time , ticket_create_time , "None", 0 , 0, 0 , customerName ,customer_category , "Ok" , "Ok" , severity , opened_by , resolved_by , "Not Raised",assigned_to , resolver_assignment_group , time_of_day , type_of_day, 0 , day, week, state, fault] , category]
                            

            
            
            
                    else: #That ticket was never raised to any one
                        #print(ticketNo , "Not Raised At all via email")
                        customerName = str(str(ticketData[0][1]))

                        result = [[i, ticketNo ," None", fault_occur_time ,  ticket_create_time , "Not Raised" , 0 , 0 ,0 , customerName , "Ok" , "None" , "None" ,severity , opened_by , resolved_by , "Not Raised",assigned_to , resolver_assignment_group , time_of_day , type_of_day ,  "Within SLA" , day, week, state, fault] , category]
                        #SD_Data.append(SD_temp_list)
                        
            
                elif GC_Category == "GC_Proactive": #This is a proactive ticket
                    #print(ticketNo , " This is a Proactive Ticket")

                    category = "L1"


                    y = y + 1
                    sheet_name = "Sheet2"
            

                    #print(i , ticketNo , " --> ", "Pro Active")
                    
                    ticket_Email_Data = getFirstResponse(ticketNo) #get the first sent email
                    #Extract the emails being responded to which are our customer
                    str_emails_address = ticket_Email_Data

                    nce_fault_occur_time = ""
                    nce_fault_clear_time = ""
                    repair_time_difference = ""
                    repair_integrity_status = ""
                    overallIntegrity = ""
                    resolution_status = ""
                    mttc_time_difference = ""
                    mttn_time_difference = ""
                    NCE_mttc = 0
                    SNOW_mttc = 0

                    NCE_Ticket_Data = getNCE_TicketData(ticketNo)   # this returns 2 values , fault occur and fault clear in a list
                    



                    if str_emails_address is not None:   # This means that the ticket was raised via email

                        #print(ticketNo , "Raised Via Email")

                        notification_status = "Raised"
                        first_notification_time = str(str_emails_address[2])
                        MTTR = getTimeDifference(fault_occur_time , first_notification_time)
                        Data_Integrity = check_Data_Integrity(first_notification_time , fault_occur_time)

                        MTTR_SLA = get_MTTA_SLA(MTTR)
                        fault = ticket_Email_Data[1]
                        SNOW_mttn = getTimeDifference(fault_occur_time , first_notification_time)


                        if snow_fault_clear_time == "Not Resolved":
                            resolution_status = "Not Resolved"
                        else:
                            resolution_status = "Resolved"

                        
                        

                        if NCE_Ticket_Data is not None:        

                            nce_fault_occur_time = str(NCE_Ticket_Data[0])
                            nce_fault_clear_time = str(NCE_Ticket_Data[1])   

                            data_integrity = check_Data_Integrity( nce_fault_occur_time,  fault_occur_time)
                            time_difference = data_integrity[1]
                            integrity_status = data_integrity[0]

                            NCE_mttn = getTimeDifference(nce_fault_occur_time , first_notification_time)
                            NCE_mttc = getTimeDifference( nce_fault_occur_time, ticket_create_time)

                            if NCE_mttn < -1: #IF mEAN TIME TO NOTIFY IS less than 0 , USE THE SNOW Fault occur time
                                
                                NCE_mttn = getTimeDifference(ticket_create_time ,first_notification_time )
                                
                                if NCE_mttn < -1: #if Its still less that 0 , return error
                                    NCE_mttn = 0
                            
                            if NCE_mttc < -1: #IF mEAN TIME TO NOTIFY IS less than 0 , USE THE SNOW Fault occur time
                                
                                NCE_mttc = getTimeDifference( ticket_create_time, fault_occur_time)
                                
                                if NCE_mttc < -1: #if Its still less that 0 , return error
                                    NCE_mttc = SNOW_mttc
                            
                            mttc_time_difference = 0
                            mttn_time_difference = 0

                            try:
                                NCE_mttc = int(NCE_mttc)
                                SNOW_mttc = int(SNOW_mttc)

                                mttc_time_difference = NCE_mttc - SNOW_mttc
                            except ValueError as e:
                                mttc_time_difference = 0

                                print(f"Error converting values to integers: {e}")
                                # Handle the error as needed

                            try:
                                NCE_mttn = int(NCE_mttn)
                                SNOW_mttn = int(SNOW_mttn)

                                mttn_time_difference = NCE_mttn - SNOW_mttn
                            except ValueError as e:
                                mttn_time_difference = 0

                                print(f"Error converting values to integers: {e}")
                                # Handle the error as needed


                            #get Repair time integrity
                            repairTime_data_integrity = check_Repair_Time_Data_Integrity(nce_fault_clear_time , snow_fault_clear_time )
                            
                            repair_integrity_status = repairTime_data_integrity[0]
                            repair_time_difference = repairTime_data_integrity[1]                       
                            

                        else:

                            time_difference = 0
                            integrity_status = "Ok"
                            NCE_mttn = SNOW_mttn    
                            repair_integrity_status = "Ok"  
                            repair_time_difference = 0
                            nce_fault_occur_time = fault_occur_time
                            NCE_mttc = SNOW_mttc
                            nce_fault_clear_time = snow_fault_clear_time
                            
                            mttc_time_difference = 0
                            mttn_time_difference = 0


                        overallIntegrity = getOverallIntegrity(repair_integrity_status, integrity_status)

                        result = [[y,
                                ticketNo ,
                                opened_by ,
                                nce_fault_occur_time,
                                    fault_occur_time ,
                                    ticket_create_time ,
                                    first_notification_time ,
                                        time_difference,
                                        integrity_status ,
                                            NCE_mttn, 
                                            SNOW_mttn ,                                          
                                            SNOW_mttc,
                                            NCE_mttc,
                                                MTTR_SLA ,
                                                notification_status, 
                                                time_of_day ,
                                                    type_of_day ,
                                                    assigned_to ,
                                                    resolved_by ,                                                  
                                                    resolver_assignment_group ,
                                                        severity ,
                                                        nce_fault_clear_time,
                                                        snow_fault_clear_time ,
                                                        repair_time_difference,
                                                            repair_integrity_status ,
                                                            overallIntegrity,
                                                            resolution_status,
                                                            mttc_time_difference,
                                                                mttn_time_difference,
                                                                responsible_party,
                                                                "OPCO",
                                                                    day,
                                                                    week, 
                                                                    state,
                                                                    fault ] , category]
                                   
                        '''print  (k,
                                ticketNo , state , 
                                " - L1 ",                           
                                nce_fault_occur_time,
                                    fault_occur_time ,
                                    ticket_create_time ,
                                    first_notification_time ,
                                        time_difference,
                                        integrity_status ,
                                            NCE_mttn, 
                                            SNOW_mttn ,
                                            NCE_mttc,
                                            SNOW_mttc,
                                                MTTR_SLA ,
                                                opened_by ,
                                                notification_status, 
                                                time_of_day ,
                                                    type_of_day ,
                                                    assigned_to ,
                                                    resolved_by ,                                                  
                                                    resolver_assignment_group ,
                                                        severity ,
                                                        nce_fault_clear_time,
                                                        snow_fault_clear_time ,
                                                        repair_time_difference,
                                                            repair_integrity_status ,
                                                            overallIntegrity,
                                                            
                                                            
                                                                )#resolution_status,  mttc_time_difference, mttn_time_difference,
                        '''
                   
                   
                    else:

                        #print(ticketNo , "Proactive ticket not raised via Email")
                        
                        result = [[y,
                                ticketNo ,
                                opened_by ,
                                nce_fault_occur_time,
                                    fault_occur_time ,
                                    ticket_create_time ,
                                    "None" ,
                                        0,
                                        "Ok" ,
                                            0, 
                                            0 ,                                          
                                            NCE_mttc,
                                            SNOW_mttc,
                                                "Within SLA" ,
                                                "Not Raised", 
                                                time_of_day ,
                                                    type_of_day ,
                                                    assigned_to ,
                                                    resolved_by ,                                                  
                                                    resolver_assignment_group ,
                                                        severity ,
                                                        nce_fault_clear_time,
                                                        snow_fault_clear_time ,
                                                        repair_time_difference,
                                                            repair_integrity_status ,
                                                            overallIntegrity,
                                                            resolution_status,
                                                            mttc_time_difference,
                                                                mttn_time_difference,
                                                                responsible_party,
                                                                "OPCO",
                                                                    day,
                                                                    week, 
                                                                    state,
                                                                    fault ] , category]


            return result    

        except Exception as e:

            logging.error(traceback.format_exc())

def get_All_SNOW_Tickets(earliest_timestamp):
    #Get all tickets with a lesser timestamp from SNOW
    sql_query = f"SELECT Number FROM snow_data WHERE Created < '{earliest_timestamp}' order by Created desc" #Get even other tickets , to ensure all records are updated 
    # Execute the SQL query
    cursor.execute(sql_query)
    # Fetch All row returned by the query
    ticketData = cursor.fetchall()
    # Convert fetched data into a list of incident numbers
    old_Tickets = [row[0] for row in ticketData]
    return  old_Tickets
    #print("Older Tickets To Process: " , len(old_Tickets))

def get_Older_Tickets(file_path):
    # Read the first sheet
    df1 = pd.read_excel(file_path)

    # Read the second sheet
    df2 = pd.read_excel(file_path, sheet_name="Sheet2")  # Assuming the second sheet
    
    if len(df1) < 1 or len(df2) < 1:
        now = datetime.now()
        print("\nFile Is Empty, selecting all Tickets Earlier Than:", now)
        old_Tickets = get_All_SNOW_Tickets(now)
        return old_Tickets
    else:
        # Convert 'Ticket Create Time' column to datetime format
        try:
            df1['Ticket Create Time'] = pd.to_datetime(df1['Ticket Create Time'])
            df2['Ticket Create Time'] = pd.to_datetime(df2['Ticket Create Time'])
        except ValueError:
            # Handle non-convertible values gracefully (e.g., set them to NaT)
            df1['Ticket Create Time'] = pd.to_datetime(df1['Ticket Create Time'], errors='coerce')
            df2['Ticket Create Time'] = pd.to_datetime(df2['Ticket Create Time'], errors='coerce')

        # Get the earliest timestamp from both DataFrames
        earliest_timestamp_S1 = df1['Ticket Create Time'].min()
        earliest_timestamp_S2 = df2['Ticket Create Time'].min()

        # Pick the earliest of the two timestamps
        earliest_timestamp = min(earliest_timestamp_S1, earliest_timestamp_S2)
        print("Getting Tickets Earlier Than:", earliest_timestamp)
        old_Tickets = get_All_SNOW_Tickets(earliest_timestamp)
        
        return old_Tickets
    
def get_SNOW_Tickets():
    sql_query = f"SELECT Number FROM snow_data order by Created desc"
    cursor.execute(sql_query)
    # Fetch All row returned by the query
    ticketData = cursor.fetchall()
     # Convert fetched data into a list of incident numbers
    all_Tickets = [row[0] for row in ticketData]
    return all_Tickets

def get_Remaining_Tickets(file_path):
    try:
        incident_numbers = []
        # Read the first sheet
        df1 = pd.read_excel(file_path)

        # Read the second sheet
        df2 = pd.read_excel(file_path, "Sheet2")  # Assuming the second sheet

        # Extract the 'Incident No.' column from both filtered DataFrames and combine them into a single list
        incident_numbers = df1['Incident No'].tolist() + df2['Incident No'].tolist()

        return incident_numbers


    except Exception as e:
        print(f"Error: {e}")
        return None, None

def get_Ticket_List(file_path , update_status):

    tickets_To_Update = []
    if update_status == "Y":
        print("--------------Updating New Records--------------")

        updateDBRecords()

        delete_Non_Closed_Tickets(file_path)
        remaining_tickets = get_Remaining_Tickets(file_path)
        SNOW_Tickets = get_SNOW_Tickets()

        print("In Get_Ticket_List fn -----> Existing Records:" , len(remaining_tickets))
        print("Snow Tickets    :" , len(SNOW_Tickets))

        # Remove records from list_one that are also present in list_two
        tickets_To_Update = [record for record in SNOW_Tickets if record not in remaining_tickets]    
        print("tickets to update", len(tickets_To_Update))   
        return tickets_To_Update
    else:

        #No records have been updated from SNOW so continue updating any missing records.
        #get the last record in the file
        print("--------------Updating Older Records--------------")
        tickets_To_Update = get_Older_Tickets(file_path)
        
        return tickets_To_Update

def delete_Non_Closed_Tickets(file_path):
    # Load the Excel workbook
    wb = load_workbook(file_path)
    
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        if sheet == "Sheet1":
            col_index = 25
        elif sheet == "Sheet2":
            col_index = 34
        else: 
            continue
        
        rows_to_delete = []
        # Iterate through rows
        for row in ws.iter_rows(min_row=2):
            if row[col_index - 1].value != "Closed":
                rows_to_delete.append(row)
        
        print("Records To Update in " , sheet , " --> " , len(rows_to_delete) )
        # Delete the rows
        for row in rows_to_delete:
            ws.delete_rows(row[0].row)
    
    # Save the modified workbook
    wb.save(file_path)

def determine_Record_Update():

    ticketsDifference = 0
    print("***************Determine if Ticket Update is Needed********\n")  
    
    # Define the file path
    file_path = f"C:\\Users\\{username}\\MTN Group\\CIM Reports Hub - Documents\\SD_L1_Performance_MasterFile.xlsx"

    #we can check if number of tickets has changed since last run, basically what we have in SNOW vs What is in SDLI Masterfile
    try:
        # Load the workbook
        wb = openpyxl.load_workbook(file_path)
        # Select the first worksheet
        ws = wb.active
        # Count the number of rows with data
        num_rows = ws.max_row - 1

        snowTickets = get_SNOW_Tickets()
        Local_DB_Tickets = len(snowTickets)    

        print(f"Number of rows in the Excel file: {num_rows} \n Rows in snow : {Local_DB_Tickets}")
       
    except Exception as e:
        print(f"An error occurred: {e}")

    ticketsDifference = num_rows - Local_DB_Tickets
   

    try:
        # Get the last modified time of the file
        last_modified_time = os.path.getmtime(file_path)
        print("Modified Time:", datetime.fromtimestamp(last_modified_time))        
        
        # Get the current time
        current_time = time.time()
        print("Current Time :", datetime.fromtimestamp(current_time))

        # Calculate the difference in seconds
        time_difference = int((current_time - last_modified_time) /60)
        print("Time Difference:", time_difference, "Minutes")

        # Check if the time difference is greater than 4 hours
        if time_difference > 240 or not ticketsDifference == 0:  # Convert 240 seconds to hours
            # Do something if the file was last modified more than 4 hours ago
            print("The File needs an update..")
            # Example: You can perform any action here, such as sending a notification or performing a task.
            return "Y"
        else:
            # Do something else if the file was last modified within the last 4 hours
            print("The file was last modified within the last 4 hours.")
            return "N"
            
    except FileNotFoundError:
        print("File not found.")
    except Exception as e:
        print(f"An error occurred: {e}")

def main():  
    global L1_data 
    global SD_Data

    verify_Data_File_Integrity()    

    close_status = close_excel_file(file_path)
    print(close_status , "\n")

    update_status = determine_Record_Update()  
    print("------> Üpdate Status", update_status)

    ticket_List = get_Ticket_List(file_path , update_status) 
    print("-------> Length of ticket list",len(ticket_List))
    sys.exit(1)  

    if not len(ticket_List) == 0:
        #ticket_List = ['INC4266452']
        print("Records to update" , len(ticket_List) , "\n")

        #my_list = range(1, 10001)  # Assuming you have a list of 10000 elements
        # Loop through the list
        for index, ticketNo in enumerate(ticket_List, start=1):
            try:
            # Do something with the current item
            #print("Item:", ticketNo )
                processed_result= process_Ticket(ticketNo)   
            
                result = processed_result[0]
                category = processed_result[1]
                
                
                if category == "SD":
                    ticket_Create_Time = str(result[4])
                    week = str(result[23])
                    SD_Data.append(result)
                elif category == "L1":
                    ticket_Create_Time = str(result[5])
                    week = str(result[32])

                    L1_data.append(result)
                print(index , ticketNo ,ticket_Create_Time ,"-->" , category , week ) 
                
            
                        
                # Check if it's a multiple of 100
                if index % 50 == 0:

                    print("\nReached", index, "records ::" , "Saving Records")        
                    append_data_to_excel_v2(file_path, "Sheet2", L1_data)
                    append_data_to_excel_v2(file_path, "Sheet1", SD_Data)   
                    print("")         
                
                    L1_data.clear()
                    SD_Data.clear()
                    create_Master_Copy()
                    upload_To_SharePoint()
                    # Do something after every 100 records
            except Exception as e:
                print(e)
        
        print("\nReached", index, "records ::" , "Saving Records")        
        append_data_to_excel_v2(file_path, "Sheet2", L1_data)
        append_data_to_excel_v2(file_path, "Sheet1", SD_Data)   
        print("")         
        
        L1_data.clear()
        SD_Data.clear()
        create_Master_Copy()
        upload_To_SharePoint()

    else:
        print("All records updated")

    # Register the on_interrupt function for SIGINT, SIGTERM, and other signals
    signal.signal(signal.SIGINT, on_interrupt)
    signal.signal(signal.SIGTERM, on_interrupt)

main()

