import subprocess
import psutil
import mysql.connector
import win32com.client
import requests
from datetime import datetime
import pywintypes
import signal
import sys
import subprocess
from datetime import datetime, timedelta
import getpass
import os
import re
import gzip
import shutil

import time
import urllib.request
from urllib.error import URLError

#Get logged in username
username = getpass.getuser()
print("Current User Logged In: ", username)
  # Connect to the MySQL database
cnx = mysql.connector.connect(
host='127.0.0.1',
user='vicky',
password='ILoveBlack@2022',
# user='lyn',
# password='0gDUnc55',
database='CSCMailsBackup',
charset='utf8mb4',
collation='utf8mb4_general_ci'
)


#Check if outlook is open, if open, proceed, Else open and proceed
def is_outlook_open():
    for process in psutil.process_iter(['name']):
        if process.info['name'] == 'OUTLOOK.EXE':
            return True
    return False
def open_outlook():
    subprocess.Popen(["C:\\Program Files\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE"])  # Replace with the actual path

if not is_outlook_open():
    open_outlook()
else:
    print("Outlook is already open.")

def on_interrupt(signal, frame):
    # Handle the interrupt here
    print("Interrupt signal received:", signal)
    print("Closing Connection")
    # Close the cursor and database connection
    cursor.close()
    cnx.close()
    sys.exit(0)


def insertDB(Mailsubject , SentOn , ticketNo , Engineer , recipients , recipientsNames , mailType , category, to_emails_str , to_names_str , cc_emails_str , cc_names_str):

    def execute_query(query , values):
        try:
            
            # Execute the query
            cursor.execute(query , values)

            # Commit the changes to the database
            cnx.commit()

            # Close the cursor and database connection
            #cursor.close()
            #cnx.close()

            return "OK"  # Return "OK" if the query execution is successful

        except mysql.connector.Error as err:
            return "Fail: {}".format(err)  # Return "Fail" along with the error message

    table_name = "mails"
    sql = "INSERT INTO {} (MailSubject, MailDate , IncidentNo,Engineer,EmailAddress ,EmailName,MailType ,MailCategory,To_Email,To_Name , CC_Name, CC_Email) VALUES (%s, %s, %s , %s, %s, %s , %s , %s, %s, %s , %s , %s)".format(table_name)
    values = (Mailsubject , SentOn , ticketNo , Engineer , recipients , recipientsNames , mailType , category, to_emails_str , to_names_str , cc_emails_str , cc_names_str)

    return  execute_query(sql , values)


def getEngineer(body):
    index = body.find("Service Desk Engineer")
    if index > 0:
        tmpStr = body[0:index].strip()        
        spaceindex1 = tmpStr.rfind("\n")
        engineer = tmpStr[spaceindex1:len(tmpStr)].strip()
        if len(engineer) >= 50:
            return "Not Found"
        else:
            return engineer
    else:
        index = body.find("L1 Support Engineer")
        tmpStr = body[0:index].strip()        
        spaceindex1 = tmpStr.rfind("\n")
        engineer = tmpStr[spaceindex1:len(tmpStr)].strip()
        if len(engineer) >= 50:
            return "Not Found"
        else:
            return engineer

def getTicketNo(subject):
    # Using regular expression to find the pattern "INC" followed by exactly 7 digits
    matches = re.findall(r'INC[1-9]\d{6}\b', subject)
    
    if matches:
        # If matches are found, return the first match
        return matches[0]
    else:
        # If no matches found, try to find "RITM" and return it
        index = subject.rfind("RITM")
        if index >= 0:
            ticketNo = subject[index:index + 11]
            return ticketNo
        else:
            # If neither "INC" nor "RITM" found, return None
            return None


def getMailType(subject):
    subject = subject.lower()
    #print(subject)
    #print(subject.find('re:'))

    if subject.find('re:') == 0:
        return "Reply"
    else:
        return "Notification"

def add_second_to_timestamp(timestamp_str):
    # Convert the input timestamp string to a datetime object
    dt_obj = datetime.strptime(timestamp_str, '%Y-%m-%d %H:%M:%S')

    # Add a second to the datetime object
    dt_obj_with_second = dt_obj + timedelta(seconds=2)

    # Convert the updated datetime object back to a timestamp string
    updated_timestamp_str = dt_obj_with_second.strftime('%Y-%m-%d %H:%M:%S')

    return updated_timestamp_str

def convert_To_GMT(timestamp_str):
    # Convert the input timestamp string to a datetime object
    dt_obj = datetime.strptime(timestamp_str, '%Y-%m-%d %H:%M:%S')

    # Get the current system timezone
    system_timezone = datetime.now().astimezone().tzinfo
    #print(system_timezone)
   
    

    # Get the offset from GMT for the current system timezone
    current_offset = system_timezone.utcoffset(datetime.now())

    # Adjust the datetime object to GMT
    dt_obj_GMT = dt_obj - current_offset

    # Convert the GMT datetime object back to a timestamp string
    GMT_timestamp_str = dt_obj_GMT.strftime('%Y-%m-%d %H:%M:%S')

    #print("Before Conversion: " , timestamp_str)
    #print("After Conversion: " , GMT_timestamp_str)

    
    return GMT_timestamp_str

def getLastRecordDate(MailCategory):
        
    # Execute the SQL query
    query = f'SELECT * FROM Mails WHERE MailCategory = "{MailCategory}" ORDER BY MailDate DESC LIMIT 1'
    cursor.execute(query)
    result = cursor.fetchone()

    if result:
        # Extract the desired column value from the query result
        column_value = result[2]  # Assuming the desired value is in the second column (index 2)
        # Convert the column value to a datetime object
        column_datetime = datetime.combine(column_value.date(), column_value.time())
        return column_datetime

def extract_email(address):
    # Parsing out email addresses
    if "cn=" in address:
        start_index = address.find("cn=")
        end_index = address.find("-", start_index)
        if end_index != -1:
            return address[start_index + 3:end_index]
    # If parsing fails, return the original address
    return address

def updateInboxMails(InboxItems):

    print(" ")
    print("Total Emails in Inbox: " , len(InboxItems))

    undelivered = "Undelivered"
    reply = "reply"
    category = "Inbox"
    subject_filters = ["Automatic reply:", "automatique", "Undelivered"]
    sender_filters = ["postmaster", "Fixed-CSC@bayobab.africa"]

    latestEmail = str(str(getLastRecordDate("Inbox")))
    latestEmail = add_second_to_timestamp(latestEmail)
    
    #restriction = f"@SQL=urn:schemas:httpmail:datereceived > '{latestEmail}'"  
    subject_condition = " AND NOT ("
    for subject_filter in subject_filters:
        subject_condition += f"\"urn:schemas:httpmail:subject\" LIKE '%{subject_filter}%' OR "
    subject_condition = subject_condition[:-4]  # Remove the last " OR "
    subject_condition += ")"
    

    restriction = f"@SQL=urn:schemas:httpmail:datereceived > '{latestEmail}' AND NOT (\"urn:schemas:httpmail:fromname\" LIKE '%{sender_filters[0]}%' OR \"urn:schemas:httpmail:fromname\" LIKE '{sender_filters[1]}')" + subject_condition
    print(restriction)
    InboxItems = Inbox.Items.Restrict(restriction)
    print("Filtered Items in Inbox: " , len(InboxItems))

    print(latestEmail)
    index = len(InboxItems) - 1

    for i in range(0 , len(InboxItems)):    
                        
        remainingItems = len(InboxItems) - index
        email = InboxItems[index]

        try:

            Mailsubject = str(email.Subject)

            if Mailsubject.find(undelivered) == -1 or Mailsubject.find(reply) == -1:
                Mailsubject = str(email.Subject)                
                
                # Access the SentOn property
                ReceivedTime = email.ReceivedTime
                ReceivedTime = str(email.ReceivedTime)
                ReceivedTime = ReceivedTime[0:19]                 
                ReceivedTime = convert_To_GMT(ReceivedTime)
           
                #print()
                #print(index , "Received On : " , ReceivedTime  , "Subject : " ,Mailsubject)
                to_emails = []

                try:
                    sender_email = None
                    sender_name = None
                    to_emails = []
                    to_names = []
                    cc_emails = []
                    cc_names = []

                    for recipient in email.Recipients:
                        if recipient.Type == 1:  # Type 1 represents "To" recipients
                            to_email = None
                            to_name = None
                            if recipient.AddressEntry:
                                try:
                                    if recipient.AddressEntry.Type == "EX":
                                        to_email = recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress
                                    elif recipient.AddressEntry.Type == "SMTP":
                                        to_email = recipient.AddressEntry.Address
                                    to_name = recipient.Name
                                except AttributeError:
                                    to_email = None  # or assign a default value
                            to_emails.append(to_email)
                            to_names.append(to_name)
                        elif recipient.Type == 2:  # Type 2 represents "CC" recipients
                            cc_email = None
                            cc_name = None
                            if recipient.AddressEntry:
                                try:
                                    if recipient.AddressEntry.Type == "EX":
                                        cc_email = recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress
                                    elif recipient.AddressEntry.Type == "SMTP":
                                        cc_email = recipient.AddressEntry.Address
                                    cc_name = recipient.Name
                                except AttributeError:
                                    cc_email = None  # or assign a default value
                            cc_emails.append(cc_email)
                            cc_names.append(cc_name)

                    # Retrieve sender's email address and name
                    if email.SenderEmailType == "SMTP":
                        sender_email = email.SenderEmailAddress
                    else:
                        try:
                            sender_email = email.Sender.GetExchangeUser().PrimarySmtpAddress
                        except AttributeError:
                            sender_email = None  # or assign a default value
                    sender_name = email.SenderName

                except pywintypes.com_error as e:
                    print(f"Error occurred while accessing recipients property: {e}")

                to_emails_str = ';'.join(filter(None, to_emails))
                to_names_str = ';'.join(filter(None, to_names))
                cc_emails_str = ';'.join(filter(None, cc_emails))
                cc_names_str = ';'.join(filter(None, cc_names))
                

             
                senderName = str(email.SenderName)
                ticketNo = getTicketNo(Mailsubject)
                mailType = getMailType(Mailsubject)
                Engineer = ""

                currentRecordDateString = ReceivedTime
                format_string = "%Y-%m-%d %H:%M:%S"
                currentRecordDateString = datetime.strptime(currentRecordDateString , format_string)

                queryStatus = insertDB(Mailsubject , ReceivedTime , ticketNo , Engineer , sender_email , sender_name , mailType , category , to_emails_str , to_names_str , cc_emails_str , cc_names_str)  
                queryStatus = "Ok"
                print(remainingItems  , "--> " ,  ReceivedTime ," : ", queryStatus , "Rem : " , index , "-->", ticketNo , "-->", Mailsubject)


                #print("CC Emails" , to_emails_str)
                #print("CC Names" , cc_names , "\n")

                #print("To Emails" , to_emails) 
                #print("To Names" , to_names , "\n") 

                #print("Sender" , sender_email)
             

        except pywintypes.com_error as e:
                        print(f"Error occurred while accessing recepients property: {e}")
        index =index - 1
        
def updateSentMails(SentItems):

    print(" ")
    print("Total Emails in SentBox: " , len(SentItems))
    category = "Sent"


    latestEmail = str(getLastRecordDate("Sent"))
    latestEmail = add_second_to_timestamp(latestEmail)
    
    restriction = f"@SQL=urn:schemas:httpmail:date > '{latestEmail}'"  
    print(restriction)

    SentItems = Sentbox.Items.Restrict(restriction)
    print("Filtered Items in Sentbox: " , len(SentItems))

    #print(latestEmail)
    index = len(SentItems) - 1

    for i in range(0 , len(SentItems)):    
                        
        remainingItems = len(SentItems) - index
        email = SentItems[index]
       
        recipients_list = []
        recepients_name = []

        # Check if SentOn property is accessible
        try:
            # Access the SentOn property               
            SentOn = str(email.SentOn)
            SentOn = SentOn[0:19]
            SentOn = convert_To_GMT(SentOn)

            recipients = email.Recipients

            
            #print("Recepients", email.Recipients)
            for recipient in email.Recipients:
                recipient_address = recipient.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E")
                recipients_list.append(recipient_address)
                recepients_name.append(recipient.Name)
                #print(recipient.Type  , recipient)

# Extracting CC field
            

            #print(email.Recipients)
            recipients = "; ".join(recipients_list)
            recipientsNames = "; ".join(recepients_name)               
            body = str(email.Body)
            Mailsubject = str(email.Subject)       
            Engineer = getEngineer(body)
            ticketNo = getTicketNo(Mailsubject)
            mailType = getMailType(Mailsubject)

            '''
            #print(count ," :")
            print("Subject : " ,Mailsubject)
            print("Sent On : " , SentOn)
            print("Ticket No: " , ticketNo)
            print("Engineer : " , Engineer)
            print("Mail Type ", mailType)
            print("Recipients Mails:", recipients)
            print("Recipients Names:", recipientsNames)
            print("-----------------------------------------")
            print("")
            '''
            

            # Append data to the DB
            queryStatus = insertDB(Mailsubject , SentOn , ticketNo , Engineer , recipients , recipientsNames , mailType , category,"" ,"" ,"","")
            #queryStatus = "Ok"
            #query = "use CSC_Emails_DataLake"
            #result = execute_query(query)
           
            print(remainingItems  , "--> " ,  SentOn ," : ", queryStatus , "Rem : " , index , "-->", ticketNo , "-->", Mailsubject)
            
        except pywintypes.com_error as e:
            print(f"Error occurred while accessing recepients property: {e}")


        index =index - 1

branch_name = 'master'
# repo_dir=r'C:\Users\lynette.mutuku\Documents\Bayobab-Py-N-Dump'
repo_dir=r'C:\Users\monitoring.server1\QE GIT\Bayobab-Py-N-Dump'

def run_command(command, check=True):
    """Run a command in the shell and handle exceptions."""
    try:
        subprocess.run(command, shell=True, check=check)
    except subprocess.CalledProcessError as e:
        print(f"Error running command: {e}")

    #     """ Helper function to run shell commands """
    # result = subprocess.run(command, shell=True, capture_output=True, text=True)
    # if result.returncode != 0:
    #     print(f"Command failed with error: {result.stderr}")
    #     raise subprocess.CalledProcessError(result.returncode, command)
    # return result.stdout

# def automate_rebase(repo_dir, file_name='dump.sql.gz', keep_commits_count=20):
#     """Keep only the latest 20 commits that affect the specified file."""
#     os.chdir(repo_dir)
    
#     # Find commits related to the file
#     commits = run_command(f"git log --pretty=format:'%H' -- {file_name}").split('\n')

#     if len(commits) <= keep_commits_count:
#         print("Number of commits affecting the file is within the limit. No rebase needed.")
#         return

#     # Commits to be dropped
#     commits_to_drop = commits[:-keep_commits_count]

#     # Create a rebase script to drop excess commits
#     rebase_script = "\n".join(f"drop {commit}" for commit in commits_to_drop)
#     with open("rebase_script.txt", "w") as file:
#         file.write(rebase_script)

#     # Perform the rebase
#     try:
#         # Rebase only the commits related to the file
#         os.environ['GIT_EDITOR'] = ':'  # Setting GIT_EDITOR to ':' to prevent Vim from opening
#         run_command(f"git rebase -i --autosquash HEAD~{len(commits)} < rebase_script.txt", check=True)
#         print("Rebase completed. Only the latest 20 commits are retained.")
#     except subprocess.CalledProcessError:
#         print("Rebase failed. Please resolve conflicts or handle the rebase manually.")
#         return

#     # Force-push the updated branch to the remote repository
#     try:
#         run_command("git push --force", check=True)
#         print("Force-pushed the changes to the remote repository.")
#     except subprocess.CalledProcessError:
#         print("Failed to push the changes to the remote repository.")

def automate_rebase(repo_dir, file_name='dump.sql.gz', keep_commits_count=20):
    """Automate Git rebase and handle commits related to a specific file."""
    os.chdir(repo_dir)
    
    # Find commits related to the file
    try:
        commits = subprocess.check_output(
            f"git log --pretty=format:'%H' -- {file_name}", shell=True
        ).decode().strip().split('\n')
    except subprocess.CalledProcessError as e:
        print(f"Error finding commits: {e}")
        return

    if not commits:
        print("No commits found for the specified file.")
        return

    print(f"Commits affecting {file_name}: {commits}")

    # Verify the number of commits to avoid invalid upstream
    try:
        commit_count = int(subprocess.check_output("git rev-list --count HEAD", shell=True).decode().strip())
        if commit_count <= keep_commits_count:
            print("Number of commits is less than or equal to the number of commits to keep. Skipping rebase.")
            return
    except subprocess.CalledProcessError as e:
        print(f"Error counting commits: {e}")
        return

    # Create a rebase script to drop commits
    rebase_script = []
    for commit in commits:
        rebase_script.append(f"drop {commit}")

    # Add commands to the script to execute the rebase
    rebase_script.append("pick HEAD")
    rebase_script = "\n".join(rebase_script)

    with open("rebase_script.txt", "w") as file:
        file.write(rebase_script)

    # Perform the rebase
    try:
        subprocess.run(f"git rebase -i HEAD~{keep_commits_count} < rebase_script.txt", shell=True, check=True)
        print("Rebase completed.")
    except subprocess.CalledProcessError as e:
        print(f"Rebase failed: {e}")
        print("Handling rebase exception...")
        handle_rebase_exception()


def handle_rebase_exception():
    """Handle rebase exception by aborting or cleaning up."""
    rebase_dir = os.path.join('.git', 'rebase-merge')
    if os.path.exists(rebase_dir):
        print("Rebase is in progress. Attempting to abort...")
        run_command("git rebase --abort")
        print("Rebase aborted.")
    else:
        print("No rebase in progress. Removing rebase-merge directory.")
        run_command(f"rm -fr {rebase_dir}")
        print("Rebase-merge directory removed.")



def mysql_dump(host, user, password, database, output_file='dump.sql'):
    if username == "monitoring.server1":
    # if username == "lynette.mutuku":

        # local directory with the Git Repo. Should always exist 
        # dump_dir = r'C:\Users\lynette.mutuku\Documents\Bayobab-Py-N-Dump'
        dump_dir = r'C:\Users\monitoring.server1\QE GIT\Bayobab-Py-N-Dump'

        output_file_path = os.path.join(dump_dir, output_file)
        print(output_file_path)

        dump_command = [
            'mysqldump',
            '-h', host,
            '-u', user,
            '-p' + password,
            '--databases', database,
            '--result-file=' + output_file_path
        ]

        mysql_bin_path = r'C:\Program Files\MySQL\MySQL Server 8.0\bin'

        try:
            os.chdir(mysql_bin_path)  # Change directory to MySQL bin directory
            subprocess.run(dump_command, check=True)
            print("MySQL dump completed successfully.")

            return output_file_path
        except subprocess.CalledProcessError as e:
            print("Error:", e)
    else:
        print("User " , username ," Not Allowed to dump db")

def compress_and_delete_sql_file(file_name):
    # Set input and output paths
    input_path = os.path.join(os.path.dirname(__file__), file_name)
    output_path = input_path + '.gz'
    
    # Compress the file
    with open(input_path, 'rb') as f_in:
        with gzip.open(output_path, 'wb') as f_out:
            shutil.copyfileobj(f_in, f_out)
    
    # Delete the original file
    try:
        os.remove(input_path)
        print(f"Original file '{file_name}' deleted successfully.")
    except OSError as e:
        print(f"Error deleting file: {e}")

    return output_path

def git_push(repo_dir, branch_name, file_name='dump.sql.gz'):
        now = datetime.now()

        # Format the date and time
        formatted_date_time = now.strftime('%d/%m/%Y %H:%M:%S')

        # Determine the day suffix
        day = now.day
        if 10 <= day <= 20:
            suffix = 'th'
        else:
            suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(day % 10, 'th')

        commit_message = f'Automated push on {formatted_date_time} GMT'

        # Navigate to the repository directory
        os.chdir(repo_dir)
        print('Ín git push', os.chdir(repo_dir))
        print('Repo directory:', os.getcwd())

        # Check if .git directory exists
        if not os.path.exists(os.path.join(repo_dir, '.git')):
            print("Error: Not a Git repository.")
            return
    
        # run_command(f"git pull origin {branch_name}")

        # Add changes
        run_command("git add dump.sql.gz")
        # run_command("git add rebase_script.txt")

        print('Staged')
        
        try:
            run_command(f"git commit -m \"{commit_message}\"")
        except subprocess.CalledProcessError as e:
            print(f"Error committing changes: {e}")
            return

        # Automate rebase for specific file
        # automate_rebase(repo_dir, file_name)

        # Force push to the remote repository
        run_command(f"git push --force origin {branch_name}")
        print("File pushed to GitHub successfully.")
    

# Register the on_interrupt function for SIGINT, SIGTERM, and other signals
signal.signal(signal.SIGINT, on_interrupt)
signal.signal(signal.SIGTERM, on_interrupt)

#Define time wait period for next update
waitTime = 600

while (True):
    try:
# Create a cursor to execute SQL queries
        cursor = cnx.cursor()
        if(cursor):
            print("Connection Established to Database")
        else:
            print("Connection Failed")
    except mysql.connector.Error as err:
        print("Fail: {}".format(err))  # Return "Fail" along with the error message
    try:
        try:
                # Check internet connectivity
                urllib.request.urlopen('http://www.google.com', timeout=1)
                print("Internet connection is available.")
                badEmail = 0     

                try:
                    outlook = win32com.client.Dispatch("Outlook.Application")
                    namespace = outlook.GetNamespace("MAPI")
                    for account in namespace.Accounts:
                        if account.DisplayName == "Fixed-CSC@bayobab.africa":
                            store = account.DeliveryStore
                            Inbox = store.GetDefaultFolder(6)  # 6 represents the Inbox Items folder
                            Sentbox = store.GetDefaultFolder(5)  # 5 represents the Sent Items folder
                        
                            break

                    mysql_dump('localhost', 'vicky', 'ILoveBlack@2022', 'cscmailsbackup')
                    # mysql_dump('localhost', 'lyn', '0gDUnc55', 'cscmailsbackup')
                    compress_and_delete_sql_file('dump.sql')
                    git_push(repo_dir, branch_name)


                    print(" ")
                    print("Selected Account:", account.DisplayName) 
                    InboxItems = Inbox.Items  
                    SentItems = Sentbox.Items

                    updateInboxMails(InboxItems)
                    updateSentMails(SentItems)
                    mysql_dump('localhost', 'vicky', 'ILoveBlack@2022', 'cscmailsbackup')
                    # mysql_dump('localhost', 'lyn', '0gDUnc55', 'cscmailsbackup')
                    compress_and_delete_sql_file('dump.sql')
                    git_push(repo_dir, branch_name)
                   

                    print("Calling Gods_Eye")
                    # Use subprocess to run the other Python file
                    subprocess.run(['python', "Gods_Eye_Final.py"])


                    print(" ")
                    print("Waiting for " , waitTime ," Seconds to proceed")
                    time.sleep(waitTime)          

                    
                except pywintypes.com_error as e:
                    print(f"Error occurred while accessing Outlook: {e}")


        except URLError as e:
                print("Internet connection is not available.")
                # Your code when the internet is disconnected goes here
                # Replace this with your actual code
            # Wait for a period of time before checking the internet connection again
                time.sleep(5)

    except Exception as e:
        print(f"An error occurred: {e}")