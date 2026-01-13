import os
import time


def get_username():
    # Get the username from the environment variables
    username = os.getenv('USERNAME')
    if username is None:
        # If USERNAME variable is not found, try USERPROFILE variable
        userprofile = os.getenv('USERPROFILE')
        if userprofile is not None:
            # Extract the username from the USERPROFILE path
            username = os.path.basename(userprofile)
    return username

# Get the username
username = get_username()

print(username)

# Specify the path to your Excel file
excel_file_path = f"C:\\Users\\{username}\\MTN Group\\CIM Reports Hub - Documents\\Automation_Scripts\\Automated_Scripts - Edit Capable.xlsm"

# Define the list of run times
run_times = ["00:10", "04:10", "08:10", "12:10", "16:10", "20:10"]

def run_update(current_time):
    # Open the Excel file with the default application
    os.startfile(excel_file_path)
    print("Running Update at:", current_time)

# Open the Excel file on startup
run_update(time.strftime("%H:%M"))

# Enter the loop to run the script at specified times
while True:

    if username == "monitoring.server1":
            # Add 3 hours to the current time
         current_time = time.strftime("%H:%M", time.localtime(time.time() + 3*60*60))
    else:
        # Get the current time
        current_time = time.strftime("%H:%M")

    # Check if the current time matches any of the specified run times
    if current_time in run_times:
        run_update(current_time)

    # Find the next run time
    next_run_time_index = 0
    for i, run_time in enumerate(run_times):
        if run_time > current_time:
            next_run_time_index = i
            break
    next_run_time = run_times[next_run_time_index]

    # Calculate the time until the next run
    current_time_minutes = int(current_time[:2]) * 60 + int(current_time[3:])
    next_run_time_minutes = int(next_run_time[:2]) * 60 + int(next_run_time[3:])
    time_until_next_run = next_run_time_minutes - current_time_minutes

    # Display the time until the next run
    print("Updating again in", time_until_next_run, "minutes")

    # Wait for 1 minute before checking the time again
    time.sleep(60)
