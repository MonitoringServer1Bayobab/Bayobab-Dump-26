import zipfile
import os
 
def extract_zip(zip_file_path, extract_to):
    # Check if the zip file exists
    if not os.path.exists(zip_file_path):
        print(f"Error: Zip file '{zip_file_path}' does not exist.")
        return
   
    # Create the extraction directory if it doesn't exist
    os.makedirs(extract_to, exist_ok=True)
   
    # Extract the zip file
    with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
        zip_ref.extractall(extract_to)
   
    print(f"Files extracted from '{zip_file_path}' to '{extract_to}'.")
 
zip_file_path = r"C:\Users\lynette.mutuku\Downloads\L3VPN Service E2E Quality Report_Weekly_20250331000000_20250407135912.zip"
# zip_file_path = r"C:\Users\lynette.mutuku\Downloads\HistoricalAlarms20250401114544630.zip"
extract_to = r"C:\Users\lynette.mutuku\Documents\Availability L3VPN"
# extract_to = r"C:\Users\lynette.mutuku\Downloads\HistoricalAlarms20250401114544630"
 
extract_zip(zip_file_path, extract_to)