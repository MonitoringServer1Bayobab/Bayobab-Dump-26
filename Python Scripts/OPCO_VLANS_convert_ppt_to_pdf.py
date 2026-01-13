import win32com.client
import os

def batch_convert_ppt_to_pdf(folder_path):
    if not os.path.exists(folder_path):
        print(f"❌ Folder not found: {folder_path}")
        return

    # Start PowerPoint application
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = 1  # Set to 0 if you want to run it silently

    # Loop through all files in the folder
    for file in os.listdir(folder_path):
        if file.lower().endswith(('.ppt', '.pptx')):
            input_path = os.path.join(folder_path, file)
            output_path = os.path.splitext(input_path)[0] + '.pdf'

            try:
                presentation = powerpoint.Presentations.Open(input_path, WithWindow=False)
                presentation.SaveAs(output_path, FileFormat=32)  # 32 = PDF
                presentation.Close()
                print(f"✅ Converted: {file} → {os.path.basename(output_path)}")
            except Exception as e:
                print(f"❌ Failed to convert {file}: {e}")

    powerpoint.Quit()
    print("🏁 Batch conversion completed.")

folder = r"C:\Users\lynette.mutuku\Downloads\November"
batch_convert_ppt_to_pdf(folder)
