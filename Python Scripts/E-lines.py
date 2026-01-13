import pandas as pd
import os

# === File Paths ===
file_a = r'C:\Users\lynette.mutuku\Downloads\SA_Circuits.xlsx'
bayobab1 = r'C:\Users\lynette.mutuku\Downloads\Bayobab_File1.xlsx'
bayobab2 = r'C:\Users\lynette.mutuku\Downloads\Bayobab_File2.xlsx'
output_file = r'C:\Users\lynette.mutuku\Downloads\Common.xlsx'

# === Load SITEIDs from File A ===
df_a = pd.read_excel(file_a)
site_ids = set(df_a['CUSTOMER_BSIDE_ID'].astype(str).tolist())
# print("✅ Loaded SITEIDs:", site_ids)
print("🔢 SITEID count:", len(site_ids))


xls_b = pd.ExcelFile(bayobab1)

# === Create Excel writer ===
with pd.ExcelWriter(output_file, engine='openpyxl', mode='w') as writer:
    for sheet_name in xls_b.sheet_names:
        print(f"\n🔍 Checking sheet: {sheet_name}")
        df_b_sheet = pd.read_excel(bayobab1, sheet_name=sheet_name)
        
        matched_rows = []
        matched_siteids = set()

        # === Search for matching rows and track SITEIDs ===
        for idx, row in df_b_sheet.iterrows():
            row_str = row.astype(str)
            common_ids = site_ids.intersection(set(row_str))
            if common_ids:
                matched_rows.append(row)
                matched_siteids.update(common_ids)

        # === Write to output file ===
        if matched_rows:
            result_df = pd.DataFrame(matched_rows)
            result_df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"✅ Found {len(matched_rows)} matching row(s) in '{sheet_name}'.")
            print(f"   Matched SITEIDs: {matched_siteids}")
        else:
            pd.DataFrame([["No matches found"]]).to_excel(writer, sheet_name=sheet_name, header=False, index=False)
            print(f"⚠️ No matches found in '{sheet_name}'.")

