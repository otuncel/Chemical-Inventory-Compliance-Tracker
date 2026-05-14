import json
import os
import requests
import win32com.client
import time

def get_pubchem_info(cas):
    """Fetches the chemical compound name from the PubChem API."""
    try:
        url = f"https://pubchem.ncbi.nlm.nih.gov/rest/pug/compound/name/{cas}/property/Title/JSON"
        response = requests.get(url, timeout=5)
        if response.status_code == 200:
            data = response.json()
            return data['PropertyTable']['Properties'][0]['Title']
    except Exception:
        return None
    return None

def run_chemical_audit():
    # Set working directory to the script's location
    current_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(current_dir)
    db_name = 'regulation_db.json'
    
    print("--- CHEMICAL INVENTORY AUDIT STARTED ---")

    # Load the regulatory database
    try:
        if os.path.exists(db_name):
            with open(db_name, 'r', encoding='utf-8') as f:
                reg_db = json.load(f)
        else:
            print(f"WARNING: {db_name} not found. Proceeding with PubChem lookups only.")
            reg_db = {}
    except Exception as e:
        print(f"Database Error: {e}")
        return

    try:
        # Connect to the currently open Excel application
        excel = win32com.client.GetActiveObject("Excel.Application")
        wb = excel.ActiveWorkbook
        ws = wb.ActiveSheet
        
        print(f"Connected to File: {wb.Name}")
        print(f"Active Sheet: {ws.Name}")
        
        # Find the last used row based on Column B (CAS_No)
        last_row = ws.Cells(ws.Rows.Count, 2).End(-4162).Row

        for i in range(2, last_row + 1):
            cas_value = ws.Cells(i, 2).Value # Column B: CAS_No
            if not cas_value:
                continue
            
            cas = str(cas_value).strip()
            print(f"Processing: {cas}")

            # 1. PubChem Name Lookup (Fill only if Chemical_Name cell is empty)
            if not ws.Cells(i, 1).Value:
                name = get_pubchem_info(cas)
                if name:
                    ws.Cells(i, 1).Value = name
                    print(f"  > Name found: {name}")

            # 2. Regulatory Compliance Check
            status = "Compliant"
            action = "Checked & Verified"
            color_index = 4 # Green (Default)

            if cas in reg_db:
                status = reg_db[cas].get('status', 'Restricted')
                action = reg_db[cas].get('desc', 'Listed in regulation database')
                
                # Assign color based on status
                if status == "Prohibited":
                    color_index = 3 # Red
                elif status == "Restricted":
                    color_index = 6 # Yellow/Orange

            # 3. Update Excel Columns
            ws.Cells(i, 5).Value = status         # Compliance_Status
            ws.Cells(i, 6).Value = action         # Action_Required
            ws.Cells(i, 5).Interior.ColorIndex = color_index # Color the cell background
            
            # Short delay to prevent API rate limiting
            time.sleep(0.1)

        print("\n--- ALL RECORDS UPDATED SUCCESSFULLY ---")

    except Exception as e:
        print(f"ERROR: {e}")
        print("Tip: Ensure you are not in 'Edit Mode' (active cursor inside a cell) in Excel.")

if __name__ == "__main__":
    run_chemical_audit()
    # Keep the terminal open for a few seconds to review the results
    time.sleep(2)
