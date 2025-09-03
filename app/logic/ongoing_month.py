import openpyxl
import pandas as pd
import json
import os

# Class to help with file path management
class ResourceHelper:
    @staticmethod
    def get_path(relative_path: str) -> str:
        base_path = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base_path, relative_path)

# Get the path to the JSON configuration file
file_path_json = ResourceHelper.get_path('../config/inputan.json')

# Read the JSON configuration file
with open(file_path_json, 'r') as file:
    json_data = json.load(file)

# ======== Path to Excel files ========
data_final_path = json_data["final_file"]
data_summary_path = json_data["summary_file"]

# ======== Load data from file B ========
data_summary = pd.read_excel(data_summary_path, sheet_name='ITM Summary', header=json_data["header_month1"])
data_summary.columns = data_summary.columns.str.strip()

# Debug
print("Column names in file B:", data_summary.columns.tolist())

# ======== Load file A and select the sheet ========
wb = openpyxl.load_workbook(data_final_path)
ws = wb['ITM Summary']

# Mapping Column
columns_to_update = {
    'No.': 'A', 
    'Month': 'C', 
    'Company': 'D', 
    'Name of Vessel': 'F',
    'Buyer': 'G', 
    'End user': 'H', 
    'Load Port': 'I',
    'ETA/ATA': 'J',
    'ETB': 'L', 
    'ETD': 'N', 
    'Total': 'BJ', 
    'Lay': 'BK', 
    'can': 'BM',
    '%': 'BO', 
    'Status': 'BP', 
    'TM (AR)': 'BR', 
    'M (AD)': 'BS',
    'ASH (AD)': 'BT', 
    'ASH (AR)': 'BU', 
    'TS (AD)': 'BV', 
    'TS (AR)': 'BW',
    'CV (AD)': 'BX', 
    'CV (AR)': 'BY', 
    'CV (NAR)': 'BZ'
}

start_row = 4
data_count = start_row + json_data["data_count_month1"]

# ======== Fill standard columns ========
for column_name, excel_column in columns_to_update.items():
    for index in range(start_row - 4, data_count - 3):
        try:
            ws[f'{excel_column}{start_row + index}'] = data_summary[column_name].iloc[index]
        except KeyError:
            print(f"Column '{column_name}' not found in file B.")
        except IndexError:
            print(f"Insufficient data in column '{column_name}' for index {index}.")

# ======== Format date function ========
def convert_to_date_format(date_value):
    try:
        if isinstance(date_value, str) and '/' in date_value:
            day, month = date_value.split('/')
            date_converted = pd.to_datetime(f'2024-{month}-{day}', format='%Y-%m-%d')
        else:
            date_converted = pd.to_datetime(date_value)
        return f"{date_converted.day}.{date_converted.strftime('%b')}"
    except Exception as e:
        print(f"Error during value conversion '{date_value}': {e}")
        return None

# ======== Format ETA/ATA, ETB, ETD ========
for row in range(start_row, data_count + 1):
    try:
        eta = ws[f'J{row}'].value
        etb = ws[f'L{row}'].value
        etd = ws[f'N{row}'].value

        if eta:
            ws[f'K{row}'] = convert_to_date_format(eta)
        if etb:
            ws[f'M{row}'] = convert_to_date_format(etb)
        if etd:
            ws[f'O{row}'] = convert_to_date_format(etd)
    except Exception as e:
        print(f"Error converting dates at row {row}: {e}")

# ======== Format Lay and Can ========
for row in range(start_row, data_count + 1):
    try:
        lay = ws[f'BK{row}'].value
        can = ws[f'BM{row}'].value

        if lay:
            ws[f'BL{row}'] = convert_to_date_format(lay)
        if can:
            ws[f'BN{row}'] = convert_to_date_format(can)
    except Exception as e:
        print(f"Error converting Lay/Can at row {row}: {e}")

# ======== No Mahakam based on Load Port ========
no_mahakam = 1
for row in range(start_row, data_count + 1):
    load_port = ws[f'I{row}'].value
    if load_port == 'BoCT':
        ws[f'B{row}'] = 0
    elif load_port in ['SMD Anc', 'Bunyut']:
        ws[f'B{row}'] = no_mahakam
        no_mahakam += 1
    else:
        ws[f'B{row}'] = no_mahakam
        no_mahakam += 1

# ======== Type of Shipment based on Name of Vessel ========
for row in range(start_row, data_count + 1):
    name_vessel = ws[f'F{row}'].value
    if isinstance(name_vessel, str):
        vessel_upper = name_vessel.upper()
        if vessel_upper.startswith('MV'):
            ws[f'E{row}'] = 'Vessel'
        elif vessel_upper.startswith('BG'):
            ws[f'E{row}'] = 'Direct Shipment'
        elif 'DUMP TRUCK' in vessel_upper:
            ws[f'E{row}'] = 'Dump Truck'
        else:
            ws[f'E{row}'] = None

# ======== Fill columns for BoCT only ========
data_b = data_summary  # assuming 'data_summary' is same as 'data_b'

for row in range(start_row, data_count + 1):
    load_port = ws[f'I{row}'].value
    if load_port == 'BoCT':
        try:
            offset = row - start_row
            ws[f'P{row}'] = data_b['IMM-WB.HCV.LS'].iloc[offset]
            ws[f'Q{row}'] = data_b['IMM-WB.MCV.HS'].iloc[offset]
            ws[f'R{row}'] = data_b['IMM-EB.MCV.LS'].iloc[offset]
            ws[f'S{row}'] = data_b['IMM-EB.MCV.MS'].iloc[offset]
            ws[f'T{row}'] = data_b['IMM-EB.MCV.HS'].iloc[offset]
            ws[f'U{row}'] = data_b['TCM.HCV.LS'].iloc[offset]
            ws[f'V{row}'] = data_b['TCM.HCV.HS'].iloc[offset]
            ws[f'W{row}'] = data_b['TCM.LCV.MS.HA'].iloc[offset]
            ws[f'X{row}'] = data_b['BEK.MCV.LS'].iloc[offset]
            ws[f'Y{row}'] = data_b['BEK.HCV.MS'].iloc[offset]
            ws[f'Z{row}'] = data_b['JBG'].iloc[offset]
            ws[f'AA{row}'] = data_b['GPK'].iloc[offset]
            ws[f'AB{row}'] = data_b['TIS'].iloc[offset]

            # Third Party
            ws[f'AC{row}'] = data_b['EBH.HCV'].iloc[offset]
            ws[f'AE{row}'] = data_b['KMIA.MCV.LS'].iloc[offset]
            ws[f'AF{row}'] = data_b['BBE.MCV'].iloc[offset]
            ws[f'AG{row}'] = data_b['MBL.MCV'].iloc[offset]
            ws[f'AH{row}'] = data_b['MBL.56.MCV'].iloc[offset]
            ws[f'AJ{row}'] = data_b['EMJ.MCV'].iloc[offset]
            ws[f'AL{row}'] = data_b['KBM.MCV'].iloc[offset]
            ws[f'AM{row}'] = data_b['IKJ.MCV'].iloc[offset]
            ws[f'AN{row}'] = data_b['MKE.LCV'].iloc[offset]
            ws[f'AQ{row}'] = data_b['KJA.LCV.LS'].iloc[offset]
            ws[f'AS{row}'] = data_b['DMP.LCV'].iloc[offset] 
            ws[f'AT{row}'] = data_b['MCM.LCV'].iloc[offset]
            ws[f'AU{row}'] = data_b['BMM.LCV'].iloc[offset]
            ws[f'AW{row}'] = data_b['BBA.LCV'].iloc[offset]
            ws[f'AX{row}'] = data_b['MML.LCV'].iloc[offset]
            ws[f'BA{row}'] = data_b['KPM.LCV'].iloc[offset]
            ws[f'BB{row}'] = data_b['KJM.LCV'].iloc[offset]
            ws[f'BF{row}'] = data_b['BUM.LCV'].iloc[offset]
            ws[f'BG{row}'] = data_b['BISM.LCV'].iloc[offset]
        except IndexError:
            print(f"There is not enough data to fill the column in the row {row}.")

# ======== Save workbook ========
wb.save(data_final_path)
print("Excel file has been updated and saved.")
wb.close()

print("The columns have been successfully updated and saved back to the same file... :)")

# Main function placeholder
def main():
    print("ongoing_month processing")

# Run main if script executed directly
if __name__ == "__main__":
    main()