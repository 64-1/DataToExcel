from openpyxl import load_workbook

def update_resistance_values(directory, excel_name, target_time=3600, tolerance=10):
    path = os.path.join(directory, excel_name)
    wb = load_workbook(path)
    ws = wb.active  # assuming the data is in the active sheet

    # Assuming time values are in column A and resistance columns for CH1 to CH6 start from a specific column
    time_col = 'A'
    resistance_columns = {'CH1': 'C', 'CH2': 'D', 'CH3': 'E', 'CH4': 'F', 'CH5': 'G', 'CH6': 'H'}  # Update as necessary

    # Find the row where time is approximately 3600
    for row in ws.iter_rows(min_row=2):  # Skipping the header row
        time_value = row[ws[time_col].column - 1].value
        if target_time - tolerance <= time_value <= target_time + tolerance:
            # Extract resistance values
            resistances = {ch: row[ws[col].column - 1].value for ch, col in resistance_columns.items()}
            update_sheet2(wb, resistances)
            break

    wb.save(path)
    wb.close()

def update_sheet2(wb, resistances):
    # Access or create the sheet to update
    if 'Sheet2' in wb.sheetnames:
        ws = wb['Sheet2']
    else:
        ws = wb.create_sheet('Sheet2')

    # Assuming Sheet2 is prepared to receive the data in a structured format
    # For example, if the sheet is prepared with channels in rows and properties in columns
    start_row = 2  # Adjust starting row as needed
    for ch, resistance in resistances.items():
        # Assuming each channel's resistance goes in a different row consecutively
        row = start_row + int(ch[2:]) - 1  # CH1 starts at row 2, CH2 at row 3, etc.
        ws.cell(row=row, column=3, value=resistance)  # Assuming resistance values go in column 3

    print("Sheet2 updated with resistance values.")
  
