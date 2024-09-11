from openpyxl import load_workbook
import os

def update_resistance_values(directory, excel_name, target_time=3600, tolerance=10):
    path = os.path.join(directory, excel_name)
    wb = load_workbook(path)
    ws = wb.active  # assuming the data is in the active sheet

    # Initialize column indices
    col_indices = {
        "time": None,
        "Set Temperature": None,
        "Set Current": None,
        "CH1 resistance": None,
        "CH2 resistance": None,
        "CH3 resistance": None,
        "CH4 resistance": None,
        "CH5 resistance": None,
        "CH6 resistance": None
    }

    # Single pass to find all relevant columns
    for cell in ws[1]:  # ws[1] accesses the first row
        for key in col_indices:
            if cell.value and key.lower() in cell.value.lower():
                col_indices[key] = cell.column - 1  # Store 0-based index for later use

    # Check if all necessary columns were found
    if None in col_indices.values():
        missing_cols = [k for k, v in col_indices.items() if v is None]
        print(f"Missing columns: {', '.join(missing_cols)}")
        return

    # Process all rows where time is approximately 3600 and retrieve relevant data
    for row in ws.iter_rows(min_row=2):  # Skipping the header row
        time_value = row[col_indices["time"]].value
        if target_time - tolerance <= time_value <= target_time + tolerance:
            for i in range(1, 7):  # For each channel
                data = {
                    "Channel": f"CH{i}",
                    "Set Temperature": row[col_indices["Set Temperature"]].value,
                    "Set Current": row[col_indices["Set Current"]].value,
                    "Resistance": row[col_indices[f"CH{i} resistance"]].value
                }
                print(data)
                fill_sheet2(directory, excel_name, data)

    wb.save(path)
    wb.close()

def fill_sheet2(directory, excel_name, data):
    path = os.path.join(directory, excel_name)
    wb = load_workbook(path)
    ws = wb['Sheet2']  # Directly access 'Sheet2'

    print(f"Attempting to fill data for: {data}")  # Debug information

    # Find the correct row and column based on the Channel, Temperature, and Current
    found = False
    for row in range(2, ws.max_row + 1):
        channel = ws.cell(row=row, column=1).value
        temperature = ws.cell(row=row, column=2).value
        if channel == data["Channel"] and temperature == data["Set Temperature"]:
            for col in range(3, ws.max_column + 1):
                current = ws.cell(row=1, column=col).value
                if current == data["Set Current"]:
                    ws.cell(row=row, column=col, value=data["Resistance"])
                    print(f"Data filled at Row: {row}, Column: {col}")  # Confirm data placement
                    found = True
                    break
        if found:
            break

    if not found:
        print("No matching row and column found for the data provided.")  # Debug if no place found

    wb.save(path)
    wb.close()

# Example usage:
directory = os.getcwd()
excel_name = 'Result.xlsx'
update_resistance_values(directory, excel_name)
