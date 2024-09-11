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

    # Find the row where time is approximately 3600 and retrieve relevant data
    for row in ws.iter_rows(min_row=2):  # Skipping the header row
        time_value = row[col_indices["time"]].value
        if target_time - tolerance <= time_value <= target_time + tolerance:
            values = {
                "Set Temperature": row[col_indices["Set Temperature"]].value,
                "Set Current": row[col_indices["Set Current"]].value,
                "Resistances": [row[col_indices[f"CH{i} resistance"]].value for i in range(1, 7)]
            }
            # Here you would call your function to update sheet2
            print("Values found:", values)  # Replace this with your update function
            break

    wb.save(path)
    wb.close()

# Example usage:
directory = os.getcwd()
excel_name = 'Result.xlsx'
update_resistance_values(directory, excel_name)
