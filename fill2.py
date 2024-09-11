from openpyxl import load_workbook
import os

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
excel_name = 'test.xlsx'
update_resistance_values(directory, excel_name)
