from openpyxl import load_workbook
import os

def fill_sheet2(directory, excel_name, data):
    path = os.path.join(directory, excel_name)
    wb = load_workbook(path)
    ws = wb.get_sheet_by_name('Sheet2')  # Assuming 'Sheet2' already exists

    # Assuming the structure of data is a list of dictionaries with keys 'Channel', 'Set Temperature', 'Set Current', and 'Resistance'
    for entry in data:
        # Find the correct row based on the Channel and Temperature
        channel = entry['Channel']
        temperature = entry['Set Temperature']
        current = entry['Set Current']
        resistance = entry['Resistance']
        
        # Find the row
        for row in range(2, ws.max_row + 1):
            if ws.cell(row=row, column=1).value == channel and ws.cell(row=row, column=2).value == temperature:
                # Find the correct column for the current
                for col in range(3, ws.max_column + 1):
                    if ws.cell(row=1, column=col).value == current:
                        ws.cell(row=row, column=col, value=resistance)
                        break

    wb.save(path)
    wb.close()

# Sample data structure
data = [
    {'Channel': 'CH1', 'Set Temperature': 50, 'Set Current': 0.3, 'Resistance': 10},
    # Add more data as needed
]

directory = os.getcwd()
excel_name = 'Result.xlsx'
fill_sheet2(directory, excel_name, data)
