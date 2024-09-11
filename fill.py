from openpyxl import load_workbook
import os

def update_resistance_values(directory, excel_name, target_time=3600, tolerance=10):
    path = os.path.join(directory, excel_name)
    wb = load_workbook(path)
    ws = wb.active  # assuming the data is in the active sheet

    # Dynamically find the time column
    time_col_index = None
    for col in ws.iter_cols(min_row=1, max_row=1, values_only=True):
        if "time" in col[0].lower():  # Assuming the header contains "time"
            time_col_index = col[0].column
            break

    if not time_col_index:
        print("Time column not found.")
        return

    # Assuming temperature and current are immediately after time in the next two columns
    temp_col_index = time_col_index + 1
    current_col_index = time_col_index + 2

    # Find the row where time is approximately 3600 and retrieve temperature and current
    for row in ws.iter_rows(min_row=2):  # Skipping the header row
        time_value = row[time_col_index - 1].value  # Adjusting index for 0-based list access
        if target_time - tolerance <= time_value <= target_time + tolerance:
            set_temp = row[temp_col_index - 1].value
            set_current = row[current_col_index - 1].value
            update_sheet2(wb, set_temp, set_current)
            break

    wb.save(path)
    wb.close()

def update_sheet2(wb, set_temp, set_current):
    # Access or create the sheet to update
    if 'Sheet2' in wb.sheetnames:
        ws = wb['Sheet2']
    else:
        ws = wb.create_sheet('Sheet2')

    # Update Sheet2 with the temperature and current values
    # Assuming row 2 is where the data should start (adjust as needed)
    ws.cell(row=2, column=1, value="Set Temperature")
    ws.cell(row=2, column=2, value=set_temp)
    ws.cell(row=3, column=1, value="Set Current")
    ws.cell(row=3, column=2, value=set_current)

    print("Sheet2 updated with set temperature and current values.")
