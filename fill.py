import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import Alignment

def update_resistance_values(directory, excel_name, target_time=3600, tolerance=10):
    path = os.path.join(directory, excel_name)
    wb = load_workbook(path)
    ws = wb.active  # assuming the data is in the active sheet

    # Dynamically find the time column
    time_col_index = None
    for col in ws.iter_cols(min_row=1, max_row=1, values_only=True):
        if "time" in col.value.lower():
            time_col_index = col.column
            break

    # Assuming temperature and current are immediately after time in the next two columns
    temp_col_index = None
    for col in ws.iter_cols(min_row=1, max_row=1, values_only=True):
        if "Set Temperature" in col.value:
            temp_col_index = col.column
            break
    current_col_index = None
    for col in ws.iter_cols(min_row=1, max_row=1, values_only=True):
        if "Set Current" in col.value:
            current_col_index = col.column
            break
    CH1resistance_col_index = None
    for col in ws.iter_cols(min_row=1, max_row=1, values_only=True):
        if "CH1 resistance" in col.value:
            CH1resistance_col_index = col.column
            break
    CH2resistance_col_index = None
    for col in ws.iter_cols(min_row=1, max_row=1, values_only=True):
        if "CH2 resistance" in col.value:
            CH2resistance_col_index = col.column
            break
    CH3resistance_col_index = None
    for col in ws.iter_cols(min_row=1, max_row=1, values_only=True):
        if "CH3 resistance" in col.value:
            CH3resistance_col_index = col.column
            break

    CH4resistance_col_index = None
    for col in ws.iter_cols(min_row=1, max_row=1, values_only=True):
        if "CH4 resistance" in col.value:
            CH4resistance_col_index = col.column
            break
    
    CH5resistance_col_index = None
    for col in ws.iter_cols(min_row=1, max_row=1, values_only=True):
        if "CH5 resistance" in col.value:
            CH5resistance_col_index = col.column
            break
    
    CH6resistance_col_index = None
    for col in ws.iter_cols(min_row=1, max_row=1, values_only=True):
        if "CH6 resistance" in col.value:
            CH6resistance_col_index = col.column
            break
    
    CHValues = []
    # Find the row where time is approximately 3600 and retrieve temperature and current
    for row in ws.iter_rows(min_row=2):  # Skipping the header row
        time_value = row[time_col_index - 1].value  # Adjusting index for 0-based list access
        if target_time - tolerance <= time_value <= target_time + tolerance:
            set_temp = row[temp_col_index - 1].value
            set_current = row[current_col_index - 1].value
            set_CH1 = row[CH1resistance_col_index-1].value
            set_CH2 = row[CH2resistance_col_index-1].value
            set_CH3 = row[CH3resistance_col_index-1].value
            set_CH4 = row[CH4resistance_col_index-1].value
            set_CH5 = row[CH5resistance_col_index-1].value
            set_CH6 = row[CH6resistance_col_index-1].value
            CHValues.append(set_CH1)
            CHValues.append(set_CH2)
            CHValues.append(set_CH3)
            CHValues.append(set_CH4)
            CHValues.append(set_CH5)
            CHValues.append(set_CH6)
            # update_sheet2(wb, set_temp, set_current)
            break

    wb.save(path)
    wb.close()

directory = os.getcwd()
excel_name = 'Result.xlsx'
update_resistance_values(directory, excel_name)
