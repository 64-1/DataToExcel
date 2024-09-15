import pandas as pd
import numpy as np
import shutil
import os
import sys
from pathlib import Path
from openpyxl import load_workbook
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QFileDialog, QVBoxLayout



def read_folder_and_create_excel(directory):
    # Put all the folder names in the directory into a list
    folders = [folder for folder in os.listdir(directory) if os.path.isdir(os.path.join(directory, folder)) and '_duplicate' in folder]
    
    # Extract the temperature and current setting and save them in a dataframe
    for folder in folders:
        temperatures = []
        currents = []

        if 'C' in folder and 'A' in folder:
            temp_index = folder.index('C')
            curr_index = folder.index('A')

            try:
                temperature = float(folder[:temp_index])
                current = float(folder[temp_index+1:curr_index])
                temperatures.append(temperature)
                currents.append(current)
            except ValueError:
                continue
            
        filenumber = len([name for name in os.listdir(os.path.join(directory, folder)) if os.path.isfile(os.path.join(directory, folder, name))])

        df = pd.DataFrame({
            'Set Temperature (C)': temperatures * filenumber,
            'Set Current (A)': currents * filenumber
        })

        excel_path = os.path.join(directory, folder, f"{folder}.xlsx")

        # Write to excel file created
        if not os.path.exists(excel_path):
            df.to_excel(excel_path, index=False)
        else:
            print("Excel file already exists.")

def duplicate_and_rename_folder(directory):
    folders = [folder for folder in os.listdir(directory) if os.path.isdir(os.path.join(directory, folder)) and 'A' in folder and 'C' in folder and '_duplicate' not in folder]
    for folder in folders:
        source_folder_path = os.path.join(directory, folder)
        target_folder_path = os.path.join(directory, folder +'_duplicates')
        os.makedirs(target_folder_path, exist_ok=True)

        for filename in os.listdir(source_folder_path):
            original_file = os.path.join(source_folder_path, filename)
            duplicated_file = os.path.join(target_folder_path, filename)
            shutil.copy(original_file, duplicated_file)

            if filename.endswith('.all'):
                txt_filename = filename[:-4] + '.txt'
                txt_file_path = os.path.join(target_folder_path, txt_filename)
                shutil.move(duplicated_file, txt_file_path)
        
def read_and_update_excel(directory):
    folders = [folder for folder in os.listdir(directory) if os.path.isdir(os.path.join(directory, folder)) and '_duplicates' in folder]
    for folder in folders:
        folder_path = os.path.join(directory, folder)
        excel_path = os.path.join(folder_path, f"{folder}.xlsx")

        if os.path.exists(excel_path):
            df = pd.read_excel(excel_path)
        else:
            read_folder_and_create_excel(directory)
            df = pd.read_excel(excel_path)
        columns = ['time (s)']
        measurements = ['actual temperature (C)']
        for ch in range(1, 7):
            measurements += [f'CH{ch} actual current (A)', f'CH{ch} actual Voltage', f'CH{ch} resistance (ohm)']
        columns.extend(measurements)

        new_data = pd.DataFrame(columns=columns)

        # Use folder_path instead of folder
        for filename in os.listdir(folder_path):
            if filename.endswith('.txt'):
                file_path = os.path.join(folder_path, filename)
                with open(file_path, 'r') as file:
                    lines = [line.strip().split() for line in file.readlines()[:7]]
                    data = np.array(lines, dtype=float)

                    time = data[0, 0]
                    temperature = data[0, 1]
                    currents = data[1:, 6]
                    voltages = data[1:, 0]
                    resistances = voltages / currents

                    row = [time, temperature] + [val for pair in zip(currents, voltages, resistances) for val in pair]
                    new_data.loc[len(new_data)] = row

        if 'df' in locals():
            full_df = pd.merge(df, new_data, left_index=True, right_index=True)
            full_df.to_excel(excel_path, index=False)
            path = Path(excel_path)
            new_file_name = path.stem.replace('_duplicates', '') + path.suffix
            # Define the destination path (moving the file to its parent directory)
            destination_path = path.parent.parent / new_file_name

            # Move the file to its parent directory
            shutil.move(str(path), str(destination_path))
        else:
            new_data.to_excel(excel_path, index=False)

        # Remove the duplicates folder after processing
        shutil.rmtree(folder_path)

def combine_excel_files(directory, output_filename):
    # Get all Excel files in the directory and subdirectories
    excel_files = [os.path.join(dp, f) for dp, dn, filenames in os.walk(directory) for f in filenames if f.endswith('.xlsx')]
    
    # Exclude 'Result.xlsx' if it already exists
    excel_files = [file for file in excel_files if not file.endswith(output_filename)]
    
    df_list = [pd.read_excel(file) for file in excel_files]
    
    if df_list:
        combined_df = pd.concat(df_list, ignore_index=True)
        output_path = os.path.join(directory, output_filename)
        combined_df.to_excel(output_path, index=False)
    else:
        print("No Excel files found to combine.")

def remove_other_excel_files(directory, keep_file):
    for dp, dn, filenames in os.walk(directory):
        for file in filenames:
            if file.endswith('.xlsx') and file != keep_file:
                file_path = os.path.join(dp, file)
                os.remove(file_path)

def add_sheet_excel(directory, excel_name):
    # Put all the folder names in the directory into a list
    path = os.path.join(directory, excel_name)
    folders = [folder for folder in os.listdir(directory) if os.path.isdir(os.path.join(directory, folder))]
    temp_values = []
    current_values = []
    # Extract the temperature and current setting and save them in a dataframe
    for folder in folders:
        if 'C' in folder and 'A' in folder:
            temp_index = folder.index('C')
            curr_index = folder.index('A')

            try:
                temperature = float(folder[:temp_index])
                current = float(folder[temp_index+1:curr_index])
                temp_values.append(temperature)
                current_values.append(current)
            except ValueError:
                continue

    T = sorted(list(set(temp_values)))
    I = sorted(list(set(current_values)))

    index = pd.MultiIndex.from_product([range(1, 7), T], names=["Channel", "T (C)\\I (A)"])
    df = pd.DataFrame(index=index, columns=I)

    with pd.ExcelWriter(path, engine='openpyxl', mode='a') as writer:
        df.to_excel(writer, sheet_name='Sheet2')

    wb = load_workbook(path)
    ws = wb['Sheet2']

    # Unmerge all merged cells in the worksheet
    merged_cells = list(ws.merged_cells)
    for merged_cell in merged_cells:
        ws.unmerge_cells(range_string=str(merged_cell))

    # Write the channel labels and temperature values into the appropriate cells
    start_row = 2  # Adjust based on header rows in your DataFrame
    total_temperatures = len(T)
    for ch in range(1, 7):
        for i in range(total_temperatures):
            row = start_row + (ch - 1) * total_temperatures + i
            ws.cell(row=row, column=1, value=f'CH{ch}')  # Channel
            ws.cell(row=row, column=2, value=T[i])       # Temperature

    wb.save(path)
    wb.close()

def sort_column(directory, excel_name):
    target_column = 0  
    path = os.path.join(directory, excel_name)
    # Load the existing workbook
    workbook = load_workbook(path)
    sheet = workbook.active  # Get the active sheet

    # Read all data from the sheet
    data = list(sheet.iter_rows(values_only=True))

    # Separate headers and data
    labels = data[0]    # Don't sort the headers
    data = data[1:]     # Data begins on the second row

    # Sort data by the target column
    data.sort(key=lambda x: x[target_column])

    # Write sorted data back into the same sheet
    for idx, label in enumerate(labels):
        sheet.cell(row=1, column=idx+1, value=label)

    for idx_r, row in enumerate(data):
        for idx_c, value in enumerate(row):
            sheet.cell(row=idx_r+2, column=idx_c+1, value=value)

    # Save the modified workbook back to the same file
    workbook.save(path)

def update_resistance_values(directory, excel_name):
    path = os.path.join(directory, excel_name)
    wb = load_workbook(path)
    ws = wb.active  # Assuming the data is in the active sheet

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

    # Find all relevant columns in the header row
    for cell in ws[1]:  # ws[1] accesses the first row
        for key in col_indices:
            if cell.value and key.lower() in str(cell.value).lower():
                col_indices[key] = cell.column

    # Debug to check column indices
    print("Column Indices:", col_indices)

    # Build a dictionary to store the maximum time for each (Set Temperature, Set Current)
    max_time_rows = {}  # Key: (Set Temperature, Set Current), Value: (max_time, row_number)

    # Iterate over all data rows to find the maximum time for each combination
    for row_number in range(2, ws.max_row + 1):
        try:
            set_temp = float(ws.cell(row=row_number, column=col_indices["Set Temperature"]).value)
            set_current = float(ws.cell(row=row_number, column=col_indices["Set Current"]).value)
            time_value = float(ws.cell(row=row_number, column=col_indices["time"]).value)
        except (TypeError, ValueError):
            continue  # Skip rows with invalid data

        key = (set_temp, set_current)
        if key not in max_time_rows or time_value > max_time_rows[key][0]:
            max_time_rows[key] = (time_value, row_number)

    # Open Sheet2
    ws_sheet2 = wb['Sheet2']

    # Get the currents from the header row of Sheet2
    currents = []
    for col in range(3, ws_sheet2.max_column + 1):
        current_value = ws_sheet2.cell(row=1, column=col).value
        try:
            currents.append(float(current_value))
        except (TypeError, ValueError):
            continue  # Skip if the value cannot be converted to float

    # Process each row with the maximum time
    for (set_temp, set_current), (max_time, row_number) in max_time_rows.items():
        data = {
            "Set Temperature": set_temp,
            "Set Current": set_current,
        }
        for i in range(1, 7):  # For each channel
            resistance_key = f"CH{i} resistance"
            resistance_value = ws.cell(row=row_number, column=col_indices[resistance_key]).value
            data["Channel"] = f"CH{i}"
            data["Resistance"] = resistance_value
            fill_sheet2(ws_sheet2, currents, data)

    wb.save(path)
    wb.close()

def fill_sheet2(ws_sheet2, currents, data):
    # Convert the Set Current to float for comparison
    try:
        set_current = float(data["Set Current"])
    except (TypeError, ValueError):
        print(f"Invalid Set Current: {data['Set Current']}")
        return

    # Find the column index for the Set Current
    try:
        current_index = currents.index(set_current) + 3  # +3 because currents start from column 3
    except ValueError:
        print(f"Set Current {set_current} not found in currents.")
        return

    # Find the row that matches the Channel and Set Temperature
    found_row = None
    for row in range(2, ws_sheet2.max_row + 1):
        channel_value = ws_sheet2.cell(row=row, column=1).value
        temperature_value = ws_sheet2.cell(row=row, column=2).value

        # Ensure temperature values are comparable as floats
        try:
            temp_value = float(temperature_value)
            set_temp = float(data["Set Temperature"])
        except (TypeError, ValueError):
            continue  # Skip rows with invalid temperature values

        if channel_value == data["Channel"] and temp_value == set_temp:
            found_row = row
            break

    if found_row is None:
        print(f"Row with Channel {data['Channel']} and Temperature {data['Set Temperature']} not found.")
        return

    # Write the Resistance value into the correct cell
    ws_sheet2.cell(row=found_row, column=current_index, value=data["Resistance"])


# Add the FolderSelector class for the GUI
class FolderSelector(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.selected_directory = None

    def initUI(self):
        self.setWindowTitle('Select Folder for Operation')
        self.setGeometry(100, 100, 400, 100)

        self.layout = QVBoxLayout()

        self.btn_select_folder = QPushButton('Select Folder', self)
        self.btn_select_folder.clicked.connect(self.open_folder_dialog)
        self.layout.addWidget(self.btn_select_folder)

        self.setLayout(self.layout)

    def open_folder_dialog(self):
        folder = QFileDialog.getExistingDirectory(self, 'Select Folder')
        if folder:
            self.selected_directory = folder
            self.close()  # Close the GUI after selection

def main(directory_path):
    duplicate_and_rename_folder(directory_path)
    read_folder_and_create_excel(directory_path)
    read_and_update_excel(directory_path)
    combine_excel_files(directory_path, 'Result.xlsx')
    sort_column(directory_path, 'Result.xlsx')
    remove_other_excel_files(directory_path, 'Result.xlsx')
    add_sheet_excel(directory_path, 'Result.xlsx')
    update_resistance_values(directory_path, 'Result.xlsx')

if __name__ == "__main__":
    app = QApplication(sys.argv)
    selector = FolderSelector()
    selector.show()
    app.exec_()

    # After the GUI is closed, check if a directory was selected
    if selector.selected_directory:
        selected_directory = selector.selected_directory
        main(selected_directory)
    else:
        print("No folder was selected.")
