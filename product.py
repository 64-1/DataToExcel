import os
import sys
import numpy as np
import pandas as pd
from openpyxl import load_workbook
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QFileDialog, QVBoxLayout

def read_data_from_folders(directory):
    folders = [folder for folder in os.listdir(directory) if os.path.isdir(os.path.join(directory, folder)) and 'A' in folder and 'C' in folder]
    all_data = []
    for folder in folders:
        folder_path = os.path.join(directory, folder)

        if 'C' in folder and 'A' in folder:
            temp_index = folder.index('C')
            curr_index = folder.index('A')

            try:
                temperature = float(folder[:temp_index])
                current = float(folder[temp_index+1:curr_index])
            except ValueError:
                continue
        else:
            continue

        # Process all .txt files in the folder
        for filename in os.listdir(folder_path):
            if filename.endswith('.txt'):
                file_path = os.path.join(folder_path, filename)
                with open(file_path, 'r') as file:
                    lines = [line.strip().split() for line in file.readlines()[:7]]
                    data = np.array(lines, dtype=float)

                    time = data[0, 0]
                    set_temperature = temperature
                    set_current = current
                    actual_temperature = data[0, 1]
                    currents = data[1:, 6]
                    voltages = data[1:, 0]
                    resistances = voltages / currents

                    # For each channel, create a row
                    for i, ch in enumerate(range(1, 7)):
                        row = {
                            'time': time,
                            'Set Temperature': set_temperature,
                            'Set Current': set_current,
                            'actual temperature': actual_temperature,
                            'Channel': f'CH{ch}',
                            'Actual Current': currents[i],
                            'Actual Voltage': voltages[i],
                            'Resistance': resistances[i]
                        }
                        all_data.append(row)
    df = pd.DataFrame(all_data)
    return df

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

    # Sort data by the target column (Set Temperature)
    data.sort(key=lambda x: x[target_column])

    # Write sorted data back into the same sheet
    for idx, label in enumerate(labels):
        sheet.cell(row=1, column=idx+1, value=label)

    for idx_r, row in enumerate(data):
        for idx_c, value in enumerate(row):
            sheet.cell(row=idx_r+2, column=idx_c+1, value=value)

    # Save the modified workbook back to the same file
    workbook.save(path)

def add_sheet_excel(directory, excel_name):
    path = os.path.join(directory, excel_name)
    # Read the main data
    df = pd.read_excel(path)

    # Get unique temperatures and currents
    T = sorted(df['Set Temperature'].unique())
    I = sorted(df['Set Current'].unique())

    # Create a workbook and add Sheet2
    wb = load_workbook(path)
    if 'Sheet2' in wb.sheetnames:
        ws = wb['Sheet2']
    else:
        ws = wb.create_sheet('Sheet2')

    start_row = 1
    for ch in sorted(df['Channel'].unique()):
        # Add Channel label
        ws.cell(row=start_row, column=1, value=ch)

        # Get data for this channel
        ch_data = df[df['Channel'] == ch]

        # Pivot the data so that temperatures are rows and currents are columns
        pivot_table = ch_data.pivot_table(values='Resistance', index='Set Temperature', columns='Set Current')

        # Sort the index and columns
        pivot_table = pivot_table.reindex(index=T, columns=I)

        # Convert pivot_table to list of lists
        # First row: header with currents
        header = ['T (C)\\I (A)'] + [str(i) for i in I]
        data_rows = [header]
        for temp in T:
            row = [temp]
            for current in I:
                value = pivot_table.loc[temp, current] if (temp in pivot_table.index and current in pivot_table.columns) else None
                row.append(value)
            data_rows.append(row)

        # Write the data
        for i, row_data in enumerate(data_rows):
            for j, value in enumerate(row_data):
                ws.cell(row=start_row + i + 1, column=j + 1, value=value)

        # Move to the next channel, adding 2 empty rows
        start_row = start_row + len(data_rows) + 2

    # Save the workbook
    wb.save(path)
    wb.close()

def update_resistance_values(directory, excel_name):
    path = os.path.join(directory, excel_name)
    wb = load_workbook(path)
    ws_main = wb.active  # Assuming the data is in the active sheet

    # Initialize column indices
    col_indices = {
        "time": None,
        "Set Temperature": None,
        "Set Current": None,
        "Channel": None,
        "Resistance": None
    }

    # Find all relevant columns in the header row
    for cell in ws_main[1]:
        for key in col_indices:
            if cell.value and key.lower() in str(cell.value).lower():
                col_indices[key] = cell.column

    # Build a dictionary to store the maximum time for each (Channel, Set Temperature, Set Current)
    max_time_rows = {}  # Key: (Channel, Set Temperature, Set Current), Value: (max_time, row_number, resistance)

    # Iterate over all data rows to find the maximum time for each combination
    for row_number in range(2, ws_main.max_row + 1):
        try:
            set_temp = float(ws_main.cell(row=row_number, column=col_indices["Set Temperature"]).value)
            set_current = float(ws_main.cell(row=row_number, column=col_indices["Set Current"]).value)
            time_value = float(ws_main.cell(row=row_number, column=col_indices["time"]).value)
            channel = ws_main.cell(row=row_number, column=col_indices["Channel"]).value
            resistance = ws_main.cell(row=row_number, column=col_indices["Resistance"]).value
        except (TypeError, ValueError):
            continue  # Skip rows with invalid data

        key = (channel, set_temp, set_current)
        if key not in max_time_rows or time_value > max_time_rows[key][0]:
            max_time_rows[key] = (time_value, row_number, resistance)

    # Open Sheet2
    ws_sheet2 = wb['Sheet2']

    # Build a mapping of channel to starting row in Sheet2
    channel_start_rows = {}
    row_num = 1
    while row_num <= ws_sheet2.max_row:
        cell_value = ws_sheet2.cell(row=row_num, column=1).value
        if cell_value and isinstance(cell_value, str) and cell_value.startswith('CH'):
            channel = cell_value
            channel_start_rows[channel] = row_num
            # Skip the header rows (assume header is 2 rows)
            num_data_rows = len(set(df['Set Temperature'])) + 1  # +1 for header
            row_num += num_data_rows + 2  # +2 for data rows and empty rows
        else:
            row_num += 1

    # Process each entry with the maximum time
    for (channel, set_temp, set_current), (max_time, row_number, resistance_value) in max_time_rows.items():
        # Find the starting row for the channel in Sheet2
        if channel not in channel_start_rows:
            print(f"Channel {channel} not found in Sheet2.")
            continue
        start_row = channel_start_rows[channel] + 1  # Data starts 1 row below the channel label

        # Find the currents in the header row (start_row + 1)
        currents_row = start_row + 1
        currents = {}
        for col in range(2, ws_sheet2.max_column + 1):
            current_value = ws_sheet2.cell(row=currents_row, column=col).value
            try:
                current_value = float(current_value)
                currents[current_value] = col
            except (TypeError, ValueError):
                continue  # Skip if the value cannot be converted to float

        if set_current not in currents:
            print(f"Set Current {set_current} not found in Sheet2 for {channel}.")
            continue
        current_col = currents[set_current]

        # Find the row that matches the Set Temperature
        found_row = None
        for row in range(currents_row + 1, currents_row + 1 + len(set(df['Set Temperature']))):
            temperature_value = ws_sheet2.cell(row=row, column=1).value
            try:
                temp_value = float(temperature_value)
            except (TypeError, ValueError):
                continue  # Skip rows with invalid temperature values

            if temp_value == set_temp:
                found_row = row
                break

        if found_row is None:
            print(f"Temperature {set_temp} not found in Sheet2 for {channel}.")
            continue

        # Write the Resistance value into the correct cell
        ws_sheet2.cell(row=found_row, column=current_col, value=resistance_value)

    wb.save(path)
    wb.close()

# PyQt5 GUI for folder selection
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
    print("Reading data from folders...")
    df = read_data_from_folders(directory_path)
    print("Data read successfully.")

    # Write data to 'Result.xlsx'
    output_path = os.path.join(directory_path, 'Result.xlsx')
    df.to_excel(output_path, index=False)
    print(f"Data written to {output_path}")

    # Sort the data
    sort_column(directory_path, 'Result.xlsx')
    print("Data sorted by Set Temperature.")

    # Add Sheet2 and format it
    add_sheet_excel(directory_path, 'Result.xlsx')
    print("Sheet2 added and formatted.")

    # Update resistance values in Sheet2
    update_resistance_values(directory_path, 'Result.xlsx')
    print("Resistance values updated in Sheet2.")

    print("Processing completed.")

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
            
