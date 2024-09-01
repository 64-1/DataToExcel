import pandas as pd
import numpy as np
import shutil
import os

def read_folder_and_create_excel(directory):
    # Put all the folder names in the directory into a list
    items = os.listdir(directory)

    # filter out files, only keeping the directories
    folders = [item for item in items if os.path.isdir(os.path.join(directory, item))]

    temperatures = []
    currents = []

    # Extract the temperature and current setting and save them in a dataframe
    for folder in folders:
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

    df = pd.DataFrame({
        'Set Temperature (C)': temperatures * 15,
        'Set Current (A)': currents * 15
    })

    excel_path = os.path.join(directory, "output.xlsx")

    # Write to excel file created
    if not os.path.exists(excel_path):
        df.to_excel(excel_path, index=False)
        print("Excel file has been created successfully!")
    else:
        print("Excel file already exists.")

def duplicate_and_rename_folder(directory):
    folders = [folder for folder in os.listdir(directory) if os.path.isdir(os.path.join(directory, folder))]
    folders.sort()
    if folders:
        first_folder_path = os.path.join(directory, folders[0])
        target_folder_path = os.path.join(directory, folders[0]+'_duplicates')
        os.makedirs(target_folder_path, exist_ok=True)

        for filename in os.listdir(first_folder_path):
            original_file = os.path.join(first_folder_path, filename)
            duplicated_file = os.path.join(target_folder_path, filename)
            shutil.copy(original_file, duplicated_file)

            if filename.endswith('.all'):
                txt_filename = filename[:-4] + '.txt'
                txt_file_path = os.path.join(target_folder_path, txt_filename)
                shutil.move(duplicated_file, txt_file_path)
        
        print(f"Files from {folders[0]} have been duplicated and converted where necessary.")

def read_and_update_excel(directory, excel_path):
    if os.path.exists(excel_path):
        df = pd.read_excel(excel_path)
    else:
        read_folder_and_create_excel(directory)
        df = pd.read.excel(excel_path)
    
    target_folder_path = os.path.join(directory, os.listdir(directory)[0] + '_duplicates')

    columns = ['time (s)']
    measurements = ['actual temperature (C)']
    for ch in range(1, 7):
        measurements +=[f'CH{ch} actual current (A)', f'CH{ch} actual Voltage', f'CH{ch} resistance (ohm)']
    columns.extend(measurements)

    new_data = pd.DataFrame(columns=columns)

    for filename in os.listdir(target_folder_path):
        if filename.endswith('.txt'):
            file_path = os.path.join(target_folder_path, filename)
            with open(file_path, 'r') as file:
                lines = [line.strip().split() for line in file.readlines()[:7]]
                data = np.array(lines, dtype=float)

                time = data[0, 0]
                temperature = data[0, 1]
                currents = data[1:, 6]
                volatges = data[1:, 0]
                resistances = volatges / currents

                row = [time, temperature] + [val for pair in zip(currents, volatges, resistances) for val in pair]
                new_data.loc[len(new_data)] = row
    
    # full_df = pd.concat([df, new_data], axis=1, ignore_index=False)
    full_df = pd.merge(df, new_data, left_index=True, right_index=True)

    full_df.to_excel(excel_path, index=False)
    print("Excel file has been updated successfully!")

directory_path = os.getcwd()
excel_path = os.path.join(directory_path, "output.xlsx")
duplicate_and_rename_folder(directory_path)
read_folder_and_create_excel(directory_path)
read_and_update_excel(directory_path, excel_path)