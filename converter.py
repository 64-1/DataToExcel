import pandas as pd
import numpy as np
import shutil
import os

def read_folder_and_create_excel(directory):
    # Put all the folder names in the directory into a list
    folders = [folder for folder in os.listdir(directory) if os.path.isdir(os.path.join(directory, folder)) and '_duplicates' in folder]

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
            
        filenumber = len([name for name in os.listdir('.') if os.path.isfile(name)])

        df = pd.DataFrame({
            'Set Temperature (C)': temperatures * filenumber,
            'Set Current (A)': currents * filenumber
        })

        excel_path = os.path.join(directory, folder, f"{folder}.xlsx")

        # Write to excel file created
        if not os.path.exists(excel_path):
            df.to_excel(excel_path, index=False)
            print("Excel file has been created successfully!")
        else:
            print("Excel file already exists.")

def duplicate_and_rename_folder(directory):
    folders = [folder for folder in os.listdir(directory) if os.path.isdir(os.path.join(directory, folder)) and 'A' in folder and 'C' in folder and '_duplicates' not in folder]
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
        
        print(f"Files from {folders[0]} have been duplicated and converted where necessary.")

def read_and_update_excel(directory):
    folders = [folder for folder in os.listdir(directory) if os.path.isdir(os.path.join(directory, folder)) and '_duplicates' in folder]
    for folder in folders:
        folder_path = os.path.join(directory, folder)
        excel_path = os.path.join(folder_path, f"{folder[:-11]}.xlsx")

        if os.path.exists(excel_path):
            df = pd.read_excel(excel_path)
        else:
            print(f"Starting a new excel file for {folder[:-11]}.")
            df = pd.DataFrame()

        columns = ['time (s)']
        measurements = ['actual temperature (C)']
        for ch in range(1, 7):
            measurements +=[f'CH{ch} actual current (A)', f'CH{ch} actual Voltage', f'CH{ch} resistance (ohm)']
        columns.extend(measurements)

        new_data = pd.DataFrame(columns=columns)

        for filename in os.listdir(folder):
            if filename.endswith('.txt'):
                file_path = os.path.join(folder, folder, filename)
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
        
        if not df.empty:
            # full_df = pd.concat([df, new_data], axis=1, ignore_index=False)
            full_df = pd.merge(df, new_data, left_index=True, right_index=True, how='outer')
            full_df.to_excel(excel_path, index=False)
            print("Excel file has been updated successfully!")
        else:
            new_data.to_excel(excel_path, index=False)
            print(f"Excel file {excel_path} has been created and data added successfully!")
        
        shutil.rmtree(os.path.join(directory, folder))
        print(f"Duplicated folder {folder} has been deleted")

def main():
    directory_path = os.getcwd()
    duplicate_and_rename_folder(directory_path)
    read_folder_and_create_excel(directory_path)
    read_and_update_excel(directory_path)

if __name__ == "__main__":
    main()
