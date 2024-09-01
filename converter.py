import pandas as pd
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
        'Set Temperature': temperatures,
        'Set Current': currents
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
directory_path = os.getcwd()
# read_folder_and_create_excel(directory_path)
duplicate_and_rename_folder(directory_path)