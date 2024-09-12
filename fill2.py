import pandas as pd
from openpyxl import load_workbook

def fill_sheet2(directory, excel_name, data, mode='a'):
    path = os.path.join(directory, excel_name)
    # Create a DataFrame from the data dictionary
    df = pd.DataFrame([data])

    # Load the workbook and specify the engine to maintain compatibility
    book = load_workbook(path)
    writer = pd.ExcelWriter(path, engine='openpyxl')  # Define the writer
    writer.book = book
    writer.sheets = {ws.title: ws for ws in book.worksheets}

    # Write the data to 'Sheet2'; if 'Sheet2' doesn't exist, it will create it
    if 'Sheet2' not in writer.sheets:
        df.to_excel(writer, sheet_name='Sheet2', index=False)
    else:
        # Read existing data
        existing_df = pd.read_excel(path, sheet_name='Sheet2')
        # Concatenate new data with the existing data
        new_df = pd.concat([existing_df, df], ignore_index=True)
        # Write back to the Excel file
        new_df.to_excel(writer, sheet_name='Sheet2', index=False)

    # Save the changes
    writer.save()
    writer.close()
