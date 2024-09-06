# Removes rows with blank cells in specified columns from an Excel file.
# Utilizes the openpyxl library, preserving original formatting (e.g., fonts, colors).
# Ensures the visual presentation of the Excel file remains unchanged.
# It can handle multiple columns for blank cell checks.

import openpyxl
import os

def validate_input_file(input_file, sheet_name):
    """Validates if the input file exists and the sheet name is correct."""
    if not os.path.isfile(input_file):
        print(f"Input file not found: {input_file}!")
        return None
    try:
        workbook = openpyxl.load_workbook(input_file)
        sheet = workbook[sheet_name]
        return sheet
    except KeyError:
        print(f"Sheet not found: {sheet_name}!")
        return None

def find_target_columns(sheet, target_columns):
    """Finds and returns the indices of the target columns."""
    target_column_indices = {}
    for cell in sheet[1]:  # Assuming the first row is the header row
        if cell.value in target_columns:
            target_column_indices[cell.value] = cell.column
    
    missing_columns = [col for col in target_columns if col not in target_column_indices]
    if missing_columns:
        print(f"Target columns not found: {', '.join(missing_columns)}!")
        return None
    return target_column_indices

def clean_data(sheet, target_column_indices):
    """Removes rows with blank cells in the specified columns."""
    rows_to_delete = []
    for row in sheet.iter_rows(min_row=2):  # Start from the second row to skip the header
        for col_name, col_index in target_column_indices.items():
            cell = row[col_index - 1]  # Adjust for 0-based indexing
            if not cell.value:  # If the cell is empty
                rows_to_delete.append(row[0].row)
                break  # No need to check other columns for this row

    for row in reversed(rows_to_delete):  # Delete rows from bottom to top
        sheet.delete_rows(row)

def save_workbook(workbook, output_file_path):
    """Saves the workbook to the specified file path, handling overwrites and new file creation."""
    output_dir = os.path.dirname(output_file_path)
    
    if not os.path.exists(output_dir):
        try:
            os.makedirs(output_dir)
        except OSError as e:
            print(f"Error creating directory {output_dir}: {e}")
            return False
    
    if os.path.isfile(output_file_path):
        print(f"Output file already exists: {output_file_path}.")
        while True:
            response = input("Do you want to overwrite it, create a new file, or cancel? (overwrite/new/cancel): ").strip().lower()
            if response == 'overwrite':
                break
            elif response == 'new':
                base, ext = os.path.splitext(output_file_path)
                counter = 1
                new_output_file_path = f"{base}_{counter}{ext}"
                while os.path.isfile(new_output_file_path):
                    counter += 1
                    new_output_file_path = f"{base}_{counter}{ext}"
                output_file_path = new_output_file_path
                break
            elif response == 'cancel':
                print("Process canceled!")
                return False
            else:
                print("Invalid response. Please enter 'overwrite', 'new', or 'cancel'.")
    
    try:
        workbook.save(output_file_path)
        return True
    except Exception as e:
        print(f"Error saving file: {e}")
        return False

def main():
    # --- UPDATE INPUT FILE PATH ---
    input_file_path = r'C:\Users\Shade\Desktop\Email.xlsx'  # Update with your input file path
    # --- UPDATE SHEET NAME ---
    sheet_name = 'NEW COPY(in)'  # Update with your sheet name
    # --- UPDATE TARGET COLUMNS ---
    target_columns = ['Email']  # Update with your column names, you can add more columns as needed
    # --- UPDATE OUTPUT FILE PATH ---
    output_file_path = r'C:\Users\Shade\Desktop\output_file.xlsx'  # Update with your desired output file path

    sheet = validate_input_file(input_file_path, sheet_name)
    if sheet is None:
        return

    target_column_indices = find_target_columns(sheet, target_columns)
    if target_column_indices is None:
        return

    clean_data(sheet, target_column_indices)

    if save_workbook(sheet.parent, output_file_path):
        formatted_columns = ', '.join(f'"{col}"' for col in target_columns)
        print(f"Rows with missing values have been removed from the columns: {formatted_columns}.")
        print(f"Cleaned file saved as: {output_file_path}.")

if __name__ == "__main__":
    main()
