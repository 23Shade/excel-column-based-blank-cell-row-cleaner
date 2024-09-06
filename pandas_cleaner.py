# Removes rows with blank cells in specified columns from an Excel file.
# Utilizes the pandas library for efficient data processing.
# Note: This script does not preserve original formatting (e.g., fonts, colors).
# It can handle multiple columns for blank cell checks.

import pandas as pd
import os

def validate_input_file(input_file, sheet_name):
    """Validates if the input file exists and the sheet name is correct."""
    if not os.path.isfile(input_file):
        print(f"Input file not found: {input_file}!")
        return False
    try:
        pd.read_excel(input_file, sheet_name=sheet_name)
    except ValueError:
        print(f"Sheet not found: {sheet_name}!")
        return False
    return True

def clean_data(input_file, sheet_name, target_columns, output_file):
    """Removes rows with blank cells in specified columns and saves the cleaned data."""
    try:
        # Read the data into a DataFrame
        df = pd.read_excel(input_file, sheet_name=sheet_name)
    except Exception as e:
        print(f"Error reading the Excel file: {e}")
        return
    
    # Check if the target columns exist in the DataFrame
    missing_columns = [col for col in target_columns if col not in df.columns]
    if missing_columns:
        print(f"Target columns not found: {', '.join(missing_columns)}!")
        return
    
    # Remove rows where any target column is missing or NaN
    df_cleaned = df.dropna(subset=target_columns)
    
    output_dir = os.path.dirname(output_file)
    if not os.path.exists(output_dir):
        try:
            os.makedirs(output_dir)
        except OSError as e:
            print(f"Error creating directory {output_dir}: {e}")
            return
    
    # Handle case where the output file already exists
    if os.path.isfile(output_file):
        print(f"Output file already exists: {output_file}.")
        while True:
            response = input("Do you want to overwrite it, create a new file, or cancel? (overwrite/new/cancel): ").strip().lower()
            if response == 'overwrite':
                break
            elif response == 'new':
                base, ext = os.path.splitext(output_file)
                counter = 1
                new_output_file = f"{base}_{counter}{ext}"
                while os.path.isfile(new_output_file):
                    counter += 1
                    new_output_file = f"{base}_{counter}{ext}"
                output_file = new_output_file
                break
            elif response == 'cancel':
                print("Process canceled!")
                return
            else:
                print("Invalid response. Please enter 'overwrite', 'new', or 'cancel'.")
    
    # Save the cleaned DataFrame back to an Excel file
    try:
        df_cleaned.to_excel(output_file, index=False)
        formatted_columns = ', '.join([f'"{col}"' for col in target_columns])
        print(f"Rows with missing values have been removed from the columns: {formatted_columns}.")
        print(f"Cleaned file saved as: {output_file}.")
    except Exception as e:
        print(f"Error saving file: {e}")

def main():
    # --- UPDATE INPUT FILE PATH ---
    input_file_path = r'C:\Users\your_user\Desktop\your_file.xlsx'  # Update with your input file path
    # --- UPDATE SHEET NAME ---
    sheet_name = 'your_sheet_name'  # Update with your sheet name
    # --- UPDATE TARGET COLUMNS ---
    target_columns = ['your_column_name']  # Update with your column names, you can add more columns as needed
    # --- UPDATE OUTPUT FILE PATH ---
    output_file_path = r'C:\Users\your_user\Desktop\output_file.xlsx'  # Update with your desired output file path
    
    if validate_input_file(input_file_path, sheet_name):
        clean_data(input_file_path, sheet_name, target_columns, output_file_path)

if __name__ == "__main__":
    main()
