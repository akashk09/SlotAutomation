import os
from datetime import datetime, time

from openpyxl import load_workbook as loadwb

def parse_time(value):
    if isinstance(value, (datetime, time)):
        return value.time() if isinstance(value, datetime) else value
    if isinstance(value, (int, float)):
        hours = int(value)
        return time(hours, 0) if 0 <= hours <= 23 else None
    if isinstance(value, str):
        try:
            hours = int(value.strip())
            return time(hours, 0) if 0 <= hours <= 23 else None
        except ValueError:
            return None
    return None

def validate_excel_file(file_path):
    if not os.path.exists(file_path):
        print(f"Error: File not found at {file_path}")
        return

    workbook = loadwb(filename=file_path)
    sheet = workbook.sheets[0]  # Assume we're working with the first sheet

    mandatory_columns = {
        'MIPS': float,
        'Date': datetime,
        'Start Time': time,
        'End time': time,
        'Project Name': str
    }

    # Find column indices
    header_row = sheet.range('A3:Z3').value
    column_indices = {col.lower(): i for i, col in enumerate(header_row) if col}

    missing_columns = set(mandatory_columns.keys()) - set(column_indices.keys())
    if missing_columns:
        print(f"Error: Missing mandatory columns: {', '.join(missing_columns)}")
        return

    errors_found = False
    for row_index in range(3, 8):  # Check rows 4 to 8 (0-based index)
        row_values = sheet.range(f'A{row_index+1}:Z{row_index+1}').value
        
        # Skip completely empty rows
        if not any(row_values):
            continue

        for col_name, expected_type in mandatory_columns.items():
            value = row_values[column_indices[col_name.lower()]]
            
            if value is None or (isinstance(value, str) and value.strip() == ''):
                print(f"Error in row {row_index+1}: {col_name} is empty.")
                errors_found = True
                continue

            if col_name == 'MIPS':
                try:
                    float(value)
                except ValueError:
                    print(f"Error in row {row_index+1}: MIPS '{value}' is not a valid number.")
                    errors_found = True
            
            elif col_name == 'Date':
                if not isinstance(value, datetime):
                    try:
                        datetime.strptime(str(value), '%Y-%m-%d')
                    except ValueError:
                        print(f"Error in row {row_index+1}: Invalid date format '{value}'. Use YYYY-MM-DD.")
                        errors_found = True
            
            elif col_name in ['Start Time', 'End time']:
                parsed_time = parse_time(value)
                if parsed_time is None:
                    print(f"Error in row {row_index+1}: Invalid time format '{value}' for {col_name}. Use whole numbers from 0 to 23.")
                    errors_found = True
            
            elif col_name == 'Project Name':
                if not isinstance(value, str) or value.strip() == '':
                    print(f"Error in row {row_index+1}: Project Name '{value}' is not a valid string.")
                    errors_found = True

    if not errors_found:
        print("Validation complete. No errors found.")
    else:
        print("Validation complete. Errors were found. Please check the comments above for details.")

def process_excel_file(file_path):
    validate_excel_file(file_path)

# Specify the local path to the Excel file
file_path = r"C:\Users\Admin\Desktop\python\Weekday_Slot_Request.xlsx"

# Process the local Excel file
process_excel_file(file_path)