import openpyxl
from datetime import datetime, time
import os

def parse_time(value):
    if isinstance(value, (datetime, time)):
        return value
    if isinstance(value, (int, float)):
        hours = int(value)
        if 0 <= hours <= 23:
            return time(hours, 0)
    if isinstance(value, str):
        try:
            hours = int(value.strip())
            if 0 <= hours <= 23:
                return time(hours, 0)
        except ValueError:
            pass
    return None

def validate_excel_file(file_path):
    if not os.path.exists(file_path):
        print(f"Error: File not found at {file_path}")
        return

    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active

    mandatory_columns = {
        'MIPS': float,
        'Date': (str, datetime),
        'Start Time': (int, float, str, time, datetime),
        'End time': (int, float, str, time, datetime),
        'Project Name': str
    }

    # Find column indices
    column_indices = {}
    header_row = sheet[3]
    for col in header_row:
        if col.value:
            normalized_col_name = col.value.strip().lower()
            if normalized_col_name in [col.lower() for col in mandatory_columns]:
                column_indices[normalized_col_name] = col.column

    if len(column_indices) != len(mandatory_columns):
        print("Error: Not all mandatory columns are present in the sheet.")
        return

    errors_found = False

    for row_index in range(4, 9):  # Check rows 4 to 8
        row = sheet[row_index]
        row_values = [row[col_index - 1].value for col_index in column_indices.values()]
        
        # Check if row has any data in mandatory columns
        if not any(value is not None and str(value).strip() != '' for value in row_values):
            continue  # Skip completely empty rows

        for col_name, col_index in column_indices.items():
            value = row[col_index - 1].value

            if value is None or (isinstance(value, str) and value.strip() == ''):
                print(f"Error in row {row_index}: {col_name.capitalize()} is empty.")
                errors_found = True
                continue

            if col_name == 'mips':
                try:
                    float(value)
                except ValueError:
                    print(f"Error in row {row_index}: MIPS '{value}' is not a valid number.")
                    errors_found = True

            elif col_name == 'date':
                if not isinstance(value, datetime):
                    try:
                        datetime.strptime(str(value), '%Y-%m-%d')
                    except ValueError:
                        print(f"Error in row {row_index}: Invalid date format '{value}'. Use YYYY-MM-DD.")
                        errors_found = True

            elif col_name in ['start time', 'end time']:
                parsed_time = parse_time(value)
                if parsed_time is None:
                    print(f"Error in row {row_index}: Invalid time format '{value}' for {col_name.capitalize()}. Use whole numbers from 0 to 23.")
                    errors_found = True

            elif col_name == 'project name':
                if not isinstance(value, str) or value.strip() == '':
                    print(f"Error in row {row_index}: Project Name '{value}' is not a valid string.")
                    errors_found = True

    if not errors_found:
        print("Validation complete. No errors found.")
    else:
        print("Validation complete. Errors were found. Please check the comments above for details.")

def process_excel_file(file_path):
    validate_excel_file(file_path)

# Specify the local path to the Excel file
file_path = "C:\\Users\\Admin\\Desktop\\python\\Weekday_Slot_Request.xlsx"

# Process the local Excel file
process_excel_file(file_path)