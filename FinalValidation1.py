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
    sheet = workbook.active  # Get the active sheet

    # Define the mandatory columns with their expected data types
    mandatory_columns = {
        'MIPS': float,
        'Date': datetime,
        'Start Time': time,
        'End time': time,
        'Project Name': str
    }

    # Extract the header row (assuming headers are on row 3)
    header_row = [cell.value.strip() if cell.value else '' for cell in sheet[3]]  # Strip spaces and handle None
    header_row_lower = [header.lower() for header in header_row]  # Convert all headers to lowercase for comparison

    # Create a dictionary to find the column indices
    column_indices = {header: index for index, header in enumerate(header_row_lower) if header}

    # Check for missing mandatory columns
    missing_columns = set([col.lower() for col in mandatory_columns.keys()]) - set(column_indices.keys())
    if missing_columns:
        print(f"Error: Missing mandatory columns: {', '.join(missing_columns)}")
        return

    errors_found = False
    all_rows_empty = True  # Flag to check if all rows are empty

    # Iterate over rows 4 to 8
    for row in sheet.iter_rows(min_row=4, max_row=8):  # Check rows 4 to 8 (1-based index)
        row_values = [cell.value for cell in row]

        # Check if all mandatory columns in the row are empty
        if all(
            row_values[column_indices[col_name.lower()]] is None or 
            (isinstance(row_values[column_indices[col_name.lower()]], str) and row_values[column_indices[col_name.lower()]].strip() == '')
            for col_name in mandatory_columns.keys()
        ):
            # Skip the row if all mandatory columns are empty
            continue

        all_rows_empty = False  # Found at least one non-empty row

        # Only validate rows with at least one filled mandatory column
        for col_name, expected_type in mandatory_columns.items():
            col_index = column_indices[col_name.lower()]  # Get index using lowercase column name
            value = row_values[col_index]

            if value is None or (isinstance(value, str) and value.strip() == ''):
                print(f"Error in row {row[0].row}: {col_name} is empty.")
                errors_found = True
                continue

            if col_name == 'MIPS':
                try:
                    float(value)
                except ValueError:
                    print(f"Error in row {row[0].row}: MIPS '{value}' is not a valid number.")
                    errors_found = True

            elif col_name == 'Date':
                if not isinstance(value, datetime):
                    try:
                        datetime.strptime(str(value), '%Y-%m-%d')
                    except ValueError:
                        print(f"Error in row {row[0].row}: Invalid date format '{value}'. Use YYYY-MM-DD.")
                        errors_found = True

            elif col_name in ['Start Time', 'End time']:
                parsed_time = parse_time(value)
                if parsed_time is None:
                    print(f"Error in row {row[0].row}: Invalid time format '{value}' for {col_name}. Use whole numbers from 0 to 23.")
                    errors_found = True

            elif col_name == 'Project Name':
                if not isinstance(value, str) or value.strip() == '':
                    print(f"Error in row {row[0].row}: Project Name '{value}' is not a valid string.")
                    errors_found = True

    if all_rows_empty:
        print("The sheet is empty.")
    elif not errors_found:
        print("Validation complete. No errors found.")
    else:
        print("Validation complete. Errors were found. Please check the comments above for details.")

def process_excel_file(file_path):
    validate_excel_file(file_path)

# Specify the local path to the Excel file
file_path = r"C:\Users\Admin\Desktop\python\Weekday_Slot_Request.xlsx"

# Process the local Excel file
process_excel_file(file_path)
