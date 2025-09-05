

FILE_PATH = "C:/Users/raghunath_b/PycharmProjects/Web_Automation_Framework/Sample Book.xlsx"
SHEET_IDENTIFIER = "Sample sheet 1"
HEADER_SUBSTRING = "Email"

def get_data_as_string(FILE_PATH: str, SHEET_IDENTIFIER: str | int, HEADER_SUBSTRING: str) -> str:
    from utilities.ExcelReader import ExcelReader
    ExcelReader = ExcelReader()
    workbook = ExcelReader.get_work_book(FILE_PATH)
    if not workbook:
        print(f"Error: Could not load workbook from '{FILE_PATH}'")
    else:
        sheet = None
        if isinstance(SHEET_IDENTIFIER, str):
            sheet = ExcelReader.get_sheet_by_name(FILE_PATH, SHEET_IDENTIFIER)
        elif isinstance(SHEET_IDENTIFIER, int):
            sheet = ExcelReader.get_sheet_by_index(workbook, SHEET_IDENTIFIER)
        else:
            print("Error: SHEET_IDENTIFIER must be a string (sheet name) or an integer (sheet index).")

        if not sheet:
            print(f"Error: Sheet '{SHEET_IDENTIFIER}' not found in '{FILE_PATH}'")
        else:
            all_data_rows = ExcelReader.read_sheet(sheet)

            if not all_data_rows:
                print(f"No data found in sheet '{SHEET_IDENTIFIER}'.")
            else:
                first_data_row = all_data_rows[0]

                found_header_key = None
                for header_key in first_data_row.keys():
                    if HEADER_SUBSTRING.lower() in header_key.lower():
                        found_header_key = header_key
                        break

                if found_header_key:
                    name_data = first_data_row[found_header_key]
                    return name_data
                else:
                    print(f"Header containing '{HEADER_SUBSTRING}' not found in the sheet.")

