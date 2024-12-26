import os
import win32com.client as win32

def initialize_excel():
    """
    Initializes and opens the Excel application.
    Returns the Excel application instance.
    """
    try:
        excel = win32.Dispatch('Excel.Application')
        excel.Visible = False  # Keep Excel visible
        print("Excel application initialized and visible.")
        return excel
    except Exception as e:
        print(f"Failed to initialize Excel application: {e}")
        return None

def copyExcel(templatePath, excel):
    """
    Copies specific data ranges from the 'sort_times' sheet in the template Excel file
    directly to the clipboard, retaining all formatting (e.g., borders, fonts, colors).
    """
    if not os.path.exists(templatePath):
        print(f"Template file not found: {templatePath}")
        return

    workbook = None
    try:
        # Open the workbook
        workbook = excel.Workbooks.Open(templatePath, UpdateLinks=0, ReadOnly=True)
        print(f"Workbook '{templatePath}' opened successfully.")

        # Access the 'sort_times' sheet
        sheet_name = 'sort_times'
        if not any(sheet.Name == sheet_name for sheet in workbook.Sheets):
            print(f"Sheet '{sheet_name}' not found in the workbook.")
            return
        sheet = workbook.Sheets(sheet_name)
        print(f"Sheet '{sheet_name}' accessed successfully.")

        # Copy specific ranges directly to clipboard
        print("Copying ranges with formatting...")
        sheet.Range("A1:D4").Copy()  # Local Sort Plan
        print("Local Sort Plan copied.")
        sheet.Range("F1:I6").Copy()  # Root Cause of Delay
        print("Root Cause of Delay copied.")
        sheet.Range("K1:N13").Copy()  # Outbound Truck Routes
        print("Outbound Truck Routes copied.")

        # Excel remains open until explicitly closed by the program
        print("Data copied to clipboard. Excel remains open.")
    except Exception as e:
        print(f"An error occurred: {e}")
    # finally:
    #     if workbook:
    #         try:
    #             workbook.Close(False)
    #             print("Original workbook closed.")
    #         except Exception as e:
    #             print(f"Failed to close workbook: {e}")

def close_excel(excel):
    """
    Closes the Excel application if it is running.
    """
    if excel:
        try:
            excel.Quit()
            print("Excel application closed.")
        except Exception as quit_error:
            print(f"Error closing Excel: {quit_error}")
