import os
import win32com.client as win32
import psutil

def kill_excel_processes():
    for process in psutil.process_iter(['pid', 'name']):
        if process.info['name'] == 'EXCEL.EXE':
            try:
                process.terminate()
                print(f"Terminated Excel process with PID {process.info['pid']}.")
            except Exception as e:
                print(f"Failed to terminate process {process.info['pid']}: {e}")


def initialize_excel():
    try:
        excel = win32.Dispatch('Excel.Application')
        excel.Visible = False
        print("Excel application initialized and visible.")
        return excel
    except Exception as e:
        print(f"Failed to initialize Excel application: {e}")
        return None


def copyExcel(templatePath, excel):
    if not os.path.exists(templatePath):
        print(f"Template file not found: {templatePath}")
        return None

    workbook = None
    try:
        # Open the workbook
        workbook = excel.Workbooks.Open(templatePath, UpdateLinks=0, ReadOnly=False)
        print(f"Workbook '{templatePath}' opened successfully.")

        # Access the 'sort_times' sheet
        sheet_name = 'sort_times'
        if not any(sheet.Name == sheet_name for sheet in workbook.Sheets):
            print(f"Sheet '{sheet_name}' not found in the workbook.")
            return None
        sheet = workbook.Sheets(sheet_name)
        print(f"Sheet '{sheet_name}' accessed successfully.")

        # Handle existing 'Temp_Aggregate' sheet
        temp_sheet_name = "Temp_Aggregate"
        if any(sheet.Name == temp_sheet_name for sheet in workbook.Sheets):
            workbook.Sheets(temp_sheet_name).Delete()
            print(f"Existing temporary sheet '{temp_sheet_name}' deleted.")

        # Create a temporary sheet for aggregation
        temp_sheet = workbook.Sheets.Add()
        temp_sheet.Name = temp_sheet_name
        print(f"Temporary sheet '{temp_sheet_name}' created for aggregation.")

        # Aggregate ranges
        print("Aggregating ranges into the temporary sheet...")

        # Local Sort Plan
        temp_sheet.Cells(1, 1).Value = "Local Sort Plan:"
        sheet.Range("A1:D4").Copy()
        temp_sheet.Range("A2").PasteSpecial(Paste=-4104)  # Paste with formatting

        # Root Cause of Delay
        temp_sheet.Cells(7, 1).Value = "Root Cause of Delay:"
        sheet.Range("A7:D12").Copy()
        temp_sheet.Range("A8").PasteSpecial(Paste=-4104)  # Paste with formatting

        # Outbound Truck Routes
        temp_sheet.Cells(15, 1).Value = "Outbound Truck Routes:"
        sheet.Range("A15:D26").Copy()
        temp_sheet.Range("A16").PasteSpecial(Paste=-4104)  # Paste with formatting

        # Copy the entire aggregated range to clipboard
        temp_sheet.UsedRange.Copy()
        print("Aggregated data copied to clipboard with formatting.")

        # Return the workbook for cleanup later
        return workbook

    except Exception as e:
        print(f"An error occurred: {e}")
        return None


def close_excel(excel, workbook):
    if workbook:
        try:
            # Delete the temporary sheet if it exists
            if any(sheet.Name == "Temp_Aggregate" for sheet in workbook.Sheets):
                workbook.Sheets("Temp_Aggregate").Delete()
                print("Temporary sheet deleted.")
            workbook.Close(SaveChanges=False)
            print("Workbook closed without saving changes.")
        except Exception as e:
            print(f"Failed to close workbook: {e}")
    if excel:
        try:
            excel.Quit()
            print("Excel application closed.")
        except Exception as quit_error:
            print(f"Error closing Excel: {quit_error}")
