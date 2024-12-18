import win32com.client as win32
import os

def copyExcel(filePath):
    # Initialize Excel COM application
    excel = win32.Dispatch('Excel.Application')
    # excel.Visible = False  # Run Excel in the background

    workbook = None
    try:
        # Open the source workbook
        workbook = excel.Workbooks.Open(filePath)
        sheet = workbook.Sheets('sort_times')  # Ensure the sheet exists

        # Create a new workbook for copying data
        new_workbook = excel.Workbooks.Add()
        new_sheet = new_workbook.Sheets(1)

        # Copy and paste ranges with formatting
        new_sheet.Cells(1, 1).Value = "Local Sort Plan:"
        sheet.Range('A1:D4').Copy()
        new_sheet.Range('A2').PasteSpecial(-4104)  # xlPasteAll

        new_sheet.Cells(7, 1).Value = "Root Cause of Delay:"
        sheet.Range('F1:I6').Copy()
        new_sheet.Range('A8').PasteSpecial(-4104)

        new_sheet.Cells(15, 1).Value = "Outbound Truck Routes:"
        sheet.Range('K1:N13').Copy()
        new_sheet.Range('A16').PasteSpecial(-4104)

        # Copy sheet to clipboard
        new_sheet.UsedRange.Copy()
        print("Copied data with formatting to clipboard. Paste it into your email.")

    except Exception as e:
        print(f"An error occurred in copyExcel: {e}")

    finally:
        # Cleanup Excel
        if workbook:
            workbook.Close(False)
        excel.Quit()
        del excel
