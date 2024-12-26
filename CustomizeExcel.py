# from openpyxl.styles import Border, Side, PatternFill

# Border Control
def addBorder(sheet, range):
    try:
        # Get the range object
        rng = sheet.Range(range)
        
        # Set the border styles
        medium_style = 1  # xlContinuous
        medium_weight = 2  # xlMedium

        # Iterate over each cell in the range
        for cell in rng:
            for border_id in [7, 8, 9, 10]:
                border = cell.Borders(border_id)
                border.LineStyle = medium_style
                border.Weight = medium_weight

        print(f"Borders applied to every cell in range {range}.")
    except Exception as e:
        print(f"Failed to apply borders to range {range}: {e}")

# Portion Control
def adjustCol(sheet, col):
    maxLen = 0

    # find longest cell to base expansion
    for cell in sheet[col]:
        if cell.value:
            maxLen = max(maxLen, len(str(cell.value)))
    
    adjustedWid = maxLen + 1
    sheet.column_dimensions[col].width = adjustedWid

# Color Control
# def changeRowColor(sheet, rowNum, desiredColor):
#     # Yellow fill
#     yellowFill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

#     # Green fill
#     greenFill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")

#     for cell in sheet[rowNum]:
#         if desiredColor == 'green':
#             cell.fill = greenFill
#         elif desiredColor == 'yellow':
#             cell.fill = yellowFill

