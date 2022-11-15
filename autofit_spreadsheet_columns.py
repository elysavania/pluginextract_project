"""
2022-11-15 NOJ:
A helper function using openpyxl to autofit all columns in a given spreadsheet.
The helper accepts a filename, iterates through each column in each worksheet to determine the length of the longest cell in the column, and sets the column with based off that length.
"""
from openpyxl import load_workbook


def autofit_spreadsheet_columns(spreadsheet_filename):
    workbook = load_workbook(filename=spreadsheet_filename)
    for worksheet in workbook:
        # 2022-11-15 NOJ: Source https://stackoverflow.com/a/39530676
        for column in worksheet.columns:
            column_name = column[0].column_letter  # Get the column name
            max_length = 0
            for cell in column:
                try:  # Necessary to avoid error on empty cells
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            # 2022-11-15 NOJ: A scaling factor and fixed padding (both determined empricially) is required to increase cell width because we are using a font that is thinner than monospace.
            adjusted_width = (max_length + 4) * 1.15
            worksheet.column_dimensions[column_name].width = adjusted_width

    workbook.save(filename=spreadsheet_filename)
