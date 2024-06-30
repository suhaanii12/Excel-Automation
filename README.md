# Excel-Automation

Step-by-step explanation of the script:

openpyxl: This is the main library used for reading from and writing to Excel files.

Font and PatternFill from openpyxl.styles: These are used for formatting cells (e.g., setting font properties and background colors).

load_workbook('sample.xlsx'): Opens the Excel file named sample.xlsx.

wb.active: Selects the active sheet in the workbook

iter_rows: Iterates over the rows in the specified range (from row 1 to 5 and column 1 to 3).

values_only=True: Returns the values of the cells, not the cell objects.

data: A list of lists containing the data to be written to the Excel file.

sheet.append(row): Appends each row of data to the sheet.

Font(bold=True, color="FFFFFF"): Creates a font object with bold text and white color.

PatternFill(start_color="000000", end_color="000000", fill_type="solid"): Creates a fill object with a black background.

sheet["1:1"]: Selects the first row (header row).

ages: A list comprehension that extracts the age values from rows 2 to 4, column 2.

sum(ages) / len(ages): Calculates the average age.

sheet["E1"] = "Average Age": Writes the label "Average Age" to cell E1.

sheet["E2"] = average_age: Writes the calculated average age to cell E2.

wb.save('sample.xlsx'): Saves the changes made to the Excel file.

print("Data written and formatted successfully."): Prints a confirmation message.
