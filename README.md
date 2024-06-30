# Excel-Automation

Step-by-step explanation of the script:

1. openpyxl: This is the main library used for reading from and writing to Excel files.

2. Font and PatternFill from openpyxl.styles: These are used for formatting cells (e.g., setting font properties and background colors).

3. load_workbook('sample.xlsx'): Opens the Excel file named sample.xlsx.

4. wb.active: Selects the active sheet in the workbook

5. iter_rows: Iterates over the rows in the specified range (from row 1 to 5 and column 1 to 3).

6. values_only=True: Returns the values of the cells, not the cell objects.

7. sheet.append(row): Appends each row of data to the sheet.

8. Font(bold=True, color="FFFFFF"): Creates a font object with bold text and white color.

9. PatternFill(start_color="000000", end_color="000000", fill_type="solid"): Creates a fill object with a black background.

10. sheet["1:1"]: Selects the first row (header row).

11. sum(ages) / len(ages): Calculates the average age.

12. sheet["E1"] = "Average Age": Writes the label "Average Age" to cell E1.

13. wb.save('sample.xlsx'): Saves the changes made to the Excel file.
