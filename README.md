This Python script processes an Excel workbook named 'transactions.xlsx' and performs the following steps:

1. **Load the Workbook**: The script uses the `openpyxl` library to load the workbook. The active sheet is assumed to be 'Sheet1'.

2. **Process Each Row**: 
   - The script iterates over each row starting from the second row (index 2) to the last row in the sheet.
   - For each row, it extracts the value from column 3 and multiplies it by 0.9 to calculate a corrected price.

3. **Create a New Column**: 
   - It adds a new column to the same row where the corrected price is stored.

4. **Add a Bar Chart**:
   - The script creates a bar chart using `openpyxl.chart.BarChart`.
   - It references the data for the chart from columns 4 and above, starting from row 2.
   - The chart is added to the workbook at cell 'e2'.

5. **Save the Workbook**: Finally, it saves the modified workbook with the same filename as before.

The script outputs a new Excel file named 'transactions2.xlsx' that includes the original data along with the corrected prices and the newly created bar chart. This process allows for easy visualization of changes to financial data in an Excel spreadsheet.