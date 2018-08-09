# TallyRange

A VBA function to tally the unique values in a range. Because Excel has no simple means to count unique elements of a range where not all contents of the range are known previously, Tally Range is necessary. COUNTIF cannot be used unless all unique elements of a range are listed in a separate location, which can be tedious when the data isn't known.


<b> How to Import: </b> 
1. Download (or Save As) the TallyRangeVBA.bas file to any location on your computer.
2. Open the Excel file you wish to use.
3. Press Alt and F11 on your keyboard to open the Visual Basic Window
4. Go to File -> Import
5. Find the TallyRangeVBA.bas file you downloaded previously, click Open.
6. The TallyRange() function is now added to your Excel file. You may close the Visual Basic window and return to your spreadsheet.
7. To use the function, simply type "=TallyRange(" into a cell formula and specify a range to tally.

Important Note: Now that your spreadsheet contains VBA code that you have imported, you must save your Excel file as .xlsm. This is a "Macro-Enabled Workbook" file that saves the script code along with the spreadsheet. If you save your file as a standard .xls or .xlsx type the TallyRange function will not be included and cannot be used in the future without importing again.
