# SparesWeb_1_4_3.xls
Check the availability of spare parts using a list of Siemens product code numbers (MLFB).
All data are read from Industry Mall web page (no login needed).

Simply add product codes (MLFB) into 2-nd column (column B) and press the button in order to run the VBA code.
For each row, data is requested from Industry Mall web site and added to workbook.

Use "Read Row" to read data only for current Row or "Read All" to read all Rows.
Excel macros are using only the xmlHTTP version.
If an older version, using IE connectivity is needed, edit Sub EvaluateRow() and Sub EvaluateAll() and change the value for netMode variable:
'netMode: 0=Internet; 1=Intranet; 2=xmlHTTP version

In version 1.4.3 the line is set to:
netMode = 2 'Use 2=xmlHTTP version

# S5Convert.xls
Converter from different Step5 / Step7 data types to binary/hex and back...<br><br>
An Excel workbook for converting data, compatible with Excel, LibreOffice, etc.<br><br>
