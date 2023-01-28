# Excel Tools

![header](/docs/images/header.png)

web page:  [alxcor.github.io/cam410](https://alxcor.github.io/camxls)

## SparesWeb.xls Spare Parts Availability Analysis for MLFB

![SparesWeb.xls](/docs/images/spareweb.png)

**SparesWeb.xls:** Spare parts availability analysis.


An Excel macro-enabled workbook able to check the availability of Siemens Industry products.

Simply add product codes (MLFB) into 2-nd column (column B) and press the button in order to run the VBA code.

For each row, data is requested from Industry Mall web site and added to workbook.

The VBA code is available in a .bas file.

A ready to use Excel file is also available for download.
After download, unblock the file in order to use the macro.
Right click on the file, select 'Properties' and in the 'Properties' dialog, at 'Security' select 'Unblock'

![unblock](/docs/images/unblock.png)

Use "Read Row" to read data only for current Row or "Read All" to read all Rows.
Excel macros are using only the xmlHTTP version.
If an older version, using IE connectivity is needed, edit Sub EvaluateRow() and Sub EvaluateAll() and change the value for netMode variable:
'netMode: 0=Internet; 1=Intranet; 2=xmlHTTP version

In version 1.4.3 the line is set to:
netMode = 2 'Use 2=xmlHTTP version

Last version:
- v1.4.3 / 09.01.2023 New version based on xmlHTTP

Previous versions:
- v1.4.2 / 10.12.2022 Minor bug fixes
- v1.4.1 / 2010? Functional version using IE connectivity


## S5Convert.xls: Converter for Step5 / Step7 data types

![S5Convert](/docs/images/converter.png)

**S5Convert.xls:** Converter from different Step5 / Step7 data types to binary/hex and back...

An Excel workbook for converting data, compatible with Excel, LibreOffice, etc.


