# Excel Tools

![header](/docs/images/header.png)

web page:  [alxcor.github.io/cam410](https://alxcor.github.io/camxls)

## SparesWeb.xls Spare Parts Availability Analysis for MLFB

![SparesWeb.xls](/docs/images/spareweb.png)

**SparesWeb.xls:** Spare parts availability analysis.


An Excel macro-enabled workbook able to check the availability of Siemens Industry products.

Simply add product codes (MLFB) into 1st column (column A) in "Data" Worksheet and press the "Read All" or "Read Row" button in order to run the VBA code.

For each row, data is requested from Industry Mall web site and added to workbook.

The VBA code is available in a .bas file.

A ready to use Excel file is also available for download.
After download, unblock the file in order to use the macro.
Right click on the file, select 'Properties' and in the 'Properties' dialog, at 'Security' select 'Unblock'

![unblock](/docs/images/unblock.png)

Open the document with Excel and enable Macros.

A menu named Spares Web is added after Home tab, in Excel Ribbon.

![Ribbon Menu](/docs/images/sparewebmenu.png)

- "Clear All" deletes everything from "Data" worksheet.
- "Set Header" add header data and format columns in "Data" worksheet.
- "Read Row" read data from internet for the spare part code (MLFB) in column A of the selected Row in "Data" worksheet.
- "Read All" read data from internet for the spare part codes (MLFB) in column A in "Data" worksheet (up to Row 500).
- "Write Report" read data from "Data" worksheet and generates a printable report in "Report" worksheet.
- "Format Report" prepare a printable format of the data in "Report" worksheet.

To access the VBA code press ALT-F11 in Excel.

Versions:
- v1.4.4 / 01.02.2023 New version, with cmlHTTP, Ribbon menu, read data optimisations
- v1.4.3 / 09.01.2023 New version based on xmlHTTP
- v1.4.2 / 10.12.2022 Minor bug fixes
- v1.4.1 / 2010? Functional version using IE connectivity

Issues:
* [ ] [feature] ([#3][i1])

[i1]: https://github.com/alxcor/camxls/issues/3


## S5Convert.xls: Converter for Step5 / Step7 data types

![S5Convert](/docs/images/converter.png)

**S5Convert.xls:** Converter from different Step5 / Step7 data types to binary/hex and back...

An Excel workbook for converting data, compatible with Excel, LibreOffice, etc.


