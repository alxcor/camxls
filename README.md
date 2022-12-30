# SparesWeb_1_4_2.xls
Check the availability of spare parts using a list of Siemens product code numbers (MLFB).
All data are read from Industry Mall web page (no login needed).

Simply add product codes (MLFB) into 2-nd column (column B) and press the button in order to run the VBA code.
For each row, data is requested from Industry Mall web site and added to workbook.

# S5Convert.xls
Converter from different Step5 / Step7 data types to binary/hex and back...<br><br>
An Excel workbook for converting data, compatible with Excel, LibreOffice, etc.<br><br>
<table>
<tr><th>From</th><th colspan=7>To</th></tr>
<tr><td>Bin </td><td> - </td><td>Hex</td><td>Uint</td><td>Int</td><td>BCD</td><td>S5Timer</td><td>Real</td></tr>
<tr><td>Hex </td><td>Bin</td><td> - </td><td>Uint</td><td>Int</td><td>BCD</td><td>S5Timer</td><td>Real</td></tr>
<tr><td>Hex </td><td>Bin</td><td>Hex</td><td>Uint</td><td>Int</td><td>BCD</td><td>S5Timer</td><td>Real</td></tr>
<tr><td>Uint</td><td>Bin</td><td>Hex</td><td> -  </td><td>Int</td><td>BCD</td><td>S5Timer</td><td>Real</td></tr>
<tr><td>Int </td><td>Bin</td><td>Hex</td><td>Uint</td><td> - </td><td>BCD</td><td>S5Timer</td><td>Real</td></tr>
<tr><td>BCD </td><td>Bin</td><td>Hex</td><td>Uint</td><td>Int</td><td> - </td><td>S5Timer</td><td>Real</td></tr>
<tr><td>S5T </td><td>Bin</td><td>Hex</td><td>Uint</td><td>Int</td><td>BCD</td><td> -     </td><td>Real</td></tr>
<tr><td>Real</td><td>Bin</td><td>Hex</td><td>Uint</td><td>Int</td><td>BCD</td><td>S5Timer</td><td> -  </td></tr>
<tr><td>...</td><td colspan=7>...</td></tr>
</table>
