Attribute VB_Name = "Module1"
'v1.4.5 / 15.05.2024 [alxcor:240515]

Sub conClearAll(control As IRibbonControl)
    If (ActiveSheet.Name = "Data") Then
        clearAll
    ElseIf (ActiveSheet.Name = "Report") Then
        clearAllR
    Else
        MsgBox ("'Clear All' should be called from 'Data' sheet or from 'Report' sheet")
    End If
End Sub

Sub conSetHeader(control As IRibbonControl)
    If (ActiveSheet.Name = "Data") Then
        setHeader
    ElseIf (ActiveSheet.Name = "Report") Then
        setHeaderR
    Else
        MsgBox ("'Set Header' should be called from 'Data' sheet or from 'Report' sheet")
    End If
End Sub
       
Sub conReadRow(control As IRibbonControl)
    If (ActiveSheet.Name = "Data") Then
        readRow
        MsgBox "Done!"
    ElseIf (ActiveSheet.Name = "Report") Then
        readRowR
        MsgBox "Done!"
    Else
        MsgBox ("'Read Row' should be called from 'Data' sheet or from 'Report' sheet")
    End If
End Sub

Sub conReadAll(control As IRibbonControl)
    If (ActiveSheet.Name = "Data") Then
        readAll
        MsgBox "Done!"
    ElseIf (ActiveSheet.Name = "Report") Then
        readAllR
        MsgBox "Done!"
    Else
        MsgBox ("'Read All' should be called from 'Data' sheet or from 'Report' sheet")
    End If
End Sub
    
Sub conSuccessorR(control As IRibbonControl)
    If (ActiveSheet.Name = "Data") Then
        checkSuccessor
        MsgBox "Done!"
    ElseIf (ActiveSheet.Name = "Report") Then
        'checkSuccessorR
        MsgBox ("'Check Successor' should be called from 'Report' sheet")
    Else
        MsgBox ("'Check Successor' should be called from 'Data' sheet or from 'Report' sheet")
    End If
End Sub
    
Sub conSuccessorA(control As IRibbonControl)
    If (ActiveSheet.Name = "Data") Then
        checkSuccessorAll
        MsgBox "Done!"
    ElseIf (ActiveSheet.Name = "Report") Then
        MsgBox ("'Check Successor' should be called from 'Report' sheet")
    Else
        MsgBox ("'Check Successor' should be called from 'Data' sheet or from 'Report' sheet")
    End If
End Sub

Sub conReport(control As IRibbonControl)
    If (ActiveSheet.Name = "Data") Then
        Report
    ElseIf (ActiveSheet.Name = "Report") Then
        Report
    Else
        MsgBox ("'Report' should be called from 'Data' sheet or from 'Report' sheet")
    End If
End Sub
    
Sub conFormat(control As IRibbonControl)
    If (ActiveSheet.Name = "Data") Then
        FormatReport
    ElseIf (ActiveSheet.Name = "Report") Then
        FormatReport
    Else
        MsgBox ("'FormatReport' should be called from 'Data' sheet or from 'Report' sheet")
    End If
End Sub
    
Sub conSow(control As IRibbonControl)
    If (ActiveSheet.Name = "Data") Then
        displSOW
    ElseIf (ActiveSheet.Name = "Report") Then
        displSOW
    Else
        MsgBox ("'Spares on Web' should be called from 'Data' sheet or from 'Report' sheet")
    End If
End Sub

Sub conMall(control As IRibbonControl)
    If (ActiveSheet.Name = "Data") Then
        displMall
    ElseIf (ActiveSheet.Name = "Report") Then
        displMall
    Else
        MsgBox ("'Industry Mall' should be called from 'Data' sheet or from 'Report' sheet")
    End If
End Sub

Sub conSios(control As IRibbonControl)
    If (ActiveSheet.Name = "Data") Then
        displSios
    ElseIf (ActiveSheet.Name = "Report") Then
        displSios
    Else
        MsgBox ("'Industry Support' should be called from 'Data' sheet or from 'Report' sheet")
    End If
End Sub

Sub conOpenWeb(control As IRibbonControl)
    ActiveWorkbook.FollowHyperlink ("https://alxcor.github.io/camxls")
End Sub

Sub conOpenGit(control As IRibbonControl)
    ActiveWorkbook.FollowHyperlink ("https://github.com/alxcor/camxls")
End Sub
    
Sub clearAll()
    'Worksheet 'Data': Clear All
    Sheets("Data").Activate
    Cells.Clear
    Cells.ColumnWidth = 8.5
    Cells.Rows.AutoFit
    ActiveWindow.FreezePanes = False
    Cells(1, 1).Select
End Sub
    
Sub clearAllR()
    'Worksheet 'Report': Clear All
    Sheets("Report").Activate
    Cells.Clear
    Cells.ColumnWidth = 8.5
    Cells.Rows.AutoFit
    ActiveWindow.FreezePanes = False
    Cells(1, 1).Select
End Sub
    
Sub setHeader()
    'Worksheet 'Data': Set header texts on row 1
    Dim rowNumber As Integer
    Dim maxCol As Integer
    rowNumber = 1
    maxCol = 29
    'Select Data worksheet
    Sheets("Data").Activate
    DoEvents
    'Select First Row and check if the row is free for header:
    If (Cells(1, 1) <> "") Then
        If (Cells(1, 1) <> "Your Data...") Then
            Range("A1").EntireRow.Insert
        End If
    End If
    'Clear the first row
    Range(Cells(rowNumber, 1), Cells(rowNumber, maxCol)).Rows.Clear
    'Set texts
    Cells(rowNumber, 1) = "Your Data..."
    Cells(rowNumber, 2) = "MLFB"
    Cells(rowNumber, 3) = "Product Description"
    Cells(rowNumber, 4) = "Product family"
    Cells(rowNumber, 5) = "Product Lifecycle (PLM)"
    Cells(rowNumber, 6) = "PLM Effective Date"
    Cells(rowNumber, 7) = "Notes"
    Cells(rowNumber, 8) = "Price Group"
    Cells(rowNumber, 9) = "Surcharge for Raw Materials"
    Cells(rowNumber, 10) = "Metal Factor"
    Cells(rowNumber, 11) = "Export Control Regulations"
    Cells(rowNumber, 12) = "Dispatch Time"
    Cells(rowNumber, 13) = "Net Weight (kg)"
    Cells(rowNumber, 14) = "Product Dimensions (W x L x H)"
    Cells(rowNumber, 15) = "Packaging Dimension"
    Cells(rowNumber, 16) = "Package size unit of measure"
    Cells(rowNumber, 17) = "Quantity Unit"
    Cells(rowNumber, 18) = "Packaging Quantity"
    Cells(rowNumber, 19) = "EAN"
    Cells(rowNumber, 20) = "UPC"
    Cells(rowNumber, 21) = "Commodity Code"
    Cells(rowNumber, 22) = "KZ_FDB/ CatalogID"
    Cells(rowNumber, 23) = "Product Group"
    Cells(rowNumber, 24) = "Country of origin"
    Cells(rowNumber, 25) = "Compliance with the substance restrictions according to RoHS directive"
    Cells(rowNumber, 26) = "Product class"
    Cells(rowNumber, 27) = "Obligation Category for taking back electrical and electronic equipment after use"
    Cells(rowNumber, 28) = "Classifications"
    Cells(rowNumber, 29) = "Successor"
    Cells(rowNumber, 1).EntireRow.Font.Bold = True
    'Set borders
    With Range(Cells(rowNumber, 1), Cells(rowNumber, maxCol)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    'Set Column width
    setSize
End Sub

Sub setHeaderR()
    'Worksheet 'Report': Set header texts on row 1
    Dim rowNumber As Integer
    Dim maxCol As Integer
    rowNumber = 1
    maxCol = 5
    'Select Data worksheet
    Sheets("Report").Activate
    DoEvents
    'Select First Row and check if the row is free for header:
    If (Cells(1, 1) <> "") Then
        If (Cells(1, 1) <> "MLFB") Then
            Range("A1").EntireRow.Insert
        End If
    End If
    'Select Second Row and check if the row is free for header:
    If (Cells(2, 1) <> "") Then
        If (Left(Cells(2, 1), 6) <> "Active") Then
            Range("A1").EntireRow.Insert
        End If
    End If
    'Write header on first row
    rowNumberR = 1
    Sheets("Report").Cells(rowNumberR, 1).EntireRow.Font.Name = "Arial"
    Sheets("Report").Cells(rowNumberR, 1).EntireRow.Font.Size = 10
    Sheets("Report").Cells(rowNumberR, 1).EntireRow.Font.ColorIndex = xlAutomatic
    Sheets("Report").Cells(rowNumberR, 1).EntireRow.Font.TintAndShade = 0
    Sheets("Report").Cells(rowNumberR, 1) = "MLFB"
    Sheets("Report").Cells(rowNumberR, 2) = "Product Description"
    Sheets("Report").Cells(rowNumberR, 3) = "Product Lifecycle (PLM)"
    Sheets("Report").Cells(rowNumberR, 4) = "Notes"
    Sheets("Report").Cells(rowNumberR, 5) = "Dispatch Time"
    'Format cells
    Sheets("Report").Cells.VerticalAlignment = xlTop
    Sheets("Report").Cells(rowNumberR, 1).EntireColumn.WrapText = True
    Sheets("Report").Cells(rowNumberR, 1).EntireColumn.ColumnWidth = 18
    Sheets("Report").Cells(rowNumberR, 2).EntireColumn.WrapText = True
    Sheets("Report").Cells(rowNumberR, 2).EntireColumn.ColumnWidth = 20
    Sheets("Report").Cells(rowNumberR, 3).EntireColumn.WrapText = True
    Sheets("Report").Cells(rowNumberR, 3).EntireColumn.ColumnWidth = 20
    Sheets("Report").Cells(rowNumberR, 4).EntireColumn.WrapText = True
    Sheets("Report").Cells(rowNumberR, 4).EntireColumn.ColumnWidth = 20
    Sheets("Report").Cells(rowNumberR, 5).EntireColumn.WrapText = True
    Sheets("Report").Cells(rowNumberR, 5).EntireColumn.ColumnWidth = 8
    rowNumberR = 2
    Sheets("Report").Cells(rowNumberR, 1).EntireRow.Font.Name = "Arial"
    Sheets("Report").Cells(rowNumberR, 1).EntireRow.Font.Size = 10
    Sheets("Report").Cells(rowNumberR, 1).EntireRow.Font.ColorIndex = xlAutomatic
    Sheets("Report").Cells(rowNumberR, 1).EntireRow.Font.TintAndShade = 0
    'some final formatting
    Sheets("Report").Range(Cells(1, 1), Cells(1, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    Sheets("Report").Range(Cells(1, 1), Cells(1, 5)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
    Sheets("Report").Range(Cells(1, 1), Cells(1, 5)).Borders(xlEdgeBottom).TintAndShade = 0
    Sheets("Report").Range(Cells(1, 1), Cells(1, 5)).Borders(xlEdgeBottom).Weight = xlThick
    Sheets("Report").Cells(2, 1).Interior.Color = RGB(125, 242, 92)
    Sheets("Report").Cells(2, 2).Interior.Color = RGB(229, 242, 80)
    Sheets("Report").Cells(2, 3).Interior.Color = RGB(242, 135, 148)
    Sheets("Report").Cells(2, 4).Interior.Color = RGB(230, 230, 230)
    Sheets("Report").Cells(2, 5).Interior.Color = RGB(230, 230, 230)
    DoEvents
    'Set borders
    With Range(Cells(rowNumber, 1), Cells(rowNumber, maxCol)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
End Sub

Sub setSize()
    'Worksheet 'Data': Set Row/Column size
    'Select Data worksheet
    Sheets("Data").Activate
    Columns(1).EntireColumn.WrapText = False
    Columns(1).EntireColumn.AutoFit
    Columns(2).EntireColumn.WrapText = False
    Columns(2).EntireColumn.AutoFit
    Columns(3).EntireColumn.WrapText = True
    Columns(3).EntireColumn.ColumnWidth = 40
    Columns(4).EntireColumn.WrapText = True
    Columns(4).EntireColumn.ColumnWidth = 24
    Columns(5).EntireColumn.WrapText = True
    Columns(5).EntireColumn.ColumnWidth = 24
    Columns(6).EntireColumn.WrapText = True
    Columns(6).EntireColumn.ColumnWidth = 18
    Columns(7).EntireColumn.WrapText = True
    Columns(7).EntireColumn.ColumnWidth = 40
    Columns(8).EntireColumn.WrapText = True
    Columns(8).EntireColumn.ColumnWidth = 12
    Columns(9).EntireColumn.WrapText = True
    Columns(9).EntireColumn.ColumnWidth = 30
    Columns(10).EntireColumn.WrapText = True
    Columns(10).EntireColumn.ColumnWidth = 12
    Columns(11).EntireColumn.WrapText = True
    Columns(11).EntireColumn.ColumnWidth = 26
    Columns(12).EntireColumn.WrapText = True
    Columns(12).EntireColumn.ColumnWidth = 14
    Columns(13).EntireColumn.WrapText = True
    Columns(13).EntireColumn.ColumnWidth = 16
    Columns(14).EntireColumn.WrapText = True
    Columns(14).EntireColumn.ColumnWidth = 30
    Columns(15).EntireColumn.WrapText = True
    Columns(15).EntireColumn.ColumnWidth = 22
    Columns(16).EntireColumn.WrapText = True
    Columns(16).EntireColumn.ColumnWidth = 28
    Columns(17).EntireColumn.WrapText = True
    Columns(17).EntireColumn.ColumnWidth = 12
    Columns(18).EntireColumn.WrapText = True
    Columns(18).EntireColumn.ColumnWidth = 20
    Columns(19).EntireColumn.WrapText = True
    Columns(19).EntireColumn.ColumnWidth = 16
    Columns(20).EntireColumn.WrapText = True
    Columns(20).EntireColumn.ColumnWidth = 16
    Columns(21).EntireColumn.WrapText = True
    Columns(21).EntireColumn.ColumnWidth = 16
    Columns(22).EntireColumn.WrapText = True
    Columns(22).EntireColumn.ColumnWidth = 16
    Columns(23).EntireColumn.WrapText = True
    Columns(23).EntireColumn.ColumnWidth = 16
    Columns(24).EntireColumn.WrapText = True
    Columns(24).EntireColumn.ColumnWidth = 16
    Columns(25).EntireColumn.WrapText = True
    Columns(25).EntireColumn.ColumnWidth = 40
    Columns(26).EntireColumn.WrapText = True
    Columns(26).EntireColumn.ColumnWidth = 40
    Columns(27).EntireColumn.WrapText = True
    Columns(27).EntireColumn.ColumnWidth = 40
    Columns(28).EntireColumn.WrapText = True
    Columns(28).EntireColumn.ColumnWidth = 40
    Columns(29).EntireColumn.WrapText = True
    Columns(29).EntireColumn.ColumnWidth = 40
    Cells(1, 1).EntireRow.Font.Bold = True
    Cells(1, 1).EntireRow.WrapText = False
    Cells.Rows.AutoFit
End Sub

Sub setCells(rowNumber)
    'clear data for a range of cells in a row
    Dim maxCol As Integer
    Dim mlfbCode As String
    'last column is AC = column 29
    maxCol = 29
    mlfbCode = Cells(rowNumber, 1)
    'clear data for current row, columns 1 to maxCol
    With Range(Cells(rowNumber, 1), Cells(rowNumber, maxCol))
        .Clear
        .VerticalAlignment = xlTop
    End With
    Cells(rowNumber, 1).Value = mlfbCode
End Sub

Sub readRow()
    'read data for current row: on column 1 [A] should be a product code (MLFB) from Industry Mall web site
    Dim rowNumber As Long
    Dim mlfbCode As String
    Dim isSuccessor As Boolean
    '--------------------------------------------
    'Select Data worksheet
    Sheets("Data").Activate
    DoEvents
    '--------------------------------------------
    rowNumber = ActiveCell.Row
    If rowNumber < 2 Then
        MsgBox ("[EN]: Table starts on row 2; [RO]:Tabelul incepe de la randul 2!")
        GoTo EndSub
    Else
        '----------------------------------------
        Call setCells(rowNumber)
        '----------------------------------------
        mlfbCode = Cells(rowNumber, 1).Value
        isSuccessor = False
        If Len(mlfbCode) > 1 Then
            'remove successor note from code
            If (Left(mlfbCode, 8) = ("[succ.]" + vbLf)) Then
                mlfbCode = Right(mlfbCode, Len(mlfbCode) - 8)
                isSuccessor = True
            End If
            If (Left(mlfbCode, 7) = "[succ.]") Then
                mlfbCode = Right(mlfbCode, Len(mlfbCode) - 7)
                isSuccessor = True
            End If
            Call ImportSieMallIntra(mlfbCode, rowNumber)
            If (isSuccessor = True) Then
                'Add successor note to column2
                Cells(rowNumber, 2).Value = "[succ.]" + vbLf + Cells(rowNumber, 2).Value
            End If
        End If
    End If
    '--------------------------------------------
    setHeader
    setSize
    '--------------------------------------------
    Application.StatusBar = ""
EndSub:
End Sub

Sub readRowR()
    'read data starting from 'Report' worksheet
    Dim rowNumber As Long
    Dim mlfbCode As String
    '--------------------------------------------
    'Select Report worksheet
    Sheets("Report").Activate
    DoEvents
    '--------------------------------------------
    rowNumber = ActiveCell.Row
    If rowNumber < 3 Then
        MsgBox ("[EN]: Table starts on row 3; [RO]:Tabelul incepe de la randul 3!")
        GoTo EndSub
    Else
        mlfbCode = Cells(rowNumber, 1).Value
        Call setCells(rowNumber)
        Sheets("Data").Activate
        DoEvents
        setHeader
        DoEvents
        Cells(rowNumber - 1, 1).Select
        Cells(rowNumber - 1, 1).Value = mlfbCode
        DoEvents
        readRow
        ReportRow
    End If
    Sheets("Report").Activate
    Cells(rowNumber, 1).Select
    '--------------------------------------------
    Application.StatusBar = ""
EndSub:
End Sub

Sub readAll()
    'read data for all non-empty rows >= 2: on column 1 [A] should be a product code (MLFB) from Industry Mall web site
    Dim rowNumber As Long
    Dim mlfbCode As String
    Dim iCounter As Integer
    Dim isSuccessor As Boolean
    Dim maxRow As Integer
    'set a maximum of 500 rows
    maxRow = 500
    '--------------------------------------------
    'Select Data worksheet
    Sheets("Data").Activate
    DoEvents
    '--------------------------------------------
    setHeader
    '--------------------------------------------
    For rowNumber = 2 To 500
        Call setCells(rowNumber)
        '----------------------------------------
        mlfbCode = Cells(rowNumber, 1).Value
        isSuccessor = False
        If Len(mlfbCode) > 1 Then
            'remove successor note from code
            If (Left(mlfbCode, 8) = ("[succ.]" + vbLf)) Then
                mlfbCode = Right(mlfbCode, Len(mlfbCode) - 8)
               isSuccessor = True
            End If
            If (Left(mlfbCode, 7) = "[succ.]") Then
                mlfbCode = Right(mlfbCode, Len(mlfbCode) - 7)
               isSuccessor = True
            End If
            Cells(rowNumber, 1).Select
            Call ImportSieMallIntra(mlfbCode, rowNumber)
            If (isSuccessor = True) Then
                'Add successor note to column2
                Cells(rowNumber, 2).Value = "[succ.]" + vbLf + Cells(rowNumber, 2).Value
            End If
        End If
        DoEvents
    Next
    '--------------------------------------------
    setHeader
    setSize
    '--------------------------------------------
    Application.StatusBar = ""
    Range("A2").Select
    ActiveWindow.FreezePanes = True
EndSub:
End Sub

Sub readAllR()
    'read data starting from 'Report' worksheet
    Dim rowNumber As Long
    Dim mlfbCode As String
    Dim iCounter As Integer
    Dim maxRow As Integer
    'set a maximum of 500 rows
    maxRow = 500
    '--------------------------------------------
    setHeaderR
    '--------------------------------------------
    clearAll
    DoEvents
    '--------------------------------------------
    'Select Data worksheet
    Sheets("Data").Activate
    DoEvents
    '--------------------------------------------
    setHeader
    '--------------------------------------------
    For rowNumber = 2 To 500
        Cells(rowNumber, 1).Value = Sheets("Report").Cells(rowNumber + 1, 1).Value
        DoEvents
    Next
    readAll
    Report
End Sub

Sub checkSuccessor()
    'check successor; if a successor is found in column 29 a new row is added and data is read
    Dim rowNumber As Long
    Dim mlfbCode As String
    '--------------------------------------------
    'Select Data worksheet
    Sheets("Data").Activate
    DoEvents
    '--------------------------------------------
    rowNumber = ActiveCell.Row
    If rowNumber < 2 Then
        MsgBox ("[EN]: Table starts on row 2; [RO]:Tabelul incepe de la randul 2!")
        GoTo EndSub
    Else
        mlfbCode = Cells(rowNumber, 29).Value
        If (("[succ.]" + vbLf + mlfbCode <> Cells(rowNumber + 1, 1).Value) And (mlfbCode <> Cells(rowNumber + 1, 1).Value)) Then
            If (mlfbCode <> "") Then
                Cells(rowNumber + 1, 1).EntireRow.Insert
                Cells(rowNumber + 1, 1).Value = "[succ.]" + vbLf + mlfbCode
                setCells (rowNumer + 1)
                Call ImportSieMallIntra(mlfbCode, rowNumber + 1)
                'Add successor note to column2
                Cells(rowNumber + 1, 2).Value = "[succ.]" + vbLf + Cells(rowNumber + 1, 2).Value
                Cells(rowNumber + 1, 1).Activate
            End If
        End If
    End If
    '--------------------------------------------
    setHeader
    setSize
    '--------------------------------------------
    Application.StatusBar = ""
EndSub:
End Sub

Sub checkSuccessorR()
    'check successor from Report worksheet
    MsgBox "Successor should be checked from Data worksheet..."
End Sub

Sub checkSuccessorAll()
    'check all rows for successors; if a successor is found in column 29 a new row is added and data is read
    Dim rowNumber As Long
    Dim mlfbCode As String
    Dim iCounter As Integer
    Dim maxRow As Integer
    'set a maximum of 500 rows
    maxRow = 500
    '--------------------------------------------
    'Select Data worksheet
    Sheets("Data").Activate
    DoEvents
    '--------------------------------------------
    setHeader
    '--------------------------------------------
    rowNumber = 2
    While (rowNumber <= maxRow)
        Cells(rowNumber, 1).Select
        DoEvents
        mlfbCode = Cells(rowNumber, 29).Value
        If (("[succ.]" + vbLf + mlfbCode <> Cells(rowNumber + 1, 1).Value) And (mlfbCode <> Cells(rowNumber + 1, 1).Value)) Then
            If (mlfbCode <> "") Then
                Cells(rowNumber + 1, 1).EntireRow.Insert
                Cells(rowNumber + 1, 1).Value = "[succ.]" + vbLf + mlfbCode
                setCells (rowNumer + 1)
                Call ImportSieMallIntra(mlfbCode, rowNumber + 1)
                'Add successor note to column2
                Cells(rowNumber + 1, 2).Value = "[succ.]" + vbLf + Cells(rowNumber + 1, 2).Value
                maxRow = maxRow + 1
            End If
        End If
        rowNumber = rowNumber + 1
    Wend
    '--------------------------------------------
    setHeader
    setSize
    '--------------------------------------------
    Application.StatusBar = ""
    Range("A2").Select
    ActiveWindow.FreezePanes = True
EndSub:
End Sub

Sub ImportSieMallIntra(mlfbCode, rowNumber)
    'read data for a specific product code (MLFB) from Industry Mall web site
    'netMode = xmlHTTP version
    On Error GoTo ErrHand:   'disable this line to see what is the error
    Dim targetURL As String
    Dim webContent As String
    Dim index As Integer
    Dim DetailNo As Integer
    '--------------------------------------------
    Cells(rowNumber, 2).Value = mlfbCode
    Cells(rowNumber, 5).Value = "ERR: Not Found!!!"
    Cells(rowNumber, 5).Interior.Color = RGB(242, 135, 148)
    'Reading web page in buffer...
    'write status in StatusBar...
    Application.StatusBar = "Trying to connect to Industry Mall... MLFB: " + mlfbCode
    'options should be separated by space instead of +
    mlfbCode = Replace(mlfbCode, "+", "%20")
    'format spaces html style
    mlfbCode = Replace(mlfbCode, " ", "%20")
    'set web page (for scrapper)
    targetURL = "https://mall.industry.siemens.com/mall/en/WW/Catalog/Product/" + mlfbCode
    '--------------------------------------------
    'xmlHTTP version
    Application.StatusBar = "Trying to connect -via xmlHTTP- to Industry Mall... MLFB: " + mlfbCode
    Dim xmlhttp As Object
    Set xmlhttp = CreateObject("MSXML2.serverXMLHTTP")
    xmlhttp.Open "GET", targetURL, False
    xmlhttp.send
    'MsgBox (xmlhttp.responseText)
    Dim htmldoc As Object
    Set htmldoc = CreateObject("HTMLFile")
    htmldoc.body.innerHTML = xmlhttp.responseText
    Application.StatusBar = "..."
    'MsgBox (htmldoc.body.innerHTML)
    'See content of the web-page for diagnosis purposes...
    'MsgBox html.documentElement.innerHTML
    DoEvents
    index = 1
    DetailNo = 1
    'Search for ID-content in web page
    Set Product = htmldoc.getElementById("content")
    Set ProductDetails = Product.all
    Application.StatusBar = "Trying to get details for MLFB: " + mlfbCode
    For Each Detail In ProductDetails
        'MLFB code...
        If Detail.className = "productIdentifier" Then
            Cells(rowNumber, 2).Value = Detail.innerText
        End If
        'Produs details - Extract from table:
        If Detail.className = "ProductDetailsTable" Then
            'Get all details for the product
            Set Details = Detail.all
            'Count details fields
            DetailNo = Details.Length
            For index = 0 To DetailNo - 1
                'Product Description
                If (Details(index).innerText = "Product Description") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 3) = Details(index + 1).innerText
                End If
                'Product Family
                If (Details(index).innerText = "Product family") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 4) = Details(index + 1).innerText
                End If
                'Product Lifecycle(PLM)
                If (Details(index).innerText = "Product Lifecycle (PLM)") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 5) = Details(index + 1).innerText
                    'PLM status:
                    sTemp = Details(index + 1).innerText
                    iTemp = InStr(1, sTemp, "M250", vbTextCompare)
                    If iTemp > 0 Then
                        Cells(rowNumber, 5).Interior.Color = RGB(125, 242, 92)
                    End If
                    iTemp = InStr(1, sTemp, "M280", vbTextCompare)
                    If iTemp > 0 Then
                        Cells(rowNumber, 5).Interior.Color = RGB(125, 242, 92)
                    End If
                    iTemp = InStr(1, sTemp, "M300", vbTextCompare)
                    If iTemp > 0 Then
                        Cells(rowNumber, 5).Interior.Color = RGB(125, 242, 92)
                    End If
                    iTemp = InStr(1, sTemp, "M400", vbTextCompare)
                    If iTemp > 0 Then
                        Cells(rowNumber, 5).Interior.Color = RGB(229, 242, 80)
                    End If
                    iTemp = InStr(1, sTemp, "M410", vbTextCompare)
                    If iTemp > 0 Then
                        Cells(rowNumber, 5).Interior.Color = RGB(229, 242, 80)
                    End If
                    iTemp = InStr(1, sTemp, "M490", vbTextCompare)
                    If iTemp > 0 Then
                        Cells(rowNumber, 5).Interior.Color = RGB(242, 135, 148)
                    End If
                    iTemp = InStr(1, sTemp, "M500", vbTextCompare)
                    If iTemp > 0 Then
                        Cells(rowNumber, 5).Interior.Color = RGB(242, 135, 148)
                    End If
                End If
                'PLM Effective Date
                If (Details(index).innerText = "PLM Effective Date") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 6) = Details(index + 1).innerText
                End If
                'Notes
                If (Details(index).innerText = "Notes") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 7) = Details(index + 1).innerText
                    'sTemp = Details(index + 1).innerText
                    'iTemp = InStr(1, sTemp, "Successor", vbTextCompare)
                    'iTemp2 = Len(sTemp)
                    'If (iTemp > 0) And ((iTemp2 - iTemp > 11)) Then
                        'iTemp = iTemp + 11
                        'sTemp = Mid(sTemp, iTemp, iTemp2 - iTemp)
                        ''Successor: attempt to identify MLFB for successor product
                        'Cells(rowNumber, 29) = sTemp
                    'End If
                    If Cells(rowNumber, 7).Value <> "" Then
                        Cells(rowNumber, 7).Interior.Color = RGB(91, 155, 213)
                    End If
                End If
                'Price Group
                If (Details(index).innerText = "Price Group") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 8) = Details(index + 1).innerText
                End If
                'New Price Group [230209]
                If (Details(index).innerText = "Region Specific PriceGroup / Headquarter Price Group") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 8) = Details(index + 1).innerText
                End If
                'Surcharge for Raw Materials
                If (Details(index).innerText = "Surcharge for Raw Materials") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 9) = Details(index + 1).innerText
                End If
                'Surcharge for Metal Factor
                If (Details(index).innerText = "Metal Factor") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 10) = Details(index + 1).innerText
                End If
                'Export Control Regulations
                If (Details(index).innerText = "Export Control Regulations") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 11) = Details(index + 1).innerText
                End If
                'Delivery Time
                If (Details(index).innerText = "Delivery Time") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 12) = Details(index + 1).innerText
                End If
                'New Delivery Time [230209]
                If (Details(index).innerText = "Standard lead time ex-works") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 12) = Details(index + 1).innerText
                End If
                'New Delivery Time [240410]
                If (Details(index).innerText = "Estimated dispatch time (Working Days)") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 12) = Details(index + 1).innerText
                End If
                'Net Weight(kg)
                If (Details(index).innerText = "Net Weight(kg)") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 13) = Details(index + 1).innerText
                End If
                'New Net Weight(kg) [230209]
                If (Details(index).innerText = "Net Weight (kg)") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 13) = Details(index + 1).innerText
                End If
                'Product Dimensions (W x L x H)
                If (Details(index).innerText = "Product Dimensions (W x L x H)") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 14) = Details(index + 1).innerText
                End If
                'Packaging Not Dimension
                If (Details(index).innerText = "Packaging Dimension") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 15) = Details(index + 1).innerText
                End If
                'Package size unit of measure
                If (Details(index).innerText = "Package size unit of measure") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 16) = Details(index + 1).innerText
                End If
                'Quantity Unit
                If (Details(index).innerText = "Quantity Unit") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 17) = Details(index + 1).innerText
                End If
                'Packaging Quantity
                If (Details(index).innerText = "Packaging Quantity") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 18) = Details(index + 1).innerText
                End If
                'EAN
                If (Details(index).innerText = "EAN") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 19) = "'" & Details(index + 1).innerText
                End If
                'UPC
                If (Details(index).innerText = "UPC") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 20) = "'" & Details(index + 1).innerText
                End If
                'Commodity Code
                If (Details(index).innerText = "Commodity Code") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 21) = "'" & Details(index + 1).innerText
                End If
                'LKZ_FDB/ CatalogID
                If (Details(index).innerText = "LKZ_FDB/ CatalogID") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 22) = Details(index + 1).innerText
                End If
                'Product Group
                If (Details(index).innerText = "Product Group") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 23) = Details(index + 1).innerText
                End If
                'Country of origin
                If (Details(index).innerText = "Country of origin") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 24) = Details(index + 1).innerText
                End If
                'Compliance with the substance restrictions according to RoHS directive
                If (Details(index).innerText = "Compliance with the substance restrictions according to RoHS directive") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 25) = Details(index + 1).innerText
                End If
                'Product class
                If (Details(index).innerText = "Product class") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 26) = Details(index + 1).innerText
                End If
                'Obligation Category for taking back electrical and electronic equipment after use
                If (Details(index).innerText = "Obligation Category for taking back electrical and electronic equipment after use") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 27) = Details(index + 1).innerText
                End If
                'Classifications
                If (Details(index).innerText = "Classifications") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 28) = Details(index + 1).innerText
                End If
                'Successor
                If (Details(index).innerText = "Successor") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 29) = Details(index + 1).innerText
                End If
            Next
        End If
    Next
    Set xmlhttp = Nothing
    Set htmldoc = Nothing
    Application.StatusBar = ""
    GoTo EndSub
ErrHand:
    Cells(rowNumber, 3) = "Error! " & Err.Description
    Application.StatusBar = ""
EndSub:
End Sub

Sub Report()
    'generate a printable report worksheet
    Dim rowNumber As Long
    Dim rowNumberR As Long
    Dim mlfbCode As String
    Dim iCounter As Integer
    Dim maxRow As Integer
    Dim partPM, partsOK, partsAT, partsER, partsNA As Integer
    partsOK = 0
    partsAT = 0
    partsER = 0
    partsNA = 0
    maxRow = 500
    '--------------------------------------------
    Sheets("Report").Activate
    ActiveWindow.FreezePanes = False
    'Clear all data in Report worksheet
    Sheets("Report").Cells.Delete
    DoEvents
    'Write header on first row
    rowNumberR = 1
    Sheets("Report").Cells(rowNumberR, 1).EntireRow.Font.Name = "Arial"
    Sheets("Report").Cells(rowNumberR, 1).EntireRow.Font.Size = 10
    Sheets("Report").Cells(rowNumberR, 1).EntireRow.Font.ColorIndex = xlAutomatic
    Sheets("Report").Cells(rowNumberR, 1).EntireRow.Font.TintAndShade = 0
    Sheets("Report").Cells(rowNumberR, 1) = "MLFB"
    Sheets("Report").Cells(rowNumberR, 2) = "Product Description"
    Sheets("Report").Cells(rowNumberR, 3) = "Product Lifecycle (PLM)"
    Sheets("Report").Cells(rowNumberR, 4) = "Notes"
    Sheets("Report").Cells(rowNumberR, 5) = "Dispatch Time"
    'Format cells
    Sheets("Report").Cells.VerticalAlignment = xlTop
    Sheets("Report").Cells(rowNumberR, 1).EntireColumn.WrapText = True
    Sheets("Report").Cells(rowNumberR, 1).EntireColumn.ColumnWidth = 18
    Sheets("Report").Cells(rowNumberR, 2).EntireColumn.WrapText = True
    Sheets("Report").Cells(rowNumberR, 2).EntireColumn.ColumnWidth = 20
    Sheets("Report").Cells(rowNumberR, 3).EntireColumn.WrapText = True
    Sheets("Report").Cells(rowNumberR, 3).EntireColumn.ColumnWidth = 20
    Sheets("Report").Cells(rowNumberR, 4).EntireColumn.WrapText = True
    Sheets("Report").Cells(rowNumberR, 4).EntireColumn.ColumnWidth = 20
    Sheets("Report").Cells(rowNumberR, 5).EntireColumn.WrapText = True
    Sheets("Report").Cells(rowNumberR, 5).EntireColumn.ColumnWidth = 8
    rowNumberR = 2
    Sheets("Report").Cells(rowNumberR, 1).EntireRow.Font.Name = "Arial"
    Sheets("Report").Cells(rowNumberR, 1).EntireRow.Font.Size = 10
    Sheets("Report").Cells(rowNumberR, 1).EntireRow.Font.ColorIndex = xlAutomatic
    Sheets("Report").Cells(rowNumberR, 1).EntireRow.Font.TintAndShade = 0
    DoEvents
    '--------------------------------------------
    rowNumberR = 3
    For rowNumber = 2 To maxRow
        'Spare part availability ignored
        partPM = 0
        If (Sheets("Data").Cells(rowNumber, 2).Value <> "") Then
            Cells(rowNumberR, 1).Select
            'Spare part availability not yet established
            partPM = 1
            'Format cells
            Sheets("Report").Cells(rowNumberR, 1).EntireRow.Font.Name = "Arial"
            Sheets("Report").Cells(rowNumberR, 1).EntireRow.Font.Size = 8
            Sheets("Report").Cells(rowNumberR, 1).EntireRow.Font.ColorIndex = xlAutomatic
            Sheets("Report").Cells(rowNumberR, 1).EntireRow.Font.TintAndShade = 0
            For iCounter = 1 To 5
                Sheets("Report").Cells(rowNumberR, iCounter).Borders(xlEdgeBottom).LineStyle = xlContinuous
                Sheets("Report").Cells(rowNumberR, iCounter).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
                Sheets("Report").Cells(rowNumberR, iCounter).Borders(xlEdgeBottom).TintAndShade = 0
                Sheets("Report").Cells(rowNumberR, iCounter).Borders(xlEdgeBottom).Weight = xlHairline
                Sheets("Report").Cells(rowNumberR, iCounter).Borders(xlEdgeLeft).LineStyle = xlContinuous
                Sheets("Report").Cells(rowNumberR, iCounter).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
                Sheets("Report").Cells(rowNumberR, iCounter).Borders(xlEdgeLeft).TintAndShade = 0
                Sheets("Report").Cells(rowNumberR, iCounter).Borders(xlEdgeLeft).Weight = xlHairline
                Sheets("Report").Cells(rowNumberR, iCounter).Borders(xlEdgeRight).LineStyle = xlContinuous
                Sheets("Report").Cells(rowNumberR, iCounter).Borders(xlEdgeRight).ColorIndex = xlAutomatic
                Sheets("Report").Cells(rowNumberR, iCounter).Borders(xlEdgeRight).TintAndShade = 0
                Sheets("Report").Cells(rowNumberR, iCounter).Borders(xlEdgeRight).Weight = xlHairline
                Sheets("Report").Cells(rowNumberR, iCounter).Borders(xlEdgeTop).LineStyle = xlContinuous
                Sheets("Report").Cells(rowNumberR, iCounter).Borders(xlEdgeTop).ColorIndex = xlAutomatic
                Sheets("Report").Cells(rowNumberR, iCounter).Borders(xlEdgeTop).TintAndShade = 0
                Sheets("Report").Cells(rowNumberR, iCounter).Borders(xlEdgeTop).Weight = xlHairline
            Next
            mlfbCode = Sheets("Data").Cells(rowNumber, 2).Value
            Sheets("Report").Cells(rowNumberR, 1) = mlfbCode
            Sheets("Report").Cells(rowNumberR, 2) = Sheets("Data").Cells(rowNumber, 3).Value
            Sheets("Report").Cells(rowNumberR, 3) = Sheets("Data").Cells(rowNumber, 5).Value + vbCrLf + vbCrLf + Sheets("Data").Cells(rowNumber, 6).Value
            Sheets("Report").Cells(rowNumberR, 4) = Sheets("Data").Cells(rowNumber, 7).Value
            Sheets("Report").Cells(rowNumberR, 5) = Sheets("Data").Cells(rowNumber, 12).Value
            'PLM:
            Sheets("Report").Cells(rowNumberR, 3).Interior.Color = RGB(230, 230, 230)
            sTemp = Sheets("Data").Cells(rowNumber, 5).Value
            iTemp = InStr(1, sTemp, "M250", vbTextCompare)
            If iTemp > 0 Then
                partPM = 250
                Sheets("Report").Cells(rowNumberR, 3).Interior.Color = RGB(125, 242, 92)
            End If
            iTemp = InStr(1, sTemp, "M280", vbTextCompare)
            If iTemp > 0 Then
                partPM = 280
                Sheets("Report").Cells(rowNumberR, 3).Interior.Color = RGB(125, 242, 92)
            End If
            iTemp = InStr(1, sTemp, "M300", vbTextCompare)
            If iTemp > 0 Then
                partPM = 300
                Sheets("Report").Cells(rowNumberR, 3).Interior.Color = RGB(125, 242, 92)
            End If
            iTemp = InStr(1, sTemp, "M400", vbTextCompare)
            If iTemp > 0 Then
                partPM = 400
                Sheets("Report").Cells(rowNumberR, 3).Interior.Color = RGB(229, 242, 80)
            End If
            iTemp = InStr(1, sTemp, "M410", vbTextCompare)
            If iTemp > 0 Then
                partPM = 410
                Sheets("Report").Cells(rowNumberR, 3).Interior.Color = RGB(229, 242, 80)
            End If
            iTemp = InStr(1, sTemp, "M490", vbTextCompare)
            If iTemp > 0 Then
                partPM = 490
                Sheets("Report").Cells(rowNumberR, 3).Interior.Color = RGB(242, 135, 148)
            End If
            iTemp = InStr(1, sTemp, "M500", vbTextCompare)
            If iTemp > 0 Then
                partPM = 500
                Sheets("Report").Cells(rowNumberR, 3).Interior.Color = RGB(242, 135, 148)
            End If
        End If
      Select Case partPM
        Case 250, 280, 300
            partsOK = partsOK + 1
        Case 400, 410
            partsAT = partsAT + 1
        Case 490, 500
            partsER = partsER + 1
        Case 1
            partsNA = partsNA + 1
        Case Else
            'nothing to do
      End Select
      rowNumberR = rowNumberR + 1
      DoEvents
    Next
    'some final formatting
    Sheets("Report").Select
    Sheets("Report").Range(Cells(1, 1), Cells(1, 5)).Borders(xlEdgeTop).LineStyle = xlContinuous
    Sheets("Report").Range(Cells(1, 1), Cells(1, 5)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
    Sheets("Report").Range(Cells(1, 1), Cells(1, 5)).Borders(xlEdgeTop).TintAndShade = 0
    Sheets("Report").Range(Cells(1, 1), Cells(1, 5)).Borders(xlEdgeTop).Weight = xlThin
    Sheets("Report").Range(Cells(1, 1), Cells(1, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    Sheets("Report").Range(Cells(1, 1), Cells(1, 5)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
    Sheets("Report").Range(Cells(1, 1), Cells(1, 5)).Borders(xlEdgeBottom).TintAndShade = 0
    Sheets("Report").Range(Cells(1, 1), Cells(1, 5)).Borders(xlEdgeBottom).Weight = xlThick
    Sheets("Report").Cells(2, 1).Value = "Active: " & CStr(partsOK)
    Sheets("Report").Cells(2, 1).Interior.Color = RGB(125, 242, 92)
    Sheets("Report").Cells(2, 2).Value = "PhaseOut: " & CStr(partsAT)
    Sheets("Report").Cells(2, 2).Interior.Color = RGB(229, 242, 80)
    Sheets("Report").Cells(2, 3).Value = "Disc.: " & CStr(partsER)
    Sheets("Report").Cells(2, 3).Interior.Color = RGB(242, 135, 148)
    Sheets("Report").Cells(2, 4).Value = "Other: " & CStr(partsNA)
    Sheets("Report").Cells(2, 4).Interior.Color = RGB(230, 230, 230)
    Sheets("Report").Cells(2, 5).Interior.Color = RGB(230, 230, 230)
    DoEvents
    'Format report worksheet
    With Worksheets("Report").PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterHorizontally = True
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .AlignMarginsHeaderFooter = False
        .TopMargin = Application.InchesToPoints(0.7)
        .BottomMargin = Application.InchesToPoints(0.5)
        .LeftMargin = Application.InchesToPoints(0.7)
        .RightMargin = Application.InchesToPoints(0.7)
        .HeaderMargin = Application.InchesToPoints(0.5)
        .FooterMargin = Application.InchesToPoints(0.3)
        .LeftHeader = "&L&08" & "SPARE PARTS Report"
        .CenterHeader = ""
        .RightHeader = "&R&08" & "&D &T"
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = "&R&08" & "&P / &N"
        .PrintTitleRows = "$1:$1"
    End With
    DoEvents
    Sheets("Report").Activate
    Range("A3").Select
    ActiveWindow.FreezePanes = True
    Cells(2, 1).Activate
End Sub

Sub ReportRow()
    'generate a printable report row for current row
    Dim rowNumber As Long
    Dim rowNumberR As Long
    Dim mlfbCode As String
    Dim maxRow As Integer
    Sheets("Data").Activate
    rowNumber = ActiveCell.Row
    '--------------------------------------------
    Sheets("Report").Activate
    ActiveWindow.FreezePanes = False
    DoEvents
    'Write header on first row
    rowNumberR = 1
    Sheets("Report").Cells(rowNumberR, 1).EntireRow.Font.Name = "Arial"
    Sheets("Report").Cells(rowNumberR, 1).EntireRow.Font.Size = 10
    Sheets("Report").Cells(rowNumberR, 1).EntireRow.Font.ColorIndex = xlAutomatic
    Sheets("Report").Cells(rowNumberR, 1).EntireRow.Font.TintAndShade = 0
    Sheets("Report").Cells(rowNumberR, 1) = "MLFB"
    Sheets("Report").Cells(rowNumberR, 2) = "Product Description"
    Sheets("Report").Cells(rowNumberR, 3) = "Product Lifecycle (PLM)"
    Sheets("Report").Cells(rowNumberR, 4) = "Notes"
    Sheets("Report").Cells(rowNumberR, 5) = "Dispatch Time"
    'Format cells
    Sheets("Report").Cells.VerticalAlignment = xlTop
    Sheets("Report").Cells(rowNumberR, 1).EntireColumn.WrapText = True
    Sheets("Report").Cells(rowNumberR, 1).EntireColumn.ColumnWidth = 18
    Sheets("Report").Cells(rowNumberR, 2).EntireColumn.WrapText = True
    Sheets("Report").Cells(rowNumberR, 2).EntireColumn.ColumnWidth = 20
    Sheets("Report").Cells(rowNumberR, 3).EntireColumn.WrapText = True
    Sheets("Report").Cells(rowNumberR, 3).EntireColumn.ColumnWidth = 20
    Sheets("Report").Cells(rowNumberR, 4).EntireColumn.WrapText = True
    Sheets("Report").Cells(rowNumberR, 4).EntireColumn.ColumnWidth = 20
    Sheets("Report").Cells(rowNumberR, 5).EntireColumn.WrapText = True
    Sheets("Report").Cells(rowNumberR, 5).EntireColumn.ColumnWidth = 8
    rowNumberR = 2
    Sheets("Report").Cells(rowNumberR, 1).EntireRow.Font.Name = "Arial"
    Sheets("Report").Cells(rowNumberR, 1).EntireRow.Font.Size = 10
    Sheets("Report").Cells(rowNumberR, 1).EntireRow.Font.ColorIndex = xlAutomatic
    Sheets("Report").Cells(rowNumberR, 1).EntireRow.Font.TintAndShade = 0
    DoEvents
    '--------------------------------------------
    rowNumberR = rowNumber + 1
    'Spare part availability ignored
     partPM = 0
    If (Sheets("Data").Cells(rowNumber, 2).Value <> "") Then
        Cells(rowNumberR, 1).Select
        'Spare part availability not yet established
        partPM = 1
        'Format cells
        Sheets("Report").Cells(rowNumberR, 1).EntireRow.Font.Name = "Arial"
        Sheets("Report").Cells(rowNumberR, 1).EntireRow.Font.Size = 8
        Sheets("Report").Cells(rowNumberR, 1).EntireRow.Font.ColorIndex = xlAutomatic
        Sheets("Report").Cells(rowNumberR, 1).EntireRow.Font.TintAndShade = 0
        For iCounter = 1 To 5
            Sheets("Report").Cells(rowNumberR, iCounter).Borders(xlEdgeBottom).LineStyle = xlContinuous
            Sheets("Report").Cells(rowNumberR, iCounter).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
            Sheets("Report").Cells(rowNumberR, iCounter).Borders(xlEdgeBottom).TintAndShade = 0
            Sheets("Report").Cells(rowNumberR, iCounter).Borders(xlEdgeBottom).Weight = xlHairline
            Sheets("Report").Cells(rowNumberR, iCounter).Borders(xlEdgeLeft).LineStyle = xlContinuous
            Sheets("Report").Cells(rowNumberR, iCounter).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
            Sheets("Report").Cells(rowNumberR, iCounter).Borders(xlEdgeLeft).TintAndShade = 0
            Sheets("Report").Cells(rowNumberR, iCounter).Borders(xlEdgeLeft).Weight = xlHairline
            Sheets("Report").Cells(rowNumberR, iCounter).Borders(xlEdgeRight).LineStyle = xlContinuous
            Sheets("Report").Cells(rowNumberR, iCounter).Borders(xlEdgeRight).ColorIndex = xlAutomatic
            Sheets("Report").Cells(rowNumberR, iCounter).Borders(xlEdgeRight).TintAndShade = 0
            Sheets("Report").Cells(rowNumberR, iCounter).Borders(xlEdgeRight).Weight = xlHairline
            Sheets("Report").Cells(rowNumberR, iCounter).Borders(xlEdgeTop).LineStyle = xlContinuous
            Sheets("Report").Cells(rowNumberR, iCounter).Borders(xlEdgeTop).ColorIndex = xlAutomatic
            Sheets("Report").Cells(rowNumberR, iCounter).Borders(xlEdgeTop).TintAndShade = 0
            Sheets("Report").Cells(rowNumberR, iCounter).Borders(xlEdgeTop).Weight = xlHairline
        Next
        mlfbCode = Sheets("Data").Cells(rowNumber, 2).Value
        Sheets("Report").Cells(rowNumberR, 1) = mlfbCode
        Sheets("Report").Cells(rowNumberR, 2) = Sheets("Data").Cells(rowNumber, 3).Value
        Sheets("Report").Cells(rowNumberR, 3) = Sheets("Data").Cells(rowNumber, 5).Value + vbCrLf + vbCrLf + Sheets("Data").Cells(rowNumber, 6).Value
        Sheets("Report").Cells(rowNumberR, 4) = Sheets("Data").Cells(rowNumber, 7).Value
        Sheets("Report").Cells(rowNumberR, 5) = Sheets("Data").Cells(rowNumber, 12).Value
        'PLM:
        Sheets("Report").Cells(rowNumberR, 3).Interior.Color = RGB(230, 230, 230)
        sTemp = Sheets("Data").Cells(rowNumber, 5).Value
        iTemp = InStr(1, sTemp, "M250", vbTextCompare)
        If iTemp > 0 Then
            partPM = 250
            Sheets("Report").Cells(rowNumberR, 3).Interior.Color = RGB(125, 242, 92)
        End If
        iTemp = InStr(1, sTemp, "M280", vbTextCompare)
        If iTemp > 0 Then
            partPM = 280
            Sheets("Report").Cells(rowNumberR, 3).Interior.Color = RGB(125, 242, 92)
        End If
        iTemp = InStr(1, sTemp, "M300", vbTextCompare)
        If iTemp > 0 Then
            partPM = 300
            Sheets("Report").Cells(rowNumberR, 3).Interior.Color = RGB(125, 242, 92)
        End If
        iTemp = InStr(1, sTemp, "M400", vbTextCompare)
        If iTemp > 0 Then
            partPM = 400
            Sheets("Report").Cells(rowNumberR, 3).Interior.Color = RGB(229, 242, 80)
        End If
        iTemp = InStr(1, sTemp, "M410", vbTextCompare)
        If iTemp > 0 Then
            partPM = 410
            Sheets("Report").Cells(rowNumberR, 3).Interior.Color = RGB(229, 242, 80)
        End If
        iTemp = InStr(1, sTemp, "M490", vbTextCompare)
        If iTemp > 0 Then
            partPM = 490
            Sheets("Report").Cells(rowNumberR, 3).Interior.Color = RGB(242, 135, 148)
        End If
        iTemp = InStr(1, sTemp, "M500", vbTextCompare)
        If iTemp > 0 Then
            partPM = 500
            Sheets("Report").Cells(rowNumberR, 3).Interior.Color = RGB(242, 135, 148)
        End If
    End If
    DoEvents
    'some final formatting
    Sheets("Report").Select
    Sheets("Report").Range(Cells(1, 1), Cells(1, 5)).Borders(xlEdgeTop).LineStyle = xlContinuous
    Sheets("Report").Range(Cells(1, 1), Cells(1, 5)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
    Sheets("Report").Range(Cells(1, 1), Cells(1, 5)).Borders(xlEdgeTop).TintAndShade = 0
    Sheets("Report").Range(Cells(1, 1), Cells(1, 5)).Borders(xlEdgeTop).Weight = xlThin
    Sheets("Report").Range(Cells(1, 1), Cells(1, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    Sheets("Report").Range(Cells(1, 1), Cells(1, 5)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
    Sheets("Report").Range(Cells(1, 1), Cells(1, 5)).Borders(xlEdgeBottom).TintAndShade = 0
    Sheets("Report").Range(Cells(1, 1), Cells(1, 5)).Borders(xlEdgeBottom).Weight = xlThick
    Sheets("Report").Cells(2, 1).Value = "Active: ?"
    Sheets("Report").Cells(2, 1).Interior.Color = RGB(125, 242, 92)
    Sheets("Report").Cells(2, 2).Value = "PhaseOut: ?"
    Sheets("Report").Cells(2, 2).Interior.Color = RGB(229, 242, 80)
    Sheets("Report").Cells(2, 3).Value = "Disc.: ?"
    Sheets("Report").Cells(2, 3).Interior.Color = RGB(242, 135, 148)
    Sheets("Report").Cells(2, 4).Value = "Other: ?"
    Sheets("Report").Cells(2, 4).Interior.Color = RGB(230, 230, 230)
    Sheets("Report").Cells(2, 5).Interior.Color = RGB(230, 230, 230)
    DoEvents
    'Format report worksheet
    With Worksheets("Report").PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterHorizontally = True
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .AlignMarginsHeaderFooter = False
        .TopMargin = Application.InchesToPoints(0.7)
        .BottomMargin = Application.InchesToPoints(0.5)
        .LeftMargin = Application.InchesToPoints(0.7)
        .RightMargin = Application.InchesToPoints(0.7)
        .HeaderMargin = Application.InchesToPoints(0.5)
        .FooterMargin = Application.InchesToPoints(0.3)
        .LeftHeader = "&L&08" & "SPARE PARTS Report"
        .CenterHeader = ""
        .RightHeader = "&R&08" & "&D &T"
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = "&R&08" & "&P / &N"
        .PrintTitleRows = "$1:$1"
    End With
    DoEvents
    Sheets("Report").Activate
    Range("A3").Select
    ActiveWindow.FreezePanes = True
    Cells(2, 1).Activate
End Sub

Sub FormatReport()
    'Format report worksheet
    With Worksheets("Report").PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterHorizontally = True
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .AlignMarginsHeaderFooter = False
        .TopMargin = Application.InchesToPoints(0.7)
        .BottomMargin = Application.InchesToPoints(0.5)
        .LeftMargin = Application.InchesToPoints(0.7)
        .RightMargin = Application.InchesToPoints(0.7)
        .HeaderMargin = Application.InchesToPoints(0.5)
        .FooterMargin = Application.InchesToPoints(0.3)
        .LeftHeader = "&L&08" & "SPARE PARTS Report"
        .CenterHeader = ""
        .RightHeader = "&R&08" & "&D &T"
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = "&R&08" & "&P / &N"
        .PrintTitleRows = "$1:$1"
    End With
    With Worksheets("Report")
        .PrintPreview
    End With
End Sub

Sub displSOW()
    'display Spares On Web for accessories
    Dim rowNumber As Long
    Dim mlfbCode As String
    Dim mlfbOpts As String
    Dim URL As String
    Dim optPos As Integer
    Dim optLen As Integer
    Application.StatusBar = "Check Spares On Web..."
    rowNumber = ActiveCell.Row
    If rowNumber < 2 Then
        MsgBox ("[EN]: Table starts on row 2; [RO]:Tabelul incepe de la randul 2!")
        GoTo EndSub
    Else
        mlfbCode = Cells(rowNumber, 1).Value
        'remove successor note from code
        If (Left(mlfbCode, 8) = ("[succ.]" + vbLf)) Then
            mlfbCode = Right(mlfbCode, Len(mlfbCode) - 8)
        End If
        If (Left(mlfbCode, 7) = "[succ.]") Then
            mlfbCode = Right(mlfbCode, Len(mlfbCode) - 7)
        End If
        'Try to extract options
        optLen = Len(mlfbCode)
        optPos = InStr(1, mlfbCode, "-Z ", vbTextCompare)
        If ((optPos > 0) And (optPos < optLen)) Then
            mlfbOpts = Right(mlfbCode, Len(mlfbCode) - optPos - 2)
            mlfbCode = Left(mlfbCode, optPos + 1)
        Else
            mlfbOpts = ""
        End If
        'options should be separated by space instead of +
        mlfbOpts = Replace(mlfbOpts, "+", "%20")
        If Len(mlfbCode) > 1 Then
            If (Len(mlfbOpts) > 1) Then
                URL = "https://www.sow.siemens.com/?an=" + mlfbCode + "&op=" + mlfbOpts
            Else
                URL = "https://www.sow.siemens.com/?an=" + mlfbCode
            End If
            ActiveWorkbook.FollowHyperlink URL
        End If
    End If
EndSub:
    Application.StatusBar = ""
End Sub

Sub displMall()
    'display IndustryMall for a spare part
    Dim rowNumber As Long
    Dim mlfbCode As String
    Dim URL As String
    Application.StatusBar = "Check IndustryMall..."
    rowNumber = ActiveCell.Row
    If rowNumber < 2 Then
        MsgBox ("[EN]: Table starts on row 2; [RO]:Tabelul incepe de la randul 2!")
        GoTo EndSub
    Else
        mlfbCode = Cells(rowNumber, 1).Value
        'remove successor note from code
        If (Left(mlfbCode, 8) = ("[succ.]" + vbLf)) Then
            mlfbCode = Right(mlfbCode, Len(mlfbCode) - 8)
        End If
        If (Left(mlfbCode, 7) = "[succ.]") Then
            mlfbCode = Right(mlfbCode, Len(mlfbCode) - 7)
        End If
        'options should be separated by space instead of +
        mlfbCode = Replace(mlfbCode, "+", "%20")
        If Len(mlfbCode) > 1 Then
            URL = "https://mall.industry.siemens.com/mall/en/WW/Catalog/Product/" + mlfbCode
            ActiveWorkbook.FollowHyperlink URL
        End If
    End If
EndSub:
    Application.StatusBar = ""
End Sub

Sub displSios()
    'display IndustrySupport for a spare part from Report
    Dim rowNumber As Long
    Dim mlfbCode As String
    Dim mlfbOpts As String
    Dim URL As String
    Dim optPos As Integer
    Dim optLen As Integer
    '--------------------------------------------
    Application.StatusBar = "Check SIOS Support..."
    rowNumber = ActiveCell.Row
    If rowNumber < 2 Then
        MsgBox ("[EN]: Table starts on row 2; [RO]:Tabelul incepe de la randul 2!")
        GoTo EndSub
    Else
        mlfbCode = Cells(rowNumber, 1).Value
        'Try to extract options
        optLen = Len(mlfbCode)
        optPos = InStr(1, mlfbCode, "-Z ", vbTextCompare)
        If ((optPos > 3) And (optPos < optLen)) Then
            mlfbOpts = Right(mlfbCode, Len(mlfbCode) - optPos - 2)
            mlfbCode = Left(mlfbCode, optPos - 1)   'without -Z
        Else
            mlfbOpts = ""
        End If
        'remove successor note from code
        If (Left(mlfbCode, 8) = ("[succ.]" + vbLf)) Then
            mlfbCode = Right(mlfbCode, Len(mlfbCode) - 8)
        End If
        If (Left(mlfbCode, 7) = "[succ.]") Then
            mlfbCode = Right(mlfbCode, Len(mlfbCode) - 7)
        End If
        'options should be separated by space instead of +
        mlfbOpts = Replace(mlfbOpts, "+", "%20")
        If Len(mlfbCode) > 1 Then
            URL = "https://support.industry.siemens.com/cs/products/" + mlfbCode
            ActiveWorkbook.FollowHyperlink URL
        End If
    End If
EndSub:
    Application.StatusBar = ""
End Sub

