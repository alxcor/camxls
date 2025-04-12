REM  *****  BASIC  *****
'MLFB Spare Parts Availability, LibreOffice v1.4.6 / 12.04.2025 [alxcor:250412]

Sub clearAll()
    'Worksheet 'Data': Clear All
    'Select Data worksheet
    nEndColumn =  29
    nEndRow = 500
    oSheet = ThisComponent.Sheets.getByName("Data")
    ThisComponent.CurrentController.setActiveSheet(oSheet)
    'Select Cell "A1"
    oCell = oSheet.getCellByPosition(0, 0)
    ThisComponent.CurrentController.Select(oCell)
    'Remove Select, keep Focus
    oRanges = ThisComponent.createInstance("com.sun.star.sheet.SheetCellRanges")
    ThisComponent.CurrentController.Select(oRanges)
    'Remove Freeze
    ThisComponent.CurrentController.FreezeAtPosition(0,0)
    DoEvents
    If ((nEndRow > 0) And (nEndColumn > 0)) Then
        'clear data for current row, columns 1 to maxCol
        oSheet.getCellRangeByPosition(0, 0, nEndColumn, nEndRow).clearContents(127)
        oSheet.getCellRangeByPosition(0, 0, nEndColumn, nEndRow).Columns.Width = 2200
        oSheet.getCellRangeByPosition(0, 0, nEndColumn, nEndRow).Rows.OptimalHeight = TRUE
        oSheet.getCellRangeByPosition(0, 0, nEndColumn, nEndRow).HoriJustify = com.sun.star.table.CellHoriJustify.LEFT
        oSheet.getCellRangeByPosition(0, 0, nEndColumn, nEndRow).VertJustify = com.sun.star.table.CellVertJustify.TOP
    End If
    setHeader
End Sub

Sub clearAllR()
    'Worksheet 'Report': Clear All
    'Select Report worksheet
    nEndColumn =  29
    nEndRow       = 500
    oSheet = ThisComponent.Sheets.getByName("Report")
    ThisComponent.CurrentController.setActiveSheet(oSheet)
    'Select Cell "A1"
    oCell = oSheet.getCellByPosition(0, 0)
    ThisComponent.CurrentController.Select(oCell)
    'Remove Select, keep Focus
    oRanges = ThisComponent.createInstance("com.sun.star.sheet.SheetCellRanges")
    ThisComponent.CurrentController.Select(oRanges)
    'Remove Freeze
    ThisComponent.CurrentController.FreezeAtPosition(0,0)
    DoEvents
    'Select All Used Cells In Sheet
    If ((nEndRow > 0) And (nEndColumn > 0)) Then
        'clear data for current row, columns 1 to maxCol
        oSheet.getCellRangeByPosition(0, 0, nEndColumn, nEndRow).clearContents(127)
        oSheet.getCellRangeByPosition(0, 0, nEndColumn, nEndRow).Columns.Width = 2200
        oSheet.getCellRangeByPosition(0, 0, nEndColumn, nEndRow).Rows.OptimalHeight = TRUE
        oSheet.getCellRangeByPosition(0, 0, nEndColumn, nEndRow).HoriJustify = com.sun.star.table.CellHoriJustify.LEFT
        oSheet.getCellRangeByPosition(0, 0, nEndColumn, nEndRow).VertJustify = com.sun.star.table.CellVertJustify.TOP
    End If
    setHeaderR
End Sub

Sub setHeader()
    'Worksheet 'Data': Set header texts on row 1
    Dim rowNumber As Integer
    Dim maxCol As Integer
    rowNumber = 1
    maxCol = 29
    'Select Data worksheet
    oSheet = ThisComponent.Sheets.getByName("Data")
    ThisComponent.CurrentController.setActiveSheet(oSheet)
    DoEvents
    'Select 1st Row and check if the row is free for buttons
    oCell = oSheet.getCellByPosition(0, 0)
    If (oCell.Type <> EMPTY) Then
        'Insert Empty Row for Header
        oSheet.Rows.InsertByIndex(0, 1)
    End If
    'Select 2nd Row and check if the row is free for buttons
    oCell = oSheet.getCellByPosition(0, 1)
    If (oCell.Type <> EMPTY) Then
        'Insert Empty Row for Header
        oSheet.Rows.InsertByIndex(1, 1)
    End If
    'Select 3rd Row and check if the row is free for header:
    oCell = oSheet.getCellByPosition(0, 2)
    If (oCell.Type <> EMPTY) Then
        If (oCell.getString <> "Your Data...") Then
            'Insert Empty Row for Header
            oSheet.Rows.InsertByIndex(2, 1)
        End If
    End If
    DoEvents
    'Clear the 1st, 2nd, 3rd row
    oRange = oSheet.getRows().getByIndex(0)
    oRange.clearContents(127)
    oRange = oSheet.getRows().getByIndex(1)
    oRange.clearContents(127)
    oRange = oSheet.getRows().getByIndex(2)
    oRange.clearContents(127)
    oRange.CharWeight = com.sun.star.awt.FontWeight.BOLD
    'Set borders
    oBorder = CreateUnoStruct("com.sun.star.table.BorderLine")
    oBorder.OuterLineWidth = 50
    oBorder.Color = RGB(124, 124, 124)
    oRange.BottomBorder = oBorder
    'Set texts
    oSheet.getCellByPosition( 0, 2).setString("Your Data...")
    oSheet.getCellByPosition( 1, 2).setString("MLFB")
    oSheet.getCellByPosition( 2, 2).setString("Product Description")
    oSheet.getCellByPosition( 3, 2).setString("Product Family")
    oSheet.getCellByPosition( 4, 2).setString("Product Lifecycle (PLM)")
    oSheet.getCellByPosition( 5, 2).setString("PLM Effective Date")
    oSheet.getCellByPosition( 6, 2).setString("Notes")
    oSheet.getCellByPosition( 7, 2).setString("Price Group")
    oSheet.getCellByPosition( 8, 2).setString("Surcharge for Raw Materials")
    oSheet.getCellByPosition( 9, 2).setString("Metal Factor")
    oSheet.getCellByPosition(10, 2).setString("Export Control Regulations")
    oSheet.getCellByPosition(11, 2).setString("Dispatch Time")
    oSheet.getCellByPosition(12, 2).setString("Net Weight (kg)")
    oSheet.getCellByPosition(13, 2).setString("Product Dimensions (W x L x H)")
    oSheet.getCellByPosition(14, 2).setString("Packaging Dimension")
    oSheet.getCellByPosition(15, 2).setString("Package size unit of measure")
    oSheet.getCellByPosition(16, 2).setString("Quantity Unit")
    oSheet.getCellByPosition(17, 2).setString("Packaging Quantity")
    oSheet.getCellByPosition(18, 2).setString("EAN")
    oSheet.getCellByPosition(19, 2).setString("UPC")
    oSheet.getCellByPosition(20, 2).setString("Commodity Code")
    oSheet.getCellByPosition(21, 2).setString("KZ_FDB/ CatalogID")
    oSheet.getCellByPosition(22, 2).setString("Product Group")
    oSheet.getCellByPosition(23, 2).setString("Country of origin")
    oSheet.getCellByPosition(24, 2).setString("Compliance with the substance restrictions according to RoHS directive")
    oSheet.getCellByPosition(25, 2).setString("Product class")
    oSheet.getCellByPosition(26, 2).setString("Obligation Category for taking back electrical and electronic equipment after use")
    oSheet.getCellByPosition(27, 2).setString("Classifications")
    oSheet.getCellByPosition(28, 2).setString("Successor")
    'Set Column width
    setSize
End Sub

Sub setHeaderR()
    'Worksheet 'Report': Set header texts on row 1
    Dim rowNumber As Integer
    Dim maxCol As Integer
    rowNumber = 1
    maxCol = 5
    nEndRow = 500
    'Select Report worksheet
    oSheet = ThisComponent.Sheets.getByName("Report")
    ThisComponent.CurrentController.setActiveSheet(oSheet)
    DoEvents
    'Select 1st Row and check if the row is free for buttons
    oCell = oSheet.getCellByPosition(0, 0)
    If (oCell.Type <> EMPTY) Then
        'Insert Empty Row for Header
        oSheet.Rows.InsertByIndex(0, 1)
    End If
    'Select 2nd Row and check if the row is free for buttons
    oCell = oSheet.getCellByPosition(0, 1)
    If (oCell.Type <> EMPTY) Then
        'Insert Empty Row for Header
        oSheet.Rows.InsertByIndex(1, 1)
    End If
    'Select 3rd Row and check if the row is free for header:
    oCell = oSheet.getCellByPosition(0, 2)
    If (oCell.Type <> EMPTY) Then
        If (oCell.getString <> "MLFB") Then
            'Insert Empty Row for Header
            oSheet.Rows.InsertByIndex(2, 1)
        End If
    End If
    'Select 4th Row and check if the row is free for header:
    oCell = oSheet.getCellByPosition(0, 3)
    If (oCell.Type <> EMPTY) Then
        If (Left(oCell.getString, 6) <> "Active") Then
            'Insert Empty Row for Header
            oSheet.Rows.InsertByIndex(3, 1)
        End If
    End If
    DoEvents
    'Clear the 1st, 2nd, 3rd row
    oRange = oSheet.getRows().getByIndex(0)
    oRange.clearContents(127)
    oRange = oSheet.getRows().getByIndex(1)
    oRange.clearContents(127)
    oRange = oSheet.getRows().getByIndex(2)
    oRange.clearContents(127)
    oRange.CharWeight = com.sun.star.awt.FontWeight.BOLD
    'Set borders
    oBorder = CreateUnoStruct("com.sun.star.table.BorderLine")
    oBorder.OuterLineWidth = 50
    oBorder.Color = RGB(124, 124, 124)
    oSheet.getCellRangeByPosition(0, 2, 4, 2).BottomBorder = oBorder
    oSheet.getCellRangeByPosition(0, 2, 4, 3).CharWeight = com.sun.star.awt.FontWeight.BOLD
    oSheet.getCellRangeByPosition(0, 2, 4, 3).CharHeight = 10
    oSheet.getCellRangeByPosition(0, 4, 4, nEndRow).CharWeight = com.sun.star.awt.FontWeight.NORMAL
    oSheet.getCellRangeByPosition(0, 4, 4, nEndRow).CharHeight = 8
    'Set texts
    oSheet.getCellByPosition( 0, 2).setString("MLFB")
    oSheet.getCellByPosition( 1, 2).setString("Product Description")
    oSheet.getCellByPosition( 2, 2).setString("Product Lifecycle (PLM)")
    oSheet.getCellByPosition( 3, 2).setString("Notes")
    oSheet.getCellByPosition( 4, 2).setString("Dispatch Time")
    'Format cells
    oSheet.getCellRangeByPosition(0, 2, 4, nEndRow).isTextWrapped = TRUE
    oSheet.getCellRangeByPosition(0, 2, 0, nEndRow).Columns.Width = 3600
    oSheet.getCellRangeByPosition(1, 2, 1, nEndRow).Columns.Width = 4000
    oSheet.getCellRangeByPosition(2, 2, 2, nEndRow).Columns.Width = 4000
    oSheet.getCellRangeByPosition(3, 2, 3, nEndRow).Columns.Width = 4000
    oSheet.getCellRangeByPosition(4, 2, 4, nEndRow).Columns.Width = 1600
    oSheet.getCellByPosition(0, 3).CellBackColor = RGB(125, 242, 92)
    oSheet.getCellByPosition(1, 3).CellBackColor = RGB(229, 242, 80)
    oSheet.getCellByPosition(2, 3).CellBackColor = RGB(242, 135, 148)
    oSheet.getCellByPosition(3, 3).CellBackColor = RGB(230, 230, 230)
    oSheet.getCellByPosition(4, 3).CellBackColor = RGB(230, 230, 230)
    oSheet.getCellRangeByPosition(0, 0, 0, nEndRow).Rows.OptimalHeight = TRUE
    oSheet.getCellRangeByPosition(0, 0, maxCol, nEndRow).VertJustify = com.sun.star.table.CellVertJustify.TOP
    'Format report
    FormatReport
End Sub

Sub setSize()
    'Worksheet 'Data': Set Row/Column size
    nEndColumn =  29
    nEndRow = 500
    'Select Data worksheet
    oSheet = ThisComponent.Sheets.getByName("Data")
    ThisComponent.CurrentController.setActiveSheet(oSheet)
    DoEvents
    oSheet.getCellRangeByPosition(0, 2, 1, nEndRow).isTextWrapped = TRUE
    oSheet.getCellRangeByPosition(0, 2, 1, nEndRow).isTextWrapped = TRUE
    'oSheet.getCellRangeByPosition(0, 2, 0, nEndRow).Columns.OptimalWidth = TRUE
    'oSheet.getCellRangeByPosition(1, 2, 1, nEndRow).Columns.OptimalWidth = TRUE
    oSheet.getCellRangeByPosition(2, 2, 28, nEndRow).isTextWrapped = TRUE
    oSheet.getCellRangeByPosition(0, 2, 0, nEndRow).Columns.Width = 4000
    oSheet.getCellRangeByPosition(1, 2, 1, nEndRow).Columns.Width = 4000
    oSheet.getCellRangeByPosition(2, 2, 2, nEndRow).Columns.Width = 8000
    oSheet.getCellRangeByPosition(3, 2, 3, nEndRow).Columns.Width = 4800
    oSheet.getCellRangeByPosition(4, 2, 4, nEndRow).Columns.Width = 4800
    oSheet.getCellRangeByPosition(5, 2, 5, nEndRow).Columns.Width = 3600
    oSheet.getCellRangeByPosition(6, 2, 6, nEndRow).Columns.Width = 8000
    oSheet.getCellRangeByPosition(7, 2, 7, nEndRow).Columns.Width = 2400
    oSheet.getCellRangeByPosition(8, 2, 8, nEndRow).Columns.Width = 6000
    oSheet.getCellRangeByPosition(9, 2, 9, nEndRow).Columns.Width = 2400
    oSheet.getCellRangeByPosition(10, 2, 10, nEndRow).Columns.Width = 5200
    oSheet.getCellRangeByPosition(11, 2, 11, nEndRow).Columns.Width = 2800
    oSheet.getCellRangeByPosition(12, 2, 12, nEndRow).Columns.Width = 3200
    oSheet.getCellRangeByPosition(13, 2, 13, nEndRow).Columns.Width = 6000
    oSheet.getCellRangeByPosition(14, 2, 14, nEndRow).Columns.Width = 4400
    oSheet.getCellRangeByPosition(15, 2, 15, nEndRow).Columns.Width = 5600
    oSheet.getCellRangeByPosition(16, 2, 16, nEndRow).Columns.Width = 2400
    oSheet.getCellRangeByPosition(17, 2, 17, nEndRow).Columns.Width = 4000
    oSheet.getCellRangeByPosition(18, 2, 18, nEndRow).Columns.Width = 3200
    oSheet.getCellRangeByPosition(19, 2, 19, nEndRow).Columns.Width = 3200
    oSheet.getCellRangeByPosition(20, 2, 20, nEndRow).Columns.Width = 3200
    oSheet.getCellRangeByPosition(21, 2, 21, nEndRow).Columns.Width = 3200
    oSheet.getCellRangeByPosition(22, 2, 22, nEndRow).Columns.Width = 3200
    oSheet.getCellRangeByPosition(23, 2, 23, nEndRow).Columns.Width = 3200
    oSheet.getCellRangeByPosition(24, 2, 24, nEndRow).Columns.Width = 8000
    oSheet.getCellRangeByPosition(25, 2, 25, nEndRow).Columns.Width = 8000
    oSheet.getCellRangeByPosition(26, 2, 26, nEndRow).Columns.Width = 8000
    oSheet.getCellRangeByPosition(27, 2, 27, nEndRow).Columns.Width = 8000
    oSheet.getCellRangeByPosition(28, 2, 28, nEndRow).Columns.Width = 8000
    oSheet.getCellRangeByPosition(0, 0, 0, nEndRow).Rows.OptimalHeight = TRUE
    oSheet.getCellRangeByPosition(0, 0, nEndColumn, nEndRow).VertJustify = com.sun.star.table.CellVertJustify.TOP
    ThisComponent.CurrentController.FreezeAtPosition(3,3)
End Sub

Sub setCells(rowNumber)
    'clear data for a range of cells in a row
    Dim maxCol As Integer
    Dim mlfbCode As String
    'last column is AC = column 29
    maxCol = 29
    oSheet = ThisComponent.CurrentController.getActiveSheet()
    mlfbCode = oSheet.getCellByPosition(0, rowNumber).getString()
    'clear data for current row, columns 1 to maxCol
    oSheet.getCellRangeByPosition(0, rowNumber, maxCol, rowNumber).clearContents(127)
    oSheet.getCellByPosition(0, rowNumber).setString(mlfbCode)
End Sub

Sub readRow
     'read data for current row: on column 1 [A] should be a product code (MLFB) from Industry Mall web site
    Dim rowNumber As Long
    Dim mlfbCode As String
    Dim isSuccessor As Boolean
    '--------------------------------------------
    'Select Data worksheet
    oSheet = ThisComponent.Sheets.getByName("Data")
    ThisComponent.CurrentController.setActiveSheet(oSheet)
    DoEvents
    '--------------------------------------------
    oCurrentSelection = ThisComponent.getCurrentSelection()
    If oCurrentSelection.supportsService("com.sun.star.sheet.SheetCell") Then
        rowNumber = oCurrentSelection.getCellAddress().Row
    Else
        rowNumber = 0
    End If
    If rowNumber < 3 Then
        MsgBox ("[EN]: Table starts on row 4; [RO]:Tabelul incepe de la randul 4!")
        GoTo EndSub
    Else
        '----------------------------------------
        Call setCells(rowNumber)
        '----------------------------------------
        mlfbCode = oSheet.getCellByPosition(0, rowNumber).getString()
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
                mlfbCode = oSheet.getCellByPosition(1, rowNumber).getString()
                oSheet.getCellByPosition(1, rowNumber).setString("[succ.]" & CHR$(10) & mlfbCode)
            End If
        End If
    End If
    setHeader
    setSize
EndSub:
End Sub

Sub readRowR
    'read data starting from 'Report' worksheet
    Dim rowNumber As Long
    Dim mlfbCode As String
    '--------------------------------------------
    'Select Report worksheet
    oSheet = ThisComponent.Sheets.getByName("Report")
    ThisComponent.CurrentController.setActiveSheet(oSheet)
    DoEvents
    '--------------------------------------------
    oCurrentSelection = ThisComponent.getCurrentSelection()
    If oCurrentSelection.supportsService("com.sun.star.sheet.SheetCell") Then
        rowNumber = oCurrentSelection.getCellAddress().Row
    Else
        rowNumber = 0
    End If
    If rowNumber < 4 Then
        MsgBox ("[EN]: Table starts on row 5; [RO]:Tabelul incepe de la randul 5!")
        GoTo EndSub
    End If
    mlfbCode = oSheet.getCellByPosition(0, rowNumber).getString()
    Call setCells(rowNumber)
    oSheet = ThisComponent.Sheets.getByName("Data")
    ThisComponent.CurrentController.setActiveSheet(oSheet)
    DoEvents
    setHeader
    DoEvents
    oSheet.getCellByPosition(0, rowNumber - 1).setString(mlfbCode)
    ThisComponent.CurrentController.Select(oSheet.getCellByPosition(0, rowNumber - 1))
    readRow
    reportRow
    oSheet = ThisComponent.Sheets.getByName("Report")
    ThisComponent.CurrentController.setActiveSheet(oSheet)
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
    oSheet = ThisComponent.Sheets.getByName("Data")
    ThisComponent.CurrentController.setActiveSheet(oSheet)
    DoEvents
    '--------------------------------------------
    setHeader
    '--------------------------------------------
    For rowNumber = 3 To 500
        Call setCells(rowNumber)
        '----------------------------------------
        mlfbCode = oSheet.getCellByPosition(0, rowNumber).getString()
        isSuccessor = False
        If Len(mlfbCode) > 1 Then
            'remove successor note from code
            If (Left(mlfbCode, 8) = ("[succ.]" + CHR$(10))) Then
                mlfbCode = Right(mlfbCode, Len(mlfbCode) - 8)
               isSuccessor = True
            End If
            If (Left(mlfbCode, 7) = "[succ.]") Then
                mlfbCode = Right(mlfbCode, Len(mlfbCode) - 7)
               isSuccessor = True
            End If
            ThisComponent.CurrentController.Select(oSheet.getCellByPosition(0, rowNumber))
            Call ImportSieMallIntra(mlfbCode, rowNumber)
            If (isSuccessor = True) Then
                'Add successor note to column2
                mlfbCode = oSheet.getCellByPosition(1, rowNumber).getString()
                oSheet.getCellByPosition(1, rowNumber).setString("[succ.]" & CHR$(10) & mlfbCode)
            End If
        End If
        DoEvents
    Next
    '--------------------------------------------
    setHeader
    setSize
    '--------------------------------------------
    ThisComponent.CurrentController.FreezeAtPosition(3,3)
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
    oSheet = ThisComponent.Sheets.getByName("Data")
    oSheet2 = ThisComponent.Sheets.getByName("Report")
    ThisComponent.CurrentController.setActiveSheet(oSheet)
    DoEvents
    '--------------------------------------------
    setHeader
    '--------------------------------------------
    mlfbCode = ""
    For rowNumber = 3 To 500
        mlfbCode = oSheet2.getCellByPosition(0, rowNumber + 1).getString()
        oSheet.getCellByPosition(0, rowNumber).setString(mlfbCode)
        DoEvents
    Next
    readAll
    Report
    'MsgBox("Ready...")
End Sub

Sub checkSuccessor()
    'check successor; if a successor is found in column 29 a new row is added and data is read
    Dim rowNumber As Long
    Dim mlfbCode As String
    '--------------------------------------------
    'Select Data worksheet
    oSheet = ThisComponent.Sheets.getByName("Data")
    ThisComponent.CurrentController.setActiveSheet(oSheet)
    DoEvents
    '--------------------------------------------
    oCurrentSelection = ThisComponent.getCurrentSelection()
    If oCurrentSelection.supportsService("com.sun.star.sheet.SheetCell") Then
        rowNumber = oCurrentSelection.getCellAddress().Row
    Else
        rowNumber = 0
    End If
    If rowNumber < 2 Then
        MsgBox ("[EN]: Table starts on row 3; [RO]:Tabelul incepe de la randul 3!")
        GoTo EndSub
    Else
        mlfbCode = Trim(oSheet.getCellByPosition(28, rowNumber).getString())
        nextCode = Trim(oSheet.getCellByPosition(0, rowNumber + 1).getString())
        If (("[succ.]" & CHR$(10) & mlfbCode <> nextCode) And (mlfbCode <> nextCode)) Then
            If (mlfbCode <> "") Then
                oSheet.Rows.InsertByIndex(rowNumber + 1, 1)
                oSheet.getCellByPosition(0, rowNumber + 1).setString("[succ.]" + CHR$(10) + mlfbCode)
                setCells (rowNumer + 1)
                Call ImportSieMallIntra(mlfbCode, rowNumber + 1)
                'Add successor note to column2
                mlfbCode = oSheet.getCellByPosition(1, rowNumber + 1).getString()
                oSheet.getCellByPosition(1, rowNumber + 1).setString("[succ.]" & CHR$(10) & mlfbCode)
                'Select Cell and remove focus
                oCell = oSheet.getCellByPosition(0, rowNumber + 1)
                ThisComponent.CurrentController.Select(oCell)
                oRanges = ThisComponent.createInstance("com.sun.star.sheet.SheetCellRanges")
                ThisComponent.CurrentController.Select(oRanges)
            End If
        End If
    End If
    '--------------------------------------------
    setHeader
    setSize
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
    oSheet = ThisComponent.Sheets.getByName("Data")
    ThisComponent.CurrentController.setActiveSheet(oSheet)
    DoEvents
    '--------------------------------------------
    setHeader
    '--------------------------------------------
    rowNumber = 3
    While (rowNumber <= maxRow)
        oCell = oSheet.getCellByPosition(0, rowNumber)
        ThisComponent.CurrentController.Select(oCell)
        DoEvents
        mlfbCode = Trim(oSheet.getCellByPosition(28, rowNumber).getString())
        nextCode = Trim(oSheet.getCellByPosition(0, rowNumber + 1).getString())
        If (("[succ.]" & CHR$(10) & mlfbCode <> nextCode) And (mlfbCode <> nextCode)) Then
            If (mlfbCode <> "") Then
                oSheet.Rows.InsertByIndex(rowNumber + 1, 1)
                oSheet.getCellByPosition(0, rowNumber + 1).setString("[succ.]" + CHR$(10) + mlfbCode)
                setCells (rowNumer + 1)
                Call ImportSieMallIntra(mlfbCode, rowNumber + 1)
                'Add successor note to column2
                mlfbCode = oSheet.getCellByPosition(1, rowNumber + 1).getString()
                oSheet.getCellByPosition(1, rowNumber + 1).setString("[succ.]" & CHR$(10) & mlfbCode)
                'Select Cell and remove focus
                oCell = oSheet.getCellByPosition(0, rowNumber + 1)
                ThisComponent.CurrentController.Select(oCell)
                oRanges = ThisComponent.createInstance("com.sun.star.sheet.SheetCellRanges")
                ThisComponent.CurrentController.Select(oRanges)
            End If
        End If
        rowNumber = rowNumber + 1
    Wend
    '--------------------------------------------
    setHeader
    setSize
    '--------------------------------------------
    oCell = oSheet.getCellByPosition(0, 3)
    ThisComponent.CurrentController.Select(oCell)
    oRanges = ThisComponent.createInstance("com.sun.star.sheet.SheetCellRanges")
    ThisComponent.CurrentController.Select(oRanges)
    'ActiveWindow.FreezePanes = True
EndSub:   
End Sub

Sub ImportSieMallIntra(mlfbCode, rowNumber)
    'read data for a specific product code (MLFB) from Industry Mall web site
    'netMode = xmlHTTP version
    'On Error GoTo ErrHand:   'disable this line to see what is the error
    Dim targetURL As String
    Dim webContent As String
    Dim index As Integer
    Dim DetailNo As Integer
    Dim Product As Object
    '--------------------------------------------
    'Select Data worksheet
    oSheet = ThisComponent.Sheets.getByName("Data")
    ThisComponent.CurrentController.setActiveSheet(oSheet)
    DoEvents
    '--------------------------------------------
    oSheet.getCellByPosition(1, rowNumber).setString(mlfbCode)
    oSheet.getCellByPosition(4, rowNumber).setString("ERR: Not Found!!!")
    oSheet.getCellByPosition(4, rowNumber).CellBackColor = RGB(242, 135, 148)
    'Reading web page in buffer...
    'options should be separated by space instead of +
    mlfbCode = Replace(mlfbCode, "+", "%20")
    'format spaces html style
    mlfbCode = Replace(mlfbCode, " ", "%20")
    'clear front and back spaces
    mlfbCode = Replace(mlfbCode, "%20", " ")
    mlfbCode = Trim(mlfbCode)
    mlfbCode = Replace(mlfbCode, " ", "%20")
    'set web page (for scrapper)
    targetURL = "https://mall.industry.siemens.com/mall/en/WW/Catalog/Product/" + mlfbCode
    oSimpleFileAccess = createUNOService ("com.sun.star.ucb.SimpleFileAccess")
    oInpDataStream = createUNOService ("com.sun.star.io.TextInputStream")
    oInpDataStream.setInputStream(oSimpleFileAccess.openFileRead(targetURL))
    Dim delimiters() as Long
    sContent = oInpDataStream.readString(delimiters(), false)
    'identyfy MLFB by productidentifier
    result = ""
    lStartPos = instr(1, sContent, "<span class=" & chr(34) & "productidentifier" )
    If lStartPos = 0 Then
        oSheet.getCellByPosition(4, rowNumber).setString("ERR: Not Found!!!")
        GoTo EndSub
    End If
    lStartPos = lStartPos + 32
    lEndPos = instr(lStartPos, sContent, "</span>")
    sTable = mid(sContent, lStartPos, lEndPos-lStartPos)
    result = sTable
    result =  ClearData(result)   
    oSheet.getCellByPosition(1, rowNumber).setString(result)
    'identyfy MLFB by detailsPageHeader
    lStartPos = instr(1, sContent, "<span class=" & chr(34) & "detailsPageHeader" )
    If lStartPos > 32 Then
        lStartPos = lStartPos + 32
        lEndPos = instr(lStartPos, sContent, "</span>")
        sTable = mid(sContent, lStartPos, lEndPos-lStartPos)
        result = sTable
        oSheet.getCellByPosition(1, rowNumber).setString(result)
    End If
    'Identify Details
    result = IdentifyData(sContent, "Product Description")
    oSheet.getCellByPosition(2, rowNumber).setString(result)
    result = IdentifyData(sContent, "Product family")
    oSheet.getCellByPosition(3, rowNumber).setString(result)
    result = IdentifyData(sContent, "Product Lifecycle (PLM)")
    oSheet.getCellByPosition(4, rowNumber).setString(result)
    'PLM status:
    iTemp = InStr(1, result, "M250", 1)
    If iTemp > 0 Then
        oSheet.getCellByPosition(4, rowNumber).CellBackColor = RGB(125, 242, 92)
        GoTo EndPLMstatus
    End If
    iTemp = InStr(1, result, "M280", 1)
    If iTemp > 0 Then
        oSheet.getCellByPosition(4, rowNumber).CellBackColor = RGB(125, 242, 92)
        GoTo EndPLMstatus
    End If
    iTemp = InStr(1, result, "M300", 1)
    If iTemp > 0 Then
        oSheet.getCellByPosition(4, rowNumber).CellBackColor = RGB(125, 242, 92)
        GoTo EndPLMstatus
    End If
    iTemp = InStr(1, result, "M400", 1)
    If iTemp > 0 Then
        oSheet.getCellByPosition(4, rowNumber).CellBackColor = RGB(229, 242, 80)
        GoTo EndPLMstatus
    End If
    iTemp = InStr(1, result, "M410", 1)
    If iTemp > 0 Then
        oSheet.getCellByPosition(4, rowNumber).CellBackColor = RGB(229, 242, 80)
        GoTo EndPLMstatus
    End If
    iTemp = InStr(1, result, "M490", 1)
    If iTemp > 0 Then
        oSheet.getCellByPosition(4, rowNumber).CellBackColor = RGB(242, 135, 148)
        GoTo EndPLMstatus
    End If
    iTemp = InStr(1, result, "M500", 1)
    If iTemp > 0 Then
        oSheet.getCellByPosition(4, rowNumber).CellBackColor = RGB(242, 135, 148)
        GoTo EndPLMstatus
    End If
    EndPLMstatus:
    result = IdentifyData(sContent, "PLM Effective Date")
    oSheet.getCellByPosition(5, rowNumber).setString(result)
    result = IdentifyData(sContent, "Notes")
    oSheet.getCellByPosition(6, rowNumber).setString(result)
    If Len(result) > 0 Then
        oSheet.getCellByPosition(6, rowNumber).CellBackColor = RGB(91, 155, 213)
    End If
    result = IdentifyData(sContent, "Price Group")
    oSheet.getCellByPosition(7, rowNumber).setString(result)
    result = IdentifyData(sContent, "Region Specific PriceGroup / Headquarter Price Group")
    oSheet.getCellByPosition(7, rowNumber).setString(result)
    result = IdentifyData(sContent, "Surcharge for Raw Materials")
    oSheet.getCellByPosition(8, rowNumber).setString(result)
    result = IdentifyData(sContent, "Metal Factor")
    oSheet.getCellByPosition(9, rowNumber).setString(result)
    result = IdentifyData(sContent, "Export Control Regulations")
    oSheet.getCellByPosition(10, rowNumber).setString(result)
    result = IdentifyData(sContent, "Delivery Time")
    oSheet.getCellByPosition(11, rowNumber).setString(result)
    result = IdentifyData(sContent, "Standard lead time ex-works")
    oSheet.getCellByPosition(11, rowNumber).setString(result)
    result = IdentifyData(sContent, "Estimated dispatch time (Working Days)")
    oSheet.getCellByPosition(11, rowNumber).setString(result)
    result = IdentifyData(sContent, "Net Weight(kg)")
    oSheet.getCellByPosition(12, rowNumber).setString(result)
    result = IdentifyData(sContent, "Net Weight (kg)")
    oSheet.getCellByPosition(12, rowNumber).setString(result)
    result = IdentifyData(sContent, "Product Dimensions (W x L x H)")
    oSheet.getCellByPosition(13, rowNumber).setString(result)
    result = IdentifyData(sContent, "Packaging Dimension")
    oSheet.getCellByPosition(14, rowNumber).setString(result)
    result = IdentifyData(sContent, "Package size unit of measure")
    oSheet.getCellByPosition(15, rowNumber).setString(result)
    result = IdentifyData(sContent, "Quantity Unit")
    oSheet.getCellByPosition(16, rowNumber).setString(result)
    result = IdentifyData(sContent, "Packaging Quantity")
    oSheet.getCellByPosition(17, rowNumber).setString(result)
    result = IdentifyData(sContent, "EAN")
    oSheet.getCellByPosition(18, rowNumber).setString(result)
    result = IdentifyData(sContent, "UPC")
    oSheet.getCellByPosition(19, rowNumber).setString(result)
    result = IdentifyData(sContent, "Commodity Code")
    oSheet.getCellByPosition(20, rowNumber).setString(result)
    result = IdentifyData(sContent, "LKZ_FDB/ CatalogID")
    oSheet.getCellByPosition(21, rowNumber).setString(result)
    result = IdentifyData(sContent, "Product Group")
    oSheet.getCellByPosition(22, rowNumber).setString(result)
    result = IdentifyData(sContent, "Country of origin")
    oSheet.getCellByPosition(23, rowNumber).setString(result)
    result = IdentifyData(sContent, "Compliance with the substance restrictions according to RoHS directive")
    oSheet.getCellByPosition(24, rowNumber).setString(result)
    result = IdentifyData(sContent, "Product class")
    oSheet.getCellByPosition(25, rowNumber).setString(result)
    result = IdentifyData(sContent, "Obligation Category for taking back electrical and electronic equipment after use")
    oSheet.getCellByPosition(26, rowNumber).setString(result)
    result = IdentifyData(sContent, "Classifications")
    oSheet.getCellByPosition(27, rowNumber).setString(result)
    result = IdentifyData(sContent, "Successor")
    oSheet.getCellByPosition(28, rowNumber).setString(result)
    GoTo EndSub
ErrHand:
    'oSheet.getCellByPosition(2, 5).setString("Error! " & Err.Description)
    oSheet.getCellByPosition(1, 0).setString("Error! ")
EndSub:
End Sub

Function IdentifyData(sContent as String, sDetail as String) as String
    'returns detail_text column located after a detail_name column in html
    result = ""
    lStartOffset = 1
    startCaut:
    'search for detail name in details table
    lStartPos = instr(lStartOffset, sContent, "<td class=" & chr(34) & "productDetailsTable_DataLabel")
    If lStartPos < 42 Then
        GoTo isNotFound
    End If
    lStartPos = lStartPos + 42
    lEndPos = instr(lStartPos, sContent, "</td>")
    result = mid(sContent, lStartPos, lEndPos-lStartPos)
    If (result = sDetail) Then
        'detail found, try to read next column
        lStartOffset = lStartPos
        lStartPos = instr(lStartOffset, sContent, "<td>")
        If lStartPos < 4 Then
            GoTo isNotFound
        End If
        lStartPos = lStartPos + 4
        lEndPos = instr(lStartPos, sContent, "</td>")
        result = mid(sContent, lStartPos, lEndPos-lStartPos)
        result = Replace(result, CHR$(13), " ")
        result = Replace(result, CHR$(10), " ")
        result = Trim(result)
        IdentifyData =  ClearData(result)
    Else
        lStartOffset = lStartPos
        GoTo startCaut
    End If
    GoTo EndFunc
    isNotFound:
        IdentifyData =  ""
        GoTo EndFunc   
EndFunc:
End Function

Function ClearData(sContent as String) as String
    'returns text without html attributes
    result = sContent
    lStartOffset = 1
    lStartPos = 0
    lEndPos = 0
    startCaut:
    'search for attributes in text
    lStartPos = instr(lStartOffset, result, "<")
    If lStartPos < 1 Then
        GoTo isNotFound
    End If
    lEndPos = instr(lStartPos, result, ">")
    result = Trim(left(result, lStartPos - 1)) & " " & Trim(right(result, Len(result) - lEndPos))
    GoTo startCaut
    isNotFound:
    result = Replace(result, "&nbsp;", " ")
    result = Replace(result, "&amp;", "&")
    ClearData = result
End Function

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
    oBorder = CreateUnoStruct("com.sun.star.table.BorderLine")
    oBorder.OuterLineWidth = 10
    oBorder.Color = RGB(124, 124, 124)
    oSheetR = ThisComponent.Sheets.getByName("Report")
    oSheetD = ThisComponent.Sheets.getByName("Data")
    ThisComponent.CurrentController.setActiveSheet(oSheetR)
    'Clear all data in Report worksheet
    ClearAllR
    DoEvents
    'Write header on first row
    setHeaderR
    rowNumberR = 4
    For rowNumber = 3 To maxRow
        'Spare part availability ignored
        partPM = 0
        mlfbCode = Trim(oSheetD.getCellByPosition(1, rowNumber).getString())
        If (mlfbCode <> "") Then
            'Cells(rowNumberR, 1).Select
            'Spare part availability not yet established
            partPM = 1
            'Format cells
            oSheetR.getCellRangeByPosition(0, rowNumberR, 4, rowNumberR).BottomBorder = oBorder
            oSheetR.getCellRangeByPosition(0, rowNumberR, 4, rowNumberR).LeftBorder = oBorder
            oSheetR.getCellRangeByPosition(0, rowNumberR, 4, rowNumberR).RightBorder = oBorder
            oSheetR.getCellRangeByPosition(0, rowNumberR, 4, rowNumberR).CharWeight = com.sun.star.awt.FontWeight.NORMAL
            oSheetR.getCellRangeByPosition(0, rowNumberR, 4, rowNumberR).CharHeight = 8
            oSheetR.getCellByPosition(0, rowNumberR).setString(mlfbCode)
            mlfbCode = Trim(oSheetD.getCellByPosition(2, rowNumber).getString())
            oSheetR.getCellByPosition(1, rowNumberR).setString(mlfbCode)
            mlfbCode = Trim(oSheetD.getCellByPosition(4, rowNumber).getString())
            mlfbCode = mlfbCode & CHR$(10) & Trim(oSheetD.getCellByPosition(5, rowNumber).getString())
            oSheetR.getCellByPosition(2, rowNumberR).setString(mlfbCode)
            mlfbCode = Trim(oSheetD.getCellByPosition(6, rowNumber).getString())
            oSheetR.getCellByPosition(3, rowNumberR).setString(mlfbCode)
            mlfbCode = Trim(oSheetD.getCellByPosition(11, rowNumber).getString())
            oSheetR.getCellByPosition(4, rowNumberR).setString(mlfbCode)
            'PLM:
            oSheetR.getCellByPosition(2, rowNumberR).CellBackColor = RGB(230, 230, 230)
            mlfbCode = Trim(oSheetD.getCellByPosition(4, rowNumber).getString())
            iTemp = InStr(1, mlfbCode, "M250", 1)
            If iTemp > 0 Then
                partPM = 250
                oSheetR.getCellByPosition(2, rowNumberR).CellBackColor = RGB(125, 242, 92)
                GoTo EndPLMstatus
            End If
            iTemp = InStr(1, mlfbCode, "M280", 1)
            If iTemp > 0 Then
                partPM = 280
                oSheetR.getCellByPosition(2, rowNumberR).CellBackColor = RGB(125, 242, 92)
                GoTo EndPLMstatus
            End If
            iTemp = InStr(1, mlfbCode, "M300", 1)
            If iTemp > 0 Then
                partPM = 300
                oSheetR.getCellByPosition(2, rowNumberR).CellBackColor = RGB(125, 242, 92)
                GoTo EndPLMstatus
            End If
            iTemp = InStr(1, mlfbCode, "M400", 1)
            If iTemp > 0 Then
                partPM = 400
                oSheetR.getCellByPosition(2, rowNumberR).CellBackColor = RGB(229, 242, 80)
                GoTo EndPLMstatus
            End If
            iTemp = InStr(1, mlfbCode, "M410", 1)
            If iTemp > 0 Then
                partPM = 410
                oSheetR.getCellByPosition(2, rowNumberR).CellBackColor = RGB(229, 242, 80)
                GoTo EndPLMstatus
            End If
            iTemp = InStr(1, mlfbCode, "M490", 1)
            If iTemp > 0 Then
                partPM = 490
                oSheetR.getCellByPosition(2, rowNumberR).CellBackColor = RGB(242, 135, 148)
                GoTo EndPLMstatus
            End If
            iTemp = InStr(1, mlfbCode, "M500", 1)
            If iTemp > 0 Then
                partPM = 500
                oSheetR.getCellByPosition(2, rowNumberR).CellBackColor = RGB(242, 135, 148)
                GoTo EndPLMstatus
            End If
        End If
        EndPLMstatus:
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
    oSheetR.getCellByPosition(0, 3).setString("Active: " & CStr(partsOK))
    oSheetR.getCellByPosition(1, 3).setString("PhaseOut: " & CStr(partsAT))
    oSheetR.getCellByPosition(2, 3).setString("Disc: " & CStr(partsER))
    oSheetR.getCellByPosition(3, 3).setString("Other: " & CStr(partsNA))
    'Format report worksheet
    oSheetR.getCellRangeByPosition(0, rowNumberR, 4, rowNumberR).Rows.OptimalHeight = TRUE
    oSheetR.getCellRangeByPosition(0, rowNumberR, 4, rowNumberR).HoriJustify = com.sun.star.table.CellHoriJustify.LEFT
    oSheetR.getCellRangeByPosition(0, rowNumberR, 4, rowNumberR).VertJustify = com.sun.star.table.CellVertJustify.TOP
    FormatReport
End Sub

Sub ReportRow()
    'generate a printable report row for current row
    Dim rowNumber As Long
    Dim rowNumberR As Long
    Dim mlfbCode As String
    '--------------------------------------------
    oCurrentSelection = ThisComponent.getCurrentSelection()
    If oCurrentSelection.supportsService("com.sun.star.sheet.SheetCell") Then
        rowNumber = oCurrentSelection.getCellAddress().Row
    Else
        rowNumber = 0
    End If
    If rowNumber < 4 Then
        GoTo EndSub
    End If
    oBorder = CreateUnoStruct("com.sun.star.table.BorderLine")
    oBorder.OuterLineWidth = 10
    oBorder.Color = RGB(124, 124, 124)
    oSheetR = ThisComponent.Sheets.getByName("Report")
    oSheetD = ThisComponent.Sheets.getByName("Data")
    ThisComponent.CurrentController.setActiveSheet(oSheetR)
    DoEvents
    'Write header on first row
    setHeaderR
    rowNumberR = rowNumber + 1
        'Spare part availability ignored
        partPM = 0
        mlfbCode = Trim(oSheetD.getCellByPosition(1, rowNumber).getString())
        If (mlfbCode <> "") Then
            'Spare part availability not yet established
            partPM = 1
            'Format cells
            oSheetR.getCellRangeByPosition(0, rowNumberR, 4, rowNumberR).BottomBorder = oBorder
            oSheetR.getCellRangeByPosition(0, rowNumberR, 4, rowNumberR).LeftBorder = oBorder
            oSheetR.getCellRangeByPosition(0, rowNumberR, 4, rowNumberR).RightBorder = oBorder
            oSheetR.getCellRangeByPosition(0, rowNumberR, 4, rowNumberR).CharWeight = com.sun.star.awt.FontWeight.NORMAL
            oSheetR.getCellRangeByPosition(0, rowNumberR, 4, rowNumberR).CharHeight = 8
            oSheetR.getCellByPosition(0, rowNumberR).setString(mlfbCode)
            mlfbCode = Trim(oSheetD.getCellByPosition(2, rowNumber).getString())
            oSheetR.getCellByPosition(1, rowNumberR).setString(mlfbCode)
            mlfbCode = Trim(oSheetD.getCellByPosition(4, rowNumber).getString())
            mlfbCode = mlfbCode & CHR$(10) & Trim(oSheetD.getCellByPosition(5, rowNumber).getString())
            oSheetR.getCellByPosition(2, rowNumberR).setString(mlfbCode)
            mlfbCode = Trim(oSheetD.getCellByPosition(6, rowNumber).getString())
            oSheetR.getCellByPosition(3, rowNumberR).setString(mlfbCode)
            mlfbCode = Trim(oSheetD.getCellByPosition(11, rowNumber).getString())
            oSheetR.getCellByPosition(4, rowNumberR).setString(mlfbCode)
            'PLM:
            oSheetR.getCellByPosition(2, rowNumberR).CellBackColor = RGB(230, 230, 230)
            mlfbCode = Trim(oSheetD.getCellByPosition(4, rowNumber).getString())
            iTemp = InStr(1, mlfbCode, "M250", 1)
            If iTemp > 0 Then
                partPM = 250
                oSheetR.getCellByPosition(2, rowNumberR).CellBackColor = RGB(125, 242, 92)
                GoTo EndPLMstatus
            End If
            iTemp = InStr(1, mlfbCode, "M280", 1)
            If iTemp > 0 Then
                partPM = 280
                oSheetR.getCellByPosition(2, rowNumberR).CellBackColor = RGB(125, 242, 92)
                GoTo EndPLMstatus
            End If
            iTemp = InStr(1, mlfbCode, "M300", 1)
            If iTemp > 0 Then
                partPM = 300
                oSheetR.getCellByPosition(2, rowNumberR).CellBackColor = RGB(125, 242, 92)
                GoTo EndPLMstatus
            End If
            iTemp = InStr(1, mlfbCode, "M400", 1)
            If iTemp > 0 Then
                partPM = 400
                oSheetR.getCellByPosition(2, rowNumberR).CellBackColor = RGB(229, 242, 80)
                GoTo EndPLMstatus
            End If
            iTemp = InStr(1, mlfbCode, "M410", 1)
            If iTemp > 0 Then
                partPM = 410
                oSheetR.getCellByPosition(2, rowNumberR).CellBackColor = RGB(229, 242, 80)
                GoTo EndPLMstatus
            End If
            iTemp = InStr(1, mlfbCode, "M490", 1)
            If iTemp > 0 Then
                partPM = 490
                oSheetR.getCellByPosition(2, rowNumberR).CellBackColor = RGB(242, 135, 148)
                GoTo EndPLMstatus
            End If
            iTemp = InStr(1, mlfbCode, "M500", 1)
            If iTemp > 0 Then
                partPM = 500
                oSheetR.getCellByPosition(2, rowNumberR).CellBackColor = RGB(242, 135, 148)
                GoTo EndPLMstatus
            End If
        End If
        EndPLMstatus:
        DoEvents
    oSheetR.getCellByPosition(0, 3).setString("Active: ?")
    oSheetR.getCellByPosition(1, 3).setString("PhaseOut: ?")
    oSheetR.getCellByPosition(2, 3).setString("Disc: ?")
    oSheetR.getCellByPosition(3, 3).setString("Other: ?")
    oSheetR.getCellRangeByPosition(0, rowNumberR, 4, rowNumberR).Rows.OptimalHeight = TRUE
    oSheetR.getCellRangeByPosition(0, rowNumberR, 4, rowNumberR).HoriJustify = com.sun.star.table.CellHoriJustify.LEFT
    oSheetR.getCellRangeByPosition(0, rowNumberR, 4, rowNumberR).VertJustify = com.sun.star.table.CellVertJustify.TOP
    'Format report worksheet
    FormatReport
EndSub:
End Sub

Sub FormatReport()
    oSheetR = ThisComponent.Sheets.getByName("Report")
    'Print Title Rows
    Dim oTitleR As New com.sun.star.table.CellRangeAddress
    oTitleR.StartRow = 2
    oTitleR.EndRow = 3
    oTitleR.StartColumn = 0
    oTitleR.EndColumn = 0
    oSheetR.setTitleRows(oTitleR)
    oSheetR.setPrintTitleRows(TRUE)
    'Set Print Margins
    styleName = oSheetR.PageStyle
    oStyle = ThisComponent.StyleFamilies.getByName("PageStyles").getByName(styleName)
    oStyle.BottomMargin = 500 '2000
    oStyle.LeftMargin = 2000
    oStyle.RightMargin = 500 '2000
    oStyle.TopMargin = 500 '2000
    'Set Center on Page
    oStyle.CenterHorizontally = TRUE
    oStyle.CenterVertically = FALSE
    'Set Print Header and Footer
    oStyle.HeaderIsOn = TRUE
    oStyle.HeaderLeftMargin = 2000 '0
    oStyle.HeaderRightMargin = 500 '0
    oContent = oStyle.RightPageHeaderContent
    LeftText = oContent.getLeftText()
    CenterText = oContent.getCenterText()
    RightText = oContent.getRightText()
    LeftText.setString("")
    CenterText.setString("")
    RightText.setString("SPARE PARTS Report")
    oStyle.RightPageHeaderContent = oContent   
    oStyle.FooterIsOn = TRUE
    oStyle.FooterLeftMargin = 2000 '0
    oStyle.FooterRightMargin = 500 '0
    oContent = oStyle.RightPageFooterContent
    LeftText = oContent.getLeftText()
    CenterText = oContent.getCenterText()
    RightText = oContent.getRightText()
    LeftText.setString("")
    CenterText.setString("")
    RightText.setString("")
    oPageNum = ThisComponent.createInstance("com.sun.star.text.TextField.PageNumber")
    oPageTot = ThisComponent.createInstance("com.sun.star.text.TextField.PageCount")
    oTextCursor = RightText.createTextCursor
    oTextCursor.gotoEnd (False)
    RightText.setString("Page ")
    oTextCursor.gotoEnd (False)
    RightText.insertTextContent(oTextCursor, oPageNum, False)
    oTextCursor.gotoEnd (False)
    RightText.insertString(oTextCursor, " / ", False)
    oTextCursor.gotoEnd (False)
    RightText.insertTextContent(oTextCursor, oPageTot, False)
    oStyle.RightPageFooterContent = oContent
End Sub

Sub displSOW()
    'display SparesOnWeb webpage for selected product
    Dim oShell As Object
    oShell=createUnoService("com.sun.star.system.SystemShellExecute")
    Dim URL As String
    oCurrentSelection = ThisComponent.getCurrentSelection()
    If oCurrentSelection.supportsService("com.sun.star.sheet.SheetCell") Then
        rowNumber = oCurrentSelection.getCellAddress().Row
    Else
        rowNumber = 0
    End If
    If rowNumber < 3 Then
        GoTo EndSub
    End If
    oSheet = ThisComponent.getCurrentController.getActiveSheet
    mlfbCode = oSheet.getCellByPosition(0, rowNumber).getString()
    'remove successor note from code
        If (Left(mlfbCode, 8) = ("[succ.]" + CHR$(10))) Then
            mlfbCode = Right(mlfbCode, Len(mlfbCode) - 8)
        End If
        If (Left(mlfbCode, 7) = "[succ.]") Then
            mlfbCode = Right(mlfbCode, Len(mlfbCode) - 7)
        End If
        mlfbCode = Replace(mlfbCode, CHR$(10), " ")
        'Try to extract options
        optLen = Len(mlfbCode)
        optPos = InStr(1, mlfbCode, "-Z ", 1)
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
            URL = convertToUrl(URL)
            'Open the URL in default browser'
            oShell.execute(URL, "", 0)   
        End If
EndSub:
End Sub

Sub displMall()
    'display IndustryMall webpage for selected product
    Dim oShell As Object
    oShell=createUnoService("com.sun.star.system.SystemShellExecute")
    Dim URL As String
    oCurrentSelection = ThisComponent.getCurrentSelection()
    If oCurrentSelection.supportsService("com.sun.star.sheet.SheetCell") Then
        rowNumber = oCurrentSelection.getCellAddress().Row
    Else
        rowNumber = 0
    End If
    If rowNumber < 3 Then
        GoTo EndSub
    End If
    oSheet = ThisComponent.getCurrentController.getActiveSheet
    mlfbCode = oSheet.getCellByPosition(0, rowNumber).getString()
    'remove successor note from code
        If (Left(mlfbCode, 8) = ("[succ.]" + CHR$(10))) Then
            mlfbCode = Right(mlfbCode, Len(mlfbCode) - 8)
        End If
        If (Left(mlfbCode, 7) = "[succ.]") Then
            mlfbCode = Right(mlfbCode, Len(mlfbCode) - 7)
        End If
        mlfbCode = Replace(mlfbCode, CHR$(10), " ")
        'options should be separated by space instead of +
        mlfbCode = Replace(mlfbCode, "+", "%20")
        If Len(mlfbCode) > 1 Then
            URL = "https://mall.industry.siemens.com/mall/en/WW/Catalog/Product/" + mlfbCode
            URL = convertToUrl(URL)
            'Open the URL in default browser'
            oShell.execute(URL, "", 0)   
        End If
EndSub:
End Sub

Sub displSios()
    'display SIOS webpage for selected product
    Dim oShell As Object
    oShell=createUnoService("com.sun.star.system.SystemShellExecute")
    Dim URL As String
    oCurrentSelection = ThisComponent.getCurrentSelection()
    If oCurrentSelection.supportsService("com.sun.star.sheet.SheetCell") Then
        rowNumber = oCurrentSelection.getCellAddress().Row
    Else
        rowNumber = 0
    End If
    If rowNumber < 3 Then
        GoTo EndSub
    End If
    oSheet = ThisComponent.getCurrentController.getActiveSheet
    mlfbCode = oSheet.getCellByPosition(0, rowNumber).getString()
    'remove successor note from code
        If (Left(mlfbCode, 8) = ("[succ.]" + CHR$(10))) Then
            mlfbCode = Right(mlfbCode, Len(mlfbCode) - 8)
        End If
        If (Left(mlfbCode, 7) = "[succ.]") Then
            mlfbCode = Right(mlfbCode, Len(mlfbCode) - 7)
        End If
        mlfbCode = Replace(mlfbCode, CHR$(10), " ")
        'Try to extract options
        optLen = Len(mlfbCode)
        optPos = InStr(1, mlfbCode, "-Z ", 1)
        If ((optPos > 0) And (optPos < optLen)) Then
            mlfbOpts = Right(mlfbCode, Len(mlfbCode) - optPos - 2)
            mlfbCode = Left(mlfbCode, optPos + 1)
        Else
            mlfbOpts = ""
        End If
        'options should be separated by space instead of +
        mlfbOpts = Replace(mlfbOpts, "+", "%20")
        If Len(mlfbCode) > 1 Then
            URL = "https://support.industry.siemens.com/cs/products/" + mlfbCode
            URL = convertToUrl(URL)
            'Open the URL in default browser'
            oShell.execute(URL, "", 0)   
        End If
EndSub:
End Sub

Sub displGit()
    Dim oShell As Object
    oShell=createUnoService("com.sun.star.system.SystemShellExecute")
    Dim URL As String
    URL = "https://github.com/alxcor/camxls"
    URL = convertToUrl(URL)
    'Open the URL in default browser'
    oShell.execute(URL, "", 0)   
End Sub

