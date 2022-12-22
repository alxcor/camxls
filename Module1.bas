Attribute VB_Name = "Module1"
Sub setBorders(rowNumber)
    'set borders for a range of cells in a row
    Dim maxCol As Integer
    'last column is AD = column 30
    maxCol = 30
    'set borders for current row, columns 1 to maxCol
    Range(Cells(rowNumber, 1), Cells(rowNumber, maxCol)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlHairline
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlHairline
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlHairline
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlHairline
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlHairline
    End With
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub

Sub setHeaders()
    'Set header texts on row 3
    Cells(3, 1) = "No"
    Cells(3, 2) = "Your Data..."
    Cells(3, 3) = "MLFB"
    Cells(3, 4) = "Product Description"
    Cells(3, 5) = "Product family"
    Cells(3, 6) = "Product Lifecycle (PLM)"
    Cells(3, 7) = "PLM Effective Date"
    Cells(3, 8) = "Notes"
    Cells(3, 9) = "Price Group"
    Cells(3, 10) = "Surcharge for Raw Materials"
    Cells(3, 11) = "Metal Factor"
    Cells(3, 12) = "Export Control Regulations"
    Cells(3, 13) = "Delivery Time"
    Cells(3, 14) = "Net Weight (kg)"
    Cells(3, 15) = "Product Dimensions (W x L x H)"
    Cells(3, 16) = "Packaging Dimension"
    Cells(3, 17) = "Package size unit of measure"
    Cells(3, 18) = "Quantity Unit"
    Cells(3, 19) = "Packaging Quantity"
    Cells(3, 20) = "EAN"
    Cells(3, 21) = "UPC"
    Cells(3, 22) = "Commodity Code"
    Cells(3, 23) = "KZ_FDB/ CatalogID"
    Cells(3, 24) = "Product Group"
    Cells(3, 25) = "Country of origin"
    Cells(3, 26) = "Compliance with the substance restrictions according to RoHS directive"
    Cells(3, 27) = "Product class"
    Cells(3, 28) = "Obligation Category for taking back electrical and electronic equipment after use"
    Cells(3, 29) = "Classifications"
    Cells(3, 30) = "Successor"
    Cells(3, 3).EntireRow.Font.Bold = True
    Call setBorders(3)
    Range("A3:AD3").Select
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
End Sub

Sub ImportSieMallIntra(mlfbCode, rowNumber, netMode)
    'read data for a specific product code (MLFB) from Industry Mall web site
    'netMode: 0=Internet; 1=Intranet
    Dim IE
    Dim targetURL As String
    Dim webContent As String
    Dim sh
    Dim eachIE
    Dim Product As IHTMLElement
    Dim ProductDetails As IHTMLElementCollection
    Dim Detail As IHTMLElement
    Dim Details As IHTMLElementCollection
    Dim DetailNo As Integer
    Dim iCounter As Integer
    Dim index As Integer
    Dim sTemp As String
    Dim iTemp, iTemp2, PM_stat As Integer
    '--------------------------------------------------------------------------------
    'Reading web page in buffer...
    'write status in StatusBar...
    Application.StatusBar = "Trying to connect to Industry Mall... MLFB: " + mlfbCode
    'format spaces html style
    mlfbCode = Replace(mlfbCode, " ", "%20")
    'set web page (for scrapper)
    targetURL = "https://mall.industry.siemens.com/mall/en/WW/Catalog/Product/" + mlfbCode
    If (netMode = 0) Then
        'Internet version
        Application.StatusBar = "Trying to connect -via internet- to Industry Mall... MLFB: " + mlfbCode
        Set IE = New InternetExplorer
        IE.Visible = False
        IE.navigate targetURL
        'Wait until IE is done loading page: 4 = ReadyState Complete
        Do While IE.READYSTATE <> 4
            DoEvents
        Loop
        Set html = IE.document
    Else
        'Intranet version
        Application.StatusBar = "Trying to connect -via intranet- to Industry Mall... MLFB: " + mlfbCode
        Set IE = New InternetExplorerMedium
        IE.Visible = False
        IE.navigate targetURL
        'Wait until IE is done loading page: 4 = ReadyState Complete
        While IE.Busy
            DoEvents
        Wend
        index = 0
        Do
            Set sh = New Shell32.Shell
            For Each eachIE In sh.Windows
                index = index + 1
                If index < 100 Then
                    If InStr(1, eachIE.LocationURL, targetURL) Then
                        Set IE = eachIE
                        'IE.Visible = False
                        Exit Do
                    End If
                Else
                    Application.StatusBar = "Err..."
                    'MsgBox ("Err: product not found " & mlfbCode)
                    Cells(rowNumber, 3).Value = "Err: product not found " & mlfbCode
                    Exit Sub
                End If
            Next eachIE
        Loop
        Set eachIE = Nothing
        Set sh = Nothing
        Application.StatusBar = "Trying to connect -via intranet [" + index + "]- to Industry Mall... MLFB: " + mlfbCode
        While IE.Busy  'Wait for all IE
            DoEvents
        Wend
        Set html = IE.document
    End If
    Application.StatusBar = "..."
    'See content of the web-page for diagnosis purposes...
    'MsgBox html.documentElement.innerHTML
    index = 1
    DetailNo = 1
    'Search for ID-content in web page
    Set Product = html.getElementById("content")
    Set ProductDetails = Product.all
    Application.StatusBar = "Trying to get details for MLFB: " + mlfbCode
    For Each Detail In ProductDetails
        'Column 1 [A]: index of product in list...
        If Cells(rowNumber, 1).Value = "" Then
            Cells(rowNumber, 1) = rowNumber - 3
        End If
        'Column 3 [C]: MLFB code...
        If Detail.className = "productIdentifier" Then
            Cells(rowNumber, 3).Value = Detail.innerText
        End If
        'Produs details - Extract from table:
        If Detail.className = "ProductDetailsTable" Then
            'Get all details for the product
            Set Details = Detail.all
            'Count details fields
            DetailNo = Details.Length
            For index = 0 To DetailNo - 1
                'Column 4 [D]: Product Description
                If (Details(index).innerText = "Product Description") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 4) = Details(index + 1).innerText
                End If
                'Column 5 [E]: Product Family
                If (Details(index).innerText = "Product family") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 5) = Details(index + 1).innerText
                End If
                'Column 6 [F]: Product Lifecycle(PLM)
                If (Details(index).innerText = "Product Lifecycle (PLM)") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 6) = Details(index + 1).innerText
                    'PLM status:
                    sTemp = Details(index + 1).innerText
                    iTemp = InStr(1, sTemp, "M250", vbTextCompare)
                    If iTemp > 0 Then
                        Cells(rowNumber, 6).Interior.Color = RGB(0, 255, 0)
                    End If
                    iTemp = InStr(1, sTemp, "M280", vbTextCompare)
                    If iTemp > 0 Then
                        Cells(rowNumber, 6).Interior.Color = RGB(0, 255, 0)
                    End If
                    iTemp = InStr(1, sTemp, "M300", vbTextCompare)
                    If iTemp > 0 Then
                        Cells(rowNumber, 6).Interior.Color = RGB(0, 255, 0)
                    End If
                    iTemp = InStr(1, sTemp, "M400", vbTextCompare)
                    If iTemp > 0 Then
                        Cells(rowNumber, 6).Interior.Color = RGB(255, 255, 0)
                    End If
                    iTemp = InStr(1, sTemp, "M410", vbTextCompare)
                    If iTemp > 0 Then
                        Cells(rowNumber, 6).Interior.Color = RGB(255, 255, 0)
                    End If
                    iTemp = InStr(1, sTemp, "M490", vbTextCompare)
                    If iTemp > 0 Then
                        Cells(rowNumber, 6).Interior.Color = RGB(255, 0, 0)
                    End If
                    iTemp = InStr(1, sTemp, "M500", vbTextCompare)
                    If iTemp > 0 Then
                        Cells(rowNumber, 6).Interior.Color = RGB(255, 0, 0)
                    End If
                End If
                'Column 7 [G]: PLM Effective Date
                If (Details(index).innerText = "PLM Effective Date") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 7) = Details(index + 1).innerText
                End If
                'Column 8 [H]: Notes
                If (Details(index).innerText = "Notes") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 8) = Details(index + 1).innerText
                    sTemp = Details(index + 1).innerText
                    iTemp = InStr(1, sTemp, "Successor", vbTextCompare)
                    iTemp2 = Len(sTemp)
                    If (iTemp > 0) And ((iTemp2 - iTemp > 11)) Then
                        iTemp = iTemp + 11
                        sTemp = Mid(sTemp, iTemp, iTemp2 - iTemp)
                        'Column 30 [AD]: Successor: attempt to identify MLFB for successor product
                        Cells(rowNumber, 30) = sTemp
                        'Rows(rowNumber + 1).EntireRow.Insert
                        'Cells(rowNumber + 1, 1) = "Successor:"
                        'Cells(rowNumber + 1, 2) = sTemp
                    End If
                    If Cells(rowNumber, 8).Value <> "" Then
                        Cells(rowNumber, 8).Interior.Color = RGB(0, 0, 255)
                    End If
                End If
                'Column 9 [I]: Price Group
                If (Details(index).innerText = "Price Group") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 9) = Details(index + 1).innerText
                End If
                'Column 10 [J]: Surcharge for Raw Materials
                If (Details(index).innerText = "Surcharge for Raw Materials") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 10) = Details(index + 1).innerText
                End If
                'Column 11 [K]: Surcharge for Metal Factor
                If (Details(index).innerText = "Metal Factor") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 11) = Details(index + 1).innerText
                End If
                'Column 12 [L]: Export Control Regulations
                If (Details(index).innerText = "Export Control Regulations") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 12) = Details(index + 1).innerText
                End If
                'Column 13 [M]: Delivery Time
                If (Details(index).innerText = "Delivery Time") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 13) = Details(index + 1).innerText
                End If
                'Column 14 [N]: Net Weight(kg)
                If (Details(index).innerText = "Net Weight(kg)") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 14) = Details(index + 1).innerText
                End If
                'Column 15 [O]: Product Dimensions (W x L x H)
                If (Details(index).innerText = "Product Dimensions (W x L x H)") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 15) = Details(index + 1).innerText
                End If
                'Column 16 [P]: Packaging Not Dimension
                If (Details(index).innerText = "Packaging Dimension") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 16) = Details(index + 1).innerText
                End If
                'Column 17 [Q]: Package size unit of measure
                If (Details(index).innerText = "Package size unit of measure") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 17) = Details(index + 1).innerText
                End If
                'Column 18 [R]: Quantity Unit
                If (Details(index).innerText = "Quantity Unit") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 18) = Details(index + 1).innerText
                End If
                'Column 19 [S]: Packaging Quantity
                If (Details(index).innerText = "Packaging Quantity") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 19) = Details(index + 1).innerText
                End If
                'Column 20 [T]: EAN
                If (Details(index).innerText = "EAN") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 20) = "'" & Details(index + 1).innerText
                End If
                'Column 21 [U]: UPC
                If (Details(index).innerText = "UPC") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 21) = "'" & Details(index + 1).innerText
                End If
                'Column 22 [V]: Commodity Code
                If (Details(index).innerText = "Commodity Code") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 22) = "'" & Details(index + 1).innerText
                End If
                'Column 23 [W]: LKZ_FDB/ CatalogID
                If (Details(index).innerText = "LKZ_FDB/ CatalogID") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 23) = Details(index + 1).innerText
                End If
                'Column 24 [X]: Product Group
                If (Details(index).innerText = "Product Group") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 24) = Details(index + 1).innerText
                End If
                'Column 25 [Y]: Country of origin
                If (Details(index).innerText = "Country of origin") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 25) = Details(index + 1).innerText
                End If
                'Column 26 [Z]: Compliance with the substance restrictions according to RoHS directive
                If (Details(index).innerText = "Compliance with the substance restrictions according to RoHS directive") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 26) = Details(index + 1).innerText
                End If
                'Column 27 [AA]: Product class
                If (Details(index).innerText = "Product class") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 27) = Details(index + 1).innerText
                End If
                'Column 28 [AB]: Obligation Category for taking back electrical and electronic equipment after use
                If (Details(index).innerText = "Obligation Category for taking back electrical and electronic equipment after use") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 28) = Details(index + 1).innerText
                End If
                'Column 29 [AC]: Classifications
                If (Details(index).innerText = "Classifications") And (index < DetailNo - 1) Then
                    Cells(rowNumber, 29) = Details(index + 1).innerText
                End If
            Next
        End If
    Next
    'close down IE and reset status bar
    IE.Quit
    Set IE = Nothing
    Set html = Nothing
    Application.StatusBar = ""
End Sub

Sub EvaluateRow(netMode)
    'read data for current row: on column 2 [B] should be a product code (MLFB) from Industry Mall web site
    'netMode: 0=Internet; 1=Intranet
    Dim rowNumber As Long
    Dim mlfbCode As String
    Dim iCounter As Integer
    'Range("C1:AZ1").EntireColumn.Clear
    'Cells.VerticalAlignment = xlTop
    '--------------------------------------------
    rowNumber = ActiveCell.Row
    If rowNumber < 4 Then
        MsgBox ("[EN]: Table starts on row 4; [RO]:Tabelul incepe de la randul 4!")
    Else
        mlfbCode = Cells(rowNumber, 2).Value
        If Len(mlfbCode) > 1 Then
            Call ImportSieMallIntra(mlfbCode, rowNumber, netMode)
            Call setBorders(rowNumber)
        End If
    End If
    Application.StatusBar = ""
    MsgBox "Done!"
End Sub

Sub EvaluateAll(netMode)
    'read data for all non-empty rows >= 4: on column 2 [B] should be a product code (MLFB) from Industry Mall web site
    'netMode: 0=Internet; 1=Intranet
    Dim rowNumber As Long
    Dim mlfbCode As String
    Dim iCounter As Integer
    '--------------------------------------------
    'clear old data out and put titles in
    Range("C1:AZ1").EntireColumn.Clear
    Cells.VerticalAlignment = xlTop
    Call setHeaders
    '--------------------------------------------
    For rowNumber = 4 To 500
        mlfbCode = Cells(rowNumber, 2).Value
        If Len(mlfbCode) > 1 Then
            Call ImportSieMallIntra(mlfbCode, rowNumber, netMode)
            Call setBorders(rowNumber)
        End If
    Next
    '--------------------------------------------
    'do some final formatting
    Cells(3, 3).EntireColumn.WrapText = False
    Cells(3, 3).EntireColumn.AutoFit
    Cells(3, 4).EntireColumn.WrapText = True
    Cells(3, 4).EntireColumn.ColumnWidth = 40
    Cells(3, 5).EntireColumn.WrapText = False
    Cells(3, 5).EntireColumn.AutoFit
    Cells(3, 6).EntireColumn.WrapText = False
    Cells(3, 6).EntireColumn.AutoFit
    Cells(3, 7).EntireColumn.WrapText = False
    Cells(3, 7).EntireColumn.AutoFit
    Cells(3, 8).EntireColumn.WrapText = True
    Cells(3, 8).EntireColumn.ColumnWidth = 40
    Cells(3, 9).EntireColumn.WrapText = False
    Cells(3, 9).EntireColumn.AutoFit
    Cells(3, 10).EntireColumn.WrapText = False
    Cells(3, 10).EntireColumn.AutoFit
    Cells(3, 11).EntireColumn.WrapText = False
    Cells(3, 11).EntireColumn.AutoFit
    Cells(3, 12).EntireColumn.WrapText = False
    Cells(3, 12).EntireColumn.AutoFit
    Cells(3, 13).EntireColumn.WrapText = False
    Cells(3, 13).EntireColumn.AutoFit
    Cells(3, 14).EntireColumn.WrapText = False
    Cells(3, 14).EntireColumn.AutoFit
    Cells(3, 15).EntireColumn.WrapText = False
    Cells(3, 15).EntireColumn.AutoFit
    Cells(3, 16).EntireColumn.WrapText = False
    Cells(3, 16).EntireColumn.AutoFit
    Cells(3, 17).EntireColumn.WrapText = False
    Cells(3, 17).EntireColumn.AutoFit
    Cells(3, 18).EntireColumn.WrapText = False
    Cells(3, 18).EntireColumn.AutoFit
    Cells(3, 19).EntireColumn.WrapText = False
    Cells(3, 19).EntireColumn.AutoFit
    Cells(3, 20).EntireColumn.WrapText = False
    Cells(3, 20).EntireColumn.AutoFit
    Cells(3, 21).EntireColumn.WrapText = False
    Cells(3, 21).EntireColumn.AutoFit
    Cells(3, 22).EntireColumn.WrapText = False
    Cells(3, 22).EntireColumn.AutoFit
    Cells(3, 23).EntireColumn.WrapText = False
    Cells(3, 23).EntireColumn.AutoFit
    Cells(3, 24).EntireColumn.WrapText = False
    Cells(3, 24).EntireColumn.AutoFit
    Cells(3, 25).EntireColumn.WrapText = False
    Cells(3, 25).EntireColumn.AutoFit
    Cells(3, 26).EntireColumn.WrapText = True
    Cells(3, 26).EntireColumn.ColumnWidth = 40
    Cells(3, 27).EntireColumn.WrapText = True
    Cells(3, 27).EntireColumn.ColumnWidth = 40
    Cells(3, 28).EntireColumn.WrapText = True
    Cells(3, 28).EntireColumn.ColumnWidth = 40
    Cells(3, 29).EntireColumn.WrapText = True
    Cells(3, 29).EntireColumn.ColumnWidth = 40
    Cells(3, 30).EntireColumn.WrapText = False
    Cells(3, 30).EntireColumn.ColumnWidth = 40
    '--------------------------------------------
    Application.StatusBar = ""
    MsgBox "Done!"
End Sub

Sub Report()
    Dim rowNumber As Long
    Dim rowNumber2 As Long
    Dim mlfbCode As String
    Dim iCounter As Integer
    '--------------------------------------------
    Sheets("Report").Cells.Delete
    Sheets("Report").Cells(3, 1) = "MLFB"
    Sheets("Report").Cells(3, 2) = "Product Description"
    'Sheets("Report").Cells(3, 3) = "Product family"
    Sheets("Report").Cells(3, 3) = "Product Lifecycle (PLM)"
    Sheets("Report").Cells(3, 4) = "PLM Effective Date"
    'Sheets("Report").Cells(3, 5) = "Notes"
    Sheets("Report").Cells.VerticalAlignment = xlTop
    Sheets("Report").Cells(3, 1).EntireColumn.WrapText = True
    Sheets("Report").Cells(3, 1).EntireColumn.ColumnWidth = 20
    Sheets("Report").Cells(3, 2).EntireColumn.WrapText = True
    Sheets("Report").Cells(3, 2).EntireColumn.ColumnWidth = 20
    Sheets("Report").Cells(3, 3).EntireColumn.WrapText = True
    Sheets("Report").Cells(3, 3).EntireColumn.ColumnWidth = 20
    Sheets("Report").Cells(3, 4).EntireColumn.WrapText = True
    Sheets("Report").Cells(3, 4).EntireColumn.ColumnWidth = 15
    Sheets("Report").Cells(3, 5).EntireColumn.WrapText = False
    rowNumber2 = 4
    For rowNumber = 4 To 100
      If (Cells(rowNumber, 2).Value <> "") Then
        With Sheets("Report")
            .Range(.Cells(rowNumber2 + 1, 2), .Cells(rowNumber2 + 1, 5)).Merge
        End With
        Sheets("Report").Cells(rowNumber2, 1).EntireRow.Font.Name = "Arial"
        Sheets("Report").Cells(rowNumber2, 1).EntireRow.Font.Size = 8
        Sheets("Report").Cells(rowNumber2, 1).EntireRow.Font.ColorIndex = xlAutomatic
        Sheets("Report").Cells(rowNumber2, 1).EntireRow.Font.TintAndShade = 0
        Sheets("Report").Cells(rowNumber2 + 1, 1).EntireRow.Font.Name = "Arial"
        Sheets("Report").Cells(rowNumber2 + 1, 1).EntireRow.Font.Size = 8
        Sheets("Report").Cells(rowNumber2 + 1, 1).EntireRow.Font.ColorIndex = xlAutomatic
        Sheets("Report").Cells(rowNumber2 + 1, 1).EntireRow.Font.TintAndShade = 0
        For iCounter = 1 To 5
            Sheets("Report").Cells(rowNumber2, iCounter).Borders(xlEdgeBottom).LineStyle = xlContinuous
            Sheets("Report").Cells(rowNumber2, iCounter).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
            Sheets("Report").Cells(rowNumber2, iCounter).Borders(xlEdgeBottom).TintAndShade = 0
            Sheets("Report").Cells(rowNumber2, iCounter).Borders(xlEdgeBottom).Weight = xlHairline
        Next
        Sheets("Report").Cells(rowNumber2 + 1, 1).Borders(xlEdgeBottom).Weight = xlThin
        Sheets("Report").Cells(rowNumber2 + 1, 2).Borders(xlEdgeBottom).Weight = xlThin
        Sheets("Report").Cells(rowNumber2 + 1, 3).Borders(xlEdgeBottom).Weight = xlThin
        Sheets("Report").Cells(rowNumber2 + 1, 4).Borders(xlEdgeBottom).Weight = xlThin
        Sheets("Report").Cells(rowNumber2 + 1, 5).Borders(xlEdgeBottom).Weight = xlThin
        mlfbCode = Cells(rowNumber, 2).Value
        Sheets("Report").Cells(rowNumber2, 1) = mlfbCode
        Sheets("Report").Cells(rowNumber2, 2) = Cells(rowNumber, 4).Value
        Sheets("Report").Cells(rowNumber2, 3) = Cells(rowNumber, 6).Value
        Sheets("Report").Cells(rowNumber2, 4) = Cells(rowNumber, 7).Value
        Sheets("Report").Cells(rowNumber2 + 1, 1) = "Notes:"
        Sheets("Report").Cells(rowNumber2 + 1, 2) = Cells(rowNumber, 8).Value
        Sheets("Report").Cells(rowNumber2 + 1, 2).WrapText = False
        'PLM:
        sTemp = Cells(rowNumber, 6).Value
        iTemp = InStr(1, sTemp, "M250", vbTextCompare)
        If iTemp > 0 Then
            Sheets("Report").Cells(rowNumber2, 3).Interior.Color = RGB(0, 255, 0)
        End If
        iTemp = InStr(1, sTemp, "M280", vbTextCompare)
        If iTemp > 0 Then
            Sheets("Report").Cells(rowNumber2, 3).Interior.Color = RGB(0, 255, 0)
        End If
        iTemp = InStr(1, sTemp, "M300", vbTextCompare)
        If iTemp > 0 Then
            Sheets("Report").Cells(rowNumber2, 3).Interior.Color = RGB(0, 255, 0)
        End If
        iTemp = InStr(1, sTemp, "M400", vbTextCompare)
        If iTemp > 0 Then
            Sheets("Report").Cells(rowNumber2, 3).Interior.Color = RGB(255, 255, 0)
        End If
        iTemp = InStr(1, sTemp, "M410", vbTextCompare)
        If iTemp > 0 Then
            Sheets("Report").Cells(rowNumber2, 3).Interior.Color = RGB(255, 255, 0)
        End If
        iTemp = InStr(1, sTemp, "M490", vbTextCompare)
        If iTemp > 0 Then
            Sheets("Report").Cells(rowNumber2, 3).Interior.Color = RGB(255, 0, 0)
        End If
        iTemp = InStr(1, sTemp, "M500", vbTextCompare)
        If iTemp > 0 Then
            Sheets("Report").Cells(rowNumber2, 3).Interior.Color = RGB(255, 0, 0)
        End If
        'Sheets("Report").Cells(rowNumber2 + 1, 2).EntireRow.AutoFit = True
      End If
    rowNumber2 = rowNumber2 + 2
    Next
    'some final formatting
    Sheets("Report").Select
    Cells(3, 5).EntireColumn.AutoFit
    'Cells(3, 6).EntireColumn.WrapText = True
    'Cells(3, 6).EntireColumn.AutoFit
    Range("A3:E3").Select
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    Range("A1").Value = "Report:"
    Range("A1").Font.Bold = True
End Sub
