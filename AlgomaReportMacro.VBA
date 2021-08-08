Sub ImportInventoryReport():
    Dim fn$
    Dim wkb1 As Workbook
    Dim wkb2 As Workbook
    Dim sht1 As Worksheet
    Dim sht2 As Worksheet
    fn = Application.GetOpenFilename("All XLSX files (*.xlsx*), *.xlsx", , "Please choose Inventory Report: ")
    Application.ScreenUpdating = False
    Set wkb1 = ThisWorkbook
    Set wkb2 = Workbooks.Open(fn)
    Set sht1 = wkb1.Sheets("Inventory")
    Set sht2 = wkb2.Sheets("Report - ALGOMA STEEL")
    sht2.Cells.Copy
    sht1.Range("A1").PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    wkb2.Close True
    Application.ScreenUpdating = True
End Sub

Function iterateAdd(k As Integer, colArr, i As Integer) As Variant
    Dim j As Integer
    Dim tempArr As Variant
    Dim x As Long, y As Long
    Dim maxX As Long, minX As Long
    Dim maxY As Long, minY As Long
    Dim lrow As Long
    minX = LBound(colArr, 1)
    maxX = UBound(colArr, 1)
    minY = LBound(colArr, 2)
    maxY = UBound(colArr, 2)
    ReDim tempArr(minX To maxX, minY To maxY)
    lrow = Cells(Rows.Count, "A").End(xlUp).Row
    For x = minX To maxX
        For y = minY To maxY
            tempArr(x, y) = colArr(x, y)
        Next y
    Next x
    j = 1
    Do While j < lrow + 1:
        tempArr(k, j - 1) = Range(Chr(i) & j).Value
        j = j + 1
    Loop
    iterateAdd = tempArr
End Function

Sub ImportTailgateReport():
    Dim fn$
    Dim wkb1 As Workbook
    Dim wkb2 As Workbook
    Dim sht1 As Worksheet
    Dim sht2 As Worksheet
    fn = Application.GetOpenFilename("All XLSX files (*.xlsx*), *.xlsx", , "Please choose Tailgate Report: ")
    Application.ScreenUpdating = False
    Set wkb1 = ThisWorkbook
    Set wkb2 = Workbooks.Open(fn)
    Set sht1 = wkb1.Sheets("Shipped")
    Set sht2 = wkb2.Sheets("Detailed Information")
    sht2.Cells.Copy
    sht1.Range("A1").PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    wkb2.Close True
    Application.ScreenUpdating = True
End Sub

Public Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function

Public Function ArrayLen(arr As Variant) As Integer
    ArrayLen = UBound(arr) - LBound(arr) + 1
End Function

Sub FormatInventory():
    Dim wkb As Workbook
    Dim sht As Worksheet
    Set wkb = ThisWorkbook
    Set sht = wkb.Sheets("Inventory")
    Dim i As Integer
    Dim checkArr As Variant
    sht.Activate
    Application.ScreenUpdating = False
    Do While IsEmpty(Range("A1").Value)
        Rows(1).EntireRow.Delete
    Loop
    Rows(1).EntireRow.Delete
    checkArr = Array("Client Inventory", "Order", "Receiver", "Heat Number", "Mark", "Scope", "Other", "Weight per Unit", "Dimensions")
    i = 1
    Do While Not IsEmpty(Cells(1, i).Value):
        If IsInArray(Cells(1, i).Value, checkArr) Then
            i = i + 1
        Else
            Columns(i).EntireColumn.Delete
        End If
    Loop
    Range("G1").Value = "Recieved Date"
    Range("J1").Value = "Quantity"
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        Range("J" & i).Value = 1
    Next i
    i = 1
    Do While Not IsEmpty(Cells(1, i).Value):
        Columns(i).AutoFit
        i = i + 1
    Loop
    Columns("J").EntireColumn.HorizontalAlignment = xlCenter
    Application.ScreenUpdating = True
End Sub

Sub FormatShipped():
    Dim wkb As Workbook
    Dim sht As Worksheet
    Set wkb = ThisWorkbook
    Set sht = wkb.Sheets("Shipped")
    Dim i As Integer
    Dim lrow As Long
    Dim testable As String
    Dim colStore() As Variant
    sht.Activate
    Application.ScreenUpdating = False
    Do While IsEmpty(Range("A1").Value)
        Rows(1).EntireRow.Delete
    Loop
    i = 1
    Do While Not IsEmpty(Range("A" & i).Value)
        If Range("A" & i).Value <> "Shipment by Truck" And Range("A" & i).Value <> "Movement Type" Then
            Rows(i).EntireRow.Delete
        Else
           i = i + 1
        End If
    Loop
    i = 65
    Do While Not IsEmpty(Range(Chr(i) & "1").Value)
        testable = Range(Chr(i) & "1").Value
        If testable = "Movement Type" Or testable = "Client" Or testable = "Inventory" Or testable = "User" Or testable = "Vessel Project" Or testable = "Carrier Bill" Or testable = "Appointment Number" Or testable = "Vehicle Identifier" Or testable = "Billable" Or testable = "Billable Remarks" Or testable = "Product Type" Or testable = "Product Subtype" Or testable = "Seal Numbers" Or testable = "Description" Or testable = "PO Number" Or testable = "Lot Number" Or testable = "Serial Number" Or testable = "Bill of Lading" Or testable = "Barcode" Or testable = "Diameter" Or testable = "Diameter Unit" Then
            Columns(Chr(i)).EntireColumn.Delete
        ElseIf testable = "Foot Pound" Or testable = "Foot Pound Unit" Or testable = "Length" Or testable = "Length Unit" Or testable = "Width" Or testable = "Width Unit" Or testable = "Height" Or testable = "Height Unit" Or testable = "Thickness" Or testable = "Thickness Unit" Or testable = "Quantity Per Package" Or testable = "Packaging" Or testable = "Total Volume" Or testable = "Original Vessel Project" Or testable = "Freight Ton" Or testable = "LTD Qty + Class" Or testable = "UN" Or testable = "PG" Or testable = "Location" Or testable = "Customs Release" Or testable = "Shipping Line Release" Or testable = "Financial Release" Or testable = "Customer Release" Then
            Columns(Chr(i)).EntireColumn.Delete
        Else
            i = i + 1
        End If
    Loop
    i = 65
    lrow = Cells(Rows.Count, "A").End(xlUp).Row
    ReDim colStore(5, 0 To lrow - 1)
    Do While Not IsEmpty(Range(Chr(i) & "1").Value):
        testable = Range(Chr(i) & "1").Value
        Select Case testable:
            Case "ROM/SOM Number":
                Range(Chr(i) & "1").Value = "BOL #"
            Case "Other":
                Range(Chr(i) & "1").Value = "Received"
            Case "Receiver":
                colStore = iterateAdd(2, colStore, i)
            Case "Order":
                colStore = iterateAdd(3, colStore, i)
            Case "Movement Date":
                colStore = iterateAdd(1, colStore, i)
            Case "Carrier":
                colStore = iterateAdd(0, colStore, i)
            Case "Quantity":
                colStore = iterateAdd(5, colStore, i)
            Case "Total Weight":
                colStore = iterateAdd(4, colStore, i)
        End Select
        i = i + 1
    Loop
    For x = 0 To 5
        For y = 0 To UBound(colStore, 2)
        Next y
    Next x
    i = 65
    Do While Not IsEmpty(Range(Chr(i) & "1").Value):
        Select Case Chr(i):
            Case "B":
                For y = 0 To UBound(colStore, 2)
                    Range("B" & y + 1).Value = colStore(0, y)
                Next y
            Case "C":
                For y = 0 To UBound(colStore, 2)
                    Range("C" & y + 1).Value = colStore(1, y)
                Next y
            Case "D":
                For y = 0 To UBound(colStore, 2)
                    Range("D" & y + 1).Value = colStore(2, y)
                Next y
            Case "E":
                For y = 0 To UBound(colStore, 2)
                    Range("E" & y + 1).Value = colStore(3, y)
                Next y
            Case "K":
                For y = 0 To UBound(colStore, 2)
                    Range("K" & y + 1).Value = colStore(4, y)
                Next y
            Case "L":
                For y = 0 To UBound(colStore, 2)
                    Range("L" & y + 1).Value = colStore(5, y)
                Next y
        End Select
        i = i + 1
    Loop
    i = 65
    Do While Not IsEmpty(Range(Chr(i) & 1).Value):
        Columns(Chr(i)).AutoFit
        i = i + 1
    Loop
    Columns("K").EntireColumn.NumberFormat = "0.00"
    Application.ScreenUpdating = True
End Sub

Sub MakeTable()
    Application.ScreenUpdating = False
    Dim tbl As ListObject
    Dim rng As Range
    Dim lrow As Long
    lrow = Cells(Rows.Count, "A").End(xlUp).Row
    If ActiveSheet.Name = "Shipped" Then
        Set rng = Range(Range("A1"), Range("L" & lrow))
    ElseIf ActiveSheet.Name = "Inventory" Then
        Set rng = Range(Range("A1"), Range("J" & lrow))
    End If
    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, rng, , xlYes)
    tbl.TableStyle = "TableStyleLight9"
    Rows(1).EntireRow.HorizontalAlignment = xlCenter
    Application.ScreenUpdating = True
End Sub

Sub CoilBreakdown()
    Dim receiverArr() As Variant
    Dim lrow, arrayLength, i, k, tot As Integer
    Worksheets("Inventory").Activate
    lrow = Cells(Rows.Count, "A").End(xlUp).Row
    ReDim Preserve receiverArr(0)
    receiverArr(0) = Range("B2").Value
    i = 3
    Do While Not IsEmpty(Range("B" & i).Value)
        If Not IsInArray(Range("B" & i).Value, receiverArr) Then
            arrayLength = ArrayLen(receiverArr)
            ReDim Preserve receiverArr(arrayLength)
            receiverArr(arrayLength) = Range("B" & i).Value
        End If
        i = i + 1
    Loop
    Range("B" & lrow + 5 & ":" & "C" & lrow + 5).Merge
    Range("B" & lrow + 5 & ":C" & lrow + 6).Font.Bold = True
    Range("B" & lrow + 5).Font.ColorIndex = 2
    Range("B" & lrow + 5).Value = "Customer Coil Breakdown"
    Range("B" & lrow + 6).Value = "Receiver"
    Range("C" & lrow + 6).Value = "Total"
    Range("B" & lrow + 5).Interior.ColorIndex = 23
    Range("B" & lrow + 6 & ":C" & lrow + 6).Interior.ColorIndex = 37
    lrow = lrow + 6
    k = 1
    tot = 0
    For Each cust In receiverArr
        i = 2
        j = 0
        Do While Not IsEmpty(Range("B" & i).Value)
            If Range("B" & i).Value = cust Then
                j = j + 1
            End If
            i = i + 1
        Loop
        Range("B" & lrow + k).Value = cust
        Range("C" & lrow + k).Value = j
        Range("B" & lrow + k & ":C" & lrow + k).Interior.ColorIndex = 2
        k = k + 1
        tot = tot + j
    Next cust
    Range("B" & lrow + k + 1).Value = "Coil Total"
    Range("C" & lrow + k + 1).Value = tot
    Range("B" & lrow + k + 1 & ":C" & lrow + k + 1).Interior.ColorIndex = 37
    Range("B" & lrow - 1 & ":C" & lrow + k + 1).HorizontalAlignment = xlCenter
    Range("B" & lrow - 1 & ":C" & lrow + k - 1).Borders.LineStyle = xlContinuous
    Range("B" & lrow + k + 1 & ":C" & lrow + k + 1).Borders.LineStyle = xlContinuous
End Sub

Sub Main()
    For Each sht In Worksheets
        Select Case sht.Name:
            Case "Inventory":
                Sheets("Inventory").Delete
                Sheets.Add.Name = "Inventory"
            Case "Shipped":
                Sheets("Shipped").Delete
                Sheets.Add.Name = "Shipped"
        End Select
    Next sht
    ImportInventoryReport
    ImportTailgateReport
    FormatShipped
    MakeTable
    FormatInventory
    MakeTable
    CoilBreakdown
End Sub
