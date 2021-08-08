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

Sub ImportReport()
    Dim fn$
    Dim wkb1, wkb2 As Workbook
    Dim sht1, sht2 As Worksheet
    
    fn$ = Application.GetOpenFilename("All XLSM files (*.xlsx*), *.xlsx", , "Choose Inventory Report:")
    Application.ScreenUpdating = False
    Set wkb1 = ThisWorkbook
    Set wkb2 = Workbooks.Open(fn$)
    Set sht1 = wkb1.Sheets("Boscus EOM Inventory Report")
    Set sht2 = wkb2.Sheets("Report - BOSCUS")
    sht2.Cells.Copy
    sht1.Range("A1").PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    wkb2.Close
    
    Application.ScreenUpdating = True
    
End Sub

Sub FormatEOMReport()
    Dim wkb As Workbook
    Dim sht As Worksheet
    Dim i, lRow As Integer
    Dim delArr As Variant
    Dim table As ListObject
    Dim rng As Range
    
    Application.ScreenUpdating = False
    
    delArr = Array("Client Inventory", "Order", "Reception Date", "Product Type", "Quantity", "Description", "Lot Number", "Scope", "Weight per Unit", "Weight Unit", "Volume per Unit", "Volume Unit", "Quantity per Package")
    Set wkb = ThisWorkbook
    Set sht = wkb.Sheets("Boscus EOM Inventory Report")
    sht.Activate
    
    Sheets("Boscus EOM Inventory Report").Delete
    Sheets.Add.Name = "Boscus EOM Inventory Report"
    
    ImportReport
    
    Rows(1).EntireRow.Delete

    i = 1
    Do While i < 59
        If Not IsEmpty(Cells(1, i)) And Not IsInArray(Cells(1, i), delArr) Then
            Columns(i).EntireColumn.Delete
        Else
            i = i + 1
        End If
    Loop
    
    lRow = Cells(Rows.Count, "A").End(xlUp).Row
    Set rng = Range(Range("A1"), Range("M" & lRow))
    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, rng, , xlYes)
    tbl.TableStyle = "TableStyleLight9"
    Rows(1).EntireRow.HorizontalAlignment = xlCenter
    
    Application.ScreenUpdating = True
    
End Sub
