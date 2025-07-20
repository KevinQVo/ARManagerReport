Sub FormatData_ARManagerReport()
    Dim ws As Worksheet
    Dim newSheet As Worksheet
    Dim formattedDate As String
    Dim headers As Variant
    Dim fillColors As Variant
    Dim fontColors As Variant
    Dim i As Integer
    Dim lastRow As Long
    Dim iRow As Long

    ' Step 0: Rename current sheet with today's date in mm.dd.yy format
    Set ws = ActiveSheet
    formattedDate = Format(Date, "mm.dd.yy")
    On Error Resume Next
    ws.Name = "Oracle Cloud Aged (" & formattedDate & ")"
    On Error GoTo 0

    ' Step 1: Format current sheet
    With ws.Cells
        .UnMerge
        .WrapText = False
        .WrapText = True
        .WrapText = False
    End With

    ' Step 2: Create new sheet named "template"
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("template").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set newSheet = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    newSheet.Name = "template"

    ' Step 3: Define headers
    headers = Array("Sponsor/RIA", "BLK #", "Qtr", "Qtr Bucket", "Account #", "RPM", _
                    "Termination Date", "Long Title", "Total Fee Due", "Division Type", _
                    "Invoice #", "Owner", "Bucket Status", "Notes", "RIA")

    ' Step 4: Define fill colors
    fillColors = Array( _
        RGB(216, 228, 188), RGB(255, 255, 0), RGB(216, 228, 188), RGB(216, 228, 188), _
        RGB(255, 255, 0), RGB(184, 204, 228), RGB(184, 204, 228), RGB(255, 255, 0), _
        RGB(255, 255, 0), RGB(255, 255, 0), RGB(255, 255, 0), -1, -1, -1, RGB(184, 204, 228)) ' RIA = RPM style

    ' Step 5: Define font colors
    fontColors = Array( _
        RGB(0, 0, 0), RGB(0, 0, 0), RGB(0, 0, 0), RGB(0, 0, 0), RGB(0, 0, 0), _
        RGB(0, 0, 0), RGB(0, 0, 0), RGB(0, 0, 0), RGB(0, 0, 0), RGB(0, 0, 0), _
        RGB(0, 0, 0), RGB(255, 0, 0), RGB(255, 0, 0), RGB(255, 0, 0), RGB(0, 0, 0)) ' RIA now black like RPM

    ' Step 6: Apply headers and formatting
    For i = 0 To UBound(headers)
        With newSheet.Cells(1, i + 1)
            .Value = headers(i)
            .Font.Bold = True
            .Font.Color = fontColors(i)
            If fillColors(i) <> -1 Then
                .Interior.Color = fillColors(i)
            End If
        End With
    Next i

    ' Step 7: Copy data from Oracle Cloud Aged (starting from row 7)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For iRow = 7 To lastRow
        newSheet.Cells(iRow - 6, 2).Value = ws.Cells(iRow, "T").Value ' BLK # → Col B
        newSheet.Cells(iRow - 6, 5).Value = ws.Cells(iRow, "B").Value ' Account # → Col E
        newSheet.Cells(iRow - 6, 8).Value = ws.Cells(iRow, "A").Value ' Long Title → Col H
        newSheet.Cells(iRow - 6, 9).Value = ws.Cells(iRow, "M").Value ' Total Fee Due → Col I
        newSheet.Cells(iRow - 6, 10).Value = ws.Cells(iRow, "S").Value ' Division Type → Col J
        newSheet.Cells(iRow - 6, 11).Value = ws.Cells(iRow, "G").Value ' Invoice # → Col K
    Next iRow

End Sub

-----------------------

    ' Step 12: Add formula to Qtr Bucket (column D)
    newSheet.Range("D2").Formula = _
        "=IF(M2=""REFUND DUE"",""REFUND""," & _
        "IF(M2=""PAYMENT RECEIVED"",""PAYMENT RECEIVED""," & _
        "IF(M2=""KICKOUT"",""KICKOUT""," & _
        "IFERROR(IF(VLOOKUP(C2,Decodes!I:I,1,FALSE)=C2,C2,""Pre 3Q2019""),""Pre 3Q2019""))))"
    
    newSheet.Range("D2").AutoFill Destination:=newSheet.Range("D2:D" & formulaLastRow)

