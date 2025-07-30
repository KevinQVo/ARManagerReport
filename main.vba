
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
    Dim divCell As Range

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

    ' Step 2: Create new sheet named "Template"
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("template").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set newSheet = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    newSheet.Name = "Template"

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
    destRow = 2

    For sourceRow = 7 To lastRow
        newSheet.Cells(destRow, 2).Value = ws.Cells(sourceRow, "T").Value ' BLK # ? Col B
        newSheet.Cells(destRow, 5).Value = ws.Cells(sourceRow, "B").Value ' Account # ? Col E
        newSheet.Cells(destRow, 8).Value = ws.Cells(sourceRow, "A").Value ' Long Title ? Col H
        newSheet.Cells(destRow, 9).Value = ws.Cells(sourceRow, "M").Value ' Total Fee Due ? Col I
        newSheet.Cells(destRow, 10).Value = ws.Cells(sourceRow, "S").Value ' Division Type ? Col J
        newSheet.Cells(destRow, 11).Value = ws.Cells(sourceRow, "G").Value ' Invoice # ? Col K
        destRow = destRow + 1
    Next sourceRow

    ' Step 8: Formatting the Template worksheet
    With newSheet.Cells
        .Font.Name = "Arial"
        .Font.Size = 9
    End With
    
    ' Step 9: Replace 50A with 26C
    For Each divCell In newSheet.Range("J2:J" & newSheet.Cells(newSheet.Rows.Count, "J").End(xlUp).Row)
        If divCell.Value = "50A" Then
            divCell.Value = "26C"
        End If
    Next divCell
    
    ' Step 10: Sponsor/RIA (Column A)
    Dim formulaLastRow As Long
    formulaLastRow = newSheet.Cells(newSheet.Rows.Count, "B").End(xlUp).Row
    
    newSheet.Range("A2").Formula = _
        "=XLOOKUP(IF(XLOOKUP(B2,'PM Query'!A:A,'PM Query'!G:G,"""")=0,"""",XLOOKUP(B2,'PM Query'!A:A,'PM Query'!G:G,"""")),Vlookup!E:E,Vlookup!F:F,XLOOKUP(J2,Vlookup!A:A,Vlookup!C:C,""""))"
    
    newSheet.Range("A2").AutoFill Destination:=newSheet.Range("A2:A" & formulaLastRow)
    
    ' Step 11: Qtr (Column C)
    newSheet.Range("C2").Formula = "=VLOOKUP(LEFT(K2,8),Decodes!B:C,2,FALSE)"
    newSheet.Range("C2").AutoFill Destination:=newSheet.Range("C2:C" & formulaLastRow)
    
     ' Step 12: Add formula to Qtr Bucket (column D)
    newSheet.Range("D2").Formula = _
        "=IF(M2=""REFUND DUE"",""REFUND""," & _
        "IF(M2=""PAYMENT RECEIVED"",""PAYMENT RECEIVED""," & _
        "IF(M2=""KICKOUT"",""KICKOUT""," & _
        "IFERROR(IF(VLOOKUP(C2,Decodes!I:I,1,FALSE)=C2,C2,""Pre 4Q2025""),""Pre 4Q2025""))))"
    
    newSheet.Range("D2").AutoFill Destination:=newSheet.Range("D2:D" & formulaLastRow)
    
    ' Step 13: RPM (Column F)
    newSheet.Range("F2").Formula = "=VLOOKUP(B2, 'PM Query'!A:P,16,FALSE)"
    newSheet.Range("F2").AutoFill Destination:=newSheet.Range("F2:F" & formulaLastRow)
    
    ' Step 14: RPM (Column G)
    newSheet.Range("G2").Formula = _
    "=IF(VLOOKUP(B2, 'PM Query'!A:P,4,FALSE)=0,"""",VLOOKUP(B2, 'PM Query'!A:P,4,FALSE))"
    newSheet.Range("G2").AutoFill Destination:=newSheet.Range("G2:G" & formulaLastRow)
    newSheet.Range("G2:G" & formulaLastRow).NumberFormat = "mm/dd/yyyy"
    
    'Step 15: RIA (Column 0)
    With newSheet.Range("O2")
        .Formula = "=VLOOKUP(B2, 'PM Query'!A:G,7,FALSE)"
        .AutoFill Destination:=newSheet.Range("O2:O" & formulaLastRow)
    End With
    
    'Step 16: Owner (Column L)
    With newSheet.Range("L2")
        .Formula = "=VLOOKUP(K2, 'OLD Template'!$K:$L,2,FALSE)"
        .AutoFill Destination:=newSheet.Range("L2:L" & formulaLastRow)
    End With
    
    With newSheet.Range("L2:L" & lastRow)
        .Copy
        .PasteSpecial Paste:=xlPasteValues
    End With
    
    'Step 17: Bucket Status (Column M)
    With newSheet.Range("M2")
        .Formula = "=VLOOKUP(K2, 'OLD Template'!$K:$M,3,FALSE)"
        .AutoFill Destination:=newSheet.Range("M2:M" & formulaLastRow)
    End With
    
     With newSheet.Range("M2:M" & lastRow)
        .Copy
        .PasteSpecial Paste:=xlPasteValues
    End With
    
    'Step 18: Notes (Column N)
    With newSheet.Range("N2")
        .Formula = "=VLOOKUP(K2, 'OLD TEMPLATE'!$K:$N,4,FALSE)"
        .AutoFill Destination:=newSheet.Range("N2:N" & formulaLastRow)
    End With
    
    With newSheet.Range("N2:N" & lastRow)
        .Copy
        .PasteSpecial Paste:=xlPasteValues
    End With

End Sub





