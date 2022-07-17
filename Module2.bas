Attribute VB_Name = "Module2"

Sub Try()

    Dim ws As Worksheet
    Dim wsDest As Worksheet

    Set wsDest = Sheets("Sheet3")

    For Each ws In ActiveWorkbook.Sheets
        If ws.Name <> wsDest.Name Then
            ws.Range("A1", ws.Range("A1").End(xlToRight).End(xlDown)).Copy
            wsDest.Cells(Rows.Count, "A").End(xlUp).Offset(1).PasteSpecial xlPasteValues
        End If
    Next ws

End Sub

Sub Cols()
    Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1:A18") = "eM"
    Columns(1).ColumnWidth = ".0"
End Sub
