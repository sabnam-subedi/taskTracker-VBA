Attribute VB_Name = "Moduleapplycolor"
Sub ApplyColorCoding()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Tasks")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    For i = 2 To lastRow
        With ws.Range("A" & i & ":H" & i)
            Select Case ws.Cells(i, 6).Value
                Case "To-Do"
                    .Interior.Color = RGB(255, 199, 206)
                Case "In Progress"
                    .Interior.Color = RGB(189, 215, 238)
                Case "Done"
                    .Interior.Color = RGB(198, 239, 206)
            End Select
        End With
    Next i

End Sub
