Attribute VB_Name = "ModuleRemoveColor"
Sub ClearColorCoding()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Tasks")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ws.Range("A2:H" & lastRow).Interior.ColorIndex = xlNone
End Sub
