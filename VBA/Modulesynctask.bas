Attribute VB_Name = "Modulesynctask"
Sub SyncTasksToDepartments()
    Dim wsTasks As Worksheet
    Dim deptWS As Worksheet
    Dim lastRow As Long, deptLastRow As Long
    Dim i As Long
    Dim taskid As Variant, taskName As String, dueDate As Variant
    Dim priority As String, Status As String
    Dim datecreated As Variant, remaining As Long
    Dim deptName As String

    Set wsTasks = ThisWorkbook.Sheets("Tasks")
    lastRow = wsTasks.Cells(wsTasks.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRow
        taskid = wsTasks.Cells(i, 1).Value
        taskName = wsTasks.Cells(i, 2).Value
        dueDate = wsTasks.Cells(i, 3).Value
        priority = wsTasks.Cells(i, 4).Value
        deptName = wsTasks.Cells(i, 5).Value
        Status = wsTasks.Cells(i, 6).Value
        datecreated = wsTasks.Cells(i, 7).Value
        remaining = wsTasks.Cells(i, 8).Value

      
        If taskName <> "" And wsTasks.Cells(i, 5).Value <> "" Then
            On Error Resume Next
            Set deptWS = ThisWorkbook.Sheets(deptName)
            On Error GoTo 0

            If Not deptWS Is Nothing Then
                ' Check for duplicate Task ID
                If Application.CountIf(deptWS.Range("A:A"), taskid) = 0 Then
                    deptLastRow = deptWS.Cells(deptWS.Rows.Count, "A").End(xlUp).Row + 1

                    ' Write task details to the department sheet
                    deptWS.Cells(deptLastRow, 1).Value = taskid
                    deptWS.Cells(deptLastRow, 2).Value = taskName
                    deptWS.Cells(deptLastRow, 3).Value = dueDate
                    deptWS.Cells(deptLastRow, 4).Value = priority
                    deptWS.Cells(deptLastRow, 5).Value = Status
                    deptWS.Cells(deptLastRow, 6).Value = datecreated
                    deptWS.Cells(deptLastRow, 7).Value = remaining
                End If
            End If
        End If
    Next i

    MsgBox "Tasks synced to department sheets.", vbInformation
End Sub

