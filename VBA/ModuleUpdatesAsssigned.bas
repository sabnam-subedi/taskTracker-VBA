Attribute VB_Name = "ModuleUpdatesAsssigned"
Sub UpdateAssignedToDropdowns()
    Dim deptWS As Worksheet, empSheet As Worksheet
    Dim deptCell As Range, empRow As Long, lastEmpRow As Long
    Dim deptName As String, cleanDeptName As String
    Dim empName As String, empDept As String
    Dim empDict As Object: Set empDict = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet
    Dim empStr As String
    Dim empNameItem As Variant

    
    Set deptWS = ThisWorkbook.Sheets("Department_List")
    Set empSheet = ThisWorkbook.Sheets("Employee_List")

    
    lastEmpRow = empSheet.Cells(empSheet.Rows.Count, 1).End(xlUp).Row
    For empRow = 2 To lastEmpRow
        empDept = Trim(empSheet.Cells(empRow, 5).Value)
        empName = Trim(empSheet.Cells(empRow, 1).Value)
        If empDept <> "" And empName <> "" Then
            If Not empDict.exists(empDept) Then
                empDict.Add empDept, New Collection
            End If
            empDict(empDept).Add empName
        End If
    Next empRow

    
    For Each deptCell In deptWS.Range("A2:A" & deptWS.Cells(deptWS.Rows.Count, 1).End(xlUp).Row)
        deptName = Trim(deptCell.Value)
        If deptName <> "" Then
            cleanDeptName = deptName

            
            On Error Resume Next
            Set ws = ThisWorkbook.Sheets(cleanDeptName)
            On Error GoTo 0

            
            If ws Is Nothing Then
                Set ws = ThisWorkbook.Sheets.Add(After:=Sheets(Sheets.Count))
                ws.Name = cleanDeptName
                ws.Range("A1:H1").Value = Array("Task ID", "Task Name", "Due Date", "Priority", "Status", "Date Created", "Remaining Days", "Assign To")
            End If

        
            ws.Range("H2:H100").Validation.Delete

            
            If empDict.exists(deptName) Then
                empStr = ""
                For Each empNameItem In empDict(deptName)
                    empStr = empStr & empNameItem & ","
                Next empNameItem

                If Len(empStr) > 0 Then empStr = Left(empStr, Len(empStr) - 1)

                
                If Len(empStr) <= 255 Then
                    With ws.Range("H2:H100").Validation
                        .Delete
                        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=empStr
                        .IgnoreBlank = True
                        .InCellDropdown = True
                        .ShowInput = True
                        .ShowError = True
                    End With
                Else
                    MsgBox "Dropdown for '" & deptName & "' skipped (too long).", vbExclamation
                End If
            End If
        End If
        Set ws = Nothing
    Next deptCell

    MsgBox "Dropdowns updated for all departments.", vbInformation
End Sub

