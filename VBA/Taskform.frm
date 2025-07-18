VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Taskform 
   Caption         =   "Task Form"
   ClientHeight    =   10404
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   13932
   OleObjectBlob   =   "Taskform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Taskform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBoxCategory_Change()
  
End Sub

Private Sub CommandButtonCancel_Click()
    Unload Me
End Sub
Private Sub CommandButtonSubmit_Click()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Tasks")

    Dim dueDate As Date
    Dim nextRow As Long
    Dim selectedPriority As String
    Dim selectedStatus As String

    If Trim(TextBoxTaskName.Value) = "" Then
        MsgBox "Please enter a task name.", vbExclamation
        Exit Sub
    End If

    If IsDate(TextBoxdate.Value) Then
        dueDate = CDate(TextBoxdate.Value)
        If dueDate < Date Then
            MsgBox "Due Date cannot be in the past!", vbExclamation
            Exit Sub
        End If
    Else
        MsgBox "Please enter a valid date in DD/MM/YYYY format.", vbExclamation
        Exit Sub
    End If

    If OptionButtonLow.Value = True Then
        selectedPriority = "Low"
    ElseIf OptionButtonMedium.Value = True Then
        selectedPriority = "Medium"
    ElseIf OptionButtonHigh.Value = True Then
        selectedPriority = "High"
    Else
        MsgBox "Please select a priority level.", vbExclamation
        Exit Sub
    End If

    If OptionButtontodo.Value = True Then
        selectedStatus = "To-Do"
    ElseIf OptionButtonInprogress.Value = True Then
        selectedStatus = "In Progress"
    ElseIf OptionButtonDone.Value = True Then
        selectedStatus = "Done"
    Else
        MsgBox "Please select a task status.", vbExclamation
        Exit Sub
    End If

    nextRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1

    ws.Cells(nextRow, 1).Value = nextRow - 1
    ws.Cells(nextRow, 2).Value = TextBoxTaskName.Value
    ws.Cells(nextRow, 3).Value = dueDate
    ws.Cells(nextRow, 4).Value = selectedPriority
    ws.Cells(nextRow, 5).Value = ComboBoxCategory.Value
    ws.Cells(nextRow, 6).Value = selectedStatus
    ws.Cells(nextRow, 7).Value = Format(Date, "dd/mm/yyyy")
    
    Dim deptSheet As Worksheet
    On Error Resume Next
    Set deptSheet = ThisWorkbook.Sheets(ComboBoxCategory.Value)

    On Error GoTo 0

    If Not deptSheet Is Nothing Then
        Dim deptNextRow As Long
        deptNextRow = deptSheet.Cells(deptSheet.Rows.Count, "A").End(xlUp).Row + 1
    
        deptSheet.Cells(deptNextRow, 1).Value = nextRow - 1
        deptSheet.Cells(deptNextRow, 2).Value = TextBoxTaskName.Value
        deptSheet.Cells(deptNextRow, 3).Value = dueDate
        deptSheet.Cells(deptNextRow, 4).Value = selectedPriority
        deptSheet.Cells(deptNextRow, 5).Value = selectedStatus
        deptSheet.Cells(deptNextRow, 6).Value = Format(Date, "dd/mm/yyyy")
         With deptSheet.Cells(deptNextRow, 7)
        .Formula = "=C" & deptNextRow & " - TODAY()"
        .NumberFormat = "0"
    End With
    End If

    

    With ws.Cells(nextRow, 8)
        .Formula = "=C" & nextRow & " - TODAY()"
        .NumberFormat = "0"
    End With

    With ws.Range("A" & nextRow & ":G" & nextRow)
        Select Case selectedStatus
            Case "To-Do"
                .Interior.Color = RGB(255, 199, 206)
            Case "In Progress"
                .Interior.Color = RGB(189, 215, 238)
            Case "Done"
                .Interior.Color = RGB(198, 239, 206)
        End Select
    End With

    MsgBox "Task added successfully!", vbInformation

    TextBoxTaskName.Value = ""
    TextBoxdate.Value = Format(Date, "dd/mm/yyyy")
    OptionButtonLow.Value = False
    OptionButtonMedium.Value = False
    OptionButtonHigh.Value = False
    ComboBoxCategory.Value = ""
    OptionButtontodo.Value = False
    OptionButtonInprogress.Value = False
    OptionButtonDone.Value = False
    
    

End Sub

Private Sub LabelDate_Click()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    With ComboBoxCategory
        .Clear
        .RowSource = "DepartmentList"
    End With
End Sub
