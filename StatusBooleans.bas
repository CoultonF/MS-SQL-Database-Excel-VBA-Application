Attribute VB_Name = "StatusBooleans"
'This manages the Show/Hide checkboxes that are displayed in the UI Ribbon.
'The values for the show/hide are stored in a hidden sheet as either True/False

'The hidden status boolean sheet is composed of:

'Col A -> Completed
Private Const completed_col = 1
'Col B -> Canceled
Private Const canceled_col = 2
'Col C -> Hold
Private Const hold_col = 3
'Col D -> Cerner
Private Const cerner_col = 4
'Col E -> Assigned
Private Const assigned_col = 5
'Col F -> Unassigned
Private Const unassigned_col = 6

'Row 1 -> Table headers [Completed, Canceled, Hold, Cerner, Assigned, Unassigned]
Private Const header_row = 1
'Row 2 -> True/False status
Private Const status_row = 2



'To view the sheet run the showSheet function, but remember to hide it afterwards.

'Module variables
Private status(5) As Boolean
Private f As New SysFunc



Private Sub hideSheet()

    Sheets("STATUS BOOLEANS").Visible = xlSheetVeryHidden

End Sub


Private Sub showsheet()

    Sheets("STATUS BOOLEANS").Visible = xlSheetHidden
    Sheets("STATUS BOOLEANS").Visible = True
    
End Sub


Public Function getStatus(status As String)

    Select Case status
    
    Case "completed"
        getStatus = Sheets("STATUS BOOLEANS").Cells(status_row, completed_col).value
    
    Case "canceled"
        getStatus = Sheets("STATUS BOOLEANS").Cells(status_row, canceled_col).value
    
    Case "hold"
        getStatus = Sheets("STATUS BOOLEANS").Cells(status_row, hold_col).value
    
    Case "cerner"
        getStatus = Sheets("STATUS BOOLEANS").Cells(status_row, cerner_col).value
    
    Case "assigned"
        getStatus = Sheets("STATUS BOOLEANS").Cells(status_row, assigned_col).value
    
    Case "unassigned"
        getStatus = Sheets("STATUS BOOLEANS").Cells(status_row, unassigned_col).value
    
    Case Else
    
    End Select

End Function


Public Function setStatus(Optional completed As Variant, Optional canceled As Variant, Optional hold As Variant, Optional cerner As Variant, Optional assigned As Variant, Optional unassigned As Variant)

    If Not IsMissing(completed) Then
        Sheets("STATUS BOOLEANS").Cells(status_row, completed_col).value = completed
    End If

    If Not IsMissing(canceled) Then
        Sheets("STATUS BOOLEANS").Cells(status_row, canceled_col).value = canceled
    End If
    
    If Not IsMissing(hold) Then
        Sheets("STATUS BOOLEANS").Cells(status_row, hold_col).value = hold
    End If
    
    If Not IsMissing(cerner) Then
        Sheets("STATUS BOOLEANS").Cells(status_row, cerner_col).value = cerner
    End If
    
    If Not IsMissing(assigned) Then
        Sheets("STATUS BOOLEANS").Cells(status_row, assigned_col).value = assigned
    End If
    
    If Not IsMissing(unassigned) Then
        Sheets("STATUS BOOLEANS").Cells(status_row, unassigned_col).value = unassigned
    End If

End Function

Public Function resetDefaults()

    setStatus False, False, False, False, True, True

End Function
