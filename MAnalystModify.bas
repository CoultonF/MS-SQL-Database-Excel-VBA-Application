Attribute VB_Name = "MAnalystModify"
Private analystObj As New analyst
Private f As New SysFunc

Public Function showAnalystModify()

    If analystObj.isDBOwner(f.getUsername) Then
    AnalystForm.show
    Unload AnalystMenu
    Call MAnalystSelect.showAnalystSelect
    Else
    MsgBox "You need to have admin access to the Spec Database to manage users."
    End If

End Function
Public Function showAnalystModifyRibbon()

    If analystObj.isDBOwner(f.getUsername) Then
    AnalystForm.show
    Else
    MsgBox "You need to have admin access to the Spec Database to manage users."
    End If
    
End Function
