Attribute VB_Name = "UpdateListController"
Public f As New SysFunc
Public listObj As New UpdateList
Private ml_Row As Long
Public specid As Variant



Public Function list()

    Call UpdateListController.printList
    
End Function

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ml_Row = ActiveCell.row
End Sub


Public Function printList(Optional spec_id As Integer)

    ActiveSheet.Unprotect
    If spec_id = 0 Then
        specid = Cells(ActiveCell.row, f.getHeaderColumnIndexOf("SPEC_ID"))
    Else
        specid = spec_id
    End If
    On Error GoTo flag
    specid = CInt(specid)
    Dim results As Variant: results = listObj.getAllUpdatesFromDB(CInt(specid))
    If isEmpty(results) Then
        GoTo flag
    End If
    Call f.resetSheet
    Dim updateObj As New Update
    
    Call f.buildHeader(updateObj.getDefaultOrderArray(), 1, 1)
    
    Call MLoadingUI.LoadUpdatesProgressBar(results, CLng(UBound(results, 2)))
    
    f.defaultFormats "UPDATE"
    
    f.AlternateRowColors
    
    f.SetRangeBorder
    
    Exit Function
flag:

    Call f.resetSheet
    
    Call f.buildHeader(updateObj.getDefaultOrderArray(), 1, 1)
    
    f.defaultFormats "UPDATE"
    
    f.AlternateRowColors
    
    f.SetRangeBorder
End Function
