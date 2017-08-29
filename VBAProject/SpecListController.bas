Attribute VB_Name = "SpecListController"
Public listObj As SpecList
Public f As New SysFunc
Dim sentinal As Integer
Public editedSpecID As String
Public editedUpdateSpecID As String




Function Build(Optional showAllSpecItems As Boolean)

    Set listObj = New SpecList
    
    listObj.AllSpecs = showAllSpecItems

    Set f = Nothing

    Set f = New SysFunc

    Call f.resetSheet

    Call listObj.makeListAll
    

End Function

Public Function getList()
    
    Set getList = listObj

End Function

Public Function setList(ByRef list As SpecList)

    Set listObj = list

End Function

Public Function printList(Optional newListObj As SpecList)

    ActiveSheet.Unprotect
    Dim lrow As Integer
    If Not newListObj Is Nothing Then
        
        f.resetSheet
        Set listObj = newListObj
        
    End If
    
    Dim rankController As New Rank
    
    rankController.runRankingAlgorithim
    On Error GoTo flag
    
    'populates the listObj in order by SPEC ID
    listObj.inOrder
    
    Dim dict As Collection
    If Not listObj.isFiltered Then
        Set dict = listObj.getAllSpecsFromTree
    Else
        Set dict = listObj.getAllSpecsFromTree
    End If
    
    Dim specObj As New spec
    Set specObj = New spec
    Call f.buildHeader(specObj.getDefaultOrderArray(), 1, 1)
    Call MLoadingUI.LoadSpecsProgressBar(listObj.getAllSpecsFromTree, CLng(listObj.getSize))
    f.defaultFormats "SPEC"
    f.AlternateRowColors
    f.applyFilter
    f.SetRangeBorder
    'changes the update row color when a change is made
    If SpecListController.editedUpdateSpecID <> "" Then
    On Error GoTo continue
        Dim dataObj As New DataAccess
        lrow = Application.match(CLng(SpecListController.editedUpdateSpecID), Range("A:A"), 0)
        
        Cells(lrow, 1).EntireRow.Select
        Dim Data As Variant: Data = dataObj.runQuery("SELECT UPDATE_DATE, LATEST_UPDATE FROM SHAREPOINT WHERE SPEC_ID = ?", Array(CLng(SpecListController.editedUpdateSpecID)))
        Cells(lrow, 9) = Data(0, 0)
        Cells(lrow, 10) = Data(1, 0)
        Selection.FormatConditions.Delete
        Selection.Interior.ColorIndex = 19
        SpecListController.editedSpecID = ""
        ActiveWindow.ScrollRow = lrow
        
        Application.ScreenUpdating = True
    Else
    Cells(1, 1).Select
    End If
    If SpecListController.editedSpecID <> "" Then
    lrow = Application.WorksheetFunction.match(CLng(SpecListController.editedSpecID), columns(f.getHeaderColumnIndexOf("SPEC_ID")), 0)
        Cells(lrow, 1).EntireRow.Select
        Cells(lrow, f.getHeaderColumnIndexOf("UPDATE_DATE")) = dataObj.runQuery("SELECT UPDATE_DATE FROM SHAREPOINT WHERE SPEC_ID = ?", Array(CInt(SpecListController.editedSpecID)))
        Cells(lrow, f.getHeaderColumnIndexOf("LATEST_UPDATE")) = dataObj.runQuery("SELECT LATEST_UPDATE FROM SHAREPOINT WHERE SPEC_ID = ?", Array(CInt(SpecListController.editedSpecID)))
        Selection.FormatConditions.Delete
        Selection.Interior.ColorIndex = 19
        SpecListController.editedSpecID = ""
        Application.ScreenUpdating = True
    
    End If
    
continue:
    Dim editedRow As Range
    

    
    If lrow > 0 Then
        ActiveWindow.ScrollRow = lrow
        lrow = 0
    End If
    Call setFilters
    Exit Function
flag:
If sentinal < 1 Then
    Debug.Print Err.Description
    sentinal = sentinal + 1
    
    Build MUserPreferences.activate()
    Call printList
Else
sentinal = 0
Err.Clear
MsgBox "Failed to build spec list - execution timeout"

End If

End Function
Public Function setFilters()

    If ActiveSheet.AutoFilterMode = False Then
        
        Cells.AutoFilter
        
    End If
    
End Function
Public Function resetTree()
   Set listObj = Nothing
End Function
