Attribute VB_Name = "MacroStartPoints"
Public f As New SysFunc

'Sort order variables for worksheet
Public sortOrder As XlSortOrder
Public sortColumn As Long

'Filter variables for worksheet
Public w As Worksheet
Public filterArray()
Public currentFiltRange As String
Public col As Integer

'This method ensures proper loading of the Excel Ribbon UI

'List Button
Sub populateList(Control As IRibbonControl)

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.StatusBar = ""
    ActiveSheet.Unprotect
    
    If Not Cells(1, 1) = "UPDATE_ID" Then
        f.captureFilter
    End If


    Call SpecListController.printList(SpecListController.getList)


    ActiveWindow.ScrollRow = 1
    Application.ScreenUpdating = True
    
End Sub
'Updates Display
Sub getUpdates(Control As IRibbonControl)

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.StatusBar = ""
    ActiveSheet.Unprotect
    
    f.captureFilter
    
    Call UpdateListController.list
    
    Cells(1, 1).Select
    
    ActiveWindow.ScrollRow = 1
    Application.ScreenUpdating = True
    
End Sub

'List User Specs
Sub refreshList(Control As IRibbonControl)

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.StatusBar = ""
    ActiveSheet.Unprotect
    
    f.clearFilterData
    
    Call Build
    
    Call SpecListController.printList(SpecListController.getList)
    
    ActiveWindow.ScrollRow = 1
    Application.ScreenUpdating = True
    
End Sub
Sub listAll(Control As IRibbonControl)

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.StatusBar = ""
    ActiveSheet.Unprotect
    
    Debug.Print Err.Number
    
    f.clearFilterData
    
    Call Build(True)
    
    Call SpecListController.printList(SpecListController.getList)
    
    ActiveWindow.ScrollRow = 1
    Application.ScreenUpdating = True
    
End Sub


Sub Add_Edit_Delete__Action(Control As IRibbonControl)

    Application.Calculation = xlCalculationManual
    Application.StatusBar = ""
    ActiveSheet.Unprotect
    
    If UIController.isSpecListView Then
        
        Dim regex As Object
        Set regex = CreateObject("VBScript.RegExp")
        regex.pattern = ".*UPDATE.*"
        
        
        If regex.test(Cells(1, ActiveCell.column)) And Control.id = "Edit" And Cells(ActiveCell.row, f.getHeaderColumnIndexOf("LATEST_UPDATE")) <> "No Updates" And FUpdateItem.action <> "Delete" Then
        
            UpdateListController.specid = Cells(ActiveCell.row, f.getHeaderColumnIndexOf("SPEC_ID"))
            FUpdateItem.action = Control.id
            Range(Cells(ActiveCell.row, f.getHeaderColumnIndexOf("UPDATE_DATE")), Cells(ActiveCell.row, f.getHeaderColumnIndexOf("LATEST_UPDATE"))).Select
            FUpdateItem.show
            
        ElseIf Control.id = "Edit" And regex.test(Cells(1, ActiveCell.column)) And Cells(ActiveCell.row, f.getHeaderColumnIndexOf("LATEST_UPDATE")) = "No Updates" Then
        
            UpdateListController.specid = Cells(ActiveCell.row, f.getHeaderColumnIndexOf("SPEC_ID"))
            FUpdateItem.action = "Add"
            Range(Cells(ActiveCell.row, f.getHeaderColumnIndexOf("UPDATE_DATE")), Cells(ActiveCell.row, f.getHeaderColumnIndexOf("LATEST_UPDATE"))).Select
            FUpdateItem.show
        
        ElseIf Control.id = "Add" And regex.test(Cells(1, ActiveCell.column)) Then
        
            UpdateListController.specid = Cells(ActiveCell.row, f.getHeaderColumnIndexOf("SPEC_ID"))
            FUpdateItem.action = "Add"
            Range(Cells(ActiveCell.row, f.getHeaderColumnIndexOf("UPDATE_DATE")), Cells(ActiveCell.row, f.getHeaderColumnIndexOf("LATEST_UPDATE"))).Select
            FUpdateItem.show
        
        ElseIf Control.id = "Delete" And regex.test(Cells(1, ActiveCell.column)) And Cells(ActiveCell.row, f.getHeaderColumnIndexOf("LATEST_UPDATE")) <> "No Updates" Then
        
            UpdateListController.specid = Cells(ActiveCell.row, f.getHeaderColumnIndexOf("SPEC_ID"))
            FUpdateItem.action = "Delete"
            Range(Cells(ActiveCell.row, f.getHeaderColumnIndexOf("UPDATE_DATE")), Cells(ActiveCell.row, f.getHeaderColumnIndexOf("LATEST_UPDATE"))).Select
            FUpdateItem.show
        
        Else
            FSpecItem.action = Control.id
            FSpecItem.show
        End If
        'Control ID is either Add, Edit, Delete

    ElseIf UIController.isUpdateListView Then
        FUpdateItem.action = Control.id
        FUpdateItem.show

    End If
    
End Sub

Sub createReport(Control As IRibbonControl)
    Application.Calculation = xlCalculationManual
    Application.StatusBar = ""
    ActiveSheet.Unprotect
    On Error GoTo errf

    Dim savePath As String

    Cells.Select
    Range("C6").activate
    Selection.Copy
    Cells(1, 1).Select
    Workbooks.Add
    ActiveSheet.Paste
    Cells(1, 1).Select
    Dim specObj As New spec
    Dim updateObj As New Update
    Dim colValue As Variant
    Dim formatStr As String: formatStr = "SPEC"
    For Each colValue In specObj.getDefaultOrderArray()
        If f.getHeaderColumnIndexOf(CStr(colValue)) = 0 Then
            formatStr = "UPDATE"
            Exit For
        End If
    Next colValue
    f.defaultFormats (formatStr)
    If formatStr = "SPEC" Then f.deleteHiddenColumns
    ActiveWindow.freezePanes = False
    savePath = GetFolder("C:\Users\" & f.getUsername & "\")
    ActiveSheet.SaveAs Filename:=savePath
    ActiveWorkbook.Close
    
errf:
    
End Sub
Sub managePreferences(conrol As IRibbonControl)

    UserPreferences.show

End Sub
Sub quickAddUpdate(Control As IRibbonControl)
    Application.Calculation = xlCalculationManual
    Application.StatusBar = ""
    ActiveSheet.Unprotect
    If Not f.getHeaderColumnIndexOf("UPDATE_ID") Then

        'Control ID is either Add, Edit, Delete
        UpdateListController.specid = Cells(ActiveCell.row, f.getHeaderColumnIndexOf("SPEC_ID"))
        FUpdateItem.action = "Add"
        FUpdateItem.show

    Else
        FUpdateItem.action = "Add"
        FUpdateItem.show

    End If
    
End Sub

Function GetFolder(strPath As String) As String
    Application.Calculation = xlCalculationManual
    Application.StatusBar = ""
Dim fldr As FileDialog
Dim sItem As String
Set fldr = Application.FileDialog(msoFileDialogSaveAs)
With fldr
    .Title = "Select a Folder"
    .AllowMultiSelect = False
    .InitialFileName = strPath
    If .show <> -1 Then GoTo NextCode
    sItem = .SelectedItems(1)
End With
NextCode:
GetFolder = sItem
Set fldr = Nothing
End Function


Sub Checkbox1_onAction(Control As IRibbonControl, Pressed As Boolean)
'
' Code for onAction callback. Ribbon control checkBox
'
    ActiveSheet.Unprotect
    If Pressed Then
        
        StatusBooleans.setStatus completed:=True
    Else
        
        StatusBooleans.setStatus completed:=False
    End If
    
End Sub
Sub Checkbox2_onAction(Control As IRibbonControl, Pressed As Boolean)
'
' Code for onAction callback. Ribbon control checkBox
'
    ActiveSheet.Unprotect
    If Pressed Then
        
        StatusBooleans.setStatus canceled:=True
    Else
        
        StatusBooleans.setStatus canceled:=False
    End If
    
End Sub
Sub Checkbox3_onAction(Control As IRibbonControl, Pressed As Boolean)
'
' Code for onAction callback. Ribbon control checkBox
'
    ActiveSheet.Unprotect
    If Pressed Then
        
        StatusBooleans.setStatus hold:=True
    Else
        
        StatusBooleans.setStatus hold:=False
    End If
    
End Sub
Sub Checkbox4_onAction(Control As IRibbonControl, Pressed As Boolean)
'
' Code for onAction callback. Ribbon control checkBox
'
    ActiveSheet.Unprotect
    If Pressed Then
        
        StatusBooleans.setStatus cerner:=True
    Else
        
        StatusBooleans.setStatus cerner:=False
    End If
    
End Sub
Sub Checkbox5_onAction(Control As IRibbonControl, Pressed As Boolean)
'
' Code for onAction callback. Ribbon control checkBox
'
    ActiveSheet.Unprotect
    If Pressed Then
        
        StatusBooleans.setStatus assigned:=True
    Else
        
        StatusBooleans.setStatus assigned:=False
    End If
    
End Sub
Sub Checkbox6_onAction(Control As IRibbonControl, Pressed As Boolean)
'
' Code for onAction callback. Ribbon control checkBox
'
    ActiveSheet.Unprotect
    If Pressed Then
        
        StatusBooleans.setStatus unassigned:=True
    Else
        
        StatusBooleans.setStatus unassigned:=False
    End If
    
End Sub
Sub manageAnalysts(Control As IRibbonControl)

    Application.Calculation = xlCalculationManual
    MAnalystModify.showAnalystModifyRibbon

End Sub


Public Function getCompletedStatus()
    ActiveSheet.Unprotect
    getCompletedStatus = g_blnCompletedCheckboxState
    ActiveSheet.Protect
End Function
Public Function getCanceledStatus()
    ActiveSheet.Unprotect
    getCanceledStatus = g_blnCanceledCheckboxState
    ActiveSheet.Protect
End Function
Public Function getHoldStatus()
    ActiveSheet.Unprotect
    getHoldStatus = g_blnHoldCheckboxState
    ActiveSheet.Protect
End Function
Public Function getCernerStatus()
    ActiveSheet.Unprotect
    getCernerStatus = g_blnCernerCheckboxState
    ActiveSheet.Protect
End Function
Public Function getAssignedStatus()
    ActiveSheet.Unprotect
    getAssignedStatus = g_blnAssignedCheckboxState
    ActiveSheet.Protect
End Function
Public Function getUnassignedStatus()
    ActiveSheet.Unprotect
    getUnassignedStatus = g_blnUnassignedCheckboxState
    ActiveSheet.Protect
End Function
Sub rbx_onLoad(ribbon As IRibbonUI)
'
' Code for onLoad callback. Ribbon control customUI
'

    
    Set g_rbxUI = ribbon
    ActiveSheet.Protect
End Sub
