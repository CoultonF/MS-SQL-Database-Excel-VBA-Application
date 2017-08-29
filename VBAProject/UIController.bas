Attribute VB_Name = "UIController"
Public completed_bool As Boolean
Public canceled_bool As Boolean
Public onhold_bool As Boolean
Public cernerfix_bool As Boolean
Public assigned_bool As Boolean
Public unassigned_bool As Boolean

Public gobjRibbon As IRibbonUI

Public Sub OnLoad(objRibbon As IRibbonUI)

    ActiveSheet.Unprotect
    
    Set gobjRibbon = objRibbon
    
End Sub

Function invalidateRibbonUI()
Err.Clear
On Error GoTo UIERROR
Call gobjRibbon.Invalidate
Exit Function
UIERROR:
    MsgBox "The user interface failed to load properly. This may be caused by having another excel workbook open. Some functionality may not work as expected. Please close all excel workbooks and/or reopen the Spec Database.", vbExclamation, "Error Warning"
    Debug.Print Err.Description
    Err.Clear
End Function

Sub GetPressed(Control As IRibbonControl, ByRef returnValue)

    Select Case Control.id
    
        Case "completed"
            
        Case "canceled"
    
        Case "onhold"
    
        Case "cernerfix"
    
        Case "assigned"
    
        Case "unassigned"
    
        Case Else
    
    End Select

End Sub

Public Sub rxCompleted(ByRef Control As IRibbonControl, ByRef Pressed As Variant)

    Pressed = completed_bool

End Sub

Public Sub rxCanceled(ByRef Control As IRibbonControl, ByRef Pressed As Variant)

    Pressed = canceled_bool

End Sub

Public Sub rxOnHold(ByRef Control As IRibbonControl, ByRef Pressed As Variant)

    Pressed = onhold_bool

End Sub

Public Sub rxCernerFix(ByRef Control As IRibbonControl, ByRef Pressed As Variant)

    Pressed = cernerfix_bool

End Sub

Public Sub rxAssigned(ByRef Control As IRibbonControl, ByRef Pressed As Variant)

    Pressed = assigned_bool

End Sub

Public Sub rxUnassigned(ByRef Control As IRibbonControl, ByRef Pressed As Variant)

    Pressed = unassigned_bool

End Sub
Public Function isSpecListView() As Boolean

    Dim specObj As New spec
    
    Dim f As New SysFunc
    
    Dim col As Variant
    
    Dim headerRowArray As Variant: headerRowArray = f.Flatten(WorksheetFunction.Transpose(ActiveSheet.UsedRange.Rows(1)))

    For Each col In specObj.getDefaultOrderArray()
    
        If col <> headerRowArray(i) Then
        
            isSpecListView = False
        
            Exit For
            
        End If
        
        isSpecListView = True
        
        i = i + 1
    
    Next col

End Function
Public Function isUpdateListView() As Boolean

    Dim updateObj As New Update
    
    Dim f As New SysFunc
    
    Dim col As Variant
    
    Dim headerRowArray As Variant: headerRowArray = f.Flatten(WorksheetFunction.Transpose(ActiveSheet.UsedRange.Rows(1)))

    For Each col In updateObj.getDefaultOrderArray()
    
        If col <> headerRowArray(i) Then
        
            isUpdateListView = False
        
            Exit For
            
        End If
        
        isUpdateListView = True
        
        i = i + 1
    
    Next col

End Function
