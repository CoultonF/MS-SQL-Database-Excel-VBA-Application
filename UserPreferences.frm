VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserPreferences 
   Caption         =   "User Preferences"
   ClientHeight    =   9165
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7185
   OleObjectBlob   =   "UserPreferences.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserPreferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnAdd_Click()
HiddenColumns.SetFocus
    If HiddenColumns.value <> "" Then
    If HiddenColumns.ListIndex >= 0 Then
    
        VisibleColumns.AddItem HiddenColumns.value
    
        HiddenColumns.RemoveItem HiddenColumns.ListIndex

        VisibleColumns.ListIndex = VisibleColumns.ListCount - 1

        VisibleColumns.Selected(VisibleColumns.ListCount - 1) = True

    End If
    End If
End Sub

Private Sub btnMoveDown_Click()
    VisibleColumns.SetFocus
    If VisibleColumns.value <> "" Then
    Dim colValue As String: colValue = VisibleColumns.value
    
    If VisibleColumns.ListIndex + 2 = VisibleColumns.ListCount Then
        
        VisibleColumns.RemoveItem (VisibleColumns.ListIndex)
    
        VisibleColumns.AddItem colValue
        
    Else
        VisibleColumns.RemoveItem (VisibleColumns.ListIndex)
        VisibleColumns.AddItem colValue, (VisibleColumns.ListIndex + 1) Mod VisibleColumns.ListCount
    End If
    
    VisibleColumns.ListIndex = (VisibleColumns.ListIndex + 1) Mod VisibleColumns.ListCount
    
    End If
End Sub

Private Sub btnMoveUp_Click()
    VisibleColumns.SetFocus
    If VisibleColumns.value <> "" Then
    Dim colValue As String: colValue = VisibleColumns.value
    
    If VisibleColumns.ListIndex = 0 Then
        
        VisibleColumns.RemoveItem (VisibleColumns.ListIndex)
    
        VisibleColumns.AddItem colValue
        VisibleColumns.ListIndex = VisibleColumns.ListCount - 1
    ElseIf VisibleColumns.ListIndex = VisibleColumns.ListCount - 1 Then
        VisibleColumns.AddItem colValue, VisibleColumns.ListIndex - 1
        VisibleColumns.RemoveItem (VisibleColumns.ListIndex)
        
        VisibleColumns.ListIndex = VisibleColumns.ListCount - 2
    Else
        VisibleColumns.RemoveItem (VisibleColumns.ListIndex)
        VisibleColumns.AddItem colValue, (VisibleColumns.ListIndex - 1) Mod VisibleColumns.ListCount
        VisibleColumns.ListIndex = (VisibleColumns.ListIndex - 2) Mod VisibleColumns.ListCount
    End If
    
    
    End If

End Sub

Private Sub btnRemove_Click()
VisibleColumns.SetFocus
    If VisibleColumns.value <> "" Then
    If VisibleColumns.ListIndex >= 0 Then

        HiddenColumns.AddItem VisibleColumns.value
    
        VisibleColumns.RemoveItem VisibleColumns.ListIndex
        
        HiddenColumns.ListIndex = HiddenColumns.ListCount - 1
        
    End If
    End If
End Sub

Private Sub Cancel_Click()

    Call cmdReset_Click
    Unload Me

End Sub


Private Sub Reset_Click()
    
    UIController.invalidateRibbonUI

    MUserPreferences.DeleteAll
    
    WindowsRegistry.setDefaultView
    
    populateCurrentValues

End Sub

Private Sub Submit_Click()

    MUserPreferences.setStatus CStr(ShowCompleted.value), "\|COMPLETED\|"
    MUserPreferences.setStatus CStr(ShowCanceled.value), "\|CANCELED\|"
    MUserPreferences.setStatus CStr(ShowOnHold.value), "\|ONHOLD\|"
    MUserPreferences.setStatus CStr(ShowCernerfix.value), "\|CERNERFIX\|"
    MUserPreferences.setStatus CStr(ShowAssigned.value), "\|ASSIGNED\|"
    MUserPreferences.setStatus CStr(ShowUnassigned.value), "\|UNASSIGNED\|"
    MUserPreferences.setDefaultViewType (CStr(AllSpecs.value))
    MUserPreferences.setSortBy (CStr(SortField.value))
    MUserPreferences.setSortCollation (CStr(isAscending.value))
    If FilterField.value <> "" And FilterOperator <> "" Then MUserPreferences.setCustomFilterValues FilterField.value, FilterOperator.value, FilterValue.value
    If FilterField.value = "" And FilterOperator = "" And FilterValue.value = "" Then MUserPreferences.setCustomFilterValues FilterField.value, FilterOperator.value, FilterValue.value
    MUserPreferences.setVisibleColumns convertListBoxToVariant(VisibleColumns.Object)
    MUserPreferences.setHiddenColumns convertListBoxToVariant(HiddenColumns.Object)
    
    UIController.invalidateRibbonUI
    
    Unload Me
    Call cmdReset_Click
    
    Call MUserPreferences.activate
    Call SpecListController.Build(AllSpecs.value)
    Call SpecListController.printList

End Sub
Private Function convertListBoxToVariant(lb As Object) As Variant

    Dim result() As String
    
    If CallByName(lb, "listCount", VbGet) <> 0 Then
        
        ReDim result(CallByName(lb, "listCount", VbGet) - 1)
        
        Dim i As Long: i = 0
    
        For Each item In result
        
            result(i) = CallByName(lb, "list", VbGet, i)
        
            i = i + 1
        
        Next item
    
        convertListBoxToVariant = result
    End If
End Function

Private Sub UserForm_Activate()

    
    Dim specObj As New spec
    
    specObj.init
    
    FilterField.AddItem ""
    
    For Each item In specObj.getDefaultOrderArray
    
        SortField.AddItem item
        FilterField.AddItem item
    
    Next item

    FilterOperator.AddItem ""
    FilterOperator.AddItem "EQUALS"
    FilterOperator.AddItem "NOT EQUALS"
    
    populateCurrentValues

End Sub

Private Function populateCurrentValues()
    
    If CBool(WindowsRegistry.getDefaultViewType) Then
        
        AllSpecs = True
        
    Else
    
        UserSpecs = True
    
    End If
    
    SortField.value = WindowsRegistry.getDefaultViewSort
    
    isAscending.value = CBool(WindowsRegistry.getDefaultViewCollation)
    
    ShowCompleted.value = MUserPreferences.isCompleted(WindowsRegistry.getDefaultViewCheckbox)
    ShowCanceled.value = MUserPreferences.isCanceled(WindowsRegistry.getDefaultViewCheckbox)
    ShowOnHold.value = MUserPreferences.isOnHold(WindowsRegistry.getDefaultViewCheckbox)
    ShowCernerfix.value = MUserPreferences.isCernerFix(WindowsRegistry.getDefaultViewCheckbox)
    ShowAssigned.value = MUserPreferences.isAssigned(WindowsRegistry.getDefaultViewCheckbox)
    ShowUnassigned.value = MUserPreferences.isUnassigned(WindowsRegistry.getDefaultViewCheckbox)

    FilterField.value = WindowsRegistry.getDefaultViewFilterField
    
    FilterValue.value = WindowsRegistry.getDefaultViewFilterValue
    
    FilterOperator.value = MUserPreferences.getOperatorText(WindowsRegistry.getDefaultViewFilterOperator)
    
    VisibleColumns.list = MUserPreferences.getVisibleColumns
    
    HiddenColumns.list = MUserPreferences.getHiddenColumns
    
End Function


Private Sub cmdReset_Click()

    Dim ctl As MSForms.Control

    For Each ctl In Me.Controls
        Select Case TypeName(ctl)
            Case "TextBox"
                ctl.text = ""
            Case "CheckBox", "OptionButton", "ToggleButton"
                ctl.value = False
            Case "ComboBox", "ListBox"
                ctl.ListIndex = -1
        End Select
    Next ctl

End Sub

