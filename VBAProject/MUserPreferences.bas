Attribute VB_Name = "MUserPreferences"
'This Module will control the:
'   User Preferences Menu Action and state
'

Dim regex As Object

'Takes the user preferences and sets them to the required values or sets to default if not found
Public Function activate()

    Dim f As New SysFunc

    Dim listtype As Boolean

    Dim specObj As New spec

    If WindowsRegistry.registryIsEmpty Or specObj.getDefaultOrderString = "" Then

        WindowsRegistry.setDefaultView
        StatusBooleans.resetDefaults

    Else

        listtype = CBool(WindowsRegistry.getDefaultViewType)
    
        UIController.completed_bool = MUserPreferences.isCompleted(WindowsRegistry.getDefaultViewCheckbox)
        StatusBooleans.setStatus completed:=UIController.completed_bool
    
        UIController.canceled_bool = MUserPreferences.isCanceled(WindowsRegistry.getDefaultViewCheckbox)
        StatusBooleans.setStatus canceled:=UIController.canceled_bool
    
        UIController.onhold_bool = MUserPreferences.isOnHold(WindowsRegistry.getDefaultViewCheckbox)
        StatusBooleans.setStatus hold:=UIController.onhold_bool
    
        UIController.cernerfix_bool = MUserPreferences.isCernerFix(WindowsRegistry.getDefaultViewCheckbox)
        StatusBooleans.setStatus cerner:=UIController.cernerfix_bool
    
        UIController.assigned_bool = MUserPreferences.isAssigned(WindowsRegistry.getDefaultViewCheckbox)
        StatusBooleans.setStatus assigned:=UIController.assigned_bool
    
        UIController.unassigned_bool = MUserPreferences.isUnassigned(WindowsRegistry.getDefaultViewCheckbox)
        StatusBooleans.setStatus unassigned:=UIController.unassigned_bool
    
        f.setSortOrder WindowsRegistry.getDefaultViewSort, MUserPreferences.getAscendingValue
        
        If WindowsRegistry.getDefaultViewFilterField <> "" And WindowsRegistry.getDefaultViewFilterOperator <> "" Then
        f.setFilter WindowsRegistry.getDefaultViewFilterField, WindowsRegistry.getDefaultViewFilterOperator, WindowsRegistry.getDefaultViewFilterValue
        End If
        
        
    End If
    
    activate = listtype

End Function
Public Function hideSpecColumns()

    Dim HiddenColumns As Variant: HiddenColumns = Split(WindowsRegistry.getDefaultViewHiddenColumns, ", ")
    
    Dim hideCol As Variant
    
    Dim f As New SysFunc
    
    For Each hideCol In HiddenColumns
    
        Sheets(1).columns(f.getHeaderColumnIndexOf(CStr(hideCol))).Hidden = True
    
    Next hideCol

End Function
Public Function DeleteAll()

    WindowsRegistry.deleteSpecView

End Function
Public Function setStatus(value As String, status As String)
    
    Set regex = CreateObject("VBScript.RegExp")
    
    Dim checkboxStr As String: checkboxStr = WindowsRegistry.getDefaultViewCheckbox
    
    regex.pattern = status
    
    regex.IgnoreCase = True
    
    If regex.test(checkboxStr) <> CBool(value) Then
        
        If CBool(value) Then
        
            WindowsRegistry.setDefaultViewCheckbox (checkboxStr & Replace(status, "\", ""))
    
        Else
        
            WindowsRegistry.setDefaultViewCheckbox (Replace(checkboxStr, Replace(status, "\", ""), ""))
        
        End If
    
    End If
 
End Function
Public Function setSortOrder(field As String)

    WindowsRegistry.setDefaultViewSort field

End Function
Public Function getAscendingValue()
    
    If CBool(WindowsRegistry.getDefaultViewCollation) Then
        getAscendingValue = xlAscending
    Else
        getAscendingValue = xlDescending
    End If

End Function

Public Function getVisibleColumns() As Variant

    getVisibleColumns = Split(WindowsRegistry.getDefaultViewVisibleColumns, ", ")

End Function

Public Function getHiddenColumns() As Variant

    getHiddenColumns = Split(WindowsRegistry.getDefaultViewHiddenColumns, ", ")

End Function

Public Function setVisibleColumns(values As Variant)

    If isEmpty(values) Then

        WindowsRegistry.setDefaultViewVisibleColumns ""

    Else

        WindowsRegistry.setDefaultViewVisibleColumns Join(values, ", ")

    End If

End Function

Public Function setHiddenColumns(values As Variant)
    
    If isEmpty(values) Then
    
        WindowsRegistry.setDefaultViewHiddenColumns ""
    
    Else
    
        WindowsRegistry.setDefaultViewHiddenColumns Join(values, ", ")

    End If

End Function

Public Function setDefaultViewType(value As String)
    
    WindowsRegistry.setDefaultViewType value
    
End Function
Public Function setSortBy(value As String)

    WindowsRegistry.setDefaultViewSort value

End Function
Public Function setSortCollation(value As String)

    WindowsRegistry.setDefaultViewCollation value

End Function

Public Function isCompleted(registryStr As String) As Boolean
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.IgnoreCase = True
    regex.pattern = "\|COMPLETED\|"
    isCompleted = regex.test(registryStr)
    
End Function

Public Function isCanceled(registryStr As String) As Boolean

    Set regex = CreateObject("VBScript.RegExp")
    regex.IgnoreCase = True
    regex.pattern = "\|CANCELED\|"
    isCanceled = regex.test(registryStr)

End Function

Public Function isOnHold(registryStr As String) As Boolean
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.IgnoreCase = True
    regex.pattern = "\|ONHOLD\|"
    isOnHold = regex.test(registryStr)

End Function

Public Function isCernerFix(registryStr As String) As Boolean

    Set regex = CreateObject("VBScript.RegExp")
    regex.IgnoreCase = True
    regex.pattern = "\|CERNERFIX\|"
    isCernerFix = regex.test(registryStr)

End Function

Public Function isAssigned(registryStr As String) As Boolean

    Set regex = CreateObject("VBScript.RegExp")
    regex.IgnoreCase = True
    regex.pattern = "\|ASSIGNED\|"
    isAssigned = regex.test(registryStr)

End Function

Public Function isUnassigned(registryStr As String) As Boolean

    Set regex = CreateObject("VBScript.RegExp")
    regex.IgnoreCase = True
    regex.pattern = "\|UNASSIGNED\|"
    isUnassigned = regex.test(registryStr)

End Function

Public Function getOperatorText(operator As String)

    Select Case operator
    
    Case "="
        getOperatorText = "EQUALS"
    Case "<>"
        getOperatorText = "NOT EQUALS"
    Case Else
        getOperatorText = ""
    End Select

End Function

Public Function setCustomFilterValues(field As String, operator As String, search As String)

    WindowsRegistry.setDefaultViewFilterField field
    
    WindowsRegistry.setDefaultViewFilterValue search

    Select Case operator
    
    Case "EQUALS"
        WindowsRegistry.setDefaultViewFilterOperator "="
    
    Case "NOT EQUALS"
        WindowsRegistry.setDefaultViewFilterOperator "<>"
    
    Case Else
        WindowsRegistry.setDefaultViewFilterOperator ""
        
    End Select

End Function
