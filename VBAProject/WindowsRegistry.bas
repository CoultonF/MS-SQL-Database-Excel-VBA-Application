Attribute VB_Name = "WindowsRegistry"
'## THIS MODULE CONTROLS ALL THE REGISTRY KEYS THAT ARE STORED UNDER SpecDatabase
'## REGISTRY KEYS ARE SIMILAR TO WEB COOKIES. THEY STORE LOCAL DATA RELEVANT TO ONLY THE USER
'## FOR EXAMPLE, USER PREFERENCES THAT ARE MANAGED BY THE USER ONLY

'TO SEE THE REGISTRY KEYS:
    'START MENU > RUN > REGEDIT
    'GOTO COMPUTER\HKEY_CURRENT_USER\Software\VB and VBA ProgramSettings\SpecDatabase

'REGISTRY KEYS ARE STORED IN A FILE STRUCTURE
'ALL REGISTRY KEYS SHOULD BE STORED USING THE CONST SPECDATABASE APPNAME

Private Const appName = "SpecDatabase"

'VIEWSECTION SHOULD ONLY BE USED TO CONTROL DATA THAT IS IN REGARDS TO HOW THE USER VIEWS THE APPLICATION

Private Const viewSection = "DefaultView"
Public Function deleteSpecView()

    DeleteSetting appName

End Function
Public Function registryIsEmpty() As Boolean

    If getDefaultViewType = "" And getDefaultViewSort = "" And getDefaultViewFilter = "" And getDefaultViewCheckbox = "" And getDefaultViewFilterField = "" And getDefaultViewFilterValue = "" And getDefaultViewFilterOperator = "" And getDefaultViewVisibleColumns = "" And getDefaultViewHiddenColumns = "" Then registryIsEmpty = True

End Function

'IF NO SETTINGS ARE FOUND THIS WILL BE CALLED
Public Function setDefaultView()

    'ALL VALUES MUST BE IN STRING FORMAT

    'SETS THE DEFAULT TO FALSE WHICH WILL SET VIEW TO USER-SPECS ON START, "1" WOULD SHOW ALL SPECS
    setDefaultViewType "0"
    
    'SETS THE DEFAULT TO ORDER BY RANK
    setDefaultViewSort "RANK"
    
    'SETS THE DEFAULT WITH NO FILTERS
    setDefaultViewFilter ""
    
    'SETS CHECKBOX VALUES TO ALL UNCHECKED
        ' "|ASSIGNED||UNASSIGNED|" WILL CHECK THE COMPLETED, AND ONHOLD BOXES BY DEFAULT
    setDefaultViewCheckbox "|ASSIGNED||UNASSIGNED|"
    
    '1 MEANS ASCENDING, 0 MEANS DESCENDING - MUST BE IN STRING FORMAT
    setDefaultViewCollation "1"
    
    setDefaultViewFilterField ""
    
    setDefaultViewFilterOperator ""
    
    setDefaultViewFilterValue ""
    
    Dim specObj As New spec
    
    setDefaultViewVisibleColumns "SPEC_ID, RANK, STATUS, DISCIPLINE, DEPARTMENT, SUMMARY, DESCRIPTION, ANALYST, UPDATE_DATE, LATEST_UPDATE, DATE_SUBMITTED, DATE_STARTED, DATE_COMPLETED, VALUE_TO_BUSINESS, CONTACT_NAME, CONTACT_INFO"
    
    setDefaultViewHiddenColumns ""

End Function

Public Function getDefaultViewCollation() As String

    getDefaultViewCollation = GetSetting(appName, viewSection, "Collation")
    
End Function

Public Function getDefaultViewType() As String

    getDefaultViewType = GetSetting(appName, viewSection, "Type")

End Function

Public Function getDefaultViewSort() As String

    getDefaultViewSort = GetSetting(appName, viewSection, "Sort")

End Function

Public Function getDefaultViewFilter() As String

    getDefaultViewFilter = GetSetting(appName, viewSection, "Filter")

End Function

Public Function getDefaultViewCheckbox() As String

    getDefaultViewCheckbox = GetSetting(appName, viewSection, "Checkbox")

End Function

Public Function getDefaultViewFilterValue() As String

    getDefaultViewFilterValue = GetSetting(appName, viewSection, "FilterValue")

End Function

Public Function getDefaultViewFilterOperator() As String

    getDefaultViewFilterOperator = GetSetting(appName, viewSection, "FilterOperator")

End Function

Public Function getDefaultViewFilterField() As String

    getDefaultViewFilterField = GetSetting(appName, viewSection, "FilterField")

End Function
Public Function getDefaultViewVisibleColumns() As String

    getDefaultViewVisibleColumns = GetSetting(appName, viewSection, "VisibleColumns")

End Function
Public Function getDefaultViewHiddenColumns() As String

    getDefaultViewHiddenColumns = GetSetting(appName, viewSection, "HiddenColumns")

End Function
Public Function setDefaultViewVisibleColumns(columns As String)

    SaveSetting appName, viewSection, "VisibleColumns", columns

End Function
Public Function setDefaultViewHiddenColumns(columns As String)

    SaveSetting appName, viewSection, "HiddenColumns", columns

End Function
Public Function setDefaultViewCollation(collation As String)

    SaveSetting appName, viewSection, "Collation", collation

End Function

Public Function setDefaultViewType(vType As String)

    SaveSetting appName, viewSection, "Type", vType

End Function

Public Function setDefaultViewSort(sort As String)

    SaveSetting appName, viewSection, "Sort", sort

End Function

Public Function setDefaultViewFilter(filter As String)

    SaveSetting appName, viewSection, "Filter", filter

End Function

Public Function setDefaultViewCheckbox(checkbox As String)

    SaveSetting appName, viewSection, "Checkbox", checkbox

End Function

Public Function setDefaultViewFilterValue(value As String)

    SaveSetting appName, viewSection, "FilterValue", value
    
End Function

Public Function setDefaultViewFilterOperator(operator As String)

    SaveSetting appName, viewSection, "FilterOperator", operator
    
End Function

Public Function setDefaultViewFilterField(field As String)

    SaveSetting appName, viewSection, "FilterField", field
    
End Function
