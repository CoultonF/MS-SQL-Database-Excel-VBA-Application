VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SpecList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim binaryT As New Tree
Dim f As New SysFunc
Public AllSpecs As Boolean
Dim sortValue As String
Dim filterValues As String
Dim collationValue As String
Dim ShowCompleted As Boolean
Dim ShowAssigned As Boolean
Dim ShowUnassigned As Boolean
Dim showHold As Boolean
Dim ShowCernerfix As Boolean


Public Function getSize()

    getSize = binaryT.CountNodes()

End Function

Public Function getMaxSpec()

    Dim specObj As spec
    Set specObj = New spec

    Dim dataObj As DataAccess
    Set dataObj = New DataAccess
    dataObj.init
    Dim results As Variant
    results = dataObj.runQuery("SELECT MAX(SPEC_ID) FROM dbo.SPEC")(0, 0)
    
    getMaxSpec = CInt(results)

End Function


Public Function makeListAll()

    Dim specObj As spec

    'specObj.convertToSpecDict ()
    
    Dim i As Integer
    
    Dim Data As Variant:
    If AllSpecs Then
    Data = getAllSpecsFromDB()
    Else
    Data = getMainSpecsFromDB()
    End If
    'results = f.TransposeArray(results)
    If f.IsArrayAllocated(Data) Then
    For i = 0 To UBound(Data, 2)
    
        Set specObj = New spec
        
        specObj.init
        
        Dim item As Variant: item = f.sliceVariant(Data, i)
        
        Call specObj.setDictionaryToSpec(specObj.convertToSpecDict(item, specObj.getDefaultOrderArray()))
        
        Call binaryT.Insert(specObj)
        
        Next i
    Application.StatusBar = ""
    Else
    Application.StatusBar = "NO SPEC ITEMS FOUND."
    End If
    
End Function
Public Function getMainSpecsFromDB()

    Dim specObj As spec
    Set specObj = New spec
    Dim analystObj As New analyst

    Dim dataObj As DataAccess
    Set dataObj = New DataAccess
    dataObj.init
    Dim results As Variant
    Dim t As Single
    analystObj.getAnalystFirstNameFromUsername (f.getUsername())
    Dim firstName As String: firstName = analystObj.first_name
    Dim showStatus As String
        showStatus = buildStatusString(True)
    If specObj.getDefaultOrderString() = "" Then
        MUserPreferences.DeleteAll
        MUserPreferences.activate
    End If
    Dim sql As String: sql = "SELECT " & specObj.getDefaultOrderString() & " FROM [Local_DB].[dbo].[SHAREPOINT] WHERE ANALYST LIKE ? " & showStatus & " ORDER BY STATUS, RANK"
    results = dataObj.runQuery(sql, Array("%" & firstName & "%"))
    'results = dataObj.runQuery("SELECT * FROM SHAREPOINT " & filterValues & " ORDER BY " & sortValue & " " & collationValue)

    
    getMainSpecsFromDB = results

End Function
'builds the SQL String for the status checkboxes
Private Function buildStatusString(Optional firstBool As Boolean)
Dim first As Boolean: first = True

If firstBool Then first = False

    If Not StatusBooleans.getStatus("completed") Then
    
        If first Then
        first = False
        buildStatusString = "WHERE"
        Else
        buildStatusString = buildStatusString & "AND"
        End If
        
        buildStatusString = buildStatusString & " STATUS <> 'Completed' "
    End If
    If Not StatusBooleans.getStatus("canceled") Then
    
        If first Then
        first = False
        buildStatusString = "WHERE"
        Else
        buildStatusString = buildStatusString & "AND"
        End If
        
        buildStatusString = buildStatusString & " STATUS <> 'Canceled' "
    End If
    If Not StatusBooleans.getStatus("hold") Then
    
        If first Then
        first = False
        buildStatusString = "WHERE"
        Else
        buildStatusString = buildStatusString & "AND"
        End If
        
        buildStatusString = buildStatusString & " STATUS <> 'Hold' "
    End If
    If Not StatusBooleans.getStatus("cerner") Then
    
        If first Then
        first = False
        buildStatusString = "WHERE"
        Else
        buildStatusString = buildStatusString & "AND"
        End If
        
        buildStatusString = buildStatusString & " STATUS <> 'Cerner Fix' "
    End If
    If Not StatusBooleans.getStatus("assigned") Then
    
        If first Then
        first = False
        buildStatusString = "WHERE"
        Else
        buildStatusString = buildStatusString & "AND"
        End If
        
        buildStatusString = buildStatusString & " STATUS <> 'Assigned' "
    End If
    If Not StatusBooleans.getStatus("unassigned") Then
    
        If first Then
        first = False
        buildStatusString = "WHERE"
        Else
        buildStatusString = buildStatusString & "AND"
        End If
        
        buildStatusString = buildStatusString & " STATUS <> 'Unassigned' "
    End If
    
    If first Then
    
        buildStatusString = ""
    
    End If
End Function
Private Function getShowStatusSQLString()
    Dim isFirst As Boolean: isFirst = True
    Dim showStatus As String: showStatus = ""
    If ShowCompleted Then
        showStatus = "WHERE STATUS IN ("
        showStatus = showStatus & "'Completed'"
    End If
    If ShowAssigned Then
        If Not isFirst Then
            showStatus = showStatus & ", "
        Else
            showStatus = "WHERE STATUS IN ("
            isFirst = False
        End If
        showStatus = showStatus & "'Assigned'"
    End If
    If ShowUnassigned Then
        If Not isFirst Then
            showStatus = showStatus & ", "
        Else
            showStatus = "WHERE STATUS IN ("
            isFirst = False
        End If
        showStatus = showStatus & "'Unassigned'"
    End If
    If showHold Then
        If Not isFirst Then
            showStatus = showStatus & ", "
        Else
            showStatus = "WHERE STATUS IN ("
            isFirst = False
        End If
        showStatus = showStatus & "'Hold'"
    End If
    If ShowCernerfix Then
        If Not isFirst Then
            showStatus = showStatus & ", "
        Else
            showStatus = "WHERE STATUS IN ("
            isFirst = False
        End If
        showStatus = showStatus & "'Cerner Fix'"
    End If
    If showStatus <> "" Then
        showStatus = showStatus & ")"
    End If
    getShowStatusSQLString = showStatus
End Function
Public Function getAllSpecsFromDB()

    Dim specObj As spec
    Set specObj = New spec

    Dim dataObj As DataAccess
    Set dataObj = New DataAccess
    dataObj.init
    Dim results As Variant
    Dim t As Single
    Dim showStatus As String

    showStatus = buildStatusString

    If specObj.getDefaultOrderString() = "" Then
        MUserPreferences.DeleteAll
        MUserPreferences.activate
    End If
    results = dataObj.runQuery("SELECT " & specObj.getDefaultOrderString() & " FROM [Local_DB].[dbo].[SHAREPOINT] " & showStatus & " ORDER BY STATUS, RANK")
    'results = dataObj.runQuery("SELECT * FROM SHAREPOINT " & filterValues & " ORDER BY " & sortValue & " " & collationValue)

    
    getAllSpecsFromDB = results
    
End Function

Public Function getAllSpecsFromTree()
    
    Set getAllSpecsFromTree = binaryT.getCollection

End Function


'==========================================================

'||                   ADDING SPEC                        ||


'Adds a variant to the binary tree as a node object AND DB
'==========================================================

Public Function addToList(specObj As spec)

    Dim tempRank As Variant: tempRank = specObj.Rank
    Dim tempSpec As Integer: tempSpec = specObj.spec_id

    Dim dataObj As DataAccess
    Set dataObj = New DataAccess
    dataObj.init
    Dim values As Variant
    Dim params As Variant
    Dim specIndex As Long: specIndex = Application.WorksheetFunction.match("SPEC_ID", specObj.getDefaultArray(), 0) - 1
    params = f.DeleteElementAt(specIndex, specObj.getDefaultArray())
    values = f.DeleteElementAt(specIndex, specObj.toVariant())
    Dim rankIndex As Long: rankIndex = Application.WorksheetFunction.match("RANK", params, 0) - 1
    values(rankIndex) = "NULL"
    
    Call dataObj.runQuery("INSERT INTO SPEC (" & specObj.getDefaultString(params) & ") VALUES ( " & f.createParamString(UBound(params)) & " )", values)
    If IsNumeric(tempRank) And tempRank > 0 Then
        tempRank = CInt(tempRank)
        Dim rankObj As New Rank
        Call rankObj.incrementRanks(CInt(getMaxSpec), insertingRank:=CInt(tempRank))
        Call rankObj.runRankingAlgorithim
    End If
    
    specObj.Rank = tempRank
    specObj.spec_id = getMaxSpec
    Call binaryT.Insert(specObj)
    
End Function

'==========================================================

'||                   DELETE  SPEC                        ||


'Adds a variant to the binary tree as a node object AND DB
'==========================================================

Public Function removeFromList(specObj As spec)

    Dim dataObj As DataAccess
    Set dataObj = New DataAccess
    dataObj.init
    
    Call dataObj.runQuery("DELETE FROM dbo.SPEC WHERE SPEC_ID = ?", Array(specObj.spec_id))
    Call binaryT.setRootNode(binaryT.removeNode(binaryT.getRootNode, specObj.spec_id))
    
End Function

'==========================================================

'||                   Update  SPEC                        ||


'Adds a variant to the binary tree as a node object AND DB
'==========================================================

Public Function updateFromList(specObj As spec)

    Dim tempRank As Variant: tempRank = specObj.Rank
    Dim tempSpec As Integer: tempSpec = specObj.spec_id

    Dim dataObj As DataAccess
    Set dataObj = New DataAccess
    dataObj.init
    Dim values As Variant
    Dim params As Variant
    Dim specIndex As Long: specIndex = Application.WorksheetFunction.match("SPEC_ID", specObj.getDefaultArray(), 0) - 1
    params = f.DeleteElementAt(specIndex, specObj.getDefaultArray())
    values = f.DeleteElementAt(specIndex, specObj.toVariant())
    Dim rankIndex As Long: rankIndex = Application.WorksheetFunction.match("RANK", params, 0) - 1
    params = f.DeleteElementAt(rankIndex, params)
    values = f.DeleteElementAt(rankIndex, values)
    
    'NEED TO GET THE CURRENT VALUE OF THE UPDATE RANK HERE AND THEN DETERMINE IF NULL THEN RUN INSERT RANK ALGORITHM
    
    'Dim currentSpecRank As Variant: currentSpecRank = dataObj.runQuery("SELECT RANK FROM SPEC WHERE SPEC_ID = ?", Array(tempSpec))
    
    Dim sqlStr As String: sqlStr = "UPDATE [dbo].[SPEC] SET "
    Dim i As Long
    For i = 0 To UBound(params, 1)
        
        sqlStr = sqlStr & "[" & params(i) & "] = ?"
        
        If i <> UBound(params, 1) Then sqlStr = sqlStr & ", "

    Next i
    
    sqlStr = sqlStr & " WHERE SPEC_ID = " + CStr(specObj.spec_id)
    
    Call dataObj.runQuery(sqlStr, values)
    Dim rankObj As New Rank
    If tempRank = "" Then
    
        Call rankObj.incrementRanks(tempSpec, deletingRank:=True)
        Call rankObj.runRankingAlgorithim
    
    ElseIf tempRank > 0 Then
        
        Call rankObj.incrementRanks(tempSpec, updatingrank:=CInt(tempRank))
        Call rankObj.runRankingAlgorithim
    End If
    
    specObj.Rank = tempRank
    specObj.spec_id = tempSpec
    'Call binaryT.
    Call binaryT.setRootNode(binaryT.removeNode(binaryT.getRootNode, specObj.spec_id))
    Call binaryT.Insert(specObj)

End Function



Public Function inOrder()
    Call binaryT.WalkInOrder
End Function

Public Function deleteList()

    Set binaryT = New Tree

End Function

Public Function isFiltered()

    If valueFilters <> "" Or collationValue <> "" Or sortValue <> "Summary" Then
    isFiltered = True
    Else
    isFiltered = False
    End If

End Function

Private Sub Class_Initialize()

'sortValue temporarily set to build the binary tree efficiently
sortValue = "Summary"

Set f = New SysFunc

If Not SpecListController.listObj Is Nothing Then
    
    'binaryT = SpecTreeController.listing

Else

    Set binaryT = New Tree
    binaryT.init
    Call SpecListController.setList(Me)

End If


End Sub
