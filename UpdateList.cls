VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UpdateList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim f As New SysFunc

Public Function getAllUpdatesFromDB(specid As Integer)

    Dim updateObj As New Update

    Dim dataObj As New DataAccess
    
    dataObj.init
    
    Dim results As Variant
    
    results = dataObj.runQuery("SELECT " & updateObj.getDefaultOrderString() & " FROM [Local_DB].[dbo].[UPDATE] WHERE SPEC_ID = ? ORDER BY UPDATE_DATE DESC", Array(specid))
    
    getAllUpdatesFromDB = results
    
End Function
Public Function getLatestUpdate(spec_id As Variant) As Variant

    spec_id = CInt(spec_id)
    
    Dim dataObj As New DataAccess
    Dim updateObj As New Update
    
    getLatestUpdate = dataObj.runQuery("SELECT TOP 1 " & updateObj.getDefaultOrderString & " FROM [Local_DB].[dbo].[UPDATE] WHERE SPEC_ID = ? ORDER BY UPDATE_DATE DESC", Array(spec_id))

End Function
Public Function addToList(updateObj As Update)

    Dim dataObj As DataAccess
    Set dataObj = New DataAccess
    dataObj.init
    Dim values As Variant
    Dim params As Variant
    params = f.DeleteElementAt(0, updateObj.getDefaultOrderArray())
    values = f.DeleteElementAt(0, updateObj.toVariant)
    Call dataObj.runQuery("INSERT INTO [Local_DB].[dbo].[UPDATE] (" & updateObj.getDefaultString(params) & ") VALUES ( " & f.createParamString(UBound(params)) & " )", values)
    'SpecListController.editedSpecID = updateObj.spec_id
    SpecListController.editedUpdateSpecID = updateObj.spec_id
    
End Function

Public Function updateFromList(updateObj As Update)

    'Remove the node
    Call removeFromList(updateObj)
    'Add the node
    Call addToList(updateObj)
    
End Function

Public Function removeFromList(updateObj As Update)

    Dim dataObj As DataAccess
    Set dataObj = New DataAccess
    dataObj.init
    
    Call dataObj.runQuery("DELETE FROM [Local_DB].[dbo].[UPDATE] WHERE UPDATE_ID = ?", Array(updateObj.update_id))

End Function
