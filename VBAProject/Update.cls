VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Update"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_iUpdateID As Integer
Private m_sUpdateDesc As String
Private m_sUpdateDate As String
Private m_sUpdateAnalyst As String
Private m_iSpecID As Integer

Private f As New SysFunc

'==========================================================

'||                   GET ALL FIELDS                      ||
'||                 USING WHERE CLAUSE                    ||
'||                   MULTIPLE WHERE                      ||
'||      INDEXES n ARE COMPARED AGAINST EACH OTHER        ||

'==========================================================
Public Function getUpdateWhere(id As Variant, operator As Variant, value As Variant, Optional sortBy As String)

    Dim sql As String
    sql = "SELECT " & getDefaultOrderString() & " FROM dbo.UPDATE WHERE "
    Dim i As Integer
    If UBound(id) = UBound(operator) And UBound(id) = UBound(value) Then
    
        For i = 0 To UBound(id)
            If i <> 0 Then
                sql = sql & "AND"
            End If
            sql = sql & " " & id(i) & " " & operator(i) & " ? "
        
        Next i
    
        If sortBy <> "" Then
            sql = sql & " ORDER BY " & sortBy
        End If
    Dim database As DataAccess
    Set database = New DataAccess
    Call database.init
    getUpdateWhere = database.runQuery(sql, value)
    
    End If
End Function




'==========================================================

'||            MAKE VARIANT TO DICTIONARY                ||


'retuns a dictionary of key => value pairs based on
'default ordering as keys and args as values
'==========================================================

Public Function convertToSpecDict(var As Variant, keys As Variant)
Dim dict As Dictionary
Set dict = New Dictionary
Dim i As Integer
Dim id As Variant
id = keys
For i = 0 To UBound(var, 1)
    Call dict.Add(id(i), var(i))
Next i

Set convertToSpecDict = dict

End Function

'==========================================================

'||         SETS VALUES TO OBJ FROM DICTIONARY           ||


'uses key => value pairs to perform object set
'==========================================================

Public Function setDictionaryToSpec(var As Dictionary)
For Each Key In var.keys

If IsNull(var(Key)) Then
    GoTo NextIteration
End If
    Select Case Key
    Case "UPDATE_ID"
        update_id = var(Key)
    Case "UPDATE_DESC"
        update_desc = var(Key)
    Case "UPDATE_DATE"
        update_date = var(Key)
    Case "UPDATE_ANALYST"
        update_analyst = var(Key)
    Case "SPEC_ID"
        spec_id = var(Key)
    Case Else
    
        Call Err.Raise(0, "ClassModules.Spec.setDictionaryToSpec(" & Key & ") Case not caught, missing.")
    
    End Select
NextIteration:
Next

End Function

'==========================================================

'||                   DEFAULT ORDER                       ||


'determines the order of all query select *s
'==========================================================

Public Function getDefaultOrderArray()
    
    Dim returnValue As Variant
    returnValue = Array("UPDATE_ID", "UPDATE_DESC", "UPDATE_DATE", "UPDATE_ANALYST", "SPEC_ID")
    getDefaultOrderArray = returnValue

End Function

'==========================================================

'||                   STRING OF COL                      ||


'returns all column names seperated by comma

'==========================================================

Public Function getDefaultOrderString()

    Dim result As String
    Dim i As Integer
    If f.IsArrayAllocated(getDefaultOrderArray()) Then
    For i = 0 To UBound(getDefaultOrderArray(), 1) - 1
        result = result & getDefaultOrderArray()(i) + ", "
    Next i
    result = result & getDefaultOrderArray()(i)
    
    End If
    
    getDefaultOrderString = result

End Function
Public Function getDefaultString(arr As Variant)

    Dim result As String
    Dim i As Integer
    If f.IsArrayAllocated(arr) Then
    For i = 0 To UBound(arr, 1) - 1
        result = result & arr(i) + ", "
    Next i
    result = result & arr(i)
    
    End If
    
    getDefaultString = result

End Function
Public Function toVariant()

    Dim attributes As Variant: attributes = getDefaultOrderArray()
    Dim result As Variant
    Dim i As Long
    
    For i = 0 To UBound(attributes)
        Select Case i
        
            Case 0
                attributes(i) = Me.update_id
            Case 1
                attributes(i) = Me.update_desc
            Case 2
                attributes(i) = Me.update_date
            Case 3
                attributes(i) = Me.update_analyst
            Case 4
                attributes(i) = Me.spec_id
            Case Else
            
                Call Err.Raise(0 + i, "ClassModules.Spec.toVariant(" & i & ") Case not caught, missing.")
        
        End Select
        
    Next i
    
    toVariant = attributes
    
End Function

Public Function setVariantToUpdate(values As Variant)
    
    If values(0) = "Auto" Then
    Me.update_id = 0
    Else
    Me.update_id = values(0)
    End If
    Me.update_desc = values(1)
    Me.update_date = values(2)
    Me.update_analyst = values(3)
    Me.spec_id = values(4)

End Function


'==========================================================

'||                 CLASS CONSTRUCTOR                     ||
'||                 CALL .INIT ON OBJ                     ||

'==========================================================
Private Sub Class_Initialize()

    Call clearData
    
End Sub
Public Function clearData()

    m_iUpdateID = 0
    m_sUpdateDesc = ""
    m_sUpdateDate = ""
    m_sUpdateAnalyst = ""
    m_iSpecID = 0
    
End Function



'==========================================================

'||               GETTERS AND SETTERS                     ||

'==========================================================

Public Property Get update_id() As Integer

    update_id = m_iUpdateID

End Property

Public Property Let update_id(ByVal iNewValue As Integer)

    m_iUpdateID = iNewValue

End Property

Public Property Get spec_id() As Integer

    spec_id = m_iSpecID

End Property

Public Property Let spec_id(ByVal iNewValue As Integer)

    m_iSpecID = iNewValue

End Property
Public Property Get update_desc() As String

    update_desc = m_sUpdateDesc

End Property

Public Property Let update_desc(ByVal iNewValue As String)

    m_sUpdateDesc = iNewValue

End Property

Public Property Get update_date() As String

    update_date = m_sUpdateDate

End Property

Public Property Let update_date(ByVal sNewValue As String)

    m_sUpdateDate = sNewValue

End Property

Public Property Get update_analyst() As String

    update_analyst = m_sUpdateAnalyst

End Property

Public Property Let update_analyst(ByVal sNewValue As String)

    m_sUpdateAnalyst = sNewValue

End Property

