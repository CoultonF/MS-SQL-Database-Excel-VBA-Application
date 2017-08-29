VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Spec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_ispec_id As Integer
Private m_vrank As Variant
Private m_sstatus As String
Private m_sdiscipline As String
Private m_sdepartment As String
Private m_ssummary As String
Private m_sdescription As String
Private m_supdate As String
Private m_sanalyst As String
Private m_sdate_submitted As String
Private m_sdate_started As String
Private m_sdate_completed As String
Private m_svalue_to_business As String
Private m_slatest_update As String
Private m_supdate_date As String
Private m_scontact_name As String
Private m_scontact_info As String

'for access to generic functions
Private f As SysFunc


'==========================================================

'||                   GET ALL FIELDS                      ||
'||                 USING WHERE CLAUSE                    ||
'||                   MULTIPLE WHERE                      ||
'||      INDEXES n ARE COMPARED AGAINST EACH OTHER        ||

'==========================================================
Public Function getSpecWhere(id As Variant, operator As Variant, value As Variant, Optional sortBy As String)

    Dim sql As String
    sql = "SELECT " & getDefaultOrderString() & " FROM SHAREPOINT WHERE "
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
    getSpecWhere = database.runQuery(sql, value)
    
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
    Case "SPEC_ID"
        If var(Key) <> "Auto" Then
        spec_id = var(Key)
        Else
        spec_id = 0
        End If
    Case "RANK"
        If Not IsNumeric(var(Key)) Then
            var(Key) = ""
        End If
        Rank = var(Key)
    Case "STATUS"
        status = var(Key)
    Case "DISCIPLINE"
        Discipline = var(Key)
    Case "DEPARTMENT"
        Department = var(Key)
    Case "SUMMARY"
        Summary = var(Key)
    Case "DESCRIPTION"
        Description = var(Key)
    Case "ANALYST"
        analyst = var(Key)
    Case "UPDATE_DATE"
        update_date = var(Key)
    Case "DATE_SUBMITTED"
        Date_Submitted = var(Key)
    Case "DATE_STARTED"
        Date_Started = var(Key)
    Case "DATE_COMPLETED"
        Date_Completed = var(Key)
    Case "LATEST_UPDATE"
        latest_update = var(Key)
    Case "VALUE_TO_BUSINESS"
        Value_To_Business = var(Key)
    Case "CONTACT_NAME"
        Contact_Name = var(Key)
    Case "CONTACT_INFO"
        Contact_Info = var(Key)
    Case Else
    
        Call Err.Raise(0, "ClassModules.Spec.setDictionaryToSpec(" & Key & ") Case not caught, missing.")
    
    End Select
NextIteration:
Next

End Function

'==========================================================

'||                   DEFAULT ORDER                       ||


'determines the order of all query select *s for view
'return: visible + hidden columns
'==========================================================

Public Function getDefaultOrderArray() As Variant
    
    Dim returnValue As Variant
    Dim f As New SysFunc
    returnValue = Split(f.joinCommaDelimitedStrings(WindowsRegistry.getDefaultViewVisibleColumns, WindowsRegistry.getDefaultViewHiddenColumns), ", ")
    getDefaultOrderArray = returnValue

End Function
Private Function getSpecColumnNames()

    Dim dataObj As New DataAccess
    
    getSpecColumnNames = dataObj.runQuery("SELECT Column_name FROM Local_DB.INFORMATION_SCHEMA.Columns WHERE TABLE_NAME = 'SPEC'")

End Function
Public Function getDefaultArray() As Variant

    Dim f As New SysFunc
    
    getDefaultArray = f.getVariantMatchesFrom(getSpecColumnNames, Me.getDefaultOrderArray)
    
    'Dim returnValue As Variant
    'returnValue = Array("SPEC_ID", "RANK", "STATUS", "DISCIPLINE", "DEPARTMENT", "SUMMARY", "DESCRIPTION", "ANALYST", "DATE_SUBMITTED", "DATE_STARTED", "DATE_COMPLETED", "VALUE_TO_BUSINESS", "CONTACT_NAME", "CONTACT_INFO")
    'getDefaultArray = returnValue

End Function

'==========================================================

'||                   STRING OF COL                      ||


'returns all column names seperated by comma

'==========================================================

Public Function getDefaultOrderString() As String

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

    Dim attributes As Variant: attributes = getDefaultArray()
    Dim result As Variant
    Dim i As Long
    
    For i = 0 To UBound(attributes)
        
        attributes(i) = CallByName(Me, attributes(i), VbGet)
        
    Next i
    
    toVariant = attributes
    
End Function

'==========================================================

'||                 CLASS INITIALIZER                     ||
'||                  SETS PARAMETERS                      ||

'==========================================================
Public Function init(Optional spec As Integer, Optional Rank As Integer, Optional status As String, Optional Discipline As String, Optional Department As String, Optional Summary As String, Optional Description As String, Optional Update As String, Optional analyst As String, Optional Date_Submitted As String, Optional Date_Started As String, Optional Date_Completed As String, Optional Value_To_Business As String)


End Function


'==========================================================

'||                 CLASS CONSTRUCTOR                     ||
'||                 CALL .INIT ON OBJ                     ||

'==========================================================
Private Sub Class_Initialize()

    Set f = New SysFunc

    m_ispec_id = 0
    m_vrank = ""
    m_sstatus = ""
    m_sdiscipline = ""
    m_sdepartment = ""
    m_ssummary = ""
    m_sdescription = ""
    m_supdate = ""
    m_sanalyst = ""
    m_sdate_submitted = ""
    m_sdate_started = ""
    m_sdate_completed = ""
    m_svalue_to_business = ""
    m_scontact_name = ""
    m_scontact_info = ""
End Sub
Public Function clearData()
    m_ispec_id = 0
    m_vrank = Null
    m_sstatus = ""
    m_sdiscipline = ""
    m_sdepartment = ""
    m_ssummary = ""
    m_sdescription = ""
    m_supdate = ""
    m_sanalyst = ""
    m_sdate_submitted = ""
    m_sdate_started = ""
    m_sdate_completed = ""
    m_svalue_to_business = ""
    m_scontact_name = ""
    m_scontact_info = ""
End Function



'==========================================================

'||               GETTERS AND SETTERS                     ||

'==========================================================

Public Property Get spec_id() As Integer

    spec_id = m_ispec_id

End Property

Public Property Let spec_id(ByVal iNewValue As Integer)

    m_ispec_id = iNewValue

End Property

Public Property Get Rank() As Variant

    Rank = m_vrank

End Property

Public Property Let Rank(ByVal iNewValue As Variant)

    m_vrank = iNewValue

End Property
Public Property Get latest_update() As String

    latest_update = m_slatest_update

End Property

Public Property Let latest_update(ByVal iNewValue As String)

    m_slatest_update = iNewValue

End Property
Public Property Get update_date() As String

    update_date = m_supdate_date

End Property

Public Property Let update_date(ByVal iNewValue As String)

    m_supdate_date = iNewValue

End Property

Public Property Get status() As String

    status = m_sstatus

End Property

Public Property Let status(ByVal sNewValue As String)

    m_sstatus = sNewValue

End Property

Public Property Get Discipline() As String

    Discipline = m_sdiscipline

End Property

Public Property Let Discipline(ByVal sNewValue As String)

    m_sdiscipline = sNewValue

End Property

Public Property Get Department() As String

    Department = m_sdepartment

End Property

Public Property Let Department(ByVal sNewValue As String)

    m_sdepartment = sNewValue

End Property

Public Property Get Summary() As String

    Summary = m_ssummary

End Property
Public Property Get Contact_Name() As String

    Contact_Name = m_scontact_name

End Property
Public Property Let Contact_Name(ByVal sNewValue As String)

    m_scontact_name = sNewValue

End Property
Public Property Get Contact_Info() As String

    Contact_Info = m_scontact_info
    
End Property
Public Property Let Contact_Info(ByVal sNewValue As String)

    m_scontact_info = sNewValue

End Property

Public Property Let Summary(ByVal sNewValue As String)

    m_ssummary = sNewValue

End Property

Public Property Get Description() As String

    Description = m_sdescription

End Property

Public Property Let Description(ByVal sNewValue As String)

    m_sdescription = sNewValue

End Property

Public Property Get Update() As String

    Update = m_supdate

End Property

Public Property Let Update(ByVal sNewValue As String)

    m_supdate = sNewValue

End Property

Public Property Get analyst() As String

    analyst = m_sanalyst

End Property

Public Property Let analyst(ByVal sNewValue As String)

    m_sanalyst = sNewValue

End Property

Public Property Get Date_Submitted() As String

    Date_Submitted = m_sdate_submitted

End Property

Public Property Let Date_Submitted(ByVal sNewValue As String)

    m_sdate_submitted = sNewValue

End Property

Public Property Get Date_Started() As String

    Date_Started = m_sdate_started

End Property

Public Property Let Date_Started(ByVal sNewValue As String)

    m_sdate_started = sNewValue

End Property

Public Property Get Date_Completed() As String

    Date_Completed = m_sdate_completed

End Property

Public Property Let Date_Completed(ByVal sNewValue As String)

    m_sdate_completed = sNewValue

End Property

Public Property Get Value_To_Business() As String

    Value_To_Business = m_svalue_to_business

End Property

Public Property Let Value_To_Business(ByVal sNewValue As String)

    m_svalue_to_business = sNewValue

End Property


