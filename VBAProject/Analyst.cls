VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Analyst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_sFirst_Name As String
Private m_iAnalyst_ID As Variant
Private m_sLast_Name As Variant
Private m_sUsername As String

Private f As New SysFunc

Public dataObj As New DataAccess


Public Function analystExists(username As String)

    If dataObj.runQuery("SELECT Count(Analyst_ID) FROM [Local_DB].[dbo].[ANALYST] WHERE Username LIKE ?", Array(username))(0, 0) > 0 Then
        
        analystExists = True
        
    End If

End Function

Public Function find(Optional firstName, Optional lastName, Optional username, Optional isAnalyst)

    Dim sql As String: sql = "SELECT " & Me.getDefaultString & " FROM [dbo].[Analyst]"
    Dim args: Set args = CreateObject("System.Collections.ArrayList")
    Dim first As Boolean: first = True
    firstName = Trim(firstName)
    If Not firstName = "" Then
    
        If first Then
            first = False
            sql = sql & " WHERE First_Name LIKE ?"
        Else
            sql = sql & " AND First_Name LIKE ?"
        End If
    
        args.Add firstName
    
    End If
    If Not lastName = "" Then
        lastName = lastName & "%"
        If first Then
            first = False
            sql = sql & " WHERE Last_Name LIKE ?"
        Else
            sql = sql & " AND Last_Name LIKE ?"
        End If
    
        args.Add lastName
    
    End If
    If Not username = "" Then
    
        If first Then
            first = False
            sql = sql & " WHERE Username LIKE ?"
        Else
            sql = sql & " AND Username LIKE ?"
        End If
    
        args.Add username
    
    End If
    Dim arr As Variant: arr = dataObj.runQuery(sql, args.toarray())
    On Error GoTo noData
    find = Array(arr(0, 0), arr(1, 0), arr(2, 0), arr(3, 0))
noData:
End Function

Public Function getAnalystFirstNamesArray()

    getAllUserFirstNamesArray = dataObj.runQuery("SELECT First_Name FROM [Local_DB].[dbo].[ANALYST] WHERE Is_Analyst = True ORDER BY First_Name")
    
End Function
Public Function getDefaultString()

getDefaultString = "First_Name, Last_Name, Is_Analyst, Username"

End Function

Public Function getDefaultArray()

    getDefaultArray = Array("First_Name", "Last_Name", "Is_Analyst", "Username")

End Function

Public Function getAnalystFirstNameFromUsername(username As String)
    
    Dim vari As Variant: vari = dataObj.runQuery("SELECT First_Name FROM [Local_DB].[dbo].[ANALYST] WHERE Username LIKE ?", Array(username))
    
    If isEmpty(vari) Or UBound(vari, 1) < 0 Then
        vari = ""
    Else
        vari = vari(0, 0)
    End If
    first_name = vari
    getAnalystFirstNameFromUsername = vari

End Function
Public Function getAllAnalystsFirstName()

    getAllAnalystsFirstName = dataObj.runQuery("SELECT First_Name FROM ANALYST")

End Function
Public Function isAnalyst(username As String)
    
    Dim analystdata As Variant
    analystdata = dataObj.runQuery("SELECT Is_Analyst FROM [dbo].[ANALYST] WHERE Username LIKE ?", Array(username))
    If UBound(analystdata) < 0 Then
    
    isAnalyst = False
    
    Else
    
    isAnalyst = analystdata(0, 0)
    
    End If

End Function
Public Function getAllCurrentAnalystsFirstName()

    getAllCurrentAnalystsFirstName = dataObj.runQuery("SELECT First_Name FROM ANALYST WHERE Is_Analyst = 'True' ORDER BY First_Name")

End Function

Public Function addNewAnalyst(first_name As String, Last_Name As String, Is_Analyst As String, username As String)

    Dim isAnalyst As Boolean: If Is_Analyst = "Yes" Then isAnalyst = True

    Dim params As Variant: params = UBound(getDefaultArray())
    
    Dim paramString As String: paramString = getDefaultString()

    Call dataObj.runQuery("INSERT INTO [Local_DB].[dbo].[ANALYST] (" & paramString & ") VALUES ( " & f.createParamString(CInt(params)) & " )", Array(first_name, Last_Name, isAnalyst, username))

End Function
Public Function isDBOwner(username As String)

    If dataObj.runQuery("SELECT COUNT(dp.name) " & _
                        "FROM sys.sysusers us right " & _
                        "JOIN  sys.database_role_members rm ON us.uid = rm.member_principal_id " & _
                        "JOIN sys.database_principals dp ON rm.role_principal_id =  dp.principal_id WHERE us.name LIKE ? AND dp.name LIKE 'db_owner'", _
    Array("DOMAIN\" & username))(0, 0) > 0 Then
    
    isDBOwner = True
    
    Else
    'for testing assume everyone is dbowner
    isDBOwner = True
    
    End If


End Function
Public Function editAnalyst(Analyst_ID As Integer, first_name As String, Last_Name As String, Is_Analyst As String, username As String, permission As String)

    Dim isAnalyst As Boolean: If Is_Analyst = "Yes" Then isAnalyst = True

    Call dataObj.runQuery("UPDATE [Local_DB].[dbo].[ANALYST] SET [First_Name] = ? ,[Last_Name] = ? ,[Is_Analyst] = ? ,[Username] = ? WHERE Analyst_ID = ?", Array(first_name, Last_Name, isAnalyst, username, Analyst_ID))

    If userExists(username) And Is_Analyst Like "Yes" Then
    
        grantPermission username, permission
        
    ElseIf Is_Analyst Like "Yes" Then
    
        createUser username
        
        grantPermission username, permission
    
    End If

End Function
Public Function removeAnalyst(username As String)
    
    If Len(username) > 0 Then
    
    Call dataObj.runQuery("DELETE FROM [Local_DB].[dbo].[ANALYST] WHERE USERNAME LIKE ?", Array(username))

    End If

End Function


Public Function createUser(username As String)

    username = createUsername(username)
    
    Call dataObj.runQuery("CREATE USER " & username & " FOR LOGIN " & username)

End Function

Public Function userExists(username)

    If dataObj.runQuery("SELECT count(name) FROM sys.database_principals where (type = 'U') and name = 'DOMAIN\" & username & "'")(0, 0) > 0 Then userExists = True

End Function

Public Function grantPermission(username As String, permission As String)
    username = createUsername(username)
    If permission = "Admin" Then
        Call revokePermission(username, "db_datareader")
        Call revokePermission(username, "db_datawriter")
        Call dataObj.runQuery("ALTER ROLE [db_owner] ADD MEMBER " & username)
    ElseIf permission = "User" Then
        Call revokePermission(username, "db_owner")
        Call dataObj.runQuery("ALTER ROLE [db_datareader] ADD MEMBER " & username)
        Call dataObj.runQuery("ALTER ROLE [db_datawriter] ADD MEMBER " & username)
    End If
End Function
Public Function revokePermission(SQLusername As String, permission As String)

    Call dataObj.runQuery("EXEC sp_droprolemember [" & permission & "], " & SQLusername)

End Function


Private Function createUsername(username)

createUsername = Chr(34) & "DOMAIN\" & Trim(username) & Chr(34)

End Function
Public Function removeUser(username As String)
    username = sanitizeUsernameInput(username)
    username = Chr(34) & "DOMAIN\" & username & Chr(34)
    Call dataObj.runQuery("DROP USER " & username)

End Function
Private Function sanitizeUsernameInput(username As String)

    If username Like WorksheetFunction.Rept("[a-zA-Z0-9.]", Len(username)) Then sanitizeUsernameInput = username

End Function





'GETTERS AND SETTERS FOR THE CLASS DELCLARATIONS

Public Property Get first_name() As String

    first_name = m_sFirst_Name

End Property

Public Property Let first_name(ByVal iNewValue As String)

    m_sFirst_Name = iNewValue

End Property

Public Property Let Analyst_ID(ByVal iNewValue As Variant)

    m_iAnalyst_ID = iNewValue

End Property
Public Property Get Analyst_ID() As Variant

    Analyst_ID = m_iAnalyst_ID

End Property

Public Property Let Last_Name(ByVal iNewValue As Variant)

    m_sLast_Name = iNewValue

End Property
Public Property Get Last_Name() As Variant

    Last_Name = m_sLast_Name

End Property

Public Property Let username(ByVal iNewValue As String)

    m_sUsername = iNewValue

End Property
Public Property Get username() As String

    username = m_sUsername

End Property



