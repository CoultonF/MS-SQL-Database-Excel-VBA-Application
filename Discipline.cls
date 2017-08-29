VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Discipline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public dataObj As New DataAccess
Private f As New SysFunc

Public Function getAllDisciplines()

    Dim disciplines As Variant
    disciplines = dataObj.runQuery("SELECT Discipline FROM [dbo].[DISCIPLINE]", Array())
    If f.IsArrayAllocated(disciplines) Then
    
        getAllDisciplines = disciplines
    
    End If
End Function
Public Function matchDiscipline(Discipline As String)

    Dim disciplines As Variant
    Discipline = "%" & Trim(Discipline) & "%"
    matchDiscipline = dataObj.runQuery("SELECT * FROM [dbo].[DISCIPLINE] WHERE Discipline LIKE ?", Array(Discipline))

End Function

Public Function getDisciplineByID(id As Variant)

Dim vID As Integer: vID = CInt(id)
Dim Data As Variant: Data = dataObj.runQuery("SELECT Discipline FROM [dbo].[Discipline] WHERE Discipline_ID = ?", Array(vID))


If f.IsArrayAllocated(Data) Then
getDisciplineByID = Data(0, 0)
End If

End Function
Public Function addDiscipline(Discipline As String)

    Call dataObj.runQuery("INSERT INTO [dbo].[DISCIPLINE] ([Discipline]) VALUES (?)", Array(Discipline))

End Function
Public Function removeDiscipline(Discipline As String)

    Call dataObj.runQuery("DELETE FROM [dbo].[DISCIPLINE] WHERE Discipline LIKE (?)", Array(Discipline))

End Function
