VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Runtime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_iRuntime_ID As Integer
Private m_vRefresh_Load_Time As Variant
Private m_vList_Load_Time As Variant
Private m_sUsername As String

Public dataObj As New DataAccess


Public Function AddRuntimeUpdate(time As Variant)

    'dataObj.runQuery("INSERT INTO RUNTIME ([Refresh_Load_Time],[Username]) VALUES (?, ?)",Array(CDec(time),

End Function









'GETTERS AND SETTERS FOR THE CLASS DELCLARATIONS

Public Property Get Runtime_ID() As Integer

    Runtime_ID = m_iRuntime_ID

End Property

Public Property Let Runtime_ID(ByVal iNewValue As Integer)

    m_iRuntime_ID = iNewValue

End Property

Public Property Let Refresh_Load_Time(ByVal iNewValue As Variant)

    m_vRefresh_Load_Time = iNewValue

End Property
Public Property Get Refresh_Load_Time() As Variant

    Refresh_Load_Time = m_vRefresh_Load_Time

End Property

Public Property Let List_Load_Time(ByVal iNewValue As Variant)

    m_vList_Load_Time = iNewValue

End Property
Public Property Get List_Load_Time() As Variant

    List_Load_Time = m_vList_Load_Time

End Property

Public Property Let username(ByVal iNewValue As Variant)

    m_sUsername = iNewValue

End Property
Public Property Get username() As Variant

    username = m_sUsername

End Property


