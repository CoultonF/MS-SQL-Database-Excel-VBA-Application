VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Define the IProgressBar interface."
' ==========================================================================
' Module      : IProgressBar
' Type        : Class Module
' Description : Define the ProgressBar interface.
' --------------------------------------------------------------------------
' Properties  : Caption             (Get)  String
'               Caption             (Let)  String
'               ChangeRate          (Get)  Double
'               ChangeRate          (Let)  Double
'               Max                 (Get)  Long
'               Max                 (Let)  Long
'               Min                 (Get)  Long
'               Min                 (Let)  Long
'               Percent             (Get)  Double
'               Value               (Get)  Long
'               Value               (Let)  Long
' --------------------------------------------------------------------------
' Procedures  : Complete
'               Decrement
'               Hide
'               Increment
'               Reset
'               Show
' ==========================================================================

Public Property Get Caption() As String
End Property

Public Property Let Caption(ByVal Caption As String)
End Property

Public Property Get ChangeRate() As Double
End Property

Public Property Let ChangeRate(ByVal Rate As Double)
End Property

Public Property Get Max() As Long
End Property

Public Property Let Max(ByVal MaxValue As Long)
End Property

Public Property Get Min() As Long
End Property

Public Property Let Min(ByVal MinValue As Long)
End Property

Public Property Get Percent() As Double
End Property

Public Property Get value() As Long
End Property

Public Property Let value(ByVal value As Long)
End Property

Public Sub Complete()
End Sub

Public Sub Decrement()
End Sub

Public Sub Hide()
End Sub

Public Sub Increment()
End Sub

Public Sub Refresh()
End Sub

Public Sub Reset()
End Sub

Public Sub show()
End Sub
