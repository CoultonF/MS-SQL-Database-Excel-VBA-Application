VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==========================================================================
' Module      : CProgressBar
' Type        : Class Module
' Description : StatusBar implementation of IProgressBar
'
' Notes       : Many of these properties or procedures only exist to
'               fulfill the requirements of the implementation contract.
'               For example, since this is a status bar implementation,
'               there would never be an update to a 'Title'.
' --------------------------------------------------------------------------
' Properties  : IProgressBar_Caption            (Get)  String
'               IProgressBar_Caption            (Let)  String
'               IProgressBar_ChangeRate         (Get)  Double
'               IProgressBar_ChangeRate         (Let)  Double
'               IProgressBar_Max                (Get)  Long
'               IProgressBar_Max                (Let)  Long
'               IProgressBar_Min                (Get)  Long
'               IProgressBar_Min                (Let)  Long
'               IProgressBar_Percent            (Get)  Double
'               IProgressBar_Value              (Get)  Long
'               IProgressBar_Value              (Let)  Long
' --------------------------------------------------------------------------
' Procedures  : IProgressBar_Complete
'               IProgressBar_Decrement
'               IProgressBar_Hide
'               IProgressBar_Increment
'               IProgressBar_Refresh
'               IProgressBar_Reset
'               IProgressBar_Show
'               GetPercentage                          Double
' --------------------------------------------------------------------------
' Events      : OnValueChanged
'               Class_Initialize
'               Class_Terminate
' ==========================================================================

' -----------------------------------
' Option statements
' -----------------------------------

Option Explicit

' -----------------------------------
' Interface declarations
' -----------------------------------

Implements IProgressBar

' -----------------------------------
' Event declarations
' -----------------------------------

Public Event OnValueChanged(ByVal value As Long)

' -----------------------------------
' Constant declarations
' -----------------------------------
' Module Level
' ----------------

Private Const msMODULE          As String = "CProgressBar"

Private Const msDEFAULT_CAP     As String = vbNullString
Private Const mlDEFAULT_MIN     As Long = 0
Private Const mlDEFAULT_MAX     As Long = 100
Private Const mlDEFAULT_VAL     As Long = -1
Private Const mdblDEFAULT_CHG   As Double = 0.05

' -----------------------------------
' Variable declarations
' -----------------------------------
' Module Level
' ----------------

Private mbDisplayStatusBar      As Boolean
Private mdblLastPct             As Double

' IProgressBar
Private ms_Caption              As String
Private mb_Visible              As Boolean
Private ml_Min                  As Long
Private ml_Max                  As Long
Private ml_Value                As Long
Private mdbl_ChangeRate         As Double
Private mdbl_Percent            As Double

Private Property Get IProgressBar_Caption() As String
' ==========================================================================

    IProgressBar_Caption = ms_Caption

End Property

Private Property Let IProgressBar_Caption(ByVal RHS As String)
' ==========================================================================
' Description : This is the base caption string.
'               The percentage is calculated during the
'               progress change and appended automatically.
' ==========================================================================

    ms_Caption = RHS

    IProgressBar_Refresh

End Property

Private Property Get IProgressBar_ChangeRate() As Double
' ==========================================================================

    IProgressBar_ChangeRate = mdbl_ChangeRate

End Property

Private Property Let IProgressBar_ChangeRate(ByVal RHS As Double)
' ==========================================================================

    mdbl_ChangeRate = RHS

End Property

Private Property Get IProgressBar_Max() As Long
' ==========================================================================

    IProgressBar_Max = ml_Max

End Property

Private Property Let IProgressBar_Max(ByVal RHS As Long)
' ==========================================================================

    ml_Max = RHS

End Property

Private Property Get IProgressBar_Min() As Long
' ==========================================================================

    IProgressBar_Min = ml_Min

End Property

Private Property Let IProgressBar_Min(ByVal RHS As Long)
' ==========================================================================

    ml_Min = RHS

End Property

Private Property Get IProgressBar_Percent() As Double
' ==========================================================================

    IProgressBar_Percent = mdbl_Percent

End Property

Private Property Get IProgressBar_Value() As Long
' ==========================================================================

    IProgressBar_Value = ml_Value

End Property

Private Property Let IProgressBar_Value(ByVal RHS As Long)
' ==========================================================================

    If (RHS <> ml_Value) Then

        If (RHS < ml_Min) Then
            ml_Value = ml_Min
        ElseIf (RHS > ml_Max) Then
            ml_Value = ml_Max
        Else
            ml_Value = RHS
        End If

        RaiseEvent Me.OnValueChanged(ml_Value)

    End If

End Property

Private Sub IProgressBar_Complete()
' ==========================================================================

    IProgressBar_Value = IProgressBar_Max

End Sub

Private Sub IProgressBar_Decrement()
' ==========================================================================

    If (IProgressBar_Value > IProgressBar_Min) Then
        IProgressBar_Value = IProgressBar_Value - 1
    End If

End Sub

Private Sub IProgressBar_Hide()
' ==========================================================================

    mb_Visible = False
    Application.StatusBar = False
    Application.DisplayStatusBar = mbDisplayStatusBar

End Sub

Private Sub IProgressBar_Increment()
' ==========================================================================

    If (IProgressBar_Value < IProgressBar_Max) Then
        IProgressBar_Value = IProgressBar_Value + 1
    End If

End Sub

Private Sub IProgressBar_Refresh()
' ==========================================================================
' Description : Refresh the display
' ==========================================================================

    Dim sCap    As String

    ' Build the caption
    ' -----------------
    If (Len(ms_Caption) > 0) Then
        sCap = ms_Caption & " (" & Format(mdbl_Percent, "0%") & ")..."
    End If

    ' Refresh the display
    ' -------------------
    Application.StatusBar = sCap

End Sub

Private Sub IProgressBar_Reset()
' ==========================================================================

    IProgressBar_ChangeRate = mdblDEFAULT_CHG
    IProgressBar_Caption = msDEFAULT_CAP

    IProgressBar_Min = mlDEFAULT_MIN
    IProgressBar_Max = mlDEFAULT_MAX
    IProgressBar_Value = mlDEFAULT_VAL

End Sub

Private Sub IProgressBar_Show()
' ==========================================================================

    mb_Visible = True
    mbDisplayStatusBar = Application.DisplayStatusBar
    Application.DisplayStatusBar = True

End Sub

Private Function GetPercentage(ByVal Min As Long, _
                               ByVal Max As Long, _
                               ByVal Progress As Long) As Double
' ==========================================================================
' Description : Calculate the percentage for a progress bar
'
' Params      : Min         The minimum value for the bar
'               Max         The maximum value for the bar
'               Progress    The current value of the bar
'
' Returns     : Double      The current progress percentage
' ==========================================================================

    Dim dblRtn  As Double

    ' ----------------------------------------------------------------------
    ' Calculate the progress percentage
    ' ---------------------------------
    If (Max <= Min) Then
        dblRtn = 0
    Else
        dblRtn = Abs((Progress - Min) / (Max - Min))
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    GetPercentage = dblRtn

End Function

Public Sub OnValueChanged(ByVal value As Long)
' ==========================================================================
' Description : Recalculate percentage based on the new value
' ==========================================================================

    Dim bRefresh As Boolean

    ' ----------------------------------------------------------------------
    ' Get the current percentage
    ' --------------------------
    mdbl_Percent = GetPercentage(ml_Min, ml_Max, value)

    ' Refresh if needed
    ' -----------------
    bRefresh = (Abs(mdbl_Percent - mdblLastPct) > mdbl_ChangeRate)

    If bRefresh Then
        mdblLastPct = mdbl_Percent
        If mb_Visible Then
            IProgressBar_Refresh

            ' Process events at each update
            ' -----------------------------
            DoEvents

        End If
    End If

End Sub

Private Sub Class_Initialize()
' ==========================================================================
' Description : Set initial values
' ==========================================================================

    Call IProgressBar_Reset

End Sub

Private Sub Class_Terminate()
' ==========================================================================

    Application.DisplayStatusBar = mbDisplayStatusBar

End Sub
