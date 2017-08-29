VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AnalystMenu 
   Caption         =   "Add / Remove Analysts"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5415
   OleObjectBlob   =   "AnalystMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AnalystMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim i As Integer
Dim originalAnalystList As String


Private Sub CommandButton3_Click()
    Me.Hide
End Sub

Private Sub CommandButton4_Click()
    Call populateCurrentAnalysts(originalAnalystList)
    Me.Hide
End Sub

Private Sub CommandButton5_Click()
    Dim analystChange As Boolean
    analystChange = MAnalystModify.showAnalystModify
End Sub

Private Sub userform_initialize()

Call populateAnalyst

OptionButton3.value = True

End Sub
Private Sub cmdReset_Click()

    Dim ctl As MSForms.Control

    For Each ctl In Me.Controls
        Select Case TypeName(ctl)
            Case "TextBox"
                ctl.text = ""
            Case "CheckBox", "OptionButton", "ToggleButton"
                ctl.value = False
            Case "ComboBox", "ListBox"
                ctl.ListIndex = -1
        End Select
    Next ctl

End Sub
Private Function populateAnalyst()

Dim analystObj As New analyst
Dim results As Variant: results = analystObj.getAllCurrentAnalystsFirstName()
Dim i As Integer
For i = 0 To UBound(results, 2)

    ListBox1.AddItem results(0, i)

Next i
Dim f As New SysFunc
'txtAnalyst = analystObj.getAnalystFirstNameFromUsername(f.getUsername)
End Function
Public Function populateCurrentAnalysts(analysts As String)
    
    ListBox2.Clear
    Dim analyst() As String
    analyst = Split(Trim(analysts), " ")
    
    Dim i As Integer
    For i = 0 To UBound(analyst)
    
        ListBox2.AddItem (analyst(i))
    
    Next i
    originalAnalystList = analysts
    
    SortListBox

End Function

Private Sub CommandButton1_Click()

For i = 0 To ListBox1.ListCount - 1
    If ListBox1.Selected(i) = True Then ListBox2.AddItem ListBox1.list(i)
Next i
SortListBox
End Sub

Private Sub SortListBox()
    'Sorts ListBox List
    Dim i As Long
    Dim j As Long
    Dim temp As Variant
        
    With Me.ListBox2
        For j = 0 To ListBox2.ListCount - 2
            For i = 0 To ListBox2.ListCount - 2
                If .list(i) > .list(i + 1) Then
                    temp = .list(i)
                    .list(i) = .list(i + 1)
                    .list(i + 1) = temp
                End If
            Next i
        Next j
    End With
End Sub

Private Sub CommandButton2_Click()

Dim counter As Integer
counter = 0

For i = 0 To ListBox2.ListCount - 1
    If ListBox2.Selected(i - counter) Then
        ListBox2.RemoveItem (i - counter)
        counter = counter + 1
    End If
Next i

CheckBox2.value = False

End Sub

Private Sub OptionButton1_Click()

ListBox1.MultiSelect = 0
ListBox2.MultiSelect = 0

End Sub

Private Sub OptionButton2_Click()

ListBox1.MultiSelect = 1
ListBox2.MultiSelect = 1

End Sub

Private Sub OptionButton3_Click()

ListBox1.MultiSelect = 2
ListBox2.MultiSelect = 2

End Sub

Private Sub CheckBox1_Click()

If CheckBox1.value = True Then
    For i = 0 To ListBox1.ListCount - 1
        ListBox1.Selected(i) = True
    Next i
End If

If CheckBox1.value = False Then
    For i = 0 To ListBox1.ListCount - 1
        ListBox1.Selected(i) = False
    Next i
End If

End Sub

Private Sub CheckBox2_Click()

If CheckBox2.value = True Then
    For i = 0 To ListBox2.ListCount - 1
        ListBox2.Selected(i) = True
    Next i
End If

If CheckBox2.value = False Then
    For i = 0 To ListBox2.ListCount - 1
        ListBox2.Selected(i) = False
    Next i
End If

End Sub

