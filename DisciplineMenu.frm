VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DisciplineMenu 
   Caption         =   "Discipline Menu"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5370
   OleObjectBlob   =   "DisciplineMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DisciplineMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim i As Integer
Dim originalDisciplineList As String

Private Sub CommandButton3_Click()
    'MDisciplineSelect.disciplineArray = ListBox2.list
        addListToDisciplineArray
    Me.Hide
    Unload Me
End Sub
Private Function addListToDisciplineArray()

    MDisciplineSelect.disciplineArray.Clear
    For i = 0 To ListBox2.ListCount - 1
        MDisciplineSelect.disciplineArray.Add ListBox2.list(i)
    Next i

End Function

Private Sub CommandButton4_Click()
    Call populateCurrentDisciplines(originalDisciplineList)
    addListToDisciplineArray
    Me.Hide
    Unload Me
End Sub
Private Sub UserForm_Terminate()

    Call populateCurrentDisciplines(originalDisciplineList)
    'MDisciplineSelect.disciplineArray = ListBox2.list
End Sub

Private Sub CommandButton5_Click()
    Dim disciplineChange As Boolean
    disciplineChange = MDisciplineModify.showDisciplineModify
End Sub

Private Sub userform_initialize()

Set MDisciplineSelect.disciplineArray = CreateObject("System.Collections.ArrayList")

Call populatedisciplines

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
Private Function populatedisciplines()

Dim disciplineObj As New Discipline
Dim results As Variant: results = disciplineObj.getAllDisciplines()
Dim i As Integer
For i = 0 To UBound(results, 2)

    ListBox1.AddItem results(0, i)

Next i
Dim f As New SysFunc
'txtAnalyst = analystObj.getAnalystFirstNameFromUsername(f.getUsername)
End Function
Public Function populateCurrentDisciplines(disciplines As String)

    ListBox2.Clear
    Dim Discipline() As String
    Discipline = Split(Trim(disciplines), vbNewLine)
    Dim disciplineIDs: Set disciplineIDs = CreateObject("System.Collections.ArrayList")
    Dim i As Integer
    originalDisciplineList = disciplines
    On Error GoTo noArr2
    For i = 0 To UBound(Discipline)
        
        Dim disciplineObj As New Discipline
        Dim disciplineData As Variant
        disciplineData = disciplineObj.matchDiscipline(Discipline(i))
        disciplineIDs.Add (disciplineData(0, 0))
        
        'MDisciplineSelect.disciplineArray.Add disciplineData(1, 0)
        'ListBox2.AddItem (discipline(i))
    
    Next i
    Dim arrID As Variant
    arrID = disciplineIDs.toarray()
    Dim f As New SysFunc
    arrID = f.eliminateDuplicate(arrID)
    For i = 0 To UBound(arrID)
        Dim disciplineName As Variant: disciplineName = disciplineObj.getDisciplineByID(arrID(i))
        ListBox2.AddItem (disciplineName)
        Dim r As Integer: r = 0
        While r < ListBox1.ListCount
            If ListBox1.list(r) = disciplineName Then
                ListBox1.RemoveItem (r)
                r = ListBox1.ListCount
            End If
            r = r + 1
        Wend
    
    Next i
noArr2:
    SortListBox

End Function
Function collectionToArray(c As Collection) As Variant()
    Dim a() As Variant: ReDim a(0 To c.count - 1)
    Dim i As Integer
    For i = 1 To c.count
        a(i - 1) = c.item(i)
    Next
    collectionToArray = a
End Function
Private Function cmdReset()

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

End Function
Private Sub CommandButton1_Click()
Dim endCounter As Integer: endCounter = ListBox1.ListCount - 1
Dim i As Integer: i = 0
On Error GoTo endoflistitems
While i <= endCounter
    If ListBox1.Selected(i) = True Then
        ListBox2.AddItem ListBox1.list(i)
        MDisciplineSelect.disciplineArray.Add ListBox1.list(i)
        ListBox1.RemoveItem (i)
        i = i - 1
    End If
    i = i + 1
Wend
endoflistitems:
Err.Clear
SortListBox
End Sub

Public Sub SortListBox()
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
    With Me.ListBox1
        For j = 0 To ListBox1.ListCount - 2
            For i = 0 To ListBox1.ListCount - 2
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
        
        ListBox1.AddItem (ListBox2.list(i - counter))
        MDisciplineSelect.disciplineArray.Remove ListBox2.list(i - counter)
        ListBox2.RemoveItem (i - counter)
        counter = counter + 1
        
    End If
Next i

CheckBox2.value = False
SortListBox
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


