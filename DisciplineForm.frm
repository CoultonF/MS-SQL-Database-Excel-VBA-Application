VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DisciplineForm 
   Caption         =   "Disciplines"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4995
   OleObjectBlob   =   "DisciplineForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DisciplineForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAction_Click()

Dim Discipline As New Discipline
If txtDISCIPLINE.value <> "" And Trim(txtDISCIPLINE.value) <> "" Then
Discipline.addDiscipline (Trim(txtDISCIPLINE.value))
Unload Me
Unload DisciplineMenu
Call MDisciplineSelect.showDisciplineSelect
Else
    MsgBox "A required field was not entered."
End If
End Sub

Private Sub btnCancel_Click()
    Call cmdReset_Click
    Unload Me
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

Private Sub CommandButton2_Click()

Dim Discipline As New Discipline
If txtDISCIPLINE.value <> "" And Trim(txtDISCIPLINE.value) <> "" Then
Discipline.removeDiscipline (Trim(txtDISCIPLINE.value))
Unload Me
Unload DisciplineMenu
Call MDisciplineSelect.showDisciplineSelect
Else
    MsgBox "A required field was not entered."
End If

End Sub

