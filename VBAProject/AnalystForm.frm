VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AnalystForm 
   Caption         =   "Analyst Form"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6030
   OleObjectBlob   =   "AnalystForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AnalystForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' -----------------------------------
' Constant declarations
' -----------------------------------
' Global Level
' ----------------

Public action As String

'Public Const GLOBAL_CONST As String = ""

' ----------------
' Module Level
' ----------------

Private Const msMODULE As String = "FCustomer"

' -----------------------------------
' Variable declarations
' -----------------------------------
' Module Level
' ----------------


Private EnableEvents As Boolean

Private ml_Row                  As Long

Dim f As SysFunc

Private WithEvents Worksheet    As Excel.Worksheet
Attribute Worksheet.VB_VarHelpID = -1

Private Sub UpdateControls()
        
    With ActiveSheet.UsedRange.Rows(ActiveCell.row)
        txtUsername = f.getUsername
    End With

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

'SUBMIT BUTTON
Private Sub btnAction_Click()
    If txtUsername.value <> "" And txtFirstName <> "" And txtLastName <> "" And isAnalyst <> "" Then
    Dim dataObj As New DataAccess
    Dim analystObj As New analyst
    
    'If username already exists then run update query
    
    If analystObj.analystExists(txtUsername) Then
        
        Dim analystID As Integer: analystID = dataObj.runQuery("SELECT Analyst_ID FROM [dbo].[ANALYST] WHERE [Username] LIKE ?", Array(txtUsername.value))(0, 0)
        
        Call analystObj.editAnalyst(analystID, txtFirstName.value, txtLastName.value, isAnalyst.value, txtUsername.value, permission.value)
        
        
        'If the analyst is being removed from the team but the historical record is maintained then just revoke access as per below.
        
        If isAnalyst.value = "No" Then analystObj.removeUser (txtUsername.value)
        
    Else
    
    'Else if username does not exist then run insert query
    
        
        'Add the analyst information into the [Analyst] table. This does not mean the analyst will have permission, just that a row will be created in a table holding the data regarding the analyst
    
        Call analystObj.addNewAnalyst(txtFirstName.value, txtLastName.value, isAnalyst.value, txtUsername.value)
        
        
        'ON ERROR - not test instance, user permission must be added through Service Desk.
        
        Err.Clear 'Clear errors to determine the latest error state while adding permission
        
        On Error GoTo permissionDenied
        
        
        'Insert username into the list of permissions
        
        Call analystObj.createUser(txtUsername.value)
        
        Call analystObj.grantPermission(txtUsername.value, permission.value)
        
        
        'ERROR FLAG - msgbox, no permission granted

permissionDenied:

        If Err.Number <> 0 Then
        MsgBox "The permission for the user could not be added. To add permission contact Service Desk with:" & vbNewLine & "Add user " & txtUsername.value & " as db_owner on the the server '" & dataObj.Testing_Conn & "' and database 'Local_DB'"
        Err.Clear
        End If
        
        
    End If
    Unload Me
    Unload AnalystMenu
    
    Else
    MsgBox "A required field was not entered"
    End If
End Sub

'FIND BUTTON
Private Sub CommandButton1_Click()

    Dim dataObj As New DataAccess
    Dim analystObj As New analyst

    'create sql query to find the analyst data based on fields
    Dim analystdata As Variant: analystdata = analystObj.find(txtFirstName.value, txtLastName.value, txtUsername.value, isAnalyst.value)
    
    Dim sql As String
    
    sql = "SELECT TOP 1 FROM "
    clearFields
    On Error GoTo noDataFound
    txtFirstName.value = analystdata(0)
    txtLastName.value = Trim(analystdata(1))
    txtUsername.value = analystdata(3)
    If analystdata(2) Then
    isAnalyst.value = "Yes"
    Else
    isAnalyst.value = "No"
    End If
    If analystObj.isDBOwner(txtUsername.value) Then
    permission.value = "Admin"
    Else
    permission.value = "User"
    End If
noDataFound:
    'populate the data into the remaining fields

End Sub

'REMOVE BUTTON
Private Sub CommandButton2_Click()

    Dim dataObj As New DataAccess

    'Check if the username exists
        
        Dim analystObj As New analyst
        If analystObj.analystExists(txtUsername.value) Then
        
        'Run DB delete query
            
            If analystObj.isAnalyst(txtUsername) Then
                analystObj.removeUser (txtUsername)
            End If
            analystObj.removeAnalyst (txtUsername)
            
        End If
        Unload Me
        
End Sub
Private Function clearFields()

    txtFirstName.value = ""
    txtLastName.value = ""
    txtUsername.value = ""
    isAnalyst.value = ""
    permission.value = ""

End Function
Private Sub userform_initialize()

isAnalyst.AddItem "Yes"

isAnalyst.AddItem "No"

permission.AddItem "Admin"

permission.AddItem "User"



End Sub


