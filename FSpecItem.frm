VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FSpecItem 
   Caption         =   "Spec"
   ClientHeight    =   9090
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8325
   OleObjectBlob   =   "FSpecItem.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FSpecItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' ==========================================================================
' Module      : FCustomer
' Type        : Form
' Description :
' --------------------------------------------------------------------------
' Properties  : XXX
' --------------------------------------------------------------------------
' Procedures  : XXX
' --------------------------------------------------------------------------
' Events      : XXX
' --------------------------------------------------------------------------
' Dependencies: XXX
' --------------------------------------------------------------------------
' References  : XXX
' --------------------------------------------------------------------------
' Comments    :
' ==========================================================================

' -----------------------------------
' Option statements
' -----------------------------------

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

Private m_clsAnchors As CAnchors

Private Const msMODULE As String = "FCustomer"

' -----------------------------------
' Variable declarations
' -----------------------------------
' Module Level
' ----------------

Private isModeless As Boolean

Private loadnew As Boolean

Private AnalystCount As Long

Private AnalystString As String

Private EnableEvents As Boolean

Private ml_Row                  As Long
Dim f As New SysFunc

Private WithEvents Worksheet    As Excel.Worksheet
Attribute Worksheet.VB_VarHelpID = -1



Private Sub UpdateControls()
    Dim i As Long
    Dim specObj As New spec
    Dim specColumns As Variant: specColumns = specObj.getDefaultArray
    With ActiveSheet.UsedRange.Rows(ActiveCell.row)
    
        'Updates the Spec Form controls to their corresponding values
        For i = 0 To UBound(specColumns, 1)
            
            CallByName Me, "txt" & specColumns(i), VbLet, .Cells(WorksheetFunction.match(specColumns(i), ActiveSheet.Range("1:1"), 0))
        
        Next i
    
    End With

End Sub

Private Function addControls()

        If loadnew Then
        loadnew = False
        txtSPEC_ID = "Auto"
        txtSPEC_ID.Enabled = False
        txtRANK = ""
        txtANALYST = ""
        txtSTATUS = ""
        txtDISCIPLINE = ""
        txtDEPARTMENT = ""
        txtSUMMARY = ""
        txtDESCRIPTION = ""
        txtDATE_SUBMITTED = ""
        txtDATE_STARTED = ""
        txtDATE_COMPLETED = ""
        txtVALUE_TO_BUSINESS = ""
        
        End If

End Function

Private Function populateStatus()

txtSTATUS.AddItem "Assigned"
txtSTATUS.AddItem "Unassigned"
txtSTATUS.AddItem "Completed"
txtSTATUS.AddItem "Cerner Fix"
txtSTATUS.AddItem "Hold"
txtSTATUS.AddItem "Canceled"

End Function

Private Function populateValue()

txtVALUE_TO_BUSINESS.AddItem "Accreditation Requirement"

txtVALUE_TO_BUSINESS.AddItem "Cost Savings"

txtVALUE_TO_BUSINESS.AddItem "Instrument Interface"

txtVALUE_TO_BUSINESS.AddItem "New Test Added"

txtVALUE_TO_BUSINESS.AddItem "Other (Specify in description)"

txtVALUE_TO_BUSINESS.AddItem "Process Improvement"

txtVALUE_TO_BUSINESS.AddItem "Provincial Standardization"

txtVALUE_TO_BUSINESS.AddItem "Quality Improvement"

txtVALUE_TO_BUSINESS.AddItem "Revenue Generation"

'add more value to business options here
End Function

Private Function populateDepartment()

Dim dataObj As New DataAccess

dataObj.init

Dim results As Variant: results = dataObj.runQuery("SELECT DISTINCT DEPARTMENT FROM Local_DB.dbo.SPEC")

Dim i As Integer

For i = 1 To UBound(results, 2)

    txtDEPARTMENT.AddItem results(0, i)

Next i

End Function



Private Sub AnalystMenuBtn_Click()
    AnalystMenu.populateCurrentAnalysts (txtANALYST.value)
    txtANALYST = MAnalystSelect.showAnalystSelect
End Sub

Private Sub btnDateCompleted_Click()
    txtDATE_COMPLETED = MDateSelect.ShowDatePickerUserForm
End Sub

Private Sub btnDateStarted_Click()
    txtDATE_STARTED = MDateSelect.ShowDatePickerUserForm
End Sub

Private Sub btnDateSubmitted_Click()
    txtDATE_SUBMITTED = MDateSelect.ShowDatePickerUserForm
End Sub

Private Sub lblFax_Click()

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

Private Sub CommandButton1_Click()
Me.Hide

If isModeless Then
isModeless = False
Me.show


Else
isModeless = True
Me.show False


End If



End Sub


Public Function setAnchors()

    Set m_clsAnchors = New CAnchors
    
    Set m_clsAnchors.Parent = Me
    
    ' restrict minimum size of userform
    m_clsAnchors.MinimumWidth = 420.75
    m_clsAnchors.MinimumHeight = 475.5
    
    With m_clsAnchors
    
        'Anchor rules
        .Anchor("CommandButton1").AnchorStyle = enumAnchorStyleBottom Or enumAnchorStyleRight
        .Anchor("btnAction").AnchorStyle = enumAnchorStyleRight
        .Anchor("btnCancel").AnchorStyle = enumAnchorStyleRight
        .Anchor("fraspec").AnchorStyle = enumAnchorStyleLeft Or enumAnchorStyleRight Or enumAnchorStyletop Or enumAnchorStyleBottom
        .Anchor("txtSummary").AnchorStyle = enumAnchorStyleLeft Or enumAnchorStyleRight
        .Anchor("txtDescription").AnchorStyle = enumAnchorStyleLeft Or enumAnchorStyleRight Or enumAnchorStyleBottom Or enumAnchorStyletop
        .Anchor("lblPhone").AnchorStyle = enumAnchorStyleBottom
        .Anchor("txtAnalyst").AnchorStyle = enumAnchorStyleBottom
        .Anchor("AnalystMenuBtn").AnchorStyle = enumAnchorStyleBottom
        .Anchor("Label5").AnchorStyle = enumAnchorStyleBottom
        .Anchor("Label1").AnchorStyle = enumAnchorStyleBottom
        .Anchor("txtDATE_SUBMITTED").AnchorStyle = enumAnchorStyleBottom
        .Anchor("btnDateSubmitted").AnchorStyle = enumAnchorStyleBottom
        .Anchor("Label10").AnchorStyle = enumAnchorStyleBottom
        .Anchor("Label2").AnchorStyle = enumAnchorStyleBottom
        .Anchor("txtDATE_STARTED").AnchorStyle = enumAnchorStyleBottom
        .Anchor("btnDateStarted").AnchorStyle = enumAnchorStyleBottom
        .Anchor("Label3").AnchorStyle = enumAnchorStyleBottom
        .Anchor("txtDATE_COMPLETED").AnchorStyle = enumAnchorStyleBottom
        .Anchor("btnDateCompleted").AnchorStyle = enumAnchorStyleBottom
        .Anchor("Label4").AnchorStyle = enumAnchorStyleBottom
        .Anchor("txtVALUE_TO_BUSINESS").AnchorStyle = enumAnchorStyleBottom
        .Anchor("Label9").AnchorStyle = enumAnchorStyleBottom
        .Anchor("RequiredStar").AnchorStyle = enumAnchorStyleBottom
        .Anchor("fraContact").AnchorStyle = enumAnchorStyleBottom Or enumAnchorStyleLeft Or enumAnchorStyleRight
    
    End With

End Function


Private Sub CommandButton2_Click()
    DisciplineMenu.ListBox2.Clear
    DisciplineMenu.populateCurrentDisciplines (txtDISCIPLINE.value)
    txtDISCIPLINE = MDisciplineSelect.showDisciplineSelect
    
End Sub



Private Sub txtStatus_Change()

    If txtSTATUS.text = "Completed" Then
        Label3.Caption = "Date Completed:"
        RequiredStar.Visible = True
    Else
        Label3.Caption = "Est. Completion:"
        RequiredStar.Visible = False
    End If
    

End Sub

Private Sub UserForm_Activate()

    

    populateStatus
    
    populateValue
    
    populateDepartment

    FSpecItem.txtSPEC_ID.Enabled = False
    
    FSpecItem.Caption = action & " Spec"
    
    FSpecItem.btnAction.Caption = action
    
    If action = "Add" Then
    
        Call addControls
    Else
        txtSPEC_ID.Enabled = True
    End If
    
    If action = "Delete" Then
    
        txtSPEC_ID.Enabled = False
        
    Else
    
        txtSPEC_ID.Enabled = False
        
    End If
End Sub


Private Sub userform_initialize()
    
    Me.setAnchors
    
    EnableEvents = True
    
    loadnew = True
    
    Set Worksheet = Sheets(1)
    Worksheet.activate
    With Worksheet.UsedRange
        If (.Rows.count > 1) Then
            If ((ActiveCell.row <= .Rows.count) _
            And (ActiveCell.row > 1)) Then
                ml_Row = ActiveCell.row
                .Rows(ml_Row).Select
            Else
                ml_Row = 2
                Worksheet.UsedRange.Rows(ml_Row).Select
            End If
            Call UpdateControls
        Else
            ml_Row = 2
            .Rows(1).Offset(1).Select
        End If
    End With
    


End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Set m_clsAnchors = Nothing
    Call cmdReset_Click
    Unload Me
End Sub
Private Sub UserForm_Terminate()

    Set m_clsAnchors = Nothing
    Call cmdReset_Click
    Unload Me
End Sub


'THE DONE BUTTON ACTION TO SEND THE RESULTING ADD REMOVE OR EDIT OF A SPEC ITEM
Private Sub btnAction_Click()
    loadnew = True
    If RequiredStar.Visible = True And txtDATE_COMPLETED.value = "" Then
        GoTo wrongField
    End If
    If txtSUMMARY = "" Or txtDATE_SUBMITTED = "" Or txtSTATUS = "" Then
        GoTo wrongField
    End If

    Dim specObj As New spec
            
    Dim listObj As New SpecList
    
    Set listObj = SpecListController.listObj
    
    specObj.init
            
    Dim values() As String
    
    If txtRANK.text = "0" Then
        txtRANK.text = ""
    End If
    
    Dim specCol As Variant
    
    Dim i As Long
    
    For Each specCol In specObj.getDefaultArray
    
        ReDim Preserve values(i)
    
        values(i) = CallByName(Me, "txt" & specCol, VbGet)
    
        i = i + 1
        
    Next specCol
    'values = Array(txtSPEC_ID.text, txtRANK.text, txtSTATUS.text, txtDiscipline.text, txtDEPARTMENT.text, txtSUMMARY.text, txtDESCRIPTION.text, txtANALYST.text, txtDATE_SUBMITTED.text, txtDATE_STARTED.text, txtDATE_COMPLETED.text, txtVALUE_TO_BUSINESS.text, txtCONTACT_NAME.value, txtCONTACT_INFO.value)
                    
    Call specObj.setDictionaryToSpec(specObj.convertToSpecDict(values, specObj.getDefaultArray()))

    Select Case action
    
        Case "Add"
            
            Call listObj.addToList(specObj)
        
        Case "Edit"
        
            Call listObj.updateFromList(specObj)
            SpecListController.editedSpecID = CStr(specObj.spec_id)
        
        Case "Delete"
            
            Dim mbResult As Integer
            mbResult = MsgBox("Once a SPEC item is deleted the associated updates will be deleted as well. Are you sure you want to delete this SPEC?", vbYesNoCancel + vbDefaultButton2)
            If mbResult = vbYes Then
            Call listObj.removeFromList(specObj)
            End If
    End Select
    Unload Me
    Call cmdReset_Click

    If Err Then
        Debug.Print Err.Description
    End If
    
    'Call SpecListController.Build
    Call SpecListController.printList(listObj)
    Exit Sub
wrongField:
    MsgBox "A mandatory field was not entered"
    
End Sub

Private Sub btnCancel_Click()
    Set m_clsAnchors = Nothing
    Call cmdReset_Click
    Unload Me
End Sub


