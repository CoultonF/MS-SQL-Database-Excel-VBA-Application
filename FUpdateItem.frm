VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FUpdateItem 
   Caption         =   "Spec Update"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8085
   OleObjectBlob   =   "FUpdateItem.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FUpdateItem"
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

Private ml_Row                  As Long

Private f As New SysFunc

Private WithEvents Worksheet    As Excel.Worksheet
Attribute Worksheet.VB_VarHelpID = -1


Private Sub UpdateControls()
        
    If UIController.isSpecListView Then
    'get the latest update as the Spec view is being displayed
    
        Dim updateListObj As New UpdateList
        
        Dim updateItem As Variant: updateItem = updateListObj.getLatestUpdate(UpdateListController.specid)
        
        Dim f As New SysFunc
        
        Dim updateObj As New Update
        
        If f.IsArrayAllocated(updateItem) Then
        Dim col As Variant
        updateItem = f.Flatten(updateItem)
        For Each col In updateObj.getDefaultOrderArray
            Dim i As Long
            Call CallByName(Me, "txt" & col, VbLet, CStr(updateItem(i)))
            i = i + 1
        Next col

        End If
    ElseIf UIController.isUpdateListView Then
    'Update view is being displayed
        
        With ActiveSheet.UsedRange.Rows(ActiveCell.row)
        
            txtSPEC_ID = UpdateListController.specid
            
            Dim value As Variant
            
            For Each value In updateObj.getDefaultOrderArray()
            
                CallByName Me, "txt" & value, VbLet, .Cells(f.getHeaderColumnIndexOf(CStr(value)))
                
            Next value
    
        End With

    
    End If
        
    
End Sub

Private Function addControls()

        Dim analystObj As New analyst

        txtUpdate_ID = "Auto"
        txtupdate_Desc = ""
        txtUpdate_Date = ""
        txtUpdate_Analyst = analystObj.getAnalystFirstNameFromUsername(f.getUsername)
        txtSPEC_ID = UpdateListController.specid

End Function


Public Function setAnchors()

    Set m_clsAnchors = New CAnchors
    
    Set m_clsAnchors.Parent = Me
    
    ' restrict minimum size of userform
    m_clsAnchors.MinimumWidth = 408.75
    m_clsAnchors.MinimumHeight = 216
    
    With m_clsAnchors
    
        .Anchor("btnAction").AnchorStyle = enumAnchorStyleRight
        .Anchor("btnCancel").AnchorStyle = enumAnchorStyleRight
        .Anchor("txtUpdate_Desc").AnchorStyle = enumAnchorStyleRight Or enumAnchorStyleBottom Or enumAnchorStyleLeft Or enumAnchorStyletop
        .Anchor("fraspec").AnchorStyle = enumAnchorStyleLeft Or enumAnchorStyleRight Or enumAnchorStyletop Or enumAnchorStyleBottom
        .Anchor("lblPhone").AnchorStyle = enumAnchorStyleBottom
        .Anchor("AnalystMenuBtn").AnchorStyle = enumAnchorStyleBottom
        .Anchor("Label10").AnchorStyle = enumAnchorStyleBottom
        .Anchor("Label1").AnchorStyle = enumAnchorStyleBottom
        .Anchor("Label9").AnchorStyle = enumAnchorStyleBottom
        .Anchor("txtUpdate_Date").AnchorStyle = enumAnchorStyleBottom
        .Anchor("btnDateSubmitted").AnchorStyle = enumAnchorStyleBottom
        .Anchor("txtUpdate_Analyst").AnchorStyle = enumAnchorStyleBottom

    End With

End Function



Private Sub btnDateSubmitted_Click()
    txtUpdate_Date = MDateSelect.ShowDatePickerUserForm
End Sub
Private Sub AnalystMenuBtn_Click()
    AnalystMenu.populateCurrentAnalysts (txtUpdate_Analyst.value)
    txtUpdate_Analyst = MAnalystSelect.showAnalystSelect
End Sub

Private Sub UserForm_Activate()

    FUpdateItem.txtSPEC_ID.Enabled = False
    
    FUpdateItem.Caption = action & " Update"
    
    FUpdateItem.btnAction.Caption = action
    
    If action = "Add" Then
    
        Call addControls
        
        FUpdateItem.btnAction.Caption = "Add Update"
    
    ElseIf action = "Edit" Then
    
        FUpdateItem.btnAction.Caption = "Edit Update"
    
    ElseIf action = "Delete" Then
    
        FUpdateItem.btnAction.Caption = "Delete Update"
    
    End If

End Sub


Private Sub userform_initialize()
    
    Me.setAnchors

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
    Unload Me
End Sub
Private Sub UserForm_Terminate()

    Set m_clsAnchors = Nothing
    Unload Me
End Sub


'THE DONE BUTTON ACTION TO SEND THE RESULTING ADD REMOVE OR EDIT OF A SPEC ITEM
Private Sub btnAction_Click()

    If txtUpdate_Date = "" Or txtSPEC_ID = "" Or txtupdate_Desc = "" Then
        GoTo emptyField
    End If

    Dim updateObj As New Update
            
    Dim listObj As New UpdateList
            
    Dim values As Variant
            
    values = Array(txtUpdate_ID.text, txtupdate_Desc.text, txtUpdate_Date.text, txtUpdate_Analyst.text, txtSPEC_ID.text)
                    
    Call updateObj.setVariantToUpdate(values)

    Select Case action
    
        Case "Add"
            
            Call listObj.addToList(updateObj)
        
        Case "Edit"
        
            Call listObj.updateFromList(updateObj)
        
        Case "Delete"
    
            Call listObj.removeFromList(updateObj)
    
    End Select
    
    MsgBox "Done."
    
    Unload Me
    
    Call UpdateListController.printList
    Call SpecListController.printList
    
    Exit Sub
    
emptyField:
    MsgBox "A required field was left empty."
End Sub

Private Sub btnCancel_Click()
    Set m_clsAnchors = Nothing
    Unload Me
End Sub




