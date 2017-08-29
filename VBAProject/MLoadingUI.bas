Attribute VB_Name = "MLoadingUI"
' ==========================================================================
' Module      : MLoadingUI
' Type        : Module
' Description : Test the IProgressBar interface
' --------------------------------------------------------------------------
' Procedures  : TestProgressBar
' ==========================================================================

' -----------------------------------
' Option statements
' -----------------------------------

Option Explicit
Option Private Module

Private f As New SysFunc

Public Sub LoadSpecsProgressBar(list As Collection, size As Long)
' ==========================================================================
' Description : Spec Progress Bar Loading interface
' ==========================================================================

    Const lPB_MIN   As Long = 1
    Dim lPB_MAX   As Long: lPB_MAX = size
    Dim lIdx        As Long
    Dim oPB         As IProgressBar

    ' ----------------------------------------------------------------------
    ' Instantiate the object
    ' ----------------------
    Set oPB = New SpecProgressBar

    ' Set the operational parameters
    ' ------------------------------
    oPB.Caption = "Loading"
    oPB.Min = lPB_MIN
    oPB.Max = lPB_MAX
    oPB.show

    ' Increment the ProgressBar
    ' ----------------
    Dim t As Single
    t = Timer
    
    Dim specObj As New spec
    Dim specColumns As Variant: specColumns = specObj.getDefaultOrderArray
    
    'Set the cells to their respective values
    For lIdx = lPB_MIN To lPB_MAX
    
        Dim i As Long: i = 0
        
        For i = 0 To UBound(specColumns, 1)
    
            Cells(lIdx + 1, i + 1) = CallByName(list.item(CInt(lIdx)), specColumns(i), VbGet)
        
        Next i
    
        'Cells(lIdx + 1, 1) = list.item(CInt(lIdx)).spec_id
        'Cells(lIdx + 1, 2) = list.item(CInt(lIdx)).Rank
        'Cells(lIdx + 1, 3) = list.item(CInt(lIdx)).status
        'Cells(lIdx + 1, 4) = list.item(CInt(lIdx)).Discipline
        'Cells(lIdx + 1, 5) = list.item(CInt(lIdx)).Department
        'Cells(lIdx + 1, 6) = list.item(CInt(lIdx)).Summary
        'Cells(lIdx + 1, 7) = list.item(CInt(lIdx)).Description
        'Cells(lIdx + 1, 8) = list.item(CInt(lIdx)).analyst
        'Cells(lIdx + 1, 9) = list.item(CInt(lIdx)).update_date
        'Cells(lIdx + 1, 10) = list.item(CInt(lIdx)).latest_update
        'Cells(lIdx + 1, 11) = list.item(CInt(lIdx)).Date_Submitted
        'Cells(lIdx + 1, 12) = list.item(CInt(lIdx)).Date_Started
        'Cells(lIdx + 1, 13) = list.item(CInt(lIdx)).Date_Completed
        'Cells(lIdx + 1, 14) = list.item(CInt(lIdx)).Value_To_Business
        'Cells(lIdx + 1, 15) = list.item(CInt(lIdx)).Contact_Name
        'Cells(lIdx + 1, 16) = list.item(CInt(lIdx)).Contact_Info
        
        oPB.value = lIdx
    Next lIdx
    Debug.Print "Time to load worksheet data: " & Timer - t
    ' ----------------------------------------------------------------------

PROC_EXIT:

    If (Not oPB Is Nothing) Then
        oPB.Hide
    End If

    Set oPB = Nothing

End Sub
Public Sub LoadUpdatesProgressBar(list As Variant, size As Long)
' ==========================================================================
' Description : Updates Progress Bar Loading interface
' ==========================================================================

    Const lPB_MIN   As Long = 2
    Dim lPB_MAX   As Long: lPB_MAX = size + 2
    Dim lIdx        As Long
    Dim oPB         As IProgressBar

    ' ----------------------------------------------------------------------
    ' Instantiate the object
    ' ----------------------
    Set oPB = New SpecProgressBar

    ' Set the operational parameters
    ' ------------------------------
    oPB.Caption = "Spec Progress"
    oPB.Min = lPB_MIN
    oPB.Max = lPB_MAX
    oPB.show

    ' Increment the PB
    ' ----------------
    ' Set the cell values to their respective values
    For lIdx = lPB_MIN To lPB_MAX
        Cells(lIdx, 1) = list(0, CInt(lIdx - 2))
        Cells(lIdx, 2) = list(1, CInt(lIdx - 2))
        Cells(lIdx, 3) = list(2, CInt(lIdx - 2))
        Cells(lIdx, 4) = list(3, CInt(lIdx - 2))
        Cells(lIdx, 5) = list(4, CInt(lIdx - 2))
        oPB.value = lIdx
    Next lIdx

    ' ----------------------------------------------------------------------

PROC_EXIT:

    If (Not oPB Is Nothing) Then
        oPB.Hide
    End If

    Set oPB = Nothing

End Sub
