VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Rank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private dataObj As DataAccess

Public Function runRankingAlgorithim()
    
    Call issetDataObj
    
    Dim integrityCount As Variant: integrityCount = dataObj.runQuery("SELECT COUNT(SPEC_ID), MAX(RANK) FROM SPEC WHERE RANK IS NOT NULL")

    If Not checkIntegrity(integrityCount) Then
    
        Dim rankedList As Variant
        
        rankedList = dataObj.runQuery("SELECT SPEC_ID, RANK FROM SPEC WHERE RANK IS NOT NULL ORDER BY RANK")
        
        Call repairIntegrity(rankedList, integrityCount)
        
    End If

End Function

Private Function getSpecRankComparison()

    Call issetDataObj
    
    Dim sql As String: sql = "SELECT COUNT(SPEC_ID), MAX(RANK) FROM SPEC WHERE RANK IS NOT NULL"
    
    getSpecRankComparison = dataObj.runQuery(sql)

End Function

Private Function repairIntegrity(rankedList As Variant, Top As Variant)

    For count = 0 To Top(0, 0) - 1
    
        If rankedList(1, count) <> count + 1 Then
        
            Call updateRank(Rank:=count + 1, spec:=CLng(rankedList(0, count)))
            'execQuery ("UPDATE dbo.SPEC SET RANK = " & count + 1 & " WHERE SPEC_ID = " & rankedList(0, count))
            
        End If
        
    Next count
    
End Function

'==========================================================

'||                    CHECK INTERGRITY                  ||

'   results -> an array

'           CHOOSE ONLY ONE OF THE FOLLOWING:

'   rankComparator -> the target rank for update
'             spec -> the target spec ID for update

'==========================================================
Private Function checkIntegrity(results As Variant)

    If results(0, 0) = results(1, 0) Then
    
        checkIntegrity = True
        
    Else
    
        checkIntegrity = False
        
    End If
    
End Function

Private Function getIntegrityCount()

    Call issetDataObj
    
    Dim sql As String: sql = "SELECT COUNT(SPEC_ID), MAX(RANK) FROM SPEC WHERE RANK IS NOT NULL"
    
    getIntegrityCount = dataObj.runQuery(sql)

End Function

Private Function getAllRanksOrdered()

    Call issetDataObj
    
    Dim sql As String: sql = "SELECT SPEC_ID, RANK FROM SPEC WHERE RANK IS NOT NULL ORDER BY RANK"
    
    getAllRanksOrdered = dataObj.runQuery(sql)

End Function

Public Function getMaxRank()

    Call issetDataObj
    
    Dim sql As String: sql = "SELECT MAX(RANK) FROM dbo.SPEC WHERE RANK IS NOT NULL"
    
    getMaxRank = dataObj.runQuery(sql)(0, 0)

End Function

Private Function getRankForSpec(s_specID As Long)

    Call issetDataObj
    
    Dim sql As String: sql = "SELECT RANK FROM dbo.SPEC WHERE SPEC_ID = ?"
    
    getRankForSpec = dataObj.runQuery(sql, Array(s_specID))(0, 0)

End Function

'==========================================================

'||         PRIVATE USE ONLY  -  UPDATE RANK             ||

'   rank -> the VALUE being put into the row


'   WHERE CONDITIONALS - PASS ONLY ONE OF THE FOLLOWING:

'   rankComparator -> the target RANK for RANK update
'             spec -> the target SPEC_ID for RANK update

'==========================================================
Private Function updateRank(Optional Rank As Long, Optional spec As Long, Optional rankComparator As Long)
    
    Dim rankParam As String
    
    Dim sql As String
    
    Call issetDataObj
    
    'CASE 1 sql// update RANK where SPEC_ID = spec
    If Not IsMissing(spec) And spec > 0 Then
    
        If IsMissing(Rank) Then
        
            rankParam = "NULL"
            
        Else
        
            rankParam = "?"
            
        End If
        
        sql = "UPDATE dbo.SPEC SET RANK = " & rankParam & " WHERE SPEC_ID = ?"
    
        Call dataObj.runQuery(sql, Array(Rank, spec))
    
    
    'CASE 2 sql// update RANK where RANK = rankComparator
    ElseIf Not IsMissing(rankComparator) Then
        Dim args As Variant
        If IsMissing(Rank) Then
    
            rankParam = "NULL"
            args = Array(rankComparator)
            
        Else
        
            rankParam = "?"
            args = Array(Rank, rankComparator)
            
        End If
        
        sql = "UPDATE dbo.SPEC SET RANK = " & rankParam & " WHERE RANK = ?"
        dataObj.init
        Call dataObj.runQuery(sql, args)
    
    
    
    End If
    
End Function

'==========================================================

'||                INCREMENT RANKS                        ||


'   specID -> the current spec number being changed

'           CHOOSE ONLY ONE OF THE THREE OPTIONAL PARAMS
'   insertingRank -> the target rank for specID
'   updatingRank -> the target rank for specID
'   deletingRank -> set to true for specID rank removal

'==========================================================
Public Function incrementRanks(specid As Variant, Optional insertingRank As Integer, Optional updatingrank As Integer, Optional deletingRank As Boolean)
'call before new row is inserted

    
    'Check if data access is available
    Call issetDataObj

    Dim maxRank As Integer
    
    Dim integrityCount As Variant: integrityCount = getIntegrityCount()
        If Not checkIntegrity(integrityCount) Then
            Dim rankedList As Variant
            rankedList = getAllRanksOrdered()
            Call repairIntegrity(rankedList, integrityCount)
        End If
        maxRank = getMaxRank
        currentRank = getRankForSpec(CLng(specid))
        If (currentRank = insertingRank) Or (updatingrank = currentRank) Or (currentRank = deletingRank) And currentRank <> 0 Then
        Exit Function
        End If
        If insertingRank <> 0 Or Not IsNumeric(currentRank) Then
        If insertingRank = 0 And Not IsNumeric(currentRank) Then
            insertingRank = updatingrank
        End If
            If insertingRank <= maxRank And insertingRank > 0 Then
                Do While maxRank >= insertingRank
                Call updateRank(Rank:=maxRank + 1, rankComparator:=CLng(maxRank))
                maxRank = maxRank - 1
                Loop
                Call updateRank(spec:=CLng(specid), Rank:=maxRank + 1)
                'execQuery ("UPDATE dbo.SPEC SET RANK = " & maxRank + 1 & " WHERE SPEC_ID = " & specID)
                incrementRanks = 0
            ElseIf (insertingRank > maxRank) Then
                incrementRanks = maxRank + 1
                Call updateRank(spec:=CLng(specid), Rank:=maxRank + 1)
                'execQuery ("UPDATE dbo.SPEC SET RANK = " & maxRank + 1 & " WHERE SPEC_ID = " & specID)
            Else
                incrementRanks = False
            End If
        ElseIf updatingrank <> 0 Then
            If updatingrank <= maxRank And updatingrank > 0 Then
            
                If currentRank > updatingrank Then
                    Call updateRank(spec:=CLng(specid))
                    'execQuery ("UPDATE dbo.SPEC SET RANK = NULL WHERE SPEC_ID = " & specID)
                    Do While currentRank > updatingrank
                    Call updateRank(Rank:=CLng(currentRank), rankComparator:=currentRank - 1)
                    'execQuery ("UPDATE dbo.SPEC SET RANK = " & currentrank & " WHERE RANK = " & currentrank - 1)
                    currentRank = currentRank - 1
                    Loop
                ElseIf currentRank < updatingrank Then
                    Call updateRank(spec:=CLng(specid))
                    'execQuery ("UPDATE dbo.SPEC SET RANK = NULL WHERE SPEC_ID = " & specID)
                    Do While currentRank < updatingrank
                    Call updateRank(Rank:=CLng(currentRank), rankComparator:=currentRank + 1)
                    'execQuery ("UPDATE dbo.SPEC SET RANK = " & currentrank & " WHERE RANK = " & currentrank + 1)
                    currentRank = currentRank + 1
                    Loop
                End If
                maxRank = getMaxRank()
                If Not (currentRank = updatingrank And maxRank = updatingrank) Then
                Call updateRank(spec:=CLng(specid), Rank:=CLng(updatingrank))
                'execQuery ("UPDATE dbo.SPEC SET RANK = " & updatingrank & " WHERE SPEC_ID = " & specID)
                End If
            ElseIf updatingrank > maxRank Then
                Call incrementRanks(updatingrank:=maxRank, specid:=specid)
            ElseIf updatingrank = 0 Then
            MsgBox "0 is not a valid rank"
            End If
        ElseIf deletingRank Then
            Dim Rank As Integer: Rank = getRankForSpec(CLng(specid))
            Call dataObj.runQuery("UPDATE dbo.SPEC SET RANK = NULL WHERE RANK = ? AND SPEC_ID = ?", Array(Rank, specid))
            'execQuery (" & specID)
            Rank = Rank + 1
            Do While Rank <= maxRank
                Call updateRank(Rank:=Rank - 1, rankComparator:=CLng(Rank))
                'execQuery ("UPDATE dbo.SPEC SET RANK = " & rank - 1 & " WHERE RANK = " & rank)
                Rank = Rank + 1
            Loop
        Else
            incrementRanks = False
        End If
    
End Function

Private Function issetDataObj()

    If dataObj Is Nothing Then
    
        Set dataObj = New DataAccess
        
        dataObj.init
        
    End If

End Function

Private Sub Class_Initialize()

    Set dataObj = New DataAccess
    
    dataObj.init
    
End Sub
