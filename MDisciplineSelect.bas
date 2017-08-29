Attribute VB_Name = "MDisciplineSelect"
Public disciplineArray As Object

Function showDisciplineSelect(Optional currentDisciplines As Variant)

Set disciplineArray = CreateObject("System.Collections.ArrayList")

DisciplineMenu.SortListBox
DisciplineMenu.show
Dim disciplineList As String
Dim count As Integer: count = disciplineArray.count
Dim i As Integer: i = 0
While i < count
    DisciplineString = DisciplineString & vbNewLine & disciplineArray.item(i)
    i = i + 1
Wend
'For i = 0 To UBound(disciplineArray, 2)
 '   DisciplineString = DisciplineString & disciplineArray(0, i) & " "
'Next i
DisciplineString = Trim(DisciplineString)
If Len(DisciplineString) > 2 Then
DisciplineString = Right(DisciplineString, Len(DisciplineString) - 2)
End If
showDisciplineSelect = DisciplineString

End Function
