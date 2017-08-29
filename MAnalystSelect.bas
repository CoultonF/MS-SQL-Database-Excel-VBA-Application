Attribute VB_Name = "MAnalystSelect"
Function showAnalystSelect(Optional currentAnalysts As Variant)
AnalystMenu.show
Dim analystList As String
For i = 0 To AnalystMenu.ListBox2.ListCount - 1
    AnalystString = AnalystString & AnalystMenu.ListBox2.list(i, 0) & " "
Next i
AnalystString = Trim(AnalystString)
showAnalystSelect = AnalystString

End Function
