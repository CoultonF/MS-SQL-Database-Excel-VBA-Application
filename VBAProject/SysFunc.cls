VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SysFunc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function freezePanes(rowNumber As Long)
With ActiveWindow
    If .freezePanes Then .freezePanes = False
    .SplitColumn = 0
    .SplitRow = rowNumber
    .freezePanes = True
End With
End Function
Function convertBooleanYN(arg As Boolean)

    If arg Then
        convertBooleanYN = "Yes"
    Else
        convertBooleanYN = "No"
    End If

End Function
Function joinCommaDelimitedStrings(str1 As String, str2 As String)

    If str1 <> "" Then
    
        If str2 <> "" Then
        
            joinCommaDelimitedStrings = str1 & ", " & str2
        
        Else
        
            joinCommaDelimitedStrings = str1
        
        End If
    
    ElseIf str1 = "" Then
    
        joinCommaDelimitedStrings = str2
    
    End If

End Function

Function getVariantMatchesFrom(arr1 As Variant, arr2 As Variant) As Variant

Dim i As Long
Dim j As Long

arr1 = Flatten(arr1)
arr2 = Flatten(arr2)

Dim usedValues As Variant

For i = LBound(arr1, 1) To UBound(arr1, 1)
    
    For j = LBound(arr2, 1) To UBound(arr2, 1)
        
        If arr1(i) = arr2(j) And isEmpty(usedValues) Then
        
            ReDim usedValues(0)
            
            usedValues(0) = arr1(i)
        
        ElseIf arr1(i) = arr2(j) Then
        
            ReDim Preserve usedValues(UBound(usedValues, 1) + 1)
            
            usedValues(UBound(usedValues, 1)) = arr1(i)
        
        End If
            
    Next j
    
Next i

getVariantMatchesFrom = usedValues

End Function
'Converts any dimension array size into a 1D array
Function Flatten(arrIn) As Variant
Dim v, a, ub
Dim arrout As Variant
    For Each v In arrIn
        If TypeName(v) Like "*()" Then
            arrout = Flatten(v)
        Else
            If isEmpty(arrout) Then
                ReDim arrout(0)
                arrout(0) = v
            Else
                ub = UBound(arrout)
                ReDim Preserve arrout(0 To ub + 1)
                arrout(ub + 1) = v
            End If
        End If
    Next v
Flatten = arrout
End Function
Function convertYNBoolean(arg As String)

    If arg = "Yes" Then
        convertYNBoolean = True
    Else
        convertYNBoolean = False
    End If

End Function
Function setFilter(field As String, operator As String, match As String)

    Set MacroStartPoints.w = ActiveSheet
    
    Dim specObj As New spec
    
    Dim f As New SysFunc
    
    ReDim MacroStartPoints.filterArray(1 To UBound(specObj.getDefaultOrderArray, 1) + 1, 1 To 3)
    
    Set MacroStartPoints.filterArray(findColumnOfHeaderValue(field), 1) = CreateObject("System.Collections.ArrayList")
    
    MacroStartPoints.filterArray(findColumnOfHeaderValue(field), 1).Add operator & match
    
    MacroStartPoints.filterArray(findColumnOfHeaderValue(field), 2) = xlAnd
    
End Function

Function setSortOrder(field As String, ascending As XlSortOrder)

    MacroStartPoints.sortOrder = ascending
    
    MacroStartPoints.sortColumn = findColumnOfHeaderValue(field)

End Function
Function captureFilter()
    
    Call clearFilterData
    
    Set MacroStartPoints.w = ActiveSheet
    'MacroStartPoints.sortOrder = CLng(MacroStartPoints.w.Sort.SortFields.item(1).Order)
    Dim WB As New Worksheet
    Set WB = MacroStartPoints.w
    On Error GoTo noFilterFound
    If Not isEmpty(MacroStartPoints.w.AutoFilter.sort.SortFields.item(1)) Then
        MacroStartPoints.sortOrder = MacroStartPoints.w.AutoFilter.sort.SortFields.item(1).Order
        MacroStartPoints.sortColumn = MacroStartPoints.w.AutoFilter.sort.SortFields.item(1).Key.column
    End If
noFilterFound:
Err.Clear

    'Debug.Print CLng(MacroStartPoints.w.AutoFilter.Sort.SortFields.item(0).Order)
    'MacroStartPoints.sortColumn = MacroStartPoints.w.Sort.SortFields.item.Key.Column
    'Columns("A:C").Sort key1:=Columns(MacroStartPoints.sortColumn), _
      order1:=MacroStartPoints.sortOrder, Header:=xlYes
    ' Capture AutoFilter settings
    If Not MacroStartPoints.w.AutoFilter Is Nothing Then
    With MacroStartPoints.w.AutoFilter
        MacroStartPoints.currentFiltRange = .Range.Address
        With .Filters
            ReDim MacroStartPoints.filterArray(1 To .count, 1 To 3)
            Dim f As Long
            For f = 1 To .count
                With .item(f)
                    If .On Then
                        Dim j As Long: j = 0
                        Set MacroStartPoints.filterArray(f, 1) = CreateObject("System.Collections.ArrayList")
                        Dim testFilterArr As Variant: testFilterArr = MacroStartPoints.filterArray
                        Dim testFilter As Object: Set testFilter = MacroStartPoints.w.AutoFilter
                        If Not isEmpty(.Criteria1) Then
                            'Criteria1 "=Assigned"
                            MacroStartPoints.filterArray(f, 1).Add .Criteria1
                        End If
                        If .operator = 1 Or .operator = 2 Then
                            
                            'Criteria2 "=" matches anything
                            MacroStartPoints.filterArray(f, 1).Add .Criteria2
                            
                        End If
                        If .operator Then
                            'Operator 2
                            MacroStartPoints.filterArray(f, 2) = .operator
                            
                            If Application.Version < "12.0" Then
                            MacroStartPoints.filterArray(f, 3) = .Criteria2 'simply delete this line to make it work in Excel 2010
                            End If
                            
                        End If
                        
                    End If
                End With
            Next f
        End With
    End With
    End If

End Function
Private Function findColumnOfHeaderValue(headerText As String)

    Dim column As Integer: column = 1
    
    Dim specObj As New spec
    
    Dim arrayObj As Variant: arrayObj = specObj.getDefaultOrderArray
    
    For column = 1 To UBound(arrayObj, 1) + 1
    
        If arrayObj(column - 1) = headerText Then
            
            findColumnOfHeaderValue = column
            
            Exit For
            
        End If
    
    Next column

End Function
Function applyFilter()
SpecListController.setFilters
Set MacroStartPoints.w = ActiveSheet
MacroStartPoints.currentFiltRange = MacroStartPoints.w.AutoFilter.Range.Address
Dim columnRow As String
    If MacroStartPoints.sortOrder <> 0 Then
    
    
        Sheets("SPEC SHEET").AutoFilterMode = False
        Sheets("SPEC SHEET").Range("A1", Range("N1").End(xlToLeft)).Select
        columnRow = Col_To_Letter(MacroStartPoints.sortColumn)
        Sheets("SPEC SHEET").UsedRange.AutoFilter
        Sheets("SPEC SHEET").AutoFilter.sort.SortFields.Add Key:=Range(columnRow & "1:" & columnRow & Range(columnRow & Rows.count).End(xlUp).row), SortOn:=xlSortOnValues, Order:=MacroStartPoints.sortOrder
        Sheets("SPEC SHEET").AutoFilter.sort.header = xlYes
        Sheets("SPEC SHEET").AutoFilter.sort.Apply
        Sheets("SPEC SHEET").AutoFilterMode = False
    Else
        
        Sheets("SPEC SHEET").AutoFilterMode = False
        Sheets("SPEC SHEET").Range("A1", Range("N1").End(xlToLeft)).Select
        columnRow = Col_To_Letter(2)
        Sheets("SPEC SHEET").UsedRange.AutoFilter
        Sheets("SPEC SHEET").AutoFilter.sort.SortFields.Add Key:=Range(columnRow & "1:" & columnRow & Range(columnRow & Rows.count).End(xlUp).row), SortOn:=xlSortOnValues, Order:=xlAscending
        Sheets("SPEC SHEET").AutoFilter.sort.header = xlYes
        Sheets("SPEC SHEET").AutoFilter.sort.Apply
        Sheets("SPEC SHEET").AutoFilterMode = False
        
    End If
    If (Not Not MacroStartPoints.filterArray) <> 0 Then
    For MacroStartPoints.col = 1 To UBound(MacroStartPoints.filterArray, 1)
        If Not isEmpty(MacroStartPoints.filterArray(MacroStartPoints.col, 1)) Then
            Dim test2 As Variant: test2 = MacroStartPoints.filterArray
            If Not isEmpty(MacroStartPoints.filterArray(MacroStartPoints.col, 1)) Then
                Dim test As Variant: test = MacroStartPoints.currentFiltRange
                
                
                Select Case test2(MacroStartPoints.col, 1).count
                
                'The filters can be based on upto 14 different criteria within this sheet
                'This method ensures that when using the filter, the application will remember upto 14 different filters at a time. After that, only 14 will be applied regardless of if more criteria need to be matched.
                Case 0
                    If MacroStartPoints.filterArray(MacroStartPoints.col, 2) Then
                    MacroStartPoints.w.Range(MacroStartPoints.currentFiltRange).AutoFilter field:=MacroStartPoints.col, _
                    operator:=MacroStartPoints.filterArray(MacroStartPoints.col, 2)
                    Else
                    MacroStartPoints.w.Range(MacroStartPoints.currentFiltRange).AutoFilter field:=MacroStartPoints.col
                    End If
                Case 1
                    If MacroStartPoints.filterArray(MacroStartPoints.col, 2) Then
                    
                    Debug.Print MacroStartPoints.col
                    Debug.Print MacroStartPoints.filterArray(MacroStartPoints.col, 1).item(0)
                    Debug.Print MacroStartPoints.filterArray(MacroStartPoints.col, 2)
                    Debug.Print MacroStartPoints.col
                    
                    MacroStartPoints.w.Range(MacroStartPoints.currentFiltRange).AutoFilter field:=MacroStartPoints.col, _
                    Criteria1:=MacroStartPoints.filterArray(MacroStartPoints.col, 1).item(0), _
                    operator:=MacroStartPoints.filterArray(MacroStartPoints.col, 2)
                    Else
                    MacroStartPoints.w.Range(MacroStartPoints.currentFiltRange).AutoFilter field:=MacroStartPoints.col, Criteria1:=MacroStartPoints.filterArray(MacroStartPoints.col, 1).item(0)
                    End If
                Case 2
                    Dim vartest1 As Variant: vartest1 = MacroStartPoints.filterArray(MacroStartPoints.col, 1).item(0)
                    Dim vartest2 As Variant: vartest2 = MacroStartPoints.filterArray(MacroStartPoints.col, 1).item(1)
                    If MacroStartPoints.filterArray(MacroStartPoints.col, 2) Then
                    MacroStartPoints.w.Range(MacroStartPoints.currentFiltRange).AutoFilter field:=MacroStartPoints.col, _
                    Criteria1:=MacroStartPoints.filterArray(MacroStartPoints.col, 1).item(0), _
                    Criteria2:=MacroStartPoints.filterArray(MacroStartPoints.col, 1).item(1), _
                    operator:=MacroStartPoints.filterArray(MacroStartPoints.col, 2)
                    Else
                    MacroStartPoints.w.Range(MacroStartPoints.currentFiltRange).AutoFilter field:=MacroStartPoints.col, Criteria1:=MacroStartPoints.filterArray(MacroStartPoints.col, 1).item(0), Criteria2:=MacroStartPoints.filterArray(MacroStartPoints.col, 1).item(1)
                    End If
                End Select

            Else
                MacroStartPoints.w.Range(MacroStartPoints.currentFiltRange).AutoFilter field:=MacroStartPoints.col, _
                Criteria1:=MacroStartPoints.filterArray(MacroStartPoints.col, 1)
            End If
        End If
    Next col
    End If
End Function

'Converts a column number input to the corresponding excel letter of the alphabet
Function Col_To_Letter(lngCol As Long) As String
Dim vArr
vArr = Split(Cells(1, lngCol).Address(True, False), "$")
Col_To_Letter = vArr(0)
End Function
Function clearFilterData()

    Set MacroStartPoints.w = Nothing
    MacroStartPoints.filterArray = Array()
    MacroStartPoints.currentFiltRange = ""
    MacroStartPoints.col = 0
    
End Function

Public Function getUsername()
getUsername = (Environ$("Username"))
End Function
Public Function getComputer()
getComputer = Environ$("computername")
End Function
Public Function resetSheet()

    Application.DisplayAlerts = False
    'Call ActiveWorkbook.Sheets.Add
    'Sheets(2).Delete
    
    Application.Calculation = xlManual
    On Error GoTo nextLine
    ActiveSheet.ShowAllData
nextLine:
    Err.Clear
    ActiveSheet.Cells.EntireColumn.Hidden = False
    ActiveSheet.Cells.Interior.ColorIndex = 0

    Sheets("SPEC SHEET").UsedRange.Delete
    
    Sheets(1).Name = "SPEC SHEET"
    Application.DisplayAlerts = True
    
    

End Function
Function ReplaceN(ByVal str1 As Variant, strFind As String, strReplace As String, N As Long, Optional count As Long) As String
Dim i As Long, j As Long
Dim strM As String
strM = str1
If count <= 0 Then count = 1
For i = 1 To N - 1
    j = InStr(1, strM, strFind)
    strM = Mid(strM, j + Len(strFind), Len(strM))
Next i
If N <= 0 Then
    ReplaceN = str1
Else
    ReplaceN = Mid(str1, 1, Len(str1) - Len(strM)) & Replace(strM, strFind, strReplace, Start:=1, count:=count)
End If
End Function
Function SetRangeBorder()
 
 With Worksheets(1).UsedRange.Borders()
 .LineStyle = xlContinuous
 .Weight = xlThin
 .ColorIndex = 15
 End With
 
End Function
Public Function IsArrayAllocated(arr As Variant) As Boolean
        On Error Resume Next
        IsArrayAllocated = IsArray(arr) And _
                           Not IsError(LBound(arr, 1)) And _
                           LBound(arr, 1) <= UBound(arr, 1)
End Function
Public Function copyFromVariant(values As Variant, startRow As Integer, startCol As Integer, sheet As String)

Dim i As Integer
Dim k As Integer

For i = 0 To UBound(values, 1)

    For k = 0 To UBound(values, 2)
        
        Dim col As Integer: col = startCol + i
        Dim row As Integer: row = startRow + k
        Sheets(sheet).Cells(row, col).value = values(i, k)
        
    Next k

Next i

End Function
Public Function TransposeArray(MyArray As Variant) As Variant
Dim X As Long
Dim Y As Long
Dim Xupper As Long
Dim Yupper As Long
Dim tempArray As Variant
    Xupper = UBound(MyArray, 2)
    Yupper = UBound(MyArray, 1)
    ReDim tempArray(Xupper, Yupper)
    For X = 0 To Xupper
        For Y = 0 To Yupper
            tempArray(X, Y) = MyArray(Y, X)
        Next Y
    Next X
    TransposeArray = tempArray
End Function
Function sliceVariant(arr As Variant, Index As Integer)

Dim i As Integer
Dim j As Integer
Dim size As Integer: size = UBound(arr, 1)
Dim arrTemp() As Variant
ReDim arrTemp(size)

For i = 0 To UBound(arr, 1)
    arrTemp(i) = arr(i, Index)
Next i
sliceVariant = arrTemp
End Function
Public Function DeleteElementAt(ByVal Index As Integer, ByRef prLst As Variant)
       Dim i As Integer

        ' Move all element back one position
        For i = Index + 1 To UBound(prLst)
            prLst(i - 1) = prLst(i)
        Next

        ' Shrink the array by one, removing the last one
        ReDim Preserve prLst(i - 2)
        DeleteElementAt = prLst
End Function
Function createParamString(length As Integer)
Dim str As String
Dim i As Integer
If length > 0 Then
    str = "?"
End If
For i = 1 To length
str = str & " ,?"
Next i
createParamString = str
End Function
Function buildHeader(s_options As Variant, l_xPos As Long, l_yPos As Long)
Dim l_count As Long: l_count = 0
For i = l_yPos To UBound(s_options) + l_yPos
    If s_options(l_count) = "DATE_COMPLETED" And g_blnCompletedCheckboxState = True Then
        Cells(l_xPos, i) = "EST. COMPLETION"
    Else
        Cells(l_xPos, i) = s_options(l_count)
    End If
    
    Rows(l_yPos).Font.Bold = True
    l_count = l_count + 1
Next i
Call freezePanes(l_yPos)
Call SetRangeBorder
End Function

Sub AlternateRowColors()
    Dim LastRow As Long

    With ActiveSheet
        'LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        LastRow = Cells.find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    End With
    With ActiveSheet.UsedRange
        .FormatConditions.Delete
    End With
    With ActiveSheet
        
        With .UsedRange()
            .AutoFilter
            .activate
            .FormatConditions.Delete
            .FormatConditions.Add Type:=xlExpression, Formula1:="=MOD(SUBTOTAL(3,$A$1:$A2),2)"
            .FormatConditions(1).Interior.Color = RGB(240, 240, 240)
        
        End With
    
    End With
    'bring this out to a seperate funtion
    
End Sub
Function requeryRow(specid As Variant)

    lrow = Application.match(CLng(specid), Range("A:A"), 0)

End Function
Function defaultFormats(operation As String)

    Dim specObj As New spec

            With ActiveWorkbook.Sheets(1).UsedRange
            
            .WrapText = True
            
        End With
        
        ActiveWorkbook.Sheets(1).columns("A:M").Interior.ColorIndex = 0
        
        If operation = "SPEC" Or operation = "EXPORT" Then
        
            Dim value As Variant
        
            For Each value In specObj.getDefaultOrderArray
                
                If value = "SPEC_ID" Or value = "RANK" Then
                    Sheets(1).columns(getHeaderColumnIndexOf(CStr(value))).ColumnWidth = 10
                ElseIf value = "STATUS" Then
                    Sheets(1).columns(getHeaderColumnIndexOf(CStr(value))).ColumnWidth = 12
                ElseIf value = "DISCIPLINE" Or value = "DEPARTMENT" Or value = "ANALYST" Or value = "UPDATE_DATE" Or value = "CONTACT_NAME" Or value = "CONTACT_INFO" Then
                    Sheets(1).columns(getHeaderColumnIndexOf(CStr(value))).ColumnWidth = 18
                ElseIf value = "SUMMARY" Then
                    Sheets(1).columns(getHeaderColumnIndexOf(CStr(value))).ColumnWidth = 30
                ElseIf value = "DESCRIPTION" Then
                    Sheets(1).columns(getHeaderColumnIndexOf(CStr(value))).ColumnWidth = 70
                ElseIf value = "LATEST_UPDATE" Then
                    Sheets(1).columns(getHeaderColumnIndexOf(CStr(value))).ColumnWidth = 50
                ElseIf value = "DATE_SUBMITTED" Or value = "DATE_STARTED" Or value = "DATE_COMPLETED" Or value = "VALUE_TO_BUSINESS" Then
                    Sheets(1).columns(getHeaderColumnIndexOf(CStr(value))).ColumnWidth = 21
                Else
                    Sheets("SPEC SHEET").column(getHeaderColumnIndexOf(CStr(value))).ColumnWidth = 15
                End If
            
            Next value
            
            MUserPreferences.hideSpecColumns
            
        ElseIf operation = "UPDATE" Then
            ActiveWorkbook.Sheets(1).columns("A").ColumnWidth = 15
            ActiveWorkbook.Sheets(1).columns("B").ColumnWidth = 70
            ActiveWorkbook.Sheets(1).columns("C:E").ColumnWidth = 20
        End If
ActiveWorkbook.Sheets(1).UsedRange.VerticalAlignment = xlTop
If operation = "SPEC" Or operation = "UPDATE" Then
Dim r As Range
Set r = ActiveCell
Range("A8").Select
With ActiveWindow
    .freezePanes = False
    .ScrollRow = 1
    .ScrollColumn = 1
    .freezePanes = True
    .ScrollRow = r.row
End With
r.Select
End If
End Function
Function deleteHiddenColumns()
    
    Dim hiddenCol As Variant
    
    For Each hiddenCol In MUserPreferences.getHiddenColumns()
    
        columns(Me.getHeaderColumnIndexOf(CStr(hiddenCol))).EntireColumn.Delete
    
    Next hiddenCol

End Function
Function getHeaderColumnIndexOf(columnHeader As String) As Integer
    
    
    Dim response As Variant
    
    On Error Resume Next
    
    response = WorksheetFunction.match(columnHeader, ActiveSheet.Range("1:1"), 0)
    
    If IsError(response) Then
        getHeaderColumnIndexOf = 0
    Else
        getHeaderColumnIndexOf = CInt(response)
    End If

End Function
Function eliminateDuplicate(poArr As Variant) As Variant
Dim poArrNoDup()

    dupArrIndex = -1
For i = LBound(poArr) To UBound(poArr)
        dupBool = False

        For j = LBound(poArr) To i
            If poArr(i) = poArr(j) And Not i = j Then
                dupBool = True
            End If
        Next j

        If dupBool = False Then
            dupArrIndex = dupArrIndex + 1
            ReDim Preserve poArrNoDup(dupArrIndex)
            poArrNoDup(dupArrIndex) = poArr(i)
        End If
Next i

eliminateDuplicate = poArrNoDup
End Function
Public Function UnFreeze()
'Update 20140317
Dim ws As Worksheet
Application.ScreenUpdating = False
For Each ws In Application.ActiveWorkbook.Worksheets
    ws.activate
    With Application.ActiveWindow
        .freezePanes = False
    End With
Next
Application.ScreenUpdating = True
End Function


