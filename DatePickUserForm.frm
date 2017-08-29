VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DatePickUserForm 
   Caption         =   "Select Date"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3825
   OleObjectBlob   =   "DatePickUserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DatePickUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If Me.Tag = "X" Then Exit Sub
    If Cancel <> 1 Then Cancel = 1
End Sub
Private Sub UserForm_Activate()
    If MinDate.Caption = "" Then MinDate.Caption = DateSerial(101, 1, 1)
    If IsDate(MinDate.Caption) = False Then
        Me.Hide
        MsgBox ("Invalid MinDate")
        Exit Sub
    End If
    If CDate(MinDate.Caption) < DateSerial(101, 1, 1) Then
        Me.Hide
        MsgBox ("MinDate must be 1/1/101 or greater")
        Exit Sub
    End If
    If MaxDate.Caption = "" Then MaxDate.Caption = DateSerial(9999, 12, 31)
    If IsDate(MaxDate.Caption) = False Then
        Me.Hide
        MsgBox ("Invalid MaxDate")
        Exit Sub
    End If
    If CDate(MaxDate.Caption) > DateSerial(9999, 12, 31) Then
        Me.Hide
        MsgBox ("MaxDate must be 31/12/9999 or less")
        Exit Sub
    End If
    If CDate(MinDate.Caption) > CDate(MaxDate.Caption) Then
        Me.Hide
        MsgBox ("MinDate must be less than MaxDate")
        Exit Sub
    End If
    If StartDate.Caption = "" Then StartDate.Caption = Date
    If IsDate(StartDate.Caption) = False Then
        Me.Hide
        MsgBox ("Invalid StartDate")
        Exit Sub
    End If
    If CDate(StartDate.Caption) < CDate(MinDate.Caption) Or CDate(StartDate.Caption) > CDate(MaxDate.Caption) Then
        Me.Hide
        MsgBox ("StartDate must be between MinDate and MaxDate")
        Exit Sub
    End If
    If Date >= CDate(MinDate.Caption) And Date <= CDate(MaxDate.Caption) Then
        TodayCB.Enabled = True
    Else
        TodayCB.Enabled = False
    End If
    Select Case PickDateShort
    Case "d/m/yy", "dd/mm/yy", "dd/mm/yyyy", "m/d/yy", "mm/dd/yy", "mm/dd/yyyy", "yyyy/mm/dd"
        PickDateShort.Tag = PickDateShort
    Case Else
        PickDateShort.Tag = "dd/mm/yyyy"
    End Select
    PickDateLong.Tag = PickDateLong
    Call UpdateCalendarPanel(StartDate.Caption)
End Sub
Private Sub z11_Click()
    Call DateButtonClick
End Sub
Private Sub z12_Click()
    Call DateButtonClick
End Sub
Private Sub z13_Click()
    Call DateButtonClick
End Sub
Private Sub z14_Click()
    Call DateButtonClick
End Sub
Private Sub z15_Click()
    Call DateButtonClick
End Sub
Private Sub z16_Click()
    Call DateButtonClick
End Sub
Private Sub z17_Click()
    Call DateButtonClick
End Sub
Private Sub z21_Click()
    Call DateButtonClick
End Sub
Private Sub z22_Click()
    Call DateButtonClick
End Sub
Private Sub z23_Click()
    Call DateButtonClick
End Sub
Private Sub z24_Click()
    Call DateButtonClick
End Sub
Private Sub z25_Click()
    Call DateButtonClick
End Sub
Private Sub z26_Click()
    Call DateButtonClick
End Sub
Private Sub z27_Click()
    Call DateButtonClick
End Sub
Private Sub z31_Click()
    Call DateButtonClick
End Sub
Private Sub z32_Click()
    Call DateButtonClick
End Sub
Private Sub z33_Click()
    Call DateButtonClick
End Sub
Private Sub z34_Click()
    Call DateButtonClick
End Sub
Private Sub z35_Click()
    Call DateButtonClick
End Sub
Private Sub z36_Click()
    Call DateButtonClick
End Sub
Private Sub z37_Click()
    Call DateButtonClick
End Sub
Private Sub z41_Click()
    Call DateButtonClick
End Sub
Private Sub z42_Click()
    Call DateButtonClick
End Sub
Private Sub z43_Click()
    Call DateButtonClick
End Sub
Private Sub z44_Click()
    Call DateButtonClick
End Sub
Private Sub z45_Click()
    Call DateButtonClick
End Sub
Private Sub z46_Click()
    Call DateButtonClick
End Sub
Private Sub z47_Click()
    Call DateButtonClick
End Sub
Private Sub z51_Click()
    Call DateButtonClick
End Sub
Private Sub z52_Click()
    Call DateButtonClick
End Sub
Private Sub z53_Click()
    Call DateButtonClick
End Sub
Private Sub z54_Click()
    Call DateButtonClick
End Sub
Private Sub z55_Click()
    Call DateButtonClick
End Sub
Private Sub z56_Click()
    Call DateButtonClick
End Sub
Private Sub z57_Click()
    Call DateButtonClick
End Sub
Private Sub z61_Click()
    Call DateButtonClick
End Sub
Private Sub z62_Click()
    Call DateButtonClick
End Sub
Private Sub z63_Click()
    Call DateButtonClick
End Sub
Private Sub z64_Click()
    Call DateButtonClick
End Sub
Private Sub z65_Click()
    Call DateButtonClick
End Sub
Private Sub z66_Click()
    Call DateButtonClick
End Sub
Private Sub z67_Click()
    Call DateButtonClick
End Sub
Private Sub DateButtonClick()
    DayINT = Val(Me.ActiveControl.Caption)
    MonthINT = Val(ml.Caption)
    YearINT = Val(yl.Caption)
    NewDateINT = DateSerial(YearINT, MonthINT, DayINT)
    Call UpdateCalendarPanel(NewDateINT)
End Sub
Private Sub DaySB_SpinUp()
    DayINT = Val(dl.Caption)
    MonthINT = Val(ml.Caption)
    YearINT = Val(yl.Caption)
    NewDateINT = DateAdd("d", 1, DateSerial(YearINT, MonthINT, DayINT))
    If MaxDate.Caption <> "" Then
        If NewDateINT > CDate(MaxDate.Caption) Then Exit Sub
    End If
    Call UpdateCalendarPanel(NewDateINT)
End Sub
Private Sub DaySB_SpinDown()
    DayINT = Val(dl.Caption)
    MonthINT = Val(ml.Caption)
    YearINT = Val(yl.Caption)
    NewDateINT = DateAdd("d", -1, DateSerial(YearINT, MonthINT, DayINT))
    If NewDateINT < CDate(MinDate.Caption) Then Exit Sub
    Call UpdateCalendarPanel(NewDateINT)
End Sub
Private Sub MonthSB_SpinUp()
    DayINT = Val(dl.Caption)
    MonthINT = Val(ml.Caption)
    YearINT = Val(yl.Caption)
    NewDateINT = DateAdd("m", 1, DateSerial(YearINT, MonthINT, DayINT))
    If MaxDate.Caption <> "" Then
        If NewDateINT > CDate(MaxDate.Caption) Then Exit Sub
    End If
    Call UpdateCalendarPanel(NewDateINT)
End Sub
Private Sub MonthSB_SpinDown()
    DayINT = Val(dl.Caption)
    MonthINT = Val(ml.Caption)
    YearINT = Val(yl.Caption)
    NewDateINT = DateAdd("m", -1, DateSerial(YearINT, MonthINT, DayINT))
    If NewDateINT < CDate(MinDate.Caption) Then Exit Sub
    Call UpdateCalendarPanel(NewDateINT)
End Sub
Private Sub YearSB_SpinUp()
    DayINT = Val(dl.Caption)
    MonthINT = Val(ml.Caption)
    YearINT = Val(yl.Caption)
    NewDateINT = DateAdd("yyyy", 1, DateSerial(YearINT, MonthINT, DayINT))
    If MaxDate.Caption <> "" Then
        If NewDateINT > CDate(MaxDate.Caption) Then Exit Sub
    End If
    Call UpdateCalendarPanel(NewDateINT)
End Sub
Private Sub YearSB_SpinDown()
    DayINT = Val(dl.Caption)
    MonthINT = Val(ml.Caption)
    YearINT = Val(yl.Caption)
    NewDateINT = DateAdd("yyyy", -1, DateSerial(YearINT, MonthINT, DayINT))
    If NewDateINT < CDate(MinDate.Caption) Then Exit Sub
    Call UpdateCalendarPanel(NewDateINT)
End Sub
Private Sub DoneCB_Click()
    Me.Hide
End Sub
Private Sub TodayCB_Click()
    Call UpdateCalendarPanel(Date)
End Sub
Private Sub CancelCB_Click()
    Me.PickDateShort = ""
    Me.PickDateLong = ""
    Me.Hide
End Sub
Sub UpdateCalendarPanel(NewDateINT)
    YearINT = Year(NewDateINT)
    MonthINT = Month(NewDateINT)
    DayINT = Day(NewDateINT)
    FirstDayINT = DateSerial(YearINT, MonthINT, 1)
    FirstDaySTR = Format(FirstDayINT, "dddd")
    If MonthINT = 1 Then pm = 12 Else pm = MonthINT - 1
    If MonthINT = 12 Then nm = 1 Else nm = MonthINT + 1
    If UCase(Left(FirstDaySTR, 2)) = "MO" Then sc = 1: ButtonValue = 0
    If UCase(Left(FirstDaySTR, 2)) = "TU" Then sc = 2: ButtonValue = Day(DateSerial(YearINT, pm + 1, 0)) - 1
    If UCase(Left(FirstDaySTR, 2)) = "WE" Then sc = 3: ButtonValue = Day(DateSerial(YearINT, pm + 1, 0)) - 2
    If UCase(Left(FirstDaySTR, 2)) = "TH" Then sc = 4: ButtonValue = Day(DateSerial(YearINT, pm + 1, 0)) - 3
    If UCase(Left(FirstDaySTR, 2)) = "FR" Then sc = 5: ButtonValue = Day(DateSerial(YearINT, pm + 1, 0)) - 4
    If UCase(Left(FirstDaySTR, 2)) = "SA" Then sc = 6: ButtonValue = Day(DateSerial(YearINT, pm + 1, 0)) - 5
    If UCase(Left(FirstDaySTR, 2)) = "SU" Then sc = 7: ButtonValue = Day(DateSerial(YearINT, pm + 1, 0)) - 6
    LastRow = 6
    vi = True
    For rloop = 1 To 6
    For cloop = 1 To 7
    ButtonValue = ButtonValue + 1
    If rloop = 1 Then
        If cloop < sc Then en = False Else en = True
        If ButtonValue > Day(DateSerial(YearINT, pm + 1, 0)) Then ButtonValue = 1
    End If
    If rloop > 1 Then
        If ButtonValue > Day(DateSerial(YearINT, MonthINT + 1, 0)) Then
            ButtonValue = 1
            en = False
            If cloop = 1 Then
                LastRow = rloop - 1
            Else
                LastRow = rloop
            End If
        End If
    End If
    If rloop > LastRow Then vi = False
    ButtonName = "z" + Format(rloop) + Format(cloop)
    If ButtonValue = DayINT And en = True Then
        fb = True
        fs = 10
    Else
        fb = False
        fs = 8
    End If
    Me.Controls(ButtonName).Caption = ButtonValue
    Me.Controls(ButtonName).Font.Bold = fb
    Me.Controls(ButtonName).Font.size = fs
    Me.Controls(ButtonName).Enabled = en
    Me.Controls(ButtonName).Visible = vi
    Next cloop
    Next rloop
    Me.dl.Caption = DayINT
    Me.ml.Caption = MonthINT
    Me.yl.Caption = YearINT
    Me.MonthLabel.Caption = Format(DateSerial(YearINT, MonthINT, DayINT), "mmmm yyyy")
    Me.PickDateShort.Caption = Format(DateSerial(YearINT, MonthINT, DayINT), Me.PickDateShort.Tag)
    Dim N As Long
    Const csfx = "stndrdthththththth"
    N = DayINT Mod 100
    If ((Abs(N) >= 10) And (Abs(N) <= 19)) Or ((Abs(N) Mod 10) = 0) Then
        OrdinalNumber = Format(DayINT) & "th"
    Else
        OrdinalNumber = Format(DayINT) & Mid(csfx, ((Abs(N) Mod 10) * 2) - 1, 2)
    End If
    PickDateLongSTR = Format(DateSerial(YearINT, MonthINT, DayINT), "dddd, ")
    If Me.PickDateLong.Tag = "dddd, mmmm d, yyyy" Then
        PickDateLongSTR = PickDateLongSTR & Format(DateSerial(YearINT, MonthINT, DayINT), "mmmm ") & OrdinalNumber & Format(DateSerial(YearINT, MonthINT, DayINT), ", yyyy")
    Else
        PickDateLongSTR = PickDateLongSTR & OrdinalNumber & Format(DateSerial(YearINT, MonthINT, DayINT), " mmmm yyyy")
    End If
    Me.PickDateLong.Caption = PickDateLongSTR
End Sub

