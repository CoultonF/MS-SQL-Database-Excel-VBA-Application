Attribute VB_Name = "MDateSelect"

Function ShowDatePickerUserForm()
    
    Dim dateStr As String
    
    
    DatePickUserForm.MinDate = ""
    'If MinDate is not specified then 1/1/101 is assumed
    'If MinDate is specified then MinDate must be a valid date equal to or greater than 1/1/101 else an error resluts
    'MinDate must be less than MaxDate else an error results
    
    DatePickUserForm.MaxDate = ""
    'If MaxDate is not specified then 31/12/9999 is assumed
    'If MaxDate is specified then MaxDate must be a valid date equal to or less than 31/12/9999 else an error resluts
    'MaxDate must be greater than MinDate else an error results
    
    DatePickUserForm.StartDate = Date
    'If StartDate is not specified and the current system date is between MinDate and MaxDate then the current system date is assumed
    'If StartDate is not specified and the current system date is not between MinDate and MaxDate then an error results
    'If StartDate is specified then StartDate must be a valid date between MinDate and MaxDate else an error results
    
    DatePickUserForm.PickDateShort = "yyyy/mm/dd"
    'PickDateShort can be formatted using the following standard date abbreviations only: _
        "d/m/yy"        = 1/4/12 _
        "dd/mm/yy"      = 01/04/12 _
        "dd/mm/yyyy"    = 01/04/2012 _
        "m/d/yy"        = 4/1/12 _
        "mm/dd/yy"      = 04/01/12 _
        "mm/dd/yyyy"    = 04/01/2012
    'If PickDateShort is not specified or invalid then "dd/mm/yyyy" is assumed
    
    DatePickUserForm.PickDateLong = "dddd, d mmmm yyyy"
    'PickDateLong can be formatted using the following standard date abbreviations only: _
        "dddd, d mmmm yyyy"        = Sunday, 1st April 2012 _
        "dddd, mmmm d, yyyy"        = Sunday, April 1st, 2012 _
    'If PickDateLong is not specified or invalid then "dddd, d mmmm yyyy" is assumed
    
    DatePickUserForm.TodayCB.Enabled = True
    'Today Button will only be available if current date is between MinDate and MaxDate
    'Today Button can be disabled by changing the above setting to False
    
    DatePickUserForm.CancelCB.Enabled = True
    'Cancel Button will return a "" result
    'Cancel Button can be disabled by changing the above setting to False
        
    
    'Result date is retrieved from DatePickUserForm.PickDateShort & DatePickUserForm.PickDateLong (see lines below)
    DatePickUserForm.show
    result1 = DatePickUserForm.PickDateShort
    result2 = DatePickUserForm.PickDateLong
    
    dateStr = result1
    
    ShowDatePickerUserForm = result1
    
End Function
