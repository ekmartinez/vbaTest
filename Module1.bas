Attribute VB_Name = "Module1"

Sub Try()
    
    Dim dname As Date
    Range("C4") = WeekdayName(Weekday(Range("b4"), 0), True, vbUseSystemDayOfWeek)

End Sub
