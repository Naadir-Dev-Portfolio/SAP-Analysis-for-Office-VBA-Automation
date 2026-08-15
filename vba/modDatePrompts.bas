Attribute VB_Name = "modDatePrompts"
Option Explicit

' =============================================================================
' COPY-READY AFO DATE PROMPTS
' =============================================================================
' Each function is self-contained and returns one INPUT_STRING value in the AFO
' format dd.mm.yyyy; dd.mm.yyyy. Read only the recipe you need.

' Every date in the previous complete Monday-Sunday weeks.
' Example: PreviousCompleteWeeksDaily(13) returns 91 dates.
Public Function PreviousCompleteWeeksDaily( _
    Optional ByVal weeksBack As Long = 13, _
    Optional ByVal asAtDate As Date = 0) As String

    Dim thisMonday As Date
    Dim firstDate As Date
    Dim lastDate As Date
    Dim currentDate As Date
    Dim resultText As String

    If asAtDate = 0 Then asAtDate = Date
    If weeksBack < 1 Then Err.Raise 5, "PreviousCompleteWeeksDaily", _
                                        "weeksBack must be at least 1."

    thisMonday = asAtDate - Weekday(asAtDate, vbMonday) + 1
    firstDate = thisMonday - (weeksBack * 7)
    lastDate = thisMonday - 1

    For currentDate = firstDate To lastDate
        If Len(resultText) > 0 Then resultText = resultText & "; "
        resultText = resultText & Format$(currentDate, "dd.mm.yyyy")
    Next currentDate

    PreviousCompleteWeeksDaily = resultText
End Function

' One working day for each previous complete week. Monday is moved forward when
' it falls on a weekend or a date listed in Configuration > tblHolidays.
Public Function PreviousWeekWorkingDays( _
    Optional ByVal weeksBack As Long = 13, _
    Optional ByVal asAtDate As Date = 0) As String

    Dim holidays As ListObject
    Dim thisMonday As Date
    Dim workingDate As Date
    Dim isHoliday As Boolean
    Dim weekNumber As Long
    Dim resultText As String

    If asAtDate = 0 Then asAtDate = Date
    If weeksBack < 1 Then Err.Raise 5, "PreviousWeekWorkingDays", _
                                        "weeksBack must be at least 1."

    On Error Resume Next
    Set holidays = ThisWorkbook.Worksheets("Configuration").ListObjects("tblHolidays")
    On Error GoTo 0

    thisMonday = asAtDate - Weekday(asAtDate, vbMonday) + 1

    For weekNumber = weeksBack To 1 Step -1
        workingDate = thisMonday - (weekNumber * 7)

        Do
            isHoliday = False
            If Not holidays Is Nothing Then
                If Not holidays.DataBodyRange Is Nothing Then
                    isHoliday = (Application.CountIf( _
                        holidays.ListColumns("Date").DataBodyRange, CLng(workingDate)) > 0)
                End If
            End If

            If Weekday(workingDate, vbMonday) <= 5 And Not isHoliday Then Exit Do
            workingDate = workingDate + 1
        Loop

        If Len(resultText) > 0 Then resultText = resultText & "; "
        resultText = resultText & Format$(workingDate, "dd.mm.yyyy")
    Next weekNumber

    PreviousWeekWorkingDays = resultText
End Function

' Every date in the current fiscal year. The default fiscal year is April-March.
Public Function CurrentFiscalYearDates( _
    Optional ByVal asAtDate As Date = 0, _
    Optional ByVal fiscalStartMonth As Long = 4) As String

    Dim fiscalStart As Date
    Dim fiscalEnd As Date
    Dim currentDate As Date
    Dim startYear As Long
    Dim resultText As String

    If asAtDate = 0 Then asAtDate = Date
    If fiscalStartMonth < 1 Or fiscalStartMonth > 12 Then
        Err.Raise 5, "CurrentFiscalYearDates", _
                    "fiscalStartMonth must be between 1 and 12."
    End If

    startYear = Year(asAtDate)
    If Month(asAtDate) < fiscalStartMonth Then startYear = startYear - 1

    fiscalStart = DateSerial(startYear, fiscalStartMonth, 1)
    fiscalEnd = DateSerial(startYear + 1, fiscalStartMonth, 0)

    For currentDate = fiscalStart To fiscalEnd
        If Len(resultText) > 0 Then resultText = resultText & "; "
        resultText = resultText & Format$(currentDate, "dd.mm.yyyy")
    Next currentDate

    CurrentFiscalYearDates = resultText
End Function

' First working day of each month in the current fiscal year.
Public Function FiscalMonthFirstWorkingDays( _
    Optional ByVal asAtDate As Date = 0, _
    Optional ByVal fiscalStartMonth As Long = 4) As String

    Dim holidays As ListObject
    Dim fiscalStart As Date
    Dim monthStart As Date
    Dim workingDate As Date
    Dim startYear As Long
    Dim monthNumber As Long
    Dim isHoliday As Boolean
    Dim resultText As String

    If asAtDate = 0 Then asAtDate = Date
    If fiscalStartMonth < 1 Or fiscalStartMonth > 12 Then
        Err.Raise 5, "FiscalMonthFirstWorkingDays", _
                    "fiscalStartMonth must be between 1 and 12."
    End If

    On Error Resume Next
    Set holidays = ThisWorkbook.Worksheets("Configuration").ListObjects("tblHolidays")
    On Error GoTo 0

    startYear = Year(asAtDate)
    If Month(asAtDate) < fiscalStartMonth Then startYear = startYear - 1
    fiscalStart = DateSerial(startYear, fiscalStartMonth, 1)

    For monthNumber = 0 To 11
        monthStart = DateSerial(Year(fiscalStart), Month(fiscalStart) + monthNumber, 1)
        workingDate = monthStart

        Do
            isHoliday = False
            If Not holidays Is Nothing Then
                If Not holidays.DataBodyRange Is Nothing Then
                    isHoliday = (Application.CountIf( _
                        holidays.ListColumns("Date").DataBodyRange, CLng(workingDate)) > 0)
                End If
            End If

            If Weekday(workingDate, vbMonday) <= 5 And Not isHoliday Then Exit Do
            workingDate = workingDate + 1
        Loop

        If Len(resultText) > 0 Then resultText = resultText & "; "
        resultText = resultText & Format$(workingDate, "dd.mm.yyyy")
    Next monthNumber

    FiscalMonthFirstWorkingDays = resultText
End Function
