VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "calendarClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Private StylePrefix As String
Private AllowWeekdays As Boolean
Private CalendarTitle As String
Private CalendarPageLink As String

Private MonthSelected As Long
Private YearSelected As Long
Private DaySelected As Long
Private Keywords As String

Private WorkingQueryString As String

'Private GivenGroupName As String
'Private GivenDate As Date

Private Main As Object
Private CSV As Object

Public Function Execute(CsvObject As Object, MainObject As Object, optionString As String, FilterInput As String) As String
    Set CSV = CsvObject
    
    Call Init(MainObject)
    
    Execute = GetContent(optionString)
    
End Function

Friend Sub Init(MainObject As Object)
    Set Main = MainObject
    
    StylePrefix = "asa"
    AllowWeekdays = True
    
    Exit Sub

ErrorTrap:
    Call HandleError("calendarClass", "Init", Err.Number, Err.Source, Err.Description, True, False)
    End Sub

Private Function GetContent(optionString As String) As String
    On Error GoTo ErrorTrap
    
    Dim Stream As String
    Dim GivenGroupName As String
    Dim GivenDate As Date
    Dim dayLink As String
    Dim eventDetails As String
    Dim optionString2 As String
    
    GivenGroupName = Main.GetAddonOption("Group", optionString)
    GivenDate = KmaEncodeDate(Main.GetAddonOption("CurrentDate", optionString))
    If GivenDate = CDate(0) Then
        GivenDate = Int(Now())
    End If
    '
    optionString2 = "" _
        & RequestNameFormID & "=100" _
        & "&skipProcessHack=1" _
        & "&" & RequestNameGroupName & "=" & GivenGroupName _
        & "&targetdate=" & Month(GivenDate) & "/" & Day(GivenDate) & "/" & Year(GivenDate)
    eventDetails = Main.ExecuteAddon(0, "calendarResponder", optionString2, ContextPage, "", 0, "", "", 0)
    '
    Stream = "<div id=""calWrapper"">"
    Stream = Stream & "<table width=""100%"" padding=""2"" cellspacing=""0"" border=""0"">"
    Stream = Stream & "<tr>"
    Stream = Stream & "<td class=""calRowSearch"">" & GetSearch(GivenGroupName) & "</td>"
    Stream = Stream & "<td class=""calRowCalendar""><div id=""calContainer"">" & GetCalendar(GivenGroupName, GivenDate) & "</div></td>"
    Stream = Stream & "</tr>"
    Stream = Stream & "<tr>"
    Stream = Stream & "<td colspan=""2"" class=""collabEventCell""><div id=""eventContainer"">" & eventDetails & "</div></td>"
    Stream = Stream & "</tr>"
    Stream = Stream & "</table>"
    Stream = Stream & "</div>"
    
    GetContent = Stream
    
    Exit Function
ErrorTrap:
    Call HandleError("calendarClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function

Friend Function GetCalendar(GroupName As String, CurrentDate As Date) As String
    On Error GoTo ErrorTrap
    
    Dim Stream As String
    Dim DaysInMonth As Long    ' Days In Month
    Dim FirstDayofWeek As Long    ' Day Of Week that month starts on
    Dim CurrentDay As Long    ' Variable we use to hold current day of month as we write table
    Dim CurrentPosition As Long   ' Variable we use to hold current position in table
    
    Dim ClassNameHeader As String
    Dim ClassNameTable As String
    Dim ClassNameDayRow As String
    Dim ClassNameDateRow As String
    Dim ClassNameDateHighlighted As String
    Dim ClassNameFooter As String
    Dim ClassNameDateLink As String
    Dim ClassNameDateWithEvent As String
    Dim ClassNameDateRowFiller As String
    
    Dim ThisDate As Date
    Dim dayLink As String
    
    Dim Oye As String
    Dim DateString As String
    Dim HightLight As Boolean
    Dim MonthStorage As Long
    Dim YearStorage As Long
    Dim BackLink As String
    Dim ForwardLink As String
    
    Dim calendarName As String
    Dim CalendarEventsBakeName As String
    Dim CalendarEventsBakeContent As String
    Dim CalendarEventsBakeNameQuickCalendar As String
    Dim CalendarEventsBakeNameMarginaclCalendar As String

    ClassNameTable = StylePrefix & "Table"
    ClassNameHeader = StylePrefix & "Header"
    ClassNameDayRow = StylePrefix & "WeekRow"
    ClassNameDateRow = StylePrefix & "DayRow"
    ClassNameDateHighlighted = StylePrefix & "DayHighlight"
    ClassNameDateLink = StylePrefix & "Link"
    ClassNameFooter = StylePrefix & "Footer"
    ClassNameDateRowFiller = StylePrefix & "Filler"
    ClassNameDateWithEvent = StylePrefix & "Event"
    calendarName = GroupName & " Calendar"

    CalendarEventsBakeName = "CalendarEventsBake " & CDate(Date) & " " & StylePrefix & " " & MonthSelected
    'CalendarEventsBakeContent = Main.ReadBake(CalendarEventsBakeName)
    
    If CalendarEventsBakeContent <> "" Then
        Stream = CalendarEventsBakeContent
    Else
            If MonthSelected = 0 Then
                If CurrentDate = CDate(0) Then
                    CurrentDate = Date
                    MonthSelected = Month(Date)
                Else
                    MonthSelected = Month(CurrentDate)
                End If
            Else
                If MonthSelected < Month(Now) Then
                    If YearSelected = 0 Then
                        ThisDate = CDate(MonthSelected & "/" & "1/" & (Year(Date) + 1))
                    ElseIf YearSelected <> 0 Then
                        ThisDate = CDate(MonthSelected & "/" & "1/" & (YearSelected))
                    End If
                Else
                    ThisDate = CDate(MonthSelected & "/" & "1/" & Year(Date))
                End If
                CurrentDate = ThisDate
            End If

            MonthStorage = Month(CurrentDate)
            YearStorage = Year(CurrentDate)
            If MonthStorage = 1 Then
                MonthStorage = 12
                YearStorage = YearStorage - 1
            Else
                MonthStorage = MonthStorage - 1
            End If
            BackLink = RequestNameFormID & "=200&" & RequestNameGroupName & "=" & GroupName & "&targetdate=" & MonthStorage & "/1/" & YearStorage
            
            MonthStorage = Month(CurrentDate)
            YearStorage = Year(CurrentDate)
            If MonthStorage = 12 Then
                MonthStorage = 1
                YearStorage = YearStorage + 1
            Else
                MonthStorage = MonthStorage + 1
            End If
            ForwardLink = RequestNameFormID & "=200&" & RequestNameGroupName & "=" & GroupName & "&targetdate=" & MonthStorage & "/1/" & YearStorage
            

            DateString = getDateString(CurrentDate, calendarName)
            Call Main.TestPoint("DateString:" & DateString)
            
            DaysInMonth = GetDaysInMonth(Month(CurrentDate), Year(CurrentDate))
            FirstDayofWeek = GetWeekdayMonthStartsOn(CurrentDate)
            
            Stream = Stream & "<table class=""" & ClassNameTable & """>"
            If (Trim(CalendarTitle) = "") Then
                CalendarTitle = MonthName(Month(CurrentDate)) & " " & Year(CurrentDate)
            End If
            
            Stream = Stream & "<tr class=""" & ClassNameHeader & """ colspan=""7"">"
            Stream = Stream & "<td><a onClick=""GetAjax('calendarResponder','" & BackLink & "','','calContainer','',''); return false;"" href=""#""><img border=0 src=""/upload/back.gif""></a></td>"
            Stream = Stream & "<td width=100% colspan=""5"">" & CalendarTitle & "</td>"
            Stream = Stream & "<td><a onClick=""GetAjax('calendarResponder','" & ForwardLink & "','','calContainer','',''); return false;"" href=""#""><img border=0 src=""/upload/forward.gif""></a></td>"
            Stream = Stream & "</tr>"
            
            If AllowWeekdays Then
                Stream = Stream & "<tr>"
                Stream = Stream & "<td class=""" & ClassNameDayRow & """ width=""14%"">S</td>"
                Stream = Stream & "<td class=""" & ClassNameDayRow & """ width=""14%"">M</td>"
                Stream = Stream & "<td class=""" & ClassNameDayRow & """ width=""14%"">T</td>"
                Stream = Stream & "<td class=""" & ClassNameDayRow & """ width=""14%"">W</td>"
                Stream = Stream & "<td class=""" & ClassNameDayRow & """ width=""14%"">T</td>"
                Stream = Stream & "<td class=""" & ClassNameDayRow & """ width=""14%"">F</td>"
                Stream = Stream & "<td class=""" & ClassNameDayRow & """ width=""14%"">S</td>"
                Stream = Stream & "</tr>"
            End If
            
            If FirstDayofWeek <> 1 Then
                Stream = Stream & "<tr>"
                CurrentPosition = 1
                Do While CurrentPosition < FirstDayofWeek
                    Stream = Stream & "<td class=""" & ClassNameDateRowFiller & """>&nbsp;</td>"
                    CurrentPosition = CurrentPosition + 1
                Loop
            End If
            
            CurrentDay = 1
            CurrentPosition = FirstDayofWeek
            Do While CurrentDay <= DaysInMonth
                HightLight = False

                Oye = "," & CurrentDay & ","
                If InStr(1, DateString, Oye, vbTextCompare) Then
                    HightLight = True
                End If

                If CurrentPosition = 1 Then
                    Stream = Stream & "<tr>"
                End If
                
                dayLink = RequestNameFormID & "=100&" & RequestNameGroupName & "=" & GroupName & "&targetdate=" & Month(CurrentDate) & "/" & CurrentDay & "/" & Year(CurrentDate)
                
                If (CurrentDay = Day(Date)) And (MonthSelected = Month(Date)) Then
                    Stream = Stream & "<td class=""" & ClassNameDateHighlighted & """><a class=""" & ClassNameDateHighlighted & """ onClick=""GetAjax('calendarResponder','" & dayLink & "','','eventContainer','',''); return false;"" href=""#"">" & CurrentDay & "</a></td>"
                Else
                    If HightLight Then
                        Stream = Stream & "<td class=""" & ClassNameDateWithEvent & """><a class=""" & ClassNameDateWithEvent & """ onClick=""GetAjax('calendarResponder','" & dayLink & "','','eventContainer','',''); return false;"" href=""#"">" & CurrentDay & "</a></td>"
                    Else
                        Stream = Stream & "<td class=""" & ClassNameDateLink & """><a class=""" & ClassNameDateLink & """ onClick=""GetAjax('calendarResponder','" & dayLink & "','','eventContainer','',''); return false;"" href=""#"">" & CurrentDay & "</a></td>"
                    End If
                End If

                If CurrentPosition = 7 Then
                    Stream = Stream & "</tr>"
                    CurrentPosition = 0
                    End If
                    
                CurrentDay = CurrentDay + 1
                CurrentPosition = CurrentPosition + 1
            Loop
            If CurrentPosition <> 1 Then
                Do While CurrentPosition <= 7
                    Stream = Stream & "<td class=""" & ClassNameDateRowFiller & """>&nbsp;</td>"
                    CurrentPosition = CurrentPosition + 1
                Loop
                Stream = Stream & "</tr>"
                End If
            
            Stream = Stream & "<tr>"
            Stream = Stream & "<td class=""" & ClassNameFooter & """ colspan=""7""><img src=""/ccLib/images/spacer.gif"" width=""100%"" height=""10""></td>"
            Stream = Stream & "</tr>"
            Stream = Stream & "</td>"
            Stream = Stream & "</tr>"
            Stream = Stream & "</table>"

            If Stream <> "" Then
                Call Main.SaveBake(CalendarEventsBakeName, Stream, "Calendar Events", (CDate(Date) + 1))
            End If
    End If
    
    GetCalendar = Stream
    
    Exit Function
ErrorTrap:
    Call HandleError("calendarClass", "GetCalendar", Err.Number, Err.Source, Err.Description, True, False)
    End Function

Private Function getDateString(anyDateWithinTargetMonth As Date, calendarName As String) As String
    On Error GoTo ErrorTrap
    
    Dim Stream As String
    Dim CSPointer As Long
    Dim targetMonth As Long
    Dim targetYear As Long
    Dim SQL As String
    Dim StartDay As Date
    Dim EndDay As Date
    Dim EventDays As String
    Dim DaysCount As Long
    Dim CurrentCount As Long
    Dim DaysOfMonth2 As String
    Dim stream2 As String
    
    CurrentCount = 1

    targetMonth = Month(anyDateWithinTargetMonth)
    targetYear = Year(anyDateWithinTargetMonth)
    '
    If Main.GetDataSourceType("") = DataSourceTypeODBCAccess Then
        SQL = " ,datediff('d', Startdate, Enddate) as count"
    Else
        SQL = " ,(select datediff(day, Startdate, Enddate)) as count"
    End If
    SQL = "select" _
        & " distinct e.name As Name" _
        & " ,e.StartDate" _
        & " ,e.EndDate" _
        & SQL _
        & " From ((ccCalendars c" _
        & " left join ccCalendarEventRules r on r.calendarId=c.id)" _
        & " left join ccCalendarEvents e on r.calendarEventId=e.id)" _
        & " Where" _
        & " (c.name=" & KmaEncodeSQLText(calendarName) & ")" _
        & " and(" _
            & " ((month(e.enddate)=" & targetMonth & ")and(year(e.enddate)=" & targetYear & "))" _
            & " or((month(e.startdate)=" & targetMonth & ")and(year(e.startdate)=" & targetYear & "))" _
        & " )" _
        & " and (e.Active <>0)"
    '
    If (Keywords <> "") And (Keywords <> "Search By Keywords") Then
        SQL = SQL & " And (cccalendarevents.name like '%" & Keywords & "%') "
    End If
    '
    SQL = SQL & " Order BY e.StartDate ASC"
    Call Main.TestPoint("CalendarSQL1=" & SQL)
    
    CSPointer = Main.OpenCSSQL("default", SQL)
    If Main.IsCSOK(CSPointer) Then
        Do While Main.IsCSOK(CSPointer)
            StartDay = Main.GetCSDate(CSPointer, "StartDate")
            EndDay = Main.GetCSDate(CSPointer, "EndDate")
            DaysCount = (Main.GetCSInteger(CSPointer, "Count")) + 1
            DaysOfMonth2 = GetDaysInMonth(targetMonth, targetYear)
            Do While CurrentCount <= DaysCount
                If (Month(StartDay) <> targetMonth) Then
                    stream2 = stream2 & "," & Day(StartDay) & ","
                    StartDay = StartDay + 1
                    CurrentCount = CurrentCount + 1
                Else
                    Stream = Stream & "," & Day(StartDay) & ","
                    StartDay = StartDay + 1
                    CurrentCount = CurrentCount + 1
                End If
            Loop
            CurrentCount = 1
            Call Main.NextCSRecord(CSPointer)
        Loop
        Call Main.TestPoint("days with events:" & Stream)
    End If
    
    Call Main.CloseCS(CSPointer)
    
    getDateString = Stream
    
    Exit Function
ErrorTrap:
    Call HandleError("SiteSearchClass", "getDateString", Err.Number, Err.Source, Err.Description, True, False)
    End Function

Private Function GetDaysInMonth(iMonth, iYear) As String
    On Error GoTo ErrorTrap
    
    Dim dTemp
    dTemp = DateAdd("d", -1, DateSerial(iYear, iMonth + 1, 1))
    GetDaysInMonth = Day(dTemp)
    
    Exit Function
ErrorTrap:
    Call HandleError("calendarClass", "GetDaysInMonth", Err.Number, Err.Source, Err.Description, True, False)
    End Function

Function GetWeekdayMonthStartsOn(dAnyDayInTheMonth) As String
    On Error GoTo ErrorTrap
    
    Dim dTemp
    
    dTemp = DateAdd("d", -(Day(dAnyDayInTheMonth) - 1), dAnyDayInTheMonth)
    GetWeekdayMonthStartsOn = Weekday(dTemp)
    
    Exit Function
ErrorTrap:
    Call HandleError("calendarClass", "GetWeekdayMonthStartsOn", Err.Number, Err.Source, Err.Description, True, False)
    End Function

Private Function GetSearch(GivenGroupName As String) As String
    On Error GoTo ErrorTrap
    
    Dim Stream As String
    
    Stream = "<form action=""javascript:convertPost(document.getElementById('myform'));"" name=""myform"" id=""myform"">"
    Stream = Stream & "<table width=""100%"" padding=""5"" cellspacing=""0"" border=""0"">"
    Stream = Stream & "<tr>"
    Stream = Stream & "<td>" & Main.GetFormInputText(RequestNameKeywords, , , 25) & "</td>"
    Stream = Stream & "</tr>"
    Stream = Stream & "<tr>"
    Stream = Stream & "<td>" & getMonthSelect & " " & getYearSelect(GivenGroupName) & "</td>"
    Stream = Stream & "</tr>"
    Stream = Stream & "<tr>"
    Stream = Stream & "<td>" & Main.GetFormInputHidden(RequestNameGroupName, GivenGroupName) & Main.GetFormInputHidden(RequestNameFormID, 300) & Main.GetFormButton("Search") & "</td>"
    Stream = Stream & "</tr>"
    Stream = Stream & "<tr>"
    Stream = Stream & "<td><a onClick=""GetAjax('calendarResponder','" & RequestNameGroupName & "=" & GivenGroupName & "&" & RequestNameFormID & "=400','','calWrapper','',''); return false;"" href=""#"">Add an Event</a></td>"
    Stream = Stream & "</tr>"
    Stream = Stream & "</table>"
    Stream = Stream & "</form>"
    
    GetSearch = Stream
    
    Exit Function
ErrorTrap:
    Call HandleError("calendarClass", "GetSearch", Err.Number, Err.Source, Err.Description, True, False)
    End Function

Private Function getMonthSelect() As String
    On Error GoTo ErrorTrap
    
    Dim Stream As String
    
    Stream = "<select name=""" & RequestNameMonth & """>"
    Stream = Stream & "<option selected value="""">Month</option>"
    Stream = Stream & "<option value=""1"">January</option>"
    Stream = Stream & "<option value=""2"">Febuary</option>"
    Stream = Stream & "<option value=""3"">March</option>"
    Stream = Stream & "<option value=""4"">April</option>"
    Stream = Stream & "<option value=""5"">May</option>"
    Stream = Stream & "<option value=""6"">June</option>"
    Stream = Stream & "<option value=""7"">July</option>"
    Stream = Stream & "<option value=""8"">August</option>"
    Stream = Stream & "<option value=""9"">September</option>"
    Stream = Stream & "<option value=""10"">October</option>"
    Stream = Stream & "<option value=""11"">November</option>"
    Stream = Stream & "<option value=""12"">December</option>"
    Stream = Stream & "</select>"
    
    getMonthSelect = Stream
    
    Exit Function
ErrorTrap:
    Call HandleError("calendarClass", "getMonthSelect", Err.Number, Err.Source, Err.Description, True, False)
    End Function

Private Function getYearSelect(GivenGroupName As String) As String
    On Error GoTo ErrorTrap
    
    Dim Stream As String
    Dim SQL As String
    Dim CS As Long
    Dim thisYear As String
    
    SQL = "SELECT DISTINCT Year(E.StartDate) as Year"
    SQL = SQL & " FROM ccCalendars C, ccCalendarEventRules R, ccCalendarEvents E"
    SQL = SQL & " WHERE (C.Name='" & GivenGroupName & " Calendar')"
    SQL = SQL & " AND (C.ID=R.CalendarID)"
    SQL = SQL & " AND (R.CalendarEventID=E.ID)"
    
    CS = Main.OpenCSSQL("Default", SQL)
    If Main.CSOK(CS) Then
        Stream = "<select name=""" & RequestNameYear & """>"
        Stream = Stream & "<option selected>Year</option>"
        Do While Main.CSOK(CS)
            thisYear = Main.GetCS(CS, "Year")
            Stream = Stream & "<option value=""" & thisYear & """>" & thisYear & "</option>"
            Call Main.NextCSRecord(CS)
        Loop
        Stream = Stream & "</select>"
    End If
    Call Main.CloseCS(CS)
    
    getYearSelect = Stream
    
    Exit Function
ErrorTrap:
    Call HandleError("calendarClass", "getYearSelect", Err.Number, Err.Source, Err.Description, True, False)
    End Function
