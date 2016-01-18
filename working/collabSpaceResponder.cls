VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "collabSpaceResponder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Private FormID As Long
Private SourceFormID As Long
Private CurrentDay As Date
Private GroupName As String

Private Message As String

Private selKeyword As String
Private selMonth As Long
Private selYear As Long

Private Main As Object
Private CSV As Object

Public Function Execute(CsvObject As Object, MainObject As Object, OptionString As String, FilterInput As String) As String
    Set CSV = CsvObject
    
    Call Init(MainObject)
    
    Execute = GetContent(OptionString)
    
End Function

Private Sub Init(MainObject As Object)
    On Error GoTo ErrorTrap
    
    Set Main = MainObject
    
    FormID = Main.GetStreamInteger(RequestNameFormID)
    SourceFormID = Main.GetStreamInteger(RequestNameSourceFormID)
    CurrentDay = Main.GetStreamDate(RequestNameCurrentDay)
    GroupName = Main.GetStreamText(RequestNameGroupName)
    
    selMonth = Main.GetStreamInteger(RequestNameMonth)
    selYear = Main.GetStreamInteger(RequestNameYear)
    selKeyword = Main.GetStreamText(RequestNameKeywords)
    
    Call ProcessForm
    
    Exit Sub

ErrorTrap:
    Call HandleError("collabSpaceResponder", "Init", Err.Number, Err.Source, Err.Description, True, False)
    End Sub

Private Function GetContent(OptionString As String) As String
    On Error GoTo ErrorTrap
    
    Dim Stream As String
    Dim objCal As New calendarClass
    Dim SQL As String
    Dim calendarResponder As New calendarResponder

    Select Case FormID
        Case FormAddEvent
            Stream = getEventForm()
        Case FormCalSearch
        
            SQL = "SELECT E.*"
            SQL = SQL & " FROM ccCalendars C, ccCalendarEventRules R, ccCalendarEvents E"
            SQL = SQL & " WHERE (C.Name=" & Main.EncodeSQLText(GroupName & " Calendar") & ")"
            SQL = SQL & " AND (C.ID=R.CalendarID)"
            SQL = SQL & " AND (R.CalendarEventID=E.ID)"
            
            If selKeyword <> "" Then
                SQL = SQL & " AND (E.Name like '%" & selKeyword & "%')"
            End If
            If selMonth <> 0 Then
                SQL = SQL & " And ((Month(E.StartDate)=" & selMonth & ")OR(Month(E.EndDate)=" & selMonth & "))"
            End If
            If selYear <> 0 Then
                SQL = SQL & " And ((Year(E.StartDate)=" & selYear & ")OR(Year(E.StartDate)=" & selYear & "))"
            End If
            
            SQL = SQL & " ORDER BY StartTime"
            
            Stream = calendarResponder.GetCalendarEvents(SQL, #12:00:00 AM#)
            
        Case FormEvent
        
            SQL = "SELECT E.*"
            SQL = SQL & " FROM ccCalendars C, ccCalendarEventRules R, ccCalendarEvents E"
            SQL = SQL & " WHERE (C.Name=" & Main.EncodeSQLText(GroupName & " Calendar") & ")"
            SQL = SQL & " AND (C.ID=R.CalendarID)"
            SQL = SQL & " AND (R.CalendarEventID=E.ID)"
            If CurrentDay <> CDate(0) Then
                SQL = SQL & " AND ((E.StartDate <=" & Main.EncodeSQLDate(CurrentDay) & ") AND (E.EndDate >=" & Main.EncodeSQLDate(CurrentDay) & "))"
            End If
            SQL = SQL & " ORDER BY StartTime"
            
            Stream = calendarResponder.GetCalendarEvents(SQL, CurrentDay)
            
        Case FormMonth
            Call objCal.Init(Main)
            Stream = objCal.GetCalendar(GroupName, CurrentDay)
        Case Else
            Stream = Stream & Main.ExecuteAddon(0, "{E990B805-2E45-45A3-AD85-8DA22D029A73}", "Group=" & GroupName, ContextPage, "", 0, "", "", 0)
    End Select

    GetContent = Message & Stream
    
    Exit Function
ErrorTrap:
    Call HandleError("collabSpaceResponder", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
'Private Function xGetCalendarEvents(givenSQL As String) As String
'    On Error GoTo ErrorTrap
'
'    Dim Stream As String
'    Dim CS As Long
'    Dim RowNumber As Long
'    Dim RowClass As String
'    Dim dateCopy As String
'    Dim timeCopy As String
'
'
'    CS = Main.OpenCSSQL("Default", givenSQL)
'    If Main.CSOK(CS) Then
'        Stream = "<table width=""100%"" border=""0"" cellpadding=""3"" cellspacing=""0"">"
'            Stream = Stream & "<tr>"
'            Stream = Stream & "<td class=""eventHeader"">Event</td>"
'            Stream = Stream & "<td class=""eventHeader"">Start</td>"
'            Stream = Stream & "<td class=""eventHeader"">End</td>"
'            Stream = Stream & "</tr>"
'        Do While Main.CSOK(CS)
'
'            If (RowNumber Mod 2) = 0 Then
'                RowClass = "eventRowOdd"
'            Else
'                RowClass = "eventRowEven"
'            End If
'
'            Stream = Stream & "<tr>"
'            Stream = Stream & "<td class=""" & RowClass & """>" & Main.GetCSText(CS, "Name") & "</td>"
'            '
'            dateCopy = Main.GetCSText(CS, "StartDate")
'            timeCopy = Main.GetCSText(CS, "StartTime")
'            If timeCopy <> "" Then
'                dateCopy = dateCopy & " - " & timeCopy
'            End If
'            Stream = Stream & "<td class=""" & RowClass & """>" & dateCopy & "</td>"
'            '
'            dateCopy = Main.GetCSText(CS, "EndDate")
'            timeCopy = Main.GetCSText(CS, "EndTime")
'            If timeCopy <> "" Then
'                dateCopy = dateCopy & " - " & timeCopy
'            End If
'            Stream = Stream & "<td class=""" & RowClass & """>" & dateCopy & "</td>"
'            Stream = Stream & "</tr>"
'            RowNumber = RowNumber + 1
'            Call Main.NextCSRecord(CS)
'        Loop
'        Stream = Stream & "</table>"
'    Else
'        Stream = "<p style=""padding:10px;"" class=""ccError"">There are no events scheduled for the selected time.</p>"
'    End If
'    Call Main.CloseCS(CS)
'
'    GetCalendarEvents = Stream
'
'    Exit Function
'ErrorTrap:
'    Call HandleError("collabSpaceResponder", "GetCalendarEvents", Err.Number, Err.Source, Err.Description, True, False)
'    End Function

Private Function getEventForm() As String
    On Error GoTo ErrorTrap
    
    Dim Stream As String
    Dim ButtonString As String
    
    Stream = "<div id=""calWrapper"">"
    Stream = Stream & "<form action=""javascript:convertPost2(document.getElementById('eventAdd'));"" name=""eventAdd"" id=""eventAdd"">"
    Stream = Stream & "<table width=""100%"" padding=""5"" cellspacing=""0"" border=""0"">"
    
    Stream = Stream & GetFormRow("*Event Name", Main.GetFormInputText(RequestNameEvent, "", , 25))
    Stream = Stream & GetFormRow("Start Date", Main.GetFormInputDate(RequestNameStart, Date, 25))
    Stream = Stream & GetFormRow("Start Time", Main.GetFormInputText(RequestNameTimeStart, , , 25))
    Stream = Stream & GetFormRow("End Date", Main.GetFormInputDate(RequestNameEnd, Date, 25))
    Stream = Stream & GetFormRow("End Time", Main.GetFormInputText(RequestNameTimeEnd, , , 25))
    Stream = Stream & GetFormRow("Description", Main.GetFormInputText(RequestNameDescription, "", 5, 30))
    Stream = Stream & GetFormRow("Link", Main.GetFormInputText(RequestNameLink, "", , 25))
    
    ButtonString = Main.GetFormInputHidden(RequestNameGroupName, GroupName) _
                    & Main.GetFormInputHidden(RequestNameSourceFormID, FormID) _
                    & Main.GetFormButton(FormButtonAddEvent)
    
    Stream = Stream & "<tr>"
    Stream = Stream & "<td>" & ButtonString & "</td>"
    Stream = Stream & "</tr>"
    Stream = Stream & "</table>"
    Stream = Stream & "</form>"
    Stream = Stream & "</div>"

    getEventForm = Stream
    
    Exit Function
ErrorTrap:
    Call HandleError("collabSpaceResponder", "getEventForm", Err.Number, Err.Source, Err.Description, True, False)
    End Function

Private Sub ProcessForm()
    On Error GoTo ErrorTrap
    
    Select Case SourceFormID
        Case FormAddEvent
            Call ProcessFormAddEvent
    End Select
    
    Exit Sub

ErrorTrap:
    Call HandleError("collabSpaceResponder", "ProcessForm", Err.Number, Err.Source, Err.Description, True, False)
    End Sub

Private Function GetFormRow2(FieldCaption As String, Class As String) As String
    On Error GoTo ErrorTrap
    
    Dim Stream As String
    
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD colspan=""2"" "
    If Class <> "" Then
        Stream = Stream & "class=""" & Class & """"""
    End If
    Stream = Stream & ">" & FieldCaption & "</TD>"
    Stream = Stream & "</TR>"
    
    GetFormRow2 = Stream
    
    Exit Function
ErrorTrap:
    Call HandleError("collabSpaceResponder", "GetFormRow", Err.Number, Err.Source, Err.Description, True, False)
    End Function

Private Function GetFormRow(FieldCaption As String, Innards As String) As String
    On Error GoTo ErrorTrap
    
    Dim Stream As String
    
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD class=""formCaption"">" & FieldCaption & "</TD>"
    Stream = Stream & "<TD class=""formData"">" & Innards & "</TD>"
    Stream = Stream & "</TR>"
    
    GetFormRow = Stream
    
    Exit Function
ErrorTrap:
    Call HandleError("collabSpaceResponder", "GetFormRow", Err.Number, Err.Source, Err.Description, True, False)
    End Function

Private Sub ProcessFormAddEvent()
    
    Dim CS As Long
    Dim RecordID As Long
    
    If Main.GetStreamText("button") = FormButtonAdd Then
    
        CS = Main.InsertCSContent("Calendar Events")
        If Main.CSOK(CS) Then
            RecordID = Main.GetCSInteger(CS, "ID")
            Call Main.SetCS(CS, "Name", Main.GetStreamText(RequestNameEvent))
            Call Main.SetCS(CS, "StartDate", Main.GetStreamDate(RequestNameStart))
            Call Main.SetCS(CS, "StartTime", Main.GetStreamText(RequestNameTimeStart))
            Call Main.SetCS(CS, "EndDate", Main.GetStreamDate(RequestNameEnd))
            Call Main.SetCS(CS, "EndTime", Main.GetStreamText(RequestNameTimeEnd))
            Call Main.SetCS(CS, "EventLink", Main.GetStreamText(RequestNameLink))
            Call Main.SetCS(CS, "CopyFilename", Main.GetStreamText(RequestNameDescription))
            Call Main.SaveCSRecord(CS)
        End If
        Call Main.CloseCS(CS)
        
        CS = Main.InsertCSContent("Calendar Event Rules")
        If Main.CSOK(CS) Then
            Call Main.SetCS(CS, "CalendarEventID", RecordID)
            Call Main.SetCS(CS, "CalendarID", Main.GetRecordID("Calendars", GroupName & " Calendar"))
            Call Main.SaveCSRecord(CS)
        End If
        Call Main.CloseCS(CS)
        
        'Message = Message & "<div class=""ccError"">Your event has been added.</div>"
        
    End If
    
    Exit Sub
ErrorTrap:
    Call HandleError("collabSpaceResponder", "ProcessFormAddEvent", Err.Number, Err.Source, Err.Description, True, False)
End Sub


