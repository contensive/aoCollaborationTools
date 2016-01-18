VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "collabSpace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Private WorkingQueryString As String
Private GroupName As String

Private Main As Object
Private CSV As Object

Public Function Execute(CsvObject As Object, MainObject As Object, OptionString As String, FilterInput As String) As String
    Set CSV = CsvObject
    Set Main = MainObject
    
    WorkingQueryString = Main.RefreshQueryString
    If WorkingQueryString <> "" Then
        WorkingQueryString = "?" & WorkingQueryString & "&"
    Else
        WorkingQueryString = "?"
    End If
    
    Execute = GetContent(OptionString)
    
End Function

Private Function GetContent(OptionString As String) As String
    On Error GoTo ErrorTrap
    
    Dim Stream As String
    Dim calendarName As String
    Dim calendarId As Long
    Dim CS As Long
    Dim instructions As String
    Dim GroupID As Long
    
    GroupName = Trim(Main.GetAddonOption("Group", OptionString))
    If GroupName = "" Then
        GroupName = "Collaboration Space"
    End If
    GroupID = Main.GetRecordID("groups", GroupName)
    If GroupID = 0 Then
        CS = Main.InsertCSContent("groups")
        If Main.IsCSOK(CS) Then
            GroupID = Main.GetCSInteger(CS, "id")
            Call Main.SetCS(CS, "name", GroupName)
            Call Main.SetCS(CS, "caption", GroupName)
        End If
        Call Main.CloseCS(CS)
    End If
    If (Main.IsGroupMember(GroupName)) Or (Main.IsAdmin) Then
        calendarName = GroupName & " Calendar"
        calendarId = Main.GetRecordID("calendars", calendarName)
        If calendarId = 0 Then
            CS = Main.InsertCSContent("calendars")
            If Main.IsCSOK(CS) Then
                Call Main.SetCS(CS, "name", calendarName)
                Call Main.SetCS(CS, "caption", calendarName)
            End If
            Call Main.CloseCS(CS)
        End If
        instructions = Main.GetContentCopy("Collaboration Space Instructions")
        '
        ' instructions row
        '
        If instructions <> "" Then
            Stream = Stream & "<tr>"
            Stream = Stream & "<td style=""width:100%;"" colspan=""2"">" & instructions & "</td>"
            Stream = Stream & "</tr>"
        End If
        '
        ' two column row
        '
        Stream = Stream & "<tr>"
        Stream = Stream & "" _
            & "<td class=""collabLeft"">" _
            & "<div id=""directoryContainer"">" _
                & "<div class=""collabTitle"">Directory</div>" _
                & "<div class=""clearDiv"">&nbsp;</div>" _
                & Main.ExecuteAddon(0, "{E6C06E14-70FA-4C33-BD30-6FC9B845B49E}", "Group=" & GroupName, ContextPage, "", 0, "", "", 0) _
            & "</div>" _
            & "</td>"
        Stream = Stream & "" _
            & "<td class=""collabRight"">" _
            & "<div id=""eventsContainer"">" _
                & "<div class=""collabTitle"">Events</div>" _
                & "<div class=""clearDiv"">&nbsp;</div>" _
                & Main.ExecuteAddon(0, "{E990B805-2E45-45A3-AD85-8DA22D029A73}", "Group=" & GroupName, ContextPage, "", 0, "", "", 0) & "" _
            & "</div>" _
            & "<div id=""materialsContainer"">" _
                & "<div class=""clearDiv"">&nbsp;</div>" _
                & "<div class=""collabTitle"">Documents</div>" _
                & "<div class=""clearDiv"">&nbsp;</div>" _
                & Main.ExecuteAddon(0, groupMaterialsAddon, "Group=" & GroupName, ContextPage, "", 0, "", "", 0) & "" _
            & "</div>" _
            & "<div id=""minutesContainer"">" _
            & "<div class=""clearDiv"">&nbsp;</div>" _
            & "<div class=""collabTitle"">Notes</div>" _
            & "<div class=""clearDiv"">&nbsp;</div>" _
            & Main.GetContentCopy2(GroupName & " Meeting Minutes", , "Notes for " & GroupName) & "" _
            & "</div>" _
            & "</td>"
        Stream = Stream & "</tr>"
        '
        Stream = Stream & "<tr>"
        Stream = Stream & "" _
            & "<td class=""issuesContainer"" valign=""top"" colspan=""2"" width=""100%""><span class=""collabTitle"">Action Item(s)</span>" _
            & "<br />" _
            & "<br />" _
            & Main.ExecuteAddon(0, "{B6A208C8-2AEF-4108-88EB-979992F84CDC}", "Group=" & GroupName, ContextPage, "", 0, "", "", 0) & "" _
            & "</td>"
        Stream = Stream & "</tr>"
        '
        Stream = "" _
            & "<table width=""100%"" cellpadding=""10"" cellspacing=""0"" border=""0"">" _
            & Stream _
            & "</table>"
    Else
        Stream = Main.GetContentCopy2("Collaboration Space - Access Denied", , "The content you have requested has ben restricted. Please log in to access this information")
    End If
    '
    Stream = "" _
        & "<div class=""collabContainer"">" _
        & Stream _
        & "</div>"
    
    GetContent = Stream
    
    Exit Function
ErrorTrap:
    Call HandleError("collabClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function


