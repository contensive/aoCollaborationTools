VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "issueClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'
Private ErrorMessage As String
Private FormID As Long
Private SourceFormID As Long
Private IssueID As Long
Private GroupID As Long

Private ErrorString As String
Private WorkingQueryString As String
Private ButtonValue As String
Private GroupName As String
Private ReturnLink As String

Private IssueName As String
Private Summary As String
Private IssueTypeID As Long
Private SegmentID As Long
Private StatusID As Long
Private DueDate As String
Private OwnerID As Long
Private MemberID As Long
Private Comments As String

Private StyleStatusRed As String
Private StyleStatusYellow As String
Private StyleStatusGreen As String
Private StyleStatusBlue As String

Private SortField As String
Private SortDirection As String
Private filterCompleted As Boolean

Private Main As Object
Private CSV As Object

Public Function Execute(CsvObject As Object, MainObject As Object, OptionString As String, FilterInput As String) As String
    Set CSV = CsvObject
    
    Call Init(MainObject)
    
    Execute = GetContent(OptionString)
    
End Function

Private Sub Init(MainObject As Object)
    Set Main = MainObject
   
    Exit Sub
ErrorTrap:
    Call HandleError("issueClass", "Init", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Function GetContent(OptionString As String) As String
    On Error GoTo ErrorTrap
    
    If Not (Main Is Nothing) Then
    
     FormID = Main.GetStreamInteger(RequestNameFormID)
     SourceFormID = Main.GetStreamInteger(RequestNameSourceFormID)
     ButtonValue = Main.GetStreamText("buttonValue")
     WorkingQueryString = Main.RefreshQueryString
     IssueID = Main.GetStreamInteger(RequestNameIssueID)
     
     SortField = Main.GetStreamText(RequestNameSortField)
     SortDirection = Main.GetStreamText(RequestNameSortDirection)
     
     filterCompleted = Main.GetStreamBoolean(RequestNamefilterCompleted)
     
     If WorkingQueryString = "" Then
         WorkingQueryString = "?"
     Else
         WorkingQueryString = "?" & WorkingQueryString & "&"
     End If
     
     ReturnLink = Replace(Break & Break & "<a href=""" & Main.ServerPage & "?" & WorkingQueryString & """>" & Main.GetSiteProperty("Caption - Committee Return Link", "Return to Collaboration Space", True) & "</a>", "?&", "?", 1, 99, 1)
    
     StyleStatusRed = "STYLE=""font-weight: bold; BACKGROUND-COLOR: " & Main.GetSiteProperty("StatusType2 - Color", "#D13D3D") & "; COLOR: #FFFFFF;"""
     StyleStatusYellow = "STYLE=""font-weight: bold; BACKGROUND-COLOR: " & Main.GetSiteProperty("StatusType3 - Color", "#E5ED00") & ";"""
     StyleStatusGreen = "STYLE=""font-weight: bold; BACKGROUND-COLOR: " & Main.GetSiteProperty("StatusType4 - Color", "#6BC900") & ";"""
     StyleStatusBlue = "STYLE=""font-weight: bold; BACKGROUND-COLOR: " & Main.GetSiteProperty("StatusType5 - Color", "#403F9C") & "; COLOR: #FFFFFF;"""
     
     Call LoadForm
     Call ProcessForm
         
        GroupName = Main.GetAddonOption("Group", OptionString)
        GetContent = GetForm(GroupName)
    End If
    
    Exit Function
ErrorTrap:
    Call HandleError("issueClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetForm(GroupName As String) As String
    On Error GoTo ErrorTrap
    
    Dim Stream As String
    
    Select Case FormID
        Case FormEdit
            Stream = Stream & GetIssueLogEdit()
        Case Else
            Stream = Stream & GetIssueLog(GroupName)
    End Select
    
    GetForm = "<div id=""logWrapper"">" & Stream & "</div>"
    
    Exit Function
ErrorTrap:
    Call HandleError("issueClass", "GetForm", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetIssueLog(GroupName As String) As String
    On Error GoTo ErrorTrap
    
    Dim Stream As String
    Dim CSPointer As Long
    Dim LocalGroupID As Long
    Dim Class As String
    Dim RowNumber As Long
    Dim status As String
    Dim StyleStatus As String
    Dim EditLink As String
    Dim AddLink As String
    Dim Sort As String
    Dim AdditionalCriteria As String
    Dim ButtonString As String
    Dim Comments As String
    Dim CS As Long
    
    LocalGroupID = Main.GetGroupID(GroupName)
    
    If filterCompleted Then
        AdditionalCriteria = " AND (Completed=0)"
    End If
    
    Sort = SortField
    If Sort <> "" Then
        Sort = Sort & " " & SortDirection
    End If
    
    RowNumber = 1
    
    ButtonString = Main.GetFormInputHidden(RequestNameGroupID, LocalGroupID)
    ButtonString = ButtonString & Main.GetFormInputHidden(RequestNameGroupName, GroupName)
    ButtonString = ButtonString & Main.GetFormInputHidden("buttonValue", "", "buttonValue")
    ButtonString = ButtonString & Main.GetFormInputHidden(RequestNameSourceFormID, FormList)
    ButtonString = ButtonString & getButton(FormButtonAddActionItem)
    ButtonString = ButtonString & Spacer & getButton(FormButtonRefresh)
    If filterCompleted Then
        ButtonString = ButtonString & Spacer & getButton(FormButtonFilterOff)
    Else
        ButtonString = ButtonString & Spacer & getButton(FormButtonFilterOn)
    End If
    
    Stream = Stream & "<form action=""javascript:postLogMain(document.getElementById('logMain'));"" name=""logMain"" id=""logMain"">"
    Stream = Stream & "<TABLE class=""IssueTable"" CELLSPACING=""0"" CELLPADDING=""5"">"
    
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD width=100% colspan=10>" & ButtonString & "</TD>"
    Stream = Stream & "</TR>"
    
    CSPointer = Main.OpenCSContent(ContentNameIssuesLog, "(GroupID=" & Main.EncodeSQLNumber(LocalGroupID) & ")" & AdditionalCriteria, Sort)
    If Main.CSOK(CSPointer) Then
        Stream = Stream & "<TR Class=""" & ClassPanelHeader & """>"
        Stream = Stream & "<TD><nobr><span class=""" & ClassPanelHeader & """>Ref.</span></nobr></TD>"
        Stream = Stream & "<TD width=10%><nobr><a class=""" & ClassPanelHeader & """ onClick=""GetAjax('logResponder','" & GetSortableLink("name") & "','','logWrapper','',''); return false;"" href=""#"">Item Name</a></nobr></TD>"
        Stream = Stream & "<TD width=20%><nobr><span class=""" & ClassPanelHeader & """>Action Summary</span></nobr></TD>"
        Stream = Stream & "<TD width=10%><nobr><a class=""" & ClassPanelHeader & """ onClick=""GetAjax('logResponder','" & GetSortableLink("typeid") & "','','logWrapper','',''); return false;"" href=""#""><span class=""" & ClassPanelHeader & """>Type</span></a></nobr></TD>"
        Stream = Stream & "<TD width=10%><nobr><a class=""" & ClassPanelHeader & """ onClick=""GetAjax('logResponder','" & GetSortableLink("Statusid") & "','','logWrapper','',''); return false;"" href=""#"">Status</a></nobr></TD>"
        Stream = Stream & "<TD width=10%><nobr><a class=""" & ClassPanelHeader & """ onClick=""GetAjax('logResponder','" & GetSortableLink("duedate") & "','','logWrapper','',''); return false;"" href=""#"">Due Date</a></nobr></TD>"
        Stream = Stream & "<TD width=10%><nobr><a class=""" & ClassPanelHeader & """ onClick=""GetAjax('logResponder','" & GetSortableLink("ownerlastname") & "','','logWrapper','',''); return false;"" href=""#"">Owner</a></nobr></TD>"
        Stream = Stream & "<TD width=10%><nobr><a class=""" & ClassPanelHeader & """ onClick=""GetAjax('logResponder','" & GetSortableLink("memberlastname") & "','','logWrapper','',''); return false;"" href=""#"">Originator</a></nobr></TD>"
        Stream = Stream & "<TD width=20%><nobr><span class=""" & ClassPanelHeader & """>Comments</span></TD>"
        Stream = Stream & "</TR>"
        
        Do While Main.CSOK(CSPointer)
            
            If (RowNumber Mod 2) = 0 Then
                Class = ClassPanelOdd
            Else
                Class = ClassPanelEven
            End If
            
            status = Main.GetCSText(CSPointer, "Status")
            StatusID = Main.GetCSInteger(CSPointer, "StatusID")
            Comments = Main.GetCSText(CSPointer, "Comments")
            If Comments = "" Then
                Comments = "&nbsp;"
            End If
            StyleStatus = ""
            CS = Main.OpenCSContent(ContentNameStatusTypes, "ID=" & StatusID)
            If Main.CSOK(CS) Then
                StyleStatus = "STYLE=""font-weight: bold; BACKGROUND-COLOR:" & Main.GetCS(CS, "BackgroundColor") & "; COLOR:" & Main.GetCS(CS, "TextColor") & ";"""
            End If
            Call Main.CloseCS(CS)
            
            EditLink = RequestNameFormID & "=" & FormEdit & "&" & RequestNameIssueID & "=" & Main.GetCSInteger(CSPointer, "ID") & "&" & RequestNameGroupName & "=" & GroupName
            
            Stream = Stream & "<TR>"
            Stream = Stream & "<TD class=""" & Class & """ valign=top>" & Main.GetCSInteger(CSPointer, "ID") & "&nbsp;</TD>"
            Stream = Stream & "<TD class=""" & Class & """ valign=top><a onClick=""GetAjax('logResponder','" & EditLink & "','','logWrapper','',''); return false;""  href=""#"">" & Main.GetCSText(CSPointer, "Name") & "</a>&nbsp;</TD>"
            Stream = Stream & "<TD class=""" & Class & """ valign=top>" & Main.GetCSText(CSPointer, "Summary") & "&nbsp;</TD>"
            Stream = Stream & "<TD class=""" & Class & """ valign=top>" & Main.GetCSLookup(CSPointer, "TypeID") & "&nbsp;</TD>"
            Stream = Stream & "<TD class=""" & Class & """ " & StyleStatus & " valign=top>" & status & "&nbsp;</TD>"
            Stream = Stream & "<TD class=""" & Class & """ valign=top>" & Main.GetCSText(CSPointer, "DueDate") & "&nbsp;</TD>"
            Stream = Stream & "<TD class=""" & Class & """ valign=top>" & Main.GetCSText(CSPointer, "OwnerLastName") & "&nbsp;</TD>"
            Stream = Stream & "<TD class=""" & Class & """ valign=top>" & Main.GetCSText(CSPointer, "MemberLastName") & "&nbsp;</TD>"
            Stream = Stream & "<TD class=""" & Class & """ valign=top>" & Comments & "</TD>"
            Stream = Stream & "</TR>"
            Call Main.NextCSRecord(CSPointer)
            RowNumber = RowNumber + 1
        Loop
        
        End If
    
    AddLink = Main.ServerPage & WorkingQueryString & RequestNameFormID & "=" & FormEdit & "&" & RequestNameIssueID & "=0"
    
    Stream = Stream & "</TABLE>"
    Stream = Stream & Main.GetFormEnd()
    
    Call Main.CloseCS(CSPointer)
    
    If (Not Main.IsGroupMember(GroupName)) And (Not Main.IsAdmin()) Then
        Stream = Main.GetLoginPage
    End If
    
    GetIssueLog = Stream
    
    Exit Function
ErrorTrap:
    Call HandleError("issueClass", "GetIssueLog", Err.Number, Err.Source, Err.Description, True, False)
    End Function

Private Function GetIssueLogEdit() As String
    On Error GoTo ErrorTrap
    
    Dim Stream As String
    Dim CSPointer As Long
    Dim ButtonString As String
    Dim LocalMemberID As Long
    Dim LocalGroupID As Long
    Dim AdminLink As String
    Dim AdminLink2 As String
    
    LocalGroupID = Main.GetGroupID(GroupName)
    
    If Main.IsAdmin() Then
        AdminLink = "&nbsp;<a target=""_blank"" href=""" & Main.SiteProperty_AdminURL & "?cid=" & Main.GetContentID(ContentNameIssueTypes) & """>[Admin: Manage Issue Types]</a>"
        AdminLink2 = "&nbsp;<a target=""_blank"" href=""" & Main.SiteProperty_AdminURL & "?cid=" & Main.GetContentID(ContentNameStatusTypes) & """>[Admin: Manage Status Types]</a>"
    End If
    
    If IssueID = 0 Then
        LocalMemberID = Main.MemberID
    Else
        LocalMemberID = MemberID
    End If
    
    Stream = Stream & "<form action=""javascript:postLogEdit(document.getElementById('logEdit'));"" name=""logEdit"" id=""logEdit"">"
    Stream = Stream & "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=5>"
    
    Stream = Stream & GetPageHead("Issue Log Edit")
    Stream = Stream & GetFormRow("Name:", Main.GetFormInputText(RequestNameIssueName, IssueName, , 25))
    Stream = Stream & GetFormRow("Summary:", Main.GetFormInputText(RequestNameSummary, Summary, 10, 25))
    Stream = Stream & GetFormRow("Type:", Main.GetFormInputSelect(RequestNameIssueTypeID, IssueTypeID, ContentNameIssueTypes, , "Select One") & AdminLink)
    Stream = Stream & GetFormRow("Status:", Main.GetFormInputSelect(RequestNameStatusID, StatusID, ContentNameStatusTypes, "", "Select One") & AdminLink2)
    Stream = Stream & GetFormRow("Due Date:", Main.GetFormInputText(RequestNameDueDate, DueDate, , 25))
    Stream = Stream & GetFormRow("Owner:", Main.GetFormInputSelect(RequestNameOwnerID, OwnerID, "Members", , "Select One"))
    Stream = Stream & GetFormRow("Originator:", Main.GetFormInputSelect(RequestNameMemberID, MemberID, "Members", , "Select One"))
    Stream = Stream & GetFormRow("Comments:", Main.GetFormInputText(RequestNameComments, Comments, 10, 25))
    
    ButtonString = Main.GetFormInputHidden(RequestNameGroupID, LocalGroupID)
    ButtonString = ButtonString & Main.GetFormInputHidden("buttonValue", "", "buttonValue")
    ButtonString = ButtonString & Main.GetFormInputHidden(RequestNameIssueID, IssueID)
    ButtonString = ButtonString & Main.GetFormInputHidden(RequestNameSourceFormID, FormID)
    ButtonString = ButtonString & Main.GetFormInputHidden(RequestNameGroupName, GroupName)
    ButtonString = ButtonString & getButton(FormButtonSave) & Spacer
    ButtonString = ButtonString & getButton(FormButtonOK) & Spacer
    ButtonString = ButtonString & getButton(FormButtonCancel) & Spacer
    ButtonString = ButtonString & getButton(FormButtonDelete) & Spacer
    ButtonString = ButtonString & getButton(FormButtonMarkComplete) & Spacer
    
    Stream = Stream & GetFormRow("<img src=""/ccLib/images/spacer.gif"" height=1 width=150>", ButtonString)
    
    Stream = Stream & "</TABLE>"
    Stream = Stream & Main.GetFormEnd()
    
    GetIssueLogEdit = Stream
    
    Exit Function
ErrorTrap:
    Call HandleError("issueClass", "GetIssueLogEdit", Err.Number, Err.Source, Err.Description, True, False)
    End Function

Private Function GetFormRow(FieldCaption As String, Innards As String) As String
    On Error GoTo ErrorTrap
    
    Dim Stream As String
    
    If FieldCaption <> "" Then
        FieldCaption = FieldCaption
    End If
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD valign=top align=right>"
    Stream = Stream & "<img src=""/ccLib/images/spacer.gif"" width=10 height=3>" & Break
    Stream = Stream & FieldCaption
    Stream = Stream & "</TD>"
    Stream = Stream & "<TD width=""100%"">" & Innards & "</TD>"
    Stream = Stream & "</TR>"
    
    GetFormRow = Stream
    
    Exit Function
ErrorTrap:
    Call HandleError("issueClass", "GetFormRow", Err.Number, Err.Source, Err.Description, True, False)
    End Function

Private Sub LoadForm()
    On Error GoTo ErrorTrap
    
    Dim CSPointer As Long
    
    If SourceFormID = FormEdit Then
        
        GroupID = Main.GetStreamInteger(RequestNameGroupID)
        IssueName = Main.GetStreamText(RequestNameIssueName)
        Summary = Main.GetStreamText(RequestNameSummary)
        IssueTypeID = Main.GetStreamInteger(RequestNameIssueTypeID)
        SegmentID = Main.GetStreamInteger(RequestNameSegmentID)
        StatusID = Main.GetStreamInteger(RequestNameStatusID)
        DueDate = Main.GetStreamText(RequestNameDueDate)
        OwnerID = Main.GetStreamInteger(RequestNameOwnerID)
        MemberID = Main.GetStreamInteger(RequestNameMemberID)
        Comments = Main.GetStreamText(RequestNameComments)
        
    Else
        
        CSPointer = Main.OpenCSContent(ContentNameIssuesLog, "ID=" & Main.EncodeSQLNumber(IssueID))
        If Main.CSOK(CSPointer) Then
            GroupID = Main.GetCSInteger(CSPointer, "GroupID")
            IssueName = Main.GetCSText(CSPointer, "Name")
            Summary = Main.GetCSText(CSPointer, "Summary")
            IssueTypeID = Main.GetCSInteger(CSPointer, "TypeID")
            StatusID = Main.GetCSInteger(CSPointer, "StatusID")
            DueDate = Main.GetCSText(CSPointer, "DueDate")
            OwnerID = Main.GetCSInteger(CSPointer, "OwnerID")
            MemberID = Main.GetCSInteger(CSPointer, "MemberID")
            Comments = Main.GetCSText(CSPointer, "Comments")
        End If
        Call Main.CloseCS(CSPointer)
        
    End If
    
    Exit Sub
ErrorTrap:
    Call HandleError("issueClass", "LoadForm", Err.Number, Err.Source, Err.Description, True, False)
    End Sub

Private Function GetMemberLastName(GivenMemberID As Long) As String
    On Error GoTo ErrorTrap
    
    Dim Stream As String
    Dim CSPointer As String
    
    CSPointer = Main.OpenCSContent(ContentNamePeople, "ID=" & Main.EncodeSQLNumber(GivenMemberID))
    If Main.CSOK(CSPointer) Then
        Stream = Stream & Main.GetCSText(CSPointer, "LastName")
    End If
    Call Main.CloseCS(CSPointer)
    
    GetMemberLastName = Stream
    
    Exit Function
ErrorTrap:
    Call HandleError("issueClass", "GetMemberLastName", Err.Number, Err.Source, Err.Description, True, False)
    End Function

Private Sub ProcessForm()
    On Error GoTo ErrorTrap
    
    Select Case SourceFormID
        Case FormEdit
            Call ProcessFormEdit
        Case FormList
            Call ProcessFormList
    End Select
    
    Exit Sub
ErrorTrap:
    Call HandleError("issueClass", "LoadForm", Err.Number, Err.Source, Err.Description, True, False)
    End Sub

Private Sub ProcessFormEdit()
    On Error GoTo ErrorTrap
    
    Select Case ButtonValue
        Case FormButtonOK
            FormID = FormList
            Call SaveIssueLog
        Case FormButtonSave
            FormID = FormEdit
            Call SaveIssueLog
        Case FormButtonCancel
            FormID = FormList
        Case FormButtonDelete
            FormID = FormList
            Call DeleteIssueLog
        Case FormButtonMarkComplete
            FormID = FormList
            Call SaveIssueLog
            Call MarkCompleteIssueLog
    End Select
    
    Exit Sub
ErrorTrap:
    Call HandleError("issueClass", "ProcessFormEdit", Err.Number, Err.Source, Err.Description, True, False)
    End Sub

Private Sub ProcessFormList()
    On Error GoTo ErrorTrap
    
    Select Case ButtonValue
        Case FormButtonAddActionItem
            FormID = FormEdit
        Case FormButtonRefresh
            FormID = FormList
        Case FormButtonFilterOn
            filterCompleted = 1
            FormID = FormList
        Case FormButtonFilterOff
            filterCompleted = 0
            FormID = FormList
    End Select
    
    Exit Sub
ErrorTrap:
    Call HandleError("issueClass", "ProcessFormList", Err.Number, Err.Source, Err.Description, True, False)
    End Sub

Private Sub SaveIssueLog()
    On Error GoTo ErrorTrap
    
    Dim CSPointer As Long
    Dim CS As Long
    
    If ErrorString = "" Then
        If IssueID = 0 Then
            CSPointer = Main.InsertCSContent(ContentNameIssuesLog)
        Else
            CSPointer = Main.OpenCSContent(ContentNameIssuesLog, "ID=" & Main.EncodeSQLNumber(IssueID))
        End If
        If Main.CSOK(CSPointer) Then
            If IssueName = "" Then
                IssueName = "Issue"
            End If
            Call Main.SetCS(CSPointer, "Name", IssueName)
            Call Main.SetCS(CSPointer, "Summary", Summary)
            Call Main.SetCS(CSPointer, "TypeID", IssueTypeID)
            Call Main.SetCS(CSPointer, "StatusID", StatusID)
            Call Main.SetCS(CSPointer, "Status", Main.GetRecordName(ContentNameStatusTypes, StatusID))
            Call Main.SetCS(CSPointer, "DueDate", DueDate)
            Call Main.SetCS(CSPointer, "OwnerID", OwnerID)
            Call Main.SetCS(CSPointer, "MemberID", MemberID)
            Call Main.SetCS(CSPointer, "Comments", Comments)
            Call Main.SetCS(CSPointer, "GroupID", GroupID)
            
            If OwnerID <> 0 Then
                CS = Main.OpenCSContent(ContentNamePeople, "ID=" & Main.EncodeSQLNumber(OwnerID))
                If Main.CSOK(CS) Then
                    Call Main.SetCS(CSPointer, "OwnerFirstName", Main.GetCSText(CS, "FirstName"))
                    Call Main.SetCS(CSPointer, "OwnerLastName", Main.GetCSText(CS, "LastName"))
                    End If
                Call Main.CloseCS(CS)
                End If
            
            If MemberID <> 0 Then
                CS = Main.OpenCSContent(ContentNamePeople, "ID=" & Main.EncodeSQLNumber(MemberID))
                If Main.CSOK(CS) Then
                    Call Main.SetCS(CSPointer, "MemberFirstName", Main.GetCSText(CS, "FirstName"))
                    Call Main.SetCS(CSPointer, "MemberLastName", Main.GetCSText(CS, "LastName"))
                End If
                Call Main.CloseCS(CS)
            End If
            
            CS = Main.OpenCSContent(ContentNameStatusTypes, "ID=" & StatusID, , , , , "Completed")
            If Main.CSOK(CS) Then
                Call Main.SetCS(CSPointer, "Completed", Main.GetCSBoolean(CS, "Completed"))
            End If
            Call Main.CloseCS(CS)
            
            Call Main.SaveCSRecord(CSPointer)
        End If
        Call Main.CloseCS(CSPointer)
    End If
    
    Exit Sub
ErrorTrap:
    Call HandleError("issueClass", "SaveIssueLog", Err.Number, Err.Source, Err.Description, True, False)
    End Sub

'MarkCompleteIssueLog
Private Sub MarkCompleteIssueLog()
    On Error GoTo ErrorTrap
    
    If IssueID <> 0 Then
        Call Main.ExecuteSQL("default", "update IssueLogEntries set completed=1 where ID=" & IssueID)
    End If
    
    Exit Sub
ErrorTrap:
    Call HandleError("issueClass", "MarkCompleteIssueLog", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
Private Sub DeleteIssueLog()
    On Error GoTo ErrorTrap
    
    If IssueID <> 0 Then
        Call Main.DeleteContentRecords(ContentNameIssuesLog, "ID=" & IssueID)
    End If
    
    Exit Sub
ErrorTrap:
    Call HandleError("issueClass", "DeleteIssueLog", Err.Number, Err.Source, Err.Description, True, False)
    End Sub

Private Function GetPageHead(PageName As String) As String
    On Error GoTo ErrorTrap
    
    Dim Stream As String
    
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD colspan=2 width=""100%"">"
    
    If ErrorString <> "" Then
        Stream = Stream & "<span class=""ccError"">" & ErrorString & "</span>"
    Else
        Stream = Stream & Main.GetContentCopy(PageName)
        End If
    
    Stream = Stream & "</TD>"
    Stream = Stream & "</TR>"
    
    GetPageHead = Stream
    
    Exit Function
ErrorTrap:
    Call HandleError("issueClass", "GetPageHead", Err.Number, Err.Source, Err.Description, True, False)
    End Function

Private Function GetFieldSortDirection(CurrentSortField As String, CurrentSortDirection As String, SelectedSortField As String) As String
    On Error GoTo ErrorTrap
    
    If CurrentSortField = SelectedSortField Then
        Select Case CurrentSortDirection
            Case "asc"
                GetFieldSortDirection = "desc"
            Case "desc"
                GetFieldSortDirection = "asc"
            Case Else
                GetFieldSortDirection = "asc"
        End Select
    Else
        GetFieldSortDirection = "asc"
    End If
    
    Exit Function
ErrorTrap:
    Call HandleError("issueClass", "GetFieldSortDirection", Err.Number, Err.Source, Err.Description, True, False)
    End Function

Private Function GetSortableLink(FieldName As String) As String
    On Error GoTo ErrorTrap
    
    Dim Stream As String
    
    GetSortableLink = WorkingQueryString & RequestNameSortField & "=" & FieldName & "&" _
                    & RequestNameSortDirection & "=" & GetFieldSortDirection(SortField, SortDirection, FieldName) & "&" _
                    & RequestNamefilterCompleted & "=" & filterCompleted & "&" _
                    & RequestNameGroupName & "=" & GroupName
    
    Exit Function
ErrorTrap:
    Call HandleError("issueClass", "GetSortableLink", Err.Number, Err.Source, Err.Description, True, False)
    End Function

Private Function GetStatus(StatusID As Long) As String
    On Error GoTo ErrorTrap
    
    Select Case StatusID
        Case 2
            GetStatus = Main.GetSiteProperty("StatusType2", "Red")
        Case 3
            GetStatus = Main.GetSiteProperty("StatusType3", "Yellow")
        Case 4
            GetStatus = Main.GetSiteProperty("StatusType4", "Green")
        Case 5
            GetStatus = Main.GetSiteProperty("StatusType5", "Blue (Completed)")
    End Select
    
    Exit Function
ErrorTrap:
    Call HandleError("issueClass", "GetStatus", Err.Number, Err.Source, Err.Description, True, False)
    End Function

Private Function GetStatusSelect(RequestName As String, StatusID As Long) As String
    On Error GoTo ErrorTrap
    
    Dim Stream As String
    
    Stream = "<select size=""1"" name=""" & RequestName & """>"
    Stream = Stream & "<option selected value="""">Select One</option>"
    Stream = Stream & "<option "
    If StatusID = 2 Then
        Stream = Stream & "Selected "
    End If
    Stream = Stream & "value=""2"">" & Main.GetSiteProperty("StatusType2", "Red") & "</option>"
    Stream = Stream & "<option "
    If StatusID = 3 Then
        Stream = Stream & "Selected "
    End If
    Stream = Stream & "value=""3"">" & Main.GetSiteProperty("StatusType3", "Yellow") & "</option>"
    Stream = Stream & "<option "
    If StatusID = 4 Then
        Stream = Stream & "Selected "
    End If
    Stream = Stream & "value=""4"">" & Main.GetSiteProperty("StatusType4", "Green") & "</option>"
    Stream = Stream & "<option "
    If StatusID = 5 Then
        Stream = Stream & "Selected "
    End If
    Stream = Stream & "value=""5"">" & Main.GetSiteProperty("StatusType5", "Blue (Completed)") & "</option>"
    Stream = Stream & "</select>"
    
    GetStatusSelect = Stream
    
    Exit Function
ErrorTrap:
    Call HandleError("issueClass", "GetStatusSelect", Err.Number, Err.Source, Err.Description, True, False)
    End Function

Private Function getButton(ButtonValue As String) As String
    On Error GoTo ErrorTrap
    
    Dim Stream As String
    
    Stream = Main.GetFormButton(ButtonValue, , Trim(ButtonValue))
    Stream = Replace(Stream, "<input ", "<input onClick=""setButtonValue('" & Trim(ButtonValue) & "','buttonValue')"" ")
    
    getButton = Stream
    
    Exit Function
ErrorTrap:
    Call HandleError("issueClass", "getButton", Err.Number, Err.Source, Err.Description, True, False)
    End Function


