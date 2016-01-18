VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CollaborationClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'
Private ErrorMessage As String
Private FormID As Long
Private SourceFormID As Long
Private DocumentID As Long
'
Private ErrorString As String
Private WorkingQueryString As String
Private ButtonValue As String
Private GroupName As String
Private PostedGroupID As Long
Private DocumentFileName As String
Private DocName As String
Private ReturnLink As String
'
'Private Main As ccweb3.MainClass
Private Main As Object
Private CSV As Object
'
'
'
Public Function Execute(CsvObject As Object, MainObject As Object, OptionString As String, FilterInput As String) As String
    Set CSV = CsvObject
    '
    Call Init(MainObject)
    Execute = GetContent(OptionString)
    '
End Function
'
Private Sub Init(MainObject As Object)
    '
    Set Main = MainObject
    '
    FormID = Main.GetStreamInteger(RequestNameFormID)
    SourceFormID = Main.GetStreamInteger(RequestNameSourceFormID)
    DocumentID = Main.GetStreamInteger(RequestNameDocumentID)
    ButtonValue = Main.GetStreamText("Button")
    PostedGroupID = Main.GetStreamInteger(RequestNamePostedGroupID)
    DocumentFileName = Main.GetStreamText(RequestNameDocumentFileName)
    DocName = Main.GetStreamText(RequestNameDocName)
    '
    WorkingQueryString = Main.RefreshQueryString
    If WorkingQueryString = "" Then
        WorkingQueryString = "?"
    Else
        WorkingQueryString = "?" & WorkingQueryString & "&"
    End If
    '
    ReturnLink = "<div class=""CollaborationNav"">" _
        & "<a href=""" & Main.ServerPage & WorkingQueryString & """>" & Main.GetSiteProperty("Caption - Collaboration Return Link", "Return to Collaboration Space", True) _
        & "</a>" _
        & "</div>"
    '
    Call ProcessForm
    '
    'Call ProcessFileCheckout
    '
    Exit Sub
ErrorTrap:
    Call HandleError("CollaborationClass", "Init", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Function GetContent(OptionString As String) As String
    On Error GoTo ErrorTrap
    '
    Dim Copy As String
    '
    If Not (Main Is Nothing) Then
        GroupName = Main.GetAddonOption("GroupName", OptionString)
        If Main.IsGroupMember(GroupName) Or (Main.IsAdmin) Then
            GetContent = GetCollaborationTool(GroupName)
        Else
            Copy = "The content on this page has restricted access. If you have a username and password for this system," _
                & " <a href=""" & Main.ServerPage & WorkingQueryString & RequestNameMethod & "=login"">Click Here</a>." _
                & " For more information, please contact the administrator."
            GetContent = Main.GetContentCopy2("User has no access to collaborators list", , Copy)
            End If
        End If
    '
    Exit Function
ErrorTrap:
    Call HandleError("CollaborationClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetCollaborationTool(GroupName As String) As String
    On Error GoTo ErrorTrap
    '
    Dim materials As materialsClass
    '
    Select Case FormID
        Case FormResolvedConfirmation
            GetCollaborationTool = GetCollaborationTool & Main.GetContentCopy("Confirmation - Document Resolved") & ReturnLink
        Case FormCheckOutConfirmation
            GetCollaborationTool = GetCollaborationTool & Main.GetContentCopy("Confirmation - Document Check-Out") & GetResponseFileLink() & ReturnLink
        Case FormCheckInConfirmation
            GetCollaborationTool = GetCollaborationTool & Main.GetContentCopy("Confirmation - Document Check-In") & ReturnLink
        Case FormDocumentDetails
            GetCollaborationTool = GetCollaborationTool & GetDocumentDetails() & ReturnLink
        Case FormDocumentList
            Set materials = New materialsClass
            GetCollaborationTool = GetCollaborationTool & materials.GetDocumentsList(Main, GroupName) & ReturnLink
        Case FormCollaboratorList
            GetCollaborationTool = GetCollaborationTool & GetFormCollaboratorList(GroupName) & ReturnLink
        Case FormAddDocument
            GetCollaborationTool = GetCollaborationTool & GetFormAddDocument(GroupName) & ReturnLink
        Case Else
            GetCollaborationTool = GetCollaborationTool & GetFormLanding(GroupName)
    End Select
    '
    Exit Function
ErrorTrap:
    Call HandleError("CollaborationClass", "GetDocumentConsortuim", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function oldGetDocumentsList(GroupName As String) As String
    On Error GoTo ErrorTrap
    '
    Dim CSPointer As Long
    Dim Stream As String
    Dim FileCheckedOut As Boolean
    Dim MemberIDCheckedOut As Long
    Dim MemberNameCheckedOut As String
    Dim RowNumber As Long
    Dim BGColor As String
    Dim Resolved As Boolean
    Dim GroupID As Long
    '
    GroupID = Main.GetGroupID(GroupName)
    '
    Stream = Stream & "<div style=""padding-bottom: 10px;"">"
    Stream = Stream & Main.GetContentCopy("Collaboration Document List")
    Stream = Stream & "</div>"
    '
    CSPointer = Main.OpenCSContent(ContentNameDocuments, "GroupID=" & Main.EncodeSQLNumber(GroupID), "DateAdded Desc", , , , "ID, Name, DateAdded, CheckOutMemberID, Resolved, Description")
    If Main.CSOK(CSPointer) Then
        Stream = Stream & "<TABLE class=""" & ClassTable & """ cellspacing=""0"" cellpadding=""5"">"
        Stream = Stream & "<TR class=""" & ClassPanelHeader & """>"
        Stream = Stream & "<TD width=33%><b>Date Added</b></TD>"
        Stream = Stream & "<TD width=33%><b>Name</b></TD>"
        Stream = Stream & "<TD width=33%><b>File Status</b></TD>"
        Stream = Stream & "</TR>"
        Do While Main.CSOK(CSPointer)
            If (RowNumber Mod 2) = 0 Then
                BGColor = ClassPanelEven
            Else
                BGColor = ClassPanelOdd
                End If
            Resolved = Main.GetCSBoolean(CSPointer, "Resolved")
            MemberIDCheckedOut = Main.GetCSInteger(CSPointer, "CheckOutMemberID")
            MemberNameCheckedOut = GetFormattedMember(MemberIDCheckedOut)
            FileCheckedOut = Main.EncodeBoolean(MemberIDCheckedOut)
            Stream = Stream & "<TR class=""" & BGColor & """>"
            Stream = Stream & "<TD width=33%>" & Main.GetCSDate(CSPointer, "DateAdded") & "</TD>"
            Stream = Stream & "<TD width=33%><a href=""" & Main.ServerPage & WorkingQueryString & RequestNameFormID & "=" & FormDocumentDetails & "&" & RequestNameDocumentID & "=" & Main.GetCSInteger(CSPointer, "ID") & """>" & Main.GetCSText(CSPointer, "Name") & "</a></TD>"
            Stream = Stream & "<TD width=33%>"
            If Resolved Then
                Stream = Stream & "Resolved"
            Else
                If FileCheckedOut Then
                    Stream = Stream & "Checked Out by " & MemberNameCheckedOut
                Else
                    Stream = Stream & "Available for editing"
                    End If
                End If
            Stream = Stream & "</TD>"
            Stream = Stream & "</TR>"
            Call Main.NextCSRecord(CSPointer)
            RowNumber = RowNumber + 1
        Loop
        Stream = Stream & "</TABLE>"
    Else
        Stream = Stream & "<span class=""ccError"">There are currently no documents available.</span>"
    End If
    Call Main.CloseCS(CSPointer)
    '
    Stream = Stream & "<div class=""" & ClassNav & """><a href=""" & Main.ServerPage & WorkingQueryString & RequestNameFormID & "=" & FormAddDocument & """>Add a new document</a></div>"
    '
    oldGetDocumentsList = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("CollaborationClass", "oldGetDocumentsList", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetFormAddDocument(GroupName As String) As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim GroupID As Long
    '
    GroupID = Main.GetGroupID(GroupName)
    '
    Stream = Stream & Main.GetUploadFormStart()
    Stream = Stream & "<TABLE WIDTH=300 BORDER=0 CELLSPACING=0 CELLPADDING=5>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD colspan=2>"
    If ErrorMessage = "" Then
        Stream = Stream & Main.GetContentCopy2("Document Collaboration - Add a New Document", "Copy Content", "Use the form below to add a new document.")
    Else
        Stream = Stream & "<SPAN CLass=""ccError"">" & ErrorMessage & "</span>"
        End If
    Stream = Stream & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>Name:</TD>"
    Stream = Stream & "<TD>" & Main.GetFormInputText(RequestNameDocName, DocName, , 25) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD align=right>File:</TD>"
    Stream = Stream & "<TD>" & Main.GetFormInputFile(RequestNameDocumentFileName) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD><img src=""/ccLib/spacer.gif"" width=100 height=1></TD>"
    Stream = Stream & "<TD>" & Main.GetFormInputHidden(RequestNamePostedGroupID, Main.GetGroupID(GroupName)) & Main.GetFormInputHidden(RequestNameSourceFormID, FormID) & Main.GetFormButton(FormButtonUpload) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "</TABLE>"
    Stream = Stream & Main.GetUploadFormEnd()
    '
    GetFormAddDocument = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("CollaborationClass", "GetFormAddDocument", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetFormattedMember(SelectedMemberID As Long) As String
    On Error GoTo ErrorTrap
    '
    Dim CSPointer As Long
    Dim Email As String
    Dim FirstName As String
    Dim LastName As String
    '
    GetFormattedMember = ""
    '
    CSPointer = Main.OpenCSContent("People", "ID=" & Main.EncodeSQLNumber(SelectedMemberID), , , , , "FirstName, LastName, eMail")
    If Main.CSOK(CSPointer) Then
        Email = Main.GetCSText(CSPointer, "eMail")
        FirstName = Main.GetCSText(CSPointer, "FirstName")
        LastName = Main.GetCSText(CSPointer, "LastName")
        GetFormattedMember = "<a href=""mailto:" & Email & """>" & FirstName & " " & LastName & "</a>"
        End If
    Call Main.CloseCS(CSPointer)
    '
    Exit Function
ErrorTrap:
    Call HandleError("CollaborationClass", "GetMemberEMail", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetDocumentDetails() As String
    On Error GoTo ErrorTrap
    '
    Dim CSPointer As Long
    Dim Stream As String
    Dim MemberCheckOutID As Long
    Dim ModifiedDate As String
    Dim Resolved As Boolean
    Dim Description As String
    '
    Call Main.AddRefreshQueryString(RequestNameDocumentID, DocumentID)
    Call Main.AddRefreshQueryString(RequestNameFormID, FormDocumentResponse)
    '
    If ErrorMessage <> "" Then
        Stream = Stream & "<span class=""ccError"">" & ErrorMessage & "</span>"
        End If
    '
    Stream = Stream & Main.GetUploadFormStart()
    CSPointer = Main.OpenCSContent(ContentNameDocuments, "ID=" & Main.EncodeSQLNumber(DocumentID))
    If Main.CSOK(CSPointer) Then
        Resolved = Main.GetCSBoolean(CSPointer, "Resolved")
        MemberCheckOutID = Main.GetCSInteger(CSPointer, "CheckOutMemberID")
        ModifiedDate = Main.GetCS(CSPointer, "ModifiedDate")
        Description = Main.GetCSText(CSPointer, "Description")
        '
        Stream = Stream & "<div>"
        Stream = Stream & "<div class=""" & ClassDocName & """>" & Main.GetCSRecordEditLink(CSPointer) & Main.GetCSText(CSPointer, "Name") & "</div>"
        If Resolved Then
            Stream = Stream & "<b>This file has been resolved</b>" & Break
            End If
        Stream = Stream & "<b>Date Added:</b> " & Main.GetCS(CSPointer, "DateAdded") & Break
        If ModifiedDate <> "" Then
            Stream = Stream & "<b>Last Modified</b>: " & ModifiedDate & Break
            Stream = Stream & "<b>Modified By</b>: " & Main.GetCSLookup(CSPointer, "ModifiedBy")
            End If
        Stream = Stream & "</div>"
        '
        If Description <> "" Then
            Stream = Stream & "<div>" & Description & "</div>"
        End If
        '
        Stream = Stream & "<img src=""/ccLib/images/spacer.gif"" width=""1"" height=""10"">" & Break
        Stream = Stream & "<div><a target=_blank href=""" & Main.ServerFilePath & Main.GetCSText(CSPointer, "DocumentFileName") & """>" & Main.GetSiteProperty("Caption - View Document", "View Document", True) & "</a></div>"
        Stream = Stream & "<img src=""/ccLib/images/spacer.gif"" width=""1"" height=""10"">" & Break
        '
        Stream = Stream & "<div>"
        If Not Resolved Then
            If MemberCheckOutID = 0 Then
                Stream = Stream & Main.GetFormButton(FormButtonCheckOut)
                End If
            If MemberCheckOutID = Main.MemberID Then
                Stream = Stream & Main.GetFormInputFile(RequestNameDocumentFileName)
                Stream = Stream & Break & Main.GetFormButton(FormButtonCheckIn)
                End If
'            If Main.IsAdmin() Then
'                'Stream = Stream & Break & Break & Main.GetFormButton(FormButtonResolve)
'                End If
            End If
        Stream = Stream & "</div>"
    Else
        Stream = "<span class=""ccError"">The document requested is currently unavailable.</span>"
        End If
    Call Main.CloseCS(CSPointer)
    Stream = Stream & Main.GetFormInputHidden(RequestNameSourceFormID, FormID)
    Stream = Stream & Main.GetUploadFormEnd()
    '
    GetDocumentDetails = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("CollaborationClass", "GetMemberEMail", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Sub ProcessDocument()
    On Error GoTo ErrorTrap
    '
    Dim CSPointer As Long
    Dim Filename As String
    '
    If (ButtonValue = FormButtonCheckOut) Or (ButtonValue = FormButtonCheckIn) Or (ButtonValue = FormButtonResolve) Then
        '
        CSPointer = Main.OpenCSContent(ContentNameDocuments, "ID=" & Main.EncodeSQLNumber(DocumentID))
        If Main.CSOK(CSPointer) Then
            '
            If ButtonValue = FormButtonCheckOut Then
                Call Main.SetCS(CSPointer, "CheckOutMemberID", Main.MemberID)
                Call Main.SaveCSRecord(CSPointer)
                FormID = FormCheckOutConfirmation
                End If
            '
            If ButtonValue = FormButtonCheckIn Then
                If DocumentFileName = "" Then
                    ErrorMessage = Main.GetSiteProperty("Message - Response File Required", "Please select a document to check-in", True)
                    End If
                If ErrorMessage = "" Then
                    Filename = Main.GetCSFilename(CSPointer, "DocumentFileName", DocumentFileName, ContentNameDocuments)
                    Filename = Replace(Filename, DocumentFileName, "")
                    Call Main.ProcessFormInputFile(RequestNameDocumentFileName, Filename)
                    Call Main.SetCS(CSPointer, "CheckOutMemberID", 0)
                    Call Main.SaveCSRecord(CSPointer)
                    FormID = FormCheckInConfirmation
                Else
                    FormID = FormDocumentDetails
                    End If
                End If
            '
            If ButtonValue = FormButtonResolve Then
                Call Main.SetCS(CSPointer, "Resolved", 1)
                Call Main.SaveCSRecord(CSPointer)
                FormID = FormResolvedConfirmation
                End If
            '
            End If
        Call Main.CloseCS(CSPointer)
        '
        End If
    '
    Exit Sub
ErrorTrap:
    Call HandleError("CollaborationClass", "ProcessDocument", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Function GetResponseFileLink() As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim CSPointer As Long
    '
    CSPointer = Main.OpenCSContent(ContentNameDocuments, "ID=" & Main.EncodeSQLNumber(DocumentID), , , , , "DocumentFileName, Resolved")
    If Main.CSOK(CSPointer) Then
        If Main.GetCSBoolean(CSPointer, "Resolved") Then
            Stream = Break & Break & "<a target=_blank href=""" & Main.ServerFilePath & Main.GetCSText(CSPointer, "DocumentFileName") & """>" & Main.GetSiteProperty("Caption - Final Document Laink", "Click here to download the Document", True) & "</a>"
        Else
            Stream = Break & Break & "<a target=_blank href=""" & Main.ServerFilePath & Main.GetCSText(CSPointer, "DocumentFileName") & """>" & Main.GetSiteProperty("Caption - Draft Document Laink", "Click here to download the document", True) & "</a>"
            End If
    Else
        Stream = Main.GetSiteProperty("Message - Document Unavailable", "The document is currently unavailable.", True)
        End If
    Call Main.CloseCS(CSPointer)
    '
    GetResponseFileLink = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("CollaborationClass", "GetResponseFileLink", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetFormLanding(GroupName) As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    '
    Stream = Stream & "<div>" & Main.GetContentCopy(GroupName & " Collaboration Space") & "</div>"
    Stream = Stream & "<div class=""" & ClassNav & """><a href=""" & Main.ServerPage & WorkingQueryString & RequestNameFormID & "=" & FormDocumentList & """>Documents</a></div>"
    Stream = Stream & "<div class=""" & ClassNav & """><a href=""" & Main.ServerPage & WorkingQueryString & RequestNameFormID & "=" & FormCollaboratorList & """>Collaborator Member List</a></div>"
    '
    GetFormLanding = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("CollaborationClass", "GetFormLanding", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetFormCollaboratorList(GroupName As String) As String
    On Error GoTo ErrorTrap
    '
    Dim CSPointer As Long
    Dim Stream As String
    '
    Stream = Stream & Main.GetFormStart()
    '
    Stream = Stream & GetCheckedGoupList(GroupName) & Break
    '
    Stream = Stream & "<TABLE WIDTH=""100%"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""5"">"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD>"
    If ErrorString <> "" Then
        Stream = Stream & "<span class=ccError>" & ErrorString & "</span>"
    Else
        Stream = Stream & Main.GetContentCopy("Collaborator Conatct Email")
        End If
    Stream = Stream & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD>" & Main.GetSiteProperty("Caption - Email Subject", "Subject:", True) & Break
    Stream = Stream & Main.GetFormInputText(RequestNameSubject, , , 46)
    Stream = Stream & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD valign=top>" & Main.GetSiteProperty("Caption - Email Message", "Message:", True) & Break
    Stream = Stream & Main.GetFormInputText(RequestNameBody, , 10, 35)
    Stream = Stream & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD>" & Main.GetFormInputHidden(RequestNamePostedGroupID, Main.GetGroupID(GroupName)) & Main.GetFormInputHidden(RequestNameSourceFormID, FormID) & Main.GetFormButton(FormButtonSend) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD>* Only checked members will receive email.</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "</TABLE>"
    Stream = Stream & Main.GetFormEnd()
    '
    GetFormCollaboratorList = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("CollaborationClass", "GetFormCollaboratorList", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Sub ProcessFormCollaboratorList()
    On Error GoTo ErrorTrap
    '
    Dim Subject As String
    Dim Body As String
    Dim FromAddress As String
    Dim SelectedMemberString As String
    Dim SelectedMemberStringLenghth As Long
    Dim SelectedMemberStringArray() As String
    Dim SelectedMemberStringCount As Long
    Dim SelectedMemberStringPointer As Long
    Dim CSPointer As Long
    Dim SelectedMemberID As Long
    Dim Value As String
    Dim Email As String
    Dim This As String
    '
    GroupName = Main.GetGroupByID(PostedGroupID)
    '
    If Main.GetStreamText("button") = FormButtonSend Then
        '
        Subject = Main.GetStreamText(RequestNameSubject)
        Body = Main.GetStreamText(RequestNameBody)
        FromAddress = Main.MemberEmail
        '
        CSPointer = Main.OpenCSGroupMembers(GroupName, "", "LastName")
        Do While Main.CSOK(CSPointer)
            SelectedMemberID = Main.GetCSInteger(CSPointer, "ID")
            Value = Main.GetStreamBoolean(MemberCheckBoxPrefix & SelectedMemberID)
            This = MemberCheckBoxPrefix & SelectedMemberID
Call Main.TestPoint("Name: " & This)
Call Main.TestPoint("Value: " & Main.GetStreamBoolean(MemberCheckBoxPrefix & SelectedMemberID))
            If Value Then
                SelectedMemberString = SelectedMemberString & SelectedMemberID & ","
                End If
            Call Main.NextCSRecord(CSPointer)
        Loop
        Call Main.CloseCS(CSPointer)
        '
        SelectedMemberStringLenghth = Len(SelectedMemberString)
        If SelectedMemberStringLenghth > 0 Then
            SelectedMemberString = Left(SelectedMemberString, SelectedMemberStringLenghth - 1)
            End If
Call Main.TestPoint("SelectedMemberString: " & SelectedMemberString)
        '
        If Subject = "" Then
            ErrorString = ErrorString & Main.GetSiteProperty("Message - No Email Subject", "Please enter a subject for the email.", True)
            End If
        If Body = "" Then
            ErrorString = ErrorString & Main.GetSiteProperty("Message - No Email Message", "Please enter a message for the email.", True)
            End If
        If FromAddress = "" Then
            ErrorString = ErrorString & Main.GetSiteProperty("Message - No Email Email", "Your email address is not set, please edit your profile to send email using this system.", True)
            End If
        '
        If ErrorString = "" Then
            SelectedMemberStringArray = Split(SelectedMemberString, ",")
            SelectedMemberStringCount = UBound(SelectedMemberStringArray) + 1
            For SelectedMemberStringPointer = 0 To SelectedMemberStringCount - 1
                SelectedMemberID = SelectedMemberStringArray(SelectedMemberStringPointer)
                Email = GetEmail(SelectedMemberID)
                Call Main.SendEmail(Email, FromAddress, Subject, Body, , True, True)
            Next
            ErrorString = Main.GetSiteProperty("Message - Email Sent", "Your eMail has been sent", True)
            End If
        '
        End If
    '
    FormID = FormCollaboratorList
    '
    Exit Sub
ErrorTrap:
    Call HandleError("CollaborationClass", "ProcessFormCollaboratorList", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Function GetGoupList(GroupName As String) As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim CSPointer As Long
    Dim GroupID As Integer
    '
    GroupID = Main.GetGroupID(GroupName)
    '
    CSPointer = Main.OpenCSContent("Member Rules", "GroupID=" & Main.EncodeSQLNumber(GroupID), , , , , "MemberID")
    If Main.CSOK(CSPointer) Then
        Stream = Stream & "<b>" & GroupName & "</b>" & Break
        Do While Main.CSOK(CSPointer)
            Stream = Stream & "<img src=""/ccLib/images/spacer.gif"" height=5 width=150>" & Break
            Stream = Stream & GetFormattedMember(Main.GetCSInteger(CSPointer, "MemberID")) & Break
            Call Main.NextCSRecord(CSPointer)
        Loop
    Else
        Stream = "The " & GroupName & " group is currently unavailable."
        End If
    Call Main.CloseCS(CSPointer)
    '
    GetGoupList = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("CollaborationClass", "GetGoupList", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetCheckedGoupList(GroupName As String) As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim CSPointer As Long
    Dim GroupID As Integer
    Dim SelectedMemberID As Long
    '
    GroupID = Main.GetGroupID(GroupName)
    '
    CSPointer = Main.OpenCSGroupMembers(GroupName, "", "LastName")
    If Main.CSOK(CSPointer) Then
        Stream = Stream & "<b>" & GroupName & "</b>" & Break
        Do While Main.CSOK(CSPointer)
            SelectedMemberID = Main.GetCSInteger(CSPointer, "ID")
            Stream = Stream & "<img src=""/ccLib/images/spacer.gif"" height=5 width=150>" & Break
            Stream = Stream & Main.GetFormInputCheckBox(MemberCheckBoxPrefix & SelectedMemberID) & " " & GetFormattedMember(SelectedMemberID) & Break
            Call Main.NextCSRecord(CSPointer)
        Loop
    Else
        Stream = "<span class=""ccError"">The " & GroupName & " group is currently unavailable or has no members.</span>"
        End If
    Call Main.CloseCS(CSPointer)
    '
    GetCheckedGoupList = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("CollaborationClass", "GetCheckedGoupList", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetEmail(SelectedMemberID As Long) As String
    On Error GoTo ErrorTrap
    '
    Dim CSPointer As Long
    Dim Email As String
    '
    CSPointer = Main.OpenCSContent("People", "ID=" & Main.EncodeSQLNumber(SelectedMemberID), , , , , "FirstName, LastName, eMail")
    If Main.CSOK(CSPointer) Then
        Email = Main.GetCSText(CSPointer, "eMail")
        End If
    Call Main.CloseCS(CSPointer)
    '
    GetEmail = Email
    '
    Exit Function
ErrorTrap:
    Call HandleError("CollaborationClass", "GetEmail", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Sub ProcessForm()
    On Error GoTo ErrorTrap
    '
    Select Case SourceFormID
        Case FormCollaboratorList
            Call ProcessFormCollaboratorList
        Case FormDocumentDetails
            Call ProcessDocument
        Case FormAddDocument
            Call ProcessFormAddDocument
    End Select
    '
    Exit Sub
ErrorTrap:
    Call HandleError("CollaborationClass", "ProcessForm", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Sub ProcessFormAddDocument()
    On Error GoTo ErrorTrap
    '
    Dim CSPointer As Long
    Dim Filename As String
    '
    If ButtonValue = FormButtonUpload Then
        '
        If DocName = "" Then
            ErrorMessage = ErrorMessage & "Please enter a document name.<br />"
            End If
        If DocumentFileName = "" Then
            ErrorMessage = ErrorMessage & "Please select a document to upload.<br />"
            End If
        '
        If ErrorMessage = "" Then
            CSPointer = Main.InsertCSContent(ContentNameDocuments)
            If Main.CSOK(CSPointer) Then
                Filename = Main.GetCSFilename(CSPointer, "DocumentFileName", DocumentFileName, ContentNameDocuments)
                Filename = Replace(Filename, DocumentFileName, "")
                Call Main.ProcessFormInputFile(RequestNameDocumentFileName, Filename)
                Call Main.SetCS(CSPointer, "GroupID", PostedGroupID)
                Call Main.SetCS(CSPointer, "Name", DocName)
                Call Main.SaveCSRecord(CSPointer)
                End If
            Call Main.CloseCS(CSPointer)
            FormID = FormDocumentList
        Else
            FormID = FormAddDocument
            End If
        End If
    '
    Exit Sub
ErrorTrap:
    Call HandleError("CollaborationClass", "ProcessFormAddDocument", Err.Number, Err.Source, Err.Description, True, False)
    End Sub

