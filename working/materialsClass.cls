VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "materialsClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Private ErrorMessage As String
Private FormID As Long
Private SourceFormID As Long
Private DocumentID As Long

Private ErrorString As String
Private WorkingQueryString As String
Private ButtonValue As String
Private GroupName As String
Private PostedGroupID As Long
Private DocumentFileName As String
Private DocName As String
Private ReturnLink As String
'
Private PageID As Long
Private addonId As Long
Private addonGuid As String
'
'
'
Public Function Execute(CSV As Object, Main As Object, optionString As String, FilterInput As String) As String
    On Error GoTo ErrorTrap
    
    'Set CSV = CsvObject
    'Set Main = MainObject
    '
    Dim qs As String
    
    PageID = Main.GetStreamInteger("bid")
    FormID = Main.GetStreamInteger(RequestNameFormID)
    SourceFormID = Main.GetStreamInteger(RequestNameSourceFormID)
    DocumentID = Main.GetStreamInteger(RequestNameDocumentID)
    ButtonValue = Main.GetStreamText("Button")
    PostedGroupID = Main.GetStreamInteger(RequestNamePostedGroupID)
    DocumentFileName = Main.GetStreamText(RequestNameDocumentFileName)
    DocName = Main.GetStreamText(RequestNameDocName)
    GroupName = Main.GetStreamText(RequestNameGroupName)
    If GroupName = "" Then
        GroupName = Main.GetAddonOption("group", optionString)
    End If
    '
    If PageID <> 0 Then
        Call Main.AddRefreshQueryString("bid", PageID)
    Else
        addonId = Main.GetStreamInteger("addonid")
        If addonId <> 0 Then
            Call Main.AddRefreshQueryString("addonid", addonId)
        Else
            addonGuid = Main.GetStreamText("addonGuid")
            If addonGuid <> "" Then
                Call Main.AddRefreshQueryString("addonGuid", addonGuid)
            Else
            End If
        End If
    End If
    '
    WorkingQueryString = Main.RefreshQueryString
    If WorkingQueryString = "" Then
        WorkingQueryString = "?"
    Else
        WorkingQueryString = "?" & WorkingQueryString & "&"
    End If
    
    qs = Main.RefreshQueryString _
        & "&" & RequestNameGroupName & "=" & GroupName _
        & ""
    
    ReturnLink = "<div class=""returnLink""><a onClick=""GetAjax('matResponder','" & qs & "','','matContainer','',''); return false;"" href=""#"">Return to document list</a></div>"
    
    Call ProcessForm(Main)
    Execute = GetContent(Main, optionString)
    
    Exit Function
ErrorTrap:
    Call HandleError("materialsClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
'
'
Private Function GetContent(Main As Object, optionString As String) As String
    On Error GoTo ErrorTrap
    
    Dim Stream As String
    
    GroupName = Main.GetAddonOption("Group", optionString)
    
    If GroupName = "" Then
        GroupName = Main.GetStreamText(RequestNameGroupName)
    End If
    
    Select Case FormID
        Case FormResolvedConfirmation
            Stream = Main.GetContentCopy2("Confirmation - Document Resolved", , "This document has been resolved.") & ReturnLink
        Case FormCheckOutConfirmation
            Stream = Main.GetContentCopy2("Confirmation - Document Check-Out", , "This document has been checked out to you.") & GetResponseFileLink(Main) & ReturnLink
        Case FormCheckInConfirmation
            Stream = Main.GetContentCopy2("Confirmation - Document Check-In", , "Your document has been checked in.") & ReturnLink
        Case FormDocumentDetails
            Stream = GetDocumentDetails(Main) & ReturnLink
        Case FormDocumentList
            Stream = GetDocumentsList(Main, GroupName)
        Case FormCollaboratorList
            Stream = GetFormCollaboratorList(Main, GroupName) & ReturnLink
        Case FormAddDocument
            Stream = GetFormAddDocument(Main, GroupName) & ReturnLink
        Case Else
            Stream = GetDocumentsList(Main, GroupName)
    End Select
    
    GetContent = "<div id=""matContainer"">" & Stream & "</div>"
    
    Exit Function
ErrorTrap:
    Call HandleError("materialsClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
'=========================================================================================
'   Called from other classes, use only local values -- not any Globals from this class
'=========================================================================================
'
Friend Function GetDocumentsList(Main As Object, GroupName As String) As String
    On Error GoTo ErrorTrap
    
    Dim CSPointer As Long
    Dim Stream As String
    Dim FileCheckedOut As Boolean
    Dim MemberIDCheckedOut As Long
    Dim MemberNameCheckedOut As String
    Dim RowNumber As Long
    Dim BGColor As String
    Dim Resolved As Boolean
    Dim GroupID As Long
    Dim docLink As String
    Dim clearCheckout As Boolean
    
    GroupID = Main.GetGroupID(GroupName)

    CSPointer = Main.OpenCSContent(ContentNameDocuments, "GroupID=" & Main.EncodeSQLNumber(GroupID), "DateAdded Desc", , , , "ID, Name, DateAdded, CheckOutMemberID, Resolved, Description")
    If Main.CSOK(CSPointer) Then
        Stream = Stream & "<TABLE class=""" & ClassTable & """ cellspacing=""0"" cellpadding=""5"">"
        Stream = Stream & "<TR class=""" & ClassPanelHeader & """>"
        Stream = Stream & "<TD class=""headerCaption colName"">Document</TD>"
        Stream = Stream & "<TD class=""headerCaption colDateAdded"">Date Added</b></TD>"
        Stream = Stream & "<TD class=""headerCaption colFileStatus"">File Status</TD>"
        Stream = Stream & "</TR>"
        Do While Main.CSOK(CSPointer)
            If (RowNumber Mod 2) = 0 Then
                BGColor = ClassPanelEven
            Else
                BGColor = ClassPanelOdd
            End If
            Resolved = Main.GetCSBoolean(CSPointer, "Resolved")
            clearCheckout = False
            FileCheckedOut = Main.EncodeBoolean(MemberIDCheckedOut)
            MemberNameCheckedOut = ""
            If FileCheckedOut Then
                MemberIDCheckedOut = Main.GetCSInteger(CSPointer, "CheckOutMemberID")
                If MemberIDCheckedOut = 0 Then
                    clearCheckout = True
                Else
                    MemberNameCheckedOut = GetFormattedMember(Main, MemberIDCheckedOut)
                    If MemberNameCheckedOut = "" Then
                        clearCheckout = True
                    End If
                End If
                If clearCheckout Then
                    FileCheckedOut = False
                    Call Main.SetCS(CSPointer, "checkOutMemberId", 0)
                End If
            End If
            Stream = Stream & "<TR class=""" & BGColor & """>"
            
            docLink = Main.RefreshQueryString _
                & "&" & RequestNameFormID & "=" & FormDocumentDetails _
                & "&" & RequestNameDocumentID & "=" & Main.GetCSInteger(CSPointer, "ID") _
                & "&" & RequestNameGroupName & "=" & GroupName _
                & ""
            
            Stream = Stream & "<TD class=""colName""><a onClick=""GetAjax('matResponder','" & docLink & "','','matContainer','',''); return false;"" href=""#"">" & Main.GetCSText(CSPointer, "Name") & "</a></TD>"
            Stream = Stream & "<TD class=""colDateAdded"">" & FormatDateTime(Main.GetCSDate(CSPointer, "DateAdded"), vbShortDate) & "</TD>"
            Stream = Stream & "<TD class=""colFileStatus"">"
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
    
            
    docLink = Main.RefreshQueryString _
        & "&" & RequestNameFormID & "=" & FormAddDocument _
        & "&" & RequestNameGroupName & "=" & GroupName _
        & ""
        
    Stream = Stream & "<div class=""" & ClassNav & """><a onClick=""GetAjax('matResponder','" & docLink & "','','matContainer','',''); return false;"" href=""#"">Add a new document</a></div>"
    
    GetDocumentsList = Stream
    
    Exit Function
ErrorTrap:
    Call HandleError("materialsClass", "GetDocumentsList", Err.Number, Err.Source, Err.Description, True, False)
    End Function

Private Function GetFormAddDocument(Main As Object, GroupName As String) As String
    On Error GoTo ErrorTrap
    
    Dim Stream As String
    Dim GroupID As Long
    
    GroupID = Main.GetGroupID(GroupName)
    
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
    Stream = Stream & "<TD>" & Main.GetFormInputHidden(RequestNamePostedGroupID, GroupID) & Main.GetFormInputHidden(RequestNameSourceFormID, FormID) & Main.GetFormButton(FormButtonUpload) & "</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "</TABLE>"
    Stream = Stream & Main.GetUploadFormEnd()
    
    GetFormAddDocument = Stream
    
    Exit Function
ErrorTrap:
    Call HandleError("materialsClass", "GetFormAddDocument", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
' return the name formated -- or return blank if no member found
'
Private Function GetFormattedMember(Main, SelectedMemberID As Long) As String
    On Error GoTo ErrorTrap
    
    Dim CSPointer As Long
    Dim Email As String
    Dim FirstName As String
    Dim LastName As String
    '
    GetFormattedMember = ""
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
    Call HandleError("materialsClass", "GetMemberEMail", Err.Number, Err.Source, Err.Description, True, False)
    End Function

Private Function GetDocumentDetails(Main As Object) As String
    On Error GoTo ErrorTrap
    
    Dim CSPointer As Long
    Dim Stream As String
    Dim MemberCheckOutID As Long
    Dim ModifiedDate As String
    Dim Resolved As Boolean
    Dim Description As String
    Dim addonId As Long
    Dim addonGuid As String
    '
'    If PageID <> 0 Then
'        Call Main.AddRefreshQueryString("bid", PageID, True)
'    Else
'        addonId = Main.GetStreamInteger("addonid")
'        If addonId <> 0 Then
'            Call Main.AddRefreshQueryString("addonid", addonId, True)
'        Else
'            addonGuid = Main.GetStreamInteger("addonGuid")
'            If addonGuid <> "" Then
'                Call Main.AddRefreshQueryString("addonGuid", addonGuid, True)
'            Else
'            End If
'        End If
'    End If
    
    Call Main.AddRefreshQueryString(RequestNameDocumentID, DocumentID)
    Call Main.AddRefreshQueryString(RequestNameFormID, FormDocumentResponse)
    
    If ErrorMessage <> "" Then
        Stream = Stream & "<span class=""ccError"">" & ErrorMessage & "</span>"
    End If
    
    Stream = Stream & Main.GetUploadFormStart()
    CSPointer = Main.OpenCSContent(ContentNameDocuments, "ID=" & Main.EncodeSQLNumber(DocumentID))
    If Main.CSOK(CSPointer) Then
        Resolved = Main.GetCSBoolean(CSPointer, "Resolved")
        MemberCheckOutID = Main.GetCSInteger(CSPointer, "CheckOutMemberID")
        ModifiedDate = Main.GetCS(CSPointer, "ModifiedDate")
        Description = Main.GetCSText(CSPointer, "Description")
        
        Stream = Stream & "<div>"
        Stream = Stream & "<div class=""" & ClassDocName & """>" & Main.GetCSRecordEditLink(CSPointer) & Main.GetCSText(CSPointer, "Name") & "</div>"
        If Resolved Then
            Stream = Stream & "This file has been resolved" & Break
        End If
        Stream = Stream & "Date Added: " & Main.GetCS(CSPointer, "DateAdded") & Break
        If ModifiedDate <> "" Then
            Stream = Stream & "Last Modified: " & ModifiedDate & Break
            Stream = Stream & "Modified By: " & Main.GetCSLookup(CSPointer, "ModifiedBy")
        End If
        Stream = Stream & "</div>"
        
        If Description <> "" Then
            Stream = Stream & "<div>" & Description & "</div>"
        End If
        
        Stream = Stream & "<img src=""/ccLib/images/spacer.gif"" width=""1"" height=""10"">" & Break
        Stream = Stream & "<div><a target=_blank href=""" & Main.ServerFilePath & Main.GetCSText(CSPointer, "DocumentFileName") & """>" & Main.GetSiteProperty("Caption - View Document", "View Document", True) & "</a></div>"
        Stream = Stream & "<img src=""/ccLib/images/spacer.gif"" width=""1"" height=""10"">" & Break
        
        Stream = Stream & "<div>"
        If Not Resolved Then
            If MemberCheckOutID = 0 Then
                Stream = Stream & Main.GetFormButton(FormButtonCheckOut)
                End If
            If MemberCheckOutID = Main.MemberID Then
                Stream = Stream & Main.GetFormInputFile(RequestNameDocumentFileName)
                Stream = Stream & Break & Main.GetFormButton(FormButtonCheckIn)
                End If
            End If
        Stream = Stream & "</div>"
    Else
        Stream = "<span class=""ccError"">The document requested is currently unavailable.</span>"
        End If
    Call Main.CloseCS(CSPointer)
    Stream = Stream & Main.GetFormInputHidden(RequestNameSourceFormID, FormID)
    Stream = Stream & Main.GetUploadFormEnd()
    
    GetDocumentDetails = Stream
    
    Exit Function
ErrorTrap:
    Call HandleError("materialsClass", "GetMemberEMail", Err.Number, Err.Source, Err.Description, True, False)
    End Function

Private Sub ProcessDocument(Main As Object)
    On Error GoTo ErrorTrap
    
    Dim CSPointer As Long
    Dim Filename As String
    
    If (ButtonValue = FormButtonCheckOut) Or (ButtonValue = FormButtonCheckIn) Or (ButtonValue = FormButtonResolve) Then
        
        CSPointer = Main.OpenCSContent(ContentNameDocuments, "ID=" & Main.EncodeSQLNumber(DocumentID))
        If Main.CSOK(CSPointer) Then
            
            If ButtonValue = FormButtonCheckOut Then
                Call Main.SetCS(CSPointer, "CheckOutMemberID", Main.MemberID)
                Call Main.SaveCSRecord(CSPointer)
                FormID = FormCheckOutConfirmation
                End If
            
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
            
            If ButtonValue = FormButtonResolve Then
                Call Main.SetCS(CSPointer, "Resolved", 1)
                Call Main.SaveCSRecord(CSPointer)
                FormID = FormResolvedConfirmation
                End If
            
            End If
        Call Main.CloseCS(CSPointer)
        
        End If
    
    Exit Sub
ErrorTrap:
    Call HandleError("materialsClass", "ProcessDocument", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Function GetResponseFileLink(Main As Object) As String
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
    Call HandleError("materialsClass", "GetResponseFileLink", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetFormLanding(Main As Object, GroupName) As String
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
    Call HandleError("materialsClass", "GetFormLanding", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetFormCollaboratorList(Main As Object, GroupName As String) As String
    On Error GoTo ErrorTrap
    '
    Dim CSPointer As Long
    Dim Stream As String
    '
    Stream = Stream & Main.GetFormStart()
    '
    Stream = Stream & GetCheckedGoupList(Main, GroupName) & Break
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
    Call HandleError("materialsClass", "GetFormCollaboratorList", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Sub ProcessFormCollaboratorList(Main As Object)
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
                Email = GetEmail(Main, SelectedMemberID)
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
    Call HandleError("materialsClass", "ProcessFormCollaboratorList", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Function GetGoupList(Main As Object, GroupName As String) As String
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
            Stream = Stream & GetFormattedMember(Main, Main.GetCSInteger(CSPointer, "MemberID")) & Break
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
    Call HandleError("materialsClass", "GetGoupList", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetCheckedGoupList(Main As Object, GroupName As String) As String
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
            Stream = Stream & Main.GetFormInputCheckBox(MemberCheckBoxPrefix & SelectedMemberID) & " " & GetFormattedMember(Main, SelectedMemberID) & Break
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
    Call HandleError("materialsClass", "GetCheckedGoupList", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetEmail(Main As Object, SelectedMemberID As Long) As String
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
    Call HandleError("materialsClass", "GetEmail", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Sub ProcessForm(Main As Object)
    On Error GoTo ErrorTrap
    '
    Select Case SourceFormID
        Case FormCollaboratorList
            Call ProcessFormCollaboratorList(Main)
        Case FormDocumentDetails
            Call ProcessDocument(Main)
        Case FormAddDocument
            Call ProcessFormAddDocument(Main)
    End Select
    '
    Exit Sub
ErrorTrap:
    Call HandleError("materialsClass", "ProcessForm", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Sub ProcessFormAddDocument(Main As Object)
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
    Call HandleError("materialsClass", "ProcessFormAddDocument", Err.Number, Err.Source, Err.Description, True, False)
    End Sub

