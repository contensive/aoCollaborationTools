VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "directoryClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Private GroupName As String

Private Main As Object
Private CSV As Object

Public Function Execute(CsvObject As Object, MainObject As Object, optionString As String, FilterInput As String) As String
    Set CSV = CsvObject
    
    Call Init(MainObject)
    
    Execute = GetContent(optionString)
    
End Function

Private Sub Init(MainObject As Object)
    Set Main = MainObject
    
    Exit Sub

ErrorTrap:
    Call HandleError("grpDirectoryClass", "Init", Err.Number, Err.Source, Err.Description, True, False)
    End Sub

Private Function GetContent(optionString As String) As String
    On Error GoTo ErrorTrap
    
    Dim Stream As String
    Dim CS As Long
    Dim SQL As String
    Dim Caption As String
    '
    GroupName = Trim(Main.GetAddonOption("Group", optionString))
    If GroupName = "" Then
        Stream = Stream & "<p>This directory is empty</p>"
        If Main.IsAdmin Then
            Stream = Stream & Main.GetAdminHintWrapper("<p>Please set the group name for this add-on by clicking the options icon with advanced editing enabled.</p>")
        End If
        GroupName = "Collaboration Space"
    Else
        Caption = GroupName
        SQL = "select caption from ccgroups where name=" & Main.EncodeSQLText(GroupName)
        CS = Main.OpenCSSQL("", SQL)
        If Main.IsCSOK(CS) Then
            Caption = Main.GetCSText(CS, "caption")
        End If
        Call Main.CloseCS(CS)
        '
        CS = Main.OpenCSGroupMembers(GroupName, , "Name", True)
        If Main.CSOK(CS) Then
            Stream = "<div class=""dirContainer"">"
            Stream = Stream & "<table width=""100%"" cellpadding=""3"" cellspacing=""0"" border=""0"">"
            Stream = Stream & GetFormRow2(Caption, "header")
            Do While Main.CSOK(CS)
                Stream = Stream & GetMemberRow(CS)
                Call Main.NextCSRecord(CS)
            Loop
            Stream = Stream & "</table>"
            Stream = Stream & "</div>"
        End If
        Call Main.CloseCS(CS)
        If Main.IsAdmin Then
            Stream = Stream & Main.GetAdminHintWrapper("<p>To add a new person to this directory, edit their people record and add them to the group '" & GroupName & "'.</p>")
        End If
    End If
    
    GetContent = Stream
    
    Exit Function
ErrorTrap:
    Call HandleError("grpDirectoryClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function
    
Private Function GetFormRow2(FieldCaption As String, Class As String) As String
    On Error GoTo ErrorTrap
    
    Dim Stream As String
    
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD"
    If Class <> "" Then
        Stream = Stream & " class=""" & Class & """"
    End If
    Stream = Stream & ">" & FieldCaption & "</TD>"
    Stream = Stream & "</TR>"
    
    GetFormRow2 = Stream
    
    Exit Function
ErrorTrap:
    Call HandleError("grpDirectoryClass", "GetFormRow", Err.Number, Err.Source, Err.Description, True, False)
    End Function

Private Function GetMemberRow(Pointer As Long) As String
    On Error GoTo ErrorTrap
    
    Dim Stream As String
    Dim Email As String
    Dim Address2 As String
    Dim Name As String
    Dim Title As String
    Dim Company As String
    Dim Phone As String
    
    Name = Main.GetCS(Pointer, "Name")
    Email = Main.GetCS(Pointer, "Email")
    Address2 = Main.GetCS(Pointer, "Address2")
    Title = Main.GetCS(Pointer, "Title")
    Company = Main.GetCS(Pointer, "Company")
    Phone = Main.GetCS(Pointer, "Phone")
    If Name = "" Then
        Name = "Unnamed Member #" & Main.GetCSInteger(Pointer, "id")
    End If
    '
    Stream = Stream & "<tr>"
    Stream = Stream & "<td class=""membRowLeft"">"
    Stream = Stream & "<div><strong>" & Name & "</strong></div>"
    If Title <> "" Then
        Stream = Stream & "<div>" & Title & "</div>"
    End If
    If Company <> "" Then
        Stream = Stream & "<div>" & Company & "</div>"
    End If
    If Email <> "" Then
        Stream = Stream & "<a href=""mailto:" & Email & """>" & Email & "</a><br />"
    End If
    If Phone <> "" Then
        Stream = Stream & "<div>" & Phone & "</div>"
    End If
    '
    Stream = Stream & "</td>"
    GetMemberRow = Stream
    
    Exit Function
ErrorTrap:
    Call HandleError("grpDirectoryClass", "GetMemberRow", Err.Number, Err.Source, Err.Description, True, False)
    End Function

