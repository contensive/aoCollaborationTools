VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "issueResponder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private GroupName As String

Private Main As Object
Private CSV As Object

Public Function Execute(CsvObject As Object, MainObject As Object, OptionString As String, FilterInput As String) As String
    On Error GoTo ErrorTrap
    
    Set CSV = CsvObject
    
    Call Init(MainObject)
    
    Execute = GetContent(OptionString)
    
    Exit Function
ErrorTrap:
    Call HandleError("issueResponder", "Execute", Err.Number, Err.Source, Err.Description, True, False)
    End Function

Private Sub Init(MainObject As Object)
    On Error GoTo ErrorTrap
    
    Set Main = MainObject
    
    GroupName = Main.GetStreamText(RequestNameGroupName)
    
    Exit Sub

ErrorTrap:
    Call HandleError("issueResponder", "Init", Err.Number, Err.Source, Err.Description, True, False)
    End Sub

Private Function GetContent(OptionString As String) As String
    On Error GoTo ErrorTrap
    
    Dim Stream As String

    Stream = Main.ExecuteAddon(0, "{B6A208C8-2AEF-4108-88EB-979992F84CDC}", "Group=" & GroupName, ContextPage, "", 0, "", "", 0)

    GetContent = Stream
    
    Exit Function
ErrorTrap:
    Call HandleError("issueResponder", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function

