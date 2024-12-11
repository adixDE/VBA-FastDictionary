Attribute VB_Name = "AppDeclare"
'@Folder("Project.Module")
'@IgnoreModule FunctionReturnValueDiscarded,FunctionReturnValueAlwaysDiscarded,ProcedureNotUsed,ProcedureCanBeWrittenAsFunction
Option Explicit
Option Private Module
' *******************************************************************************
' ObjectName: AppDeclare
' Version:    1.0.0
' --- Description ---------------------------------------------------------------
'  core declare module
' *******************************************************************************
#If Not VBA6 Then

Private Const ProjectName As String = "Wrapper"
Private Const ProjectVersionMajor As Byte = 1
Private Const ProjectVersionMinor As Byte = 0
Private Const ProjectVersionRevision As Byte = 0

'Considered 'conditional compiler variables' of this project:
' * EnableToAny = 1   (1 = Enable - which require ValueToAny/IValueToAny)
'   Used in:  Text,DicScript/IDicScriptRead,ExcelCore
' Set/adjust the conditional compiler variable of this project via Tools>VBA Project Properties. (seperator = ':')

' ****************************************
Public Function GetAppName() As String
  GetAppName = ProjectName
End Function

Public Function GetAppVersion() As String
  If ProjectVersionRevision = 0 Then
    GetAppVersion = ProjectVersionMajor & "." & Format$(ProjectVersionMinor, "00")
  Else
    GetAppVersion = ProjectVersionMajor & "." & Format$(ProjectVersionMinor, "00") & "." & Format$(ProjectVersionRevision, "000")
  End If
End Function

Public Function GetProjectName() As String
  GetProjectName = GetAppName & Space(1) & GetAppVersion
End Function

'Public Function GetGuid(Optional ByVal GuidLen As Long = 20) As String
'  GetGuid = Text.GetGuid(GuidLen)
'End Function

Public Function GetError(Optional ByRef ErrorObject As ErrObject) As String
  
  Dim ReturnText As String: ReturnText = vbNullString
  Dim TextBlock As String: TextBlock = vbNullString
  
  If ErrorObject Is Nothing Then Set ErrorObject = Err
  If Not ErrorObject.Number = 0 Then
    ReturnText = "Unhandeled exception (" & ErrorObject.Number & ")"
    If Erl > 0 Then ReturnText = ReturnText & " within line " & Erl
    ReturnText = ReturnText & " occured. "
    
    TextBlock = Replace(ErrorObject.Description, ".", vbCrLf)
    If Len(TextBlock) > 2 Then If Right$(TextBlock, 2) = vbCrLf Then TextBlock = Left$(TextBlock, Len(TextBlock) - 2)
    
    On Error Resume Next
    If ThisWorkbook.VBProject Is Nothing Then
      Err.Clear
    Else
      If Not ErrorObject.Source = vbNullString Then
        If Not ErrorObject.Source = ThisWorkbook.VBProject.Name Then
          ReturnText = ReturnText & " within '" & ErrorObject.Source & "'"
        End If
      End If
    End If
    If Err.Number Then Err.Clear
    On Error GoTo 0
    
    If InStr(1, TextBlock, vbCrLf) > 0 Then
      ReturnText = ReturnText & vbCrLf & TextBlock
    Else
      If Len(ReturnText & TextBlock) > 80 Then ReturnText = ReturnText & vbCrLf
      ReturnText = ReturnText & TextBlock
    End If
  End If
  GetError = ReturnText
  
End Function

Public Sub RaiseError(ByVal ErrorText As String, Optional ByVal FunctionName As String)
  If Not ErrorText = vbNullString Then
    MsgBox ErrorText, vbCritical + vbOKOnly, "Exception occured" & IIf(FunctionName <> vbNullString, " within '" & FunctionName & "'", vbNullString)
  End If
End Sub

Public Sub ResetError(ByRef ErrorText As String)
  ErrorText = vbNullString
  If Err.Number Then Err.Clear
End Sub

Public Sub SaveMe(Optional ByVal RemovePersonalInformation As Boolean = False)
  With ThisWorkbook
    If Not .RemovePersonalInformation = RemovePersonalInformation Then .RemovePersonalInformation = RemovePersonalInformation
    If Not .Saved Then
      .KeepChangeHistory = False
      Application.DisplayAlerts = False
      .Save
      Application.DisplayAlerts = True
    End If
  End With
End Sub

#End If
