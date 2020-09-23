Attribute VB_Name = "ToDlg"

Option Explicit

'Printer Shit
Public Declare Function SetTextAlign Lib "gdi32.dll" (ByVal hdc As Long, ByVal wFlags As Long) As Long
Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" _
                                   (ByVal hwnd As Long, ByVal wmsg As Long, _
                                    ByVal wparam As Long, lparam As Any) As Long
Public Const TA_CENTER = 4

'Sleep and counter plus shell about dialog api
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" _
                                   (ByVal hwnd As Long, ByVal szApp As String, _
                                    ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
'ToDlg...
'05.2003 mrk Change/Add

Public Function ToDlgShow( _
              ByRef Dlg As CommonDialog) As Boolean
  On Error Resume Next
  
  With Dlg
    .CancelError = True
    .ShowSave
    ToDlgShow = CBool(Err.Number = 0)
  End With
  
  If Err.Number <> 0 Then
     Err.Clear
  End If

  On Error GoTo 0
End Function '05.2003 mrk Change/Add

Public Function ToDlgOpen( _
              ByRef Dlg As CommonDialog) As Boolean
  On Error Resume Next
  
  With Dlg
    .CancelError = True
    .ShowOpen
    ToDlgOpen = CBool(Err.Number = 0)
  End With
  
  If Err.Number <> 0 Then
     Err.Clear
  End If
  
  On Error GoTo 0
End Function '05.2003 mrk Change/Add

Public Sub ToDlgAbout( _
              ByVal Form As Form, _
              ByVal App As String, _
              ByVal OtherStuff As String, _
              ByVal Icon)
  On Error Resume Next
  
  ShellAbout Form.hwnd, _
             App, _
             OtherStuff, _
             CLng(Icon)
  
  If Err.Number <> 0 Then
     Err.Clear
  End If
  
  On Error GoTo 0
End Sub '05.2003 mrk Change/Add


