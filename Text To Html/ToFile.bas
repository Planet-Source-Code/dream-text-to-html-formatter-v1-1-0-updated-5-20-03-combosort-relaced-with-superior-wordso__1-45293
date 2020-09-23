Attribute VB_Name = "ToFile"

Option Explicit

Public Enum eToFileLoadTextType
 Default = 0 'Load all
 VBDoc = 1     'Load Code
End Enum
'
'ToFile...
'05.2003 mrk Change/Add

Public Function ToFileLoad( _
              ByVal FileName As String, _
     Optional ByVal TextType As eToFileLoadTextType = Default) As String
  On Error Resume Next
  
  Dim FF As Integer
  Dim sText As String
  Dim fText As String
  Dim TextToAdd As Long
    
  TextToAdd = 0
  
  FF = ToFileFree
  
  Select Case TextType
    
    Case eToFileLoadTextType.Default
         Open FileName For Binary Access Read As #FF
         sText = Space$(LOF(FF))
         Get FF, , sText
    
    Case eToFileLoadTextType.VBDoc
         Open FileName For Input As #FF

         Do Until EOF(FF)
            Line Input #FF, fText
            
            Select Case TextToAdd
              Case 0
                   If InStr(1, Left$(fText, 10), "Attribute") = 1 Then
                      TextToAdd = 1
                   End If
              
              Case 1
                   If InStr(1, Left$(fText, 10), "Attribute") = 0 Then
                      TextToAdd = 2
                      sText = sText & fText & vbCrLf   'load each line into txtdata
                   End If
              
              Case 2
                   'read read read
                   sText = sText & fText & vbCrLf   'load each line into txtdata
            End Select
            
            DoEvents
         Loop
  End Select
  
  If FF > 0 Then
     Close #FF
  End If

  ToFileLoad = sText
  sText = ""
  fText = ""
  If Err.Number <> 0 Then
     Err.Clear
  End If
  
  On Error GoTo 0
End Function

Public Sub ToFileSave( _
              ByVal FileName As String, _
              ByVal Text As String)
  On Error Resume Next
  
  Dim FF As Integer
  
  FF = ToFileFree
  Open FileName For Output As #FF
  Print #FF, Text
  Close #FF
  
  If Err.Number <> 0 Then
     Err.Clear
  End If
  
  On Error GoTo 0
End Sub '05.2003 mrk Change/Add

Public Sub ToFileKill( _
              ByVal FileName As String)
  On Error Resume Next
  
  If ToFileIsExist(FileName) Then
     ToFileAttrSet FileName, vbNormal
     Kill FileName
  End If
  
  If Err.Number <> 0 Then
     Err.Clear
  End If
  
  On Error GoTo 0
End Sub '05.2003 mrk Change/Add

Public Function ToFileIsExist( _
              ByVal FileName As String) As Boolean
  On Error Resume Next
  
  Dim Value As Boolean
  Value = CBool(Len(Dir$(FileName, vbArchive Or vbHidden Or vbNormal Or vbReadOnly)) > 0)
  
  If Err.Number <> 0 Then
     Err.Clear
     Value = False
  End If
  
  ToFileIsExist = Value
  
  On Error GoTo 0
End Function '05.2003 mrk Change/Add

Private Sub ToFileAttrSet( _
              ByVal FileName As String, _
              ByVal Attri As VbFileAttribute)
  On Error Resume Next
  
  If ToFileIsExist(FileName) Then
     SetAttr FileName, Attri
  End If
  
  If Err.Number <> 0 Then
     Err.Clear
  End If
  
  On Error GoTo 0
End Sub '05.2003 mrk Change/Add

Private Function ToFileAttrGet( _
              ByVal FileName As String) As VbFileAttribute
  On Error Resume Next
  
  If ToFileIsExist(FileName) Then
     ToFileAttrGet = GetAttr(FileName)
  End If
  
  If Err.Number <> 0 Then
     Err.Clear
  End If
  
  On Error GoTo 0
End Function '05.2003 mrk Change/Add


Private Function ToFileFree() As Integer
  On Error Resume Next
  
  ToFileFree = FreeFile
  
  If Err.Number <> 0 Then
     '--> ToFileFree = 0
     Err.Clear
  End If
  
  On Error GoTo 0
End Function '05.2003 mrk Change/Add
