Attribute VB_Name = "Error"
 

Public Sub LogErrorControl(Optional strInput1 As String, _
    Optional strInput2 As String, _
    Optional strInput3 As String)
    
    Dim strMsg As String
    'The msg in a msgbox explaining the erro
    '     r to the user.
    Dim strTitle As String
    'The title of that msgbox
    Dim OldErrDesc As String
    Dim OldErrNum As Long
    'The old info is an error
    'the old error info is n
    'not erased.
      Dim intFile As Integer
        'This is the file number, a handle for VB.
        
        OldErrDesc = err.Description
        OldErrNum = err.Number
        
    If lFatal = True Then
            strMsg = "Fatal"
        Else
            strMsg = "Unexpected"
    End If
        
        strMsg = strMsg & " error: " & err.Description & vbCrLf & _
        vbCrLf & "Please contact the program vendor To " & _
        "inform them of this error."
        strTitle = App.Title & " v" & App.Major & "." & App.Minor
        strTitle = strTitle & "Error #" & err.Number
        MsgBox strMsg, vbExclamation + vbOKOnly, strTitle
        
        On Error GoTo ErrWhileLogging:
        'That's in case logging the error genera
        '     tes an error.
        
        'Log the error in error.log
        intFile = FreeFile
        Open App.Path & "\errors.log" For Append As #intFile
        Print #intFile, ""
        Print #intFile, "----------------------------------------------------"


        If lFatal Then
            Print #intFile, "Fatal";
        Else
            Print #intFile, "Non-fatal";
        End If
        
        Print #intFile, " Error in " & App.Path & "\";
        Print #intFile, App.Title & " v" & App.Major & "." & App.Minor


        If Not IsNull(strInput1) Then
            Print #intFile, strInput1
        End If


        If Not IsNull(strInput2) Then
            Print #intFile, strInput2
        End If


        If Not IsNull(strInput3) Then
            Print #intFile, strInput3
        End If
        Print #intFile, Date & "  " & Time
        Print #intFile, "Error #" & OldErrNum
        Print #intFile, "" & OldErrDesc
        Close #intFile
        
        Exit Sub
        
ErrWhileLogging:
        strMsg = "Fatal Error: Could Not log error." & vbCrLf & _
        "Please contact the program vendor With the following " & _
        "error information:" & vbCrLf & vbCrLf & _
        "Err #" & OldErrNum & vbCrLf & _
        OldErrDesc


        If Not IsNull(strInput1) Then
            strMsg = strMsg & vbCrLf & strInput1
        End If


        If Not IsNull(strInput2) Then
            strMsg = strMsg & vbCrLf & strInput2
        End If


        If Not IsNull(strInput3) Then
            strMsg = strMsg & vbCrLf & strInput3
        End If
        MsgBox strMsg
        'End
End Sub

