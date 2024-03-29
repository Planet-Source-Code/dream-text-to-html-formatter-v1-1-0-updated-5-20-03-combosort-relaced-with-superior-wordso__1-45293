Attribute VB_Name = "basColor"
Option Explicit

'#--------------------------------------------------------------------------
'#
'#    File..........: basColor [Method 3]
'#    Original Author........: Will Barden 10/9/02                                              #
'#    Last Modified.: Dream /  '05.2003 mrk Change/Add
'#    ComboSort replaced with WordSort algorithm/wordlist updated
'#    Dependancies..: None
'#--------------------------------------------------------------------------
'#  apis, enums, consts, declares
'#--------------------------------------------------------------------------
' api to stop the window refreshing
Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Const COL_KEYWORD = &H800000    ' dark blue
Const COL_COMMENT = &H8000&     ' middle green
Const CHAR_COMMENT = "'"        ' comment line char

Type WORD_TYPE
    Text As String  ' word to be colored
    Color As Long   ' color to color the word :)
End Type

Type LETTER_TYPE
    Start As Long   ' first time the letter appears in the list
    Finish As Long  ' last time the letter appears in the list
End Type

'#--------------------------------------------------------------------------#
'#  variables
'#--------------------------------------------------------------------------#
Dim Words() As WORD_TYPE
Dim Letters() As LETTER_TYPE
Dim Strings() As String
Public sText As String

Public Sub InitKeyWords()
  ' initialize the array of words
    a "Alias", COL_KEYWORD
  a "And", COL_KEYWORD
  a "As", COL_KEYWORD
  a "Boolean", COL_KEYWORD
  a "(Byref", COL_KEYWORD
  a "Byref", COL_KEYWORD
  a "Byte", COL_KEYWORD
  a "ByVal", COL_KEYWORD
  a "Case", COL_KEYWORD
  a "Close", COL_KEYWORD
  a "Const", COL_KEYWORD
  a "Currency", COL_KEYWORD
  a "Declare", COL_KEYWORD
  a "Dim", COL_KEYWORD
  a "Div", COL_KEYWORD
  a "Do", COL_KEYWORD
  a "DoEvents", COL_KEYWORD
  a "Double", COL_KEYWORD
  a "Else", COL_KEYWORD
  a "End", COL_KEYWORD
  a "Enum", COL_KEYWORD
  a "Event", COL_KEYWORD
  a "Exit", COL_KEYWORD
  a "Explicit", COL_KEYWORD
  a "False", COL_KEYWORD
  a "For", COL_KEYWORD
  a "Function", COL_KEYWORD
  a "Global", COL_KEYWORD
  a "Goto", COL_KEYWORD
  a "If", COL_KEYWORD
  a "Integer", COL_KEYWORD
  a "Is", COL_KEYWORD
  a "LBound", COL_KEYWORD
  a "lib", COL_KEYWORD
  a "line", COL_KEYWORD
  a "Loop", COL_KEYWORD
  a "Long", COL_KEYWORD
  a "Me", COL_KEYWORD
  a "Mod", COL_KEYWORD
  a "Next", COL_KEYWORD
  a "Object", COL_KEYWORD
  a "Open", COL_KEYWORD
  a "Option", COL_KEYWORD
  a "Private", COL_KEYWORD
  a "Public", COL_KEYWORD
  a "ReDim", COL_KEYWORD
  a "Single", COL_KEYWORD
  a "Static", COL_KEYWORD
  a "String", COL_KEYWORD
  a "Sub", COL_KEYWORD
  a "Then", COL_KEYWORD
  a "To", COL_KEYWORD
  a "True", COL_KEYWORD
  a "Type", COL_KEYWORD
  a "Typeof", COL_KEYWORD
  a "UBound", COL_KEYWORD
  a "Until", COL_KEYWORD
  a "Wend", COL_KEYWORD
  a "While", COL_KEYWORD
  a "Xor", COL_KEYWORD
 
 'sort the array
  WordSort
 
 'build the index of letter positions
  BuildIndex
End Sub

Private Sub a( _
             ByVal Text As String, _
             ByVal Color As Long)
 'Add Text and Color To Words()
  On Error Resume Next
  Dim Index As Long
 
  Index = UBound(Words) + 1   'Err.Number <> 0 -> Index = 0
             
  ReDim Preserve Words(Index)
  With Words(Index)
   .Text = Text
   .Color = Color
  End With

  If Err.Number Then
    Err.Clear
  End If
  On Error GoTo 0
End Sub

Public Sub WordSort()
  Dim U As Long, _
      i As Long, _
      J As Long, _
      N As Long, _
      V As WORD_TYPE
' Sort Words().Text A..Z

 On Error Resume Next
 
  N = UBound(Words)
 
  U = CLng(VBA.Log(N) / VBA.Log(2))
 U = 2 ^ U - 1
 
  Do While U > 0
    For i = 1 To N - U
        For J = i To 1 Step -U
            If Words(J).Text < Words(J + U).Text Then
         Exit For
            End If
             V = Words(J)
            Words(J) = Words(J + U)
            Words(J + U) = V
        Next J
    Next i
    U = CLng(U / 2)
 Loop
End Sub

Public Sub ColorIn(RTB As RichTextBox, Optional Position As Long = 0)
  Dim lStart As Long
  Dim lFinish As Long
  Dim Text As String
  On Error Resume Next
 ' split the text into lines and color them one by one
    LockWindowUpdate RTB.hwnd
    RTB.Visible = False
    Text = RTB.Text
    basColor.sText = RTB.Text
    lStart = 1
    Do While lStart <> 2 And lStart < Len(Text)
        ' find the end of this line
        lFinish = InStr(lStart + 1, Text, vbCrLf)
        If lFinish = 0 Then lFinish = Len(Text)
            
        ' color it
        DoColor RTB, lStart, lFinish
        
        ' move up to get the next line
        lStart = lFinish + 2
    Loop
    
    ' reset the cursor
    RTB.SelStart = Position
    RTB.Visible = True
    RTB.SetFocus
    LockWindowUpdate 0&
End Sub
'//--[DoColor]--------------------------------------------------------------//
'
'  Here it is - the beast itself. This routine colors
'  a single line of text within the RTB. It will
'  split each line up into words using the custom
'  split function (SplitWords), then match each word
'  against the list of keywords.
'
Public Sub DoColor(RTB As RichTextBox, ByVal lStart As Long, ByVal lFinish As Long)
Dim sWords()    As String
Dim sLine       As String
Dim sChar       As String
Dim lCurPos     As Long
Dim lIndex      As Long
Dim lColor      As Long
Dim lPos        As Long
Dim lPos2       As Long
Dim lCom        As Long
Dim i           As Long

    ' grab the line
    sLine = Trim$(Mid$(sText, lStart, lFinish - lStart))
    ' remove the EOL
    sLine = RemoveEOL(sLine)
    ' remove the quotes so they're not colored
    sLine = RemoveStrings(sLine)
    
    ' split the line into words using our custom function
    sWords = SplitWords(sLine)
    
    ' check each word against the list
    lCurPos = 1
    ' search for each word in the array
    For i = LBound(sWords) To UBound(sWords)
    
        If Trim$(sWords(i)) <> "" Then
    
            ' check for comment whilst in the middle of a line
            If Left$(sWords(i), 1) = CHAR_COMMENT Then
            
                ' color the rest of the line
                RTB.SelStart = InStr(lStart, sText, sWords(i)) - 1
                RTB.SelLength = Len(sWords(i))
                RTB.SelColor = COL_COMMENT
            
            Else
        
                ' its a normal keyword - so color it
                ' first get the array positions from
                ' the index
                If sWords(i) = "ByVal" Or sWords(i) = "Byte" Then
                    DoEvents: End If
               'sChar = Mid(LCase$(sWords(i)), 2, 1)
                sChar = Left$(LCase$(sWords(i)), 1)
                ' if we've got a valid alphabetic char
                If sChar <> "" Then
                    ' convert this char to an index in the letters array
                    lIndex = Asc(sChar) - 97
                    ' if the index is a valid one - this
                    ' means that the text is a word, so
                    ' we should try to color it
                    If lIndex >= 0 And lIndex < UBound(Letters) Then
                        ' color the word, passing the index parameters
                        lColor = GetColor(sWords(i), _
                                    Letters(lIndex).Start, _
                                    Letters(lIndex).Finish)
                        ' if a color was returned - color the word
                        If lColor Then
                            ' locate the word in the line
                            RTB.SelStart = InStr(lStart + lCurPos - 1, sText, sWords(i)) - 1
                            RTB.SelLength = Len(sWords(i))
                            RTB.SelColor = lColor
                        End If
                    End If
                End If ' sChar <> ""
            End If ' CHAR_COMMENT
        End If ' sWords(i) <> ""
        
        ' move the current position within the line on
        lCurPos = lCurPos + Len(sWords(i))
        
    Next i
        
End Sub

'//--[DoClipBoardPaste]-----------------------------------------------------//
'
'  Call this when text has been pasted into the
'  RTB. It will grab the text, split it into lines
'  and color it.
'
Public Sub DoClipBoardPaste(RTB As RichTextBox)
Dim lCursor As Long
Dim lStart As Long
Dim lFinish As Long
Dim sText As String
Dim p1 As Long, p2 As Long
On Error Resume Next
    ' store the cursor position
    lCursor = RTB.SelStart
    
    ' add the text and color it
    LockWindowUpdate RTB.hwnd
    
    ' get the text to be pasted from the clipboard
    sText = Clipboard.GetText
    
    ' get the start point - this should be the previous
    ' vbCrLf to where the text was inserted, to make
    ' sure that if it's inserted mid-line, the whole
    ' line is colored
    lStart = InStrRev(RTB.Text, vbCrLf, RTB.SelStart) + 2
    If lStart = 0 Then lStart = RTB.SelStart
    ' also store the finish point
    lFinish = RTB.SelStart + Len(sText)
    
    ' now add the text to the box
    RTB.SelText = sText
    basColor.sText = RTB.Text
    
    ' now color each line individually starting
    ' from lStart since this is the position of
    ' the first changed line
    p1 = lStart
    Do
        ' find the next EOL character, this combined
        ' with lStart gives us the line to color
        p2 = InStr(p1, RTB.Text, vbCrLf)
        If p2 = 0 Then p2 = lFinish
                    
        ' now strip out this line and color it
        ' color it black first to remove any
        ' previous coloring..
        RTB.SelStart = p1 - 1
        RTB.SelLength = p2 - p1
        RTB.SelColor = vbBlack
        DoColor RTB, p1, p2
        
        ' move the start pointer on to just after
        ' the last EOL character - essentially onto
        ' the next actual line of text
        p1 = p2 + 2
              
        ' exit condition - keep going until we can't
        ' find any more vbCrLf (<>2) and while
        ' p1 (the start of line pointer) is lower
        ' that lFinish (the end of the text we're
        ' coloring)... easy enough
        If p1 = 2 Or p1 >= lFinish + 2 Then Exit Do
        DoEvents
    Loop
    
    ' restore the original values
    RTB.SelStart = lCursor + Len(sText)
    RTB.SelColor = vbBlack
    
    ' null the keypress (to avoid the text pasting twice)
    LockWindowUpdate 0&

End Sub

'#--------------------------------------------------------------------------#
'#  private internals
'#--------------------------------------------------------------------------#

'//--[BuildIndex]----------------------------------------------------------//
'
'  Takes the Words array and constructs an alphabetical
'  index which it puts into the Letters array.
'  Each item in the letters array accounts for a letter
'  in the alphabet - Letters(0) = "a".
'  The .Start property is the Index in the Words array
'  at which that letter starts, and the finish is the
'  same. The purpose of this is to get Hi and Lo params
'  for the GetColor (a standard binary search algorithm).
'  This saves several loops round the algorithm.
'
Private Sub BuildIndex()
Dim i As Long, J As Long
Dim sChar As String
Dim bStart As Boolean

    ' go through each letter in the alphabet
    ReDim Letters(25)
    For i = 0 To 25
        ' get the current char
        sChar = Chr$(i + 97)
        ' find the first and last instances of the letter
        For J = LBound(Words) To UBound(Words)
            If Left$(LCase$(Words(J).Text), 1) = sChar Then
                If Not bStart Then
                    ' found the start
                    bStart = True
                    Letters(i).Start = J
                End If
                ' if we've hit the end of the list
                If J = UBound(Words) Then
                    Letters(i).Finish = J
                    Exit Sub
                End If
            Else
                ' its a different char
                If bStart Then
                    ' we've found the end
                    Letters(i).Finish = J - 1
                    bStart = False
                    Exit For
                End If
                ' see if we've gone too far -
                ' there are no words beginning with
                ' this letter in the list
                If Left$(LCase$(Words(J).Text), 1) > sChar Then
                    Exit For
                End If
            End If
        Next J
    Next i

End Sub

'//--[GetColor]--------------------------------------------------------------//
'
'  Searches the Words array for a match using a standard
'  binary search algorithm, using the Lo and Hi params
'  as starting points.
'
Private Function GetColor(ByVal sWord As String, _
                          ByVal Lo As Long, _
                          ByVal Hi As Long) As Long
Dim lHi As Long
Dim lLo As Long
Dim lMid As Long
    
    ' standard binary search the words array
    ' return the color if a match is found
    lLo = Lo
    lHi = Hi
    Do While lHi >= lLo
        lMid = (lLo + lHi) \ 2
        If LCase$(Words(lMid).Text) = LCase$(sWord) Then
            GetColor = Words(lMid).Color
            Exit Do
        End If
        If LCase$(Words(lMid).Text) > LCase$(sWord) Then
            lHi = lMid - 1
        Else
            lLo = lMid + 1
        End If
    Loop
    
End Function

'//--[SplitWords]---------------------------------------------------------//
'
'  Since splitting a line into words by a single
'  character is not acceptable because we have to
'  take several end of word characters into account,
'  this routine was written.
'  It searches through the string from left to right
'  and locates the nearest word break char from a list
'  then splits at that word.
'
Private Function SplitWords(ByVal sText As String) As String()
Dim i As Long, lPos As Long
Dim sWords() As String
Dim sWordBreaks(0 To 8) As String
Dim lBreakPoints() As Long
Dim lBreak As Long
    
    ' list of word break characters
    sWordBreaks(0) = " "
    sWordBreaks(1) = "("
    sWordBreaks(2) = ")"
    sWordBreaks(3) = "<"
    sWordBreaks(4) = ">"
    sWordBreaks(5) = "."
    sWordBreaks(6) = ","
    sWordBreaks(7) = "="
    sWordBreaks(8) = CHAR_COMMENT ' comments
    ReDim lBreakPoints(UBound(sWordBreaks))

    ' get them words!
    ReDim sWords(0)
    lPos = 1
    Do
    
        ' locate the word break points
        For i = 0 To UBound(sWordBreaks)
            lBreakPoints(i) = InStr(lPos, sText, sWordBreaks(i))
        Next i
        
        ' now work out which is closest
        lBreak = Len(sText) + 1
        For i = 0 To UBound(lBreakPoints)
            If lBreakPoints(i) <> 0 Then
                If lBreakPoints(i) < lBreak Then lBreak = lBreakPoints(i)
            End If
        Next i
    
        ' now split out the word
        ' if no break point was found, then we've
        ' hit the end of the line, so add all the rest
        If lBreak = Len(sText) + 1 Then
            sWords(UBound(sWords)) = Mid$(sText, lPos)
        Else
            ' add this word - first check for a comment
            If Mid$(sText, lBreak, 1) = CHAR_COMMENT Then
                ' first add the word
                sWords(UBound(sWords)) = Mid$(sText, lPos, lBreak - lPos)
                ' then add the rest as a comment
                ReDim Preserve sWords(UBound(sWords) + 1)
                sWords(UBound(sWords)) = Mid$(sText, lBreak)
                ' now return and exit
                SplitWords = sWords
                Exit Function
            Else
                sWords(UBound(sWords)) = Mid$(sText, lPos, lBreak - lPos)
            End If
        End If
        ReDim Preserve sWords(UBound(sWords) + 1)
    
        ' move the pointer on a bit
        lPos = lBreak + 1
        
        ' setup the exit condition
        If lPos >= Len(sText) Then Exit Do
    
    Loop

    ' return the array
    SplitWords = sWords

End Function

'//--[RemoveEOL]------------------------------------------------------------//
'
'  Removes leading and trailing vbCrLf from strings
'
Private Function RemoveEOL(ByVal sText As String) As String
Dim sTmp As String
    ' remove leading or trailing vbCrLf from the string
    sTmp = sText
    If Left$(sTmp, 2) = vbCrLf Then
        sTmp = Right$(sTmp, Len(sTmp) - 2)
    End If
    If Right$(sTmp, 2) = vbCrLf Then
        sTmp = Left$(sTmp, Len(sTmp) - 2)
    End If
    RemoveEOL = sTmp
End Function

'//--[RemoveStrings]-------------------------------------------------------//
'
'  Removes any quoted strings from the text, but only
'  those that aren't within comments of course.
'
Private Function RemoveStrings(ByVal sText As String) As String
Dim lCom As Long
Dim lPos As Long
Dim lPos2 As Long

    lCom = InStr(1, sText, CHAR_COMMENT)
    lPos = InStr(1, sText, Chr$(34))
    If lPos < lCom Or lCom = 0 Then
        Do While lPos <> 0
            ' find the end " char to make a pair
            lPos2 = InStr(lPos + 1, sText, Chr$(34))
            If lPos2 <> 0 Then
                ' we've found a pair, so remove it
                sText = Mid$(sText, 1, lPos - 1) & Mid$(sText, lPos2 + 1)
                ' find the next starting " avoiding
                ' comments within strings
                lCom = InStr(lPos2 + 1, sText, CHAR_COMMENT)
                lPos = InStr(lPos2 + 1, sText, Chr$(34))
                If lPos > lCom Then Exit Do
            Else
                Exit Do
            End If
        Loop
    End If
    
    ' return
    RemoveStrings = sText
    
End Function
