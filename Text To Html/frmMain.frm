VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   Caption         =   " Text To Html Formatter     By Dream   v1.1.0"
   ClientHeight    =   4155
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10605
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   10605
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   8400
      Top             =   360
   End
   Begin SHDocVwCtl.WebBrowser WB1 
      Height          =   2535
      Left            =   5520
      TabIndex        =   0
      ToolTipText     =   " This Window Will Show You A Preview Of Your Text"
      Top             =   720
      Width           =   4935
      ExtentX         =   8705
      ExtentY         =   4471
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   8880
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   7680
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CCA
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1DDC
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1EEE
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2000
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2112
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2224
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2336
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2448
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":255A
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":266C
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":277E
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2890
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":29A2
            Key             =   "Align Right"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   10605
      _ExtentX        =   18706
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New File"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open File"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save Text/Html"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print Text/Html"
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut Text"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy Text"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste Text"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview HTML"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton cmdFormat 
      Caption         =   "Format Text/Code"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton cmdUndo 
      Caption         =   "Undo Format"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   3360
      Width           =   1695
   End
   Begin MSComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   7
      Top             =   3885
      Width           =   10605
      _ExtentX        =   18706
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3598
            MinWidth        =   3598
            Text            =   "Status:"
            TextSave        =   "Status:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10504
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            Object.Width           =   2011
            MinWidth        =   2011
            TextSave        =   "5/20/03"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   2011
            MinWidth        =   2011
            TextSave        =   "4:46 PM"
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox RTB 
      Height          =   2535
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   4471
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmMain.frx":2AB4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblMargin 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Dbl Space Margin: Off"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7080
      TabIndex        =   11
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label lblWrap 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Frame wrap: On"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5640
      TabIndex        =   10
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label lblDSpace 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Dbl Space Txt: Off"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8880
      TabIndex        =   9
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label labHtml 
      AutoSize        =   -1  'True
      Caption         =   "Html:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5520
      TabIndex        =   5
      Top             =   480
      Width           =   450
   End
   Begin VB.Label labData 
      AutoSize        =   -1  'True
      Caption         =   "Text:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   450
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
      End
      Begin VB.Menu d 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLoad 
         Caption         =   "&Load"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu ImALine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         Begin VB.Menu mnuFormatPrint 
            Caption         =   "&Html"
         End
         Begin VB.Menu dbv 
            Caption         =   "-"
         End
         Begin VB.Menu mnuUnformatPrint 
            Caption         =   "&Text"
         End
      End
      Begin VB.Menu aline 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReset 
         Caption         =   "&Reset"
      End
      Begin VB.Menu ImALine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit        (Esc)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu uhhh 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClearAll 
         Caption         =   "Clear &All"
      End
      Begin VB.Menu mnuAll 
         Caption         =   "Select &All"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuStatBar 
         Caption         =   "&StatusBar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuToolBar 
         Caption         =   "&ToolBar"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuInsert 
      Caption         =   "&Insert"
      Begin VB.Menu mnuBold 
         Caption         =   "Bold "
         Begin VB.Menu mnuBoldStart 
            Caption         =   "&Start"
         End
         Begin VB.Menu mnuBoldEnd 
            Caption         =   "&End"
         End
      End
      Begin VB.Menu mnuBreak 
         Caption         =   "Line B&reak"
      End
      Begin VB.Menu pl 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFontColr 
         Caption         =   "Font &Color"
         Begin VB.Menu mnuFCBlue 
            Caption         =   "&Blue"
         End
         Begin VB.Menu mnuFCGreen 
            Caption         =   "&Green"
         End
         Begin VB.Menu mnuFCRed 
            Caption         =   "&Red"
         End
         Begin VB.Menu mnuFCOrange 
            Caption         =   "&Orange"
         End
         Begin VB.Menu mnuFCPrpl 
            Caption         =   "&Purple"
         End
         Begin VB.Menu mnuFCYlw 
            Caption         =   "&Yellow"
         End
      End
      Begin VB.Menu mnuFont 
         Caption         =   "Font &Size"
         Begin VB.Menu mnuFont1 
            Caption         =   "&Size 1"
         End
         Begin VB.Menu mnuFont2 
            Caption         =   "&Size 2"
         End
         Begin VB.Menu mnuFont3 
            Caption         =   "&Size 3"
         End
         Begin VB.Menu mnuFont4 
            Caption         =   "&Size 4"
         End
         Begin VB.Menu mnuFont5 
            Caption         =   "&Size 5"
         End
         Begin VB.Menu mnuFont6 
            Caption         =   "&Size 6"
         End
      End
      Begin VB.Menu mnuFntEnd 
         Caption         =   "En&d Font"
      End
      Begin VB.Menu pop 
         Caption         =   "-"
      End
      Begin VB.Menu mnuT 
         Caption         =   "Table"
         Begin VB.Menu mnuTblStrt 
            Caption         =   "&Start"
         End
         Begin VB.Menu mnuTblEnd 
            Caption         =   "En&d"
         End
      End
   End
   Begin VB.Menu mnuTool 
      Caption         =   "&Tools"
      Begin VB.Menu mnuDbMargin 
         Caption         =   "&Double Space Margin"
      End
      Begin VB.Menu mnuDouble 
         Caption         =   "&Double Space Text"
      End
      Begin VB.Menu mnuFrameWrap 
         Caption         =   "&Frame Wrap On Preview"
         Checked         =   -1  'True
      End
      Begin VB.Menu bv 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormat 
         Caption         =   "&Format Text/Code"
      End
      Begin VB.Menu mnuUndo 
         Caption         =   "&Undo Format"
      End
      Begin VB.Menu wwwww 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPreview 
         Caption         =   "&Preview Html"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuExample 
         Caption         =   "&Example"
      End
      Begin VB.Menu po 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInstruct 
         Caption         =   "Text To Html &Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu sddd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCtact 
         Caption         =   "&Contact"
         Begin VB.Menu mnuRBug 
            Caption         =   "Report &Bug"
         End
      End
      Begin VB.Menu iuooi 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowAbout 
         Caption         =   "A&bout Text To Html"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'***************************************************************************
' Terms of Agreement:
' By using this code, you agree to the following terms...
' 1) You may use this code in your own programs (and may compile it into
' a program and distribute it in compiled format for languages that allow
' it) freely and with no charge.
' 2) You MAY NOT redistribute this code (for example to a web site) without
' written permission from the original author. Failure to do so is a
' violation of copyright laws.
' 3) You may link to this code from another website, but ONLY if it is not
' wrapped in a frame.
' 4) You will abide by any additional copyright restrictions which the
' author may have placed in the code or code's description.
'******************************
' Text To HTML Formatter v1.1.0
'******************************
' Example By Dream
' Date:  6th May 2003
' Email:  baddest_attitude@hotmail.com
'*************************************
' Additional Terms of Agreement:
' This program is free software; you can redistribute
' it and/or modify it under the terms of the GNU General
' Public License as published by the Free Software
' Foundation
' If you make any improvements it would be nice if you would send me a copy.
'******************************
' Text To HTML Formatter v1.1.0
'***************************************************************************

'Thanks to mrk* for his additions/modifications wherever you see this ..
'* Didnt know if he wanted his full name here or not.
'05.2003 mrk Change/Add

Dim DragStartX As Long                'RTB & WB1 resize mouse horizontal start point
Dim WebWidth As Long                  'Width of WB1 before resize
Dim TexterWidth As Long               'width of RTB before resize

Dim LockedWindow As Boolean
Dim DoubledTxt As Boolean             'Lets us know txt spaces doubled(either option)
Dim bDirty As Boolean                 'RichTextBox
Dim ResizeWindows As Boolean          'MouseDown event resize form windows(RTB & WB1)
Dim Wrapped As Boolean                'Tells us if we have previewed and WRAPPED already

Private Const meMinHeight As Single = 5595 '05.2003 mrk Change/Add
Private Const meMinWidth As Single = 10795  '05.2003 mrk Change/Add

Private Sub AddText(ByVal Text As String, _
                    Optional ByVal AddvbCrLf As Boolean = True)  ' Add text to textbox
  If AddvbCrLf Then
     RTB.SelText = Text & vbCrLf           'If last item in textbox tell it boolean false
    Else                                   'So it doesnt add the vbCrLf onto the end for
     RTB.SelText = Text                    'the next new line cause there isnt one
  End If
End Sub '05.2003 mrk Change/Add

Private Sub cmdFormat_Click()                    'Format the text
  mnuFormat_Click
End Sub

Private Sub cmdPreview_Click()                   'Preview formatted text
  Call PreviewHtml(True)                         'call wrapper (boolean for add frame)
End Sub

Private Sub cmdUndo_Click()                      'Undo format
  mnuUndo_Click
End Sub

Private Sub DblMargin()
  On Error GoTo LogError
  Dim i As Integer
  Dim TextLines As Long         'No of lines in RTB
  Dim TextBuff As String
  Dim CharRet As Long
  Dim strLine As String         'Current line in RTB being processed
  Dim Count As Integer          'Current position within strLine
  Dim LineReplacement As String 'String holding NEW Text:
                                'Replace RTB with this when done
' On Error GoTo errorkiller
  Screen.MousePointer = 11
  LockedWindow = True
  SB.Panels(1).Text = "Status: Spacing"               '## Print
  TextLines = SendMessage(RTB.hwnd, &HBA, 0, 0)    'Get number of lines in text box
  For i = 0 To TextLines - 1                           'Extract & print each line in TextBox
    TextBuff = Space(1000)
    Mid(TextBuff, 1, 1) = Chr(79 And &HFF)             'Setup buffer for the line!
    Mid(TextBuff, 2, 1) = Chr(79 \ &H100)
    CharRet = SendMessage(RTB.hwnd, &HC4, i, ByVal TextBuff) 'Get the data from the line
    strLine = Left(TextBuff, CharRet)
    SB.Panels(2).Text = "Line: " & i & " of " & TextLines - 1 & "           "
    For Count = 1 To Len(strLine)
      If Mid(strLine, Count, 1) = " " Then
          LineReplacement = LineReplacement & "&nbsp;&nbsp;"
        Else
          If i = TextLines - 1 Then
              LineReplacement = LineReplacement & Mid(strLine, Count)
            Else
              LineReplacement = LineReplacement & Mid(strLine, Count) ' & vbCrLf
          End If                   ' If you want a spacey look in RTB Uncomment...
          GoTo Done                ' the & vbCrLf above!
      End If
    Next Count
Done:
  Next i
  RTB.Text = LineReplacement
  Screen.MousePointer = 1
  SB.Panels(1).Text = "Status: Done"
  SB.Panels(2).Text = "Margin Formatted           "
  DoubledTxt = True
  LockedWindow = False
  Exit Sub
LogError:
  LockedWindow = False
  Screen.MousePointer = 1
  LockWindowUpdate 0&
  SB.Panels(1).Text = "Status: Margin Error"
  SB.Panels(2).Text = Err.Number & "  " & Err.Description & "           "
  Call LogErrorControl("Double Margin")
End Sub

Private Sub Form_Keyup(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyF1:     mnuInstruct_Click                'Show help form
    Case vbKeyEscape: Unload Me                        'Unload program
  End Select
End Sub '05.2003 mrk Change/Add

Private Sub Form_Load()
  Screen.MousePointer = 11
  'ToDo ~ mnuAbout_Click->
  cmdUndo.Enabled = False          'Disable Undo(s)
  mnuUndo.Enabled = False          'Disable Undo(s)
  Form_Resize                      'Resize Controls
  ProduceHtml                      'Produce blank html for clearing WB1
  InitKeyWords                     'Initialize RTB Keywords
  mnuExample_Click                 'Load Example file into RTB and preview it!
End Sub '05.2003 mrk Change/Add

Private Sub Form_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
  If button = 1 And (X > RTB.Left + RTB.Width) And _
                     (X < WB1.Left) Then
     If button = 1 And (Y > RTB.Top) And _
                     (Y < RTB.Top + RTB.Height) Then
        ResizeWindows = True  'This boolean tells us the mousedown occured between the
        DragStartX = X        'txtdata box and the WB1 browser meaning resize the windows
        WebWidth = WB1.Width
        TexterWidth = RTB.Width
     End If
  End If
End Sub
                        
Private Sub Form_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
  If ResizeWindows = True Then                            'Resizes the text and WB1
    RTB.Width = TexterWidth + (X - DragStartX)
    WB1.Move RTB.Left + RTB.Width + 120, WB1.Top, WebWidth - (X - DragStartX)
    frmMain.Refresh
  End If
  If (X > RTB.Left + RTB.Width) And _
     (X < WB1.Left) And (Y > RTB.Top) And _
     (Y < RTB.Top + RTB.Height) Then              'Reset mousepointer
      Screen.MousePointer = vbSizeWE
     Else
      Screen.MousePointer = vbDefault
  End If
End Sub
                   
Private Sub Form_MouseUp(button As Integer, Shift As Integer, X As Single, Y As Single)
  ResizeWindows = False  'This boolean tells us the mousedown occured within the _
             txtdata box and the WB1 browser meaning resize the windows (Switching Off)
End Sub

Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormCode Then
     'Space for MsgBoxTitle----------,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
     cancel = CBool(MsgBox("Do you wish to exit Text To Html Formatter ?", _
              vbYesNo Or vbInformation Or vbApplicationModal Or vbMsgBoxSetForeground Or vbSystemModal, _
              Me.Caption) = vbNo)
  End If
End Sub '05.2003 mrk Change/Add

Private Sub Form_Resize()                              'Lets resize controls etc
  Dim TxtTop As Integer
  Dim SBHeight As Integer
  Dim SpaceFree As Single
  If Me.WindowState = vbMinimized Then
Exit Sub
  End If
  If Me.Height < meMinHeight Then
     Me.Height = meMinHeight
Exit Sub
  End If
  If Me.Width < meMinWidth Then
     Me.Width = meMinWidth
Exit Sub
  End If
  
  Select Case SB.Visible
    Case True: SBHeight = 270
    Case False: SBHeight = 0
    Case Else: 'Blah Blah @ Microsoft
  End Select
  
  Select Case TB.Visible
    Case True: TxtTop = 720
    Case False: TxtTop = 360
    Case Else: 'Blah Blah @ Microsoft
  End Select
  
  SpaceFree = RTB.Left
  
  With cmdFormat
    .Left = SpaceFree
    .Top = ToValueSngMinIs0(Me.ScaleHeight - SBHeight - .Height - SpaceFree) 'SBHeight p.b
  End With
  With cmdUndo
    .Left = ToValueSngMinIs0(cmdFormat.Left + cmdFormat.Width + SpaceFree)
    .Top = ToValueSngMinIs0(Me.ScaleHeight - SBHeight - .Height - SpaceFree) 'SBHeight p.b
  End With
  With cmdPreview
    .Left = ToValueSngMinIs0(cmdUndo.Left + cmdUndo.Width + SpaceFree)
    .Top = ToValueSngMinIs0(Me.ScaleHeight - SBHeight - .Height - SpaceFree) 'SBHeight p.b
  End With
  
  With RTB
    .Left = SpaceFree
    .Width = ToValueSngMinIs0(Me.ScaleWidth \ 2 - SpaceFree)
    .Height = ToValueSngMinIs0(cmdFormat.Top - TxtTop - SpaceFree)           'TxtTop p.b
  End With
  With WB1
    .Left = ToValueSngMinIs0(Me.ScaleWidth \ 2 + (SpaceFree))
    .Height = ToValueSngMinIs0(cmdFormat.Top - .Top - SpaceFree)
    .Width = ToValueSngMinIs0(Me.ScaleWidth - .Left - SpaceFree)
  End With
  
  With labData
    .Left = SpaceFree
  End With
  With labHtml
    .Left = ToValueSngMinIs0(WB1.Left)
  End With
  
  With lblDSpace
    .Left = ToValueSngMinIs0(WB1.Left + WB1.Width - .Width)
    .Top = ToValueSngMinIs0(Me.ScaleHeight - SBHeight - .Height - SpaceFree) 'SBHeight p.b
  End With
  With lblMargin
    .Left = ToValueSngMinIs0(lblDSpace.Left - .Width)
    .Top = ToValueSngMinIs0(Me.ScaleHeight - SBHeight - .Height - SpaceFree) 'SBHeight p.b
  End With
  With lblWrap
    .Left = ToValueSngMinIs0(lblMargin.Left - .Width)
    .Top = ToValueSngMinIs0(Me.ScaleHeight - SBHeight - .Height - SpaceFree) 'SBHeight p.b
  End With
  
  frmMain.Refresh
End Sub '05.2003 mrk Change/Add

Private Sub Form_Unload(cancel As Integer)                    'Unload application
  Timer1.Enabled = False
  ToFileKill ToAppPath & "preview.htm"                        'Delete the preview html
End
End Sub

Private Function IsControlKey(ByVal KeyCode As Long) As Boolean  'For RTB KeyDown
 'check if the key is a control key
  Select Case KeyCode
    Case vbKeyLeft, vbKeyRight, vbKeyHome, _
         vbKeyEnd, vbKeyPageUp, vbKeyPageDown, _
         vbKeyShift, vbKeyControl
         IsControlKey = True
    Case Else
        IsControlKey = False
  End Select
End Function

Private Sub Insert(sText As String)                          'RTB Text Insertion
  RTB.SelText = sText
End Sub

Private Sub mnuAll_Click()                                   'Select all text
  RTB.SetFocus
  RTB.SelStart = 0
  RTB.SelLength = Len(RTB.Text)
End Sub

Private Sub mnuBoldEnd_Click()
  Insert "</B>"
End Sub

Private Sub mnuBoldStart_Click()
  Insert "<B>"
End Sub

Private Sub mnuBreak_Click()
  Insert "<BR>"
End Sub

Private Sub mnuClearAll_Click()
  RTB.Text = ""                                               'Clear RTB TextBox
End Sub

Private Sub mnuCopy_Click()                                   'Copy selected text
  Clipboard.Clear
  Clipboard.SetText RTB.SelText
End Sub

Private Sub mnuCut_Click()                                    'Cut selected text
  Clipboard.Clear
  Clipboard.SetText RTB.SelText
  RTB.SelText = ""
End Sub

Private Sub mnuDbMargin_Click()                               'Double Space the Margin
  mnuDbMargin.Checked = Not mnuDbMargin.Checked               'Check or not to check
  Select Case mnuDbMargin.Checked
     Case True
       lblMargin.Caption = "Dbl Space Margin: On"
     Case False
       lblMargin.Caption = "Dbl Space Margin: Off"
     Case Else
      'Blah Blah @ Microsoft
  End Select
End Sub

Private Sub mnuDelete_Click()                                 'Delete selected text
  RTB.SelText = ""
End Sub

Private Sub mnuDouble_Click()
  mnuDouble.Checked = Not mnuDouble.Checked                   'Double spacing or not
  Select Case mnuDouble.Checked
     Case True
       lblDSpace.Caption = "Dbl Space Txt: On"
     Case False
       lblDSpace.Caption = "Dbl Space Txt: Off"
     Case Else
      'Blah Blah @ Microsoft
  End Select
End Sub

Private Sub mnuExample_Click()                                'Example Text
  On Error GoTo LogError
  Dim ansa As String
  If RTB.Text <> "" Then
    ansa = MsgBox("This will clear the Text box of current Text, Proceed?", vbYesNo, "Show Example")
    If ansa = vbNo Then GoTo Nope
  End If
   cmdFormat.Enabled = True
   cmdUndo.Enabled = False
   mnuFormat.Enabled = True
   mnuUndo.Enabled = False
   LockedWindow = True
   RTB.Text = ToFileLoad(ToAppPath & "Example.txt")
   ColorIn RTB
   LockedWindow = False
   Call PreviewHtml(False) 'False is No Frame Wrap
   SB.Panels(2).Text = ToAppPath & "Example.txt           "
   Timer1.Enabled = True
Nope:
Exit Sub
LogError:
  Call LogErrorControl("Example Preview")
End Sub

Private Sub mnuExit_Click()                                   'Unload form
  Unload Me
End Sub

Private Sub mnuFCBlue_Click()
  Insert "<FONT COLOR = #000099>"
End Sub

Private Sub mnuFCGreen_Click()
  Insert "<FONT COLOR = #006600>"
End Sub

Private Sub mnuFCOrange_Click()
  Insert "<FONT COLOR = #FF6600>"
End Sub

Private Sub mnuFCPrpl_Click()
  Insert "<FONT COLOR = #990099>"
End Sub

Private Sub mnuFCRed_Click()
  Insert "<FONT COLOR = #CC0000>"
End Sub

Private Sub mnuFCYlw_Click()
  Insert "<FONT COLOR = #FFFF33>"
End Sub

Private Sub mnuFont1_Click()
  Insert "<FONT SIZE = 1>"
End Sub

Private Sub mnuFont2_Click()
  Insert "<FONT SIZE = 2>"
End Sub

Private Sub mnuFont3_Click()
  Insert "<FONT SIZE = 3>"
End Sub

Private Sub mnuFont4_Click()
  Insert "<FONT SIZE = 4>"
End Sub

Private Sub mnuFont5_Click()
  Insert "<FONT SIZE = 5>"
End Sub

Private Sub mnuFont6_Click()
  Insert "<FONT SIZE = 6>"
End Sub

Private Sub mnuFntEnd_Click()
  Insert "</FONT>"
End Sub

Private Sub mnuFormat_Click()
  On Error GoTo LogError
  LockedWindow = True                           'Stop RTB_Change sub
  Screen.MousePointer = 11                      'Format the Text to Html
  cmdFormat.Enabled = False
  cmdUndo.Enabled = True
  mnuFormat.Enabled = False
  mnuUndo.Enabled = True
  RTB.Text = Replace(RTB.Text, vbCrLf, "<br>" & vbCrLf) 'add line breaks
  If mnuDbMargin.Checked = True Then      'If Dbl Space Margin selected then If
    If mnuDouble.Checked = False Then     'double space text selected no need to
      Call DblMargin                      'call the lengthy DblMargin Sub so Skip It.
    End If
  End If
  Select Case mnuDouble.Checked    'Double space the text or not
    Case True
      DoubledTxt = True
     'This next line will replace all spaces with double spaces
      RTB.Text = Replace(RTB.Text, " ", "&nbsp;&nbsp;")
    Case False
     'These next 2 lines will replace any combo of spaces where
     'there are two or more spaces! but remove odds first(no need for single's replacement)!
      RTB.Text = Replace(RTB.Text, "   ", "&nbsp;&nbsp;&nbsp;")
      RTB.Text = Replace(RTB.Text, "  ", "&nbsp;&nbsp;")
    Case Else: 'Blah@microsoft
  End Select
    
  SB.Panels(1).Text = "Status: Formatted"
  ColorIn frmMain.RTB
  Screen.MousePointer = 1
  LockedWindow = False
Exit Sub
LogError:
  Call LogErrorControl("Formatting Function")
End Sub

Private Sub mnuFormatPrint_Click()                           'Print Formatted text
  cmdFormat_Click
  Call PrintOut(True)
End Sub

Private Sub mnuFrameWrap_Click()
  mnuFrameWrap.Checked = Not mnuFrameWrap.Checked
  Select Case mnuFrameWrap.Checked
     Case True
       lblWrap.Caption = "Frame Wrap: On"
     Case False
       lblWrap.Caption = "Frame Wrap: Off"
     Case Else
      'Blah Blah @ Microsoft
  End Select
End Sub

Private Sub mnuInstruct_Click()                              'Show instructions/help
  Dim a As String
  Dim Scr_hDC As Long
 'Get the Desktop handle
  Scr_hDC = GetDesktopWindow()
  a = ShellExecute(Scr_hDC, "Open", ToAppPath & "TtoH.hlp", "", "", SW_SHOWNORMAL)
  If a = "2" Then
    a = MsgBox("Text To Html Help Files Could Not Be Found," & vbCrLf & vbCrLf & "Do You Wish To Browse For Them?", vbYesNo, "Cannot Find Help")
    If a = vbYes Then
      With CD1
       .DialogTitle = "Load Code/Tutorial/Article"
       .CancelError = True
       .FileName = ""
       .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist
       .Filter = "All File(*.*)|*.*"     'Sets the filetypes viewed
        If ToDlgOpen(CD1) Then            'If this call returns a filename then..
          a = ShellExecute(Scr_hDC, "Open", .FileName, "", "", SW_SHOWNORMAL)
        End If
      End With
    End If
  End If
  End Sub

Private Sub mnuLoad_Click()                                  'Load file
Dim Char As String
  On Error GoTo LogError
  With CD1
    .DialogTitle = "Load Code/Tutorial/Article"
    .CancelError = True
    .FileName = ""
    .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist
    .Filter = "All File(*.*)|*.*"     'Sets the filetypes viewed
    If ToDlgOpen(CD1) Then            'If this call returns a filename then..
       If Reset = True Then
         LockedWindow = True          'Tell RTB change to cease while loading file
         Screen.MousePointer = 11
         SB.Panels(1).Text = "Status: Loading File"
         Char = Mid(.FileName, Len(.FileName) - 2)
        'Check filetype for .frm .bas & .cls
        'If VBDoc then tell ToFileLoad what kinda form!
         Select Case Char
        'You can customize the filetypes and the load procedure in the ToFile.bas
        'to load just about any filetype you wish (text type files)
           Case "frm", "bas", "cls": RTB.Text = ToFileLoad(.FileName, VBDoc) 'if a VBDoc then filter i
           Case Else: RTB.Text = ToFileLoad(.FileName)   'or ToFileLoad(.FileName, Default)
         End Select
         ColorIn RTB
         Timer1.Enabled = True
         LockedWindow = False
         Call PreviewHtml(False)
         Screen.MousePointer = 1
         SB.Panels(1).Text = "Status: File Loaded: "
         SB.Panels(2).Text = .FileName & "    "
       End If
    End If
  End With
Exit Sub
LogError:
  Call LogErrorControl("Loading File")
End Sub

Private Sub mnuNew_Click()                                  'Load New File
  If Reset = True Then
    SB.Panels(1).Text = "Status: New File"
    SB.Panels(2).Text = "FileName: None         "
    RTB.SetFocus
  End If
End Sub

Private Sub mnuPaste_Click()                               'Paste text
  On Error Resume Next
  Dim strData As String
  strData = Clipboard.GetText(vbCFText)
  RTB.SelText = strData
End Sub

Private Sub mnuPreview_Click()
  Call PreviewHtml(True)
End Sub

Private Sub mnuRBug_Click()
  OpenLink frmMain, "Mailto:dream@dream-domain.net?Subject=TtoH Bug Report"
End Sub

Private Sub mnuReset_Click()                               'Clear Program
   Reset
End Sub

Private Sub mnuSave_Click()                            'Save Text To File
  On Error GoTo LogError
  With CD1
    .DialogTitle = "Save Formatted Code/Tutorial/Article"
    .CancelError = True
    .FileName = ""
    .Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
    .Filter = "Text File (*.txt)|*.txt"      'Sets the filetypes viewed
    If ToDlgShow(CD1) Then                   'If this call returns a filename then..
       Screen.MousePointer = 11
       ToFileSave .FileName, RTB.Text        'Call the save file function
       SB.Panels(1).Text = "Status: File Saved: "
       SB.Panels(2).Text = .FileName & "          "
       Screen.MousePointer = 1
    End If
  End With
Exit Sub
LogError:
  Call LogErrorControl("Saving File")
End Sub '05.2003 mrk Change/Add

Private Sub mnuShowAbout_Click()                       'Display About Window
  ToDlgAbout Me, _
             "    " & App.ProductName, _
             "    Example By " & App.CompanyName & ", Version " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & "    " & App.LegalCopyright, _
             Me.Icon
End Sub '05.2003 mrk Change/Add

Private Sub mnuStatBar_Click()                         'Hide/Show Status Bar
  mnuStatBar.Checked = Not mnuStatBar.Checked
  Select Case mnuStatBar.Checked
    Case 1
      SB.Visible = True
      frmMain.Height = frmMain.Height + 270
    Case 0
      SB.Visible = False
      frmMain.Height = frmMain.Height - 270
    Case Else
     'Blah Blah @ Microsoft
  End Select
End Sub

Private Sub mnuTblEnd_Click()
  Insert "</table>"
End Sub

Private Sub mnuTblStrt_Click()
  Dim Inputval As String
  On Error GoTo LogError
  Inputval = InputBox("What width do you want the frame to be?", "Frame Width", "850")
  Do Until IsNumeric(Inputval) = True          ' Check to see numeric values entered!
    Inputval = InputBox("Please Enter A Numeric Value!", "User Input", "")
    ' Check the user entered a valid value
    If IsNumeric(Inputval) = True Then Exit Do
  Loop
    Insert "<TABLE border=0 width=" & Inputval & ">"
Exit Sub
LogError:
  Call LogErrorControl("Formatting Function")
End Sub

Private Sub mnuToolBar_Click()                         'Hide/Show Tool Bar
  mnuToolBar.Checked = Not mnuToolBar.Checked
  Select Case mnuToolBar.Checked
    Case 1
      TB.Visible = True                                'Show toolbar and resize
      WB1.Top = WB1.Top + 360                          'or repostion some controls
      RTB.Top = RTB.Top + 360
      labData.Top = labData.Top + 360
      labHtml.Top = labHtml.Top + 360
      frmMain.Height = frmMain.Height + 360
    Case 0
      TB.Visible = False                               'Hide toolbar and resize
      WB1.Top = WB1.Top - 360                          'or repostion some controls
      RTB.Top = RTB.Top - 360
      labData.Top = labData.Top - 360
      labHtml.Top = labHtml.Top - 360
      frmMain.Height = frmMain.Height - 360
    Case Else
     'Blah Blah @ Microsoft
  End Select
End Sub

Private Sub mnuUndo_Click()                            'Undo Format
  cmdFormat.Enabled = True
  cmdUndo.Enabled = False
  mnuFormat.Enabled = True
  mnuUndo.Enabled = False
  LockedWindow = True
  RTB.Text = Replace(RTB.Text, "<br>", "")     'Replace line breaks
  If DoubledTxt = True Then RTB.Text = Replace(RTB.Text, "&nbsp;&nbsp;", " ") 'Replace &nbsp;'s with space
  RTB.Text = Replace(RTB.Text, "&nbsp;", " ")  'Replace &nbsp; with space
  SB.Panels(1).Text = "Status: Format Undone"
  LockedWindow = False
Exit Sub
LogError:
  Call LogErrorControl("Undo Function")
End Sub

Private Sub mnuUnFormatPrint_Click()                   'Print Text
  cmdUndo_Click
  Call PrintOut(False)
End Sub

Private Sub PrintOut(Formatted As Boolean)             'Print Out Text or Html
  On Error GoTo LogError
  Dim i As Integer                                     'Text = false : Html = True
  Dim ta As Long
  Dim TextLines As Long
  Dim TextBuff As String
  Dim CharRet As Long
  On Error GoTo errorkiller
  Screen.MousePointer = 11
  SB.Panels(1).Text = "Status: Printing"               '## Print
  Printer.Print " "
  Printer.Print , , "Text To Html Formatter"
  Printer.Print " "
  Printer.Print , , Now
  ta = SetTextAlign(Printer.hdc, TA_CENTER)            'Center text on printer object
  Printer.CurrentY = (Printer.ScaleHeight / RTB.Parent.ScaleHeight) * RTB.Top
  TextLines = SendMessage(RTB.hwnd, &HBA, 0, 0)    'Get number of lines in text box
  For i = 0 To TextLines - 1                           'Extract & print each line in TextBox
    TextBuff = Space(1000)
    Printer.CurrentX = (Printer.ScaleWidth / RTB.Parent.ScaleWidth) * (RTB.Left + (RTB.Width / 2))
    Mid(TextBuff, 1, 1) = Chr(79 And &HFF)             'Setup buffer for the line!
    Mid(TextBuff, 2, 1) = Chr(79 \ &H100)
    CharRet = SendMessage(RTB.hwnd, &HC4, i, ByVal TextBuff)
    Printer.Print Left(TextBuff, CharRet)
  Next i
  ta = SetTextAlign(Printer.hdc, ta)        'Reset alignment back to original setting
  Printer.EndDoc
  Screen.MousePointer = 1
  SB.Panels(1).Text = "Status: Idle"
  SB.Panels(2).Text = "Finished Sending Data To Printer"
  Exit Sub
errorkiller:
  Screen.MousePointer = 1
   SB.Panels(1).Text = "Status: Error Printing"
   SB.Panels(2).Text = Err.Number & "   " & Err.Description & "        "
Exit Sub
LogError:
  Call LogErrorControl("Print Function")
End Sub

Private Sub ProduceHtml()                                      'Produce a blank html
  On Error GoTo LogError
  If ToFileIsExist(ToAppPath & "blank.htm") = False Then       'If doesnt exist then..
    Dim intFreeFile As Integer
    intFreeFile = FreeFile
    Open ToAppPath & "blank.htm" For Output As #intFreeFile    'Create it!
      Print , " "
    Close #intFreeFile
  End If
Exit Sub
LogError:
  Call LogErrorControl("Produce Blank Html")
End Sub

Private Function Reset() As Boolean
  On Error GoTo LogError
  If RTB.Text <> "" Then
    Dim ansa As String                  'If text present then ask to save?
    ansa = MsgBox("You will lose Unsaved Information, Continue?", vbYesNo, "Save Before Continuing")
    If ansa = vbNo Then
    GoTo EndMe
    End If
  End If
 'Now its safe to clear the form
  Wrapped = False  'After wrapping once shut down function(this resets it)
  cmdFormat.Enabled = True
  cmdUndo.Enabled = False
  mnuFormat.Enabled = True
  mnuUndo.Enabled = False
  RTB.Text = ""
  SB.Panels(1).Text = "Status: Reset"
  SB.Panels(2).Text = "       "
  WB1.Navigate ToAppPath & "blank.htm"   'Clear WB1 by loading blank html
  LockedWindow = False
Reset = True
Exit Function
EndMe:
Reset = False
Exit Function
LogError:
  Call LogErrorControl("Undo Function")
End Function

Private Sub RTB_change()                  'RTB Text changed so lets re color it
  If LockedWindow = True Then Exit Sub
  Dim Position As Long
  Position = RTB.SelStart
  ColorIn frmMain.RTB, Position
End Sub

Private Sub RTB_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim lCursor As Long
  Dim lSelectLen As Long
  Dim lStart As Long
  Dim lFinish As Long
  Dim sText As String
  On Error Resume Next                      'Here's the on the fly coloring
  Screen.MousePointer = vbDefault
  If KeyCode = vbKeyC And Shift = 2 Then Exit Sub  'check for Ctrl+C
  If KeyCode = vbKeyV And Shift = 2 Then  'check for text being pasted into the box
    Screen.MousePointer = vbHourglass
    DoClipBoardPaste RTB
    KeyCode = 0
    Screen.MousePointer = vbNormal
    Exit Sub
  End If
 'if the cursor is moving to a different
 'line then process the orginal line
  If KeyCode = 13 Or _
    KeyCode = vbKeyUp Or _
    KeyCode = vbKeyDown Then
    If bDirty Or KeyCode = 13 Then 'only color this line if it's been changed
      LockWindowUpdate RTB.hwnd  'lock the window to cancel out flickering
      lCursor = RTB.SelStart      'store the current cursor pos
      lSelectLen = RTB.SelLength  'and current selection if there is any
      If lCursor <> 0 Then        'get the line start and end
        lStart = InStrRev(RTB.Text, vbCrLf, RTB.SelStart) + 2
        If lStart = 2 Then lStart = 1
       Else
        lStart = 1
      End If
      lFinish = InStr(lCursor + 1, RTB.Text, vbCrLf)
      If lFinish = 0 Then lFinish = Len(RTB.Text)
      basColor.sText = RTB.Text            'do the coloring
      DoColor RTB, lStart, lFinish
     'if ENTER was pressed, we should color the next line
     'as well, so that if a line is broken by the ENTER
     'the new line and the old line are colored properly
      If KeyCode = 13 Then
        lStart = lCursor + 1
        lFinish = InStr(lStart, RTB.Text, vbCrLf)
        If lFinish = 0 Then lFinish = Len(RTB.Text)
        If lStart - 1 <> lFinish Then  'only color if another line exists
          RTB.SelStart = lStart - 1
          RTB.SelLength = lFinish - lStart
          RTB.SelColor = vbBlack
          DoColor RTB, lStart, lFinish
        End If
      End If
      RTB.SelStart = lCursor      'reset the properties
      RTB.SelLength = lSelectLen
      RTB.SelColor = vbBlack
      bDirty = False              'reset the flag and release the window
      LockWindowUpdate 0&
    End If
   ElseIf Not IsControlKey(KeyCode) Then
   'a different key was pressed - and
   'this will alter the line so it
   'needs recoloring when we move off it
    If Not bDirty Then
      LockWindowUpdate RTB.hwnd
      lStart = InStrRev(RTB.Text, vbCrLf, RTB.SelStart + 1) + 1 'get the line start & end
      lFinish = InStr(RTB.SelStart + 1, RTB.Text, vbCrLf)
      If lFinish = 0 Then lFinish = Len(RTB.Text)
      lCursor = RTB.SelStart     'color the line (remembering the cursor position)
      lSelectLen = RTB.SelLength
      RTB.SelStart = lStart
      RTB.SelLength = lFinish - lStart
      RTB.SelColor = vbBlack
      RTB.SelStart = lCursor
      RTB.SelLength = lSelectLen
      bDirty = True
      LockWindowUpdate 0&
    End If
  End If
End Sub

Private Sub TB_ButtonClick(ByVal button As MSComctlLib.button) 'Toolbar click
  Select Case button.Key                            ' Read them for what they say
    Case "New": mnuNew_Click
    Case "Open": mnuLoad_Click
    Case "Save": mnuSave_Click
    Case "Print"
      frmMain.PopupMenu frmMain.mnuPrint
    Case "Cut"
      mnuCut_Click
    Case "Copy"
      mnuCopy_Click
    Case "Paste"
      mnuPaste_Click
    Case Else
     'Blah Blah @ Microsoft
  End Select
End Sub

Private Sub Timer1_Timer()                               'Scroll Timer
    With SB.Panels(2)
     .Text = ToStrScrollToLeft(.Text)                    'Call textscroll function
    End With
End Sub '05.2003 mrk Change/Add

Private Function PreviewHtml(Optional Wrap As Boolean = True)
  On Error GoTo LogError
  If mnuFrameWrap.Checked = True Then        'If Frame Wrap On Preview is selected
    If Wrap = True Then                      'If preview not example from frmmain_load
      If Wrapped = False Then                'If current Text already wrapped then skip
        RTB.SelStart = 0                     'Set RTB Cursor to beginning of RTB
        mnuTblStrt_Click                     'Table start at beginning of RTB
        RTB.SelStart = Len(RTB.Text) + 1     'Set RTB Cursor to end of RTB
        mnuTblEnd_Click                      'Table end at end of RTB
        RTB.SelStart = 0                     'Set RTB Cursor back to beginning of RTB
        Wrapped = True                       'Set the Wrapped boolean to True
      End If
    End If
  End If

  ToFileSave ToAppPath & "preview.htm", RTB.Text           'Save to file
  WB1.Navigate ToAppPath & "preview.htm"                   'Preview HTML
  If cmdFormat.Enabled = False Then                        'Check cmdFormat state
    SB.Panels(1).Text = "Status: Html Previewed"           'if enabled text unformatted
   Else
    SB.Panels(1).Text = "Status: Text Previewed"           'if disabled text formatted
  End If
  SB.Panels(2).Text = "        "
  Screen.MousePointer = 1
Exit Function
LogError:
  Call LogErrorControl("Preview Html")
End Function
