VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Bill's HTML Cleaner"
   ClientHeight    =   8025
   ClientLeft      =   165
   ClientTop       =   465
   ClientWidth     =   8640
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   535
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   576
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame9 
      Caption         =   "Character Encoding"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   34
      Top             =   1680
      Width           =   8415
      Begin VB.CheckBox chkNumericEntities 
         Caption         =   "Numeric entities (#233;)"
         Height          =   255
         Left            =   6000
         TabIndex        =   41
         Top             =   240
         Width           =   2175
      End
      Begin VB.OptionButton CharEncodingOption 
         Caption         =   "Macroman"
         Height          =   255
         Index           =   5
         Left            =   4320
         TabIndex        =   40
         Tag             =   "block"
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton CharEncodingOption 
         Caption         =   "Iso2022"
         Height          =   255
         Index           =   4
         Left            =   3360
         TabIndex        =   39
         Tag             =   "block"
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton CharEncodingOption 
         Caption         =   "Utf-8"
         Height          =   255
         Index           =   3
         Left            =   2640
         TabIndex        =   38
         Tag             =   "block"
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton CharEncodingOption 
         Caption         =   "Latin1"
         Height          =   255
         Index           =   2
         Left            =   1800
         TabIndex        =   37
         Tag             =   "block"
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton CharEncodingOption 
         Caption         =   "Ascii"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   36
         Tag             =   "auto"
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton CharEncodingOption 
         Caption         =   "Raw"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   35
         Tag             =   "none"
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Cleaned HTML"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   25
      Top             =   4560
      Width           =   8415
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   495
         Left            =   7440
         TabIndex        =   29
         Top             =   2520
         Visible         =   0   'False
         Width           =   735
         ExtentX         =   1296
         ExtentY         =   873
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
      Begin RichTextLib.RichTextBox text2 
         Height          =   2535
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   4471
         _Version        =   393217
         Enabled         =   -1  'True
         HideSelection   =   0   'False
         ScrollBars      =   3
         TextRTF         =   $"Form1.frx":0442
      End
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   3015
         Left            =   7200
         TabIndex        =   28
         Top             =   240
         Width           =   1095
         Begin VB.Frame Frame8 
            BorderStyle     =   0  'None
            Caption         =   "Frame6"
            Height          =   3015
            Left            =   0
            TabIndex        =   31
            Top             =   0
            Width           =   1095
            Begin VB.CommandButton cmdKeepBody 
               Caption         =   "Keep Body"
               Height          =   255
               Left            =   0
               TabIndex        =   45
               Top             =   1080
               Width           =   1095
            End
            Begin VB.CommandButton cmdCopyResults 
               Caption         =   "Copy"
               Height          =   255
               Left            =   0
               TabIndex        =   44
               Top             =   720
               Width           =   1095
            End
            Begin VB.CommandButton cmdClearResults 
               Caption         =   "Clear"
               Height          =   255
               Left            =   0
               TabIndex        =   43
               Top             =   360
               Width           =   1095
            End
            Begin VB.CommandButton cmdValidateResults 
               Caption         =   "Validate"
               Height          =   255
               Left            =   0
               TabIndex        =   42
               Top             =   0
               Width           =   1095
            End
            Begin VB.CheckBox chkShowBrowser 
               Caption         =   "Show Browser"
               Height          =   735
               Left            =   0
               Style           =   1  'Graphical
               TabIndex        =   32
               Top             =   1560
               Width           =   1095
            End
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cleaning Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   23
      Top             =   0
      Width           =   8415
      Begin VB.CheckBox CheckBoxes 
         Caption         =   "Remove Word 2000 Formatting"
         Height          =   255
         Index           =   1
         Left            =   3720
         TabIndex        =   2
         Top             =   360
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CheckBox CheckBoxes 
         Caption         =   "Line Break before <BR>"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CheckBox CheckBoxes 
         Caption         =   "Output as XML"
         Height          =   255
         Index           =   6
         Left            =   6480
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
      Begin VB.CheckBox CheckBoxes 
         Caption         =   "Remove double linebreaks"
         Height          =   255
         Index           =   5
         Left            =   6000
         TabIndex        =   6
         Top             =   600
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CheckBox CheckBoxes 
         Caption         =   "Create Styles"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   1335
      End
      Begin VB.CheckBox CheckBoxes 
         Caption         =   "Drop Empty Paragraphs"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox CheckBoxes 
         Caption         =   "Drop Font tags"
         Height          =   255
         Index           =   4
         Left            =   3360
         TabIndex        =   5
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Doc -Type Declaration"
      Height          =   615
      Left            =   4200
      TabIndex        =   26
      Top             =   975
      Width           =   4335
      Begin VB.OptionButton doctypeOption 
         Caption         =   "loose"
         Height          =   255
         Index           =   3
         Left            =   3240
         TabIndex        =   14
         Tag             =   "loose"
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton doctypeOption 
         Caption         =   "strict"
         Height          =   255
         Index           =   2
         Left            =   2280
         TabIndex        =   13
         Tag             =   "strict"
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton doctypeOption 
         Caption         =   "auto"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   12
         Tag             =   "auto"
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton doctypeOption 
         Caption         =   "Omit"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Tag             =   "omit"
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Source HTML"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   24
      Top             =   2400
      Width           =   8415
      Begin RichTextLib.RichTextBox Text1 
         Height          =   1695
         Left            =   120
         TabIndex        =   33
         Top             =   360
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   2990
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"Form1.frx":04C4
      End
      Begin VB.Frame frame6 
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   1815
         Left            =   7200
         TabIndex        =   27
         Top             =   120
         Width           =   1095
         Begin VB.CommandButton cmdClean 
            Caption         =   "C&lean"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   0
            TabIndex        =   18
            Top             =   1320
            Width           =   1095
         End
         Begin VB.CommandButton cmdValidateSource 
            Caption         =   "Validate"
            Height          =   255
            Left            =   0
            TabIndex        =   15
            Top             =   120
            Width           =   1095
         End
         Begin VB.CommandButton cmdClearSource 
            Caption         =   "Clear"
            Height          =   255
            Left            =   0
            TabIndex        =   16
            Top             =   480
            Width           =   1095
         End
         Begin VB.CommandButton cmdCopySource 
            Caption         =   "Copy"
            Height          =   255
            Left            =   0
            TabIndex        =   17
            Top             =   840
            Width           =   1095
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Indenting"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   21
      Top             =   975
      Width           =   3975
      Begin VB.OptionButton IndentingOption 
         Caption         =   "None"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Tag             =   "none"
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton IndentingOption 
         Caption         =   "Block"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   8
         Tag             =   "auto"
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton IndentingOption 
         Caption         =   "Auto"
         Height          =   255
         Index           =   2
         Left            =   1800
         TabIndex        =   9
         Tag             =   "block"
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox IndentSpaces 
         Height          =   285
         Left            =   2760
         TabIndex        =   10
         Text            =   "4"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "spaces."
         Height          =   255
         Left            =   3240
         TabIndex        =   22
         Top             =   270
         Width           =   615
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9840
      Top             =   -120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   576
      TabIndex        =   20
      Top             =   7530
      Width           =   8640
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   6960
         TabIndex        =   19
         Top             =   30
         Width           =   1455
      End
   End
   Begin VB.Menu MeFile 
      Caption         =   "File"
      Begin VB.Menu MiOpen 
         Caption         =   "Open File"
      End
      Begin VB.Menu MiOpenWebPage 
         Caption         =   "Open Web Page"
      End
      Begin VB.Menu ms1 
         Caption         =   "-"
      End
      Begin VB.Menu MiClean 
         Caption         =   "Clean"
      End
      Begin VB.Menu ms2 
         Caption         =   "-"
      End
      Begin VB.Menu MiExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu MeHelp 
      Caption         =   "Help"
      Begin VB.Menu MiAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This project was built as a GUI front end to Andre Blavier's wrapper around the TidyCom Dll.
'Developer: Bill MacIntyre, New Brunswick Canada
'E-mail: billymac@rogers.com
'It cleans up HTML source and is a great MS Word HTML cleaner as well.
'It is also useful as a block indenting tool for HTML source.

'TidyCOM is a Windows COM component wrapping Dave Raggett's HTML Tidy, a free utility
'application from the World Wide Web Consortium that helps you clean up your web pages.
'HTML Tidy is available from the W3C as a command-line program, à la Unix.
'To better fit in the Windows environment Andre has written COM component wrapper
'for Tidy available here:
'http://perso.wanadoo.fr/ablavier/TidyCOM/

'The Tidy SourceForge object is here:
'http://tidy.sourceforge.net/



Dim LastCharTyped As Integer
Private Sub CharEncodingOption_Click(Index As Integer)

    For n = 0 To CharEncodingOption.Count - 1
        If CharEncodingOption(n).Value = True Then
            optCharacterEncoding = n
            Exit For
        End If
    Next
End Sub

Private Sub CheckBoxes_Click(Index As Integer)
    Select Case Index
        Case 0: optClean = CheckBoxes(Index).Value
        Case 1: optWord2000 = CheckBoxes(Index).Value
        Case 2: optBreakBeforeBr = CheckBoxes(Index).Value
        Case 3: optDropEmptyParas = CheckBoxes(Index).Value
        Case 4: optDropFontTags = CheckBoxes(Index).Value
        Case 5: optDoubleLineBreaks = CheckBoxes(Index).Value
        Case 6: optOutputXml = CheckBoxes(Index).Value
    End Select
End Sub

Private Sub chkNumericEntities_Click()
    optNumericEntities = chkNumericEntities.Value
End Sub

Private Sub chkShowBrowser_Click()
    WebBrowser1.Visible = chkShowBrowser.Value = 1

    If chkShowBrowser.Value = 1 Then
        WebBrowser1.Navigate "about:blank"
    End If

    text2.Refresh

End Sub

Private Sub cmdClean_Click()
    text2.Text = Clean(Text1.Text)
    Call ColorizeRTF(text2)
End Sub

Private Sub cmdClearResults_Click()
    text2.Text = ""
End Sub

Private Sub cmdClearSource_Click()
    Text1.Text = ""
End Sub

Private Sub cmdCopyResults_Click()
    Clipboard.Clear
    Clipboard.SetText text2.Text
End Sub

Private Sub cmdCopySource_Click()
    Clipboard.Clear
    Clipboard.SetText Text1.Text
End Sub

Private Sub cmdExit_Click()
    End
End Sub


Private Sub cmdKeepBody_Click()
    'keep only body
    'could do this with reg expressions but I'm too lazy
    'remove all up to <body.....>
    PosN = InStr(1, text2.Text, "<body", vbTextCompare)
    If PosN > 0 Then
        Posnn = InStr(PosN + 1, text2.Text, ">", vbTextCompare)
    End If
    If Posnn > PosN Then
        text2.Text = Right$(text2.Text, Len(text2.Text) - Posnn)
    End If
    'remove all starting at </body>
    PosN = InStr(1, text2.Text, "</body>", vbTextCompare)
    If PosN > 0 Then
        text2.Text = Left$(text2.Text, PosN - 1)
    End If
End Sub

Private Sub cmdValidateResults_Click()
    frmValidate.strValidate = text2.Text
    frmValidate.Show vbModal, Me
End Sub

Private Sub cmdValidateSource_Click()
    frmValidate.strValidate = Text1.Text
    frmValidate.Show vbModal, Me
End Sub

Private Sub doctypeOption_Click(Index As Integer)
    optDoctype = doctypeOption(Index).Tag
End Sub

Private Sub Form_Load()
    optCharEncoding = "UTF-8"
    optClean = False
    optWord2000 = True
    optBreakBeforeBr = True
    optDropEmptyParas = False
    optFixBackslash = True
    optShowWarnings = True
    optIndent = 2
    optDoctype = "omit"
    optQuoteAmpersand = False
    optIndentSpaces = 4
    optWrap = True
    optDropFontTags = False
    optEncloseBlockText = False
    optEncloseText = False
    optTidyMark = False
    optTabSize = 4
    optDoubleLineBreaks = True
    optOutputXml = False
    optCharacterEncoding = 1
    optNumericEntities = False

End Sub

Private Sub Form_Resize()

    On Error Resume Next
    Me.ScaleMode = 3
    Frame5.Width = Me.ScaleWidth - Frame5.Left - (Frame2.Left * 2) + 8
    Frame2.Width = Me.ScaleWidth - (Frame2.Left * 2)
    Frame3.Width = Me.ScaleWidth - (Frame3.Left * 2)
    Frame4.Width = Me.ScaleWidth - (Frame4.Left * 2)
    text2.Height = Me.ScaleHeight - text2.Top - Picture1.Height - 10

    CmdExit.Left = Me.ScaleWidth - (Text1.Left \ 15) - CmdExit.Width
    Frame4.Height = Me.ScaleHeight - Frame4.Top - Picture1.Height
    Text1.Height = (Frame3.Height * 15) - 350
    text2.Height = (Frame4.Height * 15) - 400
    Text1.Width = (Frame3.Width * 15) - 250 - cmdValidateSource.Width
    text2.Width = (Frame4.Width * 15) - 250 - cmdValidateSource.Width
    frame6.Left = Frame3.Width * 15 - frame6.Width - 60
    Frame7.Left = Frame4.Width * 15 - Frame7.Width - 60
    Frame9.Width = Me.ScaleWidth - (Frame9.Left * 2)

    WebBrowser1.Top = text2.Top
    WebBrowser1.Left = text2.Left
    WebBrowser1.Width = text2.Width
    WebBrowser1.Height = text2.Height

End Sub

Private Sub IndentingOption_Click(Index As Integer)
    optIndent = Index
End Sub

Private Sub IndentSpaces_Change()
    IndentSpaces.Text = Val(IndentSpaces)
    optIndentSpaces = IndentSpaces.Text
End Sub

Private Sub MiAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub MiClean_Click()
    text2.Text = Clean(Text1.Text)
End Sub

Private Sub MiExit_Click()
    cmdExit_Click
End Sub

Private Sub MiOpen_Click()
    On Error Resume Next
    CommonDialog1.FileName = ""
    CommonDialog1.CancelError = True
    CommonDialog1.DefaultExt = ".html"
    CommonDialog1.Filter = "html files|*.html;*.htm|Cold Fusion Files|*.cfm|ASP files|*.asp|All files|*.*"
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName = "" Then Exit Sub

    Text1.Text = OpenFile(CommonDialog1.FileName)

End Sub
Function OpenFile(ByVal strFilePath As String) As String
    On Error GoTo errs
    Dim TextLine
    Open strFilePath For Input As #1                       ' Open file.
    Do While Not EOF(1)                                    ' Loop until end of file.
        Line Input #1, TextLine                            ' Read line into variable.
        OpenFile = OpenFile & TextLine & vbNewLine
    Loop
    Close #1                                               ' Close file.

    Exit Function
errs:
End Function


Private Sub MiOpenWebPage_Click()

    Dim http As MSXML2.XMLHTTP
    Set http = New MSXML2.XMLHTTP
    URL = InputBox("Enter the web site URL", "Enter URL", "http://")
    http.open "get", URL, False
    http.send

    Text1.Text = http.responseText
    Call ColorizeRTF(Text1)

End Sub

Private Sub text2_Change()
    Dim Colorize As Boolean
    'only colorize on certain keypresses

    If LastCharTyped = 17 And Shift = 0 Then Colorize = True
    If InStr("|8|187|189|190|13|", "|" & LastCharTyped & "|") > 0 Then Colorize = True

    If Colorize Then Call ColorizeRTF(text2)
End Sub

Private Sub text2_KeyDown(KeyCode As Integer, Shift As Integer)
    LastCharTyped = KeyCode
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    On Error Resume Next
    Set myDoc = WebBrowser1.Document
    myDoc.Write text2.Text
    myDoc.Close
End Sub

Sub ColorizeRTF(ByRef objTextBox As Object)

    PosN = objTextBox.SelStart
    objTextBox.TextRTF = DoColorize(objTextBox.Text)
    objTextBox.SelStart = PosN
End Sub
