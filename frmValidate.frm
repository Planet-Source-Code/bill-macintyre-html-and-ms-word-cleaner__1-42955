VERSION 5.00
Begin VB.Form frmValidate 
   Caption         =   "Validation Report"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   8580
   Icon            =   "frmValidate.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5070
   ScaleWidth      =   8580
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   3375
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   120
      Width           =   8175
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   4680
      TabIndex        =   0
      Top             =   3840
      Width           =   3735
      Begin VB.CommandButton cmdExit 
         Caption         =   "&Ok"
         Height          =   375
         Left            =   2400
         TabIndex        =   1
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   855
      Left            =   0
      TabIndex        =   2
      Top             =   3600
      Width           =   3975
      Begin VB.OptionButton optShow 
         Caption         =   "Show Tidied Source"
         Height          =   255
         Index           =   2
         Left            =   1800
         TabIndex        =   9
         Top             =   600
         Width           =   2175
      End
      Begin VB.OptionButton optShow 
         Caption         =   "Show Source"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   8
         Top             =   360
         Width           =   2175
      End
      Begin VB.OptionButton optShow 
         Caption         =   "Show Validation Report"
         Height          =   255
         Index           =   0
         Left            =   1800
         TabIndex        =   7
         Top             =   120
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Show Comments"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   1695
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Show Warnings"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Show Errors"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Value           =   1  'Checked
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmValidate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strValidate As String
Dim Report As String

Private Sub Check1_Click()
    Validate
End Sub

Private Sub Check2_Click()
    Validate
End Sub

Private Sub Check3_Click()
    Validate
End Sub


Function ShowSource() As String
    ShowSource = strValidate
    'add line numbers
    n = 1
    ShowSource = "|1|" & vbTab & ShowSource
    Do
        n = n + 1
        If InStr(1, ShowSource, vbNewLine) = 0 Then Exit Do
        ShowSource = Replace(ShowSource, vbNewLine, "[CRLF]|" & n & "|" & vbTab, 1, 1, vbTextCompare)
    Loop
    ShowSource = Replace(ShowSource, "[CRLF]", vbNewLine)
End Function

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Validate()

    Text1.Text = ""

    CleanedSource = Clean(strValidate)

    If Check1.Value = 1 Then
        strvalidation = vbNewLine & "comments:" & vbNewLine
        strvalidation = strvalidation & tidy.Comments & vbNewLine
    End If

    If Check2.Value = 1 Then
        strvalidation = strvalidation & vbNewLine & "Warnings (" & tidy.TotalWarnings & "):" & vbNewLine
        For n = 0 To tidy.TotalWarnings - 1
            strvalidation = strvalidation & tidy.Warning(n) & vbNewLine
        Next
    End If

    If Check3.Value = 1 Then
        strvalidation = strvalidation & vbNewLine & "Errors (" & tidy.TotalErrors & "):" & vbNewLine
        For n = 0 To tidy.TotalErrors - 1
            strvalidation = strvalidation & tidy.Error(n) & vbNewLine
        Next
    End If
    Report = strvalidation

    If optShow(0).Value = True Then Text1.Text = Report
    If optShow(1).Value = True Then Text1.Text = ShowSource
    If optShow(2).Value = True Then Text1.Text = CleanedSource

End Sub

Private Sub Form_Load()
    Validate
    Set tidy = New TidyCOM.TidyObject
End Sub

Private Sub Form_Resize()
    Text1.Width = Me.ScaleWidth - (Text1.Left * 2)
    Frame1.Left = Me.ScaleWidth - Frame1.Width
    Frame1.Top = Me.ScaleHeight - Frame1.Height
    Frame2.Top = Me.ScaleHeight - Frame2.Height
    Text1.Height = Me.ScaleHeight - Frame2.Height - Text1.Top - 60
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set tidy = Nothing
End Sub

Private Sub optShow_Click(Index As Integer)
    Validate
End Sub
