VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   4020
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6855
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2774.676
   ScaleMode       =   0  'User
   ScaleWidth      =   6437.199
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   0
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   0
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   5520
      TabIndex        =   0
      Top             =   3600
      Width           =   1260
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   345
      Left            =   4200
      TabIndex        =   2
      Top             =   3600
      Width           =   1245
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   6310.427
      Y1              =   2401.958
      Y2              =   2401.958
   End
   Begin VB.Label lblDescription 
      Caption         =   "App Description"
      ForeColor       =   &H00000000&
      Height          =   2130
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   6525
   End
   Begin VB.Label lblTitle 
      Caption         =   "Application Title"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   960
      TabIndex        =   4
      Top             =   0
      Width           =   4725
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   112.686
      X2              =   6324.513
      Y1              =   2401.958
      Y2              =   2401.958
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   585
      Left            =   960
      TabIndex        =   5
      Top             =   480
      Width           =   4605
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
        KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
        KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL

' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                                           ' Unicode nul terminated string
Const REG_DWORD = 4                                        ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Private Sub cmdSysInfo_Click()
    Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "About " & App.Title

    lblTitle.Caption = App.Title & " Version " & App.Major & "." & App.Minor & "." & App.Revision & vbNewLine

    lblVersion.Caption = "Written by Bill MacIntyre," & vbNewLine
    lblVersion.Caption = lblVersion.Caption & "New Brunswick Canada," & vbNewLine
    lblVersion.Caption = lblVersion.Caption & "billymac@rogers.com"

    lblDescription.Caption = "This project was built as a gui front end to Andre Blavier's wrapper around the TidyCom Dll." & vbNewLine
lblDescription.Caption = lblDescription.Caption & "TidyCOM is a Windows COM component wrapping Dave Raggett's HTML Tidy, a free utility application from the World Wide Web Consortium that helps you clean up your web pages. HTML Tidy is available from the W3C as a command-line program, à la Unix. To better fit in the Windows environment Andre has written COM component wrapper for Tidy available here:" & vbNewLine
lblDescription.Caption = lblDescription.Caption & "http://perso.wanadoo.fr/ablavier/TidyCOM/" & vbNewLine
lblDescription.Caption = lblDescription.Caption & "The Tidy SourceForge object is here:" & vbNewLine
lblDescription.Caption = lblDescription.Caption & "http://tidy.sourceforge.net/"
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr

    Dim rc As Long
    Dim SysInfoPath As String

    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
        ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existance Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"

            ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
        ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If

    Call Shell(SysInfoPath, vbNormalFocus)

    Exit Sub
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                          ' Loop Counter
    Dim rc As Long                                         ' Return Code
    Dim hKey As Long                                       ' Handle To An Open Registry Key
    Dim hDepth As Long                                     '
    Dim KeyValType As Long                                 ' Data Type Of A Registry Key
    Dim tmpVal As String                                   ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                 ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey)    ' Open Registry Key

    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError         ' Handle Error...

    tmpVal = String$(1024, 0)                              ' Allocate Variable Space
    KeyValSize = 1024                                      ' Mark Variable Size

    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
            KeyValType, tmpVal, KeyValSize)                ' Get/Create Key Value

    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError         ' Handle Errors

    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then          ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)              ' Null Found, Extract From String
    Else                                                   ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                  ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                 ' Search Data Types...
        Case REG_SZ                                        ' String Registry Key Data Type
            KeyVal = tmpVal                                ' Copy String Value
        Case REG_DWORD                                     ' Double Word Registry Key Data Type
            For i = Len(tmpVal) To 1 Step -1               ' Convert Each Bit
                KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))    ' Build Value Char. By Char.
            Next
            KeyVal = Format$("&h" + KeyVal)                ' Convert Double Word To String
    End Select

    GetKeyValue = True                                     ' Return Success
    rc = RegCloseKey(hKey)                                 ' Close Registry Key
    Exit Function                                          ' Exit

GetKeyError:                                               ' Cleanup After An Error Has Occured...
    KeyVal = ""                                            ' Set Return Val To Empty String
    GetKeyValue = False                                    ' Return Failure
    rc = RegCloseKey(hKey)                                 ' Close Registry Key
End Function
