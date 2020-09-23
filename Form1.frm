VERSION 5.00
Begin VB.Form SINfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SINfo v1 (Beta)"
   ClientHeight    =   720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7830
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   720
   ScaleWidth      =   7830
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6120
      ScaleHeight     =   495
      ScaleWidth      =   1635
      TabIndex        =   5
      Top             =   120
      Width           =   1695
      Begin VB.CommandButton butToTray 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tray"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   810
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   210
         UseMaskColor    =   -1  'True
         Width           =   825
      End
      Begin VB.CommandButton butAbout 
         BackColor       =   &H00FFFFFF&
         Caption         =   "About"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   0
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   210
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton butPrefix 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Prefix"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   810
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   825
      End
      Begin VB.CommandButton butStart 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "AIM User"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   60
      TabIndex        =   3
      Top             =   30
      Width           =   2115
      Begin VB.Label lblSN 
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   150
         TabIndex        =   4
         Top             =   270
         Width           =   1455
      End
   End
   Begin VB.Timer timChecker 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   660
      Top             =   960
   End
   Begin VB.Frame Frame2 
      Caption         =   "Song Playing"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   2190
      TabIndex        =   1
      Top             =   30
      Width           =   3915
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   3330
         Top             =   120
      End
      Begin VB.Label lblSong 
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   225
         Left            =   150
         TabIndex        =   2
         Top             =   270
         Width           =   3645
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Label txtNow 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2700
      TabIndex        =   0
      Top             =   2430
      Width           =   2055
   End
End
Attribute VB_Name = "SINfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Profile As String
Public Current As String
Public ScreenName As String
'Public ParenthWnd As Long
Public X As Long
Public Button As Long

Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Private Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Const VK_SPACE = &H20
Const WM_COMMAND = &H111
Const WM_CLOSE = &H10
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const WM_KEYDOWN = &H100
Const WM_KEYUP = &H101
Const GW_CHILD = 5
Const GW_HWNDFIRST = 0
Const GW_HWNDLAST = 1
Const GW_HWNDNEXT = 2
Const GW_HWNDPREV = 3
Const GW_MAX = 5
Const GW_OWNER = 4
Const SMTO_NORMAL = &H0

Public Enum T_WindowStyle
    Hidden = 0
    Maximized = 3
    Normal = 1
    ShowOnly = 5
End Enum



'''''''''''''''''''''''''''''''''''''''''''''''''''
'           SINfo v1 AIM/WINAMP Song Tool         '
'                  Coded by det0x                 '
'        http://det0x.tk - detox@punkass.com      '
'''''''''''''''''''''''''''''''''''''''''''''''''''

'readme.txt:
'SINfo v1 (Beta 2)
'Coded using Vb6
'Tested on Winamp 2.81 and AIM 4.8

'Yea so this is SINfo... It takes the song winamp is playing and attempts to
'update your AIM profile with your current song in it.

'To use it, you need to place '%s' in your profile where you want the song
'and prefix (Now Playing:, etc...) to appear.
'You must run this while Winamp and AIM are running and playing/logged in.

'*NOTE: Don't Edit the %s or the 'Song' text in your profile
'while SINfo is running. This will cause SINfo to stop updating your profile
'and possibly cause an error.

'Fool around with it. Have fun!

'BTW.. I found a new way to update the AIM profile. This new way, to my know-
'lege does not crash AIM. Lets hope it says that way.

'Peace!



Private Sub Form_Load()
    
    Me.Show
    Me.Refresh
    
    With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "SINfo for AIM/Winamp" & vbNullChar
    End With

    txtNow.Caption = ""

End Sub

Private Sub butToTray_Click()
    Me.WindowState = vbMinimized
End Sub




Private Sub butStart_Click()

    Dim txtFile
    
    ScreenName = getSN()
    Profile = ReadRegistry(HKEY_CURRENT_USER, "Software\America Online\AOL Instant Messenger (TM)\CurrentVersion\Misc\", "BaseDataPath") & "\" & ScreenName & "\info.htm"

    If getSN() <> "" And Song() <> "" And butStart.Caption = "Start" Then
        'RunTime = 1
        timChecker.Enabled = True
        Current = "%s"
        'Click = 1
        butStart.Caption = "Stop"
        butPrefix.Enabled = False
    Else
        Call ShowWindow(X, 1)
        timChecker.Enabled = False
        butStart.Caption = "Start"
        butPrefix.Enabled = True
    End If
    
    If getSN() <> "" Or Song() <> "" And butStart.Caption = "Stop" Then

        Open Profile For Input As #1
        Line Input #1, txtFile
        Close #1

        txtFile = Replace(txtFile, Current, "%s")
        txtFile = Replace(txtFile, "[Stopped]", "")

        Open Profile For Output As #1
        Print #1, txtFile
        Close #1
    End If

End Sub


Private Sub butAbout_Click()
MsgBox "SINfo v1 Song Tool (Pre-Beta)" & vbCrLf & "Coded by det0x" & vbCrLf & "http://det0x.tk - det0x@punkass.com", vbOKOnly, "SINfo v1 Pre-Beta"
End Sub


Private Sub timChecker_Timer()

    Dim txtFile As String
    
    'NEW PROFILE UPDATER'
    X = FindWindow("#32770", "Create a Profile - Searchable Directory")
    If X <> 0 Then
        Call ShowWindow(X, 0)
        Button = FindWindowEx(X, 0, "Button", vbNullString)
        Button = FindWindowEx(X, Button, "Button", vbNullString)
    
        Call SendMessage(Button, WM_KEYDOWN, VK_SPACE, 0&)
        Call SendMessage(Button, WM_KEYUP, VK_SPACE, 0&)
        
        Delay 0.5
        
        Call SendMessage(Button, WM_KEYDOWN, VK_SPACE, 0&)
        Call SendMessage(Button, WM_KEYUP, VK_SPACE, 0&)

        Button = FindWindowEx(X, 0, "Button", vbNullString)
        Button = FindWindowEx(X, Button, "Button", vbNullString)
        Button = FindWindowEx(X, Button, "Button", vbNullString)
    
        Delay 0.5
    
        Call SendMessage(Button, WM_KEYDOWN, VK_SPACE, 0&)
        Call SendMessage(Button, WM_KEYUP, VK_SPACE, 0&)

    End If
    ''''''''''''''''''''''''''''''''
    
    '''''CLOSES THE AWAY MSG - CAN CAUSE AIM TO CRASH '''''
    'ParenthWnd = FindWindow("#32770", "*SINfo*")
    'If ParenthWnd > 0 Then
    '    Call ShowWindow(ParenthWnd, 0)
        'Delay 1
        'SendMessage ParenthWnd, WM_CLOSE, 0&, 0&
    'End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''

    If getSN() = "" Or Song() = "" Then
      timChecker.Enabled = False
      butStart.Caption = "Start"
      butPrefix.Enabled = True
      
      Open Profile For Input As #1
      Line Input #1, txtFile
      Close #1

      txtFile = Replace(txtFile, Current, "%s")
      txtFile = Replace(txtFile, "[Stopped]", "")

      Open Profile For Output As #1
      Print #1, txtFile
      Close #1
      
      Exit Sub
    End If
    
    If Current <> txtNow.Caption & Song() Then
      Call newSong
    Else
      Exit Sub
    End If
    

End Sub


Private Sub Form_Unload(Cancel As Integer)

    If lblSN.Caption <> "N/A" Then
      Dim txtFile As String

      Open Profile For Input As #1
      Line Input #1, txtFile
      Close #1

      txtFile = Replace(txtFile, Current, "%s")
      txtFile = Replace(txtFile, "[Stopped]", "")

      Open Profile For Output As #1
      Print #1, txtFile
      Close #1
    End If
    
    Call ShowWindow(X, 1)
    Shell_NotifyIcon NIM_DELETE, nid
    End
    
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim result As Long
    Dim msg As Long
    
    If Me.ScaleMode = vbPixels Then
        msg = X
    Else
        msg = X / Screen.TwipsPerPixelX
    End If
    
    Select Case msg
        Case WM_LBUTTONUP
            Me.WindowState = vbNormal
            result = SetForegroundWindow(Me.hwnd)
            Me.Show
            Shell_NotifyIcon NIM_DELETE, nid
        Case WM_LBUTTONDBLCLK
            Me.WindowState = vbNormal
            result = SetForegroundWindow(Me.hwnd)
            Me.Show
            Shell_NotifyIcon NIM_DELETE, nid
        Case WM_RBUTTONUP
            Me.WindowState = vbNormal
            result = SetForegroundWindow(Me.hwnd)
            Me.Show
            Shell_NotifyIcon NIM_DELETE, nid
    End Select
End Sub


Private Sub Form_Resize()

    If Me.WindowState = vbMinimized Then
        Me.Hide
        Shell_NotifyIcon NIM_ADD, nid
    End If

End Sub


Function Song() As String

    Dim Winamp As Long

    Winamp& = FindWindow("Winamp v1.x", vbNullString)
    If Winamp& <> 0& Then
        GoTo Start
    Else
      Song = ""
      Exit Function
    End If
    
Start:
    Dim WinAmpHwnd As Long
    Dim TitleText As String
    Dim nBytes As Integer
    Dim Idx As Long
    
    WinAmpHwnd = FindWindow("Winamp v1.x", vbNullString)

    TitleText = Space(255)
    nBytes = 256

    Call GetWindowText(WinAmpHwnd, TitleText, nBytes)
    TitleText = Left(TitleText, InStr(1, TitleText, Chr(0)) - 1)

    If Mid$(TitleText, 1, 3) = "Win" Then
      Song = "No Song Playing"
    Else
      Idx = InStr(1, TitleText, ".") + 2
      Song = Mid$(TitleText, Idx, Len(TitleText))
    End If

End Function


Function getSN() As String
 ' Snippet taken from DigitalAIM.bas
    
    Dim BuddyList As Long

    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    If BuddyList& <> 0& Then
        GoTo Start
    Else
      getSN = ""
      Exit Function
    End If

Start:
    Dim GetIt As String, clear As String
    
    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    GetIt$ = GetCaption(BuddyList&)
    
    If Right$(GetIt$, 20) = "'s Buddy List Window" Then
      clear$ = Replace(GetIt$, "'s Buddy List Window", "")
    Else
      clear$ = Replace(GetIt$, "' Buddy List Window", "")
    End If
    
    getSN = remSpaces(clear$)
    
End Function


Function remSpaces(ByVal strString As String) As String
 ' Snippet coded by Ian Ippolito
  
    Dim strResult As String
    strResult = ""
    Dim intIndex As Integer


    For intIndex = 1 To Len(strString)


    If (Mid$(strString, intIndex, 1) <> " ") Then
      strResult = strResult + Mid$(strString, intIndex, 1)
    End If
    
    Next intIndex
    remSpaces = strResult
    
End Function


Function newSong() As String

    Dim txtFile As String
    Dim timeout

    Open Profile For Input As #1
    Line Input #1, txtFile
    Close #1

    txtFile = Replace(txtFile, "[Stopped]", "")
    
      txtFile = Replace(txtFile, Current, txtNow.Caption & Song())

    Open Profile For Output As #1
    Print #1, txtFile
    Close #1

    'PART OF NEW PROFILE UPDATE'
    Call RunMenuByString("Edit &Profile...")
    ''''''''''''''''''''''''''''
    
    '''''''''''OLD PROFILE UPDATE 1 - CAN CAUSE AIM TO CRASH '''''
    'Call OpenInternet(Me, "aim:goaway?message=*SINfo*", Hidden)
    'Delay 1
    'ParenthWnd = FindWindow("#32770", "*SINfo*")
    'Call ShowWindow(ParenthWnd, 0)
    'If ParenthWnd > 0 Then
    '    SendMessage ParenthWnd, WM_CLOSE, 0&, 0&
    'End If
    
    'SendMessage ParentWnd, WM_LBUTTONDOWN
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ''''OLD PROFILE UPDATE 2 - CRASHES AIM ALSO''''
    'Call RunMenuByString("Edit &Profile...")
    
    'X = FindWindow("#32770", vbNullString)
    'Button = FindWindowEx(X, 0, "Button", vbNullString)
    'Button = FindWindowEx(X, Button, "Button", vbNullString)
    
    'Call SendMessage(Button, WM_KEYDOWN, VK_SPACE, 0&)
    'Call SendMessage(Button, WM_KEYUP, VK_SPACE, 0&)
    'Call SendMessage(Button, WM_KEYDOWN, VK_SPACE, 0&)
    'Call SendMessage(Button, WM_KEYUP, VK_SPACE, 0&)

    'Button = FindWindowEx(X, 0, "Button", vbNullString)
    'Button = FindWindowEx(X, Button, "Button", vbNullString)
    'Button = FindWindowEx(X, Button, "Button", vbNullString)
    
    'Call SendMessage(Button, WM_KEYDOWN, VK_SPACE, 0&)
    'Call SendMessage(Button, WM_KEYUP, VK_SPACE, 0&)
    ''''''''''''''''''''''''''
    
    Current = txtNow.Caption & Song()

    lblSN.Caption = getSN()
    lblSong.Caption = Song()

End Function

Function GetCaption(TheWin)
 ' Snippet taken from From Dos32.bas
    
    Dim WindowLngth As Integer, WindowTtle As String, Moo As String
    
    WindowLngth% = GetWindowTextLength(TheWin)
    WindowTtle$ = String$(WindowLngth%, 0)
    Moo$ = GetWindowText(TheWin, WindowTtle$, (WindowLngth% + 1))
    GetCaption = WindowTtle$

End Function

Public Sub OpenInternet(Parent As Form, URL As String, WindowStyle As T_WindowStyle)
    ShellExecute Parent.hwnd, "Open", URL, "", "", WindowStyle
End Sub

Private Sub butPrefix_Click()
    txtNow.Caption = InputBox("Enter text that will prefix the song in your profile:", "Change Song Prefix")
End Sub

Public Sub RunMenuByString(SearchString As String)
    Dim AOL As Long, aMenu As Long, mCount As Long
    Dim LookFor As Long, sMenu As Long, sCount As Long
    Dim LookSub As Long, sID As Long, sString As String
    AOL& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    aMenu& = GetMenu(AOL&)
    mCount& = GetMenuItemCount(aMenu&)
    For LookFor& = 0& To mCount& - 1
        sMenu& = GetSubMenu(aMenu&, LookFor&)
        sCount& = GetMenuItemCount(sMenu&)
        For LookSub& = 0 To sCount& - 1
            sID& = GetMenuItemID(sMenu&, LookSub&)
            sString$ = String$(100, " ")
            Call GetMenuString(sMenu&, sID&, sString$, 100&, 1&)
            If InStr(LCase(sString$), LCase(SearchString$)) Then
                Call SendMessageLong(AOL&, WM_COMMAND, sID&, 0&)
                Exit Sub
            End If
        Next LookSub&
    Next LookFor&
End Sub


Public Sub Delay(HowLong As Date)
Dim TempTime As Date
TempTime = DateAdd("s", HowLong, Now)
While TempTime > Now
Wend
End Sub

Private Sub Timer1_Timer()
'PART OF PROFILE UPDATER'
Call ShowWindow(X, 0)
'''''''''''''''''''''''''
End Sub
