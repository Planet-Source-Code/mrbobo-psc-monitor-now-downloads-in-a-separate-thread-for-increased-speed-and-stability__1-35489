VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTicker 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   2055
   ControlBox      =   0   'False
   Icon            =   "frmTicker.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   2055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txtComments 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   6120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtNewVote 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtNewPost 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   5160
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   120
      Top             =   1440
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   42000
      Left            =   360
      Top             =   1440
   End
   Begin VB.CommandButton cmdCloseLeft 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   2.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1830
      Picture         =   "frmTicker.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   15
      Width           =   210
   End
   Begin VB.PictureBox PicTitleBar 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   137
      TabIndex        =   1
      Top             =   0
      Width           =   2055
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1200
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTicker.frx":0FC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTicker.frx":155E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTicker.frx":1AF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTicker.frx":2092
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTicker.frx":2830
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTicker.frx":2DCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTicker.frx":3364
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTicker.frx":3B02
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   720
      Top             =   1440
   End
   Begin SHDocVwCtl.WebBrowser Brow 
      Height          =   5415
      Left            =   -360
      TabIndex        =   0
      Top             =   -480
      Width           =   4455
      ExtentX         =   7858
      ExtentY         =   9551
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   6840
      Top             =   2880
   End
   Begin VB.Menu mnuPUBase 
      Caption         =   "MnuPUBase"
      Visible         =   0   'False
      Begin VB.Menu mnuPU 
         Caption         =   "PSC Newest 10"
         Index           =   0
      End
      Begin VB.Menu mnuPU 
         Caption         =   "PSC Newest 50"
         Index           =   1
      End
      Begin VB.Menu mnuPU 
         Caption         =   "Leader Board"
         Index           =   2
      End
      Begin VB.Menu mnuPU 
         Caption         =   "PSC Search"
         Index           =   3
      End
      Begin VB.Menu mnuPU 
         Caption         =   "Ask A Pro"
         Index           =   4
      End
      Begin VB.Menu mnuPU 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuPU 
         Caption         =   "MSDN Advanced Search"
         Index           =   6
      End
      Begin VB.Menu mnuPU 
         Caption         =   "Microsoft Knowledge Base"
         Index           =   7
      End
      Begin VB.Menu mnuPU 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuPU 
         Caption         =   ""
         Index           =   9
      End
      Begin VB.Menu mnuPU 
         Caption         =   ""
         Index           =   10
      End
      Begin VB.Menu mnuPU 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnuPU 
         Caption         =   ""
         Index           =   12
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPU 
         Caption         =   ""
         Index           =   13
      End
      Begin VB.Menu mnuPU 
         Caption         =   "-"
         Index           =   14
      End
      Begin VB.Menu mnuPU 
         Caption         =   "Show Code Ticker"
         Index           =   15
      End
      Begin VB.Menu mnuPU 
         Caption         =   "Settings"
         Index           =   16
      End
      Begin VB.Menu mnuPU 
         Caption         =   "Refresh"
         Index           =   17
      End
      Begin VB.Menu mnuPU 
         Caption         =   "-"
         Index           =   18
      End
      Begin VB.Menu mnuPU 
         Caption         =   "Exit"
         Index           =   19
      End
   End
End
Attribute VB_Name = "frmTicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************
'***************Copyright PSST 2001********************************
'***************Written by MrBobo**********************************
'This code was submitted to Planet Source Code (www.planetsourcecode.com)
'If you downloaded it elsewhere, they stole it and I'll eat them alive
Public WithEvents SysIcon As Class1 'tray icon
Attribute SysIcon.VB_VarHelpID = -1
Public CInet As CWinInetConnection
Dim CurNum As Long 'last submitted ID number
Dim TimerNum As Integer 'counter to increase timer length
Dim TimerVotesNum As Integer 'counter to increase timer length
Dim TickerLoaded As Boolean 'code ticker loaded flag
Dim BrowPath As String 'location of code ticker htm file
Dim UnPaddedBrowPath As String 'location of code ticker htm file padded to URL format
Dim ShowVoteMessage As Boolean
Dim ContactingSite As Boolean
Dim CheckingFromMenu As Boolean
Dim ItsANewVote As Boolean

Private Sub Brow_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    'if the path is other than the ticker's HTM then a link in the
    'ticker was clicked so launch the default web browser to that
    'address and cancel navigation by the ticker htm
    If (URL <> BrowPath And URL <> UnPaddedBrowPath) And TickerLoaded Then
        Cancel = True
        Set pDisp = Nothing
        ShellExecute Me.hwnd, vbNullString, URL, vbNullString, "c:\", 1
    End If
End Sub

Private Sub Brow_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    If URL = BrowPath Or URL = UnPaddedBrowPath Then
        TickerLoaded = True 'set flag
        Brow.Offline = Not CInet.IsConnected 'work Off-line if not connected
    End If
End Sub

Private Sub cmdCloseLeft_Click()
    Me.Visible = False

End Sub

Private Sub Form_Load()
    Dim temp As String
    If App.PrevInstance Then
        MsgBox "PSC Monitor is already loaded!", vbInformation
        End
    End If
    If Not FileExists(IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\") + "BBdowner.exe") Then
        MsgBox "Critical error - file not found BBdowner.exe" & vbCrLf & "PSC Monitor cannot run without this file" & vbCrLf & "Exiting program", vbCritical
        End
    End If
    'delete the old file and replace with a new one
    Set CInet = New CWinInetConnection
    CInet.Refresh
    If CInet.IsConnected Then
        CInet.SetGlobalOnline
        Brow.Offline = False
        temp = GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "VoteLink", "")
        If Len(temp) <> 0 Then GetSourceDocument temp, 1
            'Thanks to Ian at PSC for providing this page - it only has the ID number of the newest submission on it - nothing else!
        GetSourceDocument "http://www.pscode.com/vb/feeds/LatestCodeId.asp?lngWId=1", 0
    Else
        Brow.Offline = True
    End If
    UnPaddedBrowPath = IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\") + "PSCScroller.htm"
    If FileExists(UnPaddedBrowPath) Then Kill UnPaddedBrowPath
    FileSave getTicker, UnPaddedBrowPath
    BrowPath = Replace("file:///" + Replace(UnPaddedBrowPath, "\", "/"), Chr(32), "%20")
    Brow.Navigate BrowPath 'load the ticker
    DoEvents
    Me.Visible = False
    Set SysIcon = New Class1 'go to the tray
    'retrieve settings
    CurNum = Val(GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "LastCount", 0))
    CurVotes = Val(GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "LastVoteCount", 0))
    CurComments = Val(GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "LastCommentCount", 0))
    NotifyMethod = GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "NotifyMethod", 1)
    CheckInterval = Val(GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "CheckInterval", 5)) * 2
    PSCCheckEnabled = Val(GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "Setting1", 1)) = 1
    PSCVoteCheckEnabled = Val(GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "Setting2", 0)) = 1
    mnuPU(9).Caption = GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "Name1", "Newest Submissions")
    mnuPU(10).Caption = GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "Name2", "PSC Search")
    mnuPU(12).Caption = GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "VoteName", "")
    'If there is a submission to check votes on show the menu
    If Len(mnuPU(12).Caption) <> 0 Then
        mnuPU(13).Caption = "Read comments on " + mnuPU(12).Caption
        mnuPU(12).Caption = "Check votes on " + mnuPU(12).Caption
    End If
    mnuPU(12).Visible = (Len(mnuPU(12).Caption) <> 0)
    mnuPU(11).Visible = mnuPU(12).Visible
    mnuPU(13).Visible = mnuPU(12).Visible
    SysIcon.Initialize hwnd, IIf(PSCCheckEnabled, ImageList1.ListImages(1).Picture, ImageList1.ListImages(4).Picture), ""
    SysIcon.ShowIcon
    SysIcon.TipText = "PSC Monitor"
    'on top
    SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 1 Or 2
    Me.Refresh
    Timer1.Enabled = PSCCheckEnabled 'monitor enabled?
    Timer3.Enabled = PSCVoteCheckEnabled 'vote checker enabled
    'no right click menu in the browser
    gLngMouseHook = SetWindowsHookEx(WH_MOUSE, AddressOf MouseHookProc, App.hInstance, GetCurrentThreadId)
    
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim msgCallBackMessage As Long
    If Timer2.Enabled Or Timer4.Enabled Then
        Timer2.Enabled = False 'stop flasher
        Timer4.Enabled = False 'stop flasher
        SysIcon.IconHandle = IIf(PSCCheckEnabled, ImageList1.ListImages(1).Picture, ImageList1.ListImages(4).Picture)
        SysIcon.TipText = "PSC Monitor"
    End If
    msgCallBackMessage = X / Screen.TwipsPerPixelX
    Select Case msgCallBackMessage
        Case WM_LBUTTONDBLCLK
            frmOptions.Show , Me
        Case WM_RBUTTONDOWN
            'show the menu "Setting" as Default menu
            'as suggested by Vlad Vissoultchev
            SetForegroundWindow Me.hwnd
            If Not ContactingSite Then PopupMenu mnuPUBase, , , , mnuPU(16)
    End Select
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    PSCCheckEnabled = False
    UnhookWindowsHookEx gLngMouseHook 'unhook browser right click
    BailOut = True 'stop monitoring connection state
    SaveSetting "PSST SOFTWARE\" + App.Title, "Settings", "LastCount", CurNum
    SaveSetting "PSST SOFTWARE\" + App.Title, "Settings", "LastVoteCount", CurVotes
    SaveSetting "PSST SOFTWARE\" + App.Title, "Settings", "LastCommentCount", CurComments
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim frm As Form
    If FileExists(UnPaddedBrowPath) Then Kill UnPaddedBrowPath 'delete the ticker htm file
    Set SysIcon = Nothing
    For Each frm In Forms
        Unload frm
        Set frm = Nothing
    Next
End Sub

Private Sub mnuPU_Click(Index As Integer)
    On Error Resume Next
    Dim iret As Long, temp As String
    Select Case Index
        Case 0 'PSC Newest 10
            ShellExecute Me.hwnd, vbNullString, "http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?grpCategories=-1&optSort=DateDescending&txtMaxNumberOfEntriesPerPage=10&blnNewestCode=TRUE&blnResetAllVariables=TRUE&lngWId=1", vbNullString, "c:\", 1
        Case 1 'PSC Newest 50
            ShellExecute Me.hwnd, vbNullString, "http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?grpCategories=-1&optSort=DateDescending&txtMaxNumberOfEntriesPerPage=50&blnNewestCode=TRUE&blnResetAllVariables=TRUE&lngWId=1", vbNullString, "c:\", 1
        Case 2 'Leader Board
            ShellExecute Me.hwnd, vbNullString, "http://www.planetsourcecode.com/vb/contest/ContestAndLeaderBoard.asp?lngWid=1", vbNullString, "c:\", 1
        Case 3 'Search
            ShellExecute Me.hwnd, vbNullString, "http://www.planetsourcecode.com/vb/scripts/search.asp?lngWid=1", vbNullString, "c:\", 1
        Case 4 'Ask A Pro
            ShellExecute Me.hwnd, vbNullString, "http://www.planetsourcecode.com/vb/discussion/AskAProMain.asp?lngWId=1", vbNullString, "c:\", 1
        Case 6 'MSDN Advanced Search
            ShellExecute Me.hwnd, vbNullString, "http://search.microsoft.com/us/dev/default.asp", vbNullString, "c:\", 1
        Case 7 'KB
            ShellExecute Me.hwnd, vbNullString, "http://search.support.microsoft.com/kb/c.asp?fr=0&SD=GN&LN=EN-US", vbNullString, "c:\", 1
        Case 9 'link 1
            ShellExecute Me.hwnd, vbNullString, GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "Link1", ""), vbNullString, "c:\", 1
        Case 10 'link 2
            ShellExecute Me.hwnd, vbNullString, GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "Link2", ""), vbNullString, "c:\", 1
        Case 12 'check votes/comments
            If Not CInet.IsConnected Then
                DoMessage Me.hwnd, "Cannot check votes/comments off-line. Make an Internet connection and try again", vbInformation
                Exit Sub
            End If
            temp = GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "VoteLink", "")
            If Len(temp) <> 0 Then
                ShowVoteMessage = True
                SysIcon.TipText = "Checking for new votes/comments on: " & GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "VoteName", "Untitled")
                SysIcon.IconHandle = ImageList1.ListImages(7).Picture
                GetSourceDocument temp, 1
            End If
        Case 13 'view comments
            If Not CInet.IsConnected Then
                DoMessage Me.hwnd, "Cannot read comments off-line. Make an Internet connection and try again", vbInformation
                Exit Sub
            End If
            temp = GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "VoteLink", "")
            If Len(temp) <> 0 Then
                SysIcon.TipText = "Downloading comments on: " & GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "VoteName", "Untitled")
                SysIcon.IconHandle = ImageList1.ListImages(7).Picture
                GetSourceDocument temp, 2
            End If
            frmComments.Tag = temp
            GetSourceDocument temp, 2
        Case 15 'code ticker
            Me.WindowState = 0
            Me.Visible = True
            SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 1 Or 2
            SizeColorCaption 'paint title bar
            TickerLoaded = False
            BrowPath = Replace("file:///" + Replace(UnPaddedBrowPath, "\", "/"), Chr(32), "%20")
            Brow.Navigate BrowPath 'load the ticker
        Case 16 'settings
            frmOptions.Show , Me
        Case 17 'refresh
            If Not CInet.IsConnected Then
                DoMessage Me.hwnd, "Cannot check for new submissions off-line. Make an Internet connection and try again", vbInformation
                Exit Sub
            End If
            If FileExists(UnPaddedBrowPath) Then Kill UnPaddedBrowPath
            TickerLoaded = False
            FileSave getTicker, UnPaddedBrowPath
            Brow.Offline = Not CInet.IsConnected
            If Me.Visible Then
                BrowPath = Replace("file:///" + Replace(UnPaddedBrowPath, "\", "/"), Chr(32), "%20")
                Brow.Navigate BrowPath 'load the ticker
            End If
            SysIcon.IconHandle = ImageList1.ListImages(3).Picture
            SysIcon.TipText = "Checking PSC for new submissions"
            'Thanks to Ian at PSC for providing this page - it only has the ID number of the newest submission on it - nothing else!
            GetSourceDocument "http://www.pscode.com/vb/feeds/LatestCodeId.asp?lngWId=1", 0
        Case 19 'exit
            Unload Me
    End Select
End Sub

Private Sub mnuPUBase_Click()
    Dim z As Long, Enabled As Boolean
    Enabled = CInet.IsConnected
    For z = 0 To 19
        Select Case z
        Case 5, 8, 11, 14, 16, 18, 19
        Case Else
            mnuPU(z).Enabled = Enabled
        End Select
    Next
    If Enabled Then
        CInet.SetGlobalOnline
    End If
    Brow.Offline = Not Enabled
End Sub

Private Sub PicTitleBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture 'alow dragging the form without a titlebar
    Call SendMessage(Me.hwnd, &HA1, 2, 0&)

End Sub

Private Sub Timer1_Timer() 'check for new submissions
    Timer1.Enabled = PSCCheckEnabled
    TimerNum = TimerNum + 1
    If TimerNum = CheckInterval Then
        TimerNum = 0
        If CInet.IsConnected Then
            CInet.SetGlobalOnline
            Brow.Offline = False
            SysIcon.IconHandle = ImageList1.ListImages(3).Picture
            SysIcon.TipText = "Checking PSC for new submissions"
            Timer1.Enabled = False
            'Thanks to Ian at PSC for providing this page - it only has the ID number of the newest submission on it - nothing else!
            GetSourceDocument "http://www.pscode.com/vb/feeds/LatestCodeId.asp?lngWId=1", 0
            TickerLoaded = False
            If FileExists(UnPaddedBrowPath) Then Kill UnPaddedBrowPath
            FileSave getTicker, UnPaddedBrowPath
            Brow.Offline = Not CInet.IsConnected
            If Me.Visible Then
                BrowPath = Replace("file:///" + Replace(UnPaddedBrowPath, "\", "/"), Chr(32), "%20")
                Brow.Navigate BrowPath 'load the ticker
            End If
        End If
    End If
End Sub
Private Sub Timer2_Timer() 'flash the icon in the tray
    If CInet.IsConnected = False Then
        If SysIcon.IconHandle = ImageList1.ListImages(3).Picture Then
            SysIcon.IconHandle = ImageList1.ListImages(5).Picture
        Else
            SysIcon.IconHandle = ImageList1.ListImages(3).Picture
        End If
        SysIcon.TipText = "PSC Monitor - No Internet connection detected"
    Else
        If SysIcon.IconHandle = ImageList1.ListImages(1).Picture Or SysIcon.IconHandle = ImageList1.ListImages(4).Picture Then
            SysIcon.IconHandle = ImageList1.ListImages(2).Picture
        Else
            SysIcon.IconHandle = IIf(PSCCheckEnabled, ImageList1.ListImages(1).Picture, ImageList1.ListImages(4).Picture)
        End If
    End If
    
End Sub
Private Sub Timer3_Timer() 'check for new votes
    'this timer is set to approximately 5 minutes
    'The reason I have done this is so that it wont
    'coincide with the timer checking for new submissions
    'which goes off on the minute. This timer interval
    'is set to 42000 with a counter of 7 which means
    'it will check for new votes every 4 minutes 54 seconds
    'which should keep it out of sync with the other timers
    Dim temp As String
    Timer3.Enabled = PSCVoteCheckEnabled
    TimerVotesNum = TimerVotesNum + 1
    If TimerVotesNum = 7 Then
        TimerVotesNum = 0
        If CInet.IsConnected Then
            SysIcon.IconHandle = ImageList1.ListImages(7).Picture
            temp = GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "VoteLink", "")
            If Len(temp) <> 0 Then GetSourceDocument temp, 1
        End If
    End If
End Sub
Private Sub Timer4_Timer() 'flash the icon in the tray
    If SysIcon.IconHandle = ImageList1.ListImages(1).Picture Or SysIcon.IconHandle = ImageList1.ListImages(4).Picture Then
        SysIcon.IconHandle = IIf(ItsANewVote, ImageList1.ListImages(6).Picture, ImageList1.ListImages(8).Picture)
    Else
        SysIcon.IconHandle = IIf(PSCCheckEnabled, ImageList1.ListImages(1).Picture, ImageList1.ListImages(4).Picture)
    End If
End Sub

Public Function getTicker() As String 'create a web page for the ticker
    getTicker = "<html><body>" + vbCrLf + _
    "<IFRAME ID=IFrame1 FRAMEBORDER=0 SCROLLING=NO" + vbCrLf + _
    "SRC=" + Chr(34) + "http://www.Planet-Source-Code.com/vb/linktous/ScrollingCode.asp?lngWId=1" + Chr(34) + vbCrLf + _
    "height=160>" + vbCrLf + _
    "Your browser does not support inline frames...However, you can click" + vbCrLf + _
    "<A href=" + Chr(34) + "http://www.Planet-Source-Code.com/vb/linktous/ScrollingCode.asp?lngWId=1" + Chr(34) + ">" + vbCrLf + _
    "here</a> to see the related document." + vbCrLf + _
    "</IFRAME>" + vbCrLf + _
    "</body></html>"

End Function
'*********************************************************
'*********CALL TO BBdowner.exe****************************
Public Sub GetSourceDocument(URL As String, mType As Long)
    Dim mPath As String, mHwnd As Long
    Select Case mType 'where to return the data
        Case 0
            mHwnd = txtNewPost.hwnd
        Case 1
            mHwnd = txtNewVote.hwnd
        Case 2
            mHwnd = txtComments.hwnd
    End Select
    mPath = IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\")
    'enter required instructions to registry
    'these will be read by BBdowner.exe
    SaveSetting "PSST SOFTWARE\PSCMonitor", "Downer", "Source", URL
    SaveSetting "PSST SOFTWARE\PSCMonitor", "Downer", "Destination", mPath
    SaveSetting "PSST SOFTWARE\PSCMonitor", "Downer", "FileName", "Source.txt"
    SaveSetting "PSST SOFTWARE\PSCMonitor", "Downer", "ReturnAddress", mHwnd
    'shell a new instance of BBdowner,exe to download the source code
    Shell IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\") + "BBdowner.exe"
End Sub
'**********************************************************
'**********SOURCE CODE PARSING ROUTINES********************
Public Sub ParseNewPostingDocument(mSource As String)
    Dim z As Long, z1 As Long, zEnd As Long, tmp As String
    ContactingSite = False
    On Error GoTo woops
    DoEvents
    Timer2.Enabled = False
    SysIcon.IconHandle = IIf(PSCCheckEnabled, ImageList1.ListImages(1).Picture, ImageList1.ListImages(4).Picture)
    SysIcon.TipText = "PSC Monitor"
    z = Val(mSource) 'Thanks to Ian at PSC there's no need to do any parsing
    If z > CurNum And CurNum <> 0 Then 'new submission
        SysIcon.TipText = "New PSC Submission Identified"
        Select Case NotifyMethod 'Action
            Case 1 'flash icon
                Timer2.Enabled = True
            Case 2 'maessage box
                DoMessage Me.hwnd, "New PSC Submission Identified", vbInformation
            Case 3 'show ticker
                If Me.Visible Then Me.WindowState = 0
                Me.Visible = True
                SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 1 Or 2
                SizeColorCaption 'paint title bar
                If (Brow.LocationURL <> BrowPath And Brow.LocationURL <> UnPaddedBrowPath) Then Brow.Navigate UnPaddedBrowPath
            Case 4 'launch browser to latest code page
                ShellExecute Me.hwnd, vbNullString, "", vbNullString, "c:\", 1
        End Select
    End If
    CurNum = z 'remember for next time
    SysIcon.IconHandle = IIf(PSCCheckEnabled, ImageList1.ListImages(1).Picture, ImageList1.ListImages(4).Picture)
    Timer1.Enabled = True 're-enable the timer
    Exit Sub
woops:
    SysIcon.IconHandle = IIf(PSCCheckEnabled, ImageList1.ListImages(1).Picture, ImageList1.ListImages(4).Picture)
    SysIcon.TipText = "Failed to download data"
    ShowVoteMessage = False
    Timer1.Enabled = True
End Sub
Public Sub ParseForVotesDocument(mSource As String)
    Dim z As Long, z1 As Long, zEnd As Long, tmp As String, ComCnt As Long, mMsg As String, mShortMsg As String
    ContactingSite = False
    On Error GoTo woops
    DoEvents
    ComCnt = CountComments(mSource)
    z1 = InStr(1, mSource, "User Rating:") 'find the number of folks who have voted
    If z1 = 0 Then GoTo woops
    z1 = z1 + 12
    tmp = Trim(Mid(mSource, z1, 16))
    If tmp = "</b> Unrated<br>" Then 'has anyone voted?
        z = 0
        GoTo done
    End If
    z = InStr(z1, mSource, "By ") 'the number comes just after this
    If z = 0 Then GoTo woops
    z = z + 3
    z1 = InStr(z, mSource, Chr(32)) 'the space after the number
    tmp = Trim(Mid(mSource, z, z1 - z)) 'the number itself
    z = Val(tmp)
done:
    'compare to last time and give the correct message
    mMsg = IIf(z > CurVotes And CurVotes <> 0, "New votes on this submission", "No new votes on this submission") & vbCrLf & z & IIf(z = 1, " person has", " people have") & " voted for this submission" & vbCrLf & vbCrLf
    mMsg = mMsg & IIf(ComCnt > CurComments, "New comment on this submission", "No new comments on this submission") & vbCrLf & ComCnt & IIf(ComCnt = 1, " person has", " people have") & " commented on this submission"
    mShortMsg = IIf(z > CurVotes And CurVotes <> 0, "New votes: ", "No new votes: ") & "Votes = " & z & IIf(ComCnt > CurComments, " - New comment: ", " - No new comments: ") & "Comments = " & ComCnt
    
    If ShowVoteMessage Then
        DoMessage Me.hwnd, mMsg, vbInformation
        ShowVoteMessage = False
        SysIcon.TipText = "PSC Monitor"
    Else
        If (CurVotes <> z And z <> 0) Or (CurComments <> ComCnt And ComCnt <> 0) Then
            ItsANewVote = (CurVotes <> z And z <> 0)
            Timer4.Enabled = True
        End If
        SysIcon.TipText = mShortMsg
    End If
    SysIcon.IconHandle = IIf(PSCCheckEnabled, ImageList1.ListImages(1).Picture, ImageList1.ListImages(4).Picture)
    Timer1.Enabled = True
    CurVotes = z 'remember for next time
    CurComments = ComCnt
    Exit Sub
woops:
    SysIcon.IconHandle = IIf(PSCCheckEnabled, ImageList1.ListImages(1).Picture, ImageList1.ListImages(4).Picture)
    SysIcon.TipText = "Failed to download data"
    Timer1.Enabled = True
End Sub
Private Sub ParseForComments(srcStr As String)
    Dim z As Long, z1 As Long, z2 As Long, z3 As Long, z4 As Long, temp As String, tmp As String
    Set mNames = New Collection
    Set mComments = New Collection
    Do
        z = z + 1
        z = InStr(z, srcStr, "<!xmp>")
        If z = 0 Then Exit Do
        z1 = InStr(z + 1, srcStr, "<!/xmp>")
        If z1 <> 0 Then
            mNames.Add Mid(srcStr, z + 6, z1 - z - 6)
            z3 = InStr(z1, srcStr, "/vb/scripts/ReportBadItem.asp")
            If z3 <> 0 Then
                temp = Mid(srcStr, z1 + 6, z3 - z1 - 6)
                z1 = InStr(1, temp, "</span", vbTextCompare)
                If z1 = 0 Then z1 = InStr(1, temp, "<form", vbTextCompare)
                If z1 <> 0 Then
                    z2 = InStrRev(temp, ">", z1, vbTextCompare)
                    temp = Mid(temp, z2 + 1, z1 - z2 - 1)
                    z4 = 1
                    Do
                        z1 = InStr(z4, temp, "&#")
                        If z1 <> 0 Then
                            z2 = InStr(z1, temp, ";")
                            If z2 <> 0 And z2 < z1 + 7 Then
                                tmp = Mid(temp, z1 + 2, z2 - z1 - 2)
                                z3 = Val(tmp)
                                If z3 <> 0 Then temp = Replace(temp, "&#" & tmp & ";", Chr(z3))
                            End If
                        Else
                        Exit Do
                        End If
                        z4 = z1 + 1
                    Loop
                    temp = Replace(temp, vbCr, vbCrLf)
                    temp = Replace(temp, "&quot;", Chr(34))
                    mComments.Add temp
                End If
            Else
                mComments.Add "PSC Monitor - Failed to identify comment"
            End If
        End If
    Loop
    SysIcon.IconHandle = IIf(PSCCheckEnabled, ImageList1.ListImages(1).Picture, ImageList1.ListImages(4).Picture)
    SysIcon.TipText = "PSC Monitor"
    If mNames.Count = 0 Then
        If Len(srcStr) > 0 Then DoMessage Me.hwnd, "No comments on this submission", vbInformation
    Else
        frmComments.Caption = "Comments on: " & GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "VoteName", "Untitled") & " (" & mNames.Count & ")"
        frmComments.txtComment.Tag = 1
        frmComments.txtComment.Text = mComments(1)
        frmComments.txtName.Text = mNames(1)
        frmComments.EnableCommands
        frmComments.Show
    End If
    
End Sub
'**********************************************************
'Hidden textboxes to recieve data returned from BBdowner.exe
Private Sub txtComments_Change()
    Dim sBuffer As String
    On Error Resume Next
    If txtComments.Text = "" Then Exit Sub
    If FileExists(txtComments.Text) Then
        sBuffer = OneGulp(txtComments.Text)
        ParseForComments sBuffer 'extract comments from source code
        Kill txtComments.Text
    Else
        SysIcon.IconHandle = IIf(PSCCheckEnabled, ImageList1.ListImages(1).Picture, ImageList1.ListImages(4).Picture)
        SysIcon.TipText = "Failed to download data"
        Timer1.Enabled = True
    End If
    txtComments.Text = ""
End Sub

Private Sub txtNewPost_Change()
    Dim sBuffer As String
    On Error Resume Next
    If txtNewPost.Text = "" Then Exit Sub
    If FileExists(txtNewPost.Text) Then
        sBuffer = OneGulp(txtNewPost.Text)
        ParseNewPostingDocument sBuffer 'look for new submission count
        Kill txtNewPost.Text
    Else
        SysIcon.IconHandle = IIf(PSCCheckEnabled, ImageList1.ListImages(1).Picture, ImageList1.ListImages(4).Picture)
        SysIcon.TipText = "Failed to download data"
        Timer1.Enabled = True
    End If
    txtNewPost.Text = ""
End Sub

Private Sub txtNewVote_Change()
    Dim sBuffer As String
    On Error Resume Next
    If txtNewVote.Text = "" Then Exit Sub
    If FileExists(txtNewVote.Text) Then
        sBuffer = OneGulp(txtNewVote.Text)
        ParseForVotesDocument sBuffer 'look for voting count
        Kill txtNewVote.Text
    Else
        SysIcon.IconHandle = IIf(PSCCheckEnabled, ImageList1.ListImages(1).Picture, ImageList1.ListImages(4).Picture)
        SysIcon.TipText = "Failed to download data"
        Timer1.Enabled = True
    End If
    txtNewVote.Text = ""
End Sub

