VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PSC Monitor"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame 
      Height          =   1935
      Index           =   5
      Left            =   -5000
      TabIndex        =   34
      Top             =   480
      Width           =   3615
      Begin VB.CheckBox ChSetting 
         Caption         =   "Check every 5 min for votes/comments"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   37
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox txtAddress 
         Height          =   285
         Index           =   5
         Left            =   240
         OLEDropMode     =   1  'Manual
         TabIndex        =   36
         Top             =   1440
         Width           =   3135
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Index           =   5
         Left            =   240
         OLEDropMode     =   1  'Manual
         TabIndex        =   35
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Tip: You can drag a link from your browser or an Internet shortcut from your Favorites"
         Height          =   855
         Index           =   2
         Left            =   1800
         TabIndex        =   44
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Address"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Name"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Frame Frame 
      Height          =   1935
      Index           =   4
      Left            =   -5000
      TabIndex        =   16
      Top             =   480
      Width           =   3615
      Begin VB.TextBox txtAddress 
         Height          =   285
         Index           =   4
         Left            =   840
         OLEDropMode     =   1  'Manual
         TabIndex        =   18
         Top             =   1440
         Width           =   2535
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Index           =   4
         Left            =   840
         OLEDropMode     =   1  'Manual
         TabIndex        =   17
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label8 
         Caption         =   "Tip: You can drag a link from your browser or an Internet shortcut from your Favorites into the textboxes below"
         Height          =   615
         Index           =   1
         Left            =   240
         TabIndex        =   43
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label4 
         Caption         =   "Address"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   1500
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Name"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   1020
         Width           =   615
      End
   End
   Begin VB.Frame Frame 
      Height          =   1935
      Index           =   2
      Left            =   -5000
      TabIndex        =   4
      Top             =   480
      Width           =   3615
      Begin VB.Image Image1 
         Height          =   240
         Index           =   7
         Left            =   120
         Picture         =   "frmOptions.frx":0000
         Top             =   1620
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   6
         Left            =   360
         Picture         =   "frmOptions.frx":058A
         Top             =   1620
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "New votes/comments"
         Height          =   255
         Index           =   6
         Left            =   720
         TabIndex        =   42
         Top             =   1650
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Checking for new votes"
         Height          =   255
         Index           =   5
         Left            =   720
         TabIndex        =   41
         Top             =   1410
         Width           =   2775
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   5
         Left            =   360
         Picture         =   "frmOptions.frx":0B14
         Top             =   1380
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   4
         Left            =   360
         Picture         =   "frmOptions.frx":12A2
         Top             =   1140
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "New submission posted on PSC"
         Height          =   255
         Index           =   4
         Left            =   720
         TabIndex        =   33
         Top             =   1170
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "Cant check - no Internet connection"
         Height          =   255
         Index           =   3
         Left            =   720
         TabIndex        =   8
         Top             =   930
         Width           =   2775
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   3
         Left            =   360
         Picture         =   "frmOptions.frx":182C
         Top             =   900
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Monitor checking for new submissions"
         Height          =   255
         Index           =   2
         Left            =   720
         TabIndex        =   7
         Top             =   690
         Width           =   2775
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   2
         Left            =   360
         Picture         =   "frmOptions.frx":1DB6
         Top             =   660
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Monitor Disabled"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   6
         Top             =   450
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   1
         Left            =   360
         Picture         =   "frmOptions.frx":2340
         Top             =   420
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Monitor Active"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   5
         Top             =   210
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   0
         Left            =   360
         Picture         =   "frmOptions.frx":2ACE
         Top             =   180
         Width           =   240
      End
   End
   Begin VB.Frame Frame 
      Height          =   1935
      Index           =   3
      Left            =   -5000
      TabIndex        =   11
      Top             =   480
      Width           =   3615
      Begin VB.TextBox txtAddress 
         Height          =   285
         Index           =   3
         Left            =   840
         OLEDropMode     =   1  'Manual
         TabIndex        =   14
         Top             =   1440
         Width           =   2535
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Index           =   3
         Left            =   840
         OLEDropMode     =   1  'Manual
         TabIndex        =   13
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label8 
         Caption         =   "Tip: You can drag a link from your browser or an Internet shortcut from your Favorites into the textboxes below"
         Height          =   615
         Index           =   0
         Left            =   240
         TabIndex        =   40
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label4 
         Caption         =   "Address"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   1500
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Name"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   1020
         Width           =   615
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   1935
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   480
      Width           =   3615
      Begin VB.Frame Frame2 
         Caption         =   "Settings"
         Height          =   1935
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   1815
         Begin VB.VScrollBar VS 
            Height          =   225
            Left            =   600
            Max             =   58
            Min             =   5
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   1455
            Value           =   55
            Width           =   255
         End
         Begin VB.CheckBox ChSetting 
            Caption         =   "Run at StartUp"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   28
            Top             =   360
            Width           =   1455
         End
         Begin VB.CheckBox ChSetting 
            Caption         =   "Enable Monitor"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   27
            Top             =   750
            Width           =   1455
         End
         Begin VB.Label lblInterval 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "minutes"
            Height          =   255
            Left            =   960
            TabIndex        =   30
            Top             =   1500
            Width           =   615
         End
         Begin VB.Label Label6 
            Caption         =   "Check PSC every"
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   1200
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Action"
         Height          =   1935
         Left            =   1920
         TabIndex        =   21
         Top             =   0
         Width           =   1695
         Begin VB.OptionButton OptNotifyMethod 
            Caption         =   "Launch Browser"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   25
            Top             =   1440
            Width           =   1455
         End
         Begin VB.OptionButton OptNotifyMethod 
            Caption         =   "Launch Ticker"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   24
            Top             =   1080
            Width           =   1335
         End
         Begin VB.OptionButton OptNotifyMethod 
            Caption         =   "Message Box"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   23
            Top             =   720
            Width           =   1455
         End
         Begin VB.OptionButton OptNotifyMethod 
            Caption         =   "Flash Icon"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   1455
         End
      End
   End
   Begin VB.PictureBox PicFocus 
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   0
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   330
      Left            =   3000
      TabIndex        =   3
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Enabled         =   0   'False
      Height          =   330
      Left            =   1920
      TabIndex        =   2
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   975
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2655
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   4683
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Options"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Status"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Link 1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Link 2"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Vote Checker"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************
'***************Copyright PSST 2001********************************
'***************Written by MrBobo**********************************
'This code was submitted to Planet Source Code (www.planetsourcecode.com)
'If you downloaded it elsewhere, they stole it and I'll eat them alive
Dim OptionsLoading As Boolean
Public Sub CheckEnabled()
    'has anything changed?
    cmdApply.Enabled = False
    If OptNotifyMethod(NotifyMethod).Value <> True Then cmdApply.Enabled = True
    If ChSetting(0).Value <> GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "Setting0", 0) Then cmdApply.Enabled = True
    If ChSetting(1).Value <> GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "Setting1", 1) Then cmdApply.Enabled = True
    If ChSetting(2).Value <> GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "Setting2", 0) Then cmdApply.Enabled = True
    If txtName(3).Text <> GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "Name1", "Newest Submissions") Then cmdApply.Enabled = True
    If txtName(4).Text <> GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "Name2", "PSC Search") Then cmdApply.Enabled = True
    If txtName(5).Text <> GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "VoteName", "") Then cmdApply.Enabled = True
    If txtAddress(3).Text <> GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "Link1", "") Then cmdApply.Enabled = True
    If txtAddress(4).Text <> GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "Link2", "") Then cmdApply.Enabled = True
    If txtAddress(5).Text <> GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "VoteLink", "") Then cmdApply.Enabled = True
    If VS.Value <> 60 - Val(GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "CheckInterval", 5)) Then cmdApply.Enabled = True
End Sub
Private Sub ChSetting_Click(Index As Integer)
    If Not OptionsLoading Then CheckEnabled
    On Error Resume Next
    PicFocus.SetFocus
End Sub
Private Sub cmdApply_Click()
    Dim z As Long 'apply changes and enter changes into registry
    If OptNotifyMethod(NotifyMethod).Value = False Then
        For z = 1 To 4
            If OptNotifyMethod(z).Value = True Then
                NotifyMethod = z
                Exit For
            End If
        Next
        SaveSetting "PSST SOFTWARE\" + App.Title, "Settings", "NotifyMethod", NotifyMethod
    End If
    If ChSetting(0).Value <> GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "Setting0", 0) Then
        If ChSetting(0).Value = 1 Then
            SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", App.Title, IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\") + App.EXEName + ".exe"
        Else
            DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", App.Title
        End If
        SaveSetting "PSST SOFTWARE\" + App.Title, "Settings", "Setting0", ChSetting(0).Value
    End If
    If ChSetting(1).Value <> GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "Setting1", 1) Then SaveSetting "PSST SOFTWARE\" + App.Title, "Settings", "Setting1", ChSetting(1).Value
    If ChSetting(2).Value <> GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "Setting2", 0) Then SaveSetting "PSST SOFTWARE\" + App.Title, "Settings", "Setting2", ChSetting(2).Value
    If txtName(3).Text <> GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "Name1", "Newest Submissions") Then
        SaveSetting "PSST SOFTWARE\" + App.Title, "Settings", "Name1", txtName(3).Text
        frmTicker.mnuPU(9).Caption = txtName(3).Text
    End If
    If txtName(4).Text <> GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "Name2", "PSC Search") Then
        SaveSetting "PSST SOFTWARE\" + App.Title, "Settings", "Name2", txtName(4).Text
        frmTicker.mnuPU(10).Caption = txtName(4).Text
    End If
    If txtName(5).Text <> GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "VoteName", "") Then
        SaveSetting "PSST SOFTWARE\" + App.Title, "Settings", "VoteName", txtName(5).Text
        frmTicker.mnuPU(12).Caption = txtName(5).Text
    End If
    If Len(txtName(5).Text) <> 0 Then
        frmTicker.mnuPU(12).Caption = "Check votes on " + txtName(5).Text
        frmTicker.mnuPU(13).Caption = "Read comments on " + txtName(5).Text
    End If
    frmTicker.mnuPU(12).Visible = (Len(frmTicker.mnuPU(12).Caption) <> 0)
    frmTicker.mnuPU(11).Visible = frmTicker.mnuPU(12).Visible
    frmTicker.mnuPU(13).Visible = frmTicker.mnuPU(12).Visible
    If txtAddress(3).Text <> GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "Link1", "") Then SaveSetting "PSST SOFTWARE\" + App.Title, "Settings", "Link1", txtAddress(3).Text
    If txtAddress(4).Text <> GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "Link2", "") Then SaveSetting "PSST SOFTWARE\" + App.Title, "Settings", "Link2", txtAddress(4).Text
    If txtAddress(5).Text <> GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "VoteLink", "") Then
        SaveSetting "PSST SOFTWARE\" + App.Title, "Settings", "VoteLink", txtAddress(5).Text
        CurVotes = 0 'reset votes count
        CurComments = 0
    End If
    If VS.Value <> 60 - Val(GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "CheckInterval", 5)) Then SaveSetting "PSST SOFTWARE\" + App.Title, "Settings", "CheckInterval", 60 - VS.Value
    CheckInterval = (60 - VS.Value) * 2
    PSCCheckEnabled = ChSetting(1).Value = 1
    PSCVoteCheckEnabled = ChSetting(2).Value = 1
    frmTicker.SysIcon.IconHandle = IIf(PSCCheckEnabled, frmTicker.ImageList1.ListImages(1).Picture, frmTicker.ImageList1.ListImages(4).Picture)
    frmTicker.SysIcon.TipText = IIf(PSCCheckEnabled, "PSC Monitor", "PSC Monitor disabled")
    frmTicker.Timer1.Enabled = PSCCheckEnabled
    cmdApply.Enabled = False
    PicFocus.SetFocus
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub



Private Sub cmdOK_Click()
    If cmdApply.Enabled Then cmdApply_Click
    Unload Me
End Sub
Private Sub Form_Load()
    OptionsLoading = True
    Me.Icon = frmTicker.Icon 'use the same icon to keep exe size to a minimum
    SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 1 Or 2 'on top
    'read settings from registry
    ChSetting(0).Value = GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "Setting0", 0)
    If ChSetting(0).Value = 1 Then
        SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", App.Title, IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\") + App.EXEName + ".exe"
    Else
        DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", App.Title
    End If
    ChSetting(1).Value = GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "Setting1", 1)
    ChSetting(2).Value = GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "Setting2", 0)
    txtName(3).Text = GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "Name1", "Newest Submissions")
    txtName(4).Text = GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "Name2", "PSC Search")
    txtName(5).Text = GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "VoteName", "")
    txtAddress(3).Text = GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "Link1", "")
    txtAddress(4).Text = GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "Link2", "")
    txtAddress(5).Text = GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "VoteLink", "")
    NotifyMethod = GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "NotifyMethod", 1)
    OptNotifyMethod(NotifyMethod).Value = True
    VS.Value = 60 - Val(GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "CheckInterval", 5))
    OptionsLoading = False
    CheckEnabled
End Sub
Private Sub OptNotifyMethod_Click(Index As Integer)
    If Not OptionsLoading Then CheckEnabled
    On Error Resume Next
    PicFocus.SetFocus
End Sub
Private Sub TabStrip1_Click()
    Dim z As Long 'show the right frame according to the tab clicked
    For z = 1 To TabStrip1.Tabs.Count
        Frame(z).Left = -5000
    Next
    Frame(TabStrip1.SelectedItem.Index).Left = 240
    If TabStrip1.SelectedItem.Index > 2 Then
        txtName(TabStrip1.SelectedItem.Index).SetFocus
    Else
        PicFocus.SetFocus
    End If
End Sub
Private Sub txtAddress_Change(Index As Integer)
    If Not OptionsLoading Then CheckEnabled
End Sub
Private Sub txtAddress_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim temp As String
    If Data.GetFormat(1) Then
        txtAddress(Index).Text = Data.GetData(1)
    ElseIf Data.GetFormat(15) Then
        temp = Data.Files(1)
        If LCase(ExtOnly(temp)) = "url" Then
            txtName(Index).Text = ChangeExt(FileOnly(temp))
            txtAddress(Index).Text = ReadINI(temp, "InternetShortcut", "URL")
        End If
    Else
        DoMessage Me.hwnd, "Format not supported", vbCritical
    End If
End Sub
Private Sub txtName_Change(Index As Integer)
    If Not OptionsLoading Then CheckEnabled
End Sub

Private Sub txtName_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim temp As String
    If Data.GetFormat(1) Then
        txtAddress(Index).Text = Data.GetData(1)
    ElseIf Data.GetFormat(15) Then
        temp = Data.Files(1)
        If LCase(ExtOnly(temp)) = "url" Then
            txtName(Index).Text = ChangeExt(FileOnly(temp))
            txtAddress(Index).Text = ReadINI(temp, "InternetShortcut", "URL")
        End If
    Else
        DoMessage Me.hwnd, "Format not supported", vbCritical
    End If
End Sub

Private Sub VS_Change()
    lblInterval.Caption = 60 - VS.Value
    If Not OptionsLoading Then CheckEnabled
End Sub
Private Sub VS_Scroll()
    VS_Change
End Sub

