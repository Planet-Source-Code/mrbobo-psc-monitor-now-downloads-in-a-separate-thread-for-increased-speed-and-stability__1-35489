VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmComments 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   5475
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   4080
      Width           =   975
   End
   Begin RichTextLib.RichTextBox txtComment 
      Height          =   3255
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5741
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmComments.frx":0000
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse to Site"
      Height          =   300
      Left            =   3600
      TabIndex        =   7
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdComment 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3240
      TabIndex        =   4
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton cmdComment 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2760
      TabIndex        =   3
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton cmdComment 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   2
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton cmdComment 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1680
      TabIndex        =   1
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Comments made by:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   180
      Width           =   1815
   End
End
Attribute VB_Name = "frmComments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this code is just to display the data contained in the
'two collections mNames and mComments
Public Sub EnableCommands()
    cmdComment(0).Enabled = txtComment.Tag > 1
    cmdComment(1).Enabled = txtComment.Tag > 1
    cmdComment(2).Enabled = txtComment.Tag < mNames.Count
    cmdComment(3).Enabled = txtComment.Tag < mNames.Count
    cmdSave.Enabled = mNames.Count > 0
End Sub

Private Sub cmdBrowse_Click()
    txtComment.SetFocus
    ShellExecute Me.hwnd, vbNullString, Me.Tag, vbNullString, "c:\", 1
End Sub

Private Sub cmdComment_Click(Index As Integer)
    Select Case Index
        Case 0
            txtComment.Tag = 1
        Case 1
            txtComment.Tag = Val(txtComment.Tag) - 1
        Case 2
            txtComment.Tag = Val(txtComment.Tag) + 1
        Case 3
            txtComment.Tag = mNames.Count
    End Select
    txtComment.Text = mComments(Val(txtComment.Tag))
    txtName.Text = mNames(Val(txtComment.Tag))
    EnableCommands
    txtComment.SetFocus
End Sub

Private Sub cmdSave_Click()
    Dim temp As String, sfile As String, z As Long
    If mComments.Count = 0 Then
        DoMessage Me.hwnd, "Nothing to save!", vbCritical
        Exit Sub
    End If
    With cmnDlg
        .ownerform = Me.hwnd
        .Filter = "Plain text (*.txt)|*.txt"
        ShowSave
        If Len(.filename) = 0 Then Exit Sub
        sfile = .filename
    End With
    If InStr(1, sfile, ".") = 0 Then
        sfile = sfile + ".txt"
    Else
        sfile = ChangeExt(sfile, "txt")
    End If
    temp = "PSC Monitor - " & Trim(Str(Now)) & vbCrLf & Me.Caption & vbCrLf & "Address: " & Me.Tag & vbCrLf & vbCrLf
    For z = 1 To mComments.Count
        temp = temp & String(50, "*") & vbCrLf & mNames(z) & vbCrLf & String(Len(mNames(z)), "¯") & vbCrLf & mComments(z) & vbCrLf & vbCrLf
    Next
    If FileExists(sfile) Then Kill sfile
    FileSave temp, sfile
End Sub

Private Sub Form_Load()
    Me.Icon = frmTicker.Icon 'use the same icon to keep exe size to a minimum
    SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 1 Or 2 'on top

End Sub
