Attribute VB_Name = "FileHandler"
'******************************************************************
'***************Copyright PSST 2001********************************
'***************Written by MrBobo**********************************
'This code was submitted to Planet Source Code (www.planetsourcecode.com)
'If you downloaded it elsewhere, they stole it and I'll eat them alive
Option Explicit
#If Win32 Then 'used to launch browser
    Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#Else
    Public Declare Function ShellExecute Lib "shell.dll" (ByVal hwnd As Integer, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer
#End If

Private Const INVALID_HANDLE_VALUE = -1
Private Const MAX_PATH = 260
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function DrawCaption Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long, pcRect As RECT, ByVal un As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As String) As Long
Public Declare Sub ReleaseCapture Lib "user32" ()
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public mNames As Collection, mComments As Collection
Public Const HKEY_LOCAL_MACHINE = &H80000002
Private Const REG_SZ = 1
Private Const DC_ACTIVE = &H1
Private Const DC_ICON = &H4
Private Const DC_TEXT = &H8
Private Const DC_GRADIENT = &H20
Public CheckInterval As Integer
Public NotifyMethod As Integer
Public PSCCheckEnabled As Boolean
Public PSCVoteCheckEnabled As Boolean
Public CurVotes As Long 'last number of votes by selected submission
Public CurComments As Long 'last number of Comments by selected submission
Public Sub SaveSettingString(hKey As Long, strPath As String, strValue As String, strData As String)
    Dim hCurKey As Long 'used to run app at startup
    Dim lRegResult As Long
    lRegResult = RegCreateKey(hKey, strPath, hCurKey)
    lRegResult = RegSetValueEx(hCurKey, strValue, 0, REG_SZ, ByVal strData, Len(strData))
    lRegResult = RegCloseKey(hCurKey)
End Sub
Public Sub DeleteValue(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String)
    Dim hCurKey As Long
    Dim lRegResult As Long
    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
    lRegResult = RegDeleteValue(hCurKey, strValue)
    lRegResult = RegCloseKey(hCurKey)
End Sub
Public Function FileExists(sSource As String) As Boolean
    If Right(sSource, 2) = ":\" Then
        Dim allDrives As String
        allDrives = Space$(64)
        Call GetLogicalDriveStrings(Len(allDrives), allDrives)
        FileExists = InStr(1, allDrives, Left(sSource, 1), 1) > 0
        Exit Function
    Else
        If Not sSource = "" Then
            Dim WFD As WIN32_FIND_DATA
            Dim hFile As Long
            hFile = FindFirstFile(sSource, WFD)
            FileExists = hFile <> INVALID_HANDLE_VALUE
            Call FindClose(hFile)
        Else
            FileExists = False
        End If
    End If
End Function
Public Sub SizeColorCaption() 'paint a titlebar
    Dim R As RECT
    SetRect R, 0, 0, frmTicker.PicTitleBar.ScaleWidth, 16
    frmTicker.PicTitleBar.Cls
    DrawCaption frmTicker.PicTitleBar.hwnd, frmTicker.PicTitleBar.hDC, R, DC_ACTIVE Or DC_ICON Or DC_TEXT Or DC_GRADIENT
    frmTicker.PicTitleBar.CurrentX = 20
    frmTicker.PicTitleBar.CurrentY = 2
    frmTicker.PicTitleBar.Print "PSC Code Ticker"
    frmTicker.ImageList1.ListImages(1).Draw frmTicker.PicTitleBar.hDC, 0, 0, 1
End Sub
Public Sub FormDrag(TheForm As Form) 'allow dragging form
    ReleaseCapture
    Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
End Sub
Public Function ReadINI(filename As String, Section As String, Key As String) As String
    Dim ret As String 'used to get path from an internet shortcut
    Dim Retlen As String
    ret = Space$(255)
    Retlen = GetPrivateProfileString(Section, Key, "", ret, Len(ret), filename)
    ret = Left$(ret, Retlen)
    ReadINI = ret
End Function
Public Function OneGulp(Src As String) As String
    On Error Resume Next
    Dim f As Integer, temp As String
    f = FreeFile
    DoEvents
    Open Src For Binary As #f
    temp = String(LOF(f), Chr$(0))
    Get #f, , temp
    Close #f
    If Left(temp, 2) = "ÿþ" Or Left(temp, 2) = "þÿ" Then temp = Replace(Right(temp, Len(temp) - 2), Chr(0), "")
    OneGulp = temp
End Function
Public Function FileOnly(ByVal filepath As String) As String
    FileOnly = Mid$(filepath, InStrRev(filepath, "\") + 1)
End Function
Public Function ExtOnly(ByVal filepath As String, Optional dot As Boolean) As String
    ExtOnly = Mid$(filepath, InStrRev(filepath, ".") + 1)
    If dot = True Then ExtOnly = "." + ExtOnly
End Function
Public Function ChangeExt(ByVal filepath As String, Optional newext As String) As String
    Dim temp As String
    If InStr(1, filepath, ".") = 0 Then
        temp = filepath
    Else
        temp = Mid$(filepath, 1, InStrRev(filepath, "."))
        temp = Left(temp, Len(temp) - 1)
    End If
    If newext <> "" Then newext = "." + newext
    ChangeExt = temp + newext
End Function


Public Sub DoMessage(mHwnd As Long, msg As String, Optional msgType As VbMsgBoxStyle)
    SetForegroundWindow mHwnd
    SetWindowPos mHwnd, -2, 0, 0, 0, 0, 1 Or 2 'not on top
    MsgBox msg, msgType
    SetWindowPos mHwnd, -1, 0, 0, 0, 0, 1 Or 2 'on top
End Sub

Public Function CountComments(srcStr As String) As Long
    Dim z As Long, z1 As Long, z2 As Long, cnt As Long
    Do
        z = z + 1
        z = InStr(z, srcStr, "<!xmp>")
        If z = 0 Then Exit Do
        z1 = InStr(z + 1, srcStr, "<!/xmp>")
        If z1 <> 0 Then cnt = cnt + 1
    Loop
    CountComments = cnt
End Function
Public Sub FileSave(Text As String, filepath As String)
    On Error Resume Next
    Dim f As Integer
    f = FreeFile
    Open filepath For Binary As #f
    Put #f, , Text
    Close #f
    Exit Sub
End Sub

