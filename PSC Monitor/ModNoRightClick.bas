Attribute VB_Name = "ModNoRightClick"
'******************************************************************
'***************Copyright PSST 2001********************************
'***************Written by MrBobo**********************************
'This code was submitted to Planet Source Code (www.planetsourcecode.com)
'If you downloaded it elsewhere, they stole it and I'll eat them alive
'hook mouse activity in the ticker window to disable right click
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lparam As Any) As Long
Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WH_MOUSE = 7
Public Type POINTAPI
    X As Long
    Y As Long
End Type
Public Type MOUSEHOOKSTRUCT
    pt As POINTAPI
    hwnd As Long
    wHitTestCode As Long
    dwExtraInfo As Long
End Type
Public gLngMouseHook As Long
Public Function MouseHookProc(ByVal nCode As Long, ByVal wParam As Long, mhs As MOUSEHOOKSTRUCT) As Long
    Dim strBuffer As String
    Dim lngBufferLen As Long
    Dim strClassName As String
    Dim lngResult As Long
    If (nCode >= 0 And wParam = WM_RBUTTONUP) Then
        strBuffer = Space(255)
        strClassName = "Internet Explorer_Server"
        Debug.Print strClassName
        lngResult = GetClassName(mhs.hwnd, strBuffer, Len(strBuffer))
        Debug.Print Left$(strBuffer, lngResult)
        If lngResult > 0 Then
            If Left$(strBuffer, lngResult) = strClassName Then
                MouseHookProc = 1
                Exit Function
            End If
        End If
    End If
    MouseHookProc = CallNextHookEx(gLngMouseHook, nCode, wParam, mhs)
End Function
