Attribute VB_Name = "CmnDlgOpenSave"
'just to save comments to file
Option Explicit
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Public Type CMDialog
    ownerform As Long
    Filter As String
    filetitle As String
    FilterIndex As Long
    filename As String
    initdir As String
    dialogtitle As String
    flags As Long
End Type
Private Const OFN_OVERWRITEPROMPT = &H2
Public cmnDlg As CMDialog
Public Sub ShowSave()
    Dim OFName As OPENFILENAME
    With cmnDlg
        OFName.lStructSize = Len(OFName)
        OFName.hwndOwner = .ownerform
        OFName.hInstance = App.hInstance
        OFName.lpstrFilter = Replace(.Filter, "|", Chr(0))
        OFName.lpstrFile = Space$(254)
        OFName.nMaxFile = 255
        OFName.lpstrFileTitle = Space$(254)
        OFName.nMaxFileTitle = 255
        OFName.lpstrInitialDir = .initdir
        OFName.lpstrTitle = .dialogtitle
        OFName.nFilterIndex = .FilterIndex
        OFName.flags = .flags Or OFN_OVERWRITEPROMPT
        If GetSaveFileName(OFName) Then
            .filename = StripTerminator(Trim$(OFName.lpstrFile))
            .filetitle = StripTerminator(Trim$(OFName.lpstrFileTitle))
            .FilterIndex = OFName.nFilterIndex
        End If
    End With
End Sub

Public Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function
