Attribute VB_Name = "mdlSave"
Option Explicit

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

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


'=======================================
'打开文件夹
'=======================================
Public Function GetDirName() As String
    Dim bi As BROWSEINFO
    Dim r As Long
    Dim pidl As Long
    Dim path As String
    Dim pos As Integer
    bi.pidlRoot = 0&

    bi.ulFlags = 1
    bi.lpszTitle = "请选取Trace文件保存路径。"
    pidl = SHBrowseForFolder(bi)
    path = Space$(512)
    r = SHGetPathFromIDList(ByVal pidl&, ByVal path)
    If r Then
    pos = InStr(path, Chr$(0))
    GetDirName = Left(path, pos - 1)
    Else: GetDirName = ""
    End If
End Function

