Attribute VB_Name = "mdlSave"
Option Explicit

'OpenFolder函数的回调函数使用
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SELCHANGED = 2
Private Const WM_USER = &H400
Private Const BFFM_SETSELECTION = (WM_USER + 102)
Private Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const BIF_STATUSTEXT = &H4
Private Const MAX_PATH = 260
Private Type BrowseInfo
   hwndOwner      As Long
   pIDLRoot       As Long
   pszDisplayName As Long
   lpszTitle      As Long
   ulFlags        As Long
   lpfnCallback   As Long
   lParam         As Long
   iImage         As Long
End Type
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private mstrAPIPath As String 'OpenFolder初始路径设置


Public Function OpenFolder(ByVal frmodtvOwner As Form, Optional strTitle As String, Optional ByVal strInitDir As String) As String
'    '----------------------------------------------------------------------------------------------------
'    '功能:选择文件夹
'    '参数:frmodtvOwner-选择文件夹的父窗体
'    '       strFolderName-指定的文件夹
'    '       strTitle-标题
'    '       strInitDir-默认打开路径
'    '返回:strFolderName-返回选择的文件夹
'    '----------------------------------------------------------------------------------------------------
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim tBrowseInfo As BrowseInfo
    
    mstrAPIPath = strInitDir & Chr(0)
    With tBrowseInfo
        .hwndOwner = frmodtvOwner.hWnd
        .lpszTitle = lstrcat(strTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_STATUSTEXT
        .lpfnCallback = AddressOfFunction(AddressOf OpenDirCallbackProc)
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    If (lpIDList) Then
       sBuffer = Space(MAX_PATH * 2)
       SHGetPathFromIDList lpIDList, sBuffer
       sBuffer = Left(sBuffer, InStr(sBuffer, Chr(0)) - 1)
       OpenFolder = sBuffer
    End If
End Function

Public Function OpenDirCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long
 '功能：OpenFolder回调函数，用来设置打开的文件的初始路径
    Dim lpIDList As Long
    Dim ret As Long
    Dim sBuffer As String
  
    On Error Resume Next
    
    Select Case uMsg
        Case BFFM_INITIALIZED
            Call SendMessage(hWnd, BFFM_SETSELECTION, 1, ByVal mstrAPIPath)
        Case BFFM_SELCHANGED
            sBuffer = Space(MAX_PATH * 2)
            ret = SHGetPathFromIDList(lp, sBuffer)
            If ret = 1 Then
                Call SendMessage(hWnd, BFFM_SETSTATUSTEXT, 0, ByVal sBuffer)
            End If
    End Select
    
    OpenDirCallbackProc = 0
End Function

Public Function AddressOfFunction(Address As Long) As Long
'功能：OpenFolder子函数
    AddressOfFunction = Address
End Function



