Attribute VB_Name = "mdlUnSendDrug"
Option Explicit

Public Const GWL_EXSTYLE = (-20)
Public Const WinStyle = &H40000
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_USERS = &H80000003

Const REG_SZ = 1
Const REG_EXPAND_SZ = 2
Const REG_DWORD = 4
Const ERROR_NONE = 0
Const ERROR_BADKEY = 2
Const ERROR_ACCESS_DENIED = 8
Const ERROR_SUCCESS = 0

Public gcnOracle As New ADODB.Connection
Public gstrDbUser As String

Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long

Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_READ = KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + READ_CONTROL
Const KEY_WRITE = KEY_SET_VALUE + KEY_CREATE_SUB_KEY + READ_CONTROL
Const KEY_EXECUTE = KEY_READ
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                       
Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String) As String
'??????????????
    Dim i As Long                                           ' ??????????
    Dim rc As Long                                          ' ????????
    Dim hKey As Long                                        ' ??????????????????????
    Dim hDepth As Long                                      '
    Dim sKeyVal As String
    Dim lKeyValType As Long                                 ' ????????????????????
    Dim tmpVal As String                                    ' ????????????????????????
    Dim KeyValSize As Long                                  ' ????????????????????
    
    ' ?? KeyRoot {HKEY_LOCAL_MACHINE...} ??????????????????
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' ????????????????
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' ????????...
    
    tmpVal = String$(1024, 0)                             ' ????????????
    KeyValSize = 1024                                       ' ????????????
    
    '------------------------------------------------------------
    ' ????????????????????...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         lKeyValType, tmpVal, KeyValSize)    ' ????/??????????????
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' ????????
      
    tmpVal = Left$(tmpVal, InStr(tmpVal, Chr(0)) - 1)

    '------------------------------------------------------------
    ' ??????????????????????...
    '------------------------------------------------------------
    Select Case lKeyValType                                  ' ????????????...
    Case REG_SZ, REG_EXPAND_SZ                              ' ??????????????????????????
        sKeyVal = tmpVal                                     ' ??????????????
    Case REG_DWORD                                          ' ??????????????????????????
        For i = Len(tmpVal) To 1 Step -1                    ' ??????????
            sKeyVal = sKeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' ??????????????????????????
        Next
        sKeyVal = Format$("&h" + sKeyVal)                     ' ??????????????????
    End Select
    
    GetKeyValue = sKeyVal                                   ' ??????
    rc = RegCloseKey(hKey)                                  ' ????????????????
    Exit Function                                           ' ????
    
GetKeyError:    ' ????????????????????...
    GetKeyValue = vbNullString                              ' ????????????????????
    rc = RegCloseKey(hKey)                                  ' ????????????????
End Function

Public Function TransColor(ByVal blnFore As Boolean, ByVal intIndex As Integer) As Long
    If blnFore Then
        Select Case intIndex
            Case 1: TransColor = vbRed
            Case 2: TransColor = vbBlue
            Case 3: TransColor = vbYellow
            Case 4: TransColor = vbGreen
            Case 5: TransColor = vbBlack
            Case Else: TransColor = vbWhite
        End Select
    Else
        Select Case intIndex
            Case 1: TransColor = vbRed
            Case 2: TransColor = vbYellow
            Case 3: TransColor = vbGreen
            Case 4: TransColor = vbWhite
            Case 5: TransColor = vbBlack
            Case Else: TransColor = vbBlue
        End Select
    End If
End Function
