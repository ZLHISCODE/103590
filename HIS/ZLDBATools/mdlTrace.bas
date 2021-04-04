Attribute VB_Name = "mdlTrace"
Option Explicit

'公共变量
'Public gfrmFind As New frmFind
Public gstrCompareExe As String
Public gstrLeft As String
Public gstrFilePath As String
Public gcolSort As Collection
Public gblnOwner As Boolean
Public gstrDBUser As String
'********************************************************************
'CommandBar命令ID
Public Enum CommandBarIDCond
    conMenu_FilePopup = 1
    conMenu_EditPopup = 2
    conMenu_ViewPopup = 8
    conMenu_HelpPopup = 9
    
    '添加一个对比功能设置
    conMenu_ComparePopup = 3
    '文件菜单
    conMenu_File_Open = 101
    conMenu_File_CompareExe = 210
    conmenu_File_Logout = 108
    conMenu_File_Exit = 109
    
    '编辑菜单
    conMenu_Edit_Trace = 201
    conMenu_Edit_Trace_1 = 2011
    conMenu_Edit_Trace_4 = 2012
    conMenu_Edit_Trace_8 = 2013
    conMenu_Edit_Trace_12 = 2014
    conMenu_Edit_ChangeReg = 2015
    conMenu_Edit_TraceOff = 202
    conMenu_Edit_CompareLeft = 211
    conMenu_Edit_Compare = 212
    
    '查看菜单
    conMenu_View_Style = 801
    conMenu_View_Style_Report = 8011
    conMenu_View_Style_Table = 8012
    conMenu_View_Filter = 802
    conMenu_View_SQLPrev = 803
    conMenu_View_SQLNext = 804
    conMenu_View_Find = 805
    conMenu_View_FindNext = 806
    conMenu_View_Refresh = 809
    conMenu_View_Close = 810
    
    '帮助菜单
    conMenu_Help_About = 901
End Enum

'CommandBar固有常量定义
Public Const XTP_ID_WINDOW_LIST = 35000 '窗体列表
Public Const XTP_ID_TOOLBARLIST = 59392 '工具栏列表
Public Const ID_INDICATOR_CAPS = 59137 '状态栏（大写）
Public Const ID_INDICATOR_NUM = 59138 '状态栏（数字）
Public Const ID_INDICATOR_SCRL = 59139 '状态栏（滚动）

'CommandBar辅助热键
Public Const FSHIFT = 4
Public Const FCONTROL = 8
Public Const FALT = 16
'********************************************************************
Public Const CB_SETDROPPEDWIDTH As Long = &H160
Public Declare Sub InitCommonControls Lib "comctl32.dll" ()
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long

'-------------------------------------------------------------
Public Const Process_Query_Information = &H400
Public Const Still_Active = &H103
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'-------------------------------------------------------------
Public Const GWL_EXSTYLE = (-20)
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'-------------------------------------------------------------
Public Const EM_LINESCROLL = &HB6 'lngW=横向行数,lngL=纵向行数
Public Const EM_SCROLL = &HB5 '按滚动条几下
Public Const EM_GETFIRSTVISIBLELINE = &HCE 'lngR(>=0)
Public Const EM_GETLINECOUNT = &HBA 'lngR(>=1,包含自动折的行)
Public Const EM_LINELENGTH = &HC1 '第一行未折行前有效
Public Const EM_GETSEL = &HB0
Public Const EM_LINEFROMCHAR = &HC9
Public Const EM_LINEINDEX = &HBB
Public Const EM_SETSEL = &HB1

Public Const FR_DOWN = &H1
Public Const FR_WHOLEWORD = &H2
Public Const FR_MATCHCASE = &H4
Public Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type
Public Type FINDTEXT
    chrg As CHARRANGE
    lpstrText As String
End Type

Public Const WM_USER = &H400
Public Const EM_EXGETSEL = (WM_USER + 52)
Public Const EM_EXSETSEL = (WM_USER + 55)
Public Const EM_FINDTEXT = (WM_USER + 56)
Public Const EM_SETTARGETDEVICE = (WM_USER + 72)
'-------------------------------------------------------------
' Reg Data Types...
Const REG_SZ = 1                         ' Unicode空终结字符串
Const REG_EXPAND_SZ = 2                  ' Unicode空终结字符串
Const REG_DWORD = 4                      ' 32-bit 数字

' 注册表关键字安全选项...
Public Const READ_CONTROL = &H20000
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' 注册表关键字根类型...
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_USERS = &H80000003

' 返回值...
Public Const ERROR_SUCCESS = 0
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long


Public Function GetShortName(ByVal strFile As String) As String
    Dim strShort As String, lngLen As Long
    
    GetShortName = strFile
    
    If InStr(strFile, " ") > 0 Then
        If gobjFile.FileExists(strFile) Then
            GetShortName = gobjFile.GetFile(strFile).ShortPath
        ElseIf gobjFile.FolderExists(strFile) Then
            GetShortName = gobjFile.GetFolder(strFile).ShortPath
        Else
            strShort = Space(255)
            lngLen = GetShortPathName(strFile, strShort, 255)
            GetShortName = Left(strShort, lngLen)
        End If
    End If
End Function


Public Function Decode(ParamArray arrPar() As Variant) As Variant
'功能：模拟Oracle的Decode函数
    Dim varValue As Variant, i As Integer
    
    i = 1
    varValue = arrPar(0)
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            Decode = arrPar(i): Exit Function
        ElseIf varValue = arrPar(i) Then
            Decode = arrPar(i + 1): Exit Function
        Else
            i = i + 2
        End If
    Loop
End Function

Public Sub CboAppendText(cboControl As Object, KeyAscii As Integer)
'功能：对ComboBox实现输入过程中自动完成的功能
'说明：在Combox.KeyPress事件中调用
    Dim strInput As String
    Dim lngIndex As Long
    Const CB_FINDSTRING = &H14C
    
    If cboControl.Style <> 0 Then Exit Sub
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyTab Then Exit Sub
    strInput = Chr(KeyAscii): KeyAscii = 0

    With cboControl
        '接着得到用户击键完成后文本框中出现的内容
        strInput = Mid(.Text, 1, .SelStart) & strInput

        '根据假想的内容得到可能的列表项
        lngIndex = SendMessage(cboControl.hWnd, CB_FINDSTRING, -1, ByVal strInput)
        If lngIndex >= 0 Then
            .ListIndex = lngIndex
            '.Text = .List(lngIndex)
            
            .SelStart = Len(strInput)
            .SelLength = Len(.Text) - Len(strInput)
        Else
            .Text = strInput
            .SelStart = Len(strInput)
        End If
    End With
End Sub

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function


