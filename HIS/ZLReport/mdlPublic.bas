Attribute VB_Name = "mdlPublic"
Option Explicit
'**************************
'       OEM代号
'
'医业  D2BDD2B5
'托普  CDD0C6D5
'**************************

'------------------------------------------------------------------------------------
Public gblnSilentMode As Boolean            '静默输出报表模式
Public gstrErrorContent As String           '错误内容
Public gcnOLEDB As ADODB.Connection
Public gcnOracle As ADODB.Connection
Public gobjRegister As Object               '注册授权部件zlRegister
Public gclsCNs As RPTDBCNs                  '管理工具设置的其他数据连接对象
Public grsConnect As ADODB.Recordset        '管理工具设置的其他数据连接记录
Public gblnManagementTool As Boolean        '管理工具调用
Public gfrmDBConnect As Object              '数据连接管理窗体对象

Public grsObject As ADODB.Recordset '当前用户所具有Select权限的对象集(用于向导或发布)
Public gblnAutoConnect As Boolean   '是否断网后自动连接数据库
Public gblnExeSQLTest   As Boolean  'SQLTest状态
Public gcolOLEDBConnect As Collection       'OLEDB数据库连接对象缓存
'------------------------------------------------------------------------------------
Public Type CustomPar
    组名 As String
    值列表 As String
    分类SQL As String
    明细SQL As String
    分类字段 As String
    明细字段 As String
    对象 As String
    格式 As Byte
End Type
Public Type ReportData
    DataName As String
    DataSet As ADODB.Recordset
End Type

'1:UBound(Array())=-1；2:Ubound(没传参数赋值)=-1；3:直接UBound()=下标越界
Public gblnError As Boolean
Public garrPars() As Variant '公共参数数组,用于DLL与外部接口
Public garrBill As Variant '打印时的票据号数组

Public glngSys As Long '程序主动调用报表执行或设计接口的系统号

'用于报表对象缓存
Public gobjReport As Report '公共报表对象,用于DLL函数调用
Public grsReport As ADODB.Recordset '最近打开的报表,用于缓存加速,更新zlReports表时需要清除
Public gdatModiTime As Date '最近打开的报表的最近修改时间,用于监视变化
Public gcolPrivs As New Collection
Public gcolRptPriv As Collection
Public gcolUserInfo As Collection

Public gblnSingleTask As Boolean '是否多报表在单任务中打印

Public glngGroup As Long '当前打开为报表组时的组ID,此时gobjReport=Nothing
Public gfrmMain As Object
Public gobjFile As New FileSystemObject

Public lngTXTProc As Long '保存默认的消息函数的地址
Public objClip As RPTItems '剪贴板对象
Public gstrColor As String
Public glngUserID As Long

Private mstrBigTable As String   '大表
Private mstrMiddleTable As String '中表
Private mstrMiddleTableRows As String


Public Const GSTR_SBC = "（＋－＊／＝＜＞）！：１２３４５６７８９０ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ；，。？｜％＃"
Public Const GSTR_DBC = "(+-*/=<>)!:1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZabcedfghijklmnopqrstuvwxyz;,.?|%#"

Private ArrayCompare(20) As Integer  '加密串数组
'------------------------------------------------------------------------------------

'错误日志处理相关变量
Private mlngErrNum As Long, mstrErrInfo As String, mbytErrType As Byte
Private mstrRecentSQL As String  '最近执行的SQL语句

'SQLLog变量
Private msngTime As Single
Private mobjLogText As TextStream

Public gblnRunLog As Boolean '是否记录使用日志
Public gblnErrLog As Boolean '是否记录运行错误
Public gblnReportRunLog As Boolean '是否记录报表运行日志
Public gblnReportUse As Boolean '是否记录报表使用痕迹

'缺省的票据宽度和高度,A4,纵向(系统以Twip作为单位存贮)
Public Const Twip_mm = 56.69286 '单位转换系数
'Public Const Twip_mm = 56.6857142857143
Public Const INIT_WIDTH = 11904
Public Const INIT_HEIGHT = 16832

Public gblnOK As Boolean
Public glngOldProc As Long, glngSelProc As Long
Public gstrFind As String
Public gblnModi As Boolean
Public gstrFonts As String

Public gstrDBUser As String '用户名
Public gstrUserName As String '用户姓名
Public gstrUserNO As String '用户编号
Public gstrLoginUser As String '登录用户名
Public gstrLoginUserName As String '登录用户姓名
Public gstrProductName As String '产品名称
Public gcnOracleConn As String '记录上次连接字符串
Public gstrComputerName As String '记录电脑名称

'API定义
Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private mlngConnectCount As Long

Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002

Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Public Const WM_GETMINMAXINFO = &H24
Public Const GWL_WNDPROC = -4
Public Const WM_CONTEXTMENU = &H7B ' 当右击文本框时，产生这条消息
Type PointAPI
    X As Long
    Y As Long
End Type
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Type Cells
    Row1 As Integer
    Col1 As Integer
    Row2 As Integer
    Col2 As Integer
    Row As Integer
End Type
Type MINMAXINFO
    ptReserved As PointAPI
    ptMaxSize As PointAPI
    ptMaxPosition As PointAPI
    ptMinTrackSize As PointAPI
    ptMaxTrackSize As PointAPI
End Type
Public Type DOCINFO
        cbSize As Long
        lpszDocName As String
        lpszOutput As String
End Type
Public Declare Function StartDoc Lib "gdi32" Alias "StartDocA" (ByVal hDC As Long, lpdi As DOCINFO) As Long
Public Declare Function EndDoc Lib "gdi32" (ByVal hDC As Long) As Long

Public Declare Function SHDeleteKey Lib "shlwapi.dll" Alias "SHDeleteKeyA" (ByVal hKey As Long, ByVal pszSubKey As String) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal Cx As Long, ByVal Cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As PointAPI) As Long
Public Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, ByVal lpPort As String, ByVal iIndex As Long, ByVal lpOutput As String, lpDevMode As Any) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long

Public Const CB_FINDSTRING = &H14C
Public Const CB_GETDROPPEDSTATE = &H157
Public Const CB_SHOWDROPDOWN = &H14F

Public Const DC_PAPERNAMES = 16 '纸张名称(每64字符为一段,以Chr(0)结束)
Public Const DC_PAPERS = 2 '纸张编号(Array or Word)
Public Const DC_BINNAMES = 12 '进纸方式(每24字符为一段,以Chr(0)结束)
Public Const DC_BINS = 6 '进纸编号(Array or Word)

Public Const REG_SZ = 1
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const LVSCW_AUTOSIZE = -1
Public Const LVSCW_AUTOSIZE_USEHEADER = -2
Public Const LVM_SETCOLUMNWIDTH = &H101E
Public Const SWP_NOMOVE = &H2

'打印纸张常量(256=自定义)
Public Const PageSize1 = "信笺， 8 1/2×11 英寸"
Public Const PageSize2 = "+A611 小型信笺， 8 1/2×11 英寸"
Public Const PageSize3 = "小型报， 11×17 英寸"
Public Const PageSize4 = "分类帐， 17×11 英寸"
Public Const PageSize5 = "法律文件， 8 1/2×14 英寸"
Public Const PageSize6 = "声明书，5 1/2×8 1/2 英寸"
Public Const PageSize7 = "行政文件，7 1/2×10 1/2 英寸"
Public Const PageSize8 = "A3, 297×420 毫米"
Public Const PageSize9 = "A4, 210×297 毫米"
Public Const PageSize10 = "A4小号， 210×297 毫米"
Public Const PageSize11 = "A5, 148×210 毫米"
Public Const PageSize12 = "B4, 250×354 毫米"
Public Const PageSize13 = "B5, 182×257 毫米"
Public Const PageSize14 = "对开本， 8 1/2×13 英寸"
Public Const PageSize15 = "四开本， 215×275 毫米"
Public Const PageSize16 = "10×14 英寸"
Public Const PageSize17 = "11×17 英寸"
Public Const PageSize18 = "便条，8 1/2×11 英寸"
Public Const PageSize19 = "#9 信封， 3 7/8×8 7/8 英寸"
Public Const PageSize20 = "#10 信封， 4 1/8×9 1/2 英寸"
Public Const PageSize21 = "#11 信封， 4 1/2×10 3/8 英寸"
Public Const PageSize22 = "#12 信封， 4 1/2×11 英寸"
Public Const PageSize23 = "#14 信封， 5×11 1/2 英寸"
Public Const PageSize24 = "C 尺寸工作单"
Public Const PageSize25 = "D 尺寸工作单"
Public Const PageSize26 = "E 尺寸工作单"
Public Const PageSize27 = "DL 型信封， 110×220 毫米"
Public Const PageSize28 = "C5 型信封， 162×229 毫米"
Public Const PageSize29 = "C3 型信封， 324×458 毫米"
Public Const PageSize30 = "C4 型信封， 229×324 毫米"
Public Const PageSize31 = "C6 型信封， 114×162 毫米"
Public Const PageSize32 = "C65 型信封，114×229 毫米"
Public Const PageSize33 = "B4 型信封， 250×353 毫米"
Public Const PageSize34 = "B5 型信封，176×250 毫米"
Public Const PageSize35 = "B6 型信封， 176×125 毫米"
Public Const PageSize36 = "信封， 110×230 毫米"
Public Const PageSize37 = "信封大王， 3 7/8×7 1/2 英寸"
Public Const PageSize38 = "信封， 3 5/8×6 1/2 英寸"
Public Const PageSize39 = "U.S. 标准复写簿， 14 7/8×11 英寸"
Public Const PageSize40 = "德国标准复写簿， 8 1/2×12 英寸"
Public Const PageSize41 = "德国法律复写簿， 8 1/2×13 英寸"

'自定义纸张
Public Const PageCustom1 = "穿孔打印纸(不分份)，241×280 毫米"
Public Const PageCustom2 = "穿孔打印纸(两等份)，241×140 毫米"
Public Const PageCustom3 = "穿孔打印纸(三等份)，241×94 毫米"

'控制TAB键的函数
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Const WH_KEYBOARD = 2
Public Const HC_ACTION = 0
Public Const HC_NOREMOVE = 3

Public glngKeyHook As Long
Public gobjTab As clsTabInput
'Html Help
Public Declare Function Htmlhelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Any) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'Public Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long

Public Const HH_DISPLAY_TOPIC = &H0

'Window版本函数
Type OSVERSIONINFO 'for GetVersionEx API call
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

'纸张打印边界控制================================================================
Public Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
'不同打印机的打印单元精度不同
Public Const PHYSICALWIDTH = 110   'Physical Width in device units
Public Const PHYSICALHEIGHT = 111  'Physical Height in device units
Public Const PHYSICALOFFSETX = 112 'Physical Printable Area x margin
Public Const PHYSICALOFFSETY = 113 'Physical Printable Area y margin
Public Const LOGPIXELSX = 88 'Number of pixels per logical inch along the screen width
Public Const LOGPIXELSY = 90
Public Const SCALINGFACTORX = 114  'Scaling factor x
Public Const SCALINGFACTORY = 115  'Scaling factor y
Public Const DRIVERVERSION = 0     'Device driver version

'WinNT自定义纸张控制================================================================
Public Const ZL_FORM_NAME = "zlBillPaper"

'Custom constants for this sample's SelectForm function
Public Const FORM_NOT_SELECTED = 0
Public Const FORM_SELECTED = 1
Public Const FORM_ADDED = 2

Public Type RECTL
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Type SIZEL
    Cx As Long
    Cy As Long
End Type
Public Type SECURITY_DESCRIPTOR
    Revision As Byte
    Sbz1 As Byte
    Control As Long
    Owner As Long
    Group As Long
    Sacl As Long  'ACL
    Dacl As Long  'ACL
End Type
'The two definitions for FORM_INFO_1 make the coding easier.
Public Type FORM_INFO_1
    Flags As Long
    pName As Long   'String
    Size As SIZEL
    ImageableArea As RECTL
End Type
Public Type sFORM_INFO_1
    Flags As Long
    pName As String
    Size As SIZEL
    ImageableArea As RECTL
End Type
'Optional functions not used in this sample, but may be useful.
Public Declare Function DeleteForm Lib "winspool.drv" Alias "DeleteFormA" (ByVal hPrinter As Long, ByVal pFormName As String) As Long
Public Declare Function EnumForms Lib "winspool.drv" Alias "EnumFormsA" (ByVal hPrinter As Long, ByVal Level As Long, ByRef pForm As Any, ByVal cbBuf As Long, ByRef pcbNeeded As Long, ByRef pcReturned As Long) As Long
Public Declare Function AddForm Lib "winspool.drv" Alias "AddFormA" (ByVal hPrinter As Long, ByVal Level As Long, pForm As Byte) As Long
Public Declare Function GetForm Lib "winspool.drv" Alias "GetFormA" (ByVal hPrinter As Long, ByVal pFormName As String, ByVal Level As Long, pForm As Byte, ByVal cbBuf As Long, pcbNeeded As Long) As Long
Public Declare Function SetForm Lib "winspool.drv" Alias "SetFormA" (ByVal hPrinter As Long, ByVal pFormName As String, ByVal Level As Long, pForm As Byte) As Long

Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByRef lpString2 As Long) As Long

'以下仅为新的打印方式使用-----------------------------------------------------------
'注意以dmFields是Long型,as Long或尾部加&符
Public Const DM_ORIENTATION = &H1&
Public Const DM_PAPERSIZE = &H2&
Public Const DM_PAPERLENGTH = &H4&
Public Const DM_PAPERWIDTH = &H8&
Public Const DM_COPIES = &H100&
Public Const DM_DEFAULTSOURCE = &H200&
Public Const DM_COLLATE = &H8000&
Public Const DM_FORMNAME = &H10000
'Constants for DocumentProperties() call
Public Const DM_COPY = 2
Public Const DM_OUT_BUFFER = DM_COPY
Public Const DM_PROMPT = 4
Public Const DM_IN_PROMPT = DM_PROMPT
Public Const DM_MODIFY = 8
Public Const DM_IN_BUFFER = DM_MODIFY
'Constants for DocumentProperties() return
Public Const IDOK = 1
Public Const IDCANCEL = 2
'Constants for DEVMODE
Public Const CCHFORMNAME = 32
Public Const CCHDEVICENAME = 32

Public Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Long
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type
Public Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, ByVal pDefault As Long) As Long
Public Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Public Declare Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal hwnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, pDevModeOutput As Any, pDevModeInput As Any, ByVal fMode As Long) As Long
Public Declare Function ResetDC Lib "gdi32" Alias "ResetDCA" (ByVal hDC As Long, lpInitData As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long

'目录选择对话框函数=================================================================
Public gstrAPIPath As String

Private Const MSTR_DBLINK_KEY As String = "zLw09OewKKO1`;owEWO-=,./w[]wwqq3##=``44314325"

Private Type BrowseInfo
  hWndOwner      As Long
  pIDLRoot       As Long
  pszDisplayName As String
  lpszTitle      As String
  ulFlags        As Long
  lpfnCallback   As Long
  lParam         As Long
  iImage         As Long
End Type

Private Const BIF_STATUSTEXT = &H4
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2

Private Const WM_USER = &H400
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SELCHANGED = 2
Private Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
Private Const BFFM_SETSELECTION = (WM_USER + 102)

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
        

'===================================================================================

'鼠标中键函数========================================================================
Public Oldwinproc As Long
Public Const WM_COMMAND = &H111
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MOUSEWHEEL = &H20A
    
Public Function FlexScroll(ByVal hwnd As Long, ByVal wMsg As Long, _
                           ByVal wParam As Long, ByVal lParam As Long) As Long
'支持滚轮的滚动
    Select Case wMsg
    Case WM_MOUSEWHEEL
        Select Case wParam
        Case -7864320  '向下滚
            SendKeys "{PGDN}"
        Case 7864320   '向上滚
            SendKeys "{PGUP}"
        End Select

    End Select
    FlexScroll = CallWindowProc(Oldwinproc, hwnd, wMsg, wParam, lParam)
End Function
'===================================================================================
Public Function Lpad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:按指定长度填制空格
    '--入参数:
    '--出参数:
    '--返  回:返回字串
    '-----------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = LenB(StrConv(strCode, vbFromUnicode))
    If lngTmp < lngLen Then
        strTmp = String(lngLen - lngTmp, strChar) & strTmp
    ElseIf lngTmp > lngLen Then  '大于长度时,自动载断
        strTmp = strCode
    End If
    Lpad = Replace(strTmp, Chr(0), strChar)
End Function
Public Function RPAD(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:按指定长度填制空格
    '--入参数:
    '--出参数:
    '--返  回:返回字串
    '-----------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = LenB(StrConv(strCode, vbFromUnicode))
    If lngTmp < lngLen Then
        strTmp = strTmp & String(lngLen - lngTmp, strChar)
    Else
        '主要有空格引起的
        strTmp = strCode
    End If
    '取掉最后半个字符
    RPAD = Replace(strTmp, Chr(0), strChar)
End Function

Public Function BrowseForFolder(ByVal hwnd As Long, ByVal Title As String, ByVal InitDir As String) As String
    Dim lpIDList As Long
    Dim szTitle As String
    Dim sBuffer As String
    Dim tBrowseInfo As BrowseInfo
    
    gstrAPIPath = InitDir & Chr(0)
    
    szTitle = Title
    
    With tBrowseInfo
        .hWndOwner = hwnd
        .lpszTitle = szTitle
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_STATUSTEXT
        .lpfnCallback = AddressOfFunction(AddressOf BrowseCallbackProc)
    End With
    
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    
    If lpIDList <> 0 Then
        sBuffer = Space(512)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, Chr(0)) - 1)
        BrowseForFolder = sBuffer
    End If
End Function
 
Private Function BrowseCallbackProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long
    Dim lpIDList As Long
    Dim ret As Long
    Dim sBuffer As String
  
    On Error Resume Next
    
    Select Case uMsg
        Case BFFM_INITIALIZED
            Call SendMessage(hwnd, BFFM_SETSELECTION, 1, ByVal gstrAPIPath)
        Case BFFM_SELCHANGED
            sBuffer = Space(512)
            ret = SHGetPathFromIDList(lp, sBuffer)
            If ret = 1 Then
                Call SendMessage(hwnd, BFFM_SETSTATUSTEXT, 0, ByVal sBuffer)
            End If
    End Select
    
    BrowseCallbackProc = 0
End Function

Private Function AddressOfFunction(Address As Long) As Long
    AddressOfFunction = Address
End Function

Public Function IsWindowsNT() As Boolean
'功能：是否WindowNT操作系统
    Const dwMaskNT = &H2&
    IsWindowsNT = (GetWinPlatform() And dwMaskNT)
End Function

Public Function IsWindows95() As Boolean
'功能：是否Window95操作系统
    Const dwMask95 = &H1&
    IsWindows95 = (GetWinPlatform() And dwMask95)
End Function
 
Private Function GetWinPlatform() As Long
    Dim osvi As OSVERSIONINFO
    Dim strCSDVersion As String
    osvi.dwOSVersionInfoSize = Len(osvi)
    If GetVersionEx(osvi) = 0 Then
        Exit Function
    End If
    GetWinPlatform = osvi.dwPlatformId
End Function

Public Function CustomHook(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'说明：
'   Code=Hook Code(HC_ACTION、HC_NOREMOVE)
'   wParam=Virtual-Key Code
'   lParam=0-15位(按键的重复次数)
'          16-23位(OEM Scan Code)
'          24位(是否扩展键,如Fx,小键盘键)
'          25-28位(保留)
'          29(ALT是否按下)
'          30(发送消息之前键是否按下)
'          31(0-正在按下,1-正在松开)
    Static blnShift As Boolean
    
    If wParam = vbKeyShift Then
        If lParam > 0 Then
            blnShift = True
        ElseIf lParam < 0 Then
            blnShift = False
        End If
    End If
    If wParam = vbKeyTab Then
        CustomHook = 1
        If blnShift Then
            If lParam > 0 Then
                gobjTab.ACT_sTabKeyDown
            ElseIf lParam < 0 Then
                gobjTab.ACT_sTabKeyUp
            End If
        Else
            If lParam > 0 Then
                gobjTab.ACT_TabKeyDown
            ElseIf lParam < 0 Then
                gobjTab.ACT_TabKeyUp
            End If
        End If
    Else
        CallNextHookEx glngKeyHook, Code, wParam, lParam
    End If
End Function

Public Sub RegReportFile()
'功能：注册中联报表文件
    Dim strSys As String * 255
    
    GetSystemDirectory strSys, 255
    
    RegSetValue HKEY_CLASSES_ROOT, ".zlr", REG_SZ, "zlReport", 7
    RegSetValue HKEY_CLASSES_ROOT, "zlReport", REG_SZ, "自定义报表文件", 7
    RegSetValue HKEY_CLASSES_ROOT, "zlReport\DefaultIcon", REG_SZ, Left(strSys, InStr(strSys, Chr(0)) - 1) & "\zl9Report.dll,0", 24
    RegSetValue HKEY_CLASSES_ROOT, "zlReport\Shell", REG_SZ, "Read", 4
    RegSetValue HKEY_CLASSES_ROOT, "zlReport\Shell\Read", REG_SZ, "打开自定义报表文件(&1)", 12
    RegSetValue HKEY_CLASSES_ROOT, "zlReport\Shell\Read\Command", REG_SZ, "NotePad.exe ""%1""", 22
End Sub

Public Function GetPaperName(ByVal intSize As Integer, Optional ByVal lngW As Long, Optional ByVal lngH As Long) As String
'功能： 根据当前打印机的设置，获取纸张名称
'参数： lngW,lngH=自定义纸张的宽高(Twip)
'返回： 纸张名称
    If intSize = 256 Then
        If CInt(lngW / Twip_mm) = 241 And CInt(lngH / Twip_mm) = 280 Then
            GetPaperName = PageCustom1
        ElseIf CInt(lngW / Twip_mm) = 241 And CInt(lngH / Twip_mm) = 140 Then
            GetPaperName = PageCustom2
        ElseIf CInt(lngW / Twip_mm) = 241 And CInt(lngH / Twip_mm) = 94 Then
            GetPaperName = PageCustom3
        Else
            GetPaperName = "用户自定义 ..."
        End If
    ElseIf intSize >= 1 And intSize <= 41 Then
        GetPaperName = Switch( _
            intSize = 1, PageSize1, intSize = 2, PageSize2, intSize = 3, PageSize3, intSize = 4, PageSize4, intSize = 5, PageSize5, _
            intSize = 6, PageSize6, intSize = 7, PageSize7, intSize = 8, PageSize8, intSize = 9, PageSize9, intSize = 10, PageSize10, _
            intSize = 11, PageSize11, intSize = 12, PageSize12, intSize = 13, PageSize13, intSize = 14, PageSize14, intSize = 15, PageSize15, _
            intSize = 16, PageSize16, intSize = 17, PageSize17, intSize = 18, PageSize18, intSize = 19, PageSize19, intSize = 20, PageSize20, _
            intSize = 21, PageSize21, intSize = 22, PageSize22, intSize = 23, PageSize23, intSize = 24, PageSize24, intSize = 25, PageSize25, _
            intSize = 26, PageSize26, intSize = 27, PageSize27, intSize = 28, PageSize28, intSize = 29, PageSize29, intSize = 30, PageSize30, _
            intSize = 31, PageSize31, intSize = 32, PageSize32, intSize = 33, PageSize33, intSize = 34, PageSize34, intSize = 35, PageSize35, _
            intSize = 36, PageSize36, intSize = 37, PageSize37, intSize = 38, PageSize38, intSize = 39, PageSize39, intSize = 40, PageSize40, _
            intSize = 41, PageSize41)
    Else
        GetPaperName = "不可测的纸张 ..."
    End If
End Function

Public Sub SetComboBoxHeight(cbo As ComboBox, lngH As Long)
'功能：设置下拉列表框尺寸,以像素为单位
    MoveWindow cbo.hwnd, cbo.Left / 15, cbo.Top / 15, cbo.Width / 15, lngH, 1
End Sub

Public Function CustomMessage(ByVal hwnd As Long, ByVal Msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
    If Msg = WM_GETMINMAXINFO Then

        Dim MinMax As MINMAXINFO
        CopyMemory MinMax, ByVal lp, Len(MinMax)
        MinMax.ptMinTrackSize.X = 9300 \ 15
        MinMax.ptMinTrackSize.Y = 6800 \ 15
        MinMax.ptMaxTrackSize.X = 1600
        MinMax.ptMaxTrackSize.Y = 1200
        CopyMemory ByVal lp, MinMax, Len(MinMax)
        CustomMessage = 1
        Exit Function
    End If
    CustomMessage = CallWindowProc(glngOldProc, hwnd, Msg, wp, lp)
End Function

Public Function ScrollExist(msh As Object) As Boolean
'功能:判断网格是否有垂直滚动条
'说明:各行高必须一样
    If msh.RowHeight(0) * msh.Rows >= msh.Height Then
        ScrollExist = True
    Else
        ScrollExist = False
    End If
End Function

Private Function mGetInvalidTable() As String
'功能：得到在最近使用的SQL语句中不能访问的表或视图
    Dim varTables As Variant
    Dim strTable As String, lngCount As Long
    Dim strInvalidTable As String
    
    varTables = Split(SQLObject(mstrRecentSQL), ",")
    
    On Error Resume Next
    For lngCount = LBound(varTables) To LBound(varTables)
        strTable = varTables(lngCount)
        
        '测试该对象是否可用
        gcnOracle.Execute "select 1 from " & strTable & " where rownum<1"
        If Err <> 0 Then
            Err.Clear
            strInvalidTable = strInvalidTable & "," & strTable
        End If
    Next
    
    If strInvalidTable <> "" Then
        '去掉第一个逗号
        mGetInvalidTable = Mid(strInvalidTable, 2)
    End If
End Function

Public Function ErrCenter() As Byte
'------------------------------------------------
'功能： 数据事务错误处理中心
'参数：
'返回： cancel      返回 0
'       resume      返回 1
'------------------------------------------------
    Dim strNote As String, strTemp As String
    Dim bytReturnType As Byte
    Dim blnExeSQLTest As Boolean
    Static mstrErrRecentSQL As String
    
    bytReturnType = 1
    If gcnOracle.Errors.count <> 0 Then
        'PL/SQL存储过程错误
        If gcnOracle.Errors(0).NativeError >= 20000 And gcnOracle.Errors(0).NativeError <= 20200 Then
            '日志变量
            mbytErrType = 1
            mlngErrNum = gcnOracle.Errors(0).NativeError
            mstrErrInfo = gcnOracle.Errors(0).Description
            
            strNote = gcnOracle.Errors(0).Description
            MsgBox Split(strNote, "[ZLSOFT]")(1), vbExclamation, App.Title
            Exit Function
        End If
        'ORACLE其它错误
        '日志变量
        mbytErrType = 2
        mlngErrNum = gcnOracle.Errors(0).NativeError
        mstrErrInfo = gcnOracle.Errors(0).Description
        
        Select Case gcnOracle.Errors(0).NativeError
        Case 1
            strNote = "已经存在相同内容的数据（要求唯一的内容[如编号、名称等]有重复）。"
            bytReturnType = 0
        Case 903
            strNote = "表名称错误。"
        Case 904
            strNote = "列名称错误。"
        Case 942
            strNote = "表或视图不存在，很可能是你不具备使用该部分数据的权限或该部分对象同义词缺失。"
            bytReturnType = 0
            
            strTemp = mGetInvalidTable()
            If strTemp <> "" Then
                mstrErrInfo = "（建议请系统管理员重新授权、修复同义词）" & vbCrLf & "请对下列对象进行检查：" & vbCrLf & vbCrLf & vbTab & strTemp
            Else
                mstrErrInfo = "（建议请系统管理员重新授权、修复同义词）" & vbCrLf & "错误SQL语句为：" & vbCrLf & vbCrLf & mstrRecentSQL
            End If
        Case 1000
            strNote = "打开的数据表太多，必要时请系统管理员修改数据库的Open_Cursors配置。"
        Case 1005
            strNote = "错误的用户名或密码。"
        Case 1017
            strNote = "错误的用户名或密码。"
            bytReturnType = 0
        Case 1031
            strNote = "没有足够的权限。"
            bytReturnType = 0
        Case 1045
            strNote = "没有联结数据库的权限。"
            bytReturnType = 0
        Case 1400
            strNote = "由于给主键或要求非空列赋予了空值，导致增加失败。"
            bytReturnType = 0
        Case 1401
            strNote = "由于赋予的值超过了列宽限制，导致增加或更新失败。"
            bytReturnType = 0
        Case 1402
            strNote = "由于赋予的值不符合视图的条件限制，导致增加或更新失败。"
            bytReturnType = 0
        Case 1403
            strNote = "由于未检索到数据，导致后续处理失败。"
        Case 1404
            strNote = "修改列操作，导致相关的索引太大。"
        Case 1405
            strNote = "取得的列值为空。"
        Case 1406
            strNote = "取得的列值被切断而缩短了。"
        Case 1407
            strNote = "由于给主键或要求非空列赋予了空值，导致更新失败。"
            bytReturnType = 0
        Case 1408
            strNote = "指定的列已经建立了索引。"
        Case 1409
            strNote = "不能进行无顺序操作(NoSort)，因为本身就没排序。"
        Case 1410
            strNote = "错误的行ID(ROWID)，行ID必须是数字和字符组成的16进制格式。"
        Case 1411
            strNote = "当前列不能存储超过64K的数据。"
            bytReturnType = 0
        Case 1412
            strNote = "当前列数据类型不能存储零长度字符串。"
            bytReturnType = 0
        Case 1413
            strNote = "错误的小数位数，导致失败。"
            bytReturnType = 0
        Case 1415
            strNote = "不能对一个标签伪列指定外连接[Outer-Join(+)]"
        Case 1416
            strNote = "两张表不能同时指向一个外连接[Outer-Join(+)]"
        Case 1417
            strNote = "一张表只能指定指向不超过一张表的外连接[Outer-Join(+)]"
        Case 1418
            strNote = "指定的索引不存在。"
        Case 1424
            strNote = "错误或无效的换码字符(通配符中只能是'%'或'_')。"
        Case 1425
            strNote = "换码字符必须是长度为1的字符。"
        Case 1426
            strNote = "数值表达式的数据溢出(太大或太小)。"
        Case 1427
            strNote = "单行子查询返回了多行。"
        Case 1428
            strNote = "函数的参数错误或超界。"
        Case 1429
            strNote = "一个二进制日期格式超界。"
        Case 1430
            strNote = "希望增加的列已经存在。"
        Case 1431
            strNote = "授权命令(GRANT)导致内在的不一致。"
        Case 1432
            strNote = "希望删除的公共同义词已经不存在。"
        Case 1433
            strNote = "希望建立的同义词已经存在。"
        Case 1434
            strNote = "希望删除的同义词已经不存在。"
        Case 1435
            strNote = "指定的用户不存在。"
            bytReturnType = 0
        Case 1438
            strNote = "数值超过了列允许的精确程度。"
        Case 1439, 1440, 1441
            strNote = "只有空值列才能修改数据类型、将精度或尺寸减小"
        Case 1536
            strNote = "某个超出表空间的空间限量。"
        Case 2290
            strNote = "由于项目值超过允许的范围（违背了检查约束），导致增加或更新失败。"
            bytReturnType = 0
        Case 2291
            strNote = "由于未填写相关表中存在的项目值(违背了外键约束)，导致增加或更新失败。"
        Case 2292
            strNote = "因为该记录已经使用，故不能删除此记录。"
            bytReturnType = 0
        Case 12203
            strNote = "由于主机串书写、配置或服务器问题，不能正常连接。"
            bytReturnType = 0
        Case Else
            strTemp = Err.Description
            If InStr(strTemp, "PLS-00201") > 0 And InStr(strTemp, "ZL_") > 0 Then
                Dim lngPos As Long
                
                lngPos = InStr(strTemp, "ZL_")
                strTemp = Mid(strTemp, lngPos)
                strTemp = Mid(strTemp, 1, InStr(strTemp, "'") - 1)
                
                strNote = "请在服务器管理工具的角色管理程序中增加对过程"" & strTemp & ""的授权。"
            Else
                strNote = "未知错误，发生在" & gcnOracle.Errors(0).Source
            End If
        End Select
        
    Else
        'VB标准错误
        '日志变量
        mbytErrType = 3
        mlngErrNum = Err.Number
        mstrErrInfo = Err.Description
        
        Select Case Err.Number
            Case 3, 3 - 2146828288
                strNote = "未采用标准返回过程"
            Case 5, 5 - 2146828288
                strNote = "无效的过程或参数"
            Case 6, 6 - 2146828288
                strNote = "数据溢出"
            Case 7, 7 - 2146828288
                strNote = "内存溢出"
            Case 9, 9 - 2146828288
                strNote = "下标超界"
            Case 10, 10 - 2146828288
                strNote = "数组是固定数组或暂时锁定"
            Case 11, 11 - 2146828288
                strNote = "除数为零太小"
            Case 13, 13 - 2146828288
                strNote = "类型不匹配"
            Case 14, 14 - 2146828288
                strNote = "超过字符串允许长度"
            Case 16, 16 - 2146828288
                strNote = "表达式太复杂"
            Case 17, 17 - 2146828288
                strNote = "不支持要求的操作"
            Case 18, 18 - 2146828288
                strNote = "发生了用户中断"
            Case 20, 20 - 2146828288
                strNote = "无错误返回"
            Case 28, 28 - 2146828288
                strNote = "堆栈空间溢出"
            Case 35, 35 - 2146828288
                strNote = "过程或函数未定义"
            Case 47, 47 - 2146828288
                strNote = " 太多的动态联结库（DLL）应用客户"
            Case 48, 48 - 2146828288
                strNote = " 调用动态联结库（DLL）错误"
            Case 49, 49 - 2146828288
                strNote = " 动态联结库（DLL）约定错误"
            Case 51, 51 - 2146828288
                strNote = "内部错误"
            Case 52, 52 - 2146828288
                strNote = "错误的文件名或文件号"
            Case 53, 53 - 2146828288
                strNote = "文件未找到"
            Case 54, 54 - 2146828288
                strNote = "文件格式错误"
            Case 55, 55 - 2146828288
                strNote = "文件已经打开"
            Case 57, 57 - 2146828288
                strNote = "设备输入 / 输出错误"
            Case 58, 58 - 2146828288
                strNote = "文件已经存在"
            Case 59, 59 - 2146828288
                strNote = "错误的记录长度"
            Case 61, 61 - 2146828288
                strNote = "磁盘满"
            Case 62, 62 - 2146828288
                strNote = "输入超过文件尾"
            Case 63, 63 - 2146828288
                strNote = "错误的记录号"
            Case 67, 67 - 2146828288
                strNote = "文件太多"
            Case 68, 68 - 2146828288
                strNote = "设备无效或不支持"
            Case 70, 70 - 2146828288
                strNote = "拒绝访问"
            Case 71, 71 - 2146828288
                strNote = "磁盘未准备好"
            Case 74, 74 - 2146828288
                strNote = "不能命名为不同的驱动器"
            Case 75, 75 - 2146828288
                strNote = "路径 / 文件访问错误"
            Case 76, 76 - 2146828288
                strNote = "路径未找到"
            Case 91, 91 - 2146828288
                strNote = "对象变量或块变量为定义(未新建实例)"
            Case 92, 92 - 2146828288
                strNote = "循环未初始化"
            Case 93, 93 - 2146828288
                strNote = "错误的模式字符串"
            Case 94, 94 - 2146828288
                strNote = "错误地使用空(Null)"
            Case 96, 96 - 2146828288
                strNote = " 由于已经使用的对象时间超过了其设置的最大元素号，导致不可能进入事件"
            Case 97, 97 - 2146828288
                strNote = "不能调用一个未建立实例的类对象函数"
            Case 98, 98 - 2146828288
                strNote = " 不能使用一个私有对象的属性和方法?参数和返回值"
            Case 321, 321 - 2146828288
                strNote = "错误的文件格式"
            Case 322, 322 - 2146828288
                strNote = "不能创建需要的临时文件"
            Case 325, 325 - 2146828288
                strNote = "资源文件中错误的格式"
            Case 380, 380 - 2146828288
                strNote = "错误的属性值"
            Case 381, 381 - 2146828288
                strNote = "错误的属性数组索引"
            Case 382, 382 - 2146828288
                strNote = "不支持的运行时设置"
            Case 383, 383 - 2146828288
                strNote = "不支持的只读属性设置"
            Case 385, 384 - 2146828288
                strNote = "需要属性数组索引"
            Case 387, 387 - 2146828288
                strNote = "不允许的设置"
            Case 393, 393 - 2146828288
                strNote = "不支持的运行时读取"
            Case 394, 394 - 2146828288
                strNote = "不支持的只写属性读取"
            Case 422, 422 - 2146828288
                strNote = "不存在的属性"
            Case 423, 423 - 2146828288
                strNote = "不存在的属性或方法"
            Case 424, 424 - 2146828288
                strNote = "要求一个对象"
            Case 429, 429 - 2146828288
                strNote = "ActiveX不能创建部件"
            Case 430, 430 - 2146828288
                strNote = "类不支持的自动化操作或不支持的界面"
            Case 432, 432 - 2146828288
                strNote = "在自动操作期间未找到文件名或类名称"
            Case 438, 438 - 2146828288
                strNote = "对象不支持该属性或方法"
            Case 440, 440 - 2146828288
                strNote = "自动化对象错误"
            Case 442, 442 - 2146828288
                strNote = "到远程类库或对象库的联结丢失，按OK进入对话移去参照"
            Case 443, 443 - 2146828288
                strNote = "自动化对象没有缺省值"
            Case 445, 445 - 2146828288
                strNote = "对象不支持这种操作"
            Case 446, 446 - 2146828288
                strNote = "对象不支持命名参数"
            Case 447, 447 - 2146828288
                strNote = "对象不支持当前本地设置"
            Case 448, 448 - 2146828288
                strNote = "命名参数未找到"
            Case 449, 449 - 2146828288
                strNote = "参数不是可选的"
            Case 450, 450 - 2146828288
                strNote = "错误的参数个数和属性分配"
            Case 451, 451 - 2146828288
                strNote = "属性赋值(Let)过程和读取(Get)过程不返回对象"
            Case 452, 452 - 2146828288
                strNote = "无效的序号"
            Case 453, 453 - 2146828288
                strNote = "指定的DLL函数未找到"
            Case 454, 454 - 2146828288
                strNote = "代码资源未找到"
            Case 455, 455 - 2146828288
                strNote = "代码资源锁定错误"
            Case 457, 457 - 2146828288
                strNote = "该关键值已经与集合的另一元素结合"
            Case 458, 458 - 2146828288
                strNote = "VB不支持的可变自动化类型"
            Case 459, 459 - 2146828288
                strNote = "对象和类不支持的事件集"
            Case 460, 460 - 2146828288
                strNote = "错误的剪贴板格式"
            Case 461, 461 - 2146828288
                strNote = "方法或数据成员未找到"
            Case 462, 462 - 2146828288
                strNote = "远程服务器不存在或无效"
            Case 463, 463 - 2146828288
                strNote = "类没有在本地注册"
            Case 481, 481 - 2146828288
                strNote = "无效的图片格式"
            Case 482, 482 - 2146828288
                strNote = "打印机错误"
            Case 735, 735 - 2146828288
                strNote = "不能将存储为临时文件"
            Case 744, 744 - 2146828288
                strNote = "未找到搜索的主题"
            Case 746, 746 - 2146828288
                strNote = "太长的复制"
            'ADO错误
            Case 3001
                strNote = "参数类型错误，或数值超过范围，或互相冲突。"
            Case 3021
                strNote = "记录超界(EOF/BOF)，或者当前记录被删除；当前应用操作需要定位当前记录。"
            Case 3219
                strNote = "上下文环境不允许当前应用操作（可能是处于尚未结束的事务）。"
            Case 3246
                strNote = "在事务执行中，不能关闭一个联结对象。"
            Case 3251
                strNote = "当前基础不支持这一应用操作。"
            Case 3265
                strNote = "ADO没找到应用程序要求的对应名称或序号。"
            Case 3367
                strNote = "对象已经存在，不能添加。"
            Case 3420
                strNote = "对象未引用。"
            Case 3421
                strNote = "当前操作使用了错误的数值类型。"
            Case 3704
                strNote = "对象关闭时，当前操作不能执行。"
            Case 3705
                strNote = "对象开启时，当前操作不能执行。"
            Case 3706
                strNote = "ADO没找到指定的支持。"
            Case 3707
                strNote = "不能采用命令对象改变一个记录集的活动连接源等属性。"
            Case 3708
                strNote = "应用程序出现错误的参数定义。"
            Case 3709
                strNote = "应用程序要求一个关闭的引用对象或无效的联结对象。"
            Case Else
                strNote = "发生在界面未知错误"
        End Select
        bytReturnType = 0
    End If
    
    
    If gblnAutoConnect Then '是否使用网络断开自动连接功能
        Dim blnConnect As Boolean
        Dim blnNumConnect As Boolean '检查次数是否重新连接
        Dim blnStatus As Boolean '是否其他错误引发的网络问题
        '通过过滤错误信息,检查是否为网络问题引发的错误。mbytErrType=2 Oracle提供的错误信息 mbytErrType=3 VB提供的错误信息
        If mbytErrType = 3 Then
            If mlngErrNum = -2147467259 Or mlngErrNum = -2147217900 Or mlngErrNum = 3709 Then
                '检查VB具体错误信息
                If CheckErrConnectInfo(mlngErrNum, strNote, mstrErrInfo, 1) Then
                    '判断相同错误,如果2次以上正常错误提示。
                    If mstrErrRecentSQL = mstrRecentSQL And mstrRecentSQL <> "" Then
                        mlngConnectCount = mlngConnectCount + 1
                        If mlngConnectCount > 2 Then
                            blnNumConnect = False  '正常错误提示
                            mlngConnectCount = 0 '还原计数器
                        Else
                            blnNumConnect = True
                        End If
                    Else
                        mstrErrRecentSQL = mstrRecentSQL
                        mlngConnectCount = 1
                        blnNumConnect = True
                    End If
                Else
                    blnConnect = False '正常错误提示
                End If
            End If
        Else
            '错误号12543 TNS: 无法连接目标主机,1012-没有登录，0028-会话被终止
            If mlngErrNum = -2147467259 Or mlngErrNum = -2147217900 Or mlngErrNum = 0 Or mlngErrNum = 12543 Or mlngErrNum = 2399 Or mlngErrNum = 2396 Or mlngErrNum = 1012 Or mlngErrNum = 28 Then
                '检查ORACLE具体错误信息
                If CheckErrConnectInfo(mlngErrNum, strNote, mstrErrInfo, 2) Then
                    '判断相同错误,如果2次以上正常错误提示。
                    If mstrErrRecentSQL = mstrRecentSQL And mstrRecentSQL <> "" Then
                        mlngConnectCount = mlngConnectCount + 1
                        If mlngConnectCount > 2 Then
                            blnNumConnect = False  '正常错误提示
                            mlngConnectCount = 0 '还原计数器
                        Else
                            blnNumConnect = True
                        End If
                    Else
                        mstrErrRecentSQL = mstrRecentSQL
                        mlngConnectCount = 1
                        blnNumConnect = True
                    End If
                Else
                    blnConnect = False '正常错误提示
                End If
            End If
        End If
        
        '自动重新连接一次,检查是否能自动重新连接
        If blnNumConnect Then '与ORACLE连接已经断开
            blnExeSQLTest = gblnExeSQLTest
            gblnExeSQLTest = True
            If CheckAdoConnction(blnStatus) Then
                If blnStatus Then
                   blnConnect = False '正常错误提示
                Else
                   blnConnect = True '提示重连
                End If
            Else
                '与ORACLE重新连接成功,不需要提示。直接返回重新执行。
                blnConnect = False
                ErrCenter = 1
                gblnExeSQLTest = blnExeSQLTest
                Exit Function
            End If
            gblnExeSQLTest = blnExeSQLTest
        End If
    End If
    
    If bytReturnType = 1 Then
        ErrCenter = frmErrAsk.ShowEdit(mlngErrNum, strNote, mstrErrInfo, blnConnect)
    Else
        Call frmErrNote.ShowEdit(mlngErrNum, strNote, mstrErrInfo, blnConnect)
        ErrCenter = 0
    End If
'
    '清除错误
    Err.Clear
End Function

Public Sub SaveErrLog()
'功能：将刚才的错误信息写入数据库错误日志
    Dim strSQL As String
    
    If mlngErrNum <> 0 And mbytErrType <> 0 And gblnErrLog Then
        On Local Error Resume Next
        If gstrComputerName = "" Then Exit Sub
        strSQL = "Zl_Zlerrorlog_Insert('" & gstrComputerName & "'," & mbytErrType & "," & mlngErrNum & "," & AdjustStr(mstrErrInfo) & ")"
        Call ExecuteProcedure(strSQL, "保存错误日志")
        mlngErrNum = 0: mstrErrInfo = "": mbytErrType = 0
    End If
End Sub

Public Function ComputerName() As String
    '功能:获取计算机名
    Dim strComputerName As String * 256
    Err = 0
    On Error Resume Next
    
    Call GetComputerName(strComputerName, 255)
    ComputerName = Trim(Replace(strComputerName, Chr(0), ""))
End Function

Public Sub ShowPercent(sngPercent As Single, objPanel As Object)
'功能:在状态条上根据百分比显示当前处理进度(█)
    Dim intAll As Integer
    intAll = objPanel.Width / frmAbout.TextWidth("█") - 4
    objPanel.Text = Format(sngPercent, "0% ") & String(intAll * sngPercent, "█")
End Sub

Public Sub SelAll(objTxt As Control)
'功能：对文本框的的文本选中
    If TypeName(objTxt) = "TextBox" Then
        objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
    ElseIf TypeName(objTxt) = "MaskEdBox" Then
        If Not IsDate(objTxt.Text) Then
            objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
        Else
            objTxt.SelStart = 0: objTxt.SelLength = 10
        End If
    End If
End Sub

Public Function CheckLen(txt As Object, intLen As Integer, strInfo As String) As Boolean
'功能：检查工本框的真实长度是否在指定限制长度内
    If LenB(StrConv(txt.Text, vbFromUnicode)) > intLen Then
        MsgBox "[" & strInfo & "]的长度不能大于（" & intLen & "）字符或（" & intLen \ 2 & "）汉字！", vbInformation, App.Title
        txt.SetFocus: Exit Function
    End If
    CheckLen = True
End Function

Public Function TLen(Str As String) As Long
'功能：返回字符串的真实长度
    TLen = LenB(StrConv(Str, vbFromUnicode))
End Function

Public Function CheckExist(strTable As String, strField As String, strValue As String, Optional lngID As Long) As Boolean
'功能：检查表strTable中字段strField的值strValue是否重复.
'说明：主要是zlReports及zlRPTGroups和编号查找
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select " & strField & " From " & strTable & " Where " & strField & "=[1] and ID<>[2]"
    Set rsTmp = OpenSQLRecord(strSQL, "CheckExist", UCase(strValue), lngID)
    If rsTmp.RecordCount > 0 Then CheckExist = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetNextID(strTable As String) As Long
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select " & Trim(strTable) & "_ID.Nextval as ID From Dual"
    Call OpenRecord(rsTmp, strSQL, "mdlPublic_GetNextID") '动态SQL
    GetNextID = rsTmp!id
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetCurrID(strTable As String) As Long
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select " & Trim(strTable) & "_ID.CurrVal as ID From Dual"
    Call OpenRecord(rsTmp, strSQL, "mdlPublic_GetCurrID") '动态SQL
    GetCurrID = rsTmp!id
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetSysNO() As String
'功能：返回当前系统所有者对应系统编号
'说明：同一所有者中可能存在多个系统(编号)
'返回：成功:"1,2,3",失败="0"
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    
    On Error GoTo errH
    
    GetSysNO = "0"
    strSQL = "Select 编号 From zlSystems Where 所有者=User"
    Call OpenRecord(rsTmp, strSQL, "mdlPublic_GetSysNO")
    If rsTmp.RecordCount > 0 Then
        GetSysNO = ""
        For i = 1 To rsTmp.RecordCount
            GetSysNO = GetSysNO & "," & rsTmp!编号
            rsTmp.MoveNext
        Next
        GetSysNO = Mid(GetSysNO, 2)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetMenuPath(ByVal lngRPTID As Long, Optional ByVal blnGroup As Boolean) As String
'功能：返回指定报表(组)发布的位置(导航台菜单或模块)
'说明：一个报表可能发布到多个位置
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim strPath1 As String, strPath2 As String
    
    On Error GoTo errH
        
    If blnGroup Then
        strSQL = "Select 1 as 标志,D.标题 as 位置" & _
            " From zlRPTGroups A,zlPrograms B,zlMenus C,zlMenus D" & _
            " Where Nvl(A.系统,0)=Nvl(B.系统,0) And A.程序ID=B.序号" & _
            " And Nvl(B.系统,0)=Nvl(C.系统,0) And B.序号=C.模块" & _
            " And C.组别='缺省' And Upper(B.部件)=Upper('zl9Report')" & _
            " And C.上级ID=D.ID And A.ID=[1]"
    Else
        strSQL = "Select 1 as 标志,D.标题 as 位置" & _
            " From zlReports A,zlPrograms B,zlMenus C,zlMenus D" & _
            " Where Nvl(A.系统,0)=Nvl(B.系统,0) And A.程序ID=B.序号" & _
            " And Nvl(B.系统,0)=Nvl(C.系统,0) And B.序号=C.模块" & _
            " And C.组别='缺省' And Upper(B.部件)=Upper('zl9Report')" & _
            " And C.上级ID=D.ID And A.ID=[1]"
        strSQL = strSQL & " Union ALL " & _
            " Select 2 as 标志,B.标题 as 位置" & _
            " From zlReports A,zlPrograms B" & _
            " Where Nvl(A.系统,0)=Nvl(B.系统,0) And A.程序ID=B.序号" & _
            " And Upper(B.部件)<>Upper('zl9Report') And A.ID=[1]"
        strSQL = strSQL & " Union ALL " & _
            " Select 2 as 标志,B.标题 as 位置" & _
            " From zlRPTPuts A,zlPrograms B" & _
            " Where A.系统=B.系统 And A.程序ID=B.序号" & _
            " And Upper(B.部件)<>Upper('zl9Report') And A.报表ID=[1]"
    End If
    Set rsTmp = OpenSQLRecord(strSQL, "GetMenuPath", lngRPTID)
    For i = 1 To rsTmp.RecordCount
        If rsTmp!标志 = 1 Then
            strPath1 = strPath1 & "," & rsTmp!位置
        ElseIf rsTmp!标志 = 2 Then
            strPath2 = strPath2 & "," & rsTmp!位置
        End If
        rsTmp.MoveNext
    Next
    If strPath1 <> "" Then strPath1 = "导航台(" & Mid(strPath1, 2) & ")"
    If strPath2 <> "" Then strPath2 = "模块(" & Mid(strPath2, 2) & ")"
    If strPath1 <> "" And strPath2 <> "" Then
        GetMenuPath = strPath1 & "," & strPath2
    Else
        GetMenuPath = IIF(strPath1 <> "", strPath1, strPath2)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ReadReport(ByVal lngRPTID As Long, Optional ByRef intMaxID As Integer, Optional ByVal blnOnlyData As Boolean) As Report
'功能：从数据库中读取指定报表到报表对象
'参数：lngRPTID=报表ID,intMaxID=设计界面处理的最大控件索引,读取过程中改变
'      blnOnlyData=只读取报表数据源
'返回：intMaxID=当前可用的最大控件索引,ReadReport=报表对象
    Dim rsReport As New ADODB.Recordset
    Dim rsFormat As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim rsPar As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim rsItem As New ADODB.Recordset
    Dim rsSub As New ADODB.Recordset
    Dim rsGraph As ADODB.Recordset
    Dim rsRelation As ADODB.Recordset
    Dim rsColProtertys As ADODB.Recordset
    Dim lng父ID As Long
    
    Dim strSQL As String, i As Integer, j As Integer
    Dim intCopyID As Integer, strReport As String
    
    Dim tmpReport As Report, tmpData As RPTData, tmpPar As RPTPar
    Dim tmpItem As RPTItem, tmpRelation As RPTRelation
    
    If gstrFonts = "" Then gstrFonts = GetScreenFonts
    
    On Error GoTo errH
        
    If ReportReaded(lngRPTID) Then
        Set rsReport = grsReport '利用缓存
    Else
        strSQL = "Select ID,编号,名称,说明,密码,打印机,进纸,票据,系统,程序ID,功能,修改时间,发布时间,打印方式,禁止开始时间,禁止结束时间 " & vbCr & _
                 "From zlReports Where ID=[1]"
        Set rsReport = OpenSQLRecord(strSQL, "ReadReport", lngRPTID)
        If Not rsReport.EOF Then '缓存处理
            Set grsReport = New ADODB.Recordset
            Set grsReport = rsReport
            gdatModiTime = grsReport!修改时间
        End If
    End If
    If Not rsReport.EOF Then
        strReport = GetFieldNames(rsReport)
        
        Set tmpReport = New Report
        tmpReport.系统 = Nvl(rsReport!系统, 0)
        tmpReport.编号 = rsReport!编号
        tmpReport.名称 = rsReport!名称
        tmpReport.说明 = Nvl(rsReport!说明)
        tmpReport.进纸 = Nvl(rsReport!进纸, 15) '缺省为自动选择
        tmpReport.打印机 = Nvl(rsReport!打印机)
        tmpReport.票据 = Nvl(rsReport!票据, 0) = 1
        tmpReport.打印方式 = Nvl(rsReport!打印方式, 0)
        tmpReport.修改时间 = rsReport!修改时间
        tmpReport.禁止开始时间 = Nvl(rsReport!禁止开始时间, 0)
        tmpReport.禁止结束时间 = Nvl(rsReport!禁止结束时间, 0)
        
        '数据源
        strSQL = "Select ID,数据连接编号,报表ID,名称,字段,对象,类型,说明 From zlRPTDatas Where 报表ID=[1] Order by 名称"
        Set rsData = OpenSQLRecord(strSQL, "ReadReport", lngRPTID)
        If Not rsData.EOF Then
            '数据源SQL
            strSQL = "Select A.源ID,A.行号,A.内容 From zlRPTSQLs A,zlRPTDatas B Where A.源ID=B.ID And B.报表ID=[1] Order by A.源ID,A.行号"
            Set rsSQL = OpenSQLRecord(strSQL, "ReadReport", lngRPTID)
            
            '数据源参数
            strSQL = "Select A.源ID,A.组名,A.序号,A.名称,A.类型,A.缺省值,A.格式,A.值列表,A.分类SQL,A.明细SQL,A.分类字段,A.明细字段,A.对象,A.锁定" & _
                    " From zlRPTPars A,zlRPTDatas B Where A.源ID=B.ID And B.报表ID=[1] Order by A.源ID,A.序号,A.名称,A.类型"
            Set rsPar = OpenSQLRecord(strSQL, "ReadReport", lngRPTID)
        End If
        For i = 1 To rsData.RecordCount
            Set tmpData = New RPTData
            tmpData.数据连接编号 = Nvl(rsData!数据连接编号, 0)
            tmpData.名称 = rsData!名称
            tmpData.类型 = rsData!类型
            tmpData.字段 = rsData!字段
            tmpData.对象 = Nvl(rsData!对象)
            tmpData.说明 = Nvl(rsData!说明)
                        
            'SQL
            tmpData.SQL = ""
            rsSQL.Filter = "源ID=" & rsData!id
            For j = 1 To rsSQL.RecordCount
                tmpData.SQL = tmpData.SQL & vbCrLf & Nvl(rsSQL!内容)
                rsSQL.MoveNext
            Next
            tmpData.SQL = Mid(tmpData.SQL, 3)
            
            '参数
            rsPar.Filter = "源ID=" & rsData!id
            For j = 1 To rsPar.RecordCount
                Set tmpPar = New RPTPar
                tmpPar.组名 = Nvl(rsPar!组名)
                tmpPar.序号 = Nvl(rsPar!序号, 0)
                tmpPar.名称 = Nvl(rsPar!名称)
                tmpPar.类型 = Nvl(rsPar!类型, 0)
                tmpPar.缺省值 = Nvl(rsPar!缺省值)
                tmpPar.格式 = Nvl(rsPar!格式, 0)
                
                tmpPar.值列表 = Nvl(rsPar!值列表)
                tmpPar.分类SQL = Nvl(rsPar!分类SQL)
                tmpPar.明细SQL = Nvl(rsPar!明细SQL)
                tmpPar.分类字段 = Nvl(rsPar!分类字段)
                tmpPar.明细字段 = Nvl(rsPar!明细字段)
                tmpPar.对象 = Nvl(rsPar!对象)
                tmpPar.是否锁定 = IIF(Nvl(rsPar!锁定, 0) = 1, True, False)
                
                '！！！以参数序号为关键字加入集合
                tmpData.Pars.Add tmpPar.组名, tmpPar.序号, tmpPar.名称, tmpPar.类型, tmpPar.缺省值, tmpPar.格式, tmpPar.值列表 _
                    , tmpPar.分类SQL, tmpPar.明细SQL, tmpPar.分类字段, tmpPar.明细字段, tmpPar.对象, "_" & tmpPar.序号, _
                    , tmpPar.是否锁定
                
                rsPar.MoveNext
            Next
            
            '！！！以数据源名称作为关键字加入集合
            tmpReport.Datas.Add tmpData.名称, tmpData.数据连接编号, tmpData.SQL, tmpData.字段, tmpData.对象, tmpData.类型 _
                , tmpData.说明, tmpData.Pars, "_" & tmpData.名称
            
            rsData.MoveNext
        Next
        
        If blnOnlyData = False Then
            '报表格式
            strSQL = "Select 报表ID,序号,说明,W,H,纸张,纸向,动态纸张,图样 From zlRPTFmts Where 报表ID=[1] Order by 序号"
            Set rsFormat = OpenSQLRecord(strSQL, "ReadReport", lngRPTID)
            For i = 1 To rsFormat.RecordCount
                If IsNull(rsFormat!纸张) And IsNull(rsFormat!W) And IsNull(rsFormat!H) _
                    And InStr(strReport, ",纸张,") > 0 And InStr(strReport, ",W,") > 0 Then
                    '兼容考虑：统一为报表统一设置
                    tmpReport.Fmts.Add rsFormat!序号, rsFormat!说明, Nvl(rsReport!W, INIT_WIDTH), Nvl(rsReport!H, INIT_HEIGHT), _
                        Nvl(rsReport!纸张, 9), Nvl(rsReport!纸向, 1), Nvl(rsReport!动态纸张, 0) = 1, Nvl(rsFormat!图样, 0), "_" & rsFormat!序号
                Else
                    '缺省为A4幅面,纵向
                    tmpReport.Fmts.Add rsFormat!序号, rsFormat!说明, Nvl(rsFormat!W, INIT_WIDTH), Nvl(rsFormat!H, INIT_HEIGHT), _
                        Nvl(rsFormat!纸张, 9), Nvl(rsFormat!纸向, 1), Nvl(rsFormat!动态纸张, 0) = 1, Nvl(rsFormat!图样, 0), "_" & rsFormat!序号
                End If
                rsFormat.MoveNext
            Next
            
            '关联报表参数对照
            strSQL = _
                "Select A.元素ID,A.关联报表ID,A.参数名,A.参数值来源,A.默认,b.名称 || '(' || b.编号 || ')' as 关联报表名称 " & vbCr & _
                "From zlrptrelation A ,zlreports B " & vbCr & _
                "Where a.关联报表id=b.id and a.报表ID=[1] "
            Set rsRelation = OpenSQLRecord(strSQL, "ReadReport", lngRPTID)
            
            '列特性设置
            strSQL = _
                "Select A.报表ID,A.元素ID,A.条件名称,A.条件字段,A.条件关系,A.条件值,A.字体颜色,A.背景颜色,A.是否加粗 " & vbCr & _
                "    ,A.是否整行应用,a.对齐 " & vbCr & _
                "From zlRPTColProterty A " & vbCr & _
                "Where a.报表ID=[1]"
            Set rsColProtertys = OpenSQLRecord(strSQL, "ReadReport", lngRPTID)
            
            '报表元素(！切记：表格在前,表格元素在后,按次序号(及XY)排序)
            strSQL = _
                "Select RowNum,系统,ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐" & vbCr & _
                "    ,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,父ID,源行号,源ID,上下间距" & vbCr & _
                "    ,左右间距,横向分栏,纵向分栏,表格线加粗,自适应行高,水平反转 " & vbCr & _
                "From zlRPTItems A " & vbCr & _
                "Where A.报表ID=[1] " & vbCr & _
                "Order by NVL(父ID,0),A.格式号,A.上级ID Desc,A.类型,A.序号,A.X,A.Y"
            Set rsItem = OpenSQLRecord(strSQL, "ReadReport", lngRPTID)
            Set rsSub = rsItem.Clone '复制记录集用于表格子项处理
            
            intMaxID = rsItem.RecordCount '控件最大索引=元素个数(+分栏总数)
            intCopyID = rsItem.RecordCount + 1 '表格分栏控件开始索引
                    
            For i = 1 To rsItem.RecordCount
                Set tmpItem = New RPTItem
                With tmpItem
                    .id = rsItem!Rownum '这里的ID为控件索引(表格子项除外),(Rownum不一定有序,但以RowNum建立的ID及上级ID关系正确)
                    .格式号 = rsItem!格式号
                    .名称 = Nvl(rsItem!名称)
                    .类型 = Nvl(rsItem!类型, 0)
                    .序号 = Nvl(rsItem!序号, 0)
                    .参照 = Nvl(rsItem!参照)
                    .性质 = Nvl(rsItem!性质, 0)
                    .内容 = Nvl(rsItem!内容)
                    .表头 = Nvl(rsItem!表头)
                    .X = Nvl(rsItem!X, 0): .Y = Nvl(rsItem!Y, 0)
                    .W = Nvl(rsItem!W, 0): .H = Nvl(rsItem!H, 0)
                    .水平反转 = Nvl(rsItem!水平反转, 0) = 1
                    If .类型 = 6 And .W < 45 Then .W = 0
                    If .类型 = 2 Or .类型 = 6 Then
                        .行高 = Nvl(rsItem!行高, 0)
                    Else
                        .行高 = Nvl(rsItem!行高, 280)
                    End If
                    .自适应行高 = Nvl(rsItem!自适应行高, 0)
                    .对齐 = Nvl(rsItem!对齐, 0) '缺省左对齐
                    .自调 = Nvl(rsItem!自调, 0) = 1
                    
                    .字体 = Nvl(rsItem!字体, "宋体") '缺省宋体9号
                    If InStr("^" & gstrFonts & "^", "^" & .字体 & "^") = 0 Then .字体 = "宋体"
                    
                    .字号 = Nvl(rsItem!字号, 9)
                    .粗体 = Nvl(rsItem!粗体, 0) = 1
                    .斜体 = Nvl(rsItem!斜体, 0) = 1
                    .下线 = Nvl(rsItem!下线, 0) = 1
                    .网格 = Nvl(rsItem!网格, 0) '缺省黑色
                    .前景 = Nvl(rsItem!前景, 0) '缺省黑色
                    .背景 = Nvl(rsItem!背景, &HFFFFFF) '缺省白色
                    .边框 = Nvl(rsItem!边框, 0) = 1
                    .源行号 = Nvl(rsItem!源行号, 0)
                    .左右间距 = Nvl(rsItem!左右间距, 0)
                    .上下间距 = Nvl(rsItem!上下间距, 0)
                    .横向分栏 = Nvl(rsItem!横向分栏, 0)
                    .纵向分栏 = Nvl(rsItem!纵向分栏, 0)
                    .表格线加粗 = Nvl(rsItem!表格线加粗, 0) = 1
                    If rsItem!源ID & "" <> "" Then
                        rsData.Filter = "ID=" & rsItem!源ID
                        If rsData.RecordCount > 0 Then
                            .数据源 = rsData!名称 & ""
                        End If
                    End If
                     
                    '缺省1栏
                    .分栏 = Nvl(rsItem!分栏, 1)
                    If .类型 <> 6 Then .分栏 = IIF(.分栏 < 1, 1, .分栏)
                    
                    .排序 = Nvl(rsItem!排序)
                    .格式 = Nvl(rsItem!格式)
                    .汇总 = Nvl(rsItem!汇总)
                    .系统 = Nvl(rsItem!系统, 0) = 1
                    
                    '图片的处理
                    If .类型 = 11 Then
                        If gobjFile.FileExists(.内容) Then
                            On Error Resume Next
                            Set .图片 = LoadPicture(.内容) '直接从本地读,加快速度
                            On Error GoTo errH
                        End If
                        If .图片 Is Nothing Then
                            Set rsGraph = New ADODB.Recordset
                            strSQL = "Select 元素ID,图片 From zlRPTGraphs Where 元素ID=[1]"
                            Set rsGraph = OpenSQLRecord(strSQL, "ReadReport", "|查询方式=1-LOB", Val(rsItem!id))
                            If Not rsGraph.EOF Then
                                Set .图片 = GetImage(rsGraph.Fields("图片"))
                            End If
                            rsGraph.Close
                        End If
                    End If
                    
                    '表格子项的处理(类型为6,7,8,9)
                    If InStr(",6,7,8,9,", "," & .类型 & ",") > 0 And Not IsNull(rsItem!上级ID) Then
                        rsSub.Filter = "ID=" & rsItem!上级ID
                        If Not rsSub.EOF Then
                            .上级ID = rsSub!Rownum '这里的上级ID对应表格控件索引
                            tmpReport.Items("_" & .上级ID).SubIDs.Add .id, "_" & .id
                        End If
                    End If
                    
                    '表格分栏索引(自定表头表有效)
                    If .类型 = 4 And .分栏 > 1 Then
                        For j = intCopyID To intCopyID + .分栏 - 2
                            .CopyIDs.Add j, "_" & j
                            intMaxID = intMaxID + 1 '一个分栏加一次
                        Next
                        intCopyID = j
                    End If
                    If rsItem!父ID & "" <> "" Then
                        rsSub.Filter = "ID=" & rsItem!父ID
                        If Not rsSub.EOF Then
                            .父ID = rsSub!Rownum '这里的上级ID对应表格控件索引
                            tmpReport.Items("_" & .父ID).SubIDs.Add .id, "_" & .id
                        End If
                    End If
 
                    '！！！以ID(控件索引)作为关键字加入集合
                    Set tmpItem = tmpReport.Items.Add(.id, .格式号, .名称, .上级ID, .类型, .序号, .参照, .性质, .内容 _
                        , .表头, .X, .Y, .W, .H, .行高, .对齐, .自调, .字体, .字号, .粗体, .下线, .斜体, .网格, .前景 _
                        , .背景, .边框, .分栏, .排序, .格式, .汇总, .表格线加粗, .自适应行高, .图片, .系统, .父ID, .SubIDs _
                        , .CopyIDs, "_" & .id, .数据源, .上下间距, .左右间距, .源行号, .横向分栏, .纵向分栏, , , .水平反转)
                    
                    '加入关联报表
                    rsRelation.Filter = "元素ID=" & rsItem!id
                    If rsRelation.RecordCount > 0 Then rsRelation.MoveFirst
                    For j = 1 To rsRelation.RecordCount
                        Set tmpRelation = New RPTRelation
                        With tmpRelation
                            .关联报表ID = Val(rsRelation!关联报表ID & "")
                            .参数名 = rsRelation!参数名 & ""
                            .参数值来源 = rsRelation!参数值来源 & ""
                            .关联报表名称 = rsRelation!关联报表名称 & ""
                            .默认 = Val(rsRelation!默认 & "")
                            tmpItem.Relations.Add .关联报表ID, .参数名, .参数值来源, .关联报表名称, .默认
                        End With
        
                        rsRelation.MoveNext
                    Next
                    '列特性设置
                    rsColProtertys.Filter = "元素ID=" & rsItem!id
                    If rsColProtertys.RecordCount > 0 Then rsColProtertys.MoveFirst
                    For j = 1 To rsColProtertys.RecordCount
                        tmpItem.ColProtertys.Add rsColProtertys!条件名称, Nvl(rsColProtertys!条件字段) _
                            , Nvl(rsColProtertys!条件关系), Nvl(rsColProtertys!条件值), rsColProtertys!字体颜色 _
                            , rsColProtertys!背景颜色, rsColProtertys!是否加粗, rsColProtertys!是否整行应用 _
                            , Nvl(rsColProtertys!对齐, 0), "_" & rsColProtertys!条件名称
                        rsColProtertys.MoveNext
                    Next
                End With
                rsItem.MoveNext
            Next
            
        End If
        
        Set ReadReport = tmpReport
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Set ReadReport = Nothing
End Function

Public Function SaveReport(lngRPTID As Long, objReport As Report, Optional objPan As Object) As Boolean
'功能:保存报表内容(对象objReport)
    Dim intCount As Integer, i As Integer, strPre As String
    Dim strSQL As String, lngSQLID As Long, lngItemID As Long
    Dim tmpData As RPTData, tmpPar As RPTPar, tmpItem As RPTItem
    Dim tmpID As RelatID, j As Integer
    Dim rsData As ADODB.Recordset
    Dim rsPar As ADODB.Recordset
    Dim rsGraph As ADODB.Recordset
    Dim rsSQL As ADODB.Recordset
    Dim lngParentID As Long
    Dim rsItem As Recordset
    Dim rsRelation As Recordset
    Dim lngTmp As Long
    Dim lngItemSubID As Long

    On Error GoTo errH
    
    If Not objPan Is Nothing Then strPre = objPan.Text
    Screen.MousePointer = 11
    gcnOracle.BeginTrans
    
    With objReport
        '计算进度总数
        If Not objPan Is Nothing Then
            intCount = .Datas.count + .Items.count + .Fmts.count + 1
            For Each tmpData In .Datas
                '数据源SQL
                If Len(Trim(tmpData.SQL)) > 0 Then intCount = intCount + UBound(Split(tmpData.SQL, vbCrLf)) + 1
                intCount = intCount + tmpData.Pars.count
            Next
        End If
        
        '报表主体(打印设置部份)
        gcnOracle.Execute _
            "Update zlReports" & _
            "   Set 打印机='" & .打印机 & "',进纸=" & .进纸 & "," & _
            "       票据=" & IIF(.票据, 1, 0) & ",打印方式=" & .打印方式 & ",修改时间=Sysdate" & ",禁止开始时间=to_date('" & Format(.禁止开始时间, "HH:mm:ss") & "','HH24:MI:SS')" & ",禁止结束时间=to_date('" & Format(.禁止结束时间, "HH:mm:ss") & "','HH24:MI:SS')" & _
            " Where ID=" & lngRPTID
        
        If Not objPan Is Nothing Then
            i = 1: Call ShowPercent(i / intCount, objPan)
        End If
        
        '报表数据源历史记录
        gcnOracle.Execute _
            "Insert Into zlRPTSQLSHistory " & vbNewLine & _
            "(报表id, 数据源名称, 修改人, 修改时间, 行号, 内容)" & vbNewLine & _
            "Select 报表id, 名称, 修改人, 修改时间, 行号, 内容 " & vbNewLine & _
            "From " & vbNewLine & _
            "   (" & vbNewLine & _
            "    Select b.报表id, b.名称, '" & gstrLoginUserName & "' 修改人, Sysdate 修改时间, a.行号, a.内容 " & vbNewLine & _
            "    From zlRPTSQLs A, zlRPTDatas B " & vbNewLine & _
            "    Where a.源id = b.Id And b.报表id = " & lngRPTID & vbNewLine & _
            "    ) A " & vbNewLine & _
            "Where Not Exists(Select 1 From zlRPTSQLSHistory " & vbNewLine & _
            "                 Where 报表id = a.报表id and 数据源名称 = a.名称 " & vbNewLine & _
            "                     And 修改时间 = a.修改时间 and 修改人 = a.修改人 ) "
            
        '报表数据源
        gcnOracle.Execute "Delete From zlRPTDatas Where 报表ID=" & lngRPTID
        
        Set rsData = New ADODB.Recordset
        rsData.CursorLocation = adUseClient
        rsData.Open "Select ID,报表ID,数据连接编号,名称,字段,对象,类型,说明 From zlRPTDatas Where ID=0", gcnOracle, adOpenStatic, adLockOptimistic
        
        For Each tmpData In .Datas
            lngSQLID = GetNextID("zlRPTDatas")
            
            rsData.AddNew
            rsData!id = lngSQLID
            rsData!报表ID = lngRPTID
            If tmpData.数据连接编号 > 0 Then
                rsData!数据连接编号 = tmpData.数据连接编号
            Else
                rsData!数据连接编号 = Null
            End If
            rsData!名称 = tmpData.名称
            rsData!字段 = tmpData.字段
            rsData!对象 = tmpData.对象
            rsData!类型 = tmpData.类型
            rsData!说明 = tmpData.说明
            rsData.Update
            
            '如果修改了名称，则同步修改数据源历史记录的名称
            If tmpData.原名称 <> "" Then
                gcnOracle.Execute "update Zlrptsqlshistory Set 数据源名称='" & tmpData.名称 & "' where 报表ID=" & lngRPTID & " And 数据源名称='" & tmpData.原名称 & "'"
                tmpData.原名称 = ""
            End If
            
            '数据源SQL
            If Len(Trim(tmpData.SQL)) > 0 Then
                Set rsSQL = New ADODB.Recordset
                rsSQL.CursorLocation = adUseClient
                rsSQL.Open "Select 源ID,行号,内容 From zlRPTSQLs Where 源ID=0", gcnOracle, adOpenKeyset, adLockOptimistic
                For j = 0 To UBound(Split(tmpData.SQL, vbCrLf))
                    rsSQL.AddNew
                    rsSQL!源ID = lngSQLID
                    rsSQL!行号 = j + 1
                    rsSQL!内容 = CStr(Split(tmpData.SQL, vbCrLf)(j))
                    rsSQL.Update
                    If Not objPan Is Nothing Then
                        i = i + 1: Call ShowPercent(i / intCount, objPan)
                    End If
                Next
            End If
            
            '数据源参数
            Set rsPar = New ADODB.Recordset
            rsPar.CursorLocation = adUseClient
            rsPar.Open "Select 源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象,锁定 " & _
                       "From zlRPTPars Where 源ID=0", gcnOracle, adOpenStatic, adLockOptimistic
            For Each tmpPar In tmpData.Pars
                rsPar.AddNew
                rsPar!源ID = lngSQLID
                rsPar!组名 = tmpPar.组名
                rsPar!序号 = tmpPar.序号
                rsPar!名称 = tmpPar.名称
                rsPar!类型 = tmpPar.类型
                rsPar!格式 = tmpPar.格式
                rsPar!缺省值 = tmpPar.缺省值
                rsPar!值列表 = tmpPar.值列表
                rsPar!分类SQL = tmpPar.分类SQL
                rsPar!明细SQL = tmpPar.明细SQL
                rsPar!分类字段 = tmpPar.分类字段
                rsPar!明细字段 = tmpPar.明细字段
                rsPar!对象 = tmpPar.对象
                rsPar!锁定 = IIF(tmpPar.是否锁定, 1, 0)
                rsPar.Update
                If Not objPan Is Nothing Then
                    i = i + 1: Call ShowPercent(i / intCount, objPan)
                End If
            Next
            
            If Not objPan Is Nothing Then
                i = i + 1: Call ShowPercent(i / intCount, objPan)
            End If
        Next
    
        '报表格式
        gcnOracle.Execute "Delete From zlRPTFmts Where 报表ID=" & lngRPTID
        For j = 1 To .Fmts.count
            gcnOracle.Execute "Insert Into zlRPTFmts(报表ID,序号,说明,W,H,纸张,纸向,动态纸张,图样) Values(" & _
                lngRPTID & "," & .Fmts(j).序号 & ",'" & .Fmts(j).说明 & "'," & .Fmts(j).W & "," & .Fmts(j).H & "," & _
                .Fmts(j).纸张 & "," & .Fmts(j).纸向 & "," & IIF(.Fmts(j).动态纸张, 1, 0) & "," & .Fmts(j).图样 & ")"
            If Not objPan Is Nothing Then
                i = i + 1: Call ShowPercent(i / intCount, objPan)
            End If
        Next
        
        '报表元素
        gcnOracle.Execute "Delete From zlRPTItems Where 上级ID is Not NULL And 报表ID=" & lngRPTID
        gcnOracle.Execute "Delete From zlRPTItems Where 上级ID is NULL And 报表ID=" & lngRPTID
        gcnOracle.Execute "Delete From zlRPTRelation Where 报表ID=" & lngRPTID
        Set rsItem = New ADODB.Recordset
        rsItem.Fields.Append "ID", adBigInt
        rsItem.Fields.Append "dataid", adBigInt
        rsItem.CursorLocation = adUseClient
        rsItem.LockType = adLockOptimistic
        rsItem.CursorType = adOpenStatic
        rsItem.Open
        
        For Each tmpItem In .Items
            '先保存卡片
            If tmpItem.类型 = Val("14-卡片元素") Then '子项除外
                
                lngItemID = GetNextID("zlRPTItems")
                rsItem.AddNew
                rsItem!id = tmpItem.id
                rsItem!dataid = lngItemID
                rsItem.Update
                lngTmp = 0
                If tmpItem.数据源 <> "" Then
                    rsData.Filter = "名称='" & tmpItem.数据源 & "'"
                    If rsData.RecordCount > 0 Then
                        lngTmp = Val(rsData!id & "")
                    End If
                End If
                gcnOracle.Execute "Insert Into zlRPTItems(ID,报表ID,格式号,名称,上级ID,类型,序号,参照,性质,内容,表头," & _
                    "X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,边框,网格,前景,背景,排序,格式,汇总,分栏,系统,父ID," & _
                    "源ID,源行号,左右间距,上下间距,横向分栏,纵向分栏,表格线加粗,自适应行高,水平反转) " & vbCr & _
                    "Values(" & _
                    lngItemID & "," & lngRPTID & "," & tmpItem.格式号 & ",'" & tmpItem.名称 & "',NULL," & tmpItem.类型 & "," & _
                    tmpItem.序号 & ",'" & tmpItem.参照 & "'," & tmpItem.性质 & ",'" & tmpItem.内容 & "','" & _
                    tmpItem.表头 & "'," & tmpItem.X & "," & tmpItem.Y & "," & tmpItem.W & "," & tmpItem.H & "," & _
                    tmpItem.行高 & "," & tmpItem.对齐 & "," & Abs(CInt(tmpItem.自调)) & ",'" & tmpItem.字体 & "'," & _
                    tmpItem.字号 & "," & Abs(CInt(tmpItem.粗体)) & "," & Abs(CInt(tmpItem.斜体)) & "," & _
                    Abs(CInt(tmpItem.下线)) & "," & Abs(CInt(tmpItem.边框)) & "," & tmpItem.网格 & "," & tmpItem.前景 & "," & _
                    tmpItem.背景 & ",'" & tmpItem.排序 & "','" & tmpItem.格式 & "','" & tmpItem.汇总 & "'," & _
                    IIF(tmpItem.分栏 = 0, 1, tmpItem.分栏) & "," & Abs(CInt(tmpItem.系统)) & "," & "Null" & _
                    "," & IIF(lngTmp = 0, "Null", lngTmp) & "," & tmpItem.源行号 & "," & tmpItem.左右间距 & _
                    "," & tmpItem.上下间距 & "," & tmpItem.横向分栏 & "," & tmpItem.纵向分栏 & _
                    "," & Abs(CInt(tmpItem.表格线加粗)) & "," & IIF(tmpItem.自适应行高, "1", "Null") & _
                    "," & IIF(tmpItem.水平反转, "1", "Null") & ")"
                
                If Not objPan Is Nothing Then
                    i = i + 1: Call ShowPercent(i / intCount, objPan)
                End If

            End If
        Next
        
        '处理其他元素
        For Each tmpItem In .Items
            '保存数据
            If InStr(",1,2,3,4,5,10,11,12,13,", "," & tmpItem.类型 & ",") > 0 Then '子项除外
                lngItemID = GetNextID("zlRPTItems")
                rsItem.AddNew
                rsItem!id = tmpItem.id
                rsItem!dataid = lngItemID
                rsItem.Update
                lngParentID = 0
                If tmpItem.父ID <> 0 Then
                    rsItem.Filter = "ID=" & tmpItem.父ID
                    If rsItem.RecordCount > 0 Then
                        lngParentID = Val(rsItem!dataid & "")
                    End If
                End If
                gcnOracle.Execute "Insert Into zlRPTItems(ID,报表ID,格式号,名称,上级ID,类型,序号,参照,性质,内容,表头," & _
                    "X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,边框,网格,前景,背景,排序,格式,汇总,分栏,系统,父ID," & _
                    "源ID,源行号,左右间距,上下间距,横向分栏,纵向分栏,表格线加粗,自适应行高,水平反转) " & vbCr & _
                    "Values(" & _
                    lngItemID & "," & lngRPTID & "," & tmpItem.格式号 & ",'" & tmpItem.名称 & "',NULL," & tmpItem.类型 & "," & _
                    tmpItem.序号 & ",'" & tmpItem.参照 & "'," & tmpItem.性质 & ",'" & tmpItem.内容 & "','" & _
                    tmpItem.表头 & "'," & tmpItem.X & "," & tmpItem.Y & "," & tmpItem.W & "," & tmpItem.H & "," & _
                    tmpItem.行高 & "," & tmpItem.对齐 & "," & Abs(CInt(tmpItem.自调)) & ",'" & tmpItem.字体 & "'," & _
                    tmpItem.字号 & "," & Abs(CInt(tmpItem.粗体)) & "," & Abs(CInt(tmpItem.斜体)) & "," & _
                    Abs(CInt(tmpItem.下线)) & "," & Abs(CInt(tmpItem.边框)) & "," & tmpItem.网格 & "," & tmpItem.前景 & "," & _
                    tmpItem.背景 & ",'" & tmpItem.排序 & "','" & tmpItem.格式 & "','" & tmpItem.汇总 & "'," & _
                    IIF(tmpItem.分栏 = 0, 1, tmpItem.分栏) & "," & Abs(CInt(tmpItem.系统)) & "," & _
                    IIF(lngParentID = 0, "Null", lngParentID) & "," & "Null" & "," & tmpItem.源行号 & "," & tmpItem.左右间距 & _
                    "," & tmpItem.上下间距 & "," & tmpItem.横向分栏 & "," & tmpItem.纵向分栏 & _
                    "," & Abs(CInt(tmpItem.表格线加粗)) & "," & IIF(tmpItem.自适应行高, "1", "Null") & _
                    "," & IIF(tmpItem.水平反转, "1", "Null") & ")"
                
                '单独处理图片字段
                If Not tmpItem.图片 Is Nothing Then
                    Set rsGraph = New ADODB.Recordset
                    rsGraph.CursorLocation = adUseClient
                    rsGraph.Open "Select 元素ID,图片 From zlRPTGraphs Where 元素ID=" & lngItemID, gcnOracle, adOpenStatic, adLockOptimistic
                    rsGraph.AddNew
                    rsGraph!元素ID = lngItemID
                    Call SaveImage(tmpItem.图片, rsGraph.Fields("图片"))
'                    If isFile(tmpItem.内容) Then
'                        '直接读取文件保存,减少消耗
'                        Call SaveFile(tmpItem.内容, rsGraph.Fields("图片"))
'                    Else
'                        Call SaveImage(tmpItem.图片, rsGraph.Fields("图片"))
'                    End If
                    rsGraph.Update
                End If
                
                If Not objPan Is Nothing Then
                    i = i + 1: Call ShowPercent(i / intCount, objPan)
                End If
                
                '处理表格的子项
                If tmpItem.类型 = Val("4-任意表") Or tmpItem.类型 = Val("5-汇总表") Then
                    For Each tmpID In tmpItem.SubIDs
                        With .Items("_" & tmpID.id)
                            lngItemSubID = GetNextID("zlRPTItems")
                            gcnOracle.Execute "Insert Into zlRPTItems(ID,报表ID,格式号,上级ID,类型,序号,内容,表头,X,Y,W,H," & _
                                "行高,对齐,字体,字号,粗体,斜体,下线,边框,网格,前景,背景,排序,格式,汇总,分栏,系统,自调," & _
                                "父ID,表格线加粗,自适应行高) " & vbCr & _
                                "Values(" & lngItemSubID & "," & lngRPTID & "," & .格式号 & "," & lngItemID & _
                                "," & .类型 & "," & .序号 & ",'" & .内容 & "','" & .表头 & "'," & .X & _
                                "," & .Y & "," & .W & "," & .H & "," & .行高 & "," & .对齐 & ",'" & .字体 & "'," & .字号 & _
                                "," & Abs(CInt(.粗体)) & "," & Abs(CInt(.斜体)) & "," & Abs(CInt(.下线)) & _
                                "," & Abs(CInt(.边框)) & "," & .网格 & "," & .前景 & "," & .背景 & ",'" & .排序 & "'" & _
                                ",'" & .格式 & "','" & .汇总 & "'," & .分栏 & "," & Abs(CInt(.系统)) & "," & Abs(CInt(.自调)) & _
                                "," & IIF(Val(lngParentID) = 0, "NULL", IIF(lngParentID = 0, "Null", lngParentID)) & _
                                "," & Abs(CInt(.表格线加粗)) & "," & IIF(.自适应行高, "1", "Null") & ")"
                            
                            '关联报表参数对照
                            For j = 1 To .Relations.count
                                gcnOracle.Execute _
                                    "Insert Into zlRPTRelation(报表ID,关联报表ID,元素ID,参数名,参数值来源,默认) " & vbCr & _
                                    "Values(" & _
                                    lngRPTID & "," & .Relations.Item(j).关联报表ID & "," & lngItemSubID & _
                                    ",'" & .Relations.Item(j).参数名 & "','" & .Relations.Item(j).参数值来源 & "'" & _
                                    "," & .Relations.Item(j).默认 & ")"
                                If Not objPan Is Nothing Then
                                    i = i + 1: Call ShowPercent(i / intCount, objPan)
                                End If
                            Next
                            '列特性设置
                            For j = 1 To .ColProtertys.count
                                gcnOracle.Execute _
                                    "Insert Into zlRPTColProterty " & vbCr & _
                                    "  (报表ID,元素ID,条件名称,条件字段,条件关系,条件值,字体颜色,背景颜色,是否加粗,是否整行应用,对齐) " & vbCr & _
                                    "Values(" & _
                                    lngRPTID & "," & lngItemSubID & ",'" & .ColProtertys.Item(j).条件名称 & "'" & _
                                    ",'" & .ColProtertys.Item(j).条件字段 & "','" & .ColProtertys.Item(j).条件关系 & "'" & _
                                    ",'" & .ColProtertys.Item(j).条件值 & "'," & Val(.ColProtertys.Item(j).字体颜色) & _
                                    "," & Val(.ColProtertys.Item(j).背景颜色) & "," & IIF(.ColProtertys.Item(j).是否加粗, 1, 0) & _
                                    "," & IIF(.ColProtertys.Item(j).是否整行应用, 1, 0) & _
                                    "," & Val(.ColProtertys.Item(j).对齐) & ")"
                                If Not objPan Is Nothing Then
                                    i = i + 1: Call ShowPercent(i / intCount, objPan)
                                End If
                            Next
                        End With
                        If Not objPan Is Nothing Then
                            i = i + 1: Call ShowPercent(i / intCount, objPan)
                        End If
                    Next
                End If
                '关联报表参数对照
                For j = 1 To tmpItem.Relations.count
                    gcnOracle.Execute _
                        "Insert Into zlRPTRelation(报表ID,关联报表ID,元素ID,参数名,参数值来源,默认) " & vbCr & _
                        "Values(" & _
                        lngRPTID & "," & tmpItem.Relations.Item(j).关联报表ID & "," & lngItemID & _
                        ",'" & tmpItem.Relations.Item(j).参数名 & "'" & _
                        ",'" & tmpItem.Relations.Item(j).参数值来源 & "'" & _
                        "," & tmpItem.Relations.Item(j).默认 & ")"
                    If Not objPan Is Nothing Then
                        i = i + 1: Call ShowPercent(i / intCount, objPan)
                    End If
                Next
            End If
        Next
    End With
    gcnOracle.CommitTrans
    SaveReport = True
    Screen.MousePointer = 0
    
    Set grsReport = Nothing '清除缓存
    
    If Not objPan Is Nothing Then objPan.Text = strPre
    Exit Function
    
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    gcnOracle.RollbackTrans
    Call SaveErrLog
    If Not objPan Is Nothing Then objPan.Text = strPre
End Function

Public Function TrimChar(Str As String) As String
'功能:去除字符串中连续的空格和回车(含两头的空格,回车),不去除TAB字符,哪怕是连续的
    Dim strTmp As String
    Dim i As Long, j As Long
    
    If Trim(Str) = "" Then TrimChar = "": Exit Function
    
    strTmp = Trim(Str)
    
    strTmp = Replace(strTmp, "  ", " ")
    strTmp = Replace(strTmp, "  ", " ")
    
'    i = InStr(strTmp, "  ")
'    Do While i > 0
'        strTmp = Left(strTmp, i) & Mid(strTmp, i + 2)
'        i = InStr(strTmp, "  ")
'    Loop
    
    strTmp = Replace(strTmp, vbCrLf & vbCrLf, vbCrLf)
    strTmp = Replace(strTmp, vbCrLf & vbCrLf, vbCrLf)
    
'    i = InStr(1, strTmp, vbCrLf & vbCrLf)
'    Do While i > 0
'        strTmp = Left(strTmp, i + 1) & Mid(strTmp, i + 4)
'        i = InStr(1, strTmp, vbCrLf & vbCrLf)
'    Loop

    If Left(strTmp, 2) = vbCrLf Then strTmp = Mid(strTmp, 3)
    If Right(strTmp, 2) = vbCrLf Then strTmp = Mid(strTmp, 1, Len(strTmp) - 2)
    TrimChar = strTmp
End Function

Public Sub CopyPars(ByVal objSPars As RPTPars, ByRef objOPars As RPTPars)
'功能：拷贝参数集对象
    Dim tmpPar As RPTPar
    
    Set objOPars = New RPTPars
    For Each tmpPar In objSPars
        With tmpPar
            objOPars.Add .组名, .序号, .名称, .类型, .缺省值, .格式, .值列表, .分类SQL, .明细SQL, .分类字段, .明细字段, .对象 _
                , "_" & .Key, .Reserve, .是否锁定
        End With
    Next
End Sub

Public Function CheckPars(strSQL As String, strMsg As String, objPars As RPTPars) As Boolean
'功能：检查SQL语句中参数符"[]"是否配对,以及参数号是否正确(非数字,不连续)
    Dim intLeft As Integer, intRight As Integer
    Dim intMin As Integer, intMax As Integer
    Dim strTmp As String, StrPar As String, strPars As String
    Dim i As Long, blnSort As Boolean
    Dim objPar As RPTPar
    
    '字符串里的特殊字符转换
    Call mdlPublic.TransSpecialChar(strSQL)
    
    For i = 1 To Len(strSQL)
        If Mid(strSQL, i, 1) = "[" Then intLeft = intLeft + 1
        If Mid(strSQL, i, 1) = "]" Then intRight = intRight + 1
    Next
    If intLeft <> intRight Then
        MsgBox "请确保参数的“[”与“]”符号成对！", vbInformation, App.Title
        Exit Function
    End If
    
    If intLeft = 0 And intRight = 0 Then CheckPars = True: Exit Function
    
    strTmp = strSQL
    intMin = 32767
    Do While InStr(strTmp, "[") > 0
        strTmp = Mid(strTmp, InStr(strTmp, "[") + 1)
        StrPar = Left(strTmp, InStr(strTmp, "]") - 1)
        If Trim(StrPar) = "" Then
            StrPar = 0
        ElseIf Not IsNumeric(StrPar) Then
            Exit Function '非数字编号
        End If
        If CInt(StrPar) < intMin Then intMin = CInt(StrPar)
        If CInt(StrPar) > intMax Then intMax = CInt(StrPar)
        If InStr(strPars, "," & CInt(StrPar)) = 0 Then strPars = strPars & "," & CInt(StrPar)
    Loop
    If intMin <> 0 Then
        strMsg = "参数号定义不是从0开始的,是否自动将后面的参数往前移？"
        If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton1, App.Title) = vbYes Then
            blnSort = True
        Else
            Exit Function '不是从0开始编号
        End If
    End If
    If strPars <> "" Then strPars = Mid(strPars, 2)
    If blnSort = False Then
        If UBound(Split(strPars, ",")) <> intMax Then
            strMsg = "参数号定义不是连续的数字编号，是否自动将后面的参数往前移？"
            If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton1, App.Title) = vbYes Then
                blnSort = True
            Else
                Exit Function '不是连续编号
            End If
        End If
    End If
    
    '自动排序
    If blnSort Then
        For i = 0 To UBound(Split(strPars, ","))
            If Split(strPars, ",")(i) <> i Then
                strSQL = Replace(strSQL, "[" & Split(strPars, ",")(i) & "]", "[" & i & "]")
                If objPars.count > UBound(Split(strPars, ",")) + 1 Then
                    For Each objPar In objPars
                        If objPar.序号 > i Then
                            objPars("_" & Val(objPar.Key) - 1).Key = Val(objPar.Key) - 1
                            objPars("_" & Val(objPar.Key) - 1).Reserve = objPar.Reserve
                            objPars("_" & Val(objPar.Key) - 1).对象 = objPar.对象
                            objPars("_" & Val(objPar.Key) - 1).分类SQL = objPar.分类SQL
                            objPars("_" & Val(objPar.Key) - 1).分类字段 = objPar.分类字段
                            objPars("_" & Val(objPar.Key) - 1).格式 = objPar.格式
                            objPars("_" & Val(objPar.Key) - 1).类型 = objPar.类型
                            objPars("_" & Val(objPar.Key) - 1).名称 = objPar.名称
                            objPars("_" & Val(objPar.Key) - 1).明细SQL = objPar.明细SQL
                            objPars("_" & Val(objPar.Key) - 1).明细字段 = objPar.明细字段
                            objPars("_" & Val(objPar.Key) - 1).缺省值 = objPar.缺省值
                            objPars("_" & Val(objPar.Key) - 1).序号 = objPar.序号 - 1
                            objPars("_" & Val(objPar.Key) - 1).值列表 = objPar.值列表
                        End If
                    Next
                    objPars.Remove "_" & objPars.count - 1
                End If
            End If
        Next
    End If
    
    '字符串里的特殊字符还原
    Call mdlPublic.TransSpecialChar(strSQL, True)
    
    CheckPars = True
End Function

Public Function GetParCount(strSQL As String) As Integer
'功能：返回SQL语句中参数的个数,以序号为准
    Dim strTmp As String, StrPar As String, strPars As String
    
    strTmp = strSQL
    
    '字符串里的特殊字符转换
    Call mdlPublic.TransSpecialChar(strTmp)
    
    Do While InStr(strTmp, "[") > 0
        strTmp = Mid(strTmp, InStr(strTmp, "[") + 1)
        StrPar = Left(strTmp, InStr(strTmp, "]") - 1)
        If Trim(StrPar) = "" Then StrPar = 0
        If InStr(strPars, "," & CInt(StrPar)) = 0 Then strPars = strPars & "," & CInt(StrPar)
    Loop
    If strPars = "" Then
        GetParCount = 0
    Else
        strPars = Mid(strPars, 2)
        GetParCount = UBound(Split(strPars, ",")) + 1
    End If
End Function

Public Function GetCboIndex(cbo As ComboBox, strFind As String) As Long
'功能：由字任串查找ComboBox的索引值
'参数：cbo=ComboBox,strFind=查找字符串
    Dim i As Integer
    If strFind = "" Then GetCboIndex = -1: Exit Function
    For i = 0 To cbo.ListCount - 1
        If cbo.List(i) = strFind Then
            GetCboIndex = i
            Exit Function
        End If
    Next
    GetCboIndex = -1
End Function

Public Function CheckSQL(ByVal strSQL As String, strErr As String, Optional ByVal objPars As RPTPars _
    , Optional ByRef strSQLref As String, Optional ByRef strFieldInfo As String _
    , Optional ByVal objDatas As RPTDatas, Optional ByVal intCurConnect As Integer) As String
'功能：根据缺省参数检查SQL语句书写是否正确
'参数：strFieldInfo=用户返回异常字段，用于提示后的错误位置定位
'      blnCheckInfo=是否检查明细SQL
'      intCurConnect=当前数据连接编号
'返回：
'     成功=SQL的字段串,包含了各个字段的名称及类型,格式如"姓名,111|年龄,111|奖金,123",类型值以ADO.Field.Type为准
'     失败=空
    Dim rsTmp As New ADODB.Recordset, tmpFld As Field
    Dim strCheck As String, strLeft As String, strRight As String
    Dim StrPar As String, bytPar As Byte, i As Integer
    Dim strSQLinfo As String
    
    strCheck = strSQL
    
    '字符串里的特殊字符转换
    Call mdlPublic.TransSpecialChar(strCheck)
    
    If Not objPars Is Nothing Then
        Do While InStr(strCheck, "[") > 0
            strLeft = Left(strCheck, InStr(strCheck, "[") - 1)
            strRight = Mid(strCheck, InStr(strCheck, "]") + 1)
            StrPar = Mid(strCheck, InStr(strCheck, "[") + 1, InStr(strCheck, "]") - InStr(strCheck, "[") - 1)
            If Trim(StrPar) = "" Then StrPar = 0
            bytPar = CByte(StrPar)
            
            '按缺省参数值替换
            If objPars("_" & CInt(bytPar)).缺省值 <> "" And Not objPars("_" & CInt(bytPar)).缺省值 Like "*…" Then
                Select Case objPars("_" & CInt(bytPar)).类型
                    Case 0 '字符
                        StrPar = "'" & Replace(objPars("_" & CInt(bytPar)).缺省值, "'", "''") & "'"
                    Case 1 '数字
                        StrPar = objPars("_" & CInt(bytPar)).缺省值
                    Case 2 '日期
                        If Left(objPars("_" & CInt(bytPar)).缺省值, 1) = "&" Then
                            StrPar = GetParSQLMacro(objPars("_" & CInt(bytPar)).缺省值)
                        Else
                            If InStr(objPars("_" & CInt(bytPar)).缺省值, ":") > 0 Then
                                '长时间格式
                                StrPar = "To_Date('" & Format(objPars("_" & CInt(bytPar)).缺省值, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            Else
                                '短时间格式
                                StrPar = "To_Date('" & Format(objPars("_" & CInt(bytPar)).缺省值, "yyyy-MM-dd") & "','YYYY-MM-DD')"
                            End If
                        End If
                    Case 3 '无类型
                        StrPar = objPars("_" & CInt(bytPar)).缺省值
                End Select
            Else '缺省值为空或为自定义项
                Select Case objPars("_" & CInt(bytPar)).类型
                    Case 0 '字符
                        StrPar = "'空串'"
                    Case 1 '数字
                        StrPar = 0
                    Case 2 '日期
                        StrPar = "Sysdate"
                    Case 3 '无类型(直接替换)
                        If objPars("_" & CInt(bytPar)).缺省值 = "固定值列表…" Then
                            '取固定值中的缺省值
                            '不好的分隔符
                            For i = 0 To UBound(Split(objPars("_" & CInt(bytPar)).值列表, "|"))
                                If Left(Split(objPars("_" & CInt(bytPar)).值列表, "|")(i), 1) = "√" Then
                                    StrPar = Split(Split(objPars("_" & CInt(bytPar)).值列表, "|")(i), ",")(1)
                                    Exit For
                                End If
                            Next
                            '没有设置缺省值则取第一个
                            If StrPar = "" Then
                                StrPar = Split(Split(objPars("_" & CInt(bytPar)).值列表, "|")(0), ",")(1)
                            End If
                        ElseIf objPars("_" & CInt(bytPar)).缺省值 = "选择器定义…" Then
                            If objPars("_" & CInt(bytPar)).值列表 <> "" Then
                                '取缺省绑定值
                                StrPar = Split(objPars("_" & CInt(bytPar)).值列表, "|")(1)
                            ElseIf objPars("_" & CInt(bytPar)).明细SQL <> "" And objPars("_" & CInt(bytPar)).明细字段 <> "" Then
                                strSQLinfo = objPars("_" & CInt(bytPar)).明细SQL
                                Call CheckParsRela(strSQLinfo, objDatas, objPars("_" & CInt(bytPar)).名称, True)
                                StrPar = GetDefaultValue(strSQLinfo, objPars("_" & CInt(bytPar)).明细字段)
                                If StrPar <> "" And InStr(StrPar, "|") > 0 Then StrPar = CStr(Split(StrPar, "|")(1))
                                
                                If objPars("_" & CInt(bytPar)).格式 = 1 Then
                                    StrPar = " In (" & StrPar & ") "
                                End If
                            Else
                                StrPar = ""
                            End If
                        Else
                            StrPar = objPars("_" & CInt(bytPar)).缺省值
                        End If
                End Select
            End If
            strCheck = strLeft & StrPar & strRight
        Loop
    End If
    
    '字符串里的特殊字符还原
    Call mdlPublic.TransSpecialChar(strCheck, True)
    
    strSQLref = strCheck
    If InStr(UCase(strCheck), "WHERE ") > 0 Then
        strCheck = Replace(UCase(strCheck), "WHERE ", "Where Rownum<1 And ")
    End If
    
    Err.Clear
    On Error Resume Next
    Call OpenRecord(rsTmp, strCheck, "mdlPublic_CheckSQL", intCurConnect)  '替换成的都是固定条件,同一数据源一般不变,测试SQL也不会大量运行
    If Err.Number = 0 Then
        strErr = ""
        For Each tmpFld In rsTmp.Fields
            If InStr(tmpFld.name, "|") > 0 Then
                strErr = "字段""" & tmpFld.name & """没有别名！"
                If strFieldInfo = "" Then strFieldInfo = tmpFld.name
                CheckSQL = "": Exit Function
            ElseIf InStr(tmpFld.name, "'") > 0 Or InStr(tmpFld.name, """") > 0 Then
                strErr = "字段名 " & tmpFld.name & " 非法！"
                If strFieldInfo = "" Then strFieldInfo = tmpFld.name
                CheckSQL = "": Exit Function
            Else
                If InStr(CheckSQL & "|", "|" & tmpFld.name & "," & tmpFld.type & "|") = 0 Then
                    CheckSQL = CheckSQL & "|" & tmpFld.name & "," & tmpFld.type
                Else
                    strErr = "在数据源中发现相同的字段项目！"
                    If strFieldInfo = "" Then strFieldInfo = tmpFld.name
                    CheckSQL = "": Exit Function
                End If
            End If
        Next
        CheckSQL = Mid(CheckSQL, 2)
    Else
        strErr = Err.Number & ":" & vbCrLf & Err.Description
        Err.Clear
    End If
    
    Exit Function
    
hErr:
    Call mdlPublic.ErrCenter
End Function

Public Function AdjustStr(Str As String) As String
'功能：将含有"'"符号的字符串调整为Oracle所能识别的字符常量
'说明：自动(必须)在两边加"'"界定符。

    Dim i As Long, strTmp As String
    
    If InStr(1, Str, "'") = 0 Then AdjustStr = "'" & Str & "'": Exit Function
    
    For i = 1 To Len(Str)
        If Mid(Str, i, 1) = "'" Then
            If i = 1 Then
                strTmp = "CHR(39)||'"
            ElseIf i = Len(Str) Then
                strTmp = strTmp & "'||CHR(39)"
            Else
                strTmp = strTmp & "'||CHR(39)||'"
            End If
        Else
            If i = 1 Then
                strTmp = "'" & Mid(Str, i, 1)
            ElseIf i = Len(Str) Then
                strTmp = strTmp & Mid(Str, i, 1) & "'"
            Else
                strTmp = strTmp & Mid(Str, i, 1)
            End If
        End If
    Next
    AdjustStr = strTmp
End Function

Public Function LevelText(ByVal objNode As Object) As String
'功能:返回树形列表中指点定结点的层次名称
    Dim strName As String
    Dim objTmp As Object
    
    strName = objNode.Text
    Set objTmp = objNode
    
    Do While Not objTmp.Parent.Parent Is Nothing
        If objTmp.Parent.Text Like "*（*）" Then
            strName = Split(objTmp.Parent.Text, "（")(0) & "." & strName
        Else
            strName = objTmp.Parent.Text & "." & strName
        End If
        Set objTmp = objTmp.Parent
    Loop
    LevelText = UCase(strName)
End Function

Public Function GetObjRECT(lngHWND As Long) As RECT
'功能:获取对象(窗体或控件)的可见尺寸描述(以象素为单位)
'说明:窗体可结合GetCaptionHeight、GetVscWidth、GetHscHeight函数使用
    Dim Area As RECT
    GetWindowRect lngHWND, Area
    GetObjRECT = Area
End Function

Public Function MakeFile(strID As String, Optional strFormat As String = "CUSTOM") As String
'功能:将资源文件中的指定资源生成磁盘文件
'参数:ID=资源号,strExt=要生成文件的扩展名(如BMP)
'返回:生成文件名
    Dim arrData() As Byte
    Dim intFile As Integer
    Dim strFile As String * 255, strR As String
    
    arrData = LoadResData(strID, strFormat)
    intFile = FreeFile
    GetTempPath 255, strFile
    strR = Trim(Left(strFile, InStr(strFile, Chr(0)) - 1)) & CLng(timer * 100) & ".AVI"
    Open strR For Binary As intFile
    Put intFile, , arrData()
    Close intFile
    MakeFile = strR
End Function

Public Sub ShowFlash(Optional strInfo As String, Optional sngPer As Single = -1, Optional frmParent As Object, Optional blnPer As Boolean)
'功能：显示或隐藏等待或进度窗体(strInfo)
'参数:strInfo=等待或进度提示信息
'     sngPer=进度
    Static blnShow As Boolean
    
    If sngPer > 1 Then sngPer = 1
    
    If strInfo = "" Then
        frmFlash.avi.Close
        Unload frmFlash
        blnShow = False
    Else
        If Not blnShow Then
            On Error Resume Next
            If sngPer = -1 Then
                '显示等待
                frmFlash.avi.Open gstrFind
                frmFlash.lbl.Caption = strInfo
                
                If frmParent Is Nothing Then
                    SetWindowPos frmFlash.hwnd, -1, (Screen.Width - frmFlash.Width) / 2 / 15, (Screen.Height - frmFlash.Height) / 2 / 15, 0, 0, 1
                    ShowWindow frmFlash.hwnd, 5
                Else
                    Err.Clear
                    frmFlash.Show , frmParent
                    If Err.Number <> 0 Then
                        Err.Clear
                        SetWindowPos frmFlash.hwnd, -1, (Screen.Width - frmFlash.Width) / 2 / 15, (Screen.Height - frmFlash.Height) / 2 / 15, 0, 0, 1
                        ShowWindow frmFlash.hwnd, 5
                    End If
                End If
                
                frmFlash.avi.Play
                frmFlash.Refresh
            Else
                '显示进度
                frmFlash.avi.Visible = False
                frmFlash.picDo.Visible = True
                frmFlash.lbl.Top = frmFlash.lbl.Top - frmFlash.lbl.Height / 2
                frmFlash.lbl.Left = frmFlash.picDo.Left
                frmFlash.lblPer.Top = frmFlash.lbl.Top
                frmFlash.lbl.Caption = strInfo
                frmFlash.lblDo.Caption = String(25 * sngPer, frmFlash.lblDo.Tag)
                If blnPer Then
                    If sngPer > 0 Then
                        frmFlash.lblPer.Caption = Int(sngPer * 100) & "%"
                    Else
                        frmFlash.lblPer.Caption = ""
                    End If
                    frmFlash.lblPer.Visible = True
                End If
                
                If frmParent Is Nothing Then
                    SetWindowPos frmFlash.hwnd, -1, (Screen.Width - frmFlash.Width) / 2 / 15, (Screen.Height - frmFlash.Height) / 2 / 15, 0, 0, 1
                    ShowWindow frmFlash.hwnd, 5
                Else
                    Err.Clear
                    frmFlash.Show , frmParent
                    If Err.Number <> 0 Then
                        Err.Clear
                        SetWindowPos frmFlash.hwnd, -1, (Screen.Width - frmFlash.Width) / 2 / 15, (Screen.Height - frmFlash.Height) / 2 / 15, 0, 0, 1
                        ShowWindow frmFlash.hwnd, 5
                    End If
                End If
                
                frmFlash.Refresh
            End If
            blnShow = True
        Else
            frmFlash.lbl.Caption = strInfo
            If sngPer >= 0 Then
                frmFlash.lblDo.Caption = String(25 * sngPer, frmFlash.lblDo.Tag)
                If sngPer > 0 Then
                    frmFlash.lblPer.Caption = Int(sngPer * 100) & "%"
                Else
                    frmFlash.lblPer.Caption = ""
                End If
            End If
            frmFlash.Refresh
        End If
    End If
End Sub

Public Sub SetHeadCenter(msh As Object)
'功能：设置表格固定行居中对齐
    Dim i As Long, j As Long
    Dim blnRedraw As Boolean
    Dim lngRow As Long, lngCol As Long

    blnRedraw = msh.Redraw: lngRow = msh.Row: lngCol = msh.Col: msh.Redraw = False
    For i = 0 To msh.FixedRows - 1
        msh.Row = i
        For j = 0 To msh.Cols - 1
            msh.Col = j
            If i <= msh.FixedRows - 2 And j <= msh.FixedCols - 1 Then '用于清册表头时,后面的条件不满足
                msh.CellAlignment = 7
            Else
                msh.CellAlignment = 4
            End If
        Next
    Next
    msh.Row = lngRow: msh.Col = lngCol: msh.Redraw = blnRedraw
End Sub

Public Function GetParSQLMacro(Str As String) As String
'功能:分析报表参数宏,并返回转换后的在SQL语句中可用的值
    Dim curDate As Date
    
    If InStr(Str, "&") = 0 Then GetParSQLMacro = Str: Exit Function
    
    curDate = Currentdate
    
    Select Case Str
        Case "&当前日期"
            GetParSQLMacro = "TO_DATE('" & Format(curDate, "yyyy-MM-dd") & "','YYYY-MM-DD')"
        Case "&当前日期时间"
            GetParSQLMacro = "Sysdate"
        Case "&当天开始时间"
            GetParSQLMacro = "TO_DATE('" & Format(curDate, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&当天结束时间"
            GetParSQLMacro = "TO_DATE('" & Format(curDate, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&前一天开始时间"
            GetParSQLMacro = "TO_DATE('" & Format(curDate - 1, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&前一天结束时间"
            GetParSQLMacro = "TO_DATE('" & Format(curDate - 1, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&前一天同时间"
            GetParSQLMacro = "Sysdate-1"
        Case "&后一天同时间"
            GetParSQLMacro = "Sysdate+1"
        Case "&后一天结束时间"
            GetParSQLMacro = "TO_DATE('" & Format(curDate + 1, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&后一天日期"
            GetParSQLMacro = "Trunc(Sysdate+1)"
        Case "&前一周日期"
            GetParSQLMacro = "Trunc(Sysdate - 7)"
        Case "&前一月日期"
            GetParSQLMacro = "TO_DATE('" & Format(DateAdd("m", -1, curDate), "yyyy-MM-dd") & "','YYYY-MM-DD')"
        Case "&前一季日期"
            GetParSQLMacro = "TO_DATE('" & Format(DateAdd("m", -3, curDate), "yyyy-MM-dd") & "','YYYY-MM-DD')"
        Case "&前一年日期"
            GetParSQLMacro = "TO_DATE('" & Format(DateAdd("yyyy", -1, curDate), "yyyy-MM-dd") & "','YYYY-MM-DD')"
        Case "&下一周日期"
            GetParSQLMacro = "Trunc(Sysdate + 7)"
        Case "&下一月日期"
            GetParSQLMacro = "TO_DATE('" & Format(DateAdd("m", 1, curDate), "yyyy-MM-dd") & "','YYYY-MM-DD')"
        Case "&下一季日期"
            GetParSQLMacro = "TO_DATE('" & Format(DateAdd("m", 3, curDate), "yyyy-MM-dd") & "','YYYY-MM-DD')"
        Case "&下一年日期"
            GetParSQLMacro = "TO_DATE('" & Format(DateAdd("yyyy", 1, curDate), "yyyy-MM-dd") & "','YYYY-MM-DD')"
        Case "&本月初时间"
            GetParSQLMacro = "TO_DATE('" & Format(Year(curDate) & "-" & Month(curDate) & "-01", "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&本月末时间"
            curDate = DateAdd("m", 1, curDate)
            curDate = CDate(Year(curDate) & "-" & Month(curDate) & "-01") - 1
            GetParSQLMacro = "TO_DATE('" & Format(curDate, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&上月初时间"
            curDate = DateAdd("m", -1, curDate)
            GetParSQLMacro = "TO_DATE('" & Format(Year(curDate) & "-" & Month(curDate) & "-01", "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&上月末时间"
            curDate = CDate(Year(curDate) & "-" & Month(curDate) & "-01") - 1
            GetParSQLMacro = "TO_DATE('" & Format(curDate, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&本年初时间"
            GetParSQLMacro = "TO_DATE('" & Format(Year(curDate) & "-01-01", "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&本年末时间"
            GetParSQLMacro = "TO_DATE('" & Format(Year(curDate) & "-12-31", "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&上年初时间"
            GetParSQLMacro = "TO_DATE('" & Format(Year(curDate) - 1 & "-01-01", "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&上年末时间"
            GetParSQLMacro = "TO_DATE('" & Format(Year(curDate) - 1 & "-12-31", "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
    End Select
End Function

Public Function GetParVBMacro(Str As String) As String
'功能:分析报表参数宏,并返回转换后的VB可用值
    Dim curDate As Date
    
    If InStr(Str, "&") = 0 Then GetParVBMacro = Str: Exit Function
    
    curDate = Currentdate
    Select Case Str
        Case "&当前日期"
            GetParVBMacro = Format(curDate, "yyyy-MM-dd")
        Case "&当前日期时间"
            GetParVBMacro = Format(curDate, "yyyy-MM-dd HH:mm:ss")
        Case "&前一周日期"
            GetParVBMacro = Format(curDate - 7, "yyyy-MM-dd")
        Case "&前一月日期"
            GetParVBMacro = Format(DateAdd("m", -1, curDate), "yyyy-MM-dd")
        Case "&前一季日期"
            GetParVBMacro = Format(DateAdd("m", -3, curDate), "yyyy-MM-dd")
        Case "&前一年日期"
            GetParVBMacro = Format(DateAdd("yyyy", -1, curDate), "yyyy-MM-dd")
        Case "&下一周日期"
            GetParVBMacro = Format(curDate + 7, "yyyy-MM-dd")
        Case "&下一月日期"
            GetParVBMacro = Format(DateAdd("m", 1, curDate), "yyyy-MM-dd")
        Case "&下一季日期"
            GetParVBMacro = Format(DateAdd("m", 3, curDate), "yyyy-MM-dd")
        Case "&下一年日期"
            GetParVBMacro = Format(DateAdd("yyyy", 1, curDate), "yyyy-MM-dd")
        Case "&当天开始时间"
            GetParVBMacro = Format(curDate, "yyyy-MM-dd 00:00:00")
        Case "&当天结束时间"
            GetParVBMacro = Format(curDate, "yyyy-MM-dd 23:59:59")
        Case "&前一天开始时间"
            GetParVBMacro = Format(curDate - 1, "yyyy-MM-dd 00:00:00")
        Case "&前一天结束时间"
            GetParVBMacro = Format(curDate - 1, "yyyy-MM-dd 23:59:59")
        Case "&前一天同时间"
            GetParVBMacro = Format(curDate - 1, "yyyy-MM-dd HH:mm:ss")
        Case "&后一天同时间"
            GetParVBMacro = Format(curDate + 1, "yyyy-MM-dd HH:mm:ss")
        Case "&后一天结束时间"
            GetParVBMacro = Format(curDate + 1, "yyyy-MM-dd 23:59:59")
        Case "&后一天日期"
            GetParVBMacro = Format(curDate + 1, "yyyy-MM-dd")
        Case "&本月初时间"
            GetParVBMacro = Format(Year(curDate) & "-" & Month(curDate) & "-01", "yyyy-MM-dd 00:00:00")
        Case "&本月末时间"
            curDate = DateAdd("m", 1, curDate)
            curDate = CDate(Year(curDate) & "-" & Month(curDate) & "-01") - 1
            GetParVBMacro = Format(curDate, "yyyy-MM-dd 23:59:59")
        Case "&上月初时间"
            curDate = DateAdd("m", -1, curDate)
            GetParVBMacro = Format(Year(curDate) & "-" & Month(curDate) & "-01", "yyyy-MM-dd 00:00:00")
        Case "&上月末时间"
            curDate = CDate(Year(curDate) & "-" & Month(curDate) & "-01") - 1
            GetParVBMacro = Format(curDate, "yyyy-MM-dd 23:59:59")
        Case "&本年初时间"
            GetParVBMacro = Format(Year(curDate) & "-01-01", "yyyy-MM-dd 00:00:00")
        Case "&本年末时间"
            GetParVBMacro = Format(Year(curDate) & "-12-31", "yyyy-MM-dd 23:59:59")
        Case "&上年初时间"
            GetParVBMacro = Format(Year(curDate) - 1 & "-01-01", "yyyy-MM-dd 00:00:00")
        Case "&上年末时间"
            GetParVBMacro = Format(Year(curDate) - 1 & "-12-31", "yyyy-MM-dd 23:59:59")
    End Select
End Function

Public Function GetParUserMacro(Str As String) As String
'功能:分析报表参数宏,并返回转换后的报表输出格式值
    Dim curDate As Date
    
    If InStr(Str, "&") = 0 Then GetParUserMacro = Str: Exit Function
    
    curDate = Currentdate
    Select Case Str
        Case "&当前日期"
            GetParUserMacro = Format(curDate, "yyyy年MM月dd日")
        Case "&当前日期时间"
            GetParUserMacro = Format(curDate, "yyyy年MM月dd日 HH:mm:ss")
        Case "&前一周日期"
            GetParUserMacro = Format(curDate - 7, "yyyy年MM月dd日")
        Case "&前一月日期"
            GetParUserMacro = Format(DateAdd("m", -1, curDate), "yyyy年MM月dd日")
        Case "&前一季日期"
            GetParUserMacro = Format(DateAdd("m", -3, curDate), "yyyy年MM月dd日")
        Case "&前一年日期"
            GetParUserMacro = Format(DateAdd("yyyy", -1, curDate), "yyyy年MM月dd日")
        Case "&下一周日期"
            GetParUserMacro = Format(curDate + 7, "yyyy年MM月dd日")
        Case "&下一月日期"
            GetParUserMacro = Format(DateAdd("m", 1, curDate), "yyyy年MM月dd日")
        Case "&下一季日期"
            GetParUserMacro = Format(DateAdd("m", 3, curDate), "yyyy年MM月dd日")
        Case "&下一年日期"
            GetParUserMacro = Format(DateAdd("yyyy", 1, curDate), "yyyy年MM月dd日")
        Case "&当天开始时间"
            GetParUserMacro = Format(curDate, "yyyy年MM月dd日 00:00:00")
        Case "&当天结束时间"
            GetParUserMacro = Format(curDate, "yyyy年MM月dd日 23:59:59")
        Case "&前一天开始时间"
            GetParUserMacro = Format(curDate - 1, "yyyy年MM月dd日 00:00:00")
        Case "&前一天结束时间"
            GetParUserMacro = Format(curDate - 1, "yyyy年MM月dd日 23:59:59")
        Case "&前一天同时间"
            GetParUserMacro = Format(curDate - 1, "yyyy年MM月dd日 HH:mm:ss")
        Case "&后一天同时间"
            GetParUserMacro = Format(curDate + 1, "yyyy年MM月dd日 HH:mm:ss")
        Case "&后一天结束时间"
            GetParUserMacro = Format(curDate + 1, "yyyy年MM月dd日 23:59:59")
        Case "&后一天日期"
            GetParUserMacro = Format(curDate + 1, "yyyy年MM月dd日")
        Case "&本月初时间"
            GetParUserMacro = Format(Year(curDate) & "-" & Month(curDate) & "-01", "yyyy年MM月dd日")
        Case "&本月末时间"
            curDate = DateAdd("m", 1, curDate)
            curDate = CDate(Year(curDate) & "-" & Month(curDate) & "-01") - 1
            GetParUserMacro = Format(curDate, "yyyy年MM月dd日")
        Case "&上月初时间"
            curDate = DateAdd("m", -1, curDate)
            GetParUserMacro = Format(Year(curDate) & "-" & Month(curDate) & "-01", "yyyy年MM月dd日")
        Case "&上月末时间"
            curDate = CDate(Year(curDate) & "-" & Month(curDate) & "-01") - 1
            GetParUserMacro = Format(curDate, "yyyy年MM月dd日")
        Case "&本年初时间"
            GetParUserMacro = Format(Year(curDate) & "-01-01", "yyyy年MM月dd日")
        Case "&本年末时间"
            GetParUserMacro = Format(Year(curDate) & "-12-31", "yyyy年MM月dd日")
        Case "&上年初时间"
            GetParUserMacro = Format(Year(curDate) - 1 & "-01-01", "yyyy年MM月dd日")
        Case "&上年末时间"
            GetParUserMacro = Format(Year(curDate) - 1 & "-12-31", "yyyy年MM月dd日")
    End Select
End Function

Public Function Currentdate() As Date
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "SELECT SYSDATE FROM DUAL"
    Call OpenRecord(rsTmp, strSQL, "mdlPublic_Currentdate")
    Currentdate = rsTmp.Fields(0).Value
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetDataName(Str As String) As String
    If InStr(Str, "[") = 0 Or InStr(Str, "]") = 0 Then
        GetDataName = Str
    Else
        GetDataName = Mid(Trim(Str), 2, Len(Trim(Str)) - 2)
    End If
End Function

Public Sub PlayWarn()
    Call Beep(2000, 50)
    Call Beep(500, 100)
End Sub

Public Sub CopyTree(tvwS As Control, tvwO As Control, Optional blnCopyBin As Boolean)
'功能：打复制树形列表内容
    Dim objNode As Object, tmpNode As Object
    
    Set tvwO.ImageList = tvwS.ImageList
    tvwO.Nodes.Clear
    
    For Each objNode In tvwS.Nodes
        With objNode
            If .Key = "Root" Then
                Set tmpNode = tvwO.Nodes.Add(, , .Key, .Text, .Image, .SelectedImage)
                tmpNode.Selected = True
                tmpNode.Expanded = .Expanded
            ElseIf .Children = 0 And (IsType(Val(.Tag), adLongVarBinary) And blnCopyBin Or IsType(Val(.Tag), adVarChar) Or IsType(Val(.Tag), adNumeric) Or IsType(Val(.Tag), adDBTimeStamp)) Then
                Set tmpNode = tvwO.Nodes.Add(.Parent.Key, 4, .Key, .Text, .Image, .SelectedImage)
                tmpNode.Expanded = .Expanded
                tmpNode.Tag = .Tag
            ElseIf .Parent.Key = "Root" Then
                Set tmpNode = tvwO.Nodes.Add(.Parent.Key, 4, .Key, .Text, .Image, .SelectedImage)
                tmpNode.Expanded = .Expanded
            End If
        End With
    Next
   
End Sub

Public Function GetItemCount(strFormula As String) As Integer
'功能：返回表列公式中数据项目的个数
    Dim strTmp As String, StrPar As String
    
    strTmp = strFormula
    
    Do While InStr(strTmp, "[") > 0
        strTmp = Mid(strTmp, InStr(strTmp, "[") + 1)
        StrPar = Left(strTmp, InStr(strTmp, "]") - 1)
        If InStr(StrPar, ".") > 0 Then GetItemCount = GetItemCount + 1
    Loop
End Function

Public Function GetNodeType(strNode As String, ByVal tvw As Control) As Long
'功能：由结点路径名返回其类型
'参数：结点名,如"A.B"
    Dim objNode As Object
    
    For Each objNode In tvw.Nodes
        If objNode.Key <> "Root" And objNode.Children = 0 And IsNumeric(objNode.Tag) Then
            If LevelText(objNode) = strNode Then
                GetNodeType = CLng(objNode.Tag)
                Exit Function
            End If
        End If
    Next
End Function

Public Function GetCellRange(msh As Control, Row As Integer, Col As Integer) As Cells
'功能：返回指定单元格的合并范围
'说明：合并的单元格只能在一个方向,且只在固定行范围内,为空的单元格不与任意单元格合并
    Dim intRowB As Integer, intRowE As Integer
    Dim intColB As Integer, intColE As Integer
    Dim i As Integer
    
    '寻找开始行
    If Row < 0 Or Col < 0 Then Exit Function
    If msh.TextMatrix(Row, Col) = "" Then
        GetCellRange.Row1 = Row
        GetCellRange.Row2 = Row
        GetCellRange.Col1 = Col
        GetCellRange.Col2 = Col
        Exit Function
    End If
    
    intRowB = Row
    For i = Row - 1 To 0 Step -1
        If i >= 0 And i <= msh.FixedRows - 1 Then
            If msh.TextMatrix(i, Col) = msh.TextMatrix(i + 1, Col) Then
                intRowB = i
            Else
                Exit For
            End If
        End If
    Next
    '寻找结束行
    intRowE = Row
    For i = Row + 1 To msh.FixedRows - 1
        If i >= 0 And i <= msh.FixedRows - 1 Then
            If msh.TextMatrix(i, Col) = msh.TextMatrix(i - 1, Col) Then
                intRowE = i
            Else
                Exit For
            End If
        End If
    Next
    '寻找开始列
    intColB = Col
    For i = Col - 1 To 0 Step -1
        If i >= 0 And i <= msh.Cols - 1 Then
            If msh.TextMatrix(Row, i) = msh.TextMatrix(Row, i + 1) Then
                intColB = i
            Else
                Exit For
            End If
        End If
    Next
    '寻找结束行
    intColE = Col
    For i = Col + 1 To msh.Cols - 1
        If i >= 0 And i <= msh.Cols - 1 Then
            If msh.TextMatrix(Row, i) = msh.TextMatrix(Row, i - 1) Then
                intColE = i
            Else
                Exit For
            End If
        End If
    Next
    
    GetCellRange.Row1 = intRowB
    GetCellRange.Row2 = intRowE
    GetCellRange.Col1 = intColB
    GetCellRange.Col2 = intColE
End Function

Public Function ReadPicture(objField As Field) As String
'功能：将指定的记录集图形字段复制为图形临时文件
'参数：objField=图形字段对象
'返回：临时产生的图片文件名

    Const BUFFER_SIZE As Integer = 10240
    Dim lngFileSize As Long, lngCurSize As Long, lngModSize As Long
    Dim intBolcks As Integer, intFile As Integer
    Dim strFile As String, strR As String * 255
    Dim arrBuffer() As Byte, j As Integer
    
    On Error GoTo errH
    
    intFile = FreeFile
    
    GetTempPath 255, strR
    strFile = Trim(Left(strR, InStr(strR, Chr(0)) - 1)) & CLng(timer * 100) & ".pic"
    
    Open strFile For Binary As intFile
    
    lngFileSize = objField.ActualSize
    lngModSize = lngFileSize Mod BUFFER_SIZE
    intBolcks = lngFileSize \ BUFFER_SIZE - IIF(lngModSize = 0, 1, 0)
    For j = 0 To intBolcks
        If j = lngFileSize \ BUFFER_SIZE Then
            lngCurSize = lngModSize
        Else
            lngCurSize = BUFFER_SIZE
        End If
        ReDim arrBuffer(lngCurSize - 1) As Byte
        arrBuffer() = objField.GetChunk(lngCurSize)
        Put intFile, , arrBuffer()
    Next
    Close intFile
    ReadPicture = strFile
    Exit Function
errH:
    Close intFile
    Kill strFile
End Function

Public Function GetParValue(frmParent As Object, strName As String) As String
'功能：从当前报表(frmparent.mobjreport)中获取指定参数的值(缺省值或传入值或最近值)
'说明：如果对应数据源在报表中未使用,则指定参数不能返回(空)
    Dim tmpPar As RPTPar, tmpData As RPTData
    
    For Each tmpData In frmParent.mobjReport.Datas
        For Each tmpPar In tmpData.Pars
            If tmpPar.名称 = strName Then
                If tmpPar.Reserve Like "*…|*" Then
                    If Split(tmpPar.Reserve, "|")(1) <> "程序传入" Then
                        GetParValue = Split(tmpPar.Reserve, "|")(1)
                    End If
                End If
                If GetParValue <> "" Then Exit Function
                If tmpPar.类型 = 2 Then
                    If Left(tmpPar.缺省值, 1) = "&" Then
                        GetParValue = GetParUserMacro(tmpPar.缺省值)
                    ElseIf InStr(tmpPar.缺省值, ":") = 0 Then
                        GetParValue = Format(tmpPar.缺省值, "yyyy年MM月dd日")
                    Else
                        GetParValue = Format(tmpPar.缺省值, "yyyy年MM月dd日 HH:mm:ss")
                    End If
                Else
                    If tmpPar.缺省值 Like "*…" Then
                        If tmpPar.值列表 Like "*|*" Then '此时存放了：显示值|绑定值
                            GetParValue = Split(tmpPar.值列表, "|")(0)
                        Else
                            GetParValue = ""
                        End If
                    Else
                        GetParValue = tmpPar.缺省值
                    End If
                End If
                Exit Function
            End If
        Next
    Next
End Function

Public Function GetUserParData(frmParent As Object, intTime As Integer) As String
'功能：获取用户传入的数据,第intTime的个,从0开始。
'说明：如果没有传入,则不能返回(为空)
    Dim i As Integer, j As Integer
    Dim arrPars As Variant
    
    arrPars = frmParent.marrPars
    
    If UBound(arrPars) <> -1 Then
        For i = 0 To UBound(arrPars)
            If InStr(CStr(arrPars(i)), "=") = 0 Then
                If j = intTime Then
                    GetUserParData = CStr(arrPars(i))
                    Exit Function
                End If
                j = j + 1
            End If
        Next
    End If
End Function

Public Function LoadPictureFromPar(frmParent As Object, ByVal strName As String) As StdPicture
'功能：根据图型元素名称从传入参数中读取图片内容
    Dim arrPars As Variant, strFile As String
    Dim i As Integer, j As Integer

    arrPars = frmParent.marrPars
    
    If UBound(arrPars) <> -1 Then
        For i = 0 To UBound(arrPars)
            If UCase(CStr(arrPars(i))) Like UCase(strName) & "=*" Then
                strFile = Mid(CStr(arrPars(i)), InStr(CStr(arrPars(i)), "=") + 1)
                If gobjFile.FileExists(strFile) Then
                    On Local Error Resume Next
                    Set LoadPictureFromPar = LoadPicture(strFile)
                    On Local Error GoTo 0
                    Exit Function
                End If
            End If
        Next
    End If
End Function

Public Function GetChartFileFromPar(frmParent As Object, ByVal strName As String) As String
'功能：从传入参数中检查是否有传入的图表文件
    Dim arrPars As Variant, strFile As String
    Dim i As Integer, j As Integer

    arrPars = frmParent.marrPars
    
    If UBound(arrPars) <> -1 Then
        For i = 0 To UBound(arrPars)
            If UCase(CStr(arrPars(i))) Like UCase(strName) & "=*" Then
                strFile = Mid(CStr(arrPars(i)), InStr(CStr(arrPars(i)), "=") + 1)
                If gobjFile.FileExists(strFile) Then
                    GetChartFileFromPar = strFile
                    Exit Function
                End If
            End If
        Next
    End If
End Function

Public Sub ShowAbout(Optional frmParent As Object)
    Dim frmShow As New frmAbout
    If frmParent Is Nothing Then
        frmShow.Show 1
    Else
        Load frmShow
        Err.Clear
        On Error Resume Next
        frmShow.Show 1, frmParent
        If Err.Number <> 0 Then
            Err.Clear
            frmShow.Show 1
        End If
    End If
End Sub

Public Function ReportLocalSet(ByVal lngSys As Long, ByVal varReport As Variant, ByVal blnOutCall As Boolean, Optional intFormat As Integer, Optional frmParent As Object) As Boolean
'功能：本地打印机设置,不能改变纸张
'参数：blnOutCall=是否外部通过接口在调用
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim frmShow As New frmLocalSet
    
    On Error GoTo errH
    
    If Printers.count = 0 Then MsgBox "在系统中没有检测到任何打印设备,请先安装打印机后再重试该操作！", vbInformation, App.Title: Exit Function
    
    If TypeName(varReport) = "String" Then
        strSQL = "Select ID,编号,名称,说明,密码,打印机,进纸,票据,打印方式,系统,程序ID,功能,修改时间,发布时间,禁止开始时间,禁止结束时间 From zlReports Where 编号=[1] And Nvl(系统,0)=[3]"
    Else
        strSQL = "Select ID,编号,名称,说明,密码,打印机,进纸,票据,打印方式,系统,程序ID,功能,修改时间,发布时间,禁止开始时间,禁止结束时间 From zlReports Where 程序ID=[2] And Nvl(系统,0)=[3]"
    End If
    Set rsTmp = OpenSQLRecord(strSQL, "LocalSet", UCase(varReport), Val(varReport), lngSys)
    If rsTmp.RecordCount = 1 Then
        frmShow.mblnOutCall = blnOutCall
        frmShow.mintFormat = intFormat
        Set frmShow.rsInfo = rsTmp
        If frmParent Is Nothing Then
            frmShow.Show 1
        Else
            Load frmShow
            Err.Clear
            On Error Resume Next
            frmShow.Show 1, frmParent
'            If Err.Number <> 0 Then
'                Err.Clear
'                frmShow.Show 1
'            End If
        End If
        ReportLocalSet = gblnOK
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ShowReport(Optional frmParent As Object, Optional objCurDLL As clsReport, Optional bytStyle As Byte) As Boolean
'功能：根据全局对象gobjReport的内容,打开并显示报表数据
'说明：
'   1.报表打印设置以本地设置为主
'   2.使用该函数之前必须初始gobjReport,garrPars,glngGroup的值
    Dim frmShow As frmReport
    
    Set frmShow = New frmReport
    Call frmShow.ShowMe(frmParent, objCurDLL, garrPars, bytStyle)
    
    If bytStyle <> 0 Then
        Unload frmShow
        Set frmShow = Nothing
    End If
    
    ShowReport = True
End Function

Public Function ShowReportForRec(Optional frmParent As Object, Optional objCurDLL As clsReport, Optional LibDatas As Object, Optional bytStyle As Byte) As Boolean
'功能：根据全局对象gobjReport的内容,打开并显示报表数据
'说明：
'   1.报表打印设置以本地设置为主
'   2.使用该函数之前必须初始gobjReport,garrPars,glngGroup的值
    Dim frmShow As frmReport
    
    Set frmShow = New frmReport
    Call frmShow.PrintReportForRec(frmParent, objCurDLL, LibDatas, garrPars, bytStyle)
    
    If bytStyle <> 0 Then
        Unload frmShow
        Set frmShow = Nothing
    End If
    
    ShowReportForRec = True
End Function

Public Function GetReportFrom(frmParent As Object, objCurDLL As clsReport, ByVal bytStyle As Byte, objfrmShow As Object, LibDatas As Object, objfrmReport As frmReport) As Boolean
'功能：根据全局对象gobjReport的内容,打开并显示报表数据
'说明：
'   1.报表打印设置以本地设置为主
'   2.使用该函数之前必须初始gobjReport,garrPars,glngGroup的值
    Set objfrmReport = New frmReport
    Set objfrmShow = objfrmReport.GetReportForm(frmParent, objCurDLL, LibDatas, garrPars, bytStyle)
    GetReportFrom = True
End Function

Public Function GetAutoFont(ByVal strText As String, ByVal lngW As Long, ByVal lngH As Long _
    , ByVal objFont As StdFont, objBase As Object, Optional ByVal blnWarp As Boolean = True _
    , Optional ByVal sngYDistance As Single) As StdFont
    
'功能：获取在指定大小区域完整输出文字所需的合适字体
'参数：strText=要输出的文字，按自动换行计算，可以包含硬回车
'      lngW,lngH=指定区域大小
'      objFont=原始输出字体
'      objBase=用于计算的临时画布(Form或PictureBox或Printer)
'      blnWarp=是否按自动换行计算
'      sngYDistance=自动换行的多行文字时的行距,缺省为0点(Point)
'返回： 最终输出字体
'说明：本函数执行时间为x/100秒，不宜大量调用。

    Dim lngX As Long, lngY As Long, i As Long
    Dim lngOneH As Long, lngLen As Long
    Dim strChar As String, strNext As String
    Dim sngSize As Currency
    Dim LINE_W As Integer
    
    strText = Replace(strText, vbCrLf, vbCr)
    strText = Replace(strText, vbLf, vbCr)
    If Not blnWarp Then strText = Replace(strText, vbCr, "")
    
    'TextWidth/TextHeight是以文字段整体计算
    'vbCr宽度为0，高度为2行；""宽度为0，高度为一行
    If Trim(Replace(strText, vbCr, "")) = "" Then Call CopyFont(objFont, GetAutoFont): Exit Function
    
    Call CopyFont(objFont, objBase.Font)
    'If objBase.TextWidth(strText) <= lngW And objBase.TextHeight(strText) <= lngH Then
    If objBase.TextWidth("A") * LenB(StrConv(strText, vbFromUnicode)) <= lngW And objBase.TextHeight(strText) <= lngH Then
        Call CopyFont(objFont, GetAutoFont): Exit Function
    End If
    
    '单元格边框固定2个像素
    LINE_W = Screen.TwipsPerPixelX * 2

    sngYDistance = objBase.ScaleY(sngYDistance, vbPoints, vbTwips)
    lngLen = Len(strText)
    lngW = lngW - 2 * LINE_W
    lngH = lngH - 2 * LINE_W
    
    Do While True
        '当前字号模拟输出计算
        lngX = LINE_W: lngY = LINE_W
        
        lngOneH = objBase.TextHeight("字")
        
        For i = 1 To lngLen
            If lngY + lngOneH > lngH Then Exit For
            
            strChar = Mid(strText, i, 1)
            If strChar = vbCr Then
                lngX = LINE_W: lngY = lngY + lngOneH + sngYDistance
            Else
                lngX = lngX + objBase.TextWidth(strChar)
                If i + 1 <= lngLen Then
                    strNext = Mid(strText, i + 1, 1)
                    If lngX + objBase.TextWidth(strNext) - LINE_W > lngW Then
                        If Not blnWarp Then Exit For
                        lngX = LINE_W: lngY = lngY + lngOneH + sngYDistance
                    End If
                End If
            End If
        Next
        
        '当前字号够用
        If i > Len(strText) Then Exit Do
        
        '当前字号过大,更小字号的处理
        sngSize = objBase.Font.Size
        Do While objBase.Font.Size = sngSize And objBase.Font.Size > 1.5
            objBase.Font.Size = objBase.Font.Size - 0.5
        Loop
        If objBase.Font.Size <= 1.5 Then Exit Do
    Loop

    Call CopyFont(objBase.Font, GetAutoFont)
End Function

Public Sub CopyFont(objSource As StdFont, objTarget As StdFont)
    If objTarget Is Nothing Then Set objTarget = New StdFont
    
    objTarget.Charset = objSource.Charset
    objTarget.Weight = objSource.Weight
    objTarget.name = objSource.name
    objTarget.Size = objSource.Size
    objTarget.Bold = objSource.Bold
    objTarget.Italic = objSource.Italic
    objTarget.Underline = objSource.Underline
    objTarget.Strikethrough = objSource.Strikethrough
End Sub

Public Function DrawCell(Dev As Object, ByVal Data As Variant, ByVal X As Long _
    , ByVal Y As Long, ByVal W As Long, ByVal H As Long _
    , Optional ByVal TW As Long, Optional ByVal TH As Long _
    , Optional BorderColor As Long, Optional ForeColor As Long _
    , Optional BackColor As Long = &HFFFFFF, Optional ByVal Font As StdFont _
    , Optional Border As String = "1111", Optional HAlign As Byte _
    , Optional VAlign As Byte = 1, Optional Wrap As Boolean _
    , Optional Ratio As Single = 1, Optional ByVal sngYDistance As Single _
    , Optional ByVal blnBold As Boolean, Optional ByVal bytShape As Byte = 0 _
    , Optional ByVal blnCellWordWrap As Boolean = False) As Boolean
'功能：在指定设备上按指定格式集输出文字或图象
'参数：
'   Dev=输出设备,为Printer或PictureBox对象
'   Data=输出内容,为线条(x)、字符串("xxx")或图象(stdPicture)。字符串不包含vbCrLf,当Data类型为数字型时,表示输出线条
'   TW,TH=输出的限定范围,超过这个范围则自动取消或缩小,为0时无效
'   Border=边框定义,上下左右,"1111"表示全画
'   Align=文字对齐,0=左,1=中,2=右,分水平对齐及垂直对齐
'   Wrap=当输出内容为字符串时,表示是否自动换行，不自动换行时,超宽部份不输出。
'        当输出内容为图片时，表示是否保持图片的宽高比例(不拉伸),同时对齐属性有效
'   Ratio=输出比例,对字体,坐标都有影响,缺省为1(100%)
'   sngYDistance=自动换行的多行文字时的行距,缺省为0点(Single类型是为了缩放计算)
'   blnBold=线条或框线输出时是否加粗
'   bytShape=框线的形状：0-方形，1-圆形
'   blnCellWordWrap=单元格高度自动调整
'说明：
'   1.在使用该函数之前,应该没有改变设备的作图初始值
'   2.输出后定位光标位置在本次输出范围的右上角

    Dim strText As String, arrText() As String
    Dim LINE_W As Integer, blnW As Boolean, blnH As Boolean
    Dim strTemp As String, i As Long
    Dim lngX As Long, lngY As Long
    Dim lngW As Long, lngH As Long
    Dim sngW As Single, sngH As Single
    Dim intOldFillStyle As Integer, intOldDrawLine As Integer
    Dim sngTmp As Single
    Dim intBase As Integer, intCellBorder As Integer
    
    On Error GoTo errH
    
    DrawCell = True
    
    intOldFillStyle = Dev.FillStyle
    intOldDrawLine = Dev.DrawWidth
    
    '范围限定
    If TW > 0 Then
        If X > TW Then Exit Function
        If X + W > TW Then W = TW - X
    End If
    If TH > 0 Then
        If Y > TH Then Exit Function
        If Y + H > TH Then H = TH - Y
    End If
    
    If TypeName(Dev) = "Printer" Then
        '计算屏幕每英寸多少缇与打印机每英寸多少缇（输出精度）的百分比
        sngTmp = Screen.TwipsPerPixelY / Printer.TwipsPerPixelY
        Dev.DrawWidth = Round(IIF(blnBold, 2, 1) * sngTmp, 0)
    Else
        Dev.DrawWidth = IIF(blnBold, 2, 1)
    End If
    
    If TypeName(Data) = "Integer" Then
        X = X * Ratio: Y = Y * Ratio: W = W * Ratio: H = H * Ratio
        If Val(Data) < 0 Then
            '注意
            '  1.必须先设置为0，再为1，物理打印机的空心框才能生效，预览不受影响。
            '  2.老报表中的矩形方框不能空心（包括新增方框元素），新增报表的矩形方框的空心没有问题。
            Dev.FillStyle = vbFSSolid: Dev.FillStyle = vbFSTransparent
            If bytShape = 0 Then
                '空心矩形
                Dev.Line (X, Y)-(X + W - IIF(W > 0, Screen.TwipsPerPixelX * Ratio, 0) _
                    , Y + H - IIF(H > 0, Screen.TwipsPerPixelY * Ratio, 0)), ForeColor, B
            Else
                '空心圆形、椭圆形
                Dev.Circle (X + W / 2, Y + H / 2), IIF(H > W, H, W) / 2, ForeColor, , , H / W
            End If
        Else
            '实心矩形
            Dev.Line (X, Y)-(X + W - IIF(W > 0, Screen.TwipsPerPixelX * Ratio, 0) _
                , Y + H - IIF(H > 0, Screen.TwipsPerPixelY * Ratio, 0)), ForeColor, BF
        End If
    ElseIf TypeName(Data) = "String" Then
        '字体
        If Font Is Nothing Then
            Set Font = New StdFont
            Font.name = "宋体"
            Font.Size = 9
        End If
        
        '不要用Set Dev.Font=Font,对像是byRef
        Dev.Font.name = Font.name
        Dev.Font.Size = Font.Size
        Dev.Font.Bold = Font.Bold
        Dev.Font.Underline = Font.Underline
        Dev.Font.Italic = Font.Italic
        Dev.Font.Strikethrough = Font.Strikethrough
        
        '因缩放后可能字体比例不对,判断时以原始大小为准
        strTemp = Replace(Data, vbCrLf, vbCr)
        strTemp = Replace(strTemp, vbLf, vbCr)
        If H >= Dev.TextHeight(Replace(strTemp, vbCr, "")) Then blnH = True '高度是否够用(加回车的算一行高度)
        
        '缩变
'暂留
'        If TypeName(Dev) = "Printer" Then
'            LINE_W = Dev.TwipsPerPixelX * 2 * Ratio '边线间隔宽度(输出时用,判断时不用)
'            intBase = Dev.TwipsPerPixelX
'        Else
            LINE_W = Screen.TwipsPerPixelX * 2 * Ratio '边线间隔宽度(输出时用,判断时不用)
            intBase = Screen.TwipsPerPixelX
'        End If
        X = -Int(-X * Ratio): Y = -Int(-Y * Ratio)
        W = -Int(-W * Ratio): H = -Int(-H * Ratio)
        Dev.Font.Size = Font.Size * Ratio
        sngYDistance = Dev.ScaleY(sngYDistance * Ratio, vbPoints, vbTwips)
        
        '背景填充
        If Not (BackColor = vbWhite) Then '白色背景暂不处理,以避免重叠覆盖
            Dev.Line (X, Y)-(X + W, Y + H), BackColor, BF
        End If
        
        '宽度是否够用(加回车的为不够用,以便拆行)
        If W > Dev.TextWidth("A") * LenB(StrConv(strTemp, vbFromUnicode)) + (LINE_W * 2) Then
            blnW = InStr(strTemp, vbCr) = 0
        Else
            blnW = False
        End If
        Dev.ForeColor = ForeColor
        
        '输出文字(边框之内再隔一线)
        '超出高度范围则不输出
        If blnH Then
            If blnW Then
                Select Case HAlign
                    Case 0
                        Dev.CurrentX = X + LINE_W * 2
                    Case 1
                        Dev.CurrentX = X + (W - Dev.TextWidth(Data)) / 2
                    Case 2
                        Dev.CurrentX = X + W - LINE_W - Dev.TextWidth(Data)
                End Select
                Select Case VAlign
                    Case 0
                        Dev.CurrentY = Y + LINE_W
                    Case 1
                        Dev.CurrentY = Y + (H - Dev.TextHeight(Data)) / 2 + LINE_W / 2
                    Case 2
                        Dev.CurrentY = Y + H - LINE_W - Dev.TextHeight(Data)
                End Select
                Dev.Print Data
            Else
                '通过0000判断非单元格（标签元素）
                If Border = "0000" Then
                    intCellBorder = 0
                    intBase = 0
                Else
                    intCellBorder = LINE_W * 2
                End If
                
                If Wrap = False And blnCellWordWrap = False Then
                    '不自动拆行时超宽部分不输出
                    strTemp = Replace(strTemp, vbCr, "")
                    strText = ""
                    For i = 1 To Len(Data)
                        If Dev.TextWidth(strText & Mid(strTemp, i, 1)) + intCellBorder + intBase > W Then Exit For
                        strText = strText & Mid(Data, i, 1)
                    Next
                    Select Case HAlign
                        Case 0
                            Dev.CurrentX = X + intCellBorder
                        Case 1
                            Dev.CurrentX = X + (W - Dev.TextWidth(strText)) / 2
                        Case 2
                            Dev.CurrentX = X + W - LINE_W - Dev.TextWidth(strText)
                    End Select
                    Select Case VAlign
                        Case 0
                            Dev.CurrentY = Y + LINE_W
                        Case 1
                            Dev.CurrentY = Y + (H - Dev.TextHeight(strText)) / 2 + LINE_W / 2
                        Case 2
                            Dev.CurrentY = Y + H - LINE_W - Dev.TextHeight(strText)
                    End Select
                    '输出截取部份
                    Dev.Print strText
                Else
                    '拆分文字成多行(在宽高范围内)
                    ReDim arrText(0) '在此,第一行不可能超高
                    For i = 1 To Len(strTemp)
                        If Mid(strTemp, i, 1) = vbCr Then
                            '多行超高则退出,超高部份不输出
                            If (Dev.TextHeight("字") + sngYDistance) * (UBound(arrText) + 2) - sngYDistance > H Then Exit For
                            ReDim Preserve arrText(UBound(arrText) + 1)
                        ElseIf Dev.TextWidth(arrText(UBound(arrText)) & Mid(strTemp, i, 1)) + intCellBorder + intBase > W Then
                            '多行超高则退出,超高部份不输出
                            If (Dev.TextHeight("字") + sngYDistance) * (UBound(arrText) + 2) - sngYDistance > H Then Exit For
                            ReDim Preserve arrText(UBound(arrText) + 1)
                        End If
                        '有可能一行一个字符宽度都不够
                        If Dev.TextWidth(arrText(UBound(arrText)) & Mid(strTemp, i, 1)) - intCellBorder <= W And Mid(strTemp, i, 1) <> vbCr Then
                            arrText(UBound(arrText)) = arrText(UBound(arrText)) & Mid(strTemp, i, 1)
                        End If
                    Next
                    
                    '输出起始坐标
                    Select Case VAlign
                        Case 0
                            Dev.CurrentY = Y + intCellBorder
                        Case 1
                            Dev.CurrentY = Y + (H - (Dev.TextHeight("A") + sngYDistance) * (UBound(arrText) + 1) + sngYDistance) / 2 + LINE_W / 2
                        Case 2
                            Dev.CurrentY = Y + H - LINE_W - (Dev.TextHeight("A") + sngYDistance) * (UBound(arrText) + 1) + sngYDistance
                    End Select
                    
                    '输出各行
                    For i = 0 To UBound(arrText)
                        Select Case HAlign
                            Case 0
                                Dev.CurrentX = X + intCellBorder
                            Case 1
                                Dev.CurrentX = X + (W - Dev.TextWidth(arrText(i))) / 2
                            Case 2
                                Dev.CurrentX = X + W - LINE_W - Dev.TextWidth(arrText(i))
                        End Select
                        If i > 0 Then Dev.CurrentY = Dev.CurrentY + sngYDistance
                        Dev.Print arrText(i)
                    Next
                End If
            End If
        End If
    Else '图形(边框之内)
        LINE_W = 15 * Ratio '边线间隔宽度(输出时用,判断时不用)
        X = X * Ratio: Y = Y * Ratio: W = W * Ratio: H = H * Ratio
        If Not Data Is Nothing Then
            If Not Wrap Then
                If Border = "0000" Then
                    Dev.PaintPicture Data, X, Y, W, H
                Else
                    Dev.PaintPicture Data, X + LINE_W, Y + LINE_W, W - LINE_W * 2, H - LINE_W * 2
                End If
            Else
                lngW = Data.Width * (15 / 26.46) * Ratio
                lngH = Data.Height * (15 / 26.46) * Ratio
                sngW = lngW / W: sngH = lngH / H
                If sngW > sngH Then
                    lngW = lngW / sngW: lngH = lngH / sngW
                Else
                    lngW = lngW / sngH: lngH = lngH / sngH
                End If
                HAlign = 1: VAlign = 1
                Select Case HAlign
                    Case 0
                        lngX = X + LINE_W
                    Case 1
                        lngX = X + LINE_W + (W - LINE_W * 2 - lngW) / 2
                    Case 2
                        lngX = X + LINE_W + (W - LINE_W - lngW)
                End Select
                Select Case VAlign
                    Case 0
                        lngY = Y + LINE_W
                    Case 1
                        lngY = Y + LINE_W + (H - LINE_W * 2 - lngH) / 2
                    Case 2
                        lngY = Y + LINE_W + (H - LINE_W - lngH)
                End Select
                Dev.PaintPicture Data, lngX, lngY, lngW, lngH
            End If
        End If
    End If
    
    If TypeName(Data) <> "Integer" Then
        '最后处理边框
        If Not (BorderColor = vbWhite And TypeName(Data) = "String") Then '白色边框暂不处理,以避免重叠覆盖
            If Mid(Border, 1, 1) Then Dev.Line (X, Y)-(X + W, Y), BorderColor
            If Mid(Border, 2, 1) Then Dev.Line (X, Y + H)-(X + W, Y + H), BorderColor
            If Mid(Border, 3, 1) Then Dev.Line (X, Y)-(X, Y + H), BorderColor
            If Mid(Border, 4, 1) Then Dev.Line (X + W, Y)-(X + W, Y + H), BorderColor
        End If
    End If
    
    Dev.FillStyle = intOldFillStyle
    Dev.DrawWidth = intOldDrawLine
    Exit Function
    
errH:
    DrawCell = False
    Dev.FillStyle = intOldFillStyle
    Dev.DrawWidth = intOldDrawLine
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Public Function ScalePicture(objDraw As Object, objPic As StdPicture, ByVal lngObjW As Long, ByVal lngObjH As Long) As StdPicture
'功能：根据指定图片和目标尺寸，返回合适比例的图片，宽高比例不变
'参数：objPic=要画的图片
'      lngObjW,lngObjH=要画的目标尺寸(Twip)
'      objDraw=用于中转处理的PictureBox(AutoRedraw=True)
    Dim W As Long, H As Long
    Dim lngW As Long, lngH As Long
    Dim sngW As Single, sngH As Single
    
    objDraw.Cls
    objDraw.BackColor = vbWhite
    objDraw.Width = lngObjW
    objDraw.Height = lngObjH
        
    '图片原始大小(Twip)
    W = objPic.Width * (15 / 26.46)
    H = objPic.Height * (15 / 26.46)
    
    sngW = W / objDraw.ScaleWidth
    sngH = H / objDraw.ScaleHeight
    If sngW > sngH Then
        lngW = W / sngW: lngH = H / sngW
    Else
        lngW = W / sngH: lngH = H / sngH
    End If
    
    '作图并返回
    objDraw.PaintPicture objPic, 0, 0, lngW, lngH
    'objDraw.PaintPicture objPic, (objDraw.ScaleWidth - lngW) / 2, (objDraw.ScaleHeight - lngH) / 2, lngW, lngH
    
    Set ScalePicture = objDraw.Image
End Function

Public Function GetFieldValue(frmParent As Object, strSource As String, Optional Convert As Boolean) As String
'功能：从数据源记录集中获取指定字段的原始值
'参数：strSource="分科费用.应收金额",Convert=是否转换为可计算格式,主要针对复合计算时的数字及日期型
'说明：
'   1.从garrData中获取指定记录集当前记录位置的值
'   2.如果字段类型为Long Raw型,则返回产生的临时文件名

    On Error Resume Next
    
    Dim strData As String, strField As String
    Dim rsTmp As ADODB.Recordset, objData As LibData
    Dim rsRaw As ADODB.Recordset
    
    strData = Left(strSource, InStr(strSource, ".") - 1)
    strField = Mid(strSource, InStr(strSource, ".") + 1)
    
    Set objData = frmParent.mLibDatas("_" & strData)
    With objData
        If .DataSet.RecordCount > 0 Then
            If Not IsNull(.DataSet.Fields(strField).Value) Then
                If Err.Number = 3265 Then
                    GetFieldValue = strSource
                    Exit Function
                End If
                If .DataSet.Fields(strField).type = adVarBinary Then    '例如:Dbms_Lob.Substr返回的Raw类型
                        Set rsRaw = New ADODB.Recordset
                      
                        rsRaw.Fields.Append strField, adLongVarBinary, 32767
                        rsRaw.CursorLocation = adUseClient
                        rsRaw.CursorType = adOpenStatic
                        rsRaw.LockType = adLockOptimistic
                        rsRaw.Open
                        
                        rsRaw.AddNew
                        rsRaw.Fields(strField) = .DataSet.Fields(strField)
                        rsRaw.Update
                        
                        GetFieldValue = ReadPicture(rsRaw.Fields(strField))
                
                ElseIf IsType(.DataSet.Fields(strField).type, adLongVarBinary) Then
                    '因为GetChunk方法使位置指针后移,不能重复读取,所以每次克隆
                    Set rsTmp = .DataSet.Clone(adLockReadOnly)
                    rsTmp.Bookmark = .DataSet.Bookmark
                    GetFieldValue = ReadPicture(rsTmp.Fields(strField))
                Else
                    If Convert Then
                        Select Case .DataSet.Fields(strField).type
                            Case adDBTimeStamp, adDBTime, adDBDate, adDate
                                GetFieldValue = "CDate(""" & .DataSet.Fields(strField).Value & """)"
                            Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                                GetFieldValue = """" & .DataSet.Fields(strField).Value & """"
                            Case Else
                                GetFieldValue = .DataSet.Fields(strField).Value
                        End Select
                    Else
                        GetFieldValue = .DataSet.Fields(strField).Value
                    End If
                End If
            Else
                If Convert Then
                    If Not IsType(.DataSet.Fields(strField).type, adLongVarBinary) Then
                        GetFieldValue = "Null"
                    End If
                End If
            End If
        End If
    End With
End Function

Public Sub SetColWidth(msh As Control)
'功能：自动调整表格列宽,以最小适合为准
    Dim arrWidth() As Long
    Dim i As Long, j As Long
    Dim lngBaseW As Long
    
    ReDim arrWidth(msh.Cols - 1)
    
    msh.Redraw = False
    Load frmFlash
    Set frmFlash.Font = msh.Font
    
    For i = 0 To msh.Cols - 1
        If msh.ColWidth(i) <> 0 Then
            lngBaseW = 0
            For j = IIF(msh.FixedRows = 0, 0, msh.FixedRows - 1) To msh.Rows - 1
                If j = msh.FixedRows - 1 And msh.FixedRows > 0 Then
                    lngBaseW = frmFlash.TextWidth(msh.TextMatrix(j, i) & "AB") + 45
                ElseIf msh.TextMatrix(j, i) <> "" Then
                    If frmFlash.TextWidth(msh.TextMatrix(j, i) & "ab") + 45 > arrWidth(i) Then
                        arrWidth(i) = frmFlash.TextWidth(msh.TextMatrix(j, i) & "AB") + 45
                    End If
                End If
            Next
            If arrWidth(i) < lngBaseW Then arrWidth(i) = lngBaseW
        End If
    Next
    
    Unload frmFlash
    
    For i = 0 To msh.Cols - 1
        If msh.ColWidth(i) <> 0 And arrWidth(i) <> 0 Then msh.ColWidth(i) = arrWidth(i)
    Next
    msh.Redraw = True
End Sub

Public Function ResetPrinterPaper(ByVal lngHWND As Long, objReport As Report, ByVal intCopys As Integer) As Boolean
'功能：恢复当前打印机的原始设定纸张
'说明：仅处理打印机的纸张部份属性
    Dim objFmt As RPTFmt
    Dim strTmp As String
    Dim strName As String
    
    Set objFmt = objReport.Fmts("_" & objReport.bytFormat)
    
    If objFmt.纸张 = Val("256-自定义纸张") Then
        If IsWindowsNT Then
            strTmp = GetRegPrinterInfo("PaperForm", objReport.编号, objFmt.说明)
            If Val(strTmp) = 1 Then
                Call SetNTPrinterPaper_Form(lngHWND, objFmt.W / Twip_mm, objFmt.H / Twip_mm, IIF(objFmt.纸向 = 0, 1, objFmt.纸向), intCopys)
            Else
                Call SetNTPrinterPaper(lngHWND, objFmt.W / Twip_mm, objFmt.H / Twip_mm, IIF(objFmt.纸向 = 0, 1, objFmt.纸向), intCopys)
            End If
        Else
            Printer.Width = objFmt.W
            Printer.Height = objFmt.H
        End If
    Else
        Printer.PaperSize = objFmt.纸张
    End If
    ResetPrinterPaper = True
End Function

Public Function SetPrinterPaper(ByVal lngHWND As Long, objReport As Report, ByVal lngH As Long, ByVal intCopys As Integer) As Boolean
'功能：动态设置当前打印机的纸张高度(自定义纸张)
'说明：仅处理打印机的纸张部份属性
    Dim objFmt As RPTFmt
    Dim strTmp As String
    Dim strDefault As String
    
    Set objFmt = objReport.Fmts("_" & objReport.bytFormat)

    SetPrinterPaper = True
    
    If IsWindowsNT Then
        strTmp = GetRegPrinterInfo("PaperForm", objReport.编号, objFmt.说明)
        If Val(strTmp) = 1 Then
            If Not SetNTPrinterPaper_Form(lngHWND, objFmt.W / Twip_mm, lngH / Twip_mm, objFmt.纸向, intCopys) Then
                SetPrinterPaper = False
            End If
        Else
            If Not SetNTPrinterPaper(lngHWND, objFmt.W / Twip_mm, lngH / Twip_mm, objFmt.纸向, intCopys) Then
                SetPrinterPaper = False
            End If
        End If
    Else
        '纸向,打印份数保持不变
        Printer.Width = objFmt.W
        Printer.Height = lngH
    End If
    
    '设置后误差大于100Twip,说明设置失败
    If Abs(Printer.Height - lngH) > 100 Then SetPrinterPaper = False
End Function

Public Function GetRegPrinterInfo(ByVal strKey As String, ByVal strCode As String, _
    ByVal strFormat As String, Optional ByVal objReport As Object, _
    Optional ByVal bytType As Byte = 1, Optional ByVal strUser As String) As String
'功能：获取注册表的打印设置信息
'参数：
'  strKey：注册表的项名
'  strCode：报表编号
'  strFormat：报表格式
'  bytType：99-报表类（Format、AllFormat）；0-格式类（默认）；1-格式表（指定）（Printer、PaperBin等）
'返回：注册表的项值

    Dim strSec As String, strSecUser As String
    Dim strValue As String
    
    strSec = "私有模块\" & App.ProductName & "\LocalSet\" & strCode
    If strUser = "" Then
        strSecUser = "私有模块\" & gstrDBUser & "\" & App.ProductName & "\LocalSet\" & strCode
    Else
        strSecUser = "私有模块\" & strUser & "\" & App.ProductName & "\LocalSet\" & strCode
    End If

    Select Case bytType
    Case 1
        strValue = GetSetting("ZLSOFT", strSec & "\" & strFormat, strKey, "")
        If strValue = "" Then strValue = GetSetting("ZLSOFT", strSec & "\所有格式", strKey, "")
        If strValue = "" Then strValue = GetSetting("ZLSOFT", strSec, strKey, "")
        If strValue = "" Then strValue = GetSetting("ZLSOFT", strSecUser, strKey, "")
        
        If Not objReport Is Nothing Then
            If strValue = "" And strKey = "Printer" Then strValue = GetSetting("ZLSOFT", strSecUser, strKey, objReport.打印机)
        End If
    Case 99
        strValue = GetSetting("ZLSOFT", strSec, strKey, "")
        If strValue = "" Then strValue = GetSetting("ZLSOFT", strSecUser, strKey, "")
    End Select

    GetRegPrinterInfo = strValue
End Function

Public Function InitPrinter(frmParent As Object, Optional ByVal intCopies As Integer = 1) As Boolean
'功能：根据注册表及frmParent.mobjReport内容初始化打印机设置(本地->服务器->当前)
'参数：intCopies=本地设置的要打印的份数
'返回：如果无打印机或纸张不对,则失败
    Dim frmMain As Object
    Dim objReport As Report
    Dim objFmt As RPTFmt
    Dim strPrinter As String
    Dim intPaperBin As Integer
    Dim intOrient As Integer
    Dim i As Integer
    Dim strFormName As String
    Dim strTmp As String
    Dim strDefault As String
    
    If Printers.count = 0 Then Exit Function
    
    If frmParent.frmParent Is Nothing Then
        Set frmMain = frmParent
    Else
        Set frmMain = frmParent.frmParent
    End If
    
    '报表对象
    Set objReport = frmParent.mobjReport
    Set objFmt = objReport.Fmts("_" & objReport.bytFormat)
    
    '本地如果只有一个打印机,默认为它
    If Printers.count = 1 Then
        strPrinter = Printers(0).DeviceName
    Else
        '本地设置
        strPrinter = GetRegPrinterInfo("Printer", objReport.编号, objFmt.说明, objReport)
    End If

    If strPrinter = "" Then
        If MsgBox("""" & objReport.名称 & """没有设置打印机,现在就设置本地打印机吗？", _
            vbQuestion + vbYesNo + vbDefaultButton1, App.Title) = vbNo Then Exit Function
        If Not ReportLocalSet(objReport.系统, objReport.编号, False, objReport.bytFormat, frmMain) Then Exit Function
        strPrinter = GetRegPrinterInfo("Printer", objReport.编号, objFmt.说明, objReport)
    End If
    If Printer.DeviceName <> strPrinter Then
        For i = 0 To Printers.count - 1
            If Printers(i).DeviceName = strPrinter Then Set Printer = Printers(i): Exit For
        Next
        If i > Printers.count - 1 Then
            If MsgBox("""" & objReport.名称 & """的打印机""" & strPrinter & """" & _
                vbCrLf & "在系统中没有安装,要设置本地打印机吗？", _
                vbQuestion + vbYesNo + vbDefaultButton1, App.Title) = vbNo Then Exit Function
            If Not ReportLocalSet(objReport.系统, objReport.编号, False, objReport.bytFormat, frmMain) Then Exit Function
            strPrinter = GetRegPrinterInfo("Printer", objReport.编号, objFmt.说明, objReport)
        End If
    End If
    If Printer.DeviceName <> strPrinter Then
        For i = 0 To Printers.count - 1
            If Printers(i).DeviceName = strPrinter Then Set Printer = Printers(i): Exit For
        Next
    End If
    InitPrinter = True
    
    '1.先按设置固定进行初始化
    On Error Resume Next
    
     '进纸方式
    strTmp = GetRegPrinterInfo("PaperBin", objReport.编号, objFmt.说明)
    intPaperBin = Val(strTmp)
    If intPaperBin = 0 Then intPaperBin = 15
    If Printer.PaperBin <> intPaperBin Then
        Printer.PaperBin = intPaperBin
    End If
    
    '纸张
    If objFmt.纸张 < Val("256-自定义纸张") Then
        Printer.PaperSize = objFmt.纸张
    End If
    
    '份数
    If Printer.Copies <> intCopies Then
        Err.Clear: Printer.Copies = intCopies
        If Err.Number <> 0 Then
            Err.Clear: Printer.Copies = 1
        End If
    End If
    
    '2.NT环境下，用API对设备进行设置
    If objFmt.纸张 = Val("256-自定义纸张") Then
        If IsWindowsNT Then
            strTmp = GetRegPrinterInfo("PaperForm", objReport.编号, objFmt.说明)
            If Val(strTmp) = 1 Then
                strFormName = GetRegPrinterInfo("PaperFormName", objReport.编号, objFmt.说明)
                If Not SetNTPrinterPaper_Form(frmMain.hwnd, objFmt.W / Twip_mm, objFmt.H / Twip_mm, IIF(objFmt.纸向 = 0, 1, objFmt.纸向), intCopies, , strFormName, Printer) Then
                    InitPrinter = False
                End If
            Else
                If Not SetNTPrinterPaper(frmMain.hwnd, objFmt.W / Twip_mm, objFmt.H / Twip_mm, IIF(objFmt.纸向 = 0, 1, objFmt.纸向), intCopies) Then
                    InitPrinter = False
                End If
            End If
        End If
    Else
        '非自定义纸张的纸向
        intOrient = IIF(objFmt.纸向 = 0, 1, objFmt.纸向)
        Printer.Orientation = intOrient
    End If
            
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Function

Private Function GetHscAlign(ByVal intCellAlign As Integer, ByVal strText As String _
    , Optional ByVal objBody As Object _
    , Optional ByVal lngCol As Long = 0) As Byte
'功能：根据表格单元的对齐设置,返回打印水平对齐设置

    Dim intColAlign As Integer

    If Not objBody Is Nothing Then
        intColAlign = objBody.ColAlignment(lngCol)
        If intCellAlign < 0 Or intCellAlign > 8 Then
            '“列特性设置”的对齐未设置（单元格），就使用设计界面的对齐（列）
            intCellAlign = intColAlign
        End If
    End If

    Select Case intCellAlign
    Case 0, 1, 2
        GetHscAlign = 0 '左
    Case 3, 4, 5
        GetHscAlign = 1 '中
    Case 6, 7, 8
        GetHscAlign = 2 '右
    Case Else
        If IsNumeric(strText) Then
            GetHscAlign = 2 '右
        Else
            GetHscAlign = 0 '左
        End If
    End Select
End Function

Private Function GetVscAlign(intAlign As Integer) As Byte
'功能：根据表格单元的对齐设置,返回打印垂直对齐设置
    Select Case intAlign
        Case 0, 3, 6
            GetVscAlign = 0 '上
        Case 1, 4, 7
            GetVscAlign = 1 '中
        Case 2, 5, 8
            GetVscAlign = 2 '下
        Case Else
            GetVscAlign = 1 '中
    End Select
End Function

Private Sub SearchCell(ByVal objGrid As Control, ByVal Row As Long, ByVal Col As Long, ByVal MaxR As Long, ByVal MaxC As Long, _
    W As Long, H As Long, strSkip As String, strSkip2 As String)
'功能：搜索表格的一个单元格的正确宽高
'参数：MaxR,MaxC=搜索的行列最大范围
'返回：W,H=该单元格的宽高(包含合并单元),strSkip=该单元格所合并的单元格,这些单元格不用再处理
    Dim lngW As Long, lngH As Long
    Dim strText As String, i As Long, j As Long
    Dim lngMin As Long, k As Long, blnPreMerge As Boolean
    Dim lngRow As Long, lngCol As Long
    
    objGrid.Row = Row
    objGrid.Col = Col
    lngRow = Row
    lngCol = Col
    strText = objGrid.Text
    lngH = objGrid.RowHeight(Row)
    lngW = objGrid.ColWidth(Col)
    
    '0-flexMergeNever,1-flexMergeFree,2-flexMergeRestrictRows,3-flexMergeRestrictColumns,4-flexMergeRestrictAll
    If strText <> "" And objGrid.MergeCells <> 0 Then
        '向右搜索横向合并单元
        If objGrid.MergeRow(Row) Then
            For i = Col + 1 To MaxC
                objGrid.Col = i
                If strText = objGrid.Text Then
                    If (objGrid.MergeCells = 3 Or objGrid.MergeCells = 4) And objGrid.Row > 0 Then
                        blnPreMerge = True
                        lngMin = IIF(Row >= objGrid.FixedRows, objGrid.FixedRows, 0)
                        For k = Row - 1 To lngMin Step -1
                            If objGrid.TextMatrix(k, i - 1) <> objGrid.TextMatrix(k, i) Then
                                blnPreMerge = False: Exit For
                            End If
                        Next
                        If blnPreMerge Then
                            lngW = lngW + objGrid.ColWidth(i)
                            strSkip = strSkip & "[" & Row & "," & i & "]"
                            strSkip2 = strSkip2 & "[(" & Row & "," & Col & ")," & Row & "," & i & "]"
                            lngCol = i
                        Else
                            Exit For
                        End If
                    Else
                        lngW = lngW + objGrid.ColWidth(i)
                        strSkip = strSkip & "[" & Row & "," & i & "]"
                        strSkip2 = strSkip2 & "[(" & Row & "," & Col & ")," & Row & "," & i & "]"
                        lngCol = i
                    End If
                Else
                    Exit For
                End If
            Next
        End If
        
        '向下搜索纵向合并单元
        objGrid.Col = Col
        If objGrid.MergeCol(Col) Then
            For i = Row + 1 To MaxR
                objGrid.Row = i
                If strText = objGrid.Text Then
                    If (objGrid.MergeCells = 2 Or objGrid.MergeCells = 4) And objGrid.Col > 0 Then
                        blnPreMerge = True
                        lngMin = IIF(Col >= objGrid.FixedCols, objGrid.FixedCols, 0)
                        For k = Col - 1 To lngMin Step -1
                            If objGrid.TextMatrix(i - 1, k) <> objGrid.TextMatrix(i, k) Then
                                blnPreMerge = False: Exit For
                            End If
                        Next
                        If blnPreMerge Then
                            lngH = lngH + objGrid.RowHeight(i)
                            strSkip = strSkip & "[" & i & "," & Col & "]"
                            strSkip2 = strSkip2 & "[(" & Row & "," & Col & ")," & i & "," & Col & "]"
                            lngRow = i
                        Else
                            Exit For
                        End If
                    Else
                        lngH = lngH + objGrid.RowHeight(i)
                        strSkip = strSkip & "[" & i & "," & Col & "]"
                        strSkip2 = strSkip2 & "[(" & Row & "," & Col & ")," & i & "," & Col & "]"
                        lngRow = i
                    End If
                Else
                    Exit For
                End If
            Next
        End If
        objGrid.Row = Row
    End If
    
    '多个单元横纵同时合并
    If lngRow > Row And lngCol > Col Then
        For i = Row + 1 To lngRow
            For j = Col + 1 To lngCol
                If InStr(strSkip, "[" & i & "," & j & "]") = 0 Then
                    strSkip = strSkip & "[" & i & "," & j & "]"
                    strSkip2 = strSkip2 & "[(" & Row & "," & Col & ")," & i & "," & j & "]"
                End If
            Next
        Next
    End If
    
    W = lngW: H = lngH
End Sub

'------------------------------------------------------------------------------------------------
'以下函数用于分析处理数据源权限------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
Public Function UserObject(Optional ByVal intConnect As Integer = 0 _
    , Optional ByVal blnIsBusinessTable As Boolean) As ADODB.Recordset
'功能：获取当前用户所具有Select 权限的所有表及视图名(包含用户自身对象及被授权对象)
'返回：成功=对象名称列表(以中英顺序排序),失败=空
'说明：！！！对中联公共用户对象,本系统将不允许查询
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = _
        "Select USER as OWNER,OBJECT_NAME,Sign(ASCII(OBJECT_NAME)-256) as Sort" & _
        " From User_Objects" & _
        " Where Object_Type in ('TABLE','VIEW') And USER<>'ZLSOFT'" & _
        " Union" & _
        " Select OWNER,OBJECT_NAME,Sign(ASCII(OBJECT_NAME)-256) as Sort" & _
        " From All_Objects O," & _
        " (Select TABLE_NAME From All_Tab_Privs Where Privilege='SELECT') G" & _
        " Where O.Object_Type in('TABLE','VIEW')" & _
        " and O.OBJECT_NAME=G.TABLE_NAME and O.Owner Not in('ZLSOFT')" & _
        "" '" Order by Sort Desc,OBJECT_NAME"
    
    strSQL = _
        "Select USER as OWNER,OBJECT_NAME,Sign(ASCII(OBJECT_NAME)-256) as Sort" & _
        " From User_Objects" & _
        " Where Object_Type in ('TABLE','VIEW')" & _
        " Union" & _
        " Select OWNER,OBJECT_NAME,Sign(ASCII(OBJECT_NAME)-256) as Sort" & _
        " From All_Objects O," & _
        " (Select TABLE_NAME From All_Tab_Privs Where Privilege='SELECT') G" & _
        " Where O.Object_Type in('TABLE','VIEW')" & _
        " and O.OBJECT_NAME=G.TABLE_NAME" & _
        "" '" Order by Sort Desc,OBJECT_NAME"
        
    strSQL = _
        "Select Owner, Object_Name, Sign(Ascii(Object_Name) - 256) As Sort" & vbNewLine & _
        "From (Select User As Owner, Object_Name" & vbNewLine & _
        "       From User_Objects" & vbNewLine & _
        "       Where Object_Type In ('TABLE', 'VIEW')" & vbNewLine & _
        "       Union" & vbNewLine & _
        "       Select Table_Schema, Table_Name" & vbNewLine & _
        "       From All_Tab_Privs" & vbNewLine & _
        "       Where Privilege = 'SELECT' And Table_Name Not Like '%_ID'" & vbNewLine & _
        "       Group By Table_Schema, Table_Name)" & vbNewLine & _
        "" '"Order By Sort Desc, Object_Name"
        
    strSQL = "Select * From (" & vbCrLf & _
             "" & strSQL & vbCrLf & _
             ")" & vbCrLf
             
    If blnIsBusinessTable Then
        strSQL = strSQL & _
                 "Where not Owner in ('SYSTEM', 'SYS', 'DEMO', 'MDSYS', 'ZLTOOLS') " & vbCrLf & _
                 "Order By Sort Desc, Object_Name, Owner"
    Else
        strSQL = strSQL & _
                 "Order By Sort Desc, Object_Name, Owner"
    End If

    On Error GoTo errH
    Call OpenRecord(rsTmp, strSQL, "mdlPublic_UserObject", intConnect)
    Set UserObject = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function TrueObject(ByVal strObject As String) As String
'功能：SQLObject函数的子函数,用于去除对象名中的无用字符
    Dim i As Integer
    '寻找第一个正常字符位置
    For i = 1 To Len(strObject)
        If InStr(Chr(32) & Chr(13) & Chr(10) & Chr(9), Mid(strObject, i, 1)) = 0 Then Exit For
    Next
    strObject = Mid(strObject, i)
    '寻找后面第一个非正常字符
    For i = 1 To Len(strObject)
        If InStr(Chr(32) & Chr(13) & Chr(10) & Chr(9), Mid(strObject, i, 1)) > 0 Then Exit For
    Next
    If i <= Len(strObject) Then strObject = Left(strObject, i - 1)
    TrueObject = strObject
End Function

Private Function GetWithAsTables(ByVal strSQL As String) As String
'功能：获取With as 之间的表名串，以逗号分隔
    Dim lngL As Long, lngR As Long, lngS As Long, strTabs As String
    Dim strTmp As String, blnFirst As Boolean
        
    strSQL = Replace(strSQL, vbCrLf, " ")
    strSQL = Replace(strSQL, vbTab, " ")
    strSQL = Replace(strSQL, "  ", " ")
    strSQL = Replace(strSQL, "  ", " ")
    strSQL = Replace(strSQL, "AS (", "AS(")
    
    lngL = InStr(1, strSQL, "WITH")
    If lngL = 0 Then
        Exit Function
    Else
        lngL = lngL + 4
        blnFirst = True
    End If
        
    Do
        lngR = InStr(lngL, strSQL, " AS(")
        If lngR = 0 Then
            Exit Do
        Else
            If Not blnFirst Then
                lngL = InStrRev(strSQL, ",", lngR) + 1
            End If
            
            strTmp = Trim(Mid(strSQL, lngL, lngR - lngL))
            '11G R2支持，例如：with T（column alias 1,column alias 2,......）
            lngS = InStr(strTmp, "(")
            If lngS > 1 Then
                strTmp = Mid(strTmp, 1, strTmp - 1)
            End If
            
            strTabs = strTabs & "," & strTmp
        End If
        
        blnFirst = False
        lngL = lngR + Len(" AS(")
    Loop
    GetWithAsTables = Mid(strTabs, 2)
End Function

Public Function SQLObject(ByVal strSQL As String, Optional ByVal strWithas As String) As String
'功能：分析SQL语句所用到的对象名
'参数：strSQL=要分析的原始SQL语句
'返回：SQL语句所访问到的对象名,如"部门表,病人费用记录,ZLHIS.人员表"
'说明：1.与Oracle SELECT语句兼容
'      2.如果SQL语句中的对象名前加有所有者前缀,则该前缀不会被截取
'      3.需要函数TrimChar;TrueObject的支持
    Dim intB As Long, intE As Long, intL As Long, intR As Long
    Dim strAnal As String, strSub As String, strObject As String
    Dim arrFrom() As String, strCur As String, strMulti As String, strTrue As String
    Dim i As Long, j As Long
    Dim lngTmp As Long
    Dim strTmp As String, strObjectSub As String
    
    On Error GoTo errH
    
    '大写化及去除多余的字符
    strAnal = UCase(TrimChar(strSQL))
    If strWithas = "" Then
        strWithas = GetWithAsTables(strAnal)
    End If
    
    If InStr(strAnal, "SELECT") = 0 Or InStr(strAnal, "FROM") = 0 Then Exit Function
    If mdlPublic.TransSpecialChar(strAnal) = False Then Exit Function
    
    '先分解处理嵌套子查询
    Do While InStr(strAnal, "(") > 0
        intB = InStr(strAnal, "("): intE = intB '匹配的左右括号位置
        intL = 1: intR = 0
        For i = intB + 1 To Len(strAnal)
            If Mid(strAnal, i, 1) = "(" Then
                intL = intL + 1
            ElseIf Mid(strAnal, i, 1) = ")" Then
                intR = intR + 1
            End If
            If intL = intR Then
                intE = i
                strTmp = Mid(strAnal, 1, intB - 1)
                lngTmp = 0
                If InStrRev(strTmp, " TABLE") > 0 Or InStrRev(strTmp, " TABLE ") > 0 Then
                    lngTmp = IIF(InStrRev(strTmp, " TABLE ") > 0, InStrRev(strTmp, " TABLE "), InStrRev(strTmp, " TABLE"))
                    strTmp = Mid(strTmp, lngTmp + 6)
                    strTmp = Trim(strTmp)
                End If
                If intE - intB - 1 <= 0 Then
                    '对于非子查询,将括号换成其它符号,以使循环继续
                    strAnal = Left(strAnal, intB - 1) & "@" & Mid(strAnal, intB + 1)
                    strAnal = Left(strAnal, intE - 1) & "@" & Mid(strAnal, intE + 1)
                ElseIf InStr(Mid(strAnal, intB + 1, intE - intB - 1), "SELECT") > 0 _
                    And InStr(Mid(strAnal, intB + 1, intE - intB - 1), "FROM") > 0 Then
                    '子查询语句
                    strSub = Mid(strAnal, intB + 1, intE - intB - 1)
                    '将该子查询部份作为为特殊对象名
                    strAnal = Replace(strAnal, Mid(strAnal, intB, intE - intB + 1), "嵌套查询")
                    '递归分析
                    strObjectSub = SQLObject(strSub, strWithas)
                    If InStr(strObject & "," & strWithas & ",", "," & strObjectSub & ",") = 0 Then
                        strObject = strObject & "," & strObjectSub
                    End If
                ElseIf strTmp = "" And lngTmp <> 0 Then
                    '去除Table动态内存表
                    strAnal = Replace(strAnal, Mid(strAnal, lngTmp + 1, intE - lngTmp + 1 + 1), "动态内存表")
                Else
                    strAnal = Left(strAnal, intB - 1) & "@" & Mid(strAnal, intB + 1)
                    strAnal = Left(strAnal, intE - 1) & "@" & Mid(strAnal, intE + 1)
                End If
                Exit For
            End If
        Next
        '无匹配右括号
        If intE = intB Then strAnal = Left(strAnal, intB - 1) & "@" & Mid(strAnal, intB + 1)
    Loop
    
    '分解分析(此时strAnal为简单查询,可能带Union等连接)
    arrFrom = Split(strAnal, "FROM")
    For i = 1 To UBound(arrFrom) '从第一个From后面部份开始
        strCur = arrFrom(i)
        If InStr(strCur, "WHERE") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "WHERE") - 1)
        ElseIf InStr(strCur, "START WITH") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "START WITH") - 1)
        ElseIf InStr(strCur, "CONNECT BY") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "CONNECT BY") - 1)
        ElseIf InStr(strCur, "GROUP") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "GROUP") - 1)
        ElseIf InStr(strCur, "HAVING") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "HAVING") - 1)
        ElseIf InStr(strCur, "ORDER") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "ORDER") - 1)
        ElseIf InStr(strCur, "UNION") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "UNION") - 1)
        ElseIf InStr(strCur, "MINUS") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "MINUS") - 1)
        ElseIf InStr(strCur, "INTERSECT") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "INTERSECT") - 1)
        Else
            strMulti = strCur
        End If
        For j = 0 To UBound(Split(strMulti, ","))
            strTrue = TrueObject(Split(strMulti, ",")(j))
            If InStr(strObject & "," & strWithas & ",", "," & strTrue & ",") = 0 And strTrue <> "嵌套查询" And strTrue <> "动态内存表" Then
                If InStr(strTrue, "'") = 0 And InStr(strTrue, "@") = 0 Then
                    strObject = strObject & "," & strTrue
                End If
            End If
        Next
    Next
    '完成
    SQLObject = Mid(strObject, 2)
    SQLObject = Replace(SQLObject, ",,", ",")
    Exit Function
errH:
    Err.Clear
End Function

Public Function CheckReportPriv(lngRPTID As Long, Optional ByVal blnReportGroup As Boolean) As Boolean
'功能：检查当前用户对某张报表(已存在)是否完全有权限访问
'参数：lngRPTID=报表ID
'返回：完全="",不完全=不能访问的对象名,如"ZLPER.部门表,ZLHIS.病人费用记录"
'说明：用于在报表管理界面打开或设计报表时检查权限
'参考：grsObject
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer, j As Integer
    Dim strOwner As String, strName As String
    Dim strSQL As String
    
    On Error GoTo errH
    
    rsTmp.CursorLocation = adUseClient
    If Not blnReportGroup Then
        strSQL = "Select 名称,对象 From zlRPTDatas Where 报表ID=[1] And Nvl(数据连接编号, 0) = 0 "
    Else
        strSQL = "Select A.名称,A.对象 " & vbCr & _
                 "From zlRPTDatas A, zlRPTSubs B " & vbCr & _
                 "Where A.报表ID=B.报表ID And Nvl(a.数据连接编号, 0) = 0 And B.组ID=[1] "
    End If
    Set rsTmp = OpenSQLRecord(strSQL, "CheckReportPriv", lngRPTID)
    For i = 1 To rsTmp.RecordCount '如果有数据访问
        If Not IsNull(rsTmp!对象) Then
            For j = 0 To UBound(Split(rsTmp!对象, ","))
                strOwner = Split(Split(rsTmp!对象, ",")(j), ".")(0)
                strName = Split(Split(rsTmp!对象, ",")(j), ".")(1)
                grsObject.Filter = "OWNER='" & strOwner & "' AND OBJECT_NAME='" & strName & "'"
                If grsObject.EOF Then Exit Function
            Next
        End If
        rsTmp.MoveNext
    Next
    CheckReportPriv = True
    Exit Function
    
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckObjectPriv(strObject As String) As String
'功能：检查当前用户对指定对象是否完全有权限访问
'参数：strObject=对象名串,如"部门表,病人费用记录"
'返回：完全=空,不完全=不能访问的对象名,如"部门表,病人费用记录"
'说明：用于在校验数据源之前检查是否有权限查询SQL语句中的对象
'参考：grsObject
    Dim i As Integer
    Dim arrObject As Variant
    
    arrObject = Split(strObject, ",")
    For i = 0 To UBound(arrObject)
        If arrObject(i) <> "DUAL" Then
            If InStr(arrObject(i), ".") = 0 Then
                grsObject.Filter = "OBJECT_NAME='" & arrObject(i) & "'"
            Else
                '如果本身就加了所有者前缀,则检查该所有者对象权限
                grsObject.Filter = "OWNER='" & Split(arrObject(i), ".")(0) & _
                    "' And OBJECT_NAME='" & Split(arrObject(i), ".")(1) & "'"
            End If
            If grsObject.EOF Then
                If InStr(CheckObjectPriv & ",", "," & arrObject(i) & ",") = 0 Then
                    CheckObjectPriv = CheckObjectPriv & "," & arrObject(i)
                End If
            End If
        End If
    Next
    If CheckObjectPriv <> "" Then CheckObjectPriv = Mid(CheckObjectPriv, 2)
End Function

Public Function ObjectOwner(ByVal strObject As String, Optional frmParent As Object, _
    Optional ByVal intConnect As Integer = 0) As String
'功能：根据对象名加上当前用户所能访问的所有者前缀(包括对同一对象名有多个所有者要求选其中之一)
'参数：strObject=对象名串,如"部门表,病人费用记录"
'返回：正常=加了所有者前缀的对象串,如"ZLPER.部门表,ZLHIS.病人费用记录",取消="取消"
'参考：grsObject
    Dim rsTmp As ADODB.Recordset
    Dim strOwner As String, strSQL As String
    Dim i As Integer, j As Integer
    Dim blnNoSel As Boolean
    Dim blnFlag As Boolean
    Dim blnNextChk As Boolean
    Dim strSelectOwner As String, strOtherConnectOwner As String
    Dim arrObject As Variant
    
    blnNextChk = True
    arrObject = Split(strObject, ",")
    For i = 0 To UBound(arrObject)
        If arrObject(i) <> "DUAL" Then
            If InStr(arrObject(i), ".") > 0 Then
                '如果本身就加了所有者前缀,则使用其本身不变
                If InStr(ObjectOwner, "," & arrObject(i)) = 0 Then
                    ObjectOwner = ObjectOwner & "," & arrObject(i)
                End If
            Else
                If intConnect > Val("0-当前登录连接") Then
                    '其他数据连接
                    strOtherConnectOwner = mdlPublic.GetDBConnectInfo(intConnect, Val("1-用户名"))
                    If strOtherConnectOwner <> "" Then
                        ObjectOwner = ObjectOwner & "," & strOtherConnectOwner & "." & arrObject(i)
                    End If
                Else
                    grsObject.Filter = "OBJECT_NAME='" & arrObject(i) & "'"
                    If grsObject.RecordCount = 1 Then
                        If InStr(ObjectOwner & ",", "," & grsObject!Owner & "." & arrObject(i) & ",") = 0 Then
                            ObjectOwner = ObjectOwner & "," & grsObject!Owner & "." & arrObject(i)
                        End If
                    ElseIf grsObject.RecordCount > 1 Then
                        '如果除后备所有者之外，只剩一个在线所有者，则直接为在线所有者
                        blnNoSel = False: strOwner = ""
                        
                        grsObject.MoveFirst
                        Do While Not grsObject.EOF
                            strOwner = strOwner & ",'" & grsObject!Owner & "'"
                            grsObject.MoveNext
                        Loop
                        grsObject.MoveFirst
                        strOwner = Mid(strOwner, 2)
                        
                        On Error GoTo errH
                        strSQL = _
                            " Select Column_Value As 所有者 From Table(Cast(zlTools.f_Str2List ('" & Replace(strOwner, "'", "") & "') as zlTools.t_StrList))" & _
                            " Minus" & _
                            " Select 所有者 From zlBakSpaces Where 所有者 IN(" & strOwner & ")"
                        strSQL = _
                            "Select A.所有者,Decode(B.所有者,Null,0,1) as 系统者 " & _
                            "From (" & strSQL & ") A,(Select Distinct 所有者 From zlSystems) B Where A.所有者=B.所有者(+)"
                        Set rsTmp = OpenSQLRecord(strSQL, "ObjectOwner")
                        If rsTmp.RecordCount = 1 Then
                            If rsTmp!系统者 = 1 Then
                                strOwner = rsTmp!所有者
                                blnNoSel = True
                            End If
                        End If
                        On Error GoTo 0
                        
                        If blnNoSel Then
                            If InStr(ObjectOwner & ",", "," & strOwner & "." & arrObject(i) & ",") = 0 Then
                                ObjectOwner = ObjectOwner & "," & strOwner & "." & arrObject(i)
                            End If
                        Else
                            '同一对象有多个所有者,则要求选择
                            Set frmSelOwner.rsObject = grsObject
                            If blnFlag = False Then
                                frmSelOwner.chkNext.Value = IIF(blnNextChk, 1, 0)
                                If frmParent Is Nothing Then
                                    frmSelOwner.Show 1
                                Else
                                    frmSelOwner.Show 1, frmParent
                                End If
                                If gblnOK Then
                                    blnFlag = frmSelOwner.chkNext.Value
                                    blnNextChk = frmSelOwner.chkNext.Value
                                    If blnFlag Then
                                        strSelectOwner = frmSelOwner.lvw.SelectedItem.Text
                                    End If
                                End If
                            End If
                            If gblnOK Then
                                If blnFlag = False Then
                                    With frmSelOwner.lvw.SelectedItem
                                        If InStr(ObjectOwner & ",", "," & .Text & "." & arrObject(i) & ",") = 0 Then
                                            ObjectOwner = ObjectOwner & "," & .Text & "." & arrObject(i)
                                        End If
                                    End With
                                    Unload frmSelOwner
                                Else
                                    If InStr(ObjectOwner & ",", "," & strSelectOwner & "." & arrObject(i) & ",") = 0 Then
                                        ObjectOwner = ObjectOwner & "," & strSelectOwner & "." & arrObject(i)
                                    End If
                                    '无法判断是否已经卸载，Unload后始终不为Nothing
                                    Unload frmSelOwner
                                End If
                            Else
                                '取消选择,也就是取消操作(调用程序),返回空
                                ObjectOwner = "取消": Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next
    If ObjectOwner <> "" Then ObjectOwner = Mid(ObjectOwner, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function SQLOwner(ByVal strSQL As String, strOwner As String) As String
'功能：将SQL语句替换成带对象所有者的形式
'参数：strSQL=原始SQL语句,strOwner=对象所有者串,如"ZLPER.部门表,ZLHIS.病人费用记录"
'返回：访问对象加了所有者前缀的SQL语句
'说明：1.本函数用于直接执行用户SQL语句,而不需要授权对象的私有同义词。
'      2.对表名与字段名相同且字段名没有带表别名,则会出错
    Dim i As Long, j As Long
    Dim intLoc As Long, blnDo As Boolean
    
    '处理成只用空格间隔
    strSQL = SpaceSQL(strSQL)
    
    For i = 0 To UBound(Split(strOwner, ","))
        '采用循环确认方式,确保替换的是表名,而不是其它语句部份或被包含在其它表名中的部份
        j = 0 '当前开始查找位置
        Do
            j = j + 1
            intLoc = InStr(j, strSQL, Split(Split(strOwner, ",")(i), ".")(1))
            If intLoc > 12 Then '至少有"SELECT FROM "
                '本身就有所有者前缀的不替换
                blnDo = True
                '右边以空格、","号、右括号结束
                blnDo = blnDo And (InStr(",) ", Mid$(strSQL, intLoc + Len(Split(Split(strOwner, ",")(i), ".")(1)), 1)) > 0)
                '左边则为","号或"FROM "
                blnDo = blnDo And (Mid$(strSQL, intLoc - 1, 1) = "," Or UCase(Mid$(strSQL, intLoc - 5, 5)) = "FROM ")
                If blnDo Then
                    strSQL = Left(strSQL, intLoc - 1) & _
                        Replace(strSQL, Split(Split(strOwner, ",")(i), ".")(1), Split(strOwner, ",")(i), intLoc, 1)
                    j = intLoc + Len(Split(strOwner, ",")(i))
                End If
            End If
        Loop Until j >= Len(strSQL)
    Next
    SQLOwner = strSQL
End Function

Public Function SpaceSQL(ByVal strSQL As String) As String
'功能：将SQL语句变换为只为空格间隔的形式,以便于分析
    Dim i As Long, j As Long, lngB As Long, lngE As Long
    Dim arrSeg() As Variant
                
    strSQL = Replace(strSQL, vbCr, " ")
    strSQL = Replace(strSQL, vbLf, " ")
    strSQL = Replace(strSQL, vbTab, " ")
    
    lngB = -1
    arrSeg = Array()
    For i = 1 To Len(strSQL)
        If Mid(strSQL, i, 1) = "'" Then
            If lngB = -1 Then
                lngB = i
            Else
                ReDim Preserve arrSeg(UBound(arrSeg) + 1)
                arrSeg(UBound(arrSeg)) = lngB & "," & i
                lngB = -1
            End If
        End If
    Next
    If lngB = -1 Then
        For i = 0 To UBound(arrSeg)
            lngB = CLng(Split(arrSeg(i), ",")(0)) + 1
            lngE = CLng(Split(arrSeg(i), ",")(1)) - 1
            For j = lngB To lngE
                If Mid(strSQL, j, 1) = " " Then
                    strSQL = Left(strSQL, j - 1) & Chr(250) & Mid(strSQL, j + 1)
                End If
            Next
        Next
    End If
    
    Do While InStr(strSQL, "  ") > 0
        strSQL = Replace(strSQL, "  ", " ")
    Loop
    
    strSQL = Replace(strSQL, Chr(250), " ")
    
    strSQL = Replace(strSQL, " ,", ",")
    strSQL = Replace(strSQL, ", ", ",")
    SpaceSQL = strSQL
End Function

Public Sub CopyReport(ByVal objS As Report, ByRef objO As Report)
'功能：拷贝报表对象,防止因Set造成地址的访问
    Dim objItem As RPTItem, objData As RPTData
    Dim objPar As RPTPar, objPars As RPTPars
    Dim i As Integer
    
    Set objO = New Report
    
    objO.系统 = objS.系统
    objO.编号 = objS.编号
    objO.名称 = objS.名称
    objO.说明 = objS.说明
    objO.打印机 = objS.打印机
    objO.进纸 = objS.进纸
    objO.票据 = objS.票据
    objO.打印方式 = objS.打印方式
    objO.禁止开始时间 = objS.禁止开始时间
    objO.禁止结束时间 = objS.禁止结束时间
    
    objO.blnLoad = objS.blnLoad
    objO.bytFormat = objS.bytFormat
    objO.intGridCount = objS.intGridCount
    objO.intGridID = objS.intGridID
    
    For i = 1 To objS.Fmts.count
        With objS.Fmts(i)
            objO.Fmts.Add .序号, .说明, .W, .H, .纸张, .纸向, .动态纸张, .图样, "_" & .序号
        End With
    Next
    
    For Each objItem In objS.Items
        With objItem
            objO.Items.Add .id, .格式号, .名称, .上级ID, .类型, .序号, .参照, .性质, .内容, .表头, .X, .Y, .W, .H _
                , .行高, .对齐, .自调, .字体, .字号, .粗体, .下线, .斜体, .网格, .前景, .背景, .边框 _
                , IIF(.分栏 < 1 And .类型 <> 6, 1, .分栏), .排序, .格式, .汇总, .表格线加粗, .自适应行高, .图片 _
                , .系统, .父ID, .SubIDs, .CopyIDs, "_" & .id, .数据源, .上下间距, .左右间距, .源行号, .横向分栏 _
                , .纵向分栏, .Relations, .ColProtertys, .水平反转
        End With
    Next
    For Each objData In objS.Datas
        With objData
            Set objPars = New RPTPars
            For Each objPar In .Pars
                objPars.Add objPar.组名, objPar.序号, objPar.名称, objPar.类型, objPar.缺省值, objPar.格式, objPar.值列表 _
                    , objPar.分类SQL, objPar.明细SQL, objPar.分类字段, objPar.明细字段, objPar.对象, "_" & objPar.序号 _
                    , objPar.Reserve, objPar.是否锁定
            Next
            objO.Datas.Add .名称, .数据连接编号, .SQL, .字段, .对象, .类型, .说明, objPars, "_" & .名称
        End With
    Next
End Sub

Public Function IncStr(ByVal strVal As String) As String
'功能：对一个字符串自动加1。
'说明：每一位进位时,如果是数字,则按十进制处理,否则按26进制处理
    Dim i As Integer, strTmp As String, bytUp As Byte, bytAdd As Byte
    
    For i = Len(strVal) To 1 Step -1
        If i = Len(strVal) Then
            bytAdd = 1
        Else
            bytAdd = 0
        End If
        If IsNumeric(Mid(strVal, i, 1)) Then
            If CByte(Mid(strVal, i, 1)) + bytAdd + bytUp < 10 Then
                strVal = Left(strVal, i - 1) & CByte(Mid(strVal, i, 1)) + bytAdd + bytUp & Mid(strVal, i + 1)
                bytUp = 0
            Else
                strVal = Left(strVal, i - 1) & "0" & Mid(strVal, i + 1)
                bytUp = 1
            End If
        Else
            If Asc(Mid(strVal, i, 1)) + bytAdd + bytUp <= Asc("Z") Then
                strVal = Left(strVal, i - 1) & Chr(Asc(Mid(strVal, i, 1)) + bytAdd + bytUp) & Mid(strVal, i + 1)
                bytUp = 0
            Else
                strVal = Left(strVal, i - 1) & "0" & Mid(strVal, i + 1)
                bytUp = 1
            End If
        End If
        If bytUp = 0 Then Exit For
    Next
    IncStr = strVal
End Function

Public Function GetNextNO(Optional ByVal blnGroup As Boolean = False) As String
'功能：获取下一个报表编号
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, blnExist As Boolean
    Const strGroup As String = "GROUP"
    Const strReport As String = "REPORT"
    
    On Error GoTo errH
    
    If Not blnGroup Then
        strSQL = "Select Max(编号) as 编号 From zlReports Where 编号 Like [1]"
    Else
        strSQL = "Select Max(编号) as 编号 From zlRPTGroups Where 编号 Like [2]"
    End If
    Set rsTmp = OpenSQLRecord(strSQL, "GetNextNO", "REPORT%", "GROUP%")
    If Not rsTmp.EOF Then
        If IsNull(rsTmp!编号) Then
            GetNextNO = IIF(blnGroup, strGroup, strReport) & "_001"
        Else
            GetNextNO = IncStr(rsTmp!编号)
        End If
    Else
        GetNextNO = IIF(blnGroup, strGroup, strReport) & "_001"
    End If
    
    Do
        blnExist = False
        blnExist = blnExist Or CheckExist("zlReports", "编号", GetNextNO)
        If Not blnExist Then blnExist = CheckExist("zlRPTGroups", "编号", GetNextNO)
        If blnExist Then GetNextNO = IncStr(GetNextNO)
    Loop Until Not blnExist
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetValue(Str As String, i As Integer) As String
    GetValue = Mid(Str, i)
    GetValue = Left(GetValue, InStr(GetValue, "]") - 1)
End Function

Public Function InDesign() As Boolean
'功能：判断当前运行程序是否在VB的工程环境中
    On Error Resume Next
    Debug.Print 1 / 0
    If Err.Number <> 0 Then Err.Clear: InDesign = True
End Function

Public Function SelMessage(ByVal hwnd As Long, ByVal Msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
    If Msg = WM_GETMINMAXINFO Then

        Dim MinMax As MINMAXINFO
        CopyMemory MinMax, ByVal lp, Len(MinMax)
        MinMax.ptMinTrackSize.X = 400
        MinMax.ptMinTrackSize.Y = 300
        MinMax.ptMaxTrackSize.X = 1600
        MinMax.ptMaxTrackSize.Y = 1200
        CopyMemory ByVal lp, MinMax, Len(MinMax)
        SelMessage = 1
        Exit Function
    End If
    SelMessage = CallWindowProc(glngSelProc, hwnd, Msg, wp, lp)
End Function

Public Function GetDBUser() As String
'功能：获取当前登录数据库用户名
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
        
    On Error GoTo errH
    If gstrDBUser <> "" Then
        GetDBUser = gstrDBUser
        Exit Function
    End If
        
    If gcnOracle Is Nothing Then Exit Function
    If gcnOracle.State = adStateClosed Then Exit Function
    If InStr(UCase(gcnOracle.ConnectionString), "USER ID=") > 0 Then
        For i = 0 To UBound(Split(UCase(gcnOracle.ConnectionString), ";"))
            If Split(UCase(gcnOracle.ConnectionString), ";")(i) Like "USER ID=*" Then
                GetDBUser = Trim(Split(Split(UCase(gcnOracle.ConnectionString), ";")(i), "=")(1))
                Exit For
            End If
        Next
    Else
        strSQL = "Select User From Dual"
        Call OpenRecord(rsTmp, strSQL, "mdlPublic_GetDBUser")
        If Not rsTmp.EOF Then GetDBUser = rsTmp!User
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetTheUserName(ByVal strUser As String) As String
'功能：获取指定用户的姓名
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
        
    On Error GoTo errH
        
    If gcnOracle Is Nothing Then Exit Function
    If gcnOracle.State = adStateClosed Then Exit Function
    If strUser = "" Then Exit Function
    strSQL = " Select A.姓名,A.编号" & _
        " From 人员表 A,上机人员表 B" & _
        " Where A.ID=B.人员ID And B.用户名='" & strUser & "'"
    Call OpenRecord(rsTmp, strSQL, "GetTheUserName")
    If Not rsTmp.EOF Then GetTheUserName = rsTmp!姓名 & ""
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub AutoSizeCol(lvw As Object)
'功能：根据自动ListView当前内容自动调整各列宽度
'参数：blnByHead=是否按列头文本调整,Col=指定列还是所有列(1-N)
    Dim i As Integer, lngW As Long
    For i = 1 To lvw.ColumnHeaders.count
        SendMessage lvw.hwnd, LVM_SETCOLUMNWIDTH, i - 1, LVSCW_AUTOSIZE
        If lvw.ColumnHeaders(i).Width < 200 Then lvw.ColumnHeaders(i).Width = 0
        If lvw.ColumnHeaders(i).Width < (TLen(lvw.ColumnHeaders(i).Text) + 2) * 90 And lvw.ColumnHeaders(i).Width <> 0 Then lvw.ColumnHeaders(i).Width = (TLen(lvw.ColumnHeaders(i).Text) + 2) * 90
    Next
End Sub

Public Function GetExpField(objFld As ADODB.Field, Optional ByVal blnDataNum As Boolean) As String
'功能：导出报表时用
'参数：blnDataNum=true 源ID按实际值导出
    Dim strTmp As String
    
    If IsNull(objFld.Value) Then
        Exit Function
    ElseIf InStr(",系统,程序ID,功能,发布时间,", "," & objFld.name & ",") > 0 Then
        Exit Function
    ElseIf objFld.name = "编号" Then
        GetExpField = "[编号]" '导入时取当前时间
    ElseIf objFld.name = "修改时间" Then
        GetExpField = "Sysdate" '导入时取当前时间
    ElseIf objFld.name = "ID" Then
        GetExpField = "[NextVal]" '导入时取"当前表_ID.NextVal"
    ElseIf objFld.name = "上级ID" Then
        GetExpField = "[CurrVal-X]" '导入时取"当前表_ID.CurrVal-X",X为上级ID不为空的开始数
    ElseIf objFld.name = "报表ID" Then
        GetExpField = "[zlReports_ID.CurrVal]" '导入时取"zlReports_ID.CurrVal"
    ElseIf objFld.name = "源ID" And blnDataNum = False Then
        GetExpField = "[zlRPTDatas_ID.CurrVal]" '导入时取"zlRPTDatas_ID.CurrVal"
    ElseIf objFld.name = "元素ID" Then
        GetExpField = "[zlRPTItems_ID.CurrVal]" '导入时取"zlRPTDatas_ID.CurrVal"
    ElseIf objFld.name = "对象" Then
        GetExpField = Replace(UCase(objFld.Value), UCase(gstrDBUser) & ".", "USER.")
    Else '导入时根据数据类型转换取值
        Select Case objFld.type
            Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                GetExpField = objFld.Value
            Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                GetExpField = objFld.Value
            Case adDBTimeStamp, adDBTime, adDBDate, adDate
                If Format(objFld.Value, "HH:mm:ss") = "00:00:00" Then
                    GetExpField = Format(objFld.Value, "yyyy-MM-dd")
                Else
                    GetExpField = Format(objFld.Value, "yyyy-MM-dd HH:mm:ss")
                End If
            Case adBinary, adVarBinary, adLongVarBinary
                '暂时不支持图片的处理
        End Select
    End If
End Function

Private Function GetFieldNames(rsTmp As ADODB.Recordset) As String
'功能：返回一个记录集所具有的字段名称串
    Dim i As Integer
    For i = 0 To rsTmp.Fields.count - 1
        GetFieldNames = GetFieldNames & "," & rsTmp.Fields(i).name
    Next
    GetFieldNames = GetFieldNames & ","
End Function

Public Function ExportReport(lngRPTID As Long, strFile As String) As Boolean
'功能：导出一张自定义报表
'参数：lngRPTID=报表ID
'      strFile=文件名
'返回：导出是否成功。
'说明：
'      1.对于已发布的报表,导出成为非发布报表
'      2.目前不支持图片元素内容的导出
    Dim objFile As FileSystemObject, objText As TextStream
    Dim rsTmp As ADODB.Recordset
    Dim rsSub As ADODB.Recordset
    Dim rsSQL As ADODB.Recordset
    Dim objFld As ADODB.Field
    Dim i As Integer, j As Integer
    Dim blnOpen As Boolean, blnSub As Boolean
    Dim strSQL As String
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    
    strSQL = "Select ID,编号,名称,说明,密码,打印机,进纸,票据,打印方式,系统,程序ID,功能,修改时间,发布时间,禁止开始时间,禁止结束时间 From zlReports Where ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "ExportReport", lngRPTID)
    If rsTmp.EOF Then
        MsgBox "没有发现指定报表的数据！", vbInformation, App.Title
        Exit Function
    End If
    
    '打开磁盘文件
    Set objFile = New FileSystemObject
    If objFile.FileExists(strFile) Then Call objFile.DeleteFile(strFile, True)
    Set objText = objFile.CreateTextFile(strFile, True)
    blnOpen = True
    
    '产生报表表头
    Call objText.WriteLine("[HEAD]")
    Call objText.WriteLine("报表编号=" & rsTmp!编号)
    Call objText.WriteLine("报表名称=" & rsTmp!名称)
    Call objText.WriteLine("报表说明=" & IIF(IsNull(rsTmp!说明), "", rsTmp!说明))
    Call objText.WriteLine("导出用户=" & gstrDBUser)
    Call objText.WriteLine("导出时间=" & Format(Currentdate, "yyyy-MM-dd HH:mm:ss"))
    Call objText.WriteLine("禁止开始时间=" & Format(rsTmp!禁止开始时间 & "", "HH:mm:ss"))
    Call objText.WriteLine("禁止结束时间=" & Format(rsTmp!禁止结束时间 & "", "HH:mm:ss"))
    
    '报表:ZLReport,以分号为行结束；以分号为一个字段结束,单分号为一条记录结束
    Call objText.WriteLine("[ZLREPORTS]")
    Call objText.WriteLine(";")
    For Each objFld In rsTmp.Fields
        Call objText.WriteLine(objFld.name & "=" & GetExpField(objFld) & ";")
    Next
    
    '报表格式
    'Set rsTmp = New ADODB.Recordset
    strSQL = "Select 报表ID,序号,说明,W,H,纸张,纸向,动态纸张,图样 From zlRPTFmts Where 报表ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "ExportReport", lngRPTID)
    If Not rsTmp.EOF Then
        Call objText.WriteLine("[ZLRPTFMTS]")
        For i = 1 To rsTmp.RecordCount
            Call objText.WriteLine(";")
            For Each objFld In rsTmp.Fields
                Call objText.WriteLine(objFld.name & "=" & GetExpField(objFld) & ";")
            Next
            rsTmp.MoveNext
        Next
    End If
    
    '报表元素
    strSQL = "Select 系统,ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体 " & vbCrLf & _
             "    ,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,ID as 原ID,父ID,源ID,上下间距,左右间距" & vbCrLf & _
             "    ,源行号,横向分栏,纵向分栏,表格线加粗,自适应行高,水平反转 " & vbCrLf & _
             "From zlRPTItems " & vbCrLf & _
             "Where 报表ID=[1] " & vbCrLf & _
             "Start With 上级ID is NULL Connect by Prior ID=上级ID"
    Set rsTmp = OpenSQLRecord(strSQL, "ExportReport", lngRPTID)
    
    strSQL = "Select 报表ID,元素ID,条件名称,条件字段,条件关系,条件值,字体颜色,背景颜色,是否加粗,是否整行应用,对齐 " & vbCrLf & _
             "From zlRPTColProterty " & vbCrLf & _
             "Where 报表ID=[1]"
    Set rsSub = OpenSQLRecord(strSQL, "ExportReport", lngRPTID)
    
    blnSub = False
    If Not rsTmp.EOF Then
        Call objText.WriteLine("[ZLRPTITEMS]")
        For i = 1 To rsTmp.RecordCount
            If blnSub Then Call objText.WriteLine("[ZLRPTITEMS]")
            Call objText.WriteLine(";")
            blnSub = False
            For Each objFld In rsTmp.Fields
                Call objText.WriteLine(objFld.name & "=" & GetExpField(objFld, True) & ";")
            Next
            rsSub.Filter = "元素ID=" & rsTmp!id
            If rsSub.RecordCount > 0 Then
                blnSub = True
                rsSub.MoveFirst
                Call objText.WriteLine("[ZLRPTCOLPROTERTY]")
                For j = 1 To rsSub.RecordCount
                    Call objText.WriteLine(";")
                    For Each objFld In rsSub.Fields
                        Call objText.WriteLine(objFld.name & "=" & GetExpField(objFld, True) & ";")
                    Next
                    
                    rsSub.MoveNext
                Next
            End If
            rsTmp.MoveNext
        Next
    End If
    
    '报表数据,'数据参数
    strSQL = "Select ID,报表ID,数据连接编号,名称,字段,对象,类型,说明,ID as 原ID From zlRPTDatas Where 报表ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "ExportReport", lngRPTID)
    
    strSQL = "Select B.源ID,B.行号,B.内容 From zlRPTDatas A,zlRPTSQLs B Where A.ID=B.源ID And A.报表ID=[1]"
    Set rsSQL = OpenSQLRecord(strSQL, "ExportReport", lngRPTID)
    
    strSQL = "Select B.源ID,B.组名,B.序号,B.名称,B.类型,B.缺省值,B.格式,B.值列表,B.分类SQL,B.明细SQL,B.分类字段" & vbCrLf & _
             "    ,B.明细字段,B.对象,B.锁定 " & vbCrLf & _
             "From zlRPTDatas A,zlRPTPars B " & vbCrLf & _
             "Where A.ID=B.源ID And A.报表ID=[1]"
    Set rsSub = OpenSQLRecord(strSQL, "ExportReport", lngRPTID)
    
    blnSub = False
    If Not rsTmp.EOF Then
        Call objText.WriteLine("[ZLRPTDATAS]")
        For i = 1 To rsTmp.RecordCount
            If blnSub Then Call objText.WriteLine("[ZLRPTDATAS]")
            
            Call objText.WriteLine(";")
            For Each objFld In rsTmp.Fields
                Call objText.WriteLine(objFld.name & "=" & GetExpField(objFld) & ";")
            Next
            
            blnSub = False
            
            rsSQL.Filter = "源ID=" & rsTmp!id
            If Not rsSQL.EOF Then
                blnSub = True
                Call objText.WriteLine("[ZLRPTSQLS]")
                For j = 1 To rsSQL.RecordCount
                    Call objText.WriteLine(";")
                    For Each objFld In rsSQL.Fields
                        Call objText.WriteLine(objFld.name & "=" & GetExpField(objFld) & ";")
                    Next
                    rsSQL.MoveNext
                Next
            End If
           
            rsSub.Filter = "源ID=" & rsTmp!id
            If Not rsSub.EOF Then
                blnSub = True
                Call objText.WriteLine("[ZLRPTPARS]")
                For j = 1 To rsSub.RecordCount
                    Call objText.WriteLine(";")
                    For Each objFld In rsSub.Fields
                        Call objText.WriteLine(objFld.name & "=" & GetExpField(objFld) & ";")
                    Next
                    rsSub.MoveNext
                Next
            End If
            
            rsTmp.MoveNext
        Next
    End If
    
    objText.Close
    Screen.MousePointer = 0
    
    ExportReport = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
    If blnOpen Then objText.Close
End Function

Public Function ImportReport(ByVal strFile As String, Optional ByVal lngCurrID As Long _
    , Optional ByVal blnOnlyData As Boolean, Optional ByVal lngGroupID As Long _
    , Optional ByVal lngClassID As Long) As String
'功能:从文件导入一张报表,如果是覆盖固定报表还要重新处理访问权限
'参数:strFile=外部文件名
'     lngCurrID=将报表导入覆盖到指定ID的已有报表
'     blnOnlyData=是否只覆盖数据源
'     lngGroupID=报表组ID,0=导入所有报表中,<>0=导入到该报表组中
'     lngClassID=报表类ID
'返回:成功="ID|编号|名称|说明",失败=""
'说明：1.导入共享报表时如果编号重复,则自动取
'      2.覆盖已有报表时,当前报表信息不变,除了纸张信息
    Dim objFile As FileSystemObject, objText As TextStream
    Dim rsReport As New ADODB.Recordset
    Dim rsFMT As New ADODB.Recordset
    Dim rsItem As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim rsPar As New ADODB.Recordset
    Dim rsCol As New ADODB.Recordset
    Dim rsProgram As New ADODB.Recordset '用来判断是否存在该模块
    Dim rsCopy As ADODB.Recordset
    
    Dim blnTran As Boolean, blnOpen As Boolean, lngUPID As Long
    Dim strLine As String, strSect As String, strFld As String, strValue As String
    Dim blnReport  As Boolean, blnFmt As Boolean, blnItem As Boolean, blnData As Boolean, blnPar As Boolean, blnSQL As Boolean
    Dim strReport As String, StrFmt As String, strItem As String, strData As String, StrPar As String, strRSQL As String
    Dim strPreNum As String, strNum As String, strName As String, strNote As String, lngRPTID As Long
    
    Dim rsCurr As New ADODB.Recordset
    Dim rsPriv As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim lngRptW As Long, lngRptH As Long, bln动态纸张 As Boolean
    Dim int纸张 As Integer, int纸向 As Integer
    Dim strObject As String, strSQL As String, i As Long
    Dim Col As New Collection
    Dim ColData As New Collection
    Dim rsItemCopy As New Recordset
    Dim lng序号 As Long
    Dim str禁止开始时间 As String, str禁止结束时间 As String
    Dim strCol As String, strTmp As String
    Dim blnCol As Boolean
    
    On Error GoTo errH
    
    If lngCurrID = 0 Then blnOnlyData = False
    
    '当前的报表信息
    If lngCurrID <> 0 Then
        strSQL = _
            "Select ID, 编号, 名称, 说明, 密码, 打印机, 进纸, 票据, 打印方式, 系统, 程序id, 功能, 修改时间" & vbNewLine & _
            "    , 发布时间, 禁止开始时间, 禁止结束时间 " & vbNewLine & _
            "From zlReports " & vbNewLine & _
            "Where ID = [1] "
        Set rsCurr = OpenSQLRecord(strSQL, "ExportReport", lngCurrID)
        If rsCurr.EOF Then Exit Function
    End If
    
    '打开报表文件
    Set objFile = New FileSystemObject
    If Not objFile.FileExists(strFile) Then Exit Function
    Set objText = objFile.OpenTextFile(strFile)
    blnOpen = True
    
    '打开新数据记录集
    If lngCurrID = 0 Then
        rsReport.CursorLocation = adUseClient
        rsReport.Open _
                "Select ID,编号,名称,说明,密码,打印机,进纸,票据,打印方式,系统,程序ID,功能" & vbNewLine & _
                "    ,修改时间,发布时间,禁止开始时间,禁止结束时间,分类ID " & vbNewLine & _
                "From zlReports " & vbNewLine & _
                "Where Rownum<1" _
            , gcnOracle, adOpenKeyset, adLockOptimistic
        strReport = GetFieldNames(rsReport)
    End If
    
    If Not blnOnlyData Then
        rsFMT.CursorLocation = adUseClient
        rsFMT.Open "Select 报表ID,序号,说明,W,H,纸张,纸向,动态纸张,图样 From zlRPTFmts Where Rownum<1", gcnOracle, adOpenKeyset, adLockOptimistic
        StrFmt = GetFieldNames(rsFMT)
        
        rsItem.CursorLocation = adUseClient
        rsItem.Open "Select 系统,ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,父ID,源ID,上下间距,左右间距,源行号,横向分栏,纵向分栏,表格线加粗,自适应行高,水平反转 From zlRPTItems Where Rownum<1", gcnOracle, adOpenKeyset, adLockOptimistic
        strItem = GetFieldNames(rsItem)
    End If
    
    rsData.CursorLocation = adUseClient
    rsData.Open "Select ID,报表ID,数据连接编号,名称,字段,对象,类型,说明 From zlRPTDatas Where Rownum<1", gcnOracle, adOpenKeyset, adLockOptimistic
    strData = GetFieldNames(rsData)
    
    rsSQL.CursorLocation = adUseClient
    rsSQL.Open "Select 源ID,行号,内容 From zlRPTSQLs Where Rownum<1", gcnOracle, adOpenKeyset, adLockOptimistic
    strRSQL = GetFieldNames(rsSQL)
    
    rsPar.CursorLocation = adUseClient
    rsPar.Open "Select 源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象,锁定 From zlRPTPars Where Rownum<1", gcnOracle, adOpenKeyset, adLockOptimistic
    StrPar = GetFieldNames(rsPar)
    
    rsCol.CursorLocation = adUseClient
    rsCol.Open "Select 报表ID,元素ID,条件名称,条件字段,条件关系,条件值,字体颜色,背景颜色,是否加粗,是否整行应用,对齐 From zlRPTColProterty Where  Rownum<1", gcnOracle, adOpenKeyset, adLockOptimistic
    strCol = GetFieldNames(rsCol)
    
    rsItemCopy.Fields.Append "ID", adBigInt, , adFldIsNullable
    rsItemCopy.Fields.Append "父ID", adBigInt, , adFldIsNullable
    rsItemCopy.Fields.Append "源ID", adBigInt, , adFldIsNullable
    rsItemCopy.CursorLocation = adUseClient
    rsItemCopy.CursorType = adOpenStatic
    rsItemCopy.LockType = adLockOptimistic
    rsItemCopy.Open
    gcnOracle.BeginTrans
    blnTran = True
            
    '覆盖固定报表时,仅保留报表信息(含所属报表组)
    If lngCurrID <> 0 Then
        If Not blnOnlyData Then gcnOracle.Execute "Delete From zlRPTFmts Where 报表ID=" & lngCurrID
        gcnOracle.Execute "Delete From zlRPTDatas Where 报表ID=" & lngCurrID
    End If
    
    Do While Not objText.AtEndOfStream
        strLine = objText.ReadLine
        
        '判断格式是否正确
        If strSect = "" And Trim(strLine) <> "" And Trim(strLine) <> "[HEAD]" Then
            objText.Close
            gcnOracle.RollbackTrans
            Exit Function
        End If
        
        '取得段号
        If Left(strLine, 1) = "[" And Right(strLine, 1) = "]" Then
            strSect = UCase(Mid(strLine, 2, Len(strLine) - 2))
        End If
        
        '处理报表头
        If strSect = "HEAD" Then
            '报表编号
            If strLine Like "报表编号=*" Then
                strNum = Mid(strLine, InStr(strLine, "=") + 1)
                strPreNum = strNum
                
                '共享报表:如果重复则另外取一个编号
                If lngCurrID = 0 Then
                    If CheckExist("zlReports", "编号", strNum) Then
                        strNum = GetNextNO
                    End If
                End If
            End If
            '报表名称
            If strLine Like "报表名称=*" Then strName = Mid(strLine, InStr(strLine, "=") + 1)
            '报表说明
            If strLine Like "报表说明=*" Then strNote = Mid(strLine, InStr(strLine, "=") + 1)
            '报表禁止时间
            If strLine Like "禁止开始时间=*" Then str禁止开始时间 = Format(Mid(strLine, InStr(strLine, "=") + 1), "HH:mm:ss")
            If strLine Like "禁止结束时间=*" Then str禁止结束时间 = Format(Mid(strLine, InStr(strLine, "=") + 1), "HH:mm:ss")
        End If
        
        '处理数据
        '新增一条记录
        If strLine = ";" Then
            If blnReport Then rsReport.Update
            If blnFmt Then rsFMT.Update
            If blnItem Then rsItem.Update
            If blnData Then rsData.Update
            If blnSQL Then rsSQL.Update
            If blnPar Then rsPar.Update
            If blnCol Then rsCol.Update
            
            Select Case strSect
                Case "ZLREPORTS"
                    If lngCurrID = 0 Then
                        rsReport.AddNew
                        If lngClassID > 0 Then rsReport!分类id = lngClassID
                        blnReport = True
                    End If
                Case "ZLRPTFMTS"
                    If Not blnOnlyData Then
                        rsFMT.AddNew: blnFmt = True
                        '兼容以前导出的报表格式,所有格式统一纸张
                        If InStr(StrFmt, ",纸张,") > 0 And int纸张 <> 0 Then
                            rsFMT!W = lngRptW
                            rsFMT!H = lngRptH
                            rsFMT!纸张 = int纸张
                            rsFMT!纸向 = int纸向
                            rsFMT!动态纸张 = IIF(bln动态纸张, 1, 0)
                        End If
                    End If
                Case "ZLRPTITEMS"
                    If Not blnOnlyData Then
                        rsItem.AddNew: blnItem = True
                    End If
                Case "ZLRPTDATAS"
                    rsData.AddNew: blnData = True
                Case "ZLRPTSQLS"
                    rsSQL.AddNew: blnSQL = True
                Case "ZLRPTPARS"
                    rsPar.AddNew: blnPar = True
                Case "ZLRPTCOLPROTERTY"
                    rsCol.AddNew: blnCol = True
            End Select
        End If

        '循环取由多行文本组成的大数据源
        If InStr(strLine, "=") > 0 And Right(strLine, 1) <> ";" And strSect <> "HEAD" Then
            Do While Not objText.AtEndOfStream And Right(strLine, 1) <> ";"
                strLine = strLine & vbCrLf & objText.ReadLine
            Loop
        End If
        
        '字段取值
        If InStr(strLine, "=") > 0 And Right(strLine, 1) = ";" And strSect <> "HEAD" Then
            strFld = Left(strLine, InStr(strLine, "=") - 1)
            strValue = Mid(strLine, InStr(strLine, "=") + 1)
            strValue = Left(strValue, Len(strValue) - 1)

            If UCase(strFld) = "原ID" And UCase(strSect) = "ZLRPTITEMS" And blnOnlyData = False Then
                Col.Add rsCopy.Fields("ID").Value, "_" & strValue
            End If
            If UCase(strFld) = "原ID" And UCase(strSect) = "ZLRPTDATAS" And blnOnlyData = False Then
                ColData.Add rsData.Fields("ID").Value, "_" & strValue
            End If
            '处理卡片数据源对照和处理控件父子关系
            If (UCase(strFld) = "源ID" Or UCase(strFld) = "父ID") And UCase(strSect) = "ZLRPTITEMS" And blnOnlyData = False Then
                rsItemCopy.Filter = "ID=" & rsItem.Fields("ID").Value
                If rsItemCopy.RecordCount = 0 Then
                    rsItemCopy.AddNew
                    rsItemCopy!id = rsItem.Fields("ID").Value
                End If
                If strValue <> "" Then
                    If UCase(strFld) = "父ID" Then
                        rsItemCopy!父ID = Val(strValue)
                    ElseIf UCase(strFld) = "源ID" Then
                        rsItemCopy!源ID = Val(strValue)
                    End If
                End If
                rsItemCopy.Update
                
                strValue = ""
            End If

            If strFld = "上级ID" Then
                If strValue = "" Then
                    lngUPID = 0
                Else
                    lngUPID = lngUPID + 1
                End If
            End If
            
            '取报表文件中的纸张设置信息,用于兼容老结构的导出报表
            If strSect = "ZLREPORTS" Then
                If UCase(strFld) = "W" Then lngRptW = Val(strValue)
                If UCase(strFld) = "H" Then lngRptH = Val(strValue)
                If strFld = "纸张" Then int纸张 = Val(strValue)
                If strFld = "纸向" Then int纸向 = Val(strValue)
                If strFld = "动态纸张" Then bln动态纸张 = Val(strValue) = 1
            End If
            
            '判断是否有该字段
            Set rsCopy = Nothing
            If strValue <> "" Then '值为空则不赋值
                Select Case strSect
                    Case "ZLREPORTS"
                        If lngCurrID = 0 Then
                            If InStr(strReport, "," & strFld & ",") > 0 Then
                                Set rsCopy = rsReport
                            End If
                        End If
                    Case "ZLRPTFMTS"
                        If Not blnOnlyData Then
                            If InStr(StrFmt, "," & strFld & ",") > 0 Then
                                Set rsCopy = rsFMT
                            End If
                        End If
                    Case "ZLRPTITEMS"
                        If Not blnOnlyData Then
                            If InStr(strItem, "," & strFld & ",") > 0 Then
                                Set rsCopy = rsItem
                            End If
                        End If
                    Case "ZLRPTDATAS"
                        If InStr(strData, "," & strFld & ",") > 0 Then Set rsCopy = rsData
                    Case "ZLRPTSQLS"
                        If InStr(strRSQL, "," & strFld & ",") > 0 Then Set rsCopy = rsSQL
                    Case "ZLRPTPARS"
                        If InStr(StrPar, "," & strFld & ",") > 0 Then Set rsCopy = rsPar
                    Case "ZLRPTCOLPROTERTY"
                        If InStr(strCol, "," & strFld & ",") > 0 Then Set rsCopy = rsCol
                End Select
            End If
            
            If Not rsCopy Is Nothing Then
                '合法性检查
                If strSect = "ZLREPORTS" And strFld = "密码" Then
                    If GetPass(strPreNum, strName) <> strValue Then
                        objText.Close
                        gcnOracle.RollbackTrans
                        Exit Function
                    End If
                End If
                '增加
                If UCase(strValue) = UCase("SysDate") Then
                    rsCopy.Fields(strFld).Value = Currentdate
                ElseIf UCase(strValue) = UCase("[编号]") Then
                    rsCopy.Fields(strFld).Value = strNum
                ElseIf strSect = "ZLREPORTS" And strFld = "密码" Then
                    rsCopy.Fields(strFld).Value = GetPass(strNum, strName)
                ElseIf UCase(strValue) = UCase("[NextVal]") Then
                    rsCopy.Fields(strFld).Value = GetNextID(strSect)
                    If UCase(strSect) = ("ZLREPORTS") Then lngRPTID = rsCopy.Fields(strFld).Value
                ElseIf UCase(strValue) = UCase("[zlReports_ID.CurrVal]") Then
                    If lngCurrID = 0 Then
                        rsCopy.Fields(strFld).Value = GetCurrID("zlReports")
                    Else
                        rsCopy.Fields(strFld).Value = lngCurrID
                    End If
                ElseIf UCase(strValue) = UCase("[zlRPTDatas_ID.CurrVal]") Then
                    rsCopy.Fields(strFld).Value = GetCurrID("zlRPTDatas")
                ElseIf UCase(strValue) = UCase("[zlRPTItems_ID.CurrVal]") Then
                    rsCopy.Fields(strFld).Value = GetCurrID("zlRPTItems")
                ElseIf UCase(strValue) = UCase("[CurrVal-X]") Then
                    rsCopy.Fields(strFld).Value = GetCurrID(strSect) - lngUPID
                ElseIf rsCopy.Fields(strFld).name = "对象" Then
                    rsCopy.Fields(strFld).Value = Replace(strValue, "USER.", UCase(gstrDBUser) & ".")
                Else
                    Select Case rsCopy.Fields(strFld).type
                        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                            rsCopy.Fields(strFld).Value = strValue
                        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                            rsCopy.Fields(strFld).Value = Val(strValue)
                        Case adDBTimeStamp, adDBTime, adDBDate, adDate
                            If IsDate(strValue) Then rsCopy.Fields(strFld).Value = CDate(strValue)
                        Case adBinary, adVarBinary, adLongVarBinary
                            '暂时不支持图片处理
                    End Select
                End If
            End If
        End If
    Loop
    
    If blnReport Then rsReport.Update
    If blnFmt Then rsFMT.Update
    If blnItem Then rsItem.Update
    If blnData Then rsData.Update
    If blnSQL Then rsSQL.Update
    If blnPar Then rsPar.Update
    If blnCol Then rsCol.Update
    '处理父ID和源ID对照
    If blnOnlyData = False Then
        rsItemCopy.Filter = "父ID <> ''"
        If rsItemCopy.RecordCount > 0 Then
            rsItemCopy.MoveFirst
            Do While Not rsItemCopy.EOF
                rsItem.Filter = "ID=" & rsItemCopy!id
                rsItem!父ID = Val(Col("_" & rsItemCopy!父ID))
                rsItem.Update
                rsItemCopy.MoveNext
            Loop
        End If
        rsItemCopy.Filter = "源ID <> ''"
        If rsItemCopy.RecordCount > 0 Then
            rsItemCopy.MoveFirst
            Do While Not rsItemCopy.EOF
                rsItem.Filter = "ID=" & rsItemCopy!id
                rsItem!源ID = Val(ColData("_" & rsItemCopy!源ID))
                rsItem.Update
                rsItemCopy.MoveNext
            Loop
        End If
    End If
        
    '更新部份报表信息
    If lngCurrID <> 0 Then
        gcnOracle.Execute _
            "Update zlReports" & _
            " Set 修改时间=Sysdate,发布时间=Decode(发布时间,NULL,NULL,Sysdate)" & _
            ",禁止开始时间=to_date('" & str禁止开始时间 & "','HH24:MI:SS')" & _
            ",禁止结束时间=to_date('" & str禁止结束时间 & "','HH24:MI:SS')" & _
            " Where ID=" & lngCurrID
    End If
    
    '对于已有报表,报表相关权限重新填写.新导入共享报表不存在重新授权
    If lngCurrID <> 0 Then
        strSQL = _
            " Select 系统,程序ID,功能,说明 From zlReports" & _
            " Where 程序ID is Not NULL And 功能 is Not NULL And ID=[1]" & _
            " Union All" & _
            " Select A.系统,A.程序ID,B.功能,C.说明" & _
            " From zlRptGroups A,zlRptSubs B,zlReports C" & _
            " Where A.ID=B.组ID And B.报表ID=C.ID And A.程序ID is Not NULL" & _
            " And B.功能 is Not NULL And B.报表ID=[1]" & _
            " Union ALL" & _
            " Select A.系统,A.程序ID,A.功能,B.说明" & _
            " From zlRPTPuts A,zlReports B Where A.报表ID=B.ID And A.报表ID=[1]"
        Set rsTmp = OpenSQLRecord(strSQL, "ExportReport", lngCurrID)
        If Not rsTmp.EOF Then
            '数据源所涉及的对象
            strSQL = _
                "Select Distinct B.对象 From zlReports A,zlRptDatas B " & _
                "Where A.ID=B.报表ID And B.对象 is Not NULL And A.ID=[1] " & _
                "    And nvl(b.数据连接编号,0) <= 0 "
            Set rsPriv = OpenSQLRecord(strSQL, "ExportReport", lngCurrID)
            Do While Not rsPriv.EOF
                For i = 0 To UBound(Split(rsPriv!对象, ","))
                    strTmp = Split(rsPriv!对象, ",")(i)
                    If InStr(strObject & ",", "," & strTmp & ",") = 0 Then
                        If InStr(",SYS,SYSTEM,ZLTOOLS,", "," & UCase(Split(strTmp, ".")(0)) & ",") = 0 Then
                            strObject = strObject & "," & strTmp
                        End If
                    End If
                Next
                rsPriv.MoveNext
            Loop
            
            '参数所涉及的对象
            strSQL = _
                "Select Distinct Replace(C.对象,'|',',') as 对象 " & _
                "From zlReports A,zlRptDatas B,zlRptPars C " & _
                "Where A.ID=B.报表ID And B.ID=C.源ID And C.对象 is Not NULL " & _
                "    And nvl(b.数据连接编号,0) <= 0 And A.ID=[1]"
            Set rsPriv = OpenSQLRecord(strSQL, "ExportReport", lngCurrID)
            Do While Not rsPriv.EOF
                For i = 0 To UBound(Split(rsPriv!对象, ","))
                    strTmp = Split(rsPriv!对象, ",")(i)
                    If InStr(strObject & ",", "," & strTmp & ",") = 0 And strTmp <> "" Then
                        If InStr(",SYS,SYSTEM,ZLTOOLS,", "," & UCase(Split(strTmp, ".")(0)) & ",") = 0 Then
                            strObject = strObject & "," & strTmp
                        End If
                    End If
                Next
                rsPriv.MoveNext
            Loop
            strObject = Mid(strObject, 2)
            
            '更新权限
            Do While Not rsTmp.EOF
                strSQL = "Select 1 From Zlprograms Where NVL(系统,0) = [1] And 序号 = [2]"
                Set rsProgram = OpenSQLRecord(strSQL, "ExportReport", Val(rsTmp!系统 & ""), Val(rsTmp!程序id & ""))
                '该系统模块存在
                If Not rsProgram.EOF Then
                    '对于已有报表,只能删除对应功能,不然如票据会删除掉其它非报表的功能
                    '因为删除了功能,操作上相应角色必须重新授权
                    gcnOracle.Execute "Delete From zlProgPrivs Where Nvl(系统,0)=" & Nvl(rsTmp!系统, 0) & " And 序号=" & rsTmp!程序id & " And 功能='" & rsTmp!功能 & "'"
                    
                    gcnOracle.Execute "Insert Into zlProgFuncs(系统,序号,功能,说明) Select " & _
                        IIF(IsNull(rsTmp!系统), "NULL", rsTmp!系统) & "," & rsTmp!程序id & ",'" & rsTmp!功能 & "','" & Nvl(rsTmp!说明) & "' From Dual" & _
                        " Where Not Exists(Select 1 From zlProgFuncs Where Nvl(系统,0)=" & Nvl(rsTmp!系统, 0) & " And 序号=" & rsTmp!程序id & " And 功能='" & rsTmp!功能 & "')"
                        
                    If strObject <> "" Then
                        For i = 0 To UBound(Split(strObject, ","))
                            gcnOracle.Execute _
                                GetInsertProgPrivs(rsTmp!系统, rsTmp!程序id, rsTmp!功能 _
                                    , Split(Split(strObject, ",")(i), ".")(1) _
                                    , Split(Split(strObject, ",")(i), ".")(0) _
                                    , "SELECT")
                        Next
                    End If
                End If
                rsTmp.MoveNext
            Loop
        End If
    End If
    
    gcnOracle.CommitTrans
    blnTran = False
    
    objText.Close
    Set grsReport = Nothing '清除缓存
    '新增导入，且导入到指定分组
    If lngCurrID = 0 And lngGroupID <> 0 Then
        On Error Resume Next
        lng序号 = 1
        Set rsTmp = New ADODB.Recordset
        strSQL = "Select Count(1) Records From zlRPTSubs Where 组ID=[1]"
        Set rsTmp = OpenSQLRecord(strSQL, "报表导入", lngGroupID)
        If Not rsTmp.EOF Then
            lng序号 = Nvl(rsTmp!Records, 0) + 1
        End If
        gcnOracle.Execute "Insert Into zlRPTSubs(组ID,报表ID,序号,功能) Values(" & lngGroupID & "," & lngRPTID & "," & lng序号 & ",'" & strName & "')"
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo errH
    End If
    If lngCurrID = 0 Then
        ImportReport = lngRPTID & "|" & strNum & "|" & strName & "|" & strNote
    Else
        ImportReport = lngCurrID & "|" & rsCurr!编号 & "|" & rsCurr!名称 & "|" & IIF(IsNull(rsCurr!说明), "", rsCurr!说明)
    End If
    Exit Function
    
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    If blnOpen Then objText.Close
    If blnTran Then gcnOracle.RollbackTrans
    Call SaveErrLog
End Function

Public Function SaveWinState(objForm As Object, Optional ByVal strProjectName As String, Optional ByVal strUserDef As String) As Boolean
'功能：保存窗体及其中各种控件的状态
'参数：objForm:要保存的窗体
'      strProjectName：当前工程名，通常可用app.ProductName传递，用以区分不同工程中的同名窗体，保证恢复的正确性；
'      strUserDef：主要适用于工程中，一个窗体多个程序使用(程序使用 set frmxxx=new frm设计窗体形式)，为了按不同应用保存恢复各自的个性化状态，需要直接确定命名。
    
    Dim objThis As Object, strTmp As String
    Dim i As Integer, blnDo As Boolean
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    If Not gcnOracle Is Nothing And strProjectName <> "" And gblnRunLog Then
        If gcnOracle.State = 1 Then
            '正常退出
            On Error Resume Next
            If gstrComputerName <> "" Then
                strSQL = "Zl_Zldiarylog_Update('" & gstrComputerName & "','" & UCase(strProjectName) & _
                        "','" & UCase(objForm.name) & "',1)"
                Call ExecuteProcedure(strSQL, "更新工作日志")
            End If
            If Err.Number <> 0 Then Err.Clear
        End If
    End If
    
    On Error Resume Next
    If Not gfrmMain Is Nothing Then Call gfrmMain.Shut任务(objForm)
    On Error GoTo 0
    
    If mdlPublic.GetMemoryParam() = False Then      '使用个性化风格
        Call DelWinState(objForm, strProjectName, strUserDef)
        SaveWinState = True: Exit Function
    End If
    
    If strProjectName <> "" Then strProjectName = strProjectName & "\"
    
    '保存窗体状态、位置、大小
    With objForm
        Select Case .WindowState
            Case 0
                SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.name & strUserDef & "\Form", "状态", objForm.WindowState & "," & .Left & "," & .Top & "," & .Width & "," & .Height
            Case 1
                SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.name & strUserDef & "\Form", "状态", 0
            Case 2
                SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.name & strUserDef & "\Form", "状态", objForm.WindowState
        End Select
    End With
    
    '保存各种控件的各种状态
    For Each objThis In objForm.Controls
        strTmp = ""
        On Error Resume Next
        If UCase(TypeName(objThis)) = UCase("Menu") Then
            If objThis.Caption Like "标准按钮*" Or _
                objThis.Caption Like "文本标签*" Or _
                objThis.Caption Like "状态栏*" Or _
                UCase(objThis.name) Like UCase("mnuViewTool*") Then
                '特殊菜单的复选
                strTmp = objThis.Checked & "," & objThis.Enabled
            Else
                strTmp = ""
            End If
        ElseIf (UCase(objThis.Tag) = "SAVE" Or UCase(objThis.name) Like "*_S" Or _
            UCase(TypeName(objThis)) = UCase("StatusBar") Or _
            UCase(TypeName(objThis)) = UCase("Toolbar") Or _
            UCase(TypeName(objThis)) = UCase("Coolbar")) And objForm.Visible Then

            blnDo = True
            If UCase(TypeName(objThis)) = UCase("Toolbar") Or UCase(objThis.Tag) = "SAVE" Or UCase(objThis.name) Like "*_S" Then
                If TypeName(objThis.Container) = "PictureBox" Then blnDo = False
            End If
            'Left,Top,Width、Height,Visible
            strTmp = strTmp & "," & objThis.Left
            If Err.Number <> 0 Then Err.Clear: strTmp = strTmp & ",-32767"
            
            strTmp = strTmp & "," & objThis.Top
            If Err.Number <> 0 Then Err.Clear: strTmp = strTmp & ",-32767"
            
            strTmp = strTmp & "," & objThis.Width
            If Err.Number <> 0 Then Err.Clear: strTmp = strTmp & ",-32767"
            
            strTmp = strTmp & "," & objThis.Height
            If Err.Number <> 0 Then Err.Clear: strTmp = strTmp & ",-32767"
            
            If blnDo Then
                strTmp = strTmp & "," & objThis.Visible
                If Err.Number <> 0 Then Err.Clear: strTmp = strTmp & ",-32767"
            Else
                strTmp = strTmp & ",-32767"
            End If
            strTmp = Mid(strTmp, 2)
        End If
        If strTmp <> "" Then
            SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.name & strUserDef & "\" & TypeName(objThis), objThis.name & "状态", strTmp
        End If
        
        Select Case UCase(TypeName(objThis))
            Case UCase("Toolbar")
                If objThis.Buttons.count > 0 Then
                    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.name & strUserDef & "\" & TypeName(objThis), objThis.name & "文本", IIF(objThis.Buttons(1).Caption <> "", 1, objThis.ButtonHeight)
                End If
            Case UCase("ListView")
                SaveListViewState objThis, strProjectName & objForm.name & strUserDef
            Case UCase("CoolBar")
                strTmp = ""
                For i = 1 To objThis.Bands.count
                    strTmp = strTmp & "," & objThis.Bands(i).NewRow
                Next
                SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.name & strUserDef & "\" & TypeName(objThis), objThis.name & "行序", Mid(strTmp, 2)
                
                strTmp = ""
                For i = 1 To objThis.Bands.count
                    strTmp = strTmp & "," & objThis.Bands(i).Visible
                Next
                SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.name & strUserDef & "\" & TypeName(objThis), objThis.name & "可见栏", Mid(strTmp, 2)
            Case UCase("VSFlexGrid")
                SaveFlexState objThis, strProjectName & objForm.name & strUserDef
        End Select
    Next
    SaveWinState = True
End Function

Public Function RestoreWinState(objForm As Object, Optional ByVal strProjectName As String, Optional ByVal strUserDef As String) As Boolean
'功能：恢复窗体的状态，当左顶边界超出时，则自动设置为0
'参数：objForm:要恢复的窗体
'      strProjectName：当前工程名，通常可用app.ProductName传递，用以区分不同工程中的同名窗体，保证恢复的正确性；
'      strUserDef：主要适用于工程中，一个窗体多个程序使用(程序使用 set frmxxx=new frm设计窗体形式)，为了按不同应用保存恢复各自的个性化状态，需要直接确定命名。
   
    Dim aryInfo() As String
    Dim strTmp As String, i As Integer
    Dim objThis As Object
    Dim blnDo As Boolean
    Dim strSave As String
    Dim strOEM As String
    Dim strSQL As String
    
    If Not gcnOracle Is Nothing And strProjectName <> "" And gblnRunLog Then
        If gcnOracle.State = 1 Then
            '导常退出
'            strSQL = "Update zlDiaryLog Set 退出原因=2,退出时间=Sysdate" & _
'                    " Where 退出原因 is NULL And 会话号 Not IN(Select SID+SERIAL# From v$Session Where USER#<>0)"
'            gcnOracle.Execute strSQL

            '进入
            On Error Resume Next
            If gstrComputerName <> "" Then
                strSQL = "Zl_Zldiarylog_Insert('" & gstrComputerName & "','" & UCase(strProjectName) & "'," & _
                    " '" & UCase(objForm.name) & "','" & UCase(objForm.Caption) & "')"
                Call ExecuteProcedure(strSQL, "保存工作日志")
            End If
            If Err.Number <> 0 Then Err.Clear
        End If
    End If
    
    On Error Resume Next
    
    If Not gfrmMain Is Nothing Then Call gfrmMain.Show任务(objForm)
    
    blnDo = mdlPublic.GetMemoryParam()      '使用个性化风格
    
    If strProjectName <> "" Then strProjectName = strProjectName & "\"
    
    '恢复窗体的状态、位置、大小
    If UCase(objForm.name) = UCase("frmReport") _
        Or UCase(objForm.name) = UCase("frmPreview") _
            Or UCase(objForm.name) = UCase("frmDesign") Then
        strTmp = "2" '特殊窗体初始最大化
    Else
        strTmp = "0," & (Screen.Width - objForm.Width) / 2 & "," & (Screen.Height - objForm.Height) / 2 & "," & objForm.Width & "," & objForm.Height
    End If
    If blnDo Then
        strSave = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.name & strUserDef & "\Form", "状态", "")
        RestoreWinState = (strSave <> "")
        If strSave = "" Then strSave = strTmp
        aryInfo = Split(strSave, ",")
    Else
        aryInfo = Split(strTmp, ",")
    End If
    With objForm
        .WindowState = aryInfo(0)
        If UBound(aryInfo) = 4 Then
            .Left = IIF(aryInfo(1) < 0, 0, aryInfo(1))
            .Top = IIF(aryInfo(2) < 0, 0, aryInfo(2))
            .Width = IIF(aryInfo(3) > Screen.Width, Screen.Width, aryInfo(3))
            .Height = IIF(aryInfo(4) > Screen.Height, Screen.Height, aryInfo(4))
        Else
            .Left = (Screen.Width - objForm.Width) / 2
            .Top = (Screen.Height - objForm.Height) / 2
        End If
    End With

    '恢复窗体中各种控件的各种状态
    For Each objThis In objForm.Controls
        
        On Error Resume Next
        If blnDo Then
            strTmp = ""
            If UCase(TypeName(objThis)) = UCase("Menu") Then
                '特殊菜单的复选
                If objThis.Caption Like "标准按钮*" Or _
                    objThis.Caption Like "文本标签*" Or _
                    objThis.Caption Like "状态栏*" Or _
                    UCase(objThis.name) Like UCase("mnuViewTool*") Then
                    strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.name & strUserDef & "\" & TypeName(objThis), objThis.name & "状态", "")
                    If UBound(Split(strTmp, ",")) = 1 Then
                        objThis.Checked = Split(strTmp, ",")(0)
                        objThis.Enabled = Split(strTmp, ",")(1)
                    End If
                End If
            ElseIf UCase(objThis.Tag) = "SAVE" Or UCase(objThis.name) Like "*_S" Or _
                UCase(TypeName(objThis)) = UCase("StatusBar") Or _
                UCase(TypeName(objThis)) = UCase("Toolbar") Or _
                UCase(TypeName(objThis)) = UCase("Coolbar") Then
                
                strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.name & strUserDef & "\" & TypeName(objThis), objThis.name & "状态", "")
                If strTmp <> "" Then
                    'Left,Top,Width、Height,Visible
                    If UBound(Split(strTmp, ",")) = 4 Then
                        If Split(strTmp, ",")(0) <> "-32767" Then objThis.Left = Split(strTmp, ",")(0)
                        If Split(strTmp, ",")(1) <> "-32767" Then objThis.Top = Split(strTmp, ",")(1)
                        If Split(strTmp, ",")(2) <> "-32767" Then objThis.Width = Split(strTmp, ",")(2)
                        If Split(strTmp, ",")(3) <> "-32767" Then objThis.Height = Split(strTmp, ",")(3)
                        If Split(strTmp, ",")(4) <> "-32767" Then objThis.Visible = Split(strTmp, ",")(4)
                    End If
                End If
            End If
        End If
        
        Select Case UCase(TypeName(objThis))
            Case UCase("StatusBar")
                '状态条试用标志
'                If zlRegInfo("授权性质") <> "1" Then
'                    If objThis.Panels(1).Bevel = sbrRaised Then
'                        objThis.Panels(1).Text = ""
'                        Set objThis.Panels(1).Picture = LoadCustomPicture("Try")
'                        objThis.Panels(1).ToolTipText = ""
'                        objThis.Height = 360
'                    End If
'                Else
                    If objThis.Panels(1).Bevel = sbrRaised Then
                        strTmp = zlRegInfo("产品简名")
                        If strTmp <> "-" Then
                            objThis.Panels(1).Text = strTmp & "软件"
                            '处理状态栏图标的OEM策略
                            If strTmp = "中联" Then
                                If zlRegInfo("授权性质") <> "1" Then
                                    objThis.Panels(1).Text = ""
                                    Set objThis.Panels(1).Picture = LoadCustomPicture("Try")
                                Else
                                    Set objThis.Panels(1).Picture = LoadCustomPicture("Logo")
                                End If
                            Else
                                strOEM = GetOEM(strTmp)
                                Set objThis.Panels(1).Picture = LoadCustomPicture(strOEM)
                                If Err <> 0 Then
                                    Err.Clear
                                Set objThis.Panels(1).Picture = LoadCustomPicture("Logo")
                                End If
                                If zlRegInfo("授权性质") <> "1" Then objThis.Panels(1).Text = strTmp & "(试用)"
                            End If
                            objThis.Panels(1).ToolTipText = ""
                            objThis.Height = 360
                        End If
                    End If
'                End If
            Case UCase("Menu")
                If UCase(objThis.name) = UCase("mnuHelpWeb") Then
                    'WEB上的中联
                    strTmp = zlRegInfo("支持商简名")
                    If strTmp <> "-" Then
                        objThis.Caption = "&WEB上的" & strTmp
                    End If
                ElseIf UCase(objThis.name) = UCase("mnuHelpWebHome") Then
                    '中联主页
                    strTmp = zlRegInfo("支持商简名")
                    If strTmp <> "-" Then
                        objThis.Caption = strTmp & "主页(&H)"
                    End If
                End If
            Case UCase("Toolbar")
                If blnDo Then
                    If objThis.Buttons.count > 0 Then
                        strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.name & strUserDef & "\" & TypeName(objThis), objThis.name & "文本", 1)
                        For i = 1 To objThis.Buttons.count
                            objThis.Buttons(i).Caption = IIF(strTmp = 1, objThis.Buttons(i).Tag, "")
                        Next
                    End If
                End If
            Case UCase("ListView")
                If blnDo Then
                    RestoreListViewState objThis, strProjectName & objForm.name & strUserDef
                End If
            Case UCase("CoolBar")
                If blnDo Then
                    strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.name & strUserDef & "\" & TypeName(objThis), objThis.name & "行序", "")
                    If UBound(Split(strTmp, ",")) >= 0 Then
                        For i = 0 To UBound(Split(strTmp, ","))
                            objThis.Bands(i + 1).NewRow = Split(strTmp, ",")(i)
                        Next
                    End If
            
                    strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.name & strUserDef & "\" & TypeName(objThis), objThis.name & "可见栏", "")
                    If UBound(Split(strTmp, ",")) >= 0 Then
                        For i = 0 To UBound(Split(strTmp, ","))
                            objThis.Bands(i + 1).Visible = Split(strTmp, ",")(i)
                        Next
                    End If
                End If
            Case UCase("VSFlexGrid")
                If blnDo Then
                    RestoreFlexState objThis, strProjectName & objForm.name & strUserDef
                End If
        End Select
    Next
End Function

Public Function RestoreFlexState(objThis As Object, strForm As String) As Boolean
    Dim i As Integer, strTmp As String
        
    On Error Resume Next
    
    strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\" & TypeName(objThis), objThis.name & "宽度", "")
    If UBound(Split(strTmp, ",")) >= 0 Then
        For i = 0 To objThis.Cols - 1
            If objThis.ColWidth(i) > 0 Then
                objThis.ColWidth(i) = Split(strTmp, ",")(i)
            End If
        Next
        RestoreFlexState = True
    End If
End Function

Public Sub SaveFlexState(objThis As Object, strForm As String)
    Dim strTmp As String, i As Integer
        
    On Error Resume Next
    
    strTmp = ""
    For i = 0 To objThis.Cols - 1
        strTmp = strTmp & "," & objThis.ColWidth(i)
    Next
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\" & TypeName(objThis), objThis.name & "宽度", Mid(strTmp, 2)
End Sub

Public Sub SaveListViewState(objLvw As Object, ByVal strForm As String, Optional strIndex As String)
'功能：保存ListView的各种特性
'参数：objLvw=ListView对象,strForm=窗体关键字
'说明：视图方式、列宽、列位置、列标题、列对齐、排序
    Dim lngCol As Long
    Dim strWidth As String
    Dim strPosition As String
    Dim strText As String
    Dim strAlign As String
    
    For lngCol = 1 To objLvw.ColumnHeaders.count
        strWidth = strWidth & "," & objLvw.ColumnHeaders(lngCol).Width
        strPosition = strPosition & "," & objLvw.ColumnHeaders(lngCol).Position
        strText = strText & "," & objLvw.ColumnHeaders(lngCol).Text
        strAlign = strAlign & "," & objLvw.ColumnHeaders(lngCol).Alignment
    Next
    
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\ListView", objLvw.name & strIndex & "视图", objLvw.View
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\ListView", objLvw.name & strIndex & "宽度", Mid(strWidth, 2)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\ListView", objLvw.name & strIndex & "位置", Mid(strPosition, 2)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\ListView", objLvw.name & strIndex & "名称", Mid(strText, 2)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\ListView", objLvw.name & strIndex & "对齐", Mid(strAlign, 2)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\ListView", objLvw.name & strIndex & "排序", objLvw.SortKey & "," & objLvw.SortOrder & "," & objLvw.Sorted
End Sub

Public Sub RestoreListViewState(objLvw As Object, ByVal strForm As String, Optional strIndex As String)
'功能：恢复ListView的各种特性
'参数：objLvw=ListView对象,strForm=窗体关键字
'说明：视图方式、列宽、列位置、列标题、列对齐、排序
    Dim lngCol As Long
    Dim strWidth As String
    Dim strPosition As String
    Dim strText As String, arrText As Variant
    Dim strAlign As String
    Dim strSort As String
    
    On Error Resume Next
    
    strText = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\ListView", objLvw.name & strIndex & "名称")
    If strText <> "" Then
        arrText = Split(strText, ",")
        If objLvw.Tag = "可变化的" Then
            '只对需要的ListView进行列变化
            objLvw.ColumnHeaders.Clear
            For lngCol = LBound(arrText) To UBound(arrText)
                '列缺省关键字为"_" & 列标题
                objLvw.ColumnHeaders.Add , "_" & arrText(lngCol), arrText(lngCol)
            Next
        Else
            '检查是否需要恢复
            '列数变了,不恢复而使用缺省
            If UBound(arrText) + 1 <> objLvw.ColumnHeaders.count Then Exit Sub
            '列标题变了,不恢复而使用缺省
            For lngCol = 1 To objLvw.ColumnHeaders.count
                If objLvw.ColumnHeaders(lngCol).Text <> arrText(lngCol - 1) Then Exit Sub
            Next
        End If
    End If
    
    '视图缺省保持初始值
    lngCol = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\ListView", objLvw.name & strIndex & "视图", -1)
    If lngCol <> -1 Then objLvw.View = lngCol
    
    '列的宽度,顺序,对齐
    strWidth = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\ListView", objLvw.name & strIndex & "宽度")
    strPosition = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\ListView", objLvw.name & strIndex & "位置")
    strAlign = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\ListView", objLvw.name & strIndex & "对齐")
    For lngCol = 1 To objLvw.ColumnHeaders.count
        '列缺省关键字为"_" & 列标题
        objLvw.ColumnHeaders(lngCol).Key = "_" & objLvw.ColumnHeaders(lngCol).Text
        If strWidth <> "" Then objLvw.ColumnHeaders(lngCol).Width = Split(strWidth, ",")(lngCol - 1)
        If strAlign <> "" Then objLvw.ColumnHeaders(lngCol).Alignment = Split(strAlign, ",")(lngCol - 1)
    Next
    For lngCol = objLvw.ColumnHeaders.count To 1 Step -1
        If strPosition <> "" Then objLvw.ColumnHeaders(lngCol).Position = Split(strPosition, ",")(lngCol - 1)
    Next
    
    '排序特性
    strSort = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\ListView", objLvw.name & strIndex & "排序")
    If strSort <> "" Then
        objLvw.SortKey = Split(strSort, ",")(0)
        objLvw.SortOrder = Split(strSort, ",")(1)
        objLvw.Sorted = Split(strSort, ",")(2)
    End If
End Sub

Public Function DelWinState(objForm As Object, Optional ByVal strProjectName As String, Optional ByVal strUserDef As String) As Boolean
'功能：删除窗体个性化设置值
'参数：objForm:要恢复的窗体
'      strProjectName：当前工程名，通常可用app.ProductName传递，用以区分不同工程中的同名窗体，保证恢复的正确性；
'      strUserDef：主要适用于工程中，一个窗体多个程序使用(程序使用 set frmxxx=new frm设计窗体形式)，为了按不同应用保存恢复各自的个性化状态，需要直接确定命名。
    Dim strProject As String
    Dim lngR As Long
    Dim objThis As Object
    
    strProject = strProjectName
    If strProjectName <> "" Then strProjectName = strProjectName & "\"
    
    For Each objThis In objForm.Controls
        lngR = RegDeleteKey(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\ZLSOFT\私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.name & strUserDef & "\" & TypeName(objThis) & Chr(0))
        If lngR <> 0 And lngR <> 2 Then Exit Function
    Next
    
    lngR = RegDeleteKey(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\ZLSOFT\私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.name & strUserDef & "\Form" & Chr(0))
    If lngR <> 0 And lngR <> 2 Then Exit Function
    lngR = RegDeleteKey(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\ZLSOFT\私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.name & strUserDef & Chr(0))
    If lngR <> 0 And lngR <> 2 Then Exit Function
    
    DelWinState = True
End Function

Public Function LoadCustomPicture(strID As String, Optional strFormat As String = "GIF") As StdPicture
'功能:将资源文件中的指定资源生成磁盘文件
'参数:ID=资源号,strExt=要生成文件的扩展名(如BMP)
'返回:生成文件名
    Dim arrData() As Byte
    Dim intFile As Integer
    Dim strFile As String * 255, strR As String
    
    arrData = LoadResData(strID, strFormat)
    intFile = FreeFile
    
    GetTempPath 255, strFile
    strR = Trim(Left(strFile, InStr(strFile, Chr(0)) - 1)) & CLng(timer * 100) & ".pic"

    Open strR For Binary As intFile
    Put intFile, , arrData()
    Close intFile
    Set LoadCustomPicture = VB.LoadPicture(strR)
    Kill strR
End Function

Public Function GetImage(objFld As ADODB.Field) As StdPicture
'功能：将指定字段中的二进制数据生成一个磁盘文件
'返回：图形对象，或内容为空的初始化了的图片
    Dim lngFileSize As Long
    Dim intFile As Integer
    Dim arrData() As Byte
    Dim strFile As String
    
    On Local Error GoTo errH
    
    If IsNull(objFld.Value) Then Exit Function
    
    lngFileSize = objFld.ActualSize
    If lngFileSize = 0 Then Exit Function
    ReDim arrData(lngFileSize - 1) As Byte
    
    intFile = FreeFile
    strFile = CurDir & "\tmp" & Int(timer * 100) & ".pic"
    Open strFile For Binary As intFile
    arrData() = objFld.GetChunk(lngFileSize)
    Put intFile, , arrData()
    Close intFile
    
    Set GetImage = VB.LoadPicture(strFile)
    Kill strFile
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function SaveImage(objGraph As StdPicture, objFld As Field) As Boolean
'功能：将指定图形存放到指定的记录集字段中
'说明：不更新记录集
    Dim intFile As Integer, strFile As String
    Dim arrData() As Byte
    
    If objGraph Is Nothing Then SaveImage = True: Exit Function
    
    On Local Error GoTo errH
    
    strFile = CurDir & "\tmp" & Int(timer * 100) & ".pic"
    Call VB.SavePicture(objGraph, strFile)
    
    intFile = FreeFile
    Open strFile For Binary Access Read As intFile
    ReDim arrData(LOF(intFile) - 1) As Byte
    Get intFile, , arrData()
    Close intFile
    Kill strFile
    
    objFld.AppendChunk arrData()
    SaveImage = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function DataUsed(objReport As Report, strData As String, Optional blnFormat As Boolean) As Boolean
'功能：判断指定数据源在指定报表格式中是否使用
'参数：strData=数据源名
'      blnFormat=是否只在报表当前的格式中判断(缺省为不,在所有格式中判断)
'说明：标签包含任意表头中的标签
    Dim tmpItem As RPTItem, tmpPar As RPTPar
    Dim strContent As String
    
    For Each tmpItem In objReport.Items
        '有分类表格必定有分类子项
        If (blnFormat And tmpItem.格式号 = objReport.bytFormat Or Not blnFormat) _
            And InStr("2,3,5,6,12,13,14,", tmpItem.类型 & ",") > 0 Then
            Select Case tmpItem.类型
                Case 2, 3, 13 '数据标签,"病人信息.姓名"
                    If InStr(tmpItem.内容, strData & ".") > 0 Then DataUsed = True: Exit Function
                Case 5 '分类表格,"病人信息"
                    strContent = tmpItem.内容
                    If strContent Like "*（*）" Then
                        strContent = Left(strContent, InStrRev(strContent, "（") - 1)
                    End If
                    If strContent = strData Then DataUsed = True: Exit For
                Case 6 '任意表格子项,"([病人信息.身高]+[2])/3"
                    If InStr(tmpItem.内容, strData & ".") > 0 Then DataUsed = True: Exit Function
                    If InStr(tmpItem.表头, strData & ".") > 0 Then DataUsed = True: Exit Function
                Case 12 '图表
                    If InStr("|" & tmpItem.内容, "|" & strData & ".") > 0 Then DataUsed = True: Exit Function
                Case 14
                    If tmpItem.数据源 = strData Then DataUsed = True: Exit Function
            End Select
        End If
    Next
End Function

Public Function MakeNamePars(objReport As Report, Optional blnFirst As Boolean) As RPTPars
'功能：从报表(objReport)所有数据源中产生参数名称唯一的参数集
'参数：blnFirst=强制取当前已有效的缺省值,不是为了重置条件
'说明：1.设计程序已限制同一报表中不同数据源之间的参数如果同名,则类型,缺省值也相同
    Dim tmpData As RPTData, tmpPar As RPTPar, StrPar As String
    Dim tmpPars As New RPTPars, strTmp As String
    
    For Each tmpData In objReport.Datas
        If DataUsed(objReport, tmpData.名称) Then
            For Each tmpPar In tmpData.Pars
                If InStr(StrPar & ",", "," & tmpPar.名称 & ",") = 0 Then
                    StrPar = StrPar & "," & tmpPar.名称
                    With tmpPar '！！以名称(唯一)关键字加入
                        If .Reserve Like "*…|*" And Not blnFirst Then
                            '当条件重置时，Reserve记录了"宏条件值|显示值"
                            '处理为缺省值为定义时的宏缺省值,Reserve为"显示值|绑定值"
                            tmpPars.Add .组名, .序号, .名称, .类型, CStr(Split(.Reserve, "|")(0)), .格式, .值列表, .分类SQL _
                                , .明细SQL, .分类字段, .明细字段, .对象, "_" & .名称, Split(.Reserve, "|")(1) & "|" & .缺省值 _
                                , .是否锁定
                        Else
                            '第一次进入(非重置条件)或其它类型的条件
                            tmpPars.Add .组名, .序号, .名称, .类型, .缺省值, .格式, .值列表, .分类SQL, .明细SQL, .分类字段 _
                                , .明细字段, .对象, "_" & .名称, .Reserve, .是否锁定
                        End If
                    End With
                End If
            Next
        End If
    Next
    Set MakeNamePars = tmpPars
End Function

Public Sub ItemAutoSize(objItem As RPTItem, ByVal strValue As String, ByVal objCalc As Object)
'功能：根据报表标签元素的内容自动调整其宽高
'参数：objCalc=用于计算实际宽高的对象
'说明：1.只改变其W,H,不改变其内容。
'      2.因为标签是循环取值,因此查询或预览每次都要调整。
    If Not objItem.自调 Then Exit Sub
    objCalc.Font.name = objItem.字体
    objCalc.Font.Size = objItem.字号
    objCalc.Font.Bold = objItem.粗体
    objCalc.Font.Italic = objItem.斜体
    objCalc.Font.Underline = objItem.下线
    
    objItem.W = objCalc.TextWidth(strValue) + objCalc.TextWidth("A")
    objItem.H = objCalc.TextHeight(strValue) + 30
End Sub

Public Function ReplaceBracket(ByVal strValue As String, Optional ByVal strReplace As String) As String
'功能：将字符串中的[]替换为指定的值
    Dim strLeft As String, strRight As String, strVar As String
    
    '[]用于Like时无效,所以要替换
    strVar = Replace(strValue, "[", "@@")
    strVar = Replace(strVar, "]", "$$")
    If Not strVar Like "*@@*$$*" Then ReplaceBracket = strValue: Exit Function
    
    Do While InStr(strValue, "[") > 0
        strLeft = Left(strValue, InStr(strValue, "[") - 1)
        strRight = Mid(strValue, InStr(strValue, "]") + 1)
        strVar = Mid(strValue, InStr(strValue, "[") + 1, InStr(strValue, "]") - InStr(strValue, "[") - 1)
            
        strValue = strLeft & strReplace & strRight
    Loop
    
    strValue = Replace(strValue, "@@", "[")
    strValue = Replace(strValue, "$$", "]")
    ReplaceBracket = strValue
End Function

Public Function GetLabelMacro(frmParent As Object, ByVal strValue As String) As String
'功能：处理标签中的宏:[n>=0],[=参数名]
'说明：不处理[页号][页数]
    Dim strLeft As String, strRight As String, strVar As String
    
    '[]用于Like时无效,所以要替换
    strVar = Replace(strValue, "[", "@@")
    strVar = Replace(strVar, "]", "$$")
    If Not strVar Like "*@@*$$*" Then GetLabelMacro = strValue: Exit Function
    If strVar Like "*@@*.*$$*" Then GetLabelMacro = strValue: Exit Function
    
    Do While InStr(strValue, "[") > 0
        strLeft = Left(strValue, InStr(strValue, "[") - 1)
        strRight = Mid(strValue, InStr(strValue, "]") + 1)
        strVar = Mid(strValue, InStr(strValue, "[") + 1, InStr(strValue, "]") - InStr(strValue, "[") - 1)
            
        If IsNumeric(strVar) Then '参数数据
            If CInt(strVar) >= 0 Then strVar = GetUserParData(frmParent, CInt(strVar))
        ElseIf Left(strVar, 1) = "=" Then '[=参数名]
            If Mid(strVar, 2) <> "" Then strVar = GetParValue(frmParent, Mid(strVar, 2))
        ElseIf strVar = "单位名称" Then
            strVar = Replace(zlRegInfo("单位名称", , -1), ";", vbCrLf)
        ElseIf strVar = "操作员姓名" Then
            strVar = gstrUserName
        ElseIf strVar = "操作员编号" Then
            strVar = gstrUserNO
        ElseIf IsDate(Format("2000-01-01", strVar)) Then '当前日期
            strVar = Format(Currentdate, strVar)
        Else
            strVar = "@@" & strVar & "$$"
        End If
        strValue = strLeft & strVar & strRight
    Loop
    
    strValue = Replace(strValue, "@@", "[")
    strValue = Replace(strValue, "$$", "]")
    GetLabelMacro = strValue
End Function

Public Function GetLabelDataName(ByVal strValue As String) As String
'功能：获取标签中所包含的数据字段名.
'返回：格式"数据源.字段|数据源.字段|..."
    Dim strLeft As String, strRight As String, strVar As String
    
    If Not BracketMatch(strValue, "[]") Then Exit Function
    
    '[]用于Like时无效,所以要替换
    strVar = Replace(strValue, "[", "@@")
    strVar = Replace(strVar, "]", "$$")
    If Not strVar Like "*@@*.*$$*" Then Exit Function
    
    Do While InStr(strValue, "[") > 0
        strLeft = Left(strValue, InStr(strValue, "[") - 1)
        strRight = Mid(strValue, InStr(strValue, "]") + 1)
        strVar = Mid(strValue, InStr(strValue, "[") + 1, InStr(strValue, "]") - InStr(strValue, "[") - 1)
            
        If InStr(strVar, ".") > 0 Then
            GetLabelDataName = GetLabelDataName & "|" & strVar
        End If
        strValue = strLeft & strVar & strRight
    Loop
    GetLabelDataName = Mid(GetLabelDataName, 2)
End Function

Public Function BracketMatch(ByVal strText As String, ByVal strBracket As String, Optional ByVal blnNesting As Boolean) As Boolean
'功能：检查指定字符串中指定的括号是否匹配
'参数：strText=要检查的字符串
'      strBracket=括号对，如"[]"
'      blnNesting=括号是否允许嵌套,即"[..[...]..]"形式
    Dim lngLeft As Long, lngRight As Long
    Dim strLast As String, i As Long
    
    If strText = "" Or Len(strBracket) <> 2 Then BracketMatch = True: Exit Function
    For i = 1 To Len(strText)
        If Mid(strText, i, 1) = Left(strBracket, 1) Then
            If Left(strBracket, 1) = strLast And Not blnNesting Then Exit Function
            lngLeft = lngLeft + 1
            strLast = Left(strBracket, 1)
        ElseIf Mid(strText, i, 1) = Right(strBracket, 1) Then
            If Right(strBracket, 1) = strLast And Not blnNesting Then Exit Function
            lngRight = lngRight + 1
            strLast = Right(strBracket, 1)
        End If
    Next
    BracketMatch = lngLeft = lngRight
End Function

Public Function GetHeadCellScript(frmSource As Object, objItem As RPTItem, R As Long, C As Long) As String
'功能：获取指定任意表指定行列的标签描述
    Dim tmpID As RelatID, strTmp As String
    
    For Each tmpID In objItem.SubIDs
        If frmSource.mobjReport.Items("_" & tmpID.id).序号 = C Then
            strTmp = frmSource.mobjReport.Items("_" & tmpID.id).表头
            strTmp = CStr(Split(Split(strTmp, "|")(R), "^")(2))
            GetHeadCellScript = strTmp
            Exit Function
        End If
    Next
End Function

Public Function GetGridStyle(objReport As Report, id As Integer) As Byte
'功能：判断任意表格的样式
'返回：0:表头的表体皆有效,1-仅表头有效,2-仅表体有效
'说明：如果表头表体皆无效，则返回两者都有效
    Dim i As Integer, tmpID As RelatID
    Dim blnBody As Boolean, blnHead As Boolean
    Dim strTmp As String
    
    If objReport.Items("_" & id).类型 <> 4 Then Exit Function
    
    blnHead = False
    blnBody = False
    For Each tmpID In objReport.Items("_" & id).SubIDs
        strTmp = objReport.Items("_" & tmpID.id).表头
        i = UBound(Split(strTmp, "|"))
        If i > 0 Then
            blnHead = True
        ElseIf i = 0 Then
            blnHead = blnHead Or (Split(Split(strTmp, "|")(i), "^")(2) <> "#")
        End If
        blnBody = blnBody Or (objReport.Items("_" & tmpID.id).内容 <> "")
        
        If blnBody And blnHead Then
            Exit For
        End If
    Next
    If blnHead And blnBody Then
        GetGridStyle = 0
    ElseIf blnHead Then
        GetGridStyle = 1
    ElseIf blnBody Then
        GetGridStyle = 2
    Else
        GetGridStyle = 0
    End If
End Function

Public Function SaveFile(strFile As String, objFld As Field) As Boolean
'功能：将指定文件存放到指定的记录集字段中
'说明：不更新记录集
    Dim intFile As Integer
    Dim arrData() As Byte
    
    On Local Error GoTo errH
    
    intFile = FreeFile
    Open strFile For Binary Access Read As intFile
    ReDim arrData(LOF(intFile) - 1) As Byte
    Get intFile, , arrData()
    Close intFile
    
    objFld.AppendChunk arrData()
    SaveFile = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetDependIDs(strName As String, frmSource As Object)
'功能：获取标签所参照的表格ID
'参数：strName=所参照的表格名
'说明：如果参照表格包含附加表格,则附加表格的ID也一并返回
    Dim objItem As RPTItem
    Dim strIDs As String
    
    For Each objItem In frmSource.mobjReport.Items
        If objItem.格式号 = frmSource.bytFormat And (objItem.类型 = 4 Or objItem.类型 = 5) _
            And ((objItem.性质 = 0 And objItem.名称 = strName) _
                Or (objItem.性质 = 1 And objItem.参照 = strName)) Then
            strIDs = strIDs & "," & objItem.id
        End If
    Next
    GetDependIDs = Mid(strIDs, 2)
End Function

Public Function GetRightWidth(lngCol As Long, lngEnd As Long, lngRow As Long, strSkip As String, strSkip2 As String, objGrid As Object) As Long
'功能：获取当前输出页中指定列右边还需要输出列的宽度
'参数：lngCol=当前输出列
'      lngEnd=当前页输出结束列
'      lngRow=当前输出行
'      strSkip=包含跳过输出单元格的串
'      objGrid=表格
    Dim i As Long, W As Long
    
    For i = lngCol + 1 To lngEnd
        If InStr(strSkip, "[" & lngRow & "," & i & "]") = 0 Or _
            (InStr(strSkip, "[" & lngRow & "," & i & "]") > 0 And _
            InStr(strSkip2, "[(" & lngRow & "," & lngCol & ")," & lngRow & "," & i & "]") = 0) Then
            W = W + objGrid.ColWidth(i)
        End If
    Next
    GetRightWidth = W
End Function

Private Function GetSubItem(frmSource As Object, objItem As RPTItem, ByVal intCol As Integer) As RPTItem
    Dim tmpID As RelatID, tmpItem As RPTItem
    
    For Each tmpID In objItem.SubIDs
        If frmSource.mobjReport.Items("_" & tmpID.id).序号 = intCol Then
            Set GetSubItem = frmSource.mobjReport.Items("_" & tmpID.id): Exit Function
        End If
    Next
End Function

Public Function PrintPage(ByVal intPage As Integer, objOut As Object, frmSource As Object, _
    Optional ByVal sngScale As Single = 1, Optional ByVal blnSure As Boolean = True, _
    Optional ByVal blnMeasure As Boolean, Optional lngMaxH As Long) As Boolean
'功能：打印(预览)一页
'参数：intPage=要输出的页号,>=0
'      objOut=输出对象,Printer或PictureBox
'      frmSource=包含数据表格的窗体(即frmReport)
'      sngScale=输出比例,Printer只能为1
'      blnSure=是否按实际可打印区域预览,缺省为True,输出到打印机时固定为False
'      blnMeasure=是否仅测试实际需要打印的纸张高度,不输出内容
'      lngMaxH=配合blnMeasure参数使用,返回测量出来的最大纸张高度(Twip)
'说明：该函数不处理换页或结束输出
'参数：frmSource.mobjReport,marrPage,mLibDatas
    Dim objFmt As RPTFmt, objSub As RPTItem, objTemp As RPTItem
    Dim arrPage As Variant, objItem As RPTItem, objPageCell As PageCell
    Dim lngCurH As Long, lngPaperW As Long, lngPaperH As Long '纸张
    Dim objBody As Object, objHead As Object, objFont As New StdFont
    Dim strValue As String, strDepend As String, objPic As StdPicture
    Dim strSkip As String, strSkip2 As String '2包含更信息的信息
    Dim arrPars As Variant, blnPressWork As Boolean  '是否套打
    Dim intBasePage As Integer, colRowIDs As Collection, objCurDLL As clsReport
    
    Dim lngPreRow As Long, lngPreCol As Long, blnHaveGrid As Boolean
    Dim LngRows As Long, lngRowB As Long, lngRowE As Long
    Dim X As Long, Y As Long, W As Long, H As Long '这些是相对尺寸
    Dim i As Long, j As Long, k As Long, l As Long, M As Long
    Dim B As Long '当前页真正有效输出的第一列
    Dim lngindex As Long, lngSize As Long, sngWidth As Single
    Dim lngChildX As Long, lngChildY As Long
    Dim arrPageCard As Variant, objPageCard As PageCard
    Dim lngX As Long, lngY As Long, lngCol As Long, lngRow As Long
    Dim lngRowHeight As Long, lngRowCount As Long, lngRowTotal As Long
    Dim lngRangeHeight As Long
    
    Dim dblSureW As Double, dblSureH As Double
    Dim colColAutoFont As Collection
    Dim strData As String, strTmp As String, strBdr As String
    '标签参照的表格当前页输出位置、尺寸
    Dim lngOX As Long, lngOY As Long, lngOW As Long, lngOH As Long
    '标签实际输出位置
    Dim lngOutX As Long, lngOutY As Long, lngDesignH As Long
    Dim blnGroup As Boolean, blnWithData As Boolean
    Dim lngForeColor As Long, lngBackColor As Long
    Dim blnPrint As Boolean                                                     '用于判断打印还是预览
    Dim blnAddition As Boolean
    
    '通过传入类型来进行判断
    blnPrint = (TypeName(objOut) = "Printer")
    
    lngCurH = 0: lngMaxH = 0
    lngindex = -1
    
    If TypeName(objOut) = "Printer" Then
        sngScale = 1
        blnSure = False
    End If
    
    arrPage = frmSource.marrPage
    arrPageCard = frmSource.marrPageCard
        
    Set objFmt = frmSource.mobjReport.Fmts("_" & frmSource.mobjReport.bytFormat)
    If objFmt.纸向 = 1 Then
        lngPaperW = objFmt.W: lngPaperH = objFmt.H
    Else
        lngPaperW = objFmt.H: lngPaperH = objFmt.W
    End If
    
    intBasePage = 1
    arrPars = frmSource.marrPars '直接访问要出属性Get错误
    If Not blnMeasure And UBound(arrPars) <> -1 Then
        For i = 0 To UBound(arrPars)
            j = InStr(CStr(arrPars(i)), "=")
            If j > 0 Then
                If UCase(Trim(Left(CStr(arrPars(i)), j - 1))) = UCase("PressWork") Then
                    '根据用户传入参数判断是否套打:测试纸张时不处理
                    If IsNumeric(Trim(Mid(CStr(arrPars(i)), j + 1))) Then
                        blnPressWork = Val(Trim(Mid(CStr(arrPars(i)), j + 1))) = 1 '全部套打
                    End If
                ElseIf UCase(Trim(Left(CStr(arrPars(i)), j - 1))) = UCase("PressWorkFirst") Then
                    If IsNumeric(Trim(Mid(CStr(arrPars(i)), j + 1))) Then
                        blnPressWork = Val(Trim(Mid(CStr(arrPars(i)), j + 1))) = 1 And intPage = 0 '首页套打
                    End If
                ElseIf UCase(Trim(Left(CStr(arrPars(i)), j - 1))) = UCase("StartPageNum") Then
                    If IsNumeric(Trim(Mid(CStr(arrPars(i)), j + 1))) Then
                        intBasePage = Val(Trim(Mid(CStr(arrPars(i)), j + 1))) '起始打印页号
                        If intBasePage = 0 Then intBasePage = 1
                    End If
                End If
            End If
        Next
    End If
    Set colRowIDs = frmSource.mcolRowIDs
    Set objCurDLL = frmSource.mobjCurDLL '直接访问说不支持属性和方法
    
    '输出表格内容
    If IsArray(arrPage) Then
        If UBound(arrPage) >= intPage Then
            If arrPage(intPage).count > 0 Then
                blnHaveGrid = True
                '循环处理当前页内的多个表格
                For Each objPageCell In arrPage(intPage)
                    
                    With objPageCell
                        Set objBody = frmSource.msh(.id)
                        Set objItem = frmSource.mobjReport.Items("_" & .id)
                        
                        '处理指定了表格单元格的标签元素
                        Call SetCellValue(IIF(blnSure, 1, 2), frmSource, objItem, .RowB)
                        
                        '自动字体设置属性的缓存
                        Set colColAutoFont = New Collection
                        For i = 0 To objBody.Cols - 1
                            colColAutoFont.Add "", "_" & i '为""表示该列尚未处理
                        Next
                        
                        objBody.Redraw = False
                        lngPreRow = objBody.Row: lngPreCol = objBody.Col
                        
                        If objItem.类型 = 4 Then
                            Set objHead = frmSource.msh(objBody.Tag)
                            objHead.Redraw = False
                        End If
                        
                        '处理固定行列交叉部份(仅分类汇总表有)
                        If .FixH > 0 And .FixW > 0 Then
                            strSkip = "": strSkip2 = "": Y = 0
                            For i = 0 To objBody.FixedRows - 1
                                If Not blnMeasure And Not blnPressWork Then
                                    objBody.Row = i: X = 0
                                    For j = 0 To objBody.FixedCols - 1
                                        objBody.Col = j
                                        If InStr(strSkip, "[" & i & "," & j & "]") = 0 Then
                                            SearchCell objBody, i, j, objBody.FixedRows - 1, objBody.FixedCols - 1, W, H, strSkip, strSkip2
                                            
                                            strBdr = "1111"
                                            If Not objItem.边框 And j = 0 Then strBdr = "1101"

                                            Set objFont = objBody.Font
                                            If objBody.Cell(flexcpFontBold, i, j) = True Then
                                                objFont.Bold = True
                                            Else
                                                objFont.Bold = False
                                            End If
                                            lngForeColor = IIF(objBody.Cell(flexcpForeColor, i, j) = &HFF0001 _
                                                                    And objBody.Cell(flexcpFontUnderline, i, j) = True _
                                                                , objBody.ForeColor _
                                                                , objBody.Cell(flexcpForeColor, i, j))
                                            '汇总表格-表头
                                            If Not DrawCell(objOut, objBody.Text, .X + X, .Y + Y, W, H, .X + .W, _
                                                        , objBody.GridColor, lngForeColor, objBody.BackColor _
                                                        , objFont, strBdr _
                                                        , GetHscAlign(objBody.Cell(flexcpAlignment, i, j), objBody.Text) _
                                                        , GetVscAlign(objBody.Cell(flexcpAlignment, i, j)) _
                                                        , True, sngScale, , objItem.表格线加粗) Then
                                                Exit Function '合并时允许换行
                                            End If
                                        End If
                                        X = X + objBody.ColWidth(j)
                                    Next
                                End If
                                
                                Y = Y + objBody.RowHeight(i)
                                
                                If blnMeasure Then
                                    lngCurH = .Y + Y
                                    If lngCurH > lngMaxH Then lngMaxH = lngCurH
                                End If
                            Next
                        End If
                                                                
                        '处理固定列部份(仅分类表格有)
                        If .FixW > 0 Then
                            strSkip = "": strSkip2 = "": Y = .FixH
                            For i = .RowB To .RowE
                                If Not blnMeasure And Not blnPressWork Then
                                    objBody.Row = i: X = 0
                                    For j = 0 To objBody.FixedCols - 1
                                        objBody.Col = j
                                        If InStr(strSkip, "[" & i & "," & j & "]") = 0 Then
                                            SearchCell objBody, i, j, .RowE, objBody.FixedCols - 1, W, H, strSkip, strSkip2
                                            
                                            strBdr = "1111"
                                            If Not objItem.边框 And j = 0 Then strBdr = "1101"
                                            Set objFont = objBody.Font
                                            If objBody.Cell(flexcpFontBold, i, j) = True Then
                                                objFont.Bold = True
                                            Else
                                                objFont.Bold = False
                                            End If
                                            lngForeColor = IIF(objBody.Cell(flexcpForeColor, i, j) = &HFF0001 _
                                                                    And objBody.Cell(flexcpFontUnderline, i, j) = True _
                                                                , objBody.ForeColor _
                                                                , objBody.Cell(flexcpForeColor, i, j))
                                            '汇总表格-表体
                                            If Not DrawCell(objOut, objBody.Text, .X + X, .Y + Y, W, H, .X + .W, _
                                                        , objBody.GridColor, lngForeColor, objBody.BackColor _
                                                        , objFont, strBdr _
                                                        , GetHscAlign(objBody.Cell(flexcpAlignment, i, j), objBody.Text, objBody, j) _
                                                        , GetVscAlign(objBody.Cell(flexcpAlignment, i, j)) _
                                                        , True, sngScale, , objItem.表格线加粗) Then
                                                Exit Function '合并时允许换行
                                            End If
                                        End If
                                        X = X + objBody.ColWidth(j)
                                    Next
                                End If
                                
                                Y = Y + objBody.RowHeight(i)
                                
                                If blnMeasure Then
                                    lngCurH = .Y + Y
                                    If lngCurH > lngMaxH Then lngMaxH = lngCurH
                                End If
                            Next
                        End If
                        
                        '处理固定行部份(都有)
                        If objItem.类型 = 5 Then
                            strSkip = "": strSkip2 = "": Y = 0
                            For i = 0 To objBody.FixedRows - 1
                                If Not blnMeasure And Not blnPressWork Then
                                    objBody.Row = i: X = .FixW
                                    For j = .ColB To .ColE
                                        objBody.Col = j
                                        If InStr(strSkip, "[" & i & "," & j & "]") = 0 Then
                                            SearchCell objBody, i, j, objBody.FixedRows - 1, .ColE, W, H, strSkip, strSkip2
                                            
                                            strBdr = "1111"
                                            If Not objItem.边框 And (j = .ColE Or (W > objBody.ColWidth(j) + 15 _
                                                And Right(strSkip, Len("[" & i & "," & .ColE & "]")) = "[" & i & "," & .ColE & "]")) Then strBdr = "1110"
                                            Set objFont = objBody.Font
                                            If objBody.Cell(flexcpFontBold, i, j) = True Then
                                                objFont.Bold = True
                                            Else
                                                objFont.Bold = False
                                            End If
                                            lngForeColor = IIF(objBody.Cell(flexcpForeColor, i, j) = &HFF0001 _
                                                                    And objBody.Cell(flexcpFontUnderline, i, j) = True _
                                                                , objBody.ForeColor _
                                                                , objBody.Cell(flexcpForeColor, i, j))
                                            '汇总表格-表头
                                            If Not DrawCell(objOut, objBody.Text, .X + X, .Y + Y, W, H, .X + .W, _
                                                        , objBody.GridColor, lngForeColor, objBody.BackColor _
                                                        , objFont, strBdr _
                                                        , GetHscAlign(objBody.Cell(flexcpAlignment, i, j), objBody.Text) _
                                                        , GetVscAlign(objBody.Cell(flexcpAlignment, i, j)) _
                                                        , True, sngScale, , objItem.表格线加粗) Then
                                                Exit Function  '合并时允许换行
                                            End If
                                        End If
                                        X = X + objBody.ColWidth(j)
                                    Next
                                End If
                                
                                Y = Y + objBody.RowHeight(i)
                                
                                If blnMeasure Then
                                    lngCurH = .Y + Y
                                    If lngCurH > lngMaxH Then lngMaxH = lngCurH
                                End If
                            Next
                        ElseIf .FixH > 0 Then
                            For k = 1 To .Copys '处理任意表格表头自动分栏
                                strSkip = "": strSkip2 = "": Y = 0
                                For i = 0 To objHead.FixedRows - 1
                                    If Not blnMeasure And Not blnPressWork Then
                                        objHead.Row = i: X = .W * (k - 1) '！！
                                        B = 0
                                        For j = .ColB To .ColE
                                            objHead.Col = j
                                            If objHead.Text = "删除线" Then lngindex = j
                                            If InStr(strSkip, "[" & i & "," & j & "]") = 0 Then
                                                '表头单元格原始数据定义
                                                strValue = GetHeadCellScript(frmSource, objItem, i, j)
                                                If strValue = "#" Then '为空
                                                    strValue = ""
                                                ElseIf strValue = "←" Then '与左边单元格相同
                                                    For l = j - 1 To 0 Step -1
                                                        strValue = GetHeadCellScript(frmSource, objItem, i, l)
                                                        If strValue <> "←" Then Exit For
                                                    Next
                                                ElseIf strValue = "↑" Then '与上边单元格相同
                                                    For l = i - 1 To 0 Step -1
                                                        strValue = GetHeadCellScript(frmSource, objItem, l, j)
                                                        If strValue <> "↑" Then Exit For
                                                    Next
                                                End If
                                                
                                                '处理页变量
                                                If InStr(strValue, "[页号]") > 0 Then
                                                    strValue = Replace(strValue, "[页号]", intPage + intBasePage)
                                                End If
                                                If InStr(strValue, "[页数]") > 0 Then
                                                    If Not IsArray(arrPage) Then
                                                        strValue = Replace(strValue, "[页数]", intBasePage)
                                                    Else
                                                        strValue = Replace(strValue, "[页数]", UBound(arrPage) + intBasePage)
                                                    End If
                                                End If
                                                If InStr(strValue, "[票据号]") > 0 Then
                                                    If IsArray(garrBill) Then
                                                        If UBound(garrBill) >= intPage Then
                                                            strValue = Replace(strValue, "[票据号]", garrBill(intPage))
                                                        Else
                                                            strValue = Replace(strValue, "[票据号]", "")
                                                        End If
                                                    Else
                                                        strValue = Replace(strValue, "[票据号]", "")
                                                    End If
                                                End If
                                                strData = GetLabelDataName(strValue)
                                                
                                                '第一栏时数据复位,以保持各栏表头一致
                                                If k = 1 Then
                                                    If strData <> "" Then
                                                        For l = 0 To UBound(Split(strData, "|"))
                                                            strTmp = Split(Split(strData, "|")(l), ".")(0)
                                                            If frmSource.mLibDatas("_" & strTmp).DataSet.RecordCount > 0 Then
                                                                frmSource.mLibDatas("_" & strTmp).DataSet.MoveFirst
                                                                '(当前页-1)表示要循环Move的次数
                                                                For M = 1 To intPage
                                                                    If Not frmSource.mLibDatas("_" & strTmp).DataSet.EOF Then
                                                                        frmSource.mLibDatas("_" & strTmp).DataSet.MoveNext
                                                                    End If
                                                                    If frmSource.mLibDatas("_" & strTmp).DataSet.EOF Then
                                                                        frmSource.mLibDatas("_" & strTmp).DataSet.MoveFirst
                                                                    End If
                                                                Next
                                                            End If
                                                        Next
                                                    End If
                                                End If
                                                
                                                '再取数据
                                                If strData <> "" Then
                                                    For l = 0 To UBound(Split(strData, "|"))
                                                        strTmp = GetFieldValue(frmSource, CStr(Split(strData, "|")(l)))
                                                        strValue = Replace(strValue, "[" & Split(strData, "|")(l) & "]", strTmp)
                                                    Next
                                                End If
                                                
                                                '再处理报表变量:[页号]、[页数]、[=参数名]、[n>=0]、[日期格式串]、[单位名称]
                                                strValue = GetLabelMacro(frmSource, strValue)
                                                
                                                '输出单元格
                                                SearchCell objHead, i, j, objHead.FixedRows - 1, .ColE, W, H, strSkip, strSkip2
                                                
                                                strBdr = "1111"
                                                If Not objItem.边框 Then
                                                    'If j = .ColB And k = 1 Then
                                                    If B = 0 And InStr(strSkip, "[" & i & "," & j - 1 & "]") = 0 And W > 0 And k = 1 Then
                                                        strBdr = "1101"
                                                    ElseIf j = .ColE Or GetRightWidth(j, .ColE, i, strSkip, strSkip2, objHead) = 0 Or (W > objHead.ColWidth(j) + 15 _
                                                        And Right(strSkip, Len("[" & i & "," & .ColE & "]")) = "[" & i & "," & .ColE & "]") Then
                                                        strBdr = "1110"
                                                    End If
                                                End If
                                                
                                                If W > 0 Then
                                                    Set objFont = objHead.Font
                                                    If objHead.Cell(flexcpFontBold, i, j) = True Then
                                                        objFont.Bold = True
                                                    Else
                                                        objFont.Bold = False
                                                    End If
                                                    lngForeColor = IIF(objHead.Cell(flexcpForeColor, i, j) = &HFF0001 _
                                                                            And objHead.Cell(flexcpFontUnderline, i, j) = True _
                                                                        , objHead.ForeColor _
                                                                        , objHead.Cell(flexcpForeColor, i, j))
                                                    If Not DrawCell(objOut, strValue, .X + X, .Y + Y, W, H, .X + .W * .Copys, _
                                                                , objHead.GridColorFixed, lngForeColor, objHead.BackColor _
                                                                , objFont, strBdr _
                                                                , GetHscAlign(objHead.Cell(flexcpAlignment, i, j), objHead.Text) _
                                                                , GetVscAlign(objHead.Cell(flexcpAlignment, i, j)) _
                                                                , True, sngScale, , objItem.表格线加粗) Then
                                                        Exit Function  '合并时允许换行
                                                    End If
                                                    B = B + 1
                                                End If
                                            End If
                                            X = X + objHead.ColWidth(j)
                                        Next
                                    End If
                                    
                                    Y = Y + objHead.RowHeight(i)
                                    
                                    If blnMeasure Then
                                        lngCurH = .Y + Y
                                        If lngCurH > lngMaxH Then lngMaxH = lngCurH
                                    End If
                                Next
                                
                                '测量高度时,只需处理一栏即可
                                If blnMeasure Then Exit For
                            Next
                        End If
                        
                        '处理数据单元
                        If objItem.类型 = 5 Then
                            strSkip = "": strSkip2 = "": Y = .FixH
                            For i = .RowB To .RowE
                                If Not blnMeasure Then
                                    objBody.Row = i: X = .FixW
                                    For j = .ColB To .ColE
                                        objBody.Col = j
                                        If InStr(strSkip, "[" & i & "," & j & "]") = 0 Then
                                            SearchCell objBody, i, j, .RowE, .ColE, W, H, strSkip, strSkip2
                                            
                                            If blnPressWork Then
                                                strBdr = "0000"
                                            Else
                                                strBdr = "1111"
                                                If Not objItem.边框 And (j = .ColE Or (W > objBody.ColWidth(j) + 15 _
                                                    And Right(strSkip, Len("[" & i & "," & .ColE & "]")) = "[" & i & "," & .ColE & "]")) Then strBdr = "1110"
                                            End If
                                            
                                            Set objFont = objBody.Font
                                            If objBody.Cell(flexcpFontBold, i, j) = True Then
                                                objFont.Bold = True
                                            Else
                                                objFont.Bold = False
                                            End If
                                            lngForeColor = IIF(objBody.Cell(flexcpForeColor, i, j) = &HFF0001 _
                                                                    And objBody.Cell(flexcpFontUnderline, i, j) = True _
                                                                , objBody.ForeColor _
                                                                , objBody.Cell(flexcpForeColor, i, j))
                                            lngBackColor = IIF(objBody.Cell(flexcpBackColor, i, j) = 0 _
                                                                , objBody.BackColor _
                                                                , objBody.Cell(flexcpBackColor, i, j))
                                            If lngForeColor = objBody.BackColor Then lngForeColor = objBody.ForeColor
                                            '汇总表格-表体
                                            If Not DrawCell(objOut, objBody.Text, .X + X, .Y + Y, W, H, .X + .W, _
                                                        , objBody.GridColor, lngForeColor, lngBackColor _
                                                        , objFont, strBdr _
                                                        , GetHscAlign(objBody.Cell(flexcpAlignment, i, j), objBody.Text, objBody, j) _
                                                        , GetVscAlign(objBody.Cell(flexcpAlignment, i, j)) _
                                                        , objItem.自调, sngScale, , objItem.表格线加粗) Then
                                                Exit Function
                                            End If
                                        End If
                                        X = X + objBody.ColWidth(j)
                                    Next
                                End If
                                
                                Y = Y + objBody.RowHeight(i)
                                
                                If blnMeasure Then
                                    lngCurH = .Y + Y
                                    If lngCurH > lngMaxH Then lngMaxH = lngCurH
                                End If
                            Next
                        ElseIf .H >= .FixH Then
                            '判断附加表格
                            lngRowTotal = 0
                            blnAddition = HaveAdditionTable(frmSource.mobjReport, objItem)
                            For k = 1 To .Copys '处理任意表格数据分栏
                                strSkip = "": strSkip2 = ""
                                
                                Y = .FixH
                                
                                If frmSource.mobjReport.票据 _
                                    And (UBound(arrPage) = intPage Or GridAtCard(frmSource.mobjReport, objItem.id)) Then
                                    '求实际行范围的总高
                                    lngRangeHeight = 0
                                    lngRowCount = 0
                                    For i = .RowB + lngRowTotal To .RowE
                                        If i > objBody.Rows - 1 Then
                                            lngRowHeight = objItem.行高
                                        Else
                                            lngRowHeight = objBody.RowHeight(i)
                                        End If
                                        
                                        If lngRangeHeight + lngRowHeight >= objBody.Height Then
                                            If lngRowHeight >= objBody.Height And lngRangeHeight = 0 Then
                                                '当实际行高大于标准行高，并且lngRangeHeight为0时，强制一行
                                                lngRowCount = 1
                                            End If
                                            Exit For
                                        Else
                                            lngRangeHeight = lngRangeHeight + lngRowHeight
                                            lngRowCount = lngRowCount + 1
                                        End If
                                    Next
                                    lngRowTotal = lngRowTotal + lngRowCount
                                                    
                                    '求表格在当前页的输出行数
                                    LngRows = 0
                                    If objBody.Height > lngRangeHeight Then
                                        '计算表体还能容纳多少标准行数
                                        Do While objBody.Height - lngRangeHeight > objItem.行高
                                            LngRows = LngRows + 1
                                            lngRangeHeight = lngRangeHeight + objItem.行高
                                        Loop
                                        LngRows = LngRows + lngRowCount
                                    ElseIf objBody.Height = lngRangeHeight Then
                                        lngRowCount = lngRowCount - 1
                                    Else
                                        '计算实际行总高与设计表格高的差还能容纳多少标准行数
                                        Do While lngRangeHeight - objBody.Height >= objItem.行高
                                            LngRows = LngRows + 1
                                            lngRangeHeight = lngRangeHeight - objItem.行高
                                        Loop
                                    End If
                                    If .Copys > 1 Then
                                        '票据，表格多栏
                                        If LngRows <= 0 Then
                                            LngRows = lngRowCount
                                        End If
                                    Else
                                        '票据，表格非多栏
                                        If LngRows <= .RowE - .RowB Then
                                            LngRows = lngRowCount
                                        End If
                                    End If
                                Else
                                    '确定每栏内的起止行范围
                                    LngRows = (IIF(.VRowE <> 0, .VRowE, .RowE) - .RowB + 1) / .Copys '每栏应输出行数
                                End If
                                
                                If k > 1 Then
                                    lngRowB = lngRowE + 1
                                Else
                                    lngRowB = .RowB + LngRows * (k - 1)
                                End If
                                lngRowE = lngRowB + LngRows - 1
                                
                                For i = lngRowB To lngRowE
                                    If i > .RowE Then
                                        '如果存在附加表格，输出虚拟空行会引起空行重叠。即不用再输出虚拟空行
                                        If blnAddition Then Exit For
                                        
                                        '判断当前附加表格底部是否存在其他附加表格（票据格式）
                                        If frmSource.mobjReport.票据 And IsBottomAdditionGrid(frmSource.mobjReport.Items, objItem) Then
                                            Exit For
                                        End If
                                        
                                        '补充的虚拟空行输出：以RowE相同列为参照
                                        If Not blnMeasure Then
                                            X = .W * (k - 1)
                                            H = objItem.行高
                                            B = 0
                                            For j = .ColB To .ColE
                                                W = objBody.ColWidth(j)
    
                                                If blnPressWork Then
                                                    strBdr = "0000"
                                                Else
                                                    strBdr = "1111"
                                                    If Not objItem.边框 Then
                                                        If B = 0 And W > 0 And k = 1 Then
                                                            strBdr = "1101"
                                                        ElseIf j = .ColE Then
                                                            strBdr = "1110"
                                                        End If
                                                    End If
                                                End If
                                                
                                                If W > 0 Then
                                                    If Not DrawCell(objOut, "", .X + X, .Y + Y, W, H, .X + .W * .Copys, _
                                                                , objBody.GridColor, objBody.ForeColor, objBody.BackColor _
                                                                , objFont, strBdr _
                                                                , GetHscAlign(objBody.CellAlignment, objBody.Text) _
                                                                , GetVscAlign(objBody.CellAlignment) _
                                                                , objItem.自调, sngScale, , objItem.表格线加粗) Then
                                                        Exit Function
                                                    End If
                                                    B = B + 1
                                                End If
    
                                                X = X + objBody.ColWidth(j)
                                            Next
                                        End If
                                        
                                        Y = Y + objItem.行高
                                    ElseIf i <= objBody.Rows - 1 Then
                                        '行未超页
                                        If GetGridStyle(frmSource.mobjReport, objBody.Index) = Val("1-只有表头，无表体") Then
                                            '忽略表体无字段设置的虚拟空行
                                            Exit For
                                        End If
                                        
                                        '数据行输出
                                        If Not blnMeasure Then
                                            '激活打印行事件：当该行有数据要打印时
                                            If Not objCurDLL Is Nothing And TypeName(objOut) = "Printer" Then
                                                For j = objBody.FixedCols To objBody.Cols - 1
                                                    If objBody.ColWidth(j) <> 0 Then
                                                        If objBody.TextMatrix(i, j) <> "" Then Exit For
                                                    End If
                                                Next
                                                If j <= objBody.Cols - 1 Then
                                                    Call objCurDLL.Act_PrintSheetRow( _
                                                            frmSource.mobjReport.编号 _
                                                            , objBody _
                                                            , intPage + intBasePage _
                                                            , i + 1 - .RowB _
                                                            , colRowIDs("_" & objBody.Index)(i))
                                                End If
                                            End If
                                            
                                            objBody.Row = i: X = .W * (k - 1)
                                            B = 0
                                            For j = .ColB To .ColE
                                                objBody.Col = j
                                                If InStr(strSkip, "[" & i & "," & j & "]") = 0 Then
                                                    SearchCell objBody, i, j _
                                                            , IIF(lngRowE > objBody.Rows - 1, objBody.Rows - 1, lngRowE) _
                                                            , .ColE, W, H, strSkip, strSkip2
                                                    
                                                    If blnPressWork Then
                                                        strBdr = "0000"
                                                    Else
                                                        strBdr = "1111"
                                                        If Not objItem.边框 Then
                                                            If B = 0 And W > 0 And k = 1 Then
                                                                strBdr = "1101"
                                                            ElseIf j = .ColE Or GetRightWidth(j, .ColE, i, strSkip, strSkip2, objBody) = 0 Then
                                                                strBdr = "1110"
                                                            End If
                                                        End If
                                                    End If
                                                    
                                                    If W > 0 Then
                                                        Set objPic = objBody.CellPicture
                                                        If Not objPic Is Nothing Then
                                                            '奇怪，每个单元的图片都不为空
                                                            If objPic.handle = 0 Then Set objPic = Nothing
                                                        End If
                                                        If Not objPic Is Nothing Then
                                                            Set objSub = GetSubItem(frmSource, objItem, j)
                                                            
                                                            strData = GetLabelDataName(objSub.内容)
                                                            If strData <> "" Then
                                                                strTmp = Split(strData, ".")(0)
                                                                On Error Resume Next
                                                                frmSource.mLibDatas("_" & strTmp).DataSet.AbsolutePosition = i + 1
                                                                Err.Clear
                                                                On Error GoTo 0
                                                            End If
                                                            strValue = GetFieldValue(frmSource, strData)
                                                            If gobjFile.FileExists(strValue) Then
                                                                '二进制字段当作图形
                                                                On Error Resume Next
                                                                Set objPic = LoadPicture(strValue)
                                                                Kill strValue
                                                                Err.Clear
                                                                On Error GoTo 0
                                                            End If
                                                            If Not DrawCell(objOut, objPic, .X + X, .Y + Y, W, H, .X + .W * .Copys, _
                                                                        , objBody.GridColor, , , , strBdr _
                                                                        , GetHscAlign(objBody.Cell(flexcpAlignment, i, j), objBody.Text) _
                                                                        , GetVscAlign(objBody.CellAlignment) _
                                                                        , True, sngScale, , objItem.表格线加粗) Then
                                                                Exit Function
                                                            End If
                                                        Else
                                                            Set objFont = objBody.Font
                                                            
                                                            '检查并设置列的自动字体,利用缓存,尽量加快速度
                                                            If colColAutoFont("_" & j) = "" Then
                                                                colColAutoFont.Remove "_" & j: colColAutoFont.Add "0", "_" & j
                                                                Set objSub = GetSubItem(frmSource, objItem, j)
                                                                If Not objSub Is Nothing Then
                                                                    '“行高（缩小字体）”与“自适应行高”为互斥关系
                                                                    If objSub.行高 = 1 Then
                                                                        colColAutoFont.Remove "_" & j: colColAutoFont.Add "1", "_" & j
                                                                    ElseIf objSub.自适应行高 = True Then
                                                                        colColAutoFont.Remove "_" & j: colColAutoFont.Add "2", "_" & j
                                                                    End If
                                                                End If
                                                            End If
                                                            Select Case colColAutoFont("_" & j)
                                                            Case "1"    '缩小字体
                                                                Set objFont = GetAutoFont(objBody.Text, W, H, objFont, objOut, objItem.自调)
                                                            Case "2"    '自适应行高
                                                                '
                                                            End Select
                                                            
                                                            If lngindex <> -1 Then
                                                                If objBody.TextMatrix(i, lngindex) = "1" Then
                                                                    objFont.Strikethrough = True
                                                                Else
                                                                    objFont.Strikethrough = False
                                                                End If
                                                            Else
                                                                objFont.Strikethrough = False
                                                            End If
                                                            
                                                            If objBody.Cell(flexcpFontBold, i, j) = True Then
                                                                objFont.Bold = True
                                                            Else
                                                                objFont.Bold = False
                                                            End If
                                                            lngForeColor = IIF(objBody.Cell(flexcpForeColor, i, j) = &HFF0001 _
                                                                                    And objBody.Cell(flexcpFontUnderline, i, j) = True _
                                                                                , objBody.ForeColor _
                                                                                , objBody.Cell(flexcpForeColor, i, j))
                                                            lngBackColor = IIF(objBody.Cell(flexcpBackColor, i, j) = 0 _
                                                                                , objBody.BackColor _
                                                                                , objBody.Cell(flexcpBackColor, i, j))
                                                            If lngForeColor = objBody.BackColor Then lngForeColor = objBody.ForeColor
                                                            '自由表格-表体
                                                            If Not DrawCell(objOut, objBody.Text, .X + X, .Y + Y, W, H, .X + .W * .Copys, _
                                                                        , objBody.GridColor, lngForeColor, lngBackColor _
                                                                        , objFont, strBdr _
                                                                        , GetHscAlign(objBody.Cell(flexcpAlignment, i, j), objBody.Text, objBody, j) _
                                                                        , GetVscAlign(objBody.ColAlignment(j)) _
                                                                        , objItem.自调, sngScale, , objItem.表格线加粗 _
                                                                        , , colColAutoFont("_" & j) = "2") Then
                                                                Exit Function
                                                            End If
                                                        End If
                                                        B = B + 1
                                                    End If
                                                End If
                                                X = X + objBody.ColWidth(j)
                                            Next
                                        End If
                                        
                                        Y = Y + objBody.RowHeight(i)
                                    End If
                                    
                                    If blnMeasure Then
                                        lngCurH = .Y + Y
                                        If lngCurH > lngMaxH Then lngMaxH = lngCurH
                                    End If
                                Next
                                
                                '测量高度时,只需处理一栏即可
                                If blnMeasure Then Exit For
                            Next
                        End If
                        objBody.Redraw = True
                        If Not objHead Is Nothing Then objHead.Redraw = True
                        objBody.Row = lngPreRow: objBody.Col = lngPreCol
                    End With
                Next
            End If
        End If
    End If
    
    '输出非表格内容
    If Not blnPressWork Then
        For Each objItem In frmSource.mobjReport.Items
            If objItem.格式号 = frmSource.bytFormat Then
                Set objFont = New StdFont
                With objItem
                    blnWithData = False
                    If objItem.父ID <> 0 Then
                        lngChildX = frmSource.mobjReport.Items("_" & objItem.父ID).X
                        lngChildY = frmSource.mobjReport.Items("_" & objItem.父ID).Y
                        If frmSource.mobjReport.Items("_" & objItem.父ID).数据源 <> "" Then
                            blnWithData = True
                        End If
                    Else
                        lngChildX = 0
                        lngChildY = 0
                    End If
                    '如果是动态打印的卡片内的内容，则在后面打印
                    If blnWithData = False Then
                        Select Case .类型
                            Case 10 '框线
                                If Not blnMeasure Then
                                    If Not DrawCell(objOut, -1, .X + lngChildX, .Y + lngChildY, .W, .H, lngPaperW, lngPaperH, 0, .前景, , , , , , , sngScale, , .粗体, IIF(.边框, 1, 0)) Then Exit Function
                                Else
                                    lngCurH = .Y + .H
                                    If lngCurH > lngPaperH Then lngCurH = lngPaperH
                                    If lngCurH > lngMaxH Then lngMaxH = lngCurH
                                End If
                            Case 11 '图片
                                Set objPic = LoadPictureFromPar(frmSource, .名称)
                                If objPic Is Nothing Then Set objPic = .图片
                                If .自调 And Not objPic Is Nothing Then
                                    .W = objPic.Width * (15 / 26.46)
                                    .H = objPic.Height * (15 / 26.46)
                                End If
                                If Not blnMeasure Then
                                    If Not DrawCell(objOut, objPic, .X + lngChildX, .Y + lngChildY, .W, .H, lngPaperW, lngPaperH, 0, .前景, , , IIF(.边框, "1111", "0000"), 0, 0, .粗体, sngScale) Then Exit Function
                                Else
                                    lngCurH = .Y + .H
                                    If lngCurH > lngPaperH Then lngCurH = lngPaperH
                                    If lngCurH > lngMaxH Then lngMaxH = lngCurH
                                End If
                            Case 14 '卡片
                                If Not blnMeasure Then
                                    If Not DrawCell(objOut, objPic, .X + lngChildX, .Y + lngChildY, .W, .H, lngPaperW, lngPaperH, 0, .前景, , , IIF(.边框, "1111", "0000"), 0, 0, .粗体, sngScale) Then Exit Function
                                Else
                                    lngCurH = .Y + .H
                                    If lngCurH > lngPaperH Then lngCurH = lngPaperH
                                    If lngCurH > lngMaxH Then lngMaxH = lngCurH
                                End If
                            Case 13 '条码
                                '获取条码内容
                                strValue = .内容
                                '处理页变量
                                If InStr(strValue, "[页号]") > 0 Then
                                    strValue = Replace(strValue, "[页号]", intPage + intBasePage)
                                End If
                                If InStr(strValue, "[页数]") > 0 Then
                                    If Not IsArray(arrPage) Then
                                        strValue = Replace(strValue, "[页数]", intBasePage)
                                    Else
                                        strValue = Replace(strValue, "[页数]", UBound(arrPage) + intBasePage)
                                    End If
                                End If
                                If InStr(strValue, "[票据号]") > 0 Then
                                    If IsArray(garrBill) Then
                                        If UBound(garrBill) >= intPage Then
                                            strValue = Replace(strValue, "[票据号]", garrBill(intPage))
                                        Else
                                            strValue = Replace(strValue, "[票据号]", "")
                                        End If
                                    Else
                                        strValue = Replace(strValue, "[票据号]", "")
                                    End If
                                End If
                                
                                '数据指针复位(可能用到多个数据源、多个字段)
                                strData = GetLabelDataName(strValue) '"数据源.字段"串
                                If strData <> "" Then
                                    For i = 0 To UBound(Split(strData, "|"))
                                        strTmp = Split(Split(strData, "|")(i), ".")(0)
                                        
                                        If frmSource.mLibDatas("_" & strTmp).DataSet.RecordCount > 0 Then
                                            frmSource.mLibDatas("_" & strTmp).DataSet.MoveFirst
                                            '(当前页-1)表示要循环Move的次数
                                            For j = 1 To intPage
                                                If Not frmSource.mLibDatas("_" & strTmp).DataSet.EOF Then
                                                    frmSource.mLibDatas("_" & strTmp).DataSet.MoveNext
                                                End If
                                                If frmSource.mLibDatas("_" & strTmp).DataSet.EOF Then
                                                    frmSource.mLibDatas("_" & strTmp).DataSet.MoveFirst
                                                End If
                                            Next
                                            If .源行号 <> 0 Then
                                                If .源行号 <= frmSource.mLibDatas("_" & strTmp).DataSet.RecordCount Then
                                                    frmSource.mLibDatas("_" & strTmp).DataSet.AbsolutePosition = .源行号
                                                End If
                                            End If
                                        End If
                                    Next
                                End If
                                
                                '先处理数据字段(查询时只取第一个值)
                                If strData <> "" Then
                                    For i = 0 To UBound(Split(strData, "|"))
                                        strTmp = GetFieldValue(frmSource, CStr(Split(strData, "|")(i)))
                                        If .格式 <> "" Then
                                            On Error Resume Next
                                            strTmp = Format(strTmp, .格式)
                                            If Err.Number <> 0 Then Err.Clear
                                            On Error GoTo 0
                                        End If
                                        strValue = Replace(strValue, "[" & Split(strData, "|")(i) & "]", strTmp)
                                    Next
                                End If
                                
                                '再处理报表变量:[页号]、[页数]、[=参数名]、[n>=0]、[日期格式串]、[单位名称]
                                strValue = GetLabelMacro(frmSource, strValue)
                                
                                '获取条码图形
                                Set objPic = Nothing
                                If strValue <> "" Then
                                    Unload frmFlash '强制初始Picture，不然切换绘制有问题
                                    If .序号 = 1 Then
                                        Set objPic = DrawBarCode128(frmFlash.picTemp, 3, strValue, Mid(.表头, 1, 1) = "1")
                                    ElseIf .序号 = 2 Then
                                        Set objPic = DrawBarCode39(frmFlash.picTemp, 3, strValue, Mid(.表头, 2, 1) = "1", Mid(.表头, 1, 1) = "1")
                                    ElseIf .序号 = 3 Then
                                        Set objPic = DrawBarCode128Auto(frmFlash.picTemp, strValue, sngWidth, .行高, Mid(.表头, 1, 1) = "1")
                                    ElseIf .序号 = 10 Then
                                        Set objPic = DrawBarCode2D(strValue, frmFlash.picTemp, lngSize)
                                    End If
                                    If Val(Mid(.表头, 3, 1)) <> 0 Then
                                        Set objPic = PictureSpin(objPic, Val(Mid(.表头, 3, 1)), frmFlash.picTemp)
                                    End If
                                    
                                    If .序号 = 3 And blnPrint = False Then
                                        '128码自动调整宽度
                                        If Val(Mid(.表头, 3, 1)) = 0 Then
                                            .W = objOut.ScaleX(sngWidth, vbMillimeters, vbTwips)
                                        Else
                                            .H = objOut.ScaleY(sngWidth, vbMillimeters, vbTwips)
                                        End If
                                    ElseIf .序号 = 10 And .自调 Then
                                        '二维条码缺省自动调整大小
                                        .W = lngSize: .H = lngSize
                                    End If
                                End If
                                
                                '输出图形
                                If Not blnMeasure Then
                                    If Not DrawCell(objOut, objPic, .X + lngChildX, .Y + lngChildY, .W, .H, lngPaperW, lngPaperH, 0, , , , IIF(.边框, "1111", "0000"), , , , sngScale) Then Exit Function
                                Else
                                    lngCurH = .Y + .H
                                    If lngCurH > lngPaperH Then lngCurH = lngPaperH
                                    If lngCurH > lngMaxH Then lngMaxH = lngCurH
                                End If
                            Case 1 '线条
                                If Not blnMeasure Then
                                    If Not DrawCell(objOut, 1, .X + lngChildX, .Y + lngChildY, .W, .H, lngPaperW, lngPaperH, .前景, .前景, , , , , , , sngScale, , .粗体) Then Exit Function
                                Else
                                    lngCurH = .Y + .H
                                    If lngCurH > lngPaperH Then lngCurH = lngPaperH
                                    If lngCurH > lngMaxH Then lngMaxH = lngCurH
                                End If
                            Case 12 '图表@@@
                                If intPage = 0 Then '只在第一页打印
                                    If Not blnMeasure Then
                                        If sngScale = 1 Then
                                            strTmp = gobjFile.GetSpecialFolder(TemporaryFolder) & "\" & gobjFile.GetTempName
                                            If frmSource.Chart(.id).SaveImageAsJpeg(strTmp, 100, False, False, False) Then
                                                Set objPic = LoadPicture(strTmp)
                                            End If
                                            If gobjFile.FileExists(strTmp) Then
                                                Call gobjFile.DeleteFile(strTmp, True)
                                            End If
                                        Else
                                            Load frmSource.Chart(9999)
                                            
                                            strTmp = GetChartFileFromPar(frmSource, .名称)
                                            If strTmp <> "" Then
                                                Call frmSource.Chart(9999).Load(strTmp)
                                                
                                                frmSource.Chart(9999).Left = 0
                                                frmSource.Chart(9999).Top = 0
                                                frmSource.Chart(9999).Width = frmSource.Chart(.id).Width * sngScale
                                                frmSource.Chart(9999).Height = frmSource.Chart(.id).Height * sngScale
                                                
                                                strTmp = gobjFile.GetSpecialFolder(TemporaryFolder) & "\" & gobjFile.GetTempName
                                                If frmSource.Chart(9999).SaveImageAsJpeg(strTmp, 100, False, False, False) Then
                                                    Set objPic = LoadPicture(strTmp)
                                                End If
                                                If gobjFile.FileExists(strTmp) Then
                                                    Call gobjFile.DeleteFile(strTmp, True)
                                                End If
                                            Else
                                                Call GetChartDataName(objItem.内容, , , , strTmp)
                                                If strTmp <> "" Then
                                                    Set objPic = GetChartPicture(frmSource.Chart(9999), frmSource.Chart(.id), objItem, frmSource.mLibDatas("_" & strTmp).DataSet, sngScale)
                                                Else
                                                    Set objPic = GetChartPicture(frmSource.Chart(9999), frmSource.Chart(.id), objItem, , sngScale)
                                                End If
                                            End If
                                            
                                            Unload frmSource.Chart(9999)
                                        End If
                                    
                                        If Not DrawCell(objOut, objPic, .X + lngChildX, .Y + lngChildY, .W, .H, lngPaperW, lngPaperH, , , , , IIF(.边框, "1111", "0000"), , , , sngScale) Then Exit Function
                                    Else
                                        lngCurH = .Y + .H
                                        If lngCurH > lngPaperH Then lngCurH = lngPaperH
                                        If lngCurH > lngMaxH Then lngMaxH = lngCurH
                                    End If
                                End If
                            Case 2, 3 '标签,标签绑定图片
                                objFont.name = .字体
                                objFont.Size = .字号
                                objFont.Bold = .粗体
                                objFont.Italic = .斜体
                                objFont.Underline = .下线
                                
                                strValue = .内容
                                '处理页变量
                                If InStr(strValue, "[页号]") > 0 Then
                                    strValue = Replace(strValue, "[页号]", intPage + intBasePage)
                                End If
                                If InStr(strValue, "[页数]") > 0 Then
                                    If Not IsArray(arrPage) Then
                                        strValue = Replace(strValue, "[页数]", intBasePage)
                                    Else
                                        strValue = Replace(strValue, "[页数]", UBound(arrPage) + intBasePage)
                                    End If
                                End If
                                If InStr(strValue, "[票据号]") > 0 Then
                                    If IsArray(garrBill) Then
                                        If UBound(garrBill) >= intPage Then
                                            strValue = Replace(strValue, "[票据号]", garrBill(intPage))
                                        Else
                                            strValue = Replace(strValue, "[票据号]", "")
                                        End If
                                    Else
                                        strValue = Replace(strValue, "[票据号]", "")
                                    End If
                                End If
                                
                                '数据指针复位(可能用到多个数据源、多个字段)
                                strData = GetLabelDataName(strValue) '"数据源.字段"串
                                If strData <> "" Then
                                    For i = 0 To UBound(Split(strData, "|"))
                                        strTmp = Split(Split(strData, "|")(i), ".")(0)
                                        
                                        If frmSource.mLibDatas("_" & strTmp).DataSet.RecordCount > 0 Then
                                            frmSource.mLibDatas("_" & strTmp).DataSet.MoveFirst
                                            '(当前页-1)表示要循环Move的次数
                                            For j = 1 To intPage
                                                If Not frmSource.mLibDatas("_" & strTmp).DataSet.EOF Then
                                                    frmSource.mLibDatas("_" & strTmp).DataSet.MoveNext
                                                End If
                                                If frmSource.mLibDatas("_" & strTmp).DataSet.EOF Then
                                                    frmSource.mLibDatas("_" & strTmp).DataSet.MoveFirst
                                                End If
                                            Next

                                            If .源行号 <> 0 Then
                                                If .源行号 <= frmSource.mLibDatas("_" & strTmp).DataSet.RecordCount Then
                                                    frmSource.mLibDatas("_" & strTmp).DataSet.AbsolutePosition = .源行号
                                                End If
                                            End If
                                        End If
                                        
                                        '先处理数据字段(查询时只取第一个值)
                                        strTmp = GetFieldValue(frmSource, CStr(Split(strData, "|")(i)))
                                        If .格式 <> "" Then
                                            On Error Resume Next
                                            strTmp = Format(strTmp, .格式)
                                            If Err.Number <> 0 Then Err.Clear
                                            On Error GoTo 0
                                        End If
                                        strValue = Replace(strValue, "[" & Split(strData, "|")(i) & "]", strTmp)
                                    Next
                                End If
                                
                                '再处理报表变量:[页号]、[页数]、[=参数名]、[n>=0]、[日期格式串]、[单位名称]
                                strValue = GetLabelMacro(frmSource, strValue)
                                If gobjFile.FileExists(strValue) Then
                                    '二进制字段当作图形
                                    On Error Resume Next
                                    Set .图片 = LoadPicture(strValue)
                                    Kill strValue
                                    Err.Clear
                                    On Error GoTo 0
                                    
                                    If .自调 And Not .图片 Is Nothing Then
                                        .W = .图片.Width * (15 / 26.46)
                                        .H = .图片.Height * (15 / 26.46)
                                    End If
                                    
                                    If Not blnMeasure Then
                                        If Not DrawCell(objOut, .图片, .X + lngChildX, .Y + lngChildY, .W, .H, lngPaperW, lngPaperH, 0, .前景, , , IIF(.边框, "1111", "0000"), 0, 0, .粗体, sngScale) Then Exit Function
                                    Else
                                        lngCurH = .Y + .H
                                        If lngCurH > lngPaperH Then lngCurH = lngPaperH
                                        If lngCurH > lngMaxH Then lngMaxH = lngCurH
                                    End If
                                Else
                                    If .自调 Then Call ItemAutoSize(objItem, strValue, objOut)
                                    If objItem.性质 > 0 And objItem.参照 <> "" And blnHaveGrid Then
                                        '计算靠齐标签的位置
                                        strDepend = GetDependIDs(.参照, frmSource)
                                        lngOX = 0: lngOY = 0: lngOH = 0: lngOW = 0: lngDesignH = 0
                                        For Each objPageCell In arrPage(intPage)
                                            If InStr("," & strDepend & ",", "," & objPageCell.id & ",") > 0 Then
                                                If lngOX = 0 And lngOY = 0 And lngOH = 0 And lngOW = 0 Then
                                                    lngOX = objPageCell.X
                                                    lngOY = objPageCell.Y
                                                    lngOW = objPageCell.W * objPageCell.Copys
                                                    lngDesignH = objPageCell.MaxH
                                                End If
                                                lngOH = lngOH + objPageCell.H
                                            End If
                                        Next
        
                                        '左右靠齐
                                        Select Case .性质
                                            Case 11, 21 '左
                                                lngOutX = lngOX
                                            Case 12, 22 '中
                                                lngOutX = lngOX + (lngOW - .W) / 2
                                            Case 13, 23 '右
                                                lngOutX = lngOX + lngOW - .W
                                        End Select
                                        '上下靠齐
                                        If frmSource.mobjReport.票据 Then
                                            lngOutY = .Y '票据时位置应该不变
                                        Else
                                            If CInt(Left(CStr(.性质), 1)) = 2 Then
                                                lngOutY = lngOY + lngOH + (.Y - (lngOY + lngDesignH))
                                            Else
                                                lngOutY = .Y
                                            End If
                                        End If
                                        If strValue <> "" Then
                                            If Not blnMeasure Then
                                                If .行高 = 1 Then Set objFont = GetAutoFont(strValue, .W, .H, objFont, objOut, True, .网格)
                                                If .水平反转 Then
                                                    If Not DrawCell(objOut, frmSource.picRotate(.id).Image, .X + lngChildX, .Y + lngChildY, .W, .H, lngPaperW, lngPaperH, 0, .前景, .背景, objFont, "0000", .对齐, 0, True, sngScale, .网格) Then Exit Function
                                                Else
                                                    If Not DrawCell(objOut, strValue, lngOutX, lngOutY, .W, .H, lngPaperW, lngPaperH, 0, .前景, .背景, objFont, IIF(.边框, "1111", "0000"), .对齐, 0, True, sngScale, .网格) Then Exit Function
                                                End If
                                            Else
                                                lngCurH = lngOutY + .H
                                                If lngCurH > lngPaperH Then lngCurH = lngPaperH
                                                If lngCurH > lngMaxH Then lngMaxH = lngCurH
                                            End If
                                        End If
                                    Else
                                        If strValue <> "" Then
                                            If Not blnMeasure Then
                                                If .行高 = 1 Then Set objFont = GetAutoFont(strValue, .W, .H, objFont, objOut, True, .网格)
                                                If .水平反转 Then
                                                    If Not DrawCell(objOut, frmSource.picRotate(.id).Image, .X + lngChildX, .Y + lngChildY, .W, .H, lngPaperW, lngPaperH, 0, .前景, .背景, objFont, "0000", .对齐, 0, True, sngScale, .网格) Then Exit Function
                                                Else
                                                    If Not DrawCell(objOut, strValue, .X + lngChildX, .Y + lngChildY, .W, .H, lngPaperW, lngPaperH, 0, .前景, .背景, objFont, IIF(.边框, "1111", "0000"), .对齐, 0, True, sngScale, .网格) Then Exit Function
                                                End If
                                            Else
                                                lngCurH = .Y + .H
                                                If lngCurH > lngPaperH Then lngCurH = lngPaperH
                                                If lngCurH > lngMaxH Then lngMaxH = lngCurH
                                            End If
                                        End If
                                    End If
                                    
                                End If
                                
                            Case Else
                                '
                        End Select
                    End If
                End With
            End If
            Set objPic = Nothing
        Next
    End If
    
    If Not blnPressWork Then
        If IsArray(arrPageCard) Then
            If UBound(arrPageCard) >= intPage Then
                If arrPageCard(intPage).count > 0 Then
                    For Each objPageCard In arrPageCard(intPage).Items
                        lngCol = 0: lngRow = 0
                        For Y = 1 To objPageCard.Item.count
                            '先输出卡片对象
                            On Error Resume Next
                            Set objTemp = frmSource.mobjReport.Items("_" & objPageCard.id)
                            If Err.Number <> 0 Then
                                On Error GoTo 0
                                Exit For
                            End If
                            On Error GoTo 0
                            If objTemp Is Nothing Then Exit Function
                            
                            '可能卡片对象后于卡片里的对象创建，因此，先输出卡片再输出卡片里的对象
                            '输出卡片
                            With objTemp
                                Set objPic = LoadPictureFromPar(frmSource, .名称)
                                If lngCol >= objPageCard.Col Then lngRow = lngRow + 1: lngCol = 0
                                lngX = lngRow * (.H + .上下间距)
                                lngY = lngCol * (.W + .左右间距)
                                If Not blnMeasure Then
                                    If Not DrawCell(objOut, objPic, .X + lngY, .Y + lngX, .W, .H, lngPaperW, lngPaperH, 0 _
                                                , .前景, , , IIF(.边框, "1111", "0000"), 0, 0, .粗体, sngScale) Then
                                        Exit Function
                                    End If
                                Else
                                    lngCurH = .Y + .H
                                    If lngCurH > lngPaperH Then lngCurH = lngPaperH
                                    If lngCurH > lngMaxH Then lngMaxH = lngCurH
                                End If
                                lngCol = lngCol + 1
                            End With
                        
                            '再输出卡片里的对象
                            For Each objItem In frmSource.mobjReport.Items
                                If objItem.格式号 = frmSource.bytFormat Then
                                    Set objFont = New StdFont
                                    With objItem
                                        If .父ID <> 0 Then
                                            lngChildX = frmSource.mobjReport.Items("_" & .父ID).X
                                            lngChildY = frmSource.mobjReport.Items("_" & .父ID).Y
                                        Else
                                            lngChildX = 0
                                            lngChildY = 0
                                        End If
                                        
                                        If .父ID = objPageCard.id Then
                                            '再输出卡片中的内容
                                            Select Case .类型
                                            Case 10 '框线
                                                If Not blnMeasure Then
                                                    If Not DrawCell(objOut, -1, .X + lngY + lngChildX, .Y + lngX + lngChildY, .W, .H, lngPaperW, lngPaperH, 0, .前景, , , , , , , sngScale, , .粗体, IIF(.边框, 1, 0)) Then Exit Function
                                                Else
                                                    lngCurH = .Y + .H
                                                    If lngCurH > lngPaperH Then lngCurH = lngPaperH
                                                    If lngCurH > lngMaxH Then lngMaxH = lngCurH
                                                End If
                                            Case 1 '线条
                                                If Not blnMeasure Then
                                                    If Not DrawCell(objOut, 1, .X + lngY + lngChildX, .Y + lngX + lngChildY, .W, .H, lngPaperW, lngPaperH, .前景, .前景, , , , , , , sngScale, , .粗体) Then Exit Function
                                                Else
                                                    lngCurH = .Y + .H
                                                    If lngCurH > lngPaperH Then lngCurH = lngPaperH
                                                    If lngCurH > lngMaxH Then lngMaxH = lngCurH
                                                End If
                                            Case 11 '图片
                                                Set objPic = LoadPictureFromPar(frmSource, .名称)
                                                If objPic Is Nothing Then Set objPic = .图片
                                                If .自调 And Not objPic Is Nothing Then
                                                    .W = objPic.Width * (15 / 26.46)
                                                    .H = objPic.Height * (15 / 26.46)
                                                End If
                                                If Not blnMeasure Then
                                                    If Not DrawCell(objOut, objPic, .X + lngY + lngChildX, .Y + lngX + lngChildY, .W, .H, lngPaperW, lngPaperH, 0, .前景, , , IIF(.边框, "1111", "0000"), 0, 0, .粗体, sngScale) Then Exit Function
                                                Else
                                                    lngCurH = .Y + .H
                                                    If lngCurH > lngPaperH Then lngCurH = lngPaperH
                                                    If lngCurH > lngMaxH Then lngMaxH = lngCurH
                                                End If
                                            Case 13 '条码
                                                '获取条码内容
                                                strValue = .内容
                                                '处理页变量
                                                If InStr(strValue, "[页号]") > 0 Then
                                                    strValue = Replace(strValue, "[页号]", intPage + intBasePage)
                                                End If
                                                If InStr(strValue, "[页数]") > 0 Then
                                                    If Not IsArray(arrPage) Then
                                                        strValue = Replace(strValue, "[页数]", intBasePage)
                                                    Else
                                                        strValue = Replace(strValue, "[页数]", UBound(arrPage) + intBasePage)
                                                    End If
                                                End If
                                                If InStr(strValue, "[票据号]") > 0 Then
                                                    If IsArray(garrBill) Then
                                                        If UBound(garrBill) >= intPage Then
                                                            strValue = Replace(strValue, "[票据号]", garrBill(intPage))
                                                        Else
                                                            strValue = Replace(strValue, "[票据号]", "")
                                                        End If
                                                    Else
                                                        strValue = Replace(strValue, "[票据号]", "")
                                                    End If
                                                End If
                                                
                                                '数据指针复位(可能用到多个数据源、多个字段)
                                                strData = GetLabelDataName(strValue) '"数据源.字段"串
                                                If strData <> "" Then
                                                    For i = 0 To UBound(Split(strData, "|"))
                                                        strTmp = Split(Split(strData, "|")(i), ".")(0)
                                                        
                                                        If frmSource.mLibDatas("_" & strTmp).DataSet.RecordCount > 0 Then
                                                            frmSource.mLibDatas("_" & strTmp).DataSet.MoveFirst
                                                            '(当前页-1)表示要循环Move的次数
                                                            For j = 1 To intPage
                                                                If Not frmSource.mLibDatas("_" & strTmp).DataSet.EOF Then
                                                                    frmSource.mLibDatas("_" & strTmp).DataSet.MoveNext
                                                                End If
                                                                If frmSource.mLibDatas("_" & strTmp).DataSet.EOF Then
                                                                    frmSource.mLibDatas("_" & strTmp).DataSet.MoveFirst
                                                                End If
                                                            Next
                                                            If frmSource.mobjReport.Items("_" & .父ID).数据源 = strTmp Then
                                                                blnGroup = False
                                                                On Error Resume Next
                                                                If frmSource.mLibDatas("_" & strTmp).DataSet!分组标识 & "" <> "" Or frmSource.mLibDatas("_" & strTmp).DataSet!分组标识 & "" = "" Then
                                                                    If Err.Number = 0 Then
                                                                        '按组动态打印
                                                                        blnGroup = True
                                                                    End If
                                                                    Err.Clear: On Error GoTo 0
                                                                    If arrPageCard(intPage).count > 0 Then
                                                                        frmSource.mLibDatas("_" & strTmp).DataSet.AbsolutePosition = Val(Mid(objPageCard.Item(Y), 1, InStr(objPageCard.Item(Y), "-") - 1))
                                                                    End If
                                                                End If
                                                            End If
                                                            If .源行号 <> 0 Then
                                                                If blnGroup Then
                                                                    '按组动态打印
                                                                    If .源行号 <= Val(Mid(objPageCard.Item(Y), InStr(objPageCard.Item(Y), "-") + 1, Len(objPageCard.Item(Y)))) - Val(Mid(objPageCard.Item(Y), 1, InStr(objPageCard.Item(Y), "-") - 1)) + 1 Then
                                                                        frmSource.mLibDatas("_" & strTmp).DataSet.AbsolutePosition = frmSource.mLibDatas("_" & strTmp).DataSet.AbsolutePosition + .源行号 - 1
                                                                    End If
                                                                Else
                                                                    If .源行号 <= frmSource.mLibDatas("_" & strTmp).DataSet.RecordCount Then
                                                                        frmSource.mLibDatas("_" & strTmp).DataSet.AbsolutePosition = .源行号
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    Next
                                                End If
                                                
                                                '先处理数据字段(查询时只取第一个值)
                                                If strData <> "" Then
                                                    For i = 0 To UBound(Split(strData, "|"))
                                                        strTmp = GetFieldValue(frmSource, CStr(Split(strData, "|")(i)))
                                                        If .格式 <> "" Then
                                                            On Error Resume Next
                                                            strTmp = Format(strTmp, .格式)
                                                            If Err.Number <> 0 Then Err.Clear
                                                            On Error GoTo 0
                                                        End If
                                                        strValue = Replace(strValue, "[" & Split(strData, "|")(i) & "]", strTmp)
                                                    Next
                                                End If
                                                
                                                '再处理报表变量:[页号]、[页数]、[=参数名]、[n>=0]、[日期格式串]、[单位名称]
                                                strValue = GetLabelMacro(frmSource, strValue)
                                                
                                                '获取条码图形
                                                Set objPic = Nothing
                                                If strValue <> "" Then
                                                    Unload frmFlash '强制初始Picture，不然切换绘制有问题
                                                    If .序号 = 1 Then
                                                        Set objPic = DrawBarCode128(frmFlash.picTemp, 3, strValue, Mid(.表头, 1, 1) = "1")
                                                    ElseIf .序号 = 2 Then
                                                        Set objPic = DrawBarCode39(frmFlash.picTemp, 3, strValue, Mid(.表头, 2, 1) = "1", Mid(.表头, 1, 1) = "1")
                                                    ElseIf .序号 = 3 Then
                                                        Set objPic = DrawBarCode128Auto(frmFlash.picTemp, strValue, sngWidth, .行高, Mid(.表头, 1, 1) = "1")
                                                    ElseIf .序号 = 10 Then
                                                        Set objPic = DrawBarCode2D(strValue, frmFlash.picTemp, lngSize)
                                                    End If
                                                    If Val(Mid(.表头, 3, 1)) <> 0 Then
                                                        Set objPic = PictureSpin(objPic, Val(Mid(.表头, 3, 1)), frmFlash.picTemp)
                                                    End If
                                                    
                                                    If .序号 = 3 Then
                                                        '128码自动调整宽度
                                                        If Val(Mid(.表头, 3, 1)) = 0 Then
                                                            .W = objOut.ScaleX(sngWidth, vbMillimeters, vbTwips)
                                                        Else
                                                            .H = objOut.ScaleY(sngWidth, vbMillimeters, vbTwips)
                                                        End If
                                                    ElseIf .序号 = 10 And .自调 Then
                                                        '二维条码缺省自动调整大小
                                                        .W = lngSize: .H = lngSize
                                                    End If
                                                End If
                                                
                                                '输出图形
                                                If Not blnMeasure Then
                                                    If Not DrawCell(objOut, objPic, .X + lngY + lngChildX, .Y + lngX + lngChildY, .W, .H, lngPaperW, lngPaperH, 0, , , , IIF(.边框, "1111", "0000"), , , , sngScale) Then Exit Function
                                                Else
                                                    lngCurH = .Y + .H
                                                    If lngCurH > lngPaperH Then lngCurH = lngPaperH
                                                    If lngCurH > lngMaxH Then lngMaxH = lngCurH
                                                End If
                                            Case 2, 3 '标签,标签绑定图片
                                                objFont.name = .字体
                                                objFont.Size = .字号
                                                objFont.Bold = .粗体
                                                objFont.Italic = .斜体
                                                objFont.Underline = .下线
                                                
                                                strValue = .内容
                                                '处理页变量
                                                If InStr(strValue, "[页号]") > 0 Then
                                                    strValue = Replace(strValue, "[页号]", intPage + intBasePage)
                                                End If
                                                If InStr(strValue, "[页数]") > 0 Then
                                                    If Not IsArray(arrPage) Then
                                                        strValue = Replace(strValue, "[页数]", intBasePage)
                                                    Else
                                                        strValue = Replace(strValue, "[页数]", UBound(arrPage) + intBasePage)
                                                    End If
                                                End If
                                                If InStr(strValue, "[票据号]") > 0 Then
                                                    If IsArray(garrBill) Then
                                                        If UBound(garrBill) >= intPage Then
                                                            strValue = Replace(strValue, "[票据号]", garrBill(intPage))
                                                        Else
                                                            strValue = Replace(strValue, "[票据号]", "")
                                                        End If
                                                    Else
                                                        strValue = Replace(strValue, "[票据号]", "")
                                                    End If
                                                End If
                                                
                                                '数据指针复位(可能用到多个数据源、多个字段)
                                                strData = GetLabelDataName(strValue) '"数据源.字段"串
                                                If strData <> "" Then
                                                    For i = 0 To UBound(Split(strData, "|"))
                                                        strTmp = Split(Split(strData, "|")(i), ".")(0)
                                                        
                                                        If frmSource.mLibDatas("_" & strTmp).DataSet.RecordCount > 0 Then
                                                            
                                                            frmSource.mLibDatas("_" & strTmp).DataSet.MoveFirst
                                                            '(当前页-1)表示要循环Move的次数
                                                            For j = 1 To intPage
                                                                If Not frmSource.mLibDatas("_" & strTmp).DataSet.EOF Then
                                                                    frmSource.mLibDatas("_" & strTmp).DataSet.MoveNext
                                                                End If
                                                                If frmSource.mLibDatas("_" & strTmp).DataSet.EOF Then
                                                                    frmSource.mLibDatas("_" & strTmp).DataSet.MoveFirst
                                                                End If
                                                            Next
                                                            If frmSource.mobjReport.Items("_" & .父ID).数据源 = strTmp Then
                                                                blnGroup = False
                                                                On Error Resume Next
                                                                If frmSource.mLibDatas("_" & strTmp).DataSet!分组标识 & "" <> "" Or frmSource.mLibDatas("_" & strTmp).DataSet!分组标识 & "" = "" Then
                                                                    If Err.Number = 0 Then
                                                                        '按组动态打印
                                                                        blnGroup = True
                                                                    End If
                                                                    Err.Clear: On Error GoTo 0
                                                                    If arrPageCard(intPage).count > 0 Then
                                                                        frmSource.mLibDatas("_" & strTmp).DataSet.AbsolutePosition = Val(Mid(objPageCard.Item(Y), 1, InStr(objPageCard.Item(Y), "-") - 1))
                                                                    End If
                                                                End If
                                                            End If
                                                            If .源行号 <> 0 Then
                                                                If blnGroup Then
                                                                    '按组动态打印
                                                                    If .源行号 <= Val(Mid(objPageCard.Item(Y), InStr(objPageCard.Item(Y), "-") + 1, Len(objPageCard.Item(Y)))) - Val(Mid(objPageCard.Item(Y), 1, InStr(objPageCard.Item(Y), "-") - 1)) + 1 Then
                                                                        frmSource.mLibDatas("_" & strTmp).DataSet.AbsolutePosition = frmSource.mLibDatas("_" & strTmp).DataSet.AbsolutePosition + .源行号 - 1
                                                                    End If
                                                                Else
                                                                    If .源行号 <= frmSource.mLibDatas("_" & strTmp).DataSet.RecordCount Then
                                                                        frmSource.mLibDatas("_" & strTmp).DataSet.AbsolutePosition = .源行号
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                        
                                                        '先处理数据字段(查询时只取第一个值)
                                                        strTmp = GetFieldValue(frmSource, CStr(Split(strData, "|")(i)))
                                                        If .格式 <> "" Then
                                                            On Error Resume Next
                                                            strTmp = Format(strTmp, .格式)
                                                            If Err.Number <> 0 Then Err.Clear
                                                            On Error GoTo 0
                                                        End If
                                                        strValue = Replace(strValue, "[" & Split(strData, "|")(i) & "]", strTmp)
                                                    Next
                                                End If
                                                
                                                '再处理报表变量:[页号]、[页数]、[=参数名]、[n>=0]、[日期格式串]、[单位名称]
                                                strValue = GetLabelMacro(frmSource, strValue)
                                                
                                                If gobjFile.FileExists(strValue) Then
                                                    '二进制字段当作图形
                                                    On Error Resume Next
                                                    Set .图片 = LoadPicture(strValue)
                                                    Kill strValue
                                                    Err.Clear
                                                    On Error GoTo 0
                                                    
                                                    If .自调 And Not .图片 Is Nothing Then
                                                        .W = .图片.Width * (15 / 26.46)
                                                        .H = .图片.Height * (15 / 26.46)
                                                    End If
                                                    
                                                    If Not blnMeasure Then
                                                        If Not DrawCell(objOut, .图片, .X + lngY + lngChildX, .Y + lngX + lngChildY, .W, .H, lngPaperW, lngPaperH, 0, .前景, , , IIF(.边框, "1111", "0000"), 0, 0, .粗体, sngScale) Then Exit Function
                                                    Else
                                                        lngCurH = .Y + .H
                                                        If lngCurH > lngPaperH Then lngCurH = lngPaperH
                                                        If lngCurH > lngMaxH Then lngMaxH = lngCurH
                                                    End If
                                                Else
                                                    If .自调 Then Call ItemAutoSize(objItem, strValue, objOut)
                                                    If objItem.性质 > 0 And objItem.参照 <> "" And blnHaveGrid Then
                                                        '计算靠齐标签的位置
                                                        strDepend = GetDependIDs(.参照, frmSource)
                                                        lngOX = 0: lngOY = 0: lngOH = 0: lngOW = 0: lngDesignH = 0
                                                        strTmp = ""
                                                        For Each objPageCell In arrPage(intPage)
                                                            If InStr("," & strDepend & ",", "," & objPageCell.id & ",") > 0 And InStr(strTmp & ",", "," & objPageCell.id & ",") = 0 Then
                                                                If lngOX = 0 And lngOY = 0 And lngOH = 0 And lngOW = 0 Then
                                                                    lngOX = objPageCell.X
                                                                    lngOY = objPageCell.Y
                                                                    lngOW = objPageCell.W * objPageCell.Copys
                                                                    lngDesignH = objPageCell.MaxH
                                                                End If
                                                                lngOH = lngOH + objPageCell.H
                                                                strTmp = strTmp & "," & objPageCell.id
                                                            End If
                                                        Next
                        
                                                        '左右靠齐
                                                        Select Case .性质
                                                            Case 11, 21 '左
                                                                lngOutX = lngOX
                                                            Case 12, 22 '中
                                                                lngOutX = lngOX + (lngOW - .W) / 2
                                                            Case 13, 23 '右
                                                                lngOutX = lngOX + lngOW - .W
                                                        End Select
                                                        '上下靠齐
                                                        If frmSource.mobjReport.票据 Then
                                                            lngOutY = .Y '票据时位置应该不变
                                                        Else
                                                            If CInt(Left(CStr(.性质), 1)) = 2 Then
                                                                lngOutY = lngOY + lngOH + (.Y - (lngOY + lngDesignH))
                                                            Else
                                                                lngOutY = .Y
                                                            End If
                                                        End If
                                                        If strValue <> "" Then
                                                            If Not blnMeasure Then
                                                                If .行高 = 1 Then Set objFont = GetAutoFont(strValue, .W, .H, objFont, objOut, True, .网格)
                                                                If .水平反转 Then
                                                                    If Not DrawCell(objOut, frmSource.picRotate(.id).Image, lngOutX + lngY, lngOutY + lngX + lngChildY, .W, .H, lngPaperW, lngPaperH, 0, .前景, .背景, objFont, "0000", .对齐, 0, True, sngScale, .网格) Then Exit Function
                                                                Else
                                                                    If Not DrawCell(objOut, strValue, lngOutX + lngY, lngOutY + lngX + lngChildY, .W, .H, lngPaperW, lngPaperH, 0, .前景, .背景, objFont, IIF(.边框, "1111", "0000"), .对齐, 0, True, sngScale, .网格) Then Exit Function
                                                                End If
                                                            Else
                                                                lngCurH = lngOutY + .H
                                                                If lngCurH > lngPaperH Then lngCurH = lngPaperH
                                                                If lngCurH > lngMaxH Then lngMaxH = lngCurH
                                                            End If
                                                        End If
                                                    Else
                                                        If strValue <> "" Then
                                                            If Not blnMeasure Then
                                                                If .行高 = 1 Then Set objFont = GetAutoFont(strValue, .W, .H, objFont, objOut, True, .网格)
                                                                If .水平反转 Then
                                                                    If Not DrawCell(objOut, frmSource.picRotate(.id).Image, .X + lngY + lngChildX, .Y + lngX + lngChildY, .W, .H, lngPaperW, lngPaperH, 0, .前景, .背景, objFont, "0000", .对齐, 0, True, sngScale, .网格) Then Exit Function
                                                                Else
                                                                    If Not DrawCell(objOut, strValue, .X + lngY + lngChildX, .Y + lngX + lngChildY, .W, .H, lngPaperW, lngPaperH, 0, .前景, .背景, objFont, IIF(.边框, "1111", "0000"), .对齐, 0, True, sngScale, .网格) Then Exit Function
                                                                End If
                                                            Else
                                                                lngCurH = .Y + .H
                                                                If lngCurH > lngPaperH Then lngCurH = lngPaperH
                                                                If lngCurH > lngMaxH Then lngMaxH = lngCurH
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End Select
                                        End If
                                    End With
                                End If
                            Next
                        Next
                    Next
                End If
            End If
        End If
    End If
    
    If Not blnMeasure Then
        '打印的实际可输出区域预览
        dblSureW = GetDeviceCaps(Printer.hDC, PHYSICALOFFSETX) / GetDeviceCaps(Printer.hDC, PHYSICALWIDTH)
        dblSureH = GetDeviceCaps(Printer.hDC, PHYSICALOFFSETY) / GetDeviceCaps(Printer.hDC, PHYSICALHEIGHT)
        If blnSure Then
            objOut.DrawStyle = 2
            objOut.Line (objOut.Width * dblSureW, objOut.Height * dblSureH)-(objOut.Width * (1 - dblSureW) - Printer.TwipsPerPixelX * sngScale, objOut.Height * (1 - dblSureH) - Printer.TwipsPerPixelY * sngScale), &H808080, B
            objOut.DrawStyle = 0
        End If
        
        '试用标志
'        strTmp = Decode(zlRegInfo("授权性质"), "2", "试用", "3", "测试", "")
'        If strTmp <> "" Then
'            Set objFont = New StdFont
'            objFont.name = "黑体"
'            objFont.Size = 24 * sngScale
'            objFont.Bold = True
'            objFont.Italic = False
'            objFont.Italic = False
'            If Not DrawCell(objOut, strTmp & "样张", objOut.Width * dblSureW + 2 * Printer.TwipsPerPixelX * sngScale, objOut.Height * dblSureH + 2 * Printer.TwipsPerPixelY * sngScale, 2500 * sngScale, 600 * sngScale, , , vbRed, vbRed, , objFont, , 1, 1, , 1) Then Exit Function
'            If Not DrawCell(objOut, strTmp & "样张", objOut.Width / 2 - 1250 * sngScale, objOut.Height / 2 - 300 * sngScale, 2500 * sngScale, 600 * sngScale, , , vbRed, vbRed, , objFont, , 1, 1, , 1) Then Exit Function
'            If Not DrawCell(objOut, strTmp & "样张", objOut.Width * (1 - dblSureW) - 2500 * sngScale - 2 * Printer.TwipsPerPixelX * sngScale, objOut.Height * (1 - dblSureH) - 600 * sngScale - 2 * Printer.TwipsPerPixelY * sngScale, 2500 * sngScale, 600 * sngScale, , , vbRed, vbRed, , objFont, , 1, 1, , 1) Then Exit Function
'        End If
    End If
    
    PrintPage = True
End Function

Public Function GetScreenFonts() As String
'功能：获取系统所支持的字体
    Dim i As Integer, strFont As String
    For i = 0 To Screen.FontCount - 1
        strFont = strFont & "^" & Screen.Fonts(i)
    Next
    GetScreenFonts = Mid(strFont, 2)
End Function

Public Function MatchIndex(ByVal cbo As Object, ByRef KeyAscii As Integer, Optional sngInterval As Single = 1) As Long
'功能：根据输入的字符串自动匹配ComboBox的选中项,并自动识别输入间隔
'参数：cbo.Hwnd=ComboBox的Hwnd属性,KeyAscii=ComboBox的KeyPress事件中的KeyAscii参数,sngInterval=指定输入间隔
'返回：-2=未加处理,其它=匹配的索引(含不匹配的索引)
'说明：请将该函数在KeyPress事件中调用。

    Static lngPreTime As Single, lngPreHwnd As Long
    Static strFind As String
    Dim sngTime As Single, lngR As Long
    
    If lngPreHwnd <> cbo.hwnd Then lngPreTime = Empty: strFind = Empty
    lngPreHwnd = cbo.hwnd
    
    If KeyAscii <> 13 Then
        sngTime = timer
        If Abs(sngTime - lngPreTime) > sngInterval Then '输入间隔(缺省为0.5秒)
            strFind = ""
        End If
        strFind = strFind & Chr(KeyAscii)
        lngPreTime = timer
        KeyAscii = 0 '使ComboBox本身的单字匹配功能失效
        MatchIndex = SendMessage(cbo.hwnd, CB_FINDSTRING, -1, ByVal strFind)
        If MatchIndex = -1 Then
            cbo.Text = strFind
            cbo.SelStart = Len(cbo.Text)
        End If
    Else
        MatchIndex = -2 '在这里对回车不作处理
    End If
End Function

Public Function ReportReaded(Optional ByVal lng报表ID As Long, _
    Optional ByVal varReport As Variant, Optional ByVal lng系统 As Long) As Boolean
'功能：判断报表缓存是否可用
'参数：lng报表ID,varReport(编号或程序ID)=用于判断当前报表缓存是否符合的条件
'      lng系统=当传入报表编号或程序ID时需要,可能为0表示共享系统
    If grsReport Is Nothing Then Exit Function
    If grsReport.State = 0 Then Exit Function
    If grsReport.EOF Or grsReport.BOF Then Exit Function
    
    If Format(grsReport!修改时间, "yyyy-MM-dd HH:mm:ss") = Format(gdatModiTime, "yyyy-MM-dd HH:mm:ss") Then
        If lng报表ID <> 0 Then
            ReportReaded = grsReport!id = lng报表ID
        Else
            If TypeName(varReport) = "String" Then
                ReportReaded = (UCase(grsReport!编号) = UCase(varReport) And Nvl(grsReport!系统, 0) = lng系统)
            Else
                ReportReaded = (Nvl(grsReport!程序id, 0) = CLng(varReport) And Nvl(grsReport!系统, 0) = lng系统)
            End If
        End If
    End If
End Function

Public Function isGroup(ByVal lngSys As Long, ByVal varReport As Variant _
    , ByRef lngID As Long, Optional ByRef strInfo As String) As Boolean
'功能：判断指定的报表是单独表还是报表组
'参数：varReport=编号或程序ID
'返回：lngID=报表或组ID

    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    '每次都读，不利用缓存，以便报表设计修改时及时更新变量gdatModiTime
    '是否报表
    If TypeName(varReport) = "String" Then
        strSQL = "Select ID,编号,名称,说明,密码,打印机,进纸,票据,打印方式,系统,程序ID,功能,修改时间,发布时间" & _
                 "    ,禁止开始时间,禁止结束时间 " & vbCr & _
                 "From zlReports " & vbCr & _
                 "Where Nvl(系统,0)=[3] And 编号=[1]"
    Else
        strSQL = "Select ID,编号,名称,说明,密码,打印机,进纸,票据,打印方式,系统,程序ID,功能,修改时间,发布时间" & _
                 "    ,禁止开始时间,禁止结束时间 " & vbCr & _
                 "From zlReports " & vbCr & _
                 "Where Nvl(系统,0)=[3] And 程序ID=[2]"
    End If
    
    Set rsTmp = OpenSQLRecord(strSQL, "isGroup", UCase(varReport), Val(varReport), lngSys)
    If Not rsTmp.EOF Then
        '缓存处理
        Set grsReport = New ADODB.Recordset
        Set grsReport = rsTmp
        gdatModiTime = grsReport!修改时间
        
        lngID = rsTmp!id
        strInfo = mdlPublic.FormatString("【[1]】[2]", Nvl(rsTmp!编号), Nvl(rsTmp!名称))
        Exit Function
    End If
    
    '是报表组
    If TypeName(varReport) = "String" Then
        strSQL = "Select ID,编号,名称 From zlRPTGroups Where Nvl(系统,0)=[3] And Upper(编号)=[1]"
    Else
        strSQL = "Select ID,编号,名称 From zlRPTGroups Where Nvl(系统,0)=[3] And 程序ID=[2]"
    End If
    Set rsTmp = OpenSQLRecord(strSQL, "isGroup", UCase(varReport), Val(varReport), lngSys)
    If Not rsTmp.EOF Then
        lngID = rsTmp!id
        strInfo = mdlPublic.FormatString("【[1]】[2]", Nvl(rsTmp!编号), Nvl(rsTmp!名称))
    End If
    isGroup = True
    Exit Function
    
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetLenStr(Str As String, lngW As Long, objBase As Object) As String
'功能：根据指定的长度截取字符串
    Dim lngTmp As Long, i As Integer
    
    For i = 1 To Len(Str)
        lngTmp = lngTmp + objBase.TextWidth(Mid(Str, i, 1))
        If lngTmp <= lngW Then
            GetLenStr = GetLenStr & Mid(Str, i, 1)
        Else
            Exit For
        End If
    Next
    If GetLenStr <> Str Then
        GetLenStr = Left(GetLenStr, Len(GetLenStr) - 1) & ".."
    End If
End Function

Public Function RemoveOrderBy(ByVal Str As String) As String
'功能：将SQL语句中最后的Order by 语句去除
    Dim i As Integer, intMax As Integer
    Dim strTmp As String
    
    strTmp = UCase(Str): intMax = -1
    Do While strTmp Like UCase("*ORDER BY*")
        i = InStr(UCase(strTmp), "ORDER BY")
        If i > intMax Then intMax = i
        strTmp = Left(strTmp, i - 1) & "12345678" & Mid(strTmp, i + 8)
    Loop
    If intMax <> -1 Then
        RemoveOrderBy = Left(Str, intMax - 1)
    Else
        RemoveOrderBy = Str
    End If
End Function

Public Function ReportCanQuery(lngRPTID As Long) As Integer
'功能：判断当前用户是否有权限对指定ID的报表进行查询
'返回：0-有权限,1-报表无权限,2-票据无权限,3-有错误
'说明：因调用报表时未传入调用位置(系统,模块),报表的多个授权位置，只要有一个位置授权即可使用
    Dim rsTmp As New ADODB.Recordset
    Dim strPriv As String, strSQL As String
    
    If gcolRptPriv Is Nothing Then
        Set gcolRptPriv = New Collection
    Else
        On Error Resume Next
        strPriv = gcolRptPriv("_" & lngRPTID)
        If Err.Number = 0 Then
            ReportCanQuery = Val(strPriv)
            Exit Function
        End If
    End If

    On Error GoTo errH

    strSQL = _
        " Select 票据,系统,程序ID,功能 From zlReports" & _
        " Where 程序ID is Not Null And 功能 Is Not Null And ID=[1]" & _
        " Union ALL" & _
        " Select A.票据,B.系统,B.程序ID,B.功能 From zlReports A,zlRPTPuts B" & _
        " Where A.ID=B.报表ID And A.ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "ReportCanQuery", lngRPTID)

    Do While Not rsTmp.EOF
        strPriv = GetPrivFunc(Nvl(rsTmp!系统, 0), rsTmp!程序id)
        If InStr(";" & strPriv & ";", ";" & rsTmp!功能 & ";") > 0 Then
            ReportCanQuery = 0: Exit Do
        Else
            ReportCanQuery = IIF(Nvl(rsTmp!票据, 0) = 0, 1, 2)
        End If
        rsTmp.MoveNext
    Loop
    
    gcolRptPriv.Add ReportCanQuery, "_" & lngRPTID
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    ReportCanQuery = 3
End Function

Public Function GetDefaultValue(ByVal strSQL As String, ByVal strFld As String _
    , Optional ByVal strDefBand As String, Optional ByVal intConnectNo As Integer = 0) As String
'功能：根据参数选择器SQL定义，返回显示字段及绑定字段的值
'参数：strFld=参数数据源字段说明串
'      strDefBand=程序传入的缺省绑定值,是否按此值过滤
'      intConnectNo=数据库连接序号；0=缺省；1>=其他
'返回：显示值|绑定值|原始记录数
    Dim rsTmp As New ADODB.Recordset
    Dim strTmp As String, i As Long
    Dim strShow As String, strBand As String
        
    '取出显示,绑定字段名
    For i = 0 To UBound(Split(strFld, "|"))
        strTmp = Split(strFld, "|")(i)
        If Split(strTmp, ",")(2) Like "*&D*" Then strShow = CStr(Split(strTmp, ",")(0))
        If Split(strTmp, ",")(2) Like "*&B*" Then strBand = CStr(Split(strTmp, ",")(0))
    Next
    If strShow = "" And strBand = "" Then Exit Function
        
    '打开参数数据源
    On Error GoTo errH
    strSQL = Replace(RemoveNote(strSQL), "[*]", "")
    Call OpenRecord(rsTmp, strSQL, "mdlPublic_GetDefaultValue", intConnectNo) '[*]在SQL的''中,类型无法处理
    i = rsTmp.RecordCount '原始记录个数
        
    '先按指定的绑定值过滤出数据行
    If Not rsTmp.EOF And strDefBand <> "" Then
        If IsType(rsTmp.Fields(strBand).type, adVarChar) Then
            rsTmp.Filter = strBand & "='" & Replace(strDefBand, "'", "''") & "'"
        ElseIf IsType(rsTmp.Fields(strBand).type, adNumeric) Then
            If Not IsNumeric(strDefBand) Then Exit Function
            rsTmp.Filter = strBand & "=" & strDefBand
        ElseIf IsType(rsTmp.Fields(strBand).type, adDBTimeStamp) Then
            If Not IsDate(strDefBand) Then Exit Function
            rsTmp.Filter = strBand & "=#" & strDefBand & "#"
        End If
    End If
    
    '再返回缺省行数据或过滤行数据
    If Not rsTmp.EOF Then
        strShow = Nvl(rsTmp.Fields(strShow).Value, "")
        strBand = Nvl(rsTmp.Fields(strBand).Value, "")
        If strShow <> "" Or strBand <> "" Then
            GetDefaultValue = strShow & "|" & strBand & "|" & i
        End If
    End If
    If GetDefaultValue = "" Then GetDefaultValue = "||1"
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
End Function

'去掉TextBox的默认右键菜单
Public Function WndMessage(ByVal hwnd As OLE_HANDLE, ByVal Msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' 如果消息不是WM_CONTEXTMENU，就调用默认的窗口函数处理
    If Msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(lngTXTProc, hwnd, Msg, wp, lp)
End Function

Public Function CheckPass(ByVal lngRPTID As Long) As Boolean
'返回假说明密码错
    Dim rsPass As New ADODB.Recordset
    Dim strPass As String, strSQL As String
            
    If ReportReaded(lngRPTID) Then
        '利用缓存
        Set rsPass = grsReport
    Else
        strSQL = "Select ID,编号,名称,说明,密码,打印机,进纸,票据,打印方式,系统,程序ID,功能,修改时间,发布时间,禁止开始时间 " & vbCrLf & _
                 "  ,禁止结束时间 " & vbCrLf & _
                 "From zlReports Where ID=[1]"
        Set rsPass = OpenSQLRecord(strSQL, "CheckPass", lngRPTID)
        If rsPass.EOF Then Exit Function
        
        '缓存处理
        Set grsReport = New ADODB.Recordset
        Set grsReport = rsPass
        gdatModiTime = grsReport!修改时间
    End If
    
    If IsNull(rsPass!密码) Then Exit Function
    strPass = GetPass(rsPass!编号, rsPass!名称)
    If strPass <> rsPass!密码 Then Exit Function
    CheckPass = True
End Function

Public Function GetPass(ByVal strCode As String, ByVal strName As String, Optional ByVal BlnSave As Boolean = False) As String
    '1-如果报表编号长度不足20位,则以空格填充
    '2-采用首位与末位异或,再与报表名称简码异或,最后再与当前位置的加密串异或的方式
    '3-如果报表编号长度超过20位,计数器复位
    Dim PStart As Integer, PEnd As Integer, PNameS As Integer, PNameE As Integer
    Dim intProcess As Integer, lngProcess As Long, strReturn As String
    
    strReturn = LCase(zlGetSymbol(strName))
    strName = IIF(strReturn = "", strName, strReturn)
    
    strReturn = ""
    intProcess = 1
    PStart = 1: PEnd = Len(strCode): PNameS = 1: PNameE = Len(strName)
    If PEnd < 20 Then strCode = strCode & String(20 - PEnd, " "): PEnd = 20
    
    Do While intProcess <= 20
        lngProcess = Asc(Mid(strCode, PStart, 1))
        lngProcess = lngProcess Xor Asc(Mid(strCode, PEnd, 1))
        lngProcess = lngProcess Xor Asc(Mid(strName, PNameS, 1))
        lngProcess = lngProcess Xor ArrayCompare(intProcess)
        
        If lngProcess < 32 Then
            lngProcess = lngProcess + 32
        ElseIf lngProcess > 127 Then
            lngProcess = lngProcess - (lngProcess - 107)
        End If
        
        If lngProcess = 34 Then
            strReturn = strReturn & """"
        ElseIf lngProcess = 39 Then
            strReturn = strReturn & IIF(BlnSave, "''", "'")
        Else
            strReturn = strReturn & Chr(lngProcess)
        End If
        
        intProcess = intProcess + 1
        PStart = PStart + 1: PEnd = PEnd - 1: PNameS = PNameS + 1
        If PNameS > PNameE Then PNameS = 1
    Loop
    GetPass = strReturn
End Function

Public Function GetCompare()
    Dim StrChange As String                     '转换串
    Dim PStart As Integer, PEnd As Integer      '位置指针
    Dim IntDO As Integer
    Dim BytThis As Byte
    
    '还原加密串
    
    StrChange = "ZL9REPORT"
    PStart = 1: PEnd = Len(StrChange)
    IntDO = 1
    
    Do While IntDO <= 20
        BytThis = ArrayCompare(IntDO)
        BytThis = BytThis Xor Asc(Mid(StrChange, PStart, 1))
        ArrayCompare(IntDO) = BytThis
        
        IntDO = IntDO + 1
        PStart = PStart + 1
        If PStart = PEnd Then PStart = 1
    Loop
End Function

Public Sub InitEnv()
    '明文"ThisProgramWriteByZT"
    ArrayCompare(1) = Asc("")
    ArrayCompare(2) = Asc("$")
    ArrayCompare(3) = Asc("P")
    ArrayCompare(4) = Asc("!")
    ArrayCompare(5) = Asc("")
    ArrayCompare(6) = 34
    ArrayCompare(7) = Asc(" ")
    ArrayCompare(8) = Asc("5")
    ArrayCompare(9) = Asc("(")
    ArrayCompare(10) = Asc("-")
    ArrayCompare(11) = Asc("T")
    ArrayCompare(12) = Asc("")
    ArrayCompare(13) = Asc("7")
    ArrayCompare(14) = Asc("9")
    ArrayCompare(15) = Asc(";")
    ArrayCompare(16) = Asc("7")
    ArrayCompare(17) = Asc("")
    ArrayCompare(18) = Asc("5")
    ArrayCompare(19) = Asc("c")
    ArrayCompare(20) = Asc("")
End Sub

Private Function GetOEM(ByVal strAsk As String) As String
    '-------------------------------------------------------------
    '功能：返回每个字线的ASCII码
    '参数：
    '返回：
    '-------------------------------------------------------------
    Dim intBit As Integer, iCount As Integer, blnCan As Boolean
    Dim strCode As String
    
    strCode = "OEM_"
    For intBit = 1 To Len(strAsk)
        '取每个字的ASCII码
        strCode = strCode & Hex(Asc(Mid(strAsk, intBit, 1)))
    Next
    GetOEM = strCode
End Function

Public Function zlGetSymbol(ByVal strInput As String, Optional ByVal bytIsWB As Byte) As String
'功能：生成字符串的简码
'入参：strInput-输入字符串；bytIsWB-是否五笔(否则为拼音)
'出参：正确返回字符串；错误返回"-"
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    If bytIsWB Then
        strSQL = "Select zlWBCode([1]) From Dual"
    Else
        strSQL = "Select zlSpellCode([1]) From Dual"
    End If
    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSQL, "zlGetSymbol", strInput)
    zlGetSymbol = Nvl(rsTmp.Fields(0).Value)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlGetSymbol = "-"
End Function

Public Function RemoveNote(ByVal strSQL As String) As String
'功能：移除SQL语句中的注释
'说明：只支持移除整行的注释
    Dim strTmp As String, i As Integer
    Dim arrLine() As String
    
    strSQL = Replace(strSQL, vbTab, " ")
    strSQL = Replace(strSQL, vbLf, vbCr)
    strSQL = Replace(strSQL, vbCr & vbCr, vbCr)
    strSQL = Replace(strSQL, vbCr & vbCr, vbCr)
    strSQL = Replace(strSQL, vbCr, vbCrLf)
    arrLine = Split(strSQL, vbCrLf)
    
    For i = 0 To UBound(arrLine)
        If Not Trim(arrLine(i)) Like "--*" Then
            RemoveNote = RemoveNote & vbCrLf & arrLine(i)
        End If
    Next
    RemoveNote = Mid(RemoveNote, 3)
End Function

Public Function ReplaceParSysNo(oldPars As RPTPars, lngSys As Long) As RPTPars
'功能：将参数集中的自定义SQL中的[系统]宏替换成正常值
    Dim i As Integer
    Dim newPars As RPTPars
    
    Call CopyPars(oldPars, newPars)
    
    For i = 1 To newPars.count
        newPars(i).明细SQL = Replace(newPars(i).明细SQL, "[系统]", lngSys)
        newPars(i).分类SQL = Replace(newPars(i).分类SQL, "[系统]", lngSys)
    Next
    Set ReplaceParSysNo = newPars
End Function

Public Function CheckParsRela(strSQL As String, ByVal objDatas As RPTDatas, ByVal strName As String _
    , Optional ByVal blnIsCheck As Boolean, Optional ByVal colValue As Collection _
    , Optional ByVal objPars As RPTPars, Optional ByRef strParName As String) As Boolean
'功能：检查是否绑定了其他参数
'      varValue=如果传入了，则表示实际的参数值
'参数：strName=SQL所属参数
'      strParName=绑定的参数名
    Dim objPar As RPTPar, objData As RPTData
      
    If InStr("Collection", TypeName(colValue)) = 0 Then Set colValue = New Collection
    If objDatas Is Nothing Then
        For Each objPar In objPars
            Call CheckParsRelaChild(strSQL, objPar, strName, colValue)
        Next
    Else
        '设计时传入数据源对象取参数
        For Each objData In objDatas
            For Each objPar In objData.Pars
                Call CheckParsRelaChild(strSQL, objPar, strName, colValue)
            Next
        Next
    End If
    If InStr(strSQL, "[=") > 0 And InStr(strSQL, "]") > 0 Then
        strParName = Mid(strSQL, InStr(strSQL, "[=") + 2)
        strParName = Mid(strParName, 1, InStr(strParName, "]") - 1)
        If blnIsCheck Then
            '将绑定的参数替换为0，保证能够正常编译
            Do While InStr(strSQL, "[=") > 0 And InStr(strSQL, "]") > 0
                strSQL = Replace(strSQL, Mid(strSQL, InStr(strSQL, "[=")), "'0'" & Mid(strSQL, InStr(strSQL, "]") + 1))
            Loop
        End If
        Exit Function
    End If
    CheckParsRela = True
End Function

Private Function CheckParsRelaChild(ByRef strSQL As String, ByVal objPar As RPTPar _
    , ByVal strName As String, Optional ByVal colValue As Collection) As Boolean
'功能：将数据源的SQL转换成Oracle可执行的SQL
'参数：
'  strSQL：数据源SQL，以及返回转换后的SQL
'  objPar：参数对象
'  strName：数据源SQL的参数名
'  colValue：参数集合对象

    Dim strTmp As String
    Dim lngTmp As Long          '0-执行； 1-SQL书写检查
    Dim bytType As Byte         '0-常规； 1-“Between ... And ...”语句
    
    If objPar.名称 <> strName Then
        If InStr(strSQL, "[=" & objPar.名称 & "]") > 0 Then
            bytType = 0
            If colValue.count = 0 Then
                strTmp = Mid(strSQL, 1, InStr(strSQL, "[=" & objPar.名称 & "]") - 1)
                strTmp = RTrim(strTmp)
                lngTmp = 0
                If InStr("=<>", Mid(strTmp, Len(strTmp))) > 0 Then
                    If Mid(strTmp, Len(strTmp) - 1) Like "[<|>][=|>]" Then
                        '运算符：不等于“<>”、大于等于“>=”、小于等于“<=”
                        strTmp = Mid(strTmp, 1, Len(strTmp) - 2)
                    Else
                        strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
                    End If
                    strTmp = RTrim(strTmp)
                    lngTmp = InStrRev(strTmp, " ")
                ElseIf (UCase(strTmp) Like "* BETWEEN" Or UCase(strTmp) Like "* BETWEEN * AND") _
                    And UCase(strSQL) Like "* BETWEEN * AND *" Then
                    lngTmp = Len(strTmp)
                    bytType = Val("1-...Between...And...")
                End If
            Else
                lngTmp = 0
            End If
            
            If lngTmp = 0 Then
                If objPar.类型 = 0 Then
                    strSQL = Replace(strSQL, "[=" & objPar.名称 & "]", "'" & IIF(colValue.count = 0, "", GetColValues(colValue, objPar.名称)) & "'")
                ElseIf objPar.类型 = 1 Then
                    strSQL = Replace(strSQL, "[=" & objPar.名称 & "]", IIF(colValue.count = 0, 0, Val(GetColValues(colValue, objPar.名称))))
                ElseIf objPar.类型 = 2 Then
                    strSQL = Replace(strSQL, "[=" & objPar.名称 & "]", IIF(colValue.count = 0, "sysdate", "to_date('" & GetColValues(colValue, objPar.名称) & "', 'YYYY-MM-DD HH24:MI:SS')"))
                ElseIf objPar.类型 = 3 Then
                    strTmp = GetColValues(colValue, objPar.名称)
                    If UCase(Trim(strTmp)) Like "IN (*)*" Then
                        strSQL = Replace(strSQL, "= [=" & objPar.名称 & "]", IIF(colValue.count = 0, "''", strTmp))
                        strSQL = Replace(strSQL, "=[=" & objPar.名称 & "]", IIF(colValue.count = 0, "''", strTmp))
                        strSQL = Replace(strSQL, "[=" & objPar.名称 & "]", IIF(colValue.count = 0, "''", strTmp))
                    Else
                        strSQL = Replace(strSQL, "[=" & objPar.名称 & "]", "'" & IIF(colValue.count = 0, "0", strTmp) & "'")
                    End If
                End If
            Else
                If bytType = 1 Then
                    Select Case objPar.类型
                    Case Val("0-字符"), Val("3-无类型")
                        If UCase(strSQL) Like "* BETWEEN [[]=" & objPar.名称 & "[]] AND *" Then
                            strSQL = Replace(strSQL, "[=" & objPar.名称 & "]", "''")
                        ElseIf UCase(strSQL) Like "* BETWEEN * AND [[]=" & objPar.名称 & "[]]*" Then
                            strSQL = Replace(strSQL, "[=" & objPar.名称 & "]", "''")
                        End If
                    Case Val("1-数值")
                        If UCase(strSQL) Like "* BETWEEN [[]=" & objPar.名称 & "[]] AND *" Then
                            strSQL = Replace(strSQL, "[=" & objPar.名称 & "]", "1")
                        ElseIf UCase(strSQL) Like "* BETWEEN * AND " & "[[]=" & objPar.名称 & "[]]*" Then
                            strSQL = Replace(strSQL, "[=" & objPar.名称 & "]", "2")
                        End If
                    Case Val("2-日期")
                        strSQL = Replace(strSQL, "[=" & objPar.名称 & "]", "sysdate")
                    End Select
                Else
                    If objPar.类型 = 0 Then
                        strSQL = Replace(strSQL, "[=" & objPar.名称 & "]", "'' Or 1=1)")
                    ElseIf objPar.类型 = 1 Then
                        strSQL = Replace(strSQL, "[=" & objPar.名称 & "]", "0 Or 1=1)")
                    ElseIf objPar.类型 = 2 Then
                        strSQL = Replace(strSQL, "[=" & objPar.名称 & "]", "sysdate Or 1=1)")
                    ElseIf objPar.类型 = 3 Then
                        strSQL = Replace(strSQL, "[=" & objPar.名称 & "]", "'0' Or 1=1)")
                    End If
                    strSQL = Mid(strSQL, 1, lngTmp) & "(" & Mid(strSQL, lngTmp + 1)
                End If
            End If
        End If
    End If
End Function

Private Function GetColValues(ByVal colValues As Collection, ByVal strParName As String) As String
'功能：获取集合中参数的值
    On Error Resume Next
    GetColValues = colValues("_" & strParName)
End Function

Public Sub GetUserName(ByVal lngSys As Long, strUserName As String, strUserNO As String)
'功能：获取登陆用户信息
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strOwner As String, strUser As String
    Dim rsData As ADODB.Recordset
    
    If gcnOracleConn <> gcnOracle.ConnectionString Then
        Set gcolUserInfo = Nothing
        gcnOracleConn = gcnOracle.ConnectionString
    End If
    If gcolUserInfo Is Nothing Then
        Set gcolUserInfo = New Collection
    Else
        On Error Resume Next
        strUser = gcolUserInfo("_" & lngSys)
        If Err.Number = 0 Then
            strUserName = Split(strUser, "_")(0)
            strUserNO = Split(strUser, "_")(1)
            glngUserID = Split(strUser, "_")(2)
            Exit Sub
        End If
    End If
    
    strUserName = gstrDBUser
    strUserNO = gstrDBUser
    
    '先假设建立了私有同义词并有权限(大部份情况)
    strSQL = _
        " Select A.姓名,A.编号,B.人员ID" & _
        " From 人员表 A,上机人员表 B,部门人员 C" & _
        " Where A.ID=B.人员ID And A.ID=C.人员ID And C.缺省=1 And B.用户名=USER"
    On Error Resume Next
    Set rsTmp = New ADODB.Recordset
    Call OpenRecord(rsTmp, strSQL, "mdlPublic_GetUserName")
    If Err.Number <> 0 And Err.Description Like "*表或视图不存在*" Then
        Err.Clear: On Error GoTo errH
        
        '再按系统所有者读取
        'Set rsTmp = New ADODB.Recordset
        strSQL = "Select 所有者 From zlSystems Where 编号=[1]"
        Set rsTmp = OpenSQLRecord(strSQL, "GetUserName", lngSys)
        If Not rsTmp.EOF Then strOwner = rsTmp!所有者 & "."
        
        strSQL = _
            " Select A.姓名,A.编号,B.人员ID" & _
            " From " & strOwner & "人员表 A," & strOwner & "上机人员表 B," & strOwner & "部门人员 C" & _
            " Where A.ID=B.人员ID And A.ID=C.人员ID And C.缺省=1 And B.用户名=USER"
        On Error Resume Next
        Set rsTmp = New ADODB.Recordset
        Call OpenRecord(rsTmp, strSQL, "mdlPublic_GetUserName")
        If Err.Number <> 0 And Err.Description Like "*表或视图不存在*" Then
            Err.Clear: On Error GoTo errH
            
            '获取用户权限对象(只获取一次)
            If grsObject Is Nothing Then Set grsObject = UserObject
            If grsObject Is Nothing Then Exit Sub
            If grsObject.State = adStateClosed Then
                Set grsObject = Nothing
                Set grsObject = UserObject
                If grsObject Is Nothing Then Exit Sub
            End If
            
            grsObject.Filter = "OBJECT_NAME='上机人员表'"
            If grsObject.EOF Then Exit Sub
            strOwner = grsObject!Owner & "."
            
            '再根据有权限的读取
            strSQL = _
                " Select A.姓名,A.编号,B.人员ID" & _
                " From " & strOwner & "人员表 A," & strOwner & "上机人员表 B," & strOwner & "部门人员 C" & _
                " Where A.ID = B.人员ID And A.ID = C.人员ID And C.缺省 = 1 And B.用户名 = USER"
            Set rsTmp = New ADODB.Recordset
            Call OpenRecord(rsTmp, strSQL, "mdlPublic_GetUserName")
        Else
            On Error GoTo errH
        End If
    Else
        On Error GoTo errH
    End If
    If Not rsTmp.EOF Then
        strUserName = rsTmp!姓名
        strUserNO = rsTmp!编号
        glngUserID = rsTmp!人员ID
    End If
    gcolUserInfo.Add strUserName & "_" & strUserNO & "_" & glngUserID, "_" & lngSys
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function ShowHelpRpt(SHwnd As Long, ByVal htmName As String, Optional Sys As Integer = 1) As Boolean
'显示帮助窗体
'SHwnd:传入窗口句柄(作为宿主窗口)
'htmName:射映在CHM中的htm文件名称
'Sys:系统,0:报表工具;1:zlhis
    Dim Path As String
    Dim strSave As String
    
    On Error GoTo ShowHelpErr
    
    ShowHelpRpt = False
    strSave = String(200, Chr$(0))
    If Sys = 0 Then
        Path = Left$(strSave, GetWindowsDirectory(strSave, Len(strSave))) + "\help\zl9server.chm"
        If Trim(Dir(Path)) = "" Then GoTo ShowHelpErr
        Call Htmlhelp(SHwnd, Path, &H0, "zlreport\" & htmName & ".htm")
    Else
    '刘兴宏:因每个报表存在帮助的情况，但目前没有相关的帮助，因此，暂取消了每个报表的有帮助的功能
    '日期:2007/09/05
'        If Mid(UCase(htmName), 5, 6) = "INSIDE" Then
            Path = Left$(strSave, GetWindowsDirectory(strSave, Len(strSave))) + "\help\zl9server.chm"
            If Trim(Dir(Path)) = "" Then GoTo ShowHelpErr
            Call Htmlhelp(SHwnd, Path, &H0, "zlreport\report.htm")
'        Else
'            Path = Left$(strSave, GetWindowsDirectory(strSave, Len(strSave))) + "\help\zl9app" & Trim(Format(Sys)) & ".chm"
'            If Trim(Dir(Path)) = "" Then GoTo ShowHelpErr
'            strSave = "zl9app" & Trim(Format(Sys)) & "rpt\" & htmName & ".htm"
'            Call Htmlhelp(SHwnd, Path, &H0, strSave)
'        End If
    End If
    ShowHelpRpt = True
    Exit Function
ShowHelpErr:
    Err.Clear
End Function

Public Function SetNTPrinterPaper(ByVal lngHWND As Long, ByVal sngWidth As Single, ByVal sngHeight As Single, _
    ByVal intOrient As Integer, ByVal intCopys As Integer, Optional ByVal blnPrompt As Boolean) As Boolean
'功能：NT环境中，设置打印机的自定义纸张尺寸
'参数：lngWidth、lngHeight=mm(毫米)
'     intOrient=1-纵向,2-横向
'     intCopys=打印份数(如果打印机支持,1-9999,不支持时不会出错,也不影响其它设置)
'说明：除了Width,Height外，其它通过本函数设置的属性不直接反映在Printer上，
'      (取DevMode也反映不出来，可能要用GetJob才能获取最近的打印文档属性)
    Dim vDevMode As DEVMODE
    Dim arrDevMode() As Byte
    Dim lngSize As Long
    
    Dim lngPrtDC As Long
    Dim lngHandle As Long
    Dim strPrtName As String
    
    lngPrtDC = Printer.hDC
    strPrtName = Printer.DeviceName
    
    If OpenPrinter(strPrtName, lngHandle, 0&) Then
        'Retrieve the size of the DEVMODE:fMode=0
        lngSize = DocumentProperties(lngHWND, lngHandle, strPrtName, 0&, 0&, 0&)
        'Reserve memory for the actual size of the DEVMODE.
        ReDim arrDevMode(1 To lngSize)
    
        'Fill the DEVMODE from the printer.
        lngSize = DocumentProperties(lngHWND, lngHandle, strPrtName, arrDevMode(1), 0&, DM_OUT_BUFFER)
        'Copy the Public (predefined) portion of the DEVMODE.
        Call CopyMemory(vDevMode, arrDevMode(1), Len(vDevMode))
        
        '设置打印文档属性
        vDevMode.dmOrientation = intOrient
        vDevMode.dmPaperSize = 256
        vDevMode.dmPaperWidth = Round(sngWidth * 10)   'in tenths of a millimeter
        vDevMode.dmPaperLength = Round(sngHeight * 10) 'in tenths of a millimeter
        vDevMode.dmCopies = intCopys
        'vDevMode.dmCollate = 0& '高级打印功能(当取消时,Copies只支持1;但不知怎么取不了)
        vDevMode.dmFields = DM_ORIENTATION Or DM_PAPERSIZE Or DM_PAPERLENGTH Or DM_PAPERWIDTH Or DM_COPIES 'Or DM_COLLATE
        
        'Copy your changes back, then update DEVMODE.
        Call CopyMemory(arrDevMode(1), vDevMode, Len(vDevMode))
        If blnPrompt Then
            lngSize = DocumentProperties(lngHWND, lngHandle, strPrtName, arrDevMode(1), arrDevMode(1), DM_IN_BUFFER Or DM_IN_PROMPT Or DM_OUT_BUFFER)
        Else
            lngSize = DocumentProperties(lngHWND, lngHandle, strPrtName, arrDevMode(1), arrDevMode(1), DM_IN_BUFFER Or DM_OUT_BUFFER)
        End If
        If lngSize = IDOK Then SetNTPrinterPaper = True
        'Reset the DEVMODE for the DC.
        lngSize = ResetDC(lngPrtDC, arrDevMode(1))
        If lngSize = 0 Then SetNTPrinterPaper = False
        
        '文本背景颜色透明模式
        Call SetBkMode(lngPrtDC, Val("1-透明"))
        
        'Close the handle when you are finished with it.
        Call ClosePrinter(lngHandle)
    End If
End Function

Public Function SetNTPrinterPaper_Form(ByVal lngHWND As Long, ByVal sngWidth As Single, ByVal sngHeight As Single, _
    ByVal intOrient As Integer, ByVal intCopys As Integer, Optional objCbo As ComboBox, _
    Optional ByVal strFormName As String, Optional objPrinter As Printer) As Boolean
'功能：NT环境中，设置打印机的自定义纸张尺寸(使用添加服务器Form方式)
'参数：lngWidth、lngHeight=mm(毫米)
'     intOrient=1-纵向,2-横向
'     intCopys=打印份数(如果打印机支持,1-9999,不支持时不会出错,也不影响其它设置)
'     objCbo=本地打印设置时，传入下拉菜单，将可用的Form加入供用户选择
'     strFormName=如果本地设置了Form时，则根据设置的Form进行打印
'说明：除了Width,Height外，其它通过本函数设置的属性不直接反映在Printer上，
'      (取DevMode也反映不出来，可能要用GetJob才能获取最近的打印文档属性)
    Dim lngSize As Long 'Size of DEVMODE
    Dim vDevMode As DEVMODE
    Dim arrDevMode() As Byte 'Working DEVMODE
    
    Dim lngPrtDC As Long 'Handle to Printer DC
    Dim lngHandle As Long 'Handle to printer
    Dim strPrtName As String
    Dim blnFormLocal As Boolean
    
    Dim vFormSize As SIZEL
    
    lngPrtDC = Printer.hDC
    strPrtName = Printer.DeviceName
    If strFormName = "" Then strFormName = ZL_FORM_NAME
    
    If OpenPrinter(strPrtName, lngHandle, 0&) Then
        'Retrieve the size of the DEVMODE.
        lngSize = DocumentProperties(lngHWND, lngHandle, strPrtName, 0&, 0&, 0&)
        'Reserve memory for the actual size of the DEVMODE.
        ReDim arrDevMode(1 To lngSize)
    
        'Fill the DEVMODE from the printer.
        lngSize = DocumentProperties(lngHWND, lngHandle, strPrtName, arrDevMode(1), 0&, DM_OUT_BUFFER)
        'Copy the Public (predefined) portion of the DEVMODE.
        Call CopyMemory(vDevMode, arrDevMode(1), Len(vDevMode))
        
        'If FormName is ZL_FORM_NAME, we must make sure it exists
        'before using it. Otherwise, it came from our EnumForms list,
        'and we do not need to check first. Note that we could have
        'passed in a Flag instead of checking for a literal name.

        'Use form ZL_FORM_NAME, adding it if necessary.
        'Set the desired size of the form needed.
        'Given in thousandths of millimeters
        vFormSize.Cx = Round(sngWidth * 1000)   'width
        vFormSize.Cy = Round(sngHeight * 1000)  'height
        
        '先删除现有的Form(如果有,因为未删掉的尺寸可能不同)
        If objCbo Is Nothing Then
            If GetFormName(lngHandle, vFormSize, strFormName) <> 0 Then
                '如果使用本地打印机Form，则不用删除后加入,直接设置
                If strFormName = ZL_FORM_NAME Then
                    If DeleteForm(lngHandle, strFormName & Chr(0)) <> 0 Then
                        '删除成功才重新加入
                        AddNewForm lngHandle, vFormSize, strFormName
                    Else
                        '未删除成功,直接利用当前Form
                        SetTheForm lngHandle, vFormSize, strFormName
                    End If
                Else
                    SetTheForm lngHandle, vFormSize, strFormName
                End If
            Else
                '没有则直接加入要用的Form
                AddNewForm lngHandle, vFormSize, strFormName
            End If
        Else
            Call GetFormName(lngHandle, vFormSize, strFormName, objCbo)
        End If
        
        If GetFormName(lngHandle, vFormSize, strFormName) = 0 Then
            Call ClosePrinter(lngHandle): Exit Function
        End If
        
        'Change the appropriate member in the DevMode.
        'In this case, you want to change the form name.
        vDevMode.dmFormName = strFormName & Chr(0)  'Must be NULL terminated!
        vDevMode.dmOrientation = intOrient
        vDevMode.dmCopies = intCopys
        'Set the dmFields bit flag to indicate what you are changing.
        vDevMode.dmFields = DM_FORMNAME Or DM_ORIENTATION Or DM_COPIES
    
        'Copy your changes back, then update DEVMODE.
        Call CopyMemory(arrDevMode(1), vDevMode, Len(vDevMode))
        lngSize = DocumentProperties(lngHWND, lngHandle, strPrtName, arrDevMode(1), arrDevMode(1), DM_IN_BUFFER Or DM_OUT_BUFFER)
        If lngSize = IDOK Then SetNTPrinterPaper_Form = True
        'Reset the DEVMODE for the DC.
        lngSize = ResetDC(lngPrtDC, arrDevMode(1))
        If lngSize = 0 Then SetNTPrinterPaper_Form = False
        If Not objPrinter Is Nothing And strFormName <> ZL_FORM_NAME Then
            '误差大于10，才重新设置width，主要是怕有些打印机使用ResetDC没有设置起Form例如TinyPDF
            If Abs(objPrinter.Width - vFormSize.Cx / 1000 * Twip_mm) > 10 Then objPrinter.Width = vFormSize.Cx / 1000 * Twip_mm
            If Abs(objPrinter.Height - vFormSize.Cy / 1000 * Twip_mm) > 10 Then objPrinter.Height = vFormSize.Cy / 1000 * Twip_mm
        End If
        
        '文本背景颜色透明模式
        Call SetBkMode(lngPrtDC, Val("1-透明"))
        
        'Close the handle when you are finished with it.
        Call ClosePrinter(lngHandle)
    End If
End Function

Public Function DelNTPrinterPaper() As Boolean
'功能：删除刚才创建的自定义纸张
    Dim lngHandle As Long
    Dim strName As String
        
    strName = Printer.DeviceName
    If OpenPrinter(strName, lngHandle, 0&) Then
        DelNTPrinterPaper = DeleteForm(lngHandle, ZL_FORM_NAME & Chr(0)) <> 0
        Call ClosePrinter(lngHandle)
    End If
End Function

Public Function GetFormName(ByVal PrinterHandle As Long, FormSize As SIZEL, ByVal FormName As String, Optional ByVal objCbo As ComboBox) As Integer
    Dim NumForms As Long, i As Long
    Dim FI1 As FORM_INFO_1
    Dim aFI1() As FORM_INFO_1           'Working FI1 array
    Dim Temp() As Byte                  'Temp FI1 array
    Dim FormIndex As Integer
    Dim BytesNeeded As Long
    Dim RetVal As Long

    'FormName = vbNullString
    FormIndex = 0
    ReDim aFI1(1)
    'First call retrieves the BytesNeeded.
    RetVal = EnumForms(PrinterHandle, 1, aFI1(0), 0&, BytesNeeded, NumForms)
    ReDim Temp(BytesNeeded)
    ReDim aFI1(BytesNeeded / Len(FI1))
    'Second call actually enumerates the supported forms.
    RetVal = EnumForms(PrinterHandle, 1, Temp(0), BytesNeeded, BytesNeeded, NumForms)
    Call CopyMemory(aFI1(0), Temp(0), BytesNeeded)
    For i = 0 To NumForms - 1
        With aFI1(i)
            If Not objCbo Is Nothing Then
                '加载可用的Form,.Flags=0的表示用户自己定义的
                 If .Flags = 0 And PtrCtoVbString(.pName) <> ZL_FORM_NAME Then
                    objCbo.AddItem PtrCtoVbString(.pName) & " " & Format(.Size.Cx / 1000, "0") & "mm(宽)×" & Format(.Size.Cy / 1000, "0") & "mm(高)"
                End If
            End If
            If PtrCtoVbString(.pName) = FormName Then '按Form名称比较
                '如果是使用本地打印机form则使用form对照的尺寸
                If FormName <> ZL_FORM_NAME Then
                    FormSize.Cx = .Size.Cx
                    FormSize.Cy = .Size.Cy
                End If
                FormIndex = i + 1
                If objCbo Is Nothing Then Exit For
            End If
        End With
    Next i
    GetFormName = FormIndex  'Returns non-zero when form is found.
End Function

Public Function SetTheForm(lngPrtHandle As Long, vFormSize As SIZEL, strFormName As String) As String
    Dim FI1 As sFORM_INFO_1
    Dim aFI1() As Byte
    Dim RetVal As Long
    
    With FI1
        .Flags = 0
        .pName = strFormName
        With .Size
            .Cx = vFormSize.Cx
            .Cy = vFormSize.Cy
        End With
        With .ImageableArea
            .Left = 0
            .Top = 0
            .Right = FI1.Size.Cx
            .Bottom = FI1.Size.Cy
        End With
    End With
    ReDim aFI1(Len(FI1))
    Call CopyMemory(aFI1(0), FI1, Len(FI1))
    
    RetVal = SetForm(lngPrtHandle, strFormName, 1, aFI1(0))
    If RetVal = 0 Then
        If Err.LastDllError = 5 Then
            MsgBox "错误:" & Err.LastDllError & vbCrLf & vbCrLf & "没有足够的权限设置自定义纸张格式。", vbExclamation, App.Title
        ElseIf Err.LastDllError = 1902 Then
            '如果用Chr(0)结尾,有时会出这个错误
            MsgBox "错误:" & Err.LastDllError & vbCrLf & vbCrLf & "指定的自定义纸张格式名称无效。", vbExclamation, App.Title
        Else
            MsgBox "错误:" & Err.LastDllError & vbCrLf & vbCrLf & "设置自定义纸张格式时发生错误。", vbExclamation, App.Title
        End If
        SetTheForm = ""
    Else
        SetTheForm = FI1.pName
    End If
End Function

Public Function AddNewForm(lngPrtHandle As Long, vFormSize As SIZEL, strFormName As String) As String
    Dim FI1 As sFORM_INFO_1
    Dim aFI1() As Byte
    Dim RetVal As Long
    
    With FI1
        .Flags = 0
        .pName = strFormName
        With .Size
            .Cx = vFormSize.Cx
            .Cy = vFormSize.Cy
        End With
        With .ImageableArea
            .Left = 0
            .Top = 0
            .Right = FI1.Size.Cx
            .Bottom = FI1.Size.Cy
        End With
    End With
    ReDim aFI1(Len(FI1))
    Call CopyMemory(aFI1(0), FI1, Len(FI1))
    RetVal = AddForm(lngPrtHandle, 1, aFI1(0))
    If RetVal = 0 Then
        If Err.LastDllError = 5 Then
            MsgBox "错误:" & Err.LastDllError & vbCrLf & vbCrLf & "没有足够的权限设置自定义纸张格式。", vbExclamation, App.Title
        ElseIf Err.LastDllError = 80 Then
            MsgBox "错误:" & Err.LastDllError & vbCrLf & vbCrLf & "指定的自定义纸张格式已经存在。", vbExclamation, App.Title
        Else
            MsgBox "错误:" & Err.LastDllError & vbCrLf & vbCrLf & "设置自定义纸张格式时发生错误。", vbExclamation, App.Title
        End If
        AddNewForm = ""
    Else
        AddNewForm = FI1.pName
    End If
End Function

Public Function PtrCtoVbString(ByVal Add As Long) As String
    Dim sTemp As String * 512, X As Long
    
    X = lstrcpy(sTemp, ByVal Add)
    If (InStr(1, sTemp, Chr(0)) = 0) Then
         PtrCtoVbString = ""
    Else
         PtrCtoVbString = Left(sTemp, InStr(1, sTemp, Chr(0)) - 1)
    End If
End Function

Public Function GetReportInfo(strFile As String) As String
'功能:获取一张外部报表的信息
'参数:strFile=外部文件名
'说明："编号;名称;版本(8/9)"
    Dim objFile As FileSystemObject, objText As TextStream
    Dim strLine As String, strSect As String, strTmp As String
    
    Set objFile = New FileSystemObject
    If Not objFile.FileExists(strFile) Then Exit Function
    Set objText = objFile.OpenTextFile(strFile)
    
    Do While Not objText.AtEndOfStream
        strLine = objText.ReadLine
        
        '判断格式是否正确
        If strSect = "" And Trim(strLine) <> "" And Trim(strLine) <> "[HEAD]" Then objText.Close: Exit Function
        
        '取得段号
        If Left(strLine, 1) = "[" And Right(strLine, 1) = "]" Then strSect = UCase(Mid(strLine, 2, Len(strLine) - 2))
        
        '处理报表头
        If strSect = "HEAD" Then
            If strLine Like "报表编号=*" Then
                strTmp = strTmp & ";" & Mid(strLine, InStr(strLine, "=") + 1)
            End If
            If strLine Like "报表名称=*" Then
                strTmp = strTmp & ";" & Mid(strLine, InStr(strLine, "=") + 1)
            End If
        ElseIf strLine = ";" Then
            If strSect = "ZLREPORTS" Then
                strTmp = strTmp & ";9"
            ElseIf strSect = "ZLREPORT" Then
                strTmp = strTmp & ";8"
            End If
            Exit Do
        End If
    Loop
    GetReportInfo = Mid(strTmp, 2)
    objText.Close
End Function

Public Function CheckFormInput(objForm As Object, Optional bln单引号 As Boolean) As Boolean
    Dim obj As Object, strText As String
    
    On Error Resume Next
    For Each obj In objForm.Controls
        If InStr("TextBox,ComboBox", TypeName(obj)) > 0 Then
            If obj.Visible And obj.Enabled Then
                Select Case TypeName(obj)
                Case "TextBox"
                    strText = obj.Text
                Case "ComboBox"
                    If obj.Style = 0 Then strText = obj.Text
                End Select
                If InStr(strText, "'") > 0 And Not bln单引号 Then
                    MsgBox "输入中存在非法字符！", vbInformation, App.Title
                    obj.SelStart = 0: obj.SelLength = Len(obj.Text)
                    obj.SetFocus: Exit Function
                End If
            End If
        End If
    Next
    CheckFormInput = True
End Function

Public Function GetEditSQL(ByVal strSQL As String, ByVal objPars As RPTPars) As String
'功能：保持格式,替换参数,返回可以直接运行的SQL
'Select * FRom 部门表 Where ID=[1]
'Select * FRom 部门表 Where ID=/*B1*/413/*E1*/
    Dim strLeft As String, strRight As String
    Dim StrPar As String, bytPar As Byte, i As Integer
    
    '字符串里的特殊字符转换
    Call mdlPublic.TransSpecialChar(strSQL)
    
    If Not objPars Is Nothing Then
        Do While InStr(strSQL, "[") > 0
            strLeft = Left(strSQL, InStr(strSQL, "[") - 1)
            strRight = Mid(strSQL, InStr(strSQL, "]") + 1)
            StrPar = Mid(strSQL, InStr(strSQL, "[") + 1, InStr(strSQL, "]") - InStr(strSQL, "[") - 1)
            If Trim(StrPar) = "" Then StrPar = 0
            bytPar = CByte(StrPar)
            
            '按缺省参数值替换
            If objPars("_" & CInt(bytPar)).缺省值 <> "" And Not objPars("_" & CInt(bytPar)).缺省值 Like "*…" Then
                Select Case objPars("_" & CInt(bytPar)).类型
                    Case 0 '字符
                        StrPar = "'" & objPars("_" & CInt(bytPar)).缺省值 & "'"
                    Case 1 '数字
                        StrPar = objPars("_" & CInt(bytPar)).缺省值
                    Case 2 '日期
                        If Left(objPars("_" & CInt(bytPar)).缺省值, 1) = "&" Then
                            StrPar = GetParSQLMacro(objPars("_" & CInt(bytPar)).缺省值)
                        Else
                            If InStr(objPars("_" & CInt(bytPar)).缺省值, ":") > 0 Then
                                '长时间格式
                                StrPar = "To_Date('" & Format(objPars("_" & CInt(bytPar)).缺省值, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            Else
                                '短时间格式
                                StrPar = "To_Date('" & Format(objPars("_" & CInt(bytPar)).缺省值, "yyyy-MM-dd") & "','YYYY-MM-DD')"
                            End If
                        End If
                    Case 3 '无类型
                        StrPar = objPars("_" & CInt(bytPar)).缺省值
                End Select
            Else '缺省值为空或为自定义项
                Select Case objPars("_" & CInt(bytPar)).类型
                    Case 0 '字符
                        StrPar = "'空串'"
                    Case 1 '数字
                        StrPar = 0
                    Case 2 '日期
                        StrPar = "Sysdate"
                    Case 3 '无类型(直接替换)
                        If objPars("_" & CInt(bytPar)).缺省值 = "固定值列表…" Then
                            '取固定值中的缺省值
                            '不好的分隔符
                            For i = 0 To UBound(Split(objPars("_" & CInt(bytPar)).值列表, "|"))
                                If Left(Split(objPars("_" & CInt(bytPar)).值列表, "|")(i), 1) = "√" Then
                                    StrPar = Split(Split(objPars("_" & CInt(bytPar)).值列表, "|")(i), ",")(1)
                                    Exit For
                                End If
                            Next
                            '没有设置缺省值则取第一个
                            If StrPar = "" Then
                                StrPar = Split(Split(objPars("_" & CInt(bytPar)).值列表, "|")(0), ",")(1)
                            End If
                        ElseIf objPars("_" & CInt(bytPar)).缺省值 = "选择器定义…" Then
                            If objPars("_" & CInt(bytPar)).值列表 <> "" Then
                                '取缺省绑定值
                                StrPar = Split(objPars("_" & CInt(bytPar)).值列表, "|")(1)
                            ElseIf objPars("_" & CInt(bytPar)).明细SQL <> "" And objPars("_" & CInt(bytPar)).明细字段 <> "" Then
                                StrPar = GetDefaultValue(objPars("_" & CInt(bytPar)).明细SQL, objPars("_" & CInt(bytPar)).明细字段)
                                If StrPar <> "" Then StrPar = CStr(Split(StrPar, "|")(1))
                                If objPars("_" & CInt(bytPar)).格式 = 1 Then
                                    StrPar = " IN (" & StrPar & ") "
                                End If
                            Else
                                StrPar = ""
                            End If
                        Else
                            StrPar = objPars("_" & CInt(bytPar)).缺省值
                        End If
                End Select
            End If
            strSQL = strLeft & "/*B" & bytPar & "*/" & StrPar & "/*E" & bytPar & "*/" & strRight
        Loop
    End If
    
    '字符串里的特殊字符还原
    Call mdlPublic.TransSpecialChar(strSQL, True)
    
    GetEditSQL = strSQL
End Function

Public Function GetParSQL(ByVal strSQL As String) As String
'功能：将SQL换成带参数的格式
'Select * FRom 部门表 Where ID=/*B1*/413/*E1*/
'Select * FRom 部门表 Where ID=[1]
    Dim strTmp As String, i As Integer
    Dim strL As String, strR As String
    Dim intMax As Integer
    
    '字符串里的特殊字符转换
    Call mdlPublic.TransSpecialChar(strSQL)
    
    On Error Resume Next
    
    strTmp = strSQL: intMax = -1
    Do While InStr(strTmp, "/*B") > 0
        strL = Left(strTmp, InStr(strTmp, "/*B") - 1)
        strR = Mid(strTmp, InStr(strTmp, "/*B") + 3)
        If Val(strR) > intMax Then intMax = Val(strR)
        strTmp = strL & strR
    Loop
    
    For i = 0 To intMax
        Do While InStr(strSQL, "/*B" & i & "*/") > 0
            strL = Left(strSQL, InStr(strSQL, "/*B" & i & "*/") - 1)
            strR = Mid(strSQL, InStr(strSQL, "/*E" & i & "*/") + Len("/*E" & i & "*/"))
            strSQL = strL & "[" & i & "]" & strR
        Loop
    Next
    
    '字符串里的特殊字符还原
    Call mdlPublic.TransSpecialChar(strSQL, True)
    
    GetParSQL = strSQL
End Function

Public Function InString(strText As String, strChars As String) As Boolean
'功能：检查在strText中是否包含strChars中指定的字符
    Dim i As Integer
    
    For i = 1 To Len(strChars)
        If InStr(strText, Mid(strChars, i, 1)) > 0 Then
            InString = True
            Exit Function
        End If
    Next
End Function

Public Function MatchString(strText As String, strChars As String) As Boolean
'功能：检查在strText中的内容是否只包含strChars中指定的字符
    Dim i As Integer
    
    For i = 1 To Len(strText)
        If InStr(strChars, Mid(strText, i, 1)) = 0 Then
            Exit Function
        End If
    Next
    
    MatchString = True
End Function

Public Function InitPar() As Boolean
'功能：系统参数初始
    On Error GoTo errH
    Static rsPar As ADODB.Recordset
    Static rsParameter As ADODB.Recordset
    Dim strSQL As String
    Dim i As Integer
    
    If rsPar Is Nothing And Not gcnOracle Is Nothing Then '静态记录集,只读取一次
        If gcnOracle.State = adStateOpen Then
            Set rsPar = New ADODB.Recordset
            strSQL = "Select 参数号,参数值 From ZLOPTIONS Where 参数号 IN(1,3)"
            Call OpenRecord(rsPar, strSQL, "mdlPublic_InitPar")
            If Not rsPar.EOF Then
                rsPar.Filter = "参数号=1"
                If Not rsPar.EOF Then gblnRunLog = Nvl(rsPar!参数值, 0) = 1
                rsPar.Filter = "参数号=3"
                If Not rsPar.EOF Then gblnErrLog = Nvl(rsPar!参数值, 0) = 1
            End If
        End If
    End If
    If Not gcnOracle Is Nothing Then
        If gcnOracle.State = adStateOpen Then
            strSQL = _
                "Select 28 参数号, zl_GetSysParameter('记录报表使用痕迹', 0, 0) 参数值 From Dual " & vbNewLine & _
                "Union All " & vbNewLine & _
                "Select 26, zl_GetSysParameter('开启报表运行日志', 0, 0) From Dual "
            Set rsParameter = OpenSQLRecord(strSQL, "")
            If rsParameter.BOF = False Then
                Do While rsParameter.EOF = False
                    Select Case rsParameter("参数号").Value
                    Case Val("26-开启报表运行日志")
                        gblnReportRunLog = Val(Nvl(rsParameter!参数值))
                    Case Val("28-记录报表使用痕迹")
                        gblnReportUse = Val(Nvl(rsParameter!参数值))
                    End Select
                    rsParameter.MoveNext
                Loop
            End If
        End If
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = IIF(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function Between(X, a, B) As Boolean
'功能：判断x是否在a和b之间
    If a < B Then
        Between = X >= a And X <= B
    Else
        Between = X >= B And X <= a
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

Public Function FormatEx(ByVal vNumber As Variant, ByVal intBit As Integer) As String
'功能：四舍五入方式格式化显示数字,保证小数点最后不出现0,小数点前要有0
'参数：vNumber=Single,Double,Currency类型的数字,intBit=最大小数位数
    Dim strNumber As String
            
    If TypeName(vNumber) = "String" Then
        If vNumber = "" Then Exit Function
        If Not IsNumeric(vNumber) Then Exit Function
        vNumber = Val(vNumber)
    End If
            
    If vNumber = 0 Then
        strNumber = 0
    ElseIf Int(vNumber) = vNumber Then
        strNumber = vNumber
    Else
        strNumber = Format(vNumber, "0." & String(intBit, "0"))
        If Left(strNumber, 1) = "." Then strNumber = "0" & strNumber
        If InStr(strNumber, ".") > 0 Then
            Do While Right(strNumber, 1) = "0"
                strNumber = Left(strNumber, Len(strNumber) - 1)
            Loop
            If Right(strNumber, 1) = "." Then strNumber = Left(strNumber, Len(strNumber) - 1)
        End If
    End If
    FormatEx = strNumber
End Function

Public Sub CboSetIndex(ByVal hWnd_combo As Long, ByVal lngindex As Long)
'功能：设置Combo控件的Index值
'为一个Combo控件选择列表项，但又不触发其Click事件
    Const CB_SETCURSEL = &H14E
    
    SendMessage hWnd_combo, CB_SETCURSEL, lngindex, 0
End Sub

Public Sub CboSetWidth(ByVal hWnd_combo As Long, ByVal lngWidth As Long)
'功能：设置Combo控件下拉列表的宽度
'此处的宽度是批下拉列表的宽度，并且是以TWIP为单位
    Const CB_SETDROPPEDWIDTH As Long = &H160

    SendMessage hWnd_combo, CB_SETDROPPEDWIDTH, lngWidth / Screen.TwipsPerPixelX, 0
End Sub

Public Sub CboSetHeight(cboControl As Object, ByVal lngHeight As Long)
'功能：设置Combo控件的下拉列表的高度
'此处的宽度是批下拉列表的高度，并且是以TWIP为单位
    SetWindowPos cboControl.hwnd, 0, 0, 0, cboControl.Width / Screen.TwipsPerPixelX, lngHeight / Screen.TwipsPerPixelY, SWP_NOMOVE
End Sub

Public Sub PressKey(bytKey As Byte)
'功能：向键盘发送一个键,类似SendKey
'参数：bytKey=VirtualKey Codes，1-254，可以用vbKeyTab,vbKeyReturn,vbKeyF4
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY, 0)
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
End Sub

Public Function GetTempPathFile(Optional ByVal strPre As String = "tmp") As String
'功能：产生一个临时文件
    Dim strPath As String, strFile As String
    
    strPath = Space(256): strFile = Space(256)
    Call GetTempPath(256, strPath)
    strPath = Left(strPath, InStr(strPath, Chr(0)) - 1)
    
    Call GetTempFileName(strPath, strPre, 0, strFile)
    strFile = Left(strFile, InStr(strFile, Chr(0)) - 1)
    
    GetTempPathFile = strFile
End Function

Public Sub CopyItem(objDest As RPTItem, objSource As RPTItem, Optional ByVal blnNew As Boolean = True)
    If blnNew Then
        Set objDest = New RPTItem
    End If
    With objDest
        .id = objSource.id
        .上级ID = objSource.上级ID
        .X = objSource.X
        .Y = objSource.Y
        .W = objSource.W
        .H = objSource.H
        .背景 = objSource.背景
        .边框 = objSource.边框
        .表头 = objSource.表头
        .参照 = objSource.参照
        .粗体 = objSource.粗体
        .对齐 = objSource.对齐
        .分栏 = objSource.分栏
        .格式 = objSource.格式
        .格式号 = objSource.格式号
        .汇总 = objSource.汇总
        .类型 = objSource.类型
        .名称 = objSource.名称
        .内容 = objSource.内容
        .排序 = objSource.排序
        .前景 = objSource.前景
        .网格 = objSource.网格
        .下线 = objSource.下线
        .斜体 = objSource.斜体
        .行高 = objSource.行高
        .性质 = objSource.性质
        .序号 = objSource.序号
        .字号 = objSource.字号
        .字体 = objSource.字体
        .自调 = objSource.自调
        .Key = objSource.Key
        Set .图片 = objSource.图片
        Set .SubIDs = objSource.SubIDs
        Set .CopyIDs = objSource.CopyIDs
    End With
End Sub

Public Sub GetChartDataName(ByVal str内容 As String, Optional strFX As String, _
    Optional strFS As String, Optional strFY As String, Optional strData As String)
'功能：根据Chart数据内容获取相关字段的名称
    Dim arrData As Variant
    
    strFX = "": strFS = "": strFY = "": strData = ""
    If str内容 <> "" Then
        arrData = Split(str内容, "|")
        
        If InStr(arrData(0), ".") > 0 Then
            strData = Split(arrData(0), ".")(0)
            strFX = Split(arrData(0), ".")(1)
        End If
        If InStr(arrData(1), ".") > 0 Then
            If strData = "" Then
                strData = Split(arrData(1), ".")(0)
            End If
            strFS = Split(arrData(1), ".")(1)
        End If
        If InStr(arrData(2), ".") > 0 Then
            If strData = "" Then
                strData = Split(arrData(2), ".")(0)
            End If
            strFY = Split(arrData(2), ".")(1)
        End If
    End If
    
    If strData Like "*（*）" Then
        strData = mdlPublic.GetStdNodeText(strData)
    End If
End Sub

Public Function SetChartDataArray(objChart As Object, rsData As ADODB.Recordset, _
    ByVal strFX As String, ByVal strFS As String, ByVal strFY As String, _
    Optional arrLabelX As Variant, Optional arrLabelS As Variant) As Boolean
'功能：设置图表数据,按照公共X轴方式
'参数：strFX=X字段,strFS=序列字段,strFY=Y字段
'返回：arrLabelX=包含X轴标签的数组
'      arrLabelS=包含序列标签的数组
    Dim colFS As New Dictionary
    Dim colFX As New Dictionary
    Dim colFY As New Dictionary
    Dim arrS As Variant, arrX As Variant, arrY As Variant
    Dim blnByDate As Boolean, strX As String, strS As String
    Dim i As Long, j As Long
    
    arrLabelX = Array()
    arrLabelS = Array()
    
    On Error GoTo errH
    
    rsData.Filter = 0
    If rsData.RecordCount = 0 Then
        SetChartDataArray = True: Exit Function
    End If
    
    blnByDate = IsType(rsData.Fields(strFX).type, adDBTimeStamp)
    For i = 1 To rsData.RecordCount
        If blnByDate Then
            strX = Format(Nvl(rsData.Fields(strFX).Value, 0), "yyyy-MM-dd HH:mm:ss")
        Else
            strX = Nvl(rsData.Fields(strFX).Value, 0)
        End If
        strS = Nvl(rsData.Fields(strFS).Value)
        
        If Not IsNull(rsData.Fields(strFS).Value) Then '不管NULL序列
            '产生序列集合
            If Not colFS.Exists("_" & strS) Then
                colFS.Add "_" & strS, strS
            End If
            
            '产生序列对应X轴点Y值集合
            If Not colFY.Exists("_" & strX & "_" & strS) Then
                colFY.Add "_" & strX & "_" & strS, Val(Nvl(rsData.Fields(strFY).Value, 0))
            Else
                '同一个序列在同一个点有多个值,则累加Y值
                colFY("_" & strX & "_" & strS) = _
                    colFY("_" & strX & "_" & strS) + Val(Nvl(rsData.Fields(strFY).Value, 0))
            End If
        End If
        
        '产生X轴点集合
        If Not colFX.Exists("_" & strX) Then
            If blnByDate Then
                colFX.Add "_" & strX, CDate(strX)
            Else
                colFX.Add "_" & strX, Val(strX)
            End If
        End If
        rsData.MoveNext
    Next
    
    With objChart.ChartGroups(1).Data
        .Layout = oc2dDataArray
        .NumSeries = colFS.count '统计序列数
        .NumPoints(1) = colFX.count '每个序列公共点数
        
        '产生X轴X值
        arrX = colFX.Items
        Call .CopyXVectorIn(1, arrX)
                
        '产生序列对应X轴点Y值
        arrS = colFS.Items
        ReDim arrY(UBound(arrX), UBound(arrS))
        For i = 0 To UBound(arrS)
            For j = 0 To UBound(arrX)
                If blnByDate Then
                    strX = Format(arrX(j), "yyyy-MM-dd HH:mm:ss")
                Else
                    strX = arrX(j)
                End If
                If colFY.Exists("_" & strX & "_" & arrS(i)) Then
                    arrY(j, i) = colFY("_" & strX & "_" & arrS(i))
                Else
                    arrY(j, i) = .HoleValue '该序列不存在的X点
                End If
            Next
        Next
        Call .CopyYArrayIn(arrY)
    End With
    
    arrLabelX = arrX
    arrLabelS = arrS
    
    SetChartDataArray = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function SetChartDataGeneral(objChart As Object, rsData As ADODB.Recordset, _
    ByVal strFX As String, ByVal strFS As String, ByVal strFY As String, Optional arrLabelS As Variant) As Boolean
'功能：设置图表数据,按照公共X轴方式
'参数：strFX=X字段,strFS=序列字段,strFY=Y字段
'返回：arrLabelX=包含X轴标签的数组
'      arrLabelS=包含序列标签的数组
    Dim colFS As New Dictionary
    Dim arrS As Variant, arrX As Variant, arrY As Variant
    Dim i As Long, j As Long
    
    arrLabelS = Array()
    
    On Error GoTo errH
    
    rsData.Filter = 0
    If rsData.RecordCount = 0 Then
        SetChartDataGeneral = True: Exit Function
    End If
    
    For i = 1 To rsData.RecordCount
        If Not IsNull(rsData.Fields(strFS).Value) Then '不管NULL序列
            If Not colFS.Exists("_" & rsData.Fields(strFS).Value) Then
                colFS.Add "_" & rsData.Fields(strFS).Value, rsData.Fields(strFS).Value
            End If
        End If
        rsData.MoveNext
    Next
    
    With objChart.ChartGroups(1).Data
        .Layout = oc2dDataGeneral
        .NumSeries = colFS.count '统计序列数
        arrS = colFS.Items
        For i = 0 To UBound(arrS)
            rsData.Filter = strFS & "='" & arrS(i) & "'"
            .NumPoints(i + 1) = rsData.RecordCount '当前序列点数
            
            '产生当前序列对应的X,Y值
            ReDim arrX(rsData.RecordCount - 1)
            ReDim arrY(rsData.RecordCount - 1)
            For j = 1 To rsData.RecordCount
                If Not IsNull(rsData.Fields(strFX).Value) Then
                    arrX(j - 1) = rsData.Fields(strFX).Value
                Else
                    arrX(j - 1) = .HoleValue
                End If
                If Not IsNull(rsData.Fields(strFY).Value) Then
                    arrY(j - 1) = rsData.Fields(strFY).Value
                Else
                    arrY(j - 1) = .HoleValue
                End If
                rsData.MoveNext
            Next
            Call .CopyXVectorIn(i + 1, arrX)
            Call .CopyYVectorIn(i + 1, arrY)
        Next
    End With
    
    arrLabelS = arrS
    
    SetChartDataGeneral = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub SetChartStyleAndData(objChart As Object, objItem As RPTItem, _
    Optional rsData As ADODB.Recordset, Optional ByVal sngScale As Single = 1, _
    Optional ByVal blnDesign As Boolean, Optional ByVal blnNoDataEmpty As Boolean)
'功能：根据当前设置的样式属性设置图表控件的样式@@@
'参数：objChart=图表控件,objItem=图表元素对象,rsData=包含数据的记录集
'      sngScale=显示比例,blnDesign=是否在报表设计环境中使用
    Dim arrTmp As Variant, strTmp As String
    Dim strFX As String, strFY As String, strFS As String
    Dim arrLabelX As Variant, arrLabelS As Variant
    Dim blnByDate As Boolean, i As Long, j As Long
        
    '图表标题
    If objItem.表头 <> "" Then
        arrTmp = Split(objItem.表头, "|")
        objChart.Header.Text = arrTmp(0)
        arrTmp = Split(arrTmp(1), ",")
        objChart.Header.Font.name = CStr(arrTmp(0))
        objChart.Header.Font.Size = Val(arrTmp(1)) * sngScale
        objChart.Header.Font.Bold = Val(arrTmp(2)) <> 0
        objChart.Header.Font.Italic = Val(arrTmp(3)) <> 0
    Else
        objChart.Header.Text = ""
    End If
    
    '图表类型:散点图因为数据处理方式不一样,所以要独立出来
    '0-Plot(散点图),1-Plot(折线图),2-Bar(条形图),3-Pie(饼图),4-StackingBar(层叠图),5-Area(面积图)
    '6-HiLo(股价图-盘高,盘低),7-HiLoOpenClose(股价图-盘高,盘低,开盘,收盘),8-Candle(股价图-阴阳烛图:盘高,盘低,开盘,收盘)
    '9-Polar(级线图),10-Radar(雷达图),11-FilledRadar(填充雷达图),12-Bubble(气泡图)
    objChart.ChartGroups(1).ChartType = IIF(objItem.序号 = 0, 1, objItem.序号)
    
    '数据内容
    Call GetChartDataName(objItem.内容, strFX, strFS, strFY)
    objChart.ChartArea.Axes("X").Title.Text = strFX '字段名作为XY轴标题
    objChart.ChartArea.Axes("Y").Title.Text = strFY
    
    '绑定数据
    arrLabelX = Array(): arrLabelS = Array()
    If Not rsData Is Nothing Then
        blnByDate = IsType(rsData.Fields(strFX).type, adDBTimeStamp)  '仅X轴支持日期/时间类型
        objChart.IsBatched = True
        objChart.ChartGroups(1).Data.NumSeries = 0
        If objItem.序号 = 0 Then
            '按DataGeneral方式设置数据
            Call SetChartDataGeneral(objChart, rsData, strFX, strFS, strFY, arrLabelS)
        Else
            '按DataArray方式设置数据
            Call SetChartDataArray(objChart, rsData, strFX, strFS, strFY, arrLabelX, arrLabelS)
        End If
        objChart.IsBatched = False '刷新内部数据
    Else
        If blnNoDataEmpty Then
            objChart.ChartGroups(1).Data.NumSeries = 0
        Else
            '初始数据状态,缺省是按Array
            If objItem.序号 <> 0 Then
                For i = 1 To objChart.ChartGroups(1).Data.NumPoints(1)
                    objChart.ChartGroups(1).Data.X(1, i) = i
                Next
            Else
                '散点图特殊显示
                objChart.ChartGroups(1).Data.X(1, 1) = 3
                objChart.ChartGroups(1).Data.X(1, 2) = 2
                objChart.ChartGroups(1).Data.X(1, 3) = 5
                objChart.ChartGroups(1).Data.X(1, 4) = 4
                objChart.ChartGroups(1).Data.X(1, 5) = 1
            End If
        End If
    End If
    If blnByDate Then '日期/时间类型时旋转显示X轴刻度
        objChart.ChartArea.Axes("X").AnnotationMethod = oc2dAnnotateTimeLabels
        objChart.ChartArea.Axes("X").AnnotationRotationAngle = -90
    Else
        objChart.ChartArea.Axes("X").AnnotationMethod = oc2dAnnotateValues
        objChart.ChartArea.Axes("X").AnnotationRotationAngle = 0
    End If
    
    objChart.IsBatched = True
    
    '图例
    objChart.ChartGroups(1).SeriesLabels.RemoveAll
    objChart.ChartGroups(1).PointLabels.RemoveAll
    If objItem.分栏 <= 1 Then
        objChart.Legend.IsShowing = False
    Else
        '对齐:右=1,左=2,上=16,右上=17,左上=18,下=32,右下=33,左下=34
        objChart.Legend.Anchor = Decode(objItem.对齐, 0, 1, 1, 32, 2, 2, 3, 16, 4, 33, 5, 34, 6, 17, 7, 18)
        '左右对齐时竖排,其它情况横排
        objChart.Legend.Orientation = Decode(objItem.对齐, 0, 1, 2, 1, 2)
        objChart.Legend.IsShowing = True
                                                    
        '序列图例
        If UBound(arrLabelS) <> -1 Then
            For i = 0 To UBound(arrLabelS)
                objChart.ChartGroups(1).SeriesLabels.Add arrLabelS(i)
            Next
        Else
            For i = 1 To objChart.ChartGroups(1).Styles.count
                If strFS <> "" Then
                    objChart.ChartGroups(1).SeriesLabels.Add strFS & i
                Else
                    objChart.ChartGroups(1).SeriesLabels.Add "序列" & i
                End If
            Next
        End If
        
        'X轴点注:目前只有饼图有效
        If objChart.ChartGroups(1).ChartType = 3 Then 'oc2dTypePie
            If UBound(arrLabelX) <> -1 Then
                For i = 0 To UBound(arrLabelX)
                    If blnByDate Then
                        strTmp = Format(arrLabelX(i), "yyyy-MM-dd HH:mm:ss")
                        strTmp = Replace(strTmp, " 00:00:00", "")
                        strTmp = Replace(strTmp, ":00:00", "")
                        strTmp = Replace(strTmp, ":00", "")
                    Else
                        strTmp = arrLabelX(i)
                    End If
                    objChart.ChartGroups(1).PointLabels.Add strTmp
                Next
            ElseIf objItem.序号 <> 0 Then 'General方式无意义
                If objChart.ChartGroups(1).Data.Layout = oc2dDataArray Then
                    For i = 1 To objChart.ChartGroups(1).Data.NumPoints(1)
                        If strFX <> "" Then
                            objChart.ChartGroups(1).PointLabels.Add strFX & objChart.ChartGroups(1).Data.X(1, i)
                        Else
                            objChart.ChartGroups(1).PointLabels.Add "点" & objChart.ChartGroups(1).Data.X(1, i)
                        End If
                    Next
                ElseIf objChart.ChartArea.Axes("X").DataMin <> objChart.ChartGroups(1).Data.HoleValue Then
                    For i = objChart.ChartArea.Axes("X").DataMin To objChart.ChartArea.Axes("X").DataMax
                        If strFX <> "" Then
                            objChart.ChartGroups(1).PointLabels.Add strFX & i
                        Else
                            objChart.ChartGroups(1).PointLabels.Add "点" & i
                        End If
                    Next
                End If
            End If
        End If
    End If
    
    '连线和结点
    ReDim arrTmp(1 To 12) As Integer
    arrTmp(1) = 2 'oc2dShapeDot(实心圆)
    arrTmp(2) = 4 'oc2dShapeTriangle(正实心三角)
    arrTmp(3) = 3 'oc2dShapeBox(实心正方形)
    arrTmp(4) = 5 'oc2dShapeDiamond(实心菱)
    arrTmp(5) = 6 'oc2dShapeStar(星号)
    arrTmp(6) = 13 'oc2dShapeDiagonalCross(叉叉)
    arrTmp(7) = 12 'oc2dShapeInvertTriangle(反实心三角)
    arrTmp(8) = 14 'oc2dShapeOpenTriangle(正空心三角)
    arrTmp(9) = 11 'oc2dShapeSquare(空心正方形)
    arrTmp(10) = 10 'oc2dShapeCircle(空心圆)
    arrTmp(11) = 15 'oc2dShapeOpenDiamond(空心菱)
    arrTmp(12) = 16 'oc2dShapeOpenInvertTriangle(空心反三角)
    'arrTmp(13) = 9 'oc2dShapeCross(加号)
    'arrTmp(14) = 8 'oc2dShapeHorizontalLine(横线)
    'arrTmp(15) = 7 'oc2dShapeVerticalLine(坚线)
    For i = 1 To objChart.ChartGroups(1).Styles.count
        If objItem.自调 Then
            '依优先级循环显示结点类型
            objChart.ChartGroups(1).Styles(i).Symbol.Shape = arrTmp(((i - 1) Mod UBound(arrTmp)) + 1)
            objChart.ChartGroups(1).Styles(i).Symbol.Size = 7 * sngScale
        Else
            objChart.ChartGroups(1).Styles(i).Symbol.Shape = oc2dShapeNone
        End If
        objChart.ChartGroups(1).Styles(i).Line.Pattern = IIF(objItem.下线, 2, 1) 'oc2dLineSolid/oc2dLineNone
        objChart.ChartGroups(1).Styles(i).Line.Width = 1 * sngScale
    Next
    
    '其它格式：数字位串,三维效果|XY轴互换
    '三维效果
    If Val(Mid(Format(objItem.格式, "00"), 1, 1)) <> 0 Then
        Select Case objItem.序号
            Case 1, 5 '折线图,面积图
                strTmp = "30,20,10"
            Case 2, 4 '条线图,层叠图
                strTmp = "10,10,10"
            Case 3 '饼图
                strTmp = "20,20,0"
            Case Else
                strTmp = "0,0,0"
        End Select
    Else
        strTmp = "0,0,0"
    End If
    '几个值不能乘以比例,控件是自动的
    objChart.ChartArea.View3D.Depth = Val(Split(strTmp, ",")(0))  '深度
    objChart.ChartArea.View3D.Elevation = Val(Split(strTmp, ",")(1))  '高度
    objChart.ChartArea.View3D.Rotation = Val(Split(strTmp, ",")(2)) '角度
    objChart.ChartArea.View3D.Shading = oc2dShadingColor
    'XY轴互换
    objChart.ChartArea.IsHorizontal = Val(Mid(Format(objItem.格式, "00"), 2, 1)) <> 0
    
    '图表网格
    If objItem.网格 <> 0 Then
        objChart.ChartArea.Axes("X").MajorGrid.Spacing.IsDefault = True
        objChart.ChartArea.Axes("Y").MajorGrid.Spacing.IsDefault = True
        
        objChart.ChartArea.Axes("X").MajorGrid.Style.Width = 1 * sngScale
        objChart.ChartArea.Axes("Y").MajorGrid.Style.Width = 1 * sngScale
    Else
        objChart.ChartArea.Axes("X").MajorGrid.Spacing.Value = 0
        objChart.ChartArea.Axes("Y").MajorGrid.Spacing.Value = 0
    End If
    objChart.ChartArea.Axes("X").AxisStyle.LineStyle.Width = 1 * sngScale
    objChart.ChartArea.Axes("Y").AxisStyle.LineStyle.Width = 1 * sngScale
    
    '图表颜色
    objChart.Interior.BackgroundColor = IIF(objItem.背景 = RGB(255, 255, 255) And blnDesign, &HEFEFEF, objItem.背景)
    objChart.Interior.ForegroundColor = objItem.前景
    '不知为什么仅设置控件前景无效,但通过属性框就有效
    objChart.ChartArea.Axes("X").AxisStyle.LineStyle.Color = objItem.前景
    objChart.ChartArea.Axes("Y").AxisStyle.LineStyle.Color = objItem.前景
        
    '图表字体
    objChart.Legend.Font.name = objItem.字体
    objChart.Legend.Font.Size = objItem.字号 * sngScale
    objChart.Legend.Font.Bold = objItem.粗体
    objChart.Legend.Font.Italic = objItem.斜体
    
    objChart.ChartArea.Axes("X").Font.name = objItem.字体 'Y轴同步变化
    objChart.ChartArea.Axes("X").Font.Size = objItem.字号 * sngScale
    objChart.ChartArea.Axes("X").Font.Bold = objItem.粗体
    objChart.ChartArea.Axes("X").Font.Italic = objItem.斜体
    
    objChart.ChartArea.Axes("X").TitleFont.name = objItem.字体 'Y轴同步变化
    objChart.ChartArea.Axes("X").TitleFont.Size = objItem.字号 * sngScale
    objChart.ChartArea.Axes("X").TitleFont.Bold = objItem.粗体
    objChart.ChartArea.Axes("X").TitleFont.Italic = objItem.斜体
    
    objChart.IsBatched = False
End Sub

Public Function GetChartPicture(objDesc As Object, objSource As Object, objItem As RPTItem, _
    Optional rsData As ADODB.Recordset, Optional ByVal sngScale As Single = 1) As StdPicture
'功能：按拷贝图表对象,按比例缩放,并获取相应的图表图片
    Dim strFX As String, strFY As String, strFS As String
    Dim arrX As Variant, arrY As Variant
    Dim blnByDate As Date, strFile As String, i As Long
        
    objDesc.Left = 0
    objDesc.Top = 0
    objDesc.Width = objSource.Width * sngScale
    objDesc.Height = objSource.Height * sngScale
        
    '图表标题
    objDesc.Header.Text = objSource.Header.Text
    objDesc.Header.Font.name = objSource.Header.Font.name
    objDesc.Header.Font.Size = objSource.Header.Font.Size * sngScale
    objDesc.Header.Font.Bold = objSource.Header.Font.Bold
    objDesc.Header.Font.Italic = objSource.Header.Font.Italic
    
    '图表类型
    objDesc.ChartGroups(1).ChartType = objSource.ChartGroups(1).ChartType
    
    '数据内容
    objDesc.ChartArea.Axes("X").Title.Text = objSource.ChartArea.Axes("X").Title.Text
    objDesc.ChartArea.Axes("Y").Title.Text = objSource.ChartArea.Axes("Y").Title.Text
    objDesc.ChartArea.Axes("X").AnnotationMethod = objSource.ChartArea.Axes("X").AnnotationMethod
    objDesc.ChartArea.Axes("X").AnnotationRotationAngle = objSource.ChartArea.Axes("X").AnnotationRotationAngle
    blnByDate = objDesc.ChartArea.Axes("X").AnnotationMethod = oc2dAnnotateTimeLabels
    
    '绑定数据
    '-----------------------------------------------------------------------------------------------
    objDesc.IsBatched = True
    
    '用文件交换时不对
'    strFile = GetTempPathFile
'    Call objSource.Save(strFile)
'    Call objDesc.Load(strFile)
'    Kill strFile
    
    Call GetChartDataName(objItem.内容, strFX, strFS, strFY)
    If strFX <> "" And strFS <> "" And strFY <> "" And Not rsData Is Nothing Then
        '使用记录集绑定比较快
        objDesc.ChartGroups(1).Data.NumSeries = 0
        If objItem.序号 = 0 Then
            Call SetChartDataGeneral(objDesc, rsData, strFX, strFS, strFY)
        Else
            Call SetChartDataArray(objDesc, rsData, strFX, strFS, strFY)
        End If
    Else
        '使用数组拷贝时,如果系列较多,会比较慢
        objDesc.ChartGroups(1).Data.Layout = objSource.ChartGroups(1).Data.Layout
        objDesc.ChartGroups(1).Data.NumSeries = 0
        objDesc.ChartGroups(1).Data.NumSeries = objSource.ChartGroups(1).Data.NumSeries
        If objDesc.ChartGroups(1).Data.NumSeries > 0 Then
            If objSource.ChartGroups(1).Data.Layout = oc2dDataArray Then
                objDesc.ChartGroups(1).Data.NumPoints(1) = objSource.ChartGroups(1).Data.NumPoints(1)
        
                If blnByDate Then
                    ReDim arrX(objDesc.ChartGroups(1).Data.NumPoints(1) - 1) As Date
                Else
                    ReDim arrX(objDesc.ChartGroups(1).Data.NumPoints(1) - 1) As Double
                End If
                Call objSource.ChartGroups(1).Data.CopyXVectorOut(1, arrX)
                Call objDesc.ChartGroups(1).Data.CopyXVectorIn(1, arrX)
        
                ReDim arrY(objDesc.ChartGroups(1).Data.NumPoints(1) - 1, objDesc.ChartGroups(1).Data.NumSeries - 1) As Double
                Call objSource.ChartGroups(1).Data.CopyYArrayOut(arrY)
                Call objDesc.ChartGroups(1).Data.CopyYArrayIn(arrY)
            Else
                For i = 1 To objSource.ChartGroups(1).Data.NumSeries
                    objDesc.ChartGroups(1).Data.NumPoints(i) = objSource.ChartGroups(1).Data.NumPoints(i)
        
                    If blnByDate Then
                        ReDim arrX(objDesc.ChartGroups(1).Data.NumPoints(i) - 1) As Date
                    Else
                        ReDim arrX(objDesc.ChartGroups(1).Data.NumPoints(i) - 1) As Double
                    End If
                    Call objSource.ChartGroups(1).Data.CopyXVectorOut(i, arrX)
                    Call objDesc.ChartGroups(1).Data.CopyXVectorIn(i, arrX)
        
                    ReDim arrY(objDesc.ChartGroups(1).Data.NumPoints(i) - 1) As Double
                    Call objSource.ChartGroups(1).Data.CopyYVectorOut(i, arrY)
                    Call objDesc.ChartGroups(1).Data.CopyYVectorIn(i, arrY)
                Next
            End If
        End If
    End If
    objDesc.IsBatched = False
    '-----------------------------------------------------------------------------------------------
    objDesc.IsBatched = True
    
    '图例
    objDesc.ChartGroups(1).SeriesLabels.RemoveAll
    objDesc.ChartGroups(1).PointLabels.RemoveAll
    
    objDesc.Legend.Anchor = objSource.Legend.Anchor
    objDesc.Legend.Orientation = objSource.Legend.Orientation
    objDesc.Legend.IsShowing = objSource.Legend.IsShowing
    
    For i = 1 To objSource.ChartGroups(1).SeriesLabels.count
        objDesc.ChartGroups(1).SeriesLabels.Add objSource.ChartGroups(1).SeriesLabels(i).Text
    Next
    For i = 1 To objSource.ChartGroups(1).PointLabels.count
        objDesc.ChartGroups(1).PointLabels.Add objSource.ChartGroups(1).PointLabels(i).Text
    Next
    
    '连线和结点
    For i = 1 To objDesc.ChartGroups(1).Styles.count
        objDesc.ChartGroups(1).Styles(i).Symbol.Shape = objSource.ChartGroups(1).Styles(i).Symbol.Shape
        objDesc.ChartGroups(1).Styles(i).Symbol.Size = objSource.ChartGroups(1).Styles(i).Symbol.Size * sngScale
        objDesc.ChartGroups(1).Styles(i).Line.Pattern = objSource.ChartGroups(1).Styles(i).Line.Pattern
        objDesc.ChartGroups(1).Styles(i).Line.Width = objSource.ChartGroups(1).Styles(i).Line.Width * sngScale
    Next
    
    '其它格式：数字位串,三维效果|XY轴互换
    '几个值不能乘以比例,控件是自动的
    objDesc.ChartArea.View3D.Depth = objSource.ChartArea.View3D.Depth
    objDesc.ChartArea.View3D.Elevation = objSource.ChartArea.View3D.Elevation
    objDesc.ChartArea.View3D.Rotation = objSource.ChartArea.View3D.Rotation
    objDesc.ChartArea.View3D.Shading = objSource.ChartArea.View3D.Shading
    'XY轴互换
    objDesc.ChartArea.IsHorizontal = objSource.ChartArea.IsHorizontal
    
    '图表网格
    objDesc.ChartArea.Axes("X").MajorGrid.Spacing.IsDefault = objSource.ChartArea.Axes("X").MajorGrid.Spacing.IsDefault
    objDesc.ChartArea.Axes("Y").MajorGrid.Spacing.IsDefault = objSource.ChartArea.Axes("Y").MajorGrid.Spacing.IsDefault
    objDesc.ChartArea.Axes("X").MajorGrid.Style.Width = objSource.ChartArea.Axes("X").MajorGrid.Style.Width * sngScale
    objDesc.ChartArea.Axes("Y").MajorGrid.Style.Width = objSource.ChartArea.Axes("Y").MajorGrid.Style.Width * sngScale
    objDesc.ChartArea.Axes("X").AxisStyle.LineStyle.Width = objSource.ChartArea.Axes("X").AxisStyle.LineStyle.Width * sngScale
    objDesc.ChartArea.Axes("Y").AxisStyle.LineStyle.Width = objSource.ChartArea.Axes("Y").AxisStyle.LineStyle.Width * sngScale
    
    '图表颜色
    objDesc.Interior.BackgroundColor = objSource.Interior.BackgroundColor
    objDesc.Interior.ForegroundColor = objSource.Interior.ForegroundColor
    objDesc.ChartArea.Axes("X").AxisStyle.LineStyle.Color = objSource.ChartArea.Axes("X").AxisStyle.LineStyle.Color
    objDesc.ChartArea.Axes("Y").AxisStyle.LineStyle.Color = objSource.ChartArea.Axes("Y").AxisStyle.LineStyle.Color
        
    '图表字体
    objDesc.Legend.Font.name = objSource.Legend.Font.name
    objDesc.Legend.Font.Size = objSource.Legend.Font.Size * sngScale
    objDesc.Legend.Font.Bold = objSource.Legend.Font.Bold
    objDesc.Legend.Font.Italic = objSource.Legend.Font.Italic
    
    objDesc.ChartArea.Axes("X").Font.name = objSource.ChartArea.Axes("X").Font.name
    objDesc.ChartArea.Axes("X").Font.Size = objSource.ChartArea.Axes("X").Font.Size * sngScale
    objDesc.ChartArea.Axes("X").Font.Bold = objSource.ChartArea.Axes("X").Font.Bold
    objDesc.ChartArea.Axes("X").Font.Italic = objSource.ChartArea.Axes("X").Font.Italic
    
    objDesc.ChartArea.Axes("X").TitleFont.name = objSource.ChartArea.Axes("X").TitleFont.name
    objDesc.ChartArea.Axes("X").TitleFont.Size = objSource.ChartArea.Axes("X").TitleFont.Size * sngScale
    objDesc.ChartArea.Axes("X").TitleFont.Bold = objSource.ChartArea.Axes("X").TitleFont.Bold
    objDesc.ChartArea.Axes("X").TitleFont.Italic = objSource.ChartArea.Axes("X").TitleFont.Italic
    
    objDesc.IsBatched = False
    
    strFile = gobjFile.GetSpecialFolder(TemporaryFolder) & "\" & gobjFile.GetTempName
    If objDesc.SaveImageAsJpeg(strFile, 100, False, False, False) Then
        Set GetChartPicture = LoadPicture(strFile)
    End If
    If gobjFile.FileExists(strFile) Then
        Call gobjFile.DeleteFile(strFile, True)
    End If
End Function

Public Function ChartInstall() As Boolean
'功能：判断Chart控件是否已经安装,如果未安装则自动注册
'返回：已安装或未安装但注册成功返回True
'      未安装且注册失败返回False
    Dim objTest As Control
    Static blnInstall As Boolean
    
    If Not blnInstall Then
        On Error Resume Next
        
        Set objTest = frmFlash.Controls.Add("C1Chart2D8.Control.1", "ChartTest")
        If Err.Number <> 0 Then
            Unload frmFlash: Err.Clear
            Call Shell("c1regsvr.exe olch2x8.ocx", vbHide)
            If Err.Number <> 0 Then
                MsgBox "图表控件未正确注册，重新安装HIS客户端可以解决这个问题。", vbExclamation, App.Title
                Exit Function
            End If
        Else
            Unload frmFlash
        End If
        blnInstall = True
    End If
    ChartInstall = True
End Function

Public Sub SQLTest(Optional ByVal strProject As String, Optional ByVal strForm As String, Optional ByVal strSQL As String, Optional ByVal strNote As String)
'功能：将部件中执行的SQL语句输出到窗体或文件中，并附加开始结束时间，执行时间
'参数：strProject=部件名称,具体可取App.Title
'      strForm=窗体名,具体可取Form.Caption
'      strSQL=将要执行的SQL语句,在Open时传入,如果不传，表示最近一次SQL执行完毕
'      strNote=SQL语句说明
    Dim strTmp As String, sngEnd As Single
    
    If gblnExeSQLTest Then Exit Sub
    mstrRecentSQL = strSQL  '保存最近执行的SQL语句
    
    If gobjRegister.GetServerName = "SQLLOG" Then
        If strSQL <> "" Then
            If mobjLogText Is Nothing Then
                On Local Error Resume Next
                Set mobjLogText = gobjFile.OpenTextFile("ReportSQL_" & gstrDBUser & "_" & Format(Date, "yyyyMMdd") & ".log", ForAppending, True, TristateFalse)
                On Local Error GoTo 0
            End If
            If Not mobjLogText Is Nothing Then
                strTmp = "[" & Format(Time, "HH:mm:ss") & "]"
                mobjLogText.WriteLine strTmp & "Application:" & strProject & "\" & strForm & IIF(strNote <> "", "," & strNote, "")
                mobjLogText.WriteLine strTmp & "SQL:" & strSQL
                msngTime = timer
            End If
        Else
            If Not mobjLogText Is Nothing Then
                sngEnd = timer
                strTmp = "[" & Format(Time, "HH:mm:ss") & "]"
                mobjLogText.WriteLine strTmp & "Expend:" & Format(sngEnd - msngTime, "0.0000")
                mobjLogText.WriteBlankLines 1
            End If
        End If
    End If
End Sub

Public Function OpenRecord(rsTmp As ADODB.Recordset, ByVal strSQL As String, ByVal strTitle As String, _
    Optional ByVal intConnect As Integer = 0, _
    Optional ByVal CursorType As CursorTypeEnum = adOpenKeyset, _
    Optional ByVal LockType As LockTypeEnum = adLockReadOnly) As ADODB.Recordset
    
    Dim cnOracle As ADODB.Connection
    
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    Call SQLTest(App.ProductName, strTitle, strSQL)
    Set cnOracle = mdlPublic.GetDBConnection(intConnect)
    rsTmp.Open strSQL, cnOracle, CursorType, LockType
    Call SQLTest
    
    Set rsTmp.ActiveConnection = Nothing
    Set OpenRecord = rsTmp
End Function

Public Function TrimEx(ByVal strText As String, Optional ByVal blnCrlf As Boolean) As String
'功能：去掉TAB字符，两边空格，回车，最后只由单空格分隔。
'说明：主要是RunSQLFile的子函数
    Dim i As Long
    
    If blnCrlf Then
        strText = Replace(strText, vbCrLf, " ")
        strText = Replace(strText, vbCr, " ")
        strText = Replace(strText, vbLf, " ")
    End If
    strText = Trim(Replace(strText, vbTab, " "))
    
    i = 5
    Do While i > 1
        strText = Replace(strText, String(i, " "), " ")
        If InStr(strText, String(i, " ")) = 0 Then i = i - 1
    Loop
    TrimEx = strText
End Function

Public Function GetDBConnectionEx(ByVal intDeviceType As Integer, ByVal Index As Integer) As ADODB.Connection
'功能：获取指定驱动的连接对象
'参数：
'  intDeviceType：0-MicroSoft ODBC；1-Oralce OLEDB
'  Index：数据连接编号

    Dim cn As New ADODB.Connection
    Dim strKey As String, strPass As String, strServer As String
    
    On Error GoTo hErr
    
    strKey = "_" & Index
        
    '数据连接
    If grsConnect.State = adStateOpen Then
        With grsConnect
            If .RecordCount > 0 Then .MoveFirst
            Do While .EOF = False
                If Nvl(!编号, 0) = Index Then
                    '初始化连接对象
                    strServer = Nvl(!IP) & _
                                IIF(Nvl(!端口) = "", ":1521", ":" & Nvl(!端口)) & _
                                IIF(Nvl(!实例名) = "", "", "/" & Nvl(!实例名))
                    strPass = Nvl(!密码)
                    '解密
                    strPass = mdlPublic.Decipher(MSTR_DBLINK_KEY, strPass)
                    Set cn = gobjRegister.GetConnection(strServer, Nvl(!用户名), strPass _
                                    , CBool(Val("0-不转换密码")) _
                                    , intDeviceType _
                                    , _
                                    , CBool(Val("0-不更新部件的连接对象")))
                    Set GetDBConnectionEx = cn
                    
                    Exit Do
                End If
                
                .MoveNext
            Loop
        End With
    Else
        Set GetDBConnectionEx = Nothing
    End If
    
    Exit Function
    
hErr:
    Call mdlPublic.ErrCenter
End Function

Public Function GetDBConnection(Optional ByVal Index As Integer = 0) As ADODB.Connection
'功能：通过数据连接编号获取对应的数据连接对象
'参数：
'  Index：数据连接编号

    Dim strKey As String, strPass As String, strServer As String
    Dim cn As New ADODB.Connection

    If Index <= 0 Then
        Set GetDBConnection = gcnOracle
    Else
        On Error GoTo hErr
        
        strKey = "_" & Index
        
        If gclsCNs.Item(strKey) Is Nothing Then
            '加载数据连接
            If grsConnect.State = adStateOpen Then
                With grsConnect
                    If .RecordCount > 0 Then .MoveFirst
                    Do While .EOF = False
                        If Nvl(!编号, 0) = Index Then
                            '初始化连接对象
                            strServer = Nvl(!IP) & _
                                        IIF(Nvl(!端口) = "", ":1521", ":" & Nvl(!端口)) & _
                                        IIF(Nvl(!实例名) = "", "", "/" & Nvl(!实例名))
                            strPass = Nvl(!密码)
                            '解密
                            strPass = mdlPublic.Decipher(MSTR_DBLINK_KEY, strPass)
                            Set cn = gobjRegister.GetConnection(strServer, Nvl(!用户名), strPass _
                                            , CBool(Val("0-不转换密码")) _
                                            , IIF(gblnManagementTool, Val("1-OraOLEDB"), Val("0-MSODBC")) _
                                            , _
                                            , CBool(Val("0-不更新部件的连接对象")))
                            Call gclsCNs.Add(Index, Nvl(!编号), cn)
                            GoTo makSet
                            
                            Exit Do
                        End If
                        
                        .MoveNext
                    Loop
                End With
            End If
        Else
makSet:
            '获取数据连接
            If Not gclsCNs.Item(strKey).Connection Is Nothing Then
                If gclsCNs.Item(strKey).Connection.State <> adStateOpen Then
                    Call gclsCNs.Item(strKey).Connection.Open
                End If
                Set GetDBConnection = gclsCNs.Item(strKey).Connection
            End If
        End If
    End If
    
    Exit Function
    
hErr:
'    Call mdlPublic.ErrCenter
End Function

Public Function OpenSQLRecord(ByVal strSQL As String, ByVal strTitle As String _
    , ParamArray arrInput() As Variant) As ADODB.Recordset
'功能：通过Command对象打开带参数SQL的记录集
'参数：strSQL=条件中包含参数的SQL语句,参数形式为"[x]"
'             x>=1为自定义参数号,"[]"之间不能有空格
'             同一个参数可多处使用,程序自动换为ADO支持的"?"号形式
'             实际使用的参数号可不连续,但传入的参数值必须连续(如SQL组合时不一定要用到的参数)
'      arrInput=不定个数的参数值,按参数号顺序依次传入,必须是明确类型
'      strTitle=用于SQLTest识别的调用窗体/模块标题
'      arrInput=第一个参数格式如果是“[数据连接=x][|查询方式=1-LOB]”。x表示数据连接的编号；
'返回：记录集，CursorLocation=adUseClient,LockType=adLockReadOnly,CursorType=adOpenStatic
'举例：
'SQL语句为="Select 姓名 From 病人信息 Where (病人ID=[3] Or 门诊号=[3] Or 姓名 Like [4]) And 性别=[5] And 登记时间 Between [1] And [2] And 险类 IN([6],[7])"
'调用方式为：Set rsPati=OpenSQLRecord(strSQL, Me.Caption, CDate(Format(rsMove!转出日期,"yyyy-MM-dd")),dtp时间.Value, lng病人ID, "张%", "男", 20, 21)
    Static cmdData As New ADODB.Command
    Static intTag As Integer
    
    Dim StrPar As String, arrPar As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim strSeq As String, intMax As Integer, i As Integer
    Dim strLog As String, varValue As Variant
    Dim strSQLtmp As String, arrStr As Variant
    Dim strTmp As String, strSQLtmp1 As String
    Dim intConnect As Integer, intQueryMode As Integer
    Dim arrInputNew() As Variant
    
    '判断是否以其他的数据连接执行记录集
    intConnect = 0
    arrInputNew = arrInput
    If UBound(arrInput) >= 0 Then
        If arrInput(0) Like "数据连接=[0-9]*" Then
            strTmp = Split(arrInput(0), "=")(1)
            intConnect = Val(strTmp)
            If UBound(arrInput) > 0 Then
                '重整参数
                ReDim Preserve arrInputNew(UBound(arrInput) - 1)
                For i = 1 To UBound(arrInput)
                    arrInputNew(i - 1) = arrInput(i)
                Next
            End If
        End If
        
        '获取指定LOB查询方式
        If arrInput(0) Like "*|查询方式=*" Then
            strTmp = Split(arrInput(0), "|")(1)
            intQueryMode = Val(Split(strTmp, "=")(1))
            If Not arrInput(0) Like "数据连接=[0-9]*" Then
                '重整参数
                ReDim Preserve arrInputNew(UBound(arrInput) - 1)
                For i = 1 To UBound(arrInput)
                    arrInputNew(i - 1) = arrInput(i)
                Next
            End If
        End If
    End If
    
    '检查如果使用了动态内存表，并且没有使用/*+ XXX*/等提示字时自动加上
    strSQLtmp = Trim(UCase(strSQL))
    If Mid(Trim(Mid(strSQLtmp, 7)), 1, 2) <> "/*" And Mid(strSQLtmp, 1, 6) = "SELECT" Then
        arrStr = Split("F_STR2LIST,F_NUM2LIST,F_NUM2LIST2,F_STR2LIST2", ",")
        For i = 0 To UBound(arrStr)
            strSQLtmp1 = strSQLtmp
            Do While InStr(strSQLtmp1, arrStr(i)) > 0
                '判断前面是否用了IN 用了则不加Rule
                '先找到最近一个SELECT
                strTmp = Mid(strSQLtmp1, 1, InStr(strSQLtmp1, arrStr(i)) - 1)
                strTmp = Replace(TrimEx(Mid(strTmp, 1, InStrRev(strTmp, "SELECT") - 1)), " ", "")
                If Len(strTmp) > 1 Then strTmp = Mid(strTmp, Len(strTmp) - 2)               '取后面3个字符
                
                If strTmp = "IN(" Then '属于in(select这种情况，则继续循环，看是否存在没有使用这种写法的其他动态内存函数
                   strSQLtmp1 = Mid(strSQLtmp1, InStr(strSQLtmp1, arrStr(i)) + Len(arrStr(i)))
                Else
                    Exit For
                End If
            Loop
        Next
        If i <= UBound(arrStr) Then
            strSQL = "Select /*+ RULE*/" & Mid(Trim(strSQL), 7)
        End If
    End If
        
    '分析自定的[x]参数
    lngLeft = InStr(1, strSQL, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strSQL, "]")
        If lngRight = 0 Then Exit Do
        '可能是正常的"[编码]名称"
        strSeq = Mid(strSQL, lngLeft + 1, lngRight - lngLeft - 1)
        If IsNumeric(strSeq) Then
            i = CInt(strSeq)
            StrPar = StrPar & "," & i
            If i > intMax Then intMax = i
        End If
        
        lngLeft = InStr(lngRight + 1, strSQL, "[")
    Loop

    '替换为"?"参数
    strLog = strSQL
    For i = 1 To intMax
        strSQL = Replace(strSQL, "[" & i & "]", "?")
        
        '产生用于SQL跟踪的语句
        varValue = arrInputNew(i - 1)
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
            strLog = Replace(strLog, "[" & i & "]", varValue)
        Case "String" '字符
            strLog = Replace(strLog, "[" & i & "]", "'" & Replace(varValue, "'", "''") & "'")
        Case "Date" '日期
            strLog = Replace(strLog, "[" & i & "]", "To_Date('" & Format(varValue, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')")
        End Select
    Next

    '清除原有参数:不然不能重复执行
    cmdData.CommandText = "" '不为空有时清除参数出错
    Do While cmdData.Parameters.count > 0
        cmdData.Parameters.Delete 0
    Loop
    
    '创建新的参数
    lngLeft = 0: lngRight = 0
    arrPar = Split(Mid(StrPar, 2), ",")
    For i = 0 To UBound(arrPar)
        varValue = arrInputNew((arrPar(i) - 1))
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarNumeric, adParamInput, 30, varValue)
        Case "String" '字符
            intMax = LenB(StrConv(varValue, vbFromUnicode))
            
            If intMax <= 2000 Then
                intMax = IIF(intMax <= 200, 200, 2000)
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarChar, adParamInput, intMax, varValue)
            Else
                If intMax < 4000 Then intMax = 4000
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adLongVarChar, adParamInput, intMax, varValue)
            End If
        Case "Date" '日期
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adDBTimeStamp, adParamInput, , varValue)
        Case "Variant()" '数组
            '这种方式可用于一些IN子句或Union语句
            '表示同一个参数的多个值,参数号不可与其它数组的参数号交叉,且要保证数组的值个数够用
            If arrPar(i) <> lngRight Then lngLeft = 0
            lngRight = arrPar(i)
            Select Case TypeName(varValue(lngLeft))
            Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarNumeric, adParamInput, 30, varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", varValue(lngLeft), 1, 1)
            Case "String" '字符
                intMax = LenB(StrConv(varValue(lngLeft), vbFromUnicode))
                            
                If intMax <= 2000 Then
                    intMax = IIF(intMax <= 200, 200, 2000)
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarChar, adParamInput, intMax, varValue(lngLeft))
                Else
                    If intMax < 4000 Then intMax = 4000
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adLongVarChar, adParamInput, intMax, varValue(lngLeft))
                End If
                
                strLog = Replace(strLog, "[" & lngRight & "]", "'" & Replace(varValue(lngLeft), "'", "''") & "'", 1, 1)
            Case "Date" '日期
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adDBTimeStamp, adParamInput, , varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", "To_Date('" & Format(varValue(lngLeft), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')", 1, 1)
            End Select
            lngLeft = lngLeft + 1 '该参数在数组中用到第几个值了
        End Select
    Next

    '执行返回记录集
    If cmdData.ActiveConnection Is Nothing Then
        Set cmdData.ActiveConnection = GetDBConnection(intConnect)    'gcnOracle '这句比较慢
    ElseIf cmdData.ActiveConnection.ConnectionString <> gcnOracle.ConnectionString _
        Or intTag <> intConnect Then
        Set cmdData.ActiveConnection = GetDBConnection(intConnect)    'gcnOracle '这句比较慢
    End If
    cmdData.CommandText = strSQL
    intTag = intConnect
    
    Call SQLTest(App.ProductName, strTitle, strLog)
    
    On Error Resume Next
    If intQueryMode = Val("1-LOB") Then
        '判断LOB字段类型的查询
        GoTo makLOB
    Else
        Set OpenSQLRecord = cmdData.Execute
    End If
    
    If Err.Number = -2147467259 Then
makLOB:
        If gcolOLEDBConnect Is Nothing Then
            Set gcolOLEDBConnect = New Collection
        End If
        'CLOB、BLOB字段类型支持
        '获取缓存连接对象
        Set gcnOLEDB = mdlPublic.GetOLEDBConnect(gcnOracle, gcolOLEDBConnect, gobjRegister)
        If gcnOLEDB Is Nothing Then
            Set gcnOLEDB = gobjRegister.ReGetConnection(Val("1-OracleOLEDB"), "", gcnOracle)
            '缓存
            Call gcolOLEDBConnect.Add(gcnOLEDB)
        End If
        If Not gcnOLEDB Is Nothing Then
            Set cmdData.ActiveConnection = gcnOLEDB
            'Set OpenSQLRecord = cmdData.Execute
            'CLOB、BLOB字段类型如果使用Command对象，记录集对象默认的锁adOpenUnspecified会引起执行慢
            '因此，改用记录集对象的Open方法
            If OpenSQLRecord Is Nothing Then
                Set OpenSQLRecord = New ADODB.Recordset
            End If
            OpenSQLRecord.Open cmdData, , adOpenStatic, adLockOptimistic
        End If
    End If
    On Error GoTo 0
    
    Set OpenSQLRecord.ActiveConnection = Nothing
    Call SQLTest
End Function

Public Sub ExecuteProcedure(strSQL As String, ByVal strFormCaption As String)
'功能：执行过程语句,并自动对过程参数进行绑定变量处理
'参数：strSQL=过程语句,可能带参数,形如"过程名(参数1,参数2,...)"。
'说明：以下几种情况过程参数不使用绑定变量,仍用老的调用方法：
'  1.参数部份是表达式,这时程序无法处理绑定变量类型和值,如"过程名(参数1,100.12*0.15,...)"
'  2.中间没有传入明确的可选参数,这时程序无法处理绑定变量类型和值,如"过程名(参数1, , ,参数3,...)"
'  3.因为该过程是自动处理,不是一定使用绑定变量,对带"'"的字符参数,仍要使用"''"形式。
    Dim cmdData As New ADODB.Command
    Dim strProc As String, StrPar As String
    Dim blnStr As Boolean, intBra As Integer
    Dim strTemp As String, i As Long
    Dim intMax As Integer, datCur As Date
    
    If Right(Trim(strSQL), 1) = ")" Then
        '清除原有参数:不然不能重复执行
'        cmdData.CommandText = "" '不为空有时清除参数出错
'        Do While cmdData.Parameters.Count > 0
'            cmdData.Parameters.Delete 0
'        Loop
        
        '执行的过程名
        strTemp = Trim(strSQL)
        strProc = Trim(Left(strTemp, InStr(strTemp, "(") - 1))
        
        '执行过程参数
        datCur = CDate(0)
        strTemp = Mid(strTemp, InStr(strTemp, "(") + 1)
        strTemp = Trim(Left(strTemp, Len(strTemp) - 1)) & ","
        For i = 1 To Len(strTemp)
            '是否在字符串内，以及表达式的括号内
            If Mid(strTemp, i, 1) = "'" Then blnStr = Not blnStr
            If Not blnStr And Mid(strTemp, i, 1) = "(" Then intBra = intBra + 1
            If Not blnStr And Mid(strTemp, i, 1) = ")" Then intBra = intBra - 1
            
            If Mid(strTemp, i, 1) = "," And Not blnStr And intBra = 0 Then
                StrPar = Trim(StrPar)
                With cmdData
                    If IsNumeric(StrPar) Then '数字
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.count, adVarNumeric, adParamInput, 30, Val(StrPar))
                    ElseIf Left(StrPar, 1) = "'" And Right(StrPar, 1) = "'" Then '字符串
                        StrPar = Mid(StrPar, 2, Len(StrPar) - 2)
                        
                        'Oracle连接符运算:'ABCD'||CHR(13)||'XXXX'||CHR(39)||'1234'
                        If InStr(Replace(StrPar, " ", ""), "'||") > 0 Then GoTo NoneVarLine
                        
                        '双"''"的绑定变量处理
                        If InStr(StrPar, "''") > 0 Then StrPar = Replace(StrPar, "''", "'")
                        
                        '电子病历处理LOB时，如果用绑定变量转换为RAW时第2000个字符不正确
                        intMax = LenB(StrConv(StrPar, vbFromUnicode))
                        If intMax = 0 Or intMax < 200 Then intMax = 200
                        If intMax > 1999 Then GoTo NoneVarLine
                        
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.count, adVarChar, adParamInput, intMax, StrPar)
                    ElseIf UCase(StrPar) Like "TO_DATE('*','*')" Then '日期
                        StrPar = Split(StrPar, "(")(1)
                        StrPar = Trim(Split(StrPar, ",")(0))
                        StrPar = Mid(StrPar, 2, Len(StrPar) - 2)
                        If StrPar = "" Then
                            'NULL值当成数字处理可兼容其他类型
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.count, adVarNumeric, adParamInput, , Null)
                        Else
                            If Not IsDate(StrPar) Then GoTo NoneVarLine
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.count, adDBTimeStamp, adParamInput, , CDate(StrPar))
                        End If
                    ElseIf UCase(StrPar) = "SYSDATE" Then '日期
                        If datCur = CDate(0) Then datCur = Currentdate
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.count, adDBTimeStamp, adParamInput, , datCur)
                    ElseIf UCase(StrPar) = "NULL" Then 'NULL值当成字符处理可兼容其他类型
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.count, adVarChar, adParamInput, 200, Null)
                    ElseIf StrPar = "" Then '可选参数当成NULL处理可能改变了缺省值:因此可选参数不能写在中间
                        GoTo NoneVarLine
                    Else '可能是其他复杂的表达式，无法处理
                        GoTo NoneVarLine
                    End If
                End With
                
                StrPar = ""
            Else
                StrPar = StrPar & Mid(strTemp, i, 1)
            End If
        Next
        
        '补充?号
        strTemp = ""
        For i = 1 To cmdData.Parameters.count
            strTemp = strTemp & ",?"
        Next
        strProc = "Call " & strProc & "(" & Mid(strTemp, 2) & ")"
        
        '执行过程
        'If cmdData.ActiveConnection Is Nothing Then
            Set cmdData.ActiveConnection = gcnOracle '这句比较慢
            cmdData.CommandType = adCmdText
        'End If
        cmdData.CommandText = strProc
        
        Call cmdData.Execute

    Else
        GoTo NoneVarLine
    End If
    Exit Sub
NoneVarLine:
    
    '说明：为了兼容新连接方式
    '1.新连接用adCmdStoredProc方式在8i下面有问题
    '2.新连接如果不使用{},则即使过程没有参数也要加()
    strSQL = "Call " & strSQL
    If InStr(strSQL, "(") = 0 Then strSQL = strSQL & "()"
    gcnOracle.Execute strSQL, , adCmdText

End Sub

Public Function ConvertSBC(ByVal strText As String) As String
'功能：转换全角字符为半角字符
    Dim i As Long, k As Long
    
    For i = 1 To Len(strText)
        k = InStr(GSTR_SBC, Mid(strText, i, 1))
        If k > 0 Then
            strText = Left(strText, i - 1) & Mid(GSTR_DBC, k, 1) & Mid(strText, i + 1)
        End If
    Next
    ConvertSBC = strText
End Function

Public Function IsType(ByVal varType As DataTypeEnum, ByVal varBase As DataTypeEnum) As Boolean
'功能：判断某个ADO字段数据类型是否与指定字段类型是同一类(如数字,日期,字符,二进制)
    Dim intA As Integer, intB As Integer
    
    Select Case varBase
        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
            intA = -1
        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
            intA = -2
        Case adDBTimeStamp, adDBTime, adDBDate, adDate
            intA = -3
        Case adBinary, adVarBinary, adLongVarBinary
            intA = -4
        Case Else
            intA = varBase
    End Select
    Select Case varType
        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
            intB = -1
        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
            intB = -2
        Case adDBTimeStamp, adDBTime, adDBDate, adDate
            intB = -3
        Case adBinary, adVarBinary, adLongVarBinary
            intB = -4
        Case Else
            intB = varType
    End Select
    
    IsType = intA = intB
End Function

Public Sub PopupButtonMenu(ToolBar As Object, Button As Object, objMenu As Object)
'功能：在下拉式工具按钮中弹出一个菜单
    Dim vRect As RECT, vDot1 As PointAPI, vDot2 As PointAPI
    
    Call GetWindowRect(ToolBar.hwnd, vRect)
    vDot1.X = vRect.Left: vDot1.Y = vRect.Top
    vDot2.X = vRect.Right: vDot2.Y = vRect.Bottom
    
    Call ScreenToClient(ToolBar.Parent.hwnd, vDot1)
    Call ScreenToClient(ToolBar.Parent.hwnd, vDot2)
    
    vDot1.X = vDot1.X * 15: vDot1.Y = vDot1.Y * 15
    vDot2.X = vDot2.X * 15: vDot2.Y = vDot2.Y * 15
    ToolBar.Parent.PopupMenu objMenu, 2, vDot1.X + Button.Left, vDot2.Y
End Sub

Public Function zlHomePage(hwnd As Long) As Boolean
'功能：根据产品发行码，联结主页
    Dim strCode As String
    
    strCode = zlRegInfo("支持商URL")
    If strCode <> "-" Then
        ShellExecute hwnd, "open", "http://" & strCode, "", "", 1
        zlHomePage = True
    End If
End Function

Public Function zlWebForum(hwnd As Long) As Boolean
'功能：根据产品发行码，联结论坛
    Dim strCode As String
    
    'strCode = zlRegInfo("支持商BBS")
    strCode = "www.zlsoft.com/techbbs/index.asp"
    If strCode <> "-" Then
        ShellExecute hwnd, "open", "http://" & strCode, "", "", 1
        zlWebForum = True
    End If
End Function

Public Function zlMailTo(hwnd As Long) As Boolean
'功能：根据产品发行码发送电子邮件
    Dim strCode As String
    strCode = zlRegInfo("支持商MAIL")
    If strCode <> "-" Then
        ShellExecute hwnd, "open", "mailto:" & strCode, "", "", 1
        zlMailTo = True
    End If
End Function

Public Function InitRegister() As Boolean
'功能：初始化注册部件对象gobjRegister

    Dim strTmp As String
    
    If gobjRegister Is Nothing Then
        On Error Resume Next
        Set gobjRegister = GetObject("", "zlRegister.clsRegister")
        Err.Clear
    End If
    
    '用于支持未通过导航台（启动程序prjMain）调用本部件的情况。
    If gobjRegister Is Nothing Then
        Set gobjRegister = CreateObject("zlRegister.clsRegister")
        Err.Clear
        If Not gobjRegister Is Nothing Then
            Call gobjRegister.zlRegInit(gcnOracle)
            strTmp = gobjRegister.zlRegCheck(False)
            If strTmp <> "" Then
                MsgBox strTmp, vbExclamation, gstrProductName
                Exit Function
            End If
        End If
    End If
    
    On Error GoTo 0
    If gobjRegister Is Nothing Then
        MsgBox "创建zlRegister部件对象失败,请检查文件是否存在并且正确注册。", vbExclamation, gstrProductName
        Exit Function
    End If
    InitRegister = True
End Function

Public Function GetPrivFunc(lngSys As Long, lngProgID As Long) As String
'功能：返回当前用户具有的指定程序的功能串
'参数：lngSys     如果是固定模块，则为0
'      lngProgId  程序序号
'返回：分号间隔的功能串,为空表示没有权限
    GetPrivFunc = gobjRegister.zlRegFunc(lngSys, lngProgID)
End Function

'--------------------------------------------------
'功能：验证系统注册授权的正确性
'参数：blnTemp-是否从未保存的临时注册信息验证
'返回：正确返回"";错误返回错误信息
'--------------------------------------------------
Public Function zlRegCheck(Optional blnTemp As Boolean) As String
    zlRegCheck = gobjRegister.zlRegCheck(blnTemp)
End Function

'--------------------------------------------------
'功能：获得指定的产品发行或注册授权信息
'参数： strItem-指定的授权项目
'       blnTemp-是否从未保存的临时注册信息验证
'       intBits-对于同时有多项信息的单位名称、产品开发商等指定获得第几个信息,0-N,为-1时表示返回";"间隔的多个
'返回：正确时返回指定的信息；错误返回""
'--------------------------------------------------
Public Function zlRegInfo(strItem As String, Optional blnTemp As Boolean, Optional intBits As Integer) As String
    zlRegInfo = gobjRegister.zlRegInfo(strItem, blnTemp, intBits)
End Function

'--------------------------------------------------
'功能：获得授权工具信息
'返回：按2的工具末位次方返回工具许可
'--------------------------------------------------
Public Function zlRegTool(Optional blnTemp As Boolean) As Long
    zlRegTool = gobjRegister.zlRegTool(blnTemp)
End Function

Public Function SetBit(ByVal strBit As String, ByVal intBit As Integer, Optional ByVal intVal As Integer = -1) As String
'功能：将指定位字符串strBit中的第intBit位设置为0或1
'参数：intVal=设置值,0或1,不传表示反转
    If Len(strBit) < intBit Then strBit = strBit & String(intBit - Len(strBit), "0")
    If intVal = -1 Then intVal = IIF(Val(Mid(strBit, intBit, 1)) = 0, 1, 0)
    SetBit = Left(strBit, intBit - 1) & intVal & Mid(strBit, intBit + 1)
End Function

'--------------------------------------------------
'功能：检查是否为网络断开或ADO断开引发的错误!
'返回：True:恢复连接成功 False恢复连接失败
'--------------------------------------------------
Public Function CheckAdoConnction(ByRef blnStatus As Boolean) As Boolean
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim blnAdoErr As Boolean
    Dim strError As String
    On Error GoTo ErrHand
    blnAdoErr = False
    blnStatus = False

    On Error GoTo ErrHand
    Err = 0
    DoEvents
    If gcnOracle.State = adStateOpen Then gcnOracle.Close
    gcnOracle.Open
    If blnAdoErr Then
        'True '是ORA-12560不能与ORACLE连接引起
        CheckAdoConnction = True
    Else
        'False '可以正常连接
        CheckAdoConnction = False
        On Error Resume Next
        '重连后判断客户端是否被禁止使用，若被禁止，则自动断开连接
        strSQL = "Select NVL(禁止使用,0)  禁止使用 From zlClients Where 工作站=SYS_CONTEXT('USERENV','TERMINAL')"
        Set rsTmp = OpenSQLRecord(strSQL, "CheckAdoConnction")
        If Err.Number <> 0 Then Err.Clear
        If Not rsTmp Is Nothing Then
            If Not rsTmp.EOF Then
                If rsTmp!禁止使用 = 1 Then
                    If gcnOracle.State = adStateOpen Then gcnOracle.Close
                    CheckAdoConnction = True
                    gblnAutoConnect = False
                    MsgBox "当前工作站已经被管理员禁用，请联系管理员解除禁用并重新登录！", vbInformation, "中联软件"
                End If
            End If
        End If
    End If
    Exit Function
ErrHand:
    If Err.Number = -2147467259 Or Err.Number = 3709 Then
        If InStr(Err.Description, "ORA-12560") > 0 Then
            blnAdoErr = True
            Resume Next
        ElseIf InStr(Err.Description, "ORA-12543") > 0 Then
            blnAdoErr = True
            Resume Next
        Else
            '其他错误引发的网络问题
            CheckAdoConnction = True
            blnStatus = True
        End If
    Else
        CheckAdoConnction = False
    End If
End Function

Public Function IsOLEDBConnection(ByVal cnMain As ADODB.Connection) As Boolean
'功能：判断当前连接是否是OraOLEDB连接
'根据Provider来判断，存在两种方式
'方式一：'Provider=OraOLEDB.Oracle.1;Password=HIS;Persist Security Info=True;User ID=ZLHIS;Data Source="DYYY";Extended Properties="PLSQLRSet=1"
'方式二：
'.Provider = "OraOLEDB.Oracle"
'.Open "PLSQLRSet=1;Data Source=" & strServer & strPersist_Security_Info, strUserName, strPassWord
'这两种方式均会自动设置.Provider属性
    '使用Like是因为可能后面增加版本如OraOLEDB.Oracle.1
    If UCase(cnMain.Provider) Like "ORAOLEDB.ORACLE*" Then
        IsOLEDBConnection = True
    End If
End Function

Public Function GetAutoConnect() As Boolean
'功能：获取是否有权限断网自动连接
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Nvl(B.参数值,Nvl(A.参数值,A.缺省值)) As 参数值" & _
        " From zlParameters A,zlUserParas B" & _
        " Where A.ID=B.参数ID(+) And A.系统 is Null And A.模块 is Null" & _
        " And Nvl(A.私有,0)=0 And Nvl(A.本机,0)=1 And A.参数名='网络断网自动重连'" & _
        " And B.用户名(+) is Null And B.机器名(+)=SYS_CONTEXT('USERENV','TERMINAL')"
    Set rsTmp = OpenSQLRecord(strSQL, "断网自动连接权限", "")
    If Not rsTmp.EOF Then
        GetAutoConnect = Val(Nvl(rsTmp!参数值, 0)) = 1
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckErrConnectInfo(ByVal strErrNum As String, ByVal strNote As String, ByVal strErrInfo As String, ByVal intType As Integer) As Boolean
    '------------------------------------------------
    '功能： 按照类型IntType(1,2)检查vb和oralce返回的具体错误信息，来判断是否为网络断开引发的错误或者是其他的错误引发
    '参数： strNote错误信息,strErrInfo错误详细信息,intType 错误类型 1：VB错误 2:ORACLE错误
    '返回： True:网络引发的错误 False:其他错误
    '------------------------------------------------
    Dim strTemp As String
    Dim i As Integer
    If intType = 1 Then
        'VB具体错误
   
        If InStr(strErrInfo, "ORA-12560") > 0 Then
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-12571") > 0 Then
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-03114") > 0 Then
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "E_FAIL") > 0 Then
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-02396") > 0 Then '超出最大空闲时间, 请重新连接 IDLE_TIME profile
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-02399") > 0 Then '超出最大连接时间, 您将被注销 connect_time profile
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-01012") > 0 Then '没有登录
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-00028") > 0 Then '会话被终止
            CheckErrConnectInfo = True
        Else
            If strErrNum = "3709" Then '3709描述：连接无法用于执行此操作。在此上下文中它可能已被关闭或无效。单独处理
                CheckErrConnectInfo = True
            Else
                If strNote = "不确定的错误" Then
                    CheckErrConnectInfo = True
                Else
                    CheckErrConnectInfo = False
                End If
            End If
        End If
    Else
        'ORACLE具体错误
        If InStr(strErrInfo, "SQLSetConnectAttr") > 0 Then
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-12560") > 0 Then
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "E_FAIL") > 0 Then
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-12571") > 0 Then
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-03114") > 0 Then
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-12543") > 0 Then
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-02396") > 0 Then '超出最大空闲时间, 请重新连接 IDLE_TIME profile
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-02399") > 0 Then '超出最大连接时间, 您将被注销 connect_time profile
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-01012") > 0 Then '没有登录
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-00028") > 0 Then '会话被终止
            CheckErrConnectInfo = True
        Else
            CheckErrConnectInfo = False
        End If
    End If
End Function

Public Function PictureSpin(objSource As StdPicture, bytSpinType As Byte, objDraw As PictureBox) As StdPicture
'功能：图片旋转(顺时针，逆时针）
'参数：objPic=原图像
'      SpinType=1-顺时针90度,2-逆时针90度
'      objTemp=绘图用的临时画布(PictureBox)
'返回：旋转后的图片

    Dim p() As Long
    Dim W As Long, H As Long
    Dim i As Long, j As Long
    
    If bytSpinType = 0 Then
        Set PictureSpin = objSource
        Exit Function
    End If
    
    '取原始像素
    objDraw.BorderStyle = 0
    objDraw.AutoRedraw = True
    objDraw.ScaleMode = vbPixels
    objDraw.Width = objDraw.Container.ScaleX(objDraw.ScaleX(objSource.Width, vbHimetric, vbPixels), vbPixels, objDraw.Container.ScaleMode)
    objDraw.Height = objDraw.Container.ScaleY(objDraw.ScaleY(objSource.Height, vbHimetric, vbPixels), vbPixels, objDraw.Container.ScaleMode)
    objDraw.PaintPicture objSource, 0, 0, objDraw.ScaleWidth, objDraw.ScaleHeight
    
    W = objDraw.ScaleWidth
    H = objDraw.ScaleHeight

    ReDim p(W - 1, H - 1)
    For i = 0 To W - 1
        For j = 0 To H - 1
            p(i, j) = objDraw.Point(i, j)
        Next j
    Next i
    
    '转换绘图
    objDraw.Cls
    objDraw.Width = objDraw.Container.ScaleY(H, vbPixels, objDraw.Container.ScaleMode)
    objDraw.Height = objDraw.Container.ScaleX(W, vbPixels, objDraw.Container.ScaleMode)
    For i = 0 To H - 1
        For j = 0 To W - 1
            If bytSpinType = 1 Then
                objDraw.PSet (H - i - 1, j), p(j, i)
            ElseIf bytSpinType = 2 Then
                objDraw.PSet (i, W - j - 1), p(j, i)
            End If
        Next j
    Next i
    
    Set PictureSpin = objDraw.Image
    objDraw.ScaleMode = vbTwips
End Function

Public Sub CboSetText(cboControl As Object, ByVal strText As String, Optional ByVal blnAfter As Boolean = True, Optional strSplit As String = "-")
'功能：根据文本串更新Combo控件的当前值
'参数：cboControl  准备设置的ComboBox控件
'      strText     输入的文本串
'      blnAfter    表示在分隔符之前或之后取值。如果没有分隔符，则取之后
'      strSplit    分隔符，通常为-
    Dim lngPos As Long
    Dim lngCount As Long
    Dim strTemp As String
    Dim blnMatch As Boolean
    
    For lngCount = 0 To cboControl.ListCount - 1
        strTemp = cboControl.List(lngCount)
        
        lngPos = InStr(strTemp, strSplit)
        If lngPos = 0 Then
            '直接返回整个字符串
            If strText = strTemp Then
                blnMatch = True
                Exit For
            End If
        Else
            If blnAfter = False Then
                '圆点之前
                If strText = Mid(strTemp, 1, lngPos - 1) Then
                    blnMatch = True
                    Exit For
                End If
            Else
                If strText = Mid(strTemp, lngPos + 1) Then
                    blnMatch = True
                    Exit For
                End If
            End If
        End If
    Next
    If blnMatch = True Then
        '已经找到
        cboControl.ListIndex = lngCount
    Else
        If blnAfter = True Then
            '这才是实际内容，如果为前则只是编码
            If strText <> "" Then
                cboControl.AddItem strText
                cboControl.ListIndex = cboControl.NewIndex
            End If
        End If
    End If
End Sub

Public Function CheckSQLPlan(ByVal strSQLCheck As String, Optional ByRef vsPlan As VSFlexGrid, _
    Optional ByVal intConnect As Integer, Optional ByRef blnSuccess As Boolean) As Boolean
'性能问题检查:
'         1.大表全表扫描zlbigtable+zlbaktables，
'         2.中型表全表扫描(如果有统计信息，User_tab_statistics:num_rows>3000(药品目录一般是这个值以上) AND num_rows<100 0000百万以内)
'         3.大表上引用基础表(非大表)的外键上的索引
'         4.大表和中型表索引全扫描（inex full scan，INDEX FAST FULL SCAN）
'         5.大表和中型表跳跃式索引扫描（INDEX SKIP SCAN）
'返回：blnReturn=true 有性能问题
    Dim rsPlan As ADODB.Recordset
    Dim i As Long, strSQL As String
    Dim j As Long, blnReturn As Boolean
    Dim rsIndex As New Recordset
    Dim rsData As ADODB.Recordset
    Dim strTable As String
    Dim rsCons_FK As New Recordset
    Dim StrPar As String
    Dim strTmp As String
    Dim strAllTable As String
    
    If intConnect > 0 Then
        blnSuccess = True
        CheckSQLPlan = False
        Exit Function
    End If
    
    Set rsPlan = GetSQLPlan(strSQLCheck, intConnect)
    If Not vsPlan Is Nothing Then
        vsPlan.Redraw = flexRDNone
        vsPlan.Rows = vsPlan.FixedRows
        vsPlan.FixedAlignment(1) = flexAlignLeftCenter
    End If
    
    blnSuccess = Not rsPlan Is Nothing
    
    If Not rsPlan Is Nothing Then
        If mstrBigTable = "" Then
            '先取大表,首次进入判断是否有zltables这张表
            '有ZLTABLES,就去B类和C类作为大表,否则取zlbigtabls和zlbaktables中的表
            If CheckTblExist("ZLTABLES") Then
               strSQL = " Select Distinct 表名 From Zltables Where 分类 In ('B1', 'B2', 'B3', 'C1', 'C2', 'C3') "
            Else
                strSQL = "Select Distinct 表名" & vbNewLine & _
                        "From Zlbigtables" & vbNewLine & _
                        "Union All" & vbNewLine & _
                        "Select Distinct 表名 From Zlbaktables"
            End If
            Call OpenRecord(rsIndex, strSQL, App.ProductName)
            Do While Not rsIndex.EOF
                mstrBigTable = mstrBigTable & "," & rsIndex!表名
                rsIndex.MoveNext
            Loop
            mstrBigTable = mstrBigTable & ","
        End If
        
        '再取中表（统计信息，User_tab_statistics:num_rows>3000）
        strSQL = "Select A.参数名,Nvl(A.参数值,A.缺省值) As 参数值" & _
            " From zlParameters A Where A.参数名 ='检查中型表'"
        Set rsData = OpenSQLRecord(strSQL, App.ProductName)
        If rsData.BOF = False Then
            StrPar = Nvl(rsData("参数值").Value, "0,0")
            If StrPar <> "0,0" Then
                If StrPar <> mstrMiddleTableRows Then
                    strSQL = "Select Table_Name as 表名 From User_Tab_Statistics Where Num_Rows > [1] And Num_Rows < [2] "
                    Set rsIndex = OpenSQLRecord(strSQL, App.ProductName, Val(Split(StrPar, ",")(0)), Val(Split(StrPar, ",")(1)))
                    mstrMiddleTable = ""
                    Do While Not rsIndex.EOF
                        If InStr("," & mstrBigTable & ",", "," & rsIndex!表名 & ",") = 0 Then
                            mstrMiddleTable = mstrMiddleTable & "," & rsIndex!表名
                        End If
                        rsIndex.MoveNext
                    Loop
                    mstrMiddleTable = mstrMiddleTable & ","
                    mstrMiddleTableRows = StrPar
                End If
            Else
                mstrMiddleTable = ""
                mstrMiddleTableRows = ""
            End If
        Else
            mstrMiddleTable = ""
            mstrMiddleTableRows = ""
        End If
        
        strAllTable = mstrMiddleTable & mstrBigTable
        
        For i = 1 To rsPlan.RecordCount
            strTmp = UCase(rsPlan!Operation & "")
            
            If Not vsPlan Is Nothing Then
                With vsPlan
                    .AddItem rsPlan!Cardinality & vbTab & Trim(rsPlan!Operation) & " " & rsPlan!name & " " & IIF(rsPlan!Bytes & "" = "" And rsPlan!cost & "" = "" And rsPlan!Time & "" = "", "", " (bytes=" & rsPlan!Bytes & " cost=" & rsPlan!cost & " time=" & Format(Time / 24 / 60 / 60, "HH:MM:SS") & ")")
                    .RowOutlineLevel(.Rows - 1) = Len(rsPlan!Operation & "") - Len(LTrim(rsPlan!Operation & ""))
                    .IsSubtotal(.Rows - 1) = True
                End With
            End If
            If InStr(strTmp, "TABLE ACCESS FULL") > 0 Then
                '判断是否是大表中表全扫描
                If InStr(strAllTable, "," & rsPlan!name & ",") > 0 Then
                    If Not vsPlan Is Nothing Then
                        vsPlan.Cell(flexcpForeColor, vsPlan.Rows - 1, 0, vsPlan.Rows - 1, vsPlan.Cols - 1) = &HFF& '红
                    End If
                    blnReturn = True
                End If
            ElseIf InStr(strTmp, "INDEX FAST FULL SCAN") > 0 Or InStr(strTmp, "INDEX FULL SCAN") > 0 Or InStr(strTmp, "INDEX SKIP SCAN") > 0 Then
                '判断是否是大表中表索引全扫描或跳跃式索引
                strTable = Split(rsPlan!name & "_", "_")(0)
                If InStr(strAllTable, "," & strTable & ",") > 0 Then
                    If Not vsPlan Is Nothing Then
                        vsPlan.Cell(flexcpForeColor, vsPlan.Rows - 1, 0, vsPlan.Rows - 1, vsPlan.Cols - 1) = &HFF& '红
                    End If
                    blnReturn = True
                End If
            ElseIf InStr(strTmp, "INDEX RANGE SCAN") > 0 Then
                '大表上使用了基础表(非大表)外键索引
                strTable = Split(rsPlan!name & "_", "_")(0)
                
                If InStr("," & mstrBigTable & ",", "," & strTable & ",") > 0 Then
                    strSQL = "Select distinct d.Table_Name, d.Index_Name, d.Column_Name,d.Column_Position" & vbNewLine & _
                        "              From User_Ind_Columns D" & vbNewLine & _
                        "              Where d.Index_Name = [1] " & vbNewLine & _
                        "              Order By d.Column_Position"
                    Set rsIndex = OpenSQLRecord(strSQL, App.ProductName, rsPlan!name & "")
                    If rsIndex.RecordCount > 0 Then
                        '找外键约束
                        Set rsCons_FK = GetConsFK(strTable, rsPlan!object_owner & "")
                        strTmp = ""
                        Do While Not rsIndex.EOF
                            strTmp = strTmp & "," & rsIndex!Column_Name
                            rsIndex.MoveNext
                        Loop
                        rsCons_FK.Filter = "Column_Name='" & Mid(strTmp, 2) & "'"
                        If rsCons_FK.RecordCount > 0 Then
                            strTable = Split(rsCons_FK!r_Constraint_Name & "_", "_")(0)
                            
                            '外键父表不是大表，则视为有性能问题
                            If InStr(mstrBigTable, "," & strTable & ",") = 0 Then
                                If Not vsPlan Is Nothing Then
                                    vsPlan.Cell(flexcpForeColor, vsPlan.Rows - 1, 0, vsPlan.Rows - 1, vsPlan.Cols - 1) = &HFF& '红
                                End If
                                blnReturn = True
                            End If
                        End If
                    End If
                End If
            End If
            
            rsPlan.MoveNext
        Next
        
        If Not vsPlan Is Nothing Then
            vsPlan.CellBorderRange 0, 0, vsPlan.Rows - 1, 0, &H808080, 0, 0, 1, 0, 0, 0
            vsPlan.CellBorderRange vsPlan.FixedRows - 1, 0, vsPlan.FixedRows - 1, vsPlan.Cols - 1, &H808080, 0, 0, 0, 1, 1, 0
            vsPlan.CellBorderRange vsPlan.Rows - 1, 0, vsPlan.Rows - 1, vsPlan.Cols - 1, &H808080, 0, 0, 0, 1, 1, 0
            vsPlan.AutoSize 0, vsPlan.Cols - 1
            vsPlan.Redraw = flexRDDirect
        End If
    End If
    
    CheckSQLPlan = blnReturn
End Function

Private Function GetConsFK(ByVal strFind As String, ByVal strOwner As String) As ADODB.Recordset
'功能：获取指定表的外键约束记录集
'参数：strFind=表名
    Dim strSQL As String
    Dim rsCons As New Recordset
    Dim rsCons_FK As New Recordset

    strSQL = "Select" & vbNewLine & _
        "        f.Constraint_Name, f.r_Constraint_Name,e.Column_Name,e.Position" & vbNewLine & _
        "       From User_Cons_Columns E, User_Constraints F" & vbNewLine & _
        "       Where e.Constraint_Name = f.Constraint_Name And e.owner = f.owner  And f.Constraint_Type = 'R' And f.Table_Name = [1] And f.owner = [2]" & vbNewLine & _
        "       order by e.constraint_name,e.position"
    Set rsCons = OpenSQLRecord(strSQL, App.ProductName, strFind, strOwner)
    Set rsCons_FK = New ADODB.Recordset
    rsCons_FK.Fields.Append "r_Constraint_Name", adVarChar, 50, adFldIsNullable
    rsCons_FK.Fields.Append "Constraint_Name", adVarChar, 50, adFldIsNullable
    rsCons_FK.Fields.Append "Column_Name", adVarChar, 100, adFldIsNullable
    rsCons_FK.CursorLocation = adUseClient
    rsCons_FK.LockType = adLockOptimistic
    rsCons_FK.CursorType = adOpenStatic
    rsCons_FK.Open
    Do While Not rsCons.EOF
        rsCons_FK.Filter = "Constraint_Name='" & rsCons!Constraint_Name & "'"
        If rsCons_FK.RecordCount = 0 Then
            rsCons_FK.AddNew
            rsCons_FK!Constraint_Name = rsCons!Constraint_Name & ""
            rsCons_FK!r_Constraint_Name = rsCons!r_Constraint_Name & ""
            rsCons_FK!Column_Name = rsCons!Column_Name & ""
        Else
            rsCons_FK!Column_Name = rsCons_FK!Column_Name & "," & rsCons!Column_Name
        End If
        rsCons_FK.Update
        rsCons.MoveNext
    Loop
    Set GetConsFK = rsCons_FK
End Function

Private Function GetSQLPlan(ByVal strSQLCheck As String, Optional ByVal intConnect As Integer = 0) As ADODB.Recordset
'功能：收集SQL的执行计划

    Dim strSQL As String, strSID As String
    Dim rsTmp As ADODB.Recordset
    Dim cnOracle As ADODB.Connection
            
    If strSQLCheck <> "" Then
        '准备连接对象
        Set cnOracle = GetDBConnection(intConnect)
        If cnOracle Is Nothing Then
            Exit Function
        End If
        
        On Error Resume Next
        strSID = Time()
        
        '执行计划
        strSQL = "explain plan set statement_id = '" & strSID & "' for " & strSQLCheck & ""
        strSQL = Replace(strSQL, "[系统]", glngSys)
        cnOracle.Execute strSQL
        If Err.Number = 0 Then
            strSQL = _
                    "Select Time From Plan_Table " & vbNewLine & _
                    "Connect By Prior ID = Parent_Id And Prior Statement_Id = Statement_Id " & vbNewLine & _
                    "Start With ID = 0 And Statement_Id = [1] " & vbNewLine & _
                    "Order By ID "
            On Error Resume Next
            Set GetSQLPlan = OpenSQLRecord(strSQL, "执行计划", "数据连接=" & intConnect, strSID)
            strSQL = _
                    "Select ID, LPad(' ', Level - 1) || Operation || ' ' || Options As Operation, Object_Name As Name" & _
                    "    ,Object_Owner, Cardinality, Bytes" & vbNewLine & _
                    "    ,Cost" & IIF(Err.Number = 0, ", Time ", ",0 as Time ") & vbNewLine & _
                    "From Plan_Table " & vbNewLine & _
                    "Connect By Prior ID = Parent_Id And Prior Statement_Id = Statement_Id " & vbNewLine & _
                    "Start With ID = 0 And Statement_Id = [1] " & vbNewLine & _
                    "Order By ID "
            Err.Clear
            Set GetSQLPlan = OpenSQLRecord(strSQL, "执行计划", "数据连接=" & intConnect, strSID)
            cnOracle.Execute "Delete plan_table"
        Else
            Set GetSQLPlan = Nothing
            Call ErrCenter
        End If
    End If
End Function

Public Function FindReport(ByVal strFind As String, ByRef lngHWND As Long, ByRef strInfo As String, _
    Optional ByVal lngSel As Long, Optional ByRef objReport As Report, Optional ByRef objRelation As RPTRelations, _
    Optional ByVal lngType As Long, Optional objParent As Object, Optional CurID As Integer) As String
'功能：查找选择报表的ID和名称
'参数：lngSel=默认选择某一行，绑定值行=lngsel则选中

    Dim strSQL As String
    Dim frmNewSelect As New frmSelect
    Dim i As Integer
    Dim bytType As Byte
    
    On Error GoTo errH
    
    strSQL = "select ID,编号,名称, 名称 || '(' || 编号 || ')' as 显示值 from zlreports"
    
    If strFind <> "" Then
        strFind = UCase(strFind)
        If IsCharChinese(strFind) Then
            strSQL = strSQL & " Where 名称 like '" & strFind & "%'"
        ElseIf IsCharAlpha(strFind) Then
            strSQL = strSQL & " Where Zlpinyincode(名称) like '" & strFind & "%'"
        ElseIf IsNumOrChar(strFind) Then
            strSQL = strSQL & " Where 编号 like '%" & strFind & "%'"
        End If
    End If
    strSQL = strSQL & " Order by 编号"
    
    With frmNewSelect
        .strMatch = strFind
        .strSQLList = strSQL
        .strFLDList = "ID," & adNumeric & ",&B|" & "编号," & adVarChar & ",&S|" & "名称," & adVarChar & ",&S|显示值," & adVarChar & ",&D"
        .strParName = "关联报表"
        .bytType = 1
        .mlngSel = lngSel
'        .mblnMulti = True
        .mblnRelationReport = True
        .mintConnect = 0
        .lngSeekHwnd = lngHWND
    
        '新修改的关联报表相关变量
        .selectlngType = lngType
        .selectCurID = CurID
        Set .selectObjReport = objReport
        Set .selectObjRelation = objRelation
        Set .selectObjParent = objParent
    End With
    
    Err.Clear
    On Error Resume Next
    
    frmNewSelect.Show 1
    If frmNewSelect.mblnOK Then
        strInfo = frmNewSelect.strOutDisp
        FindReport = frmNewSelect.strOutBand
        objRelation = frmNewSelect.selectObjRelation
        Unload frmNewSelect
    End If
    Exit Function
    
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetWinPath() As String
    '--------------------------------------------------------------------------------------------------------------
    '--功能:获取系统目录
    '--------------------------------------------------------------------------------------------------------------
    Dim Buffer As String
    Const MAX_PATH = 260
    Dim StrWinPath As String
    Dim rtn As Long
    
    Buffer = Space(MAX_PATH)
    rtn = GetWindowsDirectory(Buffer, Len(Buffer))
    StrWinPath = Left(Buffer, rtn)
    GetWinPath = StrWinPath
End Function

Public Function ShowDiff(ByVal strThisSQL As String, ByVal strNewSQL As String) As Boolean
'功能：显示两个文本的比对窗体
    Dim objFSO As TextStream
    Dim strCommand As String
    Dim lngProcess As Long
    Dim lngTemp As Long
    Dim strThisPath As String
    Dim strNewPath As String
    Const Process_Query_Information = &H400
    Const Still_Active = &H103
    Dim strSystem As String
    
    strNewPath = App.Path & "\NewSql"
    strThisPath = App.Path & "\ThisSql"
    If IsSys64 Then
        strSystem = "\syswow64"
    Else
        strSystem = "\system32"
    End If
    
    If gobjFile.FolderExists(strNewPath) Then
        Call gobjFile.DeleteFolder(strNewPath)
    End If
    If gobjFile.FolderExists(strThisPath) Then
        Call gobjFile.DeleteFolder(strThisPath)
    End If
    DoEvents
    
    Call gobjFile.CreateFolder(strNewPath)
    Call gobjFile.CreateFolder(strThisPath)
    
    DoEvents
    '文件1
    Set objFSO = gobjFile.CreateTextFile(strThisPath & "\" & "Wincmp.sql")
    objFSO.Write strThisSQL
    objFSO.Close
    '文件2
    Set objFSO = gobjFile.CreateTextFile(strNewPath & "\" & "Wincmp.sql")
    objFSO.Write strNewSQL
    objFSO.Close
    '对比
    strCommand = GetWinPath & strSystem & "\wincmp3.exe " & strThisPath & "\" & "Wincmp.sql" & " " & strNewPath & "\" & "Wincmp.sql"
    lngTemp = Shell(strCommand)
    DoEvents
    If Err <> 0 Then
        Err.Clear
        MsgBox "文件比较失败，请检查工具及文件位置是否正确:" & strSystem & "\wincmp3.exe", vbExclamation, "中联软件"
        Exit Function
    End If
    lngProcess = OpenProcess(Process_Query_Information, False, lngTemp)
    Do
        Sleep 100
        GetExitCodeProcess lngProcess, lngTemp
    Loop While lngTemp = Still_Active
    Err.Clear
    DoEvents

    On Error Resume Next
    If gobjFile.FolderExists(strNewPath) Then
        Call gobjFile.DeleteFolder(strNewPath)
    End If
    If gobjFile.FolderExists(strThisPath) Then
        Call gobjFile.DeleteFolder(strThisPath)
    End If
End Function

Public Function IsSys64() As Boolean
'功能：判断OS是32位，还是64位
'返回：True-64位；False-32位

    Dim lngMod As Long
    
    On Error GoTo errHandle
    
    lngMod = GetModuleHandle("ntdll.dll")
    If GetProcAddress(lngMod, "ZwWow64ReadVirtualMemory64") Then
       IsSys64 = True
    End If
    Exit Function
    
errHandle:
End Function

Public Function ReadFileToString(ByVal strFile As String) As String
    Dim strBuffer As String
    Dim lngHWND As Long
    Dim lngFileLen As Long

    lngHWND = FreeFile

    On Error Resume Next
    Open strFile For Binary Shared As lngHWND
    If Err.Number <> 0 Then
        MsgBox "Error " & Err.Number & vbCrLf & Err.Description & vbCrLf & "Error in ReadFileToString, File='" & strFile & "'", vbCritical
        GoTo Proc_Exit
    End If
    On Error GoTo 0
    
    lngFileLen = LOF(lngHWND)
    strBuffer = Space(lngFileLen)
    Get lngHWND, , strBuffer
    
    Close lngHWND
    
Proc_Exit:
    ReadFileToString = strBuffer
End Function

Public Sub SetCopyRelations(ByVal objRelations As RPTRelations, ByRef objRelationsCopy As RPTRelations)
'功能：复制一个关联报表对象
    Dim i As Long
    
    Set objRelationsCopy = New RPTRelations
    For i = 1 To objRelations.count
        objRelationsCopy.Add objRelations.Item(i).关联报表ID, objRelations.Item(i).参数名, objRelations.Item(i).参数值来源, objRelations.Item(i).关联报表名称, objRelations.Item(i).默认
    Next
End Sub

Public Sub SetCopyColProtertys(ByVal objColProtertys As RPTColProtertys, ByRef objColProtertysCopy As RPTColProtertys)
'功能：复制一个关联报表对象
    Dim i As Long
    
    Set objColProtertysCopy = New RPTColProtertys
    For i = 1 To objColProtertys.count
        objColProtertysCopy.Add _
            objColProtertys.Item(i).条件名称, objColProtertys.Item(i).条件字段 _
          , objColProtertys.Item(i).条件关系, objColProtertys.Item(i).条件值 _
          , objColProtertys.Item(i).字体颜色, objColProtertys.Item(i).背景颜色 _
          , objColProtertys.Item(i).是否加粗, objColProtertys.Item(i).是否整行应用 _
          , objColProtertys.Item(i).对齐, "_" & objColProtertys.Item(i).Key
    Next
End Sub

Public Function CopyNewRec(ByVal rsSource As ADODB.Recordset, Optional blnOnlyStructure As Boolean _
    , Optional ByVal strFields As String, Optional arrAppFields As Variant) As ADODB.Recordset
'编制人:朱玉宝
'修改人：刘硕
'修改日期：2014-1-6
'修改点：增加复制记录集的部分字段功能
'编制日期:2000-11-02
'复制记录集
'参数：strFields=需要复制的记录集的字段的列顺序或字段名组成的字符串
'          如：1 别名1,3 别名2,7 别名3...表示复制记录集的第1,3,7..字段组成记录集并返回
'              ID 别名1,姓名 别名2,....表示复制记录集的ID,姓名...字段组成记录集返回
'              别名*为新的记录集的列名
'              两中类型混搭容易出现列名相同的问题，请注意
'           arrAppFields=追加的字段信息：列名,类型,长度,默认值,没有默认值传Empty,没有指定长度传Empty
'      blnOnlyStructure=是否只复制结构
'在程序中，经常会涉及到相互传递记录集，而使用ADO的Clone复制产生的记录集，当其中一个记录集的数据发生变化的时候，所有副本都将发生相同的变化（通常指修改或删除），而我们往往希望这些记录集相互间保持独立
  
    Dim rsClone As ADODB.Recordset
    Dim rsTarget As ADODB.Recordset
    Dim intFields As Integer
    Dim arrFieldsName As Variant, strFieldName As String, strFieldNameAlias As String
    Dim arrTmp As Variant
    Dim i As Long
    
    If Not rsSource Is Nothing Then
        Set rsClone = rsSource.Clone
        rsClone.Filter = rsSource.Filter
    End If
    Set rsTarget = New ADODB.Recordset
    With rsTarget
        '产生记录集结构
        If Not rsClone Is Nothing Then
            If strFields = "" Then '记录集全复制模式
                arrFieldsName = Array()
                If rsClone.Fields.count > 0 Then
                    ReDim arrFieldsName(rsClone.Fields.count - 1)
                Else
                    arrFieldsName = Array()
                End If
                For intFields = 0 To rsClone.Fields.count - 1
                    arrFieldsName(intFields) = rsClone.Fields(intFields).name & ""
                    .Fields.Append rsClone.Fields(intFields).name, IIF(rsClone.Fields(intFields).type = adNumeric, adDouble, rsClone.Fields(intFields).type), rsClone.Fields(intFields).DefinedSize, adFldIsNullable    '0:表示新增
                Next
            Else '记录集部分复制模式
                If rsClone.Fields.count > 0 Then
                    arrFieldsName = Split(strFields, ",")
                    For intFields = LBound(arrFieldsName) To UBound(arrFieldsName)
                        '列包含别名
                        arrTmp = Split(arrFieldsName(intFields) & " ", " ")
                        strFieldName = Trim(arrTmp(0)): strFieldNameAlias = Trim(arrTmp(1))
                        If IsNumeric(strFieldName) Then strFieldName = rsClone.Fields(Val(strFieldName)).name & ""
                        '获取字段原名，存入数组
                        arrFieldsName(intFields) = strFieldName
                        '添加字段,若果存在别名，则新增列的列名为别名
                        .Fields.Append IIF(strFieldNameAlias = "", strFieldName, strFieldNameAlias), IIF(rsClone.Fields(strFieldName).type = adNumeric, adDouble, rsClone.Fields(strFieldName).type), rsClone.Fields(strFieldName).DefinedSize, adFldIsNullable '0:表示新增
                    Next
                End If
            End If
        End If
        '追加字段添加
        If TypeName(arrAppFields) = "Variant()" Then
            For i = LBound(arrAppFields) To UBound(arrAppFields) Step 4
                If arrAppFields(i + 2) = Empty Then
                    If arrAppFields(i + 3) = Empty Then
                        .Fields.Append arrAppFields(i), arrAppFields(i + 1), , adFldIsNullable
                    Else
                        .Fields.Append arrAppFields(i), arrAppFields(i + 1), , adFldIsNullable, arrAppFields(i + 3)
                    End If
                Else
                    If arrAppFields(i + 3) = Empty Then
                        .Fields.Append arrAppFields(i), arrAppFields(i + 1), arrAppFields(i + 2), adFldIsNullable
                    Else
                        .Fields.Append arrAppFields(i), arrAppFields(i + 1), arrAppFields(i + 2), adFldIsNullable, arrAppFields(i + 3)
                    End If
                End If
            Next
        End If
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        '复制数据
        If Not blnOnlyStructure And Not rsClone Is Nothing Then
            If rsClone.RecordCount <> 0 Then rsClone.MoveFirst
            Do While Not rsClone.EOF
                .AddNew
                For intFields = LBound(arrFieldsName) To UBound(arrFieldsName)
                    '新记录集的列按顺序添加，因此可以这样
                    .Fields(intFields).Value = rsClone.Fields(arrFieldsName(intFields)).Value
                Next
                .Update
                rsClone.MoveNext
            Loop
            If rsClone.RecordCount <> 0 Then .Filter = "": .MoveFirst
        End If
    End With
    
    Set CopyNewRec = rsTarget
End Function

Public Sub ApplyOEM(objStatus As Object)
'针对状态栏应用OEM策略
    Dim strOEM As String
    On Error Resume Next
    If gstrProductName = "" Then
        gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "")
    End If
    If gstrProductName <> "-" Then
        objStatus.Panels(1).Text = gstrProductName & "软件"
        '处理状态栏图标的OEM策略
        If gstrProductName = "中联" Then
            Set objStatus.Panels(1).Picture = LoadCustomPicture("Logo")
        Else
            strOEM = GetOEM(gstrProductName)
            Set objStatus.Panels(1).Picture = LoadCustomPicture(strOEM)
            If Err <> 0 Then
                Err.Clear
                Set objStatus.Panels(1).Picture = LoadCustomPicture("Logo")
            End If
        End If
        objStatus.Panels(1).ToolTipText = ""
        objStatus.Height = 360
    End If
End Sub

Public Function IsNumOrChar(ByVal strAsk As String) As Boolean
    '-------------------------------------------------------------
    '功能：判断指定字符串是否全部由数字和英文字母构成，用于允许数字
    '       和字母但不允许特殊字符的情况下的检测，isnumberic只能判断数字。
    '参数：（SSC编制）
    '       strAsk
    '返回：
    '-------------------------------------------------------------
    Dim i As Integer, j As Integer
    
    If Len(Trim(strAsk)) > 0 Then
        For i = 1 To Len(Trim(strAsk))
            j = Asc(Mid(Trim(strAsk), i, 1))
            If Not ((j > 47 And j < 58) Or (j > 64 And j < 91) Or (j > 96 And j < 123)) Then
                IsNumOrChar = False
                Exit Function
            End If
        Next
    End If
    IsNumOrChar = True

End Function

Public Function IsCharAlpha(ByVal strAsk As String) As Boolean
    '-------------------------------------------------------------
    '功能：判断指定字符串是否全部由英文字母构成    '
    '参数：
    '       strAsk
    '返回：
    '-------------------------------------------------------------
    Dim i As Integer, j As Integer
    
    If Len(Trim(strAsk)) > 0 Then
        For i = 1 To Len(Trim(strAsk))
            j = Asc(Mid(Trim(strAsk), i, 1))
            If Not ((j > 64 And j < 91) Or (j > 96 And j < 123)) Then
                IsCharAlpha = False
                Exit Function
            End If
        Next
    End If
    IsCharAlpha = True
End Function

Public Function IsCharChinese(ByVal strAsk As String) As Boolean
    '-------------------------------------------------------------
    '功能：判断指定字符串是否含有汉字
    '参数：
    '       strAsk
    '返回：
    '-------------------------------------------------------------
    Dim i As Integer, j As Integer
    
    If Len(Trim(strAsk)) > 0 Then
        For i = 1 To Len(Trim(strAsk))
            j = Asc(Mid(Trim(strAsk), i, 1))
            If j < 0 Then
                IsCharChinese = True
                Exit Function
            End If
        Next
    End If
    IsCharChinese = False
End Function

Public Function GetAllSubKey(ByVal KeyRoot As Long, KeyName As String) As Variant
'功能：获取注册表某项的所有子项(API方式）
'返回：子项数组
    Dim lngHKey As Long, lngRet As Long, LngIdx As Long
    Dim strName As String
    Dim arrSubKey As Variant
    
    On Error GoTo hErr
    
    arrSubKey = Array()
    LngIdx = 0: strName = String(256, Chr(0))
    lngRet = RegOpenKey(KeyRoot, KeyName, lngHKey)
    If lngRet = 0 Then
        Do
            lngRet = RegEnumKey(lngHKey, LngIdx, strName, Len(strName))
            If lngRet = 0 Then
                ReDim Preserve arrSubKey(UBound(arrSubKey) + 1)
                arrSubKey(UBound(arrSubKey)) = Left(strName, InStr(strName, Chr(0)) - 1)
                LngIdx = LngIdx + 1
            End If
        Loop Until lngRet <> 0
    End If
    RegCloseKey lngHKey
    GetAllSubKey = arrSubKey
    Exit Function
    
hErr:
    RegCloseKey lngHKey
End Function

Public Function GetMemoryParam() As Boolean
'功能：获取数据库存储的“使用个性化风格”参数值
'返回：True个性化；False非个性化

    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo hErr
    
    strSQL = "Select Nvl(a.参数值, b.缺省值) 参数值 " & vbNewLine & _
             "From zlUserParas A, zlParameters B " & vbNewLine & _
             "Where a.参数id = b.Id And a.用户名 = User And b.参数名 = '使用个性化风格' "
    Set rsTmp = OpenSQLRecord(strSQL, "使用个性化风格")
    If rsTmp.RecordCount > 0 Then
        GetMemoryParam = Val(Nvl(rsTmp!参数值)) = 1
    End If
    rsTmp.Close
    Exit Function
    
hErr:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub SetCellValue(ByVal bytOutType As Byte, ByVal objRPTForm As Object, _
    ByVal objCurItem As RPTItem, Optional ByVal lngRowBegin As Long)
'功能：处理所有“指定单元格的标签”元素
'参数：
'  bytOutType：输出类型；0-初预览；1-正式预览；2-正式打印
'  objRPTForm：窗体对象
'  objCurItem：当前表格元素
'  lngRowBegin：开始行号
'说明：
'  标签元素的格式：[表格名称(行号,列号)]
'  备注：行列号均从0开始

    '检查
    If objRPTForm Is Nothing Then Exit Sub
    If objRPTForm.mobjReport Is Nothing Then Exit Sub
    If objCurItem Is Nothing Then Exit Sub
        
    Dim intBegin As Integer, intEnd As Integer, intTmp As Integer
    Dim strValue As String, strResult As String
    Dim strHead As String, strBody As String, strTail As String
    Dim strVSF As String
    Dim lngRow As Long, lngCol As Long
    Dim objItem As RPTItem, objTmp As RPTItem
    Dim vsfObj As VSFlexGrid
    Dim blnFind As Boolean
    Dim lblTmp As Label
    
'    Set vsfObj = objRPTForm.msh(objCurItem.id)
'    If vsfObj Is Nothing Then Exit Sub
    
    For Each objItem In objRPTForm.mobjReport.Items
        '检查
        If objItem.类型 <> 2 Then GoTo makContinue
        
        On Error Resume Next
        If bytOutType > 0 Then
            '正式预览和打印以“内容”输出
            If objItem.Value = "" Then objItem.Value = objItem.内容                         '将原始的文本存到Value属性
        Else
            '初预览以“Caption”显示
            If objItem.Value = "" Then objItem.Value = objRPTForm.lbl(objItem.id).Caption   '将原始的文本存到Value属性
        End If
        strValue = objItem.Value
        
        If Err.Number <> 0 Then
            Err.Clear: On Error GoTo hErr
            GoTo makContinue
        End If
        
        If Not strValue Like "*[[]*(*,*)*[]]*" Then
            GoTo makContinue
        End If
        If strValue Like "*[[]*[[]*" Or strValue Like "*[]]*[]]*" Then
            GoTo makContinue
        End If
        
        '解析元素的内容
        intBegin = InStr(strValue, "[")
        intEnd = InStr(strValue, "]")
        If intBegin > 0 And intEnd > 0 Then
            strHead = Left(strValue, intBegin - 1)
            strTail = Mid(strValue, intEnd + 1)
            strBody = Mid(strValue, intBegin + 1, intEnd - intBegin - 1)
            
            '取表格名称
            intTmp = InStr(strBody, "(")
            If intTmp <= 0 Then intTmp = 1
            strVSF = UCase(Trim(Left(strBody, intTmp - 1)))
            
            '检查表格元素
            blnFind = False
            For Each objTmp In objRPTForm.mobjReport.Items
                If objTmp.类型 = 4 Or objTmp.类型 = 5 Then
                    If Trim(UCase(objTmp.名称)) = strVSF Then
                        Set vsfObj = objRPTForm.msh(objTmp.id)
                        blnFind = True
                        Exit For
                    End If
                End If
            Next
            If blnFind = False Then
                If bytOutType = 2 Then
                    strResult = strHead & strTail
                Else
                    strResult = strHead & "[Error：表格不存在]" & strTail
                End If
                GoSub makSet
                GoTo makContinue
            End If
            
            '取行
            strBody = Mid(strBody, intTmp + 1)
            lngRow = Val(strBody)
            
            '取列
            intTmp = InStr(strBody, ",")
            If intTmp > 0 Then
                lngCol = Val(Mid(strBody, InStr(strBody, ",") + 1))
            Else
                If bytOutType = 2 Then
                    strResult = strHead & strTail
                Else
                    strResult = strHead & "[Error：文本异常]" & strTail
                End If
                GoSub makSet
                GoTo makContinue
            End If
            
            On Error Resume Next
            strBody = vsfObj.TextMatrix(lngRowBegin + lngRow, lngCol)
            If Err.Number <> 0 Then
                Err.Clear:
                If bytOutType = 2 Then
                    strResult = strHead & strTail
                Else
                    strResult = strHead & "[Error：指定单元格不存在]" & strTail
                End If
            Else
                strResult = strHead & strBody & strTail
            End If
            On Error GoTo hErr
            GoSub makSet
        End If

makContinue:
    Next
    
    Exit Sub

makSet:
    If bytOutType > 0 Then
        '正式预览和打印以“内容”输出
        objItem.内容 = strResult
    Else
        '初预览以“Caption”显示
        For Each lblTmp In objRPTForm.lbl   '遍历lbl是防止“初预览”状态下调整报表格式引起异常
            If lblTmp.Index = objItem.id Then
                objRPTForm.lbl(objItem.id).Caption = strResult
                Exit For
            End If
        Next
    End If
    Return
    
hErr:
    Call ErrCenter
End Sub

Public Function TransSpecialChar(ByRef strSQL As String, Optional ByVal blnRestore As Boolean = False) As Boolean
'功能：转换SQL中的特殊字符；如：[]字符，避免与参数的符号冲突
'返回：True成功；False失败

    Const STR_ORIGINAL As String = "[|]|(|)"
    Const STR_TRANS As String = "<左中括号>|<右中括号>|<左括号>|<右括号>"

    Dim strResult As String, strTmp As String
    Dim arrOriginal As Variant, arrTrans As Variant, arrTemp As Variant
    Dim i As Long, j As Long, lngBegin As Long
    Dim intLen As Integer
    
    If Trim(strSQL = "") Then Exit Function
    
    On Error GoTo hErr
    
    strResult = strSQL
    If blnRestore Then
        '还原
        arrOriginal = Split(STR_TRANS, "|")
        arrTrans = Split(STR_ORIGINAL, "|")
    Else
        '转换
        arrOriginal = Split(STR_ORIGINAL, "|")
        arrTrans = Split(STR_TRANS, "|")
    End If
    
    '检查SQL字符里是否存在[]字符
    i = 1
    lngBegin = 0
    Do While Mid(strResult, i) Like "*'*"
        If Mid(strResult, i, 1) = "'" Then
            If lngBegin <= 0 Then
                '开始
                lngBegin = i
            Else
                '结束
                lngBegin = 0
            End If
        Else
            If lngBegin > 0 Then
                '查找''字符内参数的特殊字符，即：SQL语句的字符串
                strTmp = Mid(strResult, lngBegin + 1)
                If InStr(strTmp, "'") > 0 Then
                    strTmp = Left(strTmp, InStr(strTmp, "'") - 1)
                    strTmp = Replace(strTmp, arrTrans(0), arrOriginal(0))
                Else
                    strTmp = ""
                End If
                
                If Not (strTmp Like "*[[][0-9][]]*" Or strTmp Like "*[[][0-9][0-9][]]*") Then
                    For j = LBound(arrOriginal) To UBound(arrOriginal)
                        intLen = Len(arrOriginal(j))
                        If Mid(strResult, i, intLen) = arrOriginal(j) Then
                            strResult = Left(strResult, i - 1) & arrTrans(j) & Mid(strResult, i + intLen)
                        End If
                    Next
                End If
            End If
        End If
        
        i = i + 1
    Loop
    
    strSQL = strResult
    TransSpecialChar = True
    Exit Function
    
hErr:
End Function

Public Function CharCount(ByVal strString As String, ByVal strChar As String) As Long
'功能：获取字符或字符串出现的次数
'返回：字符或字符串出现的次数
    Dim lngA As Long, lngB As Long, lngC As Long
    
    lngA = Len(strString)
    lngB = Len(strChar)
    lngC = Len(Replace(strString, strChar, ""))
    CharCount = (lngA - lngC) / lngB
End Function

Public Function AtString(ByVal strVal As String) As Boolean
'功能：判断字符串是的单引号是单、双数；单数代表字符串，双数代表非字符串
'返回：True字符串；False非字符串
    
    AtString = (CharCount(strVal, "'") Mod 2) = 1
End Function

Public Sub SetControlDBConnect(ByRef vControl As Variant)
'功能：加载数据连接信息至控件

    Dim strSQL As String, strResult As String
    Dim rsTemp As ADODB.Recordset
    Dim cbiTmp As ComboItem
    
    On Error GoTo hErr
    
    '数据获取
    strSQL = _
            "Select 编号, 名称, 用户名, 密码, Ip, 端口, 实例名, 说明 " & vbNewLine & _
            "From ZlConnections " & vbNewLine & _
            "Order By 编号 "
    Set rsTemp = mdlPublic.OpenSQLRecord(strSQL, "获取数据连接信息")
    
    '数据加载
    Select Case UCase(TypeName(vControl))
    Case "COMBOBOX"
        Do While rsTemp.EOF = False
            vControl.AddItem "【" & Nvl(rsTemp!编号) & "】" & _
                             Nvl(rsTemp!名称) & _
                             ""
'                             " （" & _
'                             "IP：" & Nvl(rsTemp!IP) & "；" & _
'                             "端口：" & Nvl(rsTemp!端口) & "；" & _
'                             "服务器：" & Nvl(rsTemp!实例名) & _
'                             "）"
            vControl.ItemData(vControl.NewIndex) = Nvl(rsTemp!编号, 0)
            rsTemp.MoveNext
        Loop
    Case "RECORDSET"
        Set vControl = CopyNewRec(rsTemp)
    End Select
    rsTemp.Close
    
    Exit Sub
    
hErr:
    If ErrCenter = 1 Then Resume
End Sub

Private Function NumericPassword(ByVal password As String) As Long
    Dim Value As Long
    Dim ch As Long
    Dim shift1 As Long
    Dim shift2 As Long
    Dim i As Integer
    Dim str_len As Integer

    str_len = Len(password)
    For i = 1 To str_len
        ch = Asc(Mid$(password, i, 1))
        Value = Value Xor (ch * 2 ^ shift1)
        Value = Value Xor (ch * 2 ^ shift2)
        shift1 = (shift1 + 7) Mod 19
        shift2 = (shift2 + 13) Mod 23
    Next i
    NumericPassword = Value
End Function

Private Sub Base64EncodeByte(mInByte() As Byte, mOutByte() As Byte, Num As Integer)
    Dim tByte     As Byte
    Dim i     As Integer
    
    If Num = 1 Then
      mInByte(1) = 0
      mInByte(2) = 0
    ElseIf Num = 2 Then
      mInByte(2) = 0
    End If
    tByte = mInByte(0) And &HFC
    mOutByte(0) = tByte / 4
    tByte = ((mInByte(0) And &H3) * 16) + (mInByte(1) And &HF0) / 16
    mOutByte(1) = tByte
    tByte = ((mInByte(1) And &HF) * 4) + ((mInByte(2) And &HC0) / 64)
    mOutByte(2) = tByte
    tByte = (mInByte(2) And &H3F)
    mOutByte(3) = tByte
    For i = 0 To 3
      If mOutByte(i) >= 0 And mOutByte(i) <= 25 Then
        mOutByte(i) = mOutByte(i) + Asc("A")
      ElseIf mOutByte(i) >= 26 And mOutByte(i) <= 51 Then
        mOutByte(i) = mOutByte(i) - 26 + Asc("a")
      ElseIf mOutByte(i) >= 52 And mOutByte(i) <= 61 Then
        mOutByte(i) = mOutByte(i) - 52 + Asc("0")
      ElseIf mOutByte(i) = 62 Then
        mOutByte(i) = Asc("+")
      Else
        mOutByte(i) = Asc("/")
      End If
    Next i
    If Num = 1 Then
      mOutByte(2) = Asc("=")
      mOutByte(3) = Asc("=")
    ElseIf Num = 2 Then
      mOutByte(3) = Asc("=")
    End If
End Sub

Private Function Base64Encode(InStr1 As String) As String
    Dim mInByte(3)     As Byte, mOutByte(4)       As Byte
    Dim myByte     As Byte
    Dim i     As Integer, LenArray       As Integer, j       As Integer
    Dim myBArray()     As Byte
    Dim OutStr1     As String
    
    myBArray() = StrConv(InStr1, vbFromUnicode)
    LenArray = UBound(myBArray) + 1
    For i = 0 To LenArray Step 3
      If LenArray - i = 0 Then
        Exit For
      End If
      If LenArray - i = 2 Then
        mInByte(0) = myBArray(i)
        mInByte(1) = myBArray(i + 1)
        Base64EncodeByte mInByte, mOutByte, 2
      ElseIf LenArray - i = 1 Then
        mInByte(0) = myBArray(i)
        Base64EncodeByte mInByte, mOutByte, 1
      Else
        mInByte(0) = myBArray(i)
        mInByte(1) = myBArray(i + 1)
        mInByte(2) = myBArray(i + 2)
        Base64EncodeByte mInByte, mOutByte, 3
      End If
      For j = 0 To 3
        OutStr1 = OutStr1 & Chr(mOutByte(j))
      Next j
    Next i
    Base64Encode = OutStr1
    
End Function

Public Function Decipher(ByVal vPassword As String, ByVal vFrom_Text As String) As String
    '解密
    Const MIN_ASC = 32
    Const MAX_ASC = 126
    Const NUM_ASC = MAX_ASC - MIN_ASC + 1
    
    Dim offset As Long
    Dim str_len As Integer
    Dim i As Integer
    Dim ch As Integer
    
    vPassword = Base64Encode(vPassword) & "WIZARDPAGE"
    
    offset = NumericPassword(vPassword)
    Rnd -1
    Randomize offset

    str_len = Len(vFrom_Text)
    For i = 1 To str_len
        ch = Asc(Mid$(vFrom_Text, i, 1))
        If ch >= MIN_ASC And ch <= MAX_ASC Then
            ch = ch - MIN_ASC
            offset = Int((NUM_ASC + 1) * Rnd)
            ch = ((ch - offset) Mod NUM_ASC)
            If ch < 0 Then ch = ch + NUM_ASC
            ch = ch + MIN_ASC
            Decipher = Decipher & Chr$(ch)
        End If
    Next i
End Function

Public Function GetDBConnectInfo(ByVal intDBConnectNo As Integer, Optional ByVal bytType As Byte = 0) As String
'功能：通过intDBConnectNo参数，获取数据连接信息
'参数：
'  intDBConnectNo：数据连接编号
'  bytType：0-指定返回数据连接的名称；1-用户名

    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset

    On Error GoTo hErr
    
    strSQL = "Select 名称, 用户名, 密码, Ip, 端口, 实例名, 说明 From Zlconnections Where 编号 = [1] "
    Set rsTemp = mdlPublic.OpenSQLRecord(strSQL, "获取其他数据连接信息", intDBConnectNo)
    If rsTemp.EOF = False Then
        Select Case bytType
        Case 0
            GetDBConnectInfo = Nvl(rsTemp!名称)
        Case 1
            GetDBConnectInfo = Nvl(rsTemp!用户名)
        End Select
    End If
    rsTemp.Close
    Exit Function
    
hErr:
    If ErrCenter = 1 Then Resume
End Function

Public Function ValEx(ByVal strVal As String) As Double
    ValEx = Val(Replace(strVal, ",", ""))
End Function

Public Function GetStdNodeText(ByVal strText As String) As String
    If strText Like "*（*）" Then
        strText = Left(strText, InStrRev(strText, "（") - 1)
        GetStdNodeText = strText
    Else
        GetStdNodeText = strText
    End If
End Function

Public Function GetSysVersion(ByVal lngSysNO As Long) As String
'功能：获取指定系统的版本号
'参数：
'  lngSysNo：系统编号
'返回：版本号
    
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo hErr
    strSQL = "Select 版本号 From zlSystems Where 编号 = [1]"
    Set rsTemp = mdlPublic.OpenSQLRecord(strSQL, "获取系统版本", lngSysNO)
    If rsTemp.EOF = False Then
        GetSysVersion = mdlPublic.Nvl(rsTemp!版本号)
    End If
    rsTemp.Close
    Exit Function
    
hErr:
    If mdlPublic.ErrCenter = 1 Then Resume
End Function

Public Function GetPubIcons() As XtremeCommandBars.ImageManagerIcons
    Set GetPubIcons = frmPubIcons.imgPublic.Icons
End Function

Public Function SetPublicFontSize(ByRef frmMe As Object, ByVal bytSize As Byte, Optional ByVal strOther As String)
'功能：设置窗体及所有控件的字体大小
'参数：frmMe=需要设置字体的窗体对象
'      bytSize:设置为9号字体,0:设置为9号字体,1,设置为12号字体
'      strOther:不进行字体设置的控件父容器的集合,格式为：容器名字1,容器名字2,容器名字3,....
'说明：1.如果涉及到VsFlexGrid等表格控件，需要根据所在的环境重新调整列宽和行高
'      2.如果存在未列出的其他控件或自定义控件,需要用特定方法指定字体大小及相关处理的，需另外单独设置

    Dim objCtrol As Control
    Dim objFont As StdFont
    Dim objRPTCol As Object
    Dim i As Long, lngOldSize As Long, lngFontSize As Long
    Dim dblRate As Double
    Dim blnDo As Boolean
    Dim strContainer As String
    
    lngFontSize = IIF(bytSize = 0, 9, IIF(bytSize = 1, 12, bytSize))
    frmMe.FontSize = lngFontSize
    strOther = "," & strOther & ","
    blnDo = False
        
    For Each objCtrol In frmMe.Controls
        Select Case TypeName(objCtrol)
            Case "TabStrip", "Label", "ComboBox", "ListView", "OptionButton", "CheckBox", "DTPicker", "TextBox", "SpeedButton", _
                "DockingPane", "CommandBars", "TabControl", "CommandButton", "Frame", "RichTextBox", "MaskEdBox", "IDKindNew", _
                "VSFlexGrid", "StatusBar", "ReportControl"
                blnDo = True
            Case Else
                blnDo = False
        End Select
        
        If strOther <> ",," And blnDo Then
            '对于CommandBars用户自定义控件读取objCtrol.Container会出错
            strContainer = ""
            On Error Resume Next
            strContainer = objCtrol.Container.name
            Err.Clear: On Error GoTo 0
            If InStr(1, strOther, "," & strContainer & ",") > 0 Then
                 blnDo = False
            End If
        End If
        
        If blnDo Then
            Select Case TypeName(objCtrol)
            Case "TabStrip"
                objCtrol.Font.Size = lngFontSize
            Case "Label"
                If Not LCase(objCtrol.name) Like "*_fixed" Then
                    lngOldSize = objCtrol.Font.Size
                    dblRate = lngFontSize / lngOldSize
                    
                    objCtrol.Font.Size = lngFontSize
                    objCtrol.Height = frmMe.TextHeight("字") + 20
                    'Label宽度需要自行调整
                End If
            Case "ComboBox"
                 lngOldSize = objCtrol.Font.Size
                 dblRate = lngFontSize / lngOldSize
                 
                 objCtrol.Font.Size = lngFontSize
                 objCtrol.Width = objCtrol.Width * dblRate
            Case "ListView"
                lngOldSize = objCtrol.Font.Size
                dblRate = lngFontSize / lngOldSize
                
                objCtrol.Font.Size = lngFontSize
                For i = 1 To objCtrol.ColumnHeaders.count
                    objCtrol.ColumnHeaders(i).Width = objCtrol.ColumnHeaders(i).Width * dblRate
                Next
            Case "OptionButton"
                lngOldSize = objCtrol.Font.Size
                dblRate = lngFontSize / lngOldSize
                
                objCtrol.Font.Size = lngFontSize
                objCtrol.Width = frmMe.TextWidth("字体" & objCtrol.Caption)
                objCtrol.Height = objCtrol.Height * dblRate
            Case "CheckBox"
                lngOldSize = objCtrol.Font.Size
                dblRate = lngFontSize / lngOldSize
                
                objCtrol.Font.Size = lngFontSize
                objCtrol.Width = objCtrol.Width * dblRate
            Case "DTPicker"
                lngOldSize = objCtrol.Font.Size
                dblRate = lngFontSize / lngOldSize
                
                objCtrol.Font.Size = lngFontSize
                objCtrol.Width = frmMe.TextWidth("2012-01-01    ")
                objCtrol.Height = frmMe.TextHeight("字") + IIF(bytSize = 0, 100, 120)
            Case "TextBox"
                lngOldSize = objCtrol.Font.Size
                dblRate = lngFontSize / lngOldSize
                
                objCtrol.Font.Size = lngFontSize
                objCtrol.Width = objCtrol.Width * dblRate
                objCtrol.Height = frmMe.TextHeight("字")
            Case "MaskEdBox"
                objCtrol.FontSize = lngFontSize
                objCtrol.Width = frmMe.TextWidth(objCtrol.Mask)
                objCtrol.Height = frmMe.TextHeight("字")
            Case "ReportControl"
                lngOldSize = objCtrol.PaintManager.TextFont.Size
                dblRate = lngFontSize / lngOldSize

                Set objFont = objCtrol.PaintManager.CaptionFont
                objFont.Size = lngFontSize
                Set objCtrol.PaintManager.CaptionFont = objFont
                Set objFont = objCtrol.PaintManager.TextFont
                objFont.Size = lngFontSize
                Set objCtrol.PaintManager.TextFont = objFont
                For Each objRPTCol In objCtrol.Columns
                    objRPTCol.Width = objRPTCol.Width * dblRate
                Next
                objCtrol.Redraw
            Case "SpeedButton"
                Dim objFontTemp As New StdFont
                
                Set objFontTemp = frmMe.Font
                If bytSize = 0 Then
                    objFontTemp.Size = 12
                    dblRate = 0.8
                Else
                    objFontTemp.Size = 15.75
                    dblRate = 1 / 0.8
                End If
                Set objCtrol.Font = objFontTemp
                objCtrol.Width = objCtrol.Width * dblRate
            Case "VSFlexGrid"
                Set objCtrol.Font = frmMe.Font
                objCtrol.Font.Size = IIF(bytSize = 0, 9, 12)
            Case "DockingPane"
                Set objFont = objCtrol.PaintManager.CaptionFont
                If objFont Is Nothing Then '控件初始加载时objFont为nothing
                    Set objFont = frmMe.Font
                End If
                objFont.Size = lngFontSize
                Set objCtrol.PaintManager.CaptionFont = objFont
                
                Set objFont = objCtrol.TabPaintManager.Font
                If objFont Is Nothing Then '控件初始加载时objFont为nothing
                    Set objFont = frmMe.Font
                End If
                objFont.Size = lngFontSize
                Set objCtrol.TabPaintManager.Font = objFont

                Set objFont = objCtrol.PanelPaintManager.Font
                If objFont Is Nothing Then '控件初始加载时objFont为nothing
                    Set objFont = frmMe.Font
                End If
                objFont.Size = lngFontSize
                Set objCtrol.PanelPaintManager.Font = objFont
            Case "CommandBars"
                Set objFont = objCtrol.Options.Font
                If objFont Is Nothing Then '控件初始加载时objFont为nothing
                    Set objFont = frmMe.Font
                End If
                objFont.Size = lngFontSize
                Set objCtrol.Options.Font = objFont
            Case "TabControl"
                Set objFont = objCtrol.PaintManager.Font
                If objFont Is Nothing Then  '控件初始加载时objFont为nothing
                    Set objFont = frmMe.Font
                End If
                objFont.Size = lngFontSize
                Set objCtrol.PaintManager.Font = objFont
                objCtrol.PaintManager.Layout = xtpTabLayoutAutoSize
            Case "CommandButton"
                If Not LCase(objCtrol.name) Like "*_fixed" Then
                    lngOldSize = objCtrol.FontSize
                    dblRate = lngFontSize / lngOldSize

                    objCtrol.FontSize = lngFontSize
                    objCtrol.Width = dblRate * objCtrol.Width
                    objCtrol.Height = dblRate * objCtrol.Height
                End If
            Case "Frame"
                objCtrol.FontSize = lngFontSize
            Case "StatusBar"
                objCtrol.Font.Size = lngFontSize
            End Select
        End If
    Next
End Function

Public Function FormatString(ByVal strFormat As String, ParamArray arrParams() As Variant) As String
'功能：格式化字符串
'参数：
'  strFormat：表达式；[1-x]为参数号关键字；例子："测试值为：[1]"
'  arrParams：表达式的参数，对应strFormat中的参数号关键字
'返回：格式化后的字符串

    Dim i As Integer, intSN As Integer
    Dim strKey As String, strTmp As String
    Dim blnStart As Boolean

    FormatString = strFormat

    If Len(strFormat) > 60000 Then Exit Function
    If Not strFormat Like "*[[]*[]]*" Then Exit Function
    If UBound(arrParams) < 0 Then Exit Function

    On Error GoTo errHandle

    For i = 1 To Len(strFormat)
        If Mid(strFormat, i, 1) = "[" Then
            blnStart = True
        End If
        If blnStart Then
            If Mid(strFormat, i, 1) = "]" Then
                intSN = Val(Mid(strKey, 2))
                If intSN > 0 Then
                    If UBound(arrParams) >= intSN - 1 Then
                        strTmp = strTmp & arrParams(intSN - 1)
                    End If
                Else
                    strTmp = strTmp & Mid(strKey, 2)
                End If
                blnStart = False
                strKey = ""
            Else
                strKey = strKey & Mid(strFormat, i, 1)
            End If
        Else
            strTmp = strTmp & Mid(strFormat, i, 1)
        End If
    Next

    FormatString = strTmp
    Exit Function

errHandle:
End Function

Public Sub AddArray(ByRef cllData As Collection, ByVal strSQL As String)
'功能：将SQL写入集合
'参数：
'  cllData：集合对象
'  strSQL：SQL字符串

    Dim l As Long
    
    l = cllData.count + l
    cllData.Add strSQL, "K" & l
End Sub

Public Sub InitClass(ByVal objControl As ComboBox _
    , Optional ByVal varDefault As Variant _
    , Optional ByVal lngCurClassID As Long = 0)
'功能：初始化报表分类控件
'参数：
'  objControl：初始化的控件
'  varDefault：指定默认选项
'  lngCurClassID：当前报表分类ID，即不加载对应的分类ID到控件中
    
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo hErr
    
    objControl.Clear
    
    If lngCurClassID > 0 Then
        strSQL = _
            "Select a.ID, a.上级ID, a.名称, a.说明 " & vbCrLf & _
            "From ZlRPTClasses A " & vbCrLf & _
            "Where not a.ID in (Select ID From ZlRPTClasses Start With ID = [1] Connect By Prior ID = 上级id) " & vbCrLf & _
            "Order By a.名称 "
    Else
        strSQL = _
            "Select ID, 上级ID, 名称, 说明 " & vbCrLf & _
            "From ZlRPTClasses " & vbCrLf & _
            "Order By 名称 "
    End If
    Set rsTemp = mdlPublic.OpenSQLRecord(strSQL, "获取报表分类信息", lngCurClassID)
    
    objControl.AddItem ""
    Do While rsTemp.EOF = False
        objControl.AddItem rsTemp!名称
        objControl.ItemData(objControl.NewIndex) = rsTemp!id
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    
    '设置默认选项
    Select Case UCase(TypeName(varDefault))
    Case "INTEGER", "LONG"
        If varDefault > -1 And objControl.ListCount > 0 Then
            For i = 0 To objControl.ListCount - 1
                If varDefault = objControl.ItemData(i) Then
                    objControl.ListIndex = i
                    Exit For
                End If
            Next
        End If
    Case "STRING"
        objControl.Text = varDefault
        If varDefault <> "" And objControl.ListCount > 0 Then
            For i = 0 To objControl.ListCount - 1
                If varDefault = objControl.List(i) Then
                    objControl.ListIndex = i
                    Exit For
                End If
            Next
        End If
    End Select
    
    Exit Sub
    
hErr:
    If ErrCenter = 1 Then Resume
End Sub

Public Function IsDebugging() As Boolean
    On Error Resume Next
    Debug.Print 1 / 0
    IsDebugging = Err.Number <> 0
    On Error GoTo 0
End Function

Public Function ReportStateSwitch(ByVal lngSysNO As Long, ByVal varRPT As Variant _
    , ByVal blnGroup As Boolean, Optional ByRef strInfo As String) As Integer
'功能：获取指定报表的启停状态
'参数：
'  lngSysNO：系统号
'  varRPT：报表编号或程序ID
'  blnGroup：True-报表组
'  strInfo（实参）：报表编号和名称；如：【编号】名称
'返回：-1-异常；0-未发布或报表不存在；1-启用（发布）；2-停用（发布）
    
    Dim strSQL As String, strVar As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo hErr
    
    ReportStateSwitch = -1
    strInfo = ""
    strVar = UCase(varRPT)
    
    If blnGroup Then
        '报表组
        strSQL = "Select 发布时间, 是否停用, 编号, 名称, 0 子报表 From zlRPTGroups " & vbCr & _
                 "Where " & IIF(lngSysNO <= 0, " 系统 Is Null ", " 系统 = [1] ")
        If IsNumeric(strVar) Then
            strSQL = strSQL & " And 程序ID = [2] "
            Set rsTemp = mdlPublic.OpenSQLRecord(strSQL, "获取报表启停状态", lngSysNO, CLng(strVar))
        Else
            strSQL = strSQL & " And 编号 = [2] "
            Set rsTemp = mdlPublic.OpenSQLRecord(strSQL, "获取报表启停状态", lngSysNO, strVar)
        End If
    Else
        '报表或子报表
        strSQL = "Select a.发布时间, a.是否停用, a.编号, a.名称, b.报表id 子报表 " & vbCr & _
                 "From zlReports A, zlRPTSubs B " & vbCr & _
                 "Where a.Id = b.报表id(+) " & _
                 IIF(lngSysNO <= 0, " And a.系统 Is Null ", " And a.系统(+) = [1] ")
        If IsNumeric(strVar) Then
            strSQL = strSQL & " And a.程序ID(+) = [2] "
            Set rsTemp = mdlPublic.OpenSQLRecord(strSQL, "获取报表启停状态", lngSysNO, CLng(strVar))
        Else
            strSQL = strSQL & " And a.编号(+) = [2] "
            Set rsTemp = mdlPublic.OpenSQLRecord(strSQL, "获取报表启停状态", lngSysNO, strVar)
        End If
    End If
    
    If rsTemp.EOF Then
        ReportStateSwitch = 0
    Else
        strInfo = mdlPublic.FormatString("【[1]】[2]", mdlPublic.Nvl(rsTemp!编号), mdlPublic.Nvl(rsTemp!名称))
        If blnGroup Or Val(mdlPublic.Nvl(rsTemp!子报表)) = 0 Then
            '报表组或独立报表
            If mdlPublic.Nvl(rsTemp!发布时间) = "" And lngSysNO = 0 Then
                ReportStateSwitch = 0
            Else
                ReportStateSwitch = IIF(mdlPublic.Nvl(rsTemp!是否停用, 0) = 1, Val("2-停用"), Val("1-启用"))
            End If
        Else
            '子报表
            ReportStateSwitch = IIF(mdlPublic.Nvl(rsTemp!是否停用, 0) = 1, Val("2-停用"), Val("1-启用"))
        End If
    End If
    rsTemp.Close
    
    Exit Function
    
hErr:
    If ErrCenter = 1 Then Resume
End Function

Public Function RPTParsCondExec(ByVal vRPTID As Long, ByVal vCondID As Long, ByVal vRPTPars As RPTPars) As RPTPars
'功能：报表执行参数的“条件”选择
'参数：
'  vRPTID：报表ID
'  vCondID：条件号
'  vRPTPars：默认的报表参数对象
'返回：新的RPTPars对象

    Dim strSQL As String, strValue As String, strDefault As String
    Dim rsTmp As ADODB.Recordset
    Dim blnRetry As Boolean
    Dim objNewCond As New RPTPars
    Dim objRPTPar As RPTPar
    Dim i As Integer
    
    On Error GoTo hErr
    
    '取指定的条件
    blnRetry = True
    strSQL = "Select 参数名,参数值 From zlRptConds Where 报表ID=[1] And 条件号=[2]"
    Set rsTmp = OpenSQLRecord(strSQL, "获取报表参数信息", vRPTID, vCondID)
    blnRetry = False
    
    '产生一个参数对象
    For i = 1 To vRPTPars.count
        Set objRPTPar = vRPTPars(i)
        rsTmp.Filter = "参数名='" & objRPTPar.名称 & "'"
        If rsTmp.RecordCount > 0 Then
            '解析参数
            strValue = Nvl(rsTmp!参数值)
            strDefault = objRPTPar.缺省值
            If InStr(1, "固定值列表…,选择器定义…", objRPTPar.缺省值) <> 0 And objRPTPar.缺省值 <> "" Then
                If InStr(1, strValue, "|") > 0 Then
                    strValue = Split(strValue, "|")(1)
                    If InStr(1, strValue, "!") > 0 Then
                        strValue = Replace(strValue, "!", "|")
                    End If
                End If
            Else
                strDefault = Nvl(rsTmp!参数值)
                strValue = objRPTPar.缺省值
            End If
        Else
            '
            strDefault = objRPTPar.缺省值
            strValue = objRPTPar.Reserve
        End If
        objNewCond.Add objRPTPar.组名, objRPTPar.序号, objRPTPar.名称 _
            , objRPTPar.类型, strDefault, objRPTPar.格式, objRPTPar.值列表 _
            , objRPTPar.分类SQL, objRPTPar.明细SQL, objRPTPar.分类字段 _
            , objRPTPar.明细字段, objRPTPar.对象, "_" & objRPTPar.Key _
            , strValue, objRPTPar.是否锁定
    Next
    rsTmp.Close
    
    Set RPTParsCondExec = objNewCond
    Exit Function
    
hErr:
    If blnRetry Then
        If ErrCenter = 1 Then Resume
    Else
        Call ErrCenter
    End If
End Function

Public Function RPTParsCondSave(ByVal vReportID As Long, ByVal vCondID As Integer _
    , ByVal vPars As RPTPars, ByVal vParsDefault As RPTPars, ByVal vForm As Form _
    , Optional ByVal vIsSaveAs As Boolean = False) As Boolean
'功能：保存报表参数条件
'参数：
'  vReportID：报表ID
'  vCondID：条件号
'返回：True成功；False失败

    Dim i As Integer, j As Integer
    Dim strTmp As String, strDisp As String
    Dim strParName As String
    Dim strSQL As String, strCondName As String, strTitle As String
    Dim intCondID As Integer
    Dim rsTmp As New ADODB.Recordset
    Dim blnRetry As Boolean
    Dim objRPTPar As RPTPar
    Dim objPop As Object, lbl As Object
    
    On Error GoTo hErr
    
    '条件名称
    blnRetry = True
    If vCondID = 0 Or vIsSaveAs Then
        '取最大条件号
        strSQL = "Select Max(条件号) 条件号 From zlRptConds Where 报表ID=[1]"
        Set rsTmp = OpenSQLRecord(strSQL, "获取报表的最大条件号", vReportID)
        intCondID = Nvl(rsTmp!条件号, 0) + 1
        
        strCondName = InputBox("请输入条件名称（输入空的名称等同于取消）", "保存条件", "条件" & intCondID)
        If Trim(Replace(strCondName, "'", "")) = "" Then Exit Function
    Else
        '已有条件名称
        intCondID = vCondID
        strSQL = "Select 条件名称 From zlRptConds Where 报表ID=[1] And 条件号=[2]"
        Set rsTmp = OpenSQLRecord(strSQL, "获取报表的条件名称", vReportID, intCondID)
        strCondName = Nvl(rsTmp!条件名称)
    End If
    blnRetry = False
    
    If UCase(vForm.name) = UCase("frmReport") Then
        Set objPop = vForm.mnuPop_Cond
        Set lbl = vForm.lblName
        strTitle = vForm.mobjReport.名称
    Else
        Set objPop = vForm.PopMenu_Cond
        Set lbl = vForm.lbl
        strTitle = vForm.mstrTitle
    End If
    
    '再取值
    For i = 1 To lbl.UBound
        strParName = lbl(i).ToolTipText
        Set objRPTPar = vPars("_" & strParName)
        If objRPTPar Is Nothing Then GoTo makContinue
        
        If objRPTPar.缺省值 = "固定值列表…" Then
            Select Case objRPTPar.格式
            Case Val("0-下拉框")
                If GetCboIndex(vForm.cbo(i), vForm.cbo(i).Text) = -1 Then '是否人为输入
                    'Reserve字段保存本次条件的"宏条件值|显示值"
                    objRPTPar.Reserve = "固定值列表…|" & vForm.cbo(i).Text
                    objRPTPar.缺省值 = vForm.cbo(i).Text
                Else
                    '列表选择
                    'Reserve字段保存本次条件的"宏条件值|显示值"
                    objRPTPar.Reserve = "固定值列表…|" & vForm.cbo(i).Text
                    strTmp = objRPTPar.值列表
                    For j = 0 To UBound(Split(strTmp, "|"))
                        strDisp = Split(Split(strTmp, "|")(j), ",")(0)
                        If Left(strDisp, 1) = "√" Then
                            strDisp = Mid(strDisp, 2)
                        End If
                        If strDisp = vForm.cbo(i).Text Then
                            objRPTPar.缺省值 = Split(Split(strTmp, "|")(j), ",")(1)
                            Exit For
                        End If
                    Next
                End If
            Case Val("1-单选框")
                For j = 1 To vForm.opt.UBound
                    If vForm.opt(j).Container.Index = i Then
                        If vForm.opt(j).Value Then
                            'Reserve字段保存本次条件的"宏条件值|显示值"
                            objRPTPar.Reserve = "固定值列表…|" & vForm.opt(j).ToolTipText
                            objRPTPar.缺省值 = vForm.opt(j).Tag
                        End If
                    End If
                Next
            Case Val("2-复选框")
                'Reserve字段保存本次条件的"宏条件值|显示值"
                strTmp = objRPTPar.值列表
                For j = 0 To 1
                    strDisp = Split(Split(strTmp, "|")(j), ",")(0)
                    If vForm.chk(i).Value = 0 Then
                        If Left(strDisp, 1) <> "√" Then
                            objRPTPar.Reserve = "固定值列表…|" & strDisp
                            objRPTPar.缺省值 = Split(Split(strTmp, "|")(j), ",")(1)
                        End If
                    Else
                        If Left(strDisp, 1) = "√" Then
                            objRPTPar.Reserve = "固定值列表…|" & Mid(strDisp, 2)
                            objRPTPar.缺省值 = Split(Split(strTmp, "|")(j), ",")(1)
                        End If
                    End If
                Next
            End Select
        ElseIf objRPTPar.缺省值 = "选择器定义…" Then
            If vForm.txt(i).Tag = "" Then '是否人为输入
                'Reserve字段保存本次条件的"宏条件值|显示值"
                objRPTPar.Reserve = "选择器定义…|"
                objRPTPar.缺省值 = vForm.txt(i).Text
            Else
                '列表选择
                'Reserve字段保存本次条件的"宏条件值|显示值"
                objRPTPar.Reserve = "选择器定义…|" & vForm.txt(i).Text
                objRPTPar.缺省值 = vForm.txt(i).Tag
            End If
        Else
            Select Case objRPTPar.类型
            Case Val("0-字符"), Val("1-数字"), Val("3-无类型")
                objRPTPar.缺省值 = vForm.txt(i).Text
            Case Val("2-日期")
                If objRPTPar.缺省值 Like "&*" Then
                    objRPTPar.Reserve = objRPTPar.缺省值
                End If
                objRPTPar.缺省值 = Format(vForm.dtp(i).Value, vForm.dtp(i).CustomFormat)

'                '保存到注册表
'                If vForm.dtp(i).CustomFormat Like "*HH:mm:ss" Then
'                    SaveSetting "ZLSOFT" _
'                        , "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & vForm.name & strTitle _
'                        , lbl(i).ToolTipText & "时间" _
'                        , Format(vForm.dtp(i).Value, "HH:mm:ss")
'                End If
            End Select
        End If
        
makContinue:
    Next
    
    '产生保存串
    strSQL = ""
    For i = 1 To vParsDefault.count
        Set objRPTPar = vParsDefault(i)
        If objRPTPar.缺省值 = "固定值列表…" Then
            strSQL = strSQL & "!!" & vPars(i).名称 & "," & vPars(i).Reserve & "!" & Replace(vPars(i).缺省值, "'", "''")
        ElseIf vParsDefault(i).缺省值 = "选择器定义…" Then
            strSQL = strSQL & "!!" & vPars(i).名称 & "," & vPars(i).Reserve & "!" & Replace(vPars(i).缺省值, "'", "''")
        Else
            strSQL = strSQL & "!!" & vPars(i).名称 & "," & Replace(vPars(i).缺省值, "'", "''")
        End If
    Next
    strSQL = "zl_RptConds_Update(" & _
             vReportID & "," & _
             intCondID & "," & _
             "'" & strCondName & "'," & _
             "'" & Mid(strSQL, 3) & "'," & _
             IIF(vIsSaveAs, 0, vCondID) & ")"
    Call gcnOracle.Execute(strSQL, , adCmdStoredProc)
    
    '加入菜单
    If vCondID = 0 Or vIsSaveAs Then
        i = objPop.count
        Load objPop(i)
        With objPop(i)
            .Caption = strCondName & "(&" & intCondID & ")"
            .Visible = True
            .Tag = intCondID
        End With
    End If
    
    RPTParsCondSave = True
    Exit Function
    
hErr:
    If blnRetry Then
        If ErrCenter = 1 Then Resume
    Else
        Call ErrCenter
    End If
End Function

Public Function RPTParsCondDel(ByVal vRPTID As Long, ByVal vCondID As Integer) As Boolean
    Dim strSQL As String, strCondName As String
    Dim rsTmp As ADODB.Recordset
    Dim blnRetry As Boolean

    If vRPTID <= 0 Then Exit Function
    If vCondID <= 0 Then Exit Function
    
    On Error GoTo hErr
    
    blnRetry = True
    strSQL = "Select 条件名称 From zlRptConds Where 报表ID=[1] And 条件号=[2]"
    Set rsTmp = OpenSQLRecord(strSQL, "获取报表的参数条件", vRPTID, vCondID)
    blnRetry = False
    
    strCondName = Nvl(rsTmp!条件名称)
    If MsgBox("你确定要删除“" & strCondName & "”吗？", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Function
    
    strSQL = "zl_RptConds_Update(" & vRPTID & "," & vCondID & ",'条件名称','',0,1)"
    Call gcnOracle.Execute(strSQL, , adCmdStoredProc)
    
    RPTParsCondDel = True
    Exit Function
    
hErr:
    If blnRetry Then
        If ErrCenter = 1 Then Resume
    Else
        Call ErrCenter
    End If
End Function

Private Function CheckTblExist(ByVal strTableName As String) As Boolean
    '功能：根据表名判断表是否存在
    '参数：strTableName - 要查询的表名
    Dim strSQL As String, rsData As ADODB.Recordset
    
    On Error Resume Next
    strSQL = "select 1 from " & strTableName & " where rownum<1 "
    Set rsData = OpenSQLRecord(strSQL, "CheckTblExist")
    
    CheckTblExist = Err.Number = 0
    Err.Clear
End Function

Public Function GetDBConnectNo(ByVal objVar As RPTPar, ByVal objDatas As RPTDatas) As Integer
'功能：通过对象获取对应的数据连接编号

    Dim objData As RPTData
    Dim objPar As RPTPar
    
    If objVar Is Nothing Then Exit Function
    If objDatas Is Nothing Then Exit Function
    
    For Each objData In objDatas
        For Each objPar In objData.Pars
            If objVar.名称 = objPar.名称 Then
                GetDBConnectNo = objData.数据连接编号
                Exit Function
            End If
        Next
    Next
End Function

Public Function GetInsertProgPrivs(ByVal lngSystem As Long, ByVal lngModule As Long _
    , ByVal strFunction As String, ByVal strDBObject As String _
    , ByVal strOwner As String, ByVal strOperation As String) As String

    Dim strSQL As String

    On Error GoTo hErr
    
    strDBObject = UCase(strDBObject)
    strOwner = UCase(strOwner)
    strOperation = UCase(strOperation)
    
    strSQL = "Insert Into zlProgPrivs (系统,序号,功能,对象,所有者,权限) " & vbCrLf & _
             "Select" & _
             " " & IIF(lngSystem <= 0, "-Null", lngSystem) & _
             "," & IIF(lngModule <= 0, "-Null", lngModule) & _
             ",'" & strFunction & "'" & _
             ",'" & strDBObject & "'" & _
             ",'" & strOwner & "'" & _
             ",'" & strOperation & "' " & vbCrLf & _
             "From Dual " & vbCrLf & _
             "Where Not Exists(" & vbCrLf & _
             "            Select 1 From zlProgPrivs " & vbCrLf & _
             "            Where 系统 " & IIF(lngSystem <= 0, "Is Null", "= " & lngSystem) & vbCrLf & _
             "              And 序号 " & IIF(lngModule <= 0, "Is Null", "= " & lngModule) & vbCrLf & _
             "              And 功能 = '基本' " & vbCrLf & _
             "              And Upper(对象) = '" & strDBObject & "'" & vbCrLf & _
             "              And Upper(所有者) = '" & strOwner & "'" & vbCrLf & _
             "              And Upper(权限) = '" & strOperation & "'" & vbCrLf & _
             "            )"
    GetInsertProgPrivs = strSQL
    Exit Function
    
hErr:
    Call ErrCenter
End Function

Public Function GetOLEDBConnect(ByVal cnSource As ADODB.Connection _
    , ByVal colCache As Collection _
    , ByVal objRegister As Object) As ADODB.Connection
'功能：获取缓存集合对象中OLEDB连接对象
'参数：
'  cnSource：需要查找的连接对象
'  colCache：集合对象
'  objRegister：注册部件对象

    Dim i As Integer
    Dim strServer As String, strUser As String, strPass As String
    Dim strUserCheck As String, strServerCheck As String, strPassCheck As String

    If objRegister Is Nothing Then Exit Function
    If cnSource Is Nothing Then Exit Function
    If colCache Is Nothing Then Exit Function
    If colCache.count <= 0 Then Exit Function
    
    '检查是否缓存
    '获取连接对象中的数据库服务名、用户名信息
    Call objRegister.GetConnectionInfo(cnSource, strServerCheck, strUserCheck, strPassCheck)
    For i = 1 To colCache.count
        If Not colCache(i) Is Nothing Then
            '获取连接对象中的数据库服务名、用户名信息
            Call objRegister.GetConnectionInfo(colCache(i), strServer, strUser, strPass)
            If UCase(Trim(strUserCheck)) = UCase(Trim(strUser)) _
                And UCase(Trim(strServerCheck)) = UCase(Trim(strServer)) Then
                '找到
                Set GetOLEDBConnect = colCache(i)
                Exit Function
            End If
        End If
    Next
    
End Function

Private Function HaveAdditionTable(ByVal objRPT As Report, ByVal objRPTItem As RPTItem) As Boolean
'--------------------------------------------------------------------------------
'功能：检查RPTItem有无附加表格
'参数：
'  lngRPTItemID：RPTItem对象的ID
'返回：True有附加表格；False无附加表格
'--------------------------------------------------------------------------------

    Dim tmpItem As RPTItem
    
    HaveAdditionTable = False
    For Each tmpItem In objRPT.Items
        If objRPTItem.名称 = tmpItem.参照 And tmpItem.性质 = Val("1-附加表格") Then
            HaveAdditionTable = True
            Exit For
        End If
    Next
End Function

Public Function IsBottomAdditionGrid(ByVal objItems As RPTItems, ByVal objItem As RPTItem) As Boolean
'--------------------------------------------------------------------------------
'功能：判断参数的对象是否为最底部的附加报表
'参数：
'  objItems：Report.Items集合对象
'  objItem：Item对象
'--------------------------------------------------------------------------------

    Dim objTmp As RPTItem
    
    For Each objTmp In objItems
        '通过Y坐标判断
        If objTmp.Y > objItem.Y And objTmp.类型 = Val("4-自由表格") And objItem.参照 <> "" And objTmp.参照 = objItem.参照 Then
            IsBottomAdditionGrid = True
            Exit For
        End If
    Next
End Function

Private Function GridAtCard(ByVal objReport As Report, ByVal lngID As Long) As Boolean
'功能：判断表格对象是否在卡片对象中
'参数：
'  objReport：Report对象
'  lngID：表格对象的ID

    If objReport.Items("_" & lngID).父ID <= 0 Then Exit Function
    GridAtCard = objReport.Items("_" & objReport.Items("_" & lngID).父ID).类型 = Val("14-卡片")
End Function


