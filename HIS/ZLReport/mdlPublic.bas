Attribute VB_Name = "mdlPublic"
Option Explicit
'**************************
'       OEM����
'
'ҽҵ  D2BDD2B5
'����  CDD0C6D5
'**************************

'------------------------------------------------------------------------------------
Public gblnSilentMode As Boolean            '��Ĭ�������ģʽ
Public gstrErrorContent As String           '��������
Public gcnOLEDB As ADODB.Connection
Public gcnOracle As ADODB.Connection
Public gobjRegister As Object               'ע����Ȩ����zlRegister
Public gclsCNs As RPTDBCNs                  '���������õ������������Ӷ���
Public grsConnect As ADODB.Recordset        '���������õ������������Ӽ�¼
Public gblnManagementTool As Boolean        '�����ߵ���
Public gfrmDBConnect As Object              '�������ӹ��������

Public grsObject As ADODB.Recordset '��ǰ�û�������SelectȨ�޵Ķ���(�����򵼻򷢲�)
Public gblnAutoConnect As Boolean   '�Ƿ�������Զ��������ݿ�
Public gblnExeSQLTest   As Boolean  'SQLTest״̬
Public gcolOLEDBConnect As Collection       'OLEDB���ݿ����Ӷ��󻺴�
'------------------------------------------------------------------------------------
Public Type CustomPar
    ���� As String
    ֵ�б� As String
    ����SQL As String
    ��ϸSQL As String
    �����ֶ� As String
    ��ϸ�ֶ� As String
    ���� As String
    ��ʽ As Byte
End Type
Public Type ReportData
    DataName As String
    DataSet As ADODB.Recordset
End Type

'1:UBound(Array())=-1��2:Ubound(û��������ֵ)=-1��3:ֱ��UBound()=�±�Խ��
Public gblnError As Boolean
Public garrPars() As Variant '������������,����DLL���ⲿ�ӿ�
Public garrBill As Variant '��ӡʱ��Ʊ�ݺ�����

Public glngSys As Long '�����������ñ���ִ�л���ƽӿڵ�ϵͳ��

'���ڱ�����󻺴�
Public gobjReport As Report '�����������,����DLL��������
Public grsReport As ADODB.Recordset '����򿪵ı���,���ڻ������,����zlReports��ʱ��Ҫ���
Public gdatModiTime As Date '����򿪵ı��������޸�ʱ��,���ڼ��ӱ仯
Public gcolPrivs As New Collection
Public gcolRptPriv As Collection
Public gcolUserInfo As Collection

Public gblnSingleTask As Boolean '�Ƿ�౨���ڵ������д�ӡ

Public glngGroup As Long '��ǰ��Ϊ������ʱ����ID,��ʱgobjReport=Nothing
Public gfrmMain As Object
Public gobjFile As New FileSystemObject

Public lngTXTProc As Long '����Ĭ�ϵ���Ϣ�����ĵ�ַ
Public objClip As RPTItems '���������
Public gstrColor As String
Public glngUserID As Long

Private mstrBigTable As String   '���
Private mstrMiddleTable As String '�б�
Private mstrMiddleTableRows As String


Public Const GSTR_SBC = "���������������������������������������������£ãģţƣǣȣɣʣˣ̣ͣΣϣУѣңӣԣգ֣ףأ٣ڣ����������������������������������������������"
Public Const GSTR_DBC = "(+-*/=<>)!:1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZabcedfghijklmnopqrstuvwxyz;,.?|%#"

Private ArrayCompare(20) As Integer  '���ܴ�����
'------------------------------------------------------------------------------------

'������־������ر���
Private mlngErrNum As Long, mstrErrInfo As String, mbytErrType As Byte
Private mstrRecentSQL As String  '���ִ�е�SQL���

'SQLLog����
Private msngTime As Single
Private mobjLogText As TextStream

Public gblnRunLog As Boolean '�Ƿ��¼ʹ����־
Public gblnErrLog As Boolean '�Ƿ��¼���д���
Public gblnReportRunLog As Boolean '�Ƿ��¼����������־
Public gblnReportUse As Boolean '�Ƿ��¼����ʹ�úۼ�

'ȱʡ��Ʊ�ݿ�Ⱥ͸߶�,A4,����(ϵͳ��Twip��Ϊ��λ����)
Public Const Twip_mm = 56.69286 '��λת��ϵ��
'Public Const Twip_mm = 56.6857142857143
Public Const INIT_WIDTH = 11904
Public Const INIT_HEIGHT = 16832

Public gblnOK As Boolean
Public glngOldProc As Long, glngSelProc As Long
Public gstrFind As String
Public gblnModi As Boolean
Public gstrFonts As String

Public gstrDBUser As String '�û���
Public gstrUserName As String '�û�����
Public gstrUserNO As String '�û����
Public gstrLoginUser As String '��¼�û���
Public gstrLoginUserName As String '��¼�û�����
Public gstrProductName As String '��Ʒ����
Public gcnOracleConn As String '��¼�ϴ������ַ���
Public gstrComputerName As String '��¼��������

'API����
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
Public Const WM_CONTEXTMENU = &H7B ' ���һ��ı���ʱ������������Ϣ
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

Public Const DC_PAPERNAMES = 16 'ֽ������(ÿ64�ַ�Ϊһ��,��Chr(0)����)
Public Const DC_PAPERS = 2 'ֽ�ű��(Array or Word)
Public Const DC_BINNAMES = 12 '��ֽ��ʽ(ÿ24�ַ�Ϊһ��,��Chr(0)����)
Public Const DC_BINS = 6 '��ֽ���(Array or Word)

Public Const REG_SZ = 1
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const LVSCW_AUTOSIZE = -1
Public Const LVSCW_AUTOSIZE_USEHEADER = -2
Public Const LVM_SETCOLUMNWIDTH = &H101E
Public Const SWP_NOMOVE = &H2

'��ӡֽ�ų���(256=�Զ���)
Public Const PageSize1 = "�ż㣬 8 1/2��11 Ӣ��"
Public Const PageSize2 = "+A611 С���ż㣬 8 1/2��11 Ӣ��"
Public Const PageSize3 = "С�ͱ��� 11��17 Ӣ��"
Public Const PageSize4 = "�����ʣ� 17��11 Ӣ��"
Public Const PageSize5 = "�����ļ��� 8 1/2��14 Ӣ��"
Public Const PageSize6 = "�����飬5 1/2��8 1/2 Ӣ��"
Public Const PageSize7 = "�����ļ���7 1/2��10 1/2 Ӣ��"
Public Const PageSize8 = "A3, 297��420 ����"
Public Const PageSize9 = "A4, 210��297 ����"
Public Const PageSize10 = "A4С�ţ� 210��297 ����"
Public Const PageSize11 = "A5, 148��210 ����"
Public Const PageSize12 = "B4, 250��354 ����"
Public Const PageSize13 = "B5, 182��257 ����"
Public Const PageSize14 = "�Կ����� 8 1/2��13 Ӣ��"
Public Const PageSize15 = "�Ŀ����� 215��275 ����"
Public Const PageSize16 = "10��14 Ӣ��"
Public Const PageSize17 = "11��17 Ӣ��"
Public Const PageSize18 = "������8 1/2��11 Ӣ��"
Public Const PageSize19 = "#9 �ŷ⣬ 3 7/8��8 7/8 Ӣ��"
Public Const PageSize20 = "#10 �ŷ⣬ 4 1/8��9 1/2 Ӣ��"
Public Const PageSize21 = "#11 �ŷ⣬ 4 1/2��10 3/8 Ӣ��"
Public Const PageSize22 = "#12 �ŷ⣬ 4 1/2��11 Ӣ��"
Public Const PageSize23 = "#14 �ŷ⣬ 5��11 1/2 Ӣ��"
Public Const PageSize24 = "C �ߴ繤����"
Public Const PageSize25 = "D �ߴ繤����"
Public Const PageSize26 = "E �ߴ繤����"
Public Const PageSize27 = "DL ���ŷ⣬ 110��220 ����"
Public Const PageSize28 = "C5 ���ŷ⣬ 162��229 ����"
Public Const PageSize29 = "C3 ���ŷ⣬ 324��458 ����"
Public Const PageSize30 = "C4 ���ŷ⣬ 229��324 ����"
Public Const PageSize31 = "C6 ���ŷ⣬ 114��162 ����"
Public Const PageSize32 = "C65 ���ŷ⣬114��229 ����"
Public Const PageSize33 = "B4 ���ŷ⣬ 250��353 ����"
Public Const PageSize34 = "B5 ���ŷ⣬176��250 ����"
Public Const PageSize35 = "B6 ���ŷ⣬ 176��125 ����"
Public Const PageSize36 = "�ŷ⣬ 110��230 ����"
Public Const PageSize37 = "�ŷ������ 3 7/8��7 1/2 Ӣ��"
Public Const PageSize38 = "�ŷ⣬ 3 5/8��6 1/2 Ӣ��"
Public Const PageSize39 = "U.S. ��׼��д���� 14 7/8��11 Ӣ��"
Public Const PageSize40 = "�¹���׼��д���� 8 1/2��12 Ӣ��"
Public Const PageSize41 = "�¹����ɸ�д���� 8 1/2��13 Ӣ��"

'�Զ���ֽ��
Public Const PageCustom1 = "���״�ӡֽ(���ַ�)��241��280 ����"
Public Const PageCustom2 = "���״�ӡֽ(���ȷ�)��241��140 ����"
Public Const PageCustom3 = "���״�ӡֽ(���ȷ�)��241��94 ����"

'����TAB���ĺ���
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

'Window�汾����
Type OSVERSIONINFO 'for GetVersionEx API call
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

'ֽ�Ŵ�ӡ�߽����================================================================
Public Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
'��ͬ��ӡ���Ĵ�ӡ��Ԫ���Ȳ�ͬ
Public Const PHYSICALWIDTH = 110   'Physical Width in device units
Public Const PHYSICALHEIGHT = 111  'Physical Height in device units
Public Const PHYSICALOFFSETX = 112 'Physical Printable Area x margin
Public Const PHYSICALOFFSETY = 113 'Physical Printable Area y margin
Public Const LOGPIXELSX = 88 'Number of pixels per logical inch along the screen width
Public Const LOGPIXELSY = 90
Public Const SCALINGFACTORX = 114  'Scaling factor x
Public Const SCALINGFACTORY = 115  'Scaling factor y
Public Const DRIVERVERSION = 0     'Device driver version

'WinNT�Զ���ֽ�ſ���================================================================
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

'���½�Ϊ�µĴ�ӡ��ʽʹ��-----------------------------------------------------------
'ע����dmFields��Long��,as Long��β����&��
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

'Ŀ¼ѡ��Ի�����=================================================================
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

'����м�����========================================================================
Public Oldwinproc As Long
Public Const WM_COMMAND = &H111
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MOUSEWHEEL = &H20A
    
Public Function FlexScroll(ByVal hwnd As Long, ByVal wMsg As Long, _
                           ByVal wParam As Long, ByVal lParam As Long) As Long
'֧�ֹ��ֵĹ���
    Select Case wMsg
    Case WM_MOUSEWHEEL
        Select Case wParam
        Case -7864320  '���¹�
            SendKeys "{PGDN}"
        Case 7864320   '���Ϲ�
            SendKeys "{PGUP}"
        End Select

    End Select
    FlexScroll = CallWindowProc(Oldwinproc, hwnd, wMsg, wParam, lParam)
End Function
'===================================================================================
Public Function Lpad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ָ���������ƿո�
    '--�����:
    '--������:
    '--��  ��:�����ִ�
    '-----------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = LenB(StrConv(strCode, vbFromUnicode))
    If lngTmp < lngLen Then
        strTmp = String(lngLen - lngTmp, strChar) & strTmp
    ElseIf lngTmp > lngLen Then  '���ڳ���ʱ,�Զ��ض�
        strTmp = strCode
    End If
    Lpad = Replace(strTmp, Chr(0), strChar)
End Function
Public Function RPAD(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ָ���������ƿո�
    '--�����:
    '--������:
    '--��  ��:�����ִ�
    '-----------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = LenB(StrConv(strCode, vbFromUnicode))
    If lngTmp < lngLen Then
        strTmp = strTmp & String(lngLen - lngTmp, strChar)
    Else
        '��Ҫ�пո������
        strTmp = strCode
    End If
    'ȡ��������ַ�
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
'���ܣ��Ƿ�WindowNT����ϵͳ
    Const dwMaskNT = &H2&
    IsWindowsNT = (GetWinPlatform() And dwMaskNT)
End Function

Public Function IsWindows95() As Boolean
'���ܣ��Ƿ�Window95����ϵͳ
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
'˵����
'   Code=Hook Code(HC_ACTION��HC_NOREMOVE)
'   wParam=Virtual-Key Code
'   lParam=0-15λ(�������ظ�����)
'          16-23λ(OEM Scan Code)
'          24λ(�Ƿ���չ��,��Fx,С���̼�)
'          25-28λ(����)
'          29(ALT�Ƿ���)
'          30(������Ϣ֮ǰ���Ƿ���)
'          31(0-���ڰ���,1-�����ɿ�)
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
'���ܣ�ע�����������ļ�
    Dim strSys As String * 255
    
    GetSystemDirectory strSys, 255
    
    RegSetValue HKEY_CLASSES_ROOT, ".zlr", REG_SZ, "zlReport", 7
    RegSetValue HKEY_CLASSES_ROOT, "zlReport", REG_SZ, "�Զ��屨���ļ�", 7
    RegSetValue HKEY_CLASSES_ROOT, "zlReport\DefaultIcon", REG_SZ, Left(strSys, InStr(strSys, Chr(0)) - 1) & "\zl9Report.dll,0", 24
    RegSetValue HKEY_CLASSES_ROOT, "zlReport\Shell", REG_SZ, "Read", 4
    RegSetValue HKEY_CLASSES_ROOT, "zlReport\Shell\Read", REG_SZ, "���Զ��屨���ļ�(&1)", 12
    RegSetValue HKEY_CLASSES_ROOT, "zlReport\Shell\Read\Command", REG_SZ, "NotePad.exe ""%1""", 22
End Sub

Public Function GetPaperName(ByVal intSize As Integer, Optional ByVal lngW As Long, Optional ByVal lngH As Long) As String
'���ܣ� ���ݵ�ǰ��ӡ�������ã���ȡֽ������
'������ lngW,lngH=�Զ���ֽ�ŵĿ��(Twip)
'���أ� ֽ������
    If intSize = 256 Then
        If CInt(lngW / Twip_mm) = 241 And CInt(lngH / Twip_mm) = 280 Then
            GetPaperName = PageCustom1
        ElseIf CInt(lngW / Twip_mm) = 241 And CInt(lngH / Twip_mm) = 140 Then
            GetPaperName = PageCustom2
        ElseIf CInt(lngW / Twip_mm) = 241 And CInt(lngH / Twip_mm) = 94 Then
            GetPaperName = PageCustom3
        Else
            GetPaperName = "�û��Զ��� ..."
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
        GetPaperName = "���ɲ��ֽ�� ..."
    End If
End Function

Public Sub SetComboBoxHeight(cbo As ComboBox, lngH As Long)
'���ܣ����������б��ߴ�,������Ϊ��λ
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
'����:�ж������Ƿ��д�ֱ������
'˵��:���и߱���һ��
    If msh.RowHeight(0) * msh.Rows >= msh.Height Then
        ScrollExist = True
    Else
        ScrollExist = False
    End If
End Function

Private Function mGetInvalidTable() As String
'���ܣ��õ������ʹ�õ�SQL����в��ܷ��ʵı����ͼ
    Dim varTables As Variant
    Dim strTable As String, lngCount As Long
    Dim strInvalidTable As String
    
    varTables = Split(SQLObject(mstrRecentSQL), ",")
    
    On Error Resume Next
    For lngCount = LBound(varTables) To LBound(varTables)
        strTable = varTables(lngCount)
        
        '���Ըö����Ƿ����
        gcnOracle.Execute "select 1 from " & strTable & " where rownum<1"
        If Err <> 0 Then
            Err.Clear
            strInvalidTable = strInvalidTable & "," & strTable
        End If
    Next
    
    If strInvalidTable <> "" Then
        'ȥ����һ������
        mGetInvalidTable = Mid(strInvalidTable, 2)
    End If
End Function

Public Function ErrCenter() As Byte
'------------------------------------------------
'���ܣ� �����������������
'������
'���أ� cancel      ���� 0
'       resume      ���� 1
'------------------------------------------------
    Dim strNote As String, strTemp As String
    Dim bytReturnType As Byte
    Dim blnExeSQLTest As Boolean
    Static mstrErrRecentSQL As String
    
    bytReturnType = 1
    If gcnOracle.Errors.count <> 0 Then
        'PL/SQL�洢���̴���
        If gcnOracle.Errors(0).NativeError >= 20000 And gcnOracle.Errors(0).NativeError <= 20200 Then
            '��־����
            mbytErrType = 1
            mlngErrNum = gcnOracle.Errors(0).NativeError
            mstrErrInfo = gcnOracle.Errors(0).Description
            
            strNote = gcnOracle.Errors(0).Description
            MsgBox Split(strNote, "[ZLSOFT]")(1), vbExclamation, App.Title
            Exit Function
        End If
        'ORACLE��������
        '��־����
        mbytErrType = 2
        mlngErrNum = gcnOracle.Errors(0).NativeError
        mstrErrInfo = gcnOracle.Errors(0).Description
        
        Select Case gcnOracle.Errors(0).NativeError
        Case 1
            strNote = "�Ѿ�������ͬ���ݵ����ݣ�Ҫ��Ψһ������[���š����Ƶ�]���ظ�����"
            bytReturnType = 0
        Case 903
            strNote = "�����ƴ���"
        Case 904
            strNote = "�����ƴ���"
        Case 942
            strNote = "�����ͼ�����ڣ��ܿ������㲻�߱�ʹ�øò������ݵ�Ȩ�޻�ò��ֶ���ͬ���ȱʧ��"
            bytReturnType = 0
            
            strTemp = mGetInvalidTable()
            If strTemp <> "" Then
                mstrErrInfo = "��������ϵͳ����Ա������Ȩ���޸�ͬ��ʣ�" & vbCrLf & "������ж�����м�飺" & vbCrLf & vbCrLf & vbTab & strTemp
            Else
                mstrErrInfo = "��������ϵͳ����Ա������Ȩ���޸�ͬ��ʣ�" & vbCrLf & "����SQL���Ϊ��" & vbCrLf & vbCrLf & mstrRecentSQL
            End If
        Case 1000
            strNote = "�򿪵����ݱ�̫�࣬��Ҫʱ��ϵͳ����Ա�޸����ݿ��Open_Cursors���á�"
        Case 1005
            strNote = "������û��������롣"
        Case 1017
            strNote = "������û��������롣"
            bytReturnType = 0
        Case 1031
            strNote = "û���㹻��Ȩ�ޡ�"
            bytReturnType = 0
        Case 1045
            strNote = "û���������ݿ��Ȩ�ޡ�"
            bytReturnType = 0
        Case 1400
            strNote = "���ڸ�������Ҫ��ǿ��и����˿�ֵ����������ʧ�ܡ�"
            bytReturnType = 0
        Case 1401
            strNote = "���ڸ����ֵ�������п����ƣ��������ӻ����ʧ�ܡ�"
            bytReturnType = 0
        Case 1402
            strNote = "���ڸ����ֵ��������ͼ���������ƣ��������ӻ����ʧ�ܡ�"
            bytReturnType = 0
        Case 1403
            strNote = "����δ���������ݣ����º�������ʧ�ܡ�"
        Case 1404
            strNote = "�޸��в�����������ص�����̫��"
        Case 1405
            strNote = "ȡ�õ���ֵΪ�ա�"
        Case 1406
            strNote = "ȡ�õ���ֵ���ж϶������ˡ�"
        Case 1407
            strNote = "���ڸ�������Ҫ��ǿ��и����˿�ֵ�����¸���ʧ�ܡ�"
            bytReturnType = 0
        Case 1408
            strNote = "ָ�������Ѿ�������������"
        Case 1409
            strNote = "���ܽ�����˳�����(NoSort)����Ϊ�����û����"
        Case 1410
            strNote = "�������ID(ROWID)����ID���������ֺ��ַ���ɵ�16���Ƹ�ʽ��"
        Case 1411
            strNote = "��ǰ�в��ܴ洢����64K�����ݡ�"
            bytReturnType = 0
        Case 1412
            strNote = "��ǰ���������Ͳ��ܴ洢�㳤���ַ�����"
            bytReturnType = 0
        Case 1413
            strNote = "�����С��λ��������ʧ�ܡ�"
            bytReturnType = 0
        Case 1415
            strNote = "���ܶ�һ����ǩα��ָ��������[Outer-Join(+)]"
        Case 1416
            strNote = "���ű���ͬʱָ��һ��������[Outer-Join(+)]"
        Case 1417
            strNote = "һ�ű�ֻ��ָ��ָ�򲻳���һ�ű��������[Outer-Join(+)]"
        Case 1418
            strNote = "ָ�������������ڡ�"
        Case 1424
            strNote = "�������Ч�Ļ����ַ�(ͨ�����ֻ����'%'��'_')��"
        Case 1425
            strNote = "�����ַ������ǳ���Ϊ1���ַ���"
        Case 1426
            strNote = "��ֵ���ʽ���������(̫���̫С)��"
        Case 1427
            strNote = "�����Ӳ�ѯ�����˶��С�"
        Case 1428
            strNote = "�����Ĳ�������򳬽硣"
        Case 1429
            strNote = "һ�����������ڸ�ʽ���硣"
        Case 1430
            strNote = "ϣ�����ӵ����Ѿ����ڡ�"
        Case 1431
            strNote = "��Ȩ����(GRANT)�������ڵĲ�һ�¡�"
        Case 1432
            strNote = "ϣ��ɾ���Ĺ���ͬ����Ѿ������ڡ�"
        Case 1433
            strNote = "ϣ��������ͬ����Ѿ����ڡ�"
        Case 1434
            strNote = "ϣ��ɾ����ͬ����Ѿ������ڡ�"
        Case 1435
            strNote = "ָ�����û������ڡ�"
            bytReturnType = 0
        Case 1438
            strNote = "��ֵ������������ľ�ȷ�̶ȡ�"
        Case 1439, 1440, 1441
            strNote = "ֻ�п�ֵ�в����޸��������͡������Ȼ�ߴ��С"
        Case 1536
            strNote = "ĳ��������ռ�Ŀռ�������"
        Case 2290
            strNote = "������Ŀֵ��������ķ�Χ��Υ���˼��Լ�������������ӻ����ʧ�ܡ�"
            bytReturnType = 0
        Case 2291
            strNote = "����δ��д��ر��д��ڵ���Ŀֵ(Υ�������Լ��)���������ӻ����ʧ�ܡ�"
        Case 2292
            strNote = "��Ϊ�ü�¼�Ѿ�ʹ�ã��ʲ���ɾ���˼�¼��"
            bytReturnType = 0
        Case 12203
            strNote = "������������д�����û���������⣬�����������ӡ�"
            bytReturnType = 0
        Case Else
            strTemp = Err.Description
            If InStr(strTemp, "PLS-00201") > 0 And InStr(strTemp, "ZL_") > 0 Then
                Dim lngPos As Long
                
                lngPos = InStr(strTemp, "ZL_")
                strTemp = Mid(strTemp, lngPos)
                strTemp = Mid(strTemp, 1, InStr(strTemp, "'") - 1)
                
                strNote = "���ڷ����������ߵĽ�ɫ������������ӶԹ���"" & strTemp & ""����Ȩ��"
            Else
                strNote = "δ֪���󣬷�����" & gcnOracle.Errors(0).Source
            End If
        End Select
        
    Else
        'VB��׼����
        '��־����
        mbytErrType = 3
        mlngErrNum = Err.Number
        mstrErrInfo = Err.Description
        
        Select Case Err.Number
            Case 3, 3 - 2146828288
                strNote = "δ���ñ�׼���ع���"
            Case 5, 5 - 2146828288
                strNote = "��Ч�Ĺ��̻����"
            Case 6, 6 - 2146828288
                strNote = "�������"
            Case 7, 7 - 2146828288
                strNote = "�ڴ����"
            Case 9, 9 - 2146828288
                strNote = "�±곬��"
            Case 10, 10 - 2146828288
                strNote = "�����ǹ̶��������ʱ����"
            Case 11, 11 - 2146828288
                strNote = "����Ϊ��̫С"
            Case 13, 13 - 2146828288
                strNote = "���Ͳ�ƥ��"
            Case 14, 14 - 2146828288
                strNote = "�����ַ���������"
            Case 16, 16 - 2146828288
                strNote = "���ʽ̫����"
            Case 17, 17 - 2146828288
                strNote = "��֧��Ҫ��Ĳ���"
            Case 18, 18 - 2146828288
                strNote = "�������û��ж�"
            Case 20, 20 - 2146828288
                strNote = "�޴��󷵻�"
            Case 28, 28 - 2146828288
                strNote = "��ջ�ռ����"
            Case 35, 35 - 2146828288
                strNote = "���̻���δ����"
            Case 47, 47 - 2146828288
                strNote = " ̫��Ķ�̬����⣨DLL��Ӧ�ÿͻ�"
            Case 48, 48 - 2146828288
                strNote = " ���ö�̬����⣨DLL������"
            Case 49, 49 - 2146828288
                strNote = " ��̬����⣨DLL��Լ������"
            Case 51, 51 - 2146828288
                strNote = "�ڲ�����"
            Case 52, 52 - 2146828288
                strNote = "������ļ������ļ���"
            Case 53, 53 - 2146828288
                strNote = "�ļ�δ�ҵ�"
            Case 54, 54 - 2146828288
                strNote = "�ļ���ʽ����"
            Case 55, 55 - 2146828288
                strNote = "�ļ��Ѿ���"
            Case 57, 57 - 2146828288
                strNote = "�豸���� / �������"
            Case 58, 58 - 2146828288
                strNote = "�ļ��Ѿ�����"
            Case 59, 59 - 2146828288
                strNote = "����ļ�¼����"
            Case 61, 61 - 2146828288
                strNote = "������"
            Case 62, 62 - 2146828288
                strNote = "���볬���ļ�β"
            Case 63, 63 - 2146828288
                strNote = "����ļ�¼��"
            Case 67, 67 - 2146828288
                strNote = "�ļ�̫��"
            Case 68, 68 - 2146828288
                strNote = "�豸��Ч��֧��"
            Case 70, 70 - 2146828288
                strNote = "�ܾ�����"
            Case 71, 71 - 2146828288
                strNote = "����δ׼����"
            Case 74, 74 - 2146828288
                strNote = "��������Ϊ��ͬ��������"
            Case 75, 75 - 2146828288
                strNote = "·�� / �ļ����ʴ���"
            Case 76, 76 - 2146828288
                strNote = "·��δ�ҵ�"
            Case 91, 91 - 2146828288
                strNote = "�������������Ϊ����(δ�½�ʵ��)"
            Case 92, 92 - 2146828288
                strNote = "ѭ��δ��ʼ��"
            Case 93, 93 - 2146828288
                strNote = "�����ģʽ�ַ���"
            Case 94, 94 - 2146828288
                strNote = "�����ʹ�ÿ�(Null)"
            Case 96, 96 - 2146828288
                strNote = " �����Ѿ�ʹ�õĶ���ʱ�䳬���������õ����Ԫ�غţ����²����ܽ����¼�"
            Case 97, 97 - 2146828288
                strNote = "���ܵ���һ��δ����ʵ�����������"
            Case 98, 98 - 2146828288
                strNote = " ����ʹ��һ��˽�ж�������Ժͷ���?�����ͷ���ֵ"
            Case 321, 321 - 2146828288
                strNote = "������ļ���ʽ"
            Case 322, 322 - 2146828288
                strNote = "���ܴ�����Ҫ����ʱ�ļ�"
            Case 325, 325 - 2146828288
                strNote = "��Դ�ļ��д���ĸ�ʽ"
            Case 380, 380 - 2146828288
                strNote = "���������ֵ"
            Case 381, 381 - 2146828288
                strNote = "�����������������"
            Case 382, 382 - 2146828288
                strNote = "��֧�ֵ�����ʱ����"
            Case 383, 383 - 2146828288
                strNote = "��֧�ֵ�ֻ����������"
            Case 385, 384 - 2146828288
                strNote = "��Ҫ������������"
            Case 387, 387 - 2146828288
                strNote = "�����������"
            Case 393, 393 - 2146828288
                strNote = "��֧�ֵ�����ʱ��ȡ"
            Case 394, 394 - 2146828288
                strNote = "��֧�ֵ�ֻд���Զ�ȡ"
            Case 422, 422 - 2146828288
                strNote = "�����ڵ�����"
            Case 423, 423 - 2146828288
                strNote = "�����ڵ����Ի򷽷�"
            Case 424, 424 - 2146828288
                strNote = "Ҫ��һ������"
            Case 429, 429 - 2146828288
                strNote = "ActiveX���ܴ�������"
            Case 430, 430 - 2146828288
                strNote = "�಻֧�ֵ��Զ���������֧�ֵĽ���"
            Case 432, 432 - 2146828288
                strNote = "���Զ������ڼ�δ�ҵ��ļ�����������"
            Case 438, 438 - 2146828288
                strNote = "����֧�ָ����Ի򷽷�"
            Case 440, 440 - 2146828288
                strNote = "�Զ����������"
            Case 442, 442 - 2146828288
                strNote = "��Զ��������������ᶪʧ����OK����Ի���ȥ����"
            Case 443, 443 - 2146828288
                strNote = "�Զ�������û��ȱʡֵ"
            Case 445, 445 - 2146828288
                strNote = "����֧�����ֲ���"
            Case 446, 446 - 2146828288
                strNote = "����֧����������"
            Case 447, 447 - 2146828288
                strNote = "����֧�ֵ�ǰ��������"
            Case 448, 448 - 2146828288
                strNote = "��������δ�ҵ�"
            Case 449, 449 - 2146828288
                strNote = "�������ǿ�ѡ��"
            Case 450, 450 - 2146828288
                strNote = "����Ĳ������������Է���"
            Case 451, 451 - 2146828288
                strNote = "���Ը�ֵ(Let)���̺Ͷ�ȡ(Get)���̲����ض���"
            Case 452, 452 - 2146828288
                strNote = "��Ч�����"
            Case 453, 453 - 2146828288
                strNote = "ָ����DLL����δ�ҵ�"
            Case 454, 454 - 2146828288
                strNote = "������Դδ�ҵ�"
            Case 455, 455 - 2146828288
                strNote = "������Դ��������"
            Case 457, 457 - 2146828288
                strNote = "�ùؼ�ֵ�Ѿ��뼯�ϵ���һԪ�ؽ��"
            Case 458, 458 - 2146828288
                strNote = "VB��֧�ֵĿɱ��Զ�������"
            Case 459, 459 - 2146828288
                strNote = "������಻֧�ֵ��¼���"
            Case 460, 460 - 2146828288
                strNote = "����ļ������ʽ"
            Case 461, 461 - 2146828288
                strNote = "���������ݳ�Աδ�ҵ�"
            Case 462, 462 - 2146828288
                strNote = "Զ�̷����������ڻ���Ч"
            Case 463, 463 - 2146828288
                strNote = "��û���ڱ���ע��"
            Case 481, 481 - 2146828288
                strNote = "��Ч��ͼƬ��ʽ"
            Case 482, 482 - 2146828288
                strNote = "��ӡ������"
            Case 735, 735 - 2146828288
                strNote = "���ܽ��洢Ϊ��ʱ�ļ�"
            Case 744, 744 - 2146828288
                strNote = "δ�ҵ�����������"
            Case 746, 746 - 2146828288
                strNote = "̫���ĸ���"
            'ADO����
            Case 3001
                strNote = "�������ʹ��󣬻���ֵ������Χ�������ͻ��"
            Case 3021
                strNote = "��¼����(EOF/BOF)�����ߵ�ǰ��¼��ɾ������ǰӦ�ò�����Ҫ��λ��ǰ��¼��"
            Case 3219
                strNote = "�����Ļ���������ǰӦ�ò����������Ǵ�����δ���������񣩡�"
            Case 3246
                strNote = "������ִ���У����ܹر�һ���������"
            Case 3251
                strNote = "��ǰ������֧����һӦ�ò�����"
            Case 3265
                strNote = "ADOû�ҵ�Ӧ�ó���Ҫ��Ķ�Ӧ���ƻ���š�"
            Case 3367
                strNote = "�����Ѿ����ڣ�������ӡ�"
            Case 3420
                strNote = "����δ���á�"
            Case 3421
                strNote = "��ǰ����ʹ���˴������ֵ���͡�"
            Case 3704
                strNote = "����ر�ʱ����ǰ��������ִ�С�"
            Case 3705
                strNote = "������ʱ����ǰ��������ִ�С�"
            Case 3706
                strNote = "ADOû�ҵ�ָ����֧�֡�"
            Case 3707
                strNote = "���ܲ����������ı�һ����¼���Ļ����Դ�����ԡ�"
            Case 3708
                strNote = "Ӧ�ó�����ִ���Ĳ������塣"
            Case 3709
                strNote = "Ӧ�ó���Ҫ��һ���رյ����ö������Ч���������"
            Case Else
                strNote = "�����ڽ���δ֪����"
        End Select
        bytReturnType = 0
    End If
    
    
    If gblnAutoConnect Then '�Ƿ�ʹ������Ͽ��Զ����ӹ���
        Dim blnConnect As Boolean
        Dim blnNumConnect As Boolean '�������Ƿ���������
        Dim blnStatus As Boolean '�Ƿ�����������������������
        'ͨ�����˴�����Ϣ,����Ƿ�Ϊ�������������Ĵ���mbytErrType=2 Oracle�ṩ�Ĵ�����Ϣ mbytErrType=3 VB�ṩ�Ĵ�����Ϣ
        If mbytErrType = 3 Then
            If mlngErrNum = -2147467259 Or mlngErrNum = -2147217900 Or mlngErrNum = 3709 Then
                '���VB���������Ϣ
                If CheckErrConnectInfo(mlngErrNum, strNote, mstrErrInfo, 1) Then
                    '�ж���ͬ����,���2����������������ʾ��
                    If mstrErrRecentSQL = mstrRecentSQL And mstrRecentSQL <> "" Then
                        mlngConnectCount = mlngConnectCount + 1
                        If mlngConnectCount > 2 Then
                            blnNumConnect = False  '����������ʾ
                            mlngConnectCount = 0 '��ԭ������
                        Else
                            blnNumConnect = True
                        End If
                    Else
                        mstrErrRecentSQL = mstrRecentSQL
                        mlngConnectCount = 1
                        blnNumConnect = True
                    End If
                Else
                    blnConnect = False '����������ʾ
                End If
            End If
        Else
            '�����12543 TNS: �޷�����Ŀ������,1012-û�е�¼��0028-�Ự����ֹ
            If mlngErrNum = -2147467259 Or mlngErrNum = -2147217900 Or mlngErrNum = 0 Or mlngErrNum = 12543 Or mlngErrNum = 2399 Or mlngErrNum = 2396 Or mlngErrNum = 1012 Or mlngErrNum = 28 Then
                '���ORACLE���������Ϣ
                If CheckErrConnectInfo(mlngErrNum, strNote, mstrErrInfo, 2) Then
                    '�ж���ͬ����,���2����������������ʾ��
                    If mstrErrRecentSQL = mstrRecentSQL And mstrRecentSQL <> "" Then
                        mlngConnectCount = mlngConnectCount + 1
                        If mlngConnectCount > 2 Then
                            blnNumConnect = False  '����������ʾ
                            mlngConnectCount = 0 '��ԭ������
                        Else
                            blnNumConnect = True
                        End If
                    Else
                        mstrErrRecentSQL = mstrRecentSQL
                        mlngConnectCount = 1
                        blnNumConnect = True
                    End If
                Else
                    blnConnect = False '����������ʾ
                End If
            End If
        End If
        
        '�Զ���������һ��,����Ƿ����Զ���������
        If blnNumConnect Then '��ORACLE�����Ѿ��Ͽ�
            blnExeSQLTest = gblnExeSQLTest
            gblnExeSQLTest = True
            If CheckAdoConnction(blnStatus) Then
                If blnStatus Then
                   blnConnect = False '����������ʾ
                Else
                   blnConnect = True '��ʾ����
                End If
            Else
                '��ORACLE�������ӳɹ�,����Ҫ��ʾ��ֱ�ӷ�������ִ�С�
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
    '�������
    Err.Clear
End Function

Public Sub SaveErrLog()
'���ܣ����ղŵĴ�����Ϣд�����ݿ������־
    Dim strSQL As String
    
    If mlngErrNum <> 0 And mbytErrType <> 0 And gblnErrLog Then
        On Local Error Resume Next
        If gstrComputerName = "" Then Exit Sub
        strSQL = "Zl_Zlerrorlog_Insert('" & gstrComputerName & "'," & mbytErrType & "," & mlngErrNum & "," & AdjustStr(mstrErrInfo) & ")"
        Call ExecuteProcedure(strSQL, "���������־")
        mlngErrNum = 0: mstrErrInfo = "": mbytErrType = 0
    End If
End Sub

Public Function ComputerName() As String
    '����:��ȡ�������
    Dim strComputerName As String * 256
    Err = 0
    On Error Resume Next
    
    Call GetComputerName(strComputerName, 255)
    ComputerName = Trim(Replace(strComputerName, Chr(0), ""))
End Function

Public Sub ShowPercent(sngPercent As Single, objPanel As Object)
'����:��״̬���ϸ��ݰٷֱ���ʾ��ǰ�������(��)
    Dim intAll As Integer
    intAll = objPanel.Width / frmAbout.TextWidth("��") - 4
    objPanel.Text = Format(sngPercent, "0% ") & String(intAll * sngPercent, "��")
End Sub

Public Sub SelAll(objTxt As Control)
'���ܣ����ı���ĵ��ı�ѡ��
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
'���ܣ���鹤�������ʵ�����Ƿ���ָ�����Ƴ�����
    If LenB(StrConv(txt.Text, vbFromUnicode)) > intLen Then
        MsgBox "[" & strInfo & "]�ĳ��Ȳ��ܴ��ڣ�" & intLen & "���ַ���" & intLen \ 2 & "�����֣�", vbInformation, App.Title
        txt.SetFocus: Exit Function
    End If
    CheckLen = True
End Function

Public Function TLen(Str As String) As Long
'���ܣ������ַ�������ʵ����
    TLen = LenB(StrConv(Str, vbFromUnicode))
End Function

Public Function CheckExist(strTable As String, strField As String, strValue As String, Optional lngID As Long) As Boolean
'���ܣ�����strTable���ֶ�strField��ֵstrValue�Ƿ��ظ�.
'˵������Ҫ��zlReports��zlRPTGroups�ͱ�Ų���
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
    Call OpenRecord(rsTmp, strSQL, "mdlPublic_GetNextID") '��̬SQL
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
    Call OpenRecord(rsTmp, strSQL, "mdlPublic_GetCurrID") '��̬SQL
    GetCurrID = rsTmp!id
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetSysNO() As String
'���ܣ����ص�ǰϵͳ�����߶�Ӧϵͳ���
'˵����ͬһ�������п��ܴ��ڶ��ϵͳ(���)
'���أ��ɹ�:"1,2,3",ʧ��="0"
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    
    On Error GoTo errH
    
    GetSysNO = "0"
    strSQL = "Select ��� From zlSystems Where ������=User"
    Call OpenRecord(rsTmp, strSQL, "mdlPublic_GetSysNO")
    If rsTmp.RecordCount > 0 Then
        GetSysNO = ""
        For i = 1 To rsTmp.RecordCount
            GetSysNO = GetSysNO & "," & rsTmp!���
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
'���ܣ�����ָ������(��)������λ��(����̨�˵���ģ��)
'˵����һ��������ܷ��������λ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim strPath1 As String, strPath2 As String
    
    On Error GoTo errH
        
    If blnGroup Then
        strSQL = "Select 1 as ��־,D.���� as λ��" & _
            " From zlRPTGroups A,zlPrograms B,zlMenus C,zlMenus D" & _
            " Where Nvl(A.ϵͳ,0)=Nvl(B.ϵͳ,0) And A.����ID=B.���" & _
            " And Nvl(B.ϵͳ,0)=Nvl(C.ϵͳ,0) And B.���=C.ģ��" & _
            " And C.���='ȱʡ' And Upper(B.����)=Upper('zl9Report')" & _
            " And C.�ϼ�ID=D.ID And A.ID=[1]"
    Else
        strSQL = "Select 1 as ��־,D.���� as λ��" & _
            " From zlReports A,zlPrograms B,zlMenus C,zlMenus D" & _
            " Where Nvl(A.ϵͳ,0)=Nvl(B.ϵͳ,0) And A.����ID=B.���" & _
            " And Nvl(B.ϵͳ,0)=Nvl(C.ϵͳ,0) And B.���=C.ģ��" & _
            " And C.���='ȱʡ' And Upper(B.����)=Upper('zl9Report')" & _
            " And C.�ϼ�ID=D.ID And A.ID=[1]"
        strSQL = strSQL & " Union ALL " & _
            " Select 2 as ��־,B.���� as λ��" & _
            " From zlReports A,zlPrograms B" & _
            " Where Nvl(A.ϵͳ,0)=Nvl(B.ϵͳ,0) And A.����ID=B.���" & _
            " And Upper(B.����)<>Upper('zl9Report') And A.ID=[1]"
        strSQL = strSQL & " Union ALL " & _
            " Select 2 as ��־,B.���� as λ��" & _
            " From zlRPTPuts A,zlPrograms B" & _
            " Where A.ϵͳ=B.ϵͳ And A.����ID=B.���" & _
            " And Upper(B.����)<>Upper('zl9Report') And A.����ID=[1]"
    End If
    Set rsTmp = OpenSQLRecord(strSQL, "GetMenuPath", lngRPTID)
    For i = 1 To rsTmp.RecordCount
        If rsTmp!��־ = 1 Then
            strPath1 = strPath1 & "," & rsTmp!λ��
        ElseIf rsTmp!��־ = 2 Then
            strPath2 = strPath2 & "," & rsTmp!λ��
        End If
        rsTmp.MoveNext
    Next
    If strPath1 <> "" Then strPath1 = "����̨(" & Mid(strPath1, 2) & ")"
    If strPath2 <> "" Then strPath2 = "ģ��(" & Mid(strPath2, 2) & ")"
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
'���ܣ������ݿ��ж�ȡָ�������������
'������lngRPTID=����ID,intMaxID=��ƽ��洦������ؼ�����,��ȡ�����иı�
'      blnOnlyData=ֻ��ȡ��������Դ
'���أ�intMaxID=��ǰ���õ����ؼ�����,ReadReport=�������
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
    Dim lng��ID As Long
    
    Dim strSQL As String, i As Integer, j As Integer
    Dim intCopyID As Integer, strReport As String
    
    Dim tmpReport As Report, tmpData As RPTData, tmpPar As RPTPar
    Dim tmpItem As RPTItem, tmpRelation As RPTRelation
    
    If gstrFonts = "" Then gstrFonts = GetScreenFonts
    
    On Error GoTo errH
        
    If ReportReaded(lngRPTID) Then
        Set rsReport = grsReport '���û���
    Else
        strSQL = "Select ID,���,����,˵��,����,��ӡ��,��ֽ,Ʊ��,ϵͳ,����ID,����,�޸�ʱ��,����ʱ��,��ӡ��ʽ,��ֹ��ʼʱ��,��ֹ����ʱ�� " & vbCr & _
                 "From zlReports Where ID=[1]"
        Set rsReport = OpenSQLRecord(strSQL, "ReadReport", lngRPTID)
        If Not rsReport.EOF Then '���洦��
            Set grsReport = New ADODB.Recordset
            Set grsReport = rsReport
            gdatModiTime = grsReport!�޸�ʱ��
        End If
    End If
    If Not rsReport.EOF Then
        strReport = GetFieldNames(rsReport)
        
        Set tmpReport = New Report
        tmpReport.ϵͳ = Nvl(rsReport!ϵͳ, 0)
        tmpReport.��� = rsReport!���
        tmpReport.���� = rsReport!����
        tmpReport.˵�� = Nvl(rsReport!˵��)
        tmpReport.��ֽ = Nvl(rsReport!��ֽ, 15) 'ȱʡΪ�Զ�ѡ��
        tmpReport.��ӡ�� = Nvl(rsReport!��ӡ��)
        tmpReport.Ʊ�� = Nvl(rsReport!Ʊ��, 0) = 1
        tmpReport.��ӡ��ʽ = Nvl(rsReport!��ӡ��ʽ, 0)
        tmpReport.�޸�ʱ�� = rsReport!�޸�ʱ��
        tmpReport.��ֹ��ʼʱ�� = Nvl(rsReport!��ֹ��ʼʱ��, 0)
        tmpReport.��ֹ����ʱ�� = Nvl(rsReport!��ֹ����ʱ��, 0)
        
        '����Դ
        strSQL = "Select ID,�������ӱ��,����ID,����,�ֶ�,����,����,˵�� From zlRPTDatas Where ����ID=[1] Order by ����"
        Set rsData = OpenSQLRecord(strSQL, "ReadReport", lngRPTID)
        If Not rsData.EOF Then
            '����ԴSQL
            strSQL = "Select A.ԴID,A.�к�,A.���� From zlRPTSQLs A,zlRPTDatas B Where A.ԴID=B.ID And B.����ID=[1] Order by A.ԴID,A.�к�"
            Set rsSQL = OpenSQLRecord(strSQL, "ReadReport", lngRPTID)
            
            '����Դ����
            strSQL = "Select A.ԴID,A.����,A.���,A.����,A.����,A.ȱʡֵ,A.��ʽ,A.ֵ�б�,A.����SQL,A.��ϸSQL,A.�����ֶ�,A.��ϸ�ֶ�,A.����,A.����" & _
                    " From zlRPTPars A,zlRPTDatas B Where A.ԴID=B.ID And B.����ID=[1] Order by A.ԴID,A.���,A.����,A.����"
            Set rsPar = OpenSQLRecord(strSQL, "ReadReport", lngRPTID)
        End If
        For i = 1 To rsData.RecordCount
            Set tmpData = New RPTData
            tmpData.�������ӱ�� = Nvl(rsData!�������ӱ��, 0)
            tmpData.���� = rsData!����
            tmpData.���� = rsData!����
            tmpData.�ֶ� = rsData!�ֶ�
            tmpData.���� = Nvl(rsData!����)
            tmpData.˵�� = Nvl(rsData!˵��)
                        
            'SQL
            tmpData.SQL = ""
            rsSQL.Filter = "ԴID=" & rsData!id
            For j = 1 To rsSQL.RecordCount
                tmpData.SQL = tmpData.SQL & vbCrLf & Nvl(rsSQL!����)
                rsSQL.MoveNext
            Next
            tmpData.SQL = Mid(tmpData.SQL, 3)
            
            '����
            rsPar.Filter = "ԴID=" & rsData!id
            For j = 1 To rsPar.RecordCount
                Set tmpPar = New RPTPar
                tmpPar.���� = Nvl(rsPar!����)
                tmpPar.��� = Nvl(rsPar!���, 0)
                tmpPar.���� = Nvl(rsPar!����)
                tmpPar.���� = Nvl(rsPar!����, 0)
                tmpPar.ȱʡֵ = Nvl(rsPar!ȱʡֵ)
                tmpPar.��ʽ = Nvl(rsPar!��ʽ, 0)
                
                tmpPar.ֵ�б� = Nvl(rsPar!ֵ�б�)
                tmpPar.����SQL = Nvl(rsPar!����SQL)
                tmpPar.��ϸSQL = Nvl(rsPar!��ϸSQL)
                tmpPar.�����ֶ� = Nvl(rsPar!�����ֶ�)
                tmpPar.��ϸ�ֶ� = Nvl(rsPar!��ϸ�ֶ�)
                tmpPar.���� = Nvl(rsPar!����)
                tmpPar.�Ƿ����� = IIF(Nvl(rsPar!����, 0) = 1, True, False)
                
                '�������Բ������Ϊ�ؼ��ּ��뼯��
                tmpData.Pars.Add tmpPar.����, tmpPar.���, tmpPar.����, tmpPar.����, tmpPar.ȱʡֵ, tmpPar.��ʽ, tmpPar.ֵ�б� _
                    , tmpPar.����SQL, tmpPar.��ϸSQL, tmpPar.�����ֶ�, tmpPar.��ϸ�ֶ�, tmpPar.����, "_" & tmpPar.���, _
                    , tmpPar.�Ƿ�����
                
                rsPar.MoveNext
            Next
            
            '������������Դ������Ϊ�ؼ��ּ��뼯��
            tmpReport.Datas.Add tmpData.����, tmpData.�������ӱ��, tmpData.SQL, tmpData.�ֶ�, tmpData.����, tmpData.���� _
                , tmpData.˵��, tmpData.Pars, "_" & tmpData.����
            
            rsData.MoveNext
        Next
        
        If blnOnlyData = False Then
            '�����ʽ
            strSQL = "Select ����ID,���,˵��,W,H,ֽ��,ֽ��,��ֽ̬��,ͼ�� From zlRPTFmts Where ����ID=[1] Order by ���"
            Set rsFormat = OpenSQLRecord(strSQL, "ReadReport", lngRPTID)
            For i = 1 To rsFormat.RecordCount
                If IsNull(rsFormat!ֽ��) And IsNull(rsFormat!W) And IsNull(rsFormat!H) _
                    And InStr(strReport, ",ֽ��,") > 0 And InStr(strReport, ",W,") > 0 Then
                    '���ݿ��ǣ�ͳһΪ����ͳһ����
                    tmpReport.Fmts.Add rsFormat!���, rsFormat!˵��, Nvl(rsReport!W, INIT_WIDTH), Nvl(rsReport!H, INIT_HEIGHT), _
                        Nvl(rsReport!ֽ��, 9), Nvl(rsReport!ֽ��, 1), Nvl(rsReport!��ֽ̬��, 0) = 1, Nvl(rsFormat!ͼ��, 0), "_" & rsFormat!���
                Else
                    'ȱʡΪA4����,����
                    tmpReport.Fmts.Add rsFormat!���, rsFormat!˵��, Nvl(rsFormat!W, INIT_WIDTH), Nvl(rsFormat!H, INIT_HEIGHT), _
                        Nvl(rsFormat!ֽ��, 9), Nvl(rsFormat!ֽ��, 1), Nvl(rsFormat!��ֽ̬��, 0) = 1, Nvl(rsFormat!ͼ��, 0), "_" & rsFormat!���
                End If
                rsFormat.MoveNext
            Next
            
            '���������������
            strSQL = _
                "Select A.Ԫ��ID,A.��������ID,A.������,A.����ֵ��Դ,A.Ĭ��,b.���� || '(' || b.��� || ')' as ������������ " & vbCr & _
                "From zlrptrelation A ,zlreports B " & vbCr & _
                "Where a.��������id=b.id and a.����ID=[1] "
            Set rsRelation = OpenSQLRecord(strSQL, "ReadReport", lngRPTID)
            
            '����������
            strSQL = _
                "Select A.����ID,A.Ԫ��ID,A.��������,A.�����ֶ�,A.������ϵ,A.����ֵ,A.������ɫ,A.������ɫ,A.�Ƿ�Ӵ� " & vbCr & _
                "    ,A.�Ƿ�����Ӧ��,a.���� " & vbCr & _
                "From zlRPTColProterty A " & vbCr & _
                "Where a.����ID=[1]"
            Set rsColProtertys = OpenSQLRecord(strSQL, "ReadReport", lngRPTID)
            
            '����Ԫ��(���мǣ������ǰ,���Ԫ���ں�,�������(��XY)����)
            strSQL = _
                "Select RowNum,ϵͳ,ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����" & vbCr & _
                "    ,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,��ID,Դ�к�,ԴID,���¼��" & vbCr & _
                "    ,���Ҽ��,�������,�������,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת " & vbCr & _
                "From zlRPTItems A " & vbCr & _
                "Where A.����ID=[1] " & vbCr & _
                "Order by NVL(��ID,0),A.��ʽ��,A.�ϼ�ID Desc,A.����,A.���,A.X,A.Y"
            Set rsItem = OpenSQLRecord(strSQL, "ReadReport", lngRPTID)
            Set rsSub = rsItem.Clone '���Ƽ�¼�����ڱ�������
            
            intMaxID = rsItem.RecordCount '�ؼ��������=Ԫ�ظ���(+��������)
            intCopyID = rsItem.RecordCount + 1 '�������ؼ���ʼ����
                    
            For i = 1 To rsItem.RecordCount
                Set tmpItem = New RPTItem
                With tmpItem
                    .id = rsItem!Rownum '�����IDΪ�ؼ�����(����������),(Rownum��һ������,����RowNum������ID���ϼ�ID��ϵ��ȷ)
                    .��ʽ�� = rsItem!��ʽ��
                    .���� = Nvl(rsItem!����)
                    .���� = Nvl(rsItem!����, 0)
                    .��� = Nvl(rsItem!���, 0)
                    .���� = Nvl(rsItem!����)
                    .���� = Nvl(rsItem!����, 0)
                    .���� = Nvl(rsItem!����)
                    .��ͷ = Nvl(rsItem!��ͷ)
                    .X = Nvl(rsItem!X, 0): .Y = Nvl(rsItem!Y, 0)
                    .W = Nvl(rsItem!W, 0): .H = Nvl(rsItem!H, 0)
                    .ˮƽ��ת = Nvl(rsItem!ˮƽ��ת, 0) = 1
                    If .���� = 6 And .W < 45 Then .W = 0
                    If .���� = 2 Or .���� = 6 Then
                        .�и� = Nvl(rsItem!�и�, 0)
                    Else
                        .�и� = Nvl(rsItem!�и�, 280)
                    End If
                    .����Ӧ�и� = Nvl(rsItem!����Ӧ�и�, 0)
                    .���� = Nvl(rsItem!����, 0) 'ȱʡ�����
                    .�Ե� = Nvl(rsItem!�Ե�, 0) = 1
                    
                    .���� = Nvl(rsItem!����, "����") 'ȱʡ����9��
                    If InStr("^" & gstrFonts & "^", "^" & .���� & "^") = 0 Then .���� = "����"
                    
                    .�ֺ� = Nvl(rsItem!�ֺ�, 9)
                    .���� = Nvl(rsItem!����, 0) = 1
                    .б�� = Nvl(rsItem!б��, 0) = 1
                    .���� = Nvl(rsItem!����, 0) = 1
                    .���� = Nvl(rsItem!����, 0) 'ȱʡ��ɫ
                    .ǰ�� = Nvl(rsItem!ǰ��, 0) 'ȱʡ��ɫ
                    .���� = Nvl(rsItem!����, &HFFFFFF) 'ȱʡ��ɫ
                    .�߿� = Nvl(rsItem!�߿�, 0) = 1
                    .Դ�к� = Nvl(rsItem!Դ�к�, 0)
                    .���Ҽ�� = Nvl(rsItem!���Ҽ��, 0)
                    .���¼�� = Nvl(rsItem!���¼��, 0)
                    .������� = Nvl(rsItem!�������, 0)
                    .������� = Nvl(rsItem!�������, 0)
                    .����߼Ӵ� = Nvl(rsItem!����߼Ӵ�, 0) = 1
                    If rsItem!ԴID & "" <> "" Then
                        rsData.Filter = "ID=" & rsItem!ԴID
                        If rsData.RecordCount > 0 Then
                            .����Դ = rsData!���� & ""
                        End If
                    End If
                     
                    'ȱʡ1��
                    .���� = Nvl(rsItem!����, 1)
                    If .���� <> 6 Then .���� = IIF(.���� < 1, 1, .����)
                    
                    .���� = Nvl(rsItem!����)
                    .��ʽ = Nvl(rsItem!��ʽ)
                    .���� = Nvl(rsItem!����)
                    .ϵͳ = Nvl(rsItem!ϵͳ, 0) = 1
                    
                    'ͼƬ�Ĵ���
                    If .���� = 11 Then
                        If gobjFile.FileExists(.����) Then
                            On Error Resume Next
                            Set .ͼƬ = LoadPicture(.����) 'ֱ�Ӵӱ��ض�,�ӿ��ٶ�
                            On Error GoTo errH
                        End If
                        If .ͼƬ Is Nothing Then
                            Set rsGraph = New ADODB.Recordset
                            strSQL = "Select Ԫ��ID,ͼƬ From zlRPTGraphs Where Ԫ��ID=[1]"
                            Set rsGraph = OpenSQLRecord(strSQL, "ReadReport", "|��ѯ��ʽ=1-LOB", Val(rsItem!id))
                            If Not rsGraph.EOF Then
                                Set .ͼƬ = GetImage(rsGraph.Fields("ͼƬ"))
                            End If
                            rsGraph.Close
                        End If
                    End If
                    
                    '�������Ĵ���(����Ϊ6,7,8,9)
                    If InStr(",6,7,8,9,", "," & .���� & ",") > 0 And Not IsNull(rsItem!�ϼ�ID) Then
                        rsSub.Filter = "ID=" & rsItem!�ϼ�ID
                        If Not rsSub.EOF Then
                            .�ϼ�ID = rsSub!Rownum '������ϼ�ID��Ӧ���ؼ�����
                            tmpReport.Items("_" & .�ϼ�ID).SubIDs.Add .id, "_" & .id
                        End If
                    End If
                    
                    '����������(�Զ���ͷ����Ч)
                    If .���� = 4 And .���� > 1 Then
                        For j = intCopyID To intCopyID + .���� - 2
                            .CopyIDs.Add j, "_" & j
                            intMaxID = intMaxID + 1 'һ��������һ��
                        Next
                        intCopyID = j
                    End If
                    If rsItem!��ID & "" <> "" Then
                        rsSub.Filter = "ID=" & rsItem!��ID
                        If Not rsSub.EOF Then
                            .��ID = rsSub!Rownum '������ϼ�ID��Ӧ���ؼ�����
                            tmpReport.Items("_" & .��ID).SubIDs.Add .id, "_" & .id
                        End If
                    End If
 
                    '��������ID(�ؼ�����)��Ϊ�ؼ��ּ��뼯��
                    Set tmpItem = tmpReport.Items.Add(.id, .��ʽ��, .����, .�ϼ�ID, .����, .���, .����, .����, .���� _
                        , .��ͷ, .X, .Y, .W, .H, .�и�, .����, .�Ե�, .����, .�ֺ�, .����, .����, .б��, .����, .ǰ�� _
                        , .����, .�߿�, .����, .����, .��ʽ, .����, .����߼Ӵ�, .����Ӧ�и�, .ͼƬ, .ϵͳ, .��ID, .SubIDs _
                        , .CopyIDs, "_" & .id, .����Դ, .���¼��, .���Ҽ��, .Դ�к�, .�������, .�������, , , .ˮƽ��ת)
                    
                    '�����������
                    rsRelation.Filter = "Ԫ��ID=" & rsItem!id
                    If rsRelation.RecordCount > 0 Then rsRelation.MoveFirst
                    For j = 1 To rsRelation.RecordCount
                        Set tmpRelation = New RPTRelation
                        With tmpRelation
                            .��������ID = Val(rsRelation!��������ID & "")
                            .������ = rsRelation!������ & ""
                            .����ֵ��Դ = rsRelation!����ֵ��Դ & ""
                            .������������ = rsRelation!������������ & ""
                            .Ĭ�� = Val(rsRelation!Ĭ�� & "")
                            tmpItem.Relations.Add .��������ID, .������, .����ֵ��Դ, .������������, .Ĭ��
                        End With
        
                        rsRelation.MoveNext
                    Next
                    '����������
                    rsColProtertys.Filter = "Ԫ��ID=" & rsItem!id
                    If rsColProtertys.RecordCount > 0 Then rsColProtertys.MoveFirst
                    For j = 1 To rsColProtertys.RecordCount
                        tmpItem.ColProtertys.Add rsColProtertys!��������, Nvl(rsColProtertys!�����ֶ�) _
                            , Nvl(rsColProtertys!������ϵ), Nvl(rsColProtertys!����ֵ), rsColProtertys!������ɫ _
                            , rsColProtertys!������ɫ, rsColProtertys!�Ƿ�Ӵ�, rsColProtertys!�Ƿ�����Ӧ�� _
                            , Nvl(rsColProtertys!����, 0), "_" & rsColProtertys!��������
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
'����:���汨������(����objReport)
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
        '�����������
        If Not objPan Is Nothing Then
            intCount = .Datas.count + .Items.count + .Fmts.count + 1
            For Each tmpData In .Datas
                '����ԴSQL
                If Len(Trim(tmpData.SQL)) > 0 Then intCount = intCount + UBound(Split(tmpData.SQL, vbCrLf)) + 1
                intCount = intCount + tmpData.Pars.count
            Next
        End If
        
        '��������(��ӡ���ò���)
        gcnOracle.Execute _
            "Update zlReports" & _
            "   Set ��ӡ��='" & .��ӡ�� & "',��ֽ=" & .��ֽ & "," & _
            "       Ʊ��=" & IIF(.Ʊ��, 1, 0) & ",��ӡ��ʽ=" & .��ӡ��ʽ & ",�޸�ʱ��=Sysdate" & ",��ֹ��ʼʱ��=to_date('" & Format(.��ֹ��ʼʱ��, "HH:mm:ss") & "','HH24:MI:SS')" & ",��ֹ����ʱ��=to_date('" & Format(.��ֹ����ʱ��, "HH:mm:ss") & "','HH24:MI:SS')" & _
            " Where ID=" & lngRPTID
        
        If Not objPan Is Nothing Then
            i = 1: Call ShowPercent(i / intCount, objPan)
        End If
        
        '��������Դ��ʷ��¼
        gcnOracle.Execute _
            "Insert Into zlRPTSQLSHistory " & vbNewLine & _
            "(����id, ����Դ����, �޸���, �޸�ʱ��, �к�, ����)" & vbNewLine & _
            "Select ����id, ����, �޸���, �޸�ʱ��, �к�, ���� " & vbNewLine & _
            "From " & vbNewLine & _
            "   (" & vbNewLine & _
            "    Select b.����id, b.����, '" & gstrLoginUserName & "' �޸���, Sysdate �޸�ʱ��, a.�к�, a.���� " & vbNewLine & _
            "    From zlRPTSQLs A, zlRPTDatas B " & vbNewLine & _
            "    Where a.Դid = b.Id And b.����id = " & lngRPTID & vbNewLine & _
            "    ) A " & vbNewLine & _
            "Where Not Exists(Select 1 From zlRPTSQLSHistory " & vbNewLine & _
            "                 Where ����id = a.����id and ����Դ���� = a.���� " & vbNewLine & _
            "                     And �޸�ʱ�� = a.�޸�ʱ�� and �޸��� = a.�޸��� ) "
            
        '��������Դ
        gcnOracle.Execute "Delete From zlRPTDatas Where ����ID=" & lngRPTID
        
        Set rsData = New ADODB.Recordset
        rsData.CursorLocation = adUseClient
        rsData.Open "Select ID,����ID,�������ӱ��,����,�ֶ�,����,����,˵�� From zlRPTDatas Where ID=0", gcnOracle, adOpenStatic, adLockOptimistic
        
        For Each tmpData In .Datas
            lngSQLID = GetNextID("zlRPTDatas")
            
            rsData.AddNew
            rsData!id = lngSQLID
            rsData!����ID = lngRPTID
            If tmpData.�������ӱ�� > 0 Then
                rsData!�������ӱ�� = tmpData.�������ӱ��
            Else
                rsData!�������ӱ�� = Null
            End If
            rsData!���� = tmpData.����
            rsData!�ֶ� = tmpData.�ֶ�
            rsData!���� = tmpData.����
            rsData!���� = tmpData.����
            rsData!˵�� = tmpData.˵��
            rsData.Update
            
            '����޸������ƣ���ͬ���޸�����Դ��ʷ��¼������
            If tmpData.ԭ���� <> "" Then
                gcnOracle.Execute "update Zlrptsqlshistory Set ����Դ����='" & tmpData.���� & "' where ����ID=" & lngRPTID & " And ����Դ����='" & tmpData.ԭ���� & "'"
                tmpData.ԭ���� = ""
            End If
            
            '����ԴSQL
            If Len(Trim(tmpData.SQL)) > 0 Then
                Set rsSQL = New ADODB.Recordset
                rsSQL.CursorLocation = adUseClient
                rsSQL.Open "Select ԴID,�к�,���� From zlRPTSQLs Where ԴID=0", gcnOracle, adOpenKeyset, adLockOptimistic
                For j = 0 To UBound(Split(tmpData.SQL, vbCrLf))
                    rsSQL.AddNew
                    rsSQL!ԴID = lngSQLID
                    rsSQL!�к� = j + 1
                    rsSQL!���� = CStr(Split(tmpData.SQL, vbCrLf)(j))
                    rsSQL.Update
                    If Not objPan Is Nothing Then
                        i = i + 1: Call ShowPercent(i / intCount, objPan)
                    End If
                Next
            End If
            
            '����Դ����
            Set rsPar = New ADODB.Recordset
            rsPar.CursorLocation = adUseClient
            rsPar.Open "Select ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����,���� " & _
                       "From zlRPTPars Where ԴID=0", gcnOracle, adOpenStatic, adLockOptimistic
            For Each tmpPar In tmpData.Pars
                rsPar.AddNew
                rsPar!ԴID = lngSQLID
                rsPar!���� = tmpPar.����
                rsPar!��� = tmpPar.���
                rsPar!���� = tmpPar.����
                rsPar!���� = tmpPar.����
                rsPar!��ʽ = tmpPar.��ʽ
                rsPar!ȱʡֵ = tmpPar.ȱʡֵ
                rsPar!ֵ�б� = tmpPar.ֵ�б�
                rsPar!����SQL = tmpPar.����SQL
                rsPar!��ϸSQL = tmpPar.��ϸSQL
                rsPar!�����ֶ� = tmpPar.�����ֶ�
                rsPar!��ϸ�ֶ� = tmpPar.��ϸ�ֶ�
                rsPar!���� = tmpPar.����
                rsPar!���� = IIF(tmpPar.�Ƿ�����, 1, 0)
                rsPar.Update
                If Not objPan Is Nothing Then
                    i = i + 1: Call ShowPercent(i / intCount, objPan)
                End If
            Next
            
            If Not objPan Is Nothing Then
                i = i + 1: Call ShowPercent(i / intCount, objPan)
            End If
        Next
    
        '�����ʽ
        gcnOracle.Execute "Delete From zlRPTFmts Where ����ID=" & lngRPTID
        For j = 1 To .Fmts.count
            gcnOracle.Execute "Insert Into zlRPTFmts(����ID,���,˵��,W,H,ֽ��,ֽ��,��ֽ̬��,ͼ��) Values(" & _
                lngRPTID & "," & .Fmts(j).��� & ",'" & .Fmts(j).˵�� & "'," & .Fmts(j).W & "," & .Fmts(j).H & "," & _
                .Fmts(j).ֽ�� & "," & .Fmts(j).ֽ�� & "," & IIF(.Fmts(j).��ֽ̬��, 1, 0) & "," & .Fmts(j).ͼ�� & ")"
            If Not objPan Is Nothing Then
                i = i + 1: Call ShowPercent(i / intCount, objPan)
            End If
        Next
        
        '����Ԫ��
        gcnOracle.Execute "Delete From zlRPTItems Where �ϼ�ID is Not NULL And ����ID=" & lngRPTID
        gcnOracle.Execute "Delete From zlRPTItems Where �ϼ�ID is NULL And ����ID=" & lngRPTID
        gcnOracle.Execute "Delete From zlRPTRelation Where ����ID=" & lngRPTID
        Set rsItem = New ADODB.Recordset
        rsItem.Fields.Append "ID", adBigInt
        rsItem.Fields.Append "dataid", adBigInt
        rsItem.CursorLocation = adUseClient
        rsItem.LockType = adLockOptimistic
        rsItem.CursorType = adOpenStatic
        rsItem.Open
        
        For Each tmpItem In .Items
            '�ȱ��濨Ƭ
            If tmpItem.���� = Val("14-��ƬԪ��") Then '�������
                
                lngItemID = GetNextID("zlRPTItems")
                rsItem.AddNew
                rsItem!id = tmpItem.id
                rsItem!dataid = lngItemID
                rsItem.Update
                lngTmp = 0
                If tmpItem.����Դ <> "" Then
                    rsData.Filter = "����='" & tmpItem.����Դ & "'"
                    If rsData.RecordCount > 0 Then
                        lngTmp = Val(rsData!id & "")
                    End If
                End If
                gcnOracle.Execute "Insert Into zlRPTItems(ID,����ID,��ʽ��,����,�ϼ�ID,����,���,����,����,����,��ͷ," & _
                    "X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,�߿�,����,ǰ��,����,����,��ʽ,����,����,ϵͳ,��ID," & _
                    "ԴID,Դ�к�,���Ҽ��,���¼��,�������,�������,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת) " & vbCr & _
                    "Values(" & _
                    lngItemID & "," & lngRPTID & "," & tmpItem.��ʽ�� & ",'" & tmpItem.���� & "',NULL," & tmpItem.���� & "," & _
                    tmpItem.��� & ",'" & tmpItem.���� & "'," & tmpItem.���� & ",'" & tmpItem.���� & "','" & _
                    tmpItem.��ͷ & "'," & tmpItem.X & "," & tmpItem.Y & "," & tmpItem.W & "," & tmpItem.H & "," & _
                    tmpItem.�и� & "," & tmpItem.���� & "," & Abs(CInt(tmpItem.�Ե�)) & ",'" & tmpItem.���� & "'," & _
                    tmpItem.�ֺ� & "," & Abs(CInt(tmpItem.����)) & "," & Abs(CInt(tmpItem.б��)) & "," & _
                    Abs(CInt(tmpItem.����)) & "," & Abs(CInt(tmpItem.�߿�)) & "," & tmpItem.���� & "," & tmpItem.ǰ�� & "," & _
                    tmpItem.���� & ",'" & tmpItem.���� & "','" & tmpItem.��ʽ & "','" & tmpItem.���� & "'," & _
                    IIF(tmpItem.���� = 0, 1, tmpItem.����) & "," & Abs(CInt(tmpItem.ϵͳ)) & "," & "Null" & _
                    "," & IIF(lngTmp = 0, "Null", lngTmp) & "," & tmpItem.Դ�к� & "," & tmpItem.���Ҽ�� & _
                    "," & tmpItem.���¼�� & "," & tmpItem.������� & "," & tmpItem.������� & _
                    "," & Abs(CInt(tmpItem.����߼Ӵ�)) & "," & IIF(tmpItem.����Ӧ�и�, "1", "Null") & _
                    "," & IIF(tmpItem.ˮƽ��ת, "1", "Null") & ")"
                
                If Not objPan Is Nothing Then
                    i = i + 1: Call ShowPercent(i / intCount, objPan)
                End If

            End If
        Next
        
        '��������Ԫ��
        For Each tmpItem In .Items
            '��������
            If InStr(",1,2,3,4,5,10,11,12,13,", "," & tmpItem.���� & ",") > 0 Then '�������
                lngItemID = GetNextID("zlRPTItems")
                rsItem.AddNew
                rsItem!id = tmpItem.id
                rsItem!dataid = lngItemID
                rsItem.Update
                lngParentID = 0
                If tmpItem.��ID <> 0 Then
                    rsItem.Filter = "ID=" & tmpItem.��ID
                    If rsItem.RecordCount > 0 Then
                        lngParentID = Val(rsItem!dataid & "")
                    End If
                End If
                gcnOracle.Execute "Insert Into zlRPTItems(ID,����ID,��ʽ��,����,�ϼ�ID,����,���,����,����,����,��ͷ," & _
                    "X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,�߿�,����,ǰ��,����,����,��ʽ,����,����,ϵͳ,��ID," & _
                    "ԴID,Դ�к�,���Ҽ��,���¼��,�������,�������,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת) " & vbCr & _
                    "Values(" & _
                    lngItemID & "," & lngRPTID & "," & tmpItem.��ʽ�� & ",'" & tmpItem.���� & "',NULL," & tmpItem.���� & "," & _
                    tmpItem.��� & ",'" & tmpItem.���� & "'," & tmpItem.���� & ",'" & tmpItem.���� & "','" & _
                    tmpItem.��ͷ & "'," & tmpItem.X & "," & tmpItem.Y & "," & tmpItem.W & "," & tmpItem.H & "," & _
                    tmpItem.�и� & "," & tmpItem.���� & "," & Abs(CInt(tmpItem.�Ե�)) & ",'" & tmpItem.���� & "'," & _
                    tmpItem.�ֺ� & "," & Abs(CInt(tmpItem.����)) & "," & Abs(CInt(tmpItem.б��)) & "," & _
                    Abs(CInt(tmpItem.����)) & "," & Abs(CInt(tmpItem.�߿�)) & "," & tmpItem.���� & "," & tmpItem.ǰ�� & "," & _
                    tmpItem.���� & ",'" & tmpItem.���� & "','" & tmpItem.��ʽ & "','" & tmpItem.���� & "'," & _
                    IIF(tmpItem.���� = 0, 1, tmpItem.����) & "," & Abs(CInt(tmpItem.ϵͳ)) & "," & _
                    IIF(lngParentID = 0, "Null", lngParentID) & "," & "Null" & "," & tmpItem.Դ�к� & "," & tmpItem.���Ҽ�� & _
                    "," & tmpItem.���¼�� & "," & tmpItem.������� & "," & tmpItem.������� & _
                    "," & Abs(CInt(tmpItem.����߼Ӵ�)) & "," & IIF(tmpItem.����Ӧ�и�, "1", "Null") & _
                    "," & IIF(tmpItem.ˮƽ��ת, "1", "Null") & ")"
                
                '��������ͼƬ�ֶ�
                If Not tmpItem.ͼƬ Is Nothing Then
                    Set rsGraph = New ADODB.Recordset
                    rsGraph.CursorLocation = adUseClient
                    rsGraph.Open "Select Ԫ��ID,ͼƬ From zlRPTGraphs Where Ԫ��ID=" & lngItemID, gcnOracle, adOpenStatic, adLockOptimistic
                    rsGraph.AddNew
                    rsGraph!Ԫ��ID = lngItemID
                    Call SaveImage(tmpItem.ͼƬ, rsGraph.Fields("ͼƬ"))
'                    If isFile(tmpItem.����) Then
'                        'ֱ�Ӷ�ȡ�ļ�����,��������
'                        Call SaveFile(tmpItem.����, rsGraph.Fields("ͼƬ"))
'                    Else
'                        Call SaveImage(tmpItem.ͼƬ, rsGraph.Fields("ͼƬ"))
'                    End If
                    rsGraph.Update
                End If
                
                If Not objPan Is Nothing Then
                    i = i + 1: Call ShowPercent(i / intCount, objPan)
                End If
                
                '�����������
                If tmpItem.���� = Val("4-�����") Or tmpItem.���� = Val("5-���ܱ�") Then
                    For Each tmpID In tmpItem.SubIDs
                        With .Items("_" & tmpID.id)
                            lngItemSubID = GetNextID("zlRPTItems")
                            gcnOracle.Execute "Insert Into zlRPTItems(ID,����ID,��ʽ��,�ϼ�ID,����,���,����,��ͷ,X,Y,W,H," & _
                                "�и�,����,����,�ֺ�,����,б��,����,�߿�,����,ǰ��,����,����,��ʽ,����,����,ϵͳ,�Ե�," & _
                                "��ID,����߼Ӵ�,����Ӧ�и�) " & vbCr & _
                                "Values(" & lngItemSubID & "," & lngRPTID & "," & .��ʽ�� & "," & lngItemID & _
                                "," & .���� & "," & .��� & ",'" & .���� & "','" & .��ͷ & "'," & .X & _
                                "," & .Y & "," & .W & "," & .H & "," & .�и� & "," & .���� & ",'" & .���� & "'," & .�ֺ� & _
                                "," & Abs(CInt(.����)) & "," & Abs(CInt(.б��)) & "," & Abs(CInt(.����)) & _
                                "," & Abs(CInt(.�߿�)) & "," & .���� & "," & .ǰ�� & "," & .���� & ",'" & .���� & "'" & _
                                ",'" & .��ʽ & "','" & .���� & "'," & .���� & "," & Abs(CInt(.ϵͳ)) & "," & Abs(CInt(.�Ե�)) & _
                                "," & IIF(Val(lngParentID) = 0, "NULL", IIF(lngParentID = 0, "Null", lngParentID)) & _
                                "," & Abs(CInt(.����߼Ӵ�)) & "," & IIF(.����Ӧ�и�, "1", "Null") & ")"
                            
                            '���������������
                            For j = 1 To .Relations.count
                                gcnOracle.Execute _
                                    "Insert Into zlRPTRelation(����ID,��������ID,Ԫ��ID,������,����ֵ��Դ,Ĭ��) " & vbCr & _
                                    "Values(" & _
                                    lngRPTID & "," & .Relations.Item(j).��������ID & "," & lngItemSubID & _
                                    ",'" & .Relations.Item(j).������ & "','" & .Relations.Item(j).����ֵ��Դ & "'" & _
                                    "," & .Relations.Item(j).Ĭ�� & ")"
                                If Not objPan Is Nothing Then
                                    i = i + 1: Call ShowPercent(i / intCount, objPan)
                                End If
                            Next
                            '����������
                            For j = 1 To .ColProtertys.count
                                gcnOracle.Execute _
                                    "Insert Into zlRPTColProterty " & vbCr & _
                                    "  (����ID,Ԫ��ID,��������,�����ֶ�,������ϵ,����ֵ,������ɫ,������ɫ,�Ƿ�Ӵ�,�Ƿ�����Ӧ��,����) " & vbCr & _
                                    "Values(" & _
                                    lngRPTID & "," & lngItemSubID & ",'" & .ColProtertys.Item(j).�������� & "'" & _
                                    ",'" & .ColProtertys.Item(j).�����ֶ� & "','" & .ColProtertys.Item(j).������ϵ & "'" & _
                                    ",'" & .ColProtertys.Item(j).����ֵ & "'," & Val(.ColProtertys.Item(j).������ɫ) & _
                                    "," & Val(.ColProtertys.Item(j).������ɫ) & "," & IIF(.ColProtertys.Item(j).�Ƿ�Ӵ�, 1, 0) & _
                                    "," & IIF(.ColProtertys.Item(j).�Ƿ�����Ӧ��, 1, 0) & _
                                    "," & Val(.ColProtertys.Item(j).����) & ")"
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
                '���������������
                For j = 1 To tmpItem.Relations.count
                    gcnOracle.Execute _
                        "Insert Into zlRPTRelation(����ID,��������ID,Ԫ��ID,������,����ֵ��Դ,Ĭ��) " & vbCr & _
                        "Values(" & _
                        lngRPTID & "," & tmpItem.Relations.Item(j).��������ID & "," & lngItemID & _
                        ",'" & tmpItem.Relations.Item(j).������ & "'" & _
                        ",'" & tmpItem.Relations.Item(j).����ֵ��Դ & "'" & _
                        "," & tmpItem.Relations.Item(j).Ĭ�� & ")"
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
    
    Set grsReport = Nothing '�������
    
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
'����:ȥ���ַ����������Ŀո�ͻس�(����ͷ�Ŀո�,�س�),��ȥ��TAB�ַ�,������������
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
'���ܣ���������������
    Dim tmpPar As RPTPar
    
    Set objOPars = New RPTPars
    For Each tmpPar In objSPars
        With tmpPar
            objOPars.Add .����, .���, .����, .����, .ȱʡֵ, .��ʽ, .ֵ�б�, .����SQL, .��ϸSQL, .�����ֶ�, .��ϸ�ֶ�, .���� _
                , "_" & .Key, .Reserve, .�Ƿ�����
        End With
    Next
End Sub

Public Function CheckPars(strSQL As String, strMsg As String, objPars As RPTPars) As Boolean
'���ܣ����SQL����в�����"[]"�Ƿ����,�Լ��������Ƿ���ȷ(������,������)
    Dim intLeft As Integer, intRight As Integer
    Dim intMin As Integer, intMax As Integer
    Dim strTmp As String, StrPar As String, strPars As String
    Dim i As Long, blnSort As Boolean
    Dim objPar As RPTPar
    
    '�ַ�����������ַ�ת��
    Call mdlPublic.TransSpecialChar(strSQL)
    
    For i = 1 To Len(strSQL)
        If Mid(strSQL, i, 1) = "[" Then intLeft = intLeft + 1
        If Mid(strSQL, i, 1) = "]" Then intRight = intRight + 1
    Next
    If intLeft <> intRight Then
        MsgBox "��ȷ�������ġ�[���롰]�����ųɶԣ�", vbInformation, App.Title
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
            Exit Function '�����ֱ��
        End If
        If CInt(StrPar) < intMin Then intMin = CInt(StrPar)
        If CInt(StrPar) > intMax Then intMax = CInt(StrPar)
        If InStr(strPars, "," & CInt(StrPar)) = 0 Then strPars = strPars & "," & CInt(StrPar)
    Loop
    If intMin <> 0 Then
        strMsg = "�����Ŷ��岻�Ǵ�0��ʼ��,�Ƿ��Զ�������Ĳ�����ǰ�ƣ�"
        If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton1, App.Title) = vbYes Then
            blnSort = True
        Else
            Exit Function '���Ǵ�0��ʼ���
        End If
    End If
    If strPars <> "" Then strPars = Mid(strPars, 2)
    If blnSort = False Then
        If UBound(Split(strPars, ",")) <> intMax Then
            strMsg = "�����Ŷ��岻�����������ֱ�ţ��Ƿ��Զ�������Ĳ�����ǰ�ƣ�"
            If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton1, App.Title) = vbYes Then
                blnSort = True
            Else
                Exit Function '�����������
            End If
        End If
    End If
    
    '�Զ�����
    If blnSort Then
        For i = 0 To UBound(Split(strPars, ","))
            If Split(strPars, ",")(i) <> i Then
                strSQL = Replace(strSQL, "[" & Split(strPars, ",")(i) & "]", "[" & i & "]")
                If objPars.count > UBound(Split(strPars, ",")) + 1 Then
                    For Each objPar In objPars
                        If objPar.��� > i Then
                            objPars("_" & Val(objPar.Key) - 1).Key = Val(objPar.Key) - 1
                            objPars("_" & Val(objPar.Key) - 1).Reserve = objPar.Reserve
                            objPars("_" & Val(objPar.Key) - 1).���� = objPar.����
                            objPars("_" & Val(objPar.Key) - 1).����SQL = objPar.����SQL
                            objPars("_" & Val(objPar.Key) - 1).�����ֶ� = objPar.�����ֶ�
                            objPars("_" & Val(objPar.Key) - 1).��ʽ = objPar.��ʽ
                            objPars("_" & Val(objPar.Key) - 1).���� = objPar.����
                            objPars("_" & Val(objPar.Key) - 1).���� = objPar.����
                            objPars("_" & Val(objPar.Key) - 1).��ϸSQL = objPar.��ϸSQL
                            objPars("_" & Val(objPar.Key) - 1).��ϸ�ֶ� = objPar.��ϸ�ֶ�
                            objPars("_" & Val(objPar.Key) - 1).ȱʡֵ = objPar.ȱʡֵ
                            objPars("_" & Val(objPar.Key) - 1).��� = objPar.��� - 1
                            objPars("_" & Val(objPar.Key) - 1).ֵ�б� = objPar.ֵ�б�
                        End If
                    Next
                    objPars.Remove "_" & objPars.count - 1
                End If
            End If
        Next
    End If
    
    '�ַ�����������ַ���ԭ
    Call mdlPublic.TransSpecialChar(strSQL, True)
    
    CheckPars = True
End Function

Public Function GetParCount(strSQL As String) As Integer
'���ܣ�����SQL����в����ĸ���,�����Ϊ׼
    Dim strTmp As String, StrPar As String, strPars As String
    
    strTmp = strSQL
    
    '�ַ�����������ַ�ת��
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
'���ܣ������δ�����ComboBox������ֵ
'������cbo=ComboBox,strFind=�����ַ���
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
'���ܣ�����ȱʡ�������SQL�����д�Ƿ���ȷ
'������strFieldInfo=�û������쳣�ֶΣ�������ʾ��Ĵ���λ�ö�λ
'      blnCheckInfo=�Ƿ�����ϸSQL
'      intCurConnect=��ǰ�������ӱ��
'���أ�
'     �ɹ�=SQL���ֶδ�,�����˸����ֶε����Ƽ�����,��ʽ��"����,111|����,111|����,123",����ֵ��ADO.Field.TypeΪ׼
'     ʧ��=��
    Dim rsTmp As New ADODB.Recordset, tmpFld As Field
    Dim strCheck As String, strLeft As String, strRight As String
    Dim StrPar As String, bytPar As Byte, i As Integer
    Dim strSQLinfo As String
    
    strCheck = strSQL
    
    '�ַ�����������ַ�ת��
    Call mdlPublic.TransSpecialChar(strCheck)
    
    If Not objPars Is Nothing Then
        Do While InStr(strCheck, "[") > 0
            strLeft = Left(strCheck, InStr(strCheck, "[") - 1)
            strRight = Mid(strCheck, InStr(strCheck, "]") + 1)
            StrPar = Mid(strCheck, InStr(strCheck, "[") + 1, InStr(strCheck, "]") - InStr(strCheck, "[") - 1)
            If Trim(StrPar) = "" Then StrPar = 0
            bytPar = CByte(StrPar)
            
            '��ȱʡ����ֵ�滻
            If objPars("_" & CInt(bytPar)).ȱʡֵ <> "" And Not objPars("_" & CInt(bytPar)).ȱʡֵ Like "*��" Then
                Select Case objPars("_" & CInt(bytPar)).����
                    Case 0 '�ַ�
                        StrPar = "'" & Replace(objPars("_" & CInt(bytPar)).ȱʡֵ, "'", "''") & "'"
                    Case 1 '����
                        StrPar = objPars("_" & CInt(bytPar)).ȱʡֵ
                    Case 2 '����
                        If Left(objPars("_" & CInt(bytPar)).ȱʡֵ, 1) = "&" Then
                            StrPar = GetParSQLMacro(objPars("_" & CInt(bytPar)).ȱʡֵ)
                        Else
                            If InStr(objPars("_" & CInt(bytPar)).ȱʡֵ, ":") > 0 Then
                                '��ʱ���ʽ
                                StrPar = "To_Date('" & Format(objPars("_" & CInt(bytPar)).ȱʡֵ, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            Else
                                '��ʱ���ʽ
                                StrPar = "To_Date('" & Format(objPars("_" & CInt(bytPar)).ȱʡֵ, "yyyy-MM-dd") & "','YYYY-MM-DD')"
                            End If
                        End If
                    Case 3 '������
                        StrPar = objPars("_" & CInt(bytPar)).ȱʡֵ
                End Select
            Else 'ȱʡֵΪ�ջ�Ϊ�Զ�����
                Select Case objPars("_" & CInt(bytPar)).����
                    Case 0 '�ַ�
                        StrPar = "'�մ�'"
                    Case 1 '����
                        StrPar = 0
                    Case 2 '����
                        StrPar = "Sysdate"
                    Case 3 '������(ֱ���滻)
                        If objPars("_" & CInt(bytPar)).ȱʡֵ = "�̶�ֵ�б�" Then
                            'ȡ�̶�ֵ�е�ȱʡֵ
                            '���õķָ���
                            For i = 0 To UBound(Split(objPars("_" & CInt(bytPar)).ֵ�б�, "|"))
                                If Left(Split(objPars("_" & CInt(bytPar)).ֵ�б�, "|")(i), 1) = "��" Then
                                    StrPar = Split(Split(objPars("_" & CInt(bytPar)).ֵ�б�, "|")(i), ",")(1)
                                    Exit For
                                End If
                            Next
                            'û������ȱʡֵ��ȡ��һ��
                            If StrPar = "" Then
                                StrPar = Split(Split(objPars("_" & CInt(bytPar)).ֵ�б�, "|")(0), ",")(1)
                            End If
                        ElseIf objPars("_" & CInt(bytPar)).ȱʡֵ = "ѡ�������塭" Then
                            If objPars("_" & CInt(bytPar)).ֵ�б� <> "" Then
                                'ȡȱʡ��ֵ
                                StrPar = Split(objPars("_" & CInt(bytPar)).ֵ�б�, "|")(1)
                            ElseIf objPars("_" & CInt(bytPar)).��ϸSQL <> "" And objPars("_" & CInt(bytPar)).��ϸ�ֶ� <> "" Then
                                strSQLinfo = objPars("_" & CInt(bytPar)).��ϸSQL
                                Call CheckParsRela(strSQLinfo, objDatas, objPars("_" & CInt(bytPar)).����, True)
                                StrPar = GetDefaultValue(strSQLinfo, objPars("_" & CInt(bytPar)).��ϸ�ֶ�)
                                If StrPar <> "" And InStr(StrPar, "|") > 0 Then StrPar = CStr(Split(StrPar, "|")(1))
                                
                                If objPars("_" & CInt(bytPar)).��ʽ = 1 Then
                                    StrPar = " In (" & StrPar & ") "
                                End If
                            Else
                                StrPar = ""
                            End If
                        Else
                            StrPar = objPars("_" & CInt(bytPar)).ȱʡֵ
                        End If
                End Select
            End If
            strCheck = strLeft & StrPar & strRight
        Loop
    End If
    
    '�ַ�����������ַ���ԭ
    Call mdlPublic.TransSpecialChar(strCheck, True)
    
    strSQLref = strCheck
    If InStr(UCase(strCheck), "WHERE ") > 0 Then
        strCheck = Replace(UCase(strCheck), "WHERE ", "Where Rownum<1 And ")
    End If
    
    Err.Clear
    On Error Resume Next
    Call OpenRecord(rsTmp, strCheck, "mdlPublic_CheckSQL", intCurConnect)  '�滻�ɵĶ��ǹ̶�����,ͬһ����Դһ�㲻��,����SQLҲ�����������
    If Err.Number = 0 Then
        strErr = ""
        For Each tmpFld In rsTmp.Fields
            If InStr(tmpFld.name, "|") > 0 Then
                strErr = "�ֶ�""" & tmpFld.name & """û�б�����"
                If strFieldInfo = "" Then strFieldInfo = tmpFld.name
                CheckSQL = "": Exit Function
            ElseIf InStr(tmpFld.name, "'") > 0 Or InStr(tmpFld.name, """") > 0 Then
                strErr = "�ֶ��� " & tmpFld.name & " �Ƿ���"
                If strFieldInfo = "" Then strFieldInfo = tmpFld.name
                CheckSQL = "": Exit Function
            Else
                If InStr(CheckSQL & "|", "|" & tmpFld.name & "," & tmpFld.type & "|") = 0 Then
                    CheckSQL = CheckSQL & "|" & tmpFld.name & "," & tmpFld.type
                Else
                    strErr = "������Դ�з�����ͬ���ֶ���Ŀ��"
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
'���ܣ�������"'"���ŵ��ַ�������ΪOracle����ʶ����ַ�����
'˵�����Զ�(����)�����߼�"'"�綨����

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
'����:���������б���ָ�㶨���Ĳ������
    Dim strName As String
    Dim objTmp As Object
    
    strName = objNode.Text
    Set objTmp = objNode
    
    Do While Not objTmp.Parent.Parent Is Nothing
        If objTmp.Parent.Text Like "*��*��" Then
            strName = Split(objTmp.Parent.Text, "��")(0) & "." & strName
        Else
            strName = objTmp.Parent.Text & "." & strName
        End If
        Set objTmp = objTmp.Parent
    Loop
    LevelText = UCase(strName)
End Function

Public Function GetObjRECT(lngHWND As Long) As RECT
'����:��ȡ����(�����ؼ�)�Ŀɼ��ߴ�����(������Ϊ��λ)
'˵��:����ɽ��GetCaptionHeight��GetVscWidth��GetHscHeight����ʹ��
    Dim Area As RECT
    GetWindowRect lngHWND, Area
    GetObjRECT = Area
End Function

Public Function MakeFile(strID As String, Optional strFormat As String = "CUSTOM") As String
'����:����Դ�ļ��е�ָ����Դ���ɴ����ļ�
'����:ID=��Դ��,strExt=Ҫ�����ļ�����չ��(��BMP)
'����:�����ļ���
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
'���ܣ���ʾ�����صȴ�����ȴ���(strInfo)
'����:strInfo=�ȴ��������ʾ��Ϣ
'     sngPer=����
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
                '��ʾ�ȴ�
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
                '��ʾ����
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
'���ܣ����ñ��̶��о��ж���
    Dim i As Long, j As Long
    Dim blnRedraw As Boolean
    Dim lngRow As Long, lngCol As Long

    blnRedraw = msh.Redraw: lngRow = msh.Row: lngCol = msh.Col: msh.Redraw = False
    For i = 0 To msh.FixedRows - 1
        msh.Row = i
        For j = 0 To msh.Cols - 1
            msh.Col = j
            If i <= msh.FixedRows - 2 And j <= msh.FixedCols - 1 Then '��������ͷʱ,���������������
                msh.CellAlignment = 7
            Else
                msh.CellAlignment = 4
            End If
        Next
    Next
    msh.Row = lngRow: msh.Col = lngCol: msh.Redraw = blnRedraw
End Sub

Public Function GetParSQLMacro(Str As String) As String
'����:�������������,������ת�������SQL����п��õ�ֵ
    Dim curDate As Date
    
    If InStr(Str, "&") = 0 Then GetParSQLMacro = Str: Exit Function
    
    curDate = Currentdate
    
    Select Case Str
        Case "&��ǰ����"
            GetParSQLMacro = "TO_DATE('" & Format(curDate, "yyyy-MM-dd") & "','YYYY-MM-DD')"
        Case "&��ǰ����ʱ��"
            GetParSQLMacro = "Sysdate"
        Case "&���쿪ʼʱ��"
            GetParSQLMacro = "TO_DATE('" & Format(curDate, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&�������ʱ��"
            GetParSQLMacro = "TO_DATE('" & Format(curDate, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&ǰһ�쿪ʼʱ��"
            GetParSQLMacro = "TO_DATE('" & Format(curDate - 1, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&ǰһ�����ʱ��"
            GetParSQLMacro = "TO_DATE('" & Format(curDate - 1, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&ǰһ��ͬʱ��"
            GetParSQLMacro = "Sysdate-1"
        Case "&��һ��ͬʱ��"
            GetParSQLMacro = "Sysdate+1"
        Case "&��һ�����ʱ��"
            GetParSQLMacro = "TO_DATE('" & Format(curDate + 1, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&��һ������"
            GetParSQLMacro = "Trunc(Sysdate+1)"
        Case "&ǰһ������"
            GetParSQLMacro = "Trunc(Sysdate - 7)"
        Case "&ǰһ������"
            GetParSQLMacro = "TO_DATE('" & Format(DateAdd("m", -1, curDate), "yyyy-MM-dd") & "','YYYY-MM-DD')"
        Case "&ǰһ������"
            GetParSQLMacro = "TO_DATE('" & Format(DateAdd("m", -3, curDate), "yyyy-MM-dd") & "','YYYY-MM-DD')"
        Case "&ǰһ������"
            GetParSQLMacro = "TO_DATE('" & Format(DateAdd("yyyy", -1, curDate), "yyyy-MM-dd") & "','YYYY-MM-DD')"
        Case "&��һ������"
            GetParSQLMacro = "Trunc(Sysdate + 7)"
        Case "&��һ������"
            GetParSQLMacro = "TO_DATE('" & Format(DateAdd("m", 1, curDate), "yyyy-MM-dd") & "','YYYY-MM-DD')"
        Case "&��һ������"
            GetParSQLMacro = "TO_DATE('" & Format(DateAdd("m", 3, curDate), "yyyy-MM-dd") & "','YYYY-MM-DD')"
        Case "&��һ������"
            GetParSQLMacro = "TO_DATE('" & Format(DateAdd("yyyy", 1, curDate), "yyyy-MM-dd") & "','YYYY-MM-DD')"
        Case "&���³�ʱ��"
            GetParSQLMacro = "TO_DATE('" & Format(Year(curDate) & "-" & Month(curDate) & "-01", "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&����ĩʱ��"
            curDate = DateAdd("m", 1, curDate)
            curDate = CDate(Year(curDate) & "-" & Month(curDate) & "-01") - 1
            GetParSQLMacro = "TO_DATE('" & Format(curDate, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&���³�ʱ��"
            curDate = DateAdd("m", -1, curDate)
            GetParSQLMacro = "TO_DATE('" & Format(Year(curDate) & "-" & Month(curDate) & "-01", "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&����ĩʱ��"
            curDate = CDate(Year(curDate) & "-" & Month(curDate) & "-01") - 1
            GetParSQLMacro = "TO_DATE('" & Format(curDate, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&�����ʱ��"
            GetParSQLMacro = "TO_DATE('" & Format(Year(curDate) & "-01-01", "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&����ĩʱ��"
            GetParSQLMacro = "TO_DATE('" & Format(Year(curDate) & "-12-31", "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&�����ʱ��"
            GetParSQLMacro = "TO_DATE('" & Format(Year(curDate) - 1 & "-01-01", "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&����ĩʱ��"
            GetParSQLMacro = "TO_DATE('" & Format(Year(curDate) - 1 & "-12-31", "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
    End Select
End Function

Public Function GetParVBMacro(Str As String) As String
'����:�������������,������ת�����VB����ֵ
    Dim curDate As Date
    
    If InStr(Str, "&") = 0 Then GetParVBMacro = Str: Exit Function
    
    curDate = Currentdate
    Select Case Str
        Case "&��ǰ����"
            GetParVBMacro = Format(curDate, "yyyy-MM-dd")
        Case "&��ǰ����ʱ��"
            GetParVBMacro = Format(curDate, "yyyy-MM-dd HH:mm:ss")
        Case "&ǰһ������"
            GetParVBMacro = Format(curDate - 7, "yyyy-MM-dd")
        Case "&ǰһ������"
            GetParVBMacro = Format(DateAdd("m", -1, curDate), "yyyy-MM-dd")
        Case "&ǰһ������"
            GetParVBMacro = Format(DateAdd("m", -3, curDate), "yyyy-MM-dd")
        Case "&ǰһ������"
            GetParVBMacro = Format(DateAdd("yyyy", -1, curDate), "yyyy-MM-dd")
        Case "&��һ������"
            GetParVBMacro = Format(curDate + 7, "yyyy-MM-dd")
        Case "&��һ������"
            GetParVBMacro = Format(DateAdd("m", 1, curDate), "yyyy-MM-dd")
        Case "&��һ������"
            GetParVBMacro = Format(DateAdd("m", 3, curDate), "yyyy-MM-dd")
        Case "&��һ������"
            GetParVBMacro = Format(DateAdd("yyyy", 1, curDate), "yyyy-MM-dd")
        Case "&���쿪ʼʱ��"
            GetParVBMacro = Format(curDate, "yyyy-MM-dd 00:00:00")
        Case "&�������ʱ��"
            GetParVBMacro = Format(curDate, "yyyy-MM-dd 23:59:59")
        Case "&ǰһ�쿪ʼʱ��"
            GetParVBMacro = Format(curDate - 1, "yyyy-MM-dd 00:00:00")
        Case "&ǰһ�����ʱ��"
            GetParVBMacro = Format(curDate - 1, "yyyy-MM-dd 23:59:59")
        Case "&ǰһ��ͬʱ��"
            GetParVBMacro = Format(curDate - 1, "yyyy-MM-dd HH:mm:ss")
        Case "&��һ��ͬʱ��"
            GetParVBMacro = Format(curDate + 1, "yyyy-MM-dd HH:mm:ss")
        Case "&��һ�����ʱ��"
            GetParVBMacro = Format(curDate + 1, "yyyy-MM-dd 23:59:59")
        Case "&��һ������"
            GetParVBMacro = Format(curDate + 1, "yyyy-MM-dd")
        Case "&���³�ʱ��"
            GetParVBMacro = Format(Year(curDate) & "-" & Month(curDate) & "-01", "yyyy-MM-dd 00:00:00")
        Case "&����ĩʱ��"
            curDate = DateAdd("m", 1, curDate)
            curDate = CDate(Year(curDate) & "-" & Month(curDate) & "-01") - 1
            GetParVBMacro = Format(curDate, "yyyy-MM-dd 23:59:59")
        Case "&���³�ʱ��"
            curDate = DateAdd("m", -1, curDate)
            GetParVBMacro = Format(Year(curDate) & "-" & Month(curDate) & "-01", "yyyy-MM-dd 00:00:00")
        Case "&����ĩʱ��"
            curDate = CDate(Year(curDate) & "-" & Month(curDate) & "-01") - 1
            GetParVBMacro = Format(curDate, "yyyy-MM-dd 23:59:59")
        Case "&�����ʱ��"
            GetParVBMacro = Format(Year(curDate) & "-01-01", "yyyy-MM-dd 00:00:00")
        Case "&����ĩʱ��"
            GetParVBMacro = Format(Year(curDate) & "-12-31", "yyyy-MM-dd 23:59:59")
        Case "&�����ʱ��"
            GetParVBMacro = Format(Year(curDate) - 1 & "-01-01", "yyyy-MM-dd 00:00:00")
        Case "&����ĩʱ��"
            GetParVBMacro = Format(Year(curDate) - 1 & "-12-31", "yyyy-MM-dd 23:59:59")
    End Select
End Function

Public Function GetParUserMacro(Str As String) As String
'����:�������������,������ת����ı��������ʽֵ
    Dim curDate As Date
    
    If InStr(Str, "&") = 0 Then GetParUserMacro = Str: Exit Function
    
    curDate = Currentdate
    Select Case Str
        Case "&��ǰ����"
            GetParUserMacro = Format(curDate, "yyyy��MM��dd��")
        Case "&��ǰ����ʱ��"
            GetParUserMacro = Format(curDate, "yyyy��MM��dd�� HH:mm:ss")
        Case "&ǰһ������"
            GetParUserMacro = Format(curDate - 7, "yyyy��MM��dd��")
        Case "&ǰһ������"
            GetParUserMacro = Format(DateAdd("m", -1, curDate), "yyyy��MM��dd��")
        Case "&ǰһ������"
            GetParUserMacro = Format(DateAdd("m", -3, curDate), "yyyy��MM��dd��")
        Case "&ǰһ������"
            GetParUserMacro = Format(DateAdd("yyyy", -1, curDate), "yyyy��MM��dd��")
        Case "&��һ������"
            GetParUserMacro = Format(curDate + 7, "yyyy��MM��dd��")
        Case "&��һ������"
            GetParUserMacro = Format(DateAdd("m", 1, curDate), "yyyy��MM��dd��")
        Case "&��һ������"
            GetParUserMacro = Format(DateAdd("m", 3, curDate), "yyyy��MM��dd��")
        Case "&��һ������"
            GetParUserMacro = Format(DateAdd("yyyy", 1, curDate), "yyyy��MM��dd��")
        Case "&���쿪ʼʱ��"
            GetParUserMacro = Format(curDate, "yyyy��MM��dd�� 00:00:00")
        Case "&�������ʱ��"
            GetParUserMacro = Format(curDate, "yyyy��MM��dd�� 23:59:59")
        Case "&ǰһ�쿪ʼʱ��"
            GetParUserMacro = Format(curDate - 1, "yyyy��MM��dd�� 00:00:00")
        Case "&ǰһ�����ʱ��"
            GetParUserMacro = Format(curDate - 1, "yyyy��MM��dd�� 23:59:59")
        Case "&ǰһ��ͬʱ��"
            GetParUserMacro = Format(curDate - 1, "yyyy��MM��dd�� HH:mm:ss")
        Case "&��һ��ͬʱ��"
            GetParUserMacro = Format(curDate + 1, "yyyy��MM��dd�� HH:mm:ss")
        Case "&��һ�����ʱ��"
            GetParUserMacro = Format(curDate + 1, "yyyy��MM��dd�� 23:59:59")
        Case "&��һ������"
            GetParUserMacro = Format(curDate + 1, "yyyy��MM��dd��")
        Case "&���³�ʱ��"
            GetParUserMacro = Format(Year(curDate) & "-" & Month(curDate) & "-01", "yyyy��MM��dd��")
        Case "&����ĩʱ��"
            curDate = DateAdd("m", 1, curDate)
            curDate = CDate(Year(curDate) & "-" & Month(curDate) & "-01") - 1
            GetParUserMacro = Format(curDate, "yyyy��MM��dd��")
        Case "&���³�ʱ��"
            curDate = DateAdd("m", -1, curDate)
            GetParUserMacro = Format(Year(curDate) & "-" & Month(curDate) & "-01", "yyyy��MM��dd��")
        Case "&����ĩʱ��"
            curDate = CDate(Year(curDate) & "-" & Month(curDate) & "-01") - 1
            GetParUserMacro = Format(curDate, "yyyy��MM��dd��")
        Case "&�����ʱ��"
            GetParUserMacro = Format(Year(curDate) & "-01-01", "yyyy��MM��dd��")
        Case "&����ĩʱ��"
            GetParUserMacro = Format(Year(curDate) & "-12-31", "yyyy��MM��dd��")
        Case "&�����ʱ��"
            GetParUserMacro = Format(Year(curDate) - 1 & "-01-01", "yyyy��MM��dd��")
        Case "&����ĩʱ��"
            GetParUserMacro = Format(Year(curDate) - 1 & "-12-31", "yyyy��MM��dd��")
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
'���ܣ����������б�����
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
'���ܣ����ر��й�ʽ��������Ŀ�ĸ���
    Dim strTmp As String, StrPar As String
    
    strTmp = strFormula
    
    Do While InStr(strTmp, "[") > 0
        strTmp = Mid(strTmp, InStr(strTmp, "[") + 1)
        StrPar = Left(strTmp, InStr(strTmp, "]") - 1)
        If InStr(StrPar, ".") > 0 Then GetItemCount = GetItemCount + 1
    Loop
End Function

Public Function GetNodeType(strNode As String, ByVal tvw As Control) As Long
'���ܣ��ɽ��·��������������
'�����������,��"A.B"
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
'���ܣ�����ָ����Ԫ��ĺϲ���Χ
'˵�����ϲ��ĵ�Ԫ��ֻ����һ������,��ֻ�ڹ̶��з�Χ��,Ϊ�յĵ�Ԫ�������ⵥԪ��ϲ�
    Dim intRowB As Integer, intRowE As Integer
    Dim intColB As Integer, intColE As Integer
    Dim i As Integer
    
    'Ѱ�ҿ�ʼ��
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
    'Ѱ�ҽ�����
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
    'Ѱ�ҿ�ʼ��
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
    'Ѱ�ҽ�����
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
'���ܣ���ָ���ļ�¼��ͼ���ֶθ���Ϊͼ����ʱ�ļ�
'������objField=ͼ���ֶζ���
'���أ���ʱ������ͼƬ�ļ���

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
'���ܣ��ӵ�ǰ����(frmparent.mobjreport)�л�ȡָ��������ֵ(ȱʡֵ����ֵ�����ֵ)
'˵���������Ӧ����Դ�ڱ�����δʹ��,��ָ���������ܷ���(��)
    Dim tmpPar As RPTPar, tmpData As RPTData
    
    For Each tmpData In frmParent.mobjReport.Datas
        For Each tmpPar In tmpData.Pars
            If tmpPar.���� = strName Then
                If tmpPar.Reserve Like "*��|*" Then
                    If Split(tmpPar.Reserve, "|")(1) <> "������" Then
                        GetParValue = Split(tmpPar.Reserve, "|")(1)
                    End If
                End If
                If GetParValue <> "" Then Exit Function
                If tmpPar.���� = 2 Then
                    If Left(tmpPar.ȱʡֵ, 1) = "&" Then
                        GetParValue = GetParUserMacro(tmpPar.ȱʡֵ)
                    ElseIf InStr(tmpPar.ȱʡֵ, ":") = 0 Then
                        GetParValue = Format(tmpPar.ȱʡֵ, "yyyy��MM��dd��")
                    Else
                        GetParValue = Format(tmpPar.ȱʡֵ, "yyyy��MM��dd�� HH:mm:ss")
                    End If
                Else
                    If tmpPar.ȱʡֵ Like "*��" Then
                        If tmpPar.ֵ�б� Like "*|*" Then '��ʱ����ˣ���ʾֵ|��ֵ
                            GetParValue = Split(tmpPar.ֵ�б�, "|")(0)
                        Else
                            GetParValue = ""
                        End If
                    Else
                        GetParValue = tmpPar.ȱʡֵ
                    End If
                End If
                Exit Function
            End If
        Next
    Next
End Function

Public Function GetUserParData(frmParent As Object, intTime As Integer) As String
'���ܣ���ȡ�û����������,��intTime�ĸ�,��0��ʼ��
'˵�������û�д���,���ܷ���(Ϊ��)
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
'���ܣ�����ͼ��Ԫ�����ƴӴ�������ж�ȡͼƬ����
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
'���ܣ��Ӵ�������м���Ƿ��д����ͼ���ļ�
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
'���ܣ����ش�ӡ������,���ܸı�ֽ��
'������blnOutCall=�Ƿ��ⲿͨ���ӿ��ڵ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim frmShow As New frmLocalSet
    
    On Error GoTo errH
    
    If Printers.count = 0 Then MsgBox "��ϵͳ��û�м�⵽�κδ�ӡ�豸,���Ȱ�װ��ӡ���������Ըò�����", vbInformation, App.Title: Exit Function
    
    If TypeName(varReport) = "String" Then
        strSQL = "Select ID,���,����,˵��,����,��ӡ��,��ֽ,Ʊ��,��ӡ��ʽ,ϵͳ,����ID,����,�޸�ʱ��,����ʱ��,��ֹ��ʼʱ��,��ֹ����ʱ�� From zlReports Where ���=[1] And Nvl(ϵͳ,0)=[3]"
    Else
        strSQL = "Select ID,���,����,˵��,����,��ӡ��,��ֽ,Ʊ��,��ӡ��ʽ,ϵͳ,����ID,����,�޸�ʱ��,����ʱ��,��ֹ��ʼʱ��,��ֹ����ʱ�� From zlReports Where ����ID=[2] And Nvl(ϵͳ,0)=[3]"
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
'���ܣ�����ȫ�ֶ���gobjReport������,�򿪲���ʾ��������
'˵����
'   1.�����ӡ�����Ա�������Ϊ��
'   2.ʹ�øú���֮ǰ�����ʼgobjReport,garrPars,glngGroup��ֵ
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
'���ܣ�����ȫ�ֶ���gobjReport������,�򿪲���ʾ��������
'˵����
'   1.�����ӡ�����Ա�������Ϊ��
'   2.ʹ�øú���֮ǰ�����ʼgobjReport,garrPars,glngGroup��ֵ
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
'���ܣ�����ȫ�ֶ���gobjReport������,�򿪲���ʾ��������
'˵����
'   1.�����ӡ�����Ա�������Ϊ��
'   2.ʹ�øú���֮ǰ�����ʼgobjReport,garrPars,glngGroup��ֵ
    Set objfrmReport = New frmReport
    Set objfrmShow = objfrmReport.GetReportForm(frmParent, objCurDLL, LibDatas, garrPars, bytStyle)
    GetReportFrom = True
End Function

Public Function GetAutoFont(ByVal strText As String, ByVal lngW As Long, ByVal lngH As Long _
    , ByVal objFont As StdFont, objBase As Object, Optional ByVal blnWarp As Boolean = True _
    , Optional ByVal sngYDistance As Single) As StdFont
    
'���ܣ���ȡ��ָ����С�������������������ĺ�������
'������strText=Ҫ��������֣����Զ����м��㣬���԰���Ӳ�س�
'      lngW,lngH=ָ�������С
'      objFont=ԭʼ�������
'      objBase=���ڼ������ʱ����(Form��PictureBox��Printer)
'      blnWarp=�Ƿ��Զ����м���
'      sngYDistance=�Զ����еĶ�������ʱ���о�,ȱʡΪ0��(Point)
'���أ� �����������
'˵����������ִ��ʱ��Ϊx/100�룬���˴������á�

    Dim lngX As Long, lngY As Long, i As Long
    Dim lngOneH As Long, lngLen As Long
    Dim strChar As String, strNext As String
    Dim sngSize As Currency
    Dim LINE_W As Integer
    
    strText = Replace(strText, vbCrLf, vbCr)
    strText = Replace(strText, vbLf, vbCr)
    If Not blnWarp Then strText = Replace(strText, vbCr, "")
    
    'TextWidth/TextHeight�������ֶ��������
    'vbCr���Ϊ0���߶�Ϊ2�У�""���Ϊ0���߶�Ϊһ��
    If Trim(Replace(strText, vbCr, "")) = "" Then Call CopyFont(objFont, GetAutoFont): Exit Function
    
    Call CopyFont(objFont, objBase.Font)
    'If objBase.TextWidth(strText) <= lngW And objBase.TextHeight(strText) <= lngH Then
    If objBase.TextWidth("A") * LenB(StrConv(strText, vbFromUnicode)) <= lngW And objBase.TextHeight(strText) <= lngH Then
        Call CopyFont(objFont, GetAutoFont): Exit Function
    End If
    
    '��Ԫ��߿�̶�2������
    LINE_W = Screen.TwipsPerPixelX * 2

    sngYDistance = objBase.ScaleY(sngYDistance, vbPoints, vbTwips)
    lngLen = Len(strText)
    lngW = lngW - 2 * LINE_W
    lngH = lngH - 2 * LINE_W
    
    Do While True
        '��ǰ�ֺ�ģ���������
        lngX = LINE_W: lngY = LINE_W
        
        lngOneH = objBase.TextHeight("��")
        
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
        
        '��ǰ�ֺŹ���
        If i > Len(strText) Then Exit Do
        
        '��ǰ�ֺŹ���,��С�ֺŵĴ���
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
'���ܣ���ָ���豸�ϰ�ָ����ʽ��������ֻ�ͼ��
'������
'   Dev=����豸,ΪPrinter��PictureBox����
'   Data=�������,Ϊ����(x)���ַ���("xxx")��ͼ��(stdPicture)���ַ���������vbCrLf,��Data����Ϊ������ʱ,��ʾ�������
'   TW,TH=������޶���Χ,���������Χ���Զ�ȡ������С,Ϊ0ʱ��Ч
'   Border=�߿���,��������,"1111"��ʾȫ��
'   Align=���ֶ���,0=��,1=��,2=��,��ˮƽ���뼰��ֱ����
'   Wrap=���������Ϊ�ַ���ʱ,��ʾ�Ƿ��Զ����У����Զ�����ʱ,�����ݲ������
'        ���������ΪͼƬʱ����ʾ�Ƿ񱣳�ͼƬ�Ŀ�߱���(������),ͬʱ����������Ч
'   Ratio=�������,������,���궼��Ӱ��,ȱʡΪ1(100%)
'   sngYDistance=�Զ����еĶ�������ʱ���о�,ȱʡΪ0��(Single������Ϊ�����ż���)
'   blnBold=������������ʱ�Ƿ�Ӵ�
'   bytShape=���ߵ���״��0-���Σ�1-Բ��
'   blnCellWordWrap=��Ԫ��߶��Զ�����
'˵����
'   1.��ʹ�øú���֮ǰ,Ӧ��û�иı��豸����ͼ��ʼֵ
'   2.�����λ���λ���ڱ��������Χ�����Ͻ�

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
    
    '��Χ�޶�
    If TW > 0 Then
        If X > TW Then Exit Function
        If X + W > TW Then W = TW - X
    End If
    If TH > 0 Then
        If Y > TH Then Exit Function
        If Y + H > TH Then H = TH - Y
    End If
    
    If TypeName(Dev) = "Printer" Then
        '������ĻÿӢ���������ӡ��ÿӢ�����羣�������ȣ��İٷֱ�
        sngTmp = Screen.TwipsPerPixelY / Printer.TwipsPerPixelY
        Dev.DrawWidth = Round(IIF(blnBold, 2, 1) * sngTmp, 0)
    Else
        Dev.DrawWidth = IIF(blnBold, 2, 1)
    End If
    
    If TypeName(Data) = "Integer" Then
        X = X * Ratio: Y = Y * Ratio: W = W * Ratio: H = H * Ratio
        If Val(Data) < 0 Then
            'ע��
            '  1.����������Ϊ0����Ϊ1�������ӡ���Ŀ��Ŀ������Ч��Ԥ������Ӱ�졣
            '  2.�ϱ����еľ��η����ܿ��ģ�������������Ԫ�أ�����������ľ��η���Ŀ���û�����⡣
            Dev.FillStyle = vbFSSolid: Dev.FillStyle = vbFSTransparent
            If bytShape = 0 Then
                '���ľ���
                Dev.Line (X, Y)-(X + W - IIF(W > 0, Screen.TwipsPerPixelX * Ratio, 0) _
                    , Y + H - IIF(H > 0, Screen.TwipsPerPixelY * Ratio, 0)), ForeColor, B
            Else
                '����Բ�Ρ���Բ��
                Dev.Circle (X + W / 2, Y + H / 2), IIF(H > W, H, W) / 2, ForeColor, , , H / W
            End If
        Else
            'ʵ�ľ���
            Dev.Line (X, Y)-(X + W - IIF(W > 0, Screen.TwipsPerPixelX * Ratio, 0) _
                , Y + H - IIF(H > 0, Screen.TwipsPerPixelY * Ratio, 0)), ForeColor, BF
        End If
    ElseIf TypeName(Data) = "String" Then
        '����
        If Font Is Nothing Then
            Set Font = New StdFont
            Font.name = "����"
            Font.Size = 9
        End If
        
        '��Ҫ��Set Dev.Font=Font,������byRef
        Dev.Font.name = Font.name
        Dev.Font.Size = Font.Size
        Dev.Font.Bold = Font.Bold
        Dev.Font.Underline = Font.Underline
        Dev.Font.Italic = Font.Italic
        Dev.Font.Strikethrough = Font.Strikethrough
        
        '�����ź���������������,�ж�ʱ��ԭʼ��СΪ׼
        strTemp = Replace(Data, vbCrLf, vbCr)
        strTemp = Replace(strTemp, vbLf, vbCr)
        If H >= Dev.TextHeight(Replace(strTemp, vbCr, "")) Then blnH = True '�߶��Ƿ���(�ӻس�����һ�и߶�)
        
        '����
'����
'        If TypeName(Dev) = "Printer" Then
'            LINE_W = Dev.TwipsPerPixelX * 2 * Ratio '���߼�����(���ʱ��,�ж�ʱ����)
'            intBase = Dev.TwipsPerPixelX
'        Else
            LINE_W = Screen.TwipsPerPixelX * 2 * Ratio '���߼�����(���ʱ��,�ж�ʱ����)
            intBase = Screen.TwipsPerPixelX
'        End If
        X = -Int(-X * Ratio): Y = -Int(-Y * Ratio)
        W = -Int(-W * Ratio): H = -Int(-H * Ratio)
        Dev.Font.Size = Font.Size * Ratio
        sngYDistance = Dev.ScaleY(sngYDistance * Ratio, vbPoints, vbTwips)
        
        '�������
        If Not (BackColor = vbWhite) Then '��ɫ�����ݲ�����,�Ա����ص�����
            Dev.Line (X, Y)-(X + W, Y + H), BackColor, BF
        End If
        
        '����Ƿ���(�ӻس���Ϊ������,�Ա����)
        If W > Dev.TextWidth("A") * LenB(StrConv(strTemp, vbFromUnicode)) + (LINE_W * 2) Then
            blnW = InStr(strTemp, vbCr) = 0
        Else
            blnW = False
        End If
        Dev.ForeColor = ForeColor
        
        '�������(�߿�֮���ٸ�һ��)
        '�����߶ȷ�Χ�����
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
                'ͨ��0000�жϷǵ�Ԫ�񣨱�ǩԪ�أ�
                If Border = "0000" Then
                    intCellBorder = 0
                    intBase = 0
                Else
                    intCellBorder = LINE_W * 2
                End If
                
                If Wrap = False And blnCellWordWrap = False Then
                    '���Զ�����ʱ�����ֲ����
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
                    '�����ȡ����
                    Dev.Print strText
                Else
                    '������ֳɶ���(�ڿ�߷�Χ��)
                    ReDim arrText(0) '�ڴ�,��һ�в����ܳ���
                    For i = 1 To Len(strTemp)
                        If Mid(strTemp, i, 1) = vbCr Then
                            '���г������˳�,���߲��ݲ����
                            If (Dev.TextHeight("��") + sngYDistance) * (UBound(arrText) + 2) - sngYDistance > H Then Exit For
                            ReDim Preserve arrText(UBound(arrText) + 1)
                        ElseIf Dev.TextWidth(arrText(UBound(arrText)) & Mid(strTemp, i, 1)) + intCellBorder + intBase > W Then
                            '���г������˳�,���߲��ݲ����
                            If (Dev.TextHeight("��") + sngYDistance) * (UBound(arrText) + 2) - sngYDistance > H Then Exit For
                            ReDim Preserve arrText(UBound(arrText) + 1)
                        End If
                        '�п���һ��һ���ַ���ȶ�����
                        If Dev.TextWidth(arrText(UBound(arrText)) & Mid(strTemp, i, 1)) - intCellBorder <= W And Mid(strTemp, i, 1) <> vbCr Then
                            arrText(UBound(arrText)) = arrText(UBound(arrText)) & Mid(strTemp, i, 1)
                        End If
                    Next
                    
                    '�����ʼ����
                    Select Case VAlign
                        Case 0
                            Dev.CurrentY = Y + intCellBorder
                        Case 1
                            Dev.CurrentY = Y + (H - (Dev.TextHeight("A") + sngYDistance) * (UBound(arrText) + 1) + sngYDistance) / 2 + LINE_W / 2
                        Case 2
                            Dev.CurrentY = Y + H - LINE_W - (Dev.TextHeight("A") + sngYDistance) * (UBound(arrText) + 1) + sngYDistance
                    End Select
                    
                    '�������
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
    Else 'ͼ��(�߿�֮��)
        LINE_W = 15 * Ratio '���߼�����(���ʱ��,�ж�ʱ����)
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
        '�����߿�
        If Not (BorderColor = vbWhite And TypeName(Data) = "String") Then '��ɫ�߿��ݲ�����,�Ա����ص�����
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
'���ܣ�����ָ��ͼƬ��Ŀ��ߴ磬���غ��ʱ�����ͼƬ����߱�������
'������objPic=Ҫ����ͼƬ
'      lngObjW,lngObjH=Ҫ����Ŀ��ߴ�(Twip)
'      objDraw=������ת�����PictureBox(AutoRedraw=True)
    Dim W As Long, H As Long
    Dim lngW As Long, lngH As Long
    Dim sngW As Single, sngH As Single
    
    objDraw.Cls
    objDraw.BackColor = vbWhite
    objDraw.Width = lngObjW
    objDraw.Height = lngObjH
        
    'ͼƬԭʼ��С(Twip)
    W = objPic.Width * (15 / 26.46)
    H = objPic.Height * (15 / 26.46)
    
    sngW = W / objDraw.ScaleWidth
    sngH = H / objDraw.ScaleHeight
    If sngW > sngH Then
        lngW = W / sngW: lngH = H / sngW
    Else
        lngW = W / sngH: lngH = H / sngH
    End If
    
    '��ͼ������
    objDraw.PaintPicture objPic, 0, 0, lngW, lngH
    'objDraw.PaintPicture objPic, (objDraw.ScaleWidth - lngW) / 2, (objDraw.ScaleHeight - lngH) / 2, lngW, lngH
    
    Set ScalePicture = objDraw.Image
End Function

Public Function GetFieldValue(frmParent As Object, strSource As String, Optional Convert As Boolean) As String
'���ܣ�������Դ��¼���л�ȡָ���ֶε�ԭʼֵ
'������strSource="�ֿƷ���.Ӧ�ս��",Convert=�Ƿ�ת��Ϊ�ɼ����ʽ,��Ҫ��Ը��ϼ���ʱ�����ּ�������
'˵����
'   1.��garrData�л�ȡָ����¼����ǰ��¼λ�õ�ֵ
'   2.����ֶ�����ΪLong Raw��,�򷵻ز�������ʱ�ļ���

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
                If .DataSet.Fields(strField).type = adVarBinary Then    '����:Dbms_Lob.Substr���ص�Raw����
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
                    '��ΪGetChunk����ʹλ��ָ�����,�����ظ���ȡ,����ÿ�ο�¡
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
'���ܣ��Զ���������п�,����С�ʺ�Ϊ׼
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
'���ܣ��ָ���ǰ��ӡ����ԭʼ�趨ֽ��
'˵�����������ӡ����ֽ�Ų�������
    Dim objFmt As RPTFmt
    Dim strTmp As String
    Dim strName As String
    
    Set objFmt = objReport.Fmts("_" & objReport.bytFormat)
    
    If objFmt.ֽ�� = Val("256-�Զ���ֽ��") Then
        If IsWindowsNT Then
            strTmp = GetRegPrinterInfo("PaperForm", objReport.���, objFmt.˵��)
            If Val(strTmp) = 1 Then
                Call SetNTPrinterPaper_Form(lngHWND, objFmt.W / Twip_mm, objFmt.H / Twip_mm, IIF(objFmt.ֽ�� = 0, 1, objFmt.ֽ��), intCopys)
            Else
                Call SetNTPrinterPaper(lngHWND, objFmt.W / Twip_mm, objFmt.H / Twip_mm, IIF(objFmt.ֽ�� = 0, 1, objFmt.ֽ��), intCopys)
            End If
        Else
            Printer.Width = objFmt.W
            Printer.Height = objFmt.H
        End If
    Else
        Printer.PaperSize = objFmt.ֽ��
    End If
    ResetPrinterPaper = True
End Function

Public Function SetPrinterPaper(ByVal lngHWND As Long, objReport As Report, ByVal lngH As Long, ByVal intCopys As Integer) As Boolean
'���ܣ���̬���õ�ǰ��ӡ����ֽ�Ÿ߶�(�Զ���ֽ��)
'˵�����������ӡ����ֽ�Ų�������
    Dim objFmt As RPTFmt
    Dim strTmp As String
    Dim strDefault As String
    
    Set objFmt = objReport.Fmts("_" & objReport.bytFormat)

    SetPrinterPaper = True
    
    If IsWindowsNT Then
        strTmp = GetRegPrinterInfo("PaperForm", objReport.���, objFmt.˵��)
        If Val(strTmp) = 1 Then
            If Not SetNTPrinterPaper_Form(lngHWND, objFmt.W / Twip_mm, lngH / Twip_mm, objFmt.ֽ��, intCopys) Then
                SetPrinterPaper = False
            End If
        Else
            If Not SetNTPrinterPaper(lngHWND, objFmt.W / Twip_mm, lngH / Twip_mm, objFmt.ֽ��, intCopys) Then
                SetPrinterPaper = False
            End If
        End If
    Else
        'ֽ��,��ӡ�������ֲ���
        Printer.Width = objFmt.W
        Printer.Height = lngH
    End If
    
    '���ú�������100Twip,˵������ʧ��
    If Abs(Printer.Height - lngH) > 100 Then SetPrinterPaper = False
End Function

Public Function GetRegPrinterInfo(ByVal strKey As String, ByVal strCode As String, _
    ByVal strFormat As String, Optional ByVal objReport As Object, _
    Optional ByVal bytType As Byte = 1, Optional ByVal strUser As String) As String
'���ܣ���ȡע���Ĵ�ӡ������Ϣ
'������
'  strKey��ע��������
'  strCode��������
'  strFormat�������ʽ
'  bytType��99-�����ࣨFormat��AllFormat����0-��ʽ�ࣨĬ�ϣ���1-��ʽ��ָ������Printer��PaperBin�ȣ�
'���أ�ע������ֵ

    Dim strSec As String, strSecUser As String
    Dim strValue As String
    
    strSec = "˽��ģ��\" & App.ProductName & "\LocalSet\" & strCode
    If strUser = "" Then
        strSecUser = "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\LocalSet\" & strCode
    Else
        strSecUser = "˽��ģ��\" & strUser & "\" & App.ProductName & "\LocalSet\" & strCode
    End If

    Select Case bytType
    Case 1
        strValue = GetSetting("ZLSOFT", strSec & "\" & strFormat, strKey, "")
        If strValue = "" Then strValue = GetSetting("ZLSOFT", strSec & "\���и�ʽ", strKey, "")
        If strValue = "" Then strValue = GetSetting("ZLSOFT", strSec, strKey, "")
        If strValue = "" Then strValue = GetSetting("ZLSOFT", strSecUser, strKey, "")
        
        If Not objReport Is Nothing Then
            If strValue = "" And strKey = "Printer" Then strValue = GetSetting("ZLSOFT", strSecUser, strKey, objReport.��ӡ��)
        End If
    Case 99
        strValue = GetSetting("ZLSOFT", strSec, strKey, "")
        If strValue = "" Then strValue = GetSetting("ZLSOFT", strSecUser, strKey, "")
    End Select

    GetRegPrinterInfo = strValue
End Function

Public Function InitPrinter(frmParent As Object, Optional ByVal intCopies As Integer = 1) As Boolean
'���ܣ�����ע���frmParent.mobjReport���ݳ�ʼ����ӡ������(����->������->��ǰ)
'������intCopies=�������õ�Ҫ��ӡ�ķ���
'���أ�����޴�ӡ����ֽ�Ų���,��ʧ��
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
    
    '�������
    Set objReport = frmParent.mobjReport
    Set objFmt = objReport.Fmts("_" & objReport.bytFormat)
    
    '�������ֻ��һ����ӡ��,Ĭ��Ϊ��
    If Printers.count = 1 Then
        strPrinter = Printers(0).DeviceName
    Else
        '��������
        strPrinter = GetRegPrinterInfo("Printer", objReport.���, objFmt.˵��, objReport)
    End If

    If strPrinter = "" Then
        If MsgBox("""" & objReport.���� & """û�����ô�ӡ��,���ھ����ñ��ش�ӡ����", _
            vbQuestion + vbYesNo + vbDefaultButton1, App.Title) = vbNo Then Exit Function
        If Not ReportLocalSet(objReport.ϵͳ, objReport.���, False, objReport.bytFormat, frmMain) Then Exit Function
        strPrinter = GetRegPrinterInfo("Printer", objReport.���, objFmt.˵��, objReport)
    End If
    If Printer.DeviceName <> strPrinter Then
        For i = 0 To Printers.count - 1
            If Printers(i).DeviceName = strPrinter Then Set Printer = Printers(i): Exit For
        Next
        If i > Printers.count - 1 Then
            If MsgBox("""" & objReport.���� & """�Ĵ�ӡ��""" & strPrinter & """" & _
                vbCrLf & "��ϵͳ��û�а�װ,Ҫ���ñ��ش�ӡ����", _
                vbQuestion + vbYesNo + vbDefaultButton1, App.Title) = vbNo Then Exit Function
            If Not ReportLocalSet(objReport.ϵͳ, objReport.���, False, objReport.bytFormat, frmMain) Then Exit Function
            strPrinter = GetRegPrinterInfo("Printer", objReport.���, objFmt.˵��, objReport)
        End If
    End If
    If Printer.DeviceName <> strPrinter Then
        For i = 0 To Printers.count - 1
            If Printers(i).DeviceName = strPrinter Then Set Printer = Printers(i): Exit For
        Next
    End If
    InitPrinter = True
    
    '1.�Ȱ����ù̶����г�ʼ��
    On Error Resume Next
    
     '��ֽ��ʽ
    strTmp = GetRegPrinterInfo("PaperBin", objReport.���, objFmt.˵��)
    intPaperBin = Val(strTmp)
    If intPaperBin = 0 Then intPaperBin = 15
    If Printer.PaperBin <> intPaperBin Then
        Printer.PaperBin = intPaperBin
    End If
    
    'ֽ��
    If objFmt.ֽ�� < Val("256-�Զ���ֽ��") Then
        Printer.PaperSize = objFmt.ֽ��
    End If
    
    '����
    If Printer.Copies <> intCopies Then
        Err.Clear: Printer.Copies = intCopies
        If Err.Number <> 0 Then
            Err.Clear: Printer.Copies = 1
        End If
    End If
    
    '2.NT�����£���API���豸��������
    If objFmt.ֽ�� = Val("256-�Զ���ֽ��") Then
        If IsWindowsNT Then
            strTmp = GetRegPrinterInfo("PaperForm", objReport.���, objFmt.˵��)
            If Val(strTmp) = 1 Then
                strFormName = GetRegPrinterInfo("PaperFormName", objReport.���, objFmt.˵��)
                If Not SetNTPrinterPaper_Form(frmMain.hwnd, objFmt.W / Twip_mm, objFmt.H / Twip_mm, IIF(objFmt.ֽ�� = 0, 1, objFmt.ֽ��), intCopies, , strFormName, Printer) Then
                    InitPrinter = False
                End If
            Else
                If Not SetNTPrinterPaper(frmMain.hwnd, objFmt.W / Twip_mm, objFmt.H / Twip_mm, IIF(objFmt.ֽ�� = 0, 1, objFmt.ֽ��), intCopies) Then
                    InitPrinter = False
                End If
            End If
        End If
    Else
        '���Զ���ֽ�ŵ�ֽ��
        intOrient = IIF(objFmt.ֽ�� = 0, 1, objFmt.ֽ��)
        Printer.Orientation = intOrient
    End If
            
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Function

Private Function GetHscAlign(ByVal intCellAlign As Integer, ByVal strText As String _
    , Optional ByVal objBody As Object _
    , Optional ByVal lngCol As Long = 0) As Byte
'���ܣ����ݱ��Ԫ�Ķ�������,���ش�ӡˮƽ��������

    Dim intColAlign As Integer

    If Not objBody Is Nothing Then
        intColAlign = objBody.ColAlignment(lngCol)
        If intCellAlign < 0 Or intCellAlign > 8 Then
            '�����������á��Ķ���δ���ã���Ԫ�񣩣���ʹ����ƽ���Ķ��루�У�
            intCellAlign = intColAlign
        End If
    End If

    Select Case intCellAlign
    Case 0, 1, 2
        GetHscAlign = 0 '��
    Case 3, 4, 5
        GetHscAlign = 1 '��
    Case 6, 7, 8
        GetHscAlign = 2 '��
    Case Else
        If IsNumeric(strText) Then
            GetHscAlign = 2 '��
        Else
            GetHscAlign = 0 '��
        End If
    End Select
End Function

Private Function GetVscAlign(intAlign As Integer) As Byte
'���ܣ����ݱ��Ԫ�Ķ�������,���ش�ӡ��ֱ��������
    Select Case intAlign
        Case 0, 3, 6
            GetVscAlign = 0 '��
        Case 1, 4, 7
            GetVscAlign = 1 '��
        Case 2, 5, 8
            GetVscAlign = 2 '��
        Case Else
            GetVscAlign = 1 '��
    End Select
End Function

Private Sub SearchCell(ByVal objGrid As Control, ByVal Row As Long, ByVal Col As Long, ByVal MaxR As Long, ByVal MaxC As Long, _
    W As Long, H As Long, strSkip As String, strSkip2 As String)
'���ܣ���������һ����Ԫ�����ȷ���
'������MaxR,MaxC=�������������Χ
'���أ�W,H=�õ�Ԫ��Ŀ��(�����ϲ���Ԫ),strSkip=�õ�Ԫ�����ϲ��ĵ�Ԫ��,��Щ��Ԫ�����ٴ���
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
        '������������ϲ���Ԫ
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
        
        '������������ϲ���Ԫ
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
    
    '�����Ԫ����ͬʱ�ϲ�
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
'���º������ڷ�����������ԴȨ��------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
Public Function UserObject(Optional ByVal intConnect As Integer = 0 _
    , Optional ByVal blnIsBusinessTable As Boolean) As ADODB.Recordset
'���ܣ���ȡ��ǰ�û�������Select Ȩ�޵����б���ͼ��(�����û�������󼰱���Ȩ����)
'���أ��ɹ�=���������б�(����Ӣ˳������),ʧ��=��
'˵���������������������û�����,��ϵͳ���������ѯ
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
'���ܣ�SQLObject�������Ӻ���,����ȥ���������е������ַ�
    Dim i As Integer
    'Ѱ�ҵ�һ�������ַ�λ��
    For i = 1 To Len(strObject)
        If InStr(Chr(32) & Chr(13) & Chr(10) & Chr(9), Mid(strObject, i, 1)) = 0 Then Exit For
    Next
    strObject = Mid(strObject, i)
    'Ѱ�Һ����һ���������ַ�
    For i = 1 To Len(strObject)
        If InStr(Chr(32) & Chr(13) & Chr(10) & Chr(9), Mid(strObject, i, 1)) > 0 Then Exit For
    Next
    If i <= Len(strObject) Then strObject = Left(strObject, i - 1)
    TrueObject = strObject
End Function

Private Function GetWithAsTables(ByVal strSQL As String) As String
'���ܣ���ȡWith as ֮��ı��������Զ��ŷָ�
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
            '11G R2֧�֣����磺with T��column alias 1,column alias 2,......��
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
'���ܣ�����SQL������õ��Ķ�����
'������strSQL=Ҫ������ԭʼSQL���
'���أ�SQL��������ʵ��Ķ�����,��"���ű�,���˷��ü�¼,ZLHIS.��Ա��"
'˵����1.��Oracle SELECT������
'      2.���SQL����еĶ�����ǰ����������ǰ׺,���ǰ׺���ᱻ��ȡ
'      3.��Ҫ����TrimChar;TrueObject��֧��
    Dim intB As Long, intE As Long, intL As Long, intR As Long
    Dim strAnal As String, strSub As String, strObject As String
    Dim arrFrom() As String, strCur As String, strMulti As String, strTrue As String
    Dim i As Long, j As Long
    Dim lngTmp As Long
    Dim strTmp As String, strObjectSub As String
    
    On Error GoTo errH
    
    '��д����ȥ��������ַ�
    strAnal = UCase(TrimChar(strSQL))
    If strWithas = "" Then
        strWithas = GetWithAsTables(strAnal)
    End If
    
    If InStr(strAnal, "SELECT") = 0 Or InStr(strAnal, "FROM") = 0 Then Exit Function
    If mdlPublic.TransSpecialChar(strAnal) = False Then Exit Function
    
    '�ȷֽ⴦��Ƕ���Ӳ�ѯ
    Do While InStr(strAnal, "(") > 0
        intB = InStr(strAnal, "("): intE = intB 'ƥ�����������λ��
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
                    '���ڷ��Ӳ�ѯ,�����Ż�����������,��ʹѭ������
                    strAnal = Left(strAnal, intB - 1) & "@" & Mid(strAnal, intB + 1)
                    strAnal = Left(strAnal, intE - 1) & "@" & Mid(strAnal, intE + 1)
                ElseIf InStr(Mid(strAnal, intB + 1, intE - intB - 1), "SELECT") > 0 _
                    And InStr(Mid(strAnal, intB + 1, intE - intB - 1), "FROM") > 0 Then
                    '�Ӳ�ѯ���
                    strSub = Mid(strAnal, intB + 1, intE - intB - 1)
                    '�����Ӳ�ѯ������ΪΪ���������
                    strAnal = Replace(strAnal, Mid(strAnal, intB, intE - intB + 1), "Ƕ�ײ�ѯ")
                    '�ݹ����
                    strObjectSub = SQLObject(strSub, strWithas)
                    If InStr(strObject & "," & strWithas & ",", "," & strObjectSub & ",") = 0 Then
                        strObject = strObject & "," & strObjectSub
                    End If
                ElseIf strTmp = "" And lngTmp <> 0 Then
                    'ȥ��Table��̬�ڴ��
                    strAnal = Replace(strAnal, Mid(strAnal, lngTmp + 1, intE - lngTmp + 1 + 1), "��̬�ڴ��")
                Else
                    strAnal = Left(strAnal, intB - 1) & "@" & Mid(strAnal, intB + 1)
                    strAnal = Left(strAnal, intE - 1) & "@" & Mid(strAnal, intE + 1)
                End If
                Exit For
            End If
        Next
        '��ƥ��������
        If intE = intB Then strAnal = Left(strAnal, intB - 1) & "@" & Mid(strAnal, intB + 1)
    Loop
    
    '�ֽ����(��ʱstrAnalΪ�򵥲�ѯ,���ܴ�Union������)
    arrFrom = Split(strAnal, "FROM")
    For i = 1 To UBound(arrFrom) '�ӵ�һ��From���沿�ݿ�ʼ
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
            If InStr(strObject & "," & strWithas & ",", "," & strTrue & ",") = 0 And strTrue <> "Ƕ�ײ�ѯ" And strTrue <> "��̬�ڴ��" Then
                If InStr(strTrue, "'") = 0 And InStr(strTrue, "@") = 0 Then
                    strObject = strObject & "," & strTrue
                End If
            End If
        Next
    Next
    '���
    SQLObject = Mid(strObject, 2)
    SQLObject = Replace(SQLObject, ",,", ",")
    Exit Function
errH:
    Err.Clear
End Function

Public Function CheckReportPriv(lngRPTID As Long, Optional ByVal blnReportGroup As Boolean) As Boolean
'���ܣ���鵱ǰ�û���ĳ�ű���(�Ѵ���)�Ƿ���ȫ��Ȩ�޷���
'������lngRPTID=����ID
'���أ���ȫ="",����ȫ=���ܷ��ʵĶ�����,��"ZLPER.���ű�,ZLHIS.���˷��ü�¼"
'˵���������ڱ���������򿪻���Ʊ���ʱ���Ȩ��
'�ο���grsObject
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer, j As Integer
    Dim strOwner As String, strName As String
    Dim strSQL As String
    
    On Error GoTo errH
    
    rsTmp.CursorLocation = adUseClient
    If Not blnReportGroup Then
        strSQL = "Select ����,���� From zlRPTDatas Where ����ID=[1] And Nvl(�������ӱ��, 0) = 0 "
    Else
        strSQL = "Select A.����,A.���� " & vbCr & _
                 "From zlRPTDatas A, zlRPTSubs B " & vbCr & _
                 "Where A.����ID=B.����ID And Nvl(a.�������ӱ��, 0) = 0 And B.��ID=[1] "
    End If
    Set rsTmp = OpenSQLRecord(strSQL, "CheckReportPriv", lngRPTID)
    For i = 1 To rsTmp.RecordCount '��������ݷ���
        If Not IsNull(rsTmp!����) Then
            For j = 0 To UBound(Split(rsTmp!����, ","))
                strOwner = Split(Split(rsTmp!����, ",")(j), ".")(0)
                strName = Split(Split(rsTmp!����, ",")(j), ".")(1)
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
'���ܣ���鵱ǰ�û���ָ�������Ƿ���ȫ��Ȩ�޷���
'������strObject=��������,��"���ű�,���˷��ü�¼"
'���أ���ȫ=��,����ȫ=���ܷ��ʵĶ�����,��"���ű�,���˷��ü�¼"
'˵����������У������Դ֮ǰ����Ƿ���Ȩ�޲�ѯSQL����еĶ���
'�ο���grsObject
    Dim i As Integer
    Dim arrObject As Variant
    
    arrObject = Split(strObject, ",")
    For i = 0 To UBound(arrObject)
        If arrObject(i) <> "DUAL" Then
            If InStr(arrObject(i), ".") = 0 Then
                grsObject.Filter = "OBJECT_NAME='" & arrObject(i) & "'"
            Else
                '�������ͼ���������ǰ׺,����������߶���Ȩ��
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
'���ܣ����ݶ��������ϵ�ǰ�û����ܷ��ʵ�������ǰ׺(������ͬһ�������ж��������Ҫ��ѡ����֮һ)
'������strObject=��������,��"���ű�,���˷��ü�¼"
'���أ�����=����������ǰ׺�Ķ���,��"ZLPER.���ű�,ZLHIS.���˷��ü�¼",ȡ��="ȡ��"
'�ο���grsObject
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
                '�������ͼ���������ǰ׺,��ʹ���䱾����
                If InStr(ObjectOwner, "," & arrObject(i)) = 0 Then
                    ObjectOwner = ObjectOwner & "," & arrObject(i)
                End If
            Else
                If intConnect > Val("0-��ǰ��¼����") Then
                    '������������
                    strOtherConnectOwner = mdlPublic.GetDBConnectInfo(intConnect, Val("1-�û���"))
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
                        '�������������֮�⣬ֻʣһ�����������ߣ���ֱ��Ϊ����������
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
                            " Select Column_Value As ������ From Table(Cast(zlTools.f_Str2List ('" & Replace(strOwner, "'", "") & "') as zlTools.t_StrList))" & _
                            " Minus" & _
                            " Select ������ From zlBakSpaces Where ������ IN(" & strOwner & ")"
                        strSQL = _
                            "Select A.������,Decode(B.������,Null,0,1) as ϵͳ�� " & _
                            "From (" & strSQL & ") A,(Select Distinct ������ From zlSystems) B Where A.������=B.������(+)"
                        Set rsTmp = OpenSQLRecord(strSQL, "ObjectOwner")
                        If rsTmp.RecordCount = 1 Then
                            If rsTmp!ϵͳ�� = 1 Then
                                strOwner = rsTmp!������
                                blnNoSel = True
                            End If
                        End If
                        On Error GoTo 0
                        
                        If blnNoSel Then
                            If InStr(ObjectOwner & ",", "," & strOwner & "." & arrObject(i) & ",") = 0 Then
                                ObjectOwner = ObjectOwner & "," & strOwner & "." & arrObject(i)
                            End If
                        Else
                            'ͬһ�����ж��������,��Ҫ��ѡ��
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
                                    '�޷��ж��Ƿ��Ѿ�ж�أ�Unload��ʼ�ղ�ΪNothing
                                    Unload frmSelOwner
                                End If
                            Else
                                'ȡ��ѡ��,Ҳ����ȡ������(���ó���),���ؿ�
                                ObjectOwner = "ȡ��": Exit Function
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
'���ܣ���SQL����滻�ɴ����������ߵ���ʽ
'������strSQL=ԭʼSQL���,strOwner=���������ߴ�,��"ZLPER.���ű�,ZLHIS.���˷��ü�¼"
'���أ����ʶ������������ǰ׺��SQL���
'˵����1.����������ֱ��ִ���û�SQL���,������Ҫ��Ȩ�����˽��ͬ��ʡ�
'      2.�Ա������ֶ�����ͬ���ֶ���û�д������,������
    Dim i As Long, j As Long
    Dim intLoc As Long, blnDo As Boolean
    
    '�����ֻ�ÿո���
    strSQL = SpaceSQL(strSQL)
    
    For i = 0 To UBound(Split(strOwner, ","))
        '����ѭ��ȷ�Ϸ�ʽ,ȷ���滻���Ǳ���,������������䲿�ݻ򱻰��������������еĲ���
        j = 0 '��ǰ��ʼ����λ��
        Do
            j = j + 1
            intLoc = InStr(j, strSQL, Split(Split(strOwner, ",")(i), ".")(1))
            If intLoc > 12 Then '������"SELECT FROM "
                '�������������ǰ׺�Ĳ��滻
                blnDo = True
                '�ұ��Կո�","�š������Ž���
                blnDo = blnDo And (InStr(",) ", Mid$(strSQL, intLoc + Len(Split(Split(strOwner, ",")(i), ".")(1)), 1)) > 0)
                '�����Ϊ","�Ż�"FROM "
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
'���ܣ���SQL���任ΪֻΪ�ո�������ʽ,�Ա��ڷ���
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
'���ܣ������������,��ֹ��Set��ɵ�ַ�ķ���
    Dim objItem As RPTItem, objData As RPTData
    Dim objPar As RPTPar, objPars As RPTPars
    Dim i As Integer
    
    Set objO = New Report
    
    objO.ϵͳ = objS.ϵͳ
    objO.��� = objS.���
    objO.���� = objS.����
    objO.˵�� = objS.˵��
    objO.��ӡ�� = objS.��ӡ��
    objO.��ֽ = objS.��ֽ
    objO.Ʊ�� = objS.Ʊ��
    objO.��ӡ��ʽ = objS.��ӡ��ʽ
    objO.��ֹ��ʼʱ�� = objS.��ֹ��ʼʱ��
    objO.��ֹ����ʱ�� = objS.��ֹ����ʱ��
    
    objO.blnLoad = objS.blnLoad
    objO.bytFormat = objS.bytFormat
    objO.intGridCount = objS.intGridCount
    objO.intGridID = objS.intGridID
    
    For i = 1 To objS.Fmts.count
        With objS.Fmts(i)
            objO.Fmts.Add .���, .˵��, .W, .H, .ֽ��, .ֽ��, .��ֽ̬��, .ͼ��, "_" & .���
        End With
    Next
    
    For Each objItem In objS.Items
        With objItem
            objO.Items.Add .id, .��ʽ��, .����, .�ϼ�ID, .����, .���, .����, .����, .����, .��ͷ, .X, .Y, .W, .H _
                , .�и�, .����, .�Ե�, .����, .�ֺ�, .����, .����, .б��, .����, .ǰ��, .����, .�߿� _
                , IIF(.���� < 1 And .���� <> 6, 1, .����), .����, .��ʽ, .����, .����߼Ӵ�, .����Ӧ�и�, .ͼƬ _
                , .ϵͳ, .��ID, .SubIDs, .CopyIDs, "_" & .id, .����Դ, .���¼��, .���Ҽ��, .Դ�к�, .������� _
                , .�������, .Relations, .ColProtertys, .ˮƽ��ת
        End With
    Next
    For Each objData In objS.Datas
        With objData
            Set objPars = New RPTPars
            For Each objPar In .Pars
                objPars.Add objPar.����, objPar.���, objPar.����, objPar.����, objPar.ȱʡֵ, objPar.��ʽ, objPar.ֵ�б� _
                    , objPar.����SQL, objPar.��ϸSQL, objPar.�����ֶ�, objPar.��ϸ�ֶ�, objPar.����, "_" & objPar.��� _
                    , objPar.Reserve, objPar.�Ƿ�����
            Next
            objO.Datas.Add .����, .�������ӱ��, .SQL, .�ֶ�, .����, .����, .˵��, objPars, "_" & .����
        End With
    Next
End Sub

Public Function IncStr(ByVal strVal As String) As String
'���ܣ���һ���ַ����Զ���1��
'˵����ÿһλ��λʱ,���������,��ʮ���ƴ���,����26���ƴ���
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
'���ܣ���ȡ��һ��������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, blnExist As Boolean
    Const strGroup As String = "GROUP"
    Const strReport As String = "REPORT"
    
    On Error GoTo errH
    
    If Not blnGroup Then
        strSQL = "Select Max(���) as ��� From zlReports Where ��� Like [1]"
    Else
        strSQL = "Select Max(���) as ��� From zlRPTGroups Where ��� Like [2]"
    End If
    Set rsTmp = OpenSQLRecord(strSQL, "GetNextNO", "REPORT%", "GROUP%")
    If Not rsTmp.EOF Then
        If IsNull(rsTmp!���) Then
            GetNextNO = IIF(blnGroup, strGroup, strReport) & "_001"
        Else
            GetNextNO = IncStr(rsTmp!���)
        End If
    Else
        GetNextNO = IIF(blnGroup, strGroup, strReport) & "_001"
    End If
    
    Do
        blnExist = False
        blnExist = blnExist Or CheckExist("zlReports", "���", GetNextNO)
        If Not blnExist Then blnExist = CheckExist("zlRPTGroups", "���", GetNextNO)
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
'���ܣ��жϵ�ǰ���г����Ƿ���VB�Ĺ��̻�����
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
'���ܣ���ȡ��ǰ��¼���ݿ��û���
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
'���ܣ���ȡָ���û�������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
        
    On Error GoTo errH
        
    If gcnOracle Is Nothing Then Exit Function
    If gcnOracle.State = adStateClosed Then Exit Function
    If strUser = "" Then Exit Function
    strSQL = " Select A.����,A.���" & _
        " From ��Ա�� A,�ϻ���Ա�� B" & _
        " Where A.ID=B.��ԱID And B.�û���='" & strUser & "'"
    Call OpenRecord(rsTmp, strSQL, "GetTheUserName")
    If Not rsTmp.EOF Then GetTheUserName = rsTmp!���� & ""
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub AutoSizeCol(lvw As Object)
'���ܣ������Զ�ListView��ǰ�����Զ��������п��
'������blnByHead=�Ƿ���ͷ�ı�����,Col=ָ���л���������(1-N)
    Dim i As Integer, lngW As Long
    For i = 1 To lvw.ColumnHeaders.count
        SendMessage lvw.hwnd, LVM_SETCOLUMNWIDTH, i - 1, LVSCW_AUTOSIZE
        If lvw.ColumnHeaders(i).Width < 200 Then lvw.ColumnHeaders(i).Width = 0
        If lvw.ColumnHeaders(i).Width < (TLen(lvw.ColumnHeaders(i).Text) + 2) * 90 And lvw.ColumnHeaders(i).Width <> 0 Then lvw.ColumnHeaders(i).Width = (TLen(lvw.ColumnHeaders(i).Text) + 2) * 90
    Next
End Sub

Public Function GetExpField(objFld As ADODB.Field, Optional ByVal blnDataNum As Boolean) As String
'���ܣ���������ʱ��
'������blnDataNum=true ԴID��ʵ��ֵ����
    Dim strTmp As String
    
    If IsNull(objFld.Value) Then
        Exit Function
    ElseIf InStr(",ϵͳ,����ID,����,����ʱ��,", "," & objFld.name & ",") > 0 Then
        Exit Function
    ElseIf objFld.name = "���" Then
        GetExpField = "[���]" '����ʱȡ��ǰʱ��
    ElseIf objFld.name = "�޸�ʱ��" Then
        GetExpField = "Sysdate" '����ʱȡ��ǰʱ��
    ElseIf objFld.name = "ID" Then
        GetExpField = "[NextVal]" '����ʱȡ"��ǰ��_ID.NextVal"
    ElseIf objFld.name = "�ϼ�ID" Then
        GetExpField = "[CurrVal-X]" '����ʱȡ"��ǰ��_ID.CurrVal-X",XΪ�ϼ�ID��Ϊ�յĿ�ʼ��
    ElseIf objFld.name = "����ID" Then
        GetExpField = "[zlReports_ID.CurrVal]" '����ʱȡ"zlReports_ID.CurrVal"
    ElseIf objFld.name = "ԴID" And blnDataNum = False Then
        GetExpField = "[zlRPTDatas_ID.CurrVal]" '����ʱȡ"zlRPTDatas_ID.CurrVal"
    ElseIf objFld.name = "Ԫ��ID" Then
        GetExpField = "[zlRPTItems_ID.CurrVal]" '����ʱȡ"zlRPTDatas_ID.CurrVal"
    ElseIf objFld.name = "����" Then
        GetExpField = Replace(UCase(objFld.Value), UCase(gstrDBUser) & ".", "USER.")
    Else '����ʱ������������ת��ȡֵ
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
                '��ʱ��֧��ͼƬ�Ĵ���
        End Select
    End If
End Function

Private Function GetFieldNames(rsTmp As ADODB.Recordset) As String
'���ܣ�����һ����¼�������е��ֶ����ƴ�
    Dim i As Integer
    For i = 0 To rsTmp.Fields.count - 1
        GetFieldNames = GetFieldNames & "," & rsTmp.Fields(i).name
    Next
    GetFieldNames = GetFieldNames & ","
End Function

Public Function ExportReport(lngRPTID As Long, strFile As String) As Boolean
'���ܣ�����һ���Զ��屨��
'������lngRPTID=����ID
'      strFile=�ļ���
'���أ������Ƿ�ɹ���
'˵����
'      1.�����ѷ����ı���,������Ϊ�Ƿ�������
'      2.Ŀǰ��֧��ͼƬԪ�����ݵĵ���
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
    
    strSQL = "Select ID,���,����,˵��,����,��ӡ��,��ֽ,Ʊ��,��ӡ��ʽ,ϵͳ,����ID,����,�޸�ʱ��,����ʱ��,��ֹ��ʼʱ��,��ֹ����ʱ�� From zlReports Where ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "ExportReport", lngRPTID)
    If rsTmp.EOF Then
        MsgBox "û�з���ָ����������ݣ�", vbInformation, App.Title
        Exit Function
    End If
    
    '�򿪴����ļ�
    Set objFile = New FileSystemObject
    If objFile.FileExists(strFile) Then Call objFile.DeleteFile(strFile, True)
    Set objText = objFile.CreateTextFile(strFile, True)
    blnOpen = True
    
    '���������ͷ
    Call objText.WriteLine("[HEAD]")
    Call objText.WriteLine("������=" & rsTmp!���)
    Call objText.WriteLine("��������=" & rsTmp!����)
    Call objText.WriteLine("����˵��=" & IIF(IsNull(rsTmp!˵��), "", rsTmp!˵��))
    Call objText.WriteLine("�����û�=" & gstrDBUser)
    Call objText.WriteLine("����ʱ��=" & Format(Currentdate, "yyyy-MM-dd HH:mm:ss"))
    Call objText.WriteLine("��ֹ��ʼʱ��=" & Format(rsTmp!��ֹ��ʼʱ�� & "", "HH:mm:ss"))
    Call objText.WriteLine("��ֹ����ʱ��=" & Format(rsTmp!��ֹ����ʱ�� & "", "HH:mm:ss"))
    
    '����:ZLReport,�Էֺ�Ϊ�н������Էֺ�Ϊһ���ֶν���,���ֺ�Ϊһ����¼����
    Call objText.WriteLine("[ZLREPORTS]")
    Call objText.WriteLine(";")
    For Each objFld In rsTmp.Fields
        Call objText.WriteLine(objFld.name & "=" & GetExpField(objFld) & ";")
    Next
    
    '�����ʽ
    'Set rsTmp = New ADODB.Recordset
    strSQL = "Select ����ID,���,˵��,W,H,ֽ��,ֽ��,��ֽ̬��,ͼ�� From zlRPTFmts Where ����ID=[1]"
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
    
    '����Ԫ��
    strSQL = "Select ϵͳ,ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,���� " & vbCrLf & _
             "    ,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ID as ԭID,��ID,ԴID,���¼��,���Ҽ��" & vbCrLf & _
             "    ,Դ�к�,�������,�������,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת " & vbCrLf & _
             "From zlRPTItems " & vbCrLf & _
             "Where ����ID=[1] " & vbCrLf & _
             "Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID"
    Set rsTmp = OpenSQLRecord(strSQL, "ExportReport", lngRPTID)
    
    strSQL = "Select ����ID,Ԫ��ID,��������,�����ֶ�,������ϵ,����ֵ,������ɫ,������ɫ,�Ƿ�Ӵ�,�Ƿ�����Ӧ��,���� " & vbCrLf & _
             "From zlRPTColProterty " & vbCrLf & _
             "Where ����ID=[1]"
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
            rsSub.Filter = "Ԫ��ID=" & rsTmp!id
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
    
    '��������,'���ݲ���
    strSQL = "Select ID,����ID,�������ӱ��,����,�ֶ�,����,����,˵��,ID as ԭID From zlRPTDatas Where ����ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "ExportReport", lngRPTID)
    
    strSQL = "Select B.ԴID,B.�к�,B.���� From zlRPTDatas A,zlRPTSQLs B Where A.ID=B.ԴID And A.����ID=[1]"
    Set rsSQL = OpenSQLRecord(strSQL, "ExportReport", lngRPTID)
    
    strSQL = "Select B.ԴID,B.����,B.���,B.����,B.����,B.ȱʡֵ,B.��ʽ,B.ֵ�б�,B.����SQL,B.��ϸSQL,B.�����ֶ�" & vbCrLf & _
             "    ,B.��ϸ�ֶ�,B.����,B.���� " & vbCrLf & _
             "From zlRPTDatas A,zlRPTPars B " & vbCrLf & _
             "Where A.ID=B.ԴID And A.����ID=[1]"
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
            
            rsSQL.Filter = "ԴID=" & rsTmp!id
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
           
            rsSub.Filter = "ԴID=" & rsTmp!id
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
'����:���ļ�����һ�ű���,����Ǹ��ǹ̶�����Ҫ���´������Ȩ��
'����:strFile=�ⲿ�ļ���
'     lngCurrID=�������븲�ǵ�ָ��ID�����б���
'     blnOnlyData=�Ƿ�ֻ��������Դ
'     lngGroupID=������ID,0=�������б�����,<>0=���뵽�ñ�������
'     lngClassID=������ID
'����:�ɹ�="ID|���|����|˵��",ʧ��=""
'˵����1.���빲����ʱ�������ظ�,���Զ�ȡ
'      2.�������б���ʱ,��ǰ������Ϣ����,����ֽ����Ϣ
    Dim objFile As FileSystemObject, objText As TextStream
    Dim rsReport As New ADODB.Recordset
    Dim rsFMT As New ADODB.Recordset
    Dim rsItem As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim rsPar As New ADODB.Recordset
    Dim rsCol As New ADODB.Recordset
    Dim rsProgram As New ADODB.Recordset '�����ж��Ƿ���ڸ�ģ��
    Dim rsCopy As ADODB.Recordset
    
    Dim blnTran As Boolean, blnOpen As Boolean, lngUPID As Long
    Dim strLine As String, strSect As String, strFld As String, strValue As String
    Dim blnReport  As Boolean, blnFmt As Boolean, blnItem As Boolean, blnData As Boolean, blnPar As Boolean, blnSQL As Boolean
    Dim strReport As String, StrFmt As String, strItem As String, strData As String, StrPar As String, strRSQL As String
    Dim strPreNum As String, strNum As String, strName As String, strNote As String, lngRPTID As Long
    
    Dim rsCurr As New ADODB.Recordset
    Dim rsPriv As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim lngRptW As Long, lngRptH As Long, bln��ֽ̬�� As Boolean
    Dim intֽ�� As Integer, intֽ�� As Integer
    Dim strObject As String, strSQL As String, i As Long
    Dim Col As New Collection
    Dim ColData As New Collection
    Dim rsItemCopy As New Recordset
    Dim lng��� As Long
    Dim str��ֹ��ʼʱ�� As String, str��ֹ����ʱ�� As String
    Dim strCol As String, strTmp As String
    Dim blnCol As Boolean
    
    On Error GoTo errH
    
    If lngCurrID = 0 Then blnOnlyData = False
    
    '��ǰ�ı�����Ϣ
    If lngCurrID <> 0 Then
        strSQL = _
            "Select ID, ���, ����, ˵��, ����, ��ӡ��, ��ֽ, Ʊ��, ��ӡ��ʽ, ϵͳ, ����id, ����, �޸�ʱ��" & vbNewLine & _
            "    , ����ʱ��, ��ֹ��ʼʱ��, ��ֹ����ʱ�� " & vbNewLine & _
            "From zlReports " & vbNewLine & _
            "Where ID = [1] "
        Set rsCurr = OpenSQLRecord(strSQL, "ExportReport", lngCurrID)
        If rsCurr.EOF Then Exit Function
    End If
    
    '�򿪱����ļ�
    Set objFile = New FileSystemObject
    If Not objFile.FileExists(strFile) Then Exit Function
    Set objText = objFile.OpenTextFile(strFile)
    blnOpen = True
    
    '�������ݼ�¼��
    If lngCurrID = 0 Then
        rsReport.CursorLocation = adUseClient
        rsReport.Open _
                "Select ID,���,����,˵��,����,��ӡ��,��ֽ,Ʊ��,��ӡ��ʽ,ϵͳ,����ID,����" & vbNewLine & _
                "    ,�޸�ʱ��,����ʱ��,��ֹ��ʼʱ��,��ֹ����ʱ��,����ID " & vbNewLine & _
                "From zlReports " & vbNewLine & _
                "Where Rownum<1" _
            , gcnOracle, adOpenKeyset, adLockOptimistic
        strReport = GetFieldNames(rsReport)
    End If
    
    If Not blnOnlyData Then
        rsFMT.CursorLocation = adUseClient
        rsFMT.Open "Select ����ID,���,˵��,W,H,ֽ��,ֽ��,��ֽ̬��,ͼ�� From zlRPTFmts Where Rownum<1", gcnOracle, adOpenKeyset, adLockOptimistic
        StrFmt = GetFieldNames(rsFMT)
        
        rsItem.CursorLocation = adUseClient
        rsItem.Open "Select ϵͳ,ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,��ID,ԴID,���¼��,���Ҽ��,Դ�к�,�������,�������,����߼Ӵ�,����Ӧ�и�,ˮƽ��ת From zlRPTItems Where Rownum<1", gcnOracle, adOpenKeyset, adLockOptimistic
        strItem = GetFieldNames(rsItem)
    End If
    
    rsData.CursorLocation = adUseClient
    rsData.Open "Select ID,����ID,�������ӱ��,����,�ֶ�,����,����,˵�� From zlRPTDatas Where Rownum<1", gcnOracle, adOpenKeyset, adLockOptimistic
    strData = GetFieldNames(rsData)
    
    rsSQL.CursorLocation = adUseClient
    rsSQL.Open "Select ԴID,�к�,���� From zlRPTSQLs Where Rownum<1", gcnOracle, adOpenKeyset, adLockOptimistic
    strRSQL = GetFieldNames(rsSQL)
    
    rsPar.CursorLocation = adUseClient
    rsPar.Open "Select ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����,���� From zlRPTPars Where Rownum<1", gcnOracle, adOpenKeyset, adLockOptimistic
    StrPar = GetFieldNames(rsPar)
    
    rsCol.CursorLocation = adUseClient
    rsCol.Open "Select ����ID,Ԫ��ID,��������,�����ֶ�,������ϵ,����ֵ,������ɫ,������ɫ,�Ƿ�Ӵ�,�Ƿ�����Ӧ��,���� From zlRPTColProterty Where  Rownum<1", gcnOracle, adOpenKeyset, adLockOptimistic
    strCol = GetFieldNames(rsCol)
    
    rsItemCopy.Fields.Append "ID", adBigInt, , adFldIsNullable
    rsItemCopy.Fields.Append "��ID", adBigInt, , adFldIsNullable
    rsItemCopy.Fields.Append "ԴID", adBigInt, , adFldIsNullable
    rsItemCopy.CursorLocation = adUseClient
    rsItemCopy.CursorType = adOpenStatic
    rsItemCopy.LockType = adLockOptimistic
    rsItemCopy.Open
    gcnOracle.BeginTrans
    blnTran = True
            
    '���ǹ̶�����ʱ,������������Ϣ(������������)
    If lngCurrID <> 0 Then
        If Not blnOnlyData Then gcnOracle.Execute "Delete From zlRPTFmts Where ����ID=" & lngCurrID
        gcnOracle.Execute "Delete From zlRPTDatas Where ����ID=" & lngCurrID
    End If
    
    Do While Not objText.AtEndOfStream
        strLine = objText.ReadLine
        
        '�жϸ�ʽ�Ƿ���ȷ
        If strSect = "" And Trim(strLine) <> "" And Trim(strLine) <> "[HEAD]" Then
            objText.Close
            gcnOracle.RollbackTrans
            Exit Function
        End If
        
        'ȡ�öκ�
        If Left(strLine, 1) = "[" And Right(strLine, 1) = "]" Then
            strSect = UCase(Mid(strLine, 2, Len(strLine) - 2))
        End If
        
        '������ͷ
        If strSect = "HEAD" Then
            '������
            If strLine Like "������=*" Then
                strNum = Mid(strLine, InStr(strLine, "=") + 1)
                strPreNum = strNum
                
                '������:����ظ�������ȡһ�����
                If lngCurrID = 0 Then
                    If CheckExist("zlReports", "���", strNum) Then
                        strNum = GetNextNO
                    End If
                End If
            End If
            '��������
            If strLine Like "��������=*" Then strName = Mid(strLine, InStr(strLine, "=") + 1)
            '����˵��
            If strLine Like "����˵��=*" Then strNote = Mid(strLine, InStr(strLine, "=") + 1)
            '�����ֹʱ��
            If strLine Like "��ֹ��ʼʱ��=*" Then str��ֹ��ʼʱ�� = Format(Mid(strLine, InStr(strLine, "=") + 1), "HH:mm:ss")
            If strLine Like "��ֹ����ʱ��=*" Then str��ֹ����ʱ�� = Format(Mid(strLine, InStr(strLine, "=") + 1), "HH:mm:ss")
        End If
        
        '��������
        '����һ����¼
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
                        If lngClassID > 0 Then rsReport!����id = lngClassID
                        blnReport = True
                    End If
                Case "ZLRPTFMTS"
                    If Not blnOnlyData Then
                        rsFMT.AddNew: blnFmt = True
                        '������ǰ�����ı����ʽ,���и�ʽͳһֽ��
                        If InStr(StrFmt, ",ֽ��,") > 0 And intֽ�� <> 0 Then
                            rsFMT!W = lngRptW
                            rsFMT!H = lngRptH
                            rsFMT!ֽ�� = intֽ��
                            rsFMT!ֽ�� = intֽ��
                            rsFMT!��ֽ̬�� = IIF(bln��ֽ̬��, 1, 0)
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

        'ѭ��ȡ�ɶ����ı���ɵĴ�����Դ
        If InStr(strLine, "=") > 0 And Right(strLine, 1) <> ";" And strSect <> "HEAD" Then
            Do While Not objText.AtEndOfStream And Right(strLine, 1) <> ";"
                strLine = strLine & vbCrLf & objText.ReadLine
            Loop
        End If
        
        '�ֶ�ȡֵ
        If InStr(strLine, "=") > 0 And Right(strLine, 1) = ";" And strSect <> "HEAD" Then
            strFld = Left(strLine, InStr(strLine, "=") - 1)
            strValue = Mid(strLine, InStr(strLine, "=") + 1)
            strValue = Left(strValue, Len(strValue) - 1)

            If UCase(strFld) = "ԭID" And UCase(strSect) = "ZLRPTITEMS" And blnOnlyData = False Then
                Col.Add rsCopy.Fields("ID").Value, "_" & strValue
            End If
            If UCase(strFld) = "ԭID" And UCase(strSect) = "ZLRPTDATAS" And blnOnlyData = False Then
                ColData.Add rsData.Fields("ID").Value, "_" & strValue
            End If
            '����Ƭ����Դ���պʹ���ؼ����ӹ�ϵ
            If (UCase(strFld) = "ԴID" Or UCase(strFld) = "��ID") And UCase(strSect) = "ZLRPTITEMS" And blnOnlyData = False Then
                rsItemCopy.Filter = "ID=" & rsItem.Fields("ID").Value
                If rsItemCopy.RecordCount = 0 Then
                    rsItemCopy.AddNew
                    rsItemCopy!id = rsItem.Fields("ID").Value
                End If
                If strValue <> "" Then
                    If UCase(strFld) = "��ID" Then
                        rsItemCopy!��ID = Val(strValue)
                    ElseIf UCase(strFld) = "ԴID" Then
                        rsItemCopy!ԴID = Val(strValue)
                    End If
                End If
                rsItemCopy.Update
                
                strValue = ""
            End If

            If strFld = "�ϼ�ID" Then
                If strValue = "" Then
                    lngUPID = 0
                Else
                    lngUPID = lngUPID + 1
                End If
            End If
            
            'ȡ�����ļ��е�ֽ��������Ϣ,���ڼ����Ͻṹ�ĵ�������
            If strSect = "ZLREPORTS" Then
                If UCase(strFld) = "W" Then lngRptW = Val(strValue)
                If UCase(strFld) = "H" Then lngRptH = Val(strValue)
                If strFld = "ֽ��" Then intֽ�� = Val(strValue)
                If strFld = "ֽ��" Then intֽ�� = Val(strValue)
                If strFld = "��ֽ̬��" Then bln��ֽ̬�� = Val(strValue) = 1
            End If
            
            '�ж��Ƿ��и��ֶ�
            Set rsCopy = Nothing
            If strValue <> "" Then 'ֵΪ���򲻸�ֵ
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
                '�Ϸ��Լ��
                If strSect = "ZLREPORTS" And strFld = "����" Then
                    If GetPass(strPreNum, strName) <> strValue Then
                        objText.Close
                        gcnOracle.RollbackTrans
                        Exit Function
                    End If
                End If
                '����
                If UCase(strValue) = UCase("SysDate") Then
                    rsCopy.Fields(strFld).Value = Currentdate
                ElseIf UCase(strValue) = UCase("[���]") Then
                    rsCopy.Fields(strFld).Value = strNum
                ElseIf strSect = "ZLREPORTS" And strFld = "����" Then
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
                ElseIf rsCopy.Fields(strFld).name = "����" Then
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
                            '��ʱ��֧��ͼƬ����
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
    '����ID��ԴID����
    If blnOnlyData = False Then
        rsItemCopy.Filter = "��ID <> ''"
        If rsItemCopy.RecordCount > 0 Then
            rsItemCopy.MoveFirst
            Do While Not rsItemCopy.EOF
                rsItem.Filter = "ID=" & rsItemCopy!id
                rsItem!��ID = Val(Col("_" & rsItemCopy!��ID))
                rsItem.Update
                rsItemCopy.MoveNext
            Loop
        End If
        rsItemCopy.Filter = "ԴID <> ''"
        If rsItemCopy.RecordCount > 0 Then
            rsItemCopy.MoveFirst
            Do While Not rsItemCopy.EOF
                rsItem.Filter = "ID=" & rsItemCopy!id
                rsItem!ԴID = Val(ColData("_" & rsItemCopy!ԴID))
                rsItem.Update
                rsItemCopy.MoveNext
            Loop
        End If
    End If
        
    '���²��ݱ�����Ϣ
    If lngCurrID <> 0 Then
        gcnOracle.Execute _
            "Update zlReports" & _
            " Set �޸�ʱ��=Sysdate,����ʱ��=Decode(����ʱ��,NULL,NULL,Sysdate)" & _
            ",��ֹ��ʼʱ��=to_date('" & str��ֹ��ʼʱ�� & "','HH24:MI:SS')" & _
            ",��ֹ����ʱ��=to_date('" & str��ֹ����ʱ�� & "','HH24:MI:SS')" & _
            " Where ID=" & lngCurrID
    End If
    
    '�������б���,�������Ȩ��������д.�µ��빲��������������Ȩ
    If lngCurrID <> 0 Then
        strSQL = _
            " Select ϵͳ,����ID,����,˵�� From zlReports" & _
            " Where ����ID is Not NULL And ���� is Not NULL And ID=[1]" & _
            " Union All" & _
            " Select A.ϵͳ,A.����ID,B.����,C.˵��" & _
            " From zlRptGroups A,zlRptSubs B,zlReports C" & _
            " Where A.ID=B.��ID And B.����ID=C.ID And A.����ID is Not NULL" & _
            " And B.���� is Not NULL And B.����ID=[1]" & _
            " Union ALL" & _
            " Select A.ϵͳ,A.����ID,A.����,B.˵��" & _
            " From zlRPTPuts A,zlReports B Where A.����ID=B.ID And A.����ID=[1]"
        Set rsTmp = OpenSQLRecord(strSQL, "ExportReport", lngCurrID)
        If Not rsTmp.EOF Then
            '����Դ���漰�Ķ���
            strSQL = _
                "Select Distinct B.���� From zlReports A,zlRptDatas B " & _
                "Where A.ID=B.����ID And B.���� is Not NULL And A.ID=[1] " & _
                "    And nvl(b.�������ӱ��,0) <= 0 "
            Set rsPriv = OpenSQLRecord(strSQL, "ExportReport", lngCurrID)
            Do While Not rsPriv.EOF
                For i = 0 To UBound(Split(rsPriv!����, ","))
                    strTmp = Split(rsPriv!����, ",")(i)
                    If InStr(strObject & ",", "," & strTmp & ",") = 0 Then
                        If InStr(",SYS,SYSTEM,ZLTOOLS,", "," & UCase(Split(strTmp, ".")(0)) & ",") = 0 Then
                            strObject = strObject & "," & strTmp
                        End If
                    End If
                Next
                rsPriv.MoveNext
            Loop
            
            '�������漰�Ķ���
            strSQL = _
                "Select Distinct Replace(C.����,'|',',') as ���� " & _
                "From zlReports A,zlRptDatas B,zlRptPars C " & _
                "Where A.ID=B.����ID And B.ID=C.ԴID And C.���� is Not NULL " & _
                "    And nvl(b.�������ӱ��,0) <= 0 And A.ID=[1]"
            Set rsPriv = OpenSQLRecord(strSQL, "ExportReport", lngCurrID)
            Do While Not rsPriv.EOF
                For i = 0 To UBound(Split(rsPriv!����, ","))
                    strTmp = Split(rsPriv!����, ",")(i)
                    If InStr(strObject & ",", "," & strTmp & ",") = 0 And strTmp <> "" Then
                        If InStr(",SYS,SYSTEM,ZLTOOLS,", "," & UCase(Split(strTmp, ".")(0)) & ",") = 0 Then
                            strObject = strObject & "," & strTmp
                        End If
                    End If
                Next
                rsPriv.MoveNext
            Loop
            strObject = Mid(strObject, 2)
            
            '����Ȩ��
            Do While Not rsTmp.EOF
                strSQL = "Select 1 From Zlprograms Where NVL(ϵͳ,0) = [1] And ��� = [2]"
                Set rsProgram = OpenSQLRecord(strSQL, "ExportReport", Val(rsTmp!ϵͳ & ""), Val(rsTmp!����id & ""))
                '��ϵͳģ�����
                If Not rsProgram.EOF Then
                    '�������б���,ֻ��ɾ����Ӧ����,��Ȼ��Ʊ�ݻ�ɾ���������Ǳ���Ĺ���
                    '��Ϊɾ���˹���,��������Ӧ��ɫ����������Ȩ
                    gcnOracle.Execute "Delete From zlProgPrivs Where Nvl(ϵͳ,0)=" & Nvl(rsTmp!ϵͳ, 0) & " And ���=" & rsTmp!����id & " And ����='" & rsTmp!���� & "'"
                    
                    gcnOracle.Execute "Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Select " & _
                        IIF(IsNull(rsTmp!ϵͳ), "NULL", rsTmp!ϵͳ) & "," & rsTmp!����id & ",'" & rsTmp!���� & "','" & Nvl(rsTmp!˵��) & "' From Dual" & _
                        " Where Not Exists(Select 1 From zlProgFuncs Where Nvl(ϵͳ,0)=" & Nvl(rsTmp!ϵͳ, 0) & " And ���=" & rsTmp!����id & " And ����='" & rsTmp!���� & "')"
                        
                    If strObject <> "" Then
                        For i = 0 To UBound(Split(strObject, ","))
                            gcnOracle.Execute _
                                GetInsertProgPrivs(rsTmp!ϵͳ, rsTmp!����id, rsTmp!���� _
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
    Set grsReport = Nothing '�������
    '�������룬�ҵ��뵽ָ������
    If lngCurrID = 0 And lngGroupID <> 0 Then
        On Error Resume Next
        lng��� = 1
        Set rsTmp = New ADODB.Recordset
        strSQL = "Select Count(1) Records From zlRPTSubs Where ��ID=[1]"
        Set rsTmp = OpenSQLRecord(strSQL, "������", lngGroupID)
        If Not rsTmp.EOF Then
            lng��� = Nvl(rsTmp!Records, 0) + 1
        End If
        gcnOracle.Execute "Insert Into zlRPTSubs(��ID,����ID,���,����) Values(" & lngGroupID & "," & lngRPTID & "," & lng��� & ",'" & strName & "')"
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo errH
    End If
    If lngCurrID = 0 Then
        ImportReport = lngRPTID & "|" & strNum & "|" & strName & "|" & strNote
    Else
        ImportReport = lngCurrID & "|" & rsCurr!��� & "|" & rsCurr!���� & "|" & IIF(IsNull(rsCurr!˵��), "", rsCurr!˵��)
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
'���ܣ����洰�弰���и��ֿؼ���״̬
'������objForm:Ҫ����Ĵ���
'      strProjectName����ǰ��������ͨ������app.ProductName���ݣ��������ֲ�ͬ�����е�ͬ�����壬��֤�ָ�����ȷ�ԣ�
'      strUserDef����Ҫ�����ڹ����У�һ������������ʹ��(����ʹ�� set frmxxx=new frm��ƴ�����ʽ)��Ϊ�˰���ͬӦ�ñ���ָ����Եĸ��Ի�״̬����Ҫֱ��ȷ��������
    
    Dim objThis As Object, strTmp As String
    Dim i As Integer, blnDo As Boolean
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    If Not gcnOracle Is Nothing And strProjectName <> "" And gblnRunLog Then
        If gcnOracle.State = 1 Then
            '�����˳�
            On Error Resume Next
            If gstrComputerName <> "" Then
                strSQL = "Zl_Zldiarylog_Update('" & gstrComputerName & "','" & UCase(strProjectName) & _
                        "','" & UCase(objForm.name) & "',1)"
                Call ExecuteProcedure(strSQL, "���¹�����־")
            End If
            If Err.Number <> 0 Then Err.Clear
        End If
    End If
    
    On Error Resume Next
    If Not gfrmMain Is Nothing Then Call gfrmMain.Shut����(objForm)
    On Error GoTo 0
    
    If mdlPublic.GetMemoryParam() = False Then      'ʹ�ø��Ի����
        Call DelWinState(objForm, strProjectName, strUserDef)
        SaveWinState = True: Exit Function
    End If
    
    If strProjectName <> "" Then strProjectName = strProjectName & "\"
    
    '���洰��״̬��λ�á���С
    With objForm
        Select Case .WindowState
            Case 0
                SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strProjectName & objForm.name & strUserDef & "\Form", "״̬", objForm.WindowState & "," & .Left & "," & .Top & "," & .Width & "," & .Height
            Case 1
                SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strProjectName & objForm.name & strUserDef & "\Form", "״̬", 0
            Case 2
                SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strProjectName & objForm.name & strUserDef & "\Form", "״̬", objForm.WindowState
        End Select
    End With
    
    '������ֿؼ��ĸ���״̬
    For Each objThis In objForm.Controls
        strTmp = ""
        On Error Resume Next
        If UCase(TypeName(objThis)) = UCase("Menu") Then
            If objThis.Caption Like "��׼��ť*" Or _
                objThis.Caption Like "�ı���ǩ*" Or _
                objThis.Caption Like "״̬��*" Or _
                UCase(objThis.name) Like UCase("mnuViewTool*") Then
                '����˵��ĸ�ѡ
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
            'Left,Top,Width��Height,Visible
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
            SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strProjectName & objForm.name & strUserDef & "\" & TypeName(objThis), objThis.name & "״̬", strTmp
        End If
        
        Select Case UCase(TypeName(objThis))
            Case UCase("Toolbar")
                If objThis.Buttons.count > 0 Then
                    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strProjectName & objForm.name & strUserDef & "\" & TypeName(objThis), objThis.name & "�ı�", IIF(objThis.Buttons(1).Caption <> "", 1, objThis.ButtonHeight)
                End If
            Case UCase("ListView")
                SaveListViewState objThis, strProjectName & objForm.name & strUserDef
            Case UCase("CoolBar")
                strTmp = ""
                For i = 1 To objThis.Bands.count
                    strTmp = strTmp & "," & objThis.Bands(i).NewRow
                Next
                SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strProjectName & objForm.name & strUserDef & "\" & TypeName(objThis), objThis.name & "����", Mid(strTmp, 2)
                
                strTmp = ""
                For i = 1 To objThis.Bands.count
                    strTmp = strTmp & "," & objThis.Bands(i).Visible
                Next
                SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strProjectName & objForm.name & strUserDef & "\" & TypeName(objThis), objThis.name & "�ɼ���", Mid(strTmp, 2)
            Case UCase("VSFlexGrid")
                SaveFlexState objThis, strProjectName & objForm.name & strUserDef
        End Select
    Next
    SaveWinState = True
End Function

Public Function RestoreWinState(objForm As Object, Optional ByVal strProjectName As String, Optional ByVal strUserDef As String) As Boolean
'���ܣ��ָ������״̬�����󶥱߽糬��ʱ�����Զ�����Ϊ0
'������objForm:Ҫ�ָ��Ĵ���
'      strProjectName����ǰ��������ͨ������app.ProductName���ݣ��������ֲ�ͬ�����е�ͬ�����壬��֤�ָ�����ȷ�ԣ�
'      strUserDef����Ҫ�����ڹ����У�һ������������ʹ��(����ʹ�� set frmxxx=new frm��ƴ�����ʽ)��Ϊ�˰���ͬӦ�ñ���ָ����Եĸ��Ի�״̬����Ҫֱ��ȷ��������
   
    Dim aryInfo() As String
    Dim strTmp As String, i As Integer
    Dim objThis As Object
    Dim blnDo As Boolean
    Dim strSave As String
    Dim strOEM As String
    Dim strSQL As String
    
    If Not gcnOracle Is Nothing And strProjectName <> "" And gblnRunLog Then
        If gcnOracle.State = 1 Then
            '�����˳�
'            strSQL = "Update zlDiaryLog Set �˳�ԭ��=2,�˳�ʱ��=Sysdate" & _
'                    " Where �˳�ԭ�� is NULL And �Ự�� Not IN(Select SID+SERIAL# From v$Session Where USER#<>0)"
'            gcnOracle.Execute strSQL

            '����
            On Error Resume Next
            If gstrComputerName <> "" Then
                strSQL = "Zl_Zldiarylog_Insert('" & gstrComputerName & "','" & UCase(strProjectName) & "'," & _
                    " '" & UCase(objForm.name) & "','" & UCase(objForm.Caption) & "')"
                Call ExecuteProcedure(strSQL, "���湤����־")
            End If
            If Err.Number <> 0 Then Err.Clear
        End If
    End If
    
    On Error Resume Next
    
    If Not gfrmMain Is Nothing Then Call gfrmMain.Show����(objForm)
    
    blnDo = mdlPublic.GetMemoryParam()      'ʹ�ø��Ի����
    
    If strProjectName <> "" Then strProjectName = strProjectName & "\"
    
    '�ָ������״̬��λ�á���С
    If UCase(objForm.name) = UCase("frmReport") _
        Or UCase(objForm.name) = UCase("frmPreview") _
            Or UCase(objForm.name) = UCase("frmDesign") Then
        strTmp = "2" '���ⴰ���ʼ���
    Else
        strTmp = "0," & (Screen.Width - objForm.Width) / 2 & "," & (Screen.Height - objForm.Height) / 2 & "," & objForm.Width & "," & objForm.Height
    End If
    If blnDo Then
        strSave = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strProjectName & objForm.name & strUserDef & "\Form", "״̬", "")
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

    '�ָ������и��ֿؼ��ĸ���״̬
    For Each objThis In objForm.Controls
        
        On Error Resume Next
        If blnDo Then
            strTmp = ""
            If UCase(TypeName(objThis)) = UCase("Menu") Then
                '����˵��ĸ�ѡ
                If objThis.Caption Like "��׼��ť*" Or _
                    objThis.Caption Like "�ı���ǩ*" Or _
                    objThis.Caption Like "״̬��*" Or _
                    UCase(objThis.name) Like UCase("mnuViewTool*") Then
                    strTmp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strProjectName & objForm.name & strUserDef & "\" & TypeName(objThis), objThis.name & "״̬", "")
                    If UBound(Split(strTmp, ",")) = 1 Then
                        objThis.Checked = Split(strTmp, ",")(0)
                        objThis.Enabled = Split(strTmp, ",")(1)
                    End If
                End If
            ElseIf UCase(objThis.Tag) = "SAVE" Or UCase(objThis.name) Like "*_S" Or _
                UCase(TypeName(objThis)) = UCase("StatusBar") Or _
                UCase(TypeName(objThis)) = UCase("Toolbar") Or _
                UCase(TypeName(objThis)) = UCase("Coolbar") Then
                
                strTmp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strProjectName & objForm.name & strUserDef & "\" & TypeName(objThis), objThis.name & "״̬", "")
                If strTmp <> "" Then
                    'Left,Top,Width��Height,Visible
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
                '״̬�����ñ�־
'                If zlRegInfo("��Ȩ����") <> "1" Then
'                    If objThis.Panels(1).Bevel = sbrRaised Then
'                        objThis.Panels(1).Text = ""
'                        Set objThis.Panels(1).Picture = LoadCustomPicture("Try")
'                        objThis.Panels(1).ToolTipText = ""
'                        objThis.Height = 360
'                    End If
'                Else
                    If objThis.Panels(1).Bevel = sbrRaised Then
                        strTmp = zlRegInfo("��Ʒ����")
                        If strTmp <> "-" Then
                            objThis.Panels(1).Text = strTmp & "���"
                            '����״̬��ͼ���OEM����
                            If strTmp = "����" Then
                                If zlRegInfo("��Ȩ����") <> "1" Then
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
                                If zlRegInfo("��Ȩ����") <> "1" Then objThis.Panels(1).Text = strTmp & "(����)"
                            End If
                            objThis.Panels(1).ToolTipText = ""
                            objThis.Height = 360
                        End If
                    End If
'                End If
            Case UCase("Menu")
                If UCase(objThis.name) = UCase("mnuHelpWeb") Then
                    'WEB�ϵ�����
                    strTmp = zlRegInfo("֧���̼���")
                    If strTmp <> "-" Then
                        objThis.Caption = "&WEB�ϵ�" & strTmp
                    End If
                ElseIf UCase(objThis.name) = UCase("mnuHelpWebHome") Then
                    '������ҳ
                    strTmp = zlRegInfo("֧���̼���")
                    If strTmp <> "-" Then
                        objThis.Caption = strTmp & "��ҳ(&H)"
                    End If
                End If
            Case UCase("Toolbar")
                If blnDo Then
                    If objThis.Buttons.count > 0 Then
                        strTmp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strProjectName & objForm.name & strUserDef & "\" & TypeName(objThis), objThis.name & "�ı�", 1)
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
                    strTmp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strProjectName & objForm.name & strUserDef & "\" & TypeName(objThis), objThis.name & "����", "")
                    If UBound(Split(strTmp, ",")) >= 0 Then
                        For i = 0 To UBound(Split(strTmp, ","))
                            objThis.Bands(i + 1).NewRow = Split(strTmp, ",")(i)
                        Next
                    End If
            
                    strTmp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strProjectName & objForm.name & strUserDef & "\" & TypeName(objThis), objThis.name & "�ɼ���", "")
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
    
    strTmp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strForm & "\" & TypeName(objThis), objThis.name & "���", "")
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
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strForm & "\" & TypeName(objThis), objThis.name & "���", Mid(strTmp, 2)
End Sub

Public Sub SaveListViewState(objLvw As Object, ByVal strForm As String, Optional strIndex As String)
'���ܣ�����ListView�ĸ�������
'������objLvw=ListView����,strForm=����ؼ���
'˵������ͼ��ʽ���п���λ�á��б��⡢�ж��롢����
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
    
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strForm & "\ListView", objLvw.name & strIndex & "��ͼ", objLvw.View
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strForm & "\ListView", objLvw.name & strIndex & "���", Mid(strWidth, 2)
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strForm & "\ListView", objLvw.name & strIndex & "λ��", Mid(strPosition, 2)
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strForm & "\ListView", objLvw.name & strIndex & "����", Mid(strText, 2)
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strForm & "\ListView", objLvw.name & strIndex & "����", Mid(strAlign, 2)
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strForm & "\ListView", objLvw.name & strIndex & "����", objLvw.SortKey & "," & objLvw.SortOrder & "," & objLvw.Sorted
End Sub

Public Sub RestoreListViewState(objLvw As Object, ByVal strForm As String, Optional strIndex As String)
'���ܣ��ָ�ListView�ĸ�������
'������objLvw=ListView����,strForm=����ؼ���
'˵������ͼ��ʽ���п���λ�á��б��⡢�ж��롢����
    Dim lngCol As Long
    Dim strWidth As String
    Dim strPosition As String
    Dim strText As String, arrText As Variant
    Dim strAlign As String
    Dim strSort As String
    
    On Error Resume Next
    
    strText = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strForm & "\ListView", objLvw.name & strIndex & "����")
    If strText <> "" Then
        arrText = Split(strText, ",")
        If objLvw.Tag = "�ɱ仯��" Then
            'ֻ����Ҫ��ListView�����б仯
            objLvw.ColumnHeaders.Clear
            For lngCol = LBound(arrText) To UBound(arrText)
                '��ȱʡ�ؼ���Ϊ"_" & �б���
                objLvw.ColumnHeaders.Add , "_" & arrText(lngCol), arrText(lngCol)
            Next
        Else
            '����Ƿ���Ҫ�ָ�
            '��������,���ָ���ʹ��ȱʡ
            If UBound(arrText) + 1 <> objLvw.ColumnHeaders.count Then Exit Sub
            '�б������,���ָ���ʹ��ȱʡ
            For lngCol = 1 To objLvw.ColumnHeaders.count
                If objLvw.ColumnHeaders(lngCol).Text <> arrText(lngCol - 1) Then Exit Sub
            Next
        End If
    End If
    
    '��ͼȱʡ���ֳ�ʼֵ
    lngCol = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strForm & "\ListView", objLvw.name & strIndex & "��ͼ", -1)
    If lngCol <> -1 Then objLvw.View = lngCol
    
    '�еĿ��,˳��,����
    strWidth = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strForm & "\ListView", objLvw.name & strIndex & "���")
    strPosition = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strForm & "\ListView", objLvw.name & strIndex & "λ��")
    strAlign = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strForm & "\ListView", objLvw.name & strIndex & "����")
    For lngCol = 1 To objLvw.ColumnHeaders.count
        '��ȱʡ�ؼ���Ϊ"_" & �б���
        objLvw.ColumnHeaders(lngCol).Key = "_" & objLvw.ColumnHeaders(lngCol).Text
        If strWidth <> "" Then objLvw.ColumnHeaders(lngCol).Width = Split(strWidth, ",")(lngCol - 1)
        If strAlign <> "" Then objLvw.ColumnHeaders(lngCol).Alignment = Split(strAlign, ",")(lngCol - 1)
    Next
    For lngCol = objLvw.ColumnHeaders.count To 1 Step -1
        If strPosition <> "" Then objLvw.ColumnHeaders(lngCol).Position = Split(strPosition, ",")(lngCol - 1)
    Next
    
    '��������
    strSort = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & strForm & "\ListView", objLvw.name & strIndex & "����")
    If strSort <> "" Then
        objLvw.SortKey = Split(strSort, ",")(0)
        objLvw.SortOrder = Split(strSort, ",")(1)
        objLvw.Sorted = Split(strSort, ",")(2)
    End If
End Sub

Public Function DelWinState(objForm As Object, Optional ByVal strProjectName As String, Optional ByVal strUserDef As String) As Boolean
'���ܣ�ɾ��������Ի�����ֵ
'������objForm:Ҫ�ָ��Ĵ���
'      strProjectName����ǰ��������ͨ������app.ProductName���ݣ��������ֲ�ͬ�����е�ͬ�����壬��֤�ָ�����ȷ�ԣ�
'      strUserDef����Ҫ�����ڹ����У�һ������������ʹ��(����ʹ�� set frmxxx=new frm��ƴ�����ʽ)��Ϊ�˰���ͬӦ�ñ���ָ����Եĸ��Ի�״̬����Ҫֱ��ȷ��������
    Dim strProject As String
    Dim lngR As Long
    Dim objThis As Object
    
    strProject = strProjectName
    If strProjectName <> "" Then strProjectName = strProjectName & "\"
    
    For Each objThis In objForm.Controls
        lngR = RegDeleteKey(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\ZLSOFT\˽��ģ��\" & gstrDBUser & "\��������\" & strProjectName & objForm.name & strUserDef & "\" & TypeName(objThis) & Chr(0))
        If lngR <> 0 And lngR <> 2 Then Exit Function
    Next
    
    lngR = RegDeleteKey(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\ZLSOFT\˽��ģ��\" & gstrDBUser & "\��������\" & strProjectName & objForm.name & strUserDef & "\Form" & Chr(0))
    If lngR <> 0 And lngR <> 2 Then Exit Function
    lngR = RegDeleteKey(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\ZLSOFT\˽��ģ��\" & gstrDBUser & "\��������\" & strProjectName & objForm.name & strUserDef & Chr(0))
    If lngR <> 0 And lngR <> 2 Then Exit Function
    
    DelWinState = True
End Function

Public Function LoadCustomPicture(strID As String, Optional strFormat As String = "GIF") As StdPicture
'����:����Դ�ļ��е�ָ����Դ���ɴ����ļ�
'����:ID=��Դ��,strExt=Ҫ�����ļ�����չ��(��BMP)
'����:�����ļ���
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
'���ܣ���ָ���ֶ��еĶ�������������һ�������ļ�
'���أ�ͼ�ζ��󣬻�����Ϊ�յĳ�ʼ���˵�ͼƬ
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
'���ܣ���ָ��ͼ�δ�ŵ�ָ���ļ�¼���ֶ���
'˵���������¼�¼��
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
'���ܣ��ж�ָ������Դ��ָ�������ʽ���Ƿ�ʹ��
'������strData=����Դ��
'      blnFormat=�Ƿ�ֻ�ڱ���ǰ�ĸ�ʽ���ж�(ȱʡΪ��,�����и�ʽ���ж�)
'˵������ǩ���������ͷ�еı�ǩ
    Dim tmpItem As RPTItem, tmpPar As RPTPar
    Dim strContent As String
    
    For Each tmpItem In objReport.Items
        '�з�����ض��з�������
        If (blnFormat And tmpItem.��ʽ�� = objReport.bytFormat Or Not blnFormat) _
            And InStr("2,3,5,6,12,13,14,", tmpItem.���� & ",") > 0 Then
            Select Case tmpItem.����
                Case 2, 3, 13 '���ݱ�ǩ,"������Ϣ.����"
                    If InStr(tmpItem.����, strData & ".") > 0 Then DataUsed = True: Exit Function
                Case 5 '������,"������Ϣ"
                    strContent = tmpItem.����
                    If strContent Like "*��*��" Then
                        strContent = Left(strContent, InStrRev(strContent, "��") - 1)
                    End If
                    If strContent = strData Then DataUsed = True: Exit For
                Case 6 '����������,"([������Ϣ.���]+[2])/3"
                    If InStr(tmpItem.����, strData & ".") > 0 Then DataUsed = True: Exit Function
                    If InStr(tmpItem.��ͷ, strData & ".") > 0 Then DataUsed = True: Exit Function
                Case 12 'ͼ��
                    If InStr("|" & tmpItem.����, "|" & strData & ".") > 0 Then DataUsed = True: Exit Function
                Case 14
                    If tmpItem.����Դ = strData Then DataUsed = True: Exit Function
            End Select
        End If
    Next
End Function

Public Function MakeNamePars(objReport As Report, Optional blnFirst As Boolean) As RPTPars
'���ܣ��ӱ���(objReport)��������Դ�в�����������Ψһ�Ĳ�����
'������blnFirst=ǿ��ȡ��ǰ����Ч��ȱʡֵ,����Ϊ����������
'˵����1.��Ƴ���������ͬһ�����в�ͬ����Դ֮��Ĳ������ͬ��,������,ȱʡֵҲ��ͬ
    Dim tmpData As RPTData, tmpPar As RPTPar, StrPar As String
    Dim tmpPars As New RPTPars, strTmp As String
    
    For Each tmpData In objReport.Datas
        If DataUsed(objReport, tmpData.����) Then
            For Each tmpPar In tmpData.Pars
                If InStr(StrPar & ",", "," & tmpPar.���� & ",") = 0 Then
                    StrPar = StrPar & "," & tmpPar.����
                    With tmpPar '����������(Ψһ)�ؼ��ּ���
                        If .Reserve Like "*��|*" And Not blnFirst Then
                            '����������ʱ��Reserve��¼��"������ֵ|��ʾֵ"
                            '����ΪȱʡֵΪ����ʱ�ĺ�ȱʡֵ,ReserveΪ"��ʾֵ|��ֵ"
                            tmpPars.Add .����, .���, .����, .����, CStr(Split(.Reserve, "|")(0)), .��ʽ, .ֵ�б�, .����SQL _
                                , .��ϸSQL, .�����ֶ�, .��ϸ�ֶ�, .����, "_" & .����, Split(.Reserve, "|")(1) & "|" & .ȱʡֵ _
                                , .�Ƿ�����
                        Else
                            '��һ�ν���(����������)���������͵�����
                            tmpPars.Add .����, .���, .����, .����, .ȱʡֵ, .��ʽ, .ֵ�б�, .����SQL, .��ϸSQL, .�����ֶ� _
                                , .��ϸ�ֶ�, .����, "_" & .����, .Reserve, .�Ƿ�����
                        End If
                    End With
                End If
            Next
        End If
    Next
    Set MakeNamePars = tmpPars
End Function

Public Sub ItemAutoSize(objItem As RPTItem, ByVal strValue As String, ByVal objCalc As Object)
'���ܣ����ݱ����ǩԪ�ص������Զ���������
'������objCalc=���ڼ���ʵ�ʿ�ߵĶ���
'˵����1.ֻ�ı���W,H,���ı������ݡ�
'      2.��Ϊ��ǩ��ѭ��ȡֵ,��˲�ѯ��Ԥ��ÿ�ζ�Ҫ������
    If Not objItem.�Ե� Then Exit Sub
    objCalc.Font.name = objItem.����
    objCalc.Font.Size = objItem.�ֺ�
    objCalc.Font.Bold = objItem.����
    objCalc.Font.Italic = objItem.б��
    objCalc.Font.Underline = objItem.����
    
    objItem.W = objCalc.TextWidth(strValue) + objCalc.TextWidth("A")
    objItem.H = objCalc.TextHeight(strValue) + 30
End Sub

Public Function ReplaceBracket(ByVal strValue As String, Optional ByVal strReplace As String) As String
'���ܣ����ַ����е�[]�滻Ϊָ����ֵ
    Dim strLeft As String, strRight As String, strVar As String
    
    '[]����Likeʱ��Ч,����Ҫ�滻
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
'���ܣ������ǩ�еĺ�:[n>=0],[=������]
'˵����������[ҳ��][ҳ��]
    Dim strLeft As String, strRight As String, strVar As String
    
    '[]����Likeʱ��Ч,����Ҫ�滻
    strVar = Replace(strValue, "[", "@@")
    strVar = Replace(strVar, "]", "$$")
    If Not strVar Like "*@@*$$*" Then GetLabelMacro = strValue: Exit Function
    If strVar Like "*@@*.*$$*" Then GetLabelMacro = strValue: Exit Function
    
    Do While InStr(strValue, "[") > 0
        strLeft = Left(strValue, InStr(strValue, "[") - 1)
        strRight = Mid(strValue, InStr(strValue, "]") + 1)
        strVar = Mid(strValue, InStr(strValue, "[") + 1, InStr(strValue, "]") - InStr(strValue, "[") - 1)
            
        If IsNumeric(strVar) Then '��������
            If CInt(strVar) >= 0 Then strVar = GetUserParData(frmParent, CInt(strVar))
        ElseIf Left(strVar, 1) = "=" Then '[=������]
            If Mid(strVar, 2) <> "" Then strVar = GetParValue(frmParent, Mid(strVar, 2))
        ElseIf strVar = "��λ����" Then
            strVar = Replace(zlRegInfo("��λ����", , -1), ";", vbCrLf)
        ElseIf strVar = "����Ա����" Then
            strVar = gstrUserName
        ElseIf strVar = "����Ա���" Then
            strVar = gstrUserNO
        ElseIf IsDate(Format("2000-01-01", strVar)) Then '��ǰ����
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
'���ܣ���ȡ��ǩ���������������ֶ���.
'���أ���ʽ"����Դ.�ֶ�|����Դ.�ֶ�|..."
    Dim strLeft As String, strRight As String, strVar As String
    
    If Not BracketMatch(strValue, "[]") Then Exit Function
    
    '[]����Likeʱ��Ч,����Ҫ�滻
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
'���ܣ����ָ���ַ�����ָ���������Ƿ�ƥ��
'������strText=Ҫ�����ַ���
'      strBracket=���Ŷԣ���"[]"
'      blnNesting=�����Ƿ�����Ƕ��,��"[..[...]..]"��ʽ
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
'���ܣ���ȡָ�������ָ�����еı�ǩ����
    Dim tmpID As RelatID, strTmp As String
    
    For Each tmpID In objItem.SubIDs
        If frmSource.mobjReport.Items("_" & tmpID.id).��� = C Then
            strTmp = frmSource.mobjReport.Items("_" & tmpID.id).��ͷ
            strTmp = CStr(Split(Split(strTmp, "|")(R), "^")(2))
            GetHeadCellScript = strTmp
            Exit Function
        End If
    Next
End Function

Public Function GetGridStyle(objReport As Report, id As Integer) As Byte
'���ܣ��ж����������ʽ
'���أ�0:��ͷ�ı������Ч,1-����ͷ��Ч,2-��������Ч
'˵���������ͷ�������Ч���򷵻����߶���Ч
    Dim i As Integer, tmpID As RelatID
    Dim blnBody As Boolean, blnHead As Boolean
    Dim strTmp As String
    
    If objReport.Items("_" & id).���� <> 4 Then Exit Function
    
    blnHead = False
    blnBody = False
    For Each tmpID In objReport.Items("_" & id).SubIDs
        strTmp = objReport.Items("_" & tmpID.id).��ͷ
        i = UBound(Split(strTmp, "|"))
        If i > 0 Then
            blnHead = True
        ElseIf i = 0 Then
            blnHead = blnHead Or (Split(Split(strTmp, "|")(i), "^")(2) <> "#")
        End If
        blnBody = blnBody Or (objReport.Items("_" & tmpID.id).���� <> "")
        
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
'���ܣ���ָ���ļ���ŵ�ָ���ļ�¼���ֶ���
'˵���������¼�¼��
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
'���ܣ���ȡ��ǩ�����յı��ID
'������strName=�����յı����
'˵����������ձ��������ӱ��,�򸽼ӱ���IDҲһ������
    Dim objItem As RPTItem
    Dim strIDs As String
    
    For Each objItem In frmSource.mobjReport.Items
        If objItem.��ʽ�� = frmSource.bytFormat And (objItem.���� = 4 Or objItem.���� = 5) _
            And ((objItem.���� = 0 And objItem.���� = strName) _
                Or (objItem.���� = 1 And objItem.���� = strName)) Then
            strIDs = strIDs & "," & objItem.id
        End If
    Next
    GetDependIDs = Mid(strIDs, 2)
End Function

Public Function GetRightWidth(lngCol As Long, lngEnd As Long, lngRow As Long, strSkip As String, strSkip2 As String, objGrid As Object) As Long
'���ܣ���ȡ��ǰ���ҳ��ָ�����ұ߻���Ҫ����еĿ��
'������lngCol=��ǰ�����
'      lngEnd=��ǰҳ���������
'      lngRow=��ǰ�����
'      strSkip=�������������Ԫ��Ĵ�
'      objGrid=���
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
        If frmSource.mobjReport.Items("_" & tmpID.id).��� = intCol Then
            Set GetSubItem = frmSource.mobjReport.Items("_" & tmpID.id): Exit Function
        End If
    Next
End Function

Public Function PrintPage(ByVal intPage As Integer, objOut As Object, frmSource As Object, _
    Optional ByVal sngScale As Single = 1, Optional ByVal blnSure As Boolean = True, _
    Optional ByVal blnMeasure As Boolean, Optional lngMaxH As Long) As Boolean
'���ܣ���ӡ(Ԥ��)һҳ
'������intPage=Ҫ�����ҳ��,>=0
'      objOut=�������,Printer��PictureBox
'      frmSource=�������ݱ��Ĵ���(��frmReport)
'      sngScale=�������,Printerֻ��Ϊ1
'      blnSure=�Ƿ�ʵ�ʿɴ�ӡ����Ԥ��,ȱʡΪTrue,�������ӡ��ʱ�̶�ΪFalse
'      blnMeasure=�Ƿ������ʵ����Ҫ��ӡ��ֽ�Ÿ߶�,���������
'      lngMaxH=���blnMeasure����ʹ��,���ز������������ֽ�Ÿ߶�(Twip)
'˵�����ú���������ҳ��������
'������frmSource.mobjReport,marrPage,mLibDatas
    Dim objFmt As RPTFmt, objSub As RPTItem, objTemp As RPTItem
    Dim arrPage As Variant, objItem As RPTItem, objPageCell As PageCell
    Dim lngCurH As Long, lngPaperW As Long, lngPaperH As Long 'ֽ��
    Dim objBody As Object, objHead As Object, objFont As New StdFont
    Dim strValue As String, strDepend As String, objPic As StdPicture
    Dim strSkip As String, strSkip2 As String '2��������Ϣ����Ϣ
    Dim arrPars As Variant, blnPressWork As Boolean  '�Ƿ��״�
    Dim intBasePage As Integer, colRowIDs As Collection, objCurDLL As clsReport
    
    Dim lngPreRow As Long, lngPreCol As Long, blnHaveGrid As Boolean
    Dim LngRows As Long, lngRowB As Long, lngRowE As Long
    Dim X As Long, Y As Long, W As Long, H As Long '��Щ����Գߴ�
    Dim i As Long, j As Long, k As Long, l As Long, M As Long
    Dim B As Long '��ǰҳ������Ч����ĵ�һ��
    Dim lngindex As Long, lngSize As Long, sngWidth As Single
    Dim lngChildX As Long, lngChildY As Long
    Dim arrPageCard As Variant, objPageCard As PageCard
    Dim lngX As Long, lngY As Long, lngCol As Long, lngRow As Long
    Dim lngRowHeight As Long, lngRowCount As Long, lngRowTotal As Long
    Dim lngRangeHeight As Long
    
    Dim dblSureW As Double, dblSureH As Double
    Dim colColAutoFont As Collection
    Dim strData As String, strTmp As String, strBdr As String
    '��ǩ���յı��ǰҳ���λ�á��ߴ�
    Dim lngOX As Long, lngOY As Long, lngOW As Long, lngOH As Long
    '��ǩʵ�����λ��
    Dim lngOutX As Long, lngOutY As Long, lngDesignH As Long
    Dim blnGroup As Boolean, blnWithData As Boolean
    Dim lngForeColor As Long, lngBackColor As Long
    Dim blnPrint As Boolean                                                     '�����жϴ�ӡ����Ԥ��
    Dim blnAddition As Boolean
    
    'ͨ�����������������ж�
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
    If objFmt.ֽ�� = 1 Then
        lngPaperW = objFmt.W: lngPaperH = objFmt.H
    Else
        lngPaperW = objFmt.H: lngPaperH = objFmt.W
    End If
    
    intBasePage = 1
    arrPars = frmSource.marrPars 'ֱ�ӷ���Ҫ������Get����
    If Not blnMeasure And UBound(arrPars) <> -1 Then
        For i = 0 To UBound(arrPars)
            j = InStr(CStr(arrPars(i)), "=")
            If j > 0 Then
                If UCase(Trim(Left(CStr(arrPars(i)), j - 1))) = UCase("PressWork") Then
                    '�����û���������ж��Ƿ��״�:����ֽ��ʱ������
                    If IsNumeric(Trim(Mid(CStr(arrPars(i)), j + 1))) Then
                        blnPressWork = Val(Trim(Mid(CStr(arrPars(i)), j + 1))) = 1 'ȫ���״�
                    End If
                ElseIf UCase(Trim(Left(CStr(arrPars(i)), j - 1))) = UCase("PressWorkFirst") Then
                    If IsNumeric(Trim(Mid(CStr(arrPars(i)), j + 1))) Then
                        blnPressWork = Val(Trim(Mid(CStr(arrPars(i)), j + 1))) = 1 And intPage = 0 '��ҳ�״�
                    End If
                ElseIf UCase(Trim(Left(CStr(arrPars(i)), j - 1))) = UCase("StartPageNum") Then
                    If IsNumeric(Trim(Mid(CStr(arrPars(i)), j + 1))) Then
                        intBasePage = Val(Trim(Mid(CStr(arrPars(i)), j + 1))) '��ʼ��ӡҳ��
                        If intBasePage = 0 Then intBasePage = 1
                    End If
                End If
            End If
        Next
    End If
    Set colRowIDs = frmSource.mcolRowIDs
    Set objCurDLL = frmSource.mobjCurDLL 'ֱ�ӷ���˵��֧�����Ժͷ���
    
    '����������
    If IsArray(arrPage) Then
        If UBound(arrPage) >= intPage Then
            If arrPage(intPage).count > 0 Then
                blnHaveGrid = True
                'ѭ������ǰҳ�ڵĶ�����
                For Each objPageCell In arrPage(intPage)
                    
                    With objPageCell
                        Set objBody = frmSource.msh(.id)
                        Set objItem = frmSource.mobjReport.Items("_" & .id)
                        
                        '����ָ���˱��Ԫ��ı�ǩԪ��
                        Call SetCellValue(IIF(blnSure, 1, 2), frmSource, objItem, .RowB)
                        
                        '�Զ������������ԵĻ���
                        Set colColAutoFont = New Collection
                        For i = 0 To objBody.Cols - 1
                            colColAutoFont.Add "", "_" & i 'Ϊ""��ʾ������δ����
                        Next
                        
                        objBody.Redraw = False
                        lngPreRow = objBody.Row: lngPreCol = objBody.Col
                        
                        If objItem.���� = 4 Then
                            Set objHead = frmSource.msh(objBody.Tag)
                            objHead.Redraw = False
                        End If
                        
                        '����̶����н��沿��(��������ܱ���)
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
                                            If Not objItem.�߿� And j = 0 Then strBdr = "1101"

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
                                            '���ܱ��-��ͷ
                                            If Not DrawCell(objOut, objBody.Text, .X + X, .Y + Y, W, H, .X + .W, _
                                                        , objBody.GridColor, lngForeColor, objBody.BackColor _
                                                        , objFont, strBdr _
                                                        , GetHscAlign(objBody.Cell(flexcpAlignment, i, j), objBody.Text) _
                                                        , GetVscAlign(objBody.Cell(flexcpAlignment, i, j)) _
                                                        , True, sngScale, , objItem.����߼Ӵ�) Then
                                                Exit Function '�ϲ�ʱ������
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
                                                                
                        '����̶��в���(����������)
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
                                            If Not objItem.�߿� And j = 0 Then strBdr = "1101"
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
                                            '���ܱ��-����
                                            If Not DrawCell(objOut, objBody.Text, .X + X, .Y + Y, W, H, .X + .W, _
                                                        , objBody.GridColor, lngForeColor, objBody.BackColor _
                                                        , objFont, strBdr _
                                                        , GetHscAlign(objBody.Cell(flexcpAlignment, i, j), objBody.Text, objBody, j) _
                                                        , GetVscAlign(objBody.Cell(flexcpAlignment, i, j)) _
                                                        , True, sngScale, , objItem.����߼Ӵ�) Then
                                                Exit Function '�ϲ�ʱ������
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
                        
                        '����̶��в���(����)
                        If objItem.���� = 5 Then
                            strSkip = "": strSkip2 = "": Y = 0
                            For i = 0 To objBody.FixedRows - 1
                                If Not blnMeasure And Not blnPressWork Then
                                    objBody.Row = i: X = .FixW
                                    For j = .ColB To .ColE
                                        objBody.Col = j
                                        If InStr(strSkip, "[" & i & "," & j & "]") = 0 Then
                                            SearchCell objBody, i, j, objBody.FixedRows - 1, .ColE, W, H, strSkip, strSkip2
                                            
                                            strBdr = "1111"
                                            If Not objItem.�߿� And (j = .ColE Or (W > objBody.ColWidth(j) + 15 _
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
                                            '���ܱ��-��ͷ
                                            If Not DrawCell(objOut, objBody.Text, .X + X, .Y + Y, W, H, .X + .W, _
                                                        , objBody.GridColor, lngForeColor, objBody.BackColor _
                                                        , objFont, strBdr _
                                                        , GetHscAlign(objBody.Cell(flexcpAlignment, i, j), objBody.Text) _
                                                        , GetVscAlign(objBody.Cell(flexcpAlignment, i, j)) _
                                                        , True, sngScale, , objItem.����߼Ӵ�) Then
                                                Exit Function  '�ϲ�ʱ������
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
                            For k = 1 To .Copys '�����������ͷ�Զ�����
                                strSkip = "": strSkip2 = "": Y = 0
                                For i = 0 To objHead.FixedRows - 1
                                    If Not blnMeasure And Not blnPressWork Then
                                        objHead.Row = i: X = .W * (k - 1) '����
                                        B = 0
                                        For j = .ColB To .ColE
                                            objHead.Col = j
                                            If objHead.Text = "ɾ����" Then lngindex = j
                                            If InStr(strSkip, "[" & i & "," & j & "]") = 0 Then
                                                '��ͷ��Ԫ��ԭʼ���ݶ���
                                                strValue = GetHeadCellScript(frmSource, objItem, i, j)
                                                If strValue = "#" Then 'Ϊ��
                                                    strValue = ""
                                                ElseIf strValue = "��" Then '����ߵ�Ԫ����ͬ
                                                    For l = j - 1 To 0 Step -1
                                                        strValue = GetHeadCellScript(frmSource, objItem, i, l)
                                                        If strValue <> "��" Then Exit For
                                                    Next
                                                ElseIf strValue = "��" Then '���ϱߵ�Ԫ����ͬ
                                                    For l = i - 1 To 0 Step -1
                                                        strValue = GetHeadCellScript(frmSource, objItem, l, j)
                                                        If strValue <> "��" Then Exit For
                                                    Next
                                                End If
                                                
                                                '����ҳ����
                                                If InStr(strValue, "[ҳ��]") > 0 Then
                                                    strValue = Replace(strValue, "[ҳ��]", intPage + intBasePage)
                                                End If
                                                If InStr(strValue, "[ҳ��]") > 0 Then
                                                    If Not IsArray(arrPage) Then
                                                        strValue = Replace(strValue, "[ҳ��]", intBasePage)
                                                    Else
                                                        strValue = Replace(strValue, "[ҳ��]", UBound(arrPage) + intBasePage)
                                                    End If
                                                End If
                                                If InStr(strValue, "[Ʊ�ݺ�]") > 0 Then
                                                    If IsArray(garrBill) Then
                                                        If UBound(garrBill) >= intPage Then
                                                            strValue = Replace(strValue, "[Ʊ�ݺ�]", garrBill(intPage))
                                                        Else
                                                            strValue = Replace(strValue, "[Ʊ�ݺ�]", "")
                                                        End If
                                                    Else
                                                        strValue = Replace(strValue, "[Ʊ�ݺ�]", "")
                                                    End If
                                                End If
                                                strData = GetLabelDataName(strValue)
                                                
                                                '��һ��ʱ���ݸ�λ,�Ա��ָ�����ͷһ��
                                                If k = 1 Then
                                                    If strData <> "" Then
                                                        For l = 0 To UBound(Split(strData, "|"))
                                                            strTmp = Split(Split(strData, "|")(l), ".")(0)
                                                            If frmSource.mLibDatas("_" & strTmp).DataSet.RecordCount > 0 Then
                                                                frmSource.mLibDatas("_" & strTmp).DataSet.MoveFirst
                                                                '(��ǰҳ-1)��ʾҪѭ��Move�Ĵ���
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
                                                
                                                '��ȡ����
                                                If strData <> "" Then
                                                    For l = 0 To UBound(Split(strData, "|"))
                                                        strTmp = GetFieldValue(frmSource, CStr(Split(strData, "|")(l)))
                                                        strValue = Replace(strValue, "[" & Split(strData, "|")(l) & "]", strTmp)
                                                    Next
                                                End If
                                                
                                                '�ٴ��������:[ҳ��]��[ҳ��]��[=������]��[n>=0]��[���ڸ�ʽ��]��[��λ����]
                                                strValue = GetLabelMacro(frmSource, strValue)
                                                
                                                '�����Ԫ��
                                                SearchCell objHead, i, j, objHead.FixedRows - 1, .ColE, W, H, strSkip, strSkip2
                                                
                                                strBdr = "1111"
                                                If Not objItem.�߿� Then
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
                                                                , True, sngScale, , objItem.����߼Ӵ�) Then
                                                        Exit Function  '�ϲ�ʱ������
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
                                
                                '�����߶�ʱ,ֻ�账��һ������
                                If blnMeasure Then Exit For
                            Next
                        End If
                        
                        '�������ݵ�Ԫ
                        If objItem.���� = 5 Then
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
                                                If Not objItem.�߿� And (j = .ColE Or (W > objBody.ColWidth(j) + 15 _
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
                                            '���ܱ��-����
                                            If Not DrawCell(objOut, objBody.Text, .X + X, .Y + Y, W, H, .X + .W, _
                                                        , objBody.GridColor, lngForeColor, lngBackColor _
                                                        , objFont, strBdr _
                                                        , GetHscAlign(objBody.Cell(flexcpAlignment, i, j), objBody.Text, objBody, j) _
                                                        , GetVscAlign(objBody.Cell(flexcpAlignment, i, j)) _
                                                        , objItem.�Ե�, sngScale, , objItem.����߼Ӵ�) Then
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
                            '�жϸ��ӱ��
                            lngRowTotal = 0
                            blnAddition = HaveAdditionTable(frmSource.mobjReport, objItem)
                            For k = 1 To .Copys '�������������ݷ���
                                strSkip = "": strSkip2 = ""
                                
                                Y = .FixH
                                
                                If frmSource.mobjReport.Ʊ�� _
                                    And (UBound(arrPage) = intPage Or GridAtCard(frmSource.mobjReport, objItem.id)) Then
                                    '��ʵ���з�Χ���ܸ�
                                    lngRangeHeight = 0
                                    lngRowCount = 0
                                    For i = .RowB + lngRowTotal To .RowE
                                        If i > objBody.Rows - 1 Then
                                            lngRowHeight = objItem.�и�
                                        Else
                                            lngRowHeight = objBody.RowHeight(i)
                                        End If
                                        
                                        If lngRangeHeight + lngRowHeight >= objBody.Height Then
                                            If lngRowHeight >= objBody.Height And lngRangeHeight = 0 Then
                                                '��ʵ���иߴ��ڱ�׼�иߣ�����lngRangeHeightΪ0ʱ��ǿ��һ��
                                                lngRowCount = 1
                                            End If
                                            Exit For
                                        Else
                                            lngRangeHeight = lngRangeHeight + lngRowHeight
                                            lngRowCount = lngRowCount + 1
                                        End If
                                    Next
                                    lngRowTotal = lngRowTotal + lngRowCount
                                                    
                                    '�����ڵ�ǰҳ���������
                                    LngRows = 0
                                    If objBody.Height > lngRangeHeight Then
                                        '������廹�����ɶ��ٱ�׼����
                                        Do While objBody.Height - lngRangeHeight > objItem.�и�
                                            LngRows = LngRows + 1
                                            lngRangeHeight = lngRangeHeight + objItem.�и�
                                        Loop
                                        LngRows = LngRows + lngRowCount
                                    ElseIf objBody.Height = lngRangeHeight Then
                                        lngRowCount = lngRowCount - 1
                                    Else
                                        '����ʵ�����ܸ�����Ʊ��ߵĲ�����ɶ��ٱ�׼����
                                        Do While lngRangeHeight - objBody.Height >= objItem.�и�
                                            LngRows = LngRows + 1
                                            lngRangeHeight = lngRangeHeight - objItem.�и�
                                        Loop
                                    End If
                                    If .Copys > 1 Then
                                        'Ʊ�ݣ�������
                                        If LngRows <= 0 Then
                                            LngRows = lngRowCount
                                        End If
                                    Else
                                        'Ʊ�ݣ����Ƕ���
                                        If LngRows <= .RowE - .RowB Then
                                            LngRows = lngRowCount
                                        End If
                                    End If
                                Else
                                    'ȷ��ÿ���ڵ���ֹ�з�Χ
                                    LngRows = (IIF(.VRowE <> 0, .VRowE, .RowE) - .RowB + 1) / .Copys 'ÿ��Ӧ�������
                                End If
                                
                                If k > 1 Then
                                    lngRowB = lngRowE + 1
                                Else
                                    lngRowB = .RowB + LngRows * (k - 1)
                                End If
                                lngRowE = lngRowB + LngRows - 1
                                
                                For i = lngRowB To lngRowE
                                    If i > .RowE Then
                                        '������ڸ��ӱ�����������л���������ص���������������������
                                        If blnAddition Then Exit For
                                        
                                        '�жϵ�ǰ���ӱ��ײ��Ƿ�����������ӱ��Ʊ�ݸ�ʽ��
                                        If frmSource.mobjReport.Ʊ�� And IsBottomAdditionGrid(frmSource.mobjReport.Items, objItem) Then
                                            Exit For
                                        End If
                                        
                                        '�������������������RowE��ͬ��Ϊ����
                                        If Not blnMeasure Then
                                            X = .W * (k - 1)
                                            H = objItem.�и�
                                            B = 0
                                            For j = .ColB To .ColE
                                                W = objBody.ColWidth(j)
    
                                                If blnPressWork Then
                                                    strBdr = "0000"
                                                Else
                                                    strBdr = "1111"
                                                    If Not objItem.�߿� Then
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
                                                                , objItem.�Ե�, sngScale, , objItem.����߼Ӵ�) Then
                                                        Exit Function
                                                    End If
                                                    B = B + 1
                                                End If
    
                                                X = X + objBody.ColWidth(j)
                                            Next
                                        End If
                                        
                                        Y = Y + objItem.�и�
                                    ElseIf i <= objBody.Rows - 1 Then
                                        '��δ��ҳ
                                        If GetGridStyle(frmSource.mobjReport, objBody.Index) = Val("1-ֻ�б�ͷ���ޱ���") Then
                                            '���Ա������ֶ����õ��������
                                            Exit For
                                        End If
                                        
                                        '���������
                                        If Not blnMeasure Then
                                            '�����ӡ���¼���������������Ҫ��ӡʱ
                                            If Not objCurDLL Is Nothing And TypeName(objOut) = "Printer" Then
                                                For j = objBody.FixedCols To objBody.Cols - 1
                                                    If objBody.ColWidth(j) <> 0 Then
                                                        If objBody.TextMatrix(i, j) <> "" Then Exit For
                                                    End If
                                                Next
                                                If j <= objBody.Cols - 1 Then
                                                    Call objCurDLL.Act_PrintSheetRow( _
                                                            frmSource.mobjReport.��� _
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
                                                        If Not objItem.�߿� Then
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
                                                            '��֣�ÿ����Ԫ��ͼƬ����Ϊ��
                                                            If objPic.handle = 0 Then Set objPic = Nothing
                                                        End If
                                                        If Not objPic Is Nothing Then
                                                            Set objSub = GetSubItem(frmSource, objItem, j)
                                                            
                                                            strData = GetLabelDataName(objSub.����)
                                                            If strData <> "" Then
                                                                strTmp = Split(strData, ".")(0)
                                                                On Error Resume Next
                                                                frmSource.mLibDatas("_" & strTmp).DataSet.AbsolutePosition = i + 1
                                                                Err.Clear
                                                                On Error GoTo 0
                                                            End If
                                                            strValue = GetFieldValue(frmSource, strData)
                                                            If gobjFile.FileExists(strValue) Then
                                                                '�������ֶε���ͼ��
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
                                                                        , True, sngScale, , objItem.����߼Ӵ�) Then
                                                                Exit Function
                                                            End If
                                                        Else
                                                            Set objFont = objBody.Font
                                                            
                                                            '��鲢�����е��Զ�����,���û���,�����ӿ��ٶ�
                                                            If colColAutoFont("_" & j) = "" Then
                                                                colColAutoFont.Remove "_" & j: colColAutoFont.Add "0", "_" & j
                                                                Set objSub = GetSubItem(frmSource, objItem, j)
                                                                If Not objSub Is Nothing Then
                                                                    '���иߣ���С���壩���롰����Ӧ�иߡ�Ϊ�����ϵ
                                                                    If objSub.�и� = 1 Then
                                                                        colColAutoFont.Remove "_" & j: colColAutoFont.Add "1", "_" & j
                                                                    ElseIf objSub.����Ӧ�и� = True Then
                                                                        colColAutoFont.Remove "_" & j: colColAutoFont.Add "2", "_" & j
                                                                    End If
                                                                End If
                                                            End If
                                                            Select Case colColAutoFont("_" & j)
                                                            Case "1"    '��С����
                                                                Set objFont = GetAutoFont(objBody.Text, W, H, objFont, objOut, objItem.�Ե�)
                                                            Case "2"    '����Ӧ�и�
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
                                                            '���ɱ��-����
                                                            If Not DrawCell(objOut, objBody.Text, .X + X, .Y + Y, W, H, .X + .W * .Copys, _
                                                                        , objBody.GridColor, lngForeColor, lngBackColor _
                                                                        , objFont, strBdr _
                                                                        , GetHscAlign(objBody.Cell(flexcpAlignment, i, j), objBody.Text, objBody, j) _
                                                                        , GetVscAlign(objBody.ColAlignment(j)) _
                                                                        , objItem.�Ե�, sngScale, , objItem.����߼Ӵ� _
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
                                
                                '�����߶�ʱ,ֻ�账��һ������
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
    
    '����Ǳ������
    If Not blnPressWork Then
        For Each objItem In frmSource.mobjReport.Items
            If objItem.��ʽ�� = frmSource.bytFormat Then
                Set objFont = New StdFont
                With objItem
                    blnWithData = False
                    If objItem.��ID <> 0 Then
                        lngChildX = frmSource.mobjReport.Items("_" & objItem.��ID).X
                        lngChildY = frmSource.mobjReport.Items("_" & objItem.��ID).Y
                        If frmSource.mobjReport.Items("_" & objItem.��ID).����Դ <> "" Then
                            blnWithData = True
                        End If
                    Else
                        lngChildX = 0
                        lngChildY = 0
                    End If
                    '����Ƕ�̬��ӡ�Ŀ�Ƭ�ڵ����ݣ����ں����ӡ
                    If blnWithData = False Then
                        Select Case .����
                            Case 10 '����
                                If Not blnMeasure Then
                                    If Not DrawCell(objOut, -1, .X + lngChildX, .Y + lngChildY, .W, .H, lngPaperW, lngPaperH, 0, .ǰ��, , , , , , , sngScale, , .����, IIF(.�߿�, 1, 0)) Then Exit Function
                                Else
                                    lngCurH = .Y + .H
                                    If lngCurH > lngPaperH Then lngCurH = lngPaperH
                                    If lngCurH > lngMaxH Then lngMaxH = lngCurH
                                End If
                            Case 11 'ͼƬ
                                Set objPic = LoadPictureFromPar(frmSource, .����)
                                If objPic Is Nothing Then Set objPic = .ͼƬ
                                If .�Ե� And Not objPic Is Nothing Then
                                    .W = objPic.Width * (15 / 26.46)
                                    .H = objPic.Height * (15 / 26.46)
                                End If
                                If Not blnMeasure Then
                                    If Not DrawCell(objOut, objPic, .X + lngChildX, .Y + lngChildY, .W, .H, lngPaperW, lngPaperH, 0, .ǰ��, , , IIF(.�߿�, "1111", "0000"), 0, 0, .����, sngScale) Then Exit Function
                                Else
                                    lngCurH = .Y + .H
                                    If lngCurH > lngPaperH Then lngCurH = lngPaperH
                                    If lngCurH > lngMaxH Then lngMaxH = lngCurH
                                End If
                            Case 14 '��Ƭ
                                If Not blnMeasure Then
                                    If Not DrawCell(objOut, objPic, .X + lngChildX, .Y + lngChildY, .W, .H, lngPaperW, lngPaperH, 0, .ǰ��, , , IIF(.�߿�, "1111", "0000"), 0, 0, .����, sngScale) Then Exit Function
                                Else
                                    lngCurH = .Y + .H
                                    If lngCurH > lngPaperH Then lngCurH = lngPaperH
                                    If lngCurH > lngMaxH Then lngMaxH = lngCurH
                                End If
                            Case 13 '����
                                '��ȡ��������
                                strValue = .����
                                '����ҳ����
                                If InStr(strValue, "[ҳ��]") > 0 Then
                                    strValue = Replace(strValue, "[ҳ��]", intPage + intBasePage)
                                End If
                                If InStr(strValue, "[ҳ��]") > 0 Then
                                    If Not IsArray(arrPage) Then
                                        strValue = Replace(strValue, "[ҳ��]", intBasePage)
                                    Else
                                        strValue = Replace(strValue, "[ҳ��]", UBound(arrPage) + intBasePage)
                                    End If
                                End If
                                If InStr(strValue, "[Ʊ�ݺ�]") > 0 Then
                                    If IsArray(garrBill) Then
                                        If UBound(garrBill) >= intPage Then
                                            strValue = Replace(strValue, "[Ʊ�ݺ�]", garrBill(intPage))
                                        Else
                                            strValue = Replace(strValue, "[Ʊ�ݺ�]", "")
                                        End If
                                    Else
                                        strValue = Replace(strValue, "[Ʊ�ݺ�]", "")
                                    End If
                                End If
                                
                                '����ָ�븴λ(�����õ��������Դ������ֶ�)
                                strData = GetLabelDataName(strValue) '"����Դ.�ֶ�"��
                                If strData <> "" Then
                                    For i = 0 To UBound(Split(strData, "|"))
                                        strTmp = Split(Split(strData, "|")(i), ".")(0)
                                        
                                        If frmSource.mLibDatas("_" & strTmp).DataSet.RecordCount > 0 Then
                                            frmSource.mLibDatas("_" & strTmp).DataSet.MoveFirst
                                            '(��ǰҳ-1)��ʾҪѭ��Move�Ĵ���
                                            For j = 1 To intPage
                                                If Not frmSource.mLibDatas("_" & strTmp).DataSet.EOF Then
                                                    frmSource.mLibDatas("_" & strTmp).DataSet.MoveNext
                                                End If
                                                If frmSource.mLibDatas("_" & strTmp).DataSet.EOF Then
                                                    frmSource.mLibDatas("_" & strTmp).DataSet.MoveFirst
                                                End If
                                            Next
                                            If .Դ�к� <> 0 Then
                                                If .Դ�к� <= frmSource.mLibDatas("_" & strTmp).DataSet.RecordCount Then
                                                    frmSource.mLibDatas("_" & strTmp).DataSet.AbsolutePosition = .Դ�к�
                                                End If
                                            End If
                                        End If
                                    Next
                                End If
                                
                                '�ȴ��������ֶ�(��ѯʱֻȡ��һ��ֵ)
                                If strData <> "" Then
                                    For i = 0 To UBound(Split(strData, "|"))
                                        strTmp = GetFieldValue(frmSource, CStr(Split(strData, "|")(i)))
                                        If .��ʽ <> "" Then
                                            On Error Resume Next
                                            strTmp = Format(strTmp, .��ʽ)
                                            If Err.Number <> 0 Then Err.Clear
                                            On Error GoTo 0
                                        End If
                                        strValue = Replace(strValue, "[" & Split(strData, "|")(i) & "]", strTmp)
                                    Next
                                End If
                                
                                '�ٴ��������:[ҳ��]��[ҳ��]��[=������]��[n>=0]��[���ڸ�ʽ��]��[��λ����]
                                strValue = GetLabelMacro(frmSource, strValue)
                                
                                '��ȡ����ͼ��
                                Set objPic = Nothing
                                If strValue <> "" Then
                                    Unload frmFlash 'ǿ�Ƴ�ʼPicture����Ȼ�л�����������
                                    If .��� = 1 Then
                                        Set objPic = DrawBarCode128(frmFlash.picTemp, 3, strValue, Mid(.��ͷ, 1, 1) = "1")
                                    ElseIf .��� = 2 Then
                                        Set objPic = DrawBarCode39(frmFlash.picTemp, 3, strValue, Mid(.��ͷ, 2, 1) = "1", Mid(.��ͷ, 1, 1) = "1")
                                    ElseIf .��� = 3 Then
                                        Set objPic = DrawBarCode128Auto(frmFlash.picTemp, strValue, sngWidth, .�и�, Mid(.��ͷ, 1, 1) = "1")
                                    ElseIf .��� = 10 Then
                                        Set objPic = DrawBarCode2D(strValue, frmFlash.picTemp, lngSize)
                                    End If
                                    If Val(Mid(.��ͷ, 3, 1)) <> 0 Then
                                        Set objPic = PictureSpin(objPic, Val(Mid(.��ͷ, 3, 1)), frmFlash.picTemp)
                                    End If
                                    
                                    If .��� = 3 And blnPrint = False Then
                                        '128���Զ��������
                                        If Val(Mid(.��ͷ, 3, 1)) = 0 Then
                                            .W = objOut.ScaleX(sngWidth, vbMillimeters, vbTwips)
                                        Else
                                            .H = objOut.ScaleY(sngWidth, vbMillimeters, vbTwips)
                                        End If
                                    ElseIf .��� = 10 And .�Ե� Then
                                        '��ά����ȱʡ�Զ�������С
                                        .W = lngSize: .H = lngSize
                                    End If
                                End If
                                
                                '���ͼ��
                                If Not blnMeasure Then
                                    If Not DrawCell(objOut, objPic, .X + lngChildX, .Y + lngChildY, .W, .H, lngPaperW, lngPaperH, 0, , , , IIF(.�߿�, "1111", "0000"), , , , sngScale) Then Exit Function
                                Else
                                    lngCurH = .Y + .H
                                    If lngCurH > lngPaperH Then lngCurH = lngPaperH
                                    If lngCurH > lngMaxH Then lngMaxH = lngCurH
                                End If
                            Case 1 '����
                                If Not blnMeasure Then
                                    If Not DrawCell(objOut, 1, .X + lngChildX, .Y + lngChildY, .W, .H, lngPaperW, lngPaperH, .ǰ��, .ǰ��, , , , , , , sngScale, , .����) Then Exit Function
                                Else
                                    lngCurH = .Y + .H
                                    If lngCurH > lngPaperH Then lngCurH = lngPaperH
                                    If lngCurH > lngMaxH Then lngMaxH = lngCurH
                                End If
                            Case 12 'ͼ��@@@
                                If intPage = 0 Then 'ֻ�ڵ�һҳ��ӡ
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
                                            
                                            strTmp = GetChartFileFromPar(frmSource, .����)
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
                                                Call GetChartDataName(objItem.����, , , , strTmp)
                                                If strTmp <> "" Then
                                                    Set objPic = GetChartPicture(frmSource.Chart(9999), frmSource.Chart(.id), objItem, frmSource.mLibDatas("_" & strTmp).DataSet, sngScale)
                                                Else
                                                    Set objPic = GetChartPicture(frmSource.Chart(9999), frmSource.Chart(.id), objItem, , sngScale)
                                                End If
                                            End If
                                            
                                            Unload frmSource.Chart(9999)
                                        End If
                                    
                                        If Not DrawCell(objOut, objPic, .X + lngChildX, .Y + lngChildY, .W, .H, lngPaperW, lngPaperH, , , , , IIF(.�߿�, "1111", "0000"), , , , sngScale) Then Exit Function
                                    Else
                                        lngCurH = .Y + .H
                                        If lngCurH > lngPaperH Then lngCurH = lngPaperH
                                        If lngCurH > lngMaxH Then lngMaxH = lngCurH
                                    End If
                                End If
                            Case 2, 3 '��ǩ,��ǩ��ͼƬ
                                objFont.name = .����
                                objFont.Size = .�ֺ�
                                objFont.Bold = .����
                                objFont.Italic = .б��
                                objFont.Underline = .����
                                
                                strValue = .����
                                '����ҳ����
                                If InStr(strValue, "[ҳ��]") > 0 Then
                                    strValue = Replace(strValue, "[ҳ��]", intPage + intBasePage)
                                End If
                                If InStr(strValue, "[ҳ��]") > 0 Then
                                    If Not IsArray(arrPage) Then
                                        strValue = Replace(strValue, "[ҳ��]", intBasePage)
                                    Else
                                        strValue = Replace(strValue, "[ҳ��]", UBound(arrPage) + intBasePage)
                                    End If
                                End If
                                If InStr(strValue, "[Ʊ�ݺ�]") > 0 Then
                                    If IsArray(garrBill) Then
                                        If UBound(garrBill) >= intPage Then
                                            strValue = Replace(strValue, "[Ʊ�ݺ�]", garrBill(intPage))
                                        Else
                                            strValue = Replace(strValue, "[Ʊ�ݺ�]", "")
                                        End If
                                    Else
                                        strValue = Replace(strValue, "[Ʊ�ݺ�]", "")
                                    End If
                                End If
                                
                                '����ָ�븴λ(�����õ��������Դ������ֶ�)
                                strData = GetLabelDataName(strValue) '"����Դ.�ֶ�"��
                                If strData <> "" Then
                                    For i = 0 To UBound(Split(strData, "|"))
                                        strTmp = Split(Split(strData, "|")(i), ".")(0)
                                        
                                        If frmSource.mLibDatas("_" & strTmp).DataSet.RecordCount > 0 Then
                                            frmSource.mLibDatas("_" & strTmp).DataSet.MoveFirst
                                            '(��ǰҳ-1)��ʾҪѭ��Move�Ĵ���
                                            For j = 1 To intPage
                                                If Not frmSource.mLibDatas("_" & strTmp).DataSet.EOF Then
                                                    frmSource.mLibDatas("_" & strTmp).DataSet.MoveNext
                                                End If
                                                If frmSource.mLibDatas("_" & strTmp).DataSet.EOF Then
                                                    frmSource.mLibDatas("_" & strTmp).DataSet.MoveFirst
                                                End If
                                            Next

                                            If .Դ�к� <> 0 Then
                                                If .Դ�к� <= frmSource.mLibDatas("_" & strTmp).DataSet.RecordCount Then
                                                    frmSource.mLibDatas("_" & strTmp).DataSet.AbsolutePosition = .Դ�к�
                                                End If
                                            End If
                                        End If
                                        
                                        '�ȴ��������ֶ�(��ѯʱֻȡ��һ��ֵ)
                                        strTmp = GetFieldValue(frmSource, CStr(Split(strData, "|")(i)))
                                        If .��ʽ <> "" Then
                                            On Error Resume Next
                                            strTmp = Format(strTmp, .��ʽ)
                                            If Err.Number <> 0 Then Err.Clear
                                            On Error GoTo 0
                                        End If
                                        strValue = Replace(strValue, "[" & Split(strData, "|")(i) & "]", strTmp)
                                    Next
                                End If
                                
                                '�ٴ��������:[ҳ��]��[ҳ��]��[=������]��[n>=0]��[���ڸ�ʽ��]��[��λ����]
                                strValue = GetLabelMacro(frmSource, strValue)
                                If gobjFile.FileExists(strValue) Then
                                    '�������ֶε���ͼ��
                                    On Error Resume Next
                                    Set .ͼƬ = LoadPicture(strValue)
                                    Kill strValue
                                    Err.Clear
                                    On Error GoTo 0
                                    
                                    If .�Ե� And Not .ͼƬ Is Nothing Then
                                        .W = .ͼƬ.Width * (15 / 26.46)
                                        .H = .ͼƬ.Height * (15 / 26.46)
                                    End If
                                    
                                    If Not blnMeasure Then
                                        If Not DrawCell(objOut, .ͼƬ, .X + lngChildX, .Y + lngChildY, .W, .H, lngPaperW, lngPaperH, 0, .ǰ��, , , IIF(.�߿�, "1111", "0000"), 0, 0, .����, sngScale) Then Exit Function
                                    Else
                                        lngCurH = .Y + .H
                                        If lngCurH > lngPaperH Then lngCurH = lngPaperH
                                        If lngCurH > lngMaxH Then lngMaxH = lngCurH
                                    End If
                                Else
                                    If .�Ե� Then Call ItemAutoSize(objItem, strValue, objOut)
                                    If objItem.���� > 0 And objItem.���� <> "" And blnHaveGrid Then
                                        '���㿿���ǩ��λ��
                                        strDepend = GetDependIDs(.����, frmSource)
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
        
                                        '���ҿ���
                                        Select Case .����
                                            Case 11, 21 '��
                                                lngOutX = lngOX
                                            Case 12, 22 '��
                                                lngOutX = lngOX + (lngOW - .W) / 2
                                            Case 13, 23 '��
                                                lngOutX = lngOX + lngOW - .W
                                        End Select
                                        '���¿���
                                        If frmSource.mobjReport.Ʊ�� Then
                                            lngOutY = .Y 'Ʊ��ʱλ��Ӧ�ò���
                                        Else
                                            If CInt(Left(CStr(.����), 1)) = 2 Then
                                                lngOutY = lngOY + lngOH + (.Y - (lngOY + lngDesignH))
                                            Else
                                                lngOutY = .Y
                                            End If
                                        End If
                                        If strValue <> "" Then
                                            If Not blnMeasure Then
                                                If .�и� = 1 Then Set objFont = GetAutoFont(strValue, .W, .H, objFont, objOut, True, .����)
                                                If .ˮƽ��ת Then
                                                    If Not DrawCell(objOut, frmSource.picRotate(.id).Image, .X + lngChildX, .Y + lngChildY, .W, .H, lngPaperW, lngPaperH, 0, .ǰ��, .����, objFont, "0000", .����, 0, True, sngScale, .����) Then Exit Function
                                                Else
                                                    If Not DrawCell(objOut, strValue, lngOutX, lngOutY, .W, .H, lngPaperW, lngPaperH, 0, .ǰ��, .����, objFont, IIF(.�߿�, "1111", "0000"), .����, 0, True, sngScale, .����) Then Exit Function
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
                                                If .�и� = 1 Then Set objFont = GetAutoFont(strValue, .W, .H, objFont, objOut, True, .����)
                                                If .ˮƽ��ת Then
                                                    If Not DrawCell(objOut, frmSource.picRotate(.id).Image, .X + lngChildX, .Y + lngChildY, .W, .H, lngPaperW, lngPaperH, 0, .ǰ��, .����, objFont, "0000", .����, 0, True, sngScale, .����) Then Exit Function
                                                Else
                                                    If Not DrawCell(objOut, strValue, .X + lngChildX, .Y + lngChildY, .W, .H, lngPaperW, lngPaperH, 0, .ǰ��, .����, objFont, IIF(.�߿�, "1111", "0000"), .����, 0, True, sngScale, .����) Then Exit Function
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
                            '�������Ƭ����
                            On Error Resume Next
                            Set objTemp = frmSource.mobjReport.Items("_" & objPageCard.id)
                            If Err.Number <> 0 Then
                                On Error GoTo 0
                                Exit For
                            End If
                            On Error GoTo 0
                            If objTemp Is Nothing Then Exit Function
                            
                            '���ܿ�Ƭ������ڿ�Ƭ��Ķ��󴴽�����ˣ��������Ƭ�������Ƭ��Ķ���
                            '�����Ƭ
                            With objTemp
                                Set objPic = LoadPictureFromPar(frmSource, .����)
                                If lngCol >= objPageCard.Col Then lngRow = lngRow + 1: lngCol = 0
                                lngX = lngRow * (.H + .���¼��)
                                lngY = lngCol * (.W + .���Ҽ��)
                                If Not blnMeasure Then
                                    If Not DrawCell(objOut, objPic, .X + lngY, .Y + lngX, .W, .H, lngPaperW, lngPaperH, 0 _
                                                , .ǰ��, , , IIF(.�߿�, "1111", "0000"), 0, 0, .����, sngScale) Then
                                        Exit Function
                                    End If
                                Else
                                    lngCurH = .Y + .H
                                    If lngCurH > lngPaperH Then lngCurH = lngPaperH
                                    If lngCurH > lngMaxH Then lngMaxH = lngCurH
                                End If
                                lngCol = lngCol + 1
                            End With
                        
                            '�������Ƭ��Ķ���
                            For Each objItem In frmSource.mobjReport.Items
                                If objItem.��ʽ�� = frmSource.bytFormat Then
                                    Set objFont = New StdFont
                                    With objItem
                                        If .��ID <> 0 Then
                                            lngChildX = frmSource.mobjReport.Items("_" & .��ID).X
                                            lngChildY = frmSource.mobjReport.Items("_" & .��ID).Y
                                        Else
                                            lngChildX = 0
                                            lngChildY = 0
                                        End If
                                        
                                        If .��ID = objPageCard.id Then
                                            '�������Ƭ�е�����
                                            Select Case .����
                                            Case 10 '����
                                                If Not blnMeasure Then
                                                    If Not DrawCell(objOut, -1, .X + lngY + lngChildX, .Y + lngX + lngChildY, .W, .H, lngPaperW, lngPaperH, 0, .ǰ��, , , , , , , sngScale, , .����, IIF(.�߿�, 1, 0)) Then Exit Function
                                                Else
                                                    lngCurH = .Y + .H
                                                    If lngCurH > lngPaperH Then lngCurH = lngPaperH
                                                    If lngCurH > lngMaxH Then lngMaxH = lngCurH
                                                End If
                                            Case 1 '����
                                                If Not blnMeasure Then
                                                    If Not DrawCell(objOut, 1, .X + lngY + lngChildX, .Y + lngX + lngChildY, .W, .H, lngPaperW, lngPaperH, .ǰ��, .ǰ��, , , , , , , sngScale, , .����) Then Exit Function
                                                Else
                                                    lngCurH = .Y + .H
                                                    If lngCurH > lngPaperH Then lngCurH = lngPaperH
                                                    If lngCurH > lngMaxH Then lngMaxH = lngCurH
                                                End If
                                            Case 11 'ͼƬ
                                                Set objPic = LoadPictureFromPar(frmSource, .����)
                                                If objPic Is Nothing Then Set objPic = .ͼƬ
                                                If .�Ե� And Not objPic Is Nothing Then
                                                    .W = objPic.Width * (15 / 26.46)
                                                    .H = objPic.Height * (15 / 26.46)
                                                End If
                                                If Not blnMeasure Then
                                                    If Not DrawCell(objOut, objPic, .X + lngY + lngChildX, .Y + lngX + lngChildY, .W, .H, lngPaperW, lngPaperH, 0, .ǰ��, , , IIF(.�߿�, "1111", "0000"), 0, 0, .����, sngScale) Then Exit Function
                                                Else
                                                    lngCurH = .Y + .H
                                                    If lngCurH > lngPaperH Then lngCurH = lngPaperH
                                                    If lngCurH > lngMaxH Then lngMaxH = lngCurH
                                                End If
                                            Case 13 '����
                                                '��ȡ��������
                                                strValue = .����
                                                '����ҳ����
                                                If InStr(strValue, "[ҳ��]") > 0 Then
                                                    strValue = Replace(strValue, "[ҳ��]", intPage + intBasePage)
                                                End If
                                                If InStr(strValue, "[ҳ��]") > 0 Then
                                                    If Not IsArray(arrPage) Then
                                                        strValue = Replace(strValue, "[ҳ��]", intBasePage)
                                                    Else
                                                        strValue = Replace(strValue, "[ҳ��]", UBound(arrPage) + intBasePage)
                                                    End If
                                                End If
                                                If InStr(strValue, "[Ʊ�ݺ�]") > 0 Then
                                                    If IsArray(garrBill) Then
                                                        If UBound(garrBill) >= intPage Then
                                                            strValue = Replace(strValue, "[Ʊ�ݺ�]", garrBill(intPage))
                                                        Else
                                                            strValue = Replace(strValue, "[Ʊ�ݺ�]", "")
                                                        End If
                                                    Else
                                                        strValue = Replace(strValue, "[Ʊ�ݺ�]", "")
                                                    End If
                                                End If
                                                
                                                '����ָ�븴λ(�����õ��������Դ������ֶ�)
                                                strData = GetLabelDataName(strValue) '"����Դ.�ֶ�"��
                                                If strData <> "" Then
                                                    For i = 0 To UBound(Split(strData, "|"))
                                                        strTmp = Split(Split(strData, "|")(i), ".")(0)
                                                        
                                                        If frmSource.mLibDatas("_" & strTmp).DataSet.RecordCount > 0 Then
                                                            frmSource.mLibDatas("_" & strTmp).DataSet.MoveFirst
                                                            '(��ǰҳ-1)��ʾҪѭ��Move�Ĵ���
                                                            For j = 1 To intPage
                                                                If Not frmSource.mLibDatas("_" & strTmp).DataSet.EOF Then
                                                                    frmSource.mLibDatas("_" & strTmp).DataSet.MoveNext
                                                                End If
                                                                If frmSource.mLibDatas("_" & strTmp).DataSet.EOF Then
                                                                    frmSource.mLibDatas("_" & strTmp).DataSet.MoveFirst
                                                                End If
                                                            Next
                                                            If frmSource.mobjReport.Items("_" & .��ID).����Դ = strTmp Then
                                                                blnGroup = False
                                                                On Error Resume Next
                                                                If frmSource.mLibDatas("_" & strTmp).DataSet!�����ʶ & "" <> "" Or frmSource.mLibDatas("_" & strTmp).DataSet!�����ʶ & "" = "" Then
                                                                    If Err.Number = 0 Then
                                                                        '���鶯̬��ӡ
                                                                        blnGroup = True
                                                                    End If
                                                                    Err.Clear: On Error GoTo 0
                                                                    If arrPageCard(intPage).count > 0 Then
                                                                        frmSource.mLibDatas("_" & strTmp).DataSet.AbsolutePosition = Val(Mid(objPageCard.Item(Y), 1, InStr(objPageCard.Item(Y), "-") - 1))
                                                                    End If
                                                                End If
                                                            End If
                                                            If .Դ�к� <> 0 Then
                                                                If blnGroup Then
                                                                    '���鶯̬��ӡ
                                                                    If .Դ�к� <= Val(Mid(objPageCard.Item(Y), InStr(objPageCard.Item(Y), "-") + 1, Len(objPageCard.Item(Y)))) - Val(Mid(objPageCard.Item(Y), 1, InStr(objPageCard.Item(Y), "-") - 1)) + 1 Then
                                                                        frmSource.mLibDatas("_" & strTmp).DataSet.AbsolutePosition = frmSource.mLibDatas("_" & strTmp).DataSet.AbsolutePosition + .Դ�к� - 1
                                                                    End If
                                                                Else
                                                                    If .Դ�к� <= frmSource.mLibDatas("_" & strTmp).DataSet.RecordCount Then
                                                                        frmSource.mLibDatas("_" & strTmp).DataSet.AbsolutePosition = .Դ�к�
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    Next
                                                End If
                                                
                                                '�ȴ��������ֶ�(��ѯʱֻȡ��һ��ֵ)
                                                If strData <> "" Then
                                                    For i = 0 To UBound(Split(strData, "|"))
                                                        strTmp = GetFieldValue(frmSource, CStr(Split(strData, "|")(i)))
                                                        If .��ʽ <> "" Then
                                                            On Error Resume Next
                                                            strTmp = Format(strTmp, .��ʽ)
                                                            If Err.Number <> 0 Then Err.Clear
                                                            On Error GoTo 0
                                                        End If
                                                        strValue = Replace(strValue, "[" & Split(strData, "|")(i) & "]", strTmp)
                                                    Next
                                                End If
                                                
                                                '�ٴ��������:[ҳ��]��[ҳ��]��[=������]��[n>=0]��[���ڸ�ʽ��]��[��λ����]
                                                strValue = GetLabelMacro(frmSource, strValue)
                                                
                                                '��ȡ����ͼ��
                                                Set objPic = Nothing
                                                If strValue <> "" Then
                                                    Unload frmFlash 'ǿ�Ƴ�ʼPicture����Ȼ�л�����������
                                                    If .��� = 1 Then
                                                        Set objPic = DrawBarCode128(frmFlash.picTemp, 3, strValue, Mid(.��ͷ, 1, 1) = "1")
                                                    ElseIf .��� = 2 Then
                                                        Set objPic = DrawBarCode39(frmFlash.picTemp, 3, strValue, Mid(.��ͷ, 2, 1) = "1", Mid(.��ͷ, 1, 1) = "1")
                                                    ElseIf .��� = 3 Then
                                                        Set objPic = DrawBarCode128Auto(frmFlash.picTemp, strValue, sngWidth, .�и�, Mid(.��ͷ, 1, 1) = "1")
                                                    ElseIf .��� = 10 Then
                                                        Set objPic = DrawBarCode2D(strValue, frmFlash.picTemp, lngSize)
                                                    End If
                                                    If Val(Mid(.��ͷ, 3, 1)) <> 0 Then
                                                        Set objPic = PictureSpin(objPic, Val(Mid(.��ͷ, 3, 1)), frmFlash.picTemp)
                                                    End If
                                                    
                                                    If .��� = 3 Then
                                                        '128���Զ��������
                                                        If Val(Mid(.��ͷ, 3, 1)) = 0 Then
                                                            .W = objOut.ScaleX(sngWidth, vbMillimeters, vbTwips)
                                                        Else
                                                            .H = objOut.ScaleY(sngWidth, vbMillimeters, vbTwips)
                                                        End If
                                                    ElseIf .��� = 10 And .�Ե� Then
                                                        '��ά����ȱʡ�Զ�������С
                                                        .W = lngSize: .H = lngSize
                                                    End If
                                                End If
                                                
                                                '���ͼ��
                                                If Not blnMeasure Then
                                                    If Not DrawCell(objOut, objPic, .X + lngY + lngChildX, .Y + lngX + lngChildY, .W, .H, lngPaperW, lngPaperH, 0, , , , IIF(.�߿�, "1111", "0000"), , , , sngScale) Then Exit Function
                                                Else
                                                    lngCurH = .Y + .H
                                                    If lngCurH > lngPaperH Then lngCurH = lngPaperH
                                                    If lngCurH > lngMaxH Then lngMaxH = lngCurH
                                                End If
                                            Case 2, 3 '��ǩ,��ǩ��ͼƬ
                                                objFont.name = .����
                                                objFont.Size = .�ֺ�
                                                objFont.Bold = .����
                                                objFont.Italic = .б��
                                                objFont.Underline = .����
                                                
                                                strValue = .����
                                                '����ҳ����
                                                If InStr(strValue, "[ҳ��]") > 0 Then
                                                    strValue = Replace(strValue, "[ҳ��]", intPage + intBasePage)
                                                End If
                                                If InStr(strValue, "[ҳ��]") > 0 Then
                                                    If Not IsArray(arrPage) Then
                                                        strValue = Replace(strValue, "[ҳ��]", intBasePage)
                                                    Else
                                                        strValue = Replace(strValue, "[ҳ��]", UBound(arrPage) + intBasePage)
                                                    End If
                                                End If
                                                If InStr(strValue, "[Ʊ�ݺ�]") > 0 Then
                                                    If IsArray(garrBill) Then
                                                        If UBound(garrBill) >= intPage Then
                                                            strValue = Replace(strValue, "[Ʊ�ݺ�]", garrBill(intPage))
                                                        Else
                                                            strValue = Replace(strValue, "[Ʊ�ݺ�]", "")
                                                        End If
                                                    Else
                                                        strValue = Replace(strValue, "[Ʊ�ݺ�]", "")
                                                    End If
                                                End If
                                                
                                                '����ָ�븴λ(�����õ��������Դ������ֶ�)
                                                strData = GetLabelDataName(strValue) '"����Դ.�ֶ�"��
                                                If strData <> "" Then
                                                    For i = 0 To UBound(Split(strData, "|"))
                                                        strTmp = Split(Split(strData, "|")(i), ".")(0)
                                                        
                                                        If frmSource.mLibDatas("_" & strTmp).DataSet.RecordCount > 0 Then
                                                            
                                                            frmSource.mLibDatas("_" & strTmp).DataSet.MoveFirst
                                                            '(��ǰҳ-1)��ʾҪѭ��Move�Ĵ���
                                                            For j = 1 To intPage
                                                                If Not frmSource.mLibDatas("_" & strTmp).DataSet.EOF Then
                                                                    frmSource.mLibDatas("_" & strTmp).DataSet.MoveNext
                                                                End If
                                                                If frmSource.mLibDatas("_" & strTmp).DataSet.EOF Then
                                                                    frmSource.mLibDatas("_" & strTmp).DataSet.MoveFirst
                                                                End If
                                                            Next
                                                            If frmSource.mobjReport.Items("_" & .��ID).����Դ = strTmp Then
                                                                blnGroup = False
                                                                On Error Resume Next
                                                                If frmSource.mLibDatas("_" & strTmp).DataSet!�����ʶ & "" <> "" Or frmSource.mLibDatas("_" & strTmp).DataSet!�����ʶ & "" = "" Then
                                                                    If Err.Number = 0 Then
                                                                        '���鶯̬��ӡ
                                                                        blnGroup = True
                                                                    End If
                                                                    Err.Clear: On Error GoTo 0
                                                                    If arrPageCard(intPage).count > 0 Then
                                                                        frmSource.mLibDatas("_" & strTmp).DataSet.AbsolutePosition = Val(Mid(objPageCard.Item(Y), 1, InStr(objPageCard.Item(Y), "-") - 1))
                                                                    End If
                                                                End If
                                                            End If
                                                            If .Դ�к� <> 0 Then
                                                                If blnGroup Then
                                                                    '���鶯̬��ӡ
                                                                    If .Դ�к� <= Val(Mid(objPageCard.Item(Y), InStr(objPageCard.Item(Y), "-") + 1, Len(objPageCard.Item(Y)))) - Val(Mid(objPageCard.Item(Y), 1, InStr(objPageCard.Item(Y), "-") - 1)) + 1 Then
                                                                        frmSource.mLibDatas("_" & strTmp).DataSet.AbsolutePosition = frmSource.mLibDatas("_" & strTmp).DataSet.AbsolutePosition + .Դ�к� - 1
                                                                    End If
                                                                Else
                                                                    If .Դ�к� <= frmSource.mLibDatas("_" & strTmp).DataSet.RecordCount Then
                                                                        frmSource.mLibDatas("_" & strTmp).DataSet.AbsolutePosition = .Դ�к�
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                        
                                                        '�ȴ��������ֶ�(��ѯʱֻȡ��һ��ֵ)
                                                        strTmp = GetFieldValue(frmSource, CStr(Split(strData, "|")(i)))
                                                        If .��ʽ <> "" Then
                                                            On Error Resume Next
                                                            strTmp = Format(strTmp, .��ʽ)
                                                            If Err.Number <> 0 Then Err.Clear
                                                            On Error GoTo 0
                                                        End If
                                                        strValue = Replace(strValue, "[" & Split(strData, "|")(i) & "]", strTmp)
                                                    Next
                                                End If
                                                
                                                '�ٴ��������:[ҳ��]��[ҳ��]��[=������]��[n>=0]��[���ڸ�ʽ��]��[��λ����]
                                                strValue = GetLabelMacro(frmSource, strValue)
                                                
                                                If gobjFile.FileExists(strValue) Then
                                                    '�������ֶε���ͼ��
                                                    On Error Resume Next
                                                    Set .ͼƬ = LoadPicture(strValue)
                                                    Kill strValue
                                                    Err.Clear
                                                    On Error GoTo 0
                                                    
                                                    If .�Ե� And Not .ͼƬ Is Nothing Then
                                                        .W = .ͼƬ.Width * (15 / 26.46)
                                                        .H = .ͼƬ.Height * (15 / 26.46)
                                                    End If
                                                    
                                                    If Not blnMeasure Then
                                                        If Not DrawCell(objOut, .ͼƬ, .X + lngY + lngChildX, .Y + lngX + lngChildY, .W, .H, lngPaperW, lngPaperH, 0, .ǰ��, , , IIF(.�߿�, "1111", "0000"), 0, 0, .����, sngScale) Then Exit Function
                                                    Else
                                                        lngCurH = .Y + .H
                                                        If lngCurH > lngPaperH Then lngCurH = lngPaperH
                                                        If lngCurH > lngMaxH Then lngMaxH = lngCurH
                                                    End If
                                                Else
                                                    If .�Ե� Then Call ItemAutoSize(objItem, strValue, objOut)
                                                    If objItem.���� > 0 And objItem.���� <> "" And blnHaveGrid Then
                                                        '���㿿���ǩ��λ��
                                                        strDepend = GetDependIDs(.����, frmSource)
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
                        
                                                        '���ҿ���
                                                        Select Case .����
                                                            Case 11, 21 '��
                                                                lngOutX = lngOX
                                                            Case 12, 22 '��
                                                                lngOutX = lngOX + (lngOW - .W) / 2
                                                            Case 13, 23 '��
                                                                lngOutX = lngOX + lngOW - .W
                                                        End Select
                                                        '���¿���
                                                        If frmSource.mobjReport.Ʊ�� Then
                                                            lngOutY = .Y 'Ʊ��ʱλ��Ӧ�ò���
                                                        Else
                                                            If CInt(Left(CStr(.����), 1)) = 2 Then
                                                                lngOutY = lngOY + lngOH + (.Y - (lngOY + lngDesignH))
                                                            Else
                                                                lngOutY = .Y
                                                            End If
                                                        End If
                                                        If strValue <> "" Then
                                                            If Not blnMeasure Then
                                                                If .�и� = 1 Then Set objFont = GetAutoFont(strValue, .W, .H, objFont, objOut, True, .����)
                                                                If .ˮƽ��ת Then
                                                                    If Not DrawCell(objOut, frmSource.picRotate(.id).Image, lngOutX + lngY, lngOutY + lngX + lngChildY, .W, .H, lngPaperW, lngPaperH, 0, .ǰ��, .����, objFont, "0000", .����, 0, True, sngScale, .����) Then Exit Function
                                                                Else
                                                                    If Not DrawCell(objOut, strValue, lngOutX + lngY, lngOutY + lngX + lngChildY, .W, .H, lngPaperW, lngPaperH, 0, .ǰ��, .����, objFont, IIF(.�߿�, "1111", "0000"), .����, 0, True, sngScale, .����) Then Exit Function
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
                                                                If .�и� = 1 Then Set objFont = GetAutoFont(strValue, .W, .H, objFont, objOut, True, .����)
                                                                If .ˮƽ��ת Then
                                                                    If Not DrawCell(objOut, frmSource.picRotate(.id).Image, .X + lngY + lngChildX, .Y + lngX + lngChildY, .W, .H, lngPaperW, lngPaperH, 0, .ǰ��, .����, objFont, "0000", .����, 0, True, sngScale, .����) Then Exit Function
                                                                Else
                                                                    If Not DrawCell(objOut, strValue, .X + lngY + lngChildX, .Y + lngX + lngChildY, .W, .H, lngPaperW, lngPaperH, 0, .ǰ��, .����, objFont, IIF(.�߿�, "1111", "0000"), .����, 0, True, sngScale, .����) Then Exit Function
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
        '��ӡ��ʵ�ʿ��������Ԥ��
        dblSureW = GetDeviceCaps(Printer.hDC, PHYSICALOFFSETX) / GetDeviceCaps(Printer.hDC, PHYSICALWIDTH)
        dblSureH = GetDeviceCaps(Printer.hDC, PHYSICALOFFSETY) / GetDeviceCaps(Printer.hDC, PHYSICALHEIGHT)
        If blnSure Then
            objOut.DrawStyle = 2
            objOut.Line (objOut.Width * dblSureW, objOut.Height * dblSureH)-(objOut.Width * (1 - dblSureW) - Printer.TwipsPerPixelX * sngScale, objOut.Height * (1 - dblSureH) - Printer.TwipsPerPixelY * sngScale), &H808080, B
            objOut.DrawStyle = 0
        End If
        
        '���ñ�־
'        strTmp = Decode(zlRegInfo("��Ȩ����"), "2", "����", "3", "����", "")
'        If strTmp <> "" Then
'            Set objFont = New StdFont
'            objFont.name = "����"
'            objFont.Size = 24 * sngScale
'            objFont.Bold = True
'            objFont.Italic = False
'            objFont.Italic = False
'            If Not DrawCell(objOut, strTmp & "����", objOut.Width * dblSureW + 2 * Printer.TwipsPerPixelX * sngScale, objOut.Height * dblSureH + 2 * Printer.TwipsPerPixelY * sngScale, 2500 * sngScale, 600 * sngScale, , , vbRed, vbRed, , objFont, , 1, 1, , 1) Then Exit Function
'            If Not DrawCell(objOut, strTmp & "����", objOut.Width / 2 - 1250 * sngScale, objOut.Height / 2 - 300 * sngScale, 2500 * sngScale, 600 * sngScale, , , vbRed, vbRed, , objFont, , 1, 1, , 1) Then Exit Function
'            If Not DrawCell(objOut, strTmp & "����", objOut.Width * (1 - dblSureW) - 2500 * sngScale - 2 * Printer.TwipsPerPixelX * sngScale, objOut.Height * (1 - dblSureH) - 600 * sngScale - 2 * Printer.TwipsPerPixelY * sngScale, 2500 * sngScale, 600 * sngScale, , , vbRed, vbRed, , objFont, , 1, 1, , 1) Then Exit Function
'        End If
    End If
    
    PrintPage = True
End Function

Public Function GetScreenFonts() As String
'���ܣ���ȡϵͳ��֧�ֵ�����
    Dim i As Integer, strFont As String
    For i = 0 To Screen.FontCount - 1
        strFont = strFont & "^" & Screen.Fonts(i)
    Next
    GetScreenFonts = Mid(strFont, 2)
End Function

Public Function MatchIndex(ByVal cbo As Object, ByRef KeyAscii As Integer, Optional sngInterval As Single = 1) As Long
'���ܣ�����������ַ����Զ�ƥ��ComboBox��ѡ����,���Զ�ʶ��������
'������cbo.Hwnd=ComboBox��Hwnd����,KeyAscii=ComboBox��KeyPress�¼��е�KeyAscii����,sngInterval=ָ��������
'���أ�-2=δ�Ӵ���,����=ƥ�������(����ƥ�������)
'˵�����뽫�ú�����KeyPress�¼��е��á�

    Static lngPreTime As Single, lngPreHwnd As Long
    Static strFind As String
    Dim sngTime As Single, lngR As Long
    
    If lngPreHwnd <> cbo.hwnd Then lngPreTime = Empty: strFind = Empty
    lngPreHwnd = cbo.hwnd
    
    If KeyAscii <> 13 Then
        sngTime = timer
        If Abs(sngTime - lngPreTime) > sngInterval Then '������(ȱʡΪ0.5��)
            strFind = ""
        End If
        strFind = strFind & Chr(KeyAscii)
        lngPreTime = timer
        KeyAscii = 0 'ʹComboBox����ĵ���ƥ�书��ʧЧ
        MatchIndex = SendMessage(cbo.hwnd, CB_FINDSTRING, -1, ByVal strFind)
        If MatchIndex = -1 Then
            cbo.Text = strFind
            cbo.SelStart = Len(cbo.Text)
        End If
    Else
        MatchIndex = -2 '������Իس���������
    End If
End Function

Public Function ReportReaded(Optional ByVal lng����ID As Long, _
    Optional ByVal varReport As Variant, Optional ByVal lngϵͳ As Long) As Boolean
'���ܣ��жϱ������Ƿ����
'������lng����ID,varReport(��Ż����ID)=�����жϵ�ǰ�������Ƿ���ϵ�����
'      lngϵͳ=�����뱨���Ż����IDʱ��Ҫ,����Ϊ0��ʾ����ϵͳ
    If grsReport Is Nothing Then Exit Function
    If grsReport.State = 0 Then Exit Function
    If grsReport.EOF Or grsReport.BOF Then Exit Function
    
    If Format(grsReport!�޸�ʱ��, "yyyy-MM-dd HH:mm:ss") = Format(gdatModiTime, "yyyy-MM-dd HH:mm:ss") Then
        If lng����ID <> 0 Then
            ReportReaded = grsReport!id = lng����ID
        Else
            If TypeName(varReport) = "String" Then
                ReportReaded = (UCase(grsReport!���) = UCase(varReport) And Nvl(grsReport!ϵͳ, 0) = lngϵͳ)
            Else
                ReportReaded = (Nvl(grsReport!����id, 0) = CLng(varReport) And Nvl(grsReport!ϵͳ, 0) = lngϵͳ)
            End If
        End If
    End If
End Function

Public Function isGroup(ByVal lngSys As Long, ByVal varReport As Variant _
    , ByRef lngID As Long, Optional ByRef strInfo As String) As Boolean
'���ܣ��ж�ָ���ı����ǵ������Ǳ�����
'������varReport=��Ż����ID
'���أ�lngID=�������ID

    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    'ÿ�ζ����������û��棬�Ա㱨������޸�ʱ��ʱ���±���gdatModiTime
    '�Ƿ񱨱�
    If TypeName(varReport) = "String" Then
        strSQL = "Select ID,���,����,˵��,����,��ӡ��,��ֽ,Ʊ��,��ӡ��ʽ,ϵͳ,����ID,����,�޸�ʱ��,����ʱ��" & _
                 "    ,��ֹ��ʼʱ��,��ֹ����ʱ�� " & vbCr & _
                 "From zlReports " & vbCr & _
                 "Where Nvl(ϵͳ,0)=[3] And ���=[1]"
    Else
        strSQL = "Select ID,���,����,˵��,����,��ӡ��,��ֽ,Ʊ��,��ӡ��ʽ,ϵͳ,����ID,����,�޸�ʱ��,����ʱ��" & _
                 "    ,��ֹ��ʼʱ��,��ֹ����ʱ�� " & vbCr & _
                 "From zlReports " & vbCr & _
                 "Where Nvl(ϵͳ,0)=[3] And ����ID=[2]"
    End If
    
    Set rsTmp = OpenSQLRecord(strSQL, "isGroup", UCase(varReport), Val(varReport), lngSys)
    If Not rsTmp.EOF Then
        '���洦��
        Set grsReport = New ADODB.Recordset
        Set grsReport = rsTmp
        gdatModiTime = grsReport!�޸�ʱ��
        
        lngID = rsTmp!id
        strInfo = mdlPublic.FormatString("��[1]��[2]", Nvl(rsTmp!���), Nvl(rsTmp!����))
        Exit Function
    End If
    
    '�Ǳ�����
    If TypeName(varReport) = "String" Then
        strSQL = "Select ID,���,���� From zlRPTGroups Where Nvl(ϵͳ,0)=[3] And Upper(���)=[1]"
    Else
        strSQL = "Select ID,���,���� From zlRPTGroups Where Nvl(ϵͳ,0)=[3] And ����ID=[2]"
    End If
    Set rsTmp = OpenSQLRecord(strSQL, "isGroup", UCase(varReport), Val(varReport), lngSys)
    If Not rsTmp.EOF Then
        lngID = rsTmp!id
        strInfo = mdlPublic.FormatString("��[1]��[2]", Nvl(rsTmp!���), Nvl(rsTmp!����))
    End If
    isGroup = True
    Exit Function
    
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetLenStr(Str As String, lngW As Long, objBase As Object) As String
'���ܣ�����ָ���ĳ��Ƚ�ȡ�ַ���
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
'���ܣ���SQL���������Order by ���ȥ��
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
'���ܣ��жϵ�ǰ�û��Ƿ���Ȩ�޶�ָ��ID�ı�����в�ѯ
'���أ�0-��Ȩ��,1-������Ȩ��,2-Ʊ����Ȩ��,3-�д���
'˵��������ñ���ʱδ�������λ��(ϵͳ,ģ��),����Ķ����Ȩλ�ã�ֻҪ��һ��λ����Ȩ����ʹ��
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
        " Select Ʊ��,ϵͳ,����ID,���� From zlReports" & _
        " Where ����ID is Not Null And ���� Is Not Null And ID=[1]" & _
        " Union ALL" & _
        " Select A.Ʊ��,B.ϵͳ,B.����ID,B.���� From zlReports A,zlRPTPuts B" & _
        " Where A.ID=B.����ID And A.ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "ReportCanQuery", lngRPTID)

    Do While Not rsTmp.EOF
        strPriv = GetPrivFunc(Nvl(rsTmp!ϵͳ, 0), rsTmp!����id)
        If InStr(";" & strPriv & ";", ";" & rsTmp!���� & ";") > 0 Then
            ReportCanQuery = 0: Exit Do
        Else
            ReportCanQuery = IIF(Nvl(rsTmp!Ʊ��, 0) = 0, 1, 2)
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
'���ܣ����ݲ���ѡ����SQL���壬������ʾ�ֶμ����ֶε�ֵ
'������strFld=��������Դ�ֶ�˵����
'      strDefBand=�������ȱʡ��ֵ,�Ƿ񰴴�ֵ����
'      intConnectNo=���ݿ�������ţ�0=ȱʡ��1>=����
'���أ���ʾֵ|��ֵ|ԭʼ��¼��
    Dim rsTmp As New ADODB.Recordset
    Dim strTmp As String, i As Long
    Dim strShow As String, strBand As String
        
    'ȡ����ʾ,���ֶ���
    For i = 0 To UBound(Split(strFld, "|"))
        strTmp = Split(strFld, "|")(i)
        If Split(strTmp, ",")(2) Like "*&D*" Then strShow = CStr(Split(strTmp, ",")(0))
        If Split(strTmp, ",")(2) Like "*&B*" Then strBand = CStr(Split(strTmp, ",")(0))
    Next
    If strShow = "" And strBand = "" Then Exit Function
        
    '�򿪲�������Դ
    On Error GoTo errH
    strSQL = Replace(RemoveNote(strSQL), "[*]", "")
    Call OpenRecord(rsTmp, strSQL, "mdlPublic_GetDefaultValue", intConnectNo) '[*]��SQL��''��,�����޷�����
    i = rsTmp.RecordCount 'ԭʼ��¼����
        
    '�Ȱ�ָ���İ�ֵ���˳�������
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
    
    '�ٷ���ȱʡ�����ݻ����������
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

'ȥ��TextBox��Ĭ���Ҽ��˵�
Public Function WndMessage(ByVal hwnd As OLE_HANDLE, ByVal Msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' �����Ϣ����WM_CONTEXTMENU���͵���Ĭ�ϵĴ��ں�������
    If Msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(lngTXTProc, hwnd, Msg, wp, lp)
End Function

Public Function CheckPass(ByVal lngRPTID As Long) As Boolean
'���ؼ�˵�������
    Dim rsPass As New ADODB.Recordset
    Dim strPass As String, strSQL As String
            
    If ReportReaded(lngRPTID) Then
        '���û���
        Set rsPass = grsReport
    Else
        strSQL = "Select ID,���,����,˵��,����,��ӡ��,��ֽ,Ʊ��,��ӡ��ʽ,ϵͳ,����ID,����,�޸�ʱ��,����ʱ��,��ֹ��ʼʱ�� " & vbCrLf & _
                 "  ,��ֹ����ʱ�� " & vbCrLf & _
                 "From zlReports Where ID=[1]"
        Set rsPass = OpenSQLRecord(strSQL, "CheckPass", lngRPTID)
        If rsPass.EOF Then Exit Function
        
        '���洦��
        Set grsReport = New ADODB.Recordset
        Set grsReport = rsPass
        gdatModiTime = grsReport!�޸�ʱ��
    End If
    
    If IsNull(rsPass!����) Then Exit Function
    strPass = GetPass(rsPass!���, rsPass!����)
    If strPass <> rsPass!���� Then Exit Function
    CheckPass = True
End Function

Public Function GetPass(ByVal strCode As String, ByVal strName As String, Optional ByVal BlnSave As Boolean = False) As String
    '1-��������ų��Ȳ���20λ,���Կո����
    '2-������λ��ĩλ���,���뱨�����Ƽ������,������뵱ǰλ�õļ��ܴ����ķ�ʽ
    '3-��������ų��ȳ���20λ,��������λ
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
    Dim StrChange As String                     'ת����
    Dim PStart As Integer, PEnd As Integer      'λ��ָ��
    Dim IntDO As Integer
    Dim BytThis As Byte
    
    '��ԭ���ܴ�
    
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
    '����"ThisProgramWriteByZT"
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
    '���ܣ�����ÿ�����ߵ�ASCII��
    '������
    '���أ�
    '-------------------------------------------------------------
    Dim intBit As Integer, iCount As Integer, blnCan As Boolean
    Dim strCode As String
    
    strCode = "OEM_"
    For intBit = 1 To Len(strAsk)
        'ȡÿ���ֵ�ASCII��
        strCode = strCode & Hex(Asc(Mid(strAsk, intBit, 1)))
    Next
    GetOEM = strCode
End Function

Public Function zlGetSymbol(ByVal strInput As String, Optional ByVal bytIsWB As Byte) As String
'���ܣ������ַ����ļ���
'��Σ�strInput-�����ַ�����bytIsWB-�Ƿ����(����Ϊƴ��)
'���Σ���ȷ�����ַ��������󷵻�"-"
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
'���ܣ��Ƴ�SQL����е�ע��
'˵����ֻ֧���Ƴ����е�ע��
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
'���ܣ����������е��Զ���SQL�е�[ϵͳ]���滻������ֵ
    Dim i As Integer
    Dim newPars As RPTPars
    
    Call CopyPars(oldPars, newPars)
    
    For i = 1 To newPars.count
        newPars(i).��ϸSQL = Replace(newPars(i).��ϸSQL, "[ϵͳ]", lngSys)
        newPars(i).����SQL = Replace(newPars(i).����SQL, "[ϵͳ]", lngSys)
    Next
    Set ReplaceParSysNo = newPars
End Function

Public Function CheckParsRela(strSQL As String, ByVal objDatas As RPTDatas, ByVal strName As String _
    , Optional ByVal blnIsCheck As Boolean, Optional ByVal colValue As Collection _
    , Optional ByVal objPars As RPTPars, Optional ByRef strParName As String) As Boolean
'���ܣ�����Ƿ������������
'      varValue=��������ˣ����ʾʵ�ʵĲ���ֵ
'������strName=SQL��������
'      strParName=�󶨵Ĳ�����
    Dim objPar As RPTPar, objData As RPTData
      
    If InStr("Collection", TypeName(colValue)) = 0 Then Set colValue = New Collection
    If objDatas Is Nothing Then
        For Each objPar In objPars
            Call CheckParsRelaChild(strSQL, objPar, strName, colValue)
        Next
    Else
        '���ʱ��������Դ����ȡ����
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
            '���󶨵Ĳ����滻Ϊ0����֤�ܹ���������
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
'���ܣ�������Դ��SQLת����Oracle��ִ�е�SQL
'������
'  strSQL������ԴSQL���Լ�����ת�����SQL
'  objPar����������
'  strName������ԴSQL�Ĳ�����
'  colValue���������϶���

    Dim strTmp As String
    Dim lngTmp As Long          '0-ִ�У� 1-SQL��д���
    Dim bytType As Byte         '0-���棻 1-��Between ... And ...�����
    
    If objPar.���� <> strName Then
        If InStr(strSQL, "[=" & objPar.���� & "]") > 0 Then
            bytType = 0
            If colValue.count = 0 Then
                strTmp = Mid(strSQL, 1, InStr(strSQL, "[=" & objPar.���� & "]") - 1)
                strTmp = RTrim(strTmp)
                lngTmp = 0
                If InStr("=<>", Mid(strTmp, Len(strTmp))) > 0 Then
                    If Mid(strTmp, Len(strTmp) - 1) Like "[<|>][=|>]" Then
                        '������������ڡ�<>�������ڵ��ڡ�>=����С�ڵ��ڡ�<=��
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
                If objPar.���� = 0 Then
                    strSQL = Replace(strSQL, "[=" & objPar.���� & "]", "'" & IIF(colValue.count = 0, "", GetColValues(colValue, objPar.����)) & "'")
                ElseIf objPar.���� = 1 Then
                    strSQL = Replace(strSQL, "[=" & objPar.���� & "]", IIF(colValue.count = 0, 0, Val(GetColValues(colValue, objPar.����))))
                ElseIf objPar.���� = 2 Then
                    strSQL = Replace(strSQL, "[=" & objPar.���� & "]", IIF(colValue.count = 0, "sysdate", "to_date('" & GetColValues(colValue, objPar.����) & "', 'YYYY-MM-DD HH24:MI:SS')"))
                ElseIf objPar.���� = 3 Then
                    strTmp = GetColValues(colValue, objPar.����)
                    If UCase(Trim(strTmp)) Like "IN (*)*" Then
                        strSQL = Replace(strSQL, "= [=" & objPar.���� & "]", IIF(colValue.count = 0, "''", strTmp))
                        strSQL = Replace(strSQL, "=[=" & objPar.���� & "]", IIF(colValue.count = 0, "''", strTmp))
                        strSQL = Replace(strSQL, "[=" & objPar.���� & "]", IIF(colValue.count = 0, "''", strTmp))
                    Else
                        strSQL = Replace(strSQL, "[=" & objPar.���� & "]", "'" & IIF(colValue.count = 0, "0", strTmp) & "'")
                    End If
                End If
            Else
                If bytType = 1 Then
                    Select Case objPar.����
                    Case Val("0-�ַ�"), Val("3-������")
                        If UCase(strSQL) Like "* BETWEEN [[]=" & objPar.���� & "[]] AND *" Then
                            strSQL = Replace(strSQL, "[=" & objPar.���� & "]", "''")
                        ElseIf UCase(strSQL) Like "* BETWEEN * AND [[]=" & objPar.���� & "[]]*" Then
                            strSQL = Replace(strSQL, "[=" & objPar.���� & "]", "''")
                        End If
                    Case Val("1-��ֵ")
                        If UCase(strSQL) Like "* BETWEEN [[]=" & objPar.���� & "[]] AND *" Then
                            strSQL = Replace(strSQL, "[=" & objPar.���� & "]", "1")
                        ElseIf UCase(strSQL) Like "* BETWEEN * AND " & "[[]=" & objPar.���� & "[]]*" Then
                            strSQL = Replace(strSQL, "[=" & objPar.���� & "]", "2")
                        End If
                    Case Val("2-����")
                        strSQL = Replace(strSQL, "[=" & objPar.���� & "]", "sysdate")
                    End Select
                Else
                    If objPar.���� = 0 Then
                        strSQL = Replace(strSQL, "[=" & objPar.���� & "]", "'' Or 1=1)")
                    ElseIf objPar.���� = 1 Then
                        strSQL = Replace(strSQL, "[=" & objPar.���� & "]", "0 Or 1=1)")
                    ElseIf objPar.���� = 2 Then
                        strSQL = Replace(strSQL, "[=" & objPar.���� & "]", "sysdate Or 1=1)")
                    ElseIf objPar.���� = 3 Then
                        strSQL = Replace(strSQL, "[=" & objPar.���� & "]", "'0' Or 1=1)")
                    End If
                    strSQL = Mid(strSQL, 1, lngTmp) & "(" & Mid(strSQL, lngTmp + 1)
                End If
            End If
        End If
    End If
End Function

Private Function GetColValues(ByVal colValues As Collection, ByVal strParName As String) As String
'���ܣ���ȡ�����в�����ֵ
    On Error Resume Next
    GetColValues = colValues("_" & strParName)
End Function

Public Sub GetUserName(ByVal lngSys As Long, strUserName As String, strUserNO As String)
'���ܣ���ȡ��½�û���Ϣ
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
    
    '�ȼ��轨����˽��ͬ��ʲ���Ȩ��(�󲿷����)
    strSQL = _
        " Select A.����,A.���,B.��ԱID" & _
        " From ��Ա�� A,�ϻ���Ա�� B,������Ա C" & _
        " Where A.ID=B.��ԱID And A.ID=C.��ԱID And C.ȱʡ=1 And B.�û���=USER"
    On Error Resume Next
    Set rsTmp = New ADODB.Recordset
    Call OpenRecord(rsTmp, strSQL, "mdlPublic_GetUserName")
    If Err.Number <> 0 And Err.Description Like "*�����ͼ������*" Then
        Err.Clear: On Error GoTo errH
        
        '�ٰ�ϵͳ�����߶�ȡ
        'Set rsTmp = New ADODB.Recordset
        strSQL = "Select ������ From zlSystems Where ���=[1]"
        Set rsTmp = OpenSQLRecord(strSQL, "GetUserName", lngSys)
        If Not rsTmp.EOF Then strOwner = rsTmp!������ & "."
        
        strSQL = _
            " Select A.����,A.���,B.��ԱID" & _
            " From " & strOwner & "��Ա�� A," & strOwner & "�ϻ���Ա�� B," & strOwner & "������Ա C" & _
            " Where A.ID=B.��ԱID And A.ID=C.��ԱID And C.ȱʡ=1 And B.�û���=USER"
        On Error Resume Next
        Set rsTmp = New ADODB.Recordset
        Call OpenRecord(rsTmp, strSQL, "mdlPublic_GetUserName")
        If Err.Number <> 0 And Err.Description Like "*�����ͼ������*" Then
            Err.Clear: On Error GoTo errH
            
            '��ȡ�û�Ȩ�޶���(ֻ��ȡһ��)
            If grsObject Is Nothing Then Set grsObject = UserObject
            If grsObject Is Nothing Then Exit Sub
            If grsObject.State = adStateClosed Then
                Set grsObject = Nothing
                Set grsObject = UserObject
                If grsObject Is Nothing Then Exit Sub
            End If
            
            grsObject.Filter = "OBJECT_NAME='�ϻ���Ա��'"
            If grsObject.EOF Then Exit Sub
            strOwner = grsObject!Owner & "."
            
            '�ٸ�����Ȩ�޵Ķ�ȡ
            strSQL = _
                " Select A.����,A.���,B.��ԱID" & _
                " From " & strOwner & "��Ա�� A," & strOwner & "�ϻ���Ա�� B," & strOwner & "������Ա C" & _
                " Where A.ID = B.��ԱID And A.ID = C.��ԱID And C.ȱʡ = 1 And B.�û��� = USER"
            Set rsTmp = New ADODB.Recordset
            Call OpenRecord(rsTmp, strSQL, "mdlPublic_GetUserName")
        Else
            On Error GoTo errH
        End If
    Else
        On Error GoTo errH
    End If
    If Not rsTmp.EOF Then
        strUserName = rsTmp!����
        strUserNO = rsTmp!���
        glngUserID = rsTmp!��ԱID
    End If
    gcolUserInfo.Add strUserName & "_" & strUserNO & "_" & glngUserID, "_" & lngSys
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function ShowHelpRpt(SHwnd As Long, ByVal htmName As String, Optional Sys As Integer = 1) As Boolean
'��ʾ��������
'SHwnd:���봰�ھ��(��Ϊ��������)
'htmName:��ӳ��CHM�е�htm�ļ�����
'Sys:ϵͳ,0:������;1:zlhis
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
    '���˺�:��ÿ��������ڰ������������Ŀǰû����صİ�������ˣ���ȡ����ÿ��������а����Ĺ���
    '����:2007/09/05
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
'���ܣ�NT�����У����ô�ӡ�����Զ���ֽ�ųߴ�
'������lngWidth��lngHeight=mm(����)
'     intOrient=1-����,2-����
'     intCopys=��ӡ����(�����ӡ��֧��,1-9999,��֧��ʱ�������,Ҳ��Ӱ����������)
'˵��������Width,Height�⣬����ͨ�����������õ����Բ�ֱ�ӷ�ӳ��Printer�ϣ�
'      (ȡDevModeҲ��ӳ������������Ҫ��GetJob���ܻ�ȡ����Ĵ�ӡ�ĵ�����)
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
        
        '���ô�ӡ�ĵ�����
        vDevMode.dmOrientation = intOrient
        vDevMode.dmPaperSize = 256
        vDevMode.dmPaperWidth = Round(sngWidth * 10)   'in tenths of a millimeter
        vDevMode.dmPaperLength = Round(sngHeight * 10) 'in tenths of a millimeter
        vDevMode.dmCopies = intCopys
        'vDevMode.dmCollate = 0& '�߼���ӡ����(��ȡ��ʱ,Copiesֻ֧��1;����֪��ôȡ����)
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
        
        '�ı�������ɫ͸��ģʽ
        Call SetBkMode(lngPrtDC, Val("1-͸��"))
        
        'Close the handle when you are finished with it.
        Call ClosePrinter(lngHandle)
    End If
End Function

Public Function SetNTPrinterPaper_Form(ByVal lngHWND As Long, ByVal sngWidth As Single, ByVal sngHeight As Single, _
    ByVal intOrient As Integer, ByVal intCopys As Integer, Optional objCbo As ComboBox, _
    Optional ByVal strFormName As String, Optional objPrinter As Printer) As Boolean
'���ܣ�NT�����У����ô�ӡ�����Զ���ֽ�ųߴ�(ʹ����ӷ�����Form��ʽ)
'������lngWidth��lngHeight=mm(����)
'     intOrient=1-����,2-����
'     intCopys=��ӡ����(�����ӡ��֧��,1-9999,��֧��ʱ�������,Ҳ��Ӱ����������)
'     objCbo=���ش�ӡ����ʱ�����������˵��������õ�Form���빩�û�ѡ��
'     strFormName=�������������Formʱ����������õ�Form���д�ӡ
'˵��������Width,Height�⣬����ͨ�����������õ����Բ�ֱ�ӷ�ӳ��Printer�ϣ�
'      (ȡDevModeҲ��ӳ������������Ҫ��GetJob���ܻ�ȡ����Ĵ�ӡ�ĵ�����)
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
        
        '��ɾ�����е�Form(�����,��Ϊδɾ���ĳߴ���ܲ�ͬ)
        If objCbo Is Nothing Then
            If GetFormName(lngHandle, vFormSize, strFormName) <> 0 Then
                '���ʹ�ñ��ش�ӡ��Form������ɾ�������,ֱ������
                If strFormName = ZL_FORM_NAME Then
                    If DeleteForm(lngHandle, strFormName & Chr(0)) <> 0 Then
                        'ɾ���ɹ������¼���
                        AddNewForm lngHandle, vFormSize, strFormName
                    Else
                        'δɾ���ɹ�,ֱ�����õ�ǰForm
                        SetTheForm lngHandle, vFormSize, strFormName
                    End If
                Else
                    SetTheForm lngHandle, vFormSize, strFormName
                End If
            Else
                'û����ֱ�Ӽ���Ҫ�õ�Form
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
            '������10������������width����Ҫ������Щ��ӡ��ʹ��ResetDCû��������Form����TinyPDF
            If Abs(objPrinter.Width - vFormSize.Cx / 1000 * Twip_mm) > 10 Then objPrinter.Width = vFormSize.Cx / 1000 * Twip_mm
            If Abs(objPrinter.Height - vFormSize.Cy / 1000 * Twip_mm) > 10 Then objPrinter.Height = vFormSize.Cy / 1000 * Twip_mm
        End If
        
        '�ı�������ɫ͸��ģʽ
        Call SetBkMode(lngPrtDC, Val("1-͸��"))
        
        'Close the handle when you are finished with it.
        Call ClosePrinter(lngHandle)
    End If
End Function

Public Function DelNTPrinterPaper() As Boolean
'���ܣ�ɾ���ղŴ������Զ���ֽ��
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
                '���ؿ��õ�Form,.Flags=0�ı�ʾ�û��Լ������
                 If .Flags = 0 And PtrCtoVbString(.pName) <> ZL_FORM_NAME Then
                    objCbo.AddItem PtrCtoVbString(.pName) & " " & Format(.Size.Cx / 1000, "0") & "mm(��)��" & Format(.Size.Cy / 1000, "0") & "mm(��)"
                End If
            End If
            If PtrCtoVbString(.pName) = FormName Then '��Form���ƱȽ�
                '�����ʹ�ñ��ش�ӡ��form��ʹ��form���յĳߴ�
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
            MsgBox "����:" & Err.LastDllError & vbCrLf & vbCrLf & "û���㹻��Ȩ�������Զ���ֽ�Ÿ�ʽ��", vbExclamation, App.Title
        ElseIf Err.LastDllError = 1902 Then
            '�����Chr(0)��β,��ʱ����������
            MsgBox "����:" & Err.LastDllError & vbCrLf & vbCrLf & "ָ�����Զ���ֽ�Ÿ�ʽ������Ч��", vbExclamation, App.Title
        Else
            MsgBox "����:" & Err.LastDllError & vbCrLf & vbCrLf & "�����Զ���ֽ�Ÿ�ʽʱ��������", vbExclamation, App.Title
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
            MsgBox "����:" & Err.LastDllError & vbCrLf & vbCrLf & "û���㹻��Ȩ�������Զ���ֽ�Ÿ�ʽ��", vbExclamation, App.Title
        ElseIf Err.LastDllError = 80 Then
            MsgBox "����:" & Err.LastDllError & vbCrLf & vbCrLf & "ָ�����Զ���ֽ�Ÿ�ʽ�Ѿ����ڡ�", vbExclamation, App.Title
        Else
            MsgBox "����:" & Err.LastDllError & vbCrLf & vbCrLf & "�����Զ���ֽ�Ÿ�ʽʱ��������", vbExclamation, App.Title
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
'����:��ȡһ���ⲿ�������Ϣ
'����:strFile=�ⲿ�ļ���
'˵����"���;����;�汾(8/9)"
    Dim objFile As FileSystemObject, objText As TextStream
    Dim strLine As String, strSect As String, strTmp As String
    
    Set objFile = New FileSystemObject
    If Not objFile.FileExists(strFile) Then Exit Function
    Set objText = objFile.OpenTextFile(strFile)
    
    Do While Not objText.AtEndOfStream
        strLine = objText.ReadLine
        
        '�жϸ�ʽ�Ƿ���ȷ
        If strSect = "" And Trim(strLine) <> "" And Trim(strLine) <> "[HEAD]" Then objText.Close: Exit Function
        
        'ȡ�öκ�
        If Left(strLine, 1) = "[" And Right(strLine, 1) = "]" Then strSect = UCase(Mid(strLine, 2, Len(strLine) - 2))
        
        '������ͷ
        If strSect = "HEAD" Then
            If strLine Like "������=*" Then
                strTmp = strTmp & ";" & Mid(strLine, InStr(strLine, "=") + 1)
            End If
            If strLine Like "��������=*" Then
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

Public Function CheckFormInput(objForm As Object, Optional bln������ As Boolean) As Boolean
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
                If InStr(strText, "'") > 0 And Not bln������ Then
                    MsgBox "�����д��ڷǷ��ַ���", vbInformation, App.Title
                    obj.SelStart = 0: obj.SelLength = Len(obj.Text)
                    obj.SetFocus: Exit Function
                End If
            End If
        End If
    Next
    CheckFormInput = True
End Function

Public Function GetEditSQL(ByVal strSQL As String, ByVal objPars As RPTPars) As String
'���ܣ����ָ�ʽ,�滻����,���ؿ���ֱ�����е�SQL
'Select * FRom ���ű� Where ID=[1]
'Select * FRom ���ű� Where ID=/*B1*/413/*E1*/
    Dim strLeft As String, strRight As String
    Dim StrPar As String, bytPar As Byte, i As Integer
    
    '�ַ�����������ַ�ת��
    Call mdlPublic.TransSpecialChar(strSQL)
    
    If Not objPars Is Nothing Then
        Do While InStr(strSQL, "[") > 0
            strLeft = Left(strSQL, InStr(strSQL, "[") - 1)
            strRight = Mid(strSQL, InStr(strSQL, "]") + 1)
            StrPar = Mid(strSQL, InStr(strSQL, "[") + 1, InStr(strSQL, "]") - InStr(strSQL, "[") - 1)
            If Trim(StrPar) = "" Then StrPar = 0
            bytPar = CByte(StrPar)
            
            '��ȱʡ����ֵ�滻
            If objPars("_" & CInt(bytPar)).ȱʡֵ <> "" And Not objPars("_" & CInt(bytPar)).ȱʡֵ Like "*��" Then
                Select Case objPars("_" & CInt(bytPar)).����
                    Case 0 '�ַ�
                        StrPar = "'" & objPars("_" & CInt(bytPar)).ȱʡֵ & "'"
                    Case 1 '����
                        StrPar = objPars("_" & CInt(bytPar)).ȱʡֵ
                    Case 2 '����
                        If Left(objPars("_" & CInt(bytPar)).ȱʡֵ, 1) = "&" Then
                            StrPar = GetParSQLMacro(objPars("_" & CInt(bytPar)).ȱʡֵ)
                        Else
                            If InStr(objPars("_" & CInt(bytPar)).ȱʡֵ, ":") > 0 Then
                                '��ʱ���ʽ
                                StrPar = "To_Date('" & Format(objPars("_" & CInt(bytPar)).ȱʡֵ, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            Else
                                '��ʱ���ʽ
                                StrPar = "To_Date('" & Format(objPars("_" & CInt(bytPar)).ȱʡֵ, "yyyy-MM-dd") & "','YYYY-MM-DD')"
                            End If
                        End If
                    Case 3 '������
                        StrPar = objPars("_" & CInt(bytPar)).ȱʡֵ
                End Select
            Else 'ȱʡֵΪ�ջ�Ϊ�Զ�����
                Select Case objPars("_" & CInt(bytPar)).����
                    Case 0 '�ַ�
                        StrPar = "'�մ�'"
                    Case 1 '����
                        StrPar = 0
                    Case 2 '����
                        StrPar = "Sysdate"
                    Case 3 '������(ֱ���滻)
                        If objPars("_" & CInt(bytPar)).ȱʡֵ = "�̶�ֵ�б�" Then
                            'ȡ�̶�ֵ�е�ȱʡֵ
                            '���õķָ���
                            For i = 0 To UBound(Split(objPars("_" & CInt(bytPar)).ֵ�б�, "|"))
                                If Left(Split(objPars("_" & CInt(bytPar)).ֵ�б�, "|")(i), 1) = "��" Then
                                    StrPar = Split(Split(objPars("_" & CInt(bytPar)).ֵ�б�, "|")(i), ",")(1)
                                    Exit For
                                End If
                            Next
                            'û������ȱʡֵ��ȡ��һ��
                            If StrPar = "" Then
                                StrPar = Split(Split(objPars("_" & CInt(bytPar)).ֵ�б�, "|")(0), ",")(1)
                            End If
                        ElseIf objPars("_" & CInt(bytPar)).ȱʡֵ = "ѡ�������塭" Then
                            If objPars("_" & CInt(bytPar)).ֵ�б� <> "" Then
                                'ȡȱʡ��ֵ
                                StrPar = Split(objPars("_" & CInt(bytPar)).ֵ�б�, "|")(1)
                            ElseIf objPars("_" & CInt(bytPar)).��ϸSQL <> "" And objPars("_" & CInt(bytPar)).��ϸ�ֶ� <> "" Then
                                StrPar = GetDefaultValue(objPars("_" & CInt(bytPar)).��ϸSQL, objPars("_" & CInt(bytPar)).��ϸ�ֶ�)
                                If StrPar <> "" Then StrPar = CStr(Split(StrPar, "|")(1))
                                If objPars("_" & CInt(bytPar)).��ʽ = 1 Then
                                    StrPar = " IN (" & StrPar & ") "
                                End If
                            Else
                                StrPar = ""
                            End If
                        Else
                            StrPar = objPars("_" & CInt(bytPar)).ȱʡֵ
                        End If
                End Select
            End If
            strSQL = strLeft & "/*B" & bytPar & "*/" & StrPar & "/*E" & bytPar & "*/" & strRight
        Loop
    End If
    
    '�ַ�����������ַ���ԭ
    Call mdlPublic.TransSpecialChar(strSQL, True)
    
    GetEditSQL = strSQL
End Function

Public Function GetParSQL(ByVal strSQL As String) As String
'���ܣ���SQL���ɴ������ĸ�ʽ
'Select * FRom ���ű� Where ID=/*B1*/413/*E1*/
'Select * FRom ���ű� Where ID=[1]
    Dim strTmp As String, i As Integer
    Dim strL As String, strR As String
    Dim intMax As Integer
    
    '�ַ�����������ַ�ת��
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
    
    '�ַ�����������ַ���ԭ
    Call mdlPublic.TransSpecialChar(strSQL, True)
    
    GetParSQL = strSQL
End Function

Public Function InString(strText As String, strChars As String) As Boolean
'���ܣ������strText���Ƿ����strChars��ָ�����ַ�
    Dim i As Integer
    
    For i = 1 To Len(strChars)
        If InStr(strText, Mid(strChars, i, 1)) > 0 Then
            InString = True
            Exit Function
        End If
    Next
End Function

Public Function MatchString(strText As String, strChars As String) As Boolean
'���ܣ������strText�е������Ƿ�ֻ����strChars��ָ�����ַ�
    Dim i As Integer
    
    For i = 1 To Len(strText)
        If InStr(strChars, Mid(strText, i, 1)) = 0 Then
            Exit Function
        End If
    Next
    
    MatchString = True
End Function

Public Function InitPar() As Boolean
'���ܣ�ϵͳ������ʼ
    On Error GoTo errH
    Static rsPar As ADODB.Recordset
    Static rsParameter As ADODB.Recordset
    Dim strSQL As String
    Dim i As Integer
    
    If rsPar Is Nothing And Not gcnOracle Is Nothing Then '��̬��¼��,ֻ��ȡһ��
        If gcnOracle.State = adStateOpen Then
            Set rsPar = New ADODB.Recordset
            strSQL = "Select ������,����ֵ From ZLOPTIONS Where ������ IN(1,3)"
            Call OpenRecord(rsPar, strSQL, "mdlPublic_InitPar")
            If Not rsPar.EOF Then
                rsPar.Filter = "������=1"
                If Not rsPar.EOF Then gblnRunLog = Nvl(rsPar!����ֵ, 0) = 1
                rsPar.Filter = "������=3"
                If Not rsPar.EOF Then gblnErrLog = Nvl(rsPar!����ֵ, 0) = 1
            End If
        End If
    End If
    If Not gcnOracle Is Nothing Then
        If gcnOracle.State = adStateOpen Then
            strSQL = _
                "Select 28 ������, zl_GetSysParameter('��¼����ʹ�úۼ�', 0, 0) ����ֵ From Dual " & vbNewLine & _
                "Union All " & vbNewLine & _
                "Select 26, zl_GetSysParameter('��������������־', 0, 0) From Dual "
            Set rsParameter = OpenSQLRecord(strSQL, "")
            If rsParameter.BOF = False Then
                Do While rsParameter.EOF = False
                    Select Case rsParameter("������").Value
                    Case Val("26-��������������־")
                        gblnReportRunLog = Val(Nvl(rsParameter!����ֵ))
                    Case Val("28-��¼����ʹ�úۼ�")
                        gblnReportUse = Val(Nvl(rsParameter!����ֵ))
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
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Nvl = IIF(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function Between(X, a, B) As Boolean
'���ܣ��ж�x�Ƿ���a��b֮��
    If a < B Then
        Between = X >= a And X <= B
    Else
        Between = X >= B And X <= a
    End If
End Function

Public Function Decode(ParamArray arrPar() As Variant) As Variant
'���ܣ�ģ��Oracle��Decode����
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
'���ܣ��������뷽ʽ��ʽ����ʾ����,��֤С������󲻳���0,С����ǰҪ��0
'������vNumber=Single,Double,Currency���͵�����,intBit=���С��λ��
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
'���ܣ�����Combo�ؼ���Indexֵ
'Ϊһ��Combo�ؼ�ѡ���б�����ֲ�������Click�¼�
    Const CB_SETCURSEL = &H14E
    
    SendMessage hWnd_combo, CB_SETCURSEL, lngindex, 0
End Sub

Public Sub CboSetWidth(ByVal hWnd_combo As Long, ByVal lngWidth As Long)
'���ܣ�����Combo�ؼ������б�Ŀ��
'�˴��Ŀ�����������б�Ŀ�ȣ���������TWIPΪ��λ
    Const CB_SETDROPPEDWIDTH As Long = &H160

    SendMessage hWnd_combo, CB_SETDROPPEDWIDTH, lngWidth / Screen.TwipsPerPixelX, 0
End Sub

Public Sub CboSetHeight(cboControl As Object, ByVal lngHeight As Long)
'���ܣ�����Combo�ؼ��������б�ĸ߶�
'�˴��Ŀ�����������б�ĸ߶ȣ���������TWIPΪ��λ
    SetWindowPos cboControl.hwnd, 0, 0, 0, cboControl.Width / Screen.TwipsPerPixelX, lngHeight / Screen.TwipsPerPixelY, SWP_NOMOVE
End Sub

Public Sub PressKey(bytKey As Byte)
'���ܣ�����̷���һ����,����SendKey
'������bytKey=VirtualKey Codes��1-254��������vbKeyTab,vbKeyReturn,vbKeyF4
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY, 0)
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
End Sub

Public Function GetTempPathFile(Optional ByVal strPre As String = "tmp") As String
'���ܣ�����һ����ʱ�ļ�
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
        .�ϼ�ID = objSource.�ϼ�ID
        .X = objSource.X
        .Y = objSource.Y
        .W = objSource.W
        .H = objSource.H
        .���� = objSource.����
        .�߿� = objSource.�߿�
        .��ͷ = objSource.��ͷ
        .���� = objSource.����
        .���� = objSource.����
        .���� = objSource.����
        .���� = objSource.����
        .��ʽ = objSource.��ʽ
        .��ʽ�� = objSource.��ʽ��
        .���� = objSource.����
        .���� = objSource.����
        .���� = objSource.����
        .���� = objSource.����
        .���� = objSource.����
        .ǰ�� = objSource.ǰ��
        .���� = objSource.����
        .���� = objSource.����
        .б�� = objSource.б��
        .�и� = objSource.�и�
        .���� = objSource.����
        .��� = objSource.���
        .�ֺ� = objSource.�ֺ�
        .���� = objSource.����
        .�Ե� = objSource.�Ե�
        .Key = objSource.Key
        Set .ͼƬ = objSource.ͼƬ
        Set .SubIDs = objSource.SubIDs
        Set .CopyIDs = objSource.CopyIDs
    End With
End Sub

Public Sub GetChartDataName(ByVal str���� As String, Optional strFX As String, _
    Optional strFS As String, Optional strFY As String, Optional strData As String)
'���ܣ�����Chart�������ݻ�ȡ����ֶε�����
    Dim arrData As Variant
    
    strFX = "": strFS = "": strFY = "": strData = ""
    If str���� <> "" Then
        arrData = Split(str����, "|")
        
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
    
    If strData Like "*��*��" Then
        strData = mdlPublic.GetStdNodeText(strData)
    End If
End Sub

Public Function SetChartDataArray(objChart As Object, rsData As ADODB.Recordset, _
    ByVal strFX As String, ByVal strFS As String, ByVal strFY As String, _
    Optional arrLabelX As Variant, Optional arrLabelS As Variant) As Boolean
'���ܣ�����ͼ������,���չ���X�᷽ʽ
'������strFX=X�ֶ�,strFS=�����ֶ�,strFY=Y�ֶ�
'���أ�arrLabelX=����X���ǩ������
'      arrLabelS=�������б�ǩ������
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
        
        If Not IsNull(rsData.Fields(strFS).Value) Then '����NULL����
            '�������м���
            If Not colFS.Exists("_" & strS) Then
                colFS.Add "_" & strS, strS
            End If
            
            '�������ж�ӦX���Yֵ����
            If Not colFY.Exists("_" & strX & "_" & strS) Then
                colFY.Add "_" & strX & "_" & strS, Val(Nvl(rsData.Fields(strFY).Value, 0))
            Else
                'ͬһ��������ͬһ�����ж��ֵ,���ۼ�Yֵ
                colFY("_" & strX & "_" & strS) = _
                    colFY("_" & strX & "_" & strS) + Val(Nvl(rsData.Fields(strFY).Value, 0))
            End If
        End If
        
        '����X��㼯��
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
        .NumSeries = colFS.count 'ͳ��������
        .NumPoints(1) = colFX.count 'ÿ�����й�������
        
        '����X��Xֵ
        arrX = colFX.Items
        Call .CopyXVectorIn(1, arrX)
                
        '�������ж�ӦX���Yֵ
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
                    arrY(j, i) = .HoleValue '�����в����ڵ�X��
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
'���ܣ�����ͼ������,���չ���X�᷽ʽ
'������strFX=X�ֶ�,strFS=�����ֶ�,strFY=Y�ֶ�
'���أ�arrLabelX=����X���ǩ������
'      arrLabelS=�������б�ǩ������
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
        If Not IsNull(rsData.Fields(strFS).Value) Then '����NULL����
            If Not colFS.Exists("_" & rsData.Fields(strFS).Value) Then
                colFS.Add "_" & rsData.Fields(strFS).Value, rsData.Fields(strFS).Value
            End If
        End If
        rsData.MoveNext
    Next
    
    With objChart.ChartGroups(1).Data
        .Layout = oc2dDataGeneral
        .NumSeries = colFS.count 'ͳ��������
        arrS = colFS.Items
        For i = 0 To UBound(arrS)
            rsData.Filter = strFS & "='" & arrS(i) & "'"
            .NumPoints(i + 1) = rsData.RecordCount '��ǰ���е���
            
            '������ǰ���ж�Ӧ��X,Yֵ
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
'���ܣ����ݵ�ǰ���õ���ʽ��������ͼ��ؼ�����ʽ@@@
'������objChart=ͼ��ؼ�,objItem=ͼ��Ԫ�ض���,rsData=�������ݵļ�¼��
'      sngScale=��ʾ����,blnDesign=�Ƿ��ڱ�����ƻ�����ʹ��
    Dim arrTmp As Variant, strTmp As String
    Dim strFX As String, strFY As String, strFS As String
    Dim arrLabelX As Variant, arrLabelS As Variant
    Dim blnByDate As Boolean, i As Long, j As Long
        
    'ͼ�����
    If objItem.��ͷ <> "" Then
        arrTmp = Split(objItem.��ͷ, "|")
        objChart.Header.Text = arrTmp(0)
        arrTmp = Split(arrTmp(1), ",")
        objChart.Header.Font.name = CStr(arrTmp(0))
        objChart.Header.Font.Size = Val(arrTmp(1)) * sngScale
        objChart.Header.Font.Bold = Val(arrTmp(2)) <> 0
        objChart.Header.Font.Italic = Val(arrTmp(3)) <> 0
    Else
        objChart.Header.Text = ""
    End If
    
    'ͼ������:ɢ��ͼ��Ϊ���ݴ���ʽ��һ��,����Ҫ��������
    '0-Plot(ɢ��ͼ),1-Plot(����ͼ),2-Bar(����ͼ),3-Pie(��ͼ),4-StackingBar(���ͼ),5-Area(���ͼ)
    '6-HiLo(�ɼ�ͼ-�̸�,�̵�),7-HiLoOpenClose(�ɼ�ͼ-�̸�,�̵�,����,����),8-Candle(�ɼ�ͼ-������ͼ:�̸�,�̵�,����,����)
    '9-Polar(����ͼ),10-Radar(�״�ͼ),11-FilledRadar(����״�ͼ),12-Bubble(����ͼ)
    objChart.ChartGroups(1).ChartType = IIF(objItem.��� = 0, 1, objItem.���)
    
    '��������
    Call GetChartDataName(objItem.����, strFX, strFS, strFY)
    objChart.ChartArea.Axes("X").Title.Text = strFX '�ֶ�����ΪXY�����
    objChart.ChartArea.Axes("Y").Title.Text = strFY
    
    '������
    arrLabelX = Array(): arrLabelS = Array()
    If Not rsData Is Nothing Then
        blnByDate = IsType(rsData.Fields(strFX).type, adDBTimeStamp)  '��X��֧������/ʱ������
        objChart.IsBatched = True
        objChart.ChartGroups(1).Data.NumSeries = 0
        If objItem.��� = 0 Then
            '��DataGeneral��ʽ��������
            Call SetChartDataGeneral(objChart, rsData, strFX, strFS, strFY, arrLabelS)
        Else
            '��DataArray��ʽ��������
            Call SetChartDataArray(objChart, rsData, strFX, strFS, strFY, arrLabelX, arrLabelS)
        End If
        objChart.IsBatched = False 'ˢ���ڲ�����
    Else
        If blnNoDataEmpty Then
            objChart.ChartGroups(1).Data.NumSeries = 0
        Else
            '��ʼ����״̬,ȱʡ�ǰ�Array
            If objItem.��� <> 0 Then
                For i = 1 To objChart.ChartGroups(1).Data.NumPoints(1)
                    objChart.ChartGroups(1).Data.X(1, i) = i
                Next
            Else
                'ɢ��ͼ������ʾ
                objChart.ChartGroups(1).Data.X(1, 1) = 3
                objChart.ChartGroups(1).Data.X(1, 2) = 2
                objChart.ChartGroups(1).Data.X(1, 3) = 5
                objChart.ChartGroups(1).Data.X(1, 4) = 4
                objChart.ChartGroups(1).Data.X(1, 5) = 1
            End If
        End If
    End If
    If blnByDate Then '����/ʱ������ʱ��ת��ʾX��̶�
        objChart.ChartArea.Axes("X").AnnotationMethod = oc2dAnnotateTimeLabels
        objChart.ChartArea.Axes("X").AnnotationRotationAngle = -90
    Else
        objChart.ChartArea.Axes("X").AnnotationMethod = oc2dAnnotateValues
        objChart.ChartArea.Axes("X").AnnotationRotationAngle = 0
    End If
    
    objChart.IsBatched = True
    
    'ͼ��
    objChart.ChartGroups(1).SeriesLabels.RemoveAll
    objChart.ChartGroups(1).PointLabels.RemoveAll
    If objItem.���� <= 1 Then
        objChart.Legend.IsShowing = False
    Else
        '����:��=1,��=2,��=16,����=17,����=18,��=32,����=33,����=34
        objChart.Legend.Anchor = Decode(objItem.����, 0, 1, 1, 32, 2, 2, 3, 16, 4, 33, 5, 34, 6, 17, 7, 18)
        '���Ҷ���ʱ����,�����������
        objChart.Legend.Orientation = Decode(objItem.����, 0, 1, 2, 1, 2)
        objChart.Legend.IsShowing = True
                                                    
        '����ͼ��
        If UBound(arrLabelS) <> -1 Then
            For i = 0 To UBound(arrLabelS)
                objChart.ChartGroups(1).SeriesLabels.Add arrLabelS(i)
            Next
        Else
            For i = 1 To objChart.ChartGroups(1).Styles.count
                If strFS <> "" Then
                    objChart.ChartGroups(1).SeriesLabels.Add strFS & i
                Else
                    objChart.ChartGroups(1).SeriesLabels.Add "����" & i
                End If
            Next
        End If
        
        'X���ע:Ŀǰֻ�б�ͼ��Ч
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
            ElseIf objItem.��� <> 0 Then 'General��ʽ������
                If objChart.ChartGroups(1).Data.Layout = oc2dDataArray Then
                    For i = 1 To objChart.ChartGroups(1).Data.NumPoints(1)
                        If strFX <> "" Then
                            objChart.ChartGroups(1).PointLabels.Add strFX & objChart.ChartGroups(1).Data.X(1, i)
                        Else
                            objChart.ChartGroups(1).PointLabels.Add "��" & objChart.ChartGroups(1).Data.X(1, i)
                        End If
                    Next
                ElseIf objChart.ChartArea.Axes("X").DataMin <> objChart.ChartGroups(1).Data.HoleValue Then
                    For i = objChart.ChartArea.Axes("X").DataMin To objChart.ChartArea.Axes("X").DataMax
                        If strFX <> "" Then
                            objChart.ChartGroups(1).PointLabels.Add strFX & i
                        Else
                            objChart.ChartGroups(1).PointLabels.Add "��" & i
                        End If
                    Next
                End If
            End If
        End If
    End If
    
    '���ߺͽ��
    ReDim arrTmp(1 To 12) As Integer
    arrTmp(1) = 2 'oc2dShapeDot(ʵ��Բ)
    arrTmp(2) = 4 'oc2dShapeTriangle(��ʵ������)
    arrTmp(3) = 3 'oc2dShapeBox(ʵ��������)
    arrTmp(4) = 5 'oc2dShapeDiamond(ʵ����)
    arrTmp(5) = 6 'oc2dShapeStar(�Ǻ�)
    arrTmp(6) = 13 'oc2dShapeDiagonalCross(���)
    arrTmp(7) = 12 'oc2dShapeInvertTriangle(��ʵ������)
    arrTmp(8) = 14 'oc2dShapeOpenTriangle(����������)
    arrTmp(9) = 11 'oc2dShapeSquare(����������)
    arrTmp(10) = 10 'oc2dShapeCircle(����Բ)
    arrTmp(11) = 15 'oc2dShapeOpenDiamond(������)
    arrTmp(12) = 16 'oc2dShapeOpenInvertTriangle(���ķ�����)
    'arrTmp(13) = 9 'oc2dShapeCross(�Ӻ�)
    'arrTmp(14) = 8 'oc2dShapeHorizontalLine(����)
    'arrTmp(15) = 7 'oc2dShapeVerticalLine(����)
    For i = 1 To objChart.ChartGroups(1).Styles.count
        If objItem.�Ե� Then
            '�����ȼ�ѭ����ʾ�������
            objChart.ChartGroups(1).Styles(i).Symbol.Shape = arrTmp(((i - 1) Mod UBound(arrTmp)) + 1)
            objChart.ChartGroups(1).Styles(i).Symbol.Size = 7 * sngScale
        Else
            objChart.ChartGroups(1).Styles(i).Symbol.Shape = oc2dShapeNone
        End If
        objChart.ChartGroups(1).Styles(i).Line.Pattern = IIF(objItem.����, 2, 1) 'oc2dLineSolid/oc2dLineNone
        objChart.ChartGroups(1).Styles(i).Line.Width = 1 * sngScale
    Next
    
    '������ʽ������λ��,��άЧ��|XY�ụ��
    '��άЧ��
    If Val(Mid(Format(objItem.��ʽ, "00"), 1, 1)) <> 0 Then
        Select Case objItem.���
            Case 1, 5 '����ͼ,���ͼ
                strTmp = "30,20,10"
            Case 2, 4 '����ͼ,���ͼ
                strTmp = "10,10,10"
            Case 3 '��ͼ
                strTmp = "20,20,0"
            Case Else
                strTmp = "0,0,0"
        End Select
    Else
        strTmp = "0,0,0"
    End If
    '����ֵ���ܳ��Ա���,�ؼ����Զ���
    objChart.ChartArea.View3D.Depth = Val(Split(strTmp, ",")(0))  '���
    objChart.ChartArea.View3D.Elevation = Val(Split(strTmp, ",")(1))  '�߶�
    objChart.ChartArea.View3D.Rotation = Val(Split(strTmp, ",")(2)) '�Ƕ�
    objChart.ChartArea.View3D.Shading = oc2dShadingColor
    'XY�ụ��
    objChart.ChartArea.IsHorizontal = Val(Mid(Format(objItem.��ʽ, "00"), 2, 1)) <> 0
    
    'ͼ������
    If objItem.���� <> 0 Then
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
    
    'ͼ����ɫ
    objChart.Interior.BackgroundColor = IIF(objItem.���� = RGB(255, 255, 255) And blnDesign, &HEFEFEF, objItem.����)
    objChart.Interior.ForegroundColor = objItem.ǰ��
    '��֪Ϊʲô�����ÿؼ�ǰ����Ч,��ͨ�����Կ����Ч
    objChart.ChartArea.Axes("X").AxisStyle.LineStyle.Color = objItem.ǰ��
    objChart.ChartArea.Axes("Y").AxisStyle.LineStyle.Color = objItem.ǰ��
        
    'ͼ������
    objChart.Legend.Font.name = objItem.����
    objChart.Legend.Font.Size = objItem.�ֺ� * sngScale
    objChart.Legend.Font.Bold = objItem.����
    objChart.Legend.Font.Italic = objItem.б��
    
    objChart.ChartArea.Axes("X").Font.name = objItem.���� 'Y��ͬ���仯
    objChart.ChartArea.Axes("X").Font.Size = objItem.�ֺ� * sngScale
    objChart.ChartArea.Axes("X").Font.Bold = objItem.����
    objChart.ChartArea.Axes("X").Font.Italic = objItem.б��
    
    objChart.ChartArea.Axes("X").TitleFont.name = objItem.���� 'Y��ͬ���仯
    objChart.ChartArea.Axes("X").TitleFont.Size = objItem.�ֺ� * sngScale
    objChart.ChartArea.Axes("X").TitleFont.Bold = objItem.����
    objChart.ChartArea.Axes("X").TitleFont.Italic = objItem.б��
    
    objChart.IsBatched = False
End Sub

Public Function GetChartPicture(objDesc As Object, objSource As Object, objItem As RPTItem, _
    Optional rsData As ADODB.Recordset, Optional ByVal sngScale As Single = 1) As StdPicture
'���ܣ�������ͼ�����,����������,����ȡ��Ӧ��ͼ��ͼƬ
    Dim strFX As String, strFY As String, strFS As String
    Dim arrX As Variant, arrY As Variant
    Dim blnByDate As Date, strFile As String, i As Long
        
    objDesc.Left = 0
    objDesc.Top = 0
    objDesc.Width = objSource.Width * sngScale
    objDesc.Height = objSource.Height * sngScale
        
    'ͼ�����
    objDesc.Header.Text = objSource.Header.Text
    objDesc.Header.Font.name = objSource.Header.Font.name
    objDesc.Header.Font.Size = objSource.Header.Font.Size * sngScale
    objDesc.Header.Font.Bold = objSource.Header.Font.Bold
    objDesc.Header.Font.Italic = objSource.Header.Font.Italic
    
    'ͼ������
    objDesc.ChartGroups(1).ChartType = objSource.ChartGroups(1).ChartType
    
    '��������
    objDesc.ChartArea.Axes("X").Title.Text = objSource.ChartArea.Axes("X").Title.Text
    objDesc.ChartArea.Axes("Y").Title.Text = objSource.ChartArea.Axes("Y").Title.Text
    objDesc.ChartArea.Axes("X").AnnotationMethod = objSource.ChartArea.Axes("X").AnnotationMethod
    objDesc.ChartArea.Axes("X").AnnotationRotationAngle = objSource.ChartArea.Axes("X").AnnotationRotationAngle
    blnByDate = objDesc.ChartArea.Axes("X").AnnotationMethod = oc2dAnnotateTimeLabels
    
    '������
    '-----------------------------------------------------------------------------------------------
    objDesc.IsBatched = True
    
    '���ļ�����ʱ����
'    strFile = GetTempPathFile
'    Call objSource.Save(strFile)
'    Call objDesc.Load(strFile)
'    Kill strFile
    
    Call GetChartDataName(objItem.����, strFX, strFS, strFY)
    If strFX <> "" And strFS <> "" And strFY <> "" And Not rsData Is Nothing Then
        'ʹ�ü�¼���󶨱ȽϿ�
        objDesc.ChartGroups(1).Data.NumSeries = 0
        If objItem.��� = 0 Then
            Call SetChartDataGeneral(objDesc, rsData, strFX, strFS, strFY)
        Else
            Call SetChartDataArray(objDesc, rsData, strFX, strFS, strFY)
        End If
    Else
        'ʹ�����鿽��ʱ,���ϵ�н϶�,��Ƚ���
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
    
    'ͼ��
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
    
    '���ߺͽ��
    For i = 1 To objDesc.ChartGroups(1).Styles.count
        objDesc.ChartGroups(1).Styles(i).Symbol.Shape = objSource.ChartGroups(1).Styles(i).Symbol.Shape
        objDesc.ChartGroups(1).Styles(i).Symbol.Size = objSource.ChartGroups(1).Styles(i).Symbol.Size * sngScale
        objDesc.ChartGroups(1).Styles(i).Line.Pattern = objSource.ChartGroups(1).Styles(i).Line.Pattern
        objDesc.ChartGroups(1).Styles(i).Line.Width = objSource.ChartGroups(1).Styles(i).Line.Width * sngScale
    Next
    
    '������ʽ������λ��,��άЧ��|XY�ụ��
    '����ֵ���ܳ��Ա���,�ؼ����Զ���
    objDesc.ChartArea.View3D.Depth = objSource.ChartArea.View3D.Depth
    objDesc.ChartArea.View3D.Elevation = objSource.ChartArea.View3D.Elevation
    objDesc.ChartArea.View3D.Rotation = objSource.ChartArea.View3D.Rotation
    objDesc.ChartArea.View3D.Shading = objSource.ChartArea.View3D.Shading
    'XY�ụ��
    objDesc.ChartArea.IsHorizontal = objSource.ChartArea.IsHorizontal
    
    'ͼ������
    objDesc.ChartArea.Axes("X").MajorGrid.Spacing.IsDefault = objSource.ChartArea.Axes("X").MajorGrid.Spacing.IsDefault
    objDesc.ChartArea.Axes("Y").MajorGrid.Spacing.IsDefault = objSource.ChartArea.Axes("Y").MajorGrid.Spacing.IsDefault
    objDesc.ChartArea.Axes("X").MajorGrid.Style.Width = objSource.ChartArea.Axes("X").MajorGrid.Style.Width * sngScale
    objDesc.ChartArea.Axes("Y").MajorGrid.Style.Width = objSource.ChartArea.Axes("Y").MajorGrid.Style.Width * sngScale
    objDesc.ChartArea.Axes("X").AxisStyle.LineStyle.Width = objSource.ChartArea.Axes("X").AxisStyle.LineStyle.Width * sngScale
    objDesc.ChartArea.Axes("Y").AxisStyle.LineStyle.Width = objSource.ChartArea.Axes("Y").AxisStyle.LineStyle.Width * sngScale
    
    'ͼ����ɫ
    objDesc.Interior.BackgroundColor = objSource.Interior.BackgroundColor
    objDesc.Interior.ForegroundColor = objSource.Interior.ForegroundColor
    objDesc.ChartArea.Axes("X").AxisStyle.LineStyle.Color = objSource.ChartArea.Axes("X").AxisStyle.LineStyle.Color
    objDesc.ChartArea.Axes("Y").AxisStyle.LineStyle.Color = objSource.ChartArea.Axes("Y").AxisStyle.LineStyle.Color
        
    'ͼ������
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
'���ܣ��ж�Chart�ؼ��Ƿ��Ѿ���װ,���δ��װ���Զ�ע��
'���أ��Ѱ�װ��δ��װ��ע��ɹ�����True
'      δ��װ��ע��ʧ�ܷ���False
    Dim objTest As Control
    Static blnInstall As Boolean
    
    If Not blnInstall Then
        On Error Resume Next
        
        Set objTest = frmFlash.Controls.Add("C1Chart2D8.Control.1", "ChartTest")
        If Err.Number <> 0 Then
            Unload frmFlash: Err.Clear
            Call Shell("c1regsvr.exe olch2x8.ocx", vbHide)
            If Err.Number <> 0 Then
                MsgBox "ͼ��ؼ�δ��ȷע�ᣬ���°�װHIS�ͻ��˿��Խ��������⡣", vbExclamation, App.Title
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
'���ܣ���������ִ�е�SQL��������������ļ��У������ӿ�ʼ����ʱ�䣬ִ��ʱ��
'������strProject=��������,�����ȡApp.Title
'      strForm=������,�����ȡForm.Caption
'      strSQL=��Ҫִ�е�SQL���,��Openʱ����,�����������ʾ���һ��SQLִ�����
'      strNote=SQL���˵��
    Dim strTmp As String, sngEnd As Single
    
    If gblnExeSQLTest Then Exit Sub
    mstrRecentSQL = strSQL  '�������ִ�е�SQL���
    
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
'���ܣ�ȥ��TAB�ַ������߿ո񣬻س������ֻ�ɵ��ո�ָ���
'˵������Ҫ��RunSQLFile���Ӻ���
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
'���ܣ���ȡָ�����������Ӷ���
'������
'  intDeviceType��0-MicroSoft ODBC��1-Oralce OLEDB
'  Index���������ӱ��

    Dim cn As New ADODB.Connection
    Dim strKey As String, strPass As String, strServer As String
    
    On Error GoTo hErr
    
    strKey = "_" & Index
        
    '��������
    If grsConnect.State = adStateOpen Then
        With grsConnect
            If .RecordCount > 0 Then .MoveFirst
            Do While .EOF = False
                If Nvl(!���, 0) = Index Then
                    '��ʼ�����Ӷ���
                    strServer = Nvl(!IP) & _
                                IIF(Nvl(!�˿�) = "", ":1521", ":" & Nvl(!�˿�)) & _
                                IIF(Nvl(!ʵ����) = "", "", "/" & Nvl(!ʵ����))
                    strPass = Nvl(!����)
                    '����
                    strPass = mdlPublic.Decipher(MSTR_DBLINK_KEY, strPass)
                    Set cn = gobjRegister.GetConnection(strServer, Nvl(!�û���), strPass _
                                    , CBool(Val("0-��ת������")) _
                                    , intDeviceType _
                                    , _
                                    , CBool(Val("0-�����²��������Ӷ���")))
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
'���ܣ�ͨ���������ӱ�Ż�ȡ��Ӧ���������Ӷ���
'������
'  Index���������ӱ��

    Dim strKey As String, strPass As String, strServer As String
    Dim cn As New ADODB.Connection

    If Index <= 0 Then
        Set GetDBConnection = gcnOracle
    Else
        On Error GoTo hErr
        
        strKey = "_" & Index
        
        If gclsCNs.Item(strKey) Is Nothing Then
            '������������
            If grsConnect.State = adStateOpen Then
                With grsConnect
                    If .RecordCount > 0 Then .MoveFirst
                    Do While .EOF = False
                        If Nvl(!���, 0) = Index Then
                            '��ʼ�����Ӷ���
                            strServer = Nvl(!IP) & _
                                        IIF(Nvl(!�˿�) = "", ":1521", ":" & Nvl(!�˿�)) & _
                                        IIF(Nvl(!ʵ����) = "", "", "/" & Nvl(!ʵ����))
                            strPass = Nvl(!����)
                            '����
                            strPass = mdlPublic.Decipher(MSTR_DBLINK_KEY, strPass)
                            Set cn = gobjRegister.GetConnection(strServer, Nvl(!�û���), strPass _
                                            , CBool(Val("0-��ת������")) _
                                            , IIF(gblnManagementTool, Val("1-OraOLEDB"), Val("0-MSODBC")) _
                                            , _
                                            , CBool(Val("0-�����²��������Ӷ���")))
                            Call gclsCNs.Add(Index, Nvl(!���), cn)
                            GoTo makSet
                            
                            Exit Do
                        End If
                        
                        .MoveNext
                    Loop
                End With
            End If
        Else
makSet:
            '��ȡ��������
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
'���ܣ�ͨ��Command����򿪴�����SQL�ļ�¼��
'������strSQL=�����а���������SQL���,������ʽΪ"[x]"
'             x>=1Ϊ�Զ��������,"[]"֮�䲻���пո�
'             ͬһ�������ɶദʹ��,�����Զ���ΪADO֧�ֵ�"?"����ʽ
'             ʵ��ʹ�õĲ����ſɲ�����,������Ĳ���ֵ��������(��SQL���ʱ��һ��Ҫ�õ��Ĳ���)
'      arrInput=���������Ĳ���ֵ,��������˳�����δ���,��������ȷ����
'      strTitle=����SQLTestʶ��ĵ��ô���/ģ�����
'      arrInput=��һ��������ʽ����ǡ�[��������=x][|��ѯ��ʽ=1-LOB]����x��ʾ�������ӵı�ţ�
'���أ���¼����CursorLocation=adUseClient,LockType=adLockReadOnly,CursorType=adOpenStatic
'������
'SQL���Ϊ="Select ���� From ������Ϣ Where (����ID=[3] Or �����=[3] Or ���� Like [4]) And �Ա�=[5] And �Ǽ�ʱ�� Between [1] And [2] And ���� IN([6],[7])"
'���÷�ʽΪ��Set rsPati=OpenSQLRecord(strSQL, Me.Caption, CDate(Format(rsMove!ת������,"yyyy-MM-dd")),dtpʱ��.Value, lng����ID, "��%", "��", 20, 21)
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
    
    '�ж��Ƿ�����������������ִ�м�¼��
    intConnect = 0
    arrInputNew = arrInput
    If UBound(arrInput) >= 0 Then
        If arrInput(0) Like "��������=[0-9]*" Then
            strTmp = Split(arrInput(0), "=")(1)
            intConnect = Val(strTmp)
            If UBound(arrInput) > 0 Then
                '��������
                ReDim Preserve arrInputNew(UBound(arrInput) - 1)
                For i = 1 To UBound(arrInput)
                    arrInputNew(i - 1) = arrInput(i)
                Next
            End If
        End If
        
        '��ȡָ��LOB��ѯ��ʽ
        If arrInput(0) Like "*|��ѯ��ʽ=*" Then
            strTmp = Split(arrInput(0), "|")(1)
            intQueryMode = Val(Split(strTmp, "=")(1))
            If Not arrInput(0) Like "��������=[0-9]*" Then
                '��������
                ReDim Preserve arrInputNew(UBound(arrInput) - 1)
                For i = 1 To UBound(arrInput)
                    arrInputNew(i - 1) = arrInput(i)
                Next
            End If
        End If
    End If
    
    '������ʹ���˶�̬�ڴ������û��ʹ��/*+ XXX*/����ʾ��ʱ�Զ�����
    strSQLtmp = Trim(UCase(strSQL))
    If Mid(Trim(Mid(strSQLtmp, 7)), 1, 2) <> "/*" And Mid(strSQLtmp, 1, 6) = "SELECT" Then
        arrStr = Split("F_STR2LIST,F_NUM2LIST,F_NUM2LIST2,F_STR2LIST2", ",")
        For i = 0 To UBound(arrStr)
            strSQLtmp1 = strSQLtmp
            Do While InStr(strSQLtmp1, arrStr(i)) > 0
                '�ж�ǰ���Ƿ�����IN �����򲻼�Rule
                '���ҵ����һ��SELECT
                strTmp = Mid(strSQLtmp1, 1, InStr(strSQLtmp1, arrStr(i)) - 1)
                strTmp = Replace(TrimEx(Mid(strTmp, 1, InStrRev(strTmp, "SELECT") - 1)), " ", "")
                If Len(strTmp) > 1 Then strTmp = Mid(strTmp, Len(strTmp) - 2)               'ȡ����3���ַ�
                
                If strTmp = "IN(" Then '����in(select��������������ѭ�������Ƿ����û��ʹ������д����������̬�ڴ溯��
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
        
    '�����Զ���[x]����
    lngLeft = InStr(1, strSQL, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strSQL, "]")
        If lngRight = 0 Then Exit Do
        '������������"[����]����"
        strSeq = Mid(strSQL, lngLeft + 1, lngRight - lngLeft - 1)
        If IsNumeric(strSeq) Then
            i = CInt(strSeq)
            StrPar = StrPar & "," & i
            If i > intMax Then intMax = i
        End If
        
        lngLeft = InStr(lngRight + 1, strSQL, "[")
    Loop

    '�滻Ϊ"?"����
    strLog = strSQL
    For i = 1 To intMax
        strSQL = Replace(strSQL, "[" & i & "]", "?")
        
        '��������SQL���ٵ����
        varValue = arrInputNew(i - 1)
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
            strLog = Replace(strLog, "[" & i & "]", varValue)
        Case "String" '�ַ�
            strLog = Replace(strLog, "[" & i & "]", "'" & Replace(varValue, "'", "''") & "'")
        Case "Date" '����
            strLog = Replace(strLog, "[" & i & "]", "To_Date('" & Format(varValue, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')")
        End Select
    Next

    '���ԭ�в���:��Ȼ�����ظ�ִ��
    cmdData.CommandText = "" '��Ϊ����ʱ�����������
    Do While cmdData.Parameters.count > 0
        cmdData.Parameters.Delete 0
    Loop
    
    '�����µĲ���
    lngLeft = 0: lngRight = 0
    arrPar = Split(Mid(StrPar, 2), ",")
    For i = 0 To UBound(arrPar)
        varValue = arrInputNew((arrPar(i) - 1))
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarNumeric, adParamInput, 30, varValue)
        Case "String" '�ַ�
            intMax = LenB(StrConv(varValue, vbFromUnicode))
            
            If intMax <= 2000 Then
                intMax = IIF(intMax <= 200, 200, 2000)
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarChar, adParamInput, intMax, varValue)
            Else
                If intMax < 4000 Then intMax = 4000
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adLongVarChar, adParamInput, intMax, varValue)
            End If
        Case "Date" '����
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adDBTimeStamp, adParamInput, , varValue)
        Case "Variant()" '����
            '���ַ�ʽ������һЩIN�Ӿ��Union���
            '��ʾͬһ�������Ķ��ֵ,�����Ų�������������Ĳ����Ž���,��Ҫ��֤�����ֵ��������
            If arrPar(i) <> lngRight Then lngLeft = 0
            lngRight = arrPar(i)
            Select Case TypeName(varValue(lngLeft))
            Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarNumeric, adParamInput, 30, varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", varValue(lngLeft), 1, 1)
            Case "String" '�ַ�
                intMax = LenB(StrConv(varValue(lngLeft), vbFromUnicode))
                            
                If intMax <= 2000 Then
                    intMax = IIF(intMax <= 200, 200, 2000)
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarChar, adParamInput, intMax, varValue(lngLeft))
                Else
                    If intMax < 4000 Then intMax = 4000
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adLongVarChar, adParamInput, intMax, varValue(lngLeft))
                End If
                
                strLog = Replace(strLog, "[" & lngRight & "]", "'" & Replace(varValue(lngLeft), "'", "''") & "'", 1, 1)
            Case "Date" '����
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adDBTimeStamp, adParamInput, , varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", "To_Date('" & Format(varValue(lngLeft), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')", 1, 1)
            End Select
            lngLeft = lngLeft + 1 '�ò������������õ��ڼ���ֵ��
        End Select
    Next

    'ִ�з��ؼ�¼��
    If cmdData.ActiveConnection Is Nothing Then
        Set cmdData.ActiveConnection = GetDBConnection(intConnect)    'gcnOracle '���Ƚ���
    ElseIf cmdData.ActiveConnection.ConnectionString <> gcnOracle.ConnectionString _
        Or intTag <> intConnect Then
        Set cmdData.ActiveConnection = GetDBConnection(intConnect)    'gcnOracle '���Ƚ���
    End If
    cmdData.CommandText = strSQL
    intTag = intConnect
    
    Call SQLTest(App.ProductName, strTitle, strLog)
    
    On Error Resume Next
    If intQueryMode = Val("1-LOB") Then
        '�ж�LOB�ֶ����͵Ĳ�ѯ
        GoTo makLOB
    Else
        Set OpenSQLRecord = cmdData.Execute
    End If
    
    If Err.Number = -2147467259 Then
makLOB:
        If gcolOLEDBConnect Is Nothing Then
            Set gcolOLEDBConnect = New Collection
        End If
        'CLOB��BLOB�ֶ�����֧��
        '��ȡ�������Ӷ���
        Set gcnOLEDB = mdlPublic.GetOLEDBConnect(gcnOracle, gcolOLEDBConnect, gobjRegister)
        If gcnOLEDB Is Nothing Then
            Set gcnOLEDB = gobjRegister.ReGetConnection(Val("1-OracleOLEDB"), "", gcnOracle)
            '����
            Call gcolOLEDBConnect.Add(gcnOLEDB)
        End If
        If Not gcnOLEDB Is Nothing Then
            Set cmdData.ActiveConnection = gcnOLEDB
            'Set OpenSQLRecord = cmdData.Execute
            'CLOB��BLOB�ֶ��������ʹ��Command���󣬼�¼������Ĭ�ϵ���adOpenUnspecified������ִ����
            '��ˣ����ü�¼�������Open����
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
'���ܣ�ִ�й������,���Զ��Թ��̲������а󶨱�������
'������strSQL=�������,���ܴ�����,����"������(����1,����2,...)"��
'˵�������¼���������̲�����ʹ�ð󶨱���,�����ϵĵ��÷�����
'  1.���������Ǳ��ʽ,��ʱ�����޷�����󶨱������ͺ�ֵ,��"������(����1,100.12*0.15,...)"
'  2.�м�û�д�����ȷ�Ŀ�ѡ����,��ʱ�����޷�����󶨱������ͺ�ֵ,��"������(����1, , ,����3,...)"
'  3.��Ϊ�ù������Զ�����,����һ��ʹ�ð󶨱���,�Դ�"'"���ַ�����,��Ҫʹ��"''"��ʽ��
    Dim cmdData As New ADODB.Command
    Dim strProc As String, StrPar As String
    Dim blnStr As Boolean, intBra As Integer
    Dim strTemp As String, i As Long
    Dim intMax As Integer, datCur As Date
    
    If Right(Trim(strSQL), 1) = ")" Then
        '���ԭ�в���:��Ȼ�����ظ�ִ��
'        cmdData.CommandText = "" '��Ϊ����ʱ�����������
'        Do While cmdData.Parameters.Count > 0
'            cmdData.Parameters.Delete 0
'        Loop
        
        'ִ�еĹ�����
        strTemp = Trim(strSQL)
        strProc = Trim(Left(strTemp, InStr(strTemp, "(") - 1))
        
        'ִ�й��̲���
        datCur = CDate(0)
        strTemp = Mid(strTemp, InStr(strTemp, "(") + 1)
        strTemp = Trim(Left(strTemp, Len(strTemp) - 1)) & ","
        For i = 1 To Len(strTemp)
            '�Ƿ����ַ����ڣ��Լ����ʽ��������
            If Mid(strTemp, i, 1) = "'" Then blnStr = Not blnStr
            If Not blnStr And Mid(strTemp, i, 1) = "(" Then intBra = intBra + 1
            If Not blnStr And Mid(strTemp, i, 1) = ")" Then intBra = intBra - 1
            
            If Mid(strTemp, i, 1) = "," And Not blnStr And intBra = 0 Then
                StrPar = Trim(StrPar)
                With cmdData
                    If IsNumeric(StrPar) Then '����
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.count, adVarNumeric, adParamInput, 30, Val(StrPar))
                    ElseIf Left(StrPar, 1) = "'" And Right(StrPar, 1) = "'" Then '�ַ���
                        StrPar = Mid(StrPar, 2, Len(StrPar) - 2)
                        
                        'Oracle���ӷ�����:'ABCD'||CHR(13)||'XXXX'||CHR(39)||'1234'
                        If InStr(Replace(StrPar, " ", ""), "'||") > 0 Then GoTo NoneVarLine
                        
                        '˫"''"�İ󶨱�������
                        If InStr(StrPar, "''") > 0 Then StrPar = Replace(StrPar, "''", "'")
                        
                        '���Ӳ�������LOBʱ������ð󶨱���ת��ΪRAWʱ��2000���ַ�����ȷ
                        intMax = LenB(StrConv(StrPar, vbFromUnicode))
                        If intMax = 0 Or intMax < 200 Then intMax = 200
                        If intMax > 1999 Then GoTo NoneVarLine
                        
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.count, adVarChar, adParamInput, intMax, StrPar)
                    ElseIf UCase(StrPar) Like "TO_DATE('*','*')" Then '����
                        StrPar = Split(StrPar, "(")(1)
                        StrPar = Trim(Split(StrPar, ",")(0))
                        StrPar = Mid(StrPar, 2, Len(StrPar) - 2)
                        If StrPar = "" Then
                            'NULLֵ�������ִ���ɼ�����������
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.count, adVarNumeric, adParamInput, , Null)
                        Else
                            If Not IsDate(StrPar) Then GoTo NoneVarLine
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.count, adDBTimeStamp, adParamInput, , CDate(StrPar))
                        End If
                    ElseIf UCase(StrPar) = "SYSDATE" Then '����
                        If datCur = CDate(0) Then datCur = Currentdate
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.count, adDBTimeStamp, adParamInput, , datCur)
                    ElseIf UCase(StrPar) = "NULL" Then 'NULLֵ�����ַ�����ɼ�����������
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.count, adVarChar, adParamInput, 200, Null)
                    ElseIf StrPar = "" Then '��ѡ��������NULL������ܸı���ȱʡֵ:��˿�ѡ��������д���м�
                        GoTo NoneVarLine
                    Else '�������������ӵı��ʽ���޷�����
                        GoTo NoneVarLine
                    End If
                End With
                
                StrPar = ""
            Else
                StrPar = StrPar & Mid(strTemp, i, 1)
            End If
        Next
        
        '����?��
        strTemp = ""
        For i = 1 To cmdData.Parameters.count
            strTemp = strTemp & ",?"
        Next
        strProc = "Call " & strProc & "(" & Mid(strTemp, 2) & ")"
        
        'ִ�й���
        'If cmdData.ActiveConnection Is Nothing Then
            Set cmdData.ActiveConnection = gcnOracle '���Ƚ���
            cmdData.CommandType = adCmdText
        'End If
        cmdData.CommandText = strProc
        
        Call cmdData.Execute

    Else
        GoTo NoneVarLine
    End If
    Exit Sub
NoneVarLine:
    
    '˵����Ϊ�˼��������ӷ�ʽ
    '1.��������adCmdStoredProc��ʽ��8i����������
    '2.�����������ʹ��{},��ʹ����û�в���ҲҪ��()
    strSQL = "Call " & strSQL
    If InStr(strSQL, "(") = 0 Then strSQL = strSQL & "()"
    gcnOracle.Execute strSQL, , adCmdText

End Sub

Public Function ConvertSBC(ByVal strText As String) As String
'���ܣ�ת��ȫ���ַ�Ϊ����ַ�
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
'���ܣ��ж�ĳ��ADO�ֶ����������Ƿ���ָ���ֶ�������ͬһ��(������,����,�ַ�,������)
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
'���ܣ�������ʽ���߰�ť�е���һ���˵�
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
'���ܣ����ݲ�Ʒ�����룬������ҳ
    Dim strCode As String
    
    strCode = zlRegInfo("֧����URL")
    If strCode <> "-" Then
        ShellExecute hwnd, "open", "http://" & strCode, "", "", 1
        zlHomePage = True
    End If
End Function

Public Function zlWebForum(hwnd As Long) As Boolean
'���ܣ����ݲ�Ʒ�����룬������̳
    Dim strCode As String
    
    'strCode = zlRegInfo("֧����BBS")
    strCode = "www.zlsoft.com/techbbs/index.asp"
    If strCode <> "-" Then
        ShellExecute hwnd, "open", "http://" & strCode, "", "", 1
        zlWebForum = True
    End If
End Function

Public Function zlMailTo(hwnd As Long) As Boolean
'���ܣ����ݲ�Ʒ�����뷢�͵����ʼ�
    Dim strCode As String
    strCode = zlRegInfo("֧����MAIL")
    If strCode <> "-" Then
        ShellExecute hwnd, "open", "mailto:" & strCode, "", "", 1
        zlMailTo = True
    End If
End Function

Public Function InitRegister() As Boolean
'���ܣ���ʼ��ע�Ჿ������gobjRegister

    Dim strTmp As String
    
    If gobjRegister Is Nothing Then
        On Error Resume Next
        Set gobjRegister = GetObject("", "zlRegister.clsRegister")
        Err.Clear
    End If
    
    '����֧��δͨ������̨����������prjMain�����ñ������������
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
        MsgBox "����zlRegister��������ʧ��,�����ļ��Ƿ���ڲ�����ȷע�ᡣ", vbExclamation, gstrProductName
        Exit Function
    End If
    InitRegister = True
End Function

Public Function GetPrivFunc(lngSys As Long, lngProgID As Long) As String
'���ܣ����ص�ǰ�û����е�ָ������Ĺ��ܴ�
'������lngSys     ����ǹ̶�ģ�飬��Ϊ0
'      lngProgId  �������
'���أ��ֺż���Ĺ��ܴ�,Ϊ�ձ�ʾû��Ȩ��
    GetPrivFunc = gobjRegister.zlRegFunc(lngSys, lngProgID)
End Function

'--------------------------------------------------
'���ܣ���֤ϵͳע����Ȩ����ȷ��
'������blnTemp-�Ƿ��δ�������ʱע����Ϣ��֤
'���أ���ȷ����"";���󷵻ش�����Ϣ
'--------------------------------------------------
Public Function zlRegCheck(Optional blnTemp As Boolean) As String
    zlRegCheck = gobjRegister.zlRegCheck(blnTemp)
End Function

'--------------------------------------------------
'���ܣ����ָ���Ĳ�Ʒ���л�ע����Ȩ��Ϣ
'������ strItem-ָ������Ȩ��Ŀ
'       blnTemp-�Ƿ��δ�������ʱע����Ϣ��֤
'       intBits-����ͬʱ�ж�����Ϣ�ĵ�λ���ơ���Ʒ�����̵�ָ����õڼ�����Ϣ,0-N,Ϊ-1ʱ��ʾ����";"����Ķ��
'���أ���ȷʱ����ָ������Ϣ�����󷵻�""
'--------------------------------------------------
Public Function zlRegInfo(strItem As String, Optional blnTemp As Boolean, Optional intBits As Integer) As String
    zlRegInfo = gobjRegister.zlRegInfo(strItem, blnTemp, intBits)
End Function

'--------------------------------------------------
'���ܣ������Ȩ������Ϣ
'���أ���2�Ĺ���ĩλ�η����ع������
'--------------------------------------------------
Public Function zlRegTool(Optional blnTemp As Boolean) As Long
    zlRegTool = gobjRegister.zlRegTool(blnTemp)
End Function

Public Function SetBit(ByVal strBit As String, ByVal intBit As Integer, Optional ByVal intVal As Integer = -1) As String
'���ܣ���ָ��λ�ַ���strBit�еĵ�intBitλ����Ϊ0��1
'������intVal=����ֵ,0��1,������ʾ��ת
    If Len(strBit) < intBit Then strBit = strBit & String(intBit - Len(strBit), "0")
    If intVal = -1 Then intVal = IIF(Val(Mid(strBit, intBit, 1)) = 0, 1, 0)
    SetBit = Left(strBit, intBit - 1) & intVal & Mid(strBit, intBit + 1)
End Function

'--------------------------------------------------
'���ܣ�����Ƿ�Ϊ����Ͽ���ADO�Ͽ������Ĵ���!
'���أ�True:�ָ����ӳɹ� False�ָ�����ʧ��
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
        'True '��ORA-12560������ORACLE��������
        CheckAdoConnction = True
    Else
        'False '������������
        CheckAdoConnction = False
        On Error Resume Next
        '�������жϿͻ����Ƿ񱻽�ֹʹ�ã�������ֹ�����Զ��Ͽ�����
        strSQL = "Select NVL(��ֹʹ��,0)  ��ֹʹ�� From zlClients Where ����վ=SYS_CONTEXT('USERENV','TERMINAL')"
        Set rsTmp = OpenSQLRecord(strSQL, "CheckAdoConnction")
        If Err.Number <> 0 Then Err.Clear
        If Not rsTmp Is Nothing Then
            If Not rsTmp.EOF Then
                If rsTmp!��ֹʹ�� = 1 Then
                    If gcnOracle.State = adStateOpen Then gcnOracle.Close
                    CheckAdoConnction = True
                    gblnAutoConnect = False
                    MsgBox "��ǰ����վ�Ѿ�������Ա���ã�����ϵ����Ա������ò����µ�¼��", vbInformation, "�������"
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
            '����������������������
            CheckAdoConnction = True
            blnStatus = True
        End If
    Else
        CheckAdoConnction = False
    End If
End Function

Public Function IsOLEDBConnection(ByVal cnMain As ADODB.Connection) As Boolean
'���ܣ��жϵ�ǰ�����Ƿ���OraOLEDB����
'����Provider���жϣ��������ַ�ʽ
'��ʽһ��'Provider=OraOLEDB.Oracle.1;Password=HIS;Persist Security Info=True;User ID=ZLHIS;Data Source="DYYY";Extended Properties="PLSQLRSet=1"
'��ʽ����
'.Provider = "OraOLEDB.Oracle"
'.Open "PLSQLRSet=1;Data Source=" & strServer & strPersist_Security_Info, strUserName, strPassWord
'�����ַ�ʽ�����Զ�����.Provider����
    'ʹ��Like����Ϊ���ܺ������Ӱ汾��OraOLEDB.Oracle.1
    If UCase(cnMain.Provider) Like "ORAOLEDB.ORACLE*" Then
        IsOLEDBConnection = True
    End If
End Function

Public Function GetAutoConnect() As Boolean
'���ܣ���ȡ�Ƿ���Ȩ�޶����Զ�����
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Nvl(B.����ֵ,Nvl(A.����ֵ,A.ȱʡֵ)) As ����ֵ" & _
        " From zlParameters A,zlUserParas B" & _
        " Where A.ID=B.����ID(+) And A.ϵͳ is Null And A.ģ�� is Null" & _
        " And Nvl(A.˽��,0)=0 And Nvl(A.����,0)=1 And A.������='��������Զ�����'" & _
        " And B.�û���(+) is Null And B.������(+)=SYS_CONTEXT('USERENV','TERMINAL')"
    Set rsTmp = OpenSQLRecord(strSQL, "�����Զ�����Ȩ��", "")
    If Not rsTmp.EOF Then
        GetAutoConnect = Val(Nvl(rsTmp!����ֵ, 0)) = 1
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckErrConnectInfo(ByVal strErrNum As String, ByVal strNote As String, ByVal strErrInfo As String, ByVal intType As Integer) As Boolean
    '------------------------------------------------
    '���ܣ� ��������IntType(1,2)���vb��oralce���صľ��������Ϣ�����ж��Ƿ�Ϊ����Ͽ������Ĵ�������������Ĵ�������
    '������ strNote������Ϣ,strErrInfo������ϸ��Ϣ,intType �������� 1��VB���� 2:ORACLE����
    '���أ� True:���������Ĵ��� False:��������
    '------------------------------------------------
    Dim strTemp As String
    Dim i As Integer
    If intType = 1 Then
        'VB�������
   
        If InStr(strErrInfo, "ORA-12560") > 0 Then
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-12571") > 0 Then
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-03114") > 0 Then
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "E_FAIL") > 0 Then
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-02396") > 0 Then '����������ʱ��, ���������� IDLE_TIME profile
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-02399") > 0 Then '�����������ʱ��, ������ע�� connect_time profile
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-01012") > 0 Then 'û�е�¼
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-00028") > 0 Then '�Ự����ֹ
            CheckErrConnectInfo = True
        Else
            If strErrNum = "3709" Then '3709�����������޷�����ִ�д˲������ڴ����������������ѱ��رջ���Ч����������
                CheckErrConnectInfo = True
            Else
                If strNote = "��ȷ���Ĵ���" Then
                    CheckErrConnectInfo = True
                Else
                    CheckErrConnectInfo = False
                End If
            End If
        End If
    Else
        'ORACLE�������
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
        ElseIf InStr(strErrInfo, "ORA-02396") > 0 Then '����������ʱ��, ���������� IDLE_TIME profile
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-02399") > 0 Then '�����������ʱ��, ������ע�� connect_time profile
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-01012") > 0 Then 'û�е�¼
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-00028") > 0 Then '�Ự����ֹ
            CheckErrConnectInfo = True
        Else
            CheckErrConnectInfo = False
        End If
    End If
End Function

Public Function PictureSpin(objSource As StdPicture, bytSpinType As Byte, objDraw As PictureBox) As StdPicture
'���ܣ�ͼƬ��ת(˳ʱ�룬��ʱ�룩
'������objPic=ԭͼ��
'      SpinType=1-˳ʱ��90��,2-��ʱ��90��
'      objTemp=��ͼ�õ���ʱ����(PictureBox)
'���أ���ת���ͼƬ

    Dim p() As Long
    Dim W As Long, H As Long
    Dim i As Long, j As Long
    
    If bytSpinType = 0 Then
        Set PictureSpin = objSource
        Exit Function
    End If
    
    'ȡԭʼ����
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
    
    'ת����ͼ
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
'���ܣ������ı�������Combo�ؼ��ĵ�ǰֵ
'������cboControl  ׼�����õ�ComboBox�ؼ�
'      strText     ������ı���
'      blnAfter    ��ʾ�ڷָ���֮ǰ��֮��ȡֵ�����û�зָ�������ȡ֮��
'      strSplit    �ָ�����ͨ��Ϊ-
    Dim lngPos As Long
    Dim lngCount As Long
    Dim strTemp As String
    Dim blnMatch As Boolean
    
    For lngCount = 0 To cboControl.ListCount - 1
        strTemp = cboControl.List(lngCount)
        
        lngPos = InStr(strTemp, strSplit)
        If lngPos = 0 Then
            'ֱ�ӷ��������ַ���
            If strText = strTemp Then
                blnMatch = True
                Exit For
            End If
        Else
            If blnAfter = False Then
                'Բ��֮ǰ
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
        '�Ѿ��ҵ�
        cboControl.ListIndex = lngCount
    Else
        If blnAfter = True Then
            '�����ʵ�����ݣ����Ϊǰ��ֻ�Ǳ���
            If strText <> "" Then
                cboControl.AddItem strText
                cboControl.ListIndex = cboControl.NewIndex
            End If
        End If
    End If
End Sub

Public Function CheckSQLPlan(ByVal strSQLCheck As String, Optional ByRef vsPlan As VSFlexGrid, _
    Optional ByVal intConnect As Integer, Optional ByRef blnSuccess As Boolean) As Boolean
'����������:
'         1.���ȫ��ɨ��zlbigtable+zlbaktables��
'         2.���ͱ�ȫ��ɨ��(�����ͳ����Ϣ��User_tab_statistics:num_rows>3000(ҩƷĿ¼һ�������ֵ����) AND num_rows<100 0000��������)
'         3.��������û�����(�Ǵ��)������ϵ�����
'         4.�������ͱ�����ȫɨ�裨inex full scan��INDEX FAST FULL SCAN��
'         5.�������ͱ���Ծʽ����ɨ�裨INDEX SKIP SCAN��
'���أ�blnReturn=true ����������
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
            '��ȡ���,�״ν����ж��Ƿ���zltables���ű�
            '��ZLTABLES,��ȥB���C����Ϊ���,����ȡzlbigtabls��zlbaktables�еı�
            If CheckTblExist("ZLTABLES") Then
               strSQL = " Select Distinct ���� From Zltables Where ���� In ('B1', 'B2', 'B3', 'C1', 'C2', 'C3') "
            Else
                strSQL = "Select Distinct ����" & vbNewLine & _
                        "From Zlbigtables" & vbNewLine & _
                        "Union All" & vbNewLine & _
                        "Select Distinct ���� From Zlbaktables"
            End If
            Call OpenRecord(rsIndex, strSQL, App.ProductName)
            Do While Not rsIndex.EOF
                mstrBigTable = mstrBigTable & "," & rsIndex!����
                rsIndex.MoveNext
            Loop
            mstrBigTable = mstrBigTable & ","
        End If
        
        '��ȡ�б�ͳ����Ϣ��User_tab_statistics:num_rows>3000��
        strSQL = "Select A.������,Nvl(A.����ֵ,A.ȱʡֵ) As ����ֵ" & _
            " From zlParameters A Where A.������ ='������ͱ�'"
        Set rsData = OpenSQLRecord(strSQL, App.ProductName)
        If rsData.BOF = False Then
            StrPar = Nvl(rsData("����ֵ").Value, "0,0")
            If StrPar <> "0,0" Then
                If StrPar <> mstrMiddleTableRows Then
                    strSQL = "Select Table_Name as ���� From User_Tab_Statistics Where Num_Rows > [1] And Num_Rows < [2] "
                    Set rsIndex = OpenSQLRecord(strSQL, App.ProductName, Val(Split(StrPar, ",")(0)), Val(Split(StrPar, ",")(1)))
                    mstrMiddleTable = ""
                    Do While Not rsIndex.EOF
                        If InStr("," & mstrBigTable & ",", "," & rsIndex!���� & ",") = 0 Then
                            mstrMiddleTable = mstrMiddleTable & "," & rsIndex!����
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
                '�ж��Ƿ��Ǵ���б�ȫɨ��
                If InStr(strAllTable, "," & rsPlan!name & ",") > 0 Then
                    If Not vsPlan Is Nothing Then
                        vsPlan.Cell(flexcpForeColor, vsPlan.Rows - 1, 0, vsPlan.Rows - 1, vsPlan.Cols - 1) = &HFF& '��
                    End If
                    blnReturn = True
                End If
            ElseIf InStr(strTmp, "INDEX FAST FULL SCAN") > 0 Or InStr(strTmp, "INDEX FULL SCAN") > 0 Or InStr(strTmp, "INDEX SKIP SCAN") > 0 Then
                '�ж��Ƿ��Ǵ���б�����ȫɨ�����Ծʽ����
                strTable = Split(rsPlan!name & "_", "_")(0)
                If InStr(strAllTable, "," & strTable & ",") > 0 Then
                    If Not vsPlan Is Nothing Then
                        vsPlan.Cell(flexcpForeColor, vsPlan.Rows - 1, 0, vsPlan.Rows - 1, vsPlan.Cols - 1) = &HFF& '��
                    End If
                    blnReturn = True
                End If
            ElseIf InStr(strTmp, "INDEX RANGE SCAN") > 0 Then
                '�����ʹ���˻�����(�Ǵ��)�������
                strTable = Split(rsPlan!name & "_", "_")(0)
                
                If InStr("," & mstrBigTable & ",", "," & strTable & ",") > 0 Then
                    strSQL = "Select distinct d.Table_Name, d.Index_Name, d.Column_Name,d.Column_Position" & vbNewLine & _
                        "              From User_Ind_Columns D" & vbNewLine & _
                        "              Where d.Index_Name = [1] " & vbNewLine & _
                        "              Order By d.Column_Position"
                    Set rsIndex = OpenSQLRecord(strSQL, App.ProductName, rsPlan!name & "")
                    If rsIndex.RecordCount > 0 Then
                        '�����Լ��
                        Set rsCons_FK = GetConsFK(strTable, rsPlan!object_owner & "")
                        strTmp = ""
                        Do While Not rsIndex.EOF
                            strTmp = strTmp & "," & rsIndex!Column_Name
                            rsIndex.MoveNext
                        Loop
                        rsCons_FK.Filter = "Column_Name='" & Mid(strTmp, 2) & "'"
                        If rsCons_FK.RecordCount > 0 Then
                            strTable = Split(rsCons_FK!r_Constraint_Name & "_", "_")(0)
                            
                            '��������Ǵ������Ϊ����������
                            If InStr(mstrBigTable, "," & strTable & ",") = 0 Then
                                If Not vsPlan Is Nothing Then
                                    vsPlan.Cell(flexcpForeColor, vsPlan.Rows - 1, 0, vsPlan.Rows - 1, vsPlan.Cols - 1) = &HFF& '��
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
'���ܣ���ȡָ��������Լ����¼��
'������strFind=����
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
'���ܣ��ռ�SQL��ִ�мƻ�

    Dim strSQL As String, strSID As String
    Dim rsTmp As ADODB.Recordset
    Dim cnOracle As ADODB.Connection
            
    If strSQLCheck <> "" Then
        '׼�����Ӷ���
        Set cnOracle = GetDBConnection(intConnect)
        If cnOracle Is Nothing Then
            Exit Function
        End If
        
        On Error Resume Next
        strSID = Time()
        
        'ִ�мƻ�
        strSQL = "explain plan set statement_id = '" & strSID & "' for " & strSQLCheck & ""
        strSQL = Replace(strSQL, "[ϵͳ]", glngSys)
        cnOracle.Execute strSQL
        If Err.Number = 0 Then
            strSQL = _
                    "Select Time From Plan_Table " & vbNewLine & _
                    "Connect By Prior ID = Parent_Id And Prior Statement_Id = Statement_Id " & vbNewLine & _
                    "Start With ID = 0 And Statement_Id = [1] " & vbNewLine & _
                    "Order By ID "
            On Error Resume Next
            Set GetSQLPlan = OpenSQLRecord(strSQL, "ִ�мƻ�", "��������=" & intConnect, strSID)
            strSQL = _
                    "Select ID, LPad(' ', Level - 1) || Operation || ' ' || Options As Operation, Object_Name As Name" & _
                    "    ,Object_Owner, Cardinality, Bytes" & vbNewLine & _
                    "    ,Cost" & IIF(Err.Number = 0, ", Time ", ",0 as Time ") & vbNewLine & _
                    "From Plan_Table " & vbNewLine & _
                    "Connect By Prior ID = Parent_Id And Prior Statement_Id = Statement_Id " & vbNewLine & _
                    "Start With ID = 0 And Statement_Id = [1] " & vbNewLine & _
                    "Order By ID "
            Err.Clear
            Set GetSQLPlan = OpenSQLRecord(strSQL, "ִ�мƻ�", "��������=" & intConnect, strSID)
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
'���ܣ�����ѡ�񱨱��ID������
'������lngSel=Ĭ��ѡ��ĳһ�У���ֵ��=lngsel��ѡ��

    Dim strSQL As String
    Dim frmNewSelect As New frmSelect
    Dim i As Integer
    Dim bytType As Byte
    
    On Error GoTo errH
    
    strSQL = "select ID,���,����, ���� || '(' || ��� || ')' as ��ʾֵ from zlreports"
    
    If strFind <> "" Then
        strFind = UCase(strFind)
        If IsCharChinese(strFind) Then
            strSQL = strSQL & " Where ���� like '" & strFind & "%'"
        ElseIf IsCharAlpha(strFind) Then
            strSQL = strSQL & " Where Zlpinyincode(����) like '" & strFind & "%'"
        ElseIf IsNumOrChar(strFind) Then
            strSQL = strSQL & " Where ��� like '%" & strFind & "%'"
        End If
    End If
    strSQL = strSQL & " Order by ���"
    
    With frmNewSelect
        .strMatch = strFind
        .strSQLList = strSQL
        .strFLDList = "ID," & adNumeric & ",&B|" & "���," & adVarChar & ",&S|" & "����," & adVarChar & ",&S|��ʾֵ," & adVarChar & ",&D"
        .strParName = "��������"
        .bytType = 1
        .mlngSel = lngSel
'        .mblnMulti = True
        .mblnRelationReport = True
        .mintConnect = 0
        .lngSeekHwnd = lngHWND
    
        '���޸ĵĹ���������ر���
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
    '--����:��ȡϵͳĿ¼
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
'���ܣ���ʾ�����ı��ıȶԴ���
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
    '�ļ�1
    Set objFSO = gobjFile.CreateTextFile(strThisPath & "\" & "Wincmp.sql")
    objFSO.Write strThisSQL
    objFSO.Close
    '�ļ�2
    Set objFSO = gobjFile.CreateTextFile(strNewPath & "\" & "Wincmp.sql")
    objFSO.Write strNewSQL
    objFSO.Close
    '�Ա�
    strCommand = GetWinPath & strSystem & "\wincmp3.exe " & strThisPath & "\" & "Wincmp.sql" & " " & strNewPath & "\" & "Wincmp.sql"
    lngTemp = Shell(strCommand)
    DoEvents
    If Err <> 0 Then
        Err.Clear
        MsgBox "�ļ��Ƚ�ʧ�ܣ����鹤�߼��ļ�λ���Ƿ���ȷ:" & strSystem & "\wincmp3.exe", vbExclamation, "�������"
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
'���ܣ��ж�OS��32λ������64λ
'���أ�True-64λ��False-32λ

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
'���ܣ�����һ�������������
    Dim i As Long
    
    Set objRelationsCopy = New RPTRelations
    For i = 1 To objRelations.count
        objRelationsCopy.Add objRelations.Item(i).��������ID, objRelations.Item(i).������, objRelations.Item(i).����ֵ��Դ, objRelations.Item(i).������������, objRelations.Item(i).Ĭ��
    Next
End Sub

Public Sub SetCopyColProtertys(ByVal objColProtertys As RPTColProtertys, ByRef objColProtertysCopy As RPTColProtertys)
'���ܣ�����һ�������������
    Dim i As Long
    
    Set objColProtertysCopy = New RPTColProtertys
    For i = 1 To objColProtertys.count
        objColProtertysCopy.Add _
            objColProtertys.Item(i).��������, objColProtertys.Item(i).�����ֶ� _
          , objColProtertys.Item(i).������ϵ, objColProtertys.Item(i).����ֵ _
          , objColProtertys.Item(i).������ɫ, objColProtertys.Item(i).������ɫ _
          , objColProtertys.Item(i).�Ƿ�Ӵ�, objColProtertys.Item(i).�Ƿ�����Ӧ�� _
          , objColProtertys.Item(i).����, "_" & objColProtertys.Item(i).Key
    Next
End Sub

Public Function CopyNewRec(ByVal rsSource As ADODB.Recordset, Optional blnOnlyStructure As Boolean _
    , Optional ByVal strFields As String, Optional arrAppFields As Variant) As ADODB.Recordset
'������:����
'�޸��ˣ���˶
'�޸����ڣ�2014-1-6
'�޸ĵ㣺���Ӹ��Ƽ�¼���Ĳ����ֶι���
'��������:2000-11-02
'���Ƽ�¼��
'������strFields=��Ҫ���Ƶļ�¼�����ֶε���˳����ֶ�����ɵ��ַ���
'          �磺1 ����1,3 ����2,7 ����3...��ʾ���Ƽ�¼���ĵ�1,3,7..�ֶ���ɼ�¼��������
'              ID ����1,���� ����2,....��ʾ���Ƽ�¼����ID,����...�ֶ���ɼ�¼������
'              ����*Ϊ�µļ�¼��������
'              �������ͻ�����׳���������ͬ�����⣬��ע��
'           arrAppFields=׷�ӵ��ֶ���Ϣ������,����,����,Ĭ��ֵ,û��Ĭ��ֵ��Empty,û��ָ�����ȴ�Empty
'      blnOnlyStructure=�Ƿ�ֻ���ƽṹ
'�ڳ����У��������漰���໥���ݼ�¼������ʹ��ADO��Clone���Ʋ����ļ�¼����������һ����¼�������ݷ����仯��ʱ�����и�������������ͬ�ı仯��ͨ��ָ�޸Ļ�ɾ����������������ϣ����Щ��¼���໥�䱣�ֶ���
  
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
        '������¼���ṹ
        If Not rsClone Is Nothing Then
            If strFields = "" Then '��¼��ȫ����ģʽ
                arrFieldsName = Array()
                If rsClone.Fields.count > 0 Then
                    ReDim arrFieldsName(rsClone.Fields.count - 1)
                Else
                    arrFieldsName = Array()
                End If
                For intFields = 0 To rsClone.Fields.count - 1
                    arrFieldsName(intFields) = rsClone.Fields(intFields).name & ""
                    .Fields.Append rsClone.Fields(intFields).name, IIF(rsClone.Fields(intFields).type = adNumeric, adDouble, rsClone.Fields(intFields).type), rsClone.Fields(intFields).DefinedSize, adFldIsNullable    '0:��ʾ����
                Next
            Else '��¼�����ָ���ģʽ
                If rsClone.Fields.count > 0 Then
                    arrFieldsName = Split(strFields, ",")
                    For intFields = LBound(arrFieldsName) To UBound(arrFieldsName)
                        '�а�������
                        arrTmp = Split(arrFieldsName(intFields) & " ", " ")
                        strFieldName = Trim(arrTmp(0)): strFieldNameAlias = Trim(arrTmp(1))
                        If IsNumeric(strFieldName) Then strFieldName = rsClone.Fields(Val(strFieldName)).name & ""
                        '��ȡ�ֶ�ԭ������������
                        arrFieldsName(intFields) = strFieldName
                        '����ֶ�,�������ڱ������������е�����Ϊ����
                        .Fields.Append IIF(strFieldNameAlias = "", strFieldName, strFieldNameAlias), IIF(rsClone.Fields(strFieldName).type = adNumeric, adDouble, rsClone.Fields(strFieldName).type), rsClone.Fields(strFieldName).DefinedSize, adFldIsNullable '0:��ʾ����
                    Next
                End If
            End If
        End If
        '׷���ֶ����
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
        '��������
        If Not blnOnlyStructure And Not rsClone Is Nothing Then
            If rsClone.RecordCount <> 0 Then rsClone.MoveFirst
            Do While Not rsClone.EOF
                .AddNew
                For intFields = LBound(arrFieldsName) To UBound(arrFieldsName)
                    '�¼�¼�����а�˳����ӣ���˿�������
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
'���״̬��Ӧ��OEM����
    Dim strOEM As String
    On Error Resume Next
    If gstrProductName = "" Then
        gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "")
    End If
    If gstrProductName <> "-" Then
        objStatus.Panels(1).Text = gstrProductName & "���"
        '����״̬��ͼ���OEM����
        If gstrProductName = "����" Then
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
    '���ܣ��ж�ָ���ַ����Ƿ�ȫ�������ֺ�Ӣ����ĸ���ɣ�������������
    '       ����ĸ�������������ַ�������µļ�⣬isnumbericֻ���ж����֡�
    '��������SSC���ƣ�
    '       strAsk
    '���أ�
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
    '���ܣ��ж�ָ���ַ����Ƿ�ȫ����Ӣ����ĸ����    '
    '������
    '       strAsk
    '���أ�
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
    '���ܣ��ж�ָ���ַ����Ƿ��к���
    '������
    '       strAsk
    '���أ�
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
'���ܣ���ȡע���ĳ�����������(API��ʽ��
'���أ���������
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
'���ܣ���ȡ���ݿ�洢�ġ�ʹ�ø��Ի���񡱲���ֵ
'���أ�True���Ի���False�Ǹ��Ի�

    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo hErr
    
    strSQL = "Select Nvl(a.����ֵ, b.ȱʡֵ) ����ֵ " & vbNewLine & _
             "From zlUserParas A, zlParameters B " & vbNewLine & _
             "Where a.����id = b.Id And a.�û��� = User And b.������ = 'ʹ�ø��Ի����' "
    Set rsTmp = OpenSQLRecord(strSQL, "ʹ�ø��Ի����")
    If rsTmp.RecordCount > 0 Then
        GetMemoryParam = Val(Nvl(rsTmp!����ֵ)) = 1
    End If
    rsTmp.Close
    Exit Function
    
hErr:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub SetCellValue(ByVal bytOutType As Byte, ByVal objRPTForm As Object, _
    ByVal objCurItem As RPTItem, Optional ByVal lngRowBegin As Long)
'���ܣ��������С�ָ����Ԫ��ı�ǩ��Ԫ��
'������
'  bytOutType��������ͣ�0-��Ԥ����1-��ʽԤ����2-��ʽ��ӡ
'  objRPTForm���������
'  objCurItem����ǰ���Ԫ��
'  lngRowBegin����ʼ�к�
'˵����
'  ��ǩԪ�صĸ�ʽ��[�������(�к�,�к�)]
'  ��ע�����кž���0��ʼ

    '���
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
        '���
        If objItem.���� <> 2 Then GoTo makContinue
        
        On Error Resume Next
        If bytOutType > 0 Then
            '��ʽԤ���ʹ�ӡ�ԡ����ݡ����
            If objItem.Value = "" Then objItem.Value = objItem.����                         '��ԭʼ���ı��浽Value����
        Else
            '��Ԥ���ԡ�Caption����ʾ
            If objItem.Value = "" Then objItem.Value = objRPTForm.lbl(objItem.id).Caption   '��ԭʼ���ı��浽Value����
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
        
        '����Ԫ�ص�����
        intBegin = InStr(strValue, "[")
        intEnd = InStr(strValue, "]")
        If intBegin > 0 And intEnd > 0 Then
            strHead = Left(strValue, intBegin - 1)
            strTail = Mid(strValue, intEnd + 1)
            strBody = Mid(strValue, intBegin + 1, intEnd - intBegin - 1)
            
            'ȡ�������
            intTmp = InStr(strBody, "(")
            If intTmp <= 0 Then intTmp = 1
            strVSF = UCase(Trim(Left(strBody, intTmp - 1)))
            
            '�����Ԫ��
            blnFind = False
            For Each objTmp In objRPTForm.mobjReport.Items
                If objTmp.���� = 4 Or objTmp.���� = 5 Then
                    If Trim(UCase(objTmp.����)) = strVSF Then
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
                    strResult = strHead & "[Error����񲻴���]" & strTail
                End If
                GoSub makSet
                GoTo makContinue
            End If
            
            'ȡ��
            strBody = Mid(strBody, intTmp + 1)
            lngRow = Val(strBody)
            
            'ȡ��
            intTmp = InStr(strBody, ",")
            If intTmp > 0 Then
                lngCol = Val(Mid(strBody, InStr(strBody, ",") + 1))
            Else
                If bytOutType = 2 Then
                    strResult = strHead & strTail
                Else
                    strResult = strHead & "[Error���ı��쳣]" & strTail
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
                    strResult = strHead & "[Error��ָ����Ԫ�񲻴���]" & strTail
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
        '��ʽԤ���ʹ�ӡ�ԡ����ݡ����
        objItem.���� = strResult
    Else
        '��Ԥ���ԡ�Caption����ʾ
        For Each lblTmp In objRPTForm.lbl   '����lbl�Ƿ�ֹ����Ԥ����״̬�µ��������ʽ�����쳣
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
'���ܣ�ת��SQL�е������ַ����磺[]�ַ�������������ķ��ų�ͻ
'���أ�True�ɹ���Falseʧ��

    Const STR_ORIGINAL As String = "[|]|(|)"
    Const STR_TRANS As String = "<��������>|<��������>|<������>|<������>"

    Dim strResult As String, strTmp As String
    Dim arrOriginal As Variant, arrTrans As Variant, arrTemp As Variant
    Dim i As Long, j As Long, lngBegin As Long
    Dim intLen As Integer
    
    If Trim(strSQL = "") Then Exit Function
    
    On Error GoTo hErr
    
    strResult = strSQL
    If blnRestore Then
        '��ԭ
        arrOriginal = Split(STR_TRANS, "|")
        arrTrans = Split(STR_ORIGINAL, "|")
    Else
        'ת��
        arrOriginal = Split(STR_ORIGINAL, "|")
        arrTrans = Split(STR_TRANS, "|")
    End If
    
    '���SQL�ַ����Ƿ����[]�ַ�
    i = 1
    lngBegin = 0
    Do While Mid(strResult, i) Like "*'*"
        If Mid(strResult, i, 1) = "'" Then
            If lngBegin <= 0 Then
                '��ʼ
                lngBegin = i
            Else
                '����
                lngBegin = 0
            End If
        Else
            If lngBegin > 0 Then
                '����''�ַ��ڲ����������ַ�������SQL�����ַ���
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
'���ܣ���ȡ�ַ����ַ������ֵĴ���
'���أ��ַ����ַ������ֵĴ���
    Dim lngA As Long, lngB As Long, lngC As Long
    
    lngA = Len(strString)
    lngB = Len(strChar)
    lngC = Len(Replace(strString, strChar, ""))
    CharCount = (lngA - lngC) / lngB
End Function

Public Function AtString(ByVal strVal As String) As Boolean
'���ܣ��ж��ַ����ǵĵ������ǵ���˫�������������ַ�����˫��������ַ���
'���أ�True�ַ�����False���ַ���
    
    AtString = (CharCount(strVal, "'") Mod 2) = 1
End Function

Public Sub SetControlDBConnect(ByRef vControl As Variant)
'���ܣ���������������Ϣ���ؼ�

    Dim strSQL As String, strResult As String
    Dim rsTemp As ADODB.Recordset
    Dim cbiTmp As ComboItem
    
    On Error GoTo hErr
    
    '���ݻ�ȡ
    strSQL = _
            "Select ���, ����, �û���, ����, Ip, �˿�, ʵ����, ˵�� " & vbNewLine & _
            "From ZlConnections " & vbNewLine & _
            "Order By ��� "
    Set rsTemp = mdlPublic.OpenSQLRecord(strSQL, "��ȡ����������Ϣ")
    
    '���ݼ���
    Select Case UCase(TypeName(vControl))
    Case "COMBOBOX"
        Do While rsTemp.EOF = False
            vControl.AddItem "��" & Nvl(rsTemp!���) & "��" & _
                             Nvl(rsTemp!����) & _
                             ""
'                             " ��" & _
'                             "IP��" & Nvl(rsTemp!IP) & "��" & _
'                             "�˿ڣ�" & Nvl(rsTemp!�˿�) & "��" & _
'                             "��������" & Nvl(rsTemp!ʵ����) & _
'                             "��"
            vControl.ItemData(vControl.NewIndex) = Nvl(rsTemp!���, 0)
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
    '����
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
'���ܣ�ͨ��intDBConnectNo��������ȡ����������Ϣ
'������
'  intDBConnectNo���������ӱ��
'  bytType��0-ָ�������������ӵ����ƣ�1-�û���

    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset

    On Error GoTo hErr
    
    strSQL = "Select ����, �û���, ����, Ip, �˿�, ʵ����, ˵�� From Zlconnections Where ��� = [1] "
    Set rsTemp = mdlPublic.OpenSQLRecord(strSQL, "��ȡ��������������Ϣ", intDBConnectNo)
    If rsTemp.EOF = False Then
        Select Case bytType
        Case 0
            GetDBConnectInfo = Nvl(rsTemp!����)
        Case 1
            GetDBConnectInfo = Nvl(rsTemp!�û���)
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
    If strText Like "*��*��" Then
        strText = Left(strText, InStrRev(strText, "��") - 1)
        GetStdNodeText = strText
    Else
        GetStdNodeText = strText
    End If
End Function

Public Function GetSysVersion(ByVal lngSysNO As Long) As String
'���ܣ���ȡָ��ϵͳ�İ汾��
'������
'  lngSysNo��ϵͳ���
'���أ��汾��
    
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo hErr
    strSQL = "Select �汾�� From zlSystems Where ��� = [1]"
    Set rsTemp = mdlPublic.OpenSQLRecord(strSQL, "��ȡϵͳ�汾", lngSysNO)
    If rsTemp.EOF = False Then
        GetSysVersion = mdlPublic.Nvl(rsTemp!�汾��)
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
'���ܣ����ô��弰���пؼ��������С
'������frmMe=��Ҫ��������Ĵ������
'      bytSize:����Ϊ9������,0:����Ϊ9������,1,����Ϊ12������
'      strOther:�������������õĿؼ��������ļ���,��ʽΪ����������1,��������2,��������3,....
'˵����1.����漰��VsFlexGrid�ȱ��ؼ�����Ҫ�������ڵĻ������µ����п���и�
'      2.�������δ�г��������ؼ����Զ���ؼ�,��Ҫ���ض�����ָ�������С����ش���ģ������ⵥ������

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
            '����CommandBars�û��Զ���ؼ���ȡobjCtrol.Container�����
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
                    objCtrol.Height = frmMe.TextHeight("��") + 20
                    'Label�����Ҫ���е���
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
                objCtrol.Width = frmMe.TextWidth("����" & objCtrol.Caption)
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
                objCtrol.Height = frmMe.TextHeight("��") + IIF(bytSize = 0, 100, 120)
            Case "TextBox"
                lngOldSize = objCtrol.Font.Size
                dblRate = lngFontSize / lngOldSize
                
                objCtrol.Font.Size = lngFontSize
                objCtrol.Width = objCtrol.Width * dblRate
                objCtrol.Height = frmMe.TextHeight("��")
            Case "MaskEdBox"
                objCtrol.FontSize = lngFontSize
                objCtrol.Width = frmMe.TextWidth(objCtrol.Mask)
                objCtrol.Height = frmMe.TextHeight("��")
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
                If objFont Is Nothing Then '�ؼ���ʼ����ʱobjFontΪnothing
                    Set objFont = frmMe.Font
                End If
                objFont.Size = lngFontSize
                Set objCtrol.PaintManager.CaptionFont = objFont
                
                Set objFont = objCtrol.TabPaintManager.Font
                If objFont Is Nothing Then '�ؼ���ʼ����ʱobjFontΪnothing
                    Set objFont = frmMe.Font
                End If
                objFont.Size = lngFontSize
                Set objCtrol.TabPaintManager.Font = objFont

                Set objFont = objCtrol.PanelPaintManager.Font
                If objFont Is Nothing Then '�ؼ���ʼ����ʱobjFontΪnothing
                    Set objFont = frmMe.Font
                End If
                objFont.Size = lngFontSize
                Set objCtrol.PanelPaintManager.Font = objFont
            Case "CommandBars"
                Set objFont = objCtrol.Options.Font
                If objFont Is Nothing Then '�ؼ���ʼ����ʱobjFontΪnothing
                    Set objFont = frmMe.Font
                End If
                objFont.Size = lngFontSize
                Set objCtrol.Options.Font = objFont
            Case "TabControl"
                Set objFont = objCtrol.PaintManager.Font
                If objFont Is Nothing Then  '�ؼ���ʼ����ʱobjFontΪnothing
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
'���ܣ���ʽ���ַ���
'������
'  strFormat�����ʽ��[1-x]Ϊ�����Źؼ��֣����ӣ�"����ֵΪ��[1]"
'  arrParams�����ʽ�Ĳ�������ӦstrFormat�еĲ����Źؼ���
'���أ���ʽ������ַ���

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
'���ܣ���SQLд�뼯��
'������
'  cllData�����϶���
'  strSQL��SQL�ַ���

    Dim l As Long
    
    l = cllData.count + l
    cllData.Add strSQL, "K" & l
End Sub

Public Sub InitClass(ByVal objControl As ComboBox _
    , Optional ByVal varDefault As Variant _
    , Optional ByVal lngCurClassID As Long = 0)
'���ܣ���ʼ���������ؼ�
'������
'  objControl����ʼ���Ŀؼ�
'  varDefault��ָ��Ĭ��ѡ��
'  lngCurClassID����ǰ�������ID���������ض�Ӧ�ķ���ID���ؼ���
    
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo hErr
    
    objControl.Clear
    
    If lngCurClassID > 0 Then
        strSQL = _
            "Select a.ID, a.�ϼ�ID, a.����, a.˵�� " & vbCrLf & _
            "From ZlRPTClasses A " & vbCrLf & _
            "Where not a.ID in (Select ID From ZlRPTClasses Start With ID = [1] Connect By Prior ID = �ϼ�id) " & vbCrLf & _
            "Order By a.���� "
    Else
        strSQL = _
            "Select ID, �ϼ�ID, ����, ˵�� " & vbCrLf & _
            "From ZlRPTClasses " & vbCrLf & _
            "Order By ���� "
    End If
    Set rsTemp = mdlPublic.OpenSQLRecord(strSQL, "��ȡ���������Ϣ", lngCurClassID)
    
    objControl.AddItem ""
    Do While rsTemp.EOF = False
        objControl.AddItem rsTemp!����
        objControl.ItemData(objControl.NewIndex) = rsTemp!id
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    
    '����Ĭ��ѡ��
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
'���ܣ���ȡָ���������ͣ״̬
'������
'  lngSysNO��ϵͳ��
'  varRPT�������Ż����ID
'  blnGroup��True-������
'  strInfo��ʵ�Σ��������ź����ƣ��磺����š�����
'���أ�-1-�쳣��0-δ�����򱨱����ڣ�1-���ã���������2-ͣ�ã�������
    
    Dim strSQL As String, strVar As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo hErr
    
    ReportStateSwitch = -1
    strInfo = ""
    strVar = UCase(varRPT)
    
    If blnGroup Then
        '������
        strSQL = "Select ����ʱ��, �Ƿ�ͣ��, ���, ����, 0 �ӱ��� From zlRPTGroups " & vbCr & _
                 "Where " & IIF(lngSysNO <= 0, " ϵͳ Is Null ", " ϵͳ = [1] ")
        If IsNumeric(strVar) Then
            strSQL = strSQL & " And ����ID = [2] "
            Set rsTemp = mdlPublic.OpenSQLRecord(strSQL, "��ȡ������ͣ״̬", lngSysNO, CLng(strVar))
        Else
            strSQL = strSQL & " And ��� = [2] "
            Set rsTemp = mdlPublic.OpenSQLRecord(strSQL, "��ȡ������ͣ״̬", lngSysNO, strVar)
        End If
    Else
        '������ӱ���
        strSQL = "Select a.����ʱ��, a.�Ƿ�ͣ��, a.���, a.����, b.����id �ӱ��� " & vbCr & _
                 "From zlReports A, zlRPTSubs B " & vbCr & _
                 "Where a.Id = b.����id(+) " & _
                 IIF(lngSysNO <= 0, " And a.ϵͳ Is Null ", " And a.ϵͳ(+) = [1] ")
        If IsNumeric(strVar) Then
            strSQL = strSQL & " And a.����ID(+) = [2] "
            Set rsTemp = mdlPublic.OpenSQLRecord(strSQL, "��ȡ������ͣ״̬", lngSysNO, CLng(strVar))
        Else
            strSQL = strSQL & " And a.���(+) = [2] "
            Set rsTemp = mdlPublic.OpenSQLRecord(strSQL, "��ȡ������ͣ״̬", lngSysNO, strVar)
        End If
    End If
    
    If rsTemp.EOF Then
        ReportStateSwitch = 0
    Else
        strInfo = mdlPublic.FormatString("��[1]��[2]", mdlPublic.Nvl(rsTemp!���), mdlPublic.Nvl(rsTemp!����))
        If blnGroup Or Val(mdlPublic.Nvl(rsTemp!�ӱ���)) = 0 Then
            '��������������
            If mdlPublic.Nvl(rsTemp!����ʱ��) = "" And lngSysNO = 0 Then
                ReportStateSwitch = 0
            Else
                ReportStateSwitch = IIF(mdlPublic.Nvl(rsTemp!�Ƿ�ͣ��, 0) = 1, Val("2-ͣ��"), Val("1-����"))
            End If
        Else
            '�ӱ���
            ReportStateSwitch = IIF(mdlPublic.Nvl(rsTemp!�Ƿ�ͣ��, 0) = 1, Val("2-ͣ��"), Val("1-����"))
        End If
    End If
    rsTemp.Close
    
    Exit Function
    
hErr:
    If ErrCenter = 1 Then Resume
End Function

Public Function RPTParsCondExec(ByVal vRPTID As Long, ByVal vCondID As Long, ByVal vRPTPars As RPTPars) As RPTPars
'���ܣ�����ִ�в����ġ�������ѡ��
'������
'  vRPTID������ID
'  vCondID��������
'  vRPTPars��Ĭ�ϵı����������
'���أ��µ�RPTPars����

    Dim strSQL As String, strValue As String, strDefault As String
    Dim rsTmp As ADODB.Recordset
    Dim blnRetry As Boolean
    Dim objNewCond As New RPTPars
    Dim objRPTPar As RPTPar
    Dim i As Integer
    
    On Error GoTo hErr
    
    'ȡָ��������
    blnRetry = True
    strSQL = "Select ������,����ֵ From zlRptConds Where ����ID=[1] And ������=[2]"
    Set rsTmp = OpenSQLRecord(strSQL, "��ȡ���������Ϣ", vRPTID, vCondID)
    blnRetry = False
    
    '����һ����������
    For i = 1 To vRPTPars.count
        Set objRPTPar = vRPTPars(i)
        rsTmp.Filter = "������='" & objRPTPar.���� & "'"
        If rsTmp.RecordCount > 0 Then
            '��������
            strValue = Nvl(rsTmp!����ֵ)
            strDefault = objRPTPar.ȱʡֵ
            If InStr(1, "�̶�ֵ�б�,ѡ�������塭", objRPTPar.ȱʡֵ) <> 0 And objRPTPar.ȱʡֵ <> "" Then
                If InStr(1, strValue, "|") > 0 Then
                    strValue = Split(strValue, "|")(1)
                    If InStr(1, strValue, "!") > 0 Then
                        strValue = Replace(strValue, "!", "|")
                    End If
                End If
            Else
                strDefault = Nvl(rsTmp!����ֵ)
                strValue = objRPTPar.ȱʡֵ
            End If
        Else
            '
            strDefault = objRPTPar.ȱʡֵ
            strValue = objRPTPar.Reserve
        End If
        objNewCond.Add objRPTPar.����, objRPTPar.���, objRPTPar.���� _
            , objRPTPar.����, strDefault, objRPTPar.��ʽ, objRPTPar.ֵ�б� _
            , objRPTPar.����SQL, objRPTPar.��ϸSQL, objRPTPar.�����ֶ� _
            , objRPTPar.��ϸ�ֶ�, objRPTPar.����, "_" & objRPTPar.Key _
            , strValue, objRPTPar.�Ƿ�����
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
'���ܣ����汨���������
'������
'  vReportID������ID
'  vCondID��������
'���أ�True�ɹ���Falseʧ��

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
    
    '��������
    blnRetry = True
    If vCondID = 0 Or vIsSaveAs Then
        'ȡ���������
        strSQL = "Select Max(������) ������ From zlRptConds Where ����ID=[1]"
        Set rsTmp = OpenSQLRecord(strSQL, "��ȡ��������������", vReportID)
        intCondID = Nvl(rsTmp!������, 0) + 1
        
        strCondName = InputBox("�������������ƣ�����յ����Ƶ�ͬ��ȡ����", "��������", "����" & intCondID)
        If Trim(Replace(strCondName, "'", "")) = "" Then Exit Function
    Else
        '������������
        intCondID = vCondID
        strSQL = "Select �������� From zlRptConds Where ����ID=[1] And ������=[2]"
        Set rsTmp = OpenSQLRecord(strSQL, "��ȡ�������������", vReportID, intCondID)
        strCondName = Nvl(rsTmp!��������)
    End If
    blnRetry = False
    
    If UCase(vForm.name) = UCase("frmReport") Then
        Set objPop = vForm.mnuPop_Cond
        Set lbl = vForm.lblName
        strTitle = vForm.mobjReport.����
    Else
        Set objPop = vForm.PopMenu_Cond
        Set lbl = vForm.lbl
        strTitle = vForm.mstrTitle
    End If
    
    '��ȡֵ
    For i = 1 To lbl.UBound
        strParName = lbl(i).ToolTipText
        Set objRPTPar = vPars("_" & strParName)
        If objRPTPar Is Nothing Then GoTo makContinue
        
        If objRPTPar.ȱʡֵ = "�̶�ֵ�б�" Then
            Select Case objRPTPar.��ʽ
            Case Val("0-������")
                If GetCboIndex(vForm.cbo(i), vForm.cbo(i).Text) = -1 Then '�Ƿ���Ϊ����
                    'Reserve�ֶα��汾��������"������ֵ|��ʾֵ"
                    objRPTPar.Reserve = "�̶�ֵ�б�|" & vForm.cbo(i).Text
                    objRPTPar.ȱʡֵ = vForm.cbo(i).Text
                Else
                    '�б�ѡ��
                    'Reserve�ֶα��汾��������"������ֵ|��ʾֵ"
                    objRPTPar.Reserve = "�̶�ֵ�б�|" & vForm.cbo(i).Text
                    strTmp = objRPTPar.ֵ�б�
                    For j = 0 To UBound(Split(strTmp, "|"))
                        strDisp = Split(Split(strTmp, "|")(j), ",")(0)
                        If Left(strDisp, 1) = "��" Then
                            strDisp = Mid(strDisp, 2)
                        End If
                        If strDisp = vForm.cbo(i).Text Then
                            objRPTPar.ȱʡֵ = Split(Split(strTmp, "|")(j), ",")(1)
                            Exit For
                        End If
                    Next
                End If
            Case Val("1-��ѡ��")
                For j = 1 To vForm.opt.UBound
                    If vForm.opt(j).Container.Index = i Then
                        If vForm.opt(j).Value Then
                            'Reserve�ֶα��汾��������"������ֵ|��ʾֵ"
                            objRPTPar.Reserve = "�̶�ֵ�б�|" & vForm.opt(j).ToolTipText
                            objRPTPar.ȱʡֵ = vForm.opt(j).Tag
                        End If
                    End If
                Next
            Case Val("2-��ѡ��")
                'Reserve�ֶα��汾��������"������ֵ|��ʾֵ"
                strTmp = objRPTPar.ֵ�б�
                For j = 0 To 1
                    strDisp = Split(Split(strTmp, "|")(j), ",")(0)
                    If vForm.chk(i).Value = 0 Then
                        If Left(strDisp, 1) <> "��" Then
                            objRPTPar.Reserve = "�̶�ֵ�б�|" & strDisp
                            objRPTPar.ȱʡֵ = Split(Split(strTmp, "|")(j), ",")(1)
                        End If
                    Else
                        If Left(strDisp, 1) = "��" Then
                            objRPTPar.Reserve = "�̶�ֵ�б�|" & Mid(strDisp, 2)
                            objRPTPar.ȱʡֵ = Split(Split(strTmp, "|")(j), ",")(1)
                        End If
                    End If
                Next
            End Select
        ElseIf objRPTPar.ȱʡֵ = "ѡ�������塭" Then
            If vForm.txt(i).Tag = "" Then '�Ƿ���Ϊ����
                'Reserve�ֶα��汾��������"������ֵ|��ʾֵ"
                objRPTPar.Reserve = "ѡ�������塭|"
                objRPTPar.ȱʡֵ = vForm.txt(i).Text
            Else
                '�б�ѡ��
                'Reserve�ֶα��汾��������"������ֵ|��ʾֵ"
                objRPTPar.Reserve = "ѡ�������塭|" & vForm.txt(i).Text
                objRPTPar.ȱʡֵ = vForm.txt(i).Tag
            End If
        Else
            Select Case objRPTPar.����
            Case Val("0-�ַ�"), Val("1-����"), Val("3-������")
                objRPTPar.ȱʡֵ = vForm.txt(i).Text
            Case Val("2-����")
                If objRPTPar.ȱʡֵ Like "&*" Then
                    objRPTPar.Reserve = objRPTPar.ȱʡֵ
                End If
                objRPTPar.ȱʡֵ = Format(vForm.dtp(i).Value, vForm.dtp(i).CustomFormat)

'                '���浽ע���
'                If vForm.dtp(i).CustomFormat Like "*HH:mm:ss" Then
'                    SaveSetting "ZLSOFT" _
'                        , "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & vForm.name & strTitle _
'                        , lbl(i).ToolTipText & "ʱ��" _
'                        , Format(vForm.dtp(i).Value, "HH:mm:ss")
'                End If
            End Select
        End If
        
makContinue:
    Next
    
    '�������洮
    strSQL = ""
    For i = 1 To vParsDefault.count
        Set objRPTPar = vParsDefault(i)
        If objRPTPar.ȱʡֵ = "�̶�ֵ�б�" Then
            strSQL = strSQL & "!!" & vPars(i).���� & "," & vPars(i).Reserve & "!" & Replace(vPars(i).ȱʡֵ, "'", "''")
        ElseIf vParsDefault(i).ȱʡֵ = "ѡ�������塭" Then
            strSQL = strSQL & "!!" & vPars(i).���� & "," & vPars(i).Reserve & "!" & Replace(vPars(i).ȱʡֵ, "'", "''")
        Else
            strSQL = strSQL & "!!" & vPars(i).���� & "," & Replace(vPars(i).ȱʡֵ, "'", "''")
        End If
    Next
    strSQL = "zl_RptConds_Update(" & _
             vReportID & "," & _
             intCondID & "," & _
             "'" & strCondName & "'," & _
             "'" & Mid(strSQL, 3) & "'," & _
             IIF(vIsSaveAs, 0, vCondID) & ")"
    Call gcnOracle.Execute(strSQL, , adCmdStoredProc)
    
    '����˵�
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
    strSQL = "Select �������� From zlRptConds Where ����ID=[1] And ������=[2]"
    Set rsTmp = OpenSQLRecord(strSQL, "��ȡ����Ĳ�������", vRPTID, vCondID)
    blnRetry = False
    
    strCondName = Nvl(rsTmp!��������)
    If MsgBox("��ȷ��Ҫɾ����" & strCondName & "����", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Function
    
    strSQL = "zl_RptConds_Update(" & vRPTID & "," & vCondID & ",'��������','',0,1)"
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
    '���ܣ����ݱ����жϱ��Ƿ����
    '������strTableName - Ҫ��ѯ�ı���
    Dim strSQL As String, rsData As ADODB.Recordset
    
    On Error Resume Next
    strSQL = "select 1 from " & strTableName & " where rownum<1 "
    Set rsData = OpenSQLRecord(strSQL, "CheckTblExist")
    
    CheckTblExist = Err.Number = 0
    Err.Clear
End Function

Public Function GetDBConnectNo(ByVal objVar As RPTPar, ByVal objDatas As RPTDatas) As Integer
'���ܣ�ͨ�������ȡ��Ӧ���������ӱ��

    Dim objData As RPTData
    Dim objPar As RPTPar
    
    If objVar Is Nothing Then Exit Function
    If objDatas Is Nothing Then Exit Function
    
    For Each objData In objDatas
        For Each objPar In objData.Pars
            If objVar.���� = objPar.���� Then
                GetDBConnectNo = objData.�������ӱ��
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
    
    strSQL = "Insert Into zlProgPrivs (ϵͳ,���,����,����,������,Ȩ��) " & vbCrLf & _
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
             "            Where ϵͳ " & IIF(lngSystem <= 0, "Is Null", "= " & lngSystem) & vbCrLf & _
             "              And ��� " & IIF(lngModule <= 0, "Is Null", "= " & lngModule) & vbCrLf & _
             "              And ���� = '����' " & vbCrLf & _
             "              And Upper(����) = '" & strDBObject & "'" & vbCrLf & _
             "              And Upper(������) = '" & strOwner & "'" & vbCrLf & _
             "              And Upper(Ȩ��) = '" & strOperation & "'" & vbCrLf & _
             "            )"
    GetInsertProgPrivs = strSQL
    Exit Function
    
hErr:
    Call ErrCenter
End Function

Public Function GetOLEDBConnect(ByVal cnSource As ADODB.Connection _
    , ByVal colCache As Collection _
    , ByVal objRegister As Object) As ADODB.Connection
'���ܣ���ȡ���漯�϶�����OLEDB���Ӷ���
'������
'  cnSource����Ҫ���ҵ����Ӷ���
'  colCache�����϶���
'  objRegister��ע�Ჿ������

    Dim i As Integer
    Dim strServer As String, strUser As String, strPass As String
    Dim strUserCheck As String, strServerCheck As String, strPassCheck As String

    If objRegister Is Nothing Then Exit Function
    If cnSource Is Nothing Then Exit Function
    If colCache Is Nothing Then Exit Function
    If colCache.count <= 0 Then Exit Function
    
    '����Ƿ񻺴�
    '��ȡ���Ӷ����е����ݿ���������û�����Ϣ
    Call objRegister.GetConnectionInfo(cnSource, strServerCheck, strUserCheck, strPassCheck)
    For i = 1 To colCache.count
        If Not colCache(i) Is Nothing Then
            '��ȡ���Ӷ����е����ݿ���������û�����Ϣ
            Call objRegister.GetConnectionInfo(colCache(i), strServer, strUser, strPass)
            If UCase(Trim(strUserCheck)) = UCase(Trim(strUser)) _
                And UCase(Trim(strServerCheck)) = UCase(Trim(strServer)) Then
                '�ҵ�
                Set GetOLEDBConnect = colCache(i)
                Exit Function
            End If
        End If
    Next
    
End Function

Private Function HaveAdditionTable(ByVal objRPT As Report, ByVal objRPTItem As RPTItem) As Boolean
'--------------------------------------------------------------------------------
'���ܣ����RPTItem���޸��ӱ��
'������
'  lngRPTItemID��RPTItem�����ID
'���أ�True�и��ӱ��False�޸��ӱ��
'--------------------------------------------------------------------------------

    Dim tmpItem As RPTItem
    
    HaveAdditionTable = False
    For Each tmpItem In objRPT.Items
        If objRPTItem.���� = tmpItem.���� And tmpItem.���� = Val("1-���ӱ��") Then
            HaveAdditionTable = True
            Exit For
        End If
    Next
End Function

Public Function IsBottomAdditionGrid(ByVal objItems As RPTItems, ByVal objItem As RPTItem) As Boolean
'--------------------------------------------------------------------------------
'���ܣ��жϲ����Ķ����Ƿ�Ϊ��ײ��ĸ��ӱ���
'������
'  objItems��Report.Items���϶���
'  objItem��Item����
'--------------------------------------------------------------------------------

    Dim objTmp As RPTItem
    
    For Each objTmp In objItems
        'ͨ��Y�����ж�
        If objTmp.Y > objItem.Y And objTmp.���� = Val("4-���ɱ��") And objItem.���� <> "" And objTmp.���� = objItem.���� Then
            IsBottomAdditionGrid = True
            Exit For
        End If
    Next
End Function

Private Function GridAtCard(ByVal objReport As Report, ByVal lngID As Long) As Boolean
'���ܣ��жϱ������Ƿ��ڿ�Ƭ������
'������
'  objReport��Report����
'  lngID���������ID

    If objReport.Items("_" & lngID).��ID <= 0 Then Exit Function
    GridAtCard = objReport.Items("_" & objReport.Items("_" & lngID).��ID).���� = Val("14-��Ƭ")
End Function


