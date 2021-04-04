Attribute VB_Name = "mdlPrint"
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : mdlPrint
'    Project    : DrawGraph
'
'    Description: [type_description_here]
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
Option Explicit

Public Const conRatemmToTwip As Single = 56.6857142857143      '������羵ı���
Public Const mintNullRow As Integer = 1 '���¿̶��������
Private msngTwips As Single 'Screen.TwipsPerPixelX /printer.TwipsPerPixelX
Public gintEditorCurveState As Integer '��¼���µ��Ǳ༭���߻��Ǳ༭���
Private mfrmTendBody As Object
Private mlng���²�����ʾ��ʽ As Long

Public Const OFFSET_LEFT = 20
Public Const OFFSET_TOP = 20
Public Const OFFSET_RIGHT = 20
Public Const OFFSET_BOTTOM = 20

Private Type Type_Printer
    intPage As Integer
    lngWidth As Long
    lngHeight As Long
    lngLeft As Long
    lngTop As Long
    lngRight As Long
    lngBottom As Long
    intOrient As Integer
    intBin As Integer
End Type
Public gPrinter As Type_Printer

Public gobjFSO As New FileSystemObject
Public gblnPrinted As Boolean           '�Ƿ��ӡ�����µ�
Public gintHourBegin As Integer '���µ���ʼʱ��
Public gstrCaveSplit As String '���µ���־��ʱ��֮������ӷ�ʽ:����.��Ժ�ھ�ʱ...����Ժ--��ʱ
Public gvarTime As Variant
Public gstdSet As New StdFont  '��������
Public gbln��Ժ As Boolean  '�����Ƿ��Ժ

Private mintBaby As Integer  '�Ƿ���Ӥ��
Public Const gint���� As Integer = -1
Public Const gint���� As Integer = 1
Public Const gint���� As Integer = 2
Public Const gint���� As Integer = 3
Public Const gint��� As Integer = 10
Public Const gint��Һ As Integer = 9
Public Const gintBmpW As Integer = 12
Public Const gintBmpH As Integer = 12
Public Const glngMaxRows As Long = 80   '������
Public Const glngLableStep As Long = 30  '�̶������п�
Public Const glngLableWith As Long = 90 '�̶����� ��������<=3ʱ��Ĭ���ܿ��
Public Const glngColStep As Long = 16   '���������п�
Public Const glngInitRowStep As Long = 6 '���������и�
Public Const pסԺ��ʿվ As Long = 1262  'סԺ��ʿվ����
Public mbln�������� As Boolean
Public glngCurPage As Long
Public mintBmpW As Integer
Public mintBmpH As Integer


Public RGB_BLACK          As Long
Public RGB_RED            As Long
Public RGB_WRITE          As Long
Public RGB_BLUE          As Long
Public RGB_GRAY          As Long
Public RGB_FleetGRAY     As Long

Public mrsTabTime As New ADODB.Recordset '���±����Ŀʱ���
Public mrsCollect As New ADODB.Recordset '���»�����Ŀ
Public mrsWave As New ADODB.Recordset  '���²�����Ŀ

Public Type DrawClient
    ƫ����X As Long
    ƫ����Y As Long
    �̶����� As RECT
    �̶ȵ�λ As Long
    �������� As RECT
    �е�λ As Single
    ʱ���е�λ As Single
    ʱ���е�λ As Single
    �е�λ As Long
    ˫�� As Boolean 'һ�б�ʾ���У�
    ������ As Long
End Type

Public T_DrawClient As DrawClient

'--��ɫ
Private Enum Color
    ��ɫ = 0
    ���ɫ = &H404040
    ��ɫ = &HE0E0E0
    ��ɫ = 200
End Enum

Private Type BODYFLAG
    ��Ժ As Byte
    ��� As Byte
    ת�� As Byte
    ���� As Byte
    ���� As Byte
    ��Ժ As Byte
    ���� As Byte
    ���� As Byte
End Type

Public T_BodyFlag As BODYFLAG


Private Type TwipsPerPixel
    X As Single
    Y As Single
End Type
Public T_TwipsPerPixel As TwipsPerPixel

'��ӡ�Ǳ��±��ʹ��,�Ա���������������¼�������λ��
Public Type T_LPoint
    X As Long
    Y As Long
    W As Single
End Type

Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const SPI_GETWHEELSCROLLLINES = 104
Public WHEEL_SCROLL_LINES As Long
Global glngPrevWndProc As Long
Public Const SW_RESTORE = 9
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long

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

Public Const DC_PAPERNAMES = 16 'ֽ������(ÿ64�ַ�Ϊһ��,��Chr(0)����)
Public Const DC_PAPERS = 2 'ֽ�ű��(Array or Word)
Public Const DC_BINNAMES = 12 '��ֽ��ʽ(ÿ24�ַ�Ϊһ��,��Chr(0)����)
Public Const DC_BINS = 6 '��ֽ���(Array or Word)

'��ӡֽ�ų���(256=�Զ���)
Public Const PageSize1 = "�ż㣬 8 1/2 x 11 Ӣ��"
Public Const PageSize2 = "+A611 С���ż㣬 8 1/2 x 11 Ӣ��"
Public Const PageSize3 = "С�ͱ��� 11 x 17 Ӣ��"
Public Const PageSize4 = "�����ʣ� 17 x 11 Ӣ��"
Public Const PageSize5 = "�����ļ��� 8 1/2 x 14 Ӣ��"
Public Const PageSize6 = "�����飬5 1/2 x 8 1/2 Ӣ��"
Public Const PageSize7 = "�����ļ���7 1/2 x 10 1/2 Ӣ��"
Public Const PageSize8 = "A3, 297 x 420 ����"
Public Const PageSize9 = "A4, 210 x 297 ����"
Public Const PageSize10 = "A4С�ţ� 210 x 297 ����"
Public Const PageSize11 = "A5, 148 x 210 ����"
Public Const PageSize12 = "B4, 250 x 354 ����"
Public Const PageSize13 = "B5, 182 x 257 ����"
Public Const PageSize14 = "�Կ����� 8 1/2 x 13 Ӣ��"
Public Const PageSize15 = "�Ŀ����� 215 x 275 ����"
Public Const PageSize16 = "10 x 14 Ӣ��"
Public Const PageSize17 = "11 x 17 Ӣ��"
Public Const PageSize18 = "������8 1/2 x 11 Ӣ��"
Public Const PageSize19 = "#9 �ŷ⣬ 3 7/8 x 8 7/8 Ӣ��"
Public Const PageSize20 = "#10 �ŷ⣬ 4 1/8 x 9 1/2 Ӣ��"
Public Const PageSize21 = "#11 �ŷ⣬ 4 1/2 x 10 3/8 Ӣ��"
Public Const PageSize22 = "#12 �ŷ⣬ 4 1/2 x 11 Ӣ��"
Public Const PageSize23 = "#14 �ŷ⣬ 5 x 11 1/2 Ӣ��"
Public Const PageSize24 = "C �ߴ繤����"
Public Const PageSize25 = "D �ߴ繤����"
Public Const PageSize26 = "E �ߴ繤����"
Public Const PageSize27 = "DL ���ŷ⣬ 110 x 220 ����"
Public Const PageSize28 = "C5 ���ŷ⣬ 162 x 229 ����"
Public Const PageSize29 = "C3 ���ŷ⣬ 324 x 458 ����"
Public Const PageSize30 = "C4 ���ŷ⣬ 229 x 324 ����"
Public Const PageSize31 = "C6 ���ŷ⣬ 114 x 162 ����"
Public Const PageSize32 = "C65 ���ŷ⣬114 x 229 ����"
Public Const PageSize33 = "B4 ���ŷ⣬ 250 x 353 ����"
Public Const PageSize34 = "B5 ���ŷ⣬176 x 250 ����"
Public Const PageSize35 = "B6 ���ŷ⣬ 176 x 125 ����"
Public Const PageSize36 = "�ŷ⣬ 110 x 230 ����"
Public Const PageSize37 = "�ŷ������ 3 7/8 x 7 1/2 Ӣ��"
Public Const PageSize38 = "�ŷ⣬ 3 5/8 x 6 1/2 Ӣ��"
Public Const PageSize39 = "U.S. ��׼��д���� 14 7/8 x 11 Ӣ��"
Public Const PageSize40 = "�¹���׼��д���� 8 1/2 x 12 Ӣ��"
Public Const PageSize41 = "�¹����ɸ�д���� 8 1/2 x 13 Ӣ��"

Public Const conBin1 = "�ϲ�ֽ�н�ֽ"
Public Const conBin2 = "�²�ֽ�н�ֽ"
Public Const conBin3 = "�м�ֽ�н�ֽ"
Public Const conBin4 = "�ȴ��ֶ�����ÿҳֽ"
Public Const conBin5 = "�ŷ��ֽ����ֽ"
Public Const conBin6 = "�ŷ��ֽ����ֽ����Ҫ�ȴ��ֶ�����"
Public Const conBin7 = "��ǰȱʡֽ�н�ֽ"
Public Const conBin8 = "������ֽ����ֽ"
Public Const conBin9 = "С�ͽ�ֽ����ֽ"
Public Const conBin10 = "����ֽ�н�ֽ"
Public Const conBin11 = "��������ֽ����ֽ"

'ֽ�Ŵ�ӡ�߽����================================================================
Public Declare Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, ByVal lpPort As String, ByVal iIndex As Long, ByVal lpOutput As String, lpDevMode As Any) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
'��ͬ��ӡ���Ĵ�ӡ��Ԫ���Ȳ�ͬ

'Public Const PHYSICALHEIGHT = 111  'Physical Height in device units

'Public Const PHYSICALOFFSETY = 113 'Physical Printable Area y margin
Public Const LOGPIXELSX = 88 'Number of pixels per logical inch along the screen width
Public Const LOGPIXELSY = 90
Public Const SCALINGFACTORX = 114  'Scaling factor x
Public Const SCALINGFACTORY = 115  'Scaling factor y
Public Const DRIVERVERSION = 0     'Device driver version

'WinNT�Զ���ֽ�ſ���================================================================
Public Declare Function EnumForms Lib "winspool.drv" Alias "EnumFormsA" (ByVal hPrinter As Long, ByVal Level As Long, ByRef pForm As Any, ByVal cbBuf As Long, ByRef pcbNeeded As Long, ByRef pcReturned As Long) As Long
Public Declare Function AddForm Lib "winspool.drv" Alias "AddFormA" (ByVal hPrinter As Long, ByVal Level As Long, pForm As Byte) As Long
Public Declare Function DeleteForm Lib "winspool.drv" Alias "DeleteFormA" (ByVal hPrinter As Long, ByVal pFormName As String) As Long
Public Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, ByVal pDefault As Long) As Long
Public Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Public Declare Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal hWnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, pDevModeOutput As Any, pDevModeInput As Any, ByVal fMode As Long) As Long
Public Declare Function ResetDC Lib "gdi32" Alias "ResetDCA" (ByVal hDC As Long, lpInitData As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByRef lpString2 As Long) As Long

' Optional functions not used in this sample, but may be useful.
Public Declare Function GetForm Lib "winspool.drv" Alias "GetFormA" (ByVal hPrinter As Long, ByVal pFormName As String, ByVal Level As Long, pForm As Byte, ByVal cbBuf As Long, pcbNeeded As Long) As Long
Public Declare Function SetForm Lib "winspool.drv" Alias "SetFormA" (ByVal hPrinter As Long, ByVal pFormName As String, ByVal Level As Long, pForm As Byte) As Long

' Constants for DEVMODE
Public Const CCHFORMNAME = 32
Public Const CCHDEVICENAME = 32
Public Const DM_FORMNAME As Long = &H10000
Public Const DM_ORIENTATION = &H1&
Public Const DM_PAPERSIZE = &H2&
Public Const DM_PAPERLENGTH = &H4&
Public Const DM_PAPERWIDTH = &H8&
Public Const DM_COPIES = &H100&
Public Const DM_DEFAULTSOURCE = &H200&
Public Const DM_COLLATE = &H8000&

' Constants for PRINTER_DEFAULTS.DesiredAccess
Public Const PRINTER_ACCESS_ADMINISTER = &H4
Public Const PRINTER_ACCESS_USE = &H8
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)

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


' Custom constants for this sample's SelectForm function
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
    cx As Long
    cy As Long
End Type

Public Type SECURITY_DESCRIPTOR
    Revision As Byte
    Sbz1 As Byte
    Control As Long
    Owner As Long
    Group As Long
    Sacl As Long  ' ACL
    Dacl As Long  ' ACL
End Type

' The two definitions for FORM_INFO_1 make the coding easier.
Public Type FORM_INFO_1
    flags As Long
    pName As Long   ' String
    Size As SIZEL
    ImageableArea As RECTL
End Type

Public Type sFORM_INFO_1
    flags As Long
    pName As String
    Size As SIZEL
    ImageableArea As RECTL
End Type

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

Public Type PRINTER_DEFAULTS
    pDatatype As String
    pDevMode As Long    ' DEVMODE
    DesiredAccess As Long
End Type

Public Type PRINTER_INFO_2
    pServerName As String
    pPrinterName As String
    pShareName As String
    pPortName As String
    pDriverName As String
    pComment As String
    pLocation As String
    pDevMode As DEVMODE
    pSepFile As String
    pPrintProcessor As String
    pDatatype As String
    pParameters As String
    pSecurityDescriptor As SECURITY_DESCRIPTOR
    Attributes As Long
    Priority As Long
    DefaultPriority As Long
    StartTime As Long
    UntilTime As Long
    Status As Long
    cJobs As Long
    AveragePPM As Long
End Type

'===================================================================================

Public Function IntEx(vNumber As Variant) As Variant
'���ܣ�ȡ����ָ����ֵ����С����
    IntEx = -1 * Int(-1 * Val(vNumber))
End Function

Public Function GetPaperName(intSize As Integer) As String
    '���ܣ� ���ݵ�ǰ��ӡ�������ã���ȡֽ������
    '���أ� ֽ������
    If intSize = 256 Then
        GetPaperName = "�û��Զ��� ..."
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

Public Function InitPrint(objParent As Object) As Boolean
    '���ܣ�����ע���frmparent.mobjreport���ݳ�ʼ����ӡ������(����->������->��ǰ)
    '���أ�����޴�ӡ����ֽ�Ų���,��ʧ��
    Dim i As Integer, strPName As String
    Dim strPrinter As String  '��ӡ��
    Dim intPage As Integer  'ֽ��
    Dim lngWidth As Long  '�Զ���ֽ�ſ��
    Dim lngHeight As Long  '�Զ���ֽ�Ÿ߶�
    Dim intOrient As Byte  'ֽ��
    Dim intBin As Integer  '��ֽ��ʽ
    If Not ExistsPrinter Then Exit Function
    
    '��ʼ����ӡ����
    
    strPrinter = Trim(zldatabase.GetPara("���µ���ӡ��", glngSys, 1255, Printer.DeviceName))
    intPage = Val(zldatabase.GetPara("���µ�ֽ��", glngSys, 1255, Printer.PaperSize))
    lngWidth = Val(zldatabase.GetPara("���µ����", glngSys, 1255, Printer.Width))
    lngHeight = Val(zldatabase.GetPara("���µ��߶�", glngSys, 1255, Printer.Height))
    intOrient = Val(zldatabase.GetPara("���µ�ֽ��", glngSys, 1255, Printer.Orientation))
    intBin = Val(zldatabase.GetPara("���µ���ֽ", glngSys, 1255, Printer.PaperBin))

    
    '��ӡ��
    If Printer.DeviceName <> strPName Then
        For i = 0 To Printers.Count - 1
            If Printers(i).DeviceName = strPrinter Then Set Printer = Printers(i): Exit For
        Next
    End If
    On Error Resume Next
    'ֽ��
    If intPage = 256 Then
        Printer.PaperSize = 256
        Printer.Width = lngWidth
        Printer.Height = lngHeight
    Else
        Printer.PaperSize = intPage
    End If
    'ֽ��
    'ֽ��ֵ��,ֽ�ſ��ֵ����,ֽ��ԭΪ1
    Printer.Orientation = intOrient
    '��ֽ
    Printer.PaperBin = intBin
    '����
    Printer.Copies = 1
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    'WinNT�Զ���ֽ�Ŵ���
    If IsWindowsNT And intPage = 256 Then
        If AddCustomPaper(objParent.hWnd, lngWidth / conRatemmToTwip, lngHeight / conRatemmToTwip) = FORM_NOT_SELECTED Then Exit Function
    End If
    InitPrint = True
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
    osvi.dwOSVersionInfoSize = Len(osvi)
    If GetVersionEx(osvi) = 0 Then
        Exit Function
    End If
    GetWinPlatform = osvi.dwPlatformId
End Function

Public Function GetFormName(ByVal PrinterHandle As Long, FormSize As SIZEL, FormName As String) As Integer
    Dim NumForms As Long, i As Long
    Dim FI1 As FORM_INFO_1
    Dim aFI1() As FORM_INFO_1           ' Working FI1 array
    Dim Temp() As Byte                  ' Temp FI1 array
    Dim FormIndex As Integer
    Dim BytesNeeded As Long
    Dim RetVal As Long
    
    FormName = vbNullString
    FormIndex = 0
    ReDim aFI1(1)
    ' First call retrieves the BytesNeeded.
    RetVal = EnumForms(PrinterHandle, 1, aFI1(0), 0&, BytesNeeded, NumForms)
    ReDim Temp(BytesNeeded)
    ReDim aFI1(BytesNeeded / Len(FI1))
    ' Second call actually enumerates the supported forms.
    RetVal = EnumForms(PrinterHandle, 1, Temp(0), BytesNeeded, BytesNeeded, NumForms)
    Call CopyMemory(aFI1(0), Temp(0), BytesNeeded)
    For i = 0 To NumForms - 1
        With aFI1(i)
            If .Size.cx = FormSize.cx And .Size.cy = FormSize.cy Then
                ' Found the desired form
                FormName = PtrCtoVbString(.pName)
                FormIndex = i + 1
                Exit For
            End If
        End With
    Next i
    GetFormName = FormIndex  ' Returns non-zero when form is found.
End Function

Public Function AddCustomPaper(ByVal lngHwnd As Long, lngWidth As Long, lngHeight As Long) As Integer
    '���ܣ�����һ��NT��ʹ�õ��Զ���ֽ��
    '����������=mm(����)
    Dim lngSize As Long ' Size of DEVMODE
    Dim vDevMode As DEVMODE
    Dim arrDevMode() As Byte ' Working DEVMODE
    
    Dim lngHandle As Long 'Handle to printer
    Dim lngPrtDC As Long ' Handle to Printer DC
    Dim strPrtName As String
    
    Dim vFormSize As SIZEL
    
    strPrtName = Printer.DeviceName
    lngPrtDC = Printer.hDC
    
    If OpenPrinter(strPrtName, lngHandle, 0&) Then '��ȡ��ӡ�����
        ' Retrieve the size of the DEVMODE.
        lngSize = DocumentProperties(lngHwnd, lngHandle, strPrtName, 0&, 0&, 0&)
        ' Reserve memory for the actual size of the DEVMODE.
        ReDim arrDevMode(1 To lngSize)
        
        ' Fill the DEVMODE From the printer.
        lngSize = DocumentProperties(lngHwnd, lngHandle, strPrtName, arrDevMode(1), 0&, DM_OUT_BUFFER)
        ' Copy the Public (predefined) portion of the DEVMODE.
        Call CopyMemory(vDevMode, arrDevMode(1), Len(vDevMode))
        
        ' If FormName is "zlBillPaper", we must make sure it exists
        ' before using it. Otherwise, it came From our EnumForms list,
        ' And we do not need to check first. Note that we could have
        ' passed in a Flag instead of checking for a literal name.
        
        ' Use form "zlBillPaper", adding it if necessary.
        ' Set the desired size of the form needed.
        ' Given in thousandths of millimeters
        vFormSize.cx = lngWidth * 1000 ' width
        vFormSize.cy = lngHeight * 1000 ' height
        
        If GetFormName(lngHandle, vFormSize, "zlBillPaper") = 0 Then
            'Form not found - Either of the next 2 lines will work.
            'FormName = AddNewForm(lngHandle, vFormSize, "zlBillPaper")
            AddNewForm lngHandle, vFormSize, "zlBillPaper"
            If GetFormName(lngHandle, vFormSize, "zlBillPaper") = 0 Then
                Call ClosePrinter(lngHandle)
                AddCustomPaper = FORM_NOT_SELECTED   ' Selection Failed!
                Exit Function
            Else
                AddCustomPaper = FORM_ADDED  ' Form Added, Selection succeeded!
            End If
        End If
        ' Change the appropriate member in the DevMode.
        ' In this case, you want to change the form name.
        vDevMode.dmFormName = "zlBillPaper" & Chr(0)  ' Must be NULL terminated!
        ' Set the dmFields bit flag to indicate what you are changing.
        vDevMode.dmFields = DM_FORMNAME
        
        ' Copy your changes back, then update DEVMODE.
        Call CopyMemory(arrDevMode(1), vDevMode, Len(vDevMode))
        lngSize = DocumentProperties(lngHwnd, lngHandle, strPrtName, arrDevMode(1), arrDevMode(1), DM_IN_BUFFER Or DM_OUT_BUFFER)
        
        lngSize = ResetDC(lngPrtDC, arrDevMode(1))   ' Reset the DEVMODE for the DC.
        ' Close the handle when you are finished with it.
        Call ClosePrinter(lngHandle)
        ' Selection Succeeded! But was Form Added?
        If AddCustomPaper <> FORM_ADDED Then AddCustomPaper = FORM_SELECTED
    Else
        AddCustomPaper = FORM_NOT_SELECTED   ' Selection Failed!
    End If
End Function

Public Function DelCustomPaper() As Boolean
    '���ܣ�ɾ���ղŴ������Զ���ֽ��
    Dim lngHandle As Long
    Dim strName As String
    
    strName = Printer.DeviceName
    If OpenPrinter(strName, lngHandle, 0&) Then
        DelCustomPaper = (DeleteForm(lngHandle, "zlBillPaper" & Chr(0)) <> 0)
        Call ClosePrinter(lngHandle)
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

Private Sub CloseRs(RS As ADODB.Recordset)
    '���ܣ��ر�Recordset����
    On Error Resume Next
    If RS.State = ADODB.adStateOpen Then RS.Close
    Set RS = Nothing
End Sub

Public Function AddNewForm(lngPrtHandle As Long, vFormSize As SIZEL, strFormName As String) As String
    Dim FI1 As sFORM_INFO_1
    Dim aFI1() As Byte
    Dim RetVal As Long
    
    With FI1
        .flags = 0
        .pName = strFormName
        With .Size
            .cx = vFormSize.cx
            .cy = vFormSize.cy
        End With
        With .ImageableArea
            .Left = 0
            .Top = 0
            .Right = FI1.Size.cx
            .Bottom = FI1.Size.cy
        End With
    End With
    ReDim aFI1(Len(FI1))
    Call CopyMemory(aFI1(0), FI1, Len(FI1))
    RetVal = AddForm(lngPrtHandle, 1, aFI1(0))
    If RetVal = 0 Then
        If Err.LastDllError = 5 Then
            MsgBox "��û��Ȩ�����ô�ӡ��""" & Printer.DeviceName & """Ϊ�Զ���ߴ磬��ӡ������ܻ᲻������", vbExclamation, App.Title
        Else
            MsgBox "���ô�ӡ��ֽ��ʱ�������󣬱�ţ� " & Err.LastDllError, vbExclamation, App.Title
        End If
        AddNewForm = ""
    Else
        AddNewForm = FI1.pName
    End If
End Function

Public Sub ShowFlash(Optional strInfo As String, Optional sngPer As Single, Optional frmParent As Object)
    '���ܣ���ʾ�����صȴ�����ȴ���(strInfo)
    '����:strInfo=������ʾ��Ϣ
    '     sngPer=����
    Static blnShow As Boolean
    If sngPer > 1 Then sngPer = 1

    If strInfo = "" Then
        Unload frmFlash
        blnShow = False
    Else
        If Not blnShow Then
            On Error Resume Next
            frmFlash.lbl.Top = frmFlash.lbl.Top - frmFlash.lbl.Height / 2
            frmFlash.lblPer.Top = frmFlash.lbl.Top
            frmFlash.lbl.Caption = strInfo
            frmFlash.lblDo.Caption = String(25 * sngPer, frmFlash.lblDo.Tag)

            If sngPer > 0 Then
                frmFlash.lblPer.Caption = Int(sngPer * 100) & "%"
            Else
                frmFlash.lblPer.Caption = ""
            End If

            If frmParent Is Nothing Then
                SetWindowPos frmFlash.hWnd, -1, (Screen.Width - frmFlash.Width) / 2 / Screen.TwipsPerPixelX, (Screen.Height - frmFlash.Height) / 2 / Screen.TwipsPerPixelY, 0, 0, 1
                ShowWindow frmFlash.hWnd, 5
            Else
                Err.Clear
                frmFlash.Show , frmParent
                If Err.Number <> 0 Then
                    Err.Clear
                    SetWindowPos frmFlash.hWnd, -1, (Screen.Width - frmFlash.Width) / 2 / Screen.TwipsPerPixelX, (Screen.Height - frmFlash.Height) / 2 / Screen.TwipsPerPixelY, 0, 0, 1
                    ShowWindow frmFlash.hWnd, 5
                End If
            End If

            frmFlash.Refresh
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

Private Function NewPrintPage(objOut As Object, intPage As Long, Optional blnNewPage As Boolean = True) As Object
    '���ܣ���ӡ��Ԥ��һҳ����ʱ�Ե�ǰҳ����������,��������ҳ
    '������blnNewPage=ΪFalseʱ����ӡҳ�ŵ�,һ���ӡ��������������,��˲����������
    '���أ���ҳ����,����Ϊ��ӡ����PictureBox
    On Error GoTo errH
    Dim objDraw As Object, blnPrint As Boolean
    Dim lngWidth As Long, lngHeight As Long, lngOldY As Long
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    Dim strFontName As String, lngFontSize As Long, blnFontBold As Boolean
    Dim blnFontItalic As Boolean, lngFontColor As Long
    
    
    blnPrint = TypeName(objOut) = "Printer"
    '�߽���Ϣ(Twip)
    
    lngLeft = Val(zldatabase.GetPara("���µ���߾�", glngSys, 1255, OFFSET_LEFT)) * conRatemmToTwip
    lngRight = Val(zldatabase.GetPara("���µ��ұ߾�", glngSys, 1255, OFFSET_RIGHT)) * conRatemmToTwip
    lngTop = Val(zldatabase.GetPara("���µ��ϱ߾�", glngSys, 1255, OFFSET_TOP)) * conRatemmToTwip
    lngBottom = Val(zldatabase.GetPara("���µ��±߾�", glngSys, 1255, OFFSET_BOTTOM)) * conRatemmToTwip
    
    
    lngWidth = Printer.Width: lngHeight = Printer.Height
    
    'һҳ���������Ĵ���
    If Not blnPrint Then
        Set objDraw = objOut.picPage(objOut.picPage.UBound)
    Else
        Set objDraw = Printer
    End If
    strFontName = objDraw.Font.Name
    lngFontSize = objDraw.Font.Size
    blnFontBold = objDraw.Font.Bold
    blnFontItalic = objDraw.Font.Italic
    lngFontColor = objDraw.ForeColor
    lngOldY = objDraw.CurrentY
    
    '��ӡҳ��(0Ϊ����ӡ)
    If intPage <> 0 Then
        objDraw.ForeColor = 0
        objDraw.Font.Name = "����"
        objDraw.Font.Size = 9
        objDraw.Font.Bold = False
        objDraw.CurrentY = lngHeight - lngBottom - objDraw.TextHeight("��")
        objDraw.CurrentX = lngLeft + (lngWidth - lngLeft - lngRight) * (3 / 4)
        objDraw.FontTransparent = True
        objDraw.Print "���� " & intPage & " ҳ��"
    End If
    
    If Not blnPrint Then
        'Ԥ����ӡ����
        objDraw.DrawStyle = 2
        objDraw.Line (0, lngTop)-(lngWidth, lngTop), &H808080
        objDraw.Line (0, lngHeight - lngBottom)-(lngWidth, lngHeight - lngBottom), &H808080
        objDraw.Line (lngLeft, 0)-(lngLeft, lngHeight), &H808080
        objDraw.Line (lngWidth - lngRight, 0)-(lngWidth - lngRight, lngHeight), &H808080
        objDraw.DrawStyle = 0
    End If
    
    '������ҳ
    If blnNewPage Then
        intPage = intPage + 1
        If blnPrint Then
            Printer.NewPage
            Set objDraw = Printer
        Else
            Load objOut.picPage(objOut.picPage.UBound + 1)
            Set objDraw = objOut.picPage(objOut.picPage.UBound)
            objDraw.Width = Printer.Width
            objDraw.Height = Printer.Height
            objDraw.ZOrder
            objDraw.Cls
            objDraw.AutoRedraw = True
        End If
        objDraw.Font.Name = strFontName
        objDraw.Font.Size = lngFontSize
        objDraw.Font.Bold = blnFontBold
        objDraw.Font.Italic = blnFontItalic
        objDraw.ForeColor = lngFontColor
        '��ҳ���
        objDraw.CurrentX = lngLeft: objDraw.CurrentY = lngTop
    Else
        objDraw.CurrentY = lngOldY
    End If
    
    Set NewPrintPage = objDraw
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ExistsPrinter() As Boolean
    Dim lngHDc As Long
    
    If Printers.Count = 0 Then Exit Function
    
    On Error Resume Next
    lngHDc = Printer.hDC
    If Err.Number = 0 Then ExistsPrinter = True
    Err.Clear: On Error GoTo 0
End Function


Public Sub Hook(ByVal frmObject As Object)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Set mfrmTendBody = frmObject
    
    glngPrevWndProc = SetWindowLong(frmObject.hWnd, GWL_WNDPROC, AddressOf WindowProc)

    '��ȡ"�������"�еĹ�������ֵ
    Call SystemParametersInfo(SPI_GETWHEELSCROLLLINES, 0, WHEEL_SCROLL_LINES, 0)

    If WHEEL_SCROLL_LINES > frmObject.BodyEdit.ScrollBarY.Max Then WHEEL_SCROLL_LINES = frmObject.BodyEdit.ScrollBarY.Max
End Sub

Public Sub UnHook(ByVal frmObject As Object)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim lngReturnValue As Long

    lngReturnValue = SetWindowLong(frmObject.hWnd, GWL_WNDPROC, glngPrevWndProc)
    Set mfrmTendBody = Nothing
End Sub

Public Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    '******************************************************************************************************************
    '���ܣ�����ϵͳ�¼������д���
    '������
    '���أ�
    '******************************************************************************************************************
    Dim pt As POINTAPI
    Dim wzDelta
    Dim wKeys As Integer
    
    Select Case uMsg
    Case WM_MOUSEWHEEL                          '�����¼�
        wzDelta = HIWORD(wParam)
        wKeys = LOWORD(wParam)
        pt.X = LOWORD(lParam)
        pt.Y = HIWORD(lParam)
    
        '����Ļ����ת��ΪfrmCaseTendBody��������
    
        ScreenToClient mfrmTendBody.hWnd, pt

        With mfrmTendBody.BodyEdit
        
            '�ж������Ƿ���frmCaseTendBody.BodyEdit������
    
            If pt.X > .Left / Screen.TwipsPerPixelX And pt.X < (.Left + .Width) / Screen.TwipsPerPixelX And pt.Y > .Top / Screen.TwipsPerPixelY And pt.Y < (.Top + .Height) / Screen.TwipsPerPixelY Then
    
                If wKeys = 16 Then
                    'ˮƽ����
                    
                Else
                    '��ֱ����
                    If Sgn(wzDelta) = 1 Then
                        .ScrollBarY.Value = IIf(.ScrollBarY.Value - WHEEL_SCROLL_LINES < .ScrollBarY.Min, .ScrollBarY.Min, .ScrollBarY.Value = .ScrollBarY.Value - WHEEL_SCROLL_LINES)
                    Else
                        .ScrollBarY.Value = IIf(.ScrollBarY.Value + WHEEL_SCROLL_LINES > .ScrollBarY.Max, .ScrollBarY.Max, .ScrollBarY.Value + WHEEL_SCROLL_LINES)
                    End If
                End If
            End If
        End With
    Case Else                                   '�����¼�����ϵͳȱʡ����
        WindowProc = CallWindowProc(glngPrevWndProc, hw, uMsg, wParam, lParam)
    End Select
End Function

Public Function PrintOrPreviewBodyState(objOut As Object, _
                                        ByVal lng����ID As Long, _
                                        ByVal lng��ҳID As Long, _
                                        ByVal lng�ļ�ID As Long, _
                                        ByVal intBaby As Integer, _
                                        ByVal lngSectID As Long, _
                                        ByVal lngBeginY As Long, _
                                        ByVal lngBeginX As Long, _
                                        ByVal objParent As Object, _
                                        Optional ByVal blnKeepOn As Boolean = False, _
                                        Optional ByVal intBeginPage As Integer = -1, _
                                        Optional ByVal intEndPage As Integer = -1, _
                                        Optional ByVal intPageNo As Integer = -1, _
                                        Optional ByVal sngScale As Single = 1, _
                                        Optional ByVal blnMoved As Boolean) As Boolean
    '******************************************************************************************************************
    '����:��ӡ��Ԥ��ĳ������¶ȱ�
    '����:objOut=�������,����ΪPrinter��һ������(�����а����ؼ�����picPage)
    '      lngCaseRecordID=������¼id
    '      lngBeginY=��ʼ������
    '      blnKeepOn=�Ƿ񱣳�����
    '      objParent=�����ô���
    '      intBeginPage=Ҫ��ʼҳ�����,��Ϊ-1ʱ��ʾ�������.
    '      intEndPage=����ҳ������intEndPage����ʵ��ҳ����ֻ��ӡ��ʵ��ҳ��
    '      intPageNO=��ʼ��ҳ��,���Ϊ-1��ʾ����ʾҳ��
    '      sngScale=�������
    
    '����:���δ�ӡ�����Ƿ�ɹ�
    '******************************************************************************************************************
    Dim strSql As String, strNewSql As String
    '�������ò�������
    Dim intOpDays As Integer  '�������ע����
    Dim blnStopFlag As Boolean '�ٴ�����ֹͣǰ�α�ע
    Dim intOpFormat As Integer '��������ȱʡ��ʽ
    Dim bytδ����ʾλ�� As Byte 'δ��˵����ʾλ��
    Dim blnӤ�����µ���ʾ��Ժ As Boolean 'Ӥ�����µ���ʾ��Ժ��Ϣ
    Dim bln���µ���ʾ��� As Boolean '���µ���ʾ���
    Dim intRepairRows As Integer  '�����ʾ����
    Dim bln��ʾƤ�� As Boolean '���µ������ʾƤ�Խ��
    Dim bln��ӡҽԺ���� As Boolean '���µ��Ƿ��ӡҽԺ����
    Dim bln�����ʾ��Ժ As Boolean
    Dim bln���� As Boolean
    Dim bln���ܵ��� As Boolean '���µ���������ʱ������ܽ��컹�ǽ���������������
    Dim bln¼��Сʱ As Boolean '���µ�ȫ���������¼�����ʾ����Сʱ��
    Dim bln��ӡ������� As Boolean  '���µ���ӡʱ�Ƿ��ӡ�������
    Dim bln����ӡ������ As Boolean '���µ���ӡʱ�Ƿ��ӡ������(�������ʵ���Ӧ������Ч��ֻ�ǲ���ӡ�̶��У����������)
    Dim lngCurveRow As Long '�������߹̶��������
    Dim bln��Ժ As Boolean
    
    '������ͼ����
    Dim i As Integer, j As Integer
    Dim lngPicPageIndex As Integer 'Ԥ��ʱPIC������
    Dim blnPrint As Boolean  '�Ƿ��ӡ
    Dim strInfo As String '˵����Ϣ
    Dim intAllOpt As Single  '��ӡ���ܹ�����
    Dim intCurOpt As Single  '��ӡ���е��ڼ���
    Dim objDraw As Object '��ͼ����
    Dim lngHwnd As Long '���
    Dim lngDC As Long  '��ͼ�����DC
    Dim lngFont As Long
    Dim lngOldFont As Long
    Dim stdset As StdFont
    Dim lngLableStep As Long '�̶������п�
    Dim lngColStep As Long ' ���������п�
    Dim lngInitRowStep As Long '���������и�
    
    Dim lngCountPage As Long '����ҳ��
    Dim lngPage As Long
    Dim strBeginDate  As String, strBeginDate1 As String '��ʼʱ��
    Dim strEndDate As String '��ֹʱ��
    Dim strTmpDay As String, strEndDay As String
    Dim dtBegin As Date, dtEnd As Date
    Dim intDrawLineRows As Integer '���µ�����������
    Dim intDrawLineCOL As Integer '���µ��̶���������
    Dim strTmp As String, strTime As String, strTmp1 As String
    Dim lngValue As Long 'סԺ����
    Dim T_Rect As RECT
    Dim rsPart As New ADODB.Recordset  '�������²�λ��Ϣ
    Dim rsTemp As New ADODB.Recordset  '�˼�¼���벻Ҫ˳��ʹ��
    Dim rsTmp As New ADODB.Recordset
    Dim rsItems As New ADODB.Recordset 'ʹ����˲��˵����л�����Ŀ��Ϣ
    Dim rsDrawItems As New ADODB.Recordset '���µ�������Ŀ��Ϣ
    Dim rsPoints As New ADODB.Recordset '�������µ��ļ���
    Dim rsNotes As New ADODB.Recordset   '����˵����Ϣ
    Dim rsDownTab As New ADODB.Recordset '���±��������Ϣ
    Dim H_16pt As Long, W_16pt As Long
    Dim int����Ӧ�� As Integer
    Dim str���ʷ���  As String
    Dim arrTmpValue() As Variant, arrTmpNote() As Variant
    Dim arrValues() As String
    Dim strPart As String '��λ
    Dim SinX As Single, sinY As Single
    Dim intCOl As Integer
    Dim blnAdd As Boolean, blnAllow As Boolean
    Dim dbl��ֵ As Double, dblMinValue As Double, dblMaxValue As Double
    Dim lng��Ŀ��� As Long
    Dim str����˵�� As String
    Dim bln���� As Boolean  '�����Ƿ�Ϊ���
    Dim sngHTab As Single  '���±��߶�
    Dim sngHPrint As Single '�ɴ�ӡ����
    
    Dim strBegin As String, strEnd As String
    Dim str��� As String
    Dim strItemName As String, strItems As String
    Dim intƵ�� As Integer
    Dim intCol1 As Integer
    Dim str��Ŀ���� As String
    Dim int��Ŀ���� As Integer, int��Ŀ���� As Integer, int��Ժ�ײ� As Integer
    Dim int����ѹ As Integer, int����ѹ As Integer, Int�к� As Integer
    Dim blnColor As Boolean

    '���˻�����Ϣ
    Dim strPatiInfo As String
    Dim VarPatiInfo As Variant
    Dim lng����ȼ� As Long
    
    '--������������ �ڼ�¼���²���ʱ����ʱ�������
    Dim strTmpString0 As String  '��¼��ǰʱ��
    Dim strTmpString2 As String '��¼סԺ����
    Dim strTmpString1 As String '��¼����������
    Dim strNewTmpString As String
    Dim ArrNewTmpString() As String '��¼�����Ŀ��������ÿһ��ֵ����Ϣ
    Dim ArrNewString() As String '��¼���б����Ŀ��Ϣ
    Dim intDays As String '��������
    Dim strOpdays(1 To 7) As String
    Dim strOpValue(1 To 7) As String
    Dim arrOperDay
    Dim strEditors() As Variant    '��¼������Ŀ��Ϣ(��Ŀ���||��Ŀ����||��Ŀ��λ||��Ŀֵ��||��¼��||��¼ɫ||���ֵ||��Сֵ||�ٽ�ֵ��
    Dim ArrComTable() As Variant '��¼���еı��±����Ŀ (��Ŀ���||��λ+��Ŀ����|��Ŀ��λ||��Ŀֵ��||��¼Ƶ��||��Ŀ����||��Ŀ��ʾ||��Ժ�ײ�)
    Dim lng���� As Long  '��¼��������
    
    '������Ϣ
    Dim lngLeft As Long, lngTop As Long
    Dim lngRight As Long, lngButtom As Long
    Dim X As Long, Y As Long
    Dim lngCurX As Long, lngCurY As Long
    Dim dblSureW As Double, dblSureH As Double
    
    Dim M_DrawClient As DrawClient
    
    On Error GoTo ErrPrint
    
    msngTwips = 1
    
    mintBaby = intBaby
    '����ԭʼֵ:
    
    M_DrawClient.ƫ����X = T_DrawClient.ƫ����X
    M_DrawClient.ƫ����Y = T_DrawClient.ƫ����Y
    M_DrawClient.�̶����� = T_DrawClient.�̶�����
    M_DrawClient.�̶ȵ�λ = T_DrawClient.�̶ȵ�λ
    M_DrawClient.�������� = T_DrawClient.��������
    M_DrawClient.�е�λ = T_DrawClient.�е�λ
    M_DrawClient.ʱ���е�λ = T_DrawClient.ʱ���е�λ
    M_DrawClient.ʱ���е�λ = T_DrawClient.ʱ���е�λ
    M_DrawClient.�е�λ = T_DrawClient.�е�λ
    M_DrawClient.˫�� = T_DrawClient.˫��
    M_DrawClient.������ = T_DrawClient.������
    
    mintBmpW = gintBmpW
    mintBmpH = gintBmpH
    '��ȡ���²�����Ϣ
    '------------------------------------------------------------------------------------------------------------------
    intOpDays = Val(zldatabase.GetPara("�������ע����", glngSys, 1255, "10"))
    blnStopFlag = (Val(zldatabase.GetPara("�ٴ�����ֹͣǰ�α�ע", glngSys, 1255, "0")) = 1)
    bytδ����ʾλ�� = Abs(Val(zldatabase.GetPara("δ��˵����ʾλ��", glngSys, 1255, "0")))
    blnӤ�����µ���ʾ��Ժ = (zldatabase.GetPara("Ӥ�����µ���ʾ��Ժ��Ϣ", glngSys, 1255, 1) = 1)
    bln���µ���ʾ��� = (zldatabase.GetPara("���µ���ʾ���", glngSys, 1255, 1) = 1)
    intRepairRows = zldatabase.GetPara("���±������", glngSys, 1255, 8)
    bln��ʾƤ�� = (Val(zldatabase.GetPara("���µ���ʾƤ�Խ��", glngSys, 1255, "0")) = 1)
    bln��ӡҽԺ���� = (Val(zldatabase.GetPara("��ӡҽԺ����", glngSys, 1255, "1")) = 1)
    bln���ܵ��� = (Val(zldatabase.GetPara("���ܲ�����ʾ��������", glngSys, 1255, 0)) = 1)
    bln��ӡ������� = (Val(zldatabase.GetPara("����ӡ�������ͼ��", glngSys, 1255, "0")) = 0)
    bln����ӡ������ = (Val(zldatabase.GetPara("���µ�����ӡ������", glngSys, 1255, "0")) = 1)
    lngCurveRow = Val(zldatabase.GetPara("�������߹̶��������", glngSys, 1255, "0"))
    
    '--51282,������,2012-08-03,ȫ�������ʾ¼��ʱ��(DYEYҪ���ֹ�¼�����ʱ��H)
    bln¼��Сʱ = (Val(zldatabase.GetPara("ȫ�������ʾ¼��ʱ��", glngSys, 1255, 0)) = 1)
    
    '51338,������,2012-07-06
    strTmp = zldatabase.GetPara("��������ȱʡ��ʽ", glngSys, 1255, "2")
    If Val(strTmp) >= 0 And Val(strTmp) <= 2 Then
        intOpFormat = Val(strTmp)
    Else
        intOpFormat = 0
    End If
    '���˱䶯�����ʾ����
    '------------------------------------------------------------------------------------------------------------------
    Call InitPara
    
    blnPrint = TypeName(objOut) = "Printer"
    
    '���ڴ�ӡ������Ļ�����ز�ͬ���˴���Ҫȡ���Ե�����
    If blnPrint = True Then
        T_TwipsPerPixel.X = Printer.TwipsPerPixelX
        T_TwipsPerPixel.Y = Printer.TwipsPerPixelY
        msngTwips = Screen.TwipsPerPixelX / Printer.TwipsPerPixelX
        Printer.Font.Size = 9
        Printer.FontName = "����"
    Else
        T_TwipsPerPixel.X = Screen.TwipsPerPixelX
        T_TwipsPerPixel.Y = Screen.TwipsPerPixelY
        msngTwips = 1
    End If
    
    Screen.MousePointer = 11
    intAllOpt = 5
    
    '������ȴ���
    '------------------------------------------------------------------------------------------------------------------
    strInfo = "����" & IIf(blnPrint, "׼����ӡ���±�", "����Ԥ��") & ",���Ժ�..."
    Call ShowFlash(strInfo, , objParent)
    
    '��ӡǰ�����
    If blnKeepOn = False Then
        If Not blnPrint Then
            For i = objOut.picPage.UBound To 0 Step -1
                If i = 0 Then
                    objOut.picPage(i).Cls
                Else
                    Unload objOut.picPage(i)
                End If
            Next
            Set objDraw = objOut.picPage(0)
            objDraw.Width = Printer.Width * sngScale
            objDraw.Height = Printer.Height * sngScale
        Else
            Set objDraw = Printer
        End If
    Else
        If Not blnPrint Then
            i = objOut.picPage.UBound + 1
            Load objOut.picPage(i)
            Set objDraw = objOut.picPage(objOut.picPage.UBound)
            objDraw.Width = Printer.Width * sngScale
            objDraw.Height = Printer.Height * sngScale
        Else
            Set objDraw = Printer
        End If
    End If
    
    bln��Ժ = False
    '��ȡӤ��ҽ����Ϣ(ת�ƣ���Ժ)����ҽ����ҽ����ϢΪ׼��������ĸ�׳�Ժ����Ϊ׼
    strNewSql = "   (SELECT /*+ RULE */  ����ID,��ҳID,Ӥ��ʱ��,DECODE(nvl(Ӥ��,0),0, DECODE(NVL(��Ժ����,''),'',0,1), DECODE(NVL(Ӥ��ʱ��,''),'',0,1))��¼" & vbNewLine & _
                "       FROM (SELECT A.����ID,A.��ҳID,B.��ʼִ��ʱ�� Ӥ��ʱ��, A.��Ժ����,B.Ӥ��" & vbNewLine & _
                "           FROM ������ҳ A," & vbNewLine & _
                "               (SELECT B.����ID, B.��ҳID, B.Ӥ��, ��ʼִ��ʱ��" & vbNewLine & _
                "                FROM ����ҽ����¼ B, ������ĿĿ¼ C" & vbNewLine & _
                "                WHERE B.������ĿID + 0 = C.ID AND B.ҽ��״̬ = 8 AND nvl(B.Ӥ��,0)<>0  AND C.��� = 'Z'" & vbNewLine & _
                "                AND EXISTS (SELECT 1 FROM TABLE(CAST(F_STR2LIST('3,5,11') AS ZLTOOLS.T_STRLIST))" & vbNewLine & _
                "                               WHERE C.�������� = COLUMN_VALUE) And  B.����ID = [2] AND B.��ҳID = [3] AND B.Ӥ��(+) = [4]) B" & vbNewLine & _
                "           WHERE A.����ID = [2] AND A.��ҳID = [3] AND A.����ID = B.����ID(+) AND A.��ҳID = B.��ҳID(+)" & vbNewLine & _
                "           ORDER BY B.��ʼִ��ʱ�� DESC)" & vbNewLine & _
                "       WHERE ROWNUM < 2)  E"
    '��ȡ���˳�Ժǰ��ʱ����Ϣ
    '------------------------------------------------------------------------------------------------------------------
    strSql = _
       "Select Decode(b.����ʱ��,Null,a.��ʼ,b.����ʱ��) As ��ʼ,decode(E.��¼,0,Decode(Sign(NVL(E.Ӥ��ʱ��,a.��ֹ) - d.����ʱ��), 1,NVL(E.Ӥ��ʱ��,a.��ֹ) ,d.����ʱ��),NVL(E.Ӥ��ʱ��,a.��ֹ)) ��ֹ,E.��¼" & vbNewLine & _
        "       From" & vbNewLine & _
        "       (Select ����ID,��ҳid,Min(��ʼʱ��) as ��ʼ,Max(Nvl(��ֹʱ��,sysdate)) as ��ֹ" & vbNewLine & _
        "       From ���˱䶯��¼" & vbNewLine & _
        "       Where ��ʼʱ�� is Not Null And ����ID=[2] And ��ҳID=[3] Group By ����ID,��ҳid) a," & vbNewLine & _
        "       (Select ����ID,��ҳid,����ʱ�� From ������������¼ Where ����ID =[2] And ��ҳID =[3] And ���=[4]) b," & vbNewLine & _
        "       (SELECT NVL(����ʱ��,SYSDATE) ����ʱ�� FROM (select max(����ʱ��) ����ʱ�� from ���˻����ļ� A,���˻������� B" & vbNewLine & _
        "       where A.ID=B.�ļ�ID and A.ID=[1] and A.����ID=[2] and A.��ҳID=[3] and A.Ӥ��=[4])) d," & vbNewLine & _
        strNewSql & vbNewLine & _
        "       Where A.����ID=E.����ID And A.��ҳID=E.��ҳID And a.����id=b.����id(+) And a.��ҳid=b.��ҳid(+)"
        
    Set rsTemp = zldatabase.OpenSQLRecord(strSql, "mdlPrint", lng�ļ�ID, lng����ID, lng��ҳID, intBaby)
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
        lngCountPage = DateDiff("d", rsTemp!��ʼ, rsTemp!��ֹ) + 1
        lngCountPage = IIf(lngCountPage / 7 = Fix(lngCountPage / 7), lngCountPage / 7, Fix(lngCountPage / 7) + 1)
        strBeginDate = Format(rsTemp!��ʼ, "YYYY-MM-DD HH:MM:SS")
        strBeginDate1 = strBeginDate
        strEndDate = Format(rsTemp!��ֹ, "YYYY-MM-DD HH:MM:SS")
        bln��Ժ = Not (Val(rsTemp!��¼) = 0)
    Else
        CloseRs rsTemp
        GoTo ErrPrint '�������˱䶯��Ϣ�˳�
    End If
    
    gbln��Ժ = bln��Ժ
    '��ȡ�û����õ����µ���ʼʱ��(Ӥ���Գ���ʱ��Ϊ׼)
    If intBaby = 0 Then
        strSql = "select ��ʼʱ�� from ���˻����ļ� where ID=[1] and ����ID=[2] and ��ҳid=[3] and nvl(Ӥ��,0)=[4]"
        Set rsTmp = zldatabase.OpenSQLRecord(strSql, "��ȡ���µ���ʼʱ��", lng�ļ�ID, lng����ID, lng��ҳID, intBaby)
        If rsTmp.RecordCount <> 0 Then
            strBeginDate = Format(rsTmp!��ʼʱ��, "YYYY-MM-DD HH:mm:ss")
        End If
    End If
    
    If bln��Ժ = True Then
        '��Ժʱ�����Ժʱ�������ͬһ�У��򽫳�Ժʱ�����һ�У���������:��ԺҲҪ¼�����£�
        strEndDate = Format(RetrunEndTime(CDate(strBeginDate), CDate(strEndDate), gintHourBegin), "YYYY-MM-DD HH:mm:ss")
    End If
    
    bln�����ʾ��Ժ = False
    
    If CDate(Format(strBeginDate, "YYYY-MM-DD HH:MM:SS")) > CDate(Format(strBeginDate1, "YYYY-MM-DD HH:MM:SS")) Then
        bln�����ʾ��Ժ = True
    ElseIf T_BodyFlag.��Ժ = 0 And CDate(Format(strBeginDate, "YYYY-MM-DD HH:MM:SS")) = CDate(Format(strBeginDate1, "YYYY-MM-DD HH:MM:SS")) Then
        bln�����ʾ��Ժ = True
    End If
            
    intCurOpt = intCurOpt + 1
    
    Call ShowFlash(strInfo, intCurOpt / intAllOpt, objParent)
    
    '------------------------------------------------------------------------------------------------------------------
    '��1���ݣ����˵Ļ�����Ϣ
    '��ȡ���˻�����Ϣ
    
    '"����'����'�Ա�'�Ʊ�'����'��Ժ����'סԺ��:
    strPatiInfo = "''''''"
    VarPatiInfo = Split(strPatiInfo, "'")
    
    strSql = " Select  b.����,A.סԺ��,A.��Ժ���� ��Ժʱ��,b.�Ա�,A.���� From ������Ϣ B,������ҳ A Where A.����ID=B.����ID And A.����id=[1] And A.��ҳID=[2]"
    Set rsTemp = zldatabase.OpenSQLRecord(strSql, "mdlPrint", lng����ID, lng��ҳID)
    If rsTemp.BOF = False Then
        VarPatiInfo(0) = zlCommFun.Nvl(rsTemp("����").Value)
        VarPatiInfo(6) = zlCommFun.Nvl(rsTemp("סԺ��").Value)
        VarPatiInfo(5) = Format(zlCommFun.Nvl(rsTemp("��Ժʱ��").Value), "yyyy-MM-dd")
        VarPatiInfo(2) = zlCommFun.Nvl(rsTemp("�Ա�").Value)
        VarPatiInfo(1) = zlCommFun.Nvl(rsTemp("����").Value)
    End If
    
    '��Ժʱ��(������µ���ʼʱ�������Ժʱ��������ʱ��Ϊ׼)
    strSql = "select ��ʼʱ�� from ���˱䶯��¼ where ����id=[1] And ��ҳid=[2] and ��ʼԭ��=2 order by ��ʼʱ��"
    Set rsTemp = zldatabase.OpenSQLRecord(strSql, "mdlPrint", lng����ID, lng��ҳID)
    If rsTemp.BOF = False Then
        If bln�����ʾ��Ժ = True Then
            VarPatiInfo(5) = Format(zlCommFun.Nvl(rsTemp("��ʼʱ��").Value), "yyyy-MM-dd")
        End If
    End If
    
    If intBaby <> 0 Then
        
        VarPatiInfo(1) = ""
        VarPatiInfo(2) = ""
        
        strSql = "Select Decode(a.Ӥ������,Null,b.����||'֮��'||Trim(To_Char(a.���,'9')),a.Ӥ������) As Ӥ������,Ӥ���Ա�,����ʱ�� " & _
            " From ������������¼ a,������Ϣ b " & _
            " Where a.����id=[1] And a.��ҳid=[2] And a.����id=b.����id And a.���=[3]"
        Set rsTemp = zldatabase.OpenSQLRecord(strSql, "mdlPrint", lng����ID, lng��ҳID, intBaby)
        If rsTemp.BOF = False Then
            VarPatiInfo(0) = rsTemp("Ӥ������").Value
            VarPatiInfo(2) = zlCommFun.Nvl(rsTemp("Ӥ���Ա�").Value)
            VarPatiInfo(1) = "������"
            If IsNull(rsTemp("����ʱ��").Value) = False Then VarPatiInfo(5) = Format(zlCommFun.Nvl(rsTemp("����ʱ��").Value), "yyyy-MM-dd")
        End If
        
    End If
    
    If bln���µ���ʾ��� Then ReDim Preserve VarPatiInfo(UBound(VarPatiInfo) + 1)
    
    intCurOpt = intCurOpt + 1
    Call ShowFlash(strInfo, intCurOpt / intAllOpt, objParent)
    
    '��ȡ���˻���ȼ�
    strSql = "Select zl_PatitTendGrade([1],[2]) As ����ȼ� From dual"
    Set rsTemp = zldatabase.OpenSQLRecord(strSql, "����ȼ�", lng����ID, lng��ҳID)
    If rsTemp.BOF = False Then lng����ȼ� = zlCommFun.Nvl(rsTemp("����ȼ�"), 0)
    
    '��ȡ���ü�¼��
    Call InitPublicData
    
    '�������Ӧ�÷�ʽ
    int����Ӧ�� = 2
    str���ʷ��� = ""
    strSql = "Select a.Ӧ�÷�ʽ,b.��¼�� From �����¼��Ŀ a,���¼�¼��Ŀ b Where a.��Ŀ���=-1 And a.��Ŀ���=b.��Ŀ���"
    Set rsTemp = zldatabase.OpenSQLRecord(strSql, "mdlPrint")
    If rsTemp.BOF = False Then
        int����Ӧ�� = zlCommFun.Nvl(rsTemp("Ӧ�÷�ʽ").Value, 2)
        str���ʷ��� = zlCommFun.Nvl(rsTemp("��¼��").Value, "��")
    Else
        int����Ӧ�� = 0
    End If
    
    Dim int���� As Integer, int���� As Integer
    
    '-------------------------------------------------------------------------------------------------------------------
    '2��ȡ����������Ŀ(�����µ��̶�������������������-2)
    strSql = " Select A.��Ŀ���,A.�������,A.��¼��,C.��Ŀֵ��,A.��¼��,A.��¼ɫ,nvl(A.���ֵ,0) ���ֵ ,nvl(A.��Сֵ,0) ��Сֵ,A.�ٽ�ֵ," & _
        "nvl(A.��λֵ,0) ��λֵ,A.�̶ȼ��,A.��ʾ��,C.��Ŀ��λ ��λ,nvl(A.�����,2)-2 AS �����,B.��λ " & _
        " From ���¼�¼��Ŀ A,���²�λ B,�����¼��Ŀ C" & _
        " Where A.��Ŀ���=B.��Ŀ���(+) And B.ȱʡ��(+)=1" & _
        " And A.��¼��=1 And A.��Ŀ���=C.��Ŀ��� and nvl(C.Ӧ�÷�ʽ,0)=1 and C.����ȼ�>=[1]" & _
        " and nvl(C.���ò���,0) in (0,[2]) and (C.���ÿ���=1 or (C.���ÿ���=2 and Exists (select 1 from �������ÿ��� D where C.��Ŀ���=D.��Ŀ��� and D.����ID=[3])))" & _
        " Order by �������"
    Set rsTemp = zldatabase.OpenSQLRecord(strSql, "��ȡ����������Ŀ", lng����ȼ�, IIf(intBaby = 0, 1, 2), lngSectID)
    
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
        rsTemp.Filter = "��Ŀ���=" & gint����
        If rsTemp.RecordCount > 0 And bln����ӡ������ Then
            rsTemp.Filter = 0
            intDrawLineCOL = rsTemp.RecordCount - 1
        Else
            rsTemp.Filter = 0
            intDrawLineCOL = rsTemp.RecordCount
        End If
        If intDrawLineCOL <= 0 Then intDrawLineCOL = 1
    Else
        CloseRs rsTemp
        MsgBox "���κ�����������Ŀ��", vbExclamation, gstrSysName
        GoTo ErrExit
    End If
    strEditors = Array()
    int���� = -1: int���� = -1
    rsTemp.Filter = 0
    rsTemp.Sort = "�������"
    With rsTemp
        Do While Not .EOF
            strTmp = zlCommFun.Nvl(!��Ŀ���, 0) & "|| " & zlCommFun.Nvl(!��¼��) & "|| " & zlCommFun.Nvl(!��λ) & "|| " & zlCommFun.Nvl(!��Ŀֵ��) & "|| " & _
                 zlCommFun.Nvl(!��¼��) & "|| " & zlCommFun.Nvl(!��¼ɫ) & "||" & zlCommFun.Nvl(!���ֵ) & "||" & zlCommFun.Nvl(!��Сֵ) & "||" & zlCommFun.Nvl(!�ٽ�ֵ)
                
            ReDim Preserve strEditors(UBound(strEditors) + 1)
            strEditors(UBound(strEditors)) = strTmp
            If zlCommFun.Nvl(!��Ŀ���, 0) = gint���� Then
                int���� = UBound(strEditors)
            End If
        .MoveNext
        Loop
        .MoveFirst
    End With
    If int����Ӧ�� = 2 And int���� <> -1 Then
        ReDim Preserve strEditors(UBound(strEditors) + 1)
        strTmp = "-1||����||" & Split(strEditors(int����), "||")(2) & "||" & Split(strEditors(int����), "||")(3) & "||��||" & RGB_RED & "||" & _
            Split(strEditors(int����), "||")(6) & "||" & Split(strEditors(int����), "||")(7) & "||" & Split(strEditors(int����), "||")(8)
        strEditors(UBound(strEditors)) = strTmp
    End If
    
    intCurOpt = intCurOpt + 1
    Call ShowFlash(strInfo, intCurOpt / intAllOpt, objParent)
    
    '3����ȡ����������Ŀ��Ϣ�������Ŀ�����Ŀ���ܴ���һ����Ŀ�����λҲҪ��ȡ��
    ArrComTable = Array()
    strTmp = ""
    strTime = ""
    
    '��ȡ���ǻ�����Ŀ
    gstrSQL = "Select A.�������,A.��Ŀ���,A.��¼��,A.��¼��,A.��¼��,A.��¼ɫ,B.��Ŀֵ��,nvl(A.��¼Ƶ��,2) ��¼Ƶ��,A.��Ժ�ײ�,B.��Ŀ����," & _
        "   B.��Ŀ����,B.��Ŀ����,B.��Ŀ��ʾ,B.��ĿС��,B.��Ŀ��λ ��λ" & _
        "   From ���¼�¼��Ŀ A,�����¼��Ŀ B,����������Ŀ C" & _
        "   Where A.��Ŀ���=B.��Ŀ��� And B.��ĿID=C.Id(+)  And A.��¼��=2" & _
        "   And nvl(B.Ӧ�÷�ʽ,0)=1 And nvl(B.����ȼ�,0)>=[7] And nvl(B.���ò���,0) In (0,[8])" & _
        "   And (B.���ÿ���=1 Or (B.���ÿ���=2 And Exists (Select 1 From �������ÿ��� D Where D.��Ŀ���=B.��Ŀ��� And D.����id=[9])))"
        
    
    strSql = "Select Rownum-1 ��� ,��Ŀ���,��Ŀ����,��¼ɫ,��Ŀ��λ,��Ŀֵ��, ��λ,��¼Ƶ��,��Ժ�ײ�,��Ŀ����,��Ŀ��ʾ,��Ŀ���� From (" & _
            " Select A.��Ŀ���, Decode(A.��Ŀ���, 4, 'Ѫѹ', A.��¼��) ��Ŀ����,A.��¼ɫ,A.��λ ��Ŀ��λ, A.��Ŀֵ��, B.��λ," & vbNewLine & _
            "           nvl(A.��¼Ƶ��,2) ��¼Ƶ��,A.��Ժ�ײ�, nvl(A.��Ŀ����,1) ��Ŀ����, A.��Ŀ��ʾ,A.��Ŀ����" & vbNewLine & _
            " From (" & gstrSQL & " ) A," & vbNewLine & _
             "        (Select Distinct b.��Ŀ���, a.��λ" & vbNewLine & _
            "            From (Select ��Ŀ���, DECODE(��Ŀ���,3,'',���²�λ) ��λ" & vbNewLine & _
            "                           From ���˻����ļ� a, ���˻������� b, ���˻�����ϸ c" & vbNewLine & _
            "                           Where a.Id = b.�ļ�id And b.Id = c.��¼id And a.Id = [1] And Nvl(a.Ӥ��, 0) = [4] And a.����id = [2] And" & vbNewLine & _
            "                                       a.��ҳid = [3] And c.��¼���� = 1 And b.����ʱ�� Between [5] And [6] And ��ֹ�汾 Is Null) a, ���¼�¼��Ŀ b," & vbNewLine & _
            "                       �����¼��Ŀ c" & vbNewLine & _
            "            Where b.��Ŀ��� = a.��Ŀ���(+) And b.��Ŀ��� = c.��Ŀ��� And b.��¼�� = 2 And Nvl(����ȼ�, 0) >=[7]) B" & vbNewLine & _
            "   where A.��Ŀ���=B.��Ŀ��� and A.��Ŀ���<>5  order by Decode(A.��Ŀ���,3 ,0,1 ),A.�������,��Ŀ����,B.��λ)"

    If blnMoved Then
        strSql = Replace(strSql, "���˻����ļ�", "H���˻����ļ�")
        strSql = Replace(strSql, "���˻�������", "H���˻�������")
        strSql = Replace(strSql, "���˻�����ϸ", "H���˻�����ϸ")
    End If
    
    Set rsItems = zldatabase.OpenSQLRecord(strSql, "ȡ��ʼ��", lng�ļ�ID, lng����ID, lng��ҳID, intBaby, Int(CDate(strBeginDate)), CDate(strEndDate), lng����ȼ�, IIf(intBaby = 0, 1, 2), lngSectID)
    
    bln���� = False
    With rsItems
        Do While Not .EOF
            str��Ŀ���� = ""
            If Val(Nvl(!��Ŀ����, 1)) = 2 Then
                str��Ŀ���� = Trim(Nvl(!��λ)) & Nvl(!��Ŀ����)
            Else
                str��Ŀ���� = Nvl(!��Ŀ����)
            End If
            
            intƵ�� = Val(zlCommFun.Nvl(!��¼Ƶ��))
            
            If zlCommFun.Nvl(!��Ŀ��ʾ) = 4 Or IsWaveItem(Val(zlCommFun.Nvl(!��Ŀ���))) Then
                If intƵ�� > 2 Then intƵ�� = 2
            End If
            
            strTmp = zlCommFun.Nvl(!��Ŀ���) & "||" & Replace(str��Ŀ����, ";", ":") & "||" & zlCommFun.Nvl(!��Ŀ��λ) & "||" & _
                zlCommFun.Nvl(!��Ŀֵ��) & "||" & intƵ�� & "||" & zlCommFun.Nvl(!��Ŀ����, 1) & "||" & _
                zlCommFun.Nvl(!��Ŀ��ʾ) & "||" & zlCommFun.Nvl(!��Ŀ����) & "||" & zlCommFun.Nvl(!��Ժ�ײ�, 0)
            If Val(zlCommFun.Nvl(!��Ŀ���)) = gint���� Then
                bln���� = True
            End If
            
            ReDim Preserve ArrComTable(UBound(ArrComTable) + 1)
            ArrComTable(UBound(ArrComTable)) = strTmp
        .MoveNext
        Loop
    End With

    If rsItems.RecordCount > 0 Then rsItems.MoveFirst
    
    intCurOpt = intCurOpt + 1
    Call ShowFlash(strInfo, intCurOpt / intAllOpt, objParent)
    '------------------------------------------------------------------------------------------------------------------
    '4��ȷ��X��Y������λ��
    '�߽���Ϣ(Twip)
    
    Dim lngOffsetLeft As Long
    Dim lngOffsetTop As Long
    
    dblSureH = 0
    dblSureW = 0
    If blnPrint = True Then
        '����Ǵ�ӡԤ��,Ӧ����ӡ���Ŀɴ�ӡ�Ŀ�ʼ����ʼԤ��
        dblSureW = Round(GetDeviceCaps(Printer.hDC, PHYSICALOFFSETX) / GetDeviceCaps(Printer.hDC, PHYSICALWIDTH), 4)
        dblSureH = Round(GetDeviceCaps(Printer.hDC, PHYSICALOFFSETY) / GetDeviceCaps(Printer.hDC, PHYSICALHEIGHT), 4)
        On Error Resume Next
        dblSureH = (objDraw.Height * dblSureH) / T_TwipsPerPixel.Y
        dblSureW = (objDraw.Width * dblSureW) / T_TwipsPerPixel.X
    End If

    lngRight = gPrinter.lngRight
    lngButtom = gPrinter.lngBottom
     
    lngRight = lngRight * (conRatemmToTwip / T_TwipsPerPixel.X) * sngScale + dblSureW
    lngButtom = lngButtom * (conRatemmToTwip / T_TwipsPerPixel.Y) * sngScale
    lngLeft = lngBeginX * (conRatemmToTwip / T_TwipsPerPixel.X) * sngScale - dblSureW
    lngTop = (lngBeginY / T_TwipsPerPixel.Y) * sngScale
    
    H_16pt = objDraw.TextHeight("��") / T_TwipsPerPixel.Y
    W_16pt = objDraw.TextWidth("��") / T_TwipsPerPixel.X
    
    X = lngLeft: Y = lngTop
    lngCurX = X: lngCurY = Y
    
    If intDrawLineCOL <= 3 Then
        lngLableStep = (glngLableWith / intDrawLineCOL) * sngScale * msngTwips
    Else
        lngLableStep = glngLableStep * sngScale * msngTwips
    End If
    
    T_DrawClient.�̶�����.Left = lngCurX
    T_DrawClient.�̶�����.Right = lngCurX + intDrawLineCOL * lngLableStep
    
    If W_16pt > glngColStep * msngTwips Then
        lngColStep = W_16pt * sngScale
    Else
        lngColStep = glngColStep * sngScale * msngTwips
    End If
    
    If H_16pt > glngInitRowStep Then
        lngInitRowStep = (H_16pt * sngScale / 2)
    Else
        lngInitRowStep = glngInitRowStep * sngScale
    End If
    
    T_DrawClient.��������.Left = T_DrawClient.�̶�����.Right
    T_DrawClient.��������.Right = T_DrawClient.�̶�����.Right + (6 * 7 * lngColStep)
    
    Dim sigSign As Single
    sigSign = 1
    If T_DrawClient.��������.Right > objDraw.Width / T_TwipsPerPixel.X - lngRight Then
        sigSign = Round((T_DrawClient.��������.Right - (objDraw.Width / T_TwipsPerPixel.X - lngRight)) / (T_DrawClient.��������.Right - T_DrawClient.�̶�����.Right), 2)
        sigSign = Round((1 - sigSign), 2)
        If sigSign < 0.8 Then sigSign = 0.8
        lngLableStep = Fix(lngLableStep * sigSign)
        lngColStep = Fix(lngColStep * sigSign)
    End If
    
    If lngColStep < W_16pt Then lngColStep = W_16pt
    
    If lngColStep < gintBmpW Then
        mintBmpW = lngColStep
        mintBmpH = lngColStep
    End If
    
    T_DrawClient.�̶ȵ�λ = lngLableStep
    T_DrawClient.�̶�����.Right = lngCurX + intDrawLineCOL * lngLableStep
    T_DrawClient.��������.Left = T_DrawClient.�̶�����.Right
    T_DrawClient.��������.Right = T_DrawClient.�̶�����.Right + (6 * 7 * lngColStep)
    T_DrawClient.�е�λ = lngColStep
    T_DrawClient.�е�λ = lngInitRowStep
    T_DrawClient.ʱ���е�λ = 16 * msngTwips
    T_DrawClient.ƫ����X = lngLeft
    '------------------------------------------------------------------------------------------------------------------
    '������п�����߱���ܹ��ж�����
    '������±���Ŀ��������
    strSql = "Select Count(A.��Ŀ���) ��¼�� " & _
        "   From ���¼�¼��Ŀ A,�����¼��Ŀ B " & _
        "   Where A.��Ŀ���=B.��Ŀ��� And A.��¼��=[1]" & _
        "   And nvl(B.Ӧ�÷�ʽ,0)=1 And nvl(B.����ȼ�,0)>=[2] And nvl(B.���ò���,0) In (0,[3])" & _
        "   And (B.���ÿ���=1 Or (B.���ÿ���=2 And Exists (Select 1 From �������ÿ��� D Where D.��Ŀ���=B.��Ŀ��� And D.����id=[4])))"
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, "mdlPrint", 1, lng����ȼ�, IIf(intBaby = 0, 1, 2), lngSectID)
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        intDrawLineRows = zlCommFun.Nvl(rsTmp!��¼��, 0)
    Else
        CloseRs rsTmp
        GoTo ErrPrint
    End If
    
    If intDrawLineRows < 1 Then
        CloseRs rsTmp
        GoTo ErrPrint
    End If
    
    strSql = "Select nvl(A.���ֵ,0) ���ֵ,nvl(A.��Сֵ,0) ��Сֵ ,nvl(A.��λֵ,0.1) ,nvl(A.�����,0)-2  �����" & _
        "   From ���¼�¼��Ŀ A,�����¼��Ŀ B" & _
        "   Where A.��Ŀ���=B.��Ŀ��� And A.��Ŀ���=[1]" & _
        "   And nvl(B.Ӧ�÷�ʽ,0)=1 And nvl(B.����ȼ�,0)>=[2] And nvl(B.���ò���,0) In (0,[3])" & _
        "   And (B.���ÿ���=1 Or (B.���ÿ���=2 And Exists (Select 1 From �������ÿ��� D Where D.��Ŀ���=B.��Ŀ��� And D.����id=[4])))"
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, "mdlPrint", gint����, lng����ȼ�, IIf(intBaby = 0, 1, 2), lngSectID)
    If rsTmp.RecordCount > 0 Then
        '�޸����⣺51442
        dbl��ֵ = Val(zlCommFun.Nvl(rsTmp!��Сֵ, 0))
        intDrawLineRows = (Val(rsTmp!���ֵ) - IIf(dbl��ֵ > 34, 35, dbl��ֵ)) / 0.1 + IIf(Val(rsTmp!�����) < 0, 0, Val(rsTmp!�����)) + IIf(dbl��ֵ > 34, 10, 0)
        intDrawLineRows = intDrawLineRows + lngCurveRow
    End If
    
    If intDrawLineRows > glngMaxRows Then
        T_DrawClient.������ = intDrawLineRows
    Else
        T_DrawClient.������ = glngMaxRows
    End If
    
    intCurOpt = intCurOpt + 1
    Call ShowFlash(strInfo, intCurOpt / intAllOpt, objParent)
    '5��ѭ��������ҳ��ѭ��
    intCurOpt = 0
    intAllOpt = 100
    intCurOpt = intCurOpt + 1
    Call ShowFlash(strInfo, intCurOpt / intAllOpt, objParent)
    If blnPrint = False Then
        lngPicPageIndex = objOut.picPage.UBound + 1
    End If
    
    '��ʽ��ʼ���Ĳ���ѭ��ÿһҳ
    '------------------------------------------------------------------------------------------------------------------
    For lngPage = 1 To lngCountPage
        strTmpDay = Format(CDate(strBeginDate) + 7 * (lngPage - 1), "YYYY-MM-DD") '��õ�ǰҳ��ĵ�һ��������ʱ��
        If CDate(strTmpDay) < CDate(strBeginDate) Then strTmpDay = strBeginDate
        If CDate(strEndDate) < CDate(Format(zldatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")) And Not bln��Ժ Then strEndDate = Format(zldatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")
        strEndDay = Format(CDate(strTmpDay) + 6, "YYYY-MM-DD") & " 23:59:59"
        If CDate(strEndDay) > CDate(strEndDate) Then strEndDay = Format(strEndDate, "YYYY-MM-DD HH:mm:ss")
        intCurOpt = lngPage / lngCountPage
        strInfo = "����" & IIf(blnPrint, "��ӡ���±�", "Ԥ��") & ",���Ժ�..."
        Call ShowFlash(strInfo, intCurOpt, objParent)
        
        '��ҳ�Ŵ�ӡ
        If intBeginPage > 0 Then  'ֻ��ӡָ��ҳ���
            If lngPage >= intBeginPage And lngPage <= intEndPage Then
                If lngPage > intBeginPage Then  '���ڶ�ҳʱ��ʼ��ʼ��ֽ�Ż�ҳ��
                    If Not blnPrint Then
                        Load objOut.picPage(lngPicPageIndex)
                        Set objDraw = objOut.picPage(lngPicPageIndex)
                        objDraw.Cls
                        objDraw.Width = Printer.Width * sngScale
                        objDraw.Height = Printer.Height * sngScale
                        lngPicPageIndex = lngPicPageIndex + 1
                    Else
                        Printer.NewPage
                    End If
                End If
            Else
                GoTo NOPageSub
            End If
        Else  '��ӡ����ʱ
            If lngPage > 1 Then
                If Not blnPrint Then
                    Load objOut.picPage(lngPicPageIndex)
                    Set objDraw = objOut.picPage(lngPicPageIndex) 'PictureBox
                    objDraw.Cls
                    objDraw.Width = Printer.Width * sngScale
                    objDraw.Height = Printer.Height * sngScale
                    lngPicPageIndex = lngPicPageIndex + 1
                Else
                    Printer.NewPage
                End If
            End If
        End If
        
         'ҳüͼ�����
        Call frmTendFileRead.PrintRTBData(objDraw, True, lngTop)
        
        '��ȡ�����Dc
        lngDC = objDraw.hDC
        '��������
        Set stdset = New StdFont
        stdset.Name = "����"
        stdset.Size = 9 * sngScale
        Call SetFontIndirect(stdset, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        '��ӡ�ʿغ�
        strTmp = zldatabase.GetPara("�ʿغ�", glngSys, 1255, "")
        Call GetTextExtentPoint32(lngDC, strTmp, Len(strTmp), T_Size)
        T_Size.W = objDraw.TextWidth(strTmp) / T_TwipsPerPixel.X
        lngCurX = T_DrawClient.��������.Right - T_Size.W
        Call GetTextRect(objDraw, lngCurX, lngCurY, strTmp, , , , sngScale)
        Call DrawText(lngDC, strTmp, -1, T_LableRect, DT_CENTER)
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
        
        '�Ƿ��ӡҽԺ���ƣ��е�ҽԺ���µ�ҽԺ�����ܴ�����������Ҫ��ҳü��ʵ�֡���ʱ�Ͳ��ڴ�ӡע���ļ��е�ҽԺ��Ϣ��
        If bln��ӡҽԺ���� = True Then
            '��ȡҽԺ����
            stdset.Name = "����"
            stdset.Size = 18 * sngScale
            stdset.Bold = True
            Call SetFontIndirect(stdset, lngDC, objDraw)
            lngFont = CreateFontIndirect(T_Font)
            lngOldFont = SelectObject(lngDC, lngFont)
            strTmp = IIf(GetUnitName = "-", "", GetUnitName) & IIf(intBaby <> 0, "Ӥ��", "") & "���µ�"
            Call GetTextExtentPoint32(lngDC, strTmp, Len(strTmp), T_Size)
            lngCurY = T_Size.H + lngCurY
            Call GetTextRect(objDraw, 0, lngCurY, strTmp, objDraw.Width / T_TwipsPerPixel.X, True, T_Size.H, sngScale)
            Call DrawText(lngDC, strTmp, -1, T_LableRect, DT_CENTER)
            Call SelectObject(lngDC, lngOldFont)
            Call DeleteObject(lngFont)
            objDraw.Font.Size = 9 * sngScale
            Y = lngCurY + T_Size.H + 15 * msngTwips
        Else
            Y = lngCurY + 15 * msngTwips
        End If
        lngCurX = X
        lngCurY = Y
        '��ȡ���˿��ҡ����ŵ���Ϣ
    
        VarPatiInfo(3) = ""
        VarPatiInfo(4) = ""
        strTmp = "": strTime = ""
        strSql = " Select  c.���� As ����,b.���� As ����,a.����,a.��ʼԭ�� " & _
                    " From ���˱䶯��¼ a,���ű� b,���ű� c " & _
                    " Where a.����id=[1] And a.��ҳid=[2] And a.����id Is Not Null And a.����id=b.id and a.����id=c.id " & _
                    " And a.��ʼʱ��-4/24<=[3] And Nvl(a.��ֹʱ��,Sysdate)>=[4] Order By a.��ʼʱ��"
        
        Set rsTmp = zldatabase.OpenSQLRecord(strSql, "��ȡ���˿��ҡ����ŵ���Ϣ", lng����ID, lng��ҳID, CDate(strEndDay), CDate(strTmpDay))
        If rsTmp.BOF = False Then
            Do While Not rsTmp.EOF
                
                If zlCommFun.Nvl(rsTmp("����").Value) <> strTmp And zlCommFun.Nvl(rsTmp("����").Value) <> "" Then
                
                    strTmp = zlCommFun.Nvl(rsTmp("����").Value)
                    
                    If VarPatiInfo(3) = "" Then
                        VarPatiInfo(3) = strTmp
                    Else
                        VarPatiInfo(3) = VarPatiInfo(3) & "->" & strTmp
                    End If
                    
                End If
    
                If zlCommFun.Nvl(rsTmp("����").Value) <> strTime And zlCommFun.Nvl(rsTmp("����").Value) <> "" Then
                
                    strTime = zlCommFun.Nvl(rsTmp("����").Value)
                    
                    If VarPatiInfo(4) = "" Then
                        VarPatiInfo(4) = strTime
                    Else
                        VarPatiInfo(4) = VarPatiInfo(4) & "->" & strTime
                    End If
                    
                End If
                            
                rsTmp.MoveNext
            Loop
            
            If Left(VarPatiInfo(3), 2) = "->" Then VarPatiInfo(3) = Mid(VarPatiInfo(3), 3)
            If Left(VarPatiInfo(4), 2) = "->" Then VarPatiInfo(4) = Mid(VarPatiInfo(4), 3)
        End If
        
        If bln���µ���ʾ��� Then
            '��ȡ���������Ϣ
            strSql = "Select Zl_Replace_Element_Value([1],[2],[3],2,NULL,0,[4]) As ������ From Dual"
            Set rsTmp = zldatabase.OpenSQLRecord(strSql, "������", "������", lng����ID, lng��ҳID, CDate(strTmpDay))
            If rsTmp.BOF = False Then
                If intBaby = 0 Then
                    VarPatiInfo(UBound(VarPatiInfo)) = zlCommFun.Nvl(rsTmp("������").Value)
                Else
                    VarPatiInfo(UBound(VarPatiInfo)) = ""
                End If
            Else
                VarPatiInfo(UBound(VarPatiInfo)) = ""
            End If
        End If
        strPatiInfo = Join(VarPatiInfo, "'")
        
        stdset.Name = "����"
        stdset.Size = 9 * sngScale
        stdset.Bold = True
        Call SetFontIndirect(stdset, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        '���������Ϣ
        Call DrawPatiInfo(lngDC, objDraw, strPatiInfo, lngCurX, lngCurY, T_DrawClient.��������.Right, lngCurY, sngScale)
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
        '---��ʼ�����µ��ϱ��(סԺ����,סԺ����,����,ʱ��)
        Y = lngCurY: lngCurX = X: lngCurY = Y
        '1.��ȡסԺ��ʼ����
        lngValue = 0: strTmp = "": strTime = ""
        strSql = "Select zl_CalcInDaysNew([1],[2],[3],[4]) As ��ʼ���� From Dual"
        Set rsTmp = zldatabase.OpenSQLRecord(strSql, "��ȡסԺ����", lng�ļ�ID, lng����ID, lng��ҳID, Int(CDate(strTmpDay)))

        If rsTmp.BOF = False Then
            lngValue = rsTmp("��ʼ����").Value
        End If
        For i = 0 To 6
            strTmp = Format(CDate(strTmpDay) + i, "YYYY-MM-DD")
            If Right(strTmp, 5) = "01-01" Then
                'һ��ĵ�һ��
                strTime = strTmp
            ElseIf strTmp = Format(strBeginDate, "yyyy-MM-dd") Then
                '��Ժ��һ�죬д�����
                strTime = strTmp
            ElseIf i = 0 Then 'ÿҳ�ĵ�һ��
                strTime = Right(strTmp, 5)
            ElseIf Right(strTmp, 2) = "01" Then
                strTime = Right(strTmp, 5)
            Else
                strTime = Right(strTmp, 2)
            End If

            strTmpString0 = strTmpString0 & "'" & strTime
            strTmpString2 = strTmpString2 & "'" & lngValue + i
        Next i
        strTmpString0 = Mid(strTmpString0, 2)
        strTmpString2 = Mid(strTmpString2, 2)
        '2.��ȡ����ʱ��ʹ���
        strTime = ""
        '��ʾ��ǰ�ε��������
        strSql = "Select B.����ʱ�� ʱ��" & vbNewLine & _
            " From ���˻����ļ� A,���˻������� B,���˻�����ϸ C" & vbNewLine & _
            " Where A.Id=B.�ļ�ID And B.Id=C.��¼ID And A.Id=[1] And  nvl(A.Ӥ��,0)=[2]" & vbNewLine & _
            " And A.����ID=[3] and A.��ҳID=[4] and C.��¼����=4 and C.��ֹ�汾 is null" & vbNewLine & _
            " And B.����ʱ�� between [5] and [6] order by B.����ʱ��"
        If blnMoved Then
            strSql = Replace(strSql, "���˻����ļ�", "H���˻����ļ�")
            strSql = Replace(strSql, "���˻�������", "H���˻�������")
            strSql = Replace(strSql, "���˻�����ϸ", "H���˻�����ϸ")
        End If

        Set rsTmp = zldatabase.OpenSQLRecord(strSql, "��ȡ�������", lng�ļ�ID, intBaby, lng����ID, lng��ҳID, Int(CDate(strTmpDay) - 14), CDate(strEndDay))

        Do While Not rsTmp.EOF
            strTime = Format(rsTmp("ʱ��"), "YYYY-MM-DD")
            For i = 1 To 7
                If DateDiff("d", strTmpDay, strEndDay) + 1 >= i Then
                    intDays = DateDiff("d", strTime, strTmpDay) + (i - 1)

                    Select Case intDays
                        Case 0 '��ǰ�����ڵ�������ʼʱ��
                             'Modify 2012-03-05 �޸�һ������ж������
                            If Trim(strOpdays(i)) <> "" Then
                                strOpdays(i) = strTime & "/" & strOpdays(i)
                            Else
                                strOpdays(i) = strTime
                            End If
                        Case Else
                            If intDays >= 1 And intDays <= intOpDays Then '������ʼ����
                                If blnStopFlag Then '������ע�������ڴ�����ʱֹͣǰһ�α�ע
                                    strOpValue(i) = intDays
                                Else
                                    If Trim(strOpValue(i)) <> "" Then
                                        strOpValue(i) = intDays & "/" & strOpValue(i)
                                    Else
                                        strOpValue(i) = intDays
                                    End If
                                End If
                            End If
                    End Select
                End If
            Next i
            rsTmp.MoveNext
        Loop
        
        '��ȡ��ǰ��ʼ����-14��ǰ��������¼��Ϣ
        strSql = "select Nvl(Count(B.����ʱ��),0) ����" & _
            "   from ���˻����ļ� A, ���˻������� B,���˻�����ϸ C" & _
            "   where A.ID=B.�ļ�ID and B.ID=C.��¼ID and A.ID=[1] and nvl(A.Ӥ��,0)=[2]" & _
            "   and A.����ID=[3] and A.��ҳID=[4] and C.��¼����=4 and C.��ֹ�汾 is null" & _
            "   and B.����ʱ�� <[5] "
        Set rsTmp = zldatabase.OpenSQLRecord(strSql, "��ȡ�������", lng�ļ�ID, intBaby, lng����ID, lng��ҳID, Int(CDate(strTmpDay)))
        If blnMoved Then
            strSql = Replace(strSql, "���˻����ļ�", "H���˻����ļ�")
            strSql = Replace(strSql, "���˻�������", "H���˻�������")
            strSql = Replace(strSql, "���˻�����ϸ", "H���˻�����ϸ")
        End If
        
        lng���� = 0
        If rsTmp.BOF = False Then lng���� = Val(rsTmp("����"))
        
        For i = 1 To 7
            If DateDiff("d", Int(CDate(strTmpDay)), Int(CDate(strEndDay))) + 1 >= i Then
                If Trim(strOpdays(i)) <> "" Then
                    arrOperDay = Split(strOpdays(i), "/")
                Else
                    arrOperDay = Split("1", "/")
                End If
                lngValue = lng����
                If Trim(strOpdays(i)) <> "" And lngValue + UBound(arrOperDay) < 12 Then
                    strTmp = "": strTmp1 = ""
                    For j = UBound(arrOperDay) + 1 To 1 Step -1
                        lng���� = lngValue + j
                        strTmp1 = Switch(lng���� = 1, "��", lng���� = 2, "��", lng���� = 3, "��", lng���� = 4, "��", lng���� = 5, "��", lng���� = 6, _
                            "��", lng���� = 7, "��", lng���� = 8, "��", lng���� = 9, "��", lng���� = 10, "��", lng���� = 11, "��", lng���� = 12, "��")
                        If strTmp = "" Then
                            strTmp = strTmp1
                        Else
                            strTmp = strTmp & "/" & strTmp1
                        End If
                        If blnStopFlag Then Exit For
                    Next j
                    lng���� = lngValue + UBound(arrOperDay) + 1
                    If blnStopFlag Then '������ע�������ڴ�����ʱֹͣǰһ�α�ע
                        Select Case intOpFormat
                            Case 1 '��ʾ0
                                strOpValue(i) = 0
                            Case 2 '��ʾ��������
                                If strTmp = "��" Then
                                    strOpValue(i) = 0
                                Else
                                    strOpValue(i) = strTmp & "-0"
                                End If
                            Case Else '����ʾ
                                strOpValue(i) = ""
                        End Select
                    Else
                        Select Case intOpFormat
                            Case 1 '��ʾ0
                                If Trim(strOpValue(i)) <> "" Then
                                    strOpValue(i) = 0 & "/" & strOpValue(i)
                                Else
                                    strOpValue(i) = 0
                                End If
                            Case 2 '��ʾ��������
                                If Trim(strOpValue(i)) <> "" Then
                                    strOpValue(i) = strTmp & "/" & strOpValue(i)
                                Else
                                    strOpValue(i) = strTmp
                                End If
                            Case Else '����ʾ
                                If Trim(strOpValue(i)) <> "" Then
                                    strOpValue(i) = strOpValue(i)
                                Else
                                    strOpValue(i) = ""
                                End If
                        End Select
                    End If
                End If
            End If
        Next i
        
        strTmpString1 = Join(strOpValue, "'")
        
        stdset.Name = "����"
        stdset.Size = 9 * sngScale
        stdset.Bold = False
        Call SetFontIndirect(stdset, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        '3��ʼ���סԺ���ڣ�������������Ϣ
        Call DrawUpTable(lngDC, objDraw, strTmpString0 & "||" & strTmpString2 & "||" & strTmpString1, lngCurX, lngCurY, T_DrawClient.��������.Right, lngCurY, sngScale)
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
        
        '----------------------------------------------------------------------------------------------
         '�˴�����ɴ�ӡ���� �Ӷ��������µ���ӡ���и�
        If intRepairRows = 0 Then
            sngHTab = intRepairRows
        Else
            sngHTab = intRepairRows * T_DrawClient.ʱ���е�λ + IIf(bln���� = True, T_DrawClient.ʱ���е�λ / 2, 0)
        End If
        
        sngHTab = sngHTab + msngTwips * 30 + 10
        sngHPrint = objDraw.Height / T_TwipsPerPixel.Y - lngCurY - lngButtom - sngHTab - dblSureH
        T_DrawClient.�е�λ = (sngHPrint - 4 * T_DrawClient.�е�λ) / T_DrawClient.������
        T_DrawClient.�е�λ = Round(T_DrawClient.�е�λ - 0.05, 1)
        If T_DrawClient.�е�λ > 6 * msngTwips Then T_DrawClient.�е�λ = 6 * msngTwips
        If T_DrawClient.�е�λ < 6 * msngTwips Then T_DrawClient.�е�λ = 6 * msngTwips
        
        '�����иߺ��ڼ������µ��ɴ�ӡ�ı������
        If intRepairRows > 0 Then
            sngHPrint = T_DrawClient.������ * T_DrawClient.�е�λ + 4 * T_DrawClient.�е�λ
            sngHTab = objDraw.Height / T_TwipsPerPixel.Y - lngCurY - lngButtom - dblSureH - sngHPrint - (msngTwips * 30 + 10)
            sngHTab = sngHTab - IIf(bln���� = True, T_DrawClient.ʱ���е�λ / 2, 0)
            If Fix(sngHTab / T_DrawClient.ʱ���е�λ + 0.3) < intRepairRows Then intRepairRows = Fix(sngHTab / T_DrawClient.ʱ���е�λ + 0.3)
        End If
    
        stdset.Name = "����"
        stdset.Size = 9 * sngScale
        stdset.Bold = False
        Call SetFontIndirect(stdset, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        '4��ʼ���̶������������������̶�ֵ��Ϣ
        T_DrawClient.ƫ����Y = lngCurY
        mbln�������� = False
        
        rsTemp.Filter = 0
        rsTemp.Sort = "�������"
        rsTemp.MoveFirst
        str����˵�� = DrawCanvas(lngDC, objDraw, rsTemp, rsDrawItems, bln����ӡ������, sngScale)
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
        
        '5.��ȡ�����������ݺ����ת�ȱ����Ϣ
        '��ʼ�� ���µ��¼�������ת�ȱ����Ϣ
        
        '���е�ı��ּ���
        '   �ص��Ƿ��ص����.
        '   �ص���Ŀ��¼�ص���Ŀ
        '   �Ͽ�������:����һ��������,����δ��˵��
        '   ��ע:������ʱ��¼ԭֵ
        '   ����:������ע���²���������ֵС�ڵ�����Ŀ��Сֵ���ڵ�����Ŀ���ֵ�ǵ��������.����Ĭ��Ϊ��

        gstrFields = "���," & adDouble & ",18|��ֵ," & adLongVarChar & ",4000|��λ," & adLongVarChar & ",200|" & _
             "���," & adDouble & ",1|ʱ��," & adLongVarChar & ",20|��Ŀ���," & adDouble & ",18|" & _
             "����," & adDouble & ",1|�Ͽ�," & adDouble & ",1|�ص���Ŀ," & adLongVarChar & ",50|" & _
             "�ص�," & adDouble & ",5|X����," & adDouble & ",5|Y����," & adDouble & ",5|��ע," & adLongVarChar & ",50|" & _
             "����," & adLongVarChar & ",10|��ʾ," & adDouble & ",1"
        Call Record_Init(rsPoints, gstrFields)
    
        '������Ҫ������ı�����(����:2-�ϱ�;3-���ת;4-������;6-�±�,13-����,99-δ��˵��)
        '���ñ�ʾ��Ϣ�Ƿ����
        gstrFields = "ʱ��," & adLongVarChar & ",20|��Ŀ���," & adDouble & ",18|����," & adDouble & ",2|" & _
            "����," & adLongVarChar & ",200|��ɫ," & adLongVarChar & ",20|X����," & adDouble & ",20|" & _
            "Y����," & adDouble & ",20|�߶�," & adDouble & ",20|��ӡX����," & adDouble & ",20|" & _
            "����," & adInteger & ",1|��ʾ," & adDouble & ",1"
        Call Record_Init(rsNotes, gstrFields)
        
        Dim rs���� As New ADODB.Recordset
        Dim strFileds As String, strValues As String
        
        '��¼������Ϣ
        strFileds = "��Ŀ���," & adDouble & ",18|��ֵ," & adLongVarChar & ",4000|X����," & adDouble & ",5|ʱ��," & adLongVarChar & ",20"
        Call Record_Init(rs����, strFileds)
        
        Dim int��� As Integer
        
        '----��ȡ���в�λ��Ϣ
        strSql = "select ��Ŀ���,��λ,ȱʡ�� from ���²�λ"
        Call zldatabase.OpenRecordset(rsPart, strSql, "���²�λ")
        '----��ȡ�����������ݺ�δ��˵��
        strSql = "SELECT C.ID ���, a.����ʱ�� As ʱ��,C.��ʾ,C.��¼���� As ��ֵ,C.���²�λ,c.���Ժϸ�,D.��¼��,E.������Ŀ,D.��Ŀ���,DECODE(D.��Ŀ���,-1,1,C.��¼���) ��¼���,C.δ��˵�� " & _
                    "FROM ���˻����ļ� B,���˻������� A,���˻�����ϸ C,���¼�¼��Ŀ D,�����¼��Ŀ E " & _
                    "Where B.ID=A.�ļ�ID  " & _
                        "AND A.ID = C.��¼ID " & _
                        "AND B.ID=[1] " & _
                        "AND Nvl(B.Ӥ��,0)=[6] " & _
                        "AND B.����id=[2] " & _
                        "AND B.��ҳid=[3] " & _
                        "AND D.��Ŀ���=C.��Ŀ��� " & _
                        "AND C.��¼����=1 " & _
                        "AND E.��Ŀ���=D.��Ŀ��� " & _
                        "AND E.����ȼ�>=[7]  " & _
                        "AND A.����ʱ�� BETWEEN [4] And [5] And C.��ֹ�汾 Is Null " & _
                        "AND D.��¼��=1 AND (nvl(E.Ӧ�÷�ʽ,0)=1 OR ( -1=[10] and nvl(E.Ӧ�÷�ʽ,0)=2)) " & _
                        "AND nvl(E.���ò���,0) in (0,[8]) AND (E.���ÿ���=1 or ( E.���ÿ���=2 AND Exists (select 1 from �������ÿ��� D where D.��Ŀ���=E.��Ŀ��� and D.����ID=[9])))" & _
                    "Order By a.����ʱ��,DECODE(D.��Ŀ���,-1,1,0),DECODE(D.��Ŀ���,-1,1,C.��¼���)"
        If blnMoved Then
            strSql = Replace(strSql, "���˻����ļ�", "H���˻����ļ�")
            strSql = Replace(strSql, "���˻�������", "H���˻�������")
            strSql = Replace(strSql, "���˻�����ϸ", "H���˻�����ϸ")
        End If

        Set rsTmp = zldatabase.OpenSQLRecord(strSql, "��ȡ������Ŀ����", lng�ļ�ID, lng����ID, lng��ҳID, CDate(strTmpDay), CDate(strEndDay), _
            intBaby, lng����ȼ�, IIf(intBaby = 0, 1, 2), lngSectID, IIf(int����Ӧ�� = 2, -1, 0))
         
        strTmpString0 = ""
        strTmpString1 = ""
        strTmpString2 = ""
        With rsTmp
            Do While Not .EOF
                strTmp = ""
                blnAllow = False
                strPart = zlCommFun.Nvl(!���²�λ)
                lng��Ŀ��� = Val(zlCommFun.Nvl(!��Ŀ���))
                Select Case lng��Ŀ���
                    Case gint����
                        int��� = 1
                    Case Else
                        int��� = Val(zlCommFun.Nvl(!��¼���))
                End Select
                If strPart = "" Then
                    rsPart.Filter = "��Ŀ���=" & lng��Ŀ��� & " and ȱʡ��=1"
                    If rsPart.BOF = False Then
                        strPart = zlCommFun.Nvl(rsPart!��λ)
                    Else
                        Select Case lng��Ŀ���
                            Case gint����
                                strPart = "Ҹ��"
                            Case gint����
                                strPart = "��������"
                            Case Else
                                strPart = ""
                        End Select
                    End If
                End If
                
                SinX = GetXCoordinate(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss"), Format(strTmpDay, "YYYY-MM-DD HH:mm:ss"))
                strTime = GetXCoordinate(SinX, Format(strTmpDay, "YYYY-MM-DD HH:mm:ss"), False)
                SinX = GetXCoordinate(Format(Split(strTime, ",")(0), "YYYY-MM-DD HH:mm:ss"), Format(strTmpDay, "YYYY-MM-DD HH:mm:ss"))
                
                '��¼����������Ϣ
                If lng��Ŀ��� = gint���� Then
                    strFileds = "��Ŀ���|��ֵ|X����|ʱ��"
                    strValues = lng��Ŀ��� & "|" & zlCommFun.Nvl(!��ֵ) & "|" & SinX & "|" & Format(!ʱ��, "yyyy-MM-dd HH:mm:ss")
                    Call Record_Add(rs����, strFileds, strValues)
                End If
                
                If (Not IsNull(!δ��˵��)) And zlCommFun.Nvl(!��ֵ) <> "����" Then
                    rsNotes.Filter = "��Ŀ���=" & Val(zlCommFun.Nvl(!��Ŀ���)) & " AND X����=" & SinX
                    blnAdd = (rsNotes.RecordCount = 0)
                    '������Ҫ������ı�����(����:2-�ϱ�;3-���ת;4-������;6-�±�,99-δ��˵��)
                    gstrFields = "ʱ��|��Ŀ���|����|����|��ɫ|X����|Y����|�߶�|��ӡX����|����|��ʾ"  '���תȱʡ�Ǻ�ɫ,���±꼰δ��˵��ȱʡ����ɫ
                    gstrValues = Format(!ʱ��, "yyyy-MM-dd HH:mm:ss") & "|" & !��Ŀ��� & "|99|" & _
                        !δ��˵�� & "|" & RGB_BLUE & "|" & SinX & "|0|0|0|0|" & zlCommFun.Nvl(!��ʾ)
                   
                    If blnAdd Then
                        '��ȡ�ӽ��м�ʱ����ֵ��Ϊ����ֵ
                         Call Record_Add(rsNotes, gstrFields, gstrValues)
                    Else
                        If (zlCommFun.Nvl(rsNotes!��ʾ, 0) = 1 And zlCommFun.Nvl(!��ʾ, 0) = 1) Or (zlCommFun.Nvl(rsNotes!��ʾ, 0) <> 1 And zlCommFun.Nvl(!��ʾ, 0) <> 1) Then
                             blnAllow = GetCanvasCenter(CDate(Format(rsNotes!ʱ��, "YYYY-MM-DD HH:mm:ss")), CDate(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss")), CDate(Format(strTmpDay, "YYYY-MM-DD HH:mm:ss")), SinX)
                        ElseIf zlCommFun.Nvl(!��ʾ, 0) = 1 Then
                            blnAllow = True
                        End If
    
                        If blnAllow = True Then
                            If Val(rsNotes!��ʾ) = 2 Then
                                arrValues = Split(gstrValues, "|")
                                arrValues(UBound(arrValues)) = 2
                                gstrValues = Join(arrValues, "|")
                            End If
                            Call Record_Update(rsNotes, gstrFields, gstrValues, "ʱ��|" & Format(rsNotes!ʱ��, "yyyy-MM-dd HH:mm:ss"))
                        Else
                            If Val(zlCommFun.Nvl(!��ʾ, 0)) = 2 Then
                                gstrFields = "��ʾ"
                                gstrValues = "2"
                                Call Record_Update(rsNotes, gstrFields, gstrValues, "ʱ��|" & Format(rsNotes!ʱ��, "yyyy-MM-dd HH:mm:ss"))
                            End If
                        End If
                    End If
                Else
                    blnAdd = False
                    
                    rsPoints.Filter = "��Ŀ���=" & lng��Ŀ��� & " AND X����=" & SinX & " And ���=" & int���
                    
                    blnAdd = (rsPoints.RecordCount = 0)
                    
                    dbl��ֵ = Val(zlCommFun.Nvl(!��ֵ))
                    
                    For i = 0 To UBound(strEditors)
                        If Val(Split(strEditors(i), "||")(0)) = lng��Ŀ��� Then
                            Exit For
                        End If
                        
                    Next i
                    If i <= UBound(strEditors) Then
'                        If InStr(1, Split(strEditors(i), "||")(3), ";") <> 0 Then
'                            dblMinValue = Val(Split(Split(strEditors(i), "||")(3), ";")(0))
'                            dblMaxValue = Val(Split(Split(strEditors(i), "||")(3), ";")(1))
'                            If dblMaxValue = 0 Then dblMaxValue = Split(strEditors(i), "||")(6)
'                        Else
'                            dblMaxValue = Val(Split(strEditors(i), "||")(6))
'                            dblMinValue = Val(Split(strEditors(i), "||")(7))
'                        End If
                        dblMaxValue = Val(Split(strEditors(i), "||")(6))
                        dblMinValue = Val(Split(strEditors(i), "||")(7))
                    End If
                    
                    '�ٽ�ֵ���ȿ�,���������ֵ����Сֵ֮��
                    If Split(strEditors(i), "||")(8) <> "" And Val(Split(strEditors(i), "||")(8)) <= Val(Split(strEditors(i), "||")(6)) _
                        And Val(Split(strEditors(i), "||")(8)) >= Val(Split(strEditors(i), "||")(7)) Then dblMaxValue = Val(Split(strEditors(i), "||")(8))
                    
                    '��ָ�����ţ���Ŀ���ݲ������ֵ����Сֵ����Ŀ���������ʾ
                    If dbl��ֵ <= dblMinValue Then
                        dbl��ֵ = dblMinValue
                        'strTmp = "��"
                    End If
                    
                    
                    If dbl��ֵ >= dblMaxValue Then
                        dbl��ֵ = dblMaxValue
                        'strTmp = "��"
                    End If
                    
                     '���²���������ʾ��35�̶�
                    If Trim(Nvl(!��ֵ)) = "����" And lng��Ŀ��� = gint���� Then dbl��ֵ = 35
                    
                    sinY = Val(GetYCoordinate(objDraw, rsDrawItems, !��Ŀ���, dbl��ֵ, lngDC, True))
                    
                    gstrFields = "���|��ֵ|��λ|���|ʱ��|��Ŀ���|����|�Ͽ�|�ص���Ŀ|�ص�|X����|Y����|��ע|����|��ʾ"
                    gstrValues = Val(zlCommFun.Nvl(!���)) & "|" & !��ֵ & "|" & strPart & "|" & int��� & "|" & _
                                 Format(!ʱ��, "yyyy-MM-dd HH:mm:ss") & "|" & lng��Ŀ��� & "|" & Val(zlCommFun.Nvl(!���Ժϸ�, 0)) & "|" & IIf(zlCommFun.Nvl(!��ֵ, 0) = "����", 1, 0) & "|��|0|" & _
                                 SinX & "|" & sinY & "||" & strTmp & "|" & zlCommFun.Nvl(!��ʾ, 0)
                    If blnAdd Then '���
                        Call Record_Add(rsPoints, gstrFields, gstrValues)
                    Else
                        If (zlCommFun.Nvl(rsPoints!��ʾ, 0) = 1 And zlCommFun.Nvl(!��ʾ, 0) = 1) Or (zlCommFun.Nvl(rsPoints!��ʾ, 0) <> 1 And zlCommFun.Nvl(!��ʾ, 0) <> 1) Then
                            blnAllow = GetCanvasCenter(CDate(Format(rsPoints!ʱ��, "YYYY-MM-DD HH:mm:ss")), CDate(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss")), CDate(Format(strTmpDay, "YYYY-MM-DD HH:mm:ss")), SinX)
                        ElseIf zlCommFun.Nvl(!��ʾ, 0) = 1 Then
                            blnAllow = True
                        End If
                        
                       '��ȡ�ӽ��м�ʱ����ֵ��Ϊ����ֵ
                        If blnAllow = True Then
                            If Val(rsPoints!��ʾ) = 2 Then
                                arrValues = Split(gstrValues, "|")
                                arrValues(UBound(arrValues)) = 2
                                gstrValues = Join(arrValues, "|")
                            End If
                            Call Record_Update(rsPoints, gstrFields, gstrValues, "���|" & rsPoints!���)
                        Else
                            If Val(zlCommFun.Nvl(!��ʾ, 0)) = 2 Then
                                gstrFields = "��ʾ"
                                gstrValues = "2"
                                Call Record_Update(rsPoints, gstrFields, gstrValues, "���|" & rsPoints!���)
                            End If
                        End If
                    End If
                End If
            .MoveNext
            Loop
        End With
                
        '�����Ѿ��õ���������Ŀ��������Ϣ���������������º���������������
        rsPoints.Filter = ""
        arrTmpValue = Array()
        If int����Ӧ�� = 2 Then
            rsPoints.Filter = "��Ŀ���=" & gint����
            With rsPoints
                Do While Not .EOF
                    ReDim Preserve arrTmpValue(UBound(arrTmpValue) + 1)
                    arrTmpValue(UBound(arrTmpValue)) = !��� & ";" & !��Ŀ��� & ";" & !X���� & ";" & Format(!ʱ��, "yyyy-MM-DD HH:mm:ss")
                .MoveNext
                Loop
            End With
        End If
        
        '������Ϊ��������ʱ����������Ƿ�����Ϊ����
        If int���� <> -1 Then
            For i = 0 To UBound(arrTmpValue)
                '��������Ƿ����������Ӧ
                rs����.Filter = "��Ŀ���=" & gint���� & " And X����=" & Val(Split(CStr(arrTmpValue(i)), ";")(2)) & " And ʱ��='" & Format(Split(CStr(arrTmpValue(i)), ";")(3), "yyyy-MM-DD HH:mm:ss") & "'"
                
                rsPoints.Filter = "��Ŀ���=" & gint���� & " and X����=" & Val(Split(CStr(arrTmpValue(i)), ";")(2)) & " And ʱ��='" & Format(Split(CStr(arrTmpValue(i)), ";")(3), "yyyy-MM-DD HH:mm:ss") & "'"
                If rsPoints.RecordCount = 0 Then
                    If rs����.RecordCount = 0 Then
                        rsPoints.Filter = ""
                        gstrFields = "��Ŀ���": gstrValues = gint����
                        Call Record_Update(rsPoints, gstrFields, gstrValues, "���|" & Val(Split(CStr(arrTmpValue(i)), ";")(0)))
                    Else
                        rsPoints.Filter = "��Ŀ���=" & gint���� & " And X����=" & Val(Split(CStr(arrTmpValue(i)), ";")(2)) & " And ʱ��='" & Format(Split(CStr(arrTmpValue(i)), ";")(3), "yyyy-MM-DD HH:mm:ss") & "'"
                        rsPoints.Delete
                    End If
                End If
            Next i
        End If
        
        If int����Ӧ�� = 2 Then
            Set rs���� = New ADODB.Recordset
            strFileds = "���," & adDouble & ",18|��ֵ," & adLongVarChar & ",4000|��λ," & adLongVarChar & ",200|" & _
                        "���," & adDouble & ",1|ʱ��," & adLongVarChar & ",20|��Ŀ���," & adDouble & ",18|" & _
                        "����," & adDouble & ",1|�Ͽ�," & adDouble & ",1|�ص���Ŀ," & adLongVarChar & ",50|" & _
                        "�ص�," & adDouble & ",5|X����," & adDouble & ",5|Y����," & adDouble & ",5|��ע," & adLongVarChar & ",50|" & _
                        "����," & adLongVarChar & ",10|��ʾ," & adDouble & ",1"
            Call Record_Init(rs����, strFileds)
            
            rsPoints.Filter = "��Ŀ���=" & gint����
            With rsPoints
                Do While Not .EOF
                    rs����.AddNew
                    For i = 0 To .Fields.Count - 1
                        rs����.Fields(.Fields(i).Name).Value = .Fields(i).Value
                    Next i
                    rs����.Update
                .MoveNext
                Loop
            End With
            
            rsPoints.Filter = "��Ŀ���=" & gint����
            Do While Not rsPoints.EOF
                rsPoints.Delete
                rsPoints.MoveNext
            Loop
            
            rs����.Filter = ""
            rs����.Sort = "ʱ��"
            With rs����
                Do While Not .EOF
                    blnAdd = False
                    blnAllow = False
                    
                    SinX = Val(zlCommFun.Nvl(!X����))
                    sinY = Val(zlCommFun.Nvl(!Y����))
                    rsPoints.Filter = "��Ŀ���=" & Val(zlCommFun.Nvl(!��Ŀ���, 0)) & " AND X����=" & SinX
                    blnAdd = IIf(rsPoints.RecordCount = 0, True, False)
                    
                    strFileds = "���|��ֵ|��λ|���|ʱ��|��Ŀ���|����|�Ͽ�|�ص���Ŀ|�ص�|X����|Y����|��ע|����|��ʾ"
                    strValues = Val(zlCommFun.Nvl(!���)) & "|" & !��ֵ & "|" & zlCommFun.Nvl(!��λ) & "|" & Val(zlCommFun.Nvl(!���, 0)) & "|" & _
                                 Format(!ʱ��, "yyyy-MM-dd HH:mm:ss") & "|" & Val(zlCommFun.Nvl(!��Ŀ���)) & "|0|" & Val(zlCommFun.Nvl(!�Ͽ�)) & "|��|0|" & _
                                 SinX & "|" & sinY & "||" & zlCommFun.Nvl(!����) & "|" & Val(zlCommFun.Nvl(!��ʾ, 0))
                    
                    If blnAdd Then '���
                        Call Record_Add(rsPoints, strFileds, strValues)
                    Else
                        If (zlCommFun.Nvl(rsPoints!��ʾ, 0) = 1 And zlCommFun.Nvl(!��ʾ, 0) = 1) Or (zlCommFun.Nvl(rsPoints!��ʾ, 0) <> 1 And zlCommFun.Nvl(!��ʾ, 0) <> 1) Then
                            blnAllow = GetCanvasCenter(CDate(Format(rsPoints!ʱ��, "YYYY-MM-DD HH:mm:ss")), CDate(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss")), CDate(Format(strTmpDay, "YYYY-MM-DD HH:mm:ss")), SinX)
                        ElseIf zlCommFun.Nvl(!��ʾ, 0) = 1 Then
                            blnAllow = True
                        End If
                        
                        '��ȡ�ӽ��м�ʱ����ֵ��Ϊ����ֵ
                        If blnAllow = True Then
                            If Val(rsPoints!��ʾ) = 2 Then
                                arrValues = Split(strValues, "|")
                                arrValues(UBound(arrValues)) = 2
                                strValues = Join(arrValues, "|")
                            End If
                            Call Record_Update(rsPoints, strFileds, strValues, "���|" & rsPoints!���)
                        Else
                            If Val(zlCommFun.Nvl(!��ʾ, 0)) = 2 Then
                                strFileds = "��ʾ"
                                strValues = "2"
                                Call Record_Update(rsPoints, strFileds, strValues, "���|" & rsPoints!���)
                            End If
                        End If
                    End If
                .MoveNext
                Loop
            End With
        End If
        
        '��������������
        arrTmpValue = Array()
        rsPoints.Filter = "��Ŀ���=1 and ���=0"
        With rsPoints
            Do While Not .EOF
                ReDim Preserve arrTmpValue(UBound(arrTmpValue) + 1)
                arrTmpValue(UBound(arrTmpValue)) = !��� & ";" & !��Ŀ��� & ";" & !��ֵ & ";" & !X���� & ";" & !Y���� & ";" & Format(!ʱ��, "yyyy-MM-dd HH:mm:ss")
            .MoveNext
            Loop
        End With
        
        rsPoints.Filter = "��Ŀ���=1"
        If rsPoints.RecordCount > 0 Then rsPoints.MoveFirst
        For i = 0 To UBound(arrTmpValue)
            rsPoints.Filter = "��Ŀ���=1 and ���=1 and X����=" & Val(Split(CStr(arrTmpValue(i)), ";")(3)) & " And ʱ��='" & Format(Split(CStr(arrTmpValue(i)), ";")(5), "yyyy-MM-dd HH:mm:ss") & "'"
            If rsPoints.RecordCount <> 0 Then
                gstrFields = "��ע": gstrValues = Val(Split(CStr(arrTmpValue(i)), ";")(2)) & "," & Val(Split(CStr(arrTmpValue(i)), ";")(3)) & ";" & Val(Split(CStr(arrTmpValue(i)), ";")(4))
                Call Record_Update(rsPoints, gstrFields, gstrValues, "���|" & zlCommFun.Nvl(rsPoints!���))
            End If
        Next i
        
        arrTmpValue = Array()
        rsPoints.Filter = "��Ŀ���=1 and ���=1"
        With rsPoints
            Do While Not .EOF
                ReDim Preserve arrTmpValue(UBound(arrTmpValue) + 1)
                arrTmpValue(UBound(arrTmpValue)) = !��� & ";" & !��Ŀ��� & ";" & !��ֵ & ";" & !X���� & ";" & !Y���� & ";" & Format(!ʱ��, "yyyy-MM-dd HH:mm:ss")
            .MoveNext
            Loop
        End With
        
        rsPoints.Filter = "��Ŀ���=1"
        If rsPoints.RecordCount > 0 Then rsPoints.MoveFirst
        For i = 0 To UBound(arrTmpValue)
            rsPoints.Filter = "��Ŀ���=1 and ���=0 and X����=" & Val(Split(CStr(arrTmpValue(i)), ";")(3)) & " And ʱ��='" & Format(Split(CStr(arrTmpValue(i)), ";")(5), "yyyy-MM-dd HH:mm:ss") & "'"
            If rsPoints.RecordCount = 0 Then
                rsPoints.Filter = "��Ŀ���=1 and ���=1 and X����=" & Val(Split(CStr(arrTmpValue(i)), ";")(3)) & " And ʱ��='" & Format(Split(CStr(arrTmpValue(i)), ";")(5), "yyyy-MM-dd HH:mm:ss") & "'"
                rsPoints.Delete
            End If
        Next i
    
        'ɾ����ʾΪ2������
        rsPoints.Filter = ""
        rsPoints.Filter = "��ʾ=2"
        Do While Not rsPoints.EOF
            rsPoints.Delete
        rsPoints.MoveNext
        Loop
        
        rsNotes.Filter = ""
        rsNotes.Filter = "��ʾ=2"
        Do While Not rsNotes.EOF
            rsNotes.Delete
        rsNotes.MoveNext
        Loop
        
        '����δ��˵�����������ݸ���ʾ��һ��
        rsNotes.Filter = ""
        rsPoints.Filter = ""
        
        arrTmpValue = Array()
        arrTmpNote = Array()
        rsNotes.Sort = "��Ŀ���,X����"
        With rsNotes
            Do While Not .EOF
                SinX = Val(!X����)
                blnAllow = False
                rsPoints.Filter = "��Ŀ���=" & Val(!��Ŀ���) & " And X����=" & SinX
                If rsPoints.RecordCount > 0 Then
                    If (zlCommFun.Nvl(rsPoints!��ʾ, 0) = 1 And zlCommFun.Nvl(!��ʾ, 0) = 1) Or (zlCommFun.Nvl(rsPoints!��ʾ, 0) <> 1 And zlCommFun.Nvl(!��ʾ, 0) <> 1) Then
                        blnAllow = GetCanvasCenter(CDate(Format(rsPoints!ʱ��, "YYYY-MM-DD HH:mm:ss")), CDate(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss")), CDate(Format(strTmpDay, "YYYY-MM-DD HH:mm:ss")), SinX)
                    ElseIf zlCommFun.Nvl(!��ʾ, 0) = 1 Then
                        blnAllow = True
                    End If
                    If blnAllow = True Then
                        ReDim Preserve arrTmpValue(UBound(arrTmpValue) + 1)
                        arrTmpValue(UBound(arrTmpValue)) = !��Ŀ��� & ";" & SinX
                    Else
                        ReDim Preserve arrTmpNote(UBound(arrTmpNote) + 1)
                        arrTmpNote(UBound(arrTmpNote)) = !��Ŀ��� & ";" & SinX
                    End If
                End If
            .MoveNext
            Loop
        End With
        
        For i = 0 To UBound(arrTmpValue)
            rsPoints.Filter = "��Ŀ���=" & Val(Split(CStr(arrTmpValue(i)), ";")(0)) & " And X����=" & Val(Split(CStr(arrTmpValue(i)), ";")(1))
            Do While Not rsPoints.EOF
                rsPoints.Delete
            rsPoints.MoveNext
            Loop
        Next i
        
        For i = 0 To UBound(arrTmpNote)
            rsNotes.Filter = "��Ŀ���=" & Val(Split(CStr(arrTmpNote(i)), ";")(0)) & " And X����=" & Val(Split(CStr(arrTmpNote(i)), ";")(1))
            Do While Not rsNotes.EOF
                rsNotes.Delete
            rsNotes.MoveNext
            Loop
        Next i
    
'        '�������²��� ����Ϊ������Ҫ��35��������������²�������
        rsPoints.Filter = "��Ŀ���=" & gint���� & " and ��ֵ='����' and ���<>1"
        rsPoints.Sort = "ʱ��"
        With rsPoints
            Do While Not .EOF
                strTmpString0 = strTmpString0 & ";" & Format(!ʱ��, "yyyy-MM-dd HH:mm:ss") & "|" & Val(zlCommFun.Nvl(!��Ŀ���)) & "|99|" & _
                      "����|" & RGB_BLUE & "|" & !X���� & "|0|0|0|0"
                strTmpString2 = strTmpString2 & ";" & !X����
            .MoveNext
            Loop
        End With
        
        '--------���¶Ͽ����
        '����֮����δ��˵���Ͽ���ʱ�����һ��Ͽ�,���²����Ͽ�
        rsPoints.Filter = ""
        
        gstrFields = "�Ͽ�"
        gstrValues = "1"
        rsNotes.Filter = ""
        
        If rsNotes.RecordCount > 0 Then rsNotes.MoveFirst
        With rsNotes
            Do While Not .EOF
                If int����Ӧ�� = 2 And !��Ŀ��� = -1 Then
                    rsPoints.Filter = "��Ŀ���=" & gint���� & " And X����<=" & !X����
                Else
                    If !��Ŀ��� = 1 Then
                        rsPoints.Filter = "��Ŀ���=" & !��Ŀ��� & " And  ���<>1 And X����<" & !X����
                    Else
                        rsPoints.Filter = "��Ŀ���=" & !��Ŀ��� & " And X����<" & !X����
                    End If
                End If
                rsPoints.Sort = "ʱ��"
                If rsPoints.RecordCount <> 0 Then
                    rsPoints.MoveLast
                    Call Record_Update(rsPoints, gstrFields, gstrValues, "���|" & rsPoints!���)
                End If
      
            .MoveNext
            Loop
        End With
        'ʱ�䳬��һ��
        strTime = ""
        strTmp = ""
        rsPoints.Filter = ""
        
        rsPoints.Sort = "��Ŀ���,ʱ��,���"
        With rsPoints
            Do While Not .EOF
                If Not IsNull(!���) Then
                    If Not (Val(!��Ŀ���) = 1 And Val(!���) = 1) Then
                        If lng��Ŀ��� <> 0 Then
                            If lng��Ŀ��� <> !��Ŀ��� Then strTime = ""
                        End If
                        lng��Ŀ��� = !��Ŀ���
                        If strTime <> "" Then
                            If DateDiff("D", CDate(strTime), CDate(Format(!ʱ��, "YYYY-MM-DD"))) > 1 Then
                                strTmp = strTmp & "," & lngValue
                            End If
                        End If
                        strTime = Format(rsPoints!ʱ��, "YYYY-MM-DD")
                        lngValue = Val(rsPoints!���)
                    End If
                End If
                .MoveNext
            Loop
        End With
        
        strTmp = Mid(strTmp, 2)
        For i = 0 To UBound(Split(strTmp, ","))
            Call Record_Update(rsPoints, gstrFields, gstrValues, "���|" & Split(strTmp, ",")(i))
        Next i
        
        '�������²�����.��ǰһ����ĶϿ���־����Ϊ1
        rsPoints.Filter = ""
        rsPoints.Filter = "��Ŀ���=" & gint���� & " and ���<>1"
        rsPoints.Sort = "ʱ��,���"
        With rsPoints
            Do While Not .EOF
                If !��ֵ = "����" And .AbsolutePosition <> 1 Then
                    .MovePrevious '������һ�жϿ����
                    If Val(!�Ͽ�) <> 1 Then
                        lngValue = !���
                        Call Record_Update(rsPoints, gstrFields, gstrValues, "���|" & lngValue)
                    End If
                    .MoveNext
                End If
            .MoveNext
            Loop
        End With
    
        '��������δ��˵����ͬһX��������ͬ��˵��ֵ���һ��
        rsNotes.Filter = ""
        rsNotes.Sort = "X����"
        With rsNotes
            Do While Not .EOF
                If lngValue = !X���� Then
                    If InStr(1, "," & strTmp & ",", "," & zlCommFun.Nvl(!����) & ",") <> 0 Then
                       rsNotes.Delete
                    Else
                        strTmp = strTmp & "," & zlCommFun.Nvl(!����)
                    End If
                Else
                    lngValue = !X����
                    strTmp = zlCommFun.Nvl(!����)
                End If
            .MoveNext
            Loop
        End With
        
        '--��ȡ���Ժ,�����ȱ��˵��
        Dim bytShow As Byte
        Dim str���� As String
        Dim lng�к� As Long, lngColor As Long
        
        '��ȡ���������±���Ϣ
        '-----------------------------------------------------------------------
        gstrFields = "ʱ��|��Ŀ���|����|����|��ɫ|X����|Y����|�߶�|��ӡX����|����"  '���תȱʡ�Ǻ�ɫ,���±꼰δ��˵��ȱʡ����ɫ
        strSql = "" & _
                 " Select B.����ʱ�� AS ʱ��,C.��¼����,C.��Ŀ���,C.��¼����,C.��Ŀ����,C.δ��˵��" & _
                 " FROM ���˻����ļ� A, ���˻������� B, ���˻�����ϸ C" & _
                 " Where A.ID=B.�ļ�ID and  B.ID = C.��¼ID AND A.ID=[1]   AND Nvl(A.Ӥ��, 0)=[6] AND A.����id=[2] AND A.��ҳid=[3] And c.��ֹ�汾 Is Null" & _
                 " AND mod(c.��¼����,10) <> 1  AND B.����ʱ�� BETWEEN [4]  And [5]"
        If blnMoved Then
            strSql = Replace(strSql, "���˻����ļ�", "H���˻����ļ�")
            strSql = Replace(strSql, "���˻�������", "H���˻�������")
            strSql = Replace(strSql, "���˻�����ϸ", "H���˻�����ϸ")
        End If

        Set rsTmp = zldatabase.OpenSQLRecord(strSql, "��ȡ���������±����Ϣ", lng�ļ�ID, lng����ID, lng��ҳID, Int(CDate(strTmpDay)), CDate(strEndDay), intBaby, lng����ȼ�)
        With rsTmp
            Do While Not .EOF
                bytShow = 1
                str���� = Trim(zlCommFun.Nvl(!��¼����))
               
                lng�к� = IIf(!��¼���� = 2, 10, IIf(!��¼���� = 6, 11, 4))
                
                '����������ʾ��Ҫ���⴦��
                If !��¼���� = 4 Then
                    str���� = Trim(zlCommFun.Nvl(!��Ŀ����))
                    
                    If str���� = "����" Then
                        bytShow = T_BodyFlag.����
                    Else
                        bytShow = T_BodyFlag.����
                    End If
                    
                    If bytShow = 2 Then
                        str���� = str���� & gstrCaveSplit & ConvertTimeToChinese(Format(!ʱ��, "HH:mm"))
                    Else
                        str���� = !��Ŀ����
                    End If
                    lngColor = RGB_RED
                Else
                    lngColor = IIf(Not IsNumeric(Nvl(!δ��˵��)), RGB_BLUE, Val(Nvl(!δ��˵��)))
                End If
                
                If bytShow > 0 Then
                    SinX = Val(GetXCoordinate(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss"), strTmpDay))
                    
                    rsNotes.Filter = "X����=" & SinX & " and ��Ŀ���=" & lng�к� & " and ����=" & !��¼����
                    If rsNotes.BOF Then
                        gstrValues = Format(!ʱ��, "yyyy-MM-dd HH:mm:ss") & "|" & lng�к� & "|" & !��¼���� & "|" & _
                            str���� & "|" & lngColor & "|" & SinX & "|0|0|0|0"
                        Call Record_Add(rsNotes, gstrFields, gstrValues)
                    Else
                        rsNotes!ʱ�� = Format(!ʱ��, "yyyy-MM-dd HH:mm:ss")
                        rsNotes!���� = str����
                        rsNotes.Update
                    End If
                End If
                rsNotes.Filter = ""
                .MoveNext
            Loop
        End With
        
        '��ȡ���ת����Ϣ
        '-----------------------------------------------------------------------
        '������Ҫ������ı�����(����:2-�ϱ�;3-���ת;4-������;6-�±�,99-δ��˵��)
        '1-��Ժ��2-��ƣ�3-ת�ƣ�4-����
        gstrFields = "ʱ��|��Ŀ���|����|����|��ɫ|X����|Y����|�߶�|��ӡX����|����"  '���תȱʡ�Ǻ�ɫ,���±꼰δ��˵��ȱʡ����ɫ
        Set rsTmp = GetDataFromHis(lng����ID, lng��ҳID, intBaby, CDate(strTmpDay), CDate(strEndDay), 2)
        With rsTmp
            Do While Not .EOF
                If Trim(zlCommFun.Nvl(!����)) <> "" Then
                    bytShow = 0
                    lng�к� = Val(!�к�)
                    str���� = zlCommFun.Nvl(!����)
                    Select Case Val(!�к�)
                    Case 5
                        bytShow = T_BodyFlag.��Ժ
                    Case 6, 3 '6ת�룬3ת��
                        bytShow = T_BodyFlag.ת��
                    Case 7
                        bytShow = T_BodyFlag.����
                    Case 8
                        bytShow = T_BodyFlag.��Ժ
                        If intBaby > 0 Then
                            bytShow = IIf(blnӤ�����µ���ʾ��Ժ, bytShow, 0)
                        End If
                    Case 9
                        bytShow = T_BodyFlag.���
                    End Select
                    
                    If bytShow > 0 Then
                        If lng�к� = 9 And bln�����ʾ��Ժ = True Then str���� = "��Ժ"
                        'Ŀǰ3��4 �����ת�� 3-��ʾ˵���Ϳ��� 4 ��ʾ˵�������ң�ʱ��
                        If bytShow = 2 Then
                            str���� = str���� & gstrCaveSplit & ConvertTimeToChinese(Format(!ʱ��, "HH:mm"))
                        ElseIf bytShow = 3 Then
                            str���� = str���� & gstrCaveSplit & zlCommFun.Nvl(!����)
                        ElseIf bytShow = 4 Then
                            str���� = str���� & gstrCaveSplit & zlCommFun.Nvl(!����) & gstrCaveSplit & ConvertTimeToChinese(Format(!ʱ��, "HH:mm"))
                        ElseIf bytShow = 1 Then
                            str���� = str����
                        End If
                        
                        SinX = Val(GetXCoordinate(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss"), strTmpDay))
                        rsNotes.Filter = "X����=" & SinX & " and ��Ŀ���=" & lng�к� & " and ����=3"
                        
                        If rsNotes.BOF Then
                            gstrValues = Format(!ʱ��, "yyyy-MM-dd HH:mm:ss") & "|" & lng�к� & "|3|" & _
                                str���� & "|" & RGB_RED & "|" & SinX & "|0|0|0|0"
                            Call Record_Add(rsNotes, gstrFields, gstrValues)
                        Else
                            rsNotes!ʱ�� = Format(!ʱ��, "yyyy-MM-dd HH:mm:ss")
                            rsNotes!���� = str����
                            rsNotes.Update
                        End If
                    End If
                    rsNotes.Filter = ""
                End If
                .MoveNext
            Loop
        End With
        
        '��ȡӤ��������Ϣ
        If intBaby > 0 Then
            gstrFields = "ʱ��|��Ŀ���|����|����|��ɫ|X����|Y����|�߶�|��ӡX����|����"  '���תȱʡ�Ǻ�ɫ,���±꼰δ��˵��ȱʡ����ɫ
            Set rsTmp = GetDataFromHis(lng����ID, lng��ҳID, intBaby, CDate(strTmpDay), CDate(strEndDay), 3)
            With rsTmp
                Do While Not .EOF
                    bytShow = 0
                    If Trim(zlCommFun.Nvl(!����)) <> "" Then
                        lng�к� = 12
                        bytShow = T_BodyFlag.����
                        If bytShow > 0 Then
                            Select Case bytShow
                                Case 1
                                    str���� = zlCommFun.Nvl(!����)
                                Case 2
                                    str���� = zlCommFun.Nvl(!����) & gstrCaveSplit & ConvertTimeToChinese(Format(!ʱ��, "HH:mm"))
                            End Select
                            
                            SinX = Val(GetXCoordinate(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss"), strTmpDay))
                            rsNotes.Filter = "X����=" & SinX & " and ��Ŀ���=" & lng�к� & " and ����=13"
                            
                            If rsNotes.BOF Then
                                gstrValues = Format(!ʱ��, "yyyy-MM-dd HH:mm:ss") & "|" & lng�к� & "|13|" & _
                                    str���� & "|" & RGB_RED & "|" & SinX & "|0|0|0|0"
                                Call Record_Add(rsNotes, gstrFields, gstrValues)
                            Else
                                rsNotes!ʱ�� = Format(!ʱ��, "yyyy-MM-dd HH:mm:ss")
                                rsNotes!���� = str����
                                rsNotes.Update
                            End If
                        End If
                    End If
                    rsNotes.Filter = ""
                .MoveNext
                Loop
            End With
        End If
        '51512,������,2012-07-11,δ��˵����ʾλ�� 0-��ʾ������,1-��ʾ������,2-����ʾ(����)
        '��ҽ��ԺҪ��δ��˵������ʾ������ע��δ�ǵ����ߵ��������߲�����
        strTmp = ""
        Dim arrString() As String
        '�������²��� ���²���ʼ����ʾ�� 35 �����棬ֻ��δ��˵����ʾ�������������Ž���������δ��˵���У���������������±���
        If Left(strTmpString0, 1) = ";" Then
            gstrFields = "ʱ��|��Ŀ���|����|����|��ɫ|X����|Y����|�߶�|��ӡX����|����"
            If mlng���²�����ʾ��ʽ = 0 Or mlng���²�����ʾ��ʽ = 2 Then
                arrString = Split(strTmpString0, "|")
                arrString(3) = "�� "
                strTmpString0 = Join(arrString, "|")
            End If
            strTmpString0 = Mid(strTmpString0, "2")
            strTmpString2 = Mid(strTmpString2, 2)
            For i = 0 To UBound(Split(strTmpString0, ";"))
                strTmp = Split(strTmpString0, ";")(i)
                rsNotes.Filter = "����=" & IIf(bytδ����ʾλ�� = 1, 99, 6) & " and X����=" & Val(Split(strTmpString2, ";")(i))
                rsNotes.Sort = "��Ŀ���"
                If rsNotes.RecordCount > 0 Then
                    rsNotes!���� = IIf(mlng���²�����ʾ��ʽ = 0 Or mlng���²�����ʾ��ʽ = 2, "�� ", "����") & ";" & zlCommFun.Nvl(rsNotes!����)
                    rsNotes.Update
                Else
                    If mlng���²�����ʾ��ʽ = 0 Or mlng���²�����ʾ��ʽ = 2 Then strTmp = Replace(strTmp, "����", "�� ")
                    Call Record_Add(rsNotes, gstrFields, strTmp)
                    rsNotes!���� = IIf(bytδ����ʾλ�� = 1, 99, 6)
                    rsNotes.Update
                End If
            Next i
        End If
        
        '���δ��˵������ʾ����ȡ����¼��rsNote������Ϊ99�ļ�¼
        If bytδ����ʾλ�� = 2 Then
            rsNotes.Filter = "����=99"
            Do While Not rsNotes.EOF
                rsNotes.Delete
                rsNotes.MoveNext
            Loop
            rsNotes.Filter = ""
        End If
        rsPoints.Filter = 0
        '6 ������֯�ظ��ĵ�
        Call GetConverPoint(rsPoints)
        stdset.Name = "����"
        stdset.Size = 9 * sngScale
        Call SetFontIndirect(stdset, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        '7 ��ʼ�����Ϣ������
        
        strTmp = ShowPoints(lngDC, objDraw, rsPoints, strEditors, int����Ӧ��, sngScale)
        
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
        '8.����������������
        rsPoints.Filter = ""
        If strTmp <> "" And bln��ӡ������� = True Then Call CreatePoly(rsPoints, objDraw, lngDC, strTmpDay, strTmp)
        '9���˵����Ϣ
        '  �ȴ���δ��˵�����±�˵��
        Dim strText As String
        Dim SinY35 As Single, SinY42 As Single
        Dim intAscCharNum As Integer
        
        strTime = ""
        strTmp = ""
        blnAllow = False
        SinX = 0: sinY = 0
        SinY35 = GetYCoordinate(objDraw, rsDrawItems, gint����, 35, lngDC)
        SinY42 = GetYCoordinate(objDraw, rsDrawItems, gint����, 42, lngDC)
        
        rsNotes.Filter = ""
        rsNotes.Sort = "X����,��Ŀ���"
        With rsNotes
            Do While Not .EOF
                strTmp = ""
                For i = 0 To UBound(Split(!����, ";"))
                    If Not (Split(!����, ";")(i) = "����" And bytδ����ʾλ�� = 0 And Nvl(!����) = 99) And Split(!����, ";")(i) <> "" Then
                        If InStr(1, strTmp, Split(!����, ";")(i)) = 0 Then
                            strTmp = strTmp & ";" & Split(!����, ";")(i)
                        End If
                    End If
                Next i
                If Left(strTmp, "1") = ";" Then strTmp = Mid(strTmp, 2)
                If strTmp <> "" Then
                    strTime = Replace(strTmp, ";", " ")
                    If zlCommFun.Nvl(!����) = 99 Then
                        If bytδ����ʾλ�� = 1 Then '��ʾ�����µ�����
                            If blnAllow = True Then
                                If Val(zlCommFun.Nvl(!X����)) <> SinX Then
                                    sinY = SinY35
                                Else
                                    strTime = " " & strTime
                                End If
                            Else
                                sinY = SinY35
                            End If
                            SinX = Val(zlCommFun.Nvl(!X����))
                            For i = 1 To Len(strTime)
                                If sinY < T_DrawClient.�̶�����.Bottom Then
                                    strText = Mid(strTime, i, 1)
                                    Call GetTextExtentPoint32(lngDC, strText, Len(strText), T_Size)
                                    If T_DrawClient.�̶�����.Bottom - sinY > T_Size.H Then
                                        Call DrawRotateText(objDraw, lngDC, SinX, sinY, strText, Val(!��ɫ))
                                    End If
                                    If Asc(strText) < 0 Then
                                        sinY = sinY + T_Size.H
                                    Else
                                        sinY = sinY + T_Size.H / 2
                                    End If
                                End If
                            Next i
                            rsNotes!���� = 1
                            blnAllow = True
                        Else
                            rsNotes!���� = strTime
                            rsNotes!Y���� = SinY42
                            blnAllow = False
                        End If
                    ElseIf zlCommFun.Nvl(!����) = 6 Then
                        If blnAllow = True Then
                            If Val(zlCommFun.Nvl(!X����)) <> SinX Then
                                sinY = SinY35
                            Else
                                strTime = " " & strTime
                            End If
                        Else
                            sinY = SinY35
                        End If
                        SinX = Val(zlCommFun.Nvl(!X����))
                        For i = 1 To Len(strTime)
                            If i < 3 Then intAscCharNum = 0
                            If sinY < T_DrawClient.�̶�����.Bottom Then
                                strText = Mid(strTime, i, 1)
                                Call GetTextExtentPoint32(lngDC, strText, Len(strText), T_Size)
                                
                                If Asc(strText) < 0 Then
                                    If intAscCharNum Mod 2 = 1 Then sinY = sinY + T_Size.H / 2
                                End If
                                '���������Ϣ
                                If T_DrawClient.�̶�����.Bottom - sinY > T_Size.H Then
                                    Call DrawRotateText(objDraw, lngDC, SinX, sinY, strText, Val(zlCommFun.Nvl(!��ɫ)))
                                End If
                                If Asc(strText) < 0 Then
                                    sinY = sinY + T_Size.H
                                    intAscCharNum = 0
                                Else
                                    sinY = sinY + T_Size.H / 2
                                    intAscCharNum = intAscCharNum + 1
                                End If
                            End If
                        Next i
                        rsNotes!���� = 1
                        blnAllow = False
                        sinY = 0
                    Else
                        '���ת�ȱ����Ϣ ��ʼY�����������Ϊ42
                        rsNotes!Y���� = SinY42
                    End If
                End If
            .MoveNext
            Loop
        End With
        If rsNotes.RecordCount > 0 Then rsNotes.MoveFirst: rsNotes.Update
        stdset.Name = "����"
        stdset.Size = 9 * sngScale
        stdset.Bold = False
        Call SetFontIndirect(stdset, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        Call OutPutText(objDraw, rsDrawItems, lngDC, rsNotes, strTmpDay, sngScale)
        
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
        
        '��ʼ������±��������Ŀ����
        ReDim ArrNewString(0)
        Dim arrTmpString0() As String, arrTmpString1() As String, arrTmpString2() As String
        
        '��֯��ȡ���±����Ϣ
        For i = 0 To UBound(ArrComTable)
            lng��Ŀ��� = Val(Split(ArrComTable(i), "||")(0))
            str��Ŀ���� = Trim(Split(ArrComTable(i), "||")(1))
            If lng��Ŀ��� <> 4 Then
                j = InStr(1, str��Ŀ����, "(")
                If j > 0 Then
                    strItemName = Trim(Left(str��Ŀ����, j - 1))
                Else
                    strItemName = Trim(str��Ŀ����)
                End If
                If InStr(1, "," & strItems & ",", ",'" & strItemName & "',") = 0 Then
                    strItems = strItems & ",'" & strItemName & "'"
                End If
            End If
        Next i
        
        If Left(strItems, 1) = "," Then strItems = Mid(strItems, 2)
        If Not mbln�������� Then strItems = strItems & ",'����'"
        strItems = strItems & ",'����ѹ','����ѹ'"
        If Left(strItems, 1) = "," Then strItems = Mid(strItems, 2)
        
        dtBegin = Int(CDate(strTmpDay) - 1)
        dtEnd = CDate(CDate(Format(strEndDay, "YYYY-MM-DD HH:mm:ss")) + 1)
        If CDate(Format(dtBegin, "YYYY-MM-DD HH:mm:ss")) < CDate(Format(strBeginDate, "YYYY-MM-DD HH:mm:ss")) Then _
            dtBegin = CDate(Format(strBeginDate, "YYYY-MM-DD HH:mm:ss"))
        If CDate(Format(dtEnd, "YYYY-MM-DD HH:mm:ss")) > CDate(Format(strEndDate, "YYYY-MM-DD HH:mm:ss")) Then _
            dtEnd = CDate(Format(strEndDate, "YYYY-MM-DD HH:mm:ss"))

        
        '��ȡ���б����Ŀ������Ϣ
        strSql = "SELECT C.Id,a.����ʱ�� As ʱ��,C.��¼����,C.��ʾ,C.��¼���� As ���,C.���²�λ,C.δ��˵��,nvl(C.������Դ,0) ������Դ," & vbNewLine & _
            "  DECODE(E.��Ŀ����,2,C.���²�λ || D.��¼��,D.��¼��) ��Ŀ����,D.��Ŀ���,C.��ԴID,C.����,E.��Ŀ���� " & _
            "  FROM ���˻����ļ� B, ���˻������� A,���˻�����ϸ C,���¼�¼��Ŀ D,�����¼��Ŀ E " & _
            "  Where B.ID = A.�ļ�ID" & vbNewLine & _
            "  AND A.ID = C.��¼ID" & vbNewLine & _
            "  AND B.ID = [1]" & vbNewLine & _
            "  AND Nvl(B.Ӥ��, 0) = [7]" & vbNewLine & _
            "  AND B.����id = [2]" & vbNewLine & _
            "  AND B.��ҳid = [3]" & vbNewLine & _
            "  AND INSTR([6], DECODE(E.��Ŀ����, 2,C.���²�λ || D.��¼��, D.��¼��)) > 0" & vbNewLine & _
            "  AND D.��Ŀ��� = C.��Ŀ���" & vbNewLine & _
            "  AND Mod(c.��¼����,10) = 1" & vbNewLine & _
            "  AND E.��Ŀ��� = D.��Ŀ���" & vbNewLine & _
            "  AND E.����ȼ� >= [8]" & vbNewLine & _
            "  AND A.����ʱ�� BETWEEN [4] And [5]" & vbNewLine & _
            "  And C.��ֹ�汾 Is Null" & vbNewLine & _
            "  AND D.��¼�� = 2" & vbNewLine & _
            "  UNION ALL "
         '��ȡ�����±��Ļ�����Ŀ�����±�������Ŀ������ܴ��ڷ�������Ŀ��
        strSql = strSql & vbNewLine & _
            "  SELECT C.ID,a.����ʱ�� As ʱ��,C.��¼����,C.��ʾ,C.��¼���� As ���,C.���²�λ,C.δ��˵��,nvl(C.������Դ,0) ������Դ," & _
            "   D.��Ŀ����,D.��Ŀ���,C.��ԴID,C.����,D.��Ŀ����" & _
            "   FROM ���˻����ļ� B, ���˻������� A,���˻�����ϸ C,(SELECT A.��Ŀ���,A.��Ŀ����, 1 ��Ŀ����,B.����� FROM �����¼��Ŀ A,���������Ŀ B" & vbNewLine & _
            "       WHERE A.��Ŀ���=B.��� AND NOT EXISTS (SELECT C.��Ŀ��� FROM ���¼�¼��Ŀ C,���������Ŀ E WHERE C.��Ŀ���=E.��� AND C.��Ŀ���=A.��Ŀ���)" & vbNewLine & _
            "       AND NVL(A.Ӧ�÷�ʽ,0)=1 AND NVL(A.����ȼ�,0)>=[8] AND NVL(A.���ò���,0) IN (0,[9])" & vbNewLine & _
            "       AND (A.���ÿ���=1 OR (A.���ÿ���=2 AND EXISTS (SELECT 1 FROM �������ÿ��� D WHERE D.��Ŀ���=A.��Ŀ��� AND D.����ID=[10])))) D" & _
            "   Where B.ID=A.�ļ�ID And A.ID = C.��¼ID   AND B.ID=[1]  AND Nvl(B.Ӥ��,0)=[7] " & _
            "   AND B.����id=[2]  AND B.��ҳid=[3]  AND D.��Ŀ���=C.��Ŀ���  AND C.��¼����=1" & _
            "   AND A.����ʱ�� BETWEEN [4] And [5] And C.��ֹ�汾 Is Null"
            
        strSql = _
            "   Select ID,ʱ��,��¼����,��ʾ,���,���²�λ,δ��˵��,������Դ,��Ŀ����,��Ŀ���,��ԴID,����,��Ŀ���� From (" & strSql & ")" & _
            "   Order By  Decode(��Ŀ����,'����ѹ',0,1)," & strItems & ",ʱ��"
            
        If blnMoved Then
            strSql = Replace(strSql, "���˻����ļ�", "H���˻����ļ�")
            strSql = Replace(strSql, "���˻�������", "H���˻�������")
            strSql = Replace(strSql, "���˻�����ϸ", "H���˻�����ϸ")
        End If
        Set rsTmp = zldatabase.OpenSQLRecord(strSql, "Print", _
                                            lng�ļ�ID, _
                                            lng����ID, _
                                            lng��ҳID, _
                                            CDate(dtBegin), _
                                            CDate(dtEnd), _
                                            strItems, intBaby, lng����ȼ�, IIf(intBaby = 0, 1, 2), lngSectID)
                                                    
        ReDim Preserve ArrNewString(UBound(ArrComTable))
        For i = 0 To UBound(ArrComTable)
            If Split(ArrComTable(i), "||")(0) = 3 Then '������Ŀ
                lng��Ŀ��� = Val(Split(ArrComTable(i), "||")(0))
                strNewTmpString = String(42, ";")
                arrTmpString0 = Split(String(42, ";"), ";")
                arrTmpString1 = Split(String(42, ";"), ";")
                arrTmpString2 = Split(String(42, ";"), ";")
                
                ArrNewTmpString = Split(strNewTmpString, ";")
                
                rsTmp.Filter = "��Ŀ���=" & gint����
                With rsTmp
                    Do While Not .EOF
                        blnAdd = False
                        If CDate(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss")) >= CDate(Format(strTmpDay, "YYYY-MM-DD HH:mm:ss")) Then
                            intCOl = GetCurveColumn(CDate(!ʱ��), CDate(strTmpDay), gintHourBegin)
                            If intCOl > LBound(ArrNewTmpString) And intCOl <= UBound(ArrNewTmpString) Then
                            
                                If arrTmpString1(intCOl) <> "" Then
                                    If (Val(arrTmpString2(intCOl)) = 0 And Val(zlCommFun.Nvl(!��ʾ, 0)) = 0) Or _
                                        (Val(arrTmpString2(intCOl)) = 1 And Val(zlCommFun.Nvl(!��ʾ, 0)) = 1) Then
                                        
                                        '����Ǹ����ص�ʱ�����
                                        SinX = GetXCoordinate(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss"), Format(strTmpDay, "YYYY-MM-DD HH:mm:ss"))
                                        blnAdd = GetCanvasCenter(CDate(Format(arrTmpString1(intCOl), "YYYY-MM-DD HH:mm:ss")), CDate(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss")), CDate(Format(strTmpDay, "YYYY-MM-DD HH:mm:ss")), SinX)
                                    ElseIf Val(arrTmpString2(intCOl)) = 1 Then
                                        blnAdd = False
                                    Else
                                        blnAdd = True
                                    End If
                                    If blnAdd = True Then
                                        If Val(arrTmpString2(intCOl)) = 2 Then
                                            arrTmpString0(intCOl) = zlCommFun.Nvl(!���) & "," & zlCommFun.Nvl(!���²�λ)
                                            arrTmpString1(intCOl) = Format(!ʱ��, "YYYY-MM-DD HH:mm:ss")
                                            arrTmpString2(intCOl) = 2
                                            GoTo ErrNext
                                        End If
                                    Else
                                        If Val(zlCommFun.Nvl(!��ʾ, 0)) = 2 Then
                                            arrTmpString2(intCOl) = 2
                                            GoTo ErrNext
                                        End If
                                    End If
                                Else
                                    blnAdd = True
                                End If
                                
                                If blnAdd = True Then
                                    arrTmpString0(intCOl) = zlCommFun.Nvl(!���) & "," & zlCommFun.Nvl(!���²�λ)
                                    arrTmpString1(intCOl) = Format(!ʱ��, "YYYY-MM-DD HH:mm:ss")
                                    arrTmpString2(intCOl) = Val(zlCommFun.Nvl(!��ʾ, 0))
                                End If
                                
                            End If
                        End If
ErrNext:
                    .MoveNext
                    Loop

                    For intCOl = LBound(ArrNewTmpString) To UBound(ArrNewTmpString)
                        ArrNewTmpString(intCOl) = IIf(Val(arrTmpString2(intCOl)) = 2, "", arrTmpString0(intCOl))
                    Next intCOl
                    
                    strNewTmpString = Join(ArrNewTmpString, "||")
                End With
                ArrNewString(i) = strNewTmpString
            Else
                blnColor = False
                intƵ�� = Val(Split(ArrComTable(i), "||")(4))
                strTmp = Val(Split(ArrComTable(i), "||")(6)) '��Ŀ��ʾ 4��ʾ������Ŀ
                lng��Ŀ��� = Val(Split(ArrComTable(i), "||")(0))
                str��Ŀ���� = Trim(Split(ArrComTable(i), "||")(1))
                int��Ŀ���� = Val(Split(ArrComTable(i), "||")(5))
                int��Ŀ���� = Val(Split(ArrComTable(i), "||")(7))
                int��Ժ�ײ� = Val(Split(ArrComTable(i), "||")(8))
                
                If Val(strTmp) = 4 Or IsWaveItem(lng��Ŀ���) Then
                    If intƵ�� > 2 Then intƵ�� = 2 '����/������ĿƵ��ֻ���� 1 �� 2
                End If
                
                blnColor = (int��Ŀ���� = 2 And int��Ŀ���� = 1 And Val(strTmp) = 0)
                strNewTmpString = String(Val(intƵ��) * 7, ";")
              
                ArrNewTmpString = Split(strNewTmpString, ";")
                
                For j = 0 To 6
                    strBegin = DateAdd("D", j, CDate(strTmpDay))
                    If CDate(strBegin) > CDate(strEndDay) Then strBegin = strEndDay
                    int����ѹ = 0
                    int����ѹ = 0
                    Int�к� = 0
                    strTime = ""
                    intCOl = 0
                    
                    Set rsDownTab = ReturnItemRecord(rsTmp, Int(CDate(strBegin)), CDate(strBeginDate), lng��Ŀ��� & ";" & str��Ŀ���� & ";" & _
                                    intƵ�� & ";" & Val(strTmp) & ";" & int��Ŀ���� & ";" & int��Ժ�ײ�, bln���ܵ���, bln¼��Сʱ)
                    If rsDownTab.RecordCount > 0 Then rsDownTab.MoveFirst
                    rsDownTab.Sort = "ʱ��,��Ŀ���,���"
                    With rsDownTab
                        Do While Not .EOF
                            lngColor = 0
                            str��� = zlCommFun.Nvl(!��¼����)
                            intCOl = Val(!���)
                            intCOl = intCOl + j * intƵ��
                            If blnColor Then lngColor = Val(zlCommFun.Nvl(!δ��˵��))
                            
                            Select Case zlCommFun.Nvl(!��Ŀ����)
                                Case "����ѹ"
                                    If int����ѹ <> intCOl Then
                                        If Trim(ArrNewTmpString(intCOl)) <> "" Or str��� <> "" Then
                                            If InStr(1, ArrNewTmpString(intCOl), "/") > 0 Then
                                                ArrNewTmpString(intCOl) = Trim(Split(ArrNewTmpString(intCOl), "/")(0)) & "/" & str���
                                            Else
                                                ArrNewTmpString(intCOl) = "/" & str���
                                            End If
                                            If str��� = "���" Or str��� = "�ܲ�" Or str��� = "���" Or str��� = "δ��" Then ArrNewTmpString(intCOl) = str���
                                        End If
                                         int����ѹ = intCOl
                                         If ArrNewTmpString(intCOl) = "/" Then ArrNewTmpString(intCOl) = ""
                                    End If
                                Case "����ѹ"
                                    If int����ѹ <> intCOl Then
                                        If ArrNewTmpString(intCOl) <> "" Or str��� <> "" Then
                                            If InStr(1, ArrNewTmpString(intCOl), "/") > 0 Then
                                                ArrNewTmpString(intCOl) = str��� & "/" & Trim(Split(ArrNewTmpString(intCOl), "/")(1))
                                            Else
                                                ArrNewTmpString(intCOl) = str��� & "/"
                                            End If
                                        End If
                                        int����ѹ = intCOl
                                    End If
                                Case Else
                                    If Int�к� <> intCOl Then
                                        ArrNewTmpString(intCOl) = Replace(str���, "-#", "") & "-#" & lngColor
                                        Int�к� = intCOl
                                    End If
                            End Select
                        .MoveNext
                        Loop
                    End With
                    
                    If Format(strBegin, "YYYY-MM-DD") = Format(strEndDay, "YYYY-MM-DD") Then Exit For
                Next j
                strNewTmpString = Join(ArrNewTmpString, "||")
                ArrNewString(i) = strNewTmpString
            End If
        Next i
        
        '��Ŀ���||��λ+��Ŀ����||��Ŀ��λ||��Ŀֵ��||��¼Ƶ��||��Ŀ����||��Ŀ��ʾ
        For i = 0 To UBound(ArrComTable)
            strTmpString0 = ""

            If Trim(CStr(Split(ArrComTable(i), "||")(2))) <> "" Then
                strTmpString0 = Trim(CStr(Split(ArrComTable(i), "||")(1))) & "(" & Trim(CStr(Split(ArrComTable(i), "||")(2))) & ")"
            Else
                strTmpString0 = Trim(CStr(Split(ArrComTable(i), "||")(1)))
            End If
           
            ArrNewString(i) = Trim(CStr(Split(ArrComTable(i), "||")(0))) & ";" & strTmpString0 & ";" & ArrNewString(i)
        Next i
        
        '��ʾƤ�Խ��
        If bln��ʾƤ�� = True Then
            strSql = _
               "SELECT ʱ��,F_LIST2STR(CAST(COLLECT(ҩ����) AS T_STRLIST)) ҩ���� FROM (" & vbNewLine & _
                "   SELECT TO_CHAR(��ʼִ��ʱ��,'YYYY-MM-DD') ʱ��,DECODE(Ƥ�Խ��,'(+)',255,0) || '-#' || REPLACE(REPLACE(ҽ������,',',''),'-#','') || Ƥ�Խ��  ҩ����" & vbNewLine & _
                "   FROM ����ҽ����¼" & vbNewLine & _
                "   WHERE  ����ID=[1] AND ��ҳID=[2] AND Ӥ��=[3] AND Ƥ�Խ�� IS NOT NULL" & vbNewLine & _
                "   AND ��ʼִ��ʱ��  BETWEEN [4] AND [5]" & vbNewLine & _
                "   ORDER BY TO_DATE(TO_CHAR(��ʼִ��ʱ��,'YYYY-MM-DD'),'YYYY-MM-DD'),Ƥ�Խ��" & vbNewLine & _
                ") GROUP BY ʱ��"
                
            If blnMoved Then
                strSql = Replace(strSql, "���˹�����¼", "H���˹�����¼")
            End If
            
            Set rsTmp = zldatabase.OpenSQLRecord(strSql, "��ȡ���˹�����¼��Ϣ", lng����ID, lng��ҳID, intBaby, CDate(strTmpDay), CDate(strEndDay))
            
            strNewTmpString = String(7, ";")
            ArrNewTmpString = Split(strNewTmpString, ";")
            intCOl = 0
            
            Do While Not rsTmp.EOF
                intCOl = DateDiff("D", CDate(Format(strTmpDay, "YYYY-MM-DD")), CDate(Format(rsTmp!ʱ��, "YYYY-MM-DD"))) + 1
                ArrNewTmpString(intCOl) = Nvl(rsTmp!ҩ����)
                rsTmp.MoveNext
            Loop
            strNewTmpString = Join(ArrNewTmpString, "||")
            ReDim Preserve ArrNewString(UBound(ArrNewString) + 1)
            ArrNewString(UBound(ArrNewString)) = "-999;Ƥ�Խ��" & ";" & strNewTmpString
        End If
        
        lngCurX = X
'        stdset.Name = "����"
'        stdset.Size = 9 * sngScale
'        stdset.Bold = False
'        Call SetFontIndirect(stdset, lngDC, objDraw)
'        lngFont = CreateFontIndirect(T_Font)
'        lngOldFont = SelectObject(lngDC, lngFont)
        '��ʼ�滭�����Ŀ��չʾ����
        Call DrawBodyRecordItem(lngDC, objDraw, ArrNewString, rsItems, lngCurX, T_DrawClient.��������.Bottom, T_DrawClient.��������.Right, intRepairRows, lngCurY, sngScale)
'       Call SelectObject(lngDC, lngOldFont)
'       Call DeleteObject(lngFont)
        lngCurX = X
        lngCurY = lngCurY
        
        stdset.Name = "����"
        stdset.Size = 9 * sngScale
        stdset.Bold = False
        Call SetFontIndirect(stdset, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        '��ʼ��ӡ ҳ�� סԺ���� �� ����˵����Ϣ
        Call DrawBodyPageFooter(lngDC, objDraw, lngCurX, lngCurY, T_DrawClient.��������.Right, intPageNo, intEndPage, str����˵��, sngScale)
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
        'ҳ��ͼ�����
        Call frmTendFileRead.PrintRTBData(objDraw, False, lngButtom)
        
        If Not blnPrint Then objDraw.Refresh
NOPageSub:  Next

    If blnPrint = False Then Call DrawDeviceCaps(lngDC, objDraw)
     
    Call ShowFlash
    PrintOrPreviewBodyState = True
    Screen.MousePointer = 0
    Set stdset = Nothing
    GoTo ErrClare
    Exit Function
ErrPrint:
    Call ShowFlash
    Screen.MousePointer = 0
    
    If ErrCenter = 1 Then
        Resume
    End If
    GoTo ErrClare
    Call SaveErrLog
ErrExit:
    Call ShowFlash
    Screen.MousePointer = 0
    msngTwips = 1
    Err.Clear
    PrintOrPreviewBodyState = False
    Set stdset = Nothing
    GoTo ErrClare
ErrClare:
    T_DrawClient.ƫ����X = M_DrawClient.ƫ����X
    T_DrawClient.ƫ����Y = M_DrawClient.ƫ����Y
    T_DrawClient.�̶����� = M_DrawClient.�̶�����
    T_DrawClient.�̶ȵ�λ = M_DrawClient.�̶ȵ�λ
    T_DrawClient.�������� = M_DrawClient.��������
    T_DrawClient.�е�λ = M_DrawClient.�е�λ
    T_DrawClient.ʱ���е�λ = M_DrawClient.ʱ���е�λ
    T_DrawClient.ʱ���е�λ = M_DrawClient.ʱ���е�λ
    T_DrawClient.�е�λ = M_DrawClient.�е�λ
    T_DrawClient.˫�� = M_DrawClient.˫��
    T_DrawClient.������ = M_DrawClient.������
    Call ErrEmpty
    Set stdset = Nothing
End Function

Private Sub ErrEmpty()
    msngTwips = 1
    T_TwipsPerPixel.X = Screen.TwipsPerPixelX
    T_TwipsPerPixel.Y = Screen.TwipsPerPixelY
End Sub


Private Sub DrawDeviceCaps(ByVal lngDC As Long, ByVal objDraw As Object)
    Dim dblSureW As Double, dblSureH As Double
    '����Ǵ�ӡԤ��,Ӧ����ӡ���Ŀɴ�ӡ�Ŀ�ʼ����ʼԤ��
    dblSureW = GetDeviceCaps(Printer.hDC, PHYSICALOFFSETX) / GetDeviceCaps(Printer.hDC, PHYSICALWIDTH)
    dblSureH = GetDeviceCaps(Printer.hDC, PHYSICALOFFSETY) / GetDeviceCaps(Printer.hDC, PHYSICALHEIGHT)
    On Error Resume Next
    Call DrawRect(lngDC, (objDraw.Width * dblSureW) / T_TwipsPerPixel.X, (objDraw.Height * (1 - dblSureH)) / T_TwipsPerPixel.Y, _
    (objDraw.Width * (1 - dblSureW)) / T_TwipsPerPixel.X, objDraw.Height * dblSureH / T_TwipsPerPixel.Y, PS_DOT, 1, RGB_FleetGRAY)
End Sub
Public Sub DrawBodyRecordItem(ByVal lngDC As Long, ByVal objDraw As Object, strValue() As String, ByVal rsItems As ADODB.Recordset, ByVal lngX As Long, ByVal lngY As Long, _
    ByVal lngLeft As Long, ByVal intRepairRows As Integer, lngOutY As Long, Optional sngScale As Single = 1)
'-----------------------------------------------------------------------------------------------------------------------
'������˻�����Ϣ
'����:lngDC ��ͼ�����DC��strValue() ���б����Ŀ����Ϣ (��ʽ��������:��Ŀ���;����;����,��λ||����,��λ/(����) ��Ŀ���;����;����||����) ���ݺͲ�λ��ɵ������ʾ����Ŀ�ж�����
'    rsItems �������±������Ŀ, lngX ��߾�,lngY�ϱ߾�,lngLeft �ұ߾�(���Ի�ͼ������ұ߾�),intRepairRows Ҫ��ӡ�����Ŀ��������
'����:lngOutY ���ػ�ͼ����ϱ߾�
'-----------------------------------------------------------------------------------------------------------------------
    Dim lngX1 As Long, lngY1 As Long, lngCurY As Long, lngCurX As Long
    Dim lngRowHeiht As Long
    Dim arrTmpString0() As String, arrTmpString1() As String
    Dim arrTmp() As String, arrText() As String
    Dim intRow As Integer, intCOl As Integer
    Dim i As Integer
    Dim int������������ʽ As Integer
    Dim bln�೦����Է��ӷ�ĸ��ʾ As Boolean
    Dim strTmp As String, strPart As String
    Dim strPic As String
    Dim blnValue As Boolean
    Dim intValue As Integer, int����λ�� As Integer
    Dim intRowCount As Integer
    Dim intƵ�� As Integer '��¼Ƶ��
    Dim blnDataTrue As Boolean
    Dim lngColor As Long
    Dim intNum As Integer
    Dim blnOutText As Boolean '�Ƿ�����ı�
    Dim blnPrinter As Boolean
    Dim intBold As Integer, intFine As Integer
    Dim intSize As Integer
    Dim sngLen As Single, lngLen As Long
    Dim LPoint As T_LPoint
    Dim bln��ʾƤ�� As Boolean
    
    If UBound(strValue) < 0 Then Exit Sub
    If IsEmpty(strValue) = True Then Exit Sub
    
    On Error GoTo Errhand
    
    blnPrinter = False
    If TypeName(objDraw) = "Printer" Then
        blnPrinter = True
        intBold = 6
        intFine = 2
    Else
        msngTwips = 1
        intBold = 2
        intFine = 1
    End If
    
    lngCurY = lngY
    lngCurX = lngX
    blnValue = False
    intValue = 0
    int����λ�� = 0
    int������������ʽ = zldatabase.GetPara("����������", glngSys, 1255, 0)
    bln�೦����Է��ӷ�ĸ��ʾ = (Val(zldatabase.GetPara("�೦������ʾ��ʽ", glngSys, 1255, 0)) = 1)
    
    strPic = ""
    If InStr(1, strValue(0), ";") > 0 Then
        bln��ʾƤ�� = IIf(Split(strValue(UBound(strValue)), ";")(0) = "-999", True, False)
        
        For intRow = LBound(strValue) To UBound(strValue)
            arrTmpString0 = Split(strValue(intRow), ";")
            arrTmpString1 = Split(arrTmpString0(2), "||")
            
            If intRepairRows > 0 And intRepairRows > intRowCount Then
            
                If arrTmpString0(0) = "3" Then '������Ŀ
                    '��ȡ�����ɫ
                    rsItems.Filter = 0
                    rsItems.Filter = "��Ŀ���=" & gint����
                    If rsItems.RecordCount > 0 Then
                        lngColor = Val(Nvl(rsItems!��¼ɫ, RGB_RED))
                    Else
                        lngColor = RGB_RED
                    End If
                    intRowCount = intRowCount + 1
                    arrTmpString1 = Split(arrTmpString0(2), "||")
                    For intCOl = 0 To UBound(arrTmpString1)
                        If intCOl = 0 Then '��ͷ
                            Call SetTextColor(lngDC, RGB_BLACK)
                            Call GetTextExtentPoint32(lngDC, arrTmpString0(intCOl + 1), Len(arrTmpString0(intCOl + 1)), T_Size)
                            Call GetTextRect(objDraw, lngX, lngY + (T_DrawClient.ʱ���е�λ + T_DrawClient.ʱ���е�λ / 2) / 2, arrTmpString0(intCOl + 1), _
                                T_DrawClient.�̶�����.Right - lngX, True, , sngScale)
                            'Call DrawText(lngDC, arrTmpString0(intCOl + 1), -1, T_LableRect, DT_CENTER)
                            LPoint.X = lngX
                            LPoint.Y = lngY + (T_DrawClient.ʱ���е�λ + T_DrawClient.ʱ���е�λ / 2) / 2
                            LPoint.W = T_DrawClient.�̶�����.Right - lngX
                            Call DrawTabText(lngDC, objDraw, arrTmpString0(intCOl + 1), -1, T_LableRect, DT_CENTER, LPoint, sngScale)
                            Call DrawLine(lngDC, lngX, lngY, lngX, lngY + T_DrawClient.ʱ���е�λ + T_DrawClient.ʱ���е�λ / 2, PS_SOLID, intBold, RGB_BLACK)
                            Call DrawLine(lngDC, lngX, lngY + T_DrawClient.ʱ���е�λ + T_DrawClient.ʱ���е�λ / 2, T_DrawClient.�̶�����.Right, _
                                lngY + T_DrawClient.ʱ���е�λ + T_DrawClient.ʱ���е�λ / 2, PS_SOLID, IIf(intRowCount = intRepairRows, intBold, intFine), RGB_BLACK)
                            Call DrawLine(lngDC, T_DrawClient.�̶�����.Right, lngY, T_DrawClient.�̶�����.Right, _
                                lngY + T_DrawClient.ʱ���е�λ + T_DrawClient.ʱ���е�λ / 2, PS_SOLID, intBold, RGB_BLACK)
                            lngX1 = T_DrawClient.�̶�����.Right
                            lngY1 = lngCurY
                        Else
                            arrTmpString1(intCOl) = arrTmpString1(intCOl) & String(1 - UBound(Split(arrTmpString1(intCOl), ",")), ",")
                            strTmp = Split(arrTmpString1(intCOl), ",")(0)
                            strPart = Split(arrTmpString1(intCOl), ",")(1)
                            If strPart = "" Then strPart = "��������"
                            strPic = ""
                            '��ӡ����ֵ���������ӡ�� ��һ��ʼ��������
                            If IsNumeric(strTmp) Then
                                If strPart = "��������" Then
                                    Call SetTextColor(lngDC, lngColor)
                                    Call GetTextExtentPoint32(lngDC, strTmp, Len(strTmp), T_Size)
                                Else
                                    strPic = "BREATH"
                                End If
                                
                                If blnValue = False Then
                                    intValue = IIf(intCOl Mod 2 = 0, 0, 1)
                                    blnValue = True
                                    int����λ�� = 2
                                End If
                                
                                If int������������ʽ = 0 Then '˳��������ʾ
                                    If intCOl Mod 2 = intValue Then
                                        If strPic = "" Then
                                            LPoint.X = lngX1
                                            LPoint.Y = lngY
                                            Call GetTextRect(objDraw, lngX1, lngY, strTmp, T_DrawClient.�е�λ, False, , sngScale)
                                        Else
                                            Call DrawPicture(objDraw, strPic, objDraw.ScaleX(lngX1 + ((T_DrawClient.�е�λ - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2), vbPixels, vbTwips), _
                                                objDraw.ScaleY(lngY + 1, vbPixels, vbTwips), _
                                                objDraw.ScaleX(lngX1 + ((T_DrawClient.�е�λ - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2) + mintBmpW * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), _
                                                objDraw.ScaleY(lngY + 1 + mintBmpH * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), True)
                                            
                                        End If
                                    Else
                                        If strPic = "" Then
                                            LPoint.X = lngX1
                                            LPoint.Y = lngY + ((T_DrawClient.ʱ���е�λ + T_DrawClient.ʱ���е�λ / 2) - T_Size.H)
                                            Call GetTextRect(objDraw, lngX1, lngY + ((T_DrawClient.ʱ���е�λ + T_DrawClient.ʱ���е�λ / 2) - T_Size.H), _
                                                strTmp, T_DrawClient.�е�λ, False, , sngScale)
                                        Else
                                            Call DrawPicture(objDraw, strPic, objDraw.ScaleX(lngX1 + ((T_DrawClient.�е�λ - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2), _
                                                vbPixels, vbTwips), objDraw.ScaleY(lngY + ((T_DrawClient.ʱ���е�λ + T_DrawClient.ʱ���е�λ / 2) - mintBmpH * IIf(blnPrinter = True, msngTwips, 1)), vbPixels, vbTwips), _
                                                objDraw.ScaleX(lngX1 + ((T_DrawClient.�е�λ - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2) + mintBmpW * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), _
                                                objDraw.ScaleY(lngY + (T_DrawClient.ʱ���е�λ + T_DrawClient.ʱ���е�λ / 2), vbPixels, vbTwips), True)
                                        End If
                                    End If
                                    
                                Else        '������ʱ����֮��������ʾ
                                    If int����λ�� = 2 Then
                                        If strPic = "" Then
                                            LPoint.X = lngX1
                                            LPoint.Y = lngY
                                            Call GetTextRect(objDraw, lngX1, lngY, strTmp, T_DrawClient.�е�λ, False, , sngScale)
                                        Else
                                            Call DrawPicture(objDraw, strPic, objDraw.ScaleX(lngX1 + ((T_DrawClient.�е�λ - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2), vbPixels, vbTwips), _
                                                objDraw.ScaleY(lngY + 1, vbPixels, vbTwips), _
                                                objDraw.ScaleX(lngX1 + ((T_DrawClient.�е�λ - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2) + mintBmpW * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), _
                                                objDraw.ScaleY(lngY + 1 + mintBmpH * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), True)
                                        End If
                                    Else
                                        If strPic = "" Then
                                            LPoint.X = lngX1
                                            LPoint.Y = lngY + ((T_DrawClient.ʱ���е�λ + T_DrawClient.ʱ���е�λ / 2) - T_Size.H)
                                            Call GetTextRect(objDraw, lngX1, lngY + ((T_DrawClient.ʱ���е�λ + T_DrawClient.ʱ���е�λ / 2) - T_Size.H), _
                                                strTmp, T_DrawClient.�е�λ, False, , sngScale)
                                        Else
                                            Call DrawPicture(objDraw, strPic, objDraw.ScaleX(lngX1 + ((T_DrawClient.�е�λ - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2), vbPixels, vbTwips), _
                                                objDraw.ScaleY(lngY + ((T_DrawClient.ʱ���е�λ + T_DrawClient.ʱ���е�λ / 2) - mintBmpH * IIf(blnPrinter = True, msngTwips, 1)), vbPixels, vbTwips), _
                                                objDraw.ScaleX(lngX1 + ((T_DrawClient.�е�λ - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2) + mintBmpW * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), _
                                                objDraw.ScaleY(lngY + (T_DrawClient.ʱ���е�λ + T_DrawClient.ʱ���е�λ / 2), vbPixels, vbTwips), True)
                                        End If
                                    End If
                                    
                                   
                                    int����λ�� = int����λ�� + 1
                                    If int����λ�� > 2 Then int����λ�� = 1
                                End If
                                LPoint.W = T_DrawClient.�е�λ
                                If strPic = "" Then Call DrawTabText(lngDC, objDraw, strTmp, -1, T_LableRect, DT_CENTER, LPoint, sngScale) 'DrawText(lngDC, strTmp, -1, T_LableRect, DT_CENTER)
                                
                            End If
                            lngX1 = lngX1 + T_DrawClient.�е�λ
                        End If
                    Next intCOl
                    lngX1 = T_DrawClient.�̶�����.Right + T_DrawClient.�е�λ
                    lngY1 = lngY + T_DrawClient.ʱ���е�λ + T_DrawClient.ʱ���е�λ / 2
                    
                    '�����������е���
                    For intCOl = 1 To 42
                        If intCOl Mod 6 = 0 Then
                            Call DrawLine(lngDC, lngX1, lngY, lngX1, lngY1, PS_SOLID, intBold, RGB_BLACK)
                        Else
                            Call DrawLine(lngDC, lngX1, lngY, lngX1, lngY1, PS_SOLID, intFine, RGB_BLACK)
                        End If
                        lngX1 = lngX1 + T_DrawClient.�е�λ
                    Next intCOl
                    Call DrawLine(lngDC, T_DrawClient.�̶�����.Right, lngY1, T_DrawClient.��������.Right, lngY1, PS_SOLID, IIf(intRowCount = intRepairRows, intBold, intFine), RGB_BLACK)
                    
                    '��ǰY������
                    lngCurY = lngY1
                ElseIf arrTmpString0(0) <> "-999" Then '����Ƥ�Խ��
                    
                    rsItems.Filter = ""
                    rsItems.Filter = "���=" & intRow
                    If rsItems.RecordCount > 0 Then
                        intƵ�� = CInt(zlCommFun.Nvl(rsItems!��¼Ƶ��, 2))
                        If Val(Nvl(rsItems!��Ŀ��ʾ)) = 4 Or IsWaveItem(Val(Nvl(rsItems!��Ŀ���))) Then
                            If intƵ�� > 2 Then intƵ�� = 2 '����/������ĿƵ��ֻ���� 1 �� 2
                        End If
                        '���Ŀ����Ƿ�������ݣ������ھͲ���ӡ����
                        If zlCommFun.Nvl(rsItems!��Ŀ����) = 2 Then
                            
                            If Trim(Replace(arrTmpString0(2), "||", "")) = "" Then
                                blnDataTrue = False
                            Else
                                blnDataTrue = True
                            End If
                        Else
                            blnDataTrue = True
                        End If
                    Else
                        blnDataTrue = False
                    End If
                    
                    If blnDataTrue = True Then
                        lngY1 = lngCurY
                        lngX1 = lngCurX
                        
                        '����Ƶ�μ���Ҫ��ӡ�ı�������Ƿ񳬳��û����õı������
                        
                        intNum = 0
                        Select Case intƵ��
                            Case 1, 2, 6
                                intRowCount = intRowCount + 1
                            Case 3
                                intRowCount = intRowCount + 3
                            Case 4
                                intRowCount = intRowCount + 2
                            Case Else
                                intRowCount = intRowCount + 1
                        End Select
                        
                        If intRowCount > intRepairRows Then
                            intNum = intRowCount - intRepairRows
                            intRowCount = intRepairRows
                        End If
                        blnOutText = False
                        
                        For intCOl = 0 To UBound(arrTmpString1)
                            If intCOl = 0 Then '��ʼ����ͷ��Ϣ������������
                                Select Case intƵ��
                                    Case 1, 2, 6
                                        lngY1 = lngY1 + T_DrawClient.ʱ���е�λ
                                        lngRowHeiht = T_DrawClient.ʱ���е�λ / 2
                                    Case 3
                                        lngY1 = lngY1 + T_DrawClient.ʱ���е�λ * (3 - intNum)
                                        lngRowHeiht = (T_DrawClient.ʱ���е�λ * (3 - intNum)) / 2
                                    Case 4
                                        lngY1 = lngY1 + T_DrawClient.ʱ���е�λ * (2 - intNum)
                                        lngRowHeiht = (T_DrawClient.ʱ���е�λ * (2 - intNum)) / 2
                                End Select

                                Call SetTextColor(lngDC, RGB_BLACK)
                                Call GetTextExtentPoint32(lngDC, arrTmpString0(intCOl + 1), Len(arrTmpString0(intCOl + 1)), T_Size)
                                Call GetTextRect(objDraw, lngX1, lngY1 - lngRowHeiht, arrTmpString0(intCOl + 1), T_DrawClient.�̶�����.Right - lngX1, True, , sngScale)
                                'Call DrawText(lngDC, arrTmpString0(intCOl + 1), -1, T_LableRect, DT_CENTER)
                                LPoint.X = lngX1
                                LPoint.Y = lngY1 - lngRowHeiht
                                LPoint.W = T_DrawClient.�̶�����.Right - lngX1
                                Call DrawTabText(lngDC, objDraw, arrTmpString0(intCOl + 1), -1, T_LableRect, DT_CENTER, LPoint, sngScale)
                                Call DrawLine(lngDC, lngX1, lngCurY, lngX1, lngY1, PS_SOLID, intBold, RGB_BLACK)
                                Call DrawLine(lngDC, lngX1, lngY1, T_DrawClient.�̶�����.Right, lngY1, PS_SOLID, IIf(intRowCount = intRepairRows, intBold, intFine), RGB_BLACK)
                                Call DrawLine(lngDC, T_DrawClient.�̶�����.Right, lngCurY, T_DrawClient.�̶�����.Right, lngY1, PS_SOLID, intBold, RGB_BLACK)
                                
                                lngY1 = lngCurY
                                lngX1 = T_DrawClient.�̶�����.Right
                            Else  '��ʼ���л������
                                strTmp = CStr(arrTmpString1(intCOl))
                               
                                If InStr(1, strTmp, "-#") <> 0 Then
                                    If Not IsNumeric(Split(strTmp, "-#")(1)) Then
                                        lngColor = 0
                                    Else
                                        lngColor = Val(Split(strTmp, "-#")(1))
                                        strTmp = Split(strTmp, "-#")(0)
                                    End If
                                Else
                                    lngColor = 0
                                End If
                                
                                If strTmp = "*" And Val(arrTmpString0(0)) = gint��� Then strTmp = "��"
                                
                                Call SetTextColor(lngDC, lngColor)
                                
                                Call GetTextExtentPoint32(lngDC, strTmp, Len(strTmp), T_Size)
                                blnOutText = True
                                
                                If InStr(1, ",3,4,", "," & intƵ�� & ",") = 0 Then
                                    LPoint.X = lngX1
                                    LPoint.Y = lngCurY + T_DrawClient.ʱ���е�λ / 2
                                    LPoint.W = T_DrawClient.�е�λ * (6 / intƵ��)
                                    Call GetTextRect(objDraw, lngX1, lngCurY + T_DrawClient.ʱ���е�λ / 2, strTmp, T_DrawClient.�е�λ * (6 / intƵ��), True, , sngScale)
                                    lngX1 = lngX1 + T_DrawClient.�е�λ * (6 / intƵ��)
                                ElseIf intƵ�� = 3 Then
                                    LPoint.W = T_DrawClient.�е�λ * 6
                                    If intCOl Mod intƵ�� = 0 Then
                                        LPoint.X = lngX1
                                        LPoint.Y = lngCurY + T_DrawClient.ʱ���е�λ * 2 + T_DrawClient.ʱ���е�λ / 2
                                        Call GetTextRect(objDraw, lngX1, lngCurY + T_DrawClient.ʱ���е�λ * 2 + T_DrawClient.ʱ���е�λ / 2, strTmp, T_DrawClient.�е�λ * 6, True, , sngScale)
                                        If intNum <> 0 Then blnOutText = False
                                        lngX1 = lngX1 + T_DrawClient.�е�λ * 6
                                    ElseIf intCOl Mod intƵ�� = 2 Then
                                        LPoint.X = lngX1
                                        LPoint.Y = lngCurY + T_DrawClient.ʱ���е�λ + T_DrawClient.ʱ���е�λ / 2
                                        Call GetTextRect(objDraw, lngX1, lngCurY + T_DrawClient.ʱ���е�λ + T_DrawClient.ʱ���е�λ / 2, strTmp, T_DrawClient.�е�λ * 6, True, , sngScale)
                                        If intNum > 1 Then blnOutText = False
                                    Else
                                        LPoint.X = lngX1
                                        LPoint.Y = lngCurY + T_DrawClient.ʱ���е�λ / 2
                                        Call GetTextRect(objDraw, lngX1, lngCurY + T_DrawClient.ʱ���е�λ / 2, strTmp, T_DrawClient.�е�λ * 6, True, , sngScale)
                                    End If
                                    
                                ElseIf intƵ�� = 4 Then
                                    LPoint.W = T_DrawClient.�е�λ * 3
                                    If intCOl Mod 4 = 3 Then
                                        LPoint.X = lngX1
                                        LPoint.Y = lngCurY + T_DrawClient.ʱ���е�λ + T_DrawClient.ʱ���е�λ / 2
                                        Call GetTextRect(objDraw, lngX1, lngCurY + T_DrawClient.ʱ���е�λ + T_DrawClient.ʱ���е�λ / 2, strTmp, T_DrawClient.�е�λ * 3, True, , sngScale)
                                        If intNum > 0 Then blnOutText = False
                                        lngX1 = lngX1 + T_DrawClient.�е�λ * 3
                                    ElseIf intCOl Mod 4 = 0 Then
                                        LPoint.X = lngX1
                                        LPoint.Y = lngCurY + T_DrawClient.ʱ���е�λ + T_DrawClient.ʱ���е�λ / 2
                                        Call GetTextRect(objDraw, lngX1, lngCurY + T_DrawClient.ʱ���е�λ + T_DrawClient.ʱ���е�λ / 2, strTmp, T_DrawClient.�е�λ * 3, True, , sngScale)
                                        If intNum > 0 Then blnOutText = False
                                        lngX1 = lngX1 + T_DrawClient.�е�λ * 3
                                    ElseIf intCOl Mod 2 = 0 Then
                                        LPoint.X = lngX1
                                        LPoint.Y = lngCurY + T_DrawClient.ʱ���е�λ / 2
                                        Call GetTextRect(objDraw, lngX1, lngCurY + T_DrawClient.ʱ���е�λ / 2, strTmp, T_DrawClient.�е�λ * 3, True, , sngScale)
                                        lngX1 = lngX1 - T_DrawClient.�е�λ * 3
                                    ElseIf intCOl Mod 4 = 1 Then
                                        LPoint.X = lngX1
                                        LPoint.Y = lngCurY + T_DrawClient.ʱ���е�λ / 2
                                        Call GetTextRect(objDraw, lngX1, lngCurY + T_DrawClient.ʱ���е�λ / 2, strTmp, T_DrawClient.�е�λ * 3, True, , sngScale)
                                        lngX1 = lngX1 + T_DrawClient.�е�λ * 3
                                    End If
                                End If
                                
                                If blnOutText = True Then
                                    If AnsyGrade(Val(arrTmpString0(0)), strTmp, arrText) = True Then
                                        Call DrawAnsyGrade(lngDC, objDraw, arrText, LPoint, lngColor, bln�೦����Է��ӷ�ĸ��ʾ, sngScale)
                                    Else
                                        Call DrawTabText(lngDC, objDraw, strTmp, -1, T_LableRect, DT_CENTER, LPoint, sngScale)
                                    End If
                                End If
                   
                            End If
                        Next intCOl
                        
                        '����Ԫ������
                        If InStr(1, ",2,3,4,", "," & intƵ�� & ",") = 0 Then
                            lngX1 = T_DrawClient.�̶�����.Right + T_DrawClient.�е�λ * (6 / intƵ��)
                            lngY1 = lngCurY + T_DrawClient.ʱ���е�λ
                            For intCOl = 1 To intƵ�� * 7
                                Call DrawLine(lngDC, lngX1, lngCurY, lngX1, lngY1, PS_SOLID, IIf(intCOl Mod intƵ�� = 0, intBold, intFine), RGB_BLACK)
                                lngX1 = lngX1 + T_DrawClient.�е�λ * (6 / intƵ��)
                            Next intCOl
                            Call DrawLine(lngDC, T_DrawClient.�̶�����.Right, lngY1, T_DrawClient.��������.Right, lngY1, PS_SOLID, IIf(intRowCount = intRepairRows, intBold, intFine), RGB_BLACK)
                        ElseIf intƵ�� = 3 Then
                            intRowCount = intRowCount - (intƵ�� - intNum)
                            intValue = intRowCount
                            For i = 1 To 3 - intNum
                                lngX1 = T_DrawClient.�̶�����.Right + T_DrawClient.�е�λ * 6
                                lngY1 = lngCurY + T_DrawClient.ʱ���е�λ
                                For intCOl = 1 To 7
                                    Call DrawLine(lngDC, lngX1, lngCurY, lngX1, lngY1, PS_SOLID, intBold, RGB_BLACK)
                                    lngX1 = lngX1 + T_DrawClient.�е�λ * 6
                                Next intCOl
                                intRowCount = intValue + i
                                Call DrawLine(lngDC, T_DrawClient.�̶�����.Right, lngY1, T_DrawClient.��������.Right, lngY1, PS_SOLID, IIf(intRowCount = intRepairRows, intBold, intFine), RGB_BLACK)
                                
                                lngCurY = lngY1
                            Next i
                        ElseIf InStr(1, ",2,4,", "," & intƵ�� & ",") <> 0 Then
                            intRowCount = intRowCount - (intƵ�� / 2 - intNum)
                            intValue = intRowCount
                            For i = 1 To (intƵ�� / 2 - intNum)
                                lngY1 = lngCurY + T_DrawClient.ʱ���е�λ
                                lngX1 = T_DrawClient.�̶�����.Right + T_DrawClient.�е�λ * 3
                                For intCOl = 1 To 14
                                    Call DrawLine(lngDC, lngX1, lngCurY, lngX1, lngY1, PS_SOLID, IIf(intCOl Mod 2 = 0, intBold, intFine), RGB_BLACK)
                                    lngX1 = lngX1 + T_DrawClient.�е�λ * 3
                                Next intCOl
                                intRowCount = intValue + i
                                Call DrawLine(lngDC, T_DrawClient.�̶�����.Right, lngY1, T_DrawClient.��������.Right, lngY1, PS_SOLID, IIf(intRowCount = intRepairRows, intBold, intFine), RGB_BLACK)
                                lngCurY = lngY1
                            Next i
                        End If
                        
                        lngCurY = lngY1
                    End If
                End If
                
                intNum = 0
                
                'Ƥ�Խ��,ֻ�����������ݣ�����ڲ������д���
                If arrTmpString0(0) = "-999" Then
                    lngY1 = lngCurY
                    lngX1 = lngCurX
                    intƵ�� = 1
                    For intCOl = 0 To UBound(arrTmpString1)
                        If intCOl = 0 Then '��ʼ����ͷ��Ϣ������������
                            lngY1 = lngY1 + T_DrawClient.ʱ���е�λ
                            lngRowHeiht = T_DrawClient.ʱ���е�λ / 2
                               
                            Call SetTextColor(lngDC, RGB_BLACK)
                            Call GetTextExtentPoint32(lngDC, arrTmpString0(intCOl + 1), Len(arrTmpString0(intCOl + 1)), T_Size)
                            Call GetTextRect(objDraw, lngX1, lngY1 - lngRowHeiht, arrTmpString0(intCOl + 1), T_DrawClient.�̶�����.Right - lngX1, True, , sngScale)
                
                            LPoint.X = lngX1
                            LPoint.Y = lngY1 - lngRowHeiht
                            LPoint.W = T_DrawClient.�̶�����.Right - lngX1
                            Call DrawTabText(lngDC, objDraw, arrTmpString0(intCOl + 1), -1, T_LableRect, DT_CENTER, LPoint, sngScale)
                            
                            lngY1 = lngCurY
                            lngX1 = T_DrawClient.�̶�����.Right
                        Else  '��ʼ���л������
                            intNum = 1
                            strTmp = CStr(arrTmpString1(intCOl))
                            If strTmp = "" Then strTmp = "-#"
                            LPoint.X = lngX1
                            LPoint.Y = lngCurY + T_DrawClient.ʱ���е�λ / 2
                            LPoint.W = T_DrawClient.�е�λ * (6 / intƵ��)
                            '��ʼ�����Ƿ���Ҫ����
                            strPart = ""
                            
                            arrTmp = Split(strTmp, ",")
                            
                            For i = LBound(arrTmp) To UBound(arrTmp)
                                lngColor = Val(Split(arrTmp(i), "-#")(0))
                                '����������ɫ
                                Call SetTextColor(lngDC, lngColor)
                                strTmp = Replace(CStr(Split(arrTmp(i), "-#")(1)), vbCrLf, "") 'Ƥ�Խ��
                                If Trim(strTmp) <> "" Then
                                    If i < UBound(arrTmp) Then strTmp = strTmp & ","
                                    Do While True
                                        T_Size.W = objDraw.TextWidth(strTmp) / T_TwipsPerPixel.X
                                        strPic = strTmp
                                        If T_Size.W - (LPoint.W - (LPoint.X - lngX1)) > 0 Then
                                            sngLen = Round((LPoint.W - (LPoint.X - lngX1)) / T_Size.W, 2)
                                            lngLen = Len(StrConv(strTmp, vbFromUnicode)) * sngLen
                                            '�����תΪȫ��
                                            strTmp = StrConv(strTmp, vbWide)
                                            strPart = StrConv(Mid(StrConv(strTmp, vbFromUnicode), lngLen + 1), vbUnicode)
                                            strTmp = StrConv(Mid(StrConv(strTmp, vbFromUnicode), 1, lngLen), vbUnicode)
                                            '��ȡԭʼ�ַ���
                                            strPart = Mid(strPic, Len(strTmp) + 1)
                                            strTmp = Mid(strPic, 1, Len(strTmp))
                                            Call GetTextRect(objDraw, LPoint.X, LPoint.Y, CStr(strTmp), , True, , sngScale)
                                            Call DrawTabText(lngDC, objDraw, CStr(strTmp), -1, T_LableRect, DT_CENTER, LPoint, sngScale)
                                            
                                            T_Size.W = objDraw.TextWidth(strTmp) / T_TwipsPerPixel.X
                                            LPoint.X = LPoint.X + T_Size.W
                                            strTmp = strPart
                                            T_Size.W = objDraw.TextWidth("��") / T_TwipsPerPixel.X
                                            If T_Size.W - (LPoint.W - (LPoint.X - lngX1)) > 0 Then
                                                LPoint.X = lngX1
                                                LPoint.Y = LPoint.Y + T_DrawClient.ʱ���е�λ
                                                intNum = intNum + 1
                                                
                                                If intRowCount + intNum > intRepairRows Then GoTo ErrNext
                                            End If
                                            If strTmp = "" Then Exit Do
                                        Else
                                            Call GetTextRect(objDraw, LPoint.X, LPoint.Y, CStr(strTmp), , True, , sngScale)
                                            Call DrawTabText(lngDC, objDraw, CStr(strTmp), -1, T_LableRect, DT_CENTER, LPoint, sngScale)
                                            If T_Size.W + objDraw.TextWidth("��") / T_TwipsPerPixel.X - LPoint.W > 0 Then
                                                LPoint.X = lngX1
                                                LPoint.Y = LPoint.Y + T_DrawClient.ʱ���е�λ
                                            Else
                                                LPoint.X = LPoint.X + T_Size.W
                                            End If
                                    
                                            Exit Do
                                        End If
                                    Loop
                                End If
                            Next i
ErrNext:
                            
                            lngX1 = lngX1 + T_DrawClient.�е�λ * (6 / intƵ��)
                        End If
                    Next intCOl
                End If
            End If
        Next intRow
        
        '������
        If intRepairRows > 0 And intRepairRows > intRowCount Then
            intRowCount = intRowCount + 1
            For intRow = intRowCount To intRepairRows
                lngX1 = lngCurX
                lngY1 = lngCurY + T_DrawClient.ʱ���е�λ
                
                '�ո�ÿ������
'                For intCOl = 0 To 14
'                    If intCOl = 0 Then
'                        Call DrawLine(lngDc, lngX1, lngCurY, lngX1, lngY1, PS_SOLID, 2, RGB_BLACK)
'                        Call DrawLine(lngDc, lngX1, lngY1, T_DrawClient.�̶�����.Right, lngY1, PS_SOLID, IIf(intRow = intRepairRows, 2, 1), RGB_BLACK)
'                        Call DrawLine(lngDc, T_DrawClient.�̶�����.Right, lngCurY, T_DrawClient.�̶�����.Right, lngY1, PS_SOLID, 2, RGB_BLACK)
'                    Else
'
'                        lngX1 = T_DrawClient.�̶�����.Right + (T_DrawClient.�е�λ * 3) * intCOl
'                        Call DrawLine(lngDc, lngX1, lngCurY, lngX1, lngY1, PS_SOLID, IIf(intCOl Mod 2 = 0, 2, 1), RGB_BLACK)
'                        If intCOl = 14 Then
'                            Call DrawLine(lngDc, T_DrawClient.�̶�����.Right, lngY1, T_DrawClient.��������.Right, lngY1, PS_SOLID, IIf(intRow = intRepairRows, 2, 1), RGB_BLACK)
'                        End If
'                    End If
'                Next intCOl
                
                '�ո�ÿ��1��
                For intCOl = 0 To 7
                    If intCOl = 0 Then
                        Call DrawLine(lngDC, lngX1, lngCurY, lngX1, lngY1, PS_SOLID, intBold, RGB_BLACK)
                        Call DrawLine(lngDC, lngX1, lngY1, T_DrawClient.�̶�����.Right, lngY1, PS_SOLID, IIf(intRow = intRepairRows, intBold, intFine), RGB_BLACK)
                        Call DrawLine(lngDC, T_DrawClient.�̶�����.Right, lngCurY, T_DrawClient.�̶�����.Right, lngY1, PS_SOLID, intBold, RGB_BLACK)
                    Else
                        
                        lngX1 = T_DrawClient.�̶�����.Right + (T_DrawClient.�е�λ * 6) * intCOl
                        Call DrawLine(lngDC, lngX1, lngCurY, lngX1, lngY1, PS_SOLID, intBold, RGB_BLACK)
                        If intCOl = 7 Then
                            Call DrawLine(lngDC, T_DrawClient.�̶�����.Right, lngY1, T_DrawClient.��������.Right, lngY1, PS_SOLID, IIf(intRow = intRepairRows, intBold, intFine), RGB_BLACK)
                        End If
                    End If
                Next intCOl
                lngCurY = lngY1
            Next intRow
        End If
        
        lngOutY = lngCurY + 5
    Else
        lngOutY = lngCurY + 5
    End If
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub DrawBodyPageFooter(ByVal lngDC As Long, objDraw As Object, X As Long, Y As Long, ByVal LeftX As Long, ByVal intPageNo As Integer, _
    ByVal intBeginPage As Integer, Optional ByVal strInfo As String, Optional ByVal sngScale As Single = 1)
    '--------------------------------------------------------------------------------------------------------------------------------
    '���ܣ�������ײ�˵��
    '����:intPageNO=ҳ��
    '--------------------------------------------------------------------------------------------------------------------------------
    Dim blnWeek As Boolean
    Dim blnPageNo As Boolean
    Dim blnPrintCurveInfo As Boolean
    Dim strNOPage As String
    Dim lngX As Long
    Dim blnPrinter As Boolean
    
    blnPrinter = False
    If TypeName(objDraw) = "Printer" Then
        blnPrinter = True
    Else
        msngTwips = 1
    End If
    blnPrintCurveInfo = (Val(zldatabase.GetPara("���µ�����ӡ����˵��", glngSys, 1255, "0")) = 1)
    If blnPrintCurveInfo = False Then
        '��ӡ����˵����Ϣ
        Call SetTextColor(lngDC, RGB_BLACK)
        Call GetTextExtentPoint32(lngDC, strInfo, Len(strInfo), T_Size)
        Call GetTextRect(objDraw, X, Y, strInfo, 0, False, , sngScale)
        Call DrawText(lngDC, strInfo, -1, T_LableRect, DT_CENTER)
        Y = Y + IIf(blnPrinter = True, msngTwips, 1) * 30
    Else
        Y = Y + IIf(blnPrinter = True, msngTwips, 1) * 10
    End If
    
    blnWeek = (Val(zldatabase.GetPara("��ӡ����", glngSys, 1255, "0")) = 1)
    blnPageNo = (Val(zldatabase.GetPara("��ӡҳ��", glngSys, 1255, "1")) = 1)
    
    
    '��ӡҳ��
    '------------------------------------------------------------------------------------------------------------------
    If intPageNo > -1 And blnPageNo Then
        intPageNo = intPageNo + intBeginPage - 1
        strNOPage = "��   --" & CStr(intPageNo) & "--   ҳ"
    End If
    
    If blnWeek Then
        If strNOPage = "" Then
            strNOPage = "��   " & CStr(intBeginPage) & "   ��"
        Else
            strNOPage = strNOPage & "(�� " & CStr(intBeginPage) & " ��)"
        End If
    End If
    
    Call SetTextColor(lngDC, RGB_BLACK)
    Call GetTextExtentPoint32(lngDC, strNOPage, Len(strNOPage), T_Size)
    Call GetTextRect(objDraw, 0, Y, strNOPage, objDraw.Width / T_TwipsPerPixel.X, True, , sngScale)
    Call DrawText(lngDC, strNOPage, -1, T_LableRect, DT_CENTER)
    
    '�����ӡ��,����ǰ����Ա����
    '------------------------------------------------------------------------------------------------------------------
'    strNOPage = "��ӡ��:" & gstrUserName
'
'    Call SetTextColor(lngDc, RGB_BLACK)
'    Call GetTextExtentPoint32(lngDc, strNOPage, Len(strNOPage), T_Size)
'    Call GetTextRect(objDraw, LeftX - objDraw.TextWidth(strNOPage) / T_TwipsPerPixel.x, Y, strNOPage, 0, False, , sngScale)
'    Call DrawText(lngDc, strNOPage, -1, T_LableRect, DT_CENTER)

    Y = Y + T_Size.H / 2
    '--------------------------------------------------------------------------------------------------------------------------------
End Sub

Private Sub DrawTabText(ByVal lngDC As Long, ByVal objDraw As Object, ByVal strTmp As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long, LPoint As T_LPoint, Optional sngScale As Single = 1)
'---------------------------------------------------
'���� ������±���������
'---------------------------------------------------
    Dim lngFont As Long, lngOldFont As Long, intSize As Integer
    Dim stdset As StdFont
    Dim sngD As Single
    Dim blnChage As Boolean
    Dim arrText, blnGrade As Boolean
    
    On Error GoTo Errhand
    blnChage = False
    
    intSize = 9
    objDraw.Font.Size = intSize * sngScale
    If objDraw.TextWidth(strTmp) / T_TwipsPerPixel.X > LPoint.W Then
ErrGoTo:
        sngD = Round((objDraw.TextWidth(strTmp) / T_TwipsPerPixel.X - LPoint.W) / LPoint.W, 4)
        If sngD > 0 Then
            intSize = Int(Round((1 - sngD), 2) * intSize - 0.5)
            If intSize < 7 Then intSize = 7
            objDraw.Font.Size = intSize * sngScale
            If (objDraw.TextWidth(strTmp) / T_TwipsPerPixel.X) > LPoint.W And intSize > 7 Then GoTo ErrGoTo
            blnChage = True
        End If
    Else
        intSize = 9
    End If
    Set stdset = New StdFont
    stdset.Name = "����"
    stdset.Size = intSize * sngScale
    If stdset.Size < 9 Then
        stdset.Name = "Times New Roman"
    End If
    stdset.Bold = False
    Call SetFontIndirect(stdset, lngDC, objDraw)
    lngFont = CreateFontIndirect(T_Font)
    lngOldFont = SelectObject(lngDC, lngFont)
    If blnChage = True Then '���¼����������λ��
        Call GetTextRect(objDraw, LPoint.X, LPoint.Y, strTmp, LPoint.W, True, , sngScale)
    End If
    Call DrawText(lngDC, strTmp, -1, T_LableRect, DT_CENTER)
    
    Call SelectObject(lngDC, lngOldFont)
    Call DeleteObject(lngFont)
    
    '��ԭ��������
    objDraw.Font.Size = 9 * sngScale
    Set stdset = Nothing
    
    Exit Sub
Errhand:
    objDraw.Font.Size = 9 * sngScale
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub DrawAnsyGrade(ByVal lngDC As Long, ByVal objDraw As Object, arrText() As String, LPoint As T_LPoint, ByVal lngColor As Long, Optional ByVal blnFormat As Boolean = False, Optional sngScale As Single = 1)
'---------------------------------------------------
'���� ���������
'˵�� AnsyGrade=True���ܵ��ô˺���
'---------------------------------------------------
    Dim lngFont As Long, lngOldFont As Long, intSize As Integer
    Dim stdset As StdFont, stdOldset As StdFont
    Dim str1 As String, str2 As String, str3 As String, strTmp As String
    Dim lngX As Long, lngY As Long, sngH As Single, sngW As Single
    
    On Error GoTo Errhand
    
    If UBound(arrText) < 2 Then Exit Sub
    str1 = arrText(0): str2 = arrText(1): str3 = arrText(2)
    If blnFormat = True Then
        If Len(str2) > Len(str3) Then
            strTmp = str1 & str2
        Else
            strTmp = str1 & str3
        End If
    Else
        strTmp = str1 & str2 & "/" & str3
    End If
    intSize = 9
    objDraw.Font.Size = 9 * sngScale
    Set stdset = New StdFont
    stdset.Name = "����"
    stdset.Size = intSize * sngScale
    stdset.Bold = False
    Set stdOldset = stdset
    
    Call GetTextRect(objDraw, LPoint.X, LPoint.Y, strTmp, LPoint.W, True, , sngScale)
    '������
    If str1 <> "" Then
        Call SetFontIndirect(stdOldset, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        Call SetTextColor(lngDC, lngColor)
        Call DrawText(lngDC, str1, -1, T_LableRect, 0)
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
        lngX = T_LableRect.Left + (objDraw.TextWidth(str1) / T_TwipsPerPixel.X) - (objDraw.TextWidth("a") / T_TwipsPerPixel.X / 2) + msngTwips
    Else
        lngX = T_LableRect.Left
    End If
    
    If blnFormat = True Then '���ӷ�ĸ��ʾ
        intSize = 7
        objDraw.Font.Size = intSize * sngScale
        Set stdset = New StdFont
        stdset.Name = "����"
        stdset.Size = intSize * sngScale
        Call SetFontIndirect(stdset, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        Call SetTextColor(lngDC, lngColor)
        T_LableRect.Left = lngX
        lngY = T_LableRect.Top
        sngH = objDraw.TextHeight("A") / T_TwipsPerPixel.X / 2
        T_LableRect.Top = lngY - sngH
        T_LableRect.Bottom = LPoint.Y + (T_DrawClient.ʱ���е�λ / 2)
        Call DrawText(lngDC, str2, -1, T_LableRect, 0)
        lngY = T_LableRect.Top + (objDraw.TextHeight("A") / T_TwipsPerPixel.Y)
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
        '������
        objDraw.Font.Size = 9 * sngScale
        Call DrawLine(lngDC, lngX, lngY, lngX + (objDraw.TextWidth("A") / T_TwipsPerPixel.X), lngY)
        '�����ĸ
        lngY = lngY
        T_LableRect.Left = lngX
        T_LableRect.Top = lngY
        intSize = 7.5
        objDraw.Font.Size = intSize * sngScale
        Set stdset = New StdFont
        stdset.Name = "����"
        stdset.Size = intSize * sngScale
        Call SetFontIndirect(stdset, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        Call SetTextColor(lngDC, lngColor)
        Call DrawText(lngDC, str3, -1, T_LableRect, 0)
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
    Else
        If str1 <> "" Then
            '����ϱ�
            intSize = 7
            objDraw.Font.Size = intSize * sngScale
            Set stdset = New StdFont
            stdset.Name = "����"
            stdset.Size = intSize * sngScale
            Call SetFontIndirect(stdset, lngDC, objDraw)
            lngFont = CreateFontIndirect(T_Font)
            lngOldFont = SelectObject(lngDC, lngFont)
            Call SetTextColor(lngDC, lngColor)
            T_LableRect.Left = lngX
            lngY = T_LableRect.Top
            sngH = objDraw.TextHeight("A") / T_TwipsPerPixel.Y / 2
            T_LableRect.Top = lngY - sngH
            If T_LableRect.Top < (LPoint.Y - (T_DrawClient.ʱ���е�λ / 2)) Then T_LableRect.Top = (LPoint.Y - (T_DrawClient.ʱ���е�λ / 2)) - msngTwips
            Call DrawText(lngDC, str2, -1, T_LableRect, 0)
            Call SelectObject(lngDC, lngOldFont)
            Call DeleteObject(lngFont)
            lngX = lngX + (objDraw.TextWidth(str2) / T_TwipsPerPixel.X)
            '�����벿��
            Call SetFontIndirect(stdOldset, lngDC, objDraw)
            lngFont = CreateFontIndirect(T_Font)
            lngOldFont = SelectObject(lngDC, lngFont)
            Call SetTextColor(lngDC, lngColor)
            T_LableRect.Left = lngX
            T_LableRect.Top = lngY
            Call DrawText(lngDC, "/" & str3, -1, T_LableRect, 0)
            Call SelectObject(lngDC, lngOldFont)
            Call DeleteObject(lngFont)
        Else
            Call SetFontIndirect(stdOldset, lngDC, objDraw)
            lngFont = CreateFontIndirect(T_Font)
            lngOldFont = SelectObject(lngDC, lngFont)
            Call SetTextColor(lngDC, lngColor)
            Call DrawText(lngDC, str2 & "/" & str3, -1, T_LableRect, DT_CENTER)
            Call SelectObject(lngDC, lngOldFont)
            Call DeleteObject(lngFont)
        End If
    End If
    
    objDraw.Font.Size = 9 * sngScale
    Set stdset = Nothing
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function AnsyGrade(ByVal lngItemNO As Long, ByVal strText As String, arrText() As String) As Boolean
    '����������1/E �� 1+1/E�ĸ�ʽ ��1 1/E�ĸ�ʽ
    Dim intPos As Integer
    Dim ArrCode
    Dim str1 As String, str2 As String, str3 As String
    
    strText = Trim(strText)
    If strText = "" Or lngItemNO <> gint��� Then Exit Function
    If InStr(strText, "/E") = 0 Then Exit Function
    intPos = InStr(strText, "+")
    If intPos > 0 Then
        If intPos = 1 Then Exit Function
        str1 = Trim(Mid(strText, 1, intPos - 1))
        strText = Trim(Mid(strText, intPos + 1))
    Else
        intPos = InStr(strText, " ")
        If intPos > 0 Then
            If intPos = 1 Then Exit Function
            str1 = Trim(Mid(strText, 1, intPos - 1))
            strText = Trim(Mid(strText, intPos + 1))
        End If
    End If
    
    intPos = InStr(strText, "/E")
    If intPos > 0 Then
        If intPos = 1 Then Exit Function
        If intPos = Len(strText) Then Exit Function
        
        str2 = Trim(Mid(strText, 1, intPos - 1))
        str3 = Trim(Mid(strText, intPos + 1))
        If Len(str3) > 1 Then Exit Function
    End If
    
    ReDim arrText(0 To 2)
    arrText(0) = str1: arrText(1) = str2: arrText(2) = str3
    AnsyGrade = True
End Function


Private Function ShowPoints(ByVal lngDC As Long, ByVal objDraw As Object, ByVal rsPoint As ADODB.Recordset, _
    strEditors() As Variant, Optional int�������� As Integer = 1, Optional ByVal sngScale As Single = 1) As String
'-------------------------------------------------------------------------------------
'����:���������Ŀ�����ߺ�ͼ�����
'����::lngDC ��ͼ�����DC��objDraw �滭����.rsPoint ������Ŀ��ļ���(���|��ֵ|��λ|���|ʱ��|��Ŀ���|����|�Ͽ�|�ص���Ŀ|�ص�|X����|Y����|��ע|����)
'strEditors ���£����ʣ���������������Ϣ(��Ŀ���||��Ŀ����||��Ŀ��λ||��Ŀֵ��||��¼��||��¼ɫ)
'����:���ʵ�ļ��� !X���� & ";" & !Y���� & "," & !X���� & ";" & !Y����
'-------------------------------------------------------------------------------------
    Dim sinԭX As Single, sinԭY As Single
    Dim lng��Ŀ��� As Long
    Dim SinX As Single, sinY As Single  '������ʹ��
    Dim dblvalue As Double
    Dim dblMaxValue As Double, dblMinValue As Double
    Dim lngRGB As Long
    Dim strChar As String, str��λ As String, strTmp As String, strPic As String
    Dim str���� As String
    Dim lngCount As Long '�ص���Ŀ����
    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset
    Dim blnLine As Boolean
    Dim i As Integer
    Dim X1 As Single
    Dim blnPrinter As Boolean
    Dim intBold As Integer, intFine As Integer
    Dim bln�������� As Boolean
    On Error GoTo Errhand
    

    
    blnPrinter = False
    If TypeName(objDraw) = "Printer" Then
        blnPrinter = True
    Else
        msngTwips = 1
    End If
    
    If blnPrinter = True Then
        intBold = 4
        intFine = 4
    Else
        intBold = 2
        intFine = 1
    End If
    rsPoint.Filter = ""
    rsPoint.Sort = "��Ŀ���,ʱ��"
    '���Ƚ�������
    With rsPoint
        Do While Not .EOF
            For i = 0 To UBound(strEditors)
                If Val(Split(strEditors(i), "||")(0)) = Val(zlCommFun.Nvl(!��Ŀ���)) Then
                     Exit For
                End If
            Next i
            If Not (zlCommFun.Nvl(!��Ŀ���) = gint���� And Val(zlCommFun.Nvl(!���)) = 1) Then
                If zlCommFun.Nvl(!��Ŀ���) <> lng��Ŀ��� Then
                    sinԭX = 0
                    sinԭY = 0
                    lngRGB = Split(CStr(strEditors(i)), "||")(5)
                    lng��Ŀ��� = zlCommFun.Nvl(!��Ŀ���)
                End If
                If int�������� = 2 Then
                    If !��Ŀ��� = -1 Then
                        blnLine = True
                    Else
                        blnLine = True
                    End If
                Else
                    blnLine = True
                End If
                
                If sinԭX <> 0 And blnLine Then
                    Call DrawLine(lngDC, sinԭX + T_DrawClient.�е�λ / 2, sinԭY, !X���� + T_DrawClient.�е�λ / 2, !Y����, PS_SOLID, intFine, lngRGB)
                End If
                If !�Ͽ� = 0 Then
                    sinԭX = zlCommFun.Nvl(!X����, 0)
                    sinԭY = zlCommFun.Nvl(!Y����, 0)
                Else
                    sinԭX = 0
                End If
                
                If !��Ŀ��� = gint���� Then
                    If zlCommFun.Nvl(!����) = 1 Then '���Ժϸ�
                        Call SetTextColor(lngDC, lngRGB)
                        Call GetTextRect(objDraw, !X����, !Y���� - T_DrawClient.�е�λ, "v", T_DrawClient.�е�λ, True, , sngScale)
                        Call DrawText(lngDC, "v", -1, T_LableRect, DT_CENTER)
                    End If
                End If
                
                If i <= UBound(strEditors) Then
'                    If InStr(1, Split(strEditors(i), "||")(3), ";") <> 0 Then
'                        dblMinValue = Val(Split(Split(strEditors(i), "||")(3), ";")(0))
'                        dblMaxValue = Val(Split(Split(strEditors(i), "||")(3), ";")(1))
'                        If dblMaxValue = 0 Then dblMaxValue = Split(strEditors(i), "||")(6)
'                    Else
'                        dblMaxValue = Val(Split(strEditors(i), "||")(6))
'                        dblMinValue = Val(Split(strEditors(i), "||")(7))
'                    End If
                    dblMaxValue = Val(Split(strEditors(i), "||")(6))
                    dblMinValue = Val(Split(strEditors(i), "||")(7))
                End If
                
                '�ٽ�ֵ���ȿ�,���������ֵ����Сֵ֮��
                If Split(strEditors(i), "||")(8) <> "" And Val(Split(strEditors(i), "||")(8)) <= Val(Split(strEditors(i), "||")(6)) _
                    And Val(Split(strEditors(i), "||")(8)) >= Val(Split(strEditors(i), "||")(7)) Then dblMaxValue = Val(Split(strEditors(i), "||")(8))
                    
                If zlCommFun.Nvl(!��Ŀ���) = gint���� And Trim(zlCommFun.Nvl(!��ֵ)) = "����" Then
                    dblvalue = dblMinValue
                Else
                    dblvalue = Val(zlCommFun.Nvl(!��ֵ))
                End If
                
                If dblvalue > dblMaxValue Then
                    Call DrawLine(lngDC, !X���� + T_DrawClient.�е�λ / 2, !Y���� - T_DrawClient.�е�λ * 2, !X���� + T_DrawClient.�е�λ / 2, !Y����, PS_SOLID, intFine, lngRGB, True)
                ElseIf dblvalue < dblMinValue Then
                    Call DrawLine(lngDC, !X���� + T_DrawClient.�е�λ / 2, !Y���� + T_DrawClient.�е�λ * 2, !X���� + T_DrawClient.�е�λ / 2, !Y����, PS_SOLID, intFine, lngRGB, True)
                End If
            Else
                '���µ�������
                dblvalue = Split(!��ע, ",")(0)
                SinX = Val(Split(Split(!��ע, ",")(1), ";")(0))
                sinY = Val(Split(Split(!��ע, ",")(1), ";")(1))
                T_Size.H = objDraw.TextHeight("��") / T_TwipsPerPixel.Y

                If Val(!��ֵ) > Val(dblvalue) Then
                    '������ʧ�ܣ�������ͷ�ĺ�ɫʵ�ߣ��ַ��̶��á�
                    'Call DrawLine(lngDC, !X���� + T_DrawClient.�е�λ / 2, !Y����, SinX + T_DrawClient.�е�λ / 2, sinY, PS_SOLID, intFine, RGB_RED, True)
                    '����ʧ��ҲΪ����(ҽԺҪ��)
                    Call DrawLine(lngDC, !X���� + T_DrawClient.�е�λ / 2, !Y���� + (T_Size.H / 4), SinX + T_DrawClient.�е�λ / 2, sinY, IIf(blnPrinter, PS_DASH, PS_DOT), 1, RGB_RED, True)
                ElseIf Val(!��ֵ) < Val(dblvalue) Then
                    '�����³ɹ�������ɫ���ߣ��ַ��̶��á�
                    Call DrawLine(lngDC, !X���� + T_DrawClient.�е�λ / 2, !Y���� - (T_Size.H / 2), SinX + T_DrawClient.�е�λ / 2, sinY, IIf(blnPrinter, PS_DASH, PS_DOT), 1, RGB_RED, False)
                End If
            End If
            .MoveNext
        Loop
    End With
    If rsPoint.RecordCount > 0 Then rsPoint.MoveFirst
    '������е��ͼ��
    With rsPoint
        Do While Not .EOF
            str��λ = ""
            strTmp = ""
            For i = 0 To UBound(strEditors)
                If Split(CStr(strEditors(i)), "||")(0) = Val(zlCommFun.Nvl(!��Ŀ���)) Then
                     Exit For
                End If
            Next i
            If zlCommFun.Nvl(!�ص�) = 0 And zlCommFun.Nvl(!�ص���Ŀ) = "��" Then 'δ�ص�����Ŀ
                lngRGB = Split(CStr(strEditors(i)), "||")(5)
                If zlCommFun.Nvl(!��Ŀ���) = -1 And int�������� = 2 Then lngRGB = RGB_RED
                str��λ = zlCommFun.Nvl(!��λ)
                If str��λ = "" Then
                    Select Case lng��Ŀ���
                        Case gint����
                            str��λ = "Ҹ��"
                        Case gint����
                            str��λ = "��������"
                        Case Else
                            str��λ = ""
                    End Select
                End If
                strTmp = Split(CStr(strEditors(i)), "||")(4)
                strPic = ""
                strChar = ""
                Select Case zlCommFun.Nvl(!��Ŀ���)
                    Case gint����
                        strTmp = strTmp & String(2 - UBound(Split(strTmp, ",")), ",")
                        If str��λ = "����" Then
                            strChar = Split(strTmp, ",")(0)
                        ElseIf str��λ = "Ҹ��" Then
                            strChar = Split(strTmp, ",")(1)
                        Else
                            strChar = Split(strTmp, ",")(2)
                        End If
                        If zlCommFun.Nvl(!���) = 1 Then '�����·���
                            lngRGB = RGB_RED
                            strChar = "��"
                        Else
                            If strChar = "" Then strChar = "��"
                        End If
                    Case gint����
                        strChar = IIf(strTmp = "", "��", strTmp)
                    Case gint����
                        If str��λ = "����" Then
                            strPic = "PACEMAKER"
                        Else
                            strChar = IIf(strTmp = "", "+", strTmp)
                        End If
                    Case gint����
                        If str��λ = "��������" Then
                            strChar = IIf(strTmp = "", "*", strTmp)
                        Else
                            strPic = "BREATH"
                        End If
                    Case Else
                        strChar = strTmp
                End Select
                If Trim(zlCommFun.Nvl(!����)) <> "" Then
                    strChar = Trim(zlCommFun.Nvl(!����))
                    strPic = ""
                End If
                
                If !��Ŀ��� = gint���� And Trim(Nvl(!��ֵ)) = "����" And (mlng���²�����ʾ��ʽ = 0 Or mlng���²�����ʾ��ʽ = 1) Then
                    bln�������� = False
                Else
                    bln�������� = True
                End If
                                
                If strPic = "" And bln�������� Then
                    Call SetTextColor(lngDC, lngRGB)
                    Call GetTextRect(objDraw, !X����, !Y����, Trim(strChar), T_DrawClient.�е�λ, True, , sngScale)
                    Call DrawText(lngDC, Trim(strChar), -1, T_LableRect, DT_CENTER)
                    'Debug.Print T_LableRect.Left & ";" & T_LableRect.Right
                Else
                    Call DrawPicture(objDraw, strPic, objDraw.ScaleX(!X���� + ((T_DrawClient.�е�λ - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2), vbPixels, vbTwips), objDraw.ScaleY(!Y���� - mintBmpH * IIf(blnPrinter = True, msngTwips, 1) / 2, vbPixels, vbTwips), _
                        objDraw.ScaleX(!X���� + ((T_DrawClient.�е�λ - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2) + mintBmpW * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), objDraw.ScaleY(!Y���� + mintBmpH * IIf(blnPrinter = True, msngTwips, 1) / 2, vbPixels, vbTwips), True)
                End If
            
            Else  'չʾ�ص���λͼ��
                strPic = ""
                strChar = ""
                If zlCommFun.Nvl(!�ص���Ŀ) <> "��" Then '�ص�=1�Ĳ����κδ���
                    lngCount = UBound(Split(zlCommFun.Nvl(!�ص���Ŀ), ","))
                    strTmp = zlCommFun.Nvl(!�ص���Ŀ)
                    If Trim(strTmp) <> "" Then
                        str��λ = zlCommFun.Nvl(!��λ)
                        lngCount = lngCount + 2
                        strTmp = zlCommFun.Nvl(!��Ŀ���) & "," & strTmp
                        If InStr(1, "," & strTmp & ",", ",1,") <> 0 Then

                            strSql = "SELECT A.���,A.��Ƿ���,A.�����ɫ" & vbNewLine & _
                                    " FROM �����ص���� A," & vbNewLine & _
                                    "     (SELECT �ϼ����, COUNT(*) ����" & vbNewLine & _
                                    "     FROM �����ص����" & vbNewLine & _
                                    "     WHERE ��Ŀ��� IN (" & strTmp & ")" & vbNewLine & _
                                    "     GROUP BY �ϼ����) B" & vbNewLine & _
                                    " WHERE A.�ص���Ŀ = B.����" & vbNewLine & _
                                    " AND A.��� = B.�ϼ���� AND A.���=[1]"
                        Else
                            strSql = "Select A.���, A.��Ƿ���, A.�����ɫ" & vbNewLine & _
                                "  From �����ص���� A," & vbNewLine & _
                                "       (Select �ϼ����, Count(1) ����" & vbNewLine & _
                                "          from �����ص����" & vbNewLine & _
                                "         where ��Ŀ��� in (" & strTmp & ")" & vbNewLine & _
                                "         group by �ϼ����) B" & vbNewLine & _
                                " Where A.�ص���Ŀ = B.����" & vbNewLine & _
                                "   And A.��� = B.�ϼ����"
                        End If
                        
                        Set rsTmp = zldatabase.OpenSQLRecord(strSql, "�ص�", 1)
                        
                        If rsTmp.RecordCount > 0 Then
                            If IsNull(rsTmp!��Ƿ���) Then
                                strPic = zlBlobRead(9, zlCommFun.Nvl(rsTmp!���))
                            Else
                                strChar = Trim(zlCommFun.Nvl(rsTmp!��Ƿ���))
                                lngRGB = Val(zlCommFun.Nvl(rsTmp!�����ɫ, 0))
                            End If
                            If strPic = "" Then
                                Call SetTextColor(lngDC, lngRGB)
                                Call GetTextRect(objDraw, !X���� - 1, !Y����, Trim(strChar), T_DrawClient.�е�λ, True, , sngScale)
                                Call DrawText(lngDC, Trim(strChar), -1, T_LableRect, DT_CENTER)
                            Else
                                Call DrawPicture(objDraw, strPic, objDraw.ScaleX(!X���� + ((T_DrawClient.�е�λ - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2), vbPixels, vbTwips), objDraw.ScaleY(!Y���� - mintBmpH * IIf(blnPrinter = True, msngTwips, 1) / 2, vbPixels, vbTwips), _
                                    objDraw.ScaleX(!X���� + ((T_DrawClient.�е�λ - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2) + mintBmpW * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), objDraw.ScaleY(!Y���� + mintBmpH * IIf(blnPrinter = True, msngTwips, 1) / 2, vbPixels, vbTwips), False)
                                
                                Call FileSystem.Kill(strPic)
                            End If
                        End If
                    End If
                End If
            End If
        .MoveNext
        Loop
    End With
    
    '��ȡ�������ʵ���Ϣ
    If rsPoint.RecordCount > 0 Then rsPoint.MoveFirst
    rsPoint.Filter = "��Ŀ���=" & gint����
    With rsPoint
        Do While Not .EOF
            str���� = str���� & "," & !X���� & ";" & !Y����
        .MoveNext
        Loop
    End With
    If str���� <> "" Then str���� = Mid(str����, 2)
    
    ShowPoints = str����
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function GetCanvasCenter(ByVal dtBegin As Date, ByVal dtEnd As Date, ByVal dtBeginDate As Date, ByVal SinX As Single) As Boolean
'---------------------------------------------------------
'����:�жϸ�ʱ����Ƿ����м�ֵ
'����:dtbegin:���Ƚϵ�ʱ���.  dtend:Ҫ�Ƚϵ�ʱ��� . dtBeginDate ��ҳ���µ��Ŀ�ʼʱ�� .sinx��ǰ���X����
'---------------------------------------------------------
    Dim blnTrue As Boolean
    Dim strTime As String, strTmp As String
    Dim intDay As Integer, intTime As Integer, strDay As String
    
    
    intTime = (SinX - T_DrawClient.��������.Left) \ T_DrawClient.�е�λ
    intDay = intTime \ 6
    intTime = intTime Mod 6
        
    strDay = Format(DateAdd("d", intDay, dtBeginDate), "yyyy-MM-dd")
    strTmp = strDay & " " & Split(gvarTime(intTime), ",")(0) & "," & strDay & " " & Split(gvarTime(intTime), ",")(1)
    
    If intTime <= UBound(gvarTime) Then
        If gintHourBegin + intTime * 4 = 24 Then
            strTime = Format(Format(strDay, "YYYY-MM-DD") & " " & "23:59:59", "YYYY-MM-DD HH:mm:ss")
        Else
            strTime = Format(Format(strDay, "YYYY-MM-DD") & " " & gintHourBegin + intTime * 4 & ":00:00", "YYYY-MM-DD HH:mm:ss")
        End If
    End If
    
    If CDate(strTime) > CDate(Split(strTmp, ",")(1)) Then strTime = Format(Split(strTmp, ",")(1), "YYYY-MM-DD HH:mm:ss")
    
    If Abs(DateDiff("s", Format(dtBegin, "YYYY-MM-DD HH:mm:ss"), Format(strTime, "YYYY-MM-DD HH:mm:ss"))) > _
        Abs(DateDiff("s", Format(dtEnd, "YYYY-MM-DD HH:mm:ss"), Format(strTime, "YYYY-MM-DD HH:mm:ss"))) Then
        blnTrue = True
    End If

    GetCanvasCenter = blnTrue
End Function

Public Function DrawCanvas(ByVal lngDC As Long, ByVal objDraw As Object, ByVal rsTemp As ADODB.Recordset, rsDrawItems As ADODB.Recordset, Optional ByVal bln����ӡ������ As Boolean = False, Optional sngScale As Single = 1) As String
'------------------------------------------------------------------------------------------------------
'����:���̶������������������̶�ֵ��Ϣ
'����:lngDC ��ͼ�����DC��objDraw �滭����.rsTemp:����������Ŀ��¼��(A.��Ŀ���,A.�������,A.��¼��,A.��¼��,A.��¼ɫ,A.���ֵ,A.��Сֵ,A.��λֵ,C.��Ŀ��λ ��λ,A.�����-2 AS �����,B.��λ)
'����:���ظ������ߵľ�����Ϣ����( "��Ŀ���|���ֵ|��Сֵ|��λֵ|���ֵ����|��Сֵ����|��λ�̶�|��ʾģʽ|��ɫ")
'����˵����Ϣ(��Ŀ�ķ���)
'-------------------------------------------------------------------------------------------------------
    Dim str˵�� As String
    Static SlngMaxY As Long                 '��¼��һ�ε����߶ȣ��Ծ��������Ƿ���Ҫ�ػ�
    Dim lngCurX     As Long, lngCurY As Single  '��ǰλ��
    Dim lngMaxX     As Long, lngMaxY As Single  '�߽�
    Dim lngCurAlerY As Single '������
    Dim lngRow      As Long
    Dim intLables   As Integer
    Dim bln˫�� As Boolean                  '�˲������û�ָ��,bln˫��=TRUE��ʾֻ��ʾ����;������ʾʮ��
    Dim bln���� As Boolean                  '�˲������û�ָ��,���зֽ��Ǵ��߻���ϸ��
    '���¶��Ǳ�׼�߶�
    Dim intLineMode   As Integer
    Dim blnDoubleRow  As Boolean             '������Ϊһ�д�ӡ���
    Dim sinAlertness  As Single              '������,��������
    Dim lngLableStep  As Long
    Dim lngColStep    As Long
    Dim sinRowStep As Single, lngInitRowStep As Long
    Dim arrTemp()     As String
    Dim blnPrinter As Boolean
    Dim intBold As Integer, intFine As Integer
    Dim lngFont As Long, lngOldFont As Long
    Dim sinY��λ As Single '���ߵ�λ�����Bottom
    
    '�������ͼ�������(��Ŀ���,���ֵ,��Сֵ,��λֵ,���ֵ����,��Сֵ����,��λ�̶�,��ʾģʽ)
    Dim sin�̶� As Single, bln��ʾ�̶� As Boolean
    Dim sin�̶ȼ�� As Single, sinBegin�̶� As Single, dbl��λֵ As Double
    
    Dim str���ֵ���� As String, str��Сֵ���� As String

    On Error GoTo Errhand
    If TypeName(objDraw) = "Printer" Then
        blnPrinter = True
    Else
        blnPrinter = False
    End If
    
    If blnPrinter = True Then
        intBold = 6
        intFine = 2
    Else
        intBold = 2
        intFine = 1
    End If
    '����������Ŀ����ͼ����(��Ŀ���,���ֵ,��Сֵ,��λֵ,���ֵ����,��Сֵ����,��λ�̶�,��ʾģʽ)
    gstrFields = "��Ŀ���," & adDouble & ",18|���ֵ," & adDouble & ",18|��Сֵ," & adDouble & ",18|" & "��λֵ," & adDouble & _
        ",18|���ֵ����," & adLongVarChar & ",20|��Сֵ����," & adLongVarChar & ",20|" & "��λ�̶�," & adLongVarChar & ",20|��ʾģʽ," & adDouble & ",5|��ɫ," & adDouble & ",18"
    Call Record_Init(rsDrawItems, gstrFields)
    '------------------------------------------------------------------------------------------------------------------
    '����ֵ
    intLineMode = PS_SOLID
    lngLableStep = T_DrawClient.�̶ȵ�λ
    lngColStep = T_DrawClient.�е�λ
    lngInitRowStep = glngInitRowStep * IIf(blnPrinter = True, msngTwips, 1)
    sinRowStep = T_DrawClient.�е�λ
    
    '���µ��Ե�����ʾ(������ѡ����˫����ʾ��û�����̶���ʾһ��) 1��������ʾ 0��˫����ʾ
    If zldatabase.GetPara("���µ���ʾ��ʽ", glngSys, 1255, 0) = 1 Then
        bln˫�� = False
    Else
        bln˫�� = True
    End If
    'True��ʾ����ֻ���һ��,Ч����һ���̶�ֻ��ʾ������;����һ���̶���ʾʮ��,���û�������������,��blnDoubleRow�޹�
    bln���� = True
    
    If Not bln���� Then intLineMode = PS_DASHDOTDOT
    
    '�����
    rsTemp.Filter = "��Ŀ���=" & gint����
    If rsTemp.RecordCount > 0 And bln����ӡ������ = True Then
        rsTemp.Filter = 0
        intLables = rsTemp.RecordCount - 1
    Else
        rsTemp.Filter = 0
        intLables = rsTemp.RecordCount
    End If
    If intLables <= 0 Then intLables = 1
    lngCurX = T_DrawClient.ƫ����X
    lngCurY = T_DrawClient.ƫ����Y
    lngMaxX = (intLables * lngLableStep) + (7 * 6 * lngColStep) + T_DrawClient.ƫ����X  '�̶�+7*���+ƫ����X
    lngMaxY = 2 * mintNullRow * lngInitRowStep + T_DrawClient.������ * sinRowStep + T_DrawClient.ƫ����Y '��Ϊ����С�����������ʼY���꣩
       
    str˵�� = ""
        
    SlngMaxY = lngMaxY
    T_DrawClient.�̶ȵ�λ = lngLableStep
    T_DrawClient.�е�λ = sinRowStep
    T_DrawClient.�е�λ = lngColStep
    T_DrawClient.˫�� = blnDoubleRow
    
    For lngRow = 1 To intLables
        'Call DrawRect(lngDc, lngCurX - IIf(lngRow = 1, 0, 1), lngCurY, lngCurX + lngLableStep + 1, lngMaxY, PS_SOLID, IIf(lngRow = 1, 2, IIf(lngRow = intLables, 2, 1)), RGB_BLACK)
        Call DrawLine(lngDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, IIf(lngRow = 1, intBold, intFine), RGB_BLACK)
        lngCurX = lngCurX + lngLableStep
        Call DrawLine(lngDC, lngCurX - lngLableStep, lngCurY, lngCurX, lngCurY, PS_SOLID, intBold, RGB_BLACK)
        Call DrawLine(lngDC, lngCurX - lngLableStep, lngMaxY, lngCurX, lngMaxY, PS_SOLID, intBold, RGB_BLACK)
        If lngRow = intLables Then
            Call DrawLine(lngDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, intBold, RGB_BLACK)
        End If
    Next
    
    
    T_DrawClient.�̶�����.Left = T_DrawClient.ƫ����X
    T_DrawClient.�̶�����.Top = lngCurY
    T_DrawClient.�̶�����.Right = lngCurX
    T_DrawClient.�̶�����.Bottom = lngMaxY
    
    'Ĭ�����һ��������ʾ��Ŀ����
    lngCurY = lngCurY + lngInitRowStep * 2
    Call DrawLine(lngDC, T_DrawClient.ƫ����X, lngCurY, lngMaxX, lngCurY, PS_SOLID, intFine, RGB_BLACK)
    lngCurY = lngCurY + lngInitRowStep * ((mintNullRow - 1) * 2)
    '�����µ�������
    For lngRow = 0 To T_DrawClient.������
        If lngRow <> 0 Then
            lngCurY = lngCurY + sinRowStep
        End If
        '�����µ���������
        If ((blnDoubleRow Or bln˫��) And lngRow Mod 2 = 0) Or (Not blnDoubleRow And Not bln˫��) Then
            Call DrawLine(lngDC, lngCurX, lngCurY, lngMaxX, lngCurY, IIf(lngRow Mod 10 = 0, PS_SOLID, intLineMode), IIf(lngRow Mod 5 = 0 And sinRowStep >= 4 And bln����, intBold, intFine), RGB_BLACK)
        End If
    Next
    
    lngCurY = T_DrawClient.�̶�����.Top
    
    '�����µ�������
    For lngRow = 1 To 6 * 7
        lngCurX = lngCurX + lngColStep
        Call DrawLine(lngDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, IIf(lngRow Mod 6 = 0, intBold, intFine), IIf(lngRow Mod 6 = 0, RGB_RED, RGB_BLACK))
    Next
    
    lngCurX = T_DrawClient.�̶�����.Right
    T_DrawClient.��������.Left = T_DrawClient.�̶�����.Right
    T_DrawClient.��������.Top = T_DrawClient.�̶�����.Top
    T_DrawClient.��������.Right = lngMaxX
    T_DrawClient.��������.Bottom = lngMaxY
    
    '���̶ȿ�ı�ߣ��ӹ̶������10�п�ʼ��ʶ��
    intLables = 1
    rsTemp.Sort = "�������"
    With rsTemp
        Do While Not .EOF
            If Not (bln����ӡ������ = True And !��Ŀ��� = gint����) Then
                '��ʾ�̶ȿ���Ŀ�����Ƽ�����,�����¡�
                lngCurX = T_DrawClient.�̶�����.Left + ((intLables - 1) * T_DrawClient.�̶ȵ�λ)
                lngCurY = T_DrawClient.�̶�����.Top
                 
                gstdSet.Name = "����"
                gstdSet.Size = 9 * sngScale
                Call SetFontIndirect(gstdSet, lngDC, objDraw)
                lngFont = CreateFontIndirect(T_Font)
                lngOldFont = SelectObject(lngDC, lngFont)
                '���������Ŀ������
                Call SetTextColor(lngDC, zlCommFun.Nvl(!��¼ɫ, RGB_BLACK))
                Call GetTextRect(objDraw, lngCurX, lngCurY + objDraw.TextHeight(zlCommFun.Nvl(!��¼��)) / T_TwipsPerPixel.Y / 2, Trim(zlCommFun.Nvl(!��¼��)), T_DrawClient.�̶ȵ�λ, , , sngScale)
                Call DrawText(lngDC, Trim(zlCommFun.Nvl(!��¼��)), -1, T_LableRect, DT_CENTER)
                Call SelectObject(lngDC, lngOldFont)
                Call DeleteObject(lngFont)
                
                '���������С
                gstdSet.Name = "����"
                gstdSet.Size = 8 * sngScale
                Call SetFontIndirect(gstdSet, lngDC, objDraw)
                lngFont = CreateFontIndirect(T_Font)
                lngOldFont = SelectObject(lngDC, lngFont)
    
                '�����Ŀ��λ
                Call GetTextRect(objDraw, lngCurX, lngCurY + lngInitRowStep * 2 + objDraw.TextHeight(zlCommFun.Nvl(!��λ)) / T_TwipsPerPixel.Y / 2, Trim(zlCommFun.Nvl(!��λ)), T_DrawClient.�̶ȵ�λ, , , sngScale)
                Call DrawText(lngDC, Trim(zlCommFun.Nvl(!��λ, 0)), -1, T_LableRect, DT_CENTER)
                Call SelectObject(lngDC, lngOldFont)
                Call DeleteObject(lngFont)
                sinY��λ = T_LableRect.Bottom
                intLables = intLables + 1
            End If
            objDraw.Font.Size = 9 * sngScale
            'ǿ���趨����������Ŀ����ʾģʽ
            Select Case !��Ŀ���

                Case gint����  '��������ʱ����̶�
                    sin�̶ȼ�� = zlCommFun.Nvl(!�̶ȼ��, 1)
                    dbl��λֵ = 0.1
                    sinAlertness = zlCommFun.Nvl(!��ʾ��, 37)
                    arrTemp = Split(zlCommFun.Nvl(!��¼��, "��,��,��"), ",")
                    str˵�� = str˵�� & "��" & zlCommFun.Nvl(!��¼��) & "(����" & arrTemp(0) & ",Ҹ��" & arrTemp(1) & ",����" & arrTemp(2) & ")"

                Case gint����, gint����  '����/������10�ı�������̶�
                    sin�̶ȼ�� = zlCommFun.Nvl(!�̶ȼ��, 10)
                    dbl��λֵ = 2
                    sinAlertness = zlCommFun.Nvl(!��ʾ��, 0)

                    If !��Ŀ��� = gint���� Then
                        str˵�� = str˵�� & "��" & zlCommFun.Nvl(!��¼��) & "(ȱʡ��¼��" & zlCommFun.Nvl(!��¼��, "+") & ",����H)"
                    Else
                        str˵�� = str˵�� & "��" & zlCommFun.Nvl(!��¼��) & "(" & zlCommFun.Nvl(!��¼��, "��") & ")"
                    End If

                Case gint����  '������5�ı�������̶�
                    mbln�������� = True
                    dbl��λֵ = 1
                    sin�̶ȼ�� = zlCommFun.Nvl(!�̶ȼ��, 5)
                    sinAlertness = zlCommFun.Nvl(!��ʾ��, 0)
                    str˵�� = str˵�� & "��" & zlCommFun.Nvl(!��¼��) & "(��������" & zlCommFun.Nvl(!��¼��, "*") & ",������R)"
                Case Else
                    dbl��λֵ = Val(zlCommFun.Nvl(!��λֵ, 0))
                    sin�̶ȼ�� = zlCommFun.Nvl(!�̶ȼ��, Val(zlCommFun.Nvl(!��λֵ, 0)) * 10)
                    If sin�̶ȼ�� > Val(zlCommFun.Nvl(!���ֵ)) - Val(zlCommFun.Nvl(!��Сֵ)) Then
                        sin�̶ȼ�� = Val(zlCommFun.Nvl(!���ֵ)) - Val(zlCommFun.Nvl(!��Сֵ))
                    End If
                    sinAlertness = zlCommFun.Nvl(!��ʾ��, 0)
                    str˵�� = str˵�� & "��" & zlCommFun.Nvl(!��¼��) & "(" & zlCommFun.Nvl(!��¼��, "*") & ")"
            End Select

            '����ֵ
            lngCurY = lngCurY + (lngInitRowStep * 2 * mintNullRow) '�̶�ǰ4�еĸ߶Ȳ�����̶�

            '��������ж�λ����Чλ��
            lngCurY = lngCurY + (T_DrawClient.�е�λ * zlCommFun.Nvl(!�����, 0))
            Do While True
                bln��ʾ�̶� = False
                If sin�̶� = 0 Then     '�ս���ѭ������ʱȡ�����ֵ
                    sin�̶� = zlCommFun.Nvl(!���ֵ, 0)
                    sinBegin�̶� = sin�̶�
                    str���ֵ���� = T_DrawClient.��������.Left & "," & lngCurY
                Else                    '����õ�ÿ���̶ȵ�ֵ
                    sin�̶� = sin�̶� - dbl��λֵ     '���Ŀǰ��ʾģʽΪ˫������˫���ۼ�
                End If
                
                '�������õĿ̶ȼ����ʾ�̶�ֵ
                If Val(Format(sin�̶�, "#0.00")) = Val(Format(sinBegin�̶�, "#0.00")) Then bln��ʾ�̶� = True
                If bln��ʾ�̶� = True Or sin�̶� < sinBegin�̶� Then sinBegin�̶� = sinBegin�̶� - IIf(T_DrawClient.˫��, sin�̶ȼ�� * 2, sin�̶ȼ��)
                If sinBegin�̶� < 0 Then sinBegin�̶� = 0
                
                If bln��ʾ�̶� And Not (bln����ӡ������ = True And !��Ŀ��� = gint����) Then
                    '�������ֵ�������ߵ�λ�ظ�
                    If sin�̶� = Val(Nvl(!���ֵ, 0)) And lngCurY < sinY��λ Then
                        Call GetTextRect(objDraw, lngCurX, sinY��λ, Format(sin�̶�, "#0"), T_DrawClient.�̶ȵ�λ, , , sngScale)
                    ElseIf Format(lngCurY, "#0") = T_DrawClient.�̶�����.Bottom Then
                        Call GetTextRect(objDraw, lngCurX, lngCurY - (objDraw.TextHeight("1") / 2 / T_TwipsPerPixel.Y), Format(sin�̶�, "#0"), T_DrawClient.�̶ȵ�λ, , , sngScale)
                    Else
                        Call GetTextRect(objDraw, lngCurX, lngCurY, Format(sin�̶�, "#0"), T_DrawClient.�̶ȵ�λ, , , sngScale)
                    End If
                    Call DrawText(lngDC, Format(sin�̶�, "#0"), -1, T_LableRect, DT_CENTER)
                End If
                '���������Ч��Χ�ڣ����߳����������˳�
                If Val(Format(sin�̶�, "#0.00")) <= Val(Format(zlCommFun.Nvl(!��Сֵ), "#0.00")) Or Format(lngCurY, "#0") > T_DrawClient.�̶�����.Bottom Then
                    str��Сֵ���� = T_DrawClient.��������.Left & "," & lngCurY
                    '��Ӹ���Ŀ(��Ŀ���,���ֵ,��Сֵ,��λֵ,���ֵ����,��Сֵ����,��λ�̶�,��ʾģʽ)
                    gstrFields = "��Ŀ���|���ֵ|��Сֵ|��λֵ|���ֵ����|��Сֵ����|��λ�̶�|��ʾģʽ|��ɫ"
                    gstrValues = zlCommFun.Nvl(!��Ŀ���) & "|" & zlCommFun.Nvl(!���ֵ, 0) & "|" & zlCommFun.Nvl(!��Сֵ, 0) & _
                    "|" & dbl��λֵ & "|" & str���ֵ���� & "|" & str��Сֵ���� & "|" & T_DrawClient.�е�λ & "," & T_DrawClient.�е�λ & "|" & sin�̶ȼ�� & "|" & !��¼ɫ
                    Call Record_Add(rsDrawItems, gstrFields, gstrValues)
                    
                    '�����߻�ʾ��
                    If blnDoubleRow = False And (sinAlertness < Val(Nvl(!���ֵ)) And sinAlertness > Val(Nvl(!��Сֵ))) Then
                        lngCurAlerY = Val(GetYCoordinate(objDraw, rsDrawItems, Val(Nvl(!��Ŀ���)), sinAlertness))
                        Call DrawLine(lngDC, T_DrawClient.��������.Left, lngCurAlerY, lngMaxX, lngCurAlerY, intLineMode, intBold, RGB_RED)
                    End If
                    
                    Exit Do
                End If
                lngCurY = lngCurY + T_DrawClient.�е�λ
            Loop
            sinBegin�̶� = 0
            sin�̶� = 0                 '���ƴӵ�һ�п�ʼ���
            .MoveNext
        Loop
    End With
    str˵�� = "˵��:" & Mid(str˵��, 2)
    
    DrawCanvas = str˵��
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Sub DrawPatiInfo(ByVal lngDC As Long, ByVal objDraw As Object, ByVal strPatiInfo As String, ByVal lngX As Long, ByVal lngY As Long, _
    ByVal lngLeft As Long, lngOutY As Long, Optional ByVal sngScale As Single = 1)
'-----------------------------------------------------------------------------------------------------------------------
'������˻�����Ϣ
'����:lngDC ��ͼ�����DC��strPatiInfo ������Ϣ����ַ���,�ָ���Ϊ'(����:'����:'�Ա�:'�Ʊ�:'����:'��Ժ����:'סԺ������)
'     lngX ��߾�,lngY�ϱ߾�,lngLeft �ұ߾�(���Ի�ͼ������ұ߾�)
'����:lngOutY ���ػ�ͼ����ϱ߾�
'-----------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer, k As Integer, l As Long
    Dim VarPatiInfo() As String
    Dim VarPatiName() As String
    Dim bln�Ƿ������� As Boolean, blnOne As Boolean
    Dim lngCurX As Long, lngCurY As Long, lngWidth As Long
    Dim strPatiName As String '������Ϣ���ݱ���,�� ����,�Ա�
    Dim Pname_SIZE() As SIZEL  '��¼ÿ����Ϣ���Ƶ�������Ϣ
    Dim Pinfo_SIZE() As SIZEL  '��¼ÿ����Ϣ��������Ϣ
    Dim arrSngY()
    Dim h_9t As Long
    Dim lngCurW As Long
    
    Dim sngW As Single, sngLen As Single
    Dim strText As String, strText1 As String, strText2 As String
    
    VarPatiInfo = Split(strPatiInfo, "'")
    bln�Ƿ������� = (UBound(VarPatiInfo) > 6)
    'strPatiName = "����:'�Ա�:'����:'��Ժ����:'סԺ��:'����:'����:" & IIf(bln�Ƿ������� = True, "'���:", "")
    strPatiName = "����:'����:'�Ա�:'�Ʊ�:'����:'��Ժ����:'סԺ������:" & IIf(bln�Ƿ������� = True, "'���:", "")
    VarPatiName = Split(strPatiName, "'")
    ReDim Preserve Pname_SIZE(UBound(VarPatiInfo))
    ReDim Preserve Pinfo_SIZE(UBound(VarPatiInfo))
    
    
    lngCurX = lngX: lngCurY = lngY
    
    lngWidth = IIf(lngLeft - lngCurX < 0, lngCurX - lngLeft, lngLeft - lngCurX)

    arrSngY = Array()
    
    ReDim Preserve arrSngY(UBound(arrSngY) + 1)
    arrSngY(UBound(arrSngY)) = lngCurY
    
    '��ʼ��������
    For i = 0 To UBound(VarPatiInfo)
        Call GetTextExtentPoint32(lngDC, VarPatiName(i), Len(VarPatiName(i)), T_Size) '��ȡ�����׼�߶ȣ���ȡ��������Ŀ�Ȳ�׼�����ֵ�׼
        If h_9t = 0 Then h_9t = T_Size.H
        T_Size.W = objDraw.TextWidth(VarPatiName(i)) / T_TwipsPerPixel.X
        lngCurX = lngCurX + T_Size.W
        Pname_SIZE(i).cx = lngCurX - T_Size.W
        Pname_SIZE(i).cy = lngCurY
        Call GetTextExtentPoint32(lngDC, VarPatiInfo(i), Len(VarPatiInfo(i)), T_Size)
        T_Size.W = objDraw.TextWidth(VarPatiInfo(i)) / T_TwipsPerPixel.X
        lngCurX = lngCurX + T_Size.W
        lngCurY = lngCurY
        Pinfo_SIZE(i).cx = lngCurX - T_Size.W
        Pinfo_SIZE(i).cy = lngCurY
        If lngCurX > lngLeft And i <> 0 Then
            If Not (UBound(VarPatiInfo) = 7 And (Pname_SIZE(i).cx - lngX) < lngWidth / 2) Then
                Call GetTextExtentPoint32(lngDC, VarPatiName(i), Len(VarPatiName(i)), T_Size)
                T_Size.W = objDraw.TextWidth(VarPatiName(i)) / T_TwipsPerPixel.X
                lngCurX = lngX + T_Size.W
                lngCurY = lngCurY + T_Size.H + 5
                Pname_SIZE(i).cx = lngCurX - T_Size.W
                Pname_SIZE(i).cy = lngCurY
                
                Pinfo_SIZE(i).cx = lngCurX
                Pinfo_SIZE(i).cy = Pname_SIZE(i).cy
                
                '��¼ÿ�λ���ǰ��Y������
                ReDim Preserve arrSngY(UBound(arrSngY) + 1)
                arrSngY(UBound(arrSngY)) = lngCurY
            End If
        End If
    Next i
    
    k = 0
    blnOne = False
    
    '�������������X����
    For j = 0 To UBound(arrSngY)
        For i = k To UBound(VarPatiInfo)
            If CDbl(arrSngY(j)) = CDbl(Pinfo_SIZE(i).cy) Then
                lngCurW = lngCurW + objDraw.TextWidth(VarPatiName(i)) / T_TwipsPerPixel.X
                lngCurW = lngCurW + objDraw.TextWidth(VarPatiInfo(i)) / T_TwipsPerPixel.X
                If i = UBound(VarPatiInfo) Then
                    If UBound(arrSngY) = 0 Then
                        blnOne = True: GoTo OneCoordinate
                    Else
                        If i <> k Then
                            sngW = (lngWidth - lngCurW) / (i - k)
                            If sngW < 0 Then sngW = 5
                            For l = k + 1 To i
                                Pname_SIZE(l).cx = Pname_SIZE(l).cx + sngW * (l - k)
                                Pinfo_SIZE(l).cx = Pinfo_SIZE(l).cx + sngW * (l - k)
                            Next l
                        End If
                        GoTo OutPutPatiInfo
                    End If
                End If
            Else
                If i <> 0 And (i - k - 1) <> 0 Then
                    sngW = (lngWidth - lngCurW) / (i - k - 1)
                    If sngW < 0 Then sngW = 5
                    For l = k + 1 To (i - 1)
                         Pname_SIZE(l).cx = Pname_SIZE(l).cx + sngW * (l - k)
                         Pinfo_SIZE(l).cx = Pinfo_SIZE(l).cx + sngW * (l - k)
                    Next l
                End If
                Exit For
            End If
        Next i
        If blnOne = False Then lngCurW = 0
        k = i
    Next j
    
OneCoordinate:
    If blnOne = True Then
        sngW = (lngWidth - lngCurW) / UBound(VarPatiInfo)
        If sngW < 0 Then sngW = 5
        For i = 1 To UBound(VarPatiInfo)
             Pname_SIZE(i).cx = Pname_SIZE(i).cx + sngW * i
             Pinfo_SIZE(i).cx = Pinfo_SIZE(i).cx + sngW * i
        Next i
    End If
    
    Dim lngLen As Long
OutPutPatiInfo:
    '�������������Ϣ
    For i = 0 To UBound(VarPatiInfo)
        Call SetTextColor(lngDC, RGB_BLACK)
        Call GetTextRect(objDraw, Val(Pname_SIZE(i).cx), Val(Pname_SIZE(i).cy), CStr(VarPatiName(i)), , , , sngScale)
        Call DrawText(lngDC, CStr(VarPatiName(i)), -1, T_LableRect, DT_CENTER)
        
        Call SetTextColor(lngDC, RGB_BLUE)
        
        '������������������һ����ʾʣ�ಿ��
        If UBound(VarPatiInfo) = 7 And i = UBound(VarPatiInfo) Then
            strText1 = ""
            strText = Replace(VarPatiInfo(i), vbCrLf, "")
            Do While True
                T_Size.W = objDraw.TextWidth(strText) / T_TwipsPerPixel.X
                strText2 = strText
                If T_Size.W + Val(Pinfo_SIZE(i).cx) - lngLeft > 0 Then
                    sngLen = Round((lngLeft - Val(Pinfo_SIZE(i).cx)) / T_Size.W, 2)
                    lngLen = Len(StrConv(strText, vbFromUnicode)) * sngLen
                    '�����תΪȫ��
                    strText = StrConv(strText, vbWide)
                    strText1 = StrConv(Mid(StrConv(strText, vbFromUnicode), lngLen + 1), vbUnicode)
                    strText = StrConv(Mid(StrConv(strText, vbFromUnicode), 1, lngLen), vbUnicode)
                    
                    '�õ�ԭʼ�ַ����Ľ�ȡ�ĳ���
                    strText1 = Mid(strText2, Len(strText) + 1)
                    strText = Mid(strText2, 1, Len(strText))
                    Call GetTextExtentPoint32(lngDC, strText, Len(strText), T_Size)
                    Call GetTextRect(objDraw, Val(Pinfo_SIZE(i).cx), Val(Pinfo_SIZE(i).cy), CStr(strText), , , , sngScale)
                    Call DrawText(lngDC, CStr(strText), -1, T_LableRect, DT_CENTER)
                    T_Size.W = objDraw.TextWidth(strText) / T_TwipsPerPixel.X
                    Pinfo_SIZE(i).cx = Pinfo_SIZE(i).cx + T_Size.W
                    strText = strText1
                    T_Size.W = objDraw.TextWidth("��") / T_TwipsPerPixel.X
                    If Val(Pinfo_SIZE(i).cx) + T_Size.W - lngLeft > 0 Then
                        Pinfo_SIZE(i).cx = lngX
                        Pinfo_SIZE(i).cy = Pinfo_SIZE(i).cy + T_Size.H + 5
                    End If
                    lngCurY = Pinfo_SIZE(i).cy
                    If strText = "" Then Exit Do
                Else
                    Call GetTextRect(objDraw, Val(Pinfo_SIZE(i).cx), Val(Pinfo_SIZE(i).cy), CStr(strText), , , , sngScale)
                    Call DrawText(lngDC, CStr(strText), -1, T_LableRect, DT_CENTER)
                    Exit Do
                End If
            Loop
        Else
            Call GetTextRect(objDraw, Val(Pinfo_SIZE(i).cx), Val(Pinfo_SIZE(i).cy), CStr(VarPatiInfo(i)), , , , sngScale)
            Call DrawText(lngDC, CStr(VarPatiInfo(i)), -1, T_LableRect, DT_CENTER)
        End If
    Next i
    Call SetTextColor(lngDC, RGB_BLACK)
    '����Y������
    lngOutY = lngCurY + h_9t
End Sub

Public Sub DrawUpTable(ByVal lngDC As Long, ByVal objDraw As Object, ByVal strTmpString As String, _
    ByVal lngX As Long, ByVal lngY As Long, ByVal lngLeft As Long, lngOutY As Long, Optional sngScale As Single)
'-----------------------------------------------------------------------------------------------------------------------
'���һ����Ŀ����Ϣ������ סԺ����,����,������������ʱ������
'����:lngDC ��ͼ�����DC��strTmpString ��סԺ���ڣ����� ������������ɵ��ַ���
'     lngX ��߾�,lngY�ϱ߾�,lngLeft �ұ߾�(���Ի�ͼ������ұ߾�)
'����:lngOutY ���ػ�ͼ����ϱ߾�
'-----------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim ArrCode() As String
    Dim strTmp As String
    Dim arrTmpTime() As String 'סԺʱ��
    Dim arrTmpDay() As String  'סԺ����
    Dim arrOptDay() As String '��������
    Dim lngCurX As Long, lngCurY As Long, lngStartY As Long, lngStartX As Long, lngTmpX As Long
    Dim lngColor As Long
    Dim strDay As String
    Dim intBold As Integer, intFine As Integer
    
    
    If TypeName(objDraw) = "Printer" Then
        intBold = 6
        intFine = 2
    Else
        intBold = 2
        intFine = 1
    End If
    
    strDay = IIf(mintBaby = 0, "סԺ����", "��������")
    
    ArrCode = Split(strTmpString, "||")
    strTmp = strTmpString & String(2 - UBound(ArrCode), "||")
    ArrCode = Split(strTmp, "||")
    arrOptDay = Split(ArrCode(2), "'")
    arrTmpTime = Split(ArrCode(0), "'")
    arrTmpDay = Split(ArrCode(1), "'")

    lngCurX = lngX: lngStartX = lngX
    lngCurY = lngY: lngStartY = lngY
    
    '��ʼ�������
    
    'X
    Call DrawLine(lngDC, lngCurX, lngCurY, lngLeft, lngCurY, PS_SOLID, intBold, RGB_BLACK): lngCurY = lngCurY + T_DrawClient.ʱ���е�λ
    Call DrawLine(lngDC, lngCurX, lngCurY, lngLeft, lngCurY, PS_SOLID, intFine, RGB_BLACK): lngCurY = lngCurY + T_DrawClient.ʱ���е�λ
    Call DrawLine(lngDC, lngCurX, lngCurY, lngLeft, lngCurY, PS_SOLID, intFine, RGB_BLACK): lngCurY = lngCurY + T_DrawClient.ʱ���е�λ
    Call DrawLine(lngDC, lngCurX, lngCurY, lngLeft, lngCurY, PS_SOLID, intFine, RGB_BLACK): lngCurY = lngCurY + T_DrawClient.ʱ���е�λ + 6
    Call DrawLine(lngDC, lngCurX, lngCurY, lngLeft, lngCurY, PS_SOLID, intBold, RGB_BLACK)
    
    'Y
    Call DrawLine(lngDC, lngCurX, lngStartY, lngCurX, lngCurY, PS_SOLID, intBold, RGB_BLACK)
    lngCurX = T_DrawClient.�̶�����.Right

    Call DrawLine(lngDC, lngCurX, lngStartY, lngCurX, lngCurY, PS_SOLID, intBold, RGB_BLACK)

    For i = 0 To 6
        lngCurX = lngCurX + T_DrawClient.�е�λ * 6
        Call DrawLine(lngDC, lngCurX, lngStartY, lngCurX, lngCurY, PS_SOLID, intBold, RGB_BLACK)
    Next i
    
    lngCurX = T_DrawClient.�̶�����.Right
    lngCurY = lngStartY + T_DrawClient.ʱ���е�λ * 3
    'ʱ��
    For i = 0 To 6
        lngCurX = T_DrawClient.�̶�����.Right + i * T_DrawClient.�е�λ * 6
        For j = 1 To 5
            lngCurX = lngCurX + T_DrawClient.�е�λ
            Call DrawLine(lngDC, lngCurX, lngCurY, lngCurX, lngCurY + T_DrawClient.ʱ���е�λ + 6, PS_SOLID, intFine, RGB_BLACK)
        Next j
    Next i
    
    '��ʼ�����Ϣ
    '������Ϣ
    lngCurY = lngStartY
    Call SetTextColor(lngDC, RGB_BLACK)
    Call GetTextExtentPoint32(lngDC, "��     ��", Len("��     ��"), T_Size)
    Call GetTextRect(objDraw, lngStartX, lngCurY + T_DrawClient.ʱ���е�λ / 2, "��      ��", T_DrawClient.�̶�����.Right - lngStartX, True, , sngScale)
    Call DrawText(lngDC, "��     ��", -1, T_LableRect, DT_CENTER)
    lngCurX = T_DrawClient.�̶�����.Right
    For i = 0 To UBound(arrTmpTime)
        lngCurX = T_DrawClient.�̶�����.Right + i * 6 * T_DrawClient.�е�λ
        Call SetTextColor(lngDC, RGB_BLUE)
        Call GetTextExtentPoint32(lngDC, CStr(arrTmpTime(i)), Len(CStr(arrTmpTime(i))), T_Size)
        Call GetTextRect(objDraw, lngCurX, lngCurY + T_DrawClient.ʱ���е�λ / 2, CStr(arrTmpTime(i)), T_DrawClient.�е�λ * 6, True, , sngScale)
        Call DrawText(lngDC, CStr(arrTmpTime(i)), -1, T_LableRect, DT_CENTER)
    Next i
    
    lngCurY = lngStartY + T_DrawClient.ʱ���е�λ * 1
    'סԺ����
    Call SetTextColor(lngDC, RGB_BLACK)
    Call GetTextExtentPoint32(lngDC, strDay, Len(strDay), T_Size)
    Call GetTextRect(objDraw, lngStartX, lngCurY + T_DrawClient.ʱ���е�λ / 2, strDay, T_DrawClient.�̶�����.Right - lngStartX, True, , sngScale)
    Call DrawText(lngDC, strDay, -1, T_LableRect, DT_CENTER)
    lngCurX = T_DrawClient.�̶�����.Right
    
    For i = 0 To UBound(arrTmpDay)
        lngCurX = T_DrawClient.�̶�����.Right + i * 6 * T_DrawClient.�е�λ
        Call SetTextColor(lngDC, RGB_BLUE)
        Call GetTextExtentPoint32(lngDC, CStr(arrTmpDay(i)), Len(CStr(arrTmpDay(i))), T_Size)
        Call GetTextRect(objDraw, lngCurX, lngCurY + T_DrawClient.ʱ���е�λ / 2, CStr(arrTmpDay(i)), T_DrawClient.�е�λ * 6, True, , sngScale)
        Call DrawText(lngDC, CStr(arrTmpDay(i)), -1, T_LableRect, DT_CENTER)
    Next i
    
    '��/�������
    lngCurY = lngStartY + T_DrawClient.ʱ���е�λ * 2
    Call SetTextColor(lngDC, RGB_BLACK)
    Call GetTextExtentPoint32(lngDC, "����������", Len("����������"), T_Size)
    Call GetTextRect(objDraw, lngStartX, lngCurY + T_DrawClient.ʱ���е�λ / 2, "����������", T_DrawClient.�̶�����.Right - lngStartX, True, , sngScale)
    Call DrawText(lngDC, "����������", -1, T_LableRect, DT_CENTER)
    lngCurX = T_DrawClient.�̶�����.Right
    
    '51283,������,2012-07-11,����������ɫ
    lngColor = Val(zldatabase.GetPara("����������ʾ��ɫ", glngSys, 1255, "255"))
    For i = 0 To UBound(arrOptDay)
        lngCurX = T_DrawClient.�̶�����.Right + i * 6 * T_DrawClient.�е�λ
        Call SetTextColor(lngDC, lngColor)
        Call GetTextExtentPoint32(lngDC, CStr(arrOptDay(i)), Len(CStr(arrOptDay(i))), T_Size)
        Call GetTextRect(objDraw, lngCurX, lngCurY + T_DrawClient.ʱ���е�λ / 2, CStr(arrOptDay(i)), T_DrawClient.�е�λ * 6, True, , sngScale)
        Call DrawText(lngDC, CStr(arrOptDay(i)), -1, T_LableRect, DT_CENTER)
    Next i
    lngColor = 0
    'ʱ��
    lngCurY = lngStartY + T_DrawClient.ʱ���е�λ * 3
    Call SetTextColor(lngDC, RGB_BLACK)
    Call GetTextExtentPoint32(lngDC, "ʱ      ��", Len("ʱ      ��"), T_Size)
    Call GetTextRect(objDraw, lngStartX, lngCurY + T_DrawClient.ʱ���е�λ / 2, "ʱ      ��", T_DrawClient.�̶�����.Right - lngStartX, True, , sngScale)
    Call DrawText(lngDC, "ʱ      ��", -1, T_LableRect, DT_CENTER)
    lngCurX = T_DrawClient.�̶�����.Right
    
    For i = 0 To 6
        lngCurX = T_DrawClient.�̶�����.Right + i * 6 * T_DrawClient.�е�λ
        '�����������ʱ��
        For j = 0 To 5
            strTmp = ""
            Select Case j
                Case 0
                    strTmp = gintHourBegin + 4 * 0
                    lngColor = RGB_RED
                Case 1
                    strTmp = gintHourBegin + 4 * 1
                    lngColor = RGB_RED
                Case 2
                    strTmp = gintHourBegin + 4 * 2
                    lngColor = RGB_BLUE
                Case 3
                    lngColor = RGB_BLUE
                    strTmp = gintHourBegin + 4 * 3
                Case 4
                    lngColor = RGB_BLUE
                    strTmp = gintHourBegin + 4 * 4
                Case 5
                    lngColor = RGB_RED
                    strTmp = gintHourBegin + 4 * 5
            End Select
            lngColor = GetTimeColor(Val(strTmp))
            lngTmpX = lngCurX + T_DrawClient.�е�λ * j
            Call SetTextColor(lngDC, lngColor)
            Call GetTextExtentPoint32(lngDC, strTmp, Len(strTmp), T_Size)
            Call GetTextRect(objDraw, lngTmpX - 1, lngCurY + (T_DrawClient.ʱ���е�λ + 6) / 2, strTmp, T_DrawClient.�е�λ, True, , sngScale)
            Call DrawText(lngDC, strTmp, -1, T_LableRect, DT_CENTER)
        Next j
    Next i
    lngOutY = lngStartY + T_DrawClient.ʱ���е�λ * 4 + 6
End Sub

Public Sub SetFontIndirect(ByVal stdset As StdFont, ByVal lngDC As Long, ByVal objDraw As Object)

    '����:������������
    Dim BFileName() As Byte

    Dim i As Integer

    On Error GoTo Errhand
    
    objDraw.Font.Size = stdset.Size
    
    BFileName = StrConv(stdset.Name, vbFromUnicode)

    With T_Font
        For i = 1 To Len(stdset.Name)
            .lfFaceName(i - 1) = BFileName(i - 1)
        Next i

        .lfHeight = -MulDiv(stdset.Size, GetDeviceCaps(lngDC, LOGPIXELSY), 71)
        .lfWidth = 0
        .lfWeight = IIf(stdset.Bold = True, FW_BOLD, FW_NORMAL)
        .lfCharSet = stdset.Charset
        .lfUnderline = stdset.Underline
        .lfItalic = stdset.Italic
        .lfStrikeOut = stdset.Strikethrough
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function GetCurveColumn(ByVal dtDateTime As Date, _
                               ByVal dtBeginDateTime As Date, _
                               Optional ByVal intHourBegin As Integer = 4) As Integer

    '******************************************************************************************************************
    '���ܣ� ��ʱ��������
    '������
    '���أ�
    '******************************************************************************************************************
    Dim varTime As Variant

    Dim strTmp  As String

    Dim intDays As Integer

    Dim intLoop As Integer
    
    On Error GoTo Errhand
    
    GetCurveColumn = -1
    
    '��ʼ��ʱ�䷶Χ����
    Call InitDateTimeRange(varTime, intHourBegin)

    '���㵱ǰ���ʱ������һ��ĵڼ���λ����
    strTmp = Format(dtDateTime, "HH:mm:ss")
    
    For intLoop = 0 To 6
        If Format(strTmp, "HH:mm:ss") >= Format(Split(varTime(intLoop), ",")(0), "HH:mm:ss") And Format(strTmp, "HH:mm:ss") <= Format(Split(varTime(intLoop), ",")(1), "HH:mm:ss") Then
            Exit For
        End If
    Next
    
    If intLoop < 7 Then
        '���㵱���ڵ�ǰ���µ�ҳ���ǵڼ��죨0��ʾ��һ�죻1��ʾ�ڶ���.....��
        intDays = DateDiff("d", Int(dtBeginDateTime), Int(dtDateTime))
        GetCurveColumn = intDays * 6 + intLoop + 1
    End If
    
    Exit Function

Errhand:

    If ErrCenter() = 1 Then

        Resume

    End If

    Call SaveErrLog
            
End Function

Public Function GetCurveDate(ByVal intCOl As Integer, _
                             ByVal dtBeginDateTime As Date, _
                             Optional ByVal intHourBegin As Integer = 4) As String

    '-------------------------------------------------------------------------------------
    '����:�����м����ʱ�䷶Χ
    '���� intCol ��ǰ��    dtBeginDateTime ��ʼʱ��
    '���ظ�ʽΪ:��ʼʱ��;��ֹʱ��
    '-------------------------------------------------------------------------------------
    Dim varTime  As Variant

    Dim intDays  As Integer

    Dim strBegin As String

    Dim strEnd   As String

    Dim lngLoop  As Long

    Dim lng�к�  As Long

    On Error GoTo Errhand
    
    GetCurveDate = -1
    
    '��ʼ��ʱ�䷶Χ����
    Call InitDateTimeRange(varTime, intHourBegin)
    
    '���㵱ǰ�кͿ�ʼʱ�� ��������,�����¼����еĿ�ʼʱ��
    intDays = (intCOl - 1) \ 6
    strBegin = DateAdd("d", intDays, Int(dtBeginDateTime))
    strEnd = strBegin
    
    '���������ڵ�ʱ�䷶Χ
    lng�к� = (intCOl - 1) Mod 6
    
    strBegin = Format(strBegin & " " & Split(varTime(lng�к�), ",")(0), "YYYY-MM-DD HH:mm:ss")
    strEnd = Format(strEnd & " " & Split(varTime(lng�к�), ",")(1), "YYYY-MM-DD HH:mm:ss")

    GetCurveDate = strBegin & ";" & strEnd

    Exit Function

Errhand:

    If ErrCenter = 1 Then

        Resume

    End If

End Function

Public Function InitPara() As Boolean

    '******************************************************************************************************************
    '���ܣ��õ����б��ز���
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop     As Integer

    Dim strTmp      As String

    Dim strTmpBegin As String

    Dim strTmpEnd   As String

    On Error GoTo Errhand
    
    gvarTime = Split(String(6, ";"), ";")
    gintHourBegin = zldatabase.GetPara("���¿�ʼʱ��", glngSys, 1255, 4)
    strTmp = zldatabase.GetPara("���±�־�ָ���", glngSys, 1255, 0)
    mlng���²�����ʾ��ʽ = Val(zldatabase.GetPara("���²�����ʾ��ʽ", glngSys, 1255, "0"))
    If Val(strTmp) = 0 Then
        gstrCaveSplit = "����"
    ElseIf Val(strTmp) = 1 Then
        gstrCaveSplit = "��"
    Else
        gstrCaveSplit = ""
    End If
    
    '���˱䶯�����ʾ����
    '------------------------------------------------------------------------------------------------------------------
    strTmp = zldatabase.GetPara("���µ����", glngSys, 1255, "1;1;1;1;1;1;1;1")

    If UBound(Split(strTmp, ";")) >= 5 Then
        T_BodyFlag.��Ժ = Val(Split(strTmp, ";")(0))
        T_BodyFlag.��� = Val(Split(strTmp, ";")(1))
        T_BodyFlag.ת�� = Val(Split(strTmp, ";")(2))
        T_BodyFlag.���� = Val(Split(strTmp, ";")(3))
        T_BodyFlag.���� = Val(Split(strTmp, ";")(4))
        T_BodyFlag.��Ժ = Val(Split(strTmp, ";")(5))

        If UBound(Split(strTmp, ";")) >= 6 Then T_BodyFlag.���� = Val(Split(strTmp, ";")(6))
        If UBound(Split(strTmp, ";")) >= 7 Then T_BodyFlag.���� = Val(Split(strTmp, ";")(7))
    End If
    
    '�������µ�һ�������ʱ�䷶Χ
    Call InitDateTimeRange(gvarTime, gintHourBegin)
        
    InitPara = True

    Exit Function

Errhand:

    If ErrCenter() = 1 Then

        Resume

    End If

End Function

Public Function InitDateTimeRange(ByRef varTime As Variant, _
                                  Optional ByVal intHourBegin As Integer = 4) As Boolean

    '******************************************************************************************************************
    '���ܣ��������µ�һ�������ʱ�䷶Χ
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop     As Integer

    Dim strTmpBegin As String

    Dim strTmpEnd   As String
    
    On Error GoTo Errhand
    
    varTime = Split(String(6, ";"), ";")
    
    varTime(0) = "00:00:00," & Format(intHourBegin + 1, "00") & ":59:59"
    varTime(5) = (intHourBegin + 18) & ":00:00,23:59:59"
    
    For intLoop = 1 To 4
        strTmpBegin = Format(DateAdd("s", 1, CDate("2000-01-01 " & Split(varTime(intLoop - 1), ",")(1))), "HH:mm:ss")
        strTmpEnd = Format(DateAdd("h", 4, CDate("2000-01-01 " & Split(varTime(intLoop - 1), ",")(1))), "HH:mm:ss")
        varTime(intLoop) = strTmpBegin & "," & strTmpEnd
    Next
    
    InitDateTimeRange = True
    
    Exit Function

Errhand:

    If ErrCenter() = 1 Then
        Resume
    End If

    Call SaveErrLog
End Function

Public Function RetrunEndTime(ByVal dtBegin As Date, ByVal dtEnd As Date, Optional ByVal intHourBegin As Integer = 4) As Date
'**********************************************************************************
'���ܣ�������µ���ֹʱ��Ϳ�ʼʱ���Ƿ���ͬһ��Ԫ�������ͬһ��Ԫ����Ҫ����ֹʱ���Ƶ���һ��Ԫ��
'������strBegin ���µ���ʼʱ��,strEnd ���µ���ֹʱ��(���˳�Ժʱ��)
'����ֵ�����µ���ֹʱ��
'**********************************************************************************
'���󣺶��ڲ��˳�Ժ����Ժʱ����ͬһ�����ӣ���Ҫ¼����Ժ���£�ҲҪ¼���Ժ���£�����Ժ����¼�뵽��һ�����ӡ�

    Dim varTime As Variant
    Dim intLoop As Integer, strTmp As String
    Dim intBegin As Integer, intEnd As Integer
    Dim strEnd As String
    
    RetrunEndTime = dtEnd
    If Format(dtBegin, "YYYY-MM-DD") <> Format(dtEnd, "YYYY-MM-DD") Then Exit Function
    '��ʼ��ʱ�䷶Χ����
    Call InitDateTimeRange(varTime, intHourBegin)
    '1/���㿪ʼʱ�����ֹʱ���ڵڼ�������
    strTmp = Format(dtBegin, "HH:mm:ss")
    For intLoop = 0 To 6
        If Format(strTmp, "HH:mm:ss") >= Format(Split(varTime(intLoop), ",")(0), "HH:mm:ss") And Format(strTmp, "HH:mm:ss") <= Format(Split(varTime(intLoop), ",")(1), "HH:mm:ss") Then
            intBegin = intLoop
            Exit For
        End If
    Next
    strTmp = Format(dtEnd, "HH:mm:ss")
    For intLoop = 0 To 6
        If Format(strTmp, "HH:mm:ss") >= Format(Split(varTime(intLoop), ",")(0), "HH:mm:ss") And Format(strTmp, "HH:mm:ss") <= Format(Split(varTime(intLoop), ",")(1), "HH:mm:ss") Then
            intEnd = intLoop
            Exit For
        End If
    Next
    '2 ����ͬһ�о��˳�
    If intBegin <> intEnd Then Exit Function
    If intEnd > 5 Then Exit Function
    '3 �����ֹʱ������¸�ֵ
    If intEnd > 4 Then
        strEnd = Format(DateAdd("D", 1, dtEnd), "YYYY-MM-DD") & " " & Format(Split(varTime(0), ",")(1), "HH:mm:ss")
    Else
        strEnd = Format(dtEnd, "YYYY-MM-DD") & " " & Format(Split(varTime(intEnd + 1), ",")(1), "HH:mm:ss")
    End If
    
    RetrunEndTime = CDate(Format(strEnd, "YYYY-MM-DD HH:mm:ss"))
End Function

Public Function GetGridItem(ByVal int����ȼ� As Integer, ByVal byt���ò��� As Byte, ByVal lng����ID As Long, Optional int��Ŀ���� As Integer = 1) As ADODB.Recordset

    '**********************************************************************************
    '����:��ȡ���±����Ŀ
    '**********************************************************************************
    Dim rsTemp As New ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo Errhand
    
    '��ȡ�����Ŀ
    gstrSQL = "Select A.�������,A.��Ŀ���,'' ���²�λ,A.��¼��,A.��¼��,A.��¼��,A.��¼ɫ,A.���ֵ,A.��Сֵ,A.��λֵ,nvl(A.��¼Ƶ��,2) ��¼Ƶ��,A.��Ժ�ײ�,B.��Ŀ����," & _
        "   B.������,B.��Ŀֵ��,B.��Ŀ��ʾ,B.��Ŀ����,B.��Ŀ����,B.��ĿС��,B.��Ŀ��λ ��λ" & _
        "   From ���¼�¼��Ŀ A,�����¼��Ŀ B,����������Ŀ C" & _
        "   Where A.��Ŀ���=B.��Ŀ��� And B.��ĿID=C.Id(+) And A.��¼��=2 And nvl(B.��Ŀ����,1)=[4]" & _
        "   And nvl(B.Ӧ�÷�ʽ,0)=1 And nvl(B.����ȼ�,0)>=[1] And nvl(B.���ò���,0) In (0,[2])" & _
        "   And (B.���ÿ���=1 Or (B.���ÿ���=2 And Exists (Select 1 From �������ÿ��� D Where D.��Ŀ���=B.��Ŀ��� And D.����id=[3]))) order by Decode(��Ŀ���,3 ,0,1 ),�������"
        
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ�̶����±����Ŀ", int����ȼ�, byt���ò���, lng����ID, int��Ŀ����)
    Set GetGridItem = rsTemp

    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetAppendGridItem(ByVal lng�ļ�ID As Long, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal int����ȼ� As Integer, ByVal intӤ�� As Long, dt��ʼʱ�� As Date, dt����ʱ�� As Date, ByVal byt���ò��� As Byte, ByVal lng����ID As Long, Optional blnMove As Boolean = False) As ADODB.Recordset
    '**************************************************************************
    '����:��ȡ������ݵ����±����Ŀ�Լ��̶������Ŀ
    '**************************************************************************
    Dim rsTemp As New ADODB.Recordset
    Dim strSql As String

    On Error GoTo Errhand
    
    Set rsTemp = GetGridItem(int����ȼ�, byt���ò���, lng����ID, 2)
    If rsTemp.RecordCount = 0 Then
        '�����ڻ��Ŀֱ����ȡ�̶������Ŀ
        Set rsTemp = GetGridItem(int����ȼ�, byt���ò���, lng����ID, 1)
        Set GetAppendGridItem = rsTemp
        Exit Function
    End If
    With rsTemp
        Do While Not .EOF
            strSql = IIf(strSql = "", "select " & !��Ŀ��� & " ��Ŀ��� from dual", strSql & " UNION ALL select " & !��Ŀ��� & "  ��Ŀ��� from dual ")
            .MoveNext
        Loop
    End With
    
    strSql = "(" & strSql & ") F"
    '��ȡ���Ŀ
    gstrSQL = "Select distinct D.�������,D.��Ŀ���,C.���²�λ,C.���²�λ || D.��¼��  ��¼��,D.��¼��,D.��¼��,D.��¼ɫ,D.���ֵ,D.��Сֵ,D.��λֵ,nvl(D.��¼Ƶ��,2) ��¼Ƶ��,D.��Ժ�ײ�," & _
        "   E.��Ŀ����,E.������,E.��Ŀֵ��,E.��Ŀ��ʾ,E.��Ŀ����,E.��Ŀ����,E.��ĿС��,E.��Ŀ��λ ��λ" & _
        "   FROM ���˻����ļ� B, ���˻������� A,���˻�����ϸ C,���¼�¼��Ŀ D,�����¼��Ŀ E," & strSql & _
        "   Where  B.ID=A.�ļ�ID And A.ID = c.��¼ID  AND B.ID=[1]  AND Nvl(B.Ӥ��,0)=[5]  AND B.����id=[2]    AND B.��ҳid=[3] AND d.��Ŀ���=C.��Ŀ��� " & _
        "   AND c.��¼����=1 And E.��Ŀ����=2  AND E.��Ŀ���=D.��Ŀ���  AND E.����ȼ�>=[4]   AND a.����ʱ�� BETWEEN [6] And [7] And c.��ֹ�汾 Is Null " & _
        "   AND d.��¼��=2 and D.��Ŀ���=F.��Ŀ���"
    
    '��ȡ�̶������Ŀ
    strSql = "Select A.�������,A.��Ŀ���,'' ���²�λ,A.��¼��,A.��¼��,A.��¼��,A.��¼ɫ,A.���ֵ,A.��Сֵ,A.��λֵ,nvl(A.��¼Ƶ��,2) ��¼Ƶ��,A.��Ժ�ײ�,B.��Ŀ����," & _
        "   B.������,B.��Ŀֵ��,B.��Ŀ��ʾ,B.��Ŀ����,B.��Ŀ����,B.��ĿС��,B.��Ŀ��λ ��λ" & _
        "   From ���¼�¼��Ŀ A,�����¼��Ŀ B,����������Ŀ C" & _
        "   Where A.��Ŀ���=B.��Ŀ��� And B.��ĿID=C.Id(+) And A.��¼��=2 And nvl(B.��Ŀ����,1)=1" & _
        "   And nvl(B.Ӧ�÷�ʽ,0)=1 And nvl(B.����ȼ�,0)>=[4] And nvl(B.���ò���,0) In (0,[8])" & _
        "   And (B.���ÿ���=1 Or (B.���ÿ���=2 And Exists (Select 1 From �������ÿ��� D Where D.��Ŀ���=B.��Ŀ��� And D.����id=[9])))"
    
    gstrSQL = "Select �������,��Ŀ���,���²�λ,��¼��,��¼��,��¼��,��¼ɫ,���ֵ,��Сֵ,��λֵ,��¼Ƶ��,��Ժ�ײ�,��Ŀ����," & _
        "   ������,��Ŀֵ��,��Ŀ��ʾ,��Ŀ����,��Ŀ����,��ĿС��,��λ" & _
        "   From (" & gstrSQL & vbCrLf & " UNION ALL " & vbCrLf & strSql & ") order by Decode(��Ŀ���,3 ,0,1 ),�������,��¼��"
    If blnMove Then
        gstrSQL = Replace(gstrSQL, "���˻����ļ�", "H���˻����ļ�")
        gstrSQL = Replace(gstrSQL, "���˻�������", "H���˻�������")
        gstrSQL = Replace(gstrSQL, "���˻�����ϸ", "H���˻�����ϸ")
    End If

    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "", lng�ļ�ID, lng����ID, lng��ҳID, int����ȼ�, intӤ��, dt��ʼʱ��, dt����ʱ��, byt���ò���, lng����ID)
    
    Set GetAppendGridItem = rsTemp

    Exit Function

Errhand:

    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub DrawRotateText(ByVal objDraw As Object, ByVal lngDC As Long, ByVal X As Single, _
                          ByVal Y As Single, _
                          ByVal strText As String, _
                          Optional ByVal ForeColor As Long = 0, _
                          Optional ByVal sglScale As Single = 1)

    '��(X,Y)�����Text�ı�
    Dim objFont    As Object

    Dim lngFont    As Long

    Dim lngOldFont As Long

    Dim X1         As Long
    
    Dim blnPrinter As Boolean
    
    
    If TypeName(objDraw) = "Printer" Then
        blnPrinter = True
    Else
        msngTwips = 1
    End If
    '�����ı���ɫ
    Call SetTextColor(lngDC, ForeColor)

    '�����������
    If Asc(strText) < 0 And strText <> "��" Then
    
        Call GetTextRect(objDraw, X, Y, strText, T_DrawClient.�е�λ, False, , sglScale)
        Call DrawText(lngDC, strText, -1, T_LableRect, DT_CENTER)
        '��ת90���������
    Else
        Set objFont = New clsRotateFont
        Set objFont.LogFont = gstdSet
        
        If blnPrinter = True Then
            objFont.sngTwpic = msngTwips
        Else
            objFont.sngTwpic = 1
        End If
        
        objFont.Rotation = -90
        lngFont = objFont.Handle
        lngOldFont = SelectObject(lngDC, lngFont)
'        Call GetTextRect(objDraw, X, Y, strText, T_DrawClient.�е�λ, False, , sglScale)
'        X1 = T_LableRect.Right - T_LableRect.Left + (T_LableRect.Left - X) / 2
        Call TextOut(lngDC, X + T_DrawClient.�е�λ, Y, strText, LenB(StrConv(strText, vbFromUnicode)))
         
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
    End If
End Sub

Public Sub GetTextRect(ByVal objDraw As Object, ByVal lngX As Long, ByVal lngY As Long, ByVal strInput As String, _
    Optional ByVal lngWidth As Long = 0, Optional bln���� As Boolean = True, Optional ByVal lngHeght As Long = 0, Optional ByVal sngScale As Single = 1)
    
    Dim lngInputW As Long, lng1H As Long
    Dim sngSize As Single
        
    T_LableRect.Left = lngX + 1 '��������߽绮���غ�
    
    If bln���� = True Then
        T_LableRect.Top = lngY - objDraw.TextHeight("1") / 2 / T_TwipsPerPixel.Y
    Else
        T_LableRect.Top = lngY
    End If
    
    T_LableRect.Right = objDraw.TextWidth(strInput) / T_TwipsPerPixel.Y + T_LableRect.Left + 2
    T_LableRect.Bottom = objDraw.TextHeight("1") / T_TwipsPerPixel.Y + T_LableRect.Top + 2
    
    If lngWidth <> 0 Then
        '���ı���ʾ����ʾ��ȵ��м�����
        T_LableRect.Left = T_LableRect.Left + (lngWidth - objDraw.TextWidth(strInput) / T_TwipsPerPixel.Y - 1) / 2
        T_LableRect.Right = objDraw.TextWidth(strInput) / T_TwipsPerPixel.Y + T_LableRect.Left + 2
    End If
    
    If lngHeght <> 0 Then
        T_LableRect.Bottom = T_LableRect.Bottom + (lngHeght - objDraw.TextHeight(1) / T_TwipsPerPixel.Y)
    End If
    
End Sub


Public Sub DrawLine(ByVal lngDC As Long, ByVal lngSX As Long, ByVal lngSY As Long, ByVal lngDX As Long, ByVal lngDY As Long, _
    Optional ByVal lngType As Long = PS_SOLID, Optional ByVal intWidth As Integer = 1, Optional ByVal lngRGB As Long = 0, _
    Optional ByVal blnEndRow As Boolean = False, Optional ByVal blnPrinter As Boolean = False)
    
    Dim X As Long
    Dim lngPen As Long
    Dim lngOldPen As Long
    Dim sngX As Single, sngY As Single
    On Error GoTo Errhand
    '�����»��ʽ��л���
    
    If msngTwips = 0 Then msngTwips = 1
    sngX = 2 * msngTwips
    sngY = 3 * msngTwips

    lngPen = CreatePen(lngType, intWidth, lngRGB)
    lngOldPen = SelectObject(lngDC, lngPen)
    '��ͼ
    Call MoveToEx(lngDC, lngSX, lngSY, T_OldPoint)
    Call LineTo(lngDC, lngDX, lngDY)
    '���������»����¼�ͷ
    If blnEndRow Then
        If lngSY > lngDY Then '������ʧ�ܣ����ϼ�ͷ��
            For X = lngSX - sngX To lngSX + sngX
                Call MoveToEx(lngDC, X, lngSY - sngY, T_OldPoint)
                Call LineTo(lngDC, lngSX, lngSY)
            Next X
        Else '�����³ɹ�(���¼�ͷ)
            For X = lngSX - sngX To lngSX + sngX
                Call MoveToEx(lngDC, X, lngSY + sngY, T_OldPoint)
                Call LineTo(lngDC, lngSX, lngSY)
            Next X
        End If
    End If
    
    '��ԭ���ʲ�����
    Call SelectObject(lngDC, lngOldPen)
    Call DeleteObject(lngPen)
    lngPen = 0
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub DrawRect(ByVal lngDC As Long, ByVal lngSX As Long, ByVal lngSY As Long, ByVal lngDX As Long, ByVal lngDY As Long, _
    Optional ByVal lngType As Long = PS_SOLID, Optional ByVal intWidth As Integer = 1, Optional ByVal lngRGB As Long = 0)
    
    Dim lngPen As Long, lngOldPen As Long
    On Error GoTo Errhand
    '�����»��ʽ��л�һ������
    
    lngPen = CreatePen(lngType, intWidth, lngRGB)
    lngOldPen = SelectObject(lngDC, lngPen)
    '��ͼ
    Call Rectangle(lngDC, lngSX, lngSY, lngDX, lngDY)
    '��ԭ���ʲ�����
    Call SelectObject(lngDC, lngOldPen)
    Call DeleteObject(lngPen)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub OutPutText(ByVal objDraw As Object, ByVal rsDrawItems As ADODB.Recordset, ByVal lngDC As Long, ByVal rsNote As ADODB.Recordset, ByVal mstrBeginDate As String, Optional ByVal sngScale As Single = 1)

    'rsDrawItems  ��¼��Ŀ��������� ��λֵ�Ȼ�����Ϣ
    'rsNote ����˵����Ϣ
    'mstrBeginDate ���µ�ÿҳ��ʼʱ��
    '���������Ϣ:��Ժ,���,ת��,��Ժ,��������,δ��˵��,�ϱ�˵��������
    'δ��˵�����ϱ�˵��,��û�����ת�������估��������Ϣʱ,��ӡ��42-40֮��;�����40��ʼ���´�ӡ
    '��δ��˵�����ϱ�˵����,���ת����Ϣ��һ���̶ȷ������ʱ,����д������̶���,�������̶�Ҳ����Ϣ,˳��
    Dim lngMaxX As Long     '���µ����X����
    Dim lngX    As Long '��һ�е�X����
    Dim lngY    As Long 'Y����
    Dim lngY1   As Long '40 �ȹ̶�����
    Dim i       As Integer
    Dim X, Y As Long '�������ʱ������
    Dim strComment    As String, strText As String
    Dim intAscCharNum As Integer
    Dim rsTemp  As New ADODB.Recordset
    Dim strDate As String
    Dim bln�ϱ� As Boolean
    Dim bln�¼���ʾ���� As Boolean
    
    On Error GoTo Errhand
    
    bln�¼���ʾ���� = (Val(zldatabase.GetPara("���±�־��˳��������", glngSys, 1255, 0)) = 1)
    
    lngMaxX = T_DrawClient.��������.Right - T_DrawClient.�е�λ
    
    rsNote.Filter = "����<>1"

    '���ȼ��������ת������������Ϣ
    If rsNote.RecordCount = 0 Then Exit Sub
    
    rsNote.Sort = "X����,ʱ��,��Ŀ���"
    lngX = rsNote!X����
    
    With rsNote
        Do While Not .EOF
            If Trim(zlCommFun.Nvl(!����)) <> "" Then
                If Not (!���� = 2 Or !���� = 99) Then
                    
                    '���±�־��˳��������
                    If bln�¼���ʾ���� = True Then
                        If lngX <= lngMaxX Then
                            strDate = Format(Split(GetXCoordinate(lngX, mstrBeginDate, False), ",")(0), "YYYY-MM-DD")
                            If CDate(strDate) > CDate(Format(!ʱ��, "YYYY-MM-DD")) Then
                                lngX = Val(!X����)
                                !���� = 1
                            End If
                        Else
                            lngX = lngMaxX
                            !���� = 1
                        End If
                    Else
                        '����x���꣬��������������x���꣬�����У��
                        If lngX > lngMaxX Then lngX = lngMaxX
                    End If
                    
                    !��ӡX���� = IIf(lngX <= Val(!X����), !X����, lngX)
                    !�߶� = GetFontHeight(lngDC, zlCommFun.Nvl(!����))
                    .Update
                    
                    If lngX <= !X���� Then lngX = !X����
                    lngX = lngX + T_DrawClient.�е�λ
                Else
                    !�߶� = GetFontHeight(lngDC, zlCommFun.Nvl(!����))
                    .Update
                End If
            End If
            .MoveNext
        Loop
        
        If .RecordCount > 0 Then .MoveFirst
        lngY = GetYCoordinate(objDraw, rsDrawItems, gint����, 42)
        '�������ת ���������䵽�����X�����ж���ʽ��Y����
        .Filter = "��ӡX����=" & lngMaxX & " And ����<>1"
        .Sort = "ʱ��,��Ŀ���"

        Do While Not .EOF
            !Y���� = lngY
            .Update
            lngY = lngY + Val(!�߶�) + T_DrawClient.�е�λ
            .MoveNext
        Loop
        
        .Filter = "����<>1"
        .MoveFirst
        
        '����δ��˵�����ϱ����ʾλ��(Y����).
        '˵��:��û�����ת��������Ϣ������� ��ӡ�� 42-40��֮�䣬�����ӡ��40�����´�ӡ
        .Sort = "X����,ʱ��,��Ŀ���"
        Set rsTemp = .Clone

        Do While Not .EOF
            lngY = 0
            If (!���� = 2 Or !���� = 99) Then
                bln�ϱ� = False
                Set rsTemp = .Clone
                
                rsTemp.Filter = "(��ӡX����=" & !X���� & " and ����=99) or (��ӡX����=" & !X���� & " and ����=2)"
                
                If rsTemp.BOF Then
                    rsTemp.Filter = "��ӡX����=" & !X����
                End If
                
                If rsTemp.RecordCount > 0 Then
                    lngY = Val(rsTemp!Y����)
                    
                    Do While Not rsTemp.EOF
                        If bln�ϱ� = False Then
                            bln�ϱ� = IIf(rsTemp!���� = 2 Or rsTemp!���� = 99, True, False)

                            If bln�ϱ� = True Then lngY = Val(rsTemp!Y����)
                        End If
                        
                        lngY = lngY + rsTemp!�߶� + T_DrawClient.�е�λ
                        rsTemp.MoveNext
                    Loop
                    
                    lngY1 = GetYCoordinate(objDraw, rsDrawItems, gint����, 40)

                    If lngY > lngY1 Or bln�ϱ� Then lngY1 = lngY
                    
                Else '�������κ���Ϣ ��42��ʼ��ӡ
                    lngY1 = Val(!Y����)
                End If
                
                !Y���� = lngY1
                !��ӡX���� = !X����
                .Update
            End If

            .MoveNext
        Loop
        .MoveFirst
        
        Dim sigNum As Single
        Do While Not .EOF
            '�������
            strComment = Trim(zlCommFun.Nvl(!����))

            If strComment <> "" Then
                X = Val(IIf(Trim(!��ӡX����) <> "", !��ӡX����, !X����))
                Y = Val(!Y����)
                intAscCharNum = 0

                For i = 1 To Len(strComment)
                    If Y < T_DrawClient.�̶�����.Bottom Then
                        strText = Mid(strComment, i, 1)
                        Call GetTextExtentPoint32(lngDC, strText, Len(strText), T_Size)

                        If Asc(strText) < 0 Then
                            If intAscCharNum Mod 2 = 1 Then Y = Y + T_Size.H / 2
                            '��������õ���ֵ
                            sigNum = GetYCoordinate(objDraw, rsDrawItems, gint����, X & "," & Y, False)
                            Y = GetYCoordinate(objDraw, rsDrawItems, gint����, sigNum)
                        End If

                        '���������Ϣ
                        Call DrawRotateText(objDraw, lngDC, X, Y, strText, !��ɫ, sngScale)

                        If Asc(strText) < 0 Then
                            Y = Y + T_Size.H
                            intAscCharNum = 0
                        Else
                            Y = Y + T_Size.H / 2
                            intAscCharNum = intAscCharNum + 1
                        End If
                    End If
                Next i
            End If
            .MoveNext
        Loop

    End With
    
    Exit Sub

Errhand:

    If ErrCenter = 1 Then

        Resume

    End If

End Sub

Public Sub DrawPoly(ByVal lngDC As Long, ByRef PtInPoly() As POINTAPI, Optional ByVal lngStart As Long = 1)

    Dim lngRgn As Long, lngBrush As Long
    Dim lngPen As Long, lngOldPen As Long
    Dim bln����������䷽ʽ As Boolean

    'lngStart:ָ�����ʿ�ʼ������,������������,��������������߸��ǵ���(��ɫ���ܲ�һ��)
    On Error GoTo Errhand
    
    bln����������䷽ʽ = Val(zldatabase.GetPara("���������䷽ʽ", glngSys, 1255, "0")) = 1
    '������򲢻�����
    
    '����ϵͳˢ��
    If bln����������䷽ʽ = True Then
        lngBrush = CreateHatchBrush(HS_VERTICAL, RGB_RED)
    Else
        lngBrush = CreateHatchBrush(HS_BDIAGONAL, RGB_RED)
    End If
    '�������ˢ�ӳɹ�,��ѡ��
    If lngBrush <> 0 Then
        lngRgn = CreatePolygonRgn(PtInPoly(1), UBound(PtInPoly), ALTERNATE)
        FillRgn lngDC, lngRgn, lngBrush
        Call DeleteObject(lngRgn)
        Call DeleteObject(lngBrush)
    
        lngPen = CreatePen(PS_SOLID, 1, RGB_RED)
        lngOldPen = SelectObject(lngDC, lngPen)
        '��ͼ
        Polyline lngDC, PtInPoly(lngStart), UBound(PtInPoly) - lngStart
        '��ԭ���ʲ�����
        Call SelectObject(lngDC, lngOldPen)
        Call DeleteObject(lngPen)
        
    End If
    
    Exit Sub

Errhand:

    If ErrCenter = 1 Then

        Resume

    End If

End Sub

Public Function GetFontHeight(ByVal lngDC As Long, ByVal strComment As String) As Double

    '------------------------------------------------------------------------------------
    '����:�õ��ַ����߶�
    '------------------------------------------------------------------------------------
    Dim intAscCharNum As Integer

    Dim Y             As Double

    Dim strText       As String

    Dim i             As Integer
    
    On Error GoTo Errhand
    
    intAscCharNum = 0

    For i = 1 To Len(strComment)
        strText = Mid(strComment, i, 1)
         
        Call GetTextExtentPoint32(lngDC, strText, Len(strText), T_Size)
         
        If Asc(strText) < 0 Then
            If intAscCharNum Mod 2 = 1 Then Y = Y + T_Size.H / 2
        End If
         
        If Asc(strText) < 0 Then
            Y = Y + T_Size.H
            intAscCharNum = 0
        Else
            Y = Y + T_Size.H / 2
            intAscCharNum = intAscCharNum + 1
        End If

    Next i
    
    GetFontHeight = Y
    
    Exit Function

Errhand:

    If ErrCenter = 1 Then

        Resume

    End If

End Function

Public Function GetDataFromHis(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lngӤ�� As Long, ByVal dtFrom As Date, ByVal dtTo As Date, Optional ByVal bytMode As Byte = 1) As ADODB.Recordset

    '******************************************************************************************************************
    '���ܣ���ҽ����¼��ȡ��������������
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strSql As String
    Dim strNewSql As String
    Dim rsTemp As New ADODB.Recordset
    Dim RS As New ADODB.Recordset
    Dim rsCopy As New ADODB.Recordset
    Dim strFileds As String, strValue As String
    Dim lng������Ŀid As Long
    Dim blnBody As Boolean
    
    On Error GoTo Errhand
    
    blnBody = False
    Select Case bytMode

            '------------------------------------------------------------------------------------------------------------------
        Case 1              '��ҽ����¼��ȡ��������������
        
            '        dtFrom = dtFrom - 14
        
            strSql = "Select ִ��ʱ��,����,�ε�" & vbNewLine & _
                " From (Select ִ��ʱ��,����, Rownum As �ε�" & vbNewLine & _
                "       From (Select Distinct C.ִ��ʱ��,'����' As ���� " & vbNewLine & _
                "              From ����ҽ����¼ A, ������ĿĿ¼ B, ����ҽ��ִ�� C" & vbNewLine & _
                "              Where A.����id = [1] And A.��ҳid = [2] And Nvl(A.Ӥ��, 0) = [3] And A.ҽ����Ч = 1 And A.������Ŀid = B.ID And" & vbNewLine & _
                "                    A.������� = 'F' And A.ҽ��״̬ = 8 And C.ҽ��id = A.ID And C.ִ��ʱ�� < =[5] " & vbNewLine & _
                "               Union All Select a.����ʱ�� As ִ��ʱ��,'����' As ���� From ������������¼ a Where a.����id=[1] And a.��ҳid=[2] And a.����ʱ�� Is Not Null And RowNum<2) " & _
                "       Order By ִ��ʱ��)" & vbNewLine & "Where ִ��ʱ�� >= [4] And �ε� <= 12 " & vbNewLine & "Order By ִ��ʱ�� "
                
            Set GetDataFromHis = zldatabase.OpenSQLRecord(strSql, "���µ�", lng����ID, lng��ҳID, lngӤ��, dtFrom, dtTo)

            '------------------------------------------------------------------------------------------------------------------
        Case 2              '���ת��־(��Ժ,��Ժ,ת��,����)
            strFileds = "����," & adLongVarChar & ",50|ʱ��," & adDate & ",20|����," & adLongVarChar & ",100|�к�," & adDouble & ",5"
            Call Record_Init(rsCopy, strFileds)
            strFileds = "����|ʱ��|����|�к�"
            '1-��Ժ��2-��ƣ�3-ת�ƣ�4-������5-��λ�ȼ��䶯��6-����ȼ��䶯��7-����ҽʦ�ı䣻8-���λ�ʿ�ı�,9-���۲���תסԺ,10-����Ԥ��Ժ,11-����ҽʦ�䶯,12-����ҽʦ�䶯,13-����䶯
            
            '0-��ͨ;1-����;2-סԺ;3-ת��;4-����;5-��Ժ;6-תԺ;7-����;8-����;9-����;10-��Σ;11-����;12-��¼�����;14-��ǰ
            '��ȡ������¼ID
'            strSQL = "Select ID From ������ĿĿ¼ Where ���='Z' And ��������='11' "
'            Set RS = zlDatabase.OpenSQLRecord(strSQL, "���µ�")
'
'            If RS.BOF = False Then lng������Ŀid = zlCommFun.Nvl(RS("ID").Value)
        
            strSql = _
               "    Select ����,ʱ��,����,�к� From (" & vbNewLine & _
               "    Select b.���� As ����,��ʼʱ�� As ʱ��, Decode(��ʼԭ��, 2,'���',3, 'ת��',4,'����'||Decode(����,Null,'','('||����||')')) As ����,Decode(��ʼԭ��,2,9,3,6,4,7) As �к� " & vbNewLine & _
               "    From ���˱䶯��¼ A,���ű� b" & vbNewLine & _
               "    Where b.id(+)=a.����id and a.��ʼԭ�� In (2,3,4) And A.����id = [1] And A.��ҳid = [2]  And [3]=0 And A.��ʼʱ�� Between [4] And [5] " & vbNewLine & _
               "    Union " & vbNewLine & _
               "    Select ����,ʱ��,����,�к� From (Select * From (Select  b.���� As ����,A.��ʼʱ�� As ʱ��, '���' As ����,9 As �к� " & vbNewLine & _
               "    From ���˱䶯��¼ A,���ű� B" & vbNewLine & _
               "    Where b.id(+)=a.����id And a.��ʼԭ��=1 And A.����id = [1] And A.��ҳid = [2] And [3]=0 And NOT Exists " & vbNewLine & _
               "   (Select ID From ���˱䶯��¼ C Where C.��ʼԭ��=2 And C.����ID=A.����ID And C.��ҳID=A.��ҳID And [3]=0) Order By a.��ʼʱ��) Where RowNum=1) Where ʱ�� Between [4] And [5] " & vbNewLine & _
               "    )" & vbNewLine & _
               "    Union All" & vbNewLine & _
               "    Select '' As ����,ʱ��,����,�к� From (Select * From (Select ��ʼʱ�� As ʱ��, '��Ժ' As ����,5 As �к� " & vbNewLine & _
               "    From ���˱䶯��¼ A" & vbNewLine & _
               "    Where a.��ʼԭ��=1 And A.����id = [1] And A.��ҳid = [2] And [3]=0 Order By a.��ʼʱ��) Where RowNum=1) Where ʱ�� Between [4] And [5] " & vbNewLine & _
               "    Union All" & vbNewLine & _
               "    Select '' As ����,Nvl(b.��ʼִ��ʱ��,a.��Ժ����) As ʱ��, Decode(��Ժ��ʽ, '����', '��Ժ', ��Ժ��ʽ) As ����,8 As �к� " & vbNewLine & _
               "    From ������ҳ A,(Select x.����id,x.��ҳid,Max(x.��ʼִ��ʱ��) As ��ʼִ��ʱ�� From ����ҽ����¼ x,������ĿĿ¼ z Where x.����id=[1] And x.��ҳid=[2] " & vbNewLine & _
               "    And x.������Ŀid+0=z.ID And x.ҽ��״̬ in (3,8) And z.���='Z' And z.��������='11' Group By x.����id,x.��ҳid) B " & vbNewLine & _
               "    Where A.����id = [1] And A.��ҳid = [2] And A.��Ժ���� Between [4] And [5] And a.����id=b.����id(+) And a.��ҳid=b.��ҳid(+) "
        
            strSql = "Select * From (" & strSql & ") Order By ʱ��,�к� "
            
            Set RS = zldatabase.OpenSQLRecord(strSql, "���µ�", lng����ID, lng��ҳID, lngӤ��, dtFrom, dtTo)
            
            Do While Not RS.EOF
                strValue = Nvl(RS!����) & "|" & CDate(RS!ʱ��) & "|" & Nvl(RS!����) & "|" & Val(Nvl(RS!�к�))
                Call Record_Add(rsCopy, strFileds, strValue)
            RS.MoveNext
            Loop
                    
            If lngӤ�� <> 0 Then
                '��ȡӤ��ҽ����¼(ת��,��Ժ,����)
                strNewSql = _
                    "   select /*+ RULE */ ����,��ʼִ��ʱ��,decode(��������,3,'ת��',5,'��Ժ','����') ����,Decode(��������,'3',3,8) �к� From (" & vbNewLine & _
                    "   select D.���� ����,B.��ʼִ��ʱ��,C.�������� " & vbNewLine & _
                    "   from ������ҳ A,����ҽ����¼ B,������ĿĿ¼ C,���ű� D" & vbNewLine & _
                    "   where A.����ID=[1] and A.��ҳID=[2] And  A.����ID=B.����ID(+) And A.��ҳID=B.��ҳID(+) And B.Ӥ��(+)=[3]" & vbNewLine & _
                    "   and B.������Ŀid+0=C.ID  And B.ҽ��״̬=8  and C.���='Z' And   B.ִ�п���ID=D.ID(+)" & vbNewLine & _
                    "   and  exists (select 1 from Table(Cast(f_str2list('3,5,11') As zlTools.t_strlist)) where C.��������=COLUMN_VALUE) order by B.��ʼִ��ʱ�� DESC" & vbNewLine & _
                    "   ) where Rownum<2"

                Set rsTemp = zldatabase.OpenSQLRecord(strNewSql, "���µ�", lng����ID, lng��ҳID, lngӤ��)
                blnBody = (rsTemp.RecordCount > 0)
                
                '�������Ӥ������ת�ƣ���Ժҽ����Ϣ����Ҫ����ĸ����Ϣ
                If blnBody = True Then
                    rsCopy.Filter = "ʱ��>='" & CDate(rsTemp!��ʼִ��ʱ��) & "'"
                    Do While Not rsCopy.EOF
                        rsCopy.Delete
                        rsCopy.Update
                    rsCopy.MoveNext
                    Loop
                    'ɾ��ĸ�ױ��˵ĳ�Ժ��Ϣ
                    rsCopy.Filter = "�к�=8"
                    Do While Not rsCopy.EOF
                        rsCopy.Delete
                        rsCopy.Update
                    rsCopy.MoveNext
                    Loop
                    '���Ӥ��ҽ����Ϣ
                    rsTemp.MoveFirst
                    If CDate(Format(rsTemp!��ʼִ��ʱ��, "YYYY-MM-DD HH:mm:ss")) >= CDate(Format(dtFrom, "YYYY-MM-DD HH:mm:ss")) And CDate(rsTemp!��ʼִ��ʱ��) <= CDate(Format(dtTo, "YYYY-MM-DD HH:mm:ss")) Then
                        strValue = Nvl(rsTemp!����) & "|" & CDate(rsTemp!��ʼִ��ʱ��) & "|" & Nvl(rsTemp!����) & "|" & Val(Nvl(rsTemp!�к�))
                        Call Record_Add(rsCopy, strFileds, strValue)
                    End If
                End If
            End If
            
            rsCopy.Filter = 0
            'Call OutputRsData(rsCopy, True)
            Set GetDataFromHis = rsCopy

            '------------------------------------------------------------------------------------------------------------------
        Case 3              '����������¼�������/��������
        
            strSql = _
                "   Select '' As ����,a.����ʱ�� As ʱ��,'����' As ����,13 As �к� From ������������¼ a " & _
                "   Where a.����id=[1] And a.��ҳid=[2] And a.���=[3] And a.����ʱ�� Between [4] And [5]"
            Set GetDataFromHis = zldatabase.OpenSQLRecord(strSql, "���µ�", lng����ID, lng��ҳID, lngӤ��, dtFrom, dtTo)
    End Select
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function CheckFileBack(ByVal lngID As Long, ByVal blnMove As Boolean) As Boolean
'---------------------------------------------------------------
'����:����ļ��Ƿ�鵵
'---------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSql As String
    On Error GoTo Errhand
    
    CheckFileBack = False
    strSql = "Select 1 From ���˻����ļ� Where Id=[1] And �鵵�� Is Not Null"
    Set rsTemp = zldatabase.OpenSQLRecord(strSql, "����ļ��Ƿ�鵵", lngID)
    If blnMove = True Then
        strSql = Replace(strSql, "���˻����ļ�", "H���˻����ļ�")
    End If
    If rsTemp.RecordCount > 0 Then
        CheckFileBack = True
    End If
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Function ConvertTimeToChinese(ByVal strTime As String) As String

    '------------------------------------------------------------------------------------------------------------------
    '���ܣ�ת��ʱ��Ϊ���� �� 22:59 ת��Ϊ��ʮ��ʱ��ʮ�ŷ�
    '������ʱ�� ��ʽΪ Format(strtime,"HH:mm")
    '���أ�
    '------------------------------------------------------------------------------------------------------------------
    Dim strTmp1 As String

    Dim strTmp2 As String
    
    strTime = Format(strTime, "HH:mm")

    If InStr(strTime, ":") <= 0 Then Exit Function

    On Error GoTo Errhand
    
    strTmp1 = Left(strTime, InStr(strTime, ":") - 1)
    strTmp2 = Mid(strTime, InStr(strTime, ":") + 1)
    
    strTmp1 = Switch(strTmp1 = "00", "��", strTmp1 = "01", "һ", strTmp1 = "02", "��", strTmp1 = "03", "��", strTmp1 = "04", "��", strTmp1 = "05", "��", strTmp1 = "06", "��", strTmp1 = "07", "��", strTmp1 = "08", "��", strTmp1 = "09", "��", strTmp1 = "10", "ʮ", strTmp1 = "11", "ʮһ", strTmp1 = "12", "ʮ��", strTmp1 = "13", "ʮ��", strTmp1 = "14", "ʮ��", strTmp1 = "15", "ʮ��", strTmp1 = "16", "ʮ��", strTmp1 = "17", "ʮ��", strTmp1 = "18", "ʮ��", strTmp1 = "19", "ʮ��", strTmp1 = "20", "��ʮ", strTmp1 = "21", "��ʮһ", strTmp1 = "22", "��ʮ��", strTmp1 = "23", "��ʮ��")
    
    strTmp2 = Switch(strTmp2 = "00", "��", strTmp2 = "01", "һ", strTmp2 = "02", "��", strTmp2 = "03", "��", _
       strTmp2 = "04", "��", strTmp2 = "05", "��", strTmp2 = "06", "��", strTmp2 = "07", "��", _
       strTmp2 = "08", "��", strTmp2 = "09", "��", strTmp2 = "10", "ʮ", strTmp2 = "11", "ʮһ", _
       strTmp2 = "12", "ʮ��", strTmp2 = "13", "ʮ��", strTmp2 = "14", "ʮ��", strTmp2 = "15", "ʮ��", _
       strTmp2 = "16", "ʮ��", strTmp2 = "17", "ʮ��", strTmp2 = "18", "ʮ��", strTmp2 = "19", "ʮ��", _
       strTmp2 = "20", "��ʮ", strTmp2 = "21", "��ʮһ", strTmp2 = "22", "��ʮ��", strTmp2 = "23", "��ʮ��", _
       strTmp2 = "24", "��ʮ��", strTmp2 = "25", "��ʮ��", strTmp2 = "26", "��ʮ��", strTmp2 = "27", "��ʮ��", _
       strTmp2 = "28", "��ʮ��", strTmp2 = "29", "��ʮ��", strTmp2 = "30", "��ʮ", strTmp2 = "31", "��ʮһ", _
       strTmp2 = "32", "��ʮ��", strTmp2 = "33", "��ʮ��", strTmp2 = "34", "��ʮ��", strTmp2 = "35", "��ʮ��", _
       strTmp2 = "36", "��ʮ��", strTmp2 = "37", "��ʮ��", strTmp2 = "38", "��ʮ��", strTmp2 = "39", "��ʮ��", _
       strTmp2 = "40", "��ʮ", strTmp2 = "41", "��ʮһ", strTmp2 = "42", "��ʮ��", strTmp2 = "43", "��ʮ��", _
       strTmp2 = "44", "��ʮ��", strTmp2 = "45", "��ʮ��", strTmp2 = "46", "��ʮ��", strTmp2 = "47", "��ʮ��", _
       strTmp2 = "48", "��ʮ��", strTmp2 = "49", "��ʮ��", strTmp2 = "50", "��ʮ", strTmp2 = "51", "��ʮһ", _
       strTmp2 = "52", "��ʮ��", strTmp2 = "53", "��ʮ��", strTmp2 = "54", "��ʮ��", strTmp2 = "55", "��ʮ��", _
       strTmp2 = "56", "��ʮ��", strTmp2 = "57", "��ʮ��", strTmp2 = "58", "��ʮ��", strTmp2 = "59", "��ʮ��")
                    
    ConvertTimeToChinese = strTmp1 & "ʱ" & strTmp2 & "��"
    
    Exit Function

Errhand:

    If ErrCenter = 1 Then

        Resume

    End If

End Function

Public Function DrawPicture(objDraw As Object, ByVal strFile As String, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, Optional ByVal bln��Դ As Boolean = False) As Boolean

    '******************************************************************************************************************
    '���ܣ���������С�Զ��ȱ���������Ƭ�ļ�
    '����������ǰ����Ƭ�ļ�
    '���أ����ź����Ƭ�ļ�
    '******************************************************************************************************************
    Dim strTmp  As String

    Dim objMap  As StdPicture

    Dim W       As Single

    Dim H       As Single

    Dim sglPerW As Single

    Dim sglPerH As Single

    Dim sglPer  As Single

    Dim cx      As Long

    Dim cy      As Long

    On Error GoTo Errhand
    
    If strFile = "" Then Exit Function
    
    cx = X2 - X1
    cy = Y2 - Y1
    
    If bln��Դ Then
        Set objMap = VB.LoadResPicture(strFile, vbResBitmap)
    Else
        Set objMap = VB.LoadPicture(strFile)
    End If
    
    W = objMap.Width * 0.566950910348006
    H = objMap.Height * 0.566950910348006
    
    If W > cx Then sglPerW = cx / W
    If H > cy Then sglPerH = cy / H
    
    If W > cx Or H > cy Then
        sglPer = IIf(sglPerW > sglPerH, sglPerH, sglPerW)
        W = W * sglPer
        H = H * sglPer
    End If
    objDraw.PaintPicture objMap, X1, Y1, W, H
    DrawPicture = True

    Exit Function

Errhand:

    If ErrCenter = 1 Then

        Resume

    End If

End Function

Public Sub CreatePoly(rsPoint As ADODB.Recordset, ByVal objDraw As Object, ByVal lngDC As Long, ByVal strBeginDate As String, ByVal str�������� As String)

'rsPoint ��¼�� �������  ��Ŀ���,X����,Y����
    Dim arrData, arrPt

    Dim bln���� As Boolean      '����������ǵ�Ե�,���ʱ����Ӧ���������γ����������

    Dim bln�� As Boolean, bln�� As Boolean, bln��ǰ As Boolean, bln�Ͽ� As Boolean, bln��Ч As Boolean

    Dim intDO   As Integer, intMax As Integer             'intLast��¼���һ����Ч������

    Dim recttmp As RECT, SinX As Single, sinY As Single, sin��X As Single, sin��X As Single
    
    Dim str��ǰ As String, str�� As String, str�� As String

    Dim str���� As String, str���� As String

    Dim PtInPoly() As POINTAPI, intCOl As Integer, intCols As Integer, intCount As Integer
    Dim blnPrinter As Boolean
    Dim intBold As Integer, intFine As Integer

    On Error GoTo Errhand

    '1�����ʶ�Ӧ1��3������,����������ÿһ�춼��ֵ,�����γ�����
    '�γɵ����򼯺ϱ�����������,����,��װ������,�ٵ���װ������,�γ�������һ������
    '�ɵ���ɵķ������,��DrawPoly����ɷ�����������
    
    If TypeName(objDraw) = "Printer" Then
        intBold = 4
        intFine = 4
        blnPrinter = True
    Else
        intBold = 2
        intFine = 1
        blnPrinter = False
    End If
    
    rsPoint.Sort = "��Ŀ���,ʱ��"
    arrData = Split(str��������, ",")
    intMax = UBound(arrData)
    
    For intDO = 0 To intMax

        SinX = Val(Split(arrData(intDO), ";")(0))
        sinY = Val(Split(arrData(intDO), ";")(1))
        '����ǰ���ʼ������򼯺�
        intCount = intCount + 1
        ReDim Preserve PtInPoly(intCount)
        str���� = str���� & "," & SinX + T_DrawClient.�е�λ / 2 & ";" & sinY
        
        '��������,�������е���������
        If Not bln���� Then
            bln�� = False
            rsPoint.Filter = "��Ŀ���=" & gint���� & " And X����<" & Val(Split(arrData(intDO), ";")(0))
            
            If rsPoint.RecordCount <> 0 Then
               rsPoint.Sort = "X���� DESC"
                bln�Ͽ� = (rsPoint!�Ͽ� = 1)
                If Not bln�Ͽ� Then
                    rsPoint.Sort = "X���� DESC"
                    sin��X = rsPoint!X����
                
                    '���ݵ�ǰ�����ȡʱ��
                    str�� = GetXCoordinate(sin��X, strBeginDate, False)
                    str��ǰ = GetXCoordinate(Val(Split(arrData(intDO), ";")(0)), strBeginDate, False)
                    '��ǰ���ǰһʱ�����һ��û�����ݾͶϿ�
                    If DateDiff("d", CDate(Split(str��, ",")(0)), CDate(Split(str��ǰ, ",")(0))) < 2 Then
                        recttmp.Left = rsPoint!X����
                        recttmp.Top = rsPoint!Y����
                        '���������������򼯺�
                        intCount = intCount + 1
                        ReDim Preserve PtInPoly(intCount)
                        str���� = str���� & "," & rsPoint!X���� + T_DrawClient.�е�λ / 2 & ";" & rsPoint!Y����
                        bln�� = True
                    End If
                End If
            End If
        End If
        
        bln��ǰ = False
        'ȱʡ�Ǻ͵�ǰ�е���������
        rsPoint.Filter = "��Ŀ���=" & gint���� & " And X����=" & Val(Split(arrData(intDO), ";")(0))
        bln��ǰ = (rsPoint.RecordCount <> 0)

        If bln��ǰ Then
            If Not bln�� Then
                recttmp.Left = rsPoint!X����
                recttmp.Top = rsPoint!Y����
            End If

            bln�Ͽ� = (rsPoint!�Ͽ� = 1)

            '����ǰ�����������򼯺�
            If Not bln���� Then
                intCount = intCount + 1
                ReDim Preserve PtInPoly(intCount)
                str���� = str���� & "," & rsPoint!X���� + T_DrawClient.�е�λ / 2 & ";" & rsPoint!Y����
            End If
        End If

        bln�� = False

        If Not bln�Ͽ� Then
            rsPoint.Filter = "��Ŀ���=" & gint���� & " And X����>" & Val(Split(arrData(intDO), ";")(0))
            
            If rsPoint.RecordCount <> 0 Then
                rsPoint.Sort = "X���� ASC"
                sin��X = rsPoint!X����
            
                '���ݵ�ǰ�����ȡʱ��
                str�� = GetXCoordinate(sin��X, strBeginDate, False)
                str��ǰ = GetXCoordinate(Val(Split(arrData(intDO), ";")(0)), strBeginDate, False)
                '��ǰ�����һʱ�����һ��û�����ݾͶϿ�
                If DateDiff("d", CDate(Split(str��ǰ, ",")(0)), CDate(Split(str��, ",")(0))) < 2 Then
                    bln�� = True
                    recttmp.Right = rsPoint!X����
                    recttmp.Bottom = rsPoint!Y����
                    '���������������򼯺�
                    intCount = intCount + 1
                    ReDim Preserve PtInPoly(intCount)
                    str���� = str���� & "," & rsPoint!X���� + T_DrawClient.�е�λ / 2 & ";" & rsPoint!Y����
                End If
            End If
        End If
        
        '�Ȱ���߷��
        If bln���� = False Then
            If bln��ǰ = True Then
                '�����л�ǰ�е���������
                Call DrawLine(lngDC, recttmp.Left + T_DrawClient.�е�λ / 2, recttmp.Top, SinX + T_DrawClient.�е�λ / 2, sinY, PS_SOLID, intFine, RGB_RED)
            End If

            bln���� = (bln�� Or bln��) And bln��ǰ
        End If
        
        '�ҵ��ұߵķ������������
        If bln���� Then
            bln���� = False

            If bln�� = True Then
                '�жϵ�ǰ���ʶ�Ӧ����һ����������һ������X�����Ƿ����,����Ⱦͷ������
                If intDO < intMax Then
                    If recttmp.Right = Val(Split(arrData(intDO + 1), ";")(0)) Then
                        bln���� = True
                    End If
                End If
            End If
            
            
            If Not bln���� Then
                '��֯����,��������ʼ,Ȼ��ת������(���ʴ����ʼ,�ٻص�֮ǰ������,�ٻص���һ������,�γɷ������)
                intCount = 1
                str���� = Mid(str����, 2)
                arrPt = Split(str����, ",")
                intCols = UBound(arrPt)

                For intCOl = 0 To intCols
                    PtInPoly(intCount).X = Split(arrPt(intCOl), ";")(0)
                    PtInPoly(intCount).Y = Split(arrPt(intCOl), ";")(1)
                    intCount = intCount + 1
                Next

                str���� = Mid(str����, 2)
                arrPt = Split(str����, ",")
                intCols = UBound(arrPt)

                For intCOl = intCols To 0 Step -1
                    PtInPoly(intCount).X = Split(arrPt(intCOl), ";")(0)
                    PtInPoly(intCount).Y = Split(arrPt(intCOl), ";")(1)
                    intCount = intCount + 1
                Next

                '��������γɷ������
                ReDim Preserve PtInPoly(intCount)
                PtInPoly(intCount).X = PtInPoly(1).X
                PtInPoly(intCount).Y = PtInPoly(1).Y
                
                '��������
                Call DrawPoly(lngDC, PtInPoly, UBound(Split(str����, ",")) + 1)

            End If
        End If

        If Not bln���� Then
            intCount = 0
            str���� = ""
            str���� = ""
            ReDim Preserve PtInPoly(intCount)
        End If

    Next
    
    rsPoint.Filter = ""

    Exit Sub

Errhand:

    If ErrCenter() = 1 Then

        Resume

    End If

End Sub


Public Sub GetConverPoint(rsPiont As ADODB.Recordset)
'---------------------------------------------------------------------------------------
'����:������֯�غϵĵ�
'---------------------------------------------------------------------------------------
    Dim SinX, sinY As Single
    Dim rsConVerPoint As New ADODB.Recordset
    Dim strFields, strValues As String
    Dim lng��Ŀ��� As Long
    Dim strPart As String
    On Error GoTo Errhand
    
    If rsPiont.RecordCount = 0 Then Exit Sub
    
    strFields = "�ص���ʶ," & adLongVarChar & ",30|�ص���Ŀ," & adInteger & ",30|��Ŀ���," & adLongVarChar & ",18|" & _
        "���²�λ," & adLongVarChar & ",20"
    Call Record_Init(rsConVerPoint, strFields)
    
    '�����غϵĵ�
    rsPiont.Filter = ""
    rsPiont.Sort = "X����,Y����"
    With rsPiont
        Do While Not .EOF
            If SinX = Val(!X����) And sinY = Val(!Y����) Then
                strFields = "�ص���ʶ|�ص���Ŀ|��Ŀ���"
                rsConVerPoint.Filter = "�ص���ʶ='" & SinX & "," & sinY & "'"
                If rsConVerPoint.RecordCount = 0 Then
                    strValues = SinX & "," & sinY
                    strValues = strValues & "|" & 2
                    strValues = strValues & "|" & lng��Ŀ��� & "," & !��Ŀ���
                    Call Record_Add(rsConVerPoint, strFields, strValues)
                Else
                    strFields = "�ص���Ŀ|��Ŀ���"
                    strValues = Val(rsConVerPoint!�ص���Ŀ) + 1
                    strValues = strValues & "|" & rsConVerPoint!��Ŀ��� & "," & !��Ŀ���
                    Call Record_Update(rsConVerPoint, strFields, strValues, "�ص���ʶ|" & SinX & "," & sinY)
                End If
                
                If InStr(1, "," & rsConVerPoint!��Ŀ��� & ",", "," & gint���� & ",") > 0 And strPart <> "" Then
                    strFields = "���²�λ": strValues = strPart
                    Call Record_Update(rsConVerPoint, strFields, strValues, "�ص���ʶ|" & SinX & "," & sinY)
                    strPart = ""
                End If
                
                rsConVerPoint.Filter = ""
                
            End If
            SinX = Val(!X����)
            sinY = Val(!Y����)
            lng��Ŀ��� = !��Ŀ���
            If lng��Ŀ��� = gint���� Then strPart = !��λ
        .MoveNext
        Loop
    End With
    
    '��֯�����ظ���������ʶ
    
    Dim strTemp As String
    Dim rsTemp As New ADODB.Recordset
    Dim lngID As Long
   'Dim strPart As String
    Dim strOverpart As String
        
    If rsConVerPoint.RecordCount > 0 Then
        rsConVerPoint.MoveFirst
        Do While Not rsConVerPoint.EOF
                strTemp = rsConVerPoint!��Ŀ���
                strOverpart = ""
                strPart = ""
                
                '�������µ��غϵ����� ���ڲ�λ����
                If InStr(1, "," & strTemp & ",", "," & gint���� & ",") > 0 Then
                    strTemp = "0," & strTemp & ",0"
                    strTemp = Replace(strTemp, "," & gint���� & ",", ",")
                    
                    gstrSQL = " select  C.���,C.�ϼ����,C.��Ŀ���,C.���²�λ from �����ص���� C," & _
                        "   (Select ��� " & _
                        "   From �����ص���� A,(select �ϼ����,count(1) ���� " & _
                        "   from �����ص���� where ��Ŀ��� in (" & strTemp & ") or (��Ŀ���=1 and nvl(���²�λ,'Ҹ��')=[2]) group by �ϼ����) B " & _
                        "   where A.���=B.�ϼ���� and A.�ص���Ŀ=B.���� and B.����=[1]) D " & _
                        "   where C.�ϼ����=D.��� and C.��Ŀ��� is not null order by C.���"
                Else
                    gstrSQL = " select  C.���,C.�ϼ����,C.��Ŀ���,C.���²�λ from �����ص���� C," & _
                        "   (Select ��� " & _
                        "   From �����ص���� A,(select �ϼ����,count(1) ���� " & _
                        "   from �����ص���� where ��Ŀ��� in (" & strTemp & ") group by �ϼ����) B " & _
                        "   where A.���=B.�ϼ���� and A.�ص���Ŀ=B.���� and B.����=[1]) D " & _
                        "   where C.�ϼ����=D.��� and C.��Ŀ��� is not null order by C.���"
                End If
                Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ�ص���Ŀ", Val(rsConVerPoint!�ص���Ŀ), zlCommFun.Nvl(rsConVerPoint!���²�λ))
                
                If rsTemp.RecordCount > 0 Then
                    lngID = rsTemp!��Ŀ���
                    strPart = rsTemp!�ϼ����  '�ص���λ������
                    
                    Do While Not rsTemp.EOF
                        If lngID <> rsTemp!��Ŀ��� Then
                            strOverpart = strOverpart & "," & rsTemp!��Ŀ���
                        End If
                    rsTemp.MoveNext
                    Loop
                    
                    If strOverpart <> "" Then strOverpart = Mid(strOverpart, 2)
                    
                    '�����ظ��ĵ�
                    rsPiont.Filter = "X����=" & Split(rsConVerPoint!�ص���ʶ, ",")(0) & _
                        " and Y����=" & Split(rsConVerPoint!�ص���ʶ, ",")(1)
                        
                    Do While Not rsPiont.EOF
                        If lngID = rsPiont!��Ŀ��� Then
                            rsPiont!�ص���Ŀ = strOverpart
                            rsPiont!��λ = strPart
                        Else
                            rsPiont!�ص� = 1
                        End If
                    rsPiont.MoveNext
                    Loop
                    
                    rsPiont.Filter = ""
                End If
        rsConVerPoint.MoveNext
        Loop
    End If
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function GetXCoordinate(ByVal strInput As String, ByVal strBeginDate As String, Optional ByVal bln���� As Boolean = True) As String

    '����ʱ��õ�X��������X����ת��Ϊʱ�䷶Χ
    Dim SinX   As Single

    Dim intDO  As Integer, intMax As Integer

    Dim intDay As Integer, intTime As Integer

    Dim strDay As String, strTime As String

    On Error GoTo Errhand
    
    If bln���� Then
        '��һ����0,��������6
        strDay = Split(strInput, " ")(0)

        If InStr(1, strInput, " ") <> 0 Then
            strTime = Split(strInput, " ")(1)
        Else
            strTime = "00:00:00"
        End If

        intDay = DateDiff("d", CDate(strBeginDate), CDate(strInput))
        
        '�õ�����Ŀ̶�
        intMax = 5

        For intDO = 0 To intMax

            If strTime >= Split(gvarTime(intDO), ",")(0) And strTime <= Split(gvarTime(intDO), ",")(1) Then
                intTime = intDO
                Exit For
            End If
        Next
        
        '����õ�X����(ÿ��6��,������*�е�λ�õ�����)
        SinX = Format(T_DrawClient.��������.Left + (T_DrawClient.�е�λ * (intDay * 6 + intTime)), "#0.0")
        GetXCoordinate = SinX
    Else
        '����õ������ٸ��̶�
        SinX = Val(strInput)
        intTime = (SinX - T_DrawClient.��������.Left) \ T_DrawClient.�е�λ
        intDay = intTime \ 6
        intTime = intTime Mod 6
        
        strDay = Format(DateAdd("d", intDay, strBeginDate), "yyyy-MM-dd")
        strTime = gvarTime(intTime)
        GetXCoordinate = strDay & " " & Split(gvarTime(intTime), ",")(0) & "," & strDay & " " & Split(gvarTime(intTime), ",")(1)
    End If
    
    Exit Function

Errhand:

    If ErrCenter = 1 Then

        Resume

    End If

End Function


Public Function GetYCoordinate(ByVal objDraw As Object, ByVal rsDrawItems As ADODB.Recordset, ByVal int��Ŀ��� As Integer, ByVal strInput As String, Optional ByVal bln���� As Boolean = True, Optional lngDC As Long = 0, Optional ByVal blnOutput As Boolean = False) As String

    Dim lngCurX As Long, sinCurY As Single, sinScale As Single

    On Error GoTo Errhand

    '����ָ���������ݵ�Y��������Y�����������
    '���Ըú�������ȷ�Կ�����Paint_Canvas�����Ӹô���ʵ��(˼��:�ɸú����Լ��������ݼ���õ�Y����,��ת��Ϊ����,��ת��Ϊ���������ַ����к˶�,��ӡ������˵��ת������):
    '   Call GetYCoordinate(1, GetYCoordinate(1, "200," & GetYCoordinate(1, "37.5", True, False), False),true,true)
    
    rsDrawItems.Filter = "��Ŀ���=" & int��Ŀ���

    If rsDrawItems.RecordCount = 0 Then
        If int��Ŀ��� = gint���� Then rsDrawItems.Filter = "��Ŀ���=2"
    End If
    
    If rsDrawItems.RecordCount = 0 Then
        GetYCoordinate = 0
        Exit Function
    End If
    
    If bln���� Then
        '�õ���Ч������ʼ����
        lngCurX = Split(rsDrawItems!���ֵ����, ",")(0)
        sinCurY = Split(rsDrawItems!���ֵ����, ",")(1)
        
        '�������ֵ�뵱ǰֵ֮��Ĳ��,�Լ���Сֵ,����õ������ٸ��̶�,�ٸ��ݵ�λ�̶ȵõ�ʵ������
        sinScale = Format((rsDrawItems!���ֵ - Val(strInput)) / rsDrawItems!��λֵ * Val(Split(rsDrawItems!��λ�̶�, ",")(0)), "#0.0")
        GetYCoordinate = Format(sinCurY + sinScale, "#0")
        
        If blnOutput Then
            '��ָ����������ַ����к˶�
            Call SetTextColor(lngDC, RGB_BLUE)
            Call GetTextRect(objDraw, 202, GetYCoordinate, "��", T_DrawClient.�̶ȵ�λ)
            Call DrawText(lngDC, "��", -1, T_LableRect, DT_CENTER)
        End If
    Else
        '�õ����������ֵ
        lngCurX = Split(strInput, ",")(0)
        sinCurY = Split(strInput, ",")(1)
        
        '(����-���ֵ����)/��λ�̶ȵõ������ٸ��̶�
        '(���ֵ-��λ�̶�*��λֵ)�õ�ʵ������
        sinScale = Format((sinCurY - Split(rsDrawItems!���ֵ����, ",")(1)) / Val(Split(rsDrawItems!��λ�̶�, ",")(0)), "#0.0")
        GetYCoordinate = Format(rsDrawItems!���ֵ - sinScale * rsDrawItems!��λֵ, "#0.0")
    End If

    rsDrawItems.Filter = ""

    Exit Function

Errhand:

    If ErrCenter = 1 Then

        Resume

    End If

End Function

Public Function CalcMinMaxCol(ByVal strDate As String, _
                              MinCol As Long, _
                              MaxCol As Long) As Boolean

    '------------------------------------------------------------------------------------------------------------------
    '���ܣ� �����С���ʱ�䷶Χ
    '������
    '���أ�
    '------------------------------------------------------------------------------------------------------------------
    Dim aryValue() As String

    Dim dtTmp      As Date

    Dim strTmp     As String
    
    'If mvarEdit = False Then Exit Function
    
    aryValue = Split(strDate, ";")
    
    MinCol = GetCurveColumn(CDate(aryValue(0)), CDate(aryValue(0)), gintHourBegin)
    MaxCol = GetCurveColumn(CDate(aryValue(1)), CDate(aryValue(0)), gintHourBegin)
    
End Function

Public Function ReturnItemRecord(ByVal rsCollect As ADODB.Recordset, ByVal dtDate As Date, ByVal dtBegin As Date, _
    ByVal strEditor As String, ByVal bln���ܵ��� As Boolean, Optional ByVal bln¼��Сʱ As Boolean, Optional ByVal blnEdit As Boolean = False) As ADODB.Recordset
'---------------------------------------------------------------------------------------------------------------------------------------------
'����:��ȡĳһ��ı����Ŀ��ֵ��Ϣ(�����µ�չʾ�ʹ�ӡʹ��)
'������rsCollect ��Ŀ������Ϣ,dtDate ĳ������,dtBegin ���µ���ʼ����,strEditor ��Ŀ�й���Ϣ����Ŀ���;��Ŀ����;��ĿƵ��;��Ŀ��ʾ;��Ŀ����;��Ժ�ײ�
'      bln�������� ���������ܡ�������Ŀ��ʾ(True)��������,(false)��������  blnEdit �Ƿ��Ǳ༭״̬���ڱ༭����=true��
'      bln¼��Сʱ 51282,������,2012-08-03,ȫ�������ʾ¼��ʱ�� 10.30.20(DYEYҪ���ֹ�¼�����ʱ��H)
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsDayData As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim DtDay As Date
    Dim intType As Integer, intHour As Integer, intHour1 As Integer, int��� As Integer, int��� As Integer
    Dim strBegin As String, strEnd As String, strCenter As String
    Dim strFileds As String, strValues As String, strValues1 As String, strFind As String, strTime As String
    Dim dblData As Double, intδ��˵�� As Integer, int������Դ As Integer, lngID As Long, lng��ԴID As Long, int���� As Integer
    Dim lngNO As Long
    Dim i As Integer, intCount As Integer, intColFirst As Integer, strHourTime As String
    Dim bln���� As Boolean
    'Dim bln�״λ��� As Boolean '����:���ջ�����ʾʱ��
    Dim dtCurrDate As Date
    
    '��Ŀ�й���Ϣ
    Dim lngItemNO As Long, strName As String, int��¼Ƶ�� As Integer, int��Ŀ��ʾ As Integer, int��Ŀ���� As Integer, bln��Ժ�ײ� As Boolean
    Dim arrEditor() As String
    
    On Error GoTo Errhand
    
    arrEditor = Split(strEditor, ";")
    lngItemNO = Val(arrEditor(0))
    strName = arrEditor(1)
    int��¼Ƶ�� = Val(arrEditor(2))
    int��Ŀ��ʾ = Val(arrEditor(3))
    int��Ŀ���� = Val(arrEditor(4))
    bln��Ժ�ײ� = (Val(arrEditor(5)) = 1)
    '������Ŀ��������Ժ�ײ�
    If int��Ŀ���� = 4 Then bln��Ժ�ײ� = False
    bln���� = IsWaveItem(lngItemNO) '�Ƿ��ǲ�����Ŀ
    DtDay = dtDate
    
    '��ʼ����¼��
    strFileds = "ID," & adDouble & ",18|ʱ��," & adLongVarChar & ",20|��Ŀ���," & adDouble & ",18|��Ŀ����," & adLongVarChar & ",20|��¼����," & adLongVarChar & ",100|" & _
        "���²�λ," & adLongVarChar & ",20|δ��˵��," & adLongVarChar & ",100|������Դ," & adDouble & ",1|��ʾ," & adDouble & _
        ",1|��ԴID," & adDouble & ",18|����," & adDouble & ",1|���," & adDouble & ",1|����Сʱ," & adLongVarChar & ",100"
    Call Record_Init(rsDayData, strFileds)
    strFileds = "ID|ʱ��|��Ŀ���|��Ŀ����|��¼����|���²�λ|δ��˵��|������Դ|��ʾ|��ԴID|����|���|����Сʱ"
    
    If blnEdit And bln���� Then int��Ŀ��ʾ = 0
    
ErrBegin:
    dtDate = DtDay
    rsCollect.Filter = ""
    '����/������Ŀ����=2
    If int��Ŀ��ʾ = 4 Or bln���� Then
        intType = 2
        If int��¼Ƶ�� = 0 Then
            int��¼Ƶ�� = 2
        ElseIf int��¼Ƶ�� > 2 Then
            int��¼Ƶ�� = 2
        End If
        
        '���ݲ���ȷ������/������Ŀ����ǰһ��/��������ݣ����ݻ���ʱ�Σ�
        If Not bln���ܵ��� Then dtDate = CDate(dtDate) - 1
    Else
        intType = 1
    End If
    
    '��ȡ��ǰ������ʱ��
    dtCurrDate = CDate(Format(zldatabase.Currentdate, "YYYY-MM-DD HH:mm:ss"))
    
    '�������ͣ�Ƶ�κ���� �������Ҳ�����Ϣ
    mrsTabTime.Filter = "����=" & intType & " and Ƶ��=" & int��¼Ƶ��
    If mrsTabTime.RecordCount = 0 Then
        MsgBox "���ڻ�����Ŀ����������[" & IIf(intType = 2, "������Ŀ", "���±����Ŀ") & "]ʱ����Ϣ!", vbInformation, gstrSysName
        Set ReturnItemRecord = rsDayData
        Exit Function
    End If
    
    intColFirst = 1
    
    With mrsTabTime
        .MoveFirst
        '��ȡƵ��ʱ���
        Do While Not .EOF
            int��� = Val(!���)
            int��� = Val(Nvl(!���))
            intHour = CInt(24 / int��¼Ƶ��)
            strBegin = Format(IIf(IsDate(Trim(Nvl(!��ʼ))) = False, (Val(Nvl(!���)) - 1) * intHour & ":00:00", !��ʼ), "HH:mm:ss")
            strEnd = Format(IIf(IsDate(Trim(Nvl(!����))) = False, Val(Nvl(!���)) * intHour - 1 & ":59:59", !����), "HH:mm:ss")
            'ȷ��Ƶ��ʱ�䷶Χ
            If int��� = int��¼Ƶ�� Then
                If strBegin >= strEnd Then
                    strBegin = Format(dtDate, "YYYY-MM-DD") & " " & strBegin
                    strEnd = Format(DateAdd("d", 1, CDate(dtDate)), "YYYY-MM-DD") & " " & strEnd
                Else
                    strBegin = Format(dtDate, "YYYY-MM-DD") & " " & strBegin
                    strEnd = Format(dtDate, "YYYY-MM-DD") & " " & strEnd
                End If
            Else
                If strBegin >= strEnd Then
                    strBegin = Format(dtDate, "YYYY-MM-DD") & " " & strBegin
                    strEnd = strBegin
                Else
                    strBegin = Format(dtDate, "YYYY-MM-DD") & " " & strBegin
                    strEnd = Format(dtDate, "YYYY-MM-DD") & " " & strEnd
                End If
            End If
            strBegin = Format(strBegin, "YYYY-MM-DD HH:mm:ss")
            strEnd = Format(strEnd, "YYYY-MM-DD HH:mm:ss")
            '��ȡ�е�ʱ�����Ϣ
            intHour = DateDiff("H", CDate(strBegin), CDate(strEnd) + 0.00001) / 2
            strCenter = DateAdd("H", intHour, CDate(strBegin)) '�е�ʱ��
            If CDate(strCenter) > CDate(strEnd) Then strCenter = strEnd
            
            strFind = "ʱ��>='" & Format(strBegin, "YYYY-MM-DD HH:mm:ss") & "' and ʱ��<='" & Format(strEnd, "YYYY-MM-DD HH:mm:ss") & "'"
            
            lngNO = lngItemNO
            
            If int��Ŀ���� = 2 Then
                rsCollect.Filter = "��Ŀ���=" & lngItemNO & " and ��Ŀ����='" & strName & "' And " & strFind
                If lngItemNO = 4 Then 'ѪѹΪ���Ŀ�������̶���Ŀ����
                    rsCollect.Filter = "(��Ŀ���=4 And " & strFind & ") OR (��Ŀ���=5 And " & strFind & ")"
                End If
            Else
                If lngItemNO <> 4 Then
                    rsCollect.Filter = "��Ŀ���=" & lngItemNO & " And " & strFind
                Else
                    rsCollect.Filter = "(��Ŀ���=4 And " & strFind & ") OR (��Ŀ���=5 And " & strFind & ")"
                End If
            End If
            
            rsCollect.Sort = "��Ŀ���,ʱ��"
            
            If int��Ŀ��ʾ = 4 Then '������Ŀ
                dblData = 0: intδ��˵�� = 0: strValues = ""
                If lngItemNO = 4 Then 'Ѫѹ����޸�Ϊ������Ŀ ֱ�Ӱ�����Ѫѹ����
                    int��Ŀ��ʾ = 6
                    GoTo ErrBegin
                End If
                
                '�����ǰʱ��С�ڻ���ʱ���,�����л���
                If dtCurrDate < CDate(strEnd) And Not blnEdit And Not gbln��Ժ Then GoTo ErrNext
                
                intδ��˵�� = 0: int������Դ = 0: lngID = 0: lng��ԴID = 0: int���� = 0
                strValues1 = "": intHour1 = 0: strHourTime = ""
                '��ѭ������Ŀ����
                Do While Not rsCollect.EOF
                    If Val(Nvl(rsCollect!��¼����)) = 1 Then
                        If intδ��˵�� < Val(Nvl(rsCollect!δ��˵��)) Then intδ��˵�� = Val(Nvl(rsCollect!δ��˵��))
                        If InStr(1, ",0,9,", "," & Val(Nvl(rsCollect!������Դ)) & ",") = 0 Then
                            int������Դ = Val(Nvl(rsCollect!������Դ))
                            lng��ԴID = Val(Nvl(rsCollect!��ԴID))
                            int���� = Val(Nvl(rsCollect!����))
                            lngID = Val(Nvl(rsCollect!Id))
                        ElseIf lngID = 0 Then
                            lngID = Val(Nvl(rsCollect!Id))
                        End If
                        dblData = dblData + Val(Nvl(rsCollect!���))
                    Else
                        intHour1 = -1
                        strHourTime = Format(rsCollect!ʱ��, "YYYY-MM-DD HH:mm:ss") & ";" & Val(Nvl(rsCollect!Id))
                        strValues1 = Val(Nvl(rsCollect!���))
                        If Val(strValues1) < 0 Then strValues1 = ""
                        If Val(strValues1) > 24 Then strValues1 = 24
                    End If
                rsCollect.MoveNext
                Loop
                
                If rsCollect.RecordCount > 0 Then rsCollect.MoveFirst
                
                If int��Ŀ���� = 2 Then
                    '���Ŀ����λͳ������
                Else
                    '��ʼ��������Ŀ
                    Set rsTemp = SetCollectPItem(lngItemNO)
                    rsTemp.Filter = 0
                    Do While Not rsTemp.EOF
                        '����ͬ�������������� ���ڸ����Ѿ������� �˴����ٽ��л���
                        If Val(Nvl(rsTemp!���, 0)) <> lngItemNO Then
                            rsCollect.Filter = 0
                            rsCollect.Filter = "��Ŀ���=" & Val(Nvl(rsTemp!���, 0)) & " And ������Դ<>9 And ��¼����=1 " & " And " & strFind
                            Do While Not rsCollect.EOF
                                dblData = dblData + Val(Nvl(rsCollect!���))
                                If lng��ԴID = 0 Then
                                    If InStr(1, ",0,9,", "," & Val(Nvl(rsCollect!������Դ)) & ",") = 0 Then
                                        int������Դ = Val(Nvl(rsCollect!������Դ))
                                        lng��ԴID = Val(Nvl(rsCollect!��ԴID))
                                        int���� = Val(Nvl(rsCollect!����))
                                        lngID = Val(Nvl(rsCollect!Id))
                                    ElseIf lngID = 0 Then
                                        lngID = Val(Nvl(rsCollect!Id))
                                    End If
                                End If
                            rsCollect.MoveNext
                            Loop
                            If rsCollect.RecordCount > 0 Then rsCollect.MoveFirst
                        End If
                        rsTemp.MoveNext
                    Loop
                End If
                If lngID <> 0 Then
                    If bln¼��Сʱ = True And int��¼Ƶ�� = 1 Then
                        strValues1 = IIf(dblData = 0, "", IIf(strValues1 = "", "", "(" & strValues1 & "h)") & dblData)
                    Else
                        intHour1 = 0
                        strValues1 = IIf(dblData = 0, "", dblData)
                    End If
                    '51282,������,2012-07-11
                    '51282,������,2012-08-03,DYEYĿǰҪ��ȫ����ܿ����ֹ�¼�����Сʱ
                    'ȫ������״β�������ʱ����ʾ����ʱ��Сʱ�����硰������������200ml,����ͳ��ʱ��Ϊ18h ,�����Ӧ����ʾΪ��(18h)200��
'                    If blnEdit = False And int��¼Ƶ�� = 1 And (Format(dtBegin, "YYYY-MM-DD") = Format(dtDate, "YYYY-MM-DD") Or Format(dtBegin, "YYYY-MM-DD") = Format(DtDay, "YYYY-MM-DD")) Then
'                        bln�״λ��� = (Val(zlDatabase.GetPara("���ջ�����ʾʱ��", glngSys, 1255, "0")) = 1)
'                        '�������ʱ������Сʱ��
'                        intHour1 = Format(DateDiff("n", CDate(strBegin), CDate(strEnd) + 0.00001) / 60, "#0")
'                        If bln���ܵ��� = True And bln�״λ��� = True Then
'                            '���ܵ���ֻ�������µ����죬����ʱ�ο϶��ǵ��쿪ʼ���ڶ��������������µ���ʼʱ��ͻ��ܽ���ʱ�������Сʱ��ֻ��
'                            '����0����С�ڻ���ʱ�μ����Сʱ����������
'                            If Format(dtBegin, "YYYY-MM-DD") = Format(dtDate, "YYYY-MM-DD") Then
'                                '�������µ���ʼʱ��ͻ��ܽ���ʱ��������Сʱ
'                                intHour = Format(DateDiff("n", CDate(dtBegin), CDate(strEnd) + 0.00001) / 60, "#0")
'                                If intHour > 0 And intHour < intHour1 Then strValues1 = "(" & intHour & "h)" & strValues1
'                            End If
'                        ElseIf bln�״λ��� = True Then '������Ŀ�������죬�������������һ�������µ��Ŀ�ʼʱ���ڵ�һ�����ʱ���ڣ�һ�������µ��Ŀ�ʼʱ�䲻�ڵ�һ�����ʱ����
'                            '�������ڵڶ������ʱ���ڣ�Ҳ���ܲ��ڣ������������ֻ��������һ��
'                            If Format(dtBegin, "YYYY-MM-DD") = Format(DtDay, "YYYY-MM-DD") Then
'                                '�������µ���ʼʱ��ͻ��ܽ���ʱ��������Сʱ
'                                intHour = Format(DateDiff("n", CDate(dtBegin), CDate(strEnd) + 0.00001) / 60, "#0")
'                                If intHour > 0 And intHour < intHour1 Then strValues1 = "(" & intHour & "h)" & strValues1
'                            End If
'
'                            If Format(dtBegin, "YYYY-MM-DD") = Format(dtDate, "YYYY-MM-DD") Then
'                                '�������µ���ʼʱ��ͻ��ܽ���ʱ��������Сʱ
'                                intHour = Format(DateDiff("n", CDate(dtBegin), CDate(strEnd) + 0.00001) / 60, "#0")
'                                If intHour > 0 And intHour < intHour1 Then strValues1 = "(" & intHour & "h)" & strValues1
'                            End If
'                        End If
'                    End If
                    If int��Ŀ���� = 2 Then
                        strValues = lngID & "|" & CDate(strCenter) & "|" & lngItemNO & "|" & strName & "|" & _
                                        strValues1 & "|" & strName & "|" & intδ��˵�� & "|" & _
                                        int������Դ & "|" & 1 & "|" & lng��ԴID & "|" & int���� & "|" & int��� & "|" & intHour1 & ";" & strHourTime
                    Else
                        strValues = lngID & "|" & CDate(strCenter) & "|" & lngItemNO & "|" & strName & "|" & _
                                    strValues1 & "|" & "" & "|" & "" & "|" & _
                                    int������Դ & "|" & 1 & "|" & lng��ԴID & "|" & int���� & "|" & int��� & "|" & intHour1 & ";" & strHourTime
                    End If
                    Call Record_Add(rsDayData, strFileds, strValues)
                    strValues1 = ""
                End If
            ElseIf bln���� Then '������Ŀ
                intCount = 0: i = 0
                If lngNO = 4 Then intCount = 1
                
                If bln��Ժ�ײ� = True And Format(dtBegin, "YYYY-MM-DD") = Format(dtDate, "YYYY-MM-DD") And intColFirst = 1 Then 'dtBegin >= CDate(strBegin) And dtBegin <= CDate(strEnd) Then
                    int��� = 1 '��ȡ��һ������
                    GoTo ErrRead
                End If
                
                '�����ǰʱ��С�ڻ���ʱ���,�����л���
                If dtCurrDate < CDate(strEnd) And Not blnEdit And Not gbln��Ժ Then GoTo ErrNext
                
                For i = 0 To intCount
                    If i = 1 Then lngNO = 5
                    If intCount = 1 Then 'Ѫѹ��Ŀ������ȡ
                        rsCollect.Filter = 0
                        rsCollect.Filter = "��Ŀ���=" & lngNO & " And " & strFind
                    End If
                    strValues = "": strValues1 = "": strTime = "": dblData = 0
                    Do While Not rsCollect.EOF
                        If dblData <> 0 Then
                            '��ȡ���ֵ
                            If Val(strValues) < Val(Nvl(rsCollect!���)) Then
                                strValues = Val(Nvl(rsCollect!���))
                            End If
                            '��ȡ��Сֵ
                            If Val(strValues1) > Val(Nvl(rsCollect!���)) Then
                                strValues1 = Val(Nvl(rsCollect!���))
                            End If
                        Else
                            dblData = 99
                            If IsNumeric(Nvl(rsCollect!���)) Then
                                strValues = Val(Nvl(rsCollect!���))
                                strValues1 = strValues
                            Else
                                strValues = ""
                                strValues1 = ""
                            End If
                            
                            lngID = Val(Nvl(rsCollect!Id))
                            int������Դ = Val(Nvl(rsCollect!������Դ))
                            lng��ԴID = Val(Nvl(rsCollect!��ԴID))
                            int���� = Val(Nvl(rsCollect!����))
                            strTime = Nvl(rsCollect!ʱ��)
                        End If
                        rsCollect.MoveNext
                    Loop
                    
                    If dblData <> 0 Then
                        If Val(strValues) <> Val(strValues1) Then
                            strValues1 = Val(strValues1) & "-" & Val(strValues)
                        Else
                            strValues1 = IIf(strValues = "", "", Val(strValues))
                        End If
                        
                        '��������浽��¼����
                        strValues = lngID & "|" & CDate(strTime) & "|" & lngNO & "|" & IIf(lngItemNO <> 4, strName, IIf(lngNO = 4, "����ѹ", "����ѹ")) & "|" & _
                            strValues1 & "|" & "" & "|" & "" & "|" & int������Դ & "|" & _
                            1 & "|" & lng��ԴID & "|" & int���� & "|" & int��� & "|0"
                        Call Record_Add(rsDayData, strFileds, strValues)
                    End If
                Next i
            Else '�ǻ�����Ŀ
                intCount = 0: i = 0
                '--����Ѫѹ��Ҫ�ֱ�������ѹ������ѹ
                If lngNO = 4 Then intCount = 1
                
                If bln��Ժ�ײ� = True And Format(dtBegin, "YYYY-MM-DD") = Format(dtDate, "YYYY-MM-DD") And intColFirst = 1 Then 'dtBegin >= CDate(strBegin) And dtBegin <= CDate(strEnd) Then
                    int��� = 1 '��ȡ��һ������
                End If
ErrRead:
                For i = 0 To intCount
                    If i = 1 Then lngNO = 5
                    If intCount = 1 Then 'Ѫѹ��Ŀ���¹���
                        rsCollect.Filter = 0
                        rsCollect.Filter = "��Ŀ���=" & lngNO & " And " & strFind
                    End If
                    strValues = "": strValues1 = "": strTime = ""
                    Do While Not rsCollect.EOF
                        intColFirst = 2
                        '����Ѫѹ���д���
                        If lngNO = Val(Nvl(rsCollect!��Ŀ���)) Then
                            Select Case int���
                                Case 1 '��һ��
                                    If rsCollect.RecordCount > 0 Then rsCollect.MoveFirst
                                        strValues = Val(Nvl(rsCollect!Id)) & "|" & CDate(rsCollect!ʱ��) & "|" & Val(Nvl(rsCollect!��Ŀ���)) & "|" & Nvl(rsCollect!��Ŀ����) & "|" & _
                                            Nvl(rsCollect!���) & "|" & Nvl(rsCollect!���²�λ) & "|" & Nvl(rsCollect!δ��˵��) & "|" & Val(Nvl(rsCollect!������Դ)) & "|" & _
                                            Val(Nvl(rsCollect!��ʾ)) & "|" & Val(Nvl(rsCollect!��ԴID)) & "|" & Val(Nvl(rsCollect!����)) & "|" & int��� & "|0"
                                    Exit Do
                                Case 2 '�м�һ��
                                    strValues = Val(Nvl(rsCollect!Id)) & "|" & CDate(rsCollect!ʱ��) & "|" & Val(Nvl(rsCollect!��Ŀ���)) & "|" & Nvl(rsCollect!��Ŀ����) & "|" & _
                                            Nvl(rsCollect!���) & "|" & Nvl(rsCollect!���²�λ) & "|" & Nvl(rsCollect!δ��˵��) & "|" & Val(Nvl(rsCollect!������Դ)) & "|" & _
                                            Val(Nvl(rsCollect!��ʾ)) & "|" & Val(Nvl(rsCollect!��ԴID)) & "|" & Val(Nvl(rsCollect!����)) & "|" & int��� & "|0"
                                    If strValues1 <> "" Then
                                        '����Ǹ��ӽ��е�ʱ��
                                        If Abs(DateDiff("s", Format(CDate(rsCollect!ʱ��), "YYYY-MM-DD HH:mm:ss"), Format(strCenter, "YYYY-MM-DD HH:mm:ss"))) > _
                                            Abs(DateDiff("s", Format(CDate(strTime), "YYYY-MM-DD HH:mm:ss"), Format(strCenter, "YYYY-MM-DD HH:mm:ss"))) Then
                                             strValues = strValues1
                                        End If
                                    End If
                                    strValues1 = strValues
                                    strTime = rsCollect!ʱ��
                                Case Else '���һ��
                                    If rsCollect.RecordCount > 0 Then rsCollect.MoveLast
                                        strValues = Val(Nvl(rsCollect!Id)) & "|" & CDate(rsCollect!ʱ��) & "|" & Val(Nvl(rsCollect!��Ŀ���)) & "|" & Nvl(rsCollect!��Ŀ����) & "|" & _
                                            Nvl(rsCollect!���) & "|" & Nvl(rsCollect!���²�λ) & "|" & Nvl(rsCollect!δ��˵��) & "|" & Val(Nvl(rsCollect!������Դ)) & "|" & _
                                            Val(Nvl(rsCollect!��ʾ)) & "|" & Val(Nvl(rsCollect!��ԴID)) & "|" & Val(Nvl(rsCollect!����)) & "|" & int��� & "|0"
                                    Exit Do
                            End Select
                        End If
                    rsCollect.MoveNext
                    Loop
                    '��Ӽ�¼����Ϣ
                    If strValues <> "" Then
                        Call Record_Add(rsDayData, strFileds, strValues)
                    End If
                    'Call OutputRsData(rsDayData, True)
                Next i
            End If
ErrNext:
        .MoveNext
        Loop
    End With
    
    Set ReturnItemRecord = rsDayData
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub InitPublicData()
    On Error GoTo Errhand
    
    If Not (mrsTabTime Is Nothing) Then If mrsTabTime.State = 1 Then mrsTabTime.Close
    '�����б����Ŀʱ����Ϣ
    gstrSQL = "SELECT ���, ��ʼ, ����, Ƶ��,���, ����" & vbNewLine & _
                "  FROM (SELECT DECODE(���, 3, 1, ���) ���," & vbNewLine & _
                "               ��ʼ || ':00' ��ʼ," & vbNewLine & _
                "               ���� || ':59' ����," & vbNewLine & _
                "               DECODE(���, 3, 1, 2) Ƶ��,0 ���," & vbNewLine & _
                "               2 ����" & vbNewLine & _
                "          FROM �������ʱ�� WHERE ����=1" & vbNewLine & _
                "        UNION ALL" & vbNewLine & _
                "        SELECT ���, ��ʼ || ':00' ��ʼ, ���� || ':59' ����, Ƶ��,���, 1 ����" & vbNewLine & _
                "          FROM ������ĿƵ��)" & vbNewLine & _
                " ORDER BY ����, Ƶ��, ���"

    Call zldatabase.OpenRecordset(mrsTabTime, gstrSQL, "���µ�")
    
    If Not (mrsCollect Is Nothing) Then If mrsCollect.State = 1 Then mrsCollect.Close
    '��ȡ���������Ŀ
    gstrSQL = " SELECT ���,����� FROM ���������Ŀ"
    Call zldatabase.OpenRecordset(mrsCollect, gstrSQL, "���������Ŀ")
    
    If Not (mrsWave Is Nothing) Then If mrsWave.State = 1 Then mrsWave.Close
    '��������Ŀ
    gstrSQL = "��SELECT ��Ŀ��� FROM ��������Ŀ"
    Call zldatabase.OpenRecordset(mrsWave, gstrSQL, "��������Ŀ")
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function SetCollectPItem(ByVal lngItemNO As Long) As ADODB.Recordset
'---------------------------------------------------------------------------
'����:���ݸ���ĿID������֯����Ŀ
'---------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim rsCollect As New ADODB.Recordset
    Dim strFileds As String, strValues As String
    Dim lngNO As Long
    
    On Error GoTo Errhand
    
    '��ʼ����¼��
    strFileds = "���," & adDouble & ",18|�����," & adDouble & ",18"
    Call Record_Init(rsTemp, strFileds)
    Call Record_Init(rsCollect, strFileds)
    strFileds = "���|�����"
    
    mrsCollect.Filter = 0
   '���Ƽ�¼��
    With mrsCollect
        Do While Not .EOF
            strValues = Val(Nvl(!���)) & "|" & Val(Nvl(!�����))
            Call Record_Add(rsCollect, strFileds, strValues)
            .MoveNext
        Loop
    End With
    
    rsCollect.Filter = "�����=" & lngItemNO
    With rsCollect
        Do While Not .EOF
            strValues = Val(Nvl(!���)) & "|" & lngItemNO
            Call Record_Add(rsTemp, strFileds, strValues)
            lngNO = Val(Nvl(!���))
            'ѭ���ݹ����(��ȡ���������)
            Call SetCollectCItem(rsTemp, lngItemNO, lngNO)
            .MoveNext
        Loop
    End With
    
    Set SetCollectPItem = rsTemp
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub SetCollectCItem(rsTemp As ADODB.Recordset, ByVal lngParent As Long, ByVal lngItemNO As Long)
'����: SetCollectPItem ����
    
    Dim rsCollect As New ADODB.Recordset
    Dim strFileds As String, strValues As String
    Dim lngNO As Long
    
    '��ʼ����¼��
    strFileds = "���," & adDouble & ",18|�����," & adDouble & ",18"
    Call Record_Init(rsCollect, strFileds)
    strFileds = "���|�����"
    
    mrsCollect.Filter = 0
   '���Ƽ�¼��
    With mrsCollect
        Do While Not .EOF
            strValues = Val(Nvl(!���)) & "|" & Val(Nvl(!�����))
            Call Record_Add(rsCollect, strFileds, strValues)
            .MoveNext
        Loop
    End With
    
    rsCollect.Filter = "�����=" & lngItemNO
    With rsCollect
        Do While Not .EOF
            strValues = Val(Nvl(!���)) & "|" & lngParent
            Call Record_Add(rsTemp, strFileds, strValues)
            lngNO = Val(Nvl(!���))
            'ѭ���ݹ����(��ȡ���������)
            Call SetCollectCItem(rsTemp, lngParent, lngNO)
            .MoveNext
        Loop
    End With
End Sub

Public Function IsWaveItem(ByVal lngItemNO As Long) As Boolean
'����Ƿ��ǲ�����Ŀ
    If mrsWave Is Nothing Then Exit Function
    If mrsWave.State = 1 Then
        mrsWave.Filter = 0
        mrsWave.Filter = "��Ŀ���=" & lngItemNO
        IsWaveItem = (mrsWave.RecordCount > 0)
    End If
End Function

Public Function SetNTPrinterPaper(ByVal lngHwnd As Long, ByVal intWidth As Integer, ByVal intHeight As Integer, _
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
        lngSize = DocumentProperties(lngHwnd, lngHandle, strPrtName, 0&, 0&, 0&)
        'Reserve memory for the actual size of the DEVMODE.
        ReDim arrDevMode(1 To lngSize)
    
        'Fill the DEVMODE from the printer.
        lngSize = DocumentProperties(lngHwnd, lngHandle, strPrtName, arrDevMode(1), 0&, DM_OUT_BUFFER)
        'Copy the Public (predefined) portion of the DEVMODE.
        Call CopyMemory(vDevMode, arrDevMode(1), Len(vDevMode))
        
        '���ô�ӡ�ĵ�����
        vDevMode.dmOrientation = intOrient
        vDevMode.dmPaperSize = 256
        vDevMode.dmPaperWidth = intWidth * 10 'in tenths of a millimeter
        vDevMode.dmPaperLength = intHeight * 10 'in tenths of a millimeter
        vDevMode.dmCopies = intCopys
        'vDevMode.dmCollate = 0& '�߼���ӡ����(��ȡ��ʱ,Copiesֻ֧��1;����֪��ôȡ����)
        vDevMode.dmFields = DM_ORIENTATION Or DM_PAPERSIZE Or DM_PAPERLENGTH Or DM_PAPERWIDTH Or DM_COPIES 'Or DM_COLLATE
        
        'Copy your changes back, then update DEVMODE.
        Call CopyMemory(arrDevMode(1), vDevMode, Len(vDevMode))
        If blnPrompt Then
            lngSize = DocumentProperties(lngHwnd, lngHandle, strPrtName, arrDevMode(1), arrDevMode(1), DM_IN_BUFFER Or DM_IN_PROMPT Or DM_OUT_BUFFER)
        Else
            lngSize = DocumentProperties(lngHwnd, lngHandle, strPrtName, arrDevMode(1), arrDevMode(1), DM_IN_BUFFER Or DM_OUT_BUFFER)
        End If
        If lngSize = IDOK Then SetNTPrinterPaper = True
        'Reset the DEVMODE for the DC.
        lngSize = ResetDC(lngPrtDC, arrDevMode(1))
        If lngSize = 0 Then SetNTPrinterPaper = False
        
        'Close the handle when you are finished with it.
        Call ClosePrinter(lngHandle)
    End If
End Function

Public Function SetCustonPager(ByVal lngHwnd As Long, ByVal lngWidth As Long, ByVal lngHeight As Long) As Integer
'���ܣ��������Զ���ֽ��
'�����������Ϊ��λ
    If IsWindowsNT Then
        '��Ȼ����ʹ�����Ч�����ܸı�PaperSize������ֵ
        Printer.Width = lngWidth
        Printer.Height = lngHeight
        SetCustonPager = SetNTPrinterPaper(lngHwnd, lngWidth / conRatemmToTwip, lngHeight / conRatemmToTwip, Printer.Orientation, Printer.Copies)
    Else
        'Windows98ϵ�л�����ͨ����������
        Printer.PaperSize = 256
        Printer.Width = lngWidth
        Printer.Height = lngHeight
    End If
End Function

Public Function GetTimeColor(ByVal intHour As Integer) As Long
'---------------------------------------------
'���ݲ�����ȡ����ʱ����ɫ
'---------------------------------------------
    Dim blnTag As Boolean
    Dim strTmp As String
    Dim lngBegin As Long, lngEnd As Long
    Dim lngColor As Long
    strTmp = zldatabase.GetPara("����ʱ��ҹ���־", glngSys, 1255, "18;6")
    If UBound(Split(strTmp, ";")) >= 1 Then
        lngBegin = Abs(Val(Split(strTmp, ";")(0)))
        lngEnd = Abs(Val(Split(strTmp, ";")(1)))
    Else
        lngBegin = Abs(Val(strTmp))
    End If
    
    If lngBegin < lngEnd Then
        blnTag = (intHour >= lngBegin And intHour < lngEnd)
    Else
        blnTag = (intHour >= lngBegin Or intHour < lngEnd)
    End If
    If blnTag = True Then
        lngColor = RGB_RED
    Else
        lngColor = RGB_BLACK
    End If
    
    GetTimeColor = lngColor
End Function
