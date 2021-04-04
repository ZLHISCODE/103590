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

Public Const conRatemmToTwip As Single = 56.6857142857143      '毫米与缇的比率
Public Const mintNullRow As Integer = 1 '体温刻度上面空行
Private msngTwips As Single 'Screen.TwipsPerPixelX /printer.TwipsPerPixelX
Public gintEditorCurveState As Integer '记录体温单是编辑曲线还是编辑表格
Private mfrmTendBody As Object
Private mlng体温不升显示方式 As Long

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
Public gblnPrinted As Boolean           '是否打印了体温单
Public gintHourBegin As Integer '体温单开始时间
Public gstrCaveSplit As String '体温单标志于时间之间的连接方式:例如.入院于九时...或入院--九时
Public gvarTime As Variant
Public gstdSet As New StdFont  '设置字体
Public gbln出院 As Boolean  '病人是否出院

Private mintBaby As Integer  '是否是婴儿
Public Const gint心率 As Integer = -1
Public Const gint体温 As Integer = 1
Public Const gint脉搏 As Integer = 2
Public Const gint呼吸 As Integer = 3
Public Const gint大便 As Integer = 10
Public Const gint入液 As Integer = 9
Public Const gintBmpW As Integer = 12
Public Const gintBmpH As Integer = 12
Public Const glngMaxRows As Long = 80   '总列数
Public Const glngLableStep As Long = 30  '刻度区域列宽
Public Const glngLableWith As Long = 90 '刻度区域 曲线数据<=3时的默认总宽度
Public Const glngColStep As Long = 16   '体温区域列宽
Public Const glngInitRowStep As Long = 6 '体温区域列高
Public Const p住院护士站 As Long = 1262  '住院护士站参数
Public mbln呼吸曲线 As Boolean
Public glngCurPage As Long
Public mintBmpW As Integer
Public mintBmpH As Integer


Public RGB_BLACK          As Long
Public RGB_RED            As Long
Public RGB_WRITE          As Long
Public RGB_BLUE          As Long
Public RGB_GRAY          As Long
Public RGB_FleetGRAY     As Long

Public mrsTabTime As New ADODB.Recordset '体温表格项目时间段
Public mrsCollect As New ADODB.Recordset '体温汇总项目
Public mrsWave As New ADODB.Recordset  '体温波动项目

Public Type DrawClient
    偏移量X As Long
    偏移量Y As Long
    刻度区域 As RECT
    刻度单位 As Long
    体温区域 As RECT
    行单位 As Single
    时间行单位 As Single
    时间列单位 As Single
    列单位 As Long
    双倍 As Boolean '一行表示贰行？
    总列数 As Long
End Type

Public T_DrawClient As DrawClient

'--颜色
Private Enum Color
    黑色 = 0
    深灰色 = &H404040
    灰色 = &HE0E0E0
    红色 = 200
End Enum

Private Type BODYFLAG
    入院 As Byte
    入科 As Byte
    转出 As Byte
    换床 As Byte
    手术 As Byte
    出院 As Byte
    分娩 As Byte
    出生 As Byte
End Type

Public T_BodyFlag As BODYFLAG


Private Type TwipsPerPixel
    X As Single
    Y As Single
End Type
Public T_TwipsPerPixel As TwipsPerPixel

'打印是表下表格使用,以便于缩放字体后重新计算坐标位置
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

Public Const DC_PAPERNAMES = 16 '纸张名称(每64字符为一段,以Chr(0)结束)
Public Const DC_PAPERS = 2 '纸张编号(Array or Word)
Public Const DC_BINNAMES = 12 '进纸方式(每24字符为一段,以Chr(0)结束)
Public Const DC_BINS = 6 '进纸编号(Array or Word)

'打印纸张常量(256=自定义)
Public Const PageSize1 = "信笺， 8 1/2 x 11 英寸"
Public Const PageSize2 = "+A611 小型信笺， 8 1/2 x 11 英寸"
Public Const PageSize3 = "小型报， 11 x 17 英寸"
Public Const PageSize4 = "分类帐， 17 x 11 英寸"
Public Const PageSize5 = "法律文件， 8 1/2 x 14 英寸"
Public Const PageSize6 = "声明书，5 1/2 x 8 1/2 英寸"
Public Const PageSize7 = "行政文件，7 1/2 x 10 1/2 英寸"
Public Const PageSize8 = "A3, 297 x 420 毫米"
Public Const PageSize9 = "A4, 210 x 297 毫米"
Public Const PageSize10 = "A4小号， 210 x 297 毫米"
Public Const PageSize11 = "A5, 148 x 210 毫米"
Public Const PageSize12 = "B4, 250 x 354 毫米"
Public Const PageSize13 = "B5, 182 x 257 毫米"
Public Const PageSize14 = "对开本， 8 1/2 x 13 英寸"
Public Const PageSize15 = "四开本， 215 x 275 毫米"
Public Const PageSize16 = "10 x 14 英寸"
Public Const PageSize17 = "11 x 17 英寸"
Public Const PageSize18 = "便条，8 1/2 x 11 英寸"
Public Const PageSize19 = "#9 信封， 3 7/8 x 8 7/8 英寸"
Public Const PageSize20 = "#10 信封， 4 1/8 x 9 1/2 英寸"
Public Const PageSize21 = "#11 信封， 4 1/2 x 10 3/8 英寸"
Public Const PageSize22 = "#12 信封， 4 1/2 x 11 英寸"
Public Const PageSize23 = "#14 信封， 5 x 11 1/2 英寸"
Public Const PageSize24 = "C 尺寸工作单"
Public Const PageSize25 = "D 尺寸工作单"
Public Const PageSize26 = "E 尺寸工作单"
Public Const PageSize27 = "DL 型信封， 110 x 220 毫米"
Public Const PageSize28 = "C5 型信封， 162 x 229 毫米"
Public Const PageSize29 = "C3 型信封， 324 x 458 毫米"
Public Const PageSize30 = "C4 型信封， 229 x 324 毫米"
Public Const PageSize31 = "C6 型信封， 114 x 162 毫米"
Public Const PageSize32 = "C65 型信封，114 x 229 毫米"
Public Const PageSize33 = "B4 型信封， 250 x 353 毫米"
Public Const PageSize34 = "B5 型信封，176 x 250 毫米"
Public Const PageSize35 = "B6 型信封， 176 x 125 毫米"
Public Const PageSize36 = "信封， 110 x 230 毫米"
Public Const PageSize37 = "信封大王， 3 7/8 x 7 1/2 英寸"
Public Const PageSize38 = "信封， 3 5/8 x 6 1/2 英寸"
Public Const PageSize39 = "U.S. 标准复写簿， 14 7/8 x 11 英寸"
Public Const PageSize40 = "德国标准复写簿， 8 1/2 x 12 英寸"
Public Const PageSize41 = "德国法律复写簿， 8 1/2 x 13 英寸"

Public Const conBin1 = "上层纸盒进纸"
Public Const conBin2 = "下层纸盒进纸"
Public Const conBin3 = "中间纸盒进纸"
Public Const conBin4 = "等待手动插入每页纸"
Public Const conBin5 = "信封进纸器进纸"
Public Const conBin6 = "信封进纸器进纸；但要等待手动插入"
Public Const conBin7 = "当前缺省纸盒进纸"
Public Const conBin8 = "拖拉进纸器进纸"
Public Const conBin9 = "小型进纸器进纸"
Public Const conBin10 = "大型纸盒进纸"
Public Const conBin11 = "大容量进纸器进纸"

'纸张打印边界控制================================================================
Public Declare Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, ByVal lpPort As String, ByVal iIndex As Long, ByVal lpOutput As String, lpDevMode As Any) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
'不同打印机的打印单元精度不同

'Public Const PHYSICALHEIGHT = 111  'Physical Height in device units

'Public Const PHYSICALOFFSETY = 113 'Physical Printable Area y margin
Public Const LOGPIXELSX = 88 'Number of pixels per logical inch along the screen width
Public Const LOGPIXELSY = 90
Public Const SCALINGFACTORX = 114  'Scaling factor x
Public Const SCALINGFACTORY = 115  'Scaling factor y
Public Const DRIVERVERSION = 0     'Device driver version

'WinNT自定义纸张控制================================================================
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
'功能：取大于指定数值的最小整数
    IntEx = -1 * Int(-1 * Val(vNumber))
End Function

Public Function GetPaperName(intSize As Integer) As String
    '功能： 根据当前打印机的设置，获取纸张名称
    '返回： 纸张名称
    If intSize = 256 Then
        GetPaperName = "用户自定义 ..."
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

Public Function InitPrint(objParent As Object) As Boolean
    '功能：根据注册表及frmparent.mobjreport内容初始化打印机设置(本地->服务器->当前)
    '返回：如果无打印机或纸张不对,则失败
    Dim i As Integer, strPName As String
    Dim strPrinter As String  '打印机
    Dim intPage As Integer  '纸张
    Dim lngWidth As Long  '自定义纸张宽度
    Dim lngHeight As Long  '自定义纸张高度
    Dim intOrient As Byte  '纸向
    Dim intBin As Integer  '进纸方式
    If Not ExistsPrinter Then Exit Function
    
    '初始化打印参数
    
    strPrinter = Trim(zldatabase.GetPara("体温单打印机", glngSys, 1255, Printer.DeviceName))
    intPage = Val(zldatabase.GetPara("体温单纸张", glngSys, 1255, Printer.PaperSize))
    lngWidth = Val(zldatabase.GetPara("体温单宽度", glngSys, 1255, Printer.Width))
    lngHeight = Val(zldatabase.GetPara("体温单高度", glngSys, 1255, Printer.Height))
    intOrient = Val(zldatabase.GetPara("体温单纸向", glngSys, 1255, Printer.Orientation))
    intBin = Val(zldatabase.GetPara("体温单进纸", glngSys, 1255, Printer.PaperBin))

    
    '打印机
    If Printer.DeviceName <> strPName Then
        For i = 0 To Printers.Count - 1
            If Printers(i).DeviceName = strPrinter Then Set Printer = Printers(i): Exit For
        Next
    End If
    On Error Resume Next
    '纸张
    If intPage = 256 Then
        Printer.PaperSize = 256
        Printer.Width = lngWidth
        Printer.Height = lngHeight
    Else
        Printer.PaperSize = intPage
    End If
    '纸向
    '纸向赋值后,纸张宽高值交换,纸向还原为1
    Printer.Orientation = intOrient
    '进纸
    Printer.PaperBin = intBin
    '份数
    Printer.Copies = 1
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    'WinNT自定义纸张处理
    If IsWindowsNT And intPage = 256 Then
        If AddCustomPaper(objParent.hWnd, lngWidth / conRatemmToTwip, lngHeight / conRatemmToTwip) = FORM_NOT_SELECTED Then Exit Function
    End If
    InitPrint = True
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
    '功能：创建一个NT下使用的自定义纸张
    '参数：宽、度=mm(毫米)
    Dim lngSize As Long ' Size of DEVMODE
    Dim vDevMode As DEVMODE
    Dim arrDevMode() As Byte ' Working DEVMODE
    
    Dim lngHandle As Long 'Handle to printer
    Dim lngPrtDC As Long ' Handle to Printer DC
    Dim strPrtName As String
    
    Dim vFormSize As SIZEL
    
    strPrtName = Printer.DeviceName
    lngPrtDC = Printer.hDC
    
    If OpenPrinter(strPrtName, lngHandle, 0&) Then '获取打印机句柄
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
    '功能：删除刚才创建的自定义纸张
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
    '功能：关闭Recordset对象
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
            MsgBox "你没有权限设置打印机""" & Printer.DeviceName & """为自定义尺寸，打印结果可能会不正常！", vbExclamation, App.Title
        Else
            MsgBox "设置打印机纸张时发生错误，编号： " & Err.LastDllError, vbExclamation, App.Title
        End If
        AddNewForm = ""
    Else
        AddNewForm = FI1.pName
    End If
End Function

Public Sub ShowFlash(Optional strInfo As String, Optional sngPer As Single, Optional frmParent As Object)
    '功能：显示或隐藏等待或进度窗体(strInfo)
    '参数:strInfo=进度提示信息
    '     sngPer=进度
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
    '功能：打印或预览一页结束时对当前页作结束处理,并产生新页
    '参数：blnNewPage=为False时仅打印页号等,一般打印结束才这样处理,因此不管最后坐标
    '返回：新页对象,可能为打印机或PictureBox
    On Error GoTo errH
    Dim objDraw As Object, blnPrint As Boolean
    Dim lngWidth As Long, lngHeight As Long, lngOldY As Long
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    Dim strFontName As String, lngFontSize As Long, blnFontBold As Boolean
    Dim blnFontItalic As Boolean, lngFontColor As Long
    
    
    blnPrint = TypeName(objOut) = "Printer"
    '边界信息(Twip)
    
    lngLeft = Val(zldatabase.GetPara("体温单左边距", glngSys, 1255, OFFSET_LEFT)) * conRatemmToTwip
    lngRight = Val(zldatabase.GetPara("体温单右边距", glngSys, 1255, OFFSET_RIGHT)) * conRatemmToTwip
    lngTop = Val(zldatabase.GetPara("体温单上边距", glngSys, 1255, OFFSET_TOP)) * conRatemmToTwip
    lngBottom = Val(zldatabase.GetPara("体温单下边距", glngSys, 1255, OFFSET_BOTTOM)) * conRatemmToTwip
    
    
    lngWidth = Printer.Width: lngHeight = Printer.Height
    
    '一页处理结束后的处理
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
    
    '打印页号(0为不打印)
    If intPage <> 0 Then
        objDraw.ForeColor = 0
        objDraw.Font.Name = "宋体"
        objDraw.Font.Size = 9
        objDraw.Font.Bold = False
        objDraw.CurrentY = lngHeight - lngBottom - objDraw.TextHeight("字")
        objDraw.CurrentX = lngLeft + (lngWidth - lngLeft - lngRight) * (3 / 4)
        objDraw.FontTransparent = True
        objDraw.Print "·第 " & intPage & " 页·"
    End If
    
    If Not blnPrint Then
        '预览打印边线
        objDraw.DrawStyle = 2
        objDraw.Line (0, lngTop)-(lngWidth, lngTop), &H808080
        objDraw.Line (0, lngHeight - lngBottom)-(lngWidth, lngHeight - lngBottom), &H808080
        objDraw.Line (lngLeft, 0)-(lngLeft, lngHeight), &H808080
        objDraw.Line (lngWidth - lngRight, 0)-(lngWidth - lngRight, lngHeight), &H808080
        objDraw.DrawStyle = 0
    End If
    
    '产生新页
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
        '新页起点
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
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Set mfrmTendBody = frmObject
    
    glngPrevWndProc = SetWindowLong(frmObject.hWnd, GWL_WNDPROC, AddressOf WindowProc)

    '获取"控制面板"中的滚动行数值
    Call SystemParametersInfo(SPI_GETWHEELSCROLLLINES, 0, WHEEL_SCROLL_LINES, 0)

    If WHEEL_SCROLL_LINES > frmObject.BodyEdit.ScrollBarY.Max Then WHEEL_SCROLL_LINES = frmObject.BodyEdit.ScrollBarY.Max
End Sub

Public Sub UnHook(ByVal frmObject As Object)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngReturnValue As Long

    lngReturnValue = SetWindowLong(frmObject.hWnd, GWL_WNDPROC, glngPrevWndProc)
    Set mfrmTendBody = Nothing
End Sub

Public Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    '******************************************************************************************************************
    '功能：捕获系统事件并进行处理
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim pt As POINTAPI
    Dim wzDelta
    Dim wKeys As Integer
    
    Select Case uMsg
    Case WM_MOUSEWHEEL                          '滚轮事件
        wzDelta = HIWORD(wParam)
        wKeys = LOWORD(wParam)
        pt.X = LOWORD(lParam)
        pt.Y = HIWORD(lParam)
    
        '将屏幕坐标转换为frmCaseTendBody窗口坐标
    
        ScreenToClient mfrmTendBody.hWnd, pt

        With mfrmTendBody.BodyEdit
        
            '判断坐标是否在frmCaseTendBody.BodyEdit窗口内
    
            If pt.X > .Left / Screen.TwipsPerPixelX And pt.X < (.Left + .Width) / Screen.TwipsPerPixelX And pt.Y > .Top / Screen.TwipsPerPixelY And pt.Y < (.Top + .Height) / Screen.TwipsPerPixelY Then
    
                If wKeys = 16 Then
                    '水平滚动
                    
                Else
                    '垂直滚动
                    If Sgn(wzDelta) = 1 Then
                        .ScrollBarY.Value = IIf(.ScrollBarY.Value - WHEEL_SCROLL_LINES < .ScrollBarY.Min, .ScrollBarY.Min, .ScrollBarY.Value = .ScrollBarY.Value - WHEEL_SCROLL_LINES)
                    Else
                        .ScrollBarY.Value = IIf(.ScrollBarY.Value + WHEEL_SCROLL_LINES > .ScrollBarY.Max, .ScrollBarY.Max, .ScrollBarY.Value + WHEEL_SCROLL_LINES)
                    End If
                End If
            End If
        End With
    Case Else                                   '其他事件仍由系统缺省处理
        WindowProc = CallWindowProc(glngPrevWndProc, hw, uMsg, wParam, lParam)
    End Select
End Function

Public Function PrintOrPreviewBodyState(objOut As Object, _
                                        ByVal lng病人ID As Long, _
                                        ByVal lng主页ID As Long, _
                                        ByVal lng文件ID As Long, _
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
    '功能:打印或预览某七天的温度表
    '参数:objOut=输出对象,可以为Printer或一个窗体(窗体中包含控件数组picPage)
    '      lngCaseRecordID=病历记录id
    '      lngBeginY=开始纵坐标
    '      blnKeepOn=是否保持连续
    '      objParent=主调用窗体
    '      intBeginPage=要开始页面序号,当为-1时表示输出所有.
    '      intEndPage=结束页面号如果intEndPage大于实际页数就只打印到实际页数
    '      intPageNO=开始的页码,如果为-1表示不显示页码
    '      sngScale=输出比例
    
    '返回:本次打印操作是否成功
    '******************************************************************************************************************
    Dim strSql As String, strNewSql As String
    '护理设置参数变量
    Dim intOpDays As Integer  '手术后标注天数
    Dim blnStopFlag As Boolean '再次手术停止前次标注
    Dim intOpFormat As Integer '手术当天缺省格式
    Dim byt未记显示位置 As Byte '未记说明显示位置
    Dim bln婴儿体温单显示出院 As Boolean '婴儿体温单显示出院信息
    Dim bln体温单显示诊断 As Boolean '体温单显示诊断
    Dim intRepairRows As Integer  '表格显示行数
    Dim bln显示皮试 As Boolean '体温单输出显示皮试结果
    Dim bln打印医院名称 As Boolean '体温单是否打印医院名称
    Dim bln入科显示入院 As Boolean
    Dim bln波动 As Boolean
    Dim bln汇总当天 As Boolean '体温单汇总数据时今天汇总今天还是今天汇总昨天的数据
    Dim bln录入小时 As Boolean '体温单全天汇总允许录入和显示汇总小时数
    Dim bln打印脉搏短绌 As Boolean  '体温单打印时是否打印脉搏短绌
    Dim bln不打印心率列 As Boolean '体温单打印时是否打印心率列(仅在心率单独应用是有效，只是不打印刻度列，点正常输出)
    Dim lngCurveRow As Long '体温曲线固定添加行数
    Dim bln出院 As Boolean
    
    '其他绘图变量
    Dim i As Integer, j As Integer
    Dim lngPicPageIndex As Integer '预览时PIC的索引
    Dim blnPrint As Boolean  '是否打印
    Dim strInfo As String '说明信息
    Dim intAllOpt As Single  '打印的总共步骤
    Dim intCurOpt As Single  '打印进行到第几步
    Dim objDraw As Object '绘图对象
    Dim lngHwnd As Long '句柄
    Dim lngDC As Long  '绘图对象的DC
    Dim lngFont As Long
    Dim lngOldFont As Long
    Dim stdset As StdFont
    Dim lngLableStep As Long '刻度区域列宽
    Dim lngColStep As Long ' 体温区域列宽
    Dim lngInitRowStep As Long '体温区域列高
    
    Dim lngCountPage As Long '所有页数
    Dim lngPage As Long
    Dim strBeginDate  As String, strBeginDate1 As String '开始时间
    Dim strEndDate As String '终止时间
    Dim strTmpDay As String, strEndDay As String
    Dim dtBegin As Date, dtEnd As Date
    Dim intDrawLineRows As Integer '体温单区域总列数
    Dim intDrawLineCOL As Integer '体温单刻度区域列数
    Dim strTmp As String, strTime As String, strTmp1 As String
    Dim lngValue As Long '住院天数
    Dim T_Rect As RECT
    Dim rsPart As New ADODB.Recordset  '所有体温部位信息
    Dim rsTemp As New ADODB.Recordset  '此记录集请不要顺便使用
    Dim rsTmp As New ADODB.Recordset
    Dim rsItems As New ADODB.Recordset '使用与此病人的所有护理项目信息
    Dim rsDrawItems As New ADODB.Recordset '体温单各个项目信息
    Dim rsPoints As New ADODB.Recordset '所有体温单的集合
    Dim rsNotes As New ADODB.Recordset   '所有说明信息
    Dim rsDownTab As New ADODB.Recordset '表下表格数据信息
    Dim H_16pt As Long, W_16pt As Long
    Dim int心率应用 As Integer
    Dim str心率符号  As String
    Dim arrTmpValue() As Variant, arrTmpNote() As Variant
    Dim arrValues() As String
    Dim strPart As String '部位
    Dim SinX As Single, sinY As Single
    Dim intCOl As Integer
    Dim blnAdd As Boolean, blnAllow As Boolean
    Dim dbl数值 As Double, dblMinValue As Double, dblMaxValue As Double
    Dim lng项目序号 As Long
    Dim str体温说明 As String
    Dim bln呼吸 As Boolean  '呼吸是否为表格
    Dim sngHTab As Single  '表下表格高度
    Dim sngHPrint As Single '可打印区域
    
    Dim strBegin As String, strEnd As String
    Dim str结果 As String
    Dim strItemName As String, strItems As String
    Dim int频次 As Integer
    Dim intCol1 As Integer
    Dim str项目名称 As String
    Dim int项目性质 As Integer, int项目类型 As Integer, int入院首测 As Integer
    Dim int舒张压 As Integer, int收缩压 As Integer, Int列号 As Integer
    Dim blnColor As Boolean

    '病人基本信息
    Dim strPatiInfo As String
    Dim VarPatiInfo As Variant
    Dim lng护理等级 As Long
    
    '--下面三个变量 在记录体温不升时做临时储存对象
    Dim strTmpString0 As String  '记录当前时间
    Dim strTmpString2 As String '记录住院天数
    Dim strTmpString1 As String '记录手术后天数
    Dim strNewTmpString As String
    Dim ArrNewTmpString() As String '记录表格项目的列数和每一列值的信息
    Dim ArrNewString() As String '记录所有表格项目信息
    Dim intDays As String '手术天数
    Dim strOpdays(1 To 7) As String
    Dim strOpValue(1 To 7) As String
    Dim arrOperDay
    Dim strEditors() As Variant    '记录曲线项目信息(项目序号||项目名称||项目单位||项目值域||记录符||记录色||最大值||最小值||临界值）
    Dim ArrComTable() As Variant '记录所有的表下表格项目 (项目序号||部位+项目名称|项目单位||项目值域||记录频次||项目性质||项目表示||入院首测)
    Dim lng次数 As Long  '记录手术次数
    
    '坐标信息
    Dim lngLeft As Long, lngTop As Long
    Dim lngRight As Long, lngButtom As Long
    Dim X As Long, Y As Long
    Dim lngCurX As Long, lngCurY As Long
    Dim dblSureW As Double, dblSureH As Double
    
    Dim M_DrawClient As DrawClient
    
    On Error GoTo ErrPrint
    
    msngTwips = 1
    
    mintBaby = intBaby
    '保存原始值:
    
    M_DrawClient.偏移量X = T_DrawClient.偏移量X
    M_DrawClient.偏移量Y = T_DrawClient.偏移量Y
    M_DrawClient.刻度区域 = T_DrawClient.刻度区域
    M_DrawClient.刻度单位 = T_DrawClient.刻度单位
    M_DrawClient.体温区域 = T_DrawClient.体温区域
    M_DrawClient.行单位 = T_DrawClient.行单位
    M_DrawClient.时间行单位 = T_DrawClient.时间行单位
    M_DrawClient.时间列单位 = T_DrawClient.时间列单位
    M_DrawClient.列单位 = T_DrawClient.列单位
    M_DrawClient.双倍 = T_DrawClient.双倍
    M_DrawClient.总列数 = T_DrawClient.总列数
    
    mintBmpW = gintBmpW
    mintBmpH = gintBmpH
    '读取体温参数信息
    '------------------------------------------------------------------------------------------------------------------
    intOpDays = Val(zldatabase.GetPara("手术后标注天数", glngSys, 1255, "10"))
    blnStopFlag = (Val(zldatabase.GetPara("再次手术停止前次标注", glngSys, 1255, "0")) = 1)
    byt未记显示位置 = Abs(Val(zldatabase.GetPara("未记说明显示位置", glngSys, 1255, "0")))
    bln婴儿体温单显示出院 = (zldatabase.GetPara("婴儿体温单显示出院信息", glngSys, 1255, 1) = 1)
    bln体温单显示诊断 = (zldatabase.GetPara("体温单显示诊断", glngSys, 1255, 1) = 1)
    intRepairRows = zldatabase.GetPara("体温表格行数", glngSys, 1255, 8)
    bln显示皮试 = (Val(zldatabase.GetPara("体温单显示皮试结果", glngSys, 1255, "0")) = 1)
    bln打印医院名称 = (Val(zldatabase.GetPara("打印医院名称", glngSys, 1255, "1")) = 1)
    bln汇总当天 = (Val(zldatabase.GetPara("汇总波动显示当天数据", glngSys, 1255, 0)) = 1)
    bln打印脉搏短绌 = (Val(zldatabase.GetPara("不打印脉搏短绌图形", glngSys, 1255, "0")) = 0)
    bln不打印心率列 = (Val(zldatabase.GetPara("体温单不打印心率列", glngSys, 1255, "0")) = 1)
    lngCurveRow = Val(zldatabase.GetPara("体温曲线固定添加行数", glngSys, 1255, "0"))
    
    '--51282,刘鹏飞,2012-08-03,全天汇总显示录入时间(DYEY要求手工录入汇总时间H)
    bln录入小时 = (Val(zldatabase.GetPara("全天汇总显示录入时间", glngSys, 1255, 0)) = 1)
    
    '51338,刘鹏飞,2012-07-06
    strTmp = zldatabase.GetPara("手术当天缺省格式", glngSys, 1255, "2")
    If Val(strTmp) >= 0 And Val(strTmp) <= 2 Then
        intOpFormat = Val(strTmp)
    Else
        intOpFormat = 0
    End If
    '病人变动标记显示方法
    '------------------------------------------------------------------------------------------------------------------
    Call InitPara
    
    blnPrint = TypeName(objOut) = "Printer"
    
    '由于打印机和屏幕的像素不同，此处需要取各自的像素
    If blnPrint = True Then
        T_TwipsPerPixel.X = Printer.TwipsPerPixelX
        T_TwipsPerPixel.Y = Printer.TwipsPerPixelY
        msngTwips = Screen.TwipsPerPixelX / Printer.TwipsPerPixelX
        Printer.Font.Size = 9
        Printer.FontName = "宋体"
    Else
        T_TwipsPerPixel.X = Screen.TwipsPerPixelX
        T_TwipsPerPixel.Y = Screen.TwipsPerPixelY
        msngTwips = 1
    End If
    
    Screen.MousePointer = 11
    intAllOpt = 5
    
    '计算进度处理
    '------------------------------------------------------------------------------------------------------------------
    strInfo = "正在" & IIf(blnPrint, "准备打印体温表", "处理预览") & ",请稍候..."
    Call ShowFlash(strInfo, , objParent)
    
    '打印前的清除
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
    
    bln出院 = False
    '提取婴儿医嘱信息(转科，出院)存在医嘱以医嘱信息为准，否则以母亲出院日期为准
    strNewSql = "   (SELECT /*+ RULE */  病人ID,主页ID,婴儿时间,DECODE(nvl(婴儿,0),0, DECODE(NVL(出院日期,''),'',0,1), DECODE(NVL(婴儿时间,''),'',0,1))记录" & vbNewLine & _
                "       FROM (SELECT A.病人ID,A.主页ID,B.开始执行时间 婴儿时间, A.出院日期,B.婴儿" & vbNewLine & _
                "           FROM 病案主页 A," & vbNewLine & _
                "               (SELECT B.病人ID, B.主页ID, B.婴儿, 开始执行时间" & vbNewLine & _
                "                FROM 病人医嘱记录 B, 诊疗项目目录 C" & vbNewLine & _
                "                WHERE B.诊疗项目ID + 0 = C.ID AND B.医嘱状态 = 8 AND nvl(B.婴儿,0)<>0  AND C.类别 = 'Z'" & vbNewLine & _
                "                AND EXISTS (SELECT 1 FROM TABLE(CAST(F_STR2LIST('3,5,11') AS ZLTOOLS.T_STRLIST))" & vbNewLine & _
                "                               WHERE C.操作类型 = COLUMN_VALUE) And  B.病人ID = [2] AND B.主页ID = [3] AND B.婴儿(+) = [4]) B" & vbNewLine & _
                "           WHERE A.病人ID = [2] AND A.主页ID = [3] AND A.病人ID = B.病人ID(+) AND A.主页ID = B.主页ID(+)" & vbNewLine & _
                "           ORDER BY B.开始执行时间 DESC)" & vbNewLine & _
                "       WHERE ROWNUM < 2)  E"
    '提取病人出院前的时间信息
    '------------------------------------------------------------------------------------------------------------------
    strSql = _
       "Select Decode(b.出生时间,Null,a.开始,b.出生时间) As 开始,decode(E.记录,0,Decode(Sign(NVL(E.婴儿时间,a.终止) - d.发生时间), 1,NVL(E.婴儿时间,a.终止) ,d.发生时间),NVL(E.婴儿时间,a.终止)) 终止,E.记录" & vbNewLine & _
        "       From" & vbNewLine & _
        "       (Select 病人ID,主页id,Min(开始时间) as 开始,Max(Nvl(终止时间,sysdate)) as 终止" & vbNewLine & _
        "       From 病人变动记录" & vbNewLine & _
        "       Where 开始时间 is Not Null And 病人ID=[2] And 主页ID=[3] Group By 病人ID,主页id) a," & vbNewLine & _
        "       (Select 病人ID,主页id,出生时间 From 病人新生儿记录 Where 病人ID =[2] And 主页ID =[3] And 序号=[4]) b," & vbNewLine & _
        "       (SELECT NVL(发生时间,SYSDATE) 发生时间 FROM (select max(发生时间) 发生时间 from 病人护理文件 A,病人护理数据 B" & vbNewLine & _
        "       where A.ID=B.文件ID and A.ID=[1] and A.病人ID=[2] and A.主页ID=[3] and A.婴儿=[4])) d," & vbNewLine & _
        strNewSql & vbNewLine & _
        "       Where A.病人ID=E.病人ID And A.主页ID=E.主页ID And a.病人id=b.病人id(+) And a.主页id=b.主页id(+)"
        
    Set rsTemp = zldatabase.OpenSQLRecord(strSql, "mdlPrint", lng文件ID, lng病人ID, lng主页ID, intBaby)
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
        lngCountPage = DateDiff("d", rsTemp!开始, rsTemp!终止) + 1
        lngCountPage = IIf(lngCountPage / 7 = Fix(lngCountPage / 7), lngCountPage / 7, Fix(lngCountPage / 7) + 1)
        strBeginDate = Format(rsTemp!开始, "YYYY-MM-DD HH:MM:SS")
        strBeginDate1 = strBeginDate
        strEndDate = Format(rsTemp!终止, "YYYY-MM-DD HH:MM:SS")
        bln出院 = Not (Val(rsTemp!记录) = 0)
    Else
        CloseRs rsTemp
        GoTo ErrPrint '无数病人变动信息退出
    End If
    
    gbln出院 = bln出院
    '提取用户设置的体温单开始时间(婴儿以出生时间为准)
    If intBaby = 0 Then
        strSql = "select 开始时间 from 病人护理文件 where ID=[1] and 病人ID=[2] and 主页id=[3] and nvl(婴儿,0)=[4]"
        Set rsTmp = zldatabase.OpenSQLRecord(strSql, "提取体温单开始时间", lng文件ID, lng病人ID, lng主页ID, intBaby)
        If rsTmp.RecordCount <> 0 Then
            strBeginDate = Format(rsTmp!开始时间, "YYYY-MM-DD HH:mm:ss")
        End If
    End If
    
    If bln出院 = True Then
        '出院时间和入院时间如果在同一列，则将出院时间后移一列（内蒙需求:出院也要录入体温）
        strEndDate = Format(RetrunEndTime(CDate(strBeginDate), CDate(strEndDate), gintHourBegin), "YYYY-MM-DD HH:mm:ss")
    End If
    
    bln入科显示入院 = False
    
    If CDate(Format(strBeginDate, "YYYY-MM-DD HH:MM:SS")) > CDate(Format(strBeginDate1, "YYYY-MM-DD HH:MM:SS")) Then
        bln入科显示入院 = True
    ElseIf T_BodyFlag.入院 = 0 And CDate(Format(strBeginDate, "YYYY-MM-DD HH:MM:SS")) = CDate(Format(strBeginDate1, "YYYY-MM-DD HH:MM:SS")) Then
        bln入科显示入院 = True
    End If
            
    intCurOpt = intCurOpt + 1
    
    Call ShowFlash(strInfo, intCurOpt / intAllOpt, objParent)
    
    '------------------------------------------------------------------------------------------------------------------
    '第1部份：病人的基本信息
    '读取病人基本信息
    
    '"姓名'年龄'性别'科别'床号'入院日期'住院号:
    strPatiInfo = "''''''"
    VarPatiInfo = Split(strPatiInfo, "'")
    
    strSql = " Select  b.姓名,A.住院号,A.入院日期 入院时间,b.性别,A.年龄 From 病人信息 B,病案主页 A Where A.病人ID=B.病人ID And A.病人id=[1] And A.主页ID=[2]"
    Set rsTemp = zldatabase.OpenSQLRecord(strSql, "mdlPrint", lng病人ID, lng主页ID)
    If rsTemp.BOF = False Then
        VarPatiInfo(0) = zlCommFun.Nvl(rsTemp("姓名").Value)
        VarPatiInfo(6) = zlCommFun.Nvl(rsTemp("住院号").Value)
        VarPatiInfo(5) = Format(zlCommFun.Nvl(rsTemp("入院时间").Value), "yyyy-MM-dd")
        VarPatiInfo(2) = zlCommFun.Nvl(rsTemp("性别").Value)
        VarPatiInfo(1) = zlCommFun.Nvl(rsTemp("年龄").Value)
    End If
    
    '入院时间(如果体温单开始时间大于入院时间就以入科时间为准)
    strSql = "select 开始时间 from 病人变动记录 where 病人id=[1] And 主页id=[2] and 开始原因=2 order by 开始时间"
    Set rsTemp = zldatabase.OpenSQLRecord(strSql, "mdlPrint", lng病人ID, lng主页ID)
    If rsTemp.BOF = False Then
        If bln入科显示入院 = True Then
            VarPatiInfo(5) = Format(zlCommFun.Nvl(rsTemp("开始时间").Value), "yyyy-MM-dd")
        End If
    End If
    
    If intBaby <> 0 Then
        
        VarPatiInfo(1) = ""
        VarPatiInfo(2) = ""
        
        strSql = "Select Decode(a.婴儿姓名,Null,b.姓名||'之子'||Trim(To_Char(a.序号,'9')),a.婴儿姓名) As 婴儿姓名,婴儿性别,出生时间 " & _
            " From 病人新生儿记录 a,病人信息 b " & _
            " Where a.病人id=[1] And a.主页id=[2] And a.病人id=b.病人id And a.序号=[3]"
        Set rsTemp = zldatabase.OpenSQLRecord(strSql, "mdlPrint", lng病人ID, lng主页ID, intBaby)
        If rsTemp.BOF = False Then
            VarPatiInfo(0) = rsTemp("婴儿姓名").Value
            VarPatiInfo(2) = zlCommFun.Nvl(rsTemp("婴儿性别").Value)
            VarPatiInfo(1) = "新生儿"
            If IsNull(rsTemp("出生时间").Value) = False Then VarPatiInfo(5) = Format(zlCommFun.Nvl(rsTemp("出生时间").Value), "yyyy-MM-dd")
        End If
        
    End If
    
    If bln体温单显示诊断 Then ReDim Preserve VarPatiInfo(UBound(VarPatiInfo) + 1)
    
    intCurOpt = intCurOpt + 1
    Call ShowFlash(strInfo, intCurOpt / intAllOpt, objParent)
    
    '获取病人护理等级
    strSql = "Select zl_PatitTendGrade([1],[2]) As 护理等级 From dual"
    Set rsTemp = zldatabase.OpenSQLRecord(strSql, "护理等级", lng病人ID, lng主页ID)
    If rsTemp.BOF = False Then lng护理等级 = zlCommFun.Nvl(rsTemp("护理等级"), 0)
    
    '提取共用记录集
    Call InitPublicData
    
    '求出心率应用方式
    int心率应用 = 2
    str心率符号 = ""
    strSql = "Select a.应用方式,b.记录符 From 护理记录项目 a,体温记录项目 b Where a.项目序号=-1 And a.项目序号=b.项目序号"
    Set rsTemp = zldatabase.OpenSQLRecord(strSql, "mdlPrint")
    If rsTemp.BOF = False Then
        int心率应用 = zlCommFun.Nvl(rsTemp("应用方式").Value, 2)
        str心率符号 = zlCommFun.Nvl(rsTemp("记录符").Value, "○")
    Else
        int心率应用 = 0
    End If
    
    Dim int脉搏 As Integer, int心率 As Integer
    
    '-------------------------------------------------------------------------------------------------------------------
    '2提取所有曲线项目(此体温单固定有两行输出所以最高行-2)
    strSql = " Select A.项目序号,A.排列序号,A.记录名,C.项目值域,A.记录符,A.记录色,nvl(A.最大值,0) 最大值 ,nvl(A.最小值,0) 最小值,A.临界值," & _
        "nvl(A.单位值,0) 单位值,A.刻度间隔,A.警示线,C.项目单位 单位,nvl(A.最高行,2)-2 AS 最高行,B.部位 " & _
        " From 体温记录项目 A,体温部位 B,护理记录项目 C" & _
        " Where A.项目序号=B.项目序号(+) And B.缺省项(+)=1" & _
        " And A.记录法=1 And A.项目序号=C.项目序号 and nvl(C.应用方式,0)=1 and C.护理等级>=[1]" & _
        " and nvl(C.适用病人,0) in (0,[2]) and (C.适用科室=1 or (C.适用科室=2 and Exists (select 1 from 护理适用科室 D where C.项目序号=D.项目序号 and D.科室ID=[3])))" & _
        " Order by 排列序号"
    Set rsTemp = zldatabase.OpenSQLRecord(strSql, "获取所有曲线项目", lng护理等级, IIf(intBaby = 0, 1, 2), lngSectID)
    
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
        rsTemp.Filter = "项目序号=" & gint心率
        If rsTemp.RecordCount > 0 And bln不打印心率列 Then
            rsTemp.Filter = 0
            intDrawLineCOL = rsTemp.RecordCount - 1
        Else
            rsTemp.Filter = 0
            intDrawLineCOL = rsTemp.RecordCount
        End If
        If intDrawLineCOL <= 0 Then intDrawLineCOL = 1
    Else
        CloseRs rsTemp
        MsgBox "无任何体温曲线项目！", vbExclamation, gstrSysName
        GoTo ErrExit
    End If
    strEditors = Array()
    int脉搏 = -1: int心率 = -1
    rsTemp.Filter = 0
    rsTemp.Sort = "排列序号"
    With rsTemp
        Do While Not .EOF
            strTmp = zlCommFun.Nvl(!项目序号, 0) & "|| " & zlCommFun.Nvl(!记录名) & "|| " & zlCommFun.Nvl(!单位) & "|| " & zlCommFun.Nvl(!项目值域) & "|| " & _
                 zlCommFun.Nvl(!记录符) & "|| " & zlCommFun.Nvl(!记录色) & "||" & zlCommFun.Nvl(!最大值) & "||" & zlCommFun.Nvl(!最小值) & "||" & zlCommFun.Nvl(!临界值)
                
            ReDim Preserve strEditors(UBound(strEditors) + 1)
            strEditors(UBound(strEditors)) = strTmp
            If zlCommFun.Nvl(!项目序号, 0) = gint脉搏 Then
                int脉搏 = UBound(strEditors)
            End If
        .MoveNext
        Loop
        .MoveFirst
    End With
    If int心率应用 = 2 And int脉搏 <> -1 Then
        ReDim Preserve strEditors(UBound(strEditors) + 1)
        strTmp = "-1||心率||" & Split(strEditors(int脉搏), "||")(2) & "||" & Split(strEditors(int脉搏), "||")(3) & "||○||" & RGB_RED & "||" & _
            Split(strEditors(int脉搏), "||")(6) & "||" & Split(strEditors(int脉搏), "||")(7) & "||" & Split(strEditors(int脉搏), "||")(8)
        strEditors(UBound(strEditors)) = strTmp
    End If
    
    intCurOpt = intCurOpt + 1
    Call ShowFlash(strInfo, intCurOpt / intAllOpt, objParent)
    
    '3—提取所有特殊项目信息包括活动项目（活动项目可能存在一个项目多个部位也要提取）
    ArrComTable = Array()
    strTmp = ""
    strTime = ""
    
    '提取表格非汇总项目
    gstrSQL = "Select A.排列序号,A.项目序号,A.记录名,A.记录法,A.记录符,A.记录色,B.项目值域,nvl(A.记录频次,2) 记录频次,A.入院首测,B.项目性质," & _
        "   B.项目类型,B.项目长度,B.项目表示,B.项目小数,B.项目单位 单位" & _
        "   From 体温记录项目 A,护理记录项目 B,诊治所见项目 C" & _
        "   Where A.项目序号=B.项目序号 And B.项目ID=C.Id(+)  And A.记录法=2" & _
        "   And nvl(B.应用方式,0)=1 And nvl(B.护理等级,0)>=[7] And nvl(B.适用病人,0) In (0,[8])" & _
        "   And (B.适用科室=1 Or (B.适用科室=2 And Exists (Select 1 From 护理适用科室 D Where D.项目序号=B.项目序号 And D.科室id=[9])))"
        
    
    strSql = "Select Rownum-1 序号 ,项目序号,项目名称,记录色,项目单位,项目值域, 部位,记录频次,入院首测,项目性质,项目表示,项目类型 From (" & _
            " Select A.项目序号, Decode(A.项目序号, 4, '血压', A.记录名) 项目名称,A.记录色,A.单位 项目单位, A.项目值域, B.部位," & vbNewLine & _
            "           nvl(A.记录频次,2) 记录频次,A.入院首测, nvl(A.项目性质,1) 项目性质, A.项目表示,A.项目类型" & vbNewLine & _
            " From (" & gstrSQL & " ) A," & vbNewLine & _
             "        (Select Distinct b.项目序号, a.部位" & vbNewLine & _
            "            From (Select 项目序号, DECODE(项目序号,3,'',体温部位) 部位" & vbNewLine & _
            "                           From 病人护理文件 a, 病人护理数据 b, 病人护理明细 c" & vbNewLine & _
            "                           Where a.Id = b.文件id And b.Id = c.记录id And a.Id = [1] And Nvl(a.婴儿, 0) = [4] And a.病人id = [2] And" & vbNewLine & _
            "                                       a.主页id = [3] And c.记录类型 = 1 And b.发生时间 Between [5] And [6] And 终止版本 Is Null) a, 体温记录项目 b," & vbNewLine & _
            "                       护理记录项目 c" & vbNewLine & _
            "            Where b.项目序号 = a.项目序号(+) And b.项目序号 = c.项目序号 And b.记录法 = 2 And Nvl(护理等级, 0) >=[7]) B" & vbNewLine & _
            "   where A.项目序号=B.项目序号 and A.项目序号<>5  order by Decode(A.项目序号,3 ,0,1 ),A.排列序号,项目名称,B.部位)"

    If blnMoved Then
        strSql = Replace(strSql, "病人护理文件", "H病人护理文件")
        strSql = Replace(strSql, "病人护理数据", "H病人护理数据")
        strSql = Replace(strSql, "病人护理明细", "H病人护理明细")
    End If
    
    Set rsItems = zldatabase.OpenSQLRecord(strSql, "取开始行", lng文件ID, lng病人ID, lng主页ID, intBaby, Int(CDate(strBeginDate)), CDate(strEndDate), lng护理等级, IIf(intBaby = 0, 1, 2), lngSectID)
    
    bln呼吸 = False
    With rsItems
        Do While Not .EOF
            str项目名称 = ""
            If Val(Nvl(!项目性质, 1)) = 2 Then
                str项目名称 = Trim(Nvl(!部位)) & Nvl(!项目名称)
            Else
                str项目名称 = Nvl(!项目名称)
            End If
            
            int频次 = Val(zlCommFun.Nvl(!记录频次))
            
            If zlCommFun.Nvl(!项目表示) = 4 Or IsWaveItem(Val(zlCommFun.Nvl(!项目序号))) Then
                If int频次 > 2 Then int频次 = 2
            End If
            
            strTmp = zlCommFun.Nvl(!项目序号) & "||" & Replace(str项目名称, ";", ":") & "||" & zlCommFun.Nvl(!项目单位) & "||" & _
                zlCommFun.Nvl(!项目值域) & "||" & int频次 & "||" & zlCommFun.Nvl(!项目性质, 1) & "||" & _
                zlCommFun.Nvl(!项目表示) & "||" & zlCommFun.Nvl(!项目类型) & "||" & zlCommFun.Nvl(!入院首测, 0)
            If Val(zlCommFun.Nvl(!项目序号)) = gint呼吸 Then
                bln呼吸 = True
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
    '4、确定X和Y的坐标位置
    '边界信息(Twip)
    
    Dim lngOffsetLeft As Long
    Dim lngOffsetTop As Long
    
    dblSureH = 0
    dblSureW = 0
    If blnPrint = True Then
        '如果是打印预览,应按打印机的可打印的开始处开始预览
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
    
    H_16pt = objDraw.TextHeight("字") / T_TwipsPerPixel.Y
    W_16pt = objDraw.TextWidth("字") / T_TwipsPerPixel.X
    
    X = lngLeft: Y = lngTop
    lngCurX = X: lngCurY = Y
    
    If intDrawLineCOL <= 3 Then
        lngLableStep = (glngLableWith / intDrawLineCOL) * sngScale * msngTwips
    Else
        lngLableStep = glngLableStep * sngScale * msngTwips
    End If
    
    T_DrawClient.刻度区域.Left = lngCurX
    T_DrawClient.刻度区域.Right = lngCurX + intDrawLineCOL * lngLableStep
    
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
    
    T_DrawClient.体温区域.Left = T_DrawClient.刻度区域.Right
    T_DrawClient.体温区域.Right = T_DrawClient.刻度区域.Right + (6 * 7 * lngColStep)
    
    Dim sigSign As Single
    sigSign = 1
    If T_DrawClient.体温区域.Right > objDraw.Width / T_TwipsPerPixel.X - lngRight Then
        sigSign = Round((T_DrawClient.体温区域.Right - (objDraw.Width / T_TwipsPerPixel.X - lngRight)) / (T_DrawClient.体温区域.Right - T_DrawClient.刻度区域.Right), 2)
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
    
    T_DrawClient.刻度单位 = lngLableStep
    T_DrawClient.刻度区域.Right = lngCurX + intDrawLineCOL * lngLableStep
    T_DrawClient.体温区域.Left = T_DrawClient.刻度区域.Right
    T_DrawClient.体温区域.Right = T_DrawClient.刻度区域.Right + (6 * 7 * lngColStep)
    T_DrawClient.列单位 = lngColStep
    T_DrawClient.行单位 = lngInitRowStep
    T_DrawClient.时间列单位 = 16 * msngTwips
    T_DrawClient.偏移量X = lngLeft
    '------------------------------------------------------------------------------------------------------------------
    '求得首列宽，求左边标尺总共有多少行
    '求得体温表项目的总行数
    strSql = "Select Count(A.项目序号) 记录数 " & _
        "   From 体温记录项目 A,护理记录项目 B " & _
        "   Where A.项目序号=B.项目序号 And A.记录法=[1]" & _
        "   And nvl(B.应用方式,0)=1 And nvl(B.护理等级,0)>=[2] And nvl(B.适用病人,0) In (0,[3])" & _
        "   And (B.适用科室=1 Or (B.适用科室=2 And Exists (Select 1 From 护理适用科室 D Where D.项目序号=B.项目序号 And D.科室id=[4])))"
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, "mdlPrint", 1, lng护理等级, IIf(intBaby = 0, 1, 2), lngSectID)
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        intDrawLineRows = zlCommFun.Nvl(rsTmp!记录数, 0)
    Else
        CloseRs rsTmp
        GoTo ErrPrint
    End If
    
    If intDrawLineRows < 1 Then
        CloseRs rsTmp
        GoTo ErrPrint
    End If
    
    strSql = "Select nvl(A.最大值,0) 最大值,nvl(A.最小值,0) 最小值 ,nvl(A.单位值,0.1) ,nvl(A.最高行,0)-2  最高行" & _
        "   From 体温记录项目 A,护理记录项目 B" & _
        "   Where A.项目序号=B.项目序号 And A.项目序号=[1]" & _
        "   And nvl(B.应用方式,0)=1 And nvl(B.护理等级,0)>=[2] And nvl(B.适用病人,0) In (0,[3])" & _
        "   And (B.适用科室=1 Or (B.适用科室=2 And Exists (Select 1 From 护理适用科室 D Where D.项目序号=B.项目序号 And D.科室id=[4])))"
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, "mdlPrint", gint体温, lng护理等级, IIf(intBaby = 0, 1, 2), lngSectID)
    If rsTmp.RecordCount > 0 Then
        '修改问题：51442
        dbl数值 = Val(zlCommFun.Nvl(rsTmp!最小值, 0))
        intDrawLineRows = (Val(rsTmp!最大值) - IIf(dbl数值 > 34, 35, dbl数值)) / 0.1 + IIf(Val(rsTmp!最高行) < 0, 0, Val(rsTmp!最高行)) + IIf(dbl数值 > 34, 10, 0)
        intDrawLineRows = intDrawLineRows + lngCurveRow
    End If
    
    If intDrawLineRows > glngMaxRows Then
        T_DrawClient.总列数 = intDrawLineRows
    Else
        T_DrawClient.总列数 = glngMaxRows
    End If
    
    intCurOpt = intCurOpt + 1
    Call ShowFlash(strInfo, intCurOpt / intAllOpt, objParent)
    '5、循环：按总页数循环
    intCurOpt = 0
    intAllOpt = 100
    intCurOpt = intCurOpt + 1
    Call ShowFlash(strInfo, intCurOpt / intAllOpt, objParent)
    If blnPrint = False Then
        lngPicPageIndex = objOut.picPage.UBound + 1
    End If
    
    '正式开始第四步，循环每一页
    '------------------------------------------------------------------------------------------------------------------
    For lngPage = 1 To lngCountPage
        strTmpDay = Format(CDate(strBeginDate) + 7 * (lngPage - 1), "YYYY-MM-DD") '求得当前页面的第一天日期与时间
        If CDate(strTmpDay) < CDate(strBeginDate) Then strTmpDay = strBeginDate
        If CDate(strEndDate) < CDate(Format(zldatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")) And Not bln出院 Then strEndDate = Format(zldatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")
        strEndDay = Format(CDate(strTmpDay) + 6, "YYYY-MM-DD") & " 23:59:59"
        If CDate(strEndDay) > CDate(strEndDate) Then strEndDay = Format(strEndDate, "YYYY-MM-DD HH:mm:ss")
        intCurOpt = lngPage / lngCountPage
        strInfo = "正在" & IIf(blnPrint, "打印体温表", "预览") & ",请稍候..."
        Call ShowFlash(strInfo, intCurOpt, objParent)
        
        '按页号打印
        If intBeginPage > 0 Then  '只打印指定页码的
            If lngPage >= intBeginPage And lngPage <= intEndPage Then
                If lngPage > intBeginPage Then  '到第二页时开始初始化纸张或页面
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
        Else  '打印所有时
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
        
         '页眉图形输出
        Call frmTendFileRead.PrintRTBData(objDraw, True, lngTop)
        
        '获取对象的Dc
        lngDC = objDraw.hDC
        '创建字体
        Set stdset = New StdFont
        stdset.Name = "宋体"
        stdset.Size = 9 * sngScale
        Call SetFontIndirect(stdset, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        '打印质控号
        strTmp = zldatabase.GetPara("质控号", glngSys, 1255, "")
        Call GetTextExtentPoint32(lngDC, strTmp, Len(strTmp), T_Size)
        T_Size.W = objDraw.TextWidth(strTmp) / T_TwipsPerPixel.X
        lngCurX = T_DrawClient.体温区域.Right - T_Size.W
        Call GetTextRect(objDraw, lngCurX, lngCurY, strTmp, , , , sngScale)
        Call DrawText(lngDC, strTmp, -1, T_LableRect, DT_CENTER)
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
        
        '是否打印医院名称，有的医院体温单医院名可能存在两个，需要用页眉来实现。此时就不在打印注册文件中的医院信息。
        If bln打印医院名称 = True Then
            '获取医院名称
            stdset.Name = "宋体"
            stdset.Size = 18 * sngScale
            stdset.Bold = True
            Call SetFontIndirect(stdset, lngDC, objDraw)
            lngFont = CreateFontIndirect(T_Font)
            lngOldFont = SelectObject(lngDC, lngFont)
            strTmp = IIf(GetUnitName = "-", "", GetUnitName) & IIf(intBaby <> 0, "婴儿", "") & "体温单"
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
        '读取病人科室、床号等信息
    
        VarPatiInfo(3) = ""
        VarPatiInfo(4) = ""
        strTmp = "": strTime = ""
        strSql = " Select  c.名称 As 科室,b.名称 As 病区,a.床号,a.开始原因 " & _
                    " From 病人变动记录 a,部门表 b,部门表 c " & _
                    " Where a.病人id=[1] And a.主页id=[2] And a.科室id Is Not Null And a.病区id=b.id and a.科室id=c.id " & _
                    " And a.开始时间-4/24<=[3] And Nvl(a.终止时间,Sysdate)>=[4] Order By a.开始时间"
        
        Set rsTmp = zldatabase.OpenSQLRecord(strSql, "读取病人科室、床号等信息", lng病人ID, lng主页ID, CDate(strEndDay), CDate(strTmpDay))
        If rsTmp.BOF = False Then
            Do While Not rsTmp.EOF
                
                If zlCommFun.Nvl(rsTmp("科室").Value) <> strTmp And zlCommFun.Nvl(rsTmp("科室").Value) <> "" Then
                
                    strTmp = zlCommFun.Nvl(rsTmp("科室").Value)
                    
                    If VarPatiInfo(3) = "" Then
                        VarPatiInfo(3) = strTmp
                    Else
                        VarPatiInfo(3) = VarPatiInfo(3) & "->" & strTmp
                    End If
                    
                End If
    
                If zlCommFun.Nvl(rsTmp("床号").Value) <> strTime And zlCommFun.Nvl(rsTmp("床号").Value) <> "" Then
                
                    strTime = zlCommFun.Nvl(rsTmp("床号").Value)
                    
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
        
        If bln体温单显示诊断 Then
            '提取病人诊断信息
            strSql = "Select Zl_Replace_Element_Value([1],[2],[3],2,NULL,0,[4]) As 最后诊断 From Dual"
            Set rsTmp = zldatabase.OpenSQLRecord(strSql, "最后诊断", "最后诊断", lng病人ID, lng主页ID, CDate(strTmpDay))
            If rsTmp.BOF = False Then
                If intBaby = 0 Then
                    VarPatiInfo(UBound(VarPatiInfo)) = zlCommFun.Nvl(rsTmp("最后诊断").Value)
                Else
                    VarPatiInfo(UBound(VarPatiInfo)) = ""
                End If
            Else
                VarPatiInfo(UBound(VarPatiInfo)) = ""
            End If
        End If
        strPatiInfo = Join(VarPatiInfo, "'")
        
        stdset.Name = "宋体"
        stdset.Size = 9 * sngScale
        stdset.Bold = True
        Call SetFontIndirect(stdset, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        '输出病人信息
        Call DrawPatiInfo(lngDC, objDraw, strPatiInfo, lngCurX, lngCurY, T_DrawClient.体温区域.Right, lngCurY, sngScale)
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
        '---开始画体温单上表格(住院日期,住院天数,手术,时间)
        Y = lngCurY: lngCurX = X: lngCurY = Y
        '1.提取住院开始天数
        lngValue = 0: strTmp = "": strTime = ""
        strSql = "Select zl_CalcInDaysNew([1],[2],[3],[4]) As 开始天数 From Dual"
        Set rsTmp = zldatabase.OpenSQLRecord(strSql, "提取住院天数", lng文件ID, lng病人ID, lng主页ID, Int(CDate(strTmpDay)))

        If rsTmp.BOF = False Then
            lngValue = rsTmp("开始天数").Value
        End If
        For i = 0 To 6
            strTmp = Format(CDate(strTmpDay) + i, "YYYY-MM-DD")
            If Right(strTmp, 5) = "01-01" Then
                '一年的第一天
                strTime = strTmp
            ElseIf strTmp = Format(strBeginDate, "yyyy-MM-dd") Then
                '入院第一天，写上年份
                strTime = strTmp
            ElseIf i = 0 Then '每页的第一列
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
        '2.提取手术时间和次数
        strTime = ""
        '显示但前段的手术标记
        strSql = "Select B.发生时间 时间" & vbNewLine & _
            " From 病人护理文件 A,病人护理数据 B,病人护理明细 C" & vbNewLine & _
            " Where A.Id=B.文件ID And B.Id=C.记录ID And A.Id=[1] And  nvl(A.婴儿,0)=[2]" & vbNewLine & _
            " And A.病人ID=[3] and A.主页ID=[4] and C.记录类型=4 and C.终止版本 is null" & vbNewLine & _
            " And B.发生时间 between [5] and [6] order by B.发生时间"
        If blnMoved Then
            strSql = Replace(strSql, "病人护理文件", "H病人护理文件")
            strSql = Replace(strSql, "病人护理数据", "H病人护理数据")
            strSql = Replace(strSql, "病人护理明细", "H病人护理明细")
        End If

        Set rsTmp = zldatabase.OpenSQLRecord(strSql, "提取手术标记", lng文件ID, intBaby, lng病人ID, lng主页ID, Int(CDate(strTmpDay) - 14), CDate(strEndDay))

        Do While Not rsTmp.EOF
            strTime = Format(rsTmp("时间"), "YYYY-MM-DD")
            For i = 1 To 7
                If DateDiff("d", strTmpDay, strEndDay) + 1 >= i Then
                    intDays = DateDiff("d", strTime, strTmpDay) + (i - 1)

                    Select Case intDays
                        Case 0 '当前区域内的手术开始时间
                             'Modify 2012-03-05 修改一天可以有多次手术
                            If Trim(strOpdays(i)) <> "" Then
                                strOpdays(i) = strTime & "/" & strOpdays(i)
                            Else
                                strOpdays(i) = strTime
                            End If
                        Case Else
                            If intDays >= 1 And intDays <= intOpDays Then '手术开始天数
                                If blnStopFlag Then '手术标注后天数在次手术时停止前一次标注
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
        
        '提取当前开始日期-14天前的手术记录信息
        strSql = "select Nvl(Count(B.发生时间),0) 次数" & _
            "   from 病人护理文件 A, 病人护理数据 B,病人护理明细 C" & _
            "   where A.ID=B.文件ID and B.ID=C.记录ID and A.ID=[1] and nvl(A.婴儿,0)=[2]" & _
            "   and A.病人ID=[3] and A.主页ID=[4] and C.记录类型=4 and C.终止版本 is null" & _
            "   and B.发生时间 <[5] "
        Set rsTmp = zldatabase.OpenSQLRecord(strSql, "提取手术标记", lng文件ID, intBaby, lng病人ID, lng主页ID, Int(CDate(strTmpDay)))
        If blnMoved Then
            strSql = Replace(strSql, "病人护理文件", "H病人护理文件")
            strSql = Replace(strSql, "病人护理数据", "H病人护理数据")
            strSql = Replace(strSql, "病人护理明细", "H病人护理明细")
        End If
        
        lng次数 = 0
        If rsTmp.BOF = False Then lng次数 = Val(rsTmp("次数"))
        
        For i = 1 To 7
            If DateDiff("d", Int(CDate(strTmpDay)), Int(CDate(strEndDay))) + 1 >= i Then
                If Trim(strOpdays(i)) <> "" Then
                    arrOperDay = Split(strOpdays(i), "/")
                Else
                    arrOperDay = Split("1", "/")
                End If
                lngValue = lng次数
                If Trim(strOpdays(i)) <> "" And lngValue + UBound(arrOperDay) < 12 Then
                    strTmp = "": strTmp1 = ""
                    For j = UBound(arrOperDay) + 1 To 1 Step -1
                        lng次数 = lngValue + j
                        strTmp1 = Switch(lng次数 = 1, "Ⅰ", lng次数 = 2, "Ⅱ", lng次数 = 3, "Ⅲ", lng次数 = 4, "Ⅳ", lng次数 = 5, "Ⅴ", lng次数 = 6, _
                            "Ⅵ", lng次数 = 7, "Ⅶ", lng次数 = 8, "Ⅷ", lng次数 = 9, "Ⅸ", lng次数 = 10, "Ⅹ", lng次数 = 11, "Ⅺ", lng次数 = 12, "Ⅻ")
                        If strTmp = "" Then
                            strTmp = strTmp1
                        Else
                            strTmp = strTmp & "/" & strTmp1
                        End If
                        If blnStopFlag Then Exit For
                    Next j
                    lng次数 = lngValue + UBound(arrOperDay) + 1
                    If blnStopFlag Then '手术标注后天数在次手术时停止前一次标注
                        Select Case intOpFormat
                            Case 1 '显示0
                                strOpValue(i) = 0
                            Case 2 '显示手术次数
                                If strTmp = "Ⅰ" Then
                                    strOpValue(i) = 0
                                Else
                                    strOpValue(i) = strTmp & "-0"
                                End If
                            Case Else '不显示
                                strOpValue(i) = ""
                        End Select
                    Else
                        Select Case intOpFormat
                            Case 1 '显示0
                                If Trim(strOpValue(i)) <> "" Then
                                    strOpValue(i) = 0 & "/" & strOpValue(i)
                                Else
                                    strOpValue(i) = 0
                                End If
                            Case 2 '显示手术次数
                                If Trim(strOpValue(i)) <> "" Then
                                    strOpValue(i) = strTmp & "/" & strOpValue(i)
                                Else
                                    strOpValue(i) = strTmp
                                End If
                            Case Else '不显示
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
        
        stdset.Name = "宋体"
        stdset.Size = 9 * sngScale
        stdset.Bold = False
        Call SetFontIndirect(stdset, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        '3开始输出住院日期，天数，手术信息
        Call DrawUpTable(lngDC, objDraw, strTmpString0 & "||" & strTmpString2 & "||" & strTmpString1, lngCurX, lngCurY, T_DrawClient.体温区域.Right, lngCurY, sngScale)
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
        
        '----------------------------------------------------------------------------------------------
         '此处计算可打印区域 从而计算体温单打印的行高
        If intRepairRows = 0 Then
            sngHTab = intRepairRows
        Else
            sngHTab = intRepairRows * T_DrawClient.时间列单位 + IIf(bln呼吸 = True, T_DrawClient.时间列单位 / 2, 0)
        End If
        
        sngHTab = sngHTab + msngTwips * 30 + 10
        sngHPrint = objDraw.Height / T_TwipsPerPixel.Y - lngCurY - lngButtom - sngHTab - dblSureH
        T_DrawClient.行单位 = (sngHPrint - 4 * T_DrawClient.行单位) / T_DrawClient.总列数
        T_DrawClient.行单位 = Round(T_DrawClient.行单位 - 0.05, 1)
        If T_DrawClient.行单位 > 6 * msngTwips Then T_DrawClient.行单位 = 6 * msngTwips
        If T_DrawClient.行单位 < 6 * msngTwips Then T_DrawClient.行单位 = 6 * msngTwips
        
        '计算行高后在计算体温单可打印的表格行数
        If intRepairRows > 0 Then
            sngHPrint = T_DrawClient.总列数 * T_DrawClient.行单位 + 4 * T_DrawClient.行单位
            sngHTab = objDraw.Height / T_TwipsPerPixel.Y - lngCurY - lngButtom - dblSureH - sngHPrint - (msngTwips * 30 + 10)
            sngHTab = sngHTab - IIf(bln呼吸 = True, T_DrawClient.时间列单位 / 2, 0)
            If Fix(sngHTab / T_DrawClient.时间列单位 + 0.3) < intRepairRows Then intRepairRows = Fix(sngHTab / T_DrawClient.时间列单位 + 0.3)
        End If
    
        stdset.Name = "宋体"
        stdset.Size = 9 * sngScale
        stdset.Bold = False
        Call SetFontIndirect(stdset, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        '4开始画刻度区域和体温区域并输出刻度值信息
        T_DrawClient.偏移量Y = lngCurY
        mbln呼吸曲线 = False
        
        rsTemp.Filter = 0
        rsTemp.Sort = "排列序号"
        rsTemp.MoveFirst
        str体温说明 = DrawCanvas(lngDC, objDraw, rsTemp, rsDrawItems, bln不打印心率列, sngScale)
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
        
        '5.读取病人体温数据和入出转等标记信息
        '初始化 体温点记录集和入出转等标记信息
        
        '所有点的表现集合
        '   重叠是否重叠序号.
        '   重叠项目记录重叠项目
        '   断开的条件:超过一天无数据,存在未记说明
        '   备注:物理降温时记录原值
        '   符号:用来标注体温不升，或者值小于等于项目最小值大于等于项目最大值是的特殊符号.此外默认为空

        gstrFields = "序号," & adDouble & ",18|数值," & adLongVarChar & ",4000|部位," & adLongVarChar & ",200|" & _
             "标记," & adDouble & ",1|时间," & adLongVarChar & ",20|项目序号," & adDouble & ",18|" & _
             "复查," & adDouble & ",1|断开," & adDouble & ",1|重叠项目," & adLongVarChar & ",50|" & _
             "重叠," & adDouble & ",5|X坐标," & adDouble & ",5|Y坐标," & adDouble & ",5|备注," & adLongVarChar & ",50|" & _
             "符号," & adLongVarChar & ",10|显示," & adDouble & ",1"
        Call Record_Init(rsPoints, gstrFields)
    
        '所有需要输出的文本内容(类型:2-上标;3-入出转;4-手术日;6-下标,13-出生,99-未记说明)
        '禁用表示信息是否输出
        gstrFields = "时间," & adLongVarChar & ",20|项目序号," & adDouble & ",18|类型," & adDouble & ",2|" & _
            "内容," & adLongVarChar & ",200|颜色," & adLongVarChar & ",20|X坐标," & adDouble & ",20|" & _
            "Y坐标," & adDouble & ",20|高度," & adDouble & ",20|打印X坐标," & adDouble & ",20|" & _
            "禁用," & adInteger & ",1|显示," & adDouble & ",1"
        Call Record_Init(rsNotes, gstrFields)
        
        Dim rs脉搏 As New ADODB.Recordset
        Dim strFileds As String, strValues As String
        
        '记录脉搏信息
        strFileds = "项目序号," & adDouble & ",18|数值," & adLongVarChar & ",4000|X坐标," & adDouble & ",5|时间," & adLongVarChar & ",20"
        Call Record_Init(rs脉搏, strFileds)
        
        Dim int标记 As Integer
        
        '----提取所有部位信息
        strSql = "select 项目序号,部位,缺省项 from 体温部位"
        Call zldatabase.OpenRecordset(rsPart, strSql, "体温部位")
        '----读取病人体温数据和未记说明
        strSql = "SELECT C.ID 序号, a.发生时间 As 时间,C.显示,C.记录内容 As 数值,C.体温部位,c.复试合格,D.记录名,E.保留项目,D.项目序号,DECODE(D.项目序号,-1,1,C.记录标记) 记录标记,C.未记说明 " & _
                    "FROM 病人护理文件 B,病人护理数据 A,病人护理明细 C,体温记录项目 D,护理记录项目 E " & _
                    "Where B.ID=A.文件ID  " & _
                        "AND A.ID = C.记录ID " & _
                        "AND B.ID=[1] " & _
                        "AND Nvl(B.婴儿,0)=[6] " & _
                        "AND B.病人id=[2] " & _
                        "AND B.主页id=[3] " & _
                        "AND D.项目序号=C.项目序号 " & _
                        "AND C.记录类型=1 " & _
                        "AND E.项目序号=D.项目序号 " & _
                        "AND E.护理等级>=[7]  " & _
                        "AND A.发生时间 BETWEEN [4] And [5] And C.终止版本 Is Null " & _
                        "AND D.记录法=1 AND (nvl(E.应用方式,0)=1 OR ( -1=[10] and nvl(E.应用方式,0)=2)) " & _
                        "AND nvl(E.适用病人,0) in (0,[8]) AND (E.适用科室=1 or ( E.适用科室=2 AND Exists (select 1 from 护理适用科室 D where D.项目序号=E.项目序号 and D.科室ID=[9])))" & _
                    "Order By a.发生时间,DECODE(D.项目序号,-1,1,0),DECODE(D.项目序号,-1,1,C.记录标记)"
        If blnMoved Then
            strSql = Replace(strSql, "病人护理文件", "H病人护理文件")
            strSql = Replace(strSql, "病人护理数据", "H病人护理数据")
            strSql = Replace(strSql, "病人护理明细", "H病人护理明细")
        End If

        Set rsTmp = zldatabase.OpenSQLRecord(strSql, "读取曲线项目数据", lng文件ID, lng病人ID, lng主页ID, CDate(strTmpDay), CDate(strEndDay), _
            intBaby, lng护理等级, IIf(intBaby = 0, 1, 2), lngSectID, IIf(int心率应用 = 2, -1, 0))
         
        strTmpString0 = ""
        strTmpString1 = ""
        strTmpString2 = ""
        With rsTmp
            Do While Not .EOF
                strTmp = ""
                blnAllow = False
                strPart = zlCommFun.Nvl(!体温部位)
                lng项目序号 = Val(zlCommFun.Nvl(!项目序号))
                Select Case lng项目序号
                    Case gint心率
                        int标记 = 1
                    Case Else
                        int标记 = Val(zlCommFun.Nvl(!记录标记))
                End Select
                If strPart = "" Then
                    rsPart.Filter = "项目序号=" & lng项目序号 & " and 缺省项=1"
                    If rsPart.BOF = False Then
                        strPart = zlCommFun.Nvl(rsPart!部位)
                    Else
                        Select Case lng项目序号
                            Case gint体温
                                strPart = "腋温"
                            Case gint呼吸
                                strPart = "自主呼吸"
                            Case Else
                                strPart = ""
                        End Select
                    End If
                End If
                
                SinX = GetXCoordinate(Format(!时间, "YYYY-MM-DD HH:mm:ss"), Format(strTmpDay, "YYYY-MM-DD HH:mm:ss"))
                strTime = GetXCoordinate(SinX, Format(strTmpDay, "YYYY-MM-DD HH:mm:ss"), False)
                SinX = GetXCoordinate(Format(Split(strTime, ",")(0), "YYYY-MM-DD HH:mm:ss"), Format(strTmpDay, "YYYY-MM-DD HH:mm:ss"))
                
                '记录所有脉搏信息
                If lng项目序号 = gint脉搏 Then
                    strFileds = "项目序号|数值|X坐标|时间"
                    strValues = lng项目序号 & "|" & zlCommFun.Nvl(!数值) & "|" & SinX & "|" & Format(!时间, "yyyy-MM-dd HH:mm:ss")
                    Call Record_Add(rs脉搏, strFileds, strValues)
                End If
                
                If (Not IsNull(!未记说明)) And zlCommFun.Nvl(!数值) <> "不升" Then
                    rsNotes.Filter = "项目序号=" & Val(zlCommFun.Nvl(!项目序号)) & " AND X坐标=" & SinX
                    blnAdd = (rsNotes.RecordCount = 0)
                    '所有需要输出的文本内容(类型:2-上标;3-入出转;4-手术日;6-下标,99-未记说明)
                    gstrFields = "时间|项目序号|类型|内容|颜色|X坐标|Y坐标|高度|打印X坐标|禁用|显示"  '入出转缺省是红色,上下标及未记说明缺省是蓝色
                    gstrValues = Format(!时间, "yyyy-MM-dd HH:mm:ss") & "|" & !项目序号 & "|99|" & _
                        !未记说明 & "|" & RGB_BLUE & "|" & SinX & "|0|0|0|0|" & zlCommFun.Nvl(!显示)
                   
                    If blnAdd Then
                        '提取接近中间时间点的值做为本列值
                         Call Record_Add(rsNotes, gstrFields, gstrValues)
                    Else
                        If (zlCommFun.Nvl(rsNotes!显示, 0) = 1 And zlCommFun.Nvl(!显示, 0) = 1) Or (zlCommFun.Nvl(rsNotes!显示, 0) <> 1 And zlCommFun.Nvl(!显示, 0) <> 1) Then
                             blnAllow = GetCanvasCenter(CDate(Format(rsNotes!时间, "YYYY-MM-DD HH:mm:ss")), CDate(Format(!时间, "YYYY-MM-DD HH:mm:ss")), CDate(Format(strTmpDay, "YYYY-MM-DD HH:mm:ss")), SinX)
                        ElseIf zlCommFun.Nvl(!显示, 0) = 1 Then
                            blnAllow = True
                        End If
    
                        If blnAllow = True Then
                            If Val(rsNotes!显示) = 2 Then
                                arrValues = Split(gstrValues, "|")
                                arrValues(UBound(arrValues)) = 2
                                gstrValues = Join(arrValues, "|")
                            End If
                            Call Record_Update(rsNotes, gstrFields, gstrValues, "时间|" & Format(rsNotes!时间, "yyyy-MM-dd HH:mm:ss"))
                        Else
                            If Val(zlCommFun.Nvl(!显示, 0)) = 2 Then
                                gstrFields = "显示"
                                gstrValues = "2"
                                Call Record_Update(rsNotes, gstrFields, gstrValues, "时间|" & Format(rsNotes!时间, "yyyy-MM-dd HH:mm:ss"))
                            End If
                        End If
                    End If
                Else
                    blnAdd = False
                    
                    rsPoints.Filter = "项目序号=" & lng项目序号 & " AND X坐标=" & SinX & " And 标记=" & int标记
                    
                    blnAdd = (rsPoints.RecordCount = 0)
                    
                    dbl数值 = Val(zlCommFun.Nvl(!数值))
                    
                    For i = 0 To UBound(strEditors)
                        If Val(Split(strEditors(i), "||")(0)) = lng项目序号 Then
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
                    
                    '临界值不等空,并且在最大值和最小值之间
                    If Split(strEditors(i), "||")(8) <> "" And Val(Split(strEditors(i), "||")(8)) <= Val(Split(strEditors(i), "||")(6)) _
                        And Val(Split(strEditors(i), "||")(8)) >= Val(Split(strEditors(i), "||")(7)) Then dblMaxValue = Val(Split(strEditors(i), "||")(8))
                    
                    '不指定符号，项目数据操作最大值和最小值以项目本身符号显示
                    If dbl数值 <= dblMinValue Then
                        dbl数值 = dblMinValue
                        'strTmp = "·"
                    End If
                    
                    
                    If dbl数值 >= dblMaxValue Then
                        dbl数值 = dblMaxValue
                        'strTmp = "·"
                    End If
                    
                     '体温不升是在显示在35刻度
                    If Trim(Nvl(!数值)) = "不升" And lng项目序号 = gint体温 Then dbl数值 = 35
                    
                    sinY = Val(GetYCoordinate(objDraw, rsDrawItems, !项目序号, dbl数值, lngDC, True))
                    
                    gstrFields = "序号|数值|部位|标记|时间|项目序号|复查|断开|重叠项目|重叠|X坐标|Y坐标|备注|符号|显示"
                    gstrValues = Val(zlCommFun.Nvl(!序号)) & "|" & !数值 & "|" & strPart & "|" & int标记 & "|" & _
                                 Format(!时间, "yyyy-MM-dd HH:mm:ss") & "|" & lng项目序号 & "|" & Val(zlCommFun.Nvl(!复试合格, 0)) & "|" & IIf(zlCommFun.Nvl(!数值, 0) = "不升", 1, 0) & "|空|0|" & _
                                 SinX & "|" & sinY & "||" & strTmp & "|" & zlCommFun.Nvl(!显示, 0)
                    If blnAdd Then '添加
                        Call Record_Add(rsPoints, gstrFields, gstrValues)
                    Else
                        If (zlCommFun.Nvl(rsPoints!显示, 0) = 1 And zlCommFun.Nvl(!显示, 0) = 1) Or (zlCommFun.Nvl(rsPoints!显示, 0) <> 1 And zlCommFun.Nvl(!显示, 0) <> 1) Then
                            blnAllow = GetCanvasCenter(CDate(Format(rsPoints!时间, "YYYY-MM-DD HH:mm:ss")), CDate(Format(!时间, "YYYY-MM-DD HH:mm:ss")), CDate(Format(strTmpDay, "YYYY-MM-DD HH:mm:ss")), SinX)
                        ElseIf zlCommFun.Nvl(!显示, 0) = 1 Then
                            blnAllow = True
                        End If
                        
                       '提取接近中间时间点的值做为本列值
                        If blnAllow = True Then
                            If Val(rsPoints!显示) = 2 Then
                                arrValues = Split(gstrValues, "|")
                                arrValues(UBound(arrValues)) = 2
                                gstrValues = Join(arrValues, "|")
                            End If
                            Call Record_Update(rsPoints, gstrFields, gstrValues, "序号|" & rsPoints!序号)
                        Else
                            If Val(zlCommFun.Nvl(!显示, 0)) = 2 Then
                                gstrFields = "显示"
                                gstrValues = "2"
                                Call Record_Update(rsPoints, gstrFields, gstrValues, "序号|" & rsPoints!序号)
                            End If
                        End If
                    End If
                End If
            .MoveNext
            Loop
        End With
                
        '上面已经得到了所有项目的数据信息，下来处理物理降温和脉搏和心率数据
        rsPoints.Filter = ""
        arrTmpValue = Array()
        If int心率应用 = 2 Then
            rsPoints.Filter = "项目序号=" & gint心率
            With rsPoints
                Do While Not .EOF
                    ReDim Preserve arrTmpValue(UBound(arrTmpValue) + 1)
                    arrTmpValue(UBound(arrTmpValue)) = !序号 & ";" & !项目序号 & ";" & !X坐标 & ";" & Format(!时间, "yyyy-MM-DD HH:mm:ss")
                .MoveNext
                Loop
            End With
        End If
        
        '心率设为脉搏共用时，检查脉搏是否设置为可用
        If int脉搏 <> -1 Then
            For i = 0 To UBound(arrTmpValue)
                '检查心率是否与脉搏相对应
                rs脉搏.Filter = "项目序号=" & gint脉搏 & " And X坐标=" & Val(Split(CStr(arrTmpValue(i)), ";")(2)) & " And 时间='" & Format(Split(CStr(arrTmpValue(i)), ";")(3), "yyyy-MM-DD HH:mm:ss") & "'"
                
                rsPoints.Filter = "项目序号=" & gint脉搏 & " and X坐标=" & Val(Split(CStr(arrTmpValue(i)), ";")(2)) & " And 时间='" & Format(Split(CStr(arrTmpValue(i)), ";")(3), "yyyy-MM-DD HH:mm:ss") & "'"
                If rsPoints.RecordCount = 0 Then
                    If rs脉搏.RecordCount = 0 Then
                        rsPoints.Filter = ""
                        gstrFields = "项目序号": gstrValues = gint脉搏
                        Call Record_Update(rsPoints, gstrFields, gstrValues, "序号|" & Val(Split(CStr(arrTmpValue(i)), ";")(0)))
                    Else
                        rsPoints.Filter = "项目序号=" & gint心率 & " And X坐标=" & Val(Split(CStr(arrTmpValue(i)), ";")(2)) & " And 时间='" & Format(Split(CStr(arrTmpValue(i)), ";")(3), "yyyy-MM-DD HH:mm:ss") & "'"
                        rsPoints.Delete
                    End If
                End If
            Next i
        End If
        
        If int心率应用 = 2 Then
            Set rs脉搏 = New ADODB.Recordset
            strFileds = "序号," & adDouble & ",18|数值," & adLongVarChar & ",4000|部位," & adLongVarChar & ",200|" & _
                        "标记," & adDouble & ",1|时间," & adLongVarChar & ",20|项目序号," & adDouble & ",18|" & _
                        "复查," & adDouble & ",1|断开," & adDouble & ",1|重叠项目," & adLongVarChar & ",50|" & _
                        "重叠," & adDouble & ",5|X坐标," & adDouble & ",5|Y坐标," & adDouble & ",5|备注," & adLongVarChar & ",50|" & _
                        "符号," & adLongVarChar & ",10|显示," & adDouble & ",1"
            Call Record_Init(rs脉搏, strFileds)
            
            rsPoints.Filter = "项目序号=" & gint脉搏
            With rsPoints
                Do While Not .EOF
                    rs脉搏.AddNew
                    For i = 0 To .Fields.Count - 1
                        rs脉搏.Fields(.Fields(i).Name).Value = .Fields(i).Value
                    Next i
                    rs脉搏.Update
                .MoveNext
                Loop
            End With
            
            rsPoints.Filter = "项目序号=" & gint脉搏
            Do While Not rsPoints.EOF
                rsPoints.Delete
                rsPoints.MoveNext
            Loop
            
            rs脉搏.Filter = ""
            rs脉搏.Sort = "时间"
            With rs脉搏
                Do While Not .EOF
                    blnAdd = False
                    blnAllow = False
                    
                    SinX = Val(zlCommFun.Nvl(!X坐标))
                    sinY = Val(zlCommFun.Nvl(!Y坐标))
                    rsPoints.Filter = "项目序号=" & Val(zlCommFun.Nvl(!项目序号, 0)) & " AND X坐标=" & SinX
                    blnAdd = IIf(rsPoints.RecordCount = 0, True, False)
                    
                    strFileds = "序号|数值|部位|标记|时间|项目序号|复查|断开|重叠项目|重叠|X坐标|Y坐标|备注|符号|显示"
                    strValues = Val(zlCommFun.Nvl(!序号)) & "|" & !数值 & "|" & zlCommFun.Nvl(!部位) & "|" & Val(zlCommFun.Nvl(!标记, 0)) & "|" & _
                                 Format(!时间, "yyyy-MM-dd HH:mm:ss") & "|" & Val(zlCommFun.Nvl(!项目序号)) & "|0|" & Val(zlCommFun.Nvl(!断开)) & "|空|0|" & _
                                 SinX & "|" & sinY & "||" & zlCommFun.Nvl(!符号) & "|" & Val(zlCommFun.Nvl(!显示, 0))
                    
                    If blnAdd Then '添加
                        Call Record_Add(rsPoints, strFileds, strValues)
                    Else
                        If (zlCommFun.Nvl(rsPoints!显示, 0) = 1 And zlCommFun.Nvl(!显示, 0) = 1) Or (zlCommFun.Nvl(rsPoints!显示, 0) <> 1 And zlCommFun.Nvl(!显示, 0) <> 1) Then
                            blnAllow = GetCanvasCenter(CDate(Format(rsPoints!时间, "YYYY-MM-DD HH:mm:ss")), CDate(Format(!时间, "YYYY-MM-DD HH:mm:ss")), CDate(Format(strTmpDay, "YYYY-MM-DD HH:mm:ss")), SinX)
                        ElseIf zlCommFun.Nvl(!显示, 0) = 1 Then
                            blnAllow = True
                        End If
                        
                        '提取接近中间时间点的值做为本列值
                        If blnAllow = True Then
                            If Val(rsPoints!显示) = 2 Then
                                arrValues = Split(strValues, "|")
                                arrValues(UBound(arrValues)) = 2
                                strValues = Join(arrValues, "|")
                            End If
                            Call Record_Update(rsPoints, strFileds, strValues, "序号|" & rsPoints!序号)
                        Else
                            If Val(zlCommFun.Nvl(!显示, 0)) = 2 Then
                                strFileds = "显示"
                                strValues = "2"
                                Call Record_Update(rsPoints, strFileds, strValues, "序号|" & rsPoints!序号)
                            End If
                        End If
                    End If
                .MoveNext
                Loop
            End With
        End If
        
        '处理物理降温数据
        arrTmpValue = Array()
        rsPoints.Filter = "项目序号=1 and 标记=0"
        With rsPoints
            Do While Not .EOF
                ReDim Preserve arrTmpValue(UBound(arrTmpValue) + 1)
                arrTmpValue(UBound(arrTmpValue)) = !序号 & ";" & !项目序号 & ";" & !数值 & ";" & !X坐标 & ";" & !Y坐标 & ";" & Format(!时间, "yyyy-MM-dd HH:mm:ss")
            .MoveNext
            Loop
        End With
        
        rsPoints.Filter = "项目序号=1"
        If rsPoints.RecordCount > 0 Then rsPoints.MoveFirst
        For i = 0 To UBound(arrTmpValue)
            rsPoints.Filter = "项目序号=1 and 标记=1 and X坐标=" & Val(Split(CStr(arrTmpValue(i)), ";")(3)) & " And 时间='" & Format(Split(CStr(arrTmpValue(i)), ";")(5), "yyyy-MM-dd HH:mm:ss") & "'"
            If rsPoints.RecordCount <> 0 Then
                gstrFields = "备注": gstrValues = Val(Split(CStr(arrTmpValue(i)), ";")(2)) & "," & Val(Split(CStr(arrTmpValue(i)), ";")(3)) & ";" & Val(Split(CStr(arrTmpValue(i)), ";")(4))
                Call Record_Update(rsPoints, gstrFields, gstrValues, "序号|" & zlCommFun.Nvl(rsPoints!序号))
            End If
        Next i
        
        arrTmpValue = Array()
        rsPoints.Filter = "项目序号=1 and 标记=1"
        With rsPoints
            Do While Not .EOF
                ReDim Preserve arrTmpValue(UBound(arrTmpValue) + 1)
                arrTmpValue(UBound(arrTmpValue)) = !序号 & ";" & !项目序号 & ";" & !数值 & ";" & !X坐标 & ";" & !Y坐标 & ";" & Format(!时间, "yyyy-MM-dd HH:mm:ss")
            .MoveNext
            Loop
        End With
        
        rsPoints.Filter = "项目序号=1"
        If rsPoints.RecordCount > 0 Then rsPoints.MoveFirst
        For i = 0 To UBound(arrTmpValue)
            rsPoints.Filter = "项目序号=1 and 标记=0 and X坐标=" & Val(Split(CStr(arrTmpValue(i)), ";")(3)) & " And 时间='" & Format(Split(CStr(arrTmpValue(i)), ";")(5), "yyyy-MM-dd HH:mm:ss") & "'"
            If rsPoints.RecordCount = 0 Then
                rsPoints.Filter = "项目序号=1 and 标记=1 and X坐标=" & Val(Split(CStr(arrTmpValue(i)), ";")(3)) & " And 时间='" & Format(Split(CStr(arrTmpValue(i)), ";")(5), "yyyy-MM-dd HH:mm:ss") & "'"
                rsPoints.Delete
            End If
        Next i
    
        '删除显示为2的数据
        rsPoints.Filter = ""
        rsPoints.Filter = "显示=2"
        Do While Not rsPoints.EOF
            rsPoints.Delete
        rsPoints.MoveNext
        Loop
        
        rsNotes.Filter = ""
        rsNotes.Filter = "显示=2"
        Do While Not rsNotes.EOF
            rsNotes.Delete
        rsNotes.MoveNext
        Loop
        
        '处理未记说明和曲线数据该显示那一条
        rsNotes.Filter = ""
        rsPoints.Filter = ""
        
        arrTmpValue = Array()
        arrTmpNote = Array()
        rsNotes.Sort = "项目序号,X坐标"
        With rsNotes
            Do While Not .EOF
                SinX = Val(!X坐标)
                blnAllow = False
                rsPoints.Filter = "项目序号=" & Val(!项目序号) & " And X坐标=" & SinX
                If rsPoints.RecordCount > 0 Then
                    If (zlCommFun.Nvl(rsPoints!显示, 0) = 1 And zlCommFun.Nvl(!显示, 0) = 1) Or (zlCommFun.Nvl(rsPoints!显示, 0) <> 1 And zlCommFun.Nvl(!显示, 0) <> 1) Then
                        blnAllow = GetCanvasCenter(CDate(Format(rsPoints!时间, "YYYY-MM-DD HH:mm:ss")), CDate(Format(!时间, "YYYY-MM-DD HH:mm:ss")), CDate(Format(strTmpDay, "YYYY-MM-DD HH:mm:ss")), SinX)
                    ElseIf zlCommFun.Nvl(!显示, 0) = 1 Then
                        blnAllow = True
                    End If
                    If blnAllow = True Then
                        ReDim Preserve arrTmpValue(UBound(arrTmpValue) + 1)
                        arrTmpValue(UBound(arrTmpValue)) = !项目序号 & ";" & SinX
                    Else
                        ReDim Preserve arrTmpNote(UBound(arrTmpNote) + 1)
                        arrTmpNote(UBound(arrTmpNote)) = !项目序号 & ";" & SinX
                    End If
                End If
            .MoveNext
            Loop
        End With
        
        For i = 0 To UBound(arrTmpValue)
            rsPoints.Filter = "项目序号=" & Val(Split(CStr(arrTmpValue(i)), ";")(0)) & " And X坐标=" & Val(Split(CStr(arrTmpValue(i)), ";")(1))
            Do While Not rsPoints.EOF
                rsPoints.Delete
            rsPoints.MoveNext
            Loop
        Next i
        
        For i = 0 To UBound(arrTmpNote)
            rsNotes.Filter = "项目序号=" & Val(Split(CStr(arrTmpNote(i)), ";")(0)) & " And X坐标=" & Val(Split(CStr(arrTmpNote(i)), ";")(1))
            Do While Not rsNotes.EOF
                rsNotes.Delete
            rsNotes.MoveNext
            Loop
        Next i
    
'        '处理体温不升 体温为不升需要在35度下纵向输出体温不升二字
        rsPoints.Filter = "项目序号=" & gint体温 & " and 数值='不升' and 标记<>1"
        rsPoints.Sort = "时间"
        With rsPoints
            Do While Not .EOF
                strTmpString0 = strTmpString0 & ";" & Format(!时间, "yyyy-MM-dd HH:mm:ss") & "|" & Val(zlCommFun.Nvl(!项目序号)) & "|99|" & _
                      "不升|" & RGB_BLUE & "|" & !X坐标 & "|0|0|0|0"
                strTmpString2 = strTmpString2 & ";" & !X坐标
            .MoveNext
            Loop
        End With
        
        '--------更新断开标记
        '两点之间有未记说明断开，时间操作一天断开,体温不升断开
        rsPoints.Filter = ""
        
        gstrFields = "断开"
        gstrValues = "1"
        rsNotes.Filter = ""
        
        If rsNotes.RecordCount > 0 Then rsNotes.MoveFirst
        With rsNotes
            Do While Not .EOF
                If int心率应用 = 2 And !项目序号 = -1 Then
                    rsPoints.Filter = "项目序号=" & gint脉搏 & " And X坐标<=" & !X坐标
                Else
                    If !项目序号 = 1 Then
                        rsPoints.Filter = "项目序号=" & !项目序号 & " And  标记<>1 And X坐标<" & !X坐标
                    Else
                        rsPoints.Filter = "项目序号=" & !项目序号 & " And X坐标<" & !X坐标
                    End If
                End If
                rsPoints.Sort = "时间"
                If rsPoints.RecordCount <> 0 Then
                    rsPoints.MoveLast
                    Call Record_Update(rsPoints, gstrFields, gstrValues, "序号|" & rsPoints!序号)
                End If
      
            .MoveNext
            Loop
        End With
        '时间超过一天
        strTime = ""
        strTmp = ""
        rsPoints.Filter = ""
        
        rsPoints.Sort = "项目序号,时间,标记"
        With rsPoints
            Do While Not .EOF
                If Not IsNull(!序号) Then
                    If Not (Val(!项目序号) = 1 And Val(!标记) = 1) Then
                        If lng项目序号 <> 0 Then
                            If lng项目序号 <> !项目序号 Then strTime = ""
                        End If
                        lng项目序号 = !项目序号
                        If strTime <> "" Then
                            If DateDiff("D", CDate(strTime), CDate(Format(!时间, "YYYY-MM-DD"))) > 1 Then
                                strTmp = strTmp & "," & lngValue
                            End If
                        End If
                        strTime = Format(rsPoints!时间, "YYYY-MM-DD")
                        lngValue = Val(rsPoints!序号)
                    End If
                End If
                .MoveNext
            Loop
        End With
        
        strTmp = Mid(strTmp, 2)
        For i = 0 To UBound(Split(strTmp, ","))
            Call Record_Update(rsPoints, gstrFields, gstrValues, "序号|" & Split(strTmp, ",")(i))
        Next i
        
        '处理体温不升的.把前一个点的断开标志设置为1
        rsPoints.Filter = ""
        rsPoints.Filter = "项目序号=" & gint体温 & " and 标记<>1"
        rsPoints.Sort = "时间,标记"
        With rsPoints
            Do While Not .EOF
                If !数值 = "不升" And .AbsolutePosition <> 1 Then
                    .MovePrevious '更新上一行断开标记
                    If Val(!断开) <> 1 Then
                        lngValue = !序号
                        Call Record_Update(rsPoints, gstrFields, gstrValues, "序号|" & lngValue)
                    End If
                    .MoveNext
                End If
            .MoveNext
            Loop
        End With
    
        '重新整理未及说明，同一X坐标有相同的说明值输出一次
        rsNotes.Filter = ""
        rsNotes.Sort = "X坐标"
        With rsNotes
            Do While Not .EOF
                If lngValue = !X坐标 Then
                    If InStr(1, "," & strTmp & ",", "," & zlCommFun.Nvl(!内容) & ",") <> 0 Then
                       rsNotes.Delete
                    Else
                        strTmp = strTmp & "," & zlCommFun.Nvl(!内容)
                    End If
                Else
                    lngValue = !X坐标
                    strTmp = zlCommFun.Nvl(!内容)
                End If
            .MoveNext
            Loop
        End With
        
        '--提取入出院,手术等标记说明
        Dim bytShow As Byte
        Dim str内容 As String
        Dim lng行号 As Long, lngColor As Long
        
        '读取手术、上下标信息
        '-----------------------------------------------------------------------
        gstrFields = "时间|项目序号|类型|内容|颜色|X坐标|Y坐标|高度|打印X坐标|禁用"  '入出转缺省是红色,上下标及未记说明缺省是蓝色
        strSql = "" & _
                 " Select B.发生时间 AS 时间,C.记录类型,C.项目序号,C.记录内容,C.项目名称,C.未记说明" & _
                 " FROM 病人护理文件 A, 病人护理数据 B, 病人护理明细 C" & _
                 " Where A.ID=B.文件ID and  B.ID = C.记录ID AND A.ID=[1]   AND Nvl(A.婴儿, 0)=[6] AND A.病人id=[2] AND A.主页id=[3] And c.终止版本 Is Null" & _
                 " AND mod(c.记录类型,10) <> 1  AND B.发生时间 BETWEEN [4]  And [5]"
        If blnMoved Then
            strSql = Replace(strSql, "病人护理文件", "H病人护理文件")
            strSql = Replace(strSql, "病人护理数据", "H病人护理数据")
            strSql = Replace(strSql, "病人护理明细", "H病人护理明细")
        End If

        Set rsTmp = zldatabase.OpenSQLRecord(strSql, "读取手术、上下标等信息", lng文件ID, lng病人ID, lng主页ID, Int(CDate(strTmpDay)), CDate(strEndDay), intBaby, lng护理等级)
        With rsTmp
            Do While Not .EOF
                bytShow = 1
                str内容 = Trim(zlCommFun.Nvl(!记录内容))
               
                lng行号 = IIf(!记录类型 = 2, 10, IIf(!记录类型 = 6, 11, 4))
                
                '对于手术显示需要特殊处理
                If !记录类型 = 4 Then
                    str内容 = Trim(zlCommFun.Nvl(!项目名称))
                    
                    If str内容 = "分娩" Then
                        bytShow = T_BodyFlag.分娩
                    Else
                        bytShow = T_BodyFlag.手术
                    End If
                    
                    If bytShow = 2 Then
                        str内容 = str内容 & gstrCaveSplit & ConvertTimeToChinese(Format(!时间, "HH:mm"))
                    Else
                        str内容 = !项目名称
                    End If
                    lngColor = RGB_RED
                Else
                    lngColor = IIf(Not IsNumeric(Nvl(!未记说明)), RGB_BLUE, Val(Nvl(!未记说明)))
                End If
                
                If bytShow > 0 Then
                    SinX = Val(GetXCoordinate(Format(!时间, "YYYY-MM-DD HH:mm:ss"), strTmpDay))
                    
                    rsNotes.Filter = "X坐标=" & SinX & " and 项目序号=" & lng行号 & " and 类型=" & !记录类型
                    If rsNotes.BOF Then
                        gstrValues = Format(!时间, "yyyy-MM-dd HH:mm:ss") & "|" & lng行号 & "|" & !记录类型 & "|" & _
                            str内容 & "|" & lngColor & "|" & SinX & "|0|0|0|0"
                        Call Record_Add(rsNotes, gstrFields, gstrValues)
                    Else
                        rsNotes!时间 = Format(!时间, "yyyy-MM-dd HH:mm:ss")
                        rsNotes!内容 = str内容
                        rsNotes.Update
                    End If
                End If
                rsNotes.Filter = ""
                .MoveNext
            Loop
        End With
        
        '读取入出转等信息
        '-----------------------------------------------------------------------
        '所有需要输出的文本内容(类型:2-上标;3-入出转;4-手术日;6-下标,99-未记说明)
        '1-入院；2-入科；3-转科；4-换床
        gstrFields = "时间|项目序号|类型|内容|颜色|X坐标|Y坐标|高度|打印X坐标|禁用"  '入出转缺省是红色,上下标及未记说明缺省是蓝色
        Set rsTmp = GetDataFromHis(lng病人ID, lng主页ID, intBaby, CDate(strTmpDay), CDate(strEndDay), 2)
        With rsTmp
            Do While Not .EOF
                If Trim(zlCommFun.Nvl(!内容)) <> "" Then
                    bytShow = 0
                    lng行号 = Val(!行号)
                    str内容 = zlCommFun.Nvl(!内容)
                    Select Case Val(!行号)
                    Case 5
                        bytShow = T_BodyFlag.入院
                    Case 6, 3 '6转入，3转出
                        bytShow = T_BodyFlag.转出
                    Case 7
                        bytShow = T_BodyFlag.换床
                    Case 8
                        bytShow = T_BodyFlag.出院
                        If intBaby > 0 Then
                            bytShow = IIf(bln婴儿体温单显示出院, bytShow, 0)
                        End If
                    Case 9
                        bytShow = T_BodyFlag.入科
                    End Select
                    
                    If bytShow > 0 Then
                        If lng行号 = 9 And bln入科显示入院 = True Then str内容 = "入院"
                        '目前3，4 针对于转科 3-显示说明和科室 4 显示说明，科室，时间
                        If bytShow = 2 Then
                            str内容 = str内容 & gstrCaveSplit & ConvertTimeToChinese(Format(!时间, "HH:mm"))
                        ElseIf bytShow = 3 Then
                            str内容 = str内容 & gstrCaveSplit & zlCommFun.Nvl(!科室)
                        ElseIf bytShow = 4 Then
                            str内容 = str内容 & gstrCaveSplit & zlCommFun.Nvl(!科室) & gstrCaveSplit & ConvertTimeToChinese(Format(!时间, "HH:mm"))
                        ElseIf bytShow = 1 Then
                            str内容 = str内容
                        End If
                        
                        SinX = Val(GetXCoordinate(Format(!时间, "YYYY-MM-DD HH:mm:ss"), strTmpDay))
                        rsNotes.Filter = "X坐标=" & SinX & " and 项目序号=" & lng行号 & " and 类型=3"
                        
                        If rsNotes.BOF Then
                            gstrValues = Format(!时间, "yyyy-MM-dd HH:mm:ss") & "|" & lng行号 & "|3|" & _
                                str内容 & "|" & RGB_RED & "|" & SinX & "|0|0|0|0"
                            Call Record_Add(rsNotes, gstrFields, gstrValues)
                        Else
                            rsNotes!时间 = Format(!时间, "yyyy-MM-dd HH:mm:ss")
                            rsNotes!内容 = str内容
                            rsNotes.Update
                        End If
                    End If
                    rsNotes.Filter = ""
                End If
                .MoveNext
            Loop
        End With
        
        '提取婴儿出生信息
        If intBaby > 0 Then
            gstrFields = "时间|项目序号|类型|内容|颜色|X坐标|Y坐标|高度|打印X坐标|禁用"  '入出转缺省是红色,上下标及未记说明缺省是蓝色
            Set rsTmp = GetDataFromHis(lng病人ID, lng主页ID, intBaby, CDate(strTmpDay), CDate(strEndDay), 3)
            With rsTmp
                Do While Not .EOF
                    bytShow = 0
                    If Trim(zlCommFun.Nvl(!内容)) <> "" Then
                        lng行号 = 12
                        bytShow = T_BodyFlag.出生
                        If bytShow > 0 Then
                            Select Case bytShow
                                Case 1
                                    str内容 = zlCommFun.Nvl(!内容)
                                Case 2
                                    str内容 = zlCommFun.Nvl(!内容) & gstrCaveSplit & ConvertTimeToChinese(Format(!时间, "HH:mm"))
                            End Select
                            
                            SinX = Val(GetXCoordinate(Format(!时间, "YYYY-MM-DD HH:mm:ss"), strTmpDay))
                            rsNotes.Filter = "X坐标=" & SinX & " and 项目序号=" & lng行号 & " and 类型=13"
                            
                            If rsNotes.BOF Then
                                gstrValues = Format(!时间, "yyyy-MM-dd HH:mm:ss") & "|" & lng行号 & "|13|" & _
                                    str内容 & "|" & RGB_RED & "|" & SinX & "|0|0|0|0"
                                Call Record_Add(rsNotes, gstrFields, gstrValues)
                            Else
                                rsNotes!时间 = Format(!时间, "yyyy-MM-dd HH:mm:ss")
                                rsNotes!内容 = str内容
                                rsNotes.Update
                            End If
                        End If
                    End If
                    rsNotes.Filter = ""
                .MoveNext
                Loop
            End With
        End If
        '51512,刘鹏飞,2012-07-11,未记说明显示位置 0-显示在上面,1-显示在下面,2-不显示(新增)
        '大医二院要求未记说明不显示，但标注了未记的两边的体温曲线不连接
        strTmp = ""
        Dim arrString() As String
        '处理体温不升 体温不升始终显示在 35 度下面，只有未记说明显示在下面的情况，才将不升放入未记说明中，其它情况都放在下标中
        If Left(strTmpString0, 1) = ";" Then
            gstrFields = "时间|项目序号|类型|内容|颜色|X坐标|Y坐标|高度|打印X坐标|禁用"
            If mlng体温不升显示方式 = 0 Or mlng体温不升显示方式 = 2 Then
                arrString = Split(strTmpString0, "|")
                arrString(3) = "↓ "
                strTmpString0 = Join(arrString, "|")
            End If
            strTmpString0 = Mid(strTmpString0, "2")
            strTmpString2 = Mid(strTmpString2, 2)
            For i = 0 To UBound(Split(strTmpString0, ";"))
                strTmp = Split(strTmpString0, ";")(i)
                rsNotes.Filter = "类型=" & IIf(byt未记显示位置 = 1, 99, 6) & " and X坐标=" & Val(Split(strTmpString2, ";")(i))
                rsNotes.Sort = "项目序号"
                If rsNotes.RecordCount > 0 Then
                    rsNotes!内容 = IIf(mlng体温不升显示方式 = 0 Or mlng体温不升显示方式 = 2, "↓ ", "不升") & ";" & zlCommFun.Nvl(rsNotes!内容)
                    rsNotes.Update
                Else
                    If mlng体温不升显示方式 = 0 Or mlng体温不升显示方式 = 2 Then strTmp = Replace(strTmp, "不升", "↓ ")
                    Call Record_Add(rsNotes, gstrFields, strTmp)
                    rsNotes!类型 = IIf(byt未记显示位置 = 1, 99, 6)
                    rsNotes.Update
                End If
            Next i
        End If
        
        '如果未记说明不显示，将取消记录集rsNote中类型为99的记录
        If byt未记显示位置 = 2 Then
            rsNotes.Filter = "类型=99"
            Do While Not rsNotes.EOF
                rsNotes.Delete
                rsNotes.MoveNext
            Loop
            rsNotes.Filter = ""
        End If
        rsPoints.Filter = 0
        '6 计算组织重复的点
        Call GetConverPoint(rsPoints)
        stdset.Name = "宋体"
        stdset.Size = 9 * sngScale
        Call SetFontIndirect(stdset, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        '7 开始点的信息并连线
        
        strTmp = ShowPoints(lngDC, objDraw, rsPoints, strEditors, int心率应用, sngScale)
        
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
        '8.心率脉搏短轴连线
        rsPoints.Filter = ""
        If strTmp <> "" And bln打印脉搏短绌 = True Then Call CreatePoly(rsPoints, objDraw, lngDC, strTmpDay, strTmp)
        '9输出说明信息
        '  先处理未及说明和下标说明
        Dim strText As String
        Dim SinY35 As Single, SinY42 As Single
        Dim intAscCharNum As Integer
        
        strTime = ""
        strTmp = ""
        blnAllow = False
        SinX = 0: sinY = 0
        SinY35 = GetYCoordinate(objDraw, rsDrawItems, gint体温, 35, lngDC)
        SinY42 = GetYCoordinate(objDraw, rsDrawItems, gint体温, 42, lngDC)
        
        rsNotes.Filter = ""
        rsNotes.Sort = "X坐标,项目序号"
        With rsNotes
            Do While Not .EOF
                strTmp = ""
                For i = 0 To UBound(Split(!内容, ";"))
                    If Not (Split(!内容, ";")(i) = "不升" And byt未记显示位置 = 0 And Nvl(!类型) = 99) And Split(!内容, ";")(i) <> "" Then
                        If InStr(1, strTmp, Split(!内容, ";")(i)) = 0 Then
                            strTmp = strTmp & ";" & Split(!内容, ";")(i)
                        End If
                    End If
                Next i
                If Left(strTmp, "1") = ";" Then strTmp = Mid(strTmp, 2)
                If strTmp <> "" Then
                    strTime = Replace(strTmp, ";", " ")
                    If zlCommFun.Nvl(!类型) = 99 Then
                        If byt未记显示位置 = 1 Then '显示在体温的下面
                            If blnAllow = True Then
                                If Val(zlCommFun.Nvl(!X坐标)) <> SinX Then
                                    sinY = SinY35
                                Else
                                    strTime = " " & strTime
                                End If
                            Else
                                sinY = SinY35
                            End If
                            SinX = Val(zlCommFun.Nvl(!X坐标))
                            For i = 1 To Len(strTime)
                                If sinY < T_DrawClient.刻度区域.Bottom Then
                                    strText = Mid(strTime, i, 1)
                                    Call GetTextExtentPoint32(lngDC, strText, Len(strText), T_Size)
                                    If T_DrawClient.刻度区域.Bottom - sinY > T_Size.H Then
                                        Call DrawRotateText(objDraw, lngDC, SinX, sinY, strText, Val(!颜色))
                                    End If
                                    If Asc(strText) < 0 Then
                                        sinY = sinY + T_Size.H
                                    Else
                                        sinY = sinY + T_Size.H / 2
                                    End If
                                End If
                            Next i
                            rsNotes!禁用 = 1
                            blnAllow = True
                        Else
                            rsNotes!内容 = strTime
                            rsNotes!Y坐标 = SinY42
                            blnAllow = False
                        End If
                    ElseIf zlCommFun.Nvl(!类型) = 6 Then
                        If blnAllow = True Then
                            If Val(zlCommFun.Nvl(!X坐标)) <> SinX Then
                                sinY = SinY35
                            Else
                                strTime = " " & strTime
                            End If
                        Else
                            sinY = SinY35
                        End If
                        SinX = Val(zlCommFun.Nvl(!X坐标))
                        For i = 1 To Len(strTime)
                            If i < 3 Then intAscCharNum = 0
                            If sinY < T_DrawClient.刻度区域.Bottom Then
                                strText = Mid(strTime, i, 1)
                                Call GetTextExtentPoint32(lngDC, strText, Len(strText), T_Size)
                                
                                If Asc(strText) < 0 Then
                                    If intAscCharNum Mod 2 = 1 Then sinY = sinY + T_Size.H / 2
                                End If
                                '输出字体信息
                                If T_DrawClient.刻度区域.Bottom - sinY > T_Size.H Then
                                    Call DrawRotateText(objDraw, lngDC, SinX, sinY, strText, Val(zlCommFun.Nvl(!颜色)))
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
                        rsNotes!禁用 = 1
                        blnAllow = False
                        sinY = 0
                    Else
                        '入出转等标记信息 开始Y轴坐标均更新为42
                        rsNotes!Y坐标 = SinY42
                    End If
                End If
            .MoveNext
            Loop
        End With
        If rsNotes.RecordCount > 0 Then rsNotes.MoveFirst: rsNotes.Update
        stdset.Name = "宋体"
        stdset.Size = 9 * sngScale
        stdset.Bold = False
        Call SetFontIndirect(stdset, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        Call OutPutText(objDraw, rsDrawItems, lngDC, rsNotes, strTmpDay, sngScale)
        
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
        
        '开始处理表下表格（特殊项目栏）
        ReDim ArrNewString(0)
        Dim arrTmpString0() As String, arrTmpString1() As String, arrTmpString2() As String
        
        '组织提取表下表格信息
        For i = 0 To UBound(ArrComTable)
            lng项目序号 = Val(Split(ArrComTable(i), "||")(0))
            str项目名称 = Trim(Split(ArrComTable(i), "||")(1))
            If lng项目序号 <> 4 Then
                j = InStr(1, str项目名称, "(")
                If j > 0 Then
                    strItemName = Trim(Left(str项目名称, j - 1))
                Else
                    strItemName = Trim(str项目名称)
                End If
                If InStr(1, "," & strItems & ",", ",'" & strItemName & "',") = 0 Then
                    strItems = strItems & ",'" & strItemName & "'"
                End If
            End If
        Next i
        
        If Left(strItems, 1) = "," Then strItems = Mid(strItems, 2)
        If Not mbln呼吸曲线 Then strItems = strItems & ",'呼吸'"
        strItems = strItems & ",'收缩压','舒张压'"
        If Left(strItems, 1) = "," Then strItems = Mid(strItems, 2)
        
        dtBegin = Int(CDate(strTmpDay) - 1)
        dtEnd = CDate(CDate(Format(strEndDay, "YYYY-MM-DD HH:mm:ss")) + 1)
        If CDate(Format(dtBegin, "YYYY-MM-DD HH:mm:ss")) < CDate(Format(strBeginDate, "YYYY-MM-DD HH:mm:ss")) Then _
            dtBegin = CDate(Format(strBeginDate, "YYYY-MM-DD HH:mm:ss"))
        If CDate(Format(dtEnd, "YYYY-MM-DD HH:mm:ss")) > CDate(Format(strEndDate, "YYYY-MM-DD HH:mm:ss")) Then _
            dtEnd = CDate(Format(strEndDate, "YYYY-MM-DD HH:mm:ss"))

        
        '提取所有表格项目数据信息
        strSql = "SELECT C.Id,a.发生时间 As 时间,C.记录类型,C.显示,C.记录内容 As 结果,C.体温部位,C.未记说明,nvl(C.数据来源,0) 数据来源," & vbNewLine & _
            "  DECODE(E.项目性质,2,C.体温部位 || D.记录名,D.记录名) 项目名称,D.项目序号,C.来源ID,C.共用,E.项目性质 " & _
            "  FROM 病人护理文件 B, 病人护理数据 A,病人护理明细 C,体温记录项目 D,护理记录项目 E " & _
            "  Where B.ID = A.文件ID" & vbNewLine & _
            "  AND A.ID = C.记录ID" & vbNewLine & _
            "  AND B.ID = [1]" & vbNewLine & _
            "  AND Nvl(B.婴儿, 0) = [7]" & vbNewLine & _
            "  AND B.病人id = [2]" & vbNewLine & _
            "  AND B.主页id = [3]" & vbNewLine & _
            "  AND INSTR([6], DECODE(E.项目性质, 2,C.体温部位 || D.记录名, D.记录名)) > 0" & vbNewLine & _
            "  AND D.项目序号 = C.项目序号" & vbNewLine & _
            "  AND Mod(c.记录类型,10) = 1" & vbNewLine & _
            "  AND E.项目序号 = D.项目序号" & vbNewLine & _
            "  AND E.护理等级 >= [8]" & vbNewLine & _
            "  AND A.发生时间 BETWEEN [4] And [5]" & vbNewLine & _
            "  And C.终止版本 Is Null" & vbNewLine & _
            "  AND D.记录法 = 2" & vbNewLine & _
            "  UNION ALL "
         '提取非体温表格的汇总项目（体温表格汇总项目子项可能存在非体温项目）
        strSql = strSql & vbNewLine & _
            "  SELECT C.ID,a.发生时间 As 时间,C.记录类型,C.显示,C.记录内容 As 结果,C.体温部位,C.未记说明,nvl(C.数据来源,0) 数据来源," & _
            "   D.项目名称,D.项目序号,C.来源ID,C.共用,D.项目性质" & _
            "   FROM 病人护理文件 B, 病人护理数据 A,病人护理明细 C,(SELECT A.项目序号,A.项目名称, 1 项目性质,B.父序号 FROM 护理记录项目 A,护理汇总项目 B" & vbNewLine & _
            "       WHERE A.项目序号=B.序号 AND NOT EXISTS (SELECT C.项目序号 FROM 体温记录项目 C,护理汇总项目 E WHERE C.项目序号=E.序号 AND C.项目序号=A.项目序号)" & vbNewLine & _
            "       AND NVL(A.应用方式,0)=1 AND NVL(A.护理等级,0)>=[8] AND NVL(A.适用病人,0) IN (0,[9])" & vbNewLine & _
            "       AND (A.适用科室=1 OR (A.适用科室=2 AND EXISTS (SELECT 1 FROM 护理适用科室 D WHERE D.项目序号=A.项目序号 AND D.科室ID=[10])))) D" & _
            "   Where B.ID=A.文件ID And A.ID = C.记录ID   AND B.ID=[1]  AND Nvl(B.婴儿,0)=[7] " & _
            "   AND B.病人id=[2]  AND B.主页id=[3]  AND D.项目序号=C.项目序号  AND C.记录类型=1" & _
            "   AND A.发生时间 BETWEEN [4] And [5] And C.终止版本 Is Null"
            
        strSql = _
            "   Select ID,时间,记录类型,显示,结果,体温部位,未记说明,数据来源,项目名称,项目序号,来源ID,共用,项目性质 From (" & strSql & ")" & _
            "   Order By  Decode(项目名称,'收缩压',0,1)," & strItems & ",时间"
            
        If blnMoved Then
            strSql = Replace(strSql, "病人护理文件", "H病人护理文件")
            strSql = Replace(strSql, "病人护理数据", "H病人护理数据")
            strSql = Replace(strSql, "病人护理明细", "H病人护理明细")
        End If
        Set rsTmp = zldatabase.OpenSQLRecord(strSql, "Print", _
                                            lng文件ID, _
                                            lng病人ID, _
                                            lng主页ID, _
                                            CDate(dtBegin), _
                                            CDate(dtEnd), _
                                            strItems, intBaby, lng护理等级, IIf(intBaby = 0, 1, 2), lngSectID)
                                                    
        ReDim Preserve ArrNewString(UBound(ArrComTable))
        For i = 0 To UBound(ArrComTable)
            If Split(ArrComTable(i), "||")(0) = 3 Then '呼吸项目
                lng项目序号 = Val(Split(ArrComTable(i), "||")(0))
                strNewTmpString = String(42, ";")
                arrTmpString0 = Split(String(42, ";"), ";")
                arrTmpString1 = Split(String(42, ";"), ";")
                arrTmpString2 = Split(String(42, ";"), ";")
                
                ArrNewTmpString = Split(strNewTmpString, ";")
                
                rsTmp.Filter = "项目序号=" & gint呼吸
                With rsTmp
                    Do While Not .EOF
                        blnAdd = False
                        If CDate(Format(!时间, "YYYY-MM-DD HH:mm:ss")) >= CDate(Format(strTmpDay, "YYYY-MM-DD HH:mm:ss")) Then
                            intCOl = GetCurveColumn(CDate(!时间), CDate(strTmpDay), gintHourBegin)
                            If intCOl > LBound(ArrNewTmpString) And intCOl <= UBound(ArrNewTmpString) Then
                            
                                If arrTmpString1(intCOl) <> "" Then
                                    If (Val(arrTmpString2(intCOl)) = 0 And Val(zlCommFun.Nvl(!显示, 0)) = 0) Or _
                                        (Val(arrTmpString2(intCOl)) = 1 And Val(zlCommFun.Nvl(!显示, 0)) = 1) Then
                                        
                                        '检查那个离重点时间更近
                                        SinX = GetXCoordinate(Format(!时间, "YYYY-MM-DD HH:mm:ss"), Format(strTmpDay, "YYYY-MM-DD HH:mm:ss"))
                                        blnAdd = GetCanvasCenter(CDate(Format(arrTmpString1(intCOl), "YYYY-MM-DD HH:mm:ss")), CDate(Format(!时间, "YYYY-MM-DD HH:mm:ss")), CDate(Format(strTmpDay, "YYYY-MM-DD HH:mm:ss")), SinX)
                                    ElseIf Val(arrTmpString2(intCOl)) = 1 Then
                                        blnAdd = False
                                    Else
                                        blnAdd = True
                                    End If
                                    If blnAdd = True Then
                                        If Val(arrTmpString2(intCOl)) = 2 Then
                                            arrTmpString0(intCOl) = zlCommFun.Nvl(!结果) & "," & zlCommFun.Nvl(!体温部位)
                                            arrTmpString1(intCOl) = Format(!时间, "YYYY-MM-DD HH:mm:ss")
                                            arrTmpString2(intCOl) = 2
                                            GoTo ErrNext
                                        End If
                                    Else
                                        If Val(zlCommFun.Nvl(!显示, 0)) = 2 Then
                                            arrTmpString2(intCOl) = 2
                                            GoTo ErrNext
                                        End If
                                    End If
                                Else
                                    blnAdd = True
                                End If
                                
                                If blnAdd = True Then
                                    arrTmpString0(intCOl) = zlCommFun.Nvl(!结果) & "," & zlCommFun.Nvl(!体温部位)
                                    arrTmpString1(intCOl) = Format(!时间, "YYYY-MM-DD HH:mm:ss")
                                    arrTmpString2(intCOl) = Val(zlCommFun.Nvl(!显示, 0))
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
                int频次 = Val(Split(ArrComTable(i), "||")(4))
                strTmp = Val(Split(ArrComTable(i), "||")(6)) '项目表示 4表示汇总项目
                lng项目序号 = Val(Split(ArrComTable(i), "||")(0))
                str项目名称 = Trim(Split(ArrComTable(i), "||")(1))
                int项目性质 = Val(Split(ArrComTable(i), "||")(5))
                int项目类型 = Val(Split(ArrComTable(i), "||")(7))
                int入院首测 = Val(Split(ArrComTable(i), "||")(8))
                
                If Val(strTmp) = 4 Or IsWaveItem(lng项目序号) Then
                    If int频次 > 2 Then int频次 = 2 '汇总/波动项目频次只能是 1 、 2
                End If
                
                blnColor = (int项目性质 = 2 And int项目类型 = 1 And Val(strTmp) = 0)
                strNewTmpString = String(Val(int频次) * 7, ";")
              
                ArrNewTmpString = Split(strNewTmpString, ";")
                
                For j = 0 To 6
                    strBegin = DateAdd("D", j, CDate(strTmpDay))
                    If CDate(strBegin) > CDate(strEndDay) Then strBegin = strEndDay
                    int舒张压 = 0
                    int收缩压 = 0
                    Int列号 = 0
                    strTime = ""
                    intCOl = 0
                    
                    Set rsDownTab = ReturnItemRecord(rsTmp, Int(CDate(strBegin)), CDate(strBeginDate), lng项目序号 & ";" & str项目名称 & ";" & _
                                    int频次 & ";" & Val(strTmp) & ";" & int项目性质 & ";" & int入院首测, bln汇总当天, bln录入小时)
                    If rsDownTab.RecordCount > 0 Then rsDownTab.MoveFirst
                    rsDownTab.Sort = "时间,项目序号,序号"
                    With rsDownTab
                        Do While Not .EOF
                            lngColor = 0
                            str结果 = zlCommFun.Nvl(!记录内容)
                            intCOl = Val(!序号)
                            intCOl = intCOl + j * int频次
                            If blnColor Then lngColor = Val(zlCommFun.Nvl(!未记说明))
                            
                            Select Case zlCommFun.Nvl(!项目名称)
                                Case "舒张压"
                                    If int舒张压 <> intCOl Then
                                        If Trim(ArrNewTmpString(intCOl)) <> "" Or str结果 <> "" Then
                                            If InStr(1, ArrNewTmpString(intCOl), "/") > 0 Then
                                                ArrNewTmpString(intCOl) = Trim(Split(ArrNewTmpString(intCOl), "/")(0)) & "/" & str结果
                                            Else
                                                ArrNewTmpString(intCOl) = "/" & str结果
                                            End If
                                            If str结果 = "外出" Or str结果 = "拒测" Or str结果 = "请假" Or str结果 = "未测" Then ArrNewTmpString(intCOl) = str结果
                                        End If
                                         int舒张压 = intCOl
                                         If ArrNewTmpString(intCOl) = "/" Then ArrNewTmpString(intCOl) = ""
                                    End If
                                Case "收缩压"
                                    If int收缩压 <> intCOl Then
                                        If ArrNewTmpString(intCOl) <> "" Or str结果 <> "" Then
                                            If InStr(1, ArrNewTmpString(intCOl), "/") > 0 Then
                                                ArrNewTmpString(intCOl) = str结果 & "/" & Trim(Split(ArrNewTmpString(intCOl), "/")(1))
                                            Else
                                                ArrNewTmpString(intCOl) = str结果 & "/"
                                            End If
                                        End If
                                        int收缩压 = intCOl
                                    End If
                                Case Else
                                    If Int列号 <> intCOl Then
                                        ArrNewTmpString(intCOl) = Replace(str结果, "-#", "") & "-#" & lngColor
                                        Int列号 = intCOl
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
        
        '项目序号||部位+项目名称||项目单位||项目值域||记录频次||项目性质||项目表示
        For i = 0 To UBound(ArrComTable)
            strTmpString0 = ""

            If Trim(CStr(Split(ArrComTable(i), "||")(2))) <> "" Then
                strTmpString0 = Trim(CStr(Split(ArrComTable(i), "||")(1))) & "(" & Trim(CStr(Split(ArrComTable(i), "||")(2))) & ")"
            Else
                strTmpString0 = Trim(CStr(Split(ArrComTable(i), "||")(1)))
            End If
           
            ArrNewString(i) = Trim(CStr(Split(ArrComTable(i), "||")(0))) & ";" & strTmpString0 & ";" & ArrNewString(i)
        Next i
        
        '显示皮试结果
        If bln显示皮试 = True Then
            strSql = _
               "SELECT 时间,F_LIST2STR(CAST(COLLECT(药物名) AS T_STRLIST)) 药物名 FROM (" & vbNewLine & _
                "   SELECT TO_CHAR(开始执行时间,'YYYY-MM-DD') 时间,DECODE(皮试结果,'(+)',255,0) || '-#' || REPLACE(REPLACE(医嘱内容,',',''),'-#','') || 皮试结果  药物名" & vbNewLine & _
                "   FROM 病人医嘱记录" & vbNewLine & _
                "   WHERE  病人ID=[1] AND 主页ID=[2] AND 婴儿=[3] AND 皮试结果 IS NOT NULL" & vbNewLine & _
                "   AND 开始执行时间  BETWEEN [4] AND [5]" & vbNewLine & _
                "   ORDER BY TO_DATE(TO_CHAR(开始执行时间,'YYYY-MM-DD'),'YYYY-MM-DD'),皮试结果" & vbNewLine & _
                ") GROUP BY 时间"
                
            If blnMoved Then
                strSql = Replace(strSql, "病人过敏记录", "H病人过敏记录")
            End If
            
            Set rsTmp = zldatabase.OpenSQLRecord(strSql, "提取病人过敏记录信息", lng病人ID, lng主页ID, intBaby, CDate(strTmpDay), CDate(strEndDay))
            
            strNewTmpString = String(7, ";")
            ArrNewTmpString = Split(strNewTmpString, ";")
            intCOl = 0
            
            Do While Not rsTmp.EOF
                intCOl = DateDiff("D", CDate(Format(strTmpDay, "YYYY-MM-DD")), CDate(Format(rsTmp!时间, "YYYY-MM-DD"))) + 1
                ArrNewTmpString(intCOl) = Nvl(rsTmp!药物名)
                rsTmp.MoveNext
            Loop
            strNewTmpString = Join(ArrNewTmpString, "||")
            ReDim Preserve ArrNewString(UBound(ArrNewString) + 1)
            ArrNewString(UBound(ArrNewString)) = "-999;皮试结果" & ";" & strNewTmpString
        End If
        
        lngCurX = X
'        stdset.Name = "宋体"
'        stdset.Size = 9 * sngScale
'        stdset.Bold = False
'        Call SetFontIndirect(stdset, lngDC, objDraw)
'        lngFont = CreateFontIndirect(T_Font)
'        lngOldFont = SelectObject(lngDC, lngFont)
        '开始绘画表格项目并展示数据
        Call DrawBodyRecordItem(lngDC, objDraw, ArrNewString, rsItems, lngCurX, T_DrawClient.体温区域.Bottom, T_DrawClient.体温区域.Right, intRepairRows, lngCurY, sngScale)
'       Call SelectObject(lngDC, lngOldFont)
'       Call DeleteObject(lngFont)
        lngCurX = X
        lngCurY = lngCurY
        
        stdset.Name = "宋体"
        stdset.Size = 9 * sngScale
        stdset.Bold = False
        Call SetFontIndirect(stdset, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        '开始打印 页数 住院周数 和 体温说明信息
        Call DrawBodyPageFooter(lngDC, objDraw, lngCurX, lngCurY, T_DrawClient.体温区域.Right, intPageNo, intEndPage, str体温说明, sngScale)
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
        '页脚图形输出
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
    T_DrawClient.偏移量X = M_DrawClient.偏移量X
    T_DrawClient.偏移量Y = M_DrawClient.偏移量Y
    T_DrawClient.刻度区域 = M_DrawClient.刻度区域
    T_DrawClient.刻度单位 = M_DrawClient.刻度单位
    T_DrawClient.体温区域 = M_DrawClient.体温区域
    T_DrawClient.行单位 = M_DrawClient.行单位
    T_DrawClient.时间行单位 = M_DrawClient.时间行单位
    T_DrawClient.时间列单位 = M_DrawClient.时间列单位
    T_DrawClient.列单位 = M_DrawClient.列单位
    T_DrawClient.双倍 = M_DrawClient.双倍
    T_DrawClient.总列数 = M_DrawClient.总列数
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
    '如果是打印预览,应按打印机的可打印的开始处开始预览
    dblSureW = GetDeviceCaps(Printer.hDC, PHYSICALOFFSETX) / GetDeviceCaps(Printer.hDC, PHYSICALWIDTH)
    dblSureH = GetDeviceCaps(Printer.hDC, PHYSICALOFFSETY) / GetDeviceCaps(Printer.hDC, PHYSICALHEIGHT)
    On Error Resume Next
    Call DrawRect(lngDC, (objDraw.Width * dblSureW) / T_TwipsPerPixel.X, (objDraw.Height * (1 - dblSureH)) / T_TwipsPerPixel.Y, _
    (objDraw.Width * (1 - dblSureW)) / T_TwipsPerPixel.X, objDraw.Height * dblSureH / T_TwipsPerPixel.Y, PS_DOT, 1, RGB_FleetGRAY)
End Sub
Public Sub DrawBodyRecordItem(ByVal lngDC As Long, ByVal objDraw As Object, strValue() As String, ByVal rsItems As ADODB.Recordset, ByVal lngX As Long, ByVal lngY As Long, _
    ByVal lngLeft As Long, ByVal intRepairRows As Integer, lngOutY As Long, Optional sngScale As Single = 1)
'-----------------------------------------------------------------------------------------------------------------------
'输出病人基本信息
'参数:lngDC 绘图对象的DC，strValue() 所有表格项目的信息 (格式（呼吸）:项目序号;名称;内容,部位||内容,部位/(其他) 项目序号;名称;内容||内容) 内容和部位组成的数组表示此项目有多少列
'    rsItems 所有体温表格护理项目, lngX 左边距,lngY上边距,lngLeft 右边距(可以绘图的最大右边距),intRepairRows 要打印表格项目的总行数
'出参:lngOutY 返回绘图后的上边距
'-----------------------------------------------------------------------------------------------------------------------
    Dim lngX1 As Long, lngY1 As Long, lngCurY As Long, lngCurX As Long
    Dim lngRowHeiht As Long
    Dim arrTmpString0() As String, arrTmpString1() As String
    Dim arrTmp() As String, arrText() As String
    Dim intRow As Integer, intCOl As Integer
    Dim i As Integer
    Dim int呼吸表格输出格式 As Integer
    Dim bln灌肠大便以分子分母显示 As Boolean
    Dim strTmp As String, strPart As String
    Dim strPic As String
    Dim blnValue As Boolean
    Dim intValue As Integer, int呼吸位置 As Integer
    Dim intRowCount As Integer
    Dim int频次 As Integer '记录频次
    Dim blnDataTrue As Boolean
    Dim lngColor As Long
    Dim intNum As Integer
    Dim blnOutText As Boolean '是否输出文本
    Dim blnPrinter As Boolean
    Dim intBold As Integer, intFine As Integer
    Dim intSize As Integer
    Dim sngLen As Single, lngLen As Long
    Dim LPoint As T_LPoint
    Dim bln显示皮试 As Boolean
    
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
    int呼吸位置 = 0
    int呼吸表格输出格式 = zldatabase.GetPara("呼吸表格输出", glngSys, 1255, 0)
    bln灌肠大便以分子分母显示 = (Val(zldatabase.GetPara("灌肠后大便显示格式", glngSys, 1255, 0)) = 1)
    
    strPic = ""
    If InStr(1, strValue(0), ";") > 0 Then
        bln显示皮试 = IIf(Split(strValue(UBound(strValue)), ";")(0) = "-999", True, False)
        
        For intRow = LBound(strValue) To UBound(strValue)
            arrTmpString0 = Split(strValue(intRow), ";")
            arrTmpString1 = Split(arrTmpString0(2), "||")
            
            If intRepairRows > 0 And intRepairRows > intRowCount Then
            
                If arrTmpString0(0) = "3" Then '呼吸项目
                    '提取表格颜色
                    rsItems.Filter = 0
                    rsItems.Filter = "项目序号=" & gint呼吸
                    If rsItems.RecordCount > 0 Then
                        lngColor = Val(Nvl(rsItems!记录色, RGB_RED))
                    Else
                        lngColor = RGB_RED
                    End If
                    intRowCount = intRowCount + 1
                    arrTmpString1 = Split(arrTmpString0(2), "||")
                    For intCOl = 0 To UBound(arrTmpString1)
                        If intCOl = 0 Then '表头
                            Call SetTextColor(lngDC, RGB_BLACK)
                            Call GetTextExtentPoint32(lngDC, arrTmpString0(intCOl + 1), Len(arrTmpString0(intCOl + 1)), T_Size)
                            Call GetTextRect(objDraw, lngX, lngY + (T_DrawClient.时间列单位 + T_DrawClient.时间列单位 / 2) / 2, arrTmpString0(intCOl + 1), _
                                T_DrawClient.刻度区域.Right - lngX, True, , sngScale)
                            'Call DrawText(lngDC, arrTmpString0(intCOl + 1), -1, T_LableRect, DT_CENTER)
                            LPoint.X = lngX
                            LPoint.Y = lngY + (T_DrawClient.时间列单位 + T_DrawClient.时间列单位 / 2) / 2
                            LPoint.W = T_DrawClient.刻度区域.Right - lngX
                            Call DrawTabText(lngDC, objDraw, arrTmpString0(intCOl + 1), -1, T_LableRect, DT_CENTER, LPoint, sngScale)
                            Call DrawLine(lngDC, lngX, lngY, lngX, lngY + T_DrawClient.时间列单位 + T_DrawClient.时间列单位 / 2, PS_SOLID, intBold, RGB_BLACK)
                            Call DrawLine(lngDC, lngX, lngY + T_DrawClient.时间列单位 + T_DrawClient.时间列单位 / 2, T_DrawClient.刻度区域.Right, _
                                lngY + T_DrawClient.时间列单位 + T_DrawClient.时间列单位 / 2, PS_SOLID, IIf(intRowCount = intRepairRows, intBold, intFine), RGB_BLACK)
                            Call DrawLine(lngDC, T_DrawClient.刻度区域.Right, lngY, T_DrawClient.刻度区域.Right, _
                                lngY + T_DrawClient.时间列单位 + T_DrawClient.时间列单位 / 2, PS_SOLID, intBold, RGB_BLACK)
                            lngX1 = T_DrawClient.刻度区域.Right
                            lngY1 = lngCurY
                        Else
                            arrTmpString1(intCOl) = arrTmpString1(intCOl) & String(1 - UBound(Split(arrTmpString1(intCOl), ",")), ",")
                            strTmp = Split(arrTmpString1(intCOl), ",")(0)
                            strPart = Split(arrTmpString1(intCOl), ",")(1)
                            If strPart = "" Then strPart = "自主呼吸"
                            strPic = ""
                            '打印呼吸值（间隔错开打印） 第一行始终在上面
                            If IsNumeric(strTmp) Then
                                If strPart = "自主呼吸" Then
                                    Call SetTextColor(lngDC, lngColor)
                                    Call GetTextExtentPoint32(lngDC, strTmp, Len(strTmp), T_Size)
                                Else
                                    strPic = "BREATH"
                                End If
                                
                                If blnValue = False Then
                                    intValue = IIf(intCOl Mod 2 = 0, 0, 1)
                                    blnValue = True
                                    int呼吸位置 = 2
                                End If
                                
                                If int呼吸表格输出格式 = 0 Then '顺序上下显示
                                    If intCOl Mod 2 = intValue Then
                                        If strPic = "" Then
                                            LPoint.X = lngX1
                                            LPoint.Y = lngY
                                            Call GetTextRect(objDraw, lngX1, lngY, strTmp, T_DrawClient.列单位, False, , sngScale)
                                        Else
                                            Call DrawPicture(objDraw, strPic, objDraw.ScaleX(lngX1 + ((T_DrawClient.列单位 - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2), vbPixels, vbTwips), _
                                                objDraw.ScaleY(lngY + 1, vbPixels, vbTwips), _
                                                objDraw.ScaleX(lngX1 + ((T_DrawClient.列单位 - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2) + mintBmpW * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), _
                                                objDraw.ScaleY(lngY + 1 + mintBmpH * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), True)
                                            
                                        End If
                                    Else
                                        If strPic = "" Then
                                            LPoint.X = lngX1
                                            LPoint.Y = lngY + ((T_DrawClient.时间列单位 + T_DrawClient.时间列单位 / 2) - T_Size.H)
                                            Call GetTextRect(objDraw, lngX1, lngY + ((T_DrawClient.时间列单位 + T_DrawClient.时间列单位 / 2) - T_Size.H), _
                                                strTmp, T_DrawClient.列单位, False, , sngScale)
                                        Else
                                            Call DrawPicture(objDraw, strPic, objDraw.ScaleX(lngX1 + ((T_DrawClient.列单位 - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2), _
                                                vbPixels, vbTwips), objDraw.ScaleY(lngY + ((T_DrawClient.时间列单位 + T_DrawClient.时间列单位 / 2) - mintBmpH * IIf(blnPrinter = True, msngTwips, 1)), vbPixels, vbTwips), _
                                                objDraw.ScaleX(lngX1 + ((T_DrawClient.列单位 - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2) + mintBmpW * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), _
                                                objDraw.ScaleY(lngY + (T_DrawClient.时间列单位 + T_DrawClient.时间列单位 / 2), vbPixels, vbTwips), True)
                                        End If
                                    End If
                                    
                                Else        '有数据时数据之间上下显示
                                    If int呼吸位置 = 2 Then
                                        If strPic = "" Then
                                            LPoint.X = lngX1
                                            LPoint.Y = lngY
                                            Call GetTextRect(objDraw, lngX1, lngY, strTmp, T_DrawClient.列单位, False, , sngScale)
                                        Else
                                            Call DrawPicture(objDraw, strPic, objDraw.ScaleX(lngX1 + ((T_DrawClient.列单位 - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2), vbPixels, vbTwips), _
                                                objDraw.ScaleY(lngY + 1, vbPixels, vbTwips), _
                                                objDraw.ScaleX(lngX1 + ((T_DrawClient.列单位 - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2) + mintBmpW * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), _
                                                objDraw.ScaleY(lngY + 1 + mintBmpH * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), True)
                                        End If
                                    Else
                                        If strPic = "" Then
                                            LPoint.X = lngX1
                                            LPoint.Y = lngY + ((T_DrawClient.时间列单位 + T_DrawClient.时间列单位 / 2) - T_Size.H)
                                            Call GetTextRect(objDraw, lngX1, lngY + ((T_DrawClient.时间列单位 + T_DrawClient.时间列单位 / 2) - T_Size.H), _
                                                strTmp, T_DrawClient.列单位, False, , sngScale)
                                        Else
                                            Call DrawPicture(objDraw, strPic, objDraw.ScaleX(lngX1 + ((T_DrawClient.列单位 - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2), vbPixels, vbTwips), _
                                                objDraw.ScaleY(lngY + ((T_DrawClient.时间列单位 + T_DrawClient.时间列单位 / 2) - mintBmpH * IIf(blnPrinter = True, msngTwips, 1)), vbPixels, vbTwips), _
                                                objDraw.ScaleX(lngX1 + ((T_DrawClient.列单位 - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2) + mintBmpW * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), _
                                                objDraw.ScaleY(lngY + (T_DrawClient.时间列单位 + T_DrawClient.时间列单位 / 2), vbPixels, vbTwips), True)
                                        End If
                                    End If
                                    
                                   
                                    int呼吸位置 = int呼吸位置 + 1
                                    If int呼吸位置 > 2 Then int呼吸位置 = 1
                                End If
                                LPoint.W = T_DrawClient.列单位
                                If strPic = "" Then Call DrawTabText(lngDC, objDraw, strTmp, -1, T_LableRect, DT_CENTER, LPoint, sngScale) 'DrawText(lngDC, strTmp, -1, T_LableRect, DT_CENTER)
                                
                            End If
                            lngX1 = lngX1 + T_DrawClient.列单位
                        End If
                    Next intCOl
                    lngX1 = T_DrawClient.刻度区域.Right + T_DrawClient.列单位
                    lngY1 = lngY + T_DrawClient.时间列单位 + T_DrawClient.时间列单位 / 2
                    
                    '画呼吸栏所有的列
                    For intCOl = 1 To 42
                        If intCOl Mod 6 = 0 Then
                            Call DrawLine(lngDC, lngX1, lngY, lngX1, lngY1, PS_SOLID, intBold, RGB_BLACK)
                        Else
                            Call DrawLine(lngDC, lngX1, lngY, lngX1, lngY1, PS_SOLID, intFine, RGB_BLACK)
                        End If
                        lngX1 = lngX1 + T_DrawClient.列单位
                    Next intCOl
                    Call DrawLine(lngDC, T_DrawClient.刻度区域.Right, lngY1, T_DrawClient.体温区域.Right, lngY1, PS_SOLID, IIf(intRowCount = intRepairRows, intBold, intFine), RGB_BLACK)
                    
                    '当前Y轴坐标
                    lngCurY = lngY1
                ElseIf arrTmpString0(0) <> "-999" Then '不是皮试结果
                    
                    rsItems.Filter = ""
                    rsItems.Filter = "序号=" & intRow
                    If rsItems.RecordCount > 0 Then
                        int频次 = CInt(zlCommFun.Nvl(rsItems!记录频次, 2))
                        If Val(Nvl(rsItems!项目表示)) = 4 Or IsWaveItem(Val(Nvl(rsItems!项目序号))) Then
                            If int频次 > 2 Then int频次 = 2 '汇总/波动项目频次只能是 1 、 2
                        End If
                        '活动项目检查是否存在数据，不存在就不打印此行
                        If zlCommFun.Nvl(rsItems!项目性质) = 2 Then
                            
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
                        
                        '根据频次计算要打印的表格行数是否超出用户设置的表格行数
                        
                        intNum = 0
                        Select Case int频次
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
                            If intCOl = 0 Then '开始画表头信息包括标题的输出
                                Select Case int频次
                                    Case 1, 2, 6
                                        lngY1 = lngY1 + T_DrawClient.时间列单位
                                        lngRowHeiht = T_DrawClient.时间列单位 / 2
                                    Case 3
                                        lngY1 = lngY1 + T_DrawClient.时间列单位 * (3 - intNum)
                                        lngRowHeiht = (T_DrawClient.时间列单位 * (3 - intNum)) / 2
                                    Case 4
                                        lngY1 = lngY1 + T_DrawClient.时间列单位 * (2 - intNum)
                                        lngRowHeiht = (T_DrawClient.时间列单位 * (2 - intNum)) / 2
                                End Select

                                Call SetTextColor(lngDC, RGB_BLACK)
                                Call GetTextExtentPoint32(lngDC, arrTmpString0(intCOl + 1), Len(arrTmpString0(intCOl + 1)), T_Size)
                                Call GetTextRect(objDraw, lngX1, lngY1 - lngRowHeiht, arrTmpString0(intCOl + 1), T_DrawClient.刻度区域.Right - lngX1, True, , sngScale)
                                'Call DrawText(lngDC, arrTmpString0(intCOl + 1), -1, T_LableRect, DT_CENTER)
                                LPoint.X = lngX1
                                LPoint.Y = lngY1 - lngRowHeiht
                                LPoint.W = T_DrawClient.刻度区域.Right - lngX1
                                Call DrawTabText(lngDC, objDraw, arrTmpString0(intCOl + 1), -1, T_LableRect, DT_CENTER, LPoint, sngScale)
                                Call DrawLine(lngDC, lngX1, lngCurY, lngX1, lngY1, PS_SOLID, intBold, RGB_BLACK)
                                Call DrawLine(lngDC, lngX1, lngY1, T_DrawClient.刻度区域.Right, lngY1, PS_SOLID, IIf(intRowCount = intRepairRows, intBold, intFine), RGB_BLACK)
                                Call DrawLine(lngDC, T_DrawClient.刻度区域.Right, lngCurY, T_DrawClient.刻度区域.Right, lngY1, PS_SOLID, intBold, RGB_BLACK)
                                
                                lngY1 = lngCurY
                                lngX1 = T_DrawClient.刻度区域.Right
                            Else  '开始进行画表格线
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
                                
                                If strTmp = "*" And Val(arrTmpString0(0)) = gint大便 Then strTmp = "※"
                                
                                Call SetTextColor(lngDC, lngColor)
                                
                                Call GetTextExtentPoint32(lngDC, strTmp, Len(strTmp), T_Size)
                                blnOutText = True
                                
                                If InStr(1, ",3,4,", "," & int频次 & ",") = 0 Then
                                    LPoint.X = lngX1
                                    LPoint.Y = lngCurY + T_DrawClient.时间列单位 / 2
                                    LPoint.W = T_DrawClient.列单位 * (6 / int频次)
                                    Call GetTextRect(objDraw, lngX1, lngCurY + T_DrawClient.时间列单位 / 2, strTmp, T_DrawClient.列单位 * (6 / int频次), True, , sngScale)
                                    lngX1 = lngX1 + T_DrawClient.列单位 * (6 / int频次)
                                ElseIf int频次 = 3 Then
                                    LPoint.W = T_DrawClient.列单位 * 6
                                    If intCOl Mod int频次 = 0 Then
                                        LPoint.X = lngX1
                                        LPoint.Y = lngCurY + T_DrawClient.时间列单位 * 2 + T_DrawClient.时间列单位 / 2
                                        Call GetTextRect(objDraw, lngX1, lngCurY + T_DrawClient.时间列单位 * 2 + T_DrawClient.时间列单位 / 2, strTmp, T_DrawClient.列单位 * 6, True, , sngScale)
                                        If intNum <> 0 Then blnOutText = False
                                        lngX1 = lngX1 + T_DrawClient.列单位 * 6
                                    ElseIf intCOl Mod int频次 = 2 Then
                                        LPoint.X = lngX1
                                        LPoint.Y = lngCurY + T_DrawClient.时间列单位 + T_DrawClient.时间列单位 / 2
                                        Call GetTextRect(objDraw, lngX1, lngCurY + T_DrawClient.时间列单位 + T_DrawClient.时间列单位 / 2, strTmp, T_DrawClient.列单位 * 6, True, , sngScale)
                                        If intNum > 1 Then blnOutText = False
                                    Else
                                        LPoint.X = lngX1
                                        LPoint.Y = lngCurY + T_DrawClient.时间列单位 / 2
                                        Call GetTextRect(objDraw, lngX1, lngCurY + T_DrawClient.时间列单位 / 2, strTmp, T_DrawClient.列单位 * 6, True, , sngScale)
                                    End If
                                    
                                ElseIf int频次 = 4 Then
                                    LPoint.W = T_DrawClient.列单位 * 3
                                    If intCOl Mod 4 = 3 Then
                                        LPoint.X = lngX1
                                        LPoint.Y = lngCurY + T_DrawClient.时间列单位 + T_DrawClient.时间列单位 / 2
                                        Call GetTextRect(objDraw, lngX1, lngCurY + T_DrawClient.时间列单位 + T_DrawClient.时间列单位 / 2, strTmp, T_DrawClient.列单位 * 3, True, , sngScale)
                                        If intNum > 0 Then blnOutText = False
                                        lngX1 = lngX1 + T_DrawClient.列单位 * 3
                                    ElseIf intCOl Mod 4 = 0 Then
                                        LPoint.X = lngX1
                                        LPoint.Y = lngCurY + T_DrawClient.时间列单位 + T_DrawClient.时间列单位 / 2
                                        Call GetTextRect(objDraw, lngX1, lngCurY + T_DrawClient.时间列单位 + T_DrawClient.时间列单位 / 2, strTmp, T_DrawClient.列单位 * 3, True, , sngScale)
                                        If intNum > 0 Then blnOutText = False
                                        lngX1 = lngX1 + T_DrawClient.列单位 * 3
                                    ElseIf intCOl Mod 2 = 0 Then
                                        LPoint.X = lngX1
                                        LPoint.Y = lngCurY + T_DrawClient.时间列单位 / 2
                                        Call GetTextRect(objDraw, lngX1, lngCurY + T_DrawClient.时间列单位 / 2, strTmp, T_DrawClient.列单位 * 3, True, , sngScale)
                                        lngX1 = lngX1 - T_DrawClient.列单位 * 3
                                    ElseIf intCOl Mod 4 = 1 Then
                                        LPoint.X = lngX1
                                        LPoint.Y = lngCurY + T_DrawClient.时间列单位 / 2
                                        Call GetTextRect(objDraw, lngX1, lngCurY + T_DrawClient.时间列单位 / 2, strTmp, T_DrawClient.列单位 * 3, True, , sngScale)
                                        lngX1 = lngX1 + T_DrawClient.列单位 * 3
                                    End If
                                End If
                                
                                If blnOutText = True Then
                                    If AnsyGrade(Val(arrTmpString0(0)), strTmp, arrText) = True Then
                                        Call DrawAnsyGrade(lngDC, objDraw, arrText, LPoint, lngColor, bln灌肠大便以分子分母显示, sngScale)
                                    Else
                                        Call DrawTabText(lngDC, objDraw, strTmp, -1, T_LableRect, DT_CENTER, LPoint, sngScale)
                                    End If
                                End If
                   
                            End If
                        Next intCOl
                        
                        '画单元格竖线
                        If InStr(1, ",2,3,4,", "," & int频次 & ",") = 0 Then
                            lngX1 = T_DrawClient.刻度区域.Right + T_DrawClient.列单位 * (6 / int频次)
                            lngY1 = lngCurY + T_DrawClient.时间列单位
                            For intCOl = 1 To int频次 * 7
                                Call DrawLine(lngDC, lngX1, lngCurY, lngX1, lngY1, PS_SOLID, IIf(intCOl Mod int频次 = 0, intBold, intFine), RGB_BLACK)
                                lngX1 = lngX1 + T_DrawClient.列单位 * (6 / int频次)
                            Next intCOl
                            Call DrawLine(lngDC, T_DrawClient.刻度区域.Right, lngY1, T_DrawClient.体温区域.Right, lngY1, PS_SOLID, IIf(intRowCount = intRepairRows, intBold, intFine), RGB_BLACK)
                        ElseIf int频次 = 3 Then
                            intRowCount = intRowCount - (int频次 - intNum)
                            intValue = intRowCount
                            For i = 1 To 3 - intNum
                                lngX1 = T_DrawClient.刻度区域.Right + T_DrawClient.列单位 * 6
                                lngY1 = lngCurY + T_DrawClient.时间列单位
                                For intCOl = 1 To 7
                                    Call DrawLine(lngDC, lngX1, lngCurY, lngX1, lngY1, PS_SOLID, intBold, RGB_BLACK)
                                    lngX1 = lngX1 + T_DrawClient.列单位 * 6
                                Next intCOl
                                intRowCount = intValue + i
                                Call DrawLine(lngDC, T_DrawClient.刻度区域.Right, lngY1, T_DrawClient.体温区域.Right, lngY1, PS_SOLID, IIf(intRowCount = intRepairRows, intBold, intFine), RGB_BLACK)
                                
                                lngCurY = lngY1
                            Next i
                        ElseIf InStr(1, ",2,4,", "," & int频次 & ",") <> 0 Then
                            intRowCount = intRowCount - (int频次 / 2 - intNum)
                            intValue = intRowCount
                            For i = 1 To (int频次 / 2 - intNum)
                                lngY1 = lngCurY + T_DrawClient.时间列单位
                                lngX1 = T_DrawClient.刻度区域.Right + T_DrawClient.列单位 * 3
                                For intCOl = 1 To 14
                                    Call DrawLine(lngDC, lngX1, lngCurY, lngX1, lngY1, PS_SOLID, IIf(intCOl Mod 2 = 0, intBold, intFine), RGB_BLACK)
                                    lngX1 = lngX1 + T_DrawClient.列单位 * 3
                                Next intCOl
                                intRowCount = intValue + i
                                Call DrawLine(lngDC, T_DrawClient.刻度区域.Right, lngY1, T_DrawClient.体温区域.Right, lngY1, PS_SOLID, IIf(intRowCount = intRepairRows, intBold, intFine), RGB_BLACK)
                                lngCurY = lngY1
                            Next i
                        End If
                        
                        lngCurY = lngY1
                    End If
                End If
                
                intNum = 0
                
                '皮试结果,只输出标题和内容，表格在不空行中处理。
                If arrTmpString0(0) = "-999" Then
                    lngY1 = lngCurY
                    lngX1 = lngCurX
                    int频次 = 1
                    For intCOl = 0 To UBound(arrTmpString1)
                        If intCOl = 0 Then '开始画表头信息包括标题的输出
                            lngY1 = lngY1 + T_DrawClient.时间列单位
                            lngRowHeiht = T_DrawClient.时间列单位 / 2
                               
                            Call SetTextColor(lngDC, RGB_BLACK)
                            Call GetTextExtentPoint32(lngDC, arrTmpString0(intCOl + 1), Len(arrTmpString0(intCOl + 1)), T_Size)
                            Call GetTextRect(objDraw, lngX1, lngY1 - lngRowHeiht, arrTmpString0(intCOl + 1), T_DrawClient.刻度区域.Right - lngX1, True, , sngScale)
                
                            LPoint.X = lngX1
                            LPoint.Y = lngY1 - lngRowHeiht
                            LPoint.W = T_DrawClient.刻度区域.Right - lngX1
                            Call DrawTabText(lngDC, objDraw, arrTmpString0(intCOl + 1), -1, T_LableRect, DT_CENTER, LPoint, sngScale)
                            
                            lngY1 = lngCurY
                            lngX1 = T_DrawClient.刻度区域.Right
                        Else  '开始进行画表格线
                            intNum = 1
                            strTmp = CStr(arrTmpString1(intCOl))
                            If strTmp = "" Then strTmp = "-#"
                            LPoint.X = lngX1
                            LPoint.Y = lngCurY + T_DrawClient.时间列单位 / 2
                            LPoint.W = T_DrawClient.列单位 * (6 / int频次)
                            '开始计算是否需要换行
                            strPart = ""
                            
                            arrTmp = Split(strTmp, ",")
                            
                            For i = LBound(arrTmp) To UBound(arrTmp)
                                lngColor = Val(Split(arrTmp(i), "-#")(0))
                                '设置字体颜色
                                Call SetTextColor(lngDC, lngColor)
                                strTmp = Replace(CStr(Split(arrTmp(i), "-#")(1)), vbCrLf, "") '皮试结果
                                If Trim(strTmp) <> "" Then
                                    If i < UBound(arrTmp) Then strTmp = strTmp & ","
                                    Do While True
                                        T_Size.W = objDraw.TextWidth(strTmp) / T_TwipsPerPixel.X
                                        strPic = strTmp
                                        If T_Size.W - (LPoint.W - (LPoint.X - lngX1)) > 0 Then
                                            sngLen = Round((LPoint.W - (LPoint.X - lngX1)) / T_Size.W, 2)
                                            lngLen = Len(StrConv(strTmp, vbFromUnicode)) * sngLen
                                            '将半角转为全角
                                            strTmp = StrConv(strTmp, vbWide)
                                            strPart = StrConv(Mid(StrConv(strTmp, vbFromUnicode), lngLen + 1), vbUnicode)
                                            strTmp = StrConv(Mid(StrConv(strTmp, vbFromUnicode), 1, lngLen), vbUnicode)
                                            '截取原始字符串
                                            strPart = Mid(strPic, Len(strTmp) + 1)
                                            strTmp = Mid(strPic, 1, Len(strTmp))
                                            Call GetTextRect(objDraw, LPoint.X, LPoint.Y, CStr(strTmp), , True, , sngScale)
                                            Call DrawTabText(lngDC, objDraw, CStr(strTmp), -1, T_LableRect, DT_CENTER, LPoint, sngScale)
                                            
                                            T_Size.W = objDraw.TextWidth(strTmp) / T_TwipsPerPixel.X
                                            LPoint.X = LPoint.X + T_Size.W
                                            strTmp = strPart
                                            T_Size.W = objDraw.TextWidth("字") / T_TwipsPerPixel.X
                                            If T_Size.W - (LPoint.W - (LPoint.X - lngX1)) > 0 Then
                                                LPoint.X = lngX1
                                                LPoint.Y = LPoint.Y + T_DrawClient.时间列单位
                                                intNum = intNum + 1
                                                
                                                If intRowCount + intNum > intRepairRows Then GoTo ErrNext
                                            End If
                                            If strTmp = "" Then Exit Do
                                        Else
                                            Call GetTextRect(objDraw, LPoint.X, LPoint.Y, CStr(strTmp), , True, , sngScale)
                                            Call DrawTabText(lngDC, objDraw, CStr(strTmp), -1, T_LableRect, DT_CENTER, LPoint, sngScale)
                                            If T_Size.W + objDraw.TextWidth("字") / T_TwipsPerPixel.X - LPoint.W > 0 Then
                                                LPoint.X = lngX1
                                                LPoint.Y = LPoint.Y + T_DrawClient.时间列单位
                                            Else
                                                LPoint.X = LPoint.X + T_Size.W
                                            End If
                                    
                                            Exit Do
                                        End If
                                    Loop
                                End If
                            Next i
ErrNext:
                            
                            lngX1 = lngX1 + T_DrawClient.列单位 * (6 / int频次)
                        End If
                    Next intCOl
                End If
            End If
        Next intRow
        
        '补空行
        If intRepairRows > 0 And intRepairRows > intRowCount Then
            intRowCount = intRowCount + 1
            For intRow = intRowCount To intRepairRows
                lngX1 = lngCurX
                lngY1 = lngCurY + T_DrawClient.时间列单位
                
                '空格每行两列
'                For intCOl = 0 To 14
'                    If intCOl = 0 Then
'                        Call DrawLine(lngDc, lngX1, lngCurY, lngX1, lngY1, PS_SOLID, 2, RGB_BLACK)
'                        Call DrawLine(lngDc, lngX1, lngY1, T_DrawClient.刻度区域.Right, lngY1, PS_SOLID, IIf(intRow = intRepairRows, 2, 1), RGB_BLACK)
'                        Call DrawLine(lngDc, T_DrawClient.刻度区域.Right, lngCurY, T_DrawClient.刻度区域.Right, lngY1, PS_SOLID, 2, RGB_BLACK)
'                    Else
'
'                        lngX1 = T_DrawClient.刻度区域.Right + (T_DrawClient.列单位 * 3) * intCOl
'                        Call DrawLine(lngDc, lngX1, lngCurY, lngX1, lngY1, PS_SOLID, IIf(intCOl Mod 2 = 0, 2, 1), RGB_BLACK)
'                        If intCOl = 14 Then
'                            Call DrawLine(lngDc, T_DrawClient.刻度区域.Right, lngY1, T_DrawClient.体温区域.Right, lngY1, PS_SOLID, IIf(intRow = intRepairRows, 2, 1), RGB_BLACK)
'                        End If
'                    End If
'                Next intCOl
                
                '空格每行1列
                For intCOl = 0 To 7
                    If intCOl = 0 Then
                        Call DrawLine(lngDC, lngX1, lngCurY, lngX1, lngY1, PS_SOLID, intBold, RGB_BLACK)
                        Call DrawLine(lngDC, lngX1, lngY1, T_DrawClient.刻度区域.Right, lngY1, PS_SOLID, IIf(intRow = intRepairRows, intBold, intFine), RGB_BLACK)
                        Call DrawLine(lngDC, T_DrawClient.刻度区域.Right, lngCurY, T_DrawClient.刻度区域.Right, lngY1, PS_SOLID, intBold, RGB_BLACK)
                    Else
                        
                        lngX1 = T_DrawClient.刻度区域.Right + (T_DrawClient.列单位 * 6) * intCOl
                        Call DrawLine(lngDC, lngX1, lngCurY, lngX1, lngY1, PS_SOLID, intBold, RGB_BLACK)
                        If intCOl = 7 Then
                            Call DrawLine(lngDC, T_DrawClient.刻度区域.Right, lngY1, T_DrawClient.体温区域.Right, lngY1, PS_SOLID, IIf(intRow = intRepairRows, intBold, intFine), RGB_BLACK)
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
    '功能：画出最底部说明
    '参数:intPageNO=页码
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
    blnPrintCurveInfo = (Val(zldatabase.GetPara("体温单不打印曲线说明", glngSys, 1255, "0")) = 1)
    If blnPrintCurveInfo = False Then
        '打印体温说明信息
        Call SetTextColor(lngDC, RGB_BLACK)
        Call GetTextExtentPoint32(lngDC, strInfo, Len(strInfo), T_Size)
        Call GetTextRect(objDraw, X, Y, strInfo, 0, False, , sngScale)
        Call DrawText(lngDC, strInfo, -1, T_LableRect, DT_CENTER)
        Y = Y + IIf(blnPrinter = True, msngTwips, 1) * 30
    Else
        Y = Y + IIf(blnPrinter = True, msngTwips, 1) * 10
    End If
    
    blnWeek = (Val(zldatabase.GetPara("打印周数", glngSys, 1255, "0")) = 1)
    blnPageNo = (Val(zldatabase.GetPara("打印页号", glngSys, 1255, "1")) = 1)
    
    
    '打印页码
    '------------------------------------------------------------------------------------------------------------------
    If intPageNo > -1 And blnPageNo Then
        intPageNo = intPageNo + intBeginPage - 1
        strNOPage = "第   --" & CStr(intPageNo) & "--   页"
    End If
    
    If blnWeek Then
        If strNOPage = "" Then
            strNOPage = "第   " & CStr(intBeginPage) & "   周"
        Else
            strNOPage = strNOPage & "(第 " & CStr(intBeginPage) & " 周)"
        End If
    End If
    
    Call SetTextColor(lngDC, RGB_BLACK)
    Call GetTextExtentPoint32(lngDC, strNOPage, Len(strNOPage), T_Size)
    Call GetTextRect(objDraw, 0, Y, strNOPage, objDraw.Width / T_TwipsPerPixel.X, True, , sngScale)
    Call DrawText(lngDC, strNOPage, -1, T_LableRect, DT_CENTER)
    
    '输出打印人,即当前操作员姓名
    '------------------------------------------------------------------------------------------------------------------
'    strNOPage = "打印人:" & gstrUserName
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
'功能 处理表下表格字体输出
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
    stdset.Name = "宋体"
    stdset.Size = intSize * sngScale
    If stdset.Size < 9 Then
        stdset.Name = "Times New Roman"
    End If
    stdset.Bold = False
    Call SetFontIndirect(stdset, lngDC, objDraw)
    lngFont = CreateFontIndirect(T_Font)
    lngOldFont = SelectObject(lngDC, lngFont)
    If blnChage = True Then '重新计算输出字体位置
        Call GetTextRect(objDraw, LPoint.X, LPoint.Y, strTmp, LPoint.W, True, , sngScale)
    End If
    Call DrawText(lngDC, strTmp, -1, T_LableRect, DT_CENTER)
    
    Call SelectObject(lngDC, lngOldFont)
    Call DeleteObject(lngFont)
    
    '还原对象字体
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
'功能 大便次数输出
'说明 AnsyGrade=True才能调用此函数
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
    stdset.Name = "宋体"
    stdset.Size = intSize * sngScale
    stdset.Bold = False
    Set stdOldset = stdset
    
    Call GetTextRect(objDraw, LPoint.X, LPoint.Y, strTmp, LPoint.W, True, , sngScale)
    '输出左边
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
    
    If blnFormat = True Then '分子分母显示
        intSize = 7
        objDraw.Font.Size = intSize * sngScale
        Set stdset = New StdFont
        stdset.Name = "宋体"
        stdset.Size = intSize * sngScale
        Call SetFontIndirect(stdset, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        Call SetTextColor(lngDC, lngColor)
        T_LableRect.Left = lngX
        lngY = T_LableRect.Top
        sngH = objDraw.TextHeight("A") / T_TwipsPerPixel.X / 2
        T_LableRect.Top = lngY - sngH
        T_LableRect.Bottom = LPoint.Y + (T_DrawClient.时间列单位 / 2)
        Call DrawText(lngDC, str2, -1, T_LableRect, 0)
        lngY = T_LableRect.Top + (objDraw.TextHeight("A") / T_TwipsPerPixel.Y)
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
        '画横线
        objDraw.Font.Size = 9 * sngScale
        Call DrawLine(lngDC, lngX, lngY, lngX + (objDraw.TextWidth("A") / T_TwipsPerPixel.X), lngY)
        '输出分母
        lngY = lngY
        T_LableRect.Left = lngX
        T_LableRect.Top = lngY
        intSize = 7.5
        objDraw.Font.Size = intSize * sngScale
        Set stdset = New StdFont
        stdset.Name = "宋体"
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
            '输出上标
            intSize = 7
            objDraw.Font.Size = intSize * sngScale
            Set stdset = New StdFont
            stdset.Name = "宋体"
            stdset.Size = intSize * sngScale
            Call SetFontIndirect(stdset, lngDC, objDraw)
            lngFont = CreateFontIndirect(T_Font)
            lngOldFont = SelectObject(lngDC, lngFont)
            Call SetTextColor(lngDC, lngColor)
            T_LableRect.Left = lngX
            lngY = T_LableRect.Top
            sngH = objDraw.TextHeight("A") / T_TwipsPerPixel.Y / 2
            T_LableRect.Top = lngY - sngH
            If T_LableRect.Top < (LPoint.Y - (T_DrawClient.时间列单位 / 2)) Then T_LableRect.Top = (LPoint.Y - (T_DrawClient.时间列单位 / 2)) - msngTwips
            Call DrawText(lngDC, str2, -1, T_LableRect, 0)
            Call SelectObject(lngDC, lngOldFont)
            Call DeleteObject(lngFont)
            lngX = lngX + (objDraw.TextWidth(str2) / T_TwipsPerPixel.X)
            '输出后半部分
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
    '大便次数符合1/E 或 1+1/E的格式 或1 1/E的格式
    Dim intPos As Integer
    Dim ArrCode
    Dim str1 As String, str2 As String, str3 As String
    
    strText = Trim(strText)
    If strText = "" Or lngItemNO <> gint大便 Then Exit Function
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
    strEditors() As Variant, Optional int心率引用 As Integer = 1, Optional ByVal sngScale As Single = 1) As String
'-------------------------------------------------------------------------------------
'功能:输出体温项目的连线和图形输出
'参数::lngDC 绘图对象的DC，objDraw 绘画对象.rsPoint 所有项目点的集合(序号|数值|部位|标记|时间|项目序号|复查|断开|重叠项目|重叠|X坐标|Y坐标|备注|符号)
'strEditors 体温，心率，呼吸，脉搏的信息(项目序号||项目名称||项目单位||项目值域||记录符||记录色)
'返回:心率点的集合 !X坐标 & ";" & !Y坐标 & "," & !X坐标 & ";" & !Y坐标
'-------------------------------------------------------------------------------------
    Dim sin原X As Single, sin原Y As Single
    Dim lng项目序号 As Long
    Dim SinX As Single, sinY As Single  '物理降温使用
    Dim dblvalue As Double
    Dim dblMaxValue As Double, dblMinValue As Double
    Dim lngRGB As Long
    Dim strChar As String, str部位 As String, strTmp As String, strPic As String
    Dim str心率 As String
    Dim lngCount As Long '重叠项目数量
    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset
    Dim blnLine As Boolean
    Dim i As Integer
    Dim X1 As Single
    Dim blnPrinter As Boolean
    Dim intBold As Integer, intFine As Integer
    Dim bln不升符号 As Boolean
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
    rsPoint.Sort = "项目序号,时间"
    '首先进行连线
    With rsPoint
        Do While Not .EOF
            For i = 0 To UBound(strEditors)
                If Val(Split(strEditors(i), "||")(0)) = Val(zlCommFun.Nvl(!项目序号)) Then
                     Exit For
                End If
            Next i
            If Not (zlCommFun.Nvl(!项目序号) = gint体温 And Val(zlCommFun.Nvl(!标记)) = 1) Then
                If zlCommFun.Nvl(!项目序号) <> lng项目序号 Then
                    sin原X = 0
                    sin原Y = 0
                    lngRGB = Split(CStr(strEditors(i)), "||")(5)
                    lng项目序号 = zlCommFun.Nvl(!项目序号)
                End If
                If int心率引用 = 2 Then
                    If !项目序号 = -1 Then
                        blnLine = True
                    Else
                        blnLine = True
                    End If
                Else
                    blnLine = True
                End If
                
                If sin原X <> 0 And blnLine Then
                    Call DrawLine(lngDC, sin原X + T_DrawClient.列单位 / 2, sin原Y, !X坐标 + T_DrawClient.列单位 / 2, !Y坐标, PS_SOLID, intFine, lngRGB)
                End If
                If !断开 = 0 Then
                    sin原X = zlCommFun.Nvl(!X坐标, 0)
                    sin原Y = zlCommFun.Nvl(!Y坐标, 0)
                Else
                    sin原X = 0
                End If
                
                If !项目序号 = gint体温 Then
                    If zlCommFun.Nvl(!复查) = 1 Then '复试合格
                        Call SetTextColor(lngDC, lngRGB)
                        Call GetTextRect(objDraw, !X坐标, !Y坐标 - T_DrawClient.行单位, "v", T_DrawClient.列单位, True, , sngScale)
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
                
                '临界值不等空,并且在最大值和最小值之间
                If Split(strEditors(i), "||")(8) <> "" And Val(Split(strEditors(i), "||")(8)) <= Val(Split(strEditors(i), "||")(6)) _
                    And Val(Split(strEditors(i), "||")(8)) >= Val(Split(strEditors(i), "||")(7)) Then dblMaxValue = Val(Split(strEditors(i), "||")(8))
                    
                If zlCommFun.Nvl(!项目序号) = gint体温 And Trim(zlCommFun.Nvl(!数值)) = "不升" Then
                    dblvalue = dblMinValue
                Else
                    dblvalue = Val(zlCommFun.Nvl(!数值))
                End If
                
                If dblvalue > dblMaxValue Then
                    Call DrawLine(lngDC, !X坐标 + T_DrawClient.列单位 / 2, !Y坐标 - T_DrawClient.行单位 * 2, !X坐标 + T_DrawClient.列单位 / 2, !Y坐标, PS_SOLID, intFine, lngRGB, True)
                ElseIf dblvalue < dblMinValue Then
                    Call DrawLine(lngDC, !X坐标 + T_DrawClient.列单位 / 2, !Y坐标 + T_DrawClient.行单位 * 2, !X坐标 + T_DrawClient.列单位 / 2, !Y坐标, PS_SOLID, intFine, lngRGB, True)
                End If
            Else
                '体温的物理降温
                dblvalue = Split(!备注, ",")(0)
                SinX = Val(Split(Split(!备注, ",")(1), ";")(0))
                sinY = Val(Split(Split(!备注, ",")(1), ";")(1))
                T_Size.H = objDraw.TextHeight("○") / T_TwipsPerPixel.Y

                If Val(!数值) > Val(dblvalue) Then
                    '物理降温失败，画带箭头的红色实线，字符固定用○
                    'Call DrawLine(lngDC, !X坐标 + T_DrawClient.列单位 / 2, !Y坐标, SinX + T_DrawClient.列单位 / 2, sinY, PS_SOLID, intFine, RGB_RED, True)
                    '现在失败也为虚线(医院要求)
                    Call DrawLine(lngDC, !X坐标 + T_DrawClient.列单位 / 2, !Y坐标 + (T_Size.H / 4), SinX + T_DrawClient.列单位 / 2, sinY, IIf(blnPrinter, PS_DASH, PS_DOT), 1, RGB_RED, True)
                ElseIf Val(!数值) < Val(dblvalue) Then
                    '物理降温成功，画红色虚线，字符固定用○
                    Call DrawLine(lngDC, !X坐标 + T_DrawClient.列单位 / 2, !Y坐标 - (T_Size.H / 2), SinX + T_DrawClient.列单位 / 2, sinY, IIf(blnPrinter, PS_DASH, PS_DOT), 1, RGB_RED, False)
                End If
            End If
            .MoveNext
        Loop
    End With
    If rsPoint.RecordCount > 0 Then rsPoint.MoveFirst
    '输出所有点的图形
    With rsPoint
        Do While Not .EOF
            str部位 = ""
            strTmp = ""
            For i = 0 To UBound(strEditors)
                If Split(CStr(strEditors(i)), "||")(0) = Val(zlCommFun.Nvl(!项目序号)) Then
                     Exit For
                End If
            Next i
            If zlCommFun.Nvl(!重叠) = 0 And zlCommFun.Nvl(!重叠项目) = "空" Then '未重叠的项目
                lngRGB = Split(CStr(strEditors(i)), "||")(5)
                If zlCommFun.Nvl(!项目序号) = -1 And int心率引用 = 2 Then lngRGB = RGB_RED
                str部位 = zlCommFun.Nvl(!部位)
                If str部位 = "" Then
                    Select Case lng项目序号
                        Case gint体温
                            str部位 = "腋温"
                        Case gint呼吸
                            str部位 = "自主呼吸"
                        Case Else
                            str部位 = ""
                    End Select
                End If
                strTmp = Split(CStr(strEditors(i)), "||")(4)
                strPic = ""
                strChar = ""
                Select Case zlCommFun.Nvl(!项目序号)
                    Case gint体温
                        strTmp = strTmp & String(2 - UBound(Split(strTmp, ",")), ",")
                        If str部位 = "口温" Then
                            strChar = Split(strTmp, ",")(0)
                        ElseIf str部位 = "腋温" Then
                            strChar = Split(strTmp, ",")(1)
                        Else
                            strChar = Split(strTmp, ",")(2)
                        End If
                        If zlCommFun.Nvl(!标记) = 1 Then '物理降温符号
                            lngRGB = RGB_RED
                            strChar = "○"
                        Else
                            If strChar = "" Then strChar = "×"
                        End If
                    Case gint心率
                        strChar = IIf(strTmp = "", "Ο", strTmp)
                    Case gint脉搏
                        If str部位 = "起搏器" Then
                            strPic = "PACEMAKER"
                        Else
                            strChar = IIf(strTmp = "", "+", strTmp)
                        End If
                    Case gint呼吸
                        If str部位 = "自主呼吸" Then
                            strChar = IIf(strTmp = "", "*", strTmp)
                        Else
                            strPic = "BREATH"
                        End If
                    Case Else
                        strChar = strTmp
                End Select
                If Trim(zlCommFun.Nvl(!符号)) <> "" Then
                    strChar = Trim(zlCommFun.Nvl(!符号))
                    strPic = ""
                End If
                
                If !项目序号 = gint体温 And Trim(Nvl(!数值)) = "不升" And (mlng体温不升显示方式 = 0 Or mlng体温不升显示方式 = 1) Then
                    bln不升符号 = False
                Else
                    bln不升符号 = True
                End If
                                
                If strPic = "" And bln不升符号 Then
                    Call SetTextColor(lngDC, lngRGB)
                    Call GetTextRect(objDraw, !X坐标, !Y坐标, Trim(strChar), T_DrawClient.列单位, True, , sngScale)
                    Call DrawText(lngDC, Trim(strChar), -1, T_LableRect, DT_CENTER)
                    'Debug.Print T_LableRect.Left & ";" & T_LableRect.Right
                Else
                    Call DrawPicture(objDraw, strPic, objDraw.ScaleX(!X坐标 + ((T_DrawClient.列单位 - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2), vbPixels, vbTwips), objDraw.ScaleY(!Y坐标 - mintBmpH * IIf(blnPrinter = True, msngTwips, 1) / 2, vbPixels, vbTwips), _
                        objDraw.ScaleX(!X坐标 + ((T_DrawClient.列单位 - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2) + mintBmpW * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), objDraw.ScaleY(!Y坐标 + mintBmpH * IIf(blnPrinter = True, msngTwips, 1) / 2, vbPixels, vbTwips), True)
                End If
            
            Else  '展示重叠部位图标
                strPic = ""
                strChar = ""
                If zlCommFun.Nvl(!重叠项目) <> "空" Then '重叠=1的不做任何处理
                    lngCount = UBound(Split(zlCommFun.Nvl(!重叠项目), ","))
                    strTmp = zlCommFun.Nvl(!重叠项目)
                    If Trim(strTmp) <> "" Then
                        str部位 = zlCommFun.Nvl(!部位)
                        lngCount = lngCount + 2
                        strTmp = zlCommFun.Nvl(!项目序号) & "," & strTmp
                        If InStr(1, "," & strTmp & ",", ",1,") <> 0 Then

                            strSql = "SELECT A.序号,A.标记符号,A.标记颜色" & vbNewLine & _
                                    " FROM 体温重叠标记 A," & vbNewLine & _
                                    "     (SELECT 上级序号, COUNT(*) 数量" & vbNewLine & _
                                    "     FROM 体温重叠标记" & vbNewLine & _
                                    "     WHERE 项目序号 IN (" & strTmp & ")" & vbNewLine & _
                                    "     GROUP BY 上级序号) B" & vbNewLine & _
                                    " WHERE A.重叠数目 = B.数量" & vbNewLine & _
                                    " AND A.序号 = B.上级序号 AND A.序号=[1]"
                        Else
                            strSql = "Select A.序号, A.标记符号, A.标记颜色" & vbNewLine & _
                                "  From 体温重叠标记 A," & vbNewLine & _
                                "       (Select 上级序号, Count(1) 数量" & vbNewLine & _
                                "          from 体温重叠标记" & vbNewLine & _
                                "         where 项目序号 in (" & strTmp & ")" & vbNewLine & _
                                "         group by 上级序号) B" & vbNewLine & _
                                " Where A.重叠数目 = B.数量" & vbNewLine & _
                                "   And A.序号 = B.上级序号"
                        End If
                        
                        Set rsTmp = zldatabase.OpenSQLRecord(strSql, "重叠", 1)
                        
                        If rsTmp.RecordCount > 0 Then
                            If IsNull(rsTmp!标记符号) Then
                                strPic = zlBlobRead(9, zlCommFun.Nvl(rsTmp!序号))
                            Else
                                strChar = Trim(zlCommFun.Nvl(rsTmp!标记符号))
                                lngRGB = Val(zlCommFun.Nvl(rsTmp!标记颜色, 0))
                            End If
                            If strPic = "" Then
                                Call SetTextColor(lngDC, lngRGB)
                                Call GetTextRect(objDraw, !X坐标 - 1, !Y坐标, Trim(strChar), T_DrawClient.列单位, True, , sngScale)
                                Call DrawText(lngDC, Trim(strChar), -1, T_LableRect, DT_CENTER)
                            Else
                                Call DrawPicture(objDraw, strPic, objDraw.ScaleX(!X坐标 + ((T_DrawClient.列单位 - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2), vbPixels, vbTwips), objDraw.ScaleY(!Y坐标 - mintBmpH * IIf(blnPrinter = True, msngTwips, 1) / 2, vbPixels, vbTwips), _
                                    objDraw.ScaleX(!X坐标 + ((T_DrawClient.列单位 - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2) + mintBmpW * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), objDraw.ScaleY(!Y坐标 + mintBmpH * IIf(blnPrinter = True, msngTwips, 1) / 2, vbPixels, vbTwips), False)
                                
                                Call FileSystem.Kill(strPic)
                            End If
                        End If
                    End If
                End If
            End If
        .MoveNext
        Loop
    End With
    
    '提取所有心率的信息
    If rsPoint.RecordCount > 0 Then rsPoint.MoveFirst
    rsPoint.Filter = "项目序号=" & gint心率
    With rsPoint
        Do While Not .EOF
            str心率 = str心率 & "," & !X坐标 & ";" & !Y坐标
        .MoveNext
        Loop
    End With
    If str心率 <> "" Then str心率 = Mid(str心率, 2)
    
    ShowPoints = str心率
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function GetCanvasCenter(ByVal dtBegin As Date, ByVal dtEnd As Date, ByVal dtBeginDate As Date, ByVal SinX As Single) As Boolean
'---------------------------------------------------------
'功能:判断该时间点是否是中间值
'参数:dtbegin:被比较的时间段.  dtend:要比较的时间段 . dtBeginDate 本页体温单的开始时间 .sinx当前点的X坐标
'---------------------------------------------------------
    Dim blnTrue As Boolean
    Dim strTime As String, strTmp As String
    Dim intDay As Integer, intTime As Integer, strDay As String
    
    
    intTime = (SinX - T_DrawClient.体温区域.Left) \ T_DrawClient.列单位
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

Public Function DrawCanvas(ByVal lngDC As Long, ByVal objDraw As Object, ByVal rsTemp As ADODB.Recordset, rsDrawItems As ADODB.Recordset, Optional ByVal bln不打印心率列 As Boolean = False, Optional sngScale As Single = 1) As String
'------------------------------------------------------------------------------------------------------
'功能:画刻度区域和体温区域并输出刻度值信息
'参数:lngDC 绘图对象的DC，objDraw 绘画对象.rsTemp:体温曲线项目记录集(A.项目序号,A.排列序号,A.记录名,A.记录符,A.记录色,A.最大值,A.最小值,A.单位值,C.项目单位 单位,A.最高行-2 AS 最高行,B.部位)
'出参:返回各个曲线的具体信息包括( "项目序号|最大值|最小值|单位值|最大值坐标|最小值坐标|单位刻度|显示模式|颜色")
'返回说明信息(项目的符号)
'-------------------------------------------------------------------------------------------------------
    Dim str说明 As String
    Static SlngMaxY As Long                 '记录上一次的最大高度，以决定本次是否需要重画
    Dim lngCurX     As Long, lngCurY As Single  '当前位置
    Dim lngMaxX     As Long, lngMaxY As Single  '边界
    Dim lngCurAlerY As Single '警戒线
    Dim lngRow      As Long
    Dim intLables   As Integer
    Dim bln双行 As Boolean                  '此参数由用户指定,bln双行=TRUE表示只显示五行;否则显示十行
    Dim bln粗线 As Boolean                  '此参数由用户指定,大行分界是粗线还是细线
    '以下都是标准尺度
    Dim intLineMode   As Integer
    Dim blnDoubleRow  As Boolean             '贰行做为一行打印输出
    Dim sinAlertness  As Single              '警戒线,起辅助作用
    Dim lngLableStep  As Long
    Dim lngColStep    As Long
    Dim sinRowStep As Single, lngInitRowStep As Long
    Dim arrTemp()     As String
    Dim blnPrinter As Boolean
    Dim intBold As Integer, intFine As Integer
    Dim lngFont As Long, lngOldFont As Long
    Dim sinY单位 As Single '曲线单位输出的Bottom
    
    '以下与绘图区域相关(项目序号,最大值,最小值,单位值,最大值坐标,最小值坐标,单位刻度,显示模式)
    Dim sin刻度 As Single, bln显示刻度 As Boolean
    Dim sin刻度间隔 As Single, sinBegin刻度 As Single, dbl单位值 As Double
    
    Dim str最大值坐标 As String, str最小值坐标 As String

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
    '所有曲线项目的作图区域(项目序号,最大值,最小值,单位值,最大值坐标,最小值坐标,单位刻度,显示模式)
    gstrFields = "项目序号," & adDouble & ",18|最大值," & adDouble & ",18|最小值," & adDouble & ",18|" & "单位值," & adDouble & _
        ",18|最大值坐标," & adLongVarChar & ",20|最小值坐标," & adLongVarChar & ",20|" & "单位刻度," & adLongVarChar & ",20|显示模式," & adDouble & ",5|颜色," & adDouble & ",18"
    Call Record_Init(rsDrawItems, gstrFields)
    '------------------------------------------------------------------------------------------------------------------
    '赋初值
    intLineMode = PS_SOLID
    lngLableStep = T_DrawClient.刻度单位
    lngColStep = T_DrawClient.列单位
    lngInitRowStep = glngInitRowStep * IIf(blnPrinter = True, msngTwips, 1)
    sinRowStep = T_DrawClient.行单位
    
    '体温单以单格显示(不勾此选项以双格显示，没两个刻度显示一次) 1：单格显示 0：双格显示
    If zldatabase.GetPara("体温单显示格式", glngSys, 1255, 0) = 1 Then
        bln双行 = False
    Else
        bln双行 = True
    End If
    'True表示贰行只输出一行,效果是一个刻度只显示了五行;否则一个刻度显示十行,由用户调整参数决定,与blnDoubleRow无关
    bln粗线 = True
    
    If Not bln粗线 Then intLineMode = PS_DASHDOTDOT
    
    '画表格
    rsTemp.Filter = "项目序号=" & gint心率
    If rsTemp.RecordCount > 0 And bln不打印心率列 = True Then
        rsTemp.Filter = 0
        intLables = rsTemp.RecordCount - 1
    Else
        rsTemp.Filter = 0
        intLables = rsTemp.RecordCount
    End If
    If intLables <= 0 Then intLables = 1
    lngCurX = T_DrawClient.偏移量X
    lngCurY = T_DrawClient.偏移量Y
    lngMaxX = (intLables * lngLableStep) + (7 * 6 * lngColStep) + T_DrawClient.偏移量X  '刻度+7*宽度+偏移量X
    lngMaxY = 2 * mintNullRow * lngInitRowStep + T_DrawClient.总列数 * sinRowStep + T_DrawClient.偏移量Y '（为表格大小，还需加上起始Y坐标）
       
    str说明 = ""
        
    SlngMaxY = lngMaxY
    T_DrawClient.刻度单位 = lngLableStep
    T_DrawClient.行单位 = sinRowStep
    T_DrawClient.列单位 = lngColStep
    T_DrawClient.双倍 = blnDoubleRow
    
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
    
    
    T_DrawClient.刻度区域.Left = T_DrawClient.偏移量X
    T_DrawClient.刻度区域.Top = lngCurY
    T_DrawClient.刻度区域.Right = lngCurX
    T_DrawClient.刻度区域.Bottom = lngMaxY
    
    '默认添加一行用于显示项目名称
    lngCurY = lngCurY + lngInitRowStep * 2
    Call DrawLine(lngDC, T_DrawClient.偏移量X, lngCurY, lngMaxX, lngCurY, PS_SOLID, intFine, RGB_BLACK)
    lngCurY = lngCurY + lngInitRowStep * ((mintNullRow - 1) * 2)
    '画体温单所有行
    For lngRow = 0 To T_DrawClient.总列数
        If lngRow <> 0 Then
            lngCurY = lngCurY + sinRowStep
        End If
        '画体温单的所有行
        If ((blnDoubleRow Or bln双行) And lngRow Mod 2 = 0) Or (Not blnDoubleRow And Not bln双行) Then
            Call DrawLine(lngDC, lngCurX, lngCurY, lngMaxX, lngCurY, IIf(lngRow Mod 10 = 0, PS_SOLID, intLineMode), IIf(lngRow Mod 5 = 0 And sinRowStep >= 4 And bln粗线, intBold, intFine), RGB_BLACK)
        End If
    Next
    
    lngCurY = T_DrawClient.刻度区域.Top
    
    '画体温单所有列
    For lngRow = 1 To 6 * 7
        lngCurX = lngCurX + lngColStep
        Call DrawLine(lngDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, IIf(lngRow Mod 6 = 0, intBold, intFine), IIf(lngRow Mod 6 = 0, RGB_RED, RGB_BLACK))
    Next
    
    lngCurX = T_DrawClient.刻度区域.Right
    T_DrawClient.体温区域.Left = T_DrawClient.刻度区域.Right
    T_DrawClient.体温区域.Top = T_DrawClient.刻度区域.Top
    T_DrawClient.体温区域.Right = lngMaxX
    T_DrawClient.体温区域.Bottom = lngMaxY
    
    '画刻度框的标尺（从固定不变的10行开始标识）
    intLables = 1
    rsTemp.Sort = "排列序号"
    With rsTemp
        Do While Not .EOF
            If Not (bln不打印心率列 = True And !项目序号 = gint心率) Then
                '显示刻度框项目的名称及符号,如体温×
                lngCurX = T_DrawClient.刻度区域.Left + ((intLables - 1) * T_DrawClient.刻度单位)
                lngCurY = T_DrawClient.刻度区域.Top
                 
                gstdSet.Name = "宋体"
                gstdSet.Size = 9 * sngScale
                Call SetFontIndirect(gstdSet, lngDC, objDraw)
                lngFont = CreateFontIndirect(T_Font)
                lngOldFont = SelectObject(lngDC, lngFont)
                '输出体温项目的名称
                Call SetTextColor(lngDC, zlCommFun.Nvl(!记录色, RGB_BLACK))
                Call GetTextRect(objDraw, lngCurX, lngCurY + objDraw.TextHeight(zlCommFun.Nvl(!记录名)) / T_TwipsPerPixel.Y / 2, Trim(zlCommFun.Nvl(!记录名)), T_DrawClient.刻度单位, , , sngScale)
                Call DrawText(lngDC, Trim(zlCommFun.Nvl(!记录名)), -1, T_LableRect, DT_CENTER)
                Call SelectObject(lngDC, lngOldFont)
                Call DeleteObject(lngFont)
                
                '设置字体大小
                gstdSet.Name = "宋体"
                gstdSet.Size = 8 * sngScale
                Call SetFontIndirect(gstdSet, lngDC, objDraw)
                lngFont = CreateFontIndirect(T_Font)
                lngOldFont = SelectObject(lngDC, lngFont)
    
                '输出项目单位
                Call GetTextRect(objDraw, lngCurX, lngCurY + lngInitRowStep * 2 + objDraw.TextHeight(zlCommFun.Nvl(!单位)) / T_TwipsPerPixel.Y / 2, Trim(zlCommFun.Nvl(!单位)), T_DrawClient.刻度单位, , , sngScale)
                Call DrawText(lngDC, Trim(zlCommFun.Nvl(!单位, 0)), -1, T_LableRect, DT_CENTER)
                Call SelectObject(lngDC, lngOldFont)
                Call DeleteObject(lngFont)
                sinY单位 = T_LableRect.Bottom
                intLables = intLables + 1
            End If
            objDraw.Font.Size = 9 * sngScale
            '强制设定体温曲线项目的显示模式
            Select Case !项目序号

                Case gint体温  '体温整数时输出刻度
                    sin刻度间隔 = zlCommFun.Nvl(!刻度间隔, 1)
                    dbl单位值 = 0.1
                    sinAlertness = zlCommFun.Nvl(!警示线, 37)
                    arrTemp = Split(zlCommFun.Nvl(!记录符, "·,×,○"), ",")
                    str说明 = str说明 & "、" & zlCommFun.Nvl(!记录名) & "(口温" & arrTemp(0) & ",腋温" & arrTemp(1) & ",肛温" & arrTemp(2) & ")"

                Case gint脉搏, gint心率  '脉搏/心跳按10的倍数输出刻度
                    sin刻度间隔 = zlCommFun.Nvl(!刻度间隔, 10)
                    dbl单位值 = 2
                    sinAlertness = zlCommFun.Nvl(!警示线, 0)

                    If !项目序号 = gint脉搏 Then
                        str说明 = str说明 & "、" & zlCommFun.Nvl(!记录名) & "(缺省记录符" & zlCommFun.Nvl(!记录符, "+") & ",起搏器H)"
                    Else
                        str说明 = str说明 & "、" & zlCommFun.Nvl(!记录名) & "(" & zlCommFun.Nvl(!记录符, "Ο") & ")"
                    End If

                Case gint呼吸  '呼吸按5的倍数输出刻度
                    mbln呼吸曲线 = True
                    dbl单位值 = 1
                    sin刻度间隔 = zlCommFun.Nvl(!刻度间隔, 5)
                    sinAlertness = zlCommFun.Nvl(!警示线, 0)
                    str说明 = str说明 & "、" & zlCommFun.Nvl(!记录名) & "(自主呼吸" & zlCommFun.Nvl(!记录符, "*") & ",呼吸机R)"
                Case Else
                    dbl单位值 = Val(zlCommFun.Nvl(!单位值, 0))
                    sin刻度间隔 = zlCommFun.Nvl(!刻度间隔, Val(zlCommFun.Nvl(!单位值, 0)) * 10)
                    If sin刻度间隔 > Val(zlCommFun.Nvl(!最大值)) - Val(zlCommFun.Nvl(!最小值)) Then
                        sin刻度间隔 = Val(zlCommFun.Nvl(!最大值)) - Val(zlCommFun.Nvl(!最小值))
                    End If
                    sinAlertness = zlCommFun.Nvl(!警示线, 0)
                    str说明 = str说明 & "、" & zlCommFun.Nvl(!记录名) & "(" & zlCommFun.Nvl(!记录符, "*") & ")"
            End Select

            '赋初值
            lngCurY = lngCurY + (lngInitRowStep * 2 * mintNullRow) '固定前4行的高度不输出刻度

            '根据最高行定位到有效位置
            lngCurY = lngCurY + (T_DrawClient.行单位 * zlCommFun.Nvl(!最高行, 0))
            Do While True
                bln显示刻度 = False
                If sin刻度 = 0 Then     '刚进入循环，此时取的最大值
                    sin刻度 = zlCommFun.Nvl(!最大值, 0)
                    sinBegin刻度 = sin刻度
                    str最大值坐标 = T_DrawClient.体温区域.Left & "," & lngCurY
                Else                    '计算得到每个刻度的值
                    sin刻度 = sin刻度 - dbl单位值     '如果目前显示模式为双倍，则按双倍累计
                End If
                
                '根据设置的刻度间隔显示刻度值
                If Val(Format(sin刻度, "#0.00")) = Val(Format(sinBegin刻度, "#0.00")) Then bln显示刻度 = True
                If bln显示刻度 = True Or sin刻度 < sinBegin刻度 Then sinBegin刻度 = sinBegin刻度 - IIf(T_DrawClient.双倍, sin刻度间隔 * 2, sin刻度间隔)
                If sinBegin刻度 < 0 Then sinBegin刻度 = 0
                
                If bln显示刻度 And Not (bln不打印心率列 = True And !项目序号 = gint心率) Then
                    '控制最大值不与曲线单位重复
                    If sin刻度 = Val(Nvl(!最大值, 0)) And lngCurY < sinY单位 Then
                        Call GetTextRect(objDraw, lngCurX, sinY单位, Format(sin刻度, "#0"), T_DrawClient.刻度单位, , , sngScale)
                    ElseIf Format(lngCurY, "#0") = T_DrawClient.刻度区域.Bottom Then
                        Call GetTextRect(objDraw, lngCurX, lngCurY - (objDraw.TextHeight("1") / 2 / T_TwipsPerPixel.Y), Format(sin刻度, "#0"), T_DrawClient.刻度单位, , , sngScale)
                    Else
                        Call GetTextRect(objDraw, lngCurX, lngCurY, Format(sin刻度, "#0"), T_DrawClient.刻度单位, , , sngScale)
                    End If
                    Call DrawText(lngDC, Format(sin刻度, "#0"), -1, T_LableRect, DT_CENTER)
                End If
                '如果不在有效范围内，或者超出画布则退出
                If Val(Format(sin刻度, "#0.00")) <= Val(Format(zlCommFun.Nvl(!最小值), "#0.00")) Or Format(lngCurY, "#0") > T_DrawClient.刻度区域.Bottom Then
                    str最小值坐标 = T_DrawClient.体温区域.Left & "," & lngCurY
                    '添加该项目(项目序号,最大值,最小值,单位值,最大值坐标,最小值坐标,单位刻度,显示模式)
                    gstrFields = "项目序号|最大值|最小值|单位值|最大值坐标|最小值坐标|单位刻度|显示模式|颜色"
                    gstrValues = zlCommFun.Nvl(!项目序号) & "|" & zlCommFun.Nvl(!最大值, 0) & "|" & zlCommFun.Nvl(!最小值, 0) & _
                    "|" & dbl单位值 & "|" & str最大值坐标 & "|" & str最小值坐标 & "|" & T_DrawClient.行单位 & "," & T_DrawClient.列单位 & "|" & sin刻度间隔 & "|" & !记录色
                    Call Record_Add(rsDrawItems, gstrFields, gstrValues)
                    
                    '辅助线或警示线
                    If blnDoubleRow = False And (sinAlertness < Val(Nvl(!最大值)) And sinAlertness > Val(Nvl(!最小值))) Then
                        lngCurAlerY = Val(GetYCoordinate(objDraw, rsDrawItems, Val(Nvl(!项目序号)), sinAlertness))
                        Call DrawLine(lngDC, T_DrawClient.体温区域.Left, lngCurAlerY, lngMaxX, lngCurAlerY, intLineMode, intBold, RGB_RED)
                    End If
                    
                    Exit Do
                End If
                lngCurY = lngCurY + T_DrawClient.行单位
            Loop
            sinBegin刻度 = 0
            sin刻度 = 0                 '控制从第一行开始输出
            .MoveNext
        Loop
    End With
    str说明 = "说明:" & Mid(str说明, 2)
    
    DrawCanvas = str说明
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Sub DrawPatiInfo(ByVal lngDC As Long, ByVal objDraw As Object, ByVal strPatiInfo As String, ByVal lngX As Long, ByVal lngY As Long, _
    ByVal lngLeft As Long, lngOutY As Long, Optional ByVal sngScale As Single = 1)
'-----------------------------------------------------------------------------------------------------------------------
'输出病人基本信息
'参数:lngDC 绘图对象的DC，strPatiInfo 病人信息组成字符串,分隔符为'(姓名:'年龄:'性别:'科别:'床号:'入院日期:'住院病历号)
'     lngX 左边距,lngY上边距,lngLeft 右边距(可以绘图的最大右边距)
'出参:lngOutY 返回绘图后的上边距
'-----------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer, k As Integer, l As Long
    Dim VarPatiInfo() As String
    Dim VarPatiName() As String
    Dim bln是否输出诊断 As Boolean, blnOne As Boolean
    Dim lngCurX As Long, lngCurY As Long, lngWidth As Long
    Dim strPatiName As String '病人信息内容标题,如 姓名,性别
    Dim Pname_SIZE() As SIZEL  '记录每个信息名称的坐标信息
    Dim Pinfo_SIZE() As SIZEL  '记录每个信息的坐标信息
    Dim arrSngY()
    Dim h_9t As Long
    Dim lngCurW As Long
    
    Dim sngW As Single, sngLen As Single
    Dim strText As String, strText1 As String, strText2 As String
    
    VarPatiInfo = Split(strPatiInfo, "'")
    bln是否输出诊断 = (UBound(VarPatiInfo) > 6)
    'strPatiName = "姓名:'性别:'年龄:'入院日期:'住院号:'科室:'床号:" & IIf(bln是否输出诊断 = True, "'诊断:", "")
    strPatiName = "姓名:'年龄:'性别:'科别:'床号:'入院日期:'住院病历号:" & IIf(bln是否输出诊断 = True, "'诊断:", "")
    VarPatiName = Split(strPatiName, "'")
    ReDim Preserve Pname_SIZE(UBound(VarPatiInfo))
    ReDim Preserve Pinfo_SIZE(UBound(VarPatiInfo))
    
    
    lngCurX = lngX: lngCurY = lngY
    
    lngWidth = IIf(lngLeft - lngCurX < 0, lngCurX - lngLeft, lngLeft - lngCurX)

    arrSngY = Array()
    
    ReDim Preserve arrSngY(UBound(arrSngY) + 1)
    arrSngY(UBound(arrSngY)) = lngCurY
    
    '开始计算坐标
    For i = 0 To UBound(VarPatiInfo)
        Call GetTextExtentPoint32(lngDC, VarPatiName(i), Len(VarPatiName(i)), T_Size) '获取字体的准高度，获取汉字字体的宽度不准，数字的准
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
                
                '记录每次换行前的Y轴坐标
                ReDim Preserve arrSngY(UBound(arrSngY) + 1)
                arrSngY(UBound(arrSngY)) = lngCurY
            End If
        End If
    Next i
    
    k = 0
    blnOne = False
    
    '重新整理输出的X坐标
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
    '输出病人文字信息
    For i = 0 To UBound(VarPatiInfo)
        Call SetTextColor(lngDC, RGB_BLACK)
        Call GetTextRect(objDraw, Val(Pname_SIZE(i).cx), Val(Pname_SIZE(i).cy), CStr(VarPatiName(i)), , , , sngScale)
        Call DrawText(lngDC, CStr(VarPatiName(i)), -1, T_LableRect, DT_CENTER)
        
        Call SetTextColor(lngDC, RGB_BLUE)
        
        '诊断内容如果过多在下一行显示剩余部分
        If UBound(VarPatiInfo) = 7 And i = UBound(VarPatiInfo) Then
            strText1 = ""
            strText = Replace(VarPatiInfo(i), vbCrLf, "")
            Do While True
                T_Size.W = objDraw.TextWidth(strText) / T_TwipsPerPixel.X
                strText2 = strText
                If T_Size.W + Val(Pinfo_SIZE(i).cx) - lngLeft > 0 Then
                    sngLen = Round((lngLeft - Val(Pinfo_SIZE(i).cx)) / T_Size.W, 2)
                    lngLen = Len(StrConv(strText, vbFromUnicode)) * sngLen
                    '将半角转为全角
                    strText = StrConv(strText, vbWide)
                    strText1 = StrConv(Mid(StrConv(strText, vbFromUnicode), lngLen + 1), vbUnicode)
                    strText = StrConv(Mid(StrConv(strText, vbFromUnicode), 1, lngLen), vbUnicode)
                    
                    '得到原始字符串的截取的长度
                    strText1 = Mid(strText2, Len(strText) + 1)
                    strText = Mid(strText2, 1, Len(strText))
                    Call GetTextExtentPoint32(lngDC, strText, Len(strText), T_Size)
                    Call GetTextRect(objDraw, Val(Pinfo_SIZE(i).cx), Val(Pinfo_SIZE(i).cy), CStr(strText), , , , sngScale)
                    Call DrawText(lngDC, CStr(strText), -1, T_LableRect, DT_CENTER)
                    T_Size.W = objDraw.TextWidth(strText) / T_TwipsPerPixel.X
                    Pinfo_SIZE(i).cx = Pinfo_SIZE(i).cx + T_Size.W
                    strText = strText1
                    T_Size.W = objDraw.TextWidth("字") / T_TwipsPerPixel.X
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
    '返回Y轴坐标
    lngOutY = lngCurY + h_9t
End Sub

Public Sub DrawUpTable(ByVal lngDC As Long, ByVal objDraw As Object, ByVal strTmpString As String, _
    ByVal lngX As Long, ByVal lngY As Long, ByVal lngLeft As Long, lngOutY As Long, Optional sngScale As Single)
'-----------------------------------------------------------------------------------------------------------------------
'输出一般项目栏信息（包括 住院日期,天数,手术后天数和时间栏）
'参数:lngDC 绘图对象的DC，strTmpString 有住院日期，天数 和术后天数组成的字符串
'     lngX 左边距,lngY上边距,lngLeft 右边距(可以绘图的最大右边距)
'出参:lngOutY 返回绘图后的上边距
'-----------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim ArrCode() As String
    Dim strTmp As String
    Dim arrTmpTime() As String '住院时间
    Dim arrTmpDay() As String  '住院天数
    Dim arrOptDay() As String '术后天数
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
    
    strDay = IIf(mintBaby = 0, "住院天数", "出生天数")
    
    ArrCode = Split(strTmpString, "||")
    strTmp = strTmpString & String(2 - UBound(ArrCode), "||")
    ArrCode = Split(strTmp, "||")
    arrOptDay = Split(ArrCode(2), "'")
    arrTmpTime = Split(ArrCode(0), "'")
    arrTmpDay = Split(ArrCode(1), "'")

    lngCurX = lngX: lngStartX = lngX
    lngCurY = lngY: lngStartY = lngY
    
    '开始画表格栏
    
    'X
    Call DrawLine(lngDC, lngCurX, lngCurY, lngLeft, lngCurY, PS_SOLID, intBold, RGB_BLACK): lngCurY = lngCurY + T_DrawClient.时间列单位
    Call DrawLine(lngDC, lngCurX, lngCurY, lngLeft, lngCurY, PS_SOLID, intFine, RGB_BLACK): lngCurY = lngCurY + T_DrawClient.时间列单位
    Call DrawLine(lngDC, lngCurX, lngCurY, lngLeft, lngCurY, PS_SOLID, intFine, RGB_BLACK): lngCurY = lngCurY + T_DrawClient.时间列单位
    Call DrawLine(lngDC, lngCurX, lngCurY, lngLeft, lngCurY, PS_SOLID, intFine, RGB_BLACK): lngCurY = lngCurY + T_DrawClient.时间列单位 + 6
    Call DrawLine(lngDC, lngCurX, lngCurY, lngLeft, lngCurY, PS_SOLID, intBold, RGB_BLACK)
    
    'Y
    Call DrawLine(lngDC, lngCurX, lngStartY, lngCurX, lngCurY, PS_SOLID, intBold, RGB_BLACK)
    lngCurX = T_DrawClient.刻度区域.Right

    Call DrawLine(lngDC, lngCurX, lngStartY, lngCurX, lngCurY, PS_SOLID, intBold, RGB_BLACK)

    For i = 0 To 6
        lngCurX = lngCurX + T_DrawClient.列单位 * 6
        Call DrawLine(lngDC, lngCurX, lngStartY, lngCurX, lngCurY, PS_SOLID, intBold, RGB_BLACK)
    Next i
    
    lngCurX = T_DrawClient.刻度区域.Right
    lngCurY = lngStartY + T_DrawClient.时间列单位 * 3
    '时间
    For i = 0 To 6
        lngCurX = T_DrawClient.刻度区域.Right + i * T_DrawClient.列单位 * 6
        For j = 1 To 5
            lngCurX = lngCurX + T_DrawClient.列单位
            Call DrawLine(lngDC, lngCurX, lngCurY, lngCurX, lngCurY + T_DrawClient.时间列单位 + 6, PS_SOLID, intFine, RGB_BLACK)
        Next j
    Next i
    
    '开始输出信息
    '日期信息
    lngCurY = lngStartY
    Call SetTextColor(lngDC, RGB_BLACK)
    Call GetTextExtentPoint32(lngDC, "日     期", Len("日     期"), T_Size)
    Call GetTextRect(objDraw, lngStartX, lngCurY + T_DrawClient.时间列单位 / 2, "日      期", T_DrawClient.刻度区域.Right - lngStartX, True, , sngScale)
    Call DrawText(lngDC, "日     期", -1, T_LableRect, DT_CENTER)
    lngCurX = T_DrawClient.刻度区域.Right
    For i = 0 To UBound(arrTmpTime)
        lngCurX = T_DrawClient.刻度区域.Right + i * 6 * T_DrawClient.列单位
        Call SetTextColor(lngDC, RGB_BLUE)
        Call GetTextExtentPoint32(lngDC, CStr(arrTmpTime(i)), Len(CStr(arrTmpTime(i))), T_Size)
        Call GetTextRect(objDraw, lngCurX, lngCurY + T_DrawClient.时间列单位 / 2, CStr(arrTmpTime(i)), T_DrawClient.列单位 * 6, True, , sngScale)
        Call DrawText(lngDC, CStr(arrTmpTime(i)), -1, T_LableRect, DT_CENTER)
    Next i
    
    lngCurY = lngStartY + T_DrawClient.时间列单位 * 1
    '住院天数
    Call SetTextColor(lngDC, RGB_BLACK)
    Call GetTextExtentPoint32(lngDC, strDay, Len(strDay), T_Size)
    Call GetTextRect(objDraw, lngStartX, lngCurY + T_DrawClient.时间列单位 / 2, strDay, T_DrawClient.刻度区域.Right - lngStartX, True, , sngScale)
    Call DrawText(lngDC, strDay, -1, T_LableRect, DT_CENTER)
    lngCurX = T_DrawClient.刻度区域.Right
    
    For i = 0 To UBound(arrTmpDay)
        lngCurX = T_DrawClient.刻度区域.Right + i * 6 * T_DrawClient.列单位
        Call SetTextColor(lngDC, RGB_BLUE)
        Call GetTextExtentPoint32(lngDC, CStr(arrTmpDay(i)), Len(CStr(arrTmpDay(i))), T_Size)
        Call GetTextRect(objDraw, lngCurX, lngCurY + T_DrawClient.时间列单位 / 2, CStr(arrTmpDay(i)), T_DrawClient.列单位 * 6, True, , sngScale)
        Call DrawText(lngDC, CStr(arrTmpDay(i)), -1, T_LableRect, DT_CENTER)
    Next i
    
    '术/娩后天数
    lngCurY = lngStartY + T_DrawClient.时间列单位 * 2
    Call SetTextColor(lngDC, RGB_BLACK)
    Call GetTextExtentPoint32(lngDC, "手术后天数", Len("手术后天数"), T_Size)
    Call GetTextRect(objDraw, lngStartX, lngCurY + T_DrawClient.时间列单位 / 2, "手术后天数", T_DrawClient.刻度区域.Right - lngStartX, True, , sngScale)
    Call DrawText(lngDC, "手术后天数", -1, T_LableRect, DT_CENTER)
    lngCurX = T_DrawClient.刻度区域.Right
    
    '51283,刘鹏飞,2012-07-11,手术天数颜色
    lngColor = Val(zldatabase.GetPara("手术天数显示颜色", glngSys, 1255, "255"))
    For i = 0 To UBound(arrOptDay)
        lngCurX = T_DrawClient.刻度区域.Right + i * 6 * T_DrawClient.列单位
        Call SetTextColor(lngDC, lngColor)
        Call GetTextExtentPoint32(lngDC, CStr(arrOptDay(i)), Len(CStr(arrOptDay(i))), T_Size)
        Call GetTextRect(objDraw, lngCurX, lngCurY + T_DrawClient.时间列单位 / 2, CStr(arrOptDay(i)), T_DrawClient.列单位 * 6, True, , sngScale)
        Call DrawText(lngDC, CStr(arrOptDay(i)), -1, T_LableRect, DT_CENTER)
    Next i
    lngColor = 0
    '时间
    lngCurY = lngStartY + T_DrawClient.时间列单位 * 3
    Call SetTextColor(lngDC, RGB_BLACK)
    Call GetTextExtentPoint32(lngDC, "时      间", Len("时      间"), T_Size)
    Call GetTextRect(objDraw, lngStartX, lngCurY + T_DrawClient.时间列单位 / 2, "时      间", T_DrawClient.刻度区域.Right - lngStartX, True, , sngScale)
    Call DrawText(lngDC, "时      间", -1, T_LableRect, DT_CENTER)
    lngCurX = T_DrawClient.刻度区域.Right
    
    For i = 0 To 6
        lngCurX = T_DrawClient.刻度区域.Right + i * 6 * T_DrawClient.列单位
        '输出上午下午时间
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
            lngTmpX = lngCurX + T_DrawClient.列单位 * j
            Call SetTextColor(lngDC, lngColor)
            Call GetTextExtentPoint32(lngDC, strTmp, Len(strTmp), T_Size)
            Call GetTextRect(objDraw, lngTmpX - 1, lngCurY + (T_DrawClient.时间列单位 + 6) / 2, strTmp, T_DrawClient.列单位, True, , sngScale)
            Call DrawText(lngDC, strTmp, -1, T_LableRect, DT_CENTER)
        Next j
    Next i
    lngOutY = lngStartY + T_DrawClient.时间列单位 * 4 + 6
End Sub

Public Sub SetFontIndirect(ByVal stdset As StdFont, ByVal lngDC As Long, ByVal objDraw As Object)

    '功能:设置字体属性
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
    '功能： 从时间计算出列
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim varTime As Variant

    Dim strTmp  As String

    Dim intDays As Integer

    Dim intLoop As Integer
    
    On Error GoTo Errhand
    
    GetCurveColumn = -1
    
    '初始化时间范围划分
    Call InitDateTimeRange(varTime, intHourBegin)

    '计算当前天的时间是在一天的第几格位置上
    strTmp = Format(dtDateTime, "HH:mm:ss")
    
    For intLoop = 0 To 6
        If Format(strTmp, "HH:mm:ss") >= Format(Split(varTime(intLoop), ",")(0), "HH:mm:ss") And Format(strTmp, "HH:mm:ss") <= Format(Split(varTime(intLoop), ",")(1), "HH:mm:ss") Then
            Exit For
        End If
    Next
    
    If intLoop < 7 Then
        '计算当天在当前体温单页上是第几天（0表示第一天；1表示第二天.....）
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
    '功能:根据列计算出时间范围
    '参数 intCol 当前列    dtBeginDateTime 起始时间
    '返回格式为:开始时间;终止时间
    '-------------------------------------------------------------------------------------
    Dim varTime  As Variant

    Dim intDays  As Integer

    Dim strBegin As String

    Dim strEnd   As String

    Dim lngLoop  As Long

    Dim lng列号  As Long

    On Error GoTo Errhand
    
    GetCurveDate = -1
    
    '初始化时间范围划分
    Call InitDateTimeRange(varTime, intHourBegin)
    
    '结算当前列和开始时间 相差的天数,并重新计算列的开始时间
    intDays = (intCOl - 1) \ 6
    strBegin = DateAdd("d", intDays, Int(dtBeginDateTime))
    strEnd = strBegin
    
    '结算列所在的时间范围
    lng列号 = (intCOl - 1) Mod 6
    
    strBegin = Format(strBegin & " " & Split(varTime(lng列号), ",")(0), "YYYY-MM-DD HH:mm:ss")
    strEnd = Format(strEnd & " " & Split(varTime(lng列号), ",")(1), "YYYY-MM-DD HH:mm:ss")

    GetCurveDate = strBegin & ";" & strEnd

    Exit Function

Errhand:

    If ErrCenter = 1 Then

        Resume

    End If

End Function

Public Function InitPara() As Boolean

    '******************************************************************************************************************
    '功能：得到所有本地参数
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop     As Integer

    Dim strTmp      As String

    Dim strTmpBegin As String

    Dim strTmpEnd   As String

    On Error GoTo Errhand
    
    gvarTime = Split(String(6, ";"), ";")
    gintHourBegin = zldatabase.GetPara("体温开始时间", glngSys, 1255, 4)
    strTmp = zldatabase.GetPara("体温标志分隔符", glngSys, 1255, 0)
    mlng体温不升显示方式 = Val(zldatabase.GetPara("体温不升显示方式", glngSys, 1255, "0"))
    If Val(strTmp) = 0 Then
        gstrCaveSplit = "——"
    ElseIf Val(strTmp) = 1 Then
        gstrCaveSplit = "于"
    Else
        gstrCaveSplit = ""
    End If
    
    '病人变动标记显示方法
    '------------------------------------------------------------------------------------------------------------------
    strTmp = zldatabase.GetPara("体温单标记", glngSys, 1255, "1;1;1;1;1;1;1;1")

    If UBound(Split(strTmp, ";")) >= 5 Then
        T_BodyFlag.入院 = Val(Split(strTmp, ";")(0))
        T_BodyFlag.入科 = Val(Split(strTmp, ";")(1))
        T_BodyFlag.转出 = Val(Split(strTmp, ";")(2))
        T_BodyFlag.换床 = Val(Split(strTmp, ";")(3))
        T_BodyFlag.手术 = Val(Split(strTmp, ";")(4))
        T_BodyFlag.出院 = Val(Split(strTmp, ";")(5))

        If UBound(Split(strTmp, ";")) >= 6 Then T_BodyFlag.分娩 = Val(Split(strTmp, ";")(6))
        If UBound(Split(strTmp, ";")) >= 7 Then T_BodyFlag.出生 = Val(Split(strTmp, ";")(7))
    End If
    
    '罗列体温单一天的曲线时间范围
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
    '功能：罗列体温单一天的曲线时间范围
    '参数：
    '返回：
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
'功能：检查体温单终止时间和开始时间是否在同一单元格，如果在同一单元格需要将终止时间移到下一单元格
'参数：strBegin 体温单开始时间,strEnd 体温单终止时间(病人出院时间)
'返回值：体温单终止时间
'**********************************************************************************
'需求：对于病人出院和入院时间在同一个格子，既要录入入院体温，也要录入出院体温，将出院体温录入到下一个格子。

    Dim varTime As Variant
    Dim intLoop As Integer, strTmp As String
    Dim intBegin As Integer, intEnd As Integer
    Dim strEnd As String
    
    RetrunEndTime = dtEnd
    If Format(dtBegin, "YYYY-MM-DD") <> Format(dtEnd, "YYYY-MM-DD") Then Exit Function
    '初始化时间范围划分
    Call InitDateTimeRange(varTime, intHourBegin)
    '1/计算开始时间和终止时间在第几个格子
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
    '2 不在同一列就退出
    If intBegin <> intEnd Then Exit Function
    If intEnd > 5 Then Exit Function
    '3 完成终止时间的重新赋值
    If intEnd > 4 Then
        strEnd = Format(DateAdd("D", 1, dtEnd), "YYYY-MM-DD") & " " & Format(Split(varTime(0), ",")(1), "HH:mm:ss")
    Else
        strEnd = Format(dtEnd, "YYYY-MM-DD") & " " & Format(Split(varTime(intEnd + 1), ",")(1), "HH:mm:ss")
    End If
    
    RetrunEndTime = CDate(Format(strEnd, "YYYY-MM-DD HH:mm:ss"))
End Function

Public Function GetGridItem(ByVal int护理等级 As Integer, ByVal byt适用病人 As Byte, ByVal lng科室ID As Long, Optional int项目性质 As Integer = 1) As ADODB.Recordset

    '**********************************************************************************
    '功能:提取体温表格项目
    '**********************************************************************************
    Dim rsTemp As New ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo Errhand
    
    '提取表格项目
    gstrSQL = "Select A.排列序号,A.项目序号,'' 体温部位,A.记录名,A.记录法,A.记录符,A.记录色,A.最大值,A.最小值,A.单位值,nvl(A.记录频次,2) 记录频次,A.入院首测,B.项目性质," & _
        "   B.分组名,B.项目值域,B.项目表示,B.项目类型,B.项目长度,B.项目小数,B.项目单位 单位" & _
        "   From 体温记录项目 A,护理记录项目 B,诊治所见项目 C" & _
        "   Where A.项目序号=B.项目序号 And B.项目ID=C.Id(+) And A.记录法=2 And nvl(B.项目性质,1)=[4]" & _
        "   And nvl(B.应用方式,0)=1 And nvl(B.护理等级,0)>=[1] And nvl(B.适用病人,0) In (0,[2])" & _
        "   And (B.适用科室=1 Or (B.适用科室=2 And Exists (Select 1 From 护理适用科室 D Where D.项目序号=B.项目序号 And D.科室id=[3]))) order by Decode(项目序号,3 ,0,1 ),排列序号"
        
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "提取固定体温表格项目", int护理等级, byt适用病人, lng科室ID, int项目性质)
    Set GetGridItem = rsTemp

    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetAppendGridItem(ByVal lng文件ID As Long, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal int护理等级 As Integer, ByVal int婴儿 As Long, dt开始时间 As Date, dt结束时间 As Date, ByVal byt适用病人 As Byte, ByVal lng科室ID As Long, Optional blnMove As Boolean = False) As ADODB.Recordset
    '**************************************************************************
    '功能:提取活动有数据的体温表格项目以及固定表格项目
    '**************************************************************************
    Dim rsTemp As New ADODB.Recordset
    Dim strSql As String

    On Error GoTo Errhand
    
    Set rsTemp = GetGridItem(int护理等级, byt适用病人, lng科室ID, 2)
    If rsTemp.RecordCount = 0 Then
        '不存在活动项目直接提取固定表格项目
        Set rsTemp = GetGridItem(int护理等级, byt适用病人, lng科室ID, 1)
        Set GetAppendGridItem = rsTemp
        Exit Function
    End If
    With rsTemp
        Do While Not .EOF
            strSql = IIf(strSql = "", "select " & !项目序号 & " 项目序号 from dual", strSql & " UNION ALL select " & !项目序号 & "  项目序号 from dual ")
            .MoveNext
        Loop
    End With
    
    strSql = "(" & strSql & ") F"
    '提取活动项目
    gstrSQL = "Select distinct D.排列序号,D.项目序号,C.体温部位,C.体温部位 || D.记录名  记录名,D.记录法,D.记录符,D.记录色,D.最大值,D.最小值,D.单位值,nvl(D.记录频次,2) 记录频次,D.入院首测," & _
        "   E.项目性质,E.分组名,E.项目值域,E.项目表示,E.项目类型,E.项目长度,E.项目小数,E.项目单位 单位" & _
        "   FROM 病人护理文件 B, 病人护理数据 A,病人护理明细 C,体温记录项目 D,护理记录项目 E," & strSql & _
        "   Where  B.ID=A.文件ID And A.ID = c.记录ID  AND B.ID=[1]  AND Nvl(B.婴儿,0)=[5]  AND B.病人id=[2]    AND B.主页id=[3] AND d.项目序号=C.项目序号 " & _
        "   AND c.记录类型=1 And E.项目性质=2  AND E.项目序号=D.项目序号  AND E.护理等级>=[4]   AND a.发生时间 BETWEEN [6] And [7] And c.终止版本 Is Null " & _
        "   AND d.记录法=2 and D.项目序号=F.项目序号"
    
    '提取固定表格项目
    strSql = "Select A.排列序号,A.项目序号,'' 体温部位,A.记录名,A.记录法,A.记录符,A.记录色,A.最大值,A.最小值,A.单位值,nvl(A.记录频次,2) 记录频次,A.入院首测,B.项目性质," & _
        "   B.分组名,B.项目值域,B.项目表示,B.项目类型,B.项目长度,B.项目小数,B.项目单位 单位" & _
        "   From 体温记录项目 A,护理记录项目 B,诊治所见项目 C" & _
        "   Where A.项目序号=B.项目序号 And B.项目ID=C.Id(+) And A.记录法=2 And nvl(B.项目性质,1)=1" & _
        "   And nvl(B.应用方式,0)=1 And nvl(B.护理等级,0)>=[4] And nvl(B.适用病人,0) In (0,[8])" & _
        "   And (B.适用科室=1 Or (B.适用科室=2 And Exists (Select 1 From 护理适用科室 D Where D.项目序号=B.项目序号 And D.科室id=[9])))"
    
    gstrSQL = "Select 排列序号,项目序号,体温部位,记录名,记录法,记录符,记录色,最大值,最小值,单位值,记录频次,入院首测,项目性质," & _
        "   分组名,项目值域,项目表示,项目类型,项目长度,项目小数,单位" & _
        "   From (" & gstrSQL & vbCrLf & " UNION ALL " & vbCrLf & strSql & ") order by Decode(项目序号,3 ,0,1 ),排列序号,记录名"
    If blnMove Then
        gstrSQL = Replace(gstrSQL, "病人护理文件", "H病人护理文件")
        gstrSQL = Replace(gstrSQL, "病人护理数据", "H病人护理数据")
        gstrSQL = Replace(gstrSQL, "病人护理明细", "H病人护理明细")
    End If

    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "", lng文件ID, lng病人ID, lng主页ID, int护理等级, int婴儿, dt开始时间, dt结束时间, byt适用病人, lng科室ID)
    
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

    '在(X,Y)处输出Text文本
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
    '设置文本颜色
    Call SetTextColor(lngDC, ForeColor)

    '正常输出字体
    If Asc(strText) < 0 And strText <> "—" Then
    
        Call GetTextRect(objDraw, X, Y, strText, T_DrawClient.列单位, False, , sglScale)
        Call DrawText(lngDC, strText, -1, T_LableRect, DT_CENTER)
        '反转90度输出字体
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
'        Call GetTextRect(objDraw, X, Y, strText, T_DrawClient.列单位, False, , sglScale)
'        X1 = T_LableRect.Right - T_LableRect.Left + (T_LableRect.Left - X) / 2
        Call TextOut(lngDC, X + T_DrawClient.列单位, Y, strText, LenB(StrConv(strText, vbFromUnicode)))
         
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
    End If
End Sub

Public Sub GetTextRect(ByVal objDraw As Object, ByVal lngX As Long, ByVal lngY As Long, ByVal strInput As String, _
    Optional ByVal lngWidth As Long = 0, Optional bln居中 As Boolean = True, Optional ByVal lngHeght As Long = 0, Optional ByVal sngScale As Single = 1)
    
    Dim lngInputW As Long, lng1H As Long
    Dim sngSize As Single
        
    T_LableRect.Left = lngX + 1 '避免与左边界划线重合
    
    If bln居中 = True Then
        T_LableRect.Top = lngY - objDraw.TextHeight("1") / 2 / T_TwipsPerPixel.Y
    Else
        T_LableRect.Top = lngY
    End If
    
    T_LableRect.Right = objDraw.TextWidth(strInput) / T_TwipsPerPixel.Y + T_LableRect.Left + 2
    T_LableRect.Bottom = objDraw.TextHeight("1") / T_TwipsPerPixel.Y + T_LableRect.Top + 2
    
    If lngWidth <> 0 Then
        '将文本显示在所示宽度的中间区域
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
    '创建新画笔进行划线
    
    If msngTwips = 0 Then msngTwips = 1
    sngX = 2 * msngTwips
    sngY = 3 * msngTwips

    lngPen = CreatePen(lngType, intWidth, lngRGB)
    lngOldPen = SelectObject(lngDC, lngPen)
    '绘图
    Call MoveToEx(lngDC, lngSX, lngSY, T_OldPoint)
    Call LineTo(lngDC, lngDX, lngDY)
    '对于物理降温话上下箭头
    If blnEndRow Then
        If lngSY > lngDY Then '物理降温失败（向上箭头）
            For X = lngSX - sngX To lngSX + sngX
                Call MoveToEx(lngDC, X, lngSY - sngY, T_OldPoint)
                Call LineTo(lngDC, lngSX, lngSY)
            Next X
        Else '物理降温成功(向下箭头)
            For X = lngSX - sngX To lngSX + sngX
                Call MoveToEx(lngDC, X, lngSY + sngY, T_OldPoint)
                Call LineTo(lngDC, lngSX, lngSY)
            Next X
        End If
    End If
    
    '还原画笔并销毁
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
    '创建新画笔进行画一个矩形
    
    lngPen = CreatePen(lngType, intWidth, lngRGB)
    lngOldPen = SelectObject(lngDC, lngPen)
    '绘图
    Call Rectangle(lngDC, lngSX, lngSY, lngDX, lngDY)
    '还原画笔并销毁
    Call SelectObject(lngDC, lngOldPen)
    Call DeleteObject(lngPen)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub OutPutText(ByVal objDraw As Object, ByVal rsDrawItems As ADODB.Recordset, ByVal lngDC As Long, ByVal rsNote As ADODB.Recordset, ByVal mstrBeginDate As String, Optional ByVal sngScale As Single = 1)

    'rsDrawItems  记录项目的最大坐标 单位值等基本信息
    'rsNote 所有说的信息
    'mstrBeginDate 体温单每页开始时间
    '输出以下信息:入院,入科,转科,出院,手术分娩,未记说明,上标说明及出生
    '未记说明及上标说明,在没有入出转手术分娩及出生的信息时,打印在42-40之间;否则从40开始向下打印
    '除未记说明及上标说明外,入出转等信息当一个刻度发生多个时,依次写入各个刻度中,如其它刻度也有信息,顺移
    Dim lngMaxX As Long     '体温单最大X坐标
    Dim lngX    As Long '第一列的X坐标
    Dim lngY    As Long 'Y坐标
    Dim lngY1   As Long '40 度固定坐标
    Dim i       As Integer
    Dim X, Y As Long '输出内容时的坐标
    Dim strComment    As String, strText As String
    Dim intAscCharNum As Integer
    Dim rsTemp  As New ADODB.Recordset
    Dim strDate As String
    Dim bln上标 As Boolean
    Dim bln事件显示规则 As Boolean
    
    On Error GoTo Errhand
    
    bln事件显示规则 = (Val(zldatabase.GetPara("体温标志按顺序当天排列", glngSys, 1255, 0)) = 1)
    
    lngMaxX = T_DrawClient.体温区域.Right - T_DrawClient.列单位
    
    rsNote.Filter = "禁用<>1"

    '首先检查更新入出转，手术分娩信息
    If rsNote.RecordCount = 0 Then Exit Sub
    
    rsNote.Sort = "X坐标,时间,项目序号"
    lngX = rsNote!X坐标
    
    With rsNote
        Do While Not .EOF
            If Trim(zlCommFun.Nvl(!内容)) <> "" Then
                If Not (!类型 = 2 Or !类型 = 99) Then
                    
                    '体温标志按顺序当天排列
                    If bln事件显示规则 = True Then
                        If lngX <= lngMaxX Then
                            strDate = Format(Split(GetXCoordinate(lngX, mstrBeginDate, False), ",")(0), "YYYY-MM-DD")
                            If CDate(strDate) > CDate(Format(!时间, "YYYY-MM-DD")) Then
                                lngX = Val(!X坐标)
                                !禁用 = 1
                            End If
                        Else
                            lngX = lngMaxX
                            !禁用 = 1
                        End If
                    Else
                        '控制x坐标，如果超过体温最大x坐标，则进行校正
                        If lngX > lngMaxX Then lngX = lngMaxX
                    End If
                    
                    !打印X坐标 = IIf(lngX <= Val(!X坐标), !X坐标, lngX)
                    !高度 = GetFontHeight(lngDC, zlCommFun.Nvl(!内容))
                    .Update
                    
                    If lngX <= !X坐标 Then lngX = !X坐标
                    lngX = lngX + T_DrawClient.列单位
                Else
                    !高度 = GetFontHeight(lngDC, zlCommFun.Nvl(!内容))
                    .Update
                End If
            End If
            .MoveNext
        Loop
        
        If .RecordCount > 0 Then .MoveFirst
        lngY = GetYCoordinate(objDraw, rsDrawItems, gint体温, 42)
        '调整入出转 手术，分娩到达最大X坐标有多列式的Y坐标
        .Filter = "打印X坐标=" & lngMaxX & " And 禁用<>1"
        .Sort = "时间,项目序号"

        Do While Not .EOF
            !Y坐标 = lngY
            .Update
            lngY = lngY + Val(!高度) + T_DrawClient.行单位
            .MoveNext
        Loop
        
        .Filter = "禁用<>1"
        .MoveFirst
        
        '更新未记说明，上标的显示位置(Y坐标).
        '说明:在没有入出转，手术信息的情况下 打印在 42-40度之间，否则打印在40度以下打印
        .Sort = "X坐标,时间,项目序号"
        Set rsTemp = .Clone

        Do While Not .EOF
            lngY = 0
            If (!类型 = 2 Or !类型 = 99) Then
                bln上标 = False
                Set rsTemp = .Clone
                
                rsTemp.Filter = "(打印X坐标=" & !X坐标 & " and 类型=99) or (打印X坐标=" & !X坐标 & " and 类型=2)"
                
                If rsTemp.BOF Then
                    rsTemp.Filter = "打印X坐标=" & !X坐标
                End If
                
                If rsTemp.RecordCount > 0 Then
                    lngY = Val(rsTemp!Y坐标)
                    
                    Do While Not rsTemp.EOF
                        If bln上标 = False Then
                            bln上标 = IIf(rsTemp!类型 = 2 Or rsTemp!类型 = 99, True, False)

                            If bln上标 = True Then lngY = Val(rsTemp!Y坐标)
                        End If
                        
                        lngY = lngY + rsTemp!高度 + T_DrawClient.行单位
                        rsTemp.MoveNext
                    Loop
                    
                    lngY1 = GetYCoordinate(objDraw, rsDrawItems, gint体温, 40)

                    If lngY > lngY1 Or bln上标 Then lngY1 = lngY
                    
                Else '不存在任何信息 从42开始打印
                    lngY1 = Val(!Y坐标)
                End If
                
                !Y坐标 = lngY1
                !打印X坐标 = !X坐标
                .Update
            End If

            .MoveNext
        Loop
        .MoveFirst
        
        Dim sigNum As Single
        Do While Not .EOF
            '输出内容
            strComment = Trim(zlCommFun.Nvl(!内容))

            If strComment <> "" Then
                X = Val(IIf(Trim(!打印X坐标) <> "", !打印X坐标, !X坐标))
                Y = Val(!Y坐标)
                intAscCharNum = 0

                For i = 1 To Len(strComment)
                    If Y < T_DrawClient.刻度区域.Bottom Then
                        strText = Mid(strComment, i, 1)
                        Call GetTextExtentPoint32(lngDC, strText, Len(strText), T_Size)

                        If Asc(strText) < 0 Then
                            If intAscCharNum Mod 2 = 1 Then Y = Y + T_Size.H / 2
                            '根据坐标得到数值
                            sigNum = GetYCoordinate(objDraw, rsDrawItems, gint体温, X & "," & Y, False)
                            Y = GetYCoordinate(objDraw, rsDrawItems, gint体温, sigNum)
                        End If

                        '输出字体信息
                        Call DrawRotateText(objDraw, lngDC, X, Y, strText, !颜色, sngScale)

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
    Dim bln脉搏短轴填充方式 As Boolean

    'lngStart:指定心率开始的索引,用于区域连线,避免把脉搏的连线覆盖掉了(颜色可能不一样)
    On Error GoTo Errhand
    
    bln脉搏短轴填充方式 = Val(zldatabase.GetPara("脉搏短绌填充方式", glngSys, 1255, "0")) = 1
    '填充区域并划边线
    
    '创建系统刷子
    If bln脉搏短轴填充方式 = True Then
        lngBrush = CreateHatchBrush(HS_VERTICAL, RGB_RED)
    Else
        lngBrush = CreateHatchBrush(HS_BDIAGONAL, RGB_RED)
    End If
    '如果创建刷子成功,才选入
    If lngBrush <> 0 Then
        lngRgn = CreatePolygonRgn(PtInPoly(1), UBound(PtInPoly), ALTERNATE)
        FillRgn lngDC, lngRgn, lngBrush
        Call DeleteObject(lngRgn)
        Call DeleteObject(lngBrush)
    
        lngPen = CreatePen(PS_SOLID, 1, RGB_RED)
        lngOldPen = SelectObject(lngDC, lngPen)
        '绘图
        Polyline lngDC, PtInPoly(lngStart), UBound(PtInPoly) - lngStart
        '还原画笔并销毁
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
    '功能:得到字符串高度
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

Public Function GetDataFromHis(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lng婴儿 As Long, ByVal dtFrom As Date, ByVal dtTo As Date, Optional ByVal bytMode As Byte = 1) As ADODB.Recordset

    '******************************************************************************************************************
    '功能：从医嘱记录提取手术、分娩数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strSql As String
    Dim strNewSql As String
    Dim rsTemp As New ADODB.Recordset
    Dim RS As New ADODB.Recordset
    Dim rsCopy As New ADODB.Recordset
    Dim strFileds As String, strValue As String
    Dim lng诊疗项目id As Long
    Dim blnBody As Boolean
    
    On Error GoTo Errhand
    
    blnBody = False
    Select Case bytMode

            '------------------------------------------------------------------------------------------------------------------
        Case 1              '从医嘱记录提取手术、分娩数据
        
            '        dtFrom = dtFrom - 14
        
            strSql = "Select 执行时间,内容,次弟" & vbNewLine & _
                " From (Select 执行时间,内容, Rownum As 次弟" & vbNewLine & _
                "       From (Select Distinct C.执行时间,'手术' As 内容 " & vbNewLine & _
                "              From 病人医嘱记录 A, 诊疗项目目录 B, 病人医嘱执行 C" & vbNewLine & _
                "              Where A.病人id = [1] And A.主页id = [2] And Nvl(A.婴儿, 0) = [3] And A.医嘱期效 = 1 And A.诊疗项目id = B.ID And" & vbNewLine & _
                "                    A.诊疗类别 = 'F' And A.医嘱状态 = 8 And C.医嘱id = A.ID And C.执行时间 < =[5] " & vbNewLine & _
                "               Union All Select a.出生时间 As 执行时间,'分娩' As 内容 From 病人新生儿记录 a Where a.病人id=[1] And a.主页id=[2] And a.出生时间 Is Not Null And RowNum<2) " & _
                "       Order By 执行时间)" & vbNewLine & "Where 执行时间 >= [4] And 次弟 <= 12 " & vbNewLine & "Order By 执行时间 "
                
            Set GetDataFromHis = zldatabase.OpenSQLRecord(strSql, "体温单", lng病人ID, lng主页ID, lng婴儿, dtFrom, dtTo)

            '------------------------------------------------------------------------------------------------------------------
        Case 2              '入出转标志(入院,出院,转科,换床)
            strFileds = "科室," & adLongVarChar & ",50|时间," & adDate & ",20|内容," & adLongVarChar & ",100|行号," & adDouble & ",5"
            Call Record_Init(rsCopy, strFileds)
            strFileds = "科室|时间|内容|行号"
            '1-入院；2-入科；3-转科；4-换床；5-床位等级变动；6-护理等级变动；7-经治医师改变；8-责任护士改变,9-留观病人转住院,10-病人预出院,11-主治医师变动,12-主任医师变动,13-病情变动
            
            '0-普通;1-留观;2-住院;3-转科;4-术后;5-出院;6-转院;7-会诊;8-抢救;9-病重;10-病危;11-死亡;12-记录入出量;14-术前
            '提取死亡记录ID
'            strSQL = "Select ID From 诊疗项目目录 Where 类别='Z' And 操作类型='11' "
'            Set RS = zlDatabase.OpenSQLRecord(strSQL, "体温单")
'
'            If RS.BOF = False Then lng诊疗项目id = zlCommFun.Nvl(RS("ID").Value)
        
            strSql = _
               "    Select 科室,时间,内容,行号 From (" & vbNewLine & _
               "    Select b.名称 As 科室,开始时间 As 时间, Decode(开始原因, 2,'入科',3, '转入',4,'换床'||Decode(床号,Null,'','('||床号||')')) As 内容,Decode(开始原因,2,9,3,6,4,7) As 行号 " & vbNewLine & _
               "    From 病人变动记录 A,部门表 b" & vbNewLine & _
               "    Where b.id(+)=a.科室id and a.开始原因 In (2,3,4) And A.病人id = [1] And A.主页id = [2]  And [3]=0 And A.开始时间 Between [4] And [5] " & vbNewLine & _
               "    Union " & vbNewLine & _
               "    Select 科室,时间,内容,行号 From (Select * From (Select  b.名称 As 科室,A.开始时间 As 时间, '入科' As 内容,9 As 行号 " & vbNewLine & _
               "    From 病人变动记录 A,部门表 B" & vbNewLine & _
               "    Where b.id(+)=a.科室id And a.开始原因=1 And A.病人id = [1] And A.主页id = [2] And [3]=0 And NOT Exists " & vbNewLine & _
               "   (Select ID From 病人变动记录 C Where C.开始原因=2 And C.病人ID=A.病人ID And C.主页ID=A.主页ID And [3]=0) Order By a.开始时间) Where RowNum=1) Where 时间 Between [4] And [5] " & vbNewLine & _
               "    )" & vbNewLine & _
               "    Union All" & vbNewLine & _
               "    Select '' As 科室,时间,内容,行号 From (Select * From (Select 开始时间 As 时间, '入院' As 内容,5 As 行号 " & vbNewLine & _
               "    From 病人变动记录 A" & vbNewLine & _
               "    Where a.开始原因=1 And A.病人id = [1] And A.主页id = [2] And [3]=0 Order By a.开始时间) Where RowNum=1) Where 时间 Between [4] And [5] " & vbNewLine & _
               "    Union All" & vbNewLine & _
               "    Select '' As 科室,Nvl(b.开始执行时间,a.出院日期) As 时间, Decode(出院方式, '正常', '出院', 出院方式) As 内容,8 As 行号 " & vbNewLine & _
               "    From 病案主页 A,(Select x.病人id,x.主页id,Max(x.开始执行时间) As 开始执行时间 From 病人医嘱记录 x,诊疗项目目录 z Where x.病人id=[1] And x.主页id=[2] " & vbNewLine & _
               "    And x.诊疗项目id+0=z.ID And x.医嘱状态 in (3,8) And z.类别='Z' And z.操作类型='11' Group By x.病人id,x.主页id) B " & vbNewLine & _
               "    Where A.病人id = [1] And A.主页id = [2] And A.出院日期 Between [4] And [5] And a.病人id=b.病人id(+) And a.主页id=b.主页id(+) "
        
            strSql = "Select * From (" & strSql & ") Order By 时间,行号 "
            
            Set RS = zldatabase.OpenSQLRecord(strSql, "体温单", lng病人ID, lng主页ID, lng婴儿, dtFrom, dtTo)
            
            Do While Not RS.EOF
                strValue = Nvl(RS!科室) & "|" & CDate(RS!时间) & "|" & Nvl(RS!内容) & "|" & Val(Nvl(RS!行号))
                Call Record_Add(rsCopy, strFileds, strValue)
            RS.MoveNext
            Loop
                    
            If lng婴儿 <> 0 Then
                '提取婴儿医嘱记录(转科,出院,死亡)
                strNewSql = _
                    "   select /*+ RULE */ 科室,开始执行时间,decode(操作类型,3,'转出',5,'出院','死亡') 内容,Decode(操作类型,'3',3,8) 行号 From (" & vbNewLine & _
                    "   select D.名称 科室,B.开始执行时间,C.操作类型 " & vbNewLine & _
                    "   from 病案主页 A,病人医嘱记录 B,诊疗项目目录 C,部门表 D" & vbNewLine & _
                    "   where A.病人ID=[1] and A.主页ID=[2] And  A.病人ID=B.病人ID(+) And A.主页ID=B.主页ID(+) And B.婴儿(+)=[3]" & vbNewLine & _
                    "   and B.诊疗项目id+0=C.ID  And B.医嘱状态=8  and C.类别='Z' And   B.执行科室ID=D.ID(+)" & vbNewLine & _
                    "   and  exists (select 1 from Table(Cast(f_str2list('3,5,11') As zlTools.t_strlist)) where C.操作类型=COLUMN_VALUE) order by B.开始执行时间 DESC" & vbNewLine & _
                    "   ) where Rownum<2"

                Set rsTemp = zldatabase.OpenSQLRecord(strNewSql, "体温单", lng病人ID, lng主页ID, lng婴儿)
                blnBody = (rsTemp.RecordCount > 0)
                
                '如果发现婴儿存在转科，出院医嘱信息。需要更新母亲信息
                If blnBody = True Then
                    rsCopy.Filter = "时间>='" & CDate(rsTemp!开始执行时间) & "'"
                    Do While Not rsCopy.EOF
                        rsCopy.Delete
                        rsCopy.Update
                    rsCopy.MoveNext
                    Loop
                    '删除母亲本人的出院信息
                    rsCopy.Filter = "行号=8"
                    Do While Not rsCopy.EOF
                        rsCopy.Delete
                        rsCopy.Update
                    rsCopy.MoveNext
                    Loop
                    '添加婴儿医嘱信息
                    rsTemp.MoveFirst
                    If CDate(Format(rsTemp!开始执行时间, "YYYY-MM-DD HH:mm:ss")) >= CDate(Format(dtFrom, "YYYY-MM-DD HH:mm:ss")) And CDate(rsTemp!开始执行时间) <= CDate(Format(dtTo, "YYYY-MM-DD HH:mm:ss")) Then
                        strValue = Nvl(rsTemp!科室) & "|" & CDate(rsTemp!开始执行时间) & "|" & Nvl(rsTemp!内容) & "|" & Val(Nvl(rsTemp!行号))
                        Call Record_Add(rsCopy, strFileds, strValue)
                    End If
                End If
            End If
            
            rsCopy.Filter = 0
            'Call OutputRsData(rsCopy, True)
            Set GetDataFromHis = rsCopy

            '------------------------------------------------------------------------------------------------------------------
        Case 3              '从新生儿记录中提出生/分娩日期
        
            strSql = _
                "   Select '' As 科室,a.出生时间 As 时间,'出生' As 内容,13 As 行号 From 病人新生儿记录 a " & _
                "   Where a.病人id=[1] And a.主页id=[2] And a.序号=[3] And a.出生时间 Between [4] And [5]"
            Set GetDataFromHis = zldatabase.OpenSQLRecord(strSql, "体温单", lng病人ID, lng主页ID, lng婴儿, dtFrom, dtTo)
    End Select
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function CheckFileBack(ByVal lngID As Long, ByVal blnMove As Boolean) As Boolean
'---------------------------------------------------------------
'功能:检查文件是否归档
'---------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSql As String
    On Error GoTo Errhand
    
    CheckFileBack = False
    strSql = "Select 1 From 病人护理文件 Where Id=[1] And 归档人 Is Not Null"
    Set rsTemp = zldatabase.OpenSQLRecord(strSql, "检查文件是否归档", lngID)
    If blnMove = True Then
        strSql = Replace(strSql, "病人护理文件", "H病人护理文件")
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
    '功能：转换时间为汉字 如 22:59 转换为二十二时五十九分
    '参数：时间 格式为 Format(strtime,"HH:mm")
    '返回：
    '------------------------------------------------------------------------------------------------------------------
    Dim strTmp1 As String

    Dim strTmp2 As String
    
    strTime = Format(strTime, "HH:mm")

    If InStr(strTime, ":") <= 0 Then Exit Function

    On Error GoTo Errhand
    
    strTmp1 = Left(strTime, InStr(strTime, ":") - 1)
    strTmp2 = Mid(strTime, InStr(strTime, ":") + 1)
    
    strTmp1 = Switch(strTmp1 = "00", "零", strTmp1 = "01", "一", strTmp1 = "02", "二", strTmp1 = "03", "三", strTmp1 = "04", "四", strTmp1 = "05", "五", strTmp1 = "06", "六", strTmp1 = "07", "七", strTmp1 = "08", "八", strTmp1 = "09", "九", strTmp1 = "10", "十", strTmp1 = "11", "十一", strTmp1 = "12", "十二", strTmp1 = "13", "十三", strTmp1 = "14", "十四", strTmp1 = "15", "十五", strTmp1 = "16", "十六", strTmp1 = "17", "十七", strTmp1 = "18", "十八", strTmp1 = "19", "十九", strTmp1 = "20", "二十", strTmp1 = "21", "二十一", strTmp1 = "22", "二十二", strTmp1 = "23", "二十三")
    
    strTmp2 = Switch(strTmp2 = "00", "零", strTmp2 = "01", "一", strTmp2 = "02", "二", strTmp2 = "03", "三", _
       strTmp2 = "04", "四", strTmp2 = "05", "五", strTmp2 = "06", "六", strTmp2 = "07", "七", _
       strTmp2 = "08", "八", strTmp2 = "09", "九", strTmp2 = "10", "十", strTmp2 = "11", "十一", _
       strTmp2 = "12", "十二", strTmp2 = "13", "十三", strTmp2 = "14", "十四", strTmp2 = "15", "十五", _
       strTmp2 = "16", "十六", strTmp2 = "17", "十七", strTmp2 = "18", "十八", strTmp2 = "19", "十九", _
       strTmp2 = "20", "二十", strTmp2 = "21", "二十一", strTmp2 = "22", "二十二", strTmp2 = "23", "二十三", _
       strTmp2 = "24", "二十四", strTmp2 = "25", "二十五", strTmp2 = "26", "二十六", strTmp2 = "27", "二十七", _
       strTmp2 = "28", "二十八", strTmp2 = "29", "二十九", strTmp2 = "30", "三十", strTmp2 = "31", "三十一", _
       strTmp2 = "32", "三十二", strTmp2 = "33", "三十三", strTmp2 = "34", "三十四", strTmp2 = "35", "三十五", _
       strTmp2 = "36", "三十六", strTmp2 = "37", "三十七", strTmp2 = "38", "三十八", strTmp2 = "39", "三十九", _
       strTmp2 = "40", "四十", strTmp2 = "41", "四十一", strTmp2 = "42", "四十二", strTmp2 = "43", "四十三", _
       strTmp2 = "44", "四十四", strTmp2 = "45", "四十五", strTmp2 = "46", "四十六", strTmp2 = "47", "四十七", _
       strTmp2 = "48", "四十八", strTmp2 = "49", "四十九", strTmp2 = "50", "五十", strTmp2 = "51", "五十一", _
       strTmp2 = "52", "五十二", strTmp2 = "53", "五十三", strTmp2 = "54", "五十四", strTmp2 = "55", "五十五", _
       strTmp2 = "56", "五十六", strTmp2 = "57", "五十七", strTmp2 = "58", "五十八", strTmp2 = "59", "五十九")
                    
    ConvertTimeToChinese = strTmp1 & "时" & strTmp2 & "分"
    
    Exit Function

Errhand:

    If ErrCenter = 1 Then

        Resume

    End If

End Function

Public Function DrawPicture(objDraw As Object, ByVal strFile As String, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, Optional ByVal bln资源 As Boolean = False) As Boolean

    '******************************************************************************************************************
    '功能：根据相册大小自动等比例缩放相片文件
    '参数：缩放前的相片文件
    '返回：缩放后的相片文件
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
    
    If bln资源 Then
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

Public Sub CreatePoly(rsPoint As ADODB.Recordset, ByVal objDraw As Object, ByVal lngDC As Long, ByVal strBeginDate As String, ByVal str心率坐标 As String)

'rsPoint 记录集 必须包括  项目序号,X坐标,Y坐标
    Dim arrData, arrPt

    Dim bln区域 As Boolean      '不是区域就是点对点,心率必须对应脉搏才能形成区域或连线

    Dim bln左 As Boolean, bln右 As Boolean, bln当前 As Boolean, bln断开 As Boolean, bln有效 As Boolean

    Dim intDO   As Integer, intMax As Integer             'intLast记录最后一个有效的心率

    Dim recttmp As RECT, SinX As Single, sinY As Single, sin左X As Single, sin右X As Single
    
    Dim str当前 As String, str左 As String, str右 As String

    Dim str脉搏 As String, str心率 As String

    Dim PtInPoly() As POINTAPI, intCOl As Integer, intCols As Integer, intCount As Integer
    Dim blnPrinter As Boolean
    Dim intBold As Integer, intFine As Integer

    On Error GoTo Errhand

    '1个心率对应1至3个脉搏,脉搏必须在每一天都有值,否则不形成区域
    '形成的区域集合必须是连续的,所以,先装入脉搏,再倒起装入心率,形成完整的一个区域
    '由点组成的封闭区域,在DrawPoly中完成封闭区域的连线
    
    If TypeName(objDraw) = "Printer" Then
        intBold = 4
        intFine = 4
        blnPrinter = True
    Else
        intBold = 2
        intFine = 1
        blnPrinter = False
    End If
    
    rsPoint.Sort = "项目序号,时间"
    arrData = Split(str心率坐标, ",")
    intMax = UBound(arrData)
    
    For intDO = 0 To intMax

        SinX = Val(Split(arrData(intDO), ";")(0))
        sinY = Val(Split(arrData(intDO), ";")(1))
        '将当前心率加入区域集合
        intCount = intCount + 1
        ReDim Preserve PtInPoly(intCount)
        str心率 = str心率 & "," & SinX + T_DrawClient.列单位 / 2 & ";" & sinY
        
        '如果左边有,则与左列的脉搏连线
        If Not bln区域 Then
            bln左 = False
            rsPoint.Filter = "项目序号=" & gint脉搏 & " And X坐标<" & Val(Split(arrData(intDO), ";")(0))
            
            If rsPoint.RecordCount <> 0 Then
               rsPoint.Sort = "X坐标 DESC"
                bln断开 = (rsPoint!断开 = 1)
                If Not bln断开 Then
                    rsPoint.Sort = "X坐标 DESC"
                    sin左X = rsPoint!X坐标
                
                    '根据当前坐标获取时间
                    str左 = GetXCoordinate(sin左X, strBeginDate, False)
                    str当前 = GetXCoordinate(Val(Split(arrData(intDO), ";")(0)), strBeginDate, False)
                    '当前点和前一时间点间隔一天没有数据就断开
                    If DateDiff("d", CDate(Split(str左, ",")(0)), CDate(Split(str当前, ",")(0))) < 2 Then
                        recttmp.Left = rsPoint!X坐标
                        recttmp.Top = rsPoint!Y坐标
                        '将左脉搏加入区域集合
                        intCount = intCount + 1
                        ReDim Preserve PtInPoly(intCount)
                        str脉搏 = str脉搏 & "," & rsPoint!X坐标 + T_DrawClient.列单位 / 2 & ";" & rsPoint!Y坐标
                        bln左 = True
                    End If
                End If
            End If
        End If
        
        bln当前 = False
        '缺省是和当前列的脉搏连线
        rsPoint.Filter = "项目序号=" & gint脉搏 & " And X坐标=" & Val(Split(arrData(intDO), ";")(0))
        bln当前 = (rsPoint.RecordCount <> 0)

        If bln当前 Then
            If Not bln左 Then
                recttmp.Left = rsPoint!X坐标
                recttmp.Top = rsPoint!Y坐标
            End If

            bln断开 = (rsPoint!断开 = 1)

            '将当前脉搏加入区域集合
            If Not bln区域 Then
                intCount = intCount + 1
                ReDim Preserve PtInPoly(intCount)
                str脉搏 = str脉搏 & "," & rsPoint!X坐标 + T_DrawClient.列单位 / 2 & ";" & rsPoint!Y坐标
            End If
        End If

        bln右 = False

        If Not bln断开 Then
            rsPoint.Filter = "项目序号=" & gint脉搏 & " And X坐标>" & Val(Split(arrData(intDO), ";")(0))
            
            If rsPoint.RecordCount <> 0 Then
                rsPoint.Sort = "X坐标 ASC"
                sin右X = rsPoint!X坐标
            
                '根据当前坐标获取时间
                str右 = GetXCoordinate(sin右X, strBeginDate, False)
                str当前 = GetXCoordinate(Val(Split(arrData(intDO), ";")(0)), strBeginDate, False)
                '当前点和下一时间点间隔一天没有数据就断开
                If DateDiff("d", CDate(Split(str当前, ",")(0)), CDate(Split(str右, ",")(0))) < 2 Then
                    bln右 = True
                    recttmp.Right = rsPoint!X坐标
                    recttmp.Bottom = rsPoint!Y坐标
                    '将右脉搏加入区域集合
                    intCount = intCount + 1
                    ReDim Preserve PtInPoly(intCount)
                    str脉搏 = str脉搏 & "," & rsPoint!X坐标 + T_DrawClient.列单位 / 2 & ";" & rsPoint!Y坐标
                End If
            End If
        End If
        
        '先把左边封闭
        If bln区域 = False Then
            If bln当前 = True Then
                '与左列或当前列的脉搏连线
                Call DrawLine(lngDC, recttmp.Left + T_DrawClient.列单位 / 2, recttmp.Top, SinX + T_DrawClient.列单位 / 2, sinY, PS_SOLID, intFine, RGB_RED)
            End If

            bln区域 = (bln左 Or bln右) And bln当前
        End If
        
        '找到右边的封闭区进行连线
        If bln区域 Then
            bln区域 = False

            If bln右 = True Then
                '判断当前心率对应的下一个脉搏和下一个心率X坐标是否相等,不相等就封闭区域
                If intDO < intMax Then
                    If recttmp.Right = Val(Split(arrData(intDO + 1), ";")(0)) Then
                        bln区域 = True
                    End If
                End If
            End If
            
            
            If Not bln区域 Then
                '组织区域,从脉搏开始,然后转到心率(心率从最后开始,再回到之前的心率,再回到第一个脉搏,形成封闭区域)
                intCount = 1
                str脉搏 = Mid(str脉搏, 2)
                arrPt = Split(str脉搏, ",")
                intCols = UBound(arrPt)

                For intCOl = 0 To intCols
                    PtInPoly(intCount).X = Split(arrPt(intCOl), ";")(0)
                    PtInPoly(intCount).Y = Split(arrPt(intCOl), ";")(1)
                    intCount = intCount + 1
                Next

                str心率 = Mid(str心率, 2)
                arrPt = Split(str心率, ",")
                intCols = UBound(arrPt)

                For intCOl = intCols To 0 Step -1
                    PtInPoly(intCount).X = Split(arrPt(intCOl), ";")(0)
                    PtInPoly(intCount).Y = Split(arrPt(intCOl), ";")(1)
                    intCount = intCount + 1
                Next

                '加上起点形成封闭区域
                ReDim Preserve PtInPoly(intCount)
                PtInPoly(intCount).X = PtInPoly(1).X
                PtInPoly(intCount).Y = PtInPoly(1).Y
                
                '填充该区域
                Call DrawPoly(lngDC, PtInPoly, UBound(Split(str脉搏, ",")) + 1)

            End If
        End If

        If Not bln区域 Then
            intCount = 0
            str脉搏 = ""
            str心率 = ""
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
'功能:计算组织重合的点
'---------------------------------------------------------------------------------------
    Dim SinX, sinY As Single
    Dim rsConVerPoint As New ADODB.Recordset
    Dim strFields, strValues As String
    Dim lng项目序号 As Long
    Dim strPart As String
    On Error GoTo Errhand
    
    If rsPiont.RecordCount = 0 Then Exit Sub
    
    strFields = "重叠标识," & adLongVarChar & ",30|重叠数目," & adInteger & ",30|项目序号," & adLongVarChar & ",18|" & _
        "体温部位," & adLongVarChar & ",20"
    Call Record_Init(rsConVerPoint, strFields)
    
    '计算重合的点
    rsPiont.Filter = ""
    rsPiont.Sort = "X坐标,Y坐标"
    With rsPiont
        Do While Not .EOF
            If SinX = Val(!X坐标) And sinY = Val(!Y坐标) Then
                strFields = "重叠标识|重叠数目|项目序号"
                rsConVerPoint.Filter = "重叠标识='" & SinX & "," & sinY & "'"
                If rsConVerPoint.RecordCount = 0 Then
                    strValues = SinX & "," & sinY
                    strValues = strValues & "|" & 2
                    strValues = strValues & "|" & lng项目序号 & "," & !项目序号
                    Call Record_Add(rsConVerPoint, strFields, strValues)
                Else
                    strFields = "重叠数目|项目序号"
                    strValues = Val(rsConVerPoint!重叠数目) + 1
                    strValues = strValues & "|" & rsConVerPoint!项目序号 & "," & !项目序号
                    Call Record_Update(rsConVerPoint, strFields, strValues, "重叠标识|" & SinX & "," & sinY)
                End If
                
                If InStr(1, "," & rsConVerPoint!项目序号 & ",", "," & gint体温 & ",") > 0 And strPart <> "" Then
                    strFields = "体温部位": strValues = strPart
                    Call Record_Update(rsConVerPoint, strFields, strValues, "重叠标识|" & SinX & "," & sinY)
                    strPart = ""
                End If
                
                rsConVerPoint.Filter = ""
                
            End If
            SinX = Val(!X坐标)
            sinY = Val(!Y坐标)
            lng项目序号 = !项目序号
            If lng项目序号 = gint体温 Then strPart = !部位
        .MoveNext
        Loop
    End With
    
    '组织更新重复点的输出标识
    
    Dim strTemp As String
    Dim rsTemp As New ADODB.Recordset
    Dim lngID As Long
   'Dim strPart As String
    Dim strOverpart As String
        
    If rsConVerPoint.RecordCount > 0 Then
        rsConVerPoint.MoveFirst
        Do While Not rsConVerPoint.EOF
                strTemp = rsConVerPoint!项目序号
                strOverpart = ""
                strPart = ""
                
                '由于体温的重合点设置 存在部位设置
                If InStr(1, "," & strTemp & ",", "," & gint体温 & ",") > 0 Then
                    strTemp = "0," & strTemp & ",0"
                    strTemp = Replace(strTemp, "," & gint体温 & ",", ",")
                    
                    gstrSQL = " select  C.序号,C.上级序号,C.项目序号,C.体温部位 from 体温重叠标记 C," & _
                        "   (Select 序号 " & _
                        "   From 体温重叠标记 A,(select 上级序号,count(1) 数量 " & _
                        "   from 体温重叠标记 where 项目序号 in (" & strTemp & ") or (项目序号=1 and nvl(体温部位,'腋温')=[2]) group by 上级序号) B " & _
                        "   where A.序号=B.上级序号 and A.重叠数目=B.数量 and B.数量=[1]) D " & _
                        "   where C.上级序号=D.序号 and C.项目序号 is not null order by C.序号"
                Else
                    gstrSQL = " select  C.序号,C.上级序号,C.项目序号,C.体温部位 from 体温重叠标记 C," & _
                        "   (Select 序号 " & _
                        "   From 体温重叠标记 A,(select 上级序号,count(1) 数量 " & _
                        "   from 体温重叠标记 where 项目序号 in (" & strTemp & ") group by 上级序号) B " & _
                        "   where A.序号=B.上级序号 and A.重叠数目=B.数量 and B.数量=[1]) D " & _
                        "   where C.上级序号=D.序号 and C.项目序号 is not null order by C.序号"
                End If
                Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "提取重叠项目", Val(rsConVerPoint!重叠数目), zlCommFun.Nvl(rsConVerPoint!体温部位))
                
                If rsTemp.RecordCount > 0 Then
                    lngID = rsTemp!项目序号
                    strPart = rsTemp!上级序号  '重叠部位存放序号
                    
                    Do While Not rsTemp.EOF
                        If lngID <> rsTemp!项目序号 Then
                            strOverpart = strOverpart & "," & rsTemp!项目序号
                        End If
                    rsTemp.MoveNext
                    Loop
                    
                    If strOverpart <> "" Then strOverpart = Mid(strOverpart, 2)
                    
                    '更新重复的点
                    rsPiont.Filter = "X坐标=" & Split(rsConVerPoint!重叠标识, ",")(0) & _
                        " and Y坐标=" & Split(rsConVerPoint!重叠标识, ",")(1)
                        
                    Do While Not rsPiont.EOF
                        If lngID = rsPiont!项目序号 Then
                            rsPiont!重叠项目 = strOverpart
                            rsPiont!部位 = strPart
                        Else
                            rsPiont!重叠 = 1
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

Public Function GetXCoordinate(ByVal strInput As String, ByVal strBeginDate As String, Optional ByVal bln坐标 As Boolean = True) As String

    '根据时间得到X坐标或根据X坐标转换为时间范围
    Dim SinX   As Single

    Dim intDO  As Integer, intMax As Integer

    Dim intDay As Integer, intTime As Integer

    Dim strDay As String, strTime As String

    On Error GoTo Errhand
    
    If bln坐标 Then
        '第一天是0,第七天是6
        strDay = Split(strInput, " ")(0)

        If InStr(1, strInput, " ") <> 0 Then
            strTime = Split(strInput, " ")(1)
        Else
            strTime = "00:00:00"
        End If

        intDay = DateDiff("d", CDate(strBeginDate), CDate(strInput))
        
        '得到当天的刻度
        intMax = 5

        For intDO = 0 To intMax

            If strTime >= Split(gvarTime(intDO), ",")(0) And strTime <= Split(gvarTime(intDO), ",")(1) Then
                intTime = intDO
                Exit For
            End If
        Next
        
        '计算得到X坐标(每天6列,以列数*列单位得到坐标)
        SinX = Format(T_DrawClient.体温区域.Left + (T_DrawClient.列单位 * (intDay * 6 + intTime)), "#0.0")
        GetXCoordinate = SinX
    Else
        '计算得到相差多少个刻度
        SinX = Val(strInput)
        intTime = (SinX - T_DrawClient.体温区域.Left) \ T_DrawClient.列单位
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


Public Function GetYCoordinate(ByVal objDraw As Object, ByVal rsDrawItems As ADODB.Recordset, ByVal int项目序号 As Integer, ByVal strInput As String, Optional ByVal bln坐标 As Boolean = True, Optional lngDC As Long = 0, Optional ByVal blnOutput As Boolean = False) As String

    Dim lngCurX As Long, sinCurY As Single, sinScale As Single

    On Error GoTo Errhand

    '返回指定曲线数据的Y坐标或根据Y坐标计算数据
    '测试该函数的正确性可以在Paint_Canvas中增加该代码实现(思想:由该函数自己根据数据计算得到Y坐标,再转换为数据,再转换为坐标后输出字符进行核对,打印无误则说明转换无误):
    '   Call GetYCoordinate(1, GetYCoordinate(1, "200," & GetYCoordinate(1, "37.5", True, False), False),true,true)
    
    rsDrawItems.Filter = "项目序号=" & int项目序号

    If rsDrawItems.RecordCount = 0 Then
        If int项目序号 = gint心率 Then rsDrawItems.Filter = "项目序号=2"
    End If
    
    If rsDrawItems.RecordCount = 0 Then
        GetYCoordinate = 0
        Exit Function
    End If
    
    If bln坐标 Then
        '得到有效数据起始坐标
        lngCurX = Split(rsDrawItems!最大值坐标, ",")(0)
        sinCurY = Split(rsDrawItems!最大值坐标, ",")(1)
        
        '根据最大值与当前值之间的差额,以及最小值,计算得到相差多少个刻度,再根据单位刻度得到实际坐标
        sinScale = Format((rsDrawItems!最大值 - Val(strInput)) / rsDrawItems!单位值 * Val(Split(rsDrawItems!单位刻度, ",")(0)), "#0.0")
        GetYCoordinate = Format(sinCurY + sinScale, "#0")
        
        If blnOutput Then
            '在指定坐标输出字符进行核对
            Call SetTextColor(lngDC, RGB_BLUE)
            Call GetTextRect(objDraw, 202, GetYCoordinate, "·", T_DrawClient.刻度单位)
            Call DrawText(lngDC, "·", -1, T_LableRect, DT_CENTER)
        End If
    Else
        '得到传入的坐标值
        lngCurX = Split(strInput, ",")(0)
        sinCurY = Split(strInput, ",")(1)
        
        '(坐标-最大值坐标)/单位刻度得到相差多少个刻度
        '(最大值-单位刻度*单位值)得到实际数据
        sinScale = Format((sinCurY - Split(rsDrawItems!最大值坐标, ",")(1)) / Val(Split(rsDrawItems!单位刻度, ",")(0)), "#0.0")
        GetYCoordinate = Format(rsDrawItems!最大值 - sinScale * rsDrawItems!单位值, "#0.0")
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
    '功能： 获得最小最大时间范围
    '参数：
    '返回：
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
    ByVal strEditor As String, ByVal bln汇总当天 As Boolean, Optional ByVal bln录入小时 As Boolean, Optional ByVal blnEdit As Boolean = False) As ADODB.Recordset
'---------------------------------------------------------------------------------------------------------------------------------------------
'功能:获取某一天的表格项目数值信息(共体温单展示和打印使用)
'参数：rsCollect 项目数据信息,dtDate 某天日期,dtBegin 体温单开始日期,strEditor 项目有关信息：项目序号;项目名称;项目频次;项目表示;项目性质;入院首测
'      bln汇总昨天 参数：汇总、波动项目显示(True)当天数据,(false)昨天数据  blnEdit 是否是编辑状态（在编辑程序=true）
'      bln录入小时 51282,刘鹏飞,2012-08-03,全天汇总显示录入时间 10.30.20(DYEY要求手工录入汇总时间H)
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsDayData As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim DtDay As Date
    Dim intType As Integer, intHour As Integer, intHour1 As Integer, int类别 As Integer, int序号 As Integer
    Dim strBegin As String, strEnd As String, strCenter As String
    Dim strFileds As String, strValues As String, strValues1 As String, strFind As String, strTime As String
    Dim dblData As Double, int未记说明 As Integer, int数据来源 As Integer, lngID As Long, lng来源ID As Long, int共用 As Integer
    Dim lngNO As Long
    Dim i As Integer, intCount As Integer, intColFirst As Integer, strHourTime As String
    Dim bln波动 As Boolean
    'Dim bln首次汇总 As Boolean '参数:首日汇总显示时间
    Dim dtCurrDate As Date
    
    '项目有关信息
    Dim lngItemNO As Long, strName As String, int记录频次 As Integer, int项目表示 As Integer, int项目性质 As Integer, bln入院首测 As Boolean
    Dim arrEditor() As String
    
    On Error GoTo Errhand
    
    arrEditor = Split(strEditor, ";")
    lngItemNO = Val(arrEditor(0))
    strName = arrEditor(1)
    int记录频次 = Val(arrEditor(2))
    int项目表示 = Val(arrEditor(3))
    int项目性质 = Val(arrEditor(4))
    bln入院首测 = (Val(arrEditor(5)) = 1)
    '汇总项目不存在入院首测
    If int项目性质 = 4 Then bln入院首测 = False
    bln波动 = IsWaveItem(lngItemNO) '是否是波动项目
    DtDay = dtDate
    
    '初始化记录集
    strFileds = "ID," & adDouble & ",18|时间," & adLongVarChar & ",20|项目序号," & adDouble & ",18|项目名称," & adLongVarChar & ",20|记录内容," & adLongVarChar & ",100|" & _
        "体温部位," & adLongVarChar & ",20|未记说明," & adLongVarChar & ",100|数据来源," & adDouble & ",1|显示," & adDouble & _
        ",1|来源ID," & adDouble & ",18|共用," & adDouble & ",1|序号," & adDouble & ",1|汇总小时," & adLongVarChar & ",100"
    Call Record_Init(rsDayData, strFileds)
    strFileds = "ID|时间|项目序号|项目名称|记录内容|体温部位|未记说明|数据来源|显示|来源ID|共用|序号|汇总小时"
    
    If blnEdit And bln波动 Then int项目表示 = 0
    
ErrBegin:
    dtDate = DtDay
    rsCollect.Filter = ""
    '汇总/波动项目类型=2
    If int项目表示 = 4 Or bln波动 Then
        intType = 2
        If int记录频次 = 0 Then
            int记录频次 = 2
        ElseIf int记录频次 > 2 Then
            int记录频次 = 2
        End If
        
        '根据参数确定汇总/波动项目汇总前一天/当天的数据（根据汇总时段）
        If Not bln汇总当天 Then dtDate = CDate(dtDate) - 1
    Else
        intType = 1
    End If
    
    '提取当前服务器时间
    dtCurrDate = CDate(Format(zldatabase.Currentdate, "YYYY-MM-DD HH:mm:ss"))
    
    '根据类型，频次和序号 不可能找不到信息
    mrsTabTime.Filter = "类型=" & intType & " and 频次=" & int记录频次
    If mrsTabTime.RecordCount = 0 Then
        MsgBox "请在护理项目管理中设置[" & IIf(intType = 2, "汇总项目", "体温表格项目") & "]时段信息!", vbInformation, gstrSysName
        Set ReturnItemRecord = rsDayData
        Exit Function
    End If
    
    intColFirst = 1
    
    With mrsTabTime
        .MoveFirst
        '提取频次时间段
        Do While Not .EOF
            int类别 = Val(!类别)
            int序号 = Val(Nvl(!序号))
            intHour = CInt(24 / int记录频次)
            strBegin = Format(IIf(IsDate(Trim(Nvl(!开始))) = False, (Val(Nvl(!序号)) - 1) * intHour & ":00:00", !开始), "HH:mm:ss")
            strEnd = Format(IIf(IsDate(Trim(Nvl(!结束))) = False, Val(Nvl(!序号)) * intHour - 1 & ":59:59", !结束), "HH:mm:ss")
            '确定频次时间范围
            If int序号 = int记录频次 Then
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
            '获取中点时间段信息
            intHour = DateDiff("H", CDate(strBegin), CDate(strEnd) + 0.00001) / 2
            strCenter = DateAdd("H", intHour, CDate(strBegin)) '中点时间
            If CDate(strCenter) > CDate(strEnd) Then strCenter = strEnd
            
            strFind = "时间>='" & Format(strBegin, "YYYY-MM-DD HH:mm:ss") & "' and 时间<='" & Format(strEnd, "YYYY-MM-DD HH:mm:ss") & "'"
            
            lngNO = lngItemNO
            
            If int项目性质 = 2 Then
                rsCollect.Filter = "项目序号=" & lngItemNO & " and 项目名称='" & strName & "' And " & strFind
                If lngItemNO = 4 Then '血压为活动项目继续按固定项目处理
                    rsCollect.Filter = "(项目序号=4 And " & strFind & ") OR (项目序号=5 And " & strFind & ")"
                End If
            Else
                If lngItemNO <> 4 Then
                    rsCollect.Filter = "项目序号=" & lngItemNO & " And " & strFind
                Else
                    rsCollect.Filter = "(项目序号=4 And " & strFind & ") OR (项目序号=5 And " & strFind & ")"
                End If
            End If
            
            rsCollect.Sort = "项目序号,时间"
            
            If int项目表示 = 4 Then '汇总项目
                dblData = 0: int未记说明 = 0: strValues = ""
                If lngItemNO = 4 Then '血压如果修改为汇总项目 直接按波动血压处理
                    int项目表示 = 6
                    GoTo ErrBegin
                End If
                
                '如果当前时间小于汇总时间段,不进行汇总
                If dtCurrDate < CDate(strEnd) And Not blnEdit And Not gbln出院 Then GoTo ErrNext
                
                int未记说明 = 0: int数据来源 = 0: lngID = 0: lng来源ID = 0: int共用 = 0
                strValues1 = "": intHour1 = 0: strHourTime = ""
                '先循环父项目本身
                Do While Not rsCollect.EOF
                    If Val(Nvl(rsCollect!记录类型)) = 1 Then
                        If int未记说明 < Val(Nvl(rsCollect!未记说明)) Then int未记说明 = Val(Nvl(rsCollect!未记说明))
                        If InStr(1, ",0,9,", "," & Val(Nvl(rsCollect!数据来源)) & ",") = 0 Then
                            int数据来源 = Val(Nvl(rsCollect!数据来源))
                            lng来源ID = Val(Nvl(rsCollect!来源ID))
                            int共用 = Val(Nvl(rsCollect!共用))
                            lngID = Val(Nvl(rsCollect!Id))
                        ElseIf lngID = 0 Then
                            lngID = Val(Nvl(rsCollect!Id))
                        End If
                        dblData = dblData + Val(Nvl(rsCollect!结果))
                    Else
                        intHour1 = -1
                        strHourTime = Format(rsCollect!时间, "YYYY-MM-DD HH:mm:ss") & ";" & Val(Nvl(rsCollect!Id))
                        strValues1 = Val(Nvl(rsCollect!结果))
                        If Val(strValues1) < 0 Then strValues1 = ""
                        If Val(strValues1) > 24 Then strValues1 = 24
                    End If
                rsCollect.MoveNext
                Loop
                
                If rsCollect.RecordCount > 0 Then rsCollect.MoveFirst
                
                If int项目性质 = 2 Then
                    '活动项目按部位统计自身
                Else
                    '开始汇总子项目
                    Set rsTemp = SetCollectPItem(lngItemNO)
                    rsTemp.Filter = 0
                    Do While Not rsTemp.EOF
                        '对于同步过来的老数据 由于父项已经汇总了 此处不再进行汇总
                        If Val(Nvl(rsTemp!序号, 0)) <> lngItemNO Then
                            rsCollect.Filter = 0
                            rsCollect.Filter = "项目序号=" & Val(Nvl(rsTemp!序号, 0)) & " And 数据来源<>9 And 记录类型=1 " & " And " & strFind
                            Do While Not rsCollect.EOF
                                dblData = dblData + Val(Nvl(rsCollect!结果))
                                If lng来源ID = 0 Then
                                    If InStr(1, ",0,9,", "," & Val(Nvl(rsCollect!数据来源)) & ",") = 0 Then
                                        int数据来源 = Val(Nvl(rsCollect!数据来源))
                                        lng来源ID = Val(Nvl(rsCollect!来源ID))
                                        int共用 = Val(Nvl(rsCollect!共用))
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
                    If bln录入小时 = True And int记录频次 = 1 Then
                        strValues1 = IIf(dblData = 0, "", IIf(strValues1 = "", "", "(" & strValues1 & "h)") & dblData)
                    Else
                        intHour1 = 0
                        strValues1 = IIf(dblData = 0, "", dblData)
                    End If
                    '51282,刘鹏飞,2012-07-11
                    '51282,刘鹏飞,2012-08-03,DYEY目前要求全天汇总可以手工录入汇总小时
                    '全天汇总首次不满汇总时段显示汇总时间小时，比如“入量”汇总是200ml,首日统计时间为18h ,表格栏应该显示为“(18h)200”
'                    If blnEdit = False And int记录频次 = 1 And (Format(dtBegin, "YYYY-MM-DD") = Format(dtDate, "YYYY-MM-DD") Or Format(dtBegin, "YYYY-MM-DD") = Format(DtDay, "YYYY-MM-DD")) Then
'                        bln首次汇总 = (Val(zlDatabase.GetPara("首日汇总显示时间", glngSys, 1255, "0")) = 1)
'                        '计算汇总时段相差的小时数
'                        intHour1 = Format(DateDiff("n", CDate(strBegin), CDate(strEnd) + 0.00001) / 60, "#0")
'                        If bln汇总当天 = True And bln首次汇总 = True Then
'                            '汇总当天只处理体温单当天，汇总时段肯定是当天开始到第二天或当天结束，体温单开始时间和汇总结束时间相隔的小时数只有
'                            '大于0并且小于汇总时段间隔的小时才满足条件
'                            If Format(dtBegin, "YYYY-MM-DD") = Format(dtDate, "YYYY-MM-DD") Then
'                                '计算体温单开始时间和汇总结束时间相差多少小时
'                                intHour = Format(DateDiff("n", CDate(dtBegin), CDate(strEnd) + 0.00001) / 60, "#0")
'                                If intHour > 0 And intHour < intHour1 Then strValues1 = "(" & intHour & "h)" & strValues1
'                            End If
'                        ElseIf bln首次汇总 = True Then '汇总项目汇总昨天，存在两种情况，一种是体温单的开始时间在第一天汇总时段内；一种是体温单的开始时间不在第一天汇总时段内
'                            '（可能在第二天汇总时段内，也可能不在）。这两种情况只能满足其一。
'                            If Format(dtBegin, "YYYY-MM-DD") = Format(DtDay, "YYYY-MM-DD") Then
'                                '计算体温单开始时间和汇总结束时间相差多少小时
'                                intHour = Format(DateDiff("n", CDate(dtBegin), CDate(strEnd) + 0.00001) / 60, "#0")
'                                If intHour > 0 And intHour < intHour1 Then strValues1 = "(" & intHour & "h)" & strValues1
'                            End If
'
'                            If Format(dtBegin, "YYYY-MM-DD") = Format(dtDate, "YYYY-MM-DD") Then
'                                '计算体温单开始时间和汇总结束时间相差多少小时
'                                intHour = Format(DateDiff("n", CDate(dtBegin), CDate(strEnd) + 0.00001) / 60, "#0")
'                                If intHour > 0 And intHour < intHour1 Then strValues1 = "(" & intHour & "h)" & strValues1
'                            End If
'                        End If
'                    End If
                    If int项目性质 = 2 Then
                        strValues = lngID & "|" & CDate(strCenter) & "|" & lngItemNO & "|" & strName & "|" & _
                                        strValues1 & "|" & strName & "|" & int未记说明 & "|" & _
                                        int数据来源 & "|" & 1 & "|" & lng来源ID & "|" & int共用 & "|" & int序号 & "|" & intHour1 & ";" & strHourTime
                    Else
                        strValues = lngID & "|" & CDate(strCenter) & "|" & lngItemNO & "|" & strName & "|" & _
                                    strValues1 & "|" & "" & "|" & "" & "|" & _
                                    int数据来源 & "|" & 1 & "|" & lng来源ID & "|" & int共用 & "|" & int序号 & "|" & intHour1 & ";" & strHourTime
                    End If
                    Call Record_Add(rsDayData, strFileds, strValues)
                    strValues1 = ""
                End If
            ElseIf bln波动 Then '波动项目
                intCount = 0: i = 0
                If lngNO = 4 Then intCount = 1
                
                If bln入院首测 = True And Format(dtBegin, "YYYY-MM-DD") = Format(dtDate, "YYYY-MM-DD") And intColFirst = 1 Then 'dtBegin >= CDate(strBegin) And dtBegin <= CDate(strEnd) Then
                    int类别 = 1 '提取第一条数据
                    GoTo ErrRead
                End If
                
                '如果当前时间小于汇总时间段,不进行汇总
                If dtCurrDate < CDate(strEnd) And Not blnEdit And Not gbln出院 Then GoTo ErrNext
                
                For i = 0 To intCount
                    If i = 1 Then lngNO = 5
                    If intCount = 1 Then '血压项目重新提取
                        rsCollect.Filter = 0
                        rsCollect.Filter = "项目序号=" & lngNO & " And " & strFind
                    End If
                    strValues = "": strValues1 = "": strTime = "": dblData = 0
                    Do While Not rsCollect.EOF
                        If dblData <> 0 Then
                            '提取最大值
                            If Val(strValues) < Val(Nvl(rsCollect!结果)) Then
                                strValues = Val(Nvl(rsCollect!结果))
                            End If
                            '提取最小值
                            If Val(strValues1) > Val(Nvl(rsCollect!结果)) Then
                                strValues1 = Val(Nvl(rsCollect!结果))
                            End If
                        Else
                            dblData = 99
                            If IsNumeric(Nvl(rsCollect!结果)) Then
                                strValues = Val(Nvl(rsCollect!结果))
                                strValues1 = strValues
                            Else
                                strValues = ""
                                strValues1 = ""
                            End If
                            
                            lngID = Val(Nvl(rsCollect!Id))
                            int数据来源 = Val(Nvl(rsCollect!数据来源))
                            lng来源ID = Val(Nvl(rsCollect!来源ID))
                            int共用 = Val(Nvl(rsCollect!共用))
                            strTime = Nvl(rsCollect!时间)
                        End If
                        rsCollect.MoveNext
                    Loop
                    
                    If dblData <> 0 Then
                        If Val(strValues) <> Val(strValues1) Then
                            strValues1 = Val(strValues1) & "-" & Val(strValues)
                        Else
                            strValues1 = IIf(strValues = "", "", Val(strValues))
                        End If
                        
                        '将结果保存到记录集中
                        strValues = lngID & "|" & CDate(strTime) & "|" & lngNO & "|" & IIf(lngItemNO <> 4, strName, IIf(lngNO = 4, "收缩压", "舒张压")) & "|" & _
                            strValues1 & "|" & "" & "|" & "" & "|" & int数据来源 & "|" & _
                            1 & "|" & lng来源ID & "|" & int共用 & "|" & int序号 & "|0"
                        Call Record_Add(rsDayData, strFileds, strValues)
                    End If
                Next i
            Else '非汇总项目
                intCount = 0: i = 0
                '--对于血压需要分别处理收缩压和舒张压
                If lngNO = 4 Then intCount = 1
                
                If bln入院首测 = True And Format(dtBegin, "YYYY-MM-DD") = Format(dtDate, "YYYY-MM-DD") And intColFirst = 1 Then 'dtBegin >= CDate(strBegin) And dtBegin <= CDate(strEnd) Then
                    int类别 = 1 '提取第一条数据
                End If
ErrRead:
                For i = 0 To intCount
                    If i = 1 Then lngNO = 5
                    If intCount = 1 Then '血压项目重新过滤
                        rsCollect.Filter = 0
                        rsCollect.Filter = "项目序号=" & lngNO & " And " & strFind
                    End If
                    strValues = "": strValues1 = "": strTime = ""
                    Do While Not rsCollect.EOF
                        intColFirst = 2
                        '对于血压进行处理
                        If lngNO = Val(Nvl(rsCollect!项目序号)) Then
                            Select Case int类别
                                Case 1 '第一条
                                    If rsCollect.RecordCount > 0 Then rsCollect.MoveFirst
                                        strValues = Val(Nvl(rsCollect!Id)) & "|" & CDate(rsCollect!时间) & "|" & Val(Nvl(rsCollect!项目序号)) & "|" & Nvl(rsCollect!项目名称) & "|" & _
                                            Nvl(rsCollect!结果) & "|" & Nvl(rsCollect!体温部位) & "|" & Nvl(rsCollect!未记说明) & "|" & Val(Nvl(rsCollect!数据来源)) & "|" & _
                                            Val(Nvl(rsCollect!显示)) & "|" & Val(Nvl(rsCollect!来源ID)) & "|" & Val(Nvl(rsCollect!共用)) & "|" & int序号 & "|0"
                                    Exit Do
                                Case 2 '中间一条
                                    strValues = Val(Nvl(rsCollect!Id)) & "|" & CDate(rsCollect!时间) & "|" & Val(Nvl(rsCollect!项目序号)) & "|" & Nvl(rsCollect!项目名称) & "|" & _
                                            Nvl(rsCollect!结果) & "|" & Nvl(rsCollect!体温部位) & "|" & Nvl(rsCollect!未记说明) & "|" & Val(Nvl(rsCollect!数据来源)) & "|" & _
                                            Val(Nvl(rsCollect!显示)) & "|" & Val(Nvl(rsCollect!来源ID)) & "|" & Val(Nvl(rsCollect!共用)) & "|" & int序号 & "|0"
                                    If strValues1 <> "" Then
                                        '检查那个接近中点时间
                                        If Abs(DateDiff("s", Format(CDate(rsCollect!时间), "YYYY-MM-DD HH:mm:ss"), Format(strCenter, "YYYY-MM-DD HH:mm:ss"))) > _
                                            Abs(DateDiff("s", Format(CDate(strTime), "YYYY-MM-DD HH:mm:ss"), Format(strCenter, "YYYY-MM-DD HH:mm:ss"))) Then
                                             strValues = strValues1
                                        End If
                                    End If
                                    strValues1 = strValues
                                    strTime = rsCollect!时间
                                Case Else '最后一条
                                    If rsCollect.RecordCount > 0 Then rsCollect.MoveLast
                                        strValues = Val(Nvl(rsCollect!Id)) & "|" & CDate(rsCollect!时间) & "|" & Val(Nvl(rsCollect!项目序号)) & "|" & Nvl(rsCollect!项目名称) & "|" & _
                                            Nvl(rsCollect!结果) & "|" & Nvl(rsCollect!体温部位) & "|" & Nvl(rsCollect!未记说明) & "|" & Val(Nvl(rsCollect!数据来源)) & "|" & _
                                            Val(Nvl(rsCollect!显示)) & "|" & Val(Nvl(rsCollect!来源ID)) & "|" & Val(Nvl(rsCollect!共用)) & "|" & int序号 & "|0"
                                    Exit Do
                            End Select
                        End If
                    rsCollect.MoveNext
                    Loop
                    '添加记录集信息
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
    '提所有表格项目时段信息
    gstrSQL = "SELECT 序号, 开始, 结束, 频次,类别, 类型" & vbNewLine & _
                "  FROM (SELECT DECODE(类别, 3, 1, 类别) 序号," & vbNewLine & _
                "               开始 || ':00' 开始," & vbNewLine & _
                "               结束 || ':59' 结束," & vbNewLine & _
                "               DECODE(类别, 3, 1, 2) 频次,0 类别," & vbNewLine & _
                "               2 类型" & vbNewLine & _
                "          FROM 护理汇总时段 WHERE 单据=1" & vbNewLine & _
                "        UNION ALL" & vbNewLine & _
                "        SELECT 序号, 开始 || ':00' 开始, 结束 || ':59' 结束, 频次,类别, 1 类型" & vbNewLine & _
                "          FROM 护理项目频次)" & vbNewLine & _
                " ORDER BY 类型, 频次, 序号"

    Call zldatabase.OpenRecordset(mrsTabTime, gstrSQL, "体温单")
    
    If Not (mrsCollect Is Nothing) Then If mrsCollect.State = 1 Then mrsCollect.Close
    '提取护理汇总项目
    gstrSQL = " SELECT 序号,父序号 FROM 护理汇总项目"
    Call zldatabase.OpenRecordset(mrsCollect, gstrSQL, "护理汇总项目")
    
    If Not (mrsWave Is Nothing) Then If mrsWave.State = 1 Then mrsWave.Close
    '护理波动项目
    gstrSQL = "　SELECT 项目序号 FROM 护理波动项目"
    Call zldatabase.OpenRecordset(mrsWave, gstrSQL, "护理波动项目")
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function SetCollectPItem(ByVal lngItemNO As Long) As ADODB.Recordset
'---------------------------------------------------------------------------
'功能:根据父项目ID重新组织子项目
'---------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim rsCollect As New ADODB.Recordset
    Dim strFileds As String, strValues As String
    Dim lngNO As Long
    
    On Error GoTo Errhand
    
    '初始化记录集
    strFileds = "序号," & adDouble & ",18|父序号," & adDouble & ",18"
    Call Record_Init(rsTemp, strFileds)
    Call Record_Init(rsCollect, strFileds)
    strFileds = "序号|父序号"
    
    mrsCollect.Filter = 0
   '复制记录集
    With mrsCollect
        Do While Not .EOF
            strValues = Val(Nvl(!序号)) & "|" & Val(Nvl(!父序号))
            Call Record_Add(rsCollect, strFileds, strValues)
            .MoveNext
        Loop
    End With
    
    rsCollect.Filter = "父序号=" & lngItemNO
    With rsCollect
        Do While Not .EOF
            strValues = Val(Nvl(!序号)) & "|" & lngItemNO
            Call Record_Add(rsTemp, strFileds, strValues)
            lngNO = Val(Nvl(!序号))
            '循环递归调用(获取子项的子项)
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
'功能: SetCollectPItem 调用
    
    Dim rsCollect As New ADODB.Recordset
    Dim strFileds As String, strValues As String
    Dim lngNO As Long
    
    '初始化记录集
    strFileds = "序号," & adDouble & ",18|父序号," & adDouble & ",18"
    Call Record_Init(rsCollect, strFileds)
    strFileds = "序号|父序号"
    
    mrsCollect.Filter = 0
   '复制记录集
    With mrsCollect
        Do While Not .EOF
            strValues = Val(Nvl(!序号)) & "|" & Val(Nvl(!父序号))
            Call Record_Add(rsCollect, strFileds, strValues)
            .MoveNext
        Loop
    End With
    
    rsCollect.Filter = "父序号=" & lngItemNO
    With rsCollect
        Do While Not .EOF
            strValues = Val(Nvl(!序号)) & "|" & lngParent
            Call Record_Add(rsTemp, strFileds, strValues)
            lngNO = Val(Nvl(!序号))
            '循环递归调用(获取子项的子项)
            Call SetCollectCItem(rsTemp, lngParent, lngNO)
            .MoveNext
        Loop
    End With
End Sub

Public Function IsWaveItem(ByVal lngItemNO As Long) As Boolean
'检查是否是波动项目
    If mrsWave Is Nothing Then Exit Function
    If mrsWave.State = 1 Then
        mrsWave.Filter = 0
        mrsWave.Filter = "项目序号=" & lngItemNO
        IsWaveItem = (mrsWave.RecordCount > 0)
    End If
End Function

Public Function SetNTPrinterPaper(ByVal lngHwnd As Long, ByVal intWidth As Integer, ByVal intHeight As Integer, _
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
        lngSize = DocumentProperties(lngHwnd, lngHandle, strPrtName, 0&, 0&, 0&)
        'Reserve memory for the actual size of the DEVMODE.
        ReDim arrDevMode(1 To lngSize)
    
        'Fill the DEVMODE from the printer.
        lngSize = DocumentProperties(lngHwnd, lngHandle, strPrtName, arrDevMode(1), 0&, DM_OUT_BUFFER)
        'Copy the Public (predefined) portion of the DEVMODE.
        Call CopyMemory(vDevMode, arrDevMode(1), Len(vDevMode))
        
        '设置打印文档属性
        vDevMode.dmOrientation = intOrient
        vDevMode.dmPaperSize = 256
        vDevMode.dmPaperWidth = intWidth * 10 'in tenths of a millimeter
        vDevMode.dmPaperLength = intHeight * 10 'in tenths of a millimeter
        vDevMode.dmCopies = intCopys
        'vDevMode.dmCollate = 0& '高级打印功能(当取消时,Copies只支持1;但不知怎么取不了)
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
'功能：在设置自定义纸张
'参数：是以绨为单位
    If IsWindowsNT Then
        '虽然不能使宽度生效，但能改变PaperSize的属性值
        Printer.Width = lngWidth
        Printer.Height = lngHeight
        SetCustonPager = SetNTPrinterPaper(lngHwnd, lngWidth / conRatemmToTwip, lngHeight / conRatemmToTwip, Printer.Orientation, Printer.Copies)
    Else
        'Windows98系列还是以通常方法处理
        Printer.PaperSize = 256
        Printer.Width = lngWidth
        Printer.Height = lngHeight
    End If
End Function

Public Function GetTimeColor(ByVal intHour As Integer) As Long
'---------------------------------------------
'根据参数获取体温时间颜色
'---------------------------------------------
    Dim blnTag As Boolean
    Dim strTmp As String
    Dim lngBegin As Long, lngEnd As Long
    Dim lngColor As Long
    strTmp = zldatabase.GetPara("体温时间夜班标志", glngSys, 1255, "18;6")
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
