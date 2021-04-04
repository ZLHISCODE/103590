VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPDFConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mblnReset As Boolean
Private marrReset() As Byte

Private mobjCustomReport As Object ' zl9Report.clsReport    '自定义报表对象
Attribute mobjCustomReport.VB_VarHelpID = -1
Private mobjReport As Object ' zlRichEPR.cDockReport        '报告对象
Attribute mobjReport.VB_VarHelpID = -1
Private mobjTable As Object     'zlRichEPR.cEPRTable
Private mobjRchEPR As Object 'zlRichEPR.cRichEPR
Private mobjComLib As Object

Private mlngSys As Long
Private mlngModule As Long
Private mblnMoved As Boolean
Private mcnOracle As ADODB.Connection
Private mstrDBUser As String

Private mstrPrinterDeviceName As String '获取打印时打印机的名称
Private mstr报表编号 As String
Private mstr选中报表格式 As String
Private mbln使用自定义报表 As Boolean

'Window版本函数
Private Type OSVERSIONINFO 'for GetVersionEx API call
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

'注册表关键字根类型
Private Enum REGRoot
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
End Enum

'注册表数据类型
Private Enum REGValueType
    REG_SZ = 1 'Unicode空终结字符串
    REG_EXPAND_SZ = 2 'Unicode空终结字符串
    REG_BINARY = 3 '二进制数值
    REG_DWORD = 4 '32-bit 数字
    REG_DWORD_BIG_ENDIAN = 5
    REG_LINK = 6
    REG_MULTI_SZ = 7 ' 二进制数值串
End Enum

'注册表选项
Private Const READ_CONTROL = &H20000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_READ = KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + READ_CONTROL
Private Const KEY_WRITE = KEY_SET_VALUE + KEY_CREATE_SUB_KEY + READ_CONTROL
Private Const KEY_EXECUTE = KEY_READ
Private Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
      
'操作返回值
Private Const ERROR_SUCCESS = 0
Private Const ERROR_BADKEY = 2
Private Const ERROR_ACCESS_DENIED = 8
      
'打开注册表
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
'查询注册表值
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueEx_BINARY Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
'设置注册表值
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueEx_BINARY Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
'删除注册表目录/项目
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
'关闭注册表
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

'新的打印方式使用-----------------------------------------------------------
'注意以dmFields是Long型,as Long或尾部加&符
Private Const DM_ORIENTATION = &H1&
Private Const DM_PAPERSIZE = &H2&
Private Const DM_PAPERLENGTH = &H4&
Private Const DM_PAPERWIDTH = &H8&
Private Const DM_COPIES = &H100&
Private Const DM_DEFAULTSOURCE = &H200&
Private Const DM_COLLATE = &H8000&
Private Const DM_FORMNAME = &H10000
'Constants for DocumentProperties() call
Private Const DM_COPY = 2
Private Const DM_OUT_BUFFER = DM_COPY
Private Const DM_PROMPT = 4
Private Const DM_IN_PROMPT = DM_PROMPT
Private Const DM_MODIFY = 8
Private Const DM_IN_BUFFER = DM_MODIFY
'Constants for DocumentProperties() return
Private Const IDOK = 1
Private Const IDCANCEL = 2
'Constants for DEVMODE
Private Const CCHFORMNAME = 32
Private Const CCHDEVICENAME = 32

Private Type DEVMODE
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

Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, ByVal pDefault As Long) As Long
Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal hwnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, pDevModeOutput As Any, pDevModeInput As Any, ByVal fMode As Long) As Long
Private Declare Function ResetDC Lib "gdi32" Alias "ResetDCA" (ByVal hDC As Long, lpInitData As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
'创建多级文件目录
Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal DirPath As String) As Long
'纸张打印边界控制================================================================
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
'不同打印机的打印单元精度不同
Private Const PHYSICALWIDTH = 110   'Physical Width in device units
Private Const PHYSICALHEIGHT = 111  'Physical Height in device units
Private Const PHYSICALOFFSETX = 112 'Physical Printable Area x margin
Private Const PHYSICALOFFSETY = 113 'Physical Printable Area y margin
Private Const LOGPIXELSX = 88 'Number of pixels per logical inch along the screen width
Private Const LOGPIXELSY = 90
Private Const SCALINGFACTORX = 114  'Scaling factor x
Private Const SCALINGFACTORY = 115  'Scaling factor y
Private Const DRIVERVERSION = 0     'Device driver version

Public Function SetPrinterPaper(ByVal lngHwnd As Long, ByVal intWidth As Integer, ByVal intHeight As Integer, _
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
        If lngSize = IDOK Then SetPrinterPaper = True
        'Reset the DEVMODE for the DC.
        lngSize = ResetDC(lngPrtDC, arrDevMode(1))
        If lngSize = 0 Then SetPrinterPaper = False
        
        'Close the handle when you are finished with it.
        Call ClosePrinter(lngHandle)
    End If
End Function

Private Function IsWindowsNT() As Boolean
'功能：是否WindowNT操作系统
    Const dwMaskNT = &H2&
    IsWindowsNT = (GetWinPlatform() And dwMaskNT)
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

Private Function GetRegValueBinary(ByVal hKey As REGRoot, ByVal strSubKey As String, ByVal strValueName As String, arrData() As Byte) As Boolean
'功能：获取注册表中指定位置的二进制值
    Dim lngKey As Long, lngReturn As Long
    Dim lngLength As Long
    
    lngReturn = RegOpenKeyEx(hKey, strSubKey, 0, KEY_QUERY_VALUE, lngKey)
    If lngReturn <> ERROR_SUCCESS Then
        Exit Function
    End If
    
    lngReturn = RegQueryValueEx_BINARY(lngKey, strValueName, 0, REG_BINARY, ByVal 0, lngLength)
    If lngReturn <> ERROR_SUCCESS Then
        RegCloseKey lngKey
        Exit Function
    End If
    
    ReDim arrData(lngLength - 1)
    lngReturn = RegQueryValueEx_BINARY(lngKey, strValueName, 0, REG_BINARY, arrData(0), lngLength)
    If lngReturn <> ERROR_SUCCESS Then
        RegCloseKey lngKey
        Exit Function
    End If
    
    RegCloseKey lngKey
    GetRegValueBinary = True
End Function

Private Function SetRegValueBinary(ByVal hKey As REGRoot, ByVal strSubKey As String, ByVal strValueName As String, arrData() As Byte) As Boolean
'功能：设置注册表中指定位置的二进制值
'说明：
'  1.当注册表项不存在时会自动创建
'  2.如果注册表项是其他类型会变为二进制类型

    Dim lngKey As Long, lngReturn As Long
    
    lngReturn = RegOpenKeyEx(hKey, strSubKey, 0, KEY_SET_VALUE, lngKey)
    If lngReturn <> ERROR_SUCCESS Then
        Exit Function
    End If
    
    lngReturn = RegSetValueEx_BINARY(lngKey, strValueName, 0, REG_BINARY, arrData(0), UBound(arrData) + 1)
    If lngReturn <> ERROR_SUCCESS Then
        RegCloseKey lngKey
        Exit Function
    End If
    
    RegCloseKey lngKey
    SetRegValueBinary = True
End Function

Private Function DeleteRegValue(ByVal hKey As REGRoot, ByVal strSubKey As String, ByVal strValueName As String) As Boolean
'功能：删除注册表中指定位置的项目
    Dim lngLength As Long, lngReturn As Long
    Dim lngKey As Long, lngType As Long
    
    
    lngReturn = RegOpenKeyEx(hKey, strSubKey, 0, KEY_SET_VALUE, lngKey)
    If lngReturn <> ERROR_SUCCESS Then
        Exit Function
    End If
    
    lngReturn = RegDeleteValue(lngKey, strValueName)
    If lngReturn <> ERROR_SUCCESS Then
        RegCloseKey lngKey
        Exit Function
    End If
    
    RegCloseKey lngKey
    DeleteRegValue = True
End Function

'初始化
Public Sub InitInterface(ByVal cnOracle As ADODB.Connection, _
                       ByVal strDBUser As String, _
                       ByVal lngSys As Long, _
                       ByVal lngModule As Long, _
                       ByVal blnMoved As Boolean)
'cnOracle: 主程序的数据库连接
'strDBUser: 当前登录数据库用户名
'lngModule: 需要执行的功能序号
'blnMoved: 是否被转储
    Dim i As Integer
    
    For i = 0 To Printers.Count - 1
        If Printers(i).DeviceName = "TinyPDF" Then
            Set Printer = Printers(i): Exit For
        End If
    Next
    
    Set mcnOracle = cnOracle
    mlngSys = lngSys
    mstrDBUser = strDBUser
    mlngModule = lngModule
    mblnMoved = blnMoved
End Sub

Public Sub PDFConvertEX(ByVal lngAdviceID As Long, ByVal strPDFFile As String)
'功能:转换成PDF格式文件到指定位置
'lngAdviceID:医嘱ID
'strPDFFile:输出文件路径,包括文件名和扩展名
    Dim objFile As New FileSystemObject
    Dim lngReportID As Long
    
On Error GoTo ErrorHand
    If lngAdviceID = 0 Then Exit Sub
    
    If strPDFFile <> "" Then
        If Not objFile.FolderExists(objFile.GetParentFolderName(strPDFFile)) Then
            MakeSureDirectoryPathExists objFile.GetParentFolderName(strPDFFile) & "\"
        End If
    End If
    
    Call InitVariable(lngAdviceID, lngReportID)
        
    '开始转换成PDF
    If mbln使用自定义报表 Then
        Call zlReportConvert(lngAdviceID, lngReportID) '使用自定义报表
    Else                                               '使用编辑模式
        Call zlEPRPConvert(lngAdviceID)
        
        If objFile.FileExists(objFile.GetParentFolderName(strPDFFile) & "\Temp.pdf") Then
            objFile.DeleteFile objFile.GetParentFolderName(strPDFFile) & "\Temp.pdf"
        End If
        
        '使用编辑模式转换时，由于转换一份后资源未释放，导致下一次转换时文件保存路径用了上一次的路径，
        '从而覆盖上一次转换的文件,因此运用临时文件Temp.pdf 中转
        InitPDF objFile.GetParentFolderName(strPDFFile) & "\Temp.pdf"

        Call zlEPRPConvert(lngAdviceID)
    End If
    
    '由于批量转换时报表编号不一定相同，所以每转换完一报告后需将其打印机恢复为打印时的打印机，避免影响打印设置
    SaveSetting "ZLSOFT", "私有模块\" & mstrDBUser & "\zl9Report\LocalSet\" & mstr报表编号, "Printer", mstrPrinterDeviceName
    
    '恢复PDF设置
    Call TinyPDF_Reset
    Exit Sub
ErrorHand:
    If mobjComLib.ErrCenter = 1 Then Resume
End Sub

Public Sub PDFConvert(ByVal lngAdviceID As Long, ByVal strPDFFile As String)
'功能:转换成PDF格式文件到指定位置
'lngAdviceID:医嘱ID
'strPDFFile:输出文件路径,包括文件名和扩展名
    Dim i As Integer
    Dim objFile As New FileSystemObject
    
On Error GoTo ErrorHand
    If lngAdviceID = 0 Then Exit Sub
    
    If strPDFFile <> "" Then
        If Not objFile.FolderExists(objFile.GetParentFolderName(strPDFFile)) Then
            MakeSureDirectoryPathExists objFile.GetParentFolderName(strPDFFile) & "\"
        End If
    End If
    
    Call ConvertToPDF(lngAdviceID)
    
    '由于批量转换时报表编号不一定相同，所以每转换完一报告后需将其打印机恢复为打印时的打印机，避免影响打印设置
    SaveSetting "ZLSOFT", "私有模块\" & mstrDBUser & "\zl9Report\LocalSet\" & mstr报表编号, "Printer", mstrPrinterDeviceName
    
    '恢复PDF设置
    Call TinyPDF_Reset

    Exit Sub
ErrorHand:
    If mobjComLib.ErrCenter = 1 Then Resume
End Sub

Private Sub ConvertToPDF(ByVal lngAdviceID As Long)
    '功能：转换成pdf格式文件
    Dim lngReportID As Long

On Error GoTo ErrorHand
    '初始化各变量
    Call InitVariable(lngAdviceID, lngReportID)
    
     '开始转换成PDF
    If mbln使用自定义报表 Then
        Call zlReportConvert(lngAdviceID, lngReportID) '使用自定义报表
    Else
        Call zlEPRPConvert(lngAdviceID)                '使用编辑模式
    End If
    
    Exit Sub
ErrorHand:
    If mobjComLib.ErrCenter = 1 Then Resume
End Sub

Private Sub InitVariable(ByVal lngAdviceID As Long, ByRef lngReportID As Long)
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim lngFileID As Long
    Dim str选中报表格式 As String
    Dim str报表编号 As String
    Dim strRegPath As String
    
On Error GoTo ErrorHand
    
    strSql = "select 报告打印 from 影像检查记录 where 医嘱id=[1]"
    Set rsTemp = mobjComLib.zlDatabase.OpenSQLRecord(strSql, "", lngAdviceID)
    
    strRegPath = "公共模块\ZL9PACSWORK\frmReport"
    str选中报表格式 = GetSetting("ZLSOFT", strRegPath, "选中报表格式", "")
    str报表编号 = GetSetting("ZLSOFT", strRegPath, "报表编号", "")

'    If mobjCustomReport Is Nothing Then Set mobjCustomReport = New clsReport
    
    '查询病历文件ID
    strSql = "Select 病历ID From 病人医嘱报告 Where 医嘱ID= [1]"
    If mblnMoved = True Then strSql = Replace(strSql, "电子病历记录", "H电子病历记录")
    Set rsTemp = mobjComLib.zlDatabase.OpenSQLRecord(strSql, "", lngAdviceID)
    
    If rsTemp.EOF Then
         '如果没有查到记录，说明病人还没有报告，需要根据诊疗项目创建报告
        strSql = "Select l.病人来源, a.病历文件id" & vbNewLine & _
            "From 病人医嘱记录 l, 病历单据应用 a" & vbNewLine & _
            "Where l.诊疗项目id = a.诊疗项目id(+) And a.应用场合(+) = Decode(l.病人来源, 2, 2, 4 ,4, 1) And l.Id = [1]"
        Set rsTemp = mobjComLib.zlDatabase.OpenSQLRecord(strSql, "", lngAdviceID)
        
        If rsTemp.EOF = True Then
            lngFileID = 0
        Else
            lngFileID = Nvl(rsTemp!病历文件id, 0)
        End If
        lngReportID = 0
    Else
        lngReportID = rsTemp!病历ID
        
         '查找设计格式的文件ID
        strSql = "Select 文件ID,科室ID,创建人,保存人,签名级别,保存时间 From 电子病历记录  Where Id =[1]"
        If mblnMoved = True Then strSql = Replace(strSql, "电子病历记录", "H电子病历记录")
        Set rsTemp = mobjComLib.zlDatabase.OpenSQLRecord(strSql, "", lngReportID)
        
        lngFileID = rsTemp!文件ID
    End If
    
    '先判断是否使用自定义报表
    strSql = "Select 通用,编号 From 病历文件列表  Where Id =[1]"
    Set rsTemp = mobjComLib.zlDatabase.OpenSQLRecord(strSql, "提取报告打印方式", lngFileID)
    
    If rsTemp.EOF = False Then
        If Nvl(rsTemp!通用) = 2 Then
            mbln使用自定义报表 = True     '使用自定义报表格式打印
        Else
            mbln使用自定义报表 = False    '使用编辑格式打印
            '使用编辑格式打印，清空自定义报表格式的设置
            str选中报表格式 = ""
            str报表编号 = ""
        End If
        
        If str报表编号 <> "ZLCISBILL" & Format(Nvl(rsTemp!编号), "00000") & "-2" Then str选中报表格式 = ""
        str报表编号 = "ZLCISBILL" & Format(Nvl(rsTemp!编号), "00000") & "-2"
    End If
    
    mstr报表编号 = str报表编号
    mstr选中报表格式 = str选中报表格式
    
    '设置打印机之前获取原来的打印机设置，转换完后恢复原来的设置，主要是避免打印时的打印机被转换时的打印机替代
    mstrPrinterDeviceName = GetSetting("ZLSOFT", "私有模块\" & mstrDBUser & "\zl9Report\LocalSet\" & mstr报表编号, "Printer", "")
    '将打印机设置为TinyPDF打印机
    SaveSetting "ZLSOFT", "私有模块\" & mstrDBUser & "\zl9Report\LocalSet\" & mstr报表编号, "Printer", "TinyPDF"
    
    Exit Sub
ErrorHand:
    If mobjComLib.ErrCenter = 1 Then Resume
End Sub

'使用编辑模式，调用病历的打印过程转换
Private Sub zlEPRPConvert(ByVal lngAdviceID As Long)
    Dim cbrToolBar As CommandBar
    Dim cbrControl As CommandBarControl

On Error GoTo ErrorHand
    
    mobjRchEPR.InitRichEPR mcnOracle, frmObj, mlngSys, False
    
    mobjReport.zlRefresh 0, 0
    mobjReport.zlRefresh lngAdviceID, 0, , , True
    
    frmObj.cbrMain.DeleteAll
    Set cbrToolBar = frmObj.cbrMain.Add("Main", xtpBarLeft)
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, 122, "Print")
    
    mobjReport.ConfigPrint "TinyPDF", 1
    mobjReport.zlExecuteCommandBars cbrControl
    
    Exit Sub
ErrorHand:
    If mobjComLib.ErrCenter = 1 Then Resume
End Sub

'使用自定义报表转换
Private Sub zlReportConvert(ByVal lngAdviceID As Long, ByVal lngReportID As Long)
    Dim i As Integer, j As Integer
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strExseNo As String
    Dim intExseKind As Integer
    Dim intPCount As Integer
    Dim strPicPath As String
    Dim strPicFile As String
    Dim objFile As New Scripting.FileSystemObject
    Dim oPicture As StdPicture
    Dim aryPrintPara(19) As String, strFlagString As String '实际传给自定义报表的内容
    Dim aryPara(19) As String, aryFlagPara(1) As String     '报告图中的图像记录
    Dim dcmImages As New DicomImages, dcmResultImage As DicomImage
    Dim arr报表格式() As String
    Dim int格式号 As Integer
    Dim intRows As Integer, intCols As Integer
    Dim intParaCount As Integer
 
On Error GoTo ErrorHand

    strSql = "Select 记录性质, No From 病人医嘱发送 Where 医嘱id = [1]"
    If mblnMoved = True Then strSql = Replace(strSql, "病人医嘱发送", "H病人医嘱发送")
    Set rsTemp = mobjComLib.zlDatabase.OpenSQLRecord(strSql, "提前记录性质和No", lngAdviceID)
    
    If rsTemp.RecordCount = 0 Then Exit Sub
    
    strExseNo = "" & rsTemp!NO
    intExseKind = Val("" & rsTemp!记录性质)
    
    '获取图像
    strPicPath = App.Path & "\TmpImage\"
    If objFile.FolderExists(strPicPath) = False Then objFile.CreateFolder strPicPath
    
    '获取报告图像（包括标记图）生成本地文件
    '一个报告表格中可能排列多个报告图
    intPCount = 0
    strSql = "Select Id As 表格Id From 电子病历内容" & vbNewLine & _
        "       Where 文件id = [1] And 对象类型 = 3 And Substr(对象属性, Instr(对象属性, ';', 1, 18) + 1, 1) = '2'" & vbNewLine & _
        "       Order By 对象序号"
    If mblnMoved = True Then strSql = Replace(strSql, "电子病历内容", "H电子病历内容")
    
    Set rsTemp = mobjComLib.zlDatabase.OpenSQLRecord(strSql, "提取图像", lngReportID)
    
    Do While Not rsTemp.EOF
        Set mobjTable = CreateObject("zlRichEPR.cEPRTable")
        If mobjTable.GetTableFromDB(3, lngReportID, Val("" & rsTemp!表格Id)) Then
            For i = 1 To mobjTable.Pictures.Count
                strPicFile = "PACSPic" & i & ".JPG"
                If objFile.FileExists(strPicFile) Then objFile.DeleteFile strPicFile, True
                If mobjTable.Pictures(i).PictureType = 1 Then
                    Set oPicture = mobjTable.Pictures(i).DrawFinalPic
                Else
                    Set oPicture = mobjTable.Pictures(i).OrigPic
                End If
                SavePicture oPicture, strPicFile
                If objFile.FileExists(strPicFile) Then
                    '保存标记图和图象的路径
                    If mobjTable.Pictures(i).PictureType = 1 Then
                        aryFlagPara(0) = strPicFile
                    Else
                        aryPara(intPCount) = strPicFile
                        dcmImages.AddNew
                        dcmImages(dcmImages.Count).FileImport strPicFile, "BMP"
                        intPCount = intPCount + 1
                        If intPCount > UBound(aryPara) Then Exit Do
                    End If
                End If
            Next i
        End If
        
        Set mobjTable = Nothing
        rsTemp.MoveNext
    Loop
    
    '根据选择的自定义报表格式，组织图像
    '如果只选择了一种格式，则检查是否只有一个图象框,只有一个图像框的时候，自动组合图像。
    '如果选择了2种以上的格式，则对只有一个图像框的情况不作自动组合
    arr报表格式 = Split(mstr选中报表格式, ",")
    '处理没有选择格式的情况
    If UBound(arr报表格式) = -1 Then
        ReDim arr报表格式(0) As String
        arr报表格式(0) = "1-1"
    End If
    
    If UBound(arr报表格式) = 0 Then     '只有一种格式
        int格式号 = Split(arr报表格式(0), "-")(0)
        
        strSql = "Select b.名称,b.W,b.H From zlReports a, zlRptItems b" & vbNewLine & _
        "       Where a.Id = b.报表id And a.编号 = [1] And Nvl(b.下线, 0) = 1 And b.类型 = 11 And b.格式号 = [2] And b.名称 not like '标记%'"
        Set rsTemp = mobjComLib.zlDatabase.OpenSQLRecord(strSql, "查询是否需要组合图像", mstr报表编号, int格式号)
        
        If rsTemp.RecordCount = 1 And intPCount >= 1 Then
            '组合图象
            ResizeRegion intPCount, rsTemp("W"), rsTemp("H"), intRows, intCols
            Set dcmResultImage = AssembleImage(dcmImages, intRows, intCols, rsTemp("H"), rsTemp("W"))
            dcmResultImage.FileExport Right(aryPara(0), Len(aryPara(0)) - InStr(aryPara(0), "=")), "JPEG"
        End If
    End If
    
    '获取图像，调用报表
    intPCount = 0       '记录图像的数量
    
    For i = 0 To UBound(arr报表格式)
        int格式号 = Split(arr报表格式(i), "-")(0)
        
        strSql = "Select b.名称 From zlReports a, zlRptItems b" & vbNewLine & _
        "       Where a.Id = b.报表id And a.编号 = [1] And Nvl(b.下线, 0) = 1 And b.类型 = 11 And b.格式号 = [2]" & vbNewLine & _
        "       Order By b.名称" 'Trunc(b.y/567),Trunc(b.x/567)
        Set rsTemp = mobjComLib.zlDatabase.OpenSQLRecord(strSql, "提取图象框", mstr报表编号, int格式号)
        
        '装载图像数据
        intParaCount = 0
        Do While Not rsTemp.EOF
            '分别装在标记图和报告图
            If InStr(rsTemp!名称, "标记") <> 0 Then '标记图
                If aryFlagPara(0) <> "" Then strFlagString = rsTemp!名称 & "=" & aryFlagPara(0)
            Else    '报告图
                If intPCount > UBound(aryPara) Then Exit Do     '图像数量超过报告中的图像，退出
                If aryPara(intPCount) = "" Then Exit Do         '报表中的图象框比报告中的多，退出
                
                aryPrintPara(intParaCount) = rsTemp!名称 & "=" & aryPara(intPCount)
                intPCount = intPCount + 1
                intParaCount = intParaCount + 1
            End If
            rsTemp.MoveNext
        Loop
        
        '处理报表中图形比报告中少的情况
        For j = intParaCount To UBound(aryPrintPara)
            If aryPrintPara(j) Like "*=*" Then aryPrintPara(j) = ""
        Next j
        
        '调用报表
        Call mobjCustomReport.ReportOpen(mcnOracle, mlngSys, mstr报表编号, Nothing, _
            "NO=" & strExseNo, "性质=" & intExseKind, "医嘱ID=" & lngAdviceID, strFlagString, _
            aryPrintPara(0), aryPrintPara(1), aryPrintPara(2), aryPrintPara(3), aryPrintPara(4), aryPrintPara(5), _
            aryPrintPara(6), aryPrintPara(7), aryPrintPara(8), aryPrintPara(9), aryPrintPara(10), aryPrintPara(11), _
            aryPrintPara(12), aryPrintPara(13), aryPrintPara(14), aryPrintPara(15), aryPrintPara(16), aryPrintPara(17), _
            aryPrintPara(18), aryPrintPara(19), "ReportFormat=" & int格式号, 2)
    Next i
    Exit Sub
ErrorHand:
    If mobjComLib.ErrCenter = 1 Then Resume
End Sub

Public Sub InitPDF(ByVal strPDFFile As String, _
    Optional ByVal blnCopyable As Boolean, _
    Optional ByVal blnEditable As Boolean, _
    Optional ByVal blnPrintable As Boolean, _
    Optional ByVal strPassword As String, _
    Optional ByVal strAttachs As String)
'功能：指定TinyPDF打印机输出参数
'参数：strPDFFile=输出文件名，包含文件路径和文件扩展名
'                 文件路径必须存在，且会自动覆盖同名文件
'                 如果未指定，则弹出文件保存对话框
'      blnCopyable=输出的PDF文件是否允许复制内容
'      blnEditable=输出的PDF文件是否允许编辑内容
'      blnPrintable=输出的PDF文件是否允许打印输出
'      strPassword=是否要求输入密码
'      strAttachs=要加到PDF中的附件文件名(包含路径),多个以"|"分隔
'注意：该函数需要在Printer的任何打印设置之前调用(包括API方式设置)
    Dim arrData() As Byte
    Dim intSect As Integer, intAdr As Integer
    Dim intTag As Integer, strFile As String
    Dim i As Integer, j As Integer
    Dim strWord As String
    
    '读取设置
    If Not GetRegValueBinary(HKEY_CURRENT_USER, "Printers\DevModePerUser", "TinyPDF", arrData) Then Exit Sub
    If Not mblnReset Then marrReset = arrData: mblnReset = True
    
    '更改设置
    arrData(Val("&H00E0")) = &H0 '页边距
    arrData(Val("&H00E1")) = &H0 '页边距
    arrData(Val("&H00E4")) = &H0 '不自动打开
    arrData(Val("&H011C")) = &H1 '嵌入所有字体
    arrData(Val("&H0130")) = &H0 'RGB颜色(sRGB不正常)
    
    If strPassword <> "" Then
        arrData(Val("&H013C")) = &H1 '有用户密码
        For i = 1 To Len(strPassword)
            arrData(Val("&H0140") + i - 1) = Asc(Mid(strPassword, i, 1))
        Next
        arrData(Val("&H0140") + i - 1) = &H0
    Else
        arrData(Val("&H013C")) = &H0 '无用户密码
        arrData(Val("&H0140")) = &H0
    End If
    
    arrData(Val("&H0164")) = &H1 '进行安全控制
    arrData(Val("&H0168")) = &H0  '安全控制密码为空
    If blnPrintable Then
        arrData(Val("&H0189")) = &H2  '允许高分辨率打印
    Else
        arrData(Val("&H0189")) = &H0  '不允许打印
    End If
    If blnEditable Then
        arrData(Val("&H018A")) = &H4  '除提取页面之外的任何内容
    Else
        arrData(Val("&H018A")) = &H0  '不允许更改
    End If
    If blnCopyable Then
        arrData(Val("&H018C")) = &H1  '允许复制内容
    Else
        arrData(Val("&H018C")) = &H0  '不允许复制
    End If
    arrData(Val("&H0190")) = &H1  '不允许复制时，允许屏幕阅读器设备视觉受损地访问文本
    If strPDFFile <> "" Then
        arrData(Val("&H0194")) = &H2  '指定文件输出(包含路径)
    Else
        arrData(Val("&H0194")) = &H0  '提示输出
    End If
    arrData(Val("&H01A0")) = &H1  '直接覆盖文件
    
    '数据段：输出文件、附件文件
    arrData(Val("&H01C8")) = &H0
    arrData(Val("&H01C8") + 1) = &H0
    intAdr = Val("&H01CA")
    intSect = 1 '数据段序号
    intTag = 1 '1-数据内容,2-单项结束
    Do While intAdr <= 4552
        If intSect = 1 Or intSect = 2 Then '嵌入/不嵌入字体段
            If arrData(intAdr) = 0 And arrData(intAdr + 1) = 0 Then
                If intTag = 1 Then
                    intTag = 2
                ElseIf intTag = 2 Then
                    intTag = 1
                    intSect = intSect + 1
                End If
            Else
                intTag = 1
            End If
            intAdr = intAdr + 2
        ElseIf intSect = 3 Then '中间间隔段
            If arrData(intAdr) = 0 Then
                intAdr = intAdr + 1
            Else
                intSect = intSect + 1
            End If
        ElseIf intSect = 4 Or intSect = 5 Then 'RGB/CMYK配置文件段
            If arrData(intAdr) = 0 Then
                intSect = intSect + 1
            End If
            intAdr = intAdr + 1
        ElseIf intSect = 6 Then '输出目录段
            strWord = Hex(intAdr - Val("&H01C8"))
            strWord = String(4 - Len(strWord), "0") & strWord
            arrData(Val("&H0198")) = Val("&H" & Right(strWord, 2)) '低位字节
            arrData(Val("&H0198") + 1) = Val("&H" & Left(strWord, 2)) '高位字节
            
            arrData(intAdr) = 0
            arrData(intAdr + 1) = 0
            intAdr = intAdr + 2
            intSect = intSect + 1
        ElseIf intSect = 7 Then '输出文件段
            strWord = Hex(intAdr - Val("&H01C8"))
            strWord = String(4 - Len(strWord), "0") & strWord
            arrData(Val("&H019C")) = Val("&H" & Right(strWord, 2)) '低位字节
            arrData(Val("&H019C") + 1) = Val("&H" & Left(strWord, 2)) '高位字节
            
            If strPDFFile = "" Then
                arrData(intAdr) = 0
                arrData(intAdr + 1) = 0
                intAdr = intAdr + 2
            Else
                For i = 1 To Len(strPDFFile)
                    strWord = Hex(AscW(Mid(strPDFFile, i, 1)))
                    If Len(strWord) = 2 Then
                        strWord = "00" & strWord
                    End If
                    
                    arrData(intAdr + i * 2 - 2) = Val("&H" & Right(strWord, 2)) '低位Unicode
                    arrData(intAdr + i * 2 - 1) = Val("&H" & Left(strWord, 2)) '高位Unicode
                Next
                intAdr = intAdr + Len(strPDFFile) * 2
                
                arrData(intAdr) = 0
                arrData(intAdr + 1) = 0
                intAdr = intAdr + 2
            End If
            intSect = intSect + 1
        ElseIf intSect = 8 Then '中间间隔段
            strWord = Hex(intAdr - Val("&H01C8"))
            strWord = String(4 - Len(strWord), "0") & strWord
            arrData(Val("&H01A4")) = Val("&H" & Right(strWord, 2)) '低位字节
            arrData(Val("&H01A4") + 1) = Val("&H" & Left(strWord, 2)) '高位字节
            
            For i = 1 To 16
                arrData(intAdr + i - 1) = 0
            Next
            intAdr = intAdr + 16
            intSect = intSect + 1
        ElseIf intSect = 9 Then '附件文件
            '目前发现设置加载附件会导致生成的PDF打开出错
            If strAttachs = "" Then
                arrData(intAdr) = 0
                arrData(intAdr + 1) = 0
                intAdr = intAdr + 2
            Else
                For i = 0 To UBound(Split(strAttachs, "|"))
                    strFile = Split(strAttachs, "|")(i)
                    For j = 1 To Len(strFile)
                        strWord = Hex(AscW(Mid(strFile, j, 1)))
                        If Len(strWord) = 2 Then
                            strWord = "00" & strWord
                        End If
                        
                        arrData(intAdr + j * 2 - 2) = Val("&H" & Right(strWord, 2)) '低位Unicode
                        arrData(intAdr + j * 2 - 1) = Val("&H" & Left(strWord, 2)) '高位Unicode
                    Next
                    intAdr = intAdr + Len(strFile) * 2
                    
                    arrData(intAdr) = 0
                    arrData(intAdr + 1) = 0
                    intAdr = intAdr + 2
                Next
            End If
            '结束退出
            Exit Do
        End If
    Loop
    
    For i = Val("&H01A8") To Val("&H01C4") Step 4
        strWord = Hex(arrData(Val("&H01A4")) + arrData(Val("&H01A4") + 1) * 256 + (i - Val("&H01A8")) / 2 + 2)
        strWord = String(4 - Len(strWord), "0") & strWord
        
        arrData(i) = Val("&H" & Right(strWord, 2)) '低位字节
        arrData(i + 1) = Val("&H" & Left(strWord, 2)) '高位字节
        arrData(i + 2) = 0
        arrData(i + 3) = 0
    Next
    
    '保存设置
    SetRegValueBinary HKEY_CURRENT_USER, "Printers\DevModePerUser", "TinyPDF", arrData
    SetRegValueBinary HKEY_CURRENT_USER, "Printers\DevModes2", "TinyPDF", arrData
End Sub

Private Sub TinyPDF_Reset()
'功能：重置TinyPDF打印机输出参数设置
'说明：该函数在打印输出完成后调用
    If mblnReset Then
        SetRegValueBinary HKEY_CURRENT_USER, "Printers\DevModePerUser", "TinyPDF", marrReset
        SetRegValueBinary HKEY_CURRENT_USER, "Printers\DevModes2", "TinyPDF", marrReset
        Erase marrReset: mblnReset = False
    Else
        DeleteRegValue HKEY_CURRENT_USER, "Printers\DevModePerUser", "TinyPDF"
        DeleteRegValue HKEY_CURRENT_USER, "Printers\DevModes2", "TinyPDF"
    End If
End Sub

Private Sub ResizeRegion(ByVal ImageCount As Integer, ByVal RegionWidth As Long, _
    ByVal RegionHeight As Long, Rows As Integer, Cols As Integer)
'-----------------------------------------------------------------------------
'功能：根据输入的图像数量，图像区域的宽度和高度，返回最佳的图像排列行数和列数
'参数： ImageCount－－图像数量
'       RegionWidth--图像显示区域的宽度
'       RegionHeight--图像显示区域的高度
'       Rows－－[返回]最佳行数
'       Cols－－[返回]最佳列数
'返回：返回最佳行数Rows，最佳列数Cols
'-----------------------------------------------------------------------------
    Dim iCols As Integer, iRows As Integer
    Dim iBase As Integer, blnDoLoop As Integer
    Dim lngFreeCount As Long
    
    If RegionHeight = 0 Then RegionHeight = 1
    If RegionWidth = 0 Then RegionWidth = 1
    
    On Error GoTo err
    iCols = CInt(Sqr(ImageCount * RegionWidth / RegionHeight))
    iRows = CInt(Sqr(ImageCount * RegionHeight / RegionWidth))

    If iRows < 1 Then iRows = 1
    If iCols < 1 Then iCols = 1
    
    '当图像格式为如下等形式时，需要对行列进行修正
    
    '格式1：
    '图1  图2  图3  图4
    '图5  图6  图7  图8
    '空1  空2  空3  空4
    
    '格式2：
    '图1  图2  图3  图4
    '图5  图6  图7  图8
    '图9  空1  空2  空3
    
    lngFreeCount = iRows * iCols - ImageCount
    Do While lngFreeCount >= iCols Or lngFreeCount >= iRows
        If lngFreeCount >= iCols Then
            iRows = iRows - 1
        Else
            iCols = iCols - 1
        End If
        
        lngFreeCount = iRows * iCols - ImageCount
    Loop
    
    If iRows < 1 Then iRows = 1
    If iCols < 1 Then iCols = 1
    
    Do While iRows * iCols < ImageCount
        If RegionWidth / iCols > RegionHeight > iRows Then
            iCols = iCols + 1
        Else
            iRows = iRows + 1
        End If
    Loop
    
    '再次修正行列数
    lngFreeCount = iRows * iCols - ImageCount
    Do While lngFreeCount >= iCols Or lngFreeCount >= iRows
        If lngFreeCount >= iCols Then
            iRows = iRows - 1
        Else
            iCols = iCols - 1
        End If
        
        lngFreeCount = iRows * iCols - ImageCount
    Loop
    
    Rows = iRows: Cols = iCols
err:
End Sub

Private Function AssembleImage(AssembleViewer As DicomImages, ByVal intRows As Integer, ByVal intCols As Integer, _
    ByVal lngHeight As Long, ByVal lngWidth As Long) As DicomImage

'组合viewer中的显示的所有图像成一个图像

    Dim Image As New DicomImage '新图像
    Dim imgs As New DicomImages '临时存储屏幕采集的图像集
    Dim intWidth As Integer     '新图像的宽度
    Dim intHeight As Integer    '新图像的高度
    Dim Simg As New DicomImage
    Dim sZoom As Single
    Dim intImgRectWidth As Integer  '单张图像可占用的区域宽度
    Dim intImgRectHeight As Integer '单张图像可占用的区域高度
    Dim i As Integer
    Dim intMaxWidth As Integer      '拼接后图像的最大宽度
    Dim intMaxHeight As Integer     '拼接后图像的最大高度
    Dim intBorder As Integer        '图像之间的边距
    Dim intOffsetX As Integer       '拼接时X方向的位移
    Dim intOffsetY As Integer       '拼接时Y方向的位移
    Dim lngWhiteX As Long           '将图象底色改成白色的X宽度
    Dim lngWhiteY As Long           '将图象底色改成白色的Y高度
    
    If AssembleViewer.Count <= 0 Then
        '返回一个黑图**************
        Exit Function
    End If

    On Error GoTo err
    '计算新图像的宽度和高度

    '新图像的宽度和高度不能够大于intMaxWidth×intMaxHeight（宽度×高度）
    intMaxWidth = 3073
    intMaxHeight = 3073
    intBorder = 10

    intImgRectWidth = 0
    intImgRectHeight = 0

    '估算新图像的宽度和高度

    '使用原图像的宽度和高度和，并用Viewer的比例来修正。

    '估算图像的新宽高
    For i = 1 To AssembleViewer.Count
        If intImgRectWidth < AssembleViewer(i).SizeX Then intImgRectWidth = AssembleViewer(i).SizeX
        If intImgRectHeight < AssembleViewer(i).SizeY Then intImgRectHeight = AssembleViewer(i).SizeY
    Next i
    
    '计算横向和纵向图像数量
    intWidth = intImgRectWidth * intCols
    intHeight = intImgRectHeight * intRows
    
    '修正图像的宽高，不能大于最大值
    '如果大于intMaxWidth×intMaxHeight则，按照图像总长宽比，使用小于等于intMaxWidth×intMaxHeight作为新宽高,
    If intWidth > intMaxWidth Or intHeight > intMaxHeight Then
        If intHeight / intWidth > intMaxHeight / intMaxWidth Then
            intWidth = intWidth / intHeight * intMaxHeight
            intHeight = intMaxHeight
        Else
            intHeight = intHeight / intWidth * intMaxWidth
            intWidth = intMaxWidth
        End If
    End If
    
    '采集图像
    '将图像采集到临时图像集
    For i = 1 To AssembleViewer.Count
        '计算缩放比例 hj修改,解决多图合并时，放大的图象无法真正放大的问题
        sZoom = intImgRectHeight / AssembleViewer(i).SizeY
        If sZoom > intImgRectWidth / AssembleViewer(i).SizeX Then
            sZoom = intImgRectWidth / AssembleViewer(i).SizeX
        End If
        
        AssembleViewer(i).StretchToFit = False
        AssembleViewer(i).Zoom = sZoom
        '采集图像
        Set Simg = AssembleViewer(i).PrinterImage(8, 3, True, sZoom, 0, AssembleViewer(i).SizeX, 0, AssembleViewer(i).SizeY)
        imgs.Add Simg
    Next i

    '精确计算新图像的宽度和高度
    intImgRectWidth = 0
    intImgRectHeight = 0

    For i = 1 To imgs.Count
        If intImgRectWidth < imgs(i).SizeX Then intImgRectWidth = imgs(i).SizeX
        If intImgRectHeight < imgs(i).SizeY Then intImgRectHeight = imgs(i).SizeY
        imgs(i).Attributes.Add &H8, &H16, "doSOP_SecondaryCapture"
    Next i
    intImgRectWidth = intImgRectWidth + intBorder
    intImgRectHeight = intImgRectHeight + intBorder
    intWidth = intImgRectWidth * intCols
    intHeight = intImgRectHeight * intRows

    '创建新图像
    Image.Name = "print"
    Image.PatientID = "print001"
    
    Image.Attributes.Add &H8, &H16, doSOP_SecondaryCapture
    Image.Attributes.Add &H28, &H2, 3 ' samples/pixel
    Image.Attributes.Add &H28, &H4, "RGB" ' photometric interpreation  'CT都是MONOCHROME2,CR都是MONOCHROME1？
    Image.Attributes.Add &H28, &H10, intHeight  'x,Rows
    Image.Attributes.Add &H28, &H11, intWidth 'Y,Columns
    Image.Attributes.Add &H28, &H100, 8 'bits allocated
    Image.Attributes.Add &H28, &H101, 8 ' bits stored
    Image.Attributes.Add &H28, &H102, 7 ' high bit
    ReDim pix(intWidth * 3, intHeight * 3) As Byte
    For lngWhiteX = 0 To intWidth * 3
        For lngWhiteY = 0 To intHeight * 3
            pix(lngWhiteX, lngWhiteY) = 255
        Next lngWhiteY
    Next lngWhiteX
    Image.Attributes.Add &H7FE0, &H10, pix

    '拼接新图像
    For i = 1 To imgs.Count
        '计算图像内位移
        intOffsetX = (intImgRectWidth - imgs(i).SizeX - intBorder) / 2
        intOffsetY = (intImgRectHeight - imgs(i).SizeY - intBorder) / 2
        Image.Blt imgs(i), 0, 0, ((i - 1) Mod intCols) * intImgRectWidth + intOffsetX, ((i - 1) \ intCols) * intImgRectHeight + intOffsetY, imgs(i).SizeX, imgs(i).SizeY, 1, 1, 1, False
    Next i

    Set AssembleImage = Image
    Exit Function
err:
End Function

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Private Sub Class_Initialize()
    Set mobjCustomReport = CreateObject("zl9Report.clsReport")   '自定义报表对象
    Set mobjReport = CreateObject("zlRichEPR.cDockReport")       '报告对象
    Set mobjRchEPR = CreateObject("zlRichEPR.cRichEPR")
    Set mobjComLib = CreateObject("zl9ComLib.clsComLib")
End Sub

Private Sub Class_Terminate()
    If mcnOracle.State = adStateOpen Then mcnOracle.Close
    Set mcnOracle = Nothing
    Set mobjCustomReport = Nothing
    Set mobjReport = Nothing
    Set mobjRchEPR = Nothing
    Set mobjTable = Nothing
    Set mobjComLib = Nothing
    
    Unload frmObj
End Sub
