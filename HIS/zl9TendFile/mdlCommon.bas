Attribute VB_Name = "mdlCommon"
Option Explicit
'**************************
'       OEM代号
'
'医业  D2BDD2B5
'**************************
Public gobjFSO As New FileSystemObject
Public gcnOracle As ADODB.Connection
Public gstrPrivs As String
Public mlngInitClsCount As Long

'错误日志处理相关变量
Private mlngErrNum As Long, mstrErrInfo As String, mbytErrType As Byte
Private mstrRecentSQL As String  '最近执行的SQL语句

Private Const madLongVarCharDefault As Integer = 10          '字符型字段缺省长度
Private Const madDoubleDefault As Integer = 18               '数字型字段缺省长度
Private Const madDbDateDefault As Integer = 20               '日期型字段缺省长度

'SQLLog变量
Private msngTime As Single
Private mobjLogText As TextStream

Public gobjFile As New FileSystemObject

Global Const gintTends% = 1                       '打印对象是zlTFPrintTends
Global gintObjType As Integer                    '打印对象是什么类型

Global Const gintMAX_SIZE% = 255                        '最大的字符缓冲区
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_ENUMERATE_Sub_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const READ_CONTROL = &H20000
Public Const SYNCHRONIZE = &H100000
Public Const KEY_SET_VALUE = &H2
Public Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_Sub_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Public Const EM_GETLINECOUNT = &HBA&            '获取行数。
Public Const EM_GETLINE = &HC4&                '发送一行文本到指定缓冲区。

Type OSVERSIONINFO 'for GetVersionEx API call
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
' 发送指定消息到窗体，等待处理完才返回；而 PostMessage() 函数发送消息，立即返回！HWND hWnd 目标窗体的句柄。wMsg待发送的消息。wParam消息第一参数。lParam消息第二参数。
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Global gstrUnitName As String     '安装Windows时填写的单位名

Global gobjOutTo As Object        '打印输出的目标对象,可能是printer或PictureBox
Global gobjSend As Object         '要打印的对象
Public arrFormat As Variant       '对象输出到Excel的列格式数组

Global gintRowTotal As Long    '总页数
Global gintColTotal As Long    '总页面数
'每页的第一行的行号与最后一行的行号；第一列的列号与最后一列的列号
Global gintRow() As Long
Global gintCol() As Long

Global gintPage As Integer        '当前显示的页码
Global gintCopies As Integer      '打印的份数

Global gintBegin As Integer       '起始页码
Global gintShow As Integer        '预览的页数

Global gsngTotalWidth As Single   '所有页面的总宽度


Global gsngTitle As Single        '标题的高度
Global gsngUpAppRow As Single     '表上项目的高度
Global gsngDownAppRow As Single   '表下项目的高度
Global gsngFixRow As Single       '固定行的高度
Global gsngFixCol As Single       '固定列的宽度

Global gsngScaleWidth As Single   '由于设置了纵向横向引起纸张实际打印的宽度
Global gsngScaleHeight As Single  '由于设置了纵向横向引起纸张实际打印的高度
Global gsngHeight As Single       '页面的有效高度
Global gsngWidth As Single        '页面的有效宽度
Global gsngPrintedWidth() As Single '每一页面实际打印了的宽度

Global gsngScale As Single        '缩放比例
Global gcolGrid As New Collection '已打印单元格的集合

Global gfrmTemp  As New frmSample  '已打印单元格的集合
'-------------------------------------------------------------
Global gstrHeader As String           '页眉内容
Global gsngHeader As Single           '页眉位置   '以毫米为单位
Global gstrFooter As String           '页脚内容
Global gsngFooter As Single           '页脚位置   '以毫米为单位
Global gsngPageWidth As Single        '纸张宽度   以绨为单位
Global gsngPageHeight As Single       '纸张高度   以绨为单位
Global gsngPageScaleWidth As Single   '纸张实际打印的宽度   以绨为单位
Global gsngPageScaleHeight As Single  '纸张实际打印的高度   以绨为单位
Global gintSize As Integer            '纸张的尺寸,自定义为256
Global gintOri As Integer             '纸张的进纸方向.2表示横向，1表示纵向

Global gsngUp As Single               '上边距   '以毫米为单位
Global gsngDown As Single             '下边距   '以毫米为单位
Global gsngLeft As Single             '左边距   '以毫米为单位
Global gsngRight As Single            '右边距   '以毫米为单位
Global gstrTabTitle As String         '标题内容
Global gstrTitleFName As String       '标题的字体名
Global gintTitleFSize As Integer      '标题的字体大小
Global gblnTitleFBold As Boolean      '标题是否粗体
Global gblnTitleFItalic As Boolean    '标题是否斜体
Global glngTitleColor As Long         '标题的颜色
Global gstrAppRowFName As String      '表项目的字体名
Global gintAppRowFSize As Integer     '表项目的字体大小
Global gblnAppRowFBold As Boolean     '表项目是否粗体
Global gblnAppRowFItalic As Boolean   '表项目是否斜体
Global glngAppRowColor As Long        '表项目的颜色
Global gintUpAppRow As Long           '表上项目的行数
Global gintDownAppRow As Long         '表下项目的行数
Global gintTotalRow As Long           '总行数
Global gintTotalCol As Long           '总列数
Global gintFixRow As Integer          '固定行号
Global gintFixCol As Integer          '固定列号

Global gintGroups As Long             '组数

Global gstrGrant As String           '"","正式","试用","测试"

Public glng文件ID As Long
Public glng病人ID As Long
Public glng主页ID As Long
Public glng病区ID As Long
Public gint婴儿 As Integer
Public gstrSQL As String
Public gblnMoved As Boolean
Public gblnOut As Boolean '病人是否已经出院
Public frmAsk As frmTendPrintAsk        '询问窗体
Public gstr对角线 As String             '保存列号序号
Public glngHideCols As Long             '保存表格固定列数
Public glngPrintRow As Long             '从此行开始打印
Public gblnPrintMode As Boolean         '打印模式为TRUE
Public gintPrintState As Integer        '打印模式，1-续打(数据页之前打印了部分)；2-正常打印(整页打印，可用于未打印的页，或已打印的页)
Public glngSignName As Integer          '保存签名人列
Public gblnBatch As Boolean             '批量打印
Public glngDate As Long                 '记录单日期列号
Public gstrCOLDateText As String        '日期内容(请勿随便使用)
Public glngCollectColor As Long         '小结标识颜色
Public Enum enuPage
    续打
    重打
End Enum

'WinNT自定义纸张控制================================================================
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
Public Declare Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal hWnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, pDevModeOutput As Any, pDevModeInput As Any, ByVal fMode As Long) As Long
Public Declare Function ResetDC Lib "gdi32" Alias "ResetDCA" (ByVal hDC As Long, lpInitData As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

'获取显示器或者打印机的信息
Public Const HORZRES = 8            '  Horizontal width in pixels

Public Const VERTRES = 10           '  Vertical width in pixels

Public Const LOGPIXELSX = 88        '  Logical pixels/inch in X

Public Const LOGPIXELSY = 90        '  Logical pixels/inch in Y

Public Const PHYSICALOFFSETX = 112 '  Physical Printable Area x margin

Public Const PHYSICALOFFSETY = 113 '  Physical Printable Area y margin

Public Const PHYSICALHEIGHT = 111 '  Physical Height in device units

Public Const PHYSICALWIDTH = 110 '  Physical Width in device units

Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long


'--------------------------------------------------------------
'ReadVar                将打印对象的数据读到变量中
'IsWindows95            判断是否在Windows95下工作
'GetWinPlatform         返回当前的系统版本代号
'StripTerminator        去掉字符串变量中的 Chr$(0)字符
'CalculateRC            为每一个单元格计算它的位置
'CalculateHeight        计算出标题、表上项目和表下项目的高度,固定行的高度、固定列的宽度
'PrintPage              在指定设备上打印指定页
'PrintHeadFoot          打印页眉页脚
'zlOutTabAppRow         输出listview表上或表下项目
'zlOutTabAppSet         输出网格的表上或表下项目
'zlOutTitle             输出标题
'OutRow                 输出一行文字
'ConvHF                 将页眉与页脚转换成实际打印的内容
'RealPrint              输出打印机上
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Public Function ReadVar() As Boolean
'功    能：将打印对象的数据读到变量中
'编 制 人：朱玉宝
'编制日期：2011-03-01
'参    数：无
'返    回：读入参数有效则返回真
    ReadVar = True
    Dim strUserName As String
    Dim lngOffsetLeft As Long, lngOffsetTop As Long
    On Error GoTo errHandle
    gsngPageWidth = Printer.Width
    gsngPageHeight = Printer.Height
    lngOffsetLeft = Printer.ScaleX(GetDeviceCaps(Printer.hDC, PHYSICALOFFSETX), vbPixels, vbTwips)  'PHYSICALOFFSETX=112
    lngOffsetTop = Printer.ScaleY(GetDeviceCaps(Printer.hDC, PHYSICALOFFSETY), vbPixels, vbTwips)   'PHYSICALOFFSETY=113
    gsngPageScaleHeight = Printer.Height - lngOffsetTop * 2
    gsngPageScaleWidth = Printer.Width - lngOffsetLeft * 2
'    gsngPageScaleWidth = Printer.ScaleWidth
'    gsngPageScaleHeight = Printer.ScaleHeight
    
    gintSize = Printer.PaperSize
    gintOri = Printer.Orientation
'    If gintOri = 1 Then '纵向
'        gsngScaleWidth = IIf(gsngPageScaleWidth < gsngPageScaleHeight, gsngPageScaleWidth, gsngPageScaleHeight) '文档打印以纸的窄边作顶部。
'        gsngScaleHeight = IIf(gsngPageScaleWidth > gsngPageScaleHeight, gsngPageScaleWidth, gsngPageScaleHeight)
'    Else
'        gsngScaleWidth = IIf(gsngPageScaleWidth > gsngPageScaleHeight, gsngPageScaleWidth, gsngPageScaleHeight) '文档打印以纸的宽边作顶部
'        gsngScaleHeight = IIf(gsngPageScaleWidth < gsngPageScaleHeight, gsngPageScaleWidth, gsngPageScaleHeight)
'    End If
    gsngScaleWidth = gsngPageScaleWidth
    gsngScaleHeight = gsngPageScaleHeight
    With gobjSend
        '保存标题的属性
        gstrTabTitle = .Title.Text
        gstrTitleFName = .Title.Font.Name
        gintTitleFSize = .Title.Font.Size
        gblnTitleFItalic = .Title.Font.Italic
        gblnTitleFBold = .Title.Font.Bold
        glngTitleColor = .Title.Color
        '保存表上项目与表下项目的属性
        gstrAppRowFName = .AppFont.Name
        gintAppRowFSize = .AppFont.Size
        gblnAppRowFItalic = .AppFont.Italic
        gblnAppRowFBold = .AppFont.Bold
        glngAppRowColor = .AppColor
        gintUpAppRow = .UnderAppRows.Count
        gintDownAppRow = .BelowAppRows.Count
        
        If .FixRow = 0 Then .FixRow = .Body.FixedRows
        gintFixRow = .FixRow
        gintFixCol = .FixCol
        gintGroups = 1
        
        gsngDown = .EmptyDown
        gsngLeft = .EmptyLeft
        gsngRight = .EmptyRight
        gsngUp = .EmptyUp
        gsngHeader = .PageHeader
        gsngFooter = .PageFooter
        
        gstrHeader = .Header
        gstrHeader = IIf(gstrHeader = "", ";;", gstrHeader)
        gstrFooter = .Footer
        gstrFooter = IIf(gstrFooter = "", ";;", gstrFooter)
    End With
    If gsngDown < 0 Or gsngUp < 0 Or gsngLeft < 0 Or gsngRight < 0 Or gsngHeader < 0 Or gsngFooter < 0 Then
        MsgBox "页边距不能设为负值。", vbCritical, gstrSysName
        ReadVar = False
        Exit Function
    End If
    If (gsngDown + gsngUp) * conRatemmToTwip > gsngScaleHeight Then
        MsgBox "页上边距或页下边距的值太大了。", vbCritical, gstrSysName
        ReadVar = False
        Exit Function
    End If
    If (gsngLeft + gsngRight) * conRatemmToTwip > gsngScaleWidth Then
        MsgBox "页左边距或页右边距的值太大了。", vbCritical, gstrSysName
        ReadVar = False
        Exit Function
    End If
    If (gsngHeader + gsngFooter) * conRatemmToTwip > gsngScaleHeight Then
        MsgBox "页眉距或页脚距的值太大了。", vbCritical, gstrSysName
        ReadVar = False
        Exit Function
    End If
    
    Dim strKeyValue As String       '键值
    Dim lngKey As Long
    Dim lngKeySize As Long
    Dim strRegPath As String
    If IsWindows95 Then
        strRegPath = "Software\MicroSoft\Windows\CurrentVersion"
    Else
        strRegPath = "Software\MicroSoft\Windows NT\CurrentVersion"
    End If
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, strRegPath, 0, KEY_READ, lngKey) = 0 Then
        strKeyValue = Space(256)
        lngKeySize = 256
        If RegQueryValueEx(lngKey, "RegisteredOrganization", 0, 1, strKeyValue, lngKeySize) = 0 Then
            gstrUnitName = StripTerminator(strKeyValue)
        End If
        strKeyValue = Space(256)
        lngKeySize = 256
        If RegQueryValueEx(lngKey, "RegisteredOwner", 0, 1, strKeyValue, lngKeySize) = 0 Then
'            gstrUserName = StripTerminator(strKeyValue)
        End If
    End If
    RegCloseKey lngKey

    gintRowTotal = 0
    gintColTotal = 0
    gintPage = 0
    gsngTotalWidth = 0
    gintCopies = 1
    gintBegin = 1
    gintShow = 1
    Exit Function
errHandle:
    ReadVar = False
End Function

Public Function IsWindowsNT() As Boolean
'功能：是否WindowNT操作系统
    Const dwMaskNT = &H2&
    IsWindowsNT = (GetWinPlatform() And dwMaskNT)
End Function

Public Function IsWindows95() As Boolean
'功    能：判断是否在Windows95下工作
'参    数：无
'返    回：是返回True
    Const dwMask95 = &H1&
    IsWindows95 = (GetWinPlatform() And dwMask95)
End Function

Private Function GetWinPlatform() As Long
'功    能：返回当前的系统版本代号
'参    数：无
'返    回：
    Dim osvi As OSVERSIONINFO
    Dim strCSDVersion As String
    osvi.dwOSVersionInfoSize = Len(osvi)
    If GetVersionEx(osvi) = 0 Then
        Exit Function
    End If
    GetWinPlatform = osvi.dwPlatformId
End Function

Function StripTerminator(ByVal strString As String) As String
'功    能：去掉字符串变量中的 Chr$(0)字符
'编 制 人：朱玉宝
'编制日期：2011-03-01
'参    数：无
'返    回：无
    Dim intZeroPos As Integer

    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

Sub CalculateRC()
'功    能：为每一个单元格计算它的位置（页号、页面号）
'编 制 人：朱玉宝
'编制日期：2011-03-01
'参    数：无
'返    回：无
    Dim intPageRow As Long, intPageCol As Long '临时得到的页面号与页号
    Dim sngPageWidth As Single, sngPageHeight As Single    '临时得到的页宽度与页高度
    Dim sngRowHeight As Single '得出一个字的高度
    Dim intCol As Long      '实际的列数
    Dim i As Long

    Dim iTemp As Long
    Dim sngTemp As Single

    intPageCol = 1
    intPageRow = 1
    gsngTotalWidth = 0
    ReDim gsngPrintedWidth(1 To gintTotalCol)
    ReDim gintRow(1 To 2, 1 To gintTotalRow) '第一维用于该页的第一行的行号，第二维用于该页的最后一行的行号
    ReDim gintCol(1 To 2, 1 To gintTotalCol) '第一维用于该页的第一列的列号，第二维用于该页的最后一列的列号

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    On Error GoTo ErrHand
    '从第一个非固定列开始计算出每列的页面号
    gintCol(1, 1) = gintFixCol + 1
    For iTemp = gintFixCol + 1 To gintTotalCol

        '该列的宽度
        If gobjSend.Body.ColHidden(iTemp - 1) Then
            sngTemp = 0
        Else
            sngTemp = gobjSend.Body.ColWidth(iTemp - 1)
        End If
        If sngPageWidth + sngTemp > gsngWidth Then

            '超宽了
            If sngPageWidth = 0 Then

                '还没有一个非固定列,则强制跳过该列
                gsngPrintedWidth(intPageCol) = gsngFixCol + sngTemp
                gintCol(2, intPageCol) = iTemp '本页的最后一列就是本列
                If iTemp <> gintTotalCol Then  '不是要打印的最后一列
                    intPageCol = intPageCol + 1
                    If intPageCol <= gintTotalCol Then gintCol(1, intPageCol) = iTemp + 1 '本页的第一列就是本列
                End If
                gsngTotalWidth = gsngTotalWidth + sngTemp
            Else

                gsngPrintedWidth(intPageCol) = gsngFixCol + sngPageWidth
                sngPageWidth = 0
                gintCol(2, intPageCol) = iTemp - 1 '上一页的最后一列就是上一列
                intPageCol = intPageCol + 1
                '这一列放在下一页面进行计算
                '只所以再循环一次,是由于有这一列比整张纸都宽的情况
                gintCol(1, intPageCol) = iTemp      '本页的第一列就是本列
                iTemp = iTemp - 1
            End If
        Else
            'gintCol(iTemp) = intPageCol
            sngPageWidth = sngPageWidth + sngTemp
            gsngTotalWidth = gsngTotalWidth + sngTemp
        End If
    Next
    If sngPageWidth <> 0 Then '统计最后一页的宽度
          gintCol(2, intPageCol) = iTemp - 1 '上一页的最后一列就是上一列
          gsngPrintedWidth(intPageCol) = gsngFixCol + sngPageWidth
    End If

    '从第一个非固定行开始计算出每行的页号
    gintRow(1, 1) = gintFixRow + 1
    For iTemp = gintFixRow + 1 To gintTotalRow
        '该行的高度
        If gobjSend.Body.RowHidden(iTemp - 1) Then
            sngTemp = 0
        Else
            sngTemp = gobjSend.Body.RowHeightMin
        End If
        If sngPageHeight + sngTemp > gsngHeight Then
            '超高了
            If sngPageHeight = 0 Then
                '还没有一个非固定行,则强制跳过该行
                gintRow(2, intPageRow) = iTemp '本页的最后一行就是本行
                intPageRow = intPageRow + 1
                If intPageRow <= gintTotalRow Then gintRow(1, intPageRow) = iTemp + 1   '本页的第一列就是本列

            Else
                sngPageHeight = 0
                gintRow(2, intPageRow) = iTemp - 1 '上一页的最后一行就是上一行
                intPageRow = intPageRow + 1
                '只所以再循环一次,是由于有这一行比整张纸都高的情况
                gintRow(1, intPageRow) = iTemp      '本页的第一列就是本列
                iTemp = iTemp - 1
            End If
        Else
            'gintRow(iTemp) = intPageRow
            sngPageHeight = sngPageHeight + sngTemp
        End If
    Next
    If sngPageHeight <> 0 Then gintRow(2, intPageRow) = iTemp - 1 '上一页的最后一行就是上一行

    gintColTotal = intPageCol
    gintRowTotal = intPageRow
    gsngTotalWidth = gsngTotalWidth + gsngFixCol * gintColTotal
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Sub CalculateHeight()
'功    能：计算出标题、表上项目和表下项目的高度,固定行的高度、固定列的宽度
'编 制 人：朱玉宝
'编制日期：2011-03-01
'参    数：无
'返    回：无
    Dim intCol As Long, intRow As Long '临时的行列指针变量
    
    '计算出标题的高度
    gfrmTemp.Font.Name = gstrTitleFName
    gfrmTemp.Font.Size = gintTitleFSize
    gfrmTemp.Font.Bold = gblnTitleFBold
    gfrmTemp.Font.Italic = gblnTitleFItalic
    gsngTitle = gfrmTemp.TextHeight(gstrTabTitle) + 2 * conLineHigh
    '计算出表上项目和表下项目的高度
    gfrmTemp.Font.Name = gstrAppRowFName
    gfrmTemp.Font.Size = gintAppRowFSize
    gfrmTemp.Font.Bold = gblnAppRowFBold
    gfrmTemp.Font.Italic = gblnAppRowFItalic
    gsngDownAppRow = (gfrmTemp.TextHeight("jg") + conLineHigh) * gintDownAppRow + conLineHigh
    gsngUpAppRow = (gfrmTemp.TextHeight("jg") + conLineHigh) * gintUpAppRow + conLineHigh
    gsngFixRow = 0
    gsngFixCol = 0
            
    gintTotalCol = gobjSend.Body.Cols
    gintTotalRow = gobjSend.Body.Rows
    '计算出固定行的高度
    For intRow = 1 To gintFixRow
        gsngFixRow = gsngFixRow + gobjSend.Body.RowHeightMin
    Next
    '计算出固定列的宽度(打印时不输出固定列)
    For intCol = 1 To gintFixCol
        gsngFixCol = gsngFixCol + gobjSend.Body.ColWidth(intCol - 1)
    Next
    
'    If gintGroups = 1 Then
'        '计算出固定行的高度
'        grsGrid.Filter = "列号=1 and 行号<=" & CStr(gintFixRow)
'        Do Until grsGrid.EOF
'            gsngFixRow = gsngFixRow + grsGrid("高度")
'            grsGrid.MoveNext
'        Loop
'        '计算出固定列的宽度
'        grsGrid.Filter = "行号=1 and 列号<=" & CStr(gintFixCol)
'        Do Until grsGrid.EOF
'            gsngFixCol = gsngFixCol + grsGrid("宽度")
'            grsGrid.MoveNext
'        Loop
'        grsGrid.Filter = ""
'    End If
    gsngHeight = gsngScaleHeight - (gsngUp + gsngDown) * conRatemmToTwip - gsngTitle - gsngDownAppRow - gsngUpAppRow - gsngFixRow - 2 * conLineHigh
    gsngWidth = gsngScaleWidth - (gsngLeft + gsngRight) * conRatemmToTwip - gsngFixCol - 2 * conLineWide
End Sub

Public Sub PrintPage(ByVal intPage As Integer)
'功    能：在指定设备上打印指定页
'编 制 人：朱玉宝
'编制日期：2011-03-01
'参    数：intPage  打印的页码
'返    回：无
    '该页所在的页号与页面号
    Dim intPageRow As Long, intPageCol As Long
    Dim sngOriY As Single
    '如果为真表示是输出到打印机，会显示frmBusy窗口
    
    If intPage = 0 Then Exit Sub
    Debug.Print Printer.DeviceName
    
    intPageRow = 1
    intPageCol = 1
    If intPageCol = 0 Then intPageCol = gintColTotal
    Set gcolGrid = Nothing
    
    Dim sngLeft As Single, sngWidth As Single
    Dim lngOffsetLeft As Long
    lngOffsetLeft = Printer.ScaleX(GetDeviceCaps(Printer.hDC, PHYSICALOFFSETX), vbPixels, vbTwips)
    sngLeft = gsngLeft * conRatemmToTwip
    'sngWidth = gsngPrintedWidth(intPageCol)
    sngWidth = Printer.Width - lngOffsetLeft * 2
    
    If glngPrintRow = 0 Or Not gblnPrintMode Or gintPrintState > 1 Then
        zlOutTitle sngLeft, gsngUp * conRatemmToTwip - conLineHigh, sngWidth
    End If
    
    gobjOutTo.CurrentY = gsngUp * conRatemmToTwip + gsngTitle + gsngUpAppRow
            
    PrintTends intPageRow, intPageCol
    sngOriY = gobjOutTo.CurrentY
    
    If glngPrintRow = 0 Or Not gblnPrintMode Or gintPrintState > 1 Then
        zlOutTabAppSet gobjSend.UnderAppRows, sngLeft, gsngTitle + gsngUp * conRatemmToTwip - conLineHigh, sngWidth
        'PrintHeadFoot
        Call frmTendFileReader.PrintHead
        Call frmTendFileReader.PrintFoot
    End If
    
    gobjOutTo.CurrentY = sngOriY
    zlOutTabAppSet gobjSend.BelowAppRows, sngLeft, gobjOutTo.CurrentY + 100, sngWidth
    
    If gstrGrant <> "" Then
        PrintCell gstrGrant & "样稿", sngLeft, gsngUp * conRatemmToTwip, sngWidth, sngOriY - gsngUp * conRatemmToTwip, 2, RGB(255, 0, 0), , , "0000", "宋体", 48 * gsngScale
    End If
End Sub

Public Sub PrintHeadFoot()
'功    能：打印页眉页脚
'编 制 人：朱玉宝
'编制日期：2011-03-01
'参    数：无
'返    回：无
    Dim strLeft As String, strMiddle As String, strRight As String
    Dim intPos As Long
    Dim intPos1 As Long
    Dim strHeader As String, strFooter As String
    With gobjOutTo
        .FontName = gstrAppRowFName
        .FontSize = gintAppRowFSize * gsngScale
        .FontBold = gblnAppRowFBold
        .FontItalic = gblnAppRowFItalic
        .ForeColor = glngAppRowColor
    End With
    On Error Resume Next
    strHeader = ConvHF(gstrHeader)
    intPos = InStr(strHeader, ";")
    intPos1 = intPos + 1
    strLeft = Mid(strHeader, 1, intPos - 1)
    intPos = InStr(intPos1, strHeader, ";")
    strMiddle = Mid(strHeader, intPos1, intPos - intPos1)
    intPos1 = intPos + 1
    strRight = Mid(strHeader, intPos1)

    PrintCell strLeft, gsngLeft * conRatemmToTwip, gsngHeader * conRatemmToTwip, gsngWidth + gsngFixCol, gobjOutTo.TextHeight("中"), 0, _
        , , , "0000"
    PrintCell strMiddle, gsngLeft * conRatemmToTwip, gsngHeader * conRatemmToTwip, gsngWidth + gsngFixCol, gobjOutTo.TextHeight("中"), 2, _
        , , , "0000"
    PrintCell strRight, gsngLeft * conRatemmToTwip, gsngHeader * conRatemmToTwip, gsngWidth + gsngFixCol, gobjOutTo.TextHeight("中"), 1, _
        , , , "0000"
    
    strFooter = ConvHF(gstrFooter)
    intPos = InStr(strFooter, ";")
    intPos1 = intPos + 1
    strLeft = Mid(strFooter, 1, intPos - 1)
    intPos = InStr(intPos1, strFooter, ";")
    strMiddle = Mid(strFooter, intPos1, intPos - intPos1)
    intPos1 = intPos + 1
    strRight = Mid(strFooter, intPos1)

    PrintCell strLeft, gsngLeft * conRatemmToTwip, gsngScaleHeight - gsngFooter * conRatemmToTwip - gobjOutTo.TextHeight("中"), gsngWidth + gsngFixCol, gobjOutTo.TextHeight("中"), 0, _
        , , , "0000"
    PrintCell strMiddle, gsngLeft * conRatemmToTwip, gsngScaleHeight - gsngFooter * conRatemmToTwip - gobjOutTo.TextHeight("中"), gsngWidth + gsngFixCol, gobjOutTo.TextHeight("中"), 2, _
        , , , "0000"
    PrintCell strRight, gsngLeft * conRatemmToTwip, gsngScaleHeight - gsngFooter * conRatemmToTwip - gobjOutTo.TextHeight("中"), gsngWidth + gsngFixCol, gobjOutTo.TextHeight("中"), 1, _
        , , , "0000"
'    On Error GoTo 0
End Sub

Public Function zlOutTabAppRow(colItem As zlTFTabAppRow, ByVal X As Single, ByVal Y As Single, ByVal Width As Single) As Boolean
    '------------------------------------------------
    '功能： 输出表上或表下项目
    '参数：
    '   colItem:需要输出的zlPrintLvw对象的表上或表下项目
    '   X：从总宽度的Left 为X处开始打印而非输出对象的Left
    '   Y:输出对象的Y坐标
    '   Width: 打印的实际宽度
    '返回：
    '------------------------------------------------
    Dim objApp As zlTFTabAppItem            '表上表下项目对象
    Dim sngXStep As Single               'X方向平移步长
    Dim iCount As Long
    Dim sngCurrentY As Single
    Dim sngCurrentX As Single
    If colItem.Count = 0 Then Exit Function
    
    sngCurrentY = Y
    With gobjOutTo
        .FontName = gstrAppRowFName
        .FontSize = gintAppRowFSize * gsngScale
        .FontBold = gblnAppRowFBold
        .FontItalic = gblnAppRowFItalic
        .ForeColor = glngAppRowColor
        
        iCount = 0
        If colItem.Count = 1 Then
            sngXStep = Width
        Else
            sngXStep = Width / (colItem.Count - 1)
        End If
        For Each objApp In colItem
            iCount = iCount + 1
            .CurrentY = Y
            Select Case iCount
            Case Is = 1                             '最左项目
                sngCurrentX = 0
            Case Is = colItem.Count   '最右项目
                sngCurrentX = Width - .TextWidth(objApp.Text)
            Case Else                               '其他项目
                sngCurrentX = sngXStep * (iCount - 1) - .TextWidth(objApp.Text) / 2
            End Select
            PrintCell objApp.Text, X + sngCurrentX, .CurrentY, , gobjOutTo.TextHeight("中"), , _
                , , , "0000"
            
'            OutRow objApp.Text, X, sngCurrentX, Width
        Next

    End With
    zlOutTabAppRow = True
    
End Function

Public Function zlOutTabAppSet(TabAppRows As zlTFTabAppRows, ByVal X As Single, ByVal Y As Single, ByVal Width As Single) As Boolean
    '------------------------------------------------
    '功能： 输出网格的表上或表下项目
    '参数：
    '   TabAppRows:表上还是表下项目
    '   X：从总宽度的Left 为X处开始打印而非输出对象的Left
    '   Y:输出对象的Y坐标
    '   Width: 打印的实际宽度
    '返回：
    '------------------------------------------------
    
    Dim sngXStep As Single             'X方向平移步长
    Dim iCount As Long
    Dim sngCurrentY As Single
    Dim sngCurrentX As Single
    Dim objApp As zlTFTabAppItem          '表上表下项目对象
    Dim colItem As zlTFTabAppRow          '表上或表下项目行
    
    Dim strTemp As String
    
    If TabAppRows.Count = 0 Then Exit Function
    sngCurrentY = Y
    With gobjOutTo
        .FontName = gstrAppRowFName
        .FontSize = gintAppRowFSize * gsngScale
        .FontBold = gblnAppRowFBold
        .FontItalic = gblnAppRowFItalic
        .ForeColor = glngAppRowColor
        
        For Each colItem In TabAppRows
            If colItem.Count = 1 Then
                sngXStep = Width
            Else
                sngXStep = Width / (colItem.Count - 1)
            End If
            iCount = 0
            For Each objApp In colItem
                iCount = iCount + 1
                .CurrentY = sngCurrentY
                strTemp = objApp.Text
                Select Case iCount
                Case Is = 1                             '最左项目
                    sngCurrentX = 0
                Case Is = colItem.Count                 '最右项目
                    sngCurrentX = Width - .TextWidth(strTemp)
                Case Else                               '其他项目
                    sngCurrentX = sngXStep * (iCount - 1) - .TextWidth(strTemp) / 2
                End Select
               PrintCell objApp.Text, X + sngCurrentX, .CurrentY, , gobjOutTo.TextHeight("中"), , _
                     , , , "0000"
'                OutRow strTemp, X, sngCurrentX, Width
            Next
            sngCurrentY = sngCurrentY + .TextHeight("ZL")
        Next
    End With
    
    zlOutTabAppSet = True
        
End Function

Public Function zlOutTitle(ByVal X As Single, ByVal Y As Single, ByVal Width As Single) As Boolean
    '------------------------------------------------
    '功能： 输出标题
    '参数：X：从总宽度的Left 为X处开始打印而非输出对象的Left
    '      Y:输出对象的Y坐标
    '      Width: 打印的实际宽度
    '返回：无
    '------------------------------------------------
    Dim sinLeft As Single
    
    If gstrTabTitle = "" Then Exit Function
    
    With gobjOutTo
        .ForeColor = glngTitleColor
        .FontName = gstrTitleFName
        .FontSize = gintTitleFSize * gsngScale
        .FontBold = gblnTitleFBold
        .FontItalic = gblnTitleFItalic
        .CurrentY = Y
        '标题真正开始打印的位置
'        sinLeft = (gsngTotalWidth - .TextWidth(gstrTabTitle)) / 2
        PrintCell gstrTabTitle, X, .CurrentY, Width - 2 * X, gobjOutTo.TextHeight("中"), 2, _
            , , , "0000"

'        OutRow gstrTabTitle, X, sinLeft, Width
    End With
    zlOutTitle = True
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
Public Function SQLObject(ByVal strSQL As String) As String
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
    
    On Error GoTo errH
    
    '大写化及去除多余的字符
    strAnal = UCase(TrimChar(strSQL))

    If InStr(strAnal, "SELECT") = 0 Or InStr(strAnal, "FROM") = 0 Then Exit Function
    
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
                    strObject = strObject & "," & SQLObject(strSub)
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
            If InStr(strObject & ",", "," & strTrue & ",") = 0 And strTrue <> "嵌套查询" Then
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
Public Function TrimChar(str As String) As String
'功能:去除字符串中连续的空格和回车(含两头的空格,回车),不去除TAB字符,哪怕是连续的
    Dim strTmp As String
    Dim i As Long, j As Long
    
    If Trim(str) = "" Then TrimChar = "": Exit Function
    
    strTmp = Trim(str)
    
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

'64055:刘鹏飞,2013-09-04,注释该函数直接调用zlcomlib的方法，由于该函数没有使用缓存
'Public Function GetPrivFunc(lngSys As Long, lngProgID As Long) As String
''功能：返回当前用户具有的指定程序的功能串
''参数：lngSys     如果是固定模块，则为0
''      lngProgId  程序序号
''返回：分号间隔的功能串,为空表示没有权限
'    Dim rsTmp As ADODB.Recordset
'    Dim strSQL As String, strPrivs As String
'    Dim strWhere As String
'
'    On Error GoTo errH
'
''    If zlRegCheck <> "" Then Exit Function
'
'    strSQL = "Select Text as 功能 From Table(Cast(zltools.f_Reg_Func([1],[2]) as zlTools.t_Reg_Rowset))"
'    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetPrivFunc", lngSys, lngProgID)
'    Do While Not rsTmp.EOF
'        strPrivs = strPrivs & ";" & rsTmp!功能
'        rsTmp.MoveNext
'    Loop
'    GetPrivFunc = Mid(strPrivs, 2)
'    Exit Function
'errH:
'    If ErrCenter() = 1 Then Resume
'End Function

Public Function OutRow(ByVal strPrint As String, ByVal X As Single, ByVal sngLeft As Single, ByVal Width As Single) As Boolean
    '------------------------------------------------
    '功能： 输出一行文字
    '参数：X：从总宽度的Left 为X处开始打印而非输出对象的Left
    '      Y:输出对象的Y坐标
    '      Width: 打印的实际宽度
    '返回：无
    '------------------------------------------------
    Dim strTemp As String
    With gobjOutTo
        If sngLeft >= X Then '前面还有一段空白
            .CurrentX = gsngLeft * conRatemmToTwip + sngLeft - X
        Else
            Do While sngLeft + .TextWidth(strTemp) < X
                    If Len(strPrint) = 0 Then Exit Do
                    strTemp = strTemp & Left(strPrint, 1)
                    strPrint = Mid(strPrint, 2)
            Loop
            .CurrentX = gsngLeft * conRatemmToTwip + sngLeft + .TextWidth(strTemp) - X
        End If
        Dim intPageCol As Long
        intPageCol = gintPage Mod gintColTotal
        If intPageCol = 0 Then intPageCol = gintColTotal
        strTemp = ""
        Do While (.TextWidth(strPrint) > Width) Or (.CurrentX + .TextWidth(strPrint) > gsngLeft * conRatemmToTwip + gsngPrintedWidth(intPageCol))
            If Len(strPrint) = 0 Then Exit Do
            strTemp = Right(strPrint, 1) & strTemp
            strPrint = Mid(strPrint, 1, Len(strPrint) - 1)
        Loop
        If Len(strTemp) > 0 And .CurrentX < gsngLeft * conRatemmToTwip + gsngPrintedWidth(intPageCol) + .TextWidth("中") Then strPrint = strPrint & Left(strTemp, 1)
        If Len(strPrint) = 0 Then Exit Function
        gobjOutTo.Print strPrint
    End With
End Function

Public Function ConvHF(ByVal strSource As String) As String
    '------------------------------------------------
    '功能：将页眉与页脚转换成实际打印的内容
    '参数：strSource    页眉与页脚
    '返回：实际打印的内容
    '------------------------------------------------
    Dim strTemp As String
    
    strTemp = Replace(strSource, "[页码]", CStr(gintPage + gintBegin - 1))
    strTemp = Replace(strTemp, "[页数]", CStr(gintColTotal * gintRowTotal))
    strTemp = Replace(strTemp, "[时间]", Format(Time, "HH:MM:SS"))
    strTemp = Replace(strTemp, "[日期]", Format(date, "YYYY年mm月dd日"))
    strTemp = Replace(strTemp, "[用户名]", gstrUserName)
    strTemp = Replace(strTemp, "[单位名]", gstrUnitName)
    ConvHF = strTemp
End Function

Public Sub RealPrint(ByVal intBegin As Long, ByVal intEnd As Long)
    '功能： 输出打印机上
    '参数：intBegin     开始页码
    '      intEnd       结束页码
    '返回：无
    '------------------------------------------------
    Dim frmOutTemp As New frmOutStatus
    On Error Resume Next
    Screen.MousePointer = 11
    frmOutTemp.mintBegin = intBegin
    frmOutTemp.mintEnd = intEnd
    frmOutTemp.Show 1
    Unload frmOutTemp
    Set frmOutTemp = Nothing
    Screen.MousePointer = 0
End Sub


Public Sub ApplyOEM(objStatus As Object)
'针对状态栏应用OEM策略
    Dim strOEM As String
    On Error Resume Next
    
    If gstrSysName <> "-" Then
        objStatus.Panels(1).Text = gstrSysName
        '处理状态栏图标的OEM策略
        If gstrSysName = "中联软件" Then
            Set objStatus.Panels(1).Picture = LoadCustomPicture("Logo")
        Else
            strOEM = GetOEM(Mid(gstrSysName, 1, Len(gstrSysName) - 2))
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

Public Function LoadCustomPicture(strID As String) As StdPicture
'功能:将资源文件中的指定资源生成磁盘文件
'参数:ID=资源号,strExt=要生成文件的扩展名(如BMP)
'返回:生成文件名
    Dim arrData() As Byte
    Dim intFile As Integer
    Dim strFile As String * 255, strR As String
    
    arrData = LoadResData(strID, "CUSTOM")
    intFile = FreeFile
    
    GetTempPath 255, strFile
    strR = Trim(Left(strFile, InStr(strFile, Chr(0)) - 1)) & CLng(Timer * 100) & ".pic"

    Open strR For Binary As intFile
    Put intFile, , arrData()
    Close intFile
    Set LoadCustomPicture = VB.LoadPicture(strR)
    Kill strR
End Function

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

Public Function SetCustonPager(ByVal lngWidth As Long, ByVal lngHeight As Long) As Integer
'功能：在设置自定义纸张
'参数：是以绨为单位
    If IsWindowsNT Then
        '虽然不能使宽度生效，但能改变PaperSize的属性值
        Printer.Width = lngWidth
        Printer.Height = lngHeight
        SetCustonPager = SetNTPrinterPaper(gfrmTemp.hWnd, lngWidth / conRatemmToTwip, lngHeight / conRatemmToTwip, Printer.Orientation, Printer.Copies)
    Else
        'Windows98系列还是以通常方法处理
        Printer.PaperSize = 256
        Printer.Width = lngWidth
        Printer.Height = lngHeight
    End If
End Function

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function GetPrinterSet() As Boolean
'------------------------------------------------
    '功能：读取本系统注册表的打印缺省设置
    '------------------------------------------------
    Dim iCount As Long
    Dim strDeviceName As String
    Dim intPaperSize As Integer
    Dim intPaperBin As Integer
    Dim intOrientation As Long
    
    If Printers.Count = 0 Then
        GetPrinterSet = False
        Exit Function
    End If
    
    strDeviceName = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\Default", "DeviceName", Printer.DeviceName)
    If Printer.DeviceName <> strDeviceName Then
        For iCount = 0 To Printers.Count - 1
            If Printers(iCount).DeviceName = strDeviceName Then
                Set Printer = Printers(iCount)
                Exit For
            End If
        Next
    End If
    
    Err = 0
    On Error Resume Next
    Printer.PaperBin = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\Default", "PaperBin", Printer.PaperBin)
    Printer.Orientation = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\Default", "Orientation", Printer.Orientation)
    
    intPaperSize = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\Default", "PaperSize", Printer.PaperSize)
    If intPaperSize = 256 Then
        Dim lngWidth As Long
        Dim lngHeight As Long
        
        lngWidth = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\Default", "Width", Printer.Width)
        lngHeight = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\Default", "Height", Printer.Height)
        
        Call SetCustonPager(lngWidth, lngHeight)
    Else
        Printer.PaperSize = intPaperSize
    End If
    GetPrinterSet = True
End Function

Public Function ReadPageHead(objHead As RichTextBox, ByVal strKey As String) As Boolean
'################################################################################################################
'## 功能：  读取页面图片
'## 参数：  病历种类-页面编号
'## 返回：  返回获得的图片变量。
'################################################################################################################
    Dim strFile As String, strZip As String
    strZip = zlBlobRead(12, strKey, App.Path & "\Head_L.zip")
    If gobjFSO.FileExists(strZip) Then
        strFile = UnzipTendPage(strZip, "Head_S.RTF")
        objHead.LoadFile strFile, rtfRTF           '读取文件
        gobjFSO.DeleteFile strFile, True      '删除临时文件
        ReadPageHead = True
    Else
        objHead.Text = ""
    End If
End Function

Public Function ReadPageFoot(objFoot As RichTextBox, ByVal strKey As String) As Boolean
'################################################################################################################
'## 功能：  读取页面图片
'## 参数：  病历种类-页面编号
'## 返回：  返回获得的图片变量。
'################################################################################################################
    Dim strFile As String, strZip As String
    strZip = zlBlobRead(13, strKey, App.Path & "\Foot_L.zip")
    If gobjFSO.FileExists(strZip) Then
        strFile = UnzipTendPage(strZip, "Foot_S.RTF")
        objFoot.LoadFile strFile, rtfRTF           '读取文件
        gobjFSO.DeleteFile strFile, True      '删除临时文件
        ReadPageFoot = True
    Else
        objFoot.Text = ""
    End If
End Function

'################################################################################################################
'## 功能：  将指定的LOB字段复制为临时文件
'##
'## 参数：  Action      :操作类型（用以区别是操作哪个表）
'##         KeyWord     :确定数据记录的关键字，复合关键字以逗号分隔(仅5-电子病历格式为复合)
'##         strFile     :用户指定存放的文件名；不指定时，取当前路径产生文件名
'##
'## 返回：  存放内容的文件名，失败则返回零长度""
'##
'## 说明：  Action取值说明：
'##         0-病历标记图形；1-病历文件格式；2-病历文件图形；3-病历范文格式；4-病历范文图形；5-电子病历格式；6-电子病历图形；
'################################################################################################################
Public Function zlBlobRead(ByVal Action As Long, ByVal KeyWord As String, Optional ByRef strFile As String, Optional ByVal blnMoved As Boolean) As String
    
    Const conChunkSize As Integer = 10240
    Dim lngFileNum As Long, lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, strText As String
    Dim rsLob As New ADODB.Recordset
    
    Err = 0: On Error GoTo ErrHand
    
    lngFileNum = FreeFile
    If strFile = "" Then
        lngCount = 0
        Do While True
            strFile = App.Path & "\zlBlobFile" & CStr(lngCount) & ".tmp"
            If Len(Dir(strFile)) = 0 Then Exit Do
            lngCount = lngCount + 1
        Loop
    End If
    Open strFile For Binary As lngFileNum
    
    gstrSQL = "Select Zl_Lob_Read([1],[2],[3],[4]) as 片段 From Dual"
    lngCount = 0
    Do
        Set rsLob = zlDatabase.OpenSQLRecord(gstrSQL, "zlBlobRead", Action, KeyWord, lngCount, IIf(blnMoved, 1, 0))
        If rsLob.EOF Then Exit Do
        If IsNull(rsLob.Fields(0).Value) Then Exit Do
        strText = rsLob.Fields(0).Value
        
        ReDim aryChunk(Len(strText) / 2 - 1) As Byte
        For lngBound = LBound(aryChunk) To UBound(aryChunk)
            aryChunk(lngBound) = CByte("&H" & Mid(strText, lngBound * 2 + 1, 2))
        Next
        
        Put lngFileNum, , aryChunk()
        lngCount = lngCount + 1
    Loop
    Close lngFileNum
    If lngCount = 0 Then Kill strFile: strFile = ""
    zlBlobRead = strFile
    Exit Function

ErrHand:
    Close lngFileNum
    Kill strFile: zlBlobRead = ""
End Function

'################################################################################################################
'## 功能：  在压缩文件相同目录释放产生解压文件
'## 参数：  strZipFile     :压缩文件
'## 返回：  解压文件名，失败则返回零长度""
'################################################################################################################
Public Function UnzipTendPage(ByVal strZipFile As String, ByVal strTarFile As String) As String
    Dim strZipPathTmp As String
    Dim strZipPath As String
    Dim strZipFileTmp As String
    Dim strZipFileName As String
    
    On Error GoTo ErrHand
    
    If Not gobjFSO.FileExists(strZipFile) Then UnzipTendPage = "": Exit Function
    strZipPath = Left(strZipFile, Len(strZipFile) - Len(Dir(strZipFile)))
    
    strZipPath = gobjFSO.GetSpecialFolder(2)
    strZipPathTmp = strZipPath & Format(Now, "yyMMddHHmmss") & CStr(100 * Timer)
    Call gobjFSO.CreateFolder(strZipPathTmp)
    
    strZipFileTmp = strZipPathTmp ' & "\TMP.RTF"
    
    With mclsUnzip
        .ZipFile = strZipFile
        .UnzipFolder = strZipPathTmp
        .Unzip
    End With
    If gobjFSO.FolderExists(strZipFileTmp) Then
        
        strZipFileName = gobjFSO.GetFile(strZipFileTmp & "\" & strTarFile)
        Call gobjFSO.CopyFile(strZipFileName, "C:\" & strTarFile)
        
        On Error Resume Next
        gobjFSO.DeleteFolder strZipPathTmp, True
        gobjFSO.DeleteFile strZipFile, True
        
        UnzipTendPage = "C:\" & strTarFile
    Else
        UnzipTendPage = ""
    End If
ErrHand:
    Exit Function
End Function

Public Function GetTmpPath() As String
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strFileTemp As String
    Dim lngTemp As Long
    
    strFileTemp = Space(256)
    lngTemp = GetTempPath(256, strFileTemp)
    
    GetTmpPath = Mid(strFileTemp, 1, InStr(strFileTemp, Chr(0)) - 1)
End Function

Public Function FormatValue(ByVal strValue As String) As String
    Dim strRetrun As String
    strRetrun = Replace(Replace(Replace(strValue, Chr(10), ""), Chr(13), ""), Chr(1), "")
    FormatValue = strRetrun
End Function


Public Sub Record_Add(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String)
    Dim arrFields, arrValues, intField As Integer
    '添加记录
    'strFields:字段名|字段名
    'strValues:值|值
    
    '例子：
    'Dim strFields As String, strValues As String
    'strFields = "RecordID|科目ID|摘要"
    'strValues = "5188|6666|科目名称"
    'Call Record_Update(rsVoucher, strFields, strValues)

    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField = 0 Then Exit Sub

    With rsObj
        .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Public Sub Record_Update(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String, ByVal strPrimary As String, Optional ByVal blnDelete As Boolean = False)
    Dim arrFields, arrValues, intField As Integer
    '更新记录,如果不存在,则新增
    'strPrimary:字段名|值
    'strFields:字段名|字段名
    'strValues:值|值
    
    '例子：
    'Dim strFields As String, strValues As String, strPrimary As String
    'strFields = "RecordID|科目ID|摘要"
    'strValues = "5188|6666|科目名称"
    'strPrimary = "RecordID|5188"
    'Call Record_Update(rsVoucher, strFields, strValues, strPrimary, True)

    If strValues = "" Then strValues = " "
    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField < 0 Then Exit Sub

    With rsObj
        If Record_Locate(rsObj, strPrimary, blnDelete) = False Then .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Public Function Record_Locate(ByRef rsObj As ADODB.Recordset, ByVal strPrimary As String, Optional ByVal blnDelete As Boolean = False) As Boolean
    Dim arrTmp
    '定位到指定记录
    'strPrimary:主健,值
    'blnDelete=True,则该记录集存在"删除"字段
    Record_Locate = False
    
    arrTmp = Split(strPrimary, "|")
    With rsObj
        If .RecordCount = 0 Then Exit Function
        .MoveFirst
        .Find arrTmp(0) & "='" & arrTmp(1) & "'"
        If .EOF Then Exit Function
        If blnDelete Then
            Do While Not .EOF
                If !删除 = 0 Then Record_Locate = True: Exit Do
                .MoveNext
            Loop
        Else
            Record_Locate = True
        End If
    End With
End Function

Public Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)
    Dim arrFields, intField As Integer
    Dim strFieldName As String, intType As Integer, lngLength As Long
    '初始化映射记录集
    'strFields:字段名,类型,长度|字段名,类型,长度    如果长度为零,则取默认长度
    '字符型:adLongVarChar;数字型:adDouble;日期型:adDBDate
    
    '例子：
    'Dim rsVoucher As New ADODB.Recordset, strFields As String
    'strFields = "RecordID," & adDouble & ",18|科目ID," & adDouble & ",18|摘要, " & adLongVarChar & ",50|" & _
    '"删除," & adDouble & ",1"
    'Call Record_Init(rsVoucher, strFields)

    arrFields = Split(strFields, "|")
    Set rsObj = New ADODB.Recordset

    With rsObj
        If .State = 1 Then .Close
        For intField = 0 To UBound(arrFields)
            strFieldName = Split(arrFields(intField), ",")(0)
            intType = Split(arrFields(intField), ",")(1)
            lngLength = Split(arrFields(intField), ",")(2)

            '获取字段缺省长度
            If lngLength = 0 Then
                Select Case intType
                Case adDouble
                    lngLength = madDoubleDefault
                Case adVarChar
                    lngLength = madLongVarCharDefault
                Case adLongVarChar
                    lngLength = madLongVarCharDefault
                Case Else
                    lngLength = madDbDateDefault
                End Select
            End If
            .Fields.Append strFieldName, intType, lngLength, adFldIsNullable
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Public Sub OutputRsData(ByVal rsObj As ADODB.Recordset)
    Dim intCol As Integer, intCols As Integer
    Dim strValues As String
    With rsObj
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            strValues = ""
            intCols = .Fields.Count - 1
            For intCol = 0 To intCols
                strValues = strValues & "," & .Fields(intCol).Name & ":" & .Fields(intCol).Value
            Next
            Debug.Print strValues
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
End Sub
