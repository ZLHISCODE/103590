Attribute VB_Name = "mdlCommon"
Option Explicit
'**************************
'       OEM代号
'
'医业  D2BDD2B5
'**************************
Public gobjFSO As New FileSystemObject
Public gcnOracle As ADODB.Connection
Public gstrSysName As String
Public gstrDBUser As String '用户名
Public gstrPrivs As String

'错误日志处理相关变量
Private mlngErrNum As Long, mstrErrInfo As String, mbytErrType As Byte
Private mstrRecentSQL As String  '最近执行的SQL语句

'SQLLog变量
Private msngTime As Single
Private mobjLogText As TextStream

Public gobjFile As New FileSystemObject

Global Const gintTends% = 1                       '打印对象是zlPrintTends
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
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Global gstrUserName As String     '安装Windows时填写的用户名
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
Public gint婴儿 As Integer
Public gstrSQL As String
Public glngPage As Long                 '当前页面,第一页才存在打印过的内容不打印的问题
Public gblnMoved_HL As Boolean
Public frmAsk As frmTendPrintAsk        '询问窗体
Public gstr对角线 As String             '保存列号序号
Public glngPrintRow As Long             '从此行开始打印
Public gblnPrintMode As Boolean         '打印模式为TRUE
Public gintPrintState As Integer        '打印模式，1-续打；2-重打；3-连续重打

Public Enum enuPage
    续打
    重打
    连续重打
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
Public Declare Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal Hwnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, pDevModeOutput As Any, pDevModeInput As Any, ByVal fMode As Long) As Long
Public Declare Function ResetDC Lib "gdi32" Alias "ResetDCA" (ByVal hDC As Long, lpInitData As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)



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
'编 制 人：徐强
'编制日期：1999年7月12日
'参    数：无
'返    回：读入参数有效则返回真
    ReadVar = True
    On Error GoTo errHandle
    gsngPageWidth = Printer.Width
    gsngPageHeight = Printer.Height
    gsngPageScaleWidth = Printer.ScaleWidth
    gsngPageScaleHeight = Printer.ScaleHeight
    
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
            gstrUserName = StripTerminator(strKeyValue)
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
'编 制 人：徐强
'编制日期：1999年7月2日
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
'编 制 人：徐强
'编制日期：1999年7月2日
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
    On Error GoTo errHand
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
                gintCol(2, intPageRow) = iTemp '本页的最后一行就是本行
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
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Sub CalculateHeight()
'功    能：计算出标题、表上项目和表下项目的高度,固定行的高度、固定列的宽度
'编 制 人：徐强
'编制日期：1999年7月2日
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
'编 制 人：徐强
'编制日期：1999年7月5日
'参    数：intPage  打印的页码
'返    回：无
    '该页所在的页号与页面号
    Dim intPageRow As Long, intPageCol As Long
    Dim sngOriY As Single
    '如果为真表示是输出到打印机，会显示frmBusy窗口
    
    If intPage = 0 Then Exit Sub
    glngPage = intPage
    intPageRow = 1
    intPageCol = 1
    If intPageCol = 0 Then intPageCol = gintColTotal
    Set gcolGrid = Nothing
    
    Dim sngLeft As Single, sngWidth As Single
    sngLeft = gsngLeft * conRatemmToTwip
    'sngWidth = gsngWidth
    sngWidth = gsngPrintedWidth(intPageCol)
    
    If glngPrintRow = 0 Or Not gblnPrintMode Or glngPage > 1 Or gintPrintState > 1 Then
        zlOutTitle sngLeft, gsngUp * conRatemmToTwip - conLineHigh, sngWidth
    End If
    
    gobjOutTo.CurrentY = gsngUp * conRatemmToTwip + gsngTitle + gsngUpAppRow
            
    PrintTends intPageRow, intPageCol
    sngOriY = gobjOutTo.CurrentY
    
    If glngPrintRow = 0 Or Not gblnPrintMode Or glngPage > 1 Or gintPrintState > 1 Then
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
'编 制 人：徐强
'编制日期：1999年7月10日
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

Public Function zlOutTabAppRow(colItem As zlTabAppRow, ByVal x As Single, ByVal Y As Single, ByVal Width As Single) As Boolean
    '------------------------------------------------
    '功能： 输出表上或表下项目
    '参数：
    '   colItem:需要输出的zlPrintLvw对象的表上或表下项目
    '   X：从总宽度的Left 为X处开始打印而非输出对象的Left
    '   Y:输出对象的Y坐标
    '   Width: 打印的实际宽度
    '返回：
    '------------------------------------------------
    Dim objApp As zlTabAppItem            '表上表下项目对象
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
            PrintCell objApp.Text, x + sngCurrentX, .CurrentY, , gobjOutTo.TextHeight("中"), , _
                , , , "0000"
            
'            OutRow objApp.Text, X, sngCurrentX, Width
        Next

    End With
    zlOutTabAppRow = True
    
End Function

Public Function zlOutTabAppSet(TabAppRows As zlTabAppRows, ByVal x As Single, ByVal Y As Single, ByVal Width As Single) As Boolean
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
    Dim objApp As zlTabAppItem          '表上表下项目对象
    Dim colItem As zlTabAppRow          '表上或表下项目行
    
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
               PrintCell objApp.Text, x + sngCurrentX, .CurrentY, , gobjOutTo.TextHeight("中"), , _
                     , , , "0000"
'                OutRow strTemp, X, sngCurrentX, Width
            Next
            sngCurrentY = sngCurrentY + .TextHeight("ZL")
        Next
    End With
    
    zlOutTabAppSet = True
        
End Function

Public Function zlOutTitle(ByVal x As Single, ByVal Y As Single, ByVal Width As Single) As Boolean
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
        PrintCell gstrTabTitle, x, .CurrentY, Width, gobjOutTo.TextHeight("中"), 2, _
            , , , "0000"

'        OutRow gstrTabTitle, X, sinLeft, Width
    End With
    zlOutTitle = True
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
    
    bytReturnType = 1
    If gcnOracle.Errors.Count <> 0 Then
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
            strNote = "表或视图不存在，很可能是你不具备使用该部分数据的权限。"
            bytReturnType = 0
            
            strTemp = mGetInvalidTable()
            If strTemp <> "" Then
                mstrErrInfo = "请对下列对象进行检查：" & vbCrLf & vbCrLf & vbTab & strTemp
            Else
                mstrErrInfo = "错误SQL语句为：" & vbCrLf & vbCrLf & mstrRecentSQL
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
                
                strNote = "请在服务器管理工具的角色管理程序中增加对过程“" & strTemp & "”的授权。"
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

'    If bytReturnType = 1 Then
'        ErrCenter = frmErrAsk.ShowEdit(mlngErrNum, strNote, mstrErrInfo)
'    Else
'        Call frmErrNote.ShowEdit(mlngErrNum, strNote, mstrErrInfo)
'        ErrCenter = 0
'    End If
    
    '连续重打错误
    Err.Clear
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

Public Function OpenSQLRecord(ByVal strSQL As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
'功能：通过Command对象打开带参数SQL的记录集
'参数：strSQL=条件中包含参数的SQL语句,参数形式为"[x]"
'             x>=1为自定义参数号,"[]"之间不能有空格
'             同一个参数可多处使用,程序自动换为ADO支持的"?"号形式
'             实际使用的参数号可不连续,但传入的参数值必须连续(如SQL组合时不一定要用到的参数)
'      arrInput=不定个数的参数值,按参数号顺序依次传入,必须是明确类型
'      strTitle=用于SQLTest识别的调用窗体/模块标题
'返回：记录集，CursorLocation=adUseClient,LockType=adLockReadOnly,CursorType=adOpenStatic
'举例：
'SQL语句为="Select 姓名 From 病人信息 Where (病人ID=[3] Or 门诊号=[3] Or 姓名 Like [4]) And 性别=[5] And 登记时间 Between [1] And [2] And 险类 IN([6],[7])"
'调用方式为：Set rsPati=OpenSQLRecord(strSQL, Me.Caption, CDate(Format(rsMove!转出日期,"yyyy-MM-dd")),dtp时间.Value, lng病人ID, "张%", "男", 20, 21)
    Static cmdData As New ADODB.Command
    Dim strPar As String, arrPar As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim strSeq As String, intMax As Integer, i As Integer
    Dim strLog As String, varValue As Variant
    
    '分析自定的[x]参数
    lngLeft = InStr(1, strSQL, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strSQL, "]")
        
        '可能是正常的"[编码]名称"
        strSeq = Mid(strSQL, lngLeft + 1, lngRight - lngLeft - 1)
        If IsNumeric(strSeq) Then
            i = CInt(strSeq)
            strPar = strPar & "," & i
            If i > intMax Then intMax = i
        End If
        
        lngLeft = InStr(lngRight + 1, strSQL, "[")
    Loop

    '替换为"?"参数
    strLog = strSQL
    For i = 1 To intMax
        strSQL = Replace(strSQL, "[" & i & "]", "?")
        
        '产生用于SQL跟踪的语句
        varValue = arrInput(i - 1)
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
            strLog = Replace(strLog, "[" & i & "]", varValue)
        Case "String" '字符
            strLog = Replace(strLog, "[" & i & "]", "'" & Replace(varValue, "'", "''") & "'")
        Case "Date" '日期
            strLog = Replace(strLog, "[" & i & "]", "To_Date('" & Format(varValue, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')")
        End Select
    Next

    '连续重打原有参数:不然不能重复执行
    cmdData.CommandText = "" '不为空有时连续重打参数出错
    Do While cmdData.Parameters.Count > 0
        cmdData.Parameters.Delete 0
    Loop
    
    '创建新的参数
    lngLeft = 0: lngRight = 0
    arrPar = Split(Mid(strPar, 2), ",")
    For i = 0 To UBound(arrPar)
        varValue = arrInput((arrPar(i) - 1))
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarNumeric, adParamInput, 30, varValue)
        Case "String" '字符
            intMax = LenB(StrConv(varValue, vbFromUnicode))
            If intMax = 0 Or intMax < 200 Then intMax = 200
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarChar, adParamInput, intMax, varValue)
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
                If intMax = 0 Or intMax < 200 Then intMax = 200
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarChar, adParamInput, intMax, varValue(lngLeft))
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
        Set cmdData.ActiveConnection = gcnOracle '这句比较慢
    End If
    cmdData.CommandText = strSQL
    
    Call SQLTest(App.ProductName, strTitle, strLog)
    Set OpenSQLRecord = cmdData.Execute
    Call SQLTest
End Function

Public Sub SQLTest(Optional ByVal strProject As String, Optional ByVal strForm As String, Optional ByVal strSQL As String, Optional ByVal strNote As String)
'功能：将部件中执行的SQL语句输出到窗体或文件中，并附加开始结束时间，执行时间
'参数：strProject=部件名称,具体可取App.Title
'      strForm=窗体名,具体可取Form.Caption
'      strSQL=将要执行的SQL语句,在Open时传入,如果不传，表示最近一次SQL执行完毕
'      strNote=SQL语句说明
    Dim strTmp As String, sngEnd As Single
    
    mstrRecentSQL = strSQL  '保存最近执行的SQL语句
    
    If UCase(gcnOracle.ConnectionString) Like "*;SERVER=SQLLOG*" Then
        If strSQL <> "" Then
            If mobjLogText Is Nothing Then
                On Local Error Resume Next
                Set mobjLogText = gobjFile.OpenTextFile("ReportSQL_" & gstrDBUser & "_" & Format(date, "yyyyMMdd") & ".log", ForAppending, True, TristateFalse)
                On Local Error GoTo 0
            End If
            If Not mobjLogText Is Nothing Then
                strTmp = "[" & Format(Time, "HH:mm:ss") & "]"
                mobjLogText.WriteLine strTmp & "Application:" & strProject & "\" & strForm & IIf(strNote <> "", "," & strNote, "")
                mobjLogText.WriteLine strTmp & "SQL:" & strSQL
                msngTime = Timer
            End If
        Else
            If Not mobjLogText Is Nothing Then
                sngEnd = Timer
                strTmp = "[" & Format(Time, "HH:mm:ss") & "]"
                mobjLogText.WriteLine strTmp & "Expend:" & Format(sngEnd - msngTime, "0.0000")
                mobjLogText.WriteBlankLines 1
            End If
        End If
    End If
End Sub
Public Function GetPrivFunc(lngSys As Long, lngProgID As Long) As String
'功能：返回当前用户具有的指定程序的功能串
'参数：lngSys     如果是固定模块，则为0
'      lngProgId  程序序号
'返回：分号间隔的功能串,为空表示没有权限
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strPrivs As String
    Dim strWhere As String
    
    On Error GoTo errH
    
'    If zlRegCheck <> "" Then Exit Function
    
    strSQL = "Select Text as 功能 From Table(Cast(zltools.f_Reg_Func([1],[2]) as zlTools.t_Reg_Rowset))"
    Set rsTmp = OpenSQLRecord(strSQL, "GetPrivFunc", lngSys, lngProgID)
    Do While Not rsTmp.EOF
        strPrivs = strPrivs & ";" & rsTmp!功能
        rsTmp.MoveNext
    Loop
    GetPrivFunc = Mid(strPrivs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
End Function
Public Function OutRow(ByVal strPrint As String, ByVal x As Single, ByVal sngLeft As Single, ByVal Width As Single) As Boolean
    '------------------------------------------------
    '功能： 输出一行文字
    '参数：X：从总宽度的Left 为X处开始打印而非输出对象的Left
    '      Y:输出对象的Y坐标
    '      Width: 打印的实际宽度
    '返回：无
    '------------------------------------------------
    Dim strTemp As String
    With gobjOutTo
        If sngLeft >= x Then '前面还有一段空白
            .CurrentX = gsngLeft * conRatemmToTwip + sngLeft - x
        Else
            Do While sngLeft + .TextWidth(strTemp) < x
                    If Len(strPrint) = 0 Then Exit Do
                    strTemp = strTemp & Left(strPrint, 1)
                    strPrint = Mid(strPrint, 2)
            Loop
            .CurrentX = gsngLeft * conRatemmToTwip + sngLeft + .TextWidth(strTemp) - x
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
        SetCustonPager = SetNTPrinterPaper(gfrmTemp.Hwnd, lngWidth / conRatemmToTwip, lngHeight / conRatemmToTwip, Printer.Orientation, Printer.Copies)
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
    
    Err = 0: On Error GoTo errHand
    
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
        Set rsLob = OpenSQLRecord(gstrSQL, "zlBlobRead", Action, KeyWord, lngCount, IIf(blnMoved, 1, 0))
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

errHand:
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
    Dim mclsUnzip As New cUnzip
    
    On Error GoTo errHand
    
    If Not gobjFSO.FileExists(strZipFile) Then UnzipTendPage = "": Exit Function
    strZipPath = Left(strZipFile, Len(strZipFile) - Len(Dir(strZipFile)))
    
    strZipPath = GetTmpPath
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
errHand:
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

