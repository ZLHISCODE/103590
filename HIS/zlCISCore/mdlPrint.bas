Attribute VB_Name = "mdlPrint"
Option Explicit
'----------------------------------------------------------------------------------------------
'说明:
'1.该模块包含病历打印功能函数,为门诊、住院、护理公用
'2.该模块需要调用一些其它公共函数(如读取病历,API作图),必须保证这几个函数及相关子函数在其它模块中。
'3.该模块还包含体温表及麻醉记录单的打印过程
'----------------------------------------------------------------------------------------------
Public Const OFFSET_LEFT = 20
Public Const OFFSET_TOP = 20
Public Const OFFSET_RIGHT = 20
Public Const OFFSET_BOTTOM = 20

Private Const HOUR_STEP_Twips = 300 '用来决定第四小时之间的宽度 用于体温表
Private Const INTSTEPTwip = 90  '用来决定5分钟这间的宽度 用于麻醉单
Private Const STRING_WAY As String = "→"

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

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

'纸张打印边界控制================================================================
Public Declare Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, ByVal lpPort As String, ByVal iIndex As Long, ByVal lpOutput As String, lpDevMode As Any) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
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
Public Declare Function EnumForms Lib "winspool.drv" Alias "EnumFormsA" (ByVal hPrinter As Long, ByVal Level As Long, ByRef pForm As Any, ByVal cbBuf As Long, ByRef pcbNeeded As Long, ByRef pcReturned As Long) As Long
Public Declare Function AddForm Lib "winspool.drv" Alias "AddFormA" (ByVal hPrinter As Long, ByVal Level As Long, pForm As Byte) As Long
Public Declare Function DeleteForm Lib "winspool.drv" Alias "DeleteFormA" (ByVal hPrinter As Long, ByVal pFormName As String) As Long
Public Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, ByVal pDefault As Long) As Long
Public Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Public Declare Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal hwnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, pDevModeOutput As Any, pDevModeInput As Any, ByVal fMode As Long) As Long
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

' Constants for PRINTER_DEFAULTS.DesiredAccess
Public Const PRINTER_ACCESS_ADMINISTER = &H4
Public Const PRINTER_ACCESS_USE = &H8
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)

' Constants for DocumentProperties() call
Public Const DM_MODIFY = 8
Public Const DM_IN_BUFFER = DM_MODIFY
Public Const DM_COPY = 2
Public Const DM_OUT_BUFFER = DM_COPY

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
    CX As Long
    CY As Long
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
    Flags As Long
    pName As Long   ' String
    Size As SIZEL
    ImageableArea As RECTL
End Type

Public Type sFORM_INFO_1
    Flags As Long
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

Private Type CellInfoType
    FontSize As Long            '打印字体大小
    FontName As String          '打印字体
    FontBold As Boolean         '加粗
    FontItalic As Boolean       '斜体
    FontColor As OLE_COLOR      '文本颜色
    FontBackColor As OLE_COLOR  '单元格背景
    LineColor As OLE_COLOR      '边框颜色
    Text As String              '打印文本
    Merge As String             '合并单元格信息
    Height As Long              '高
    Width As Long               '宽
    HAlign As Byte
    VAlign As Byte
End Type

Private mCellArr() As CellInfoType

Dim mPageHeadDep As String                  '科室
Dim mPageHeadName As String                 '病人姓名
Dim mPageHeadNo As String                   '住院或门诊号
Dim mNewPageTop As Long                     '新页的初使化高度
Dim mNewPageInit As Long                    '新页初使化高度
Dim mPageBedNumber As String                '床号
Dim mPrintBegingPage As Long                '打印开始页
Dim mPrintEndPage As Long                   '打印结束页
Dim mPageNumber As Long                     '打印页号


'===================================================================================

Private Function GetCellWH(arrTmp() As CellInfoType, ByVal Row As Long, ByVal Col As Long, ByVal Row1 As Long, ByVal Col1 As Long, Optional blnRC As Boolean) As Long
    '得到从某一单元格到某一单元格矩形宽与高
    Dim i As Long
    Dim lngWH As Long  '临时记录宽或高
    
    lngWH = 0
    If Row > Row1 Then Exit Function
    If Col > Col1 Then Exit Function
    If blnRC Then    '求行高
        For i = LBound(arrTmp, 1) To UBound(arrTmp, 1)
            If i >= Row And i <= Row1 Then
                lngWH = lngWH + arrTmp(i, 1).Height
            End If
        Next
        GetCellWH = lngWH
    Else
        '求列宽
        For i = LBound(arrTmp, 2) To UBound(arrTmp, 2)
            If i >= Col And i <= Col1 Then
                lngWH = lngWH + arrTmp(1, i).Width
            End If
        Next
        GetCellWH = lngWH
    End If
End Function

Private Sub GridDraw(objOut As Object, objDraw As Object, ByVal lng病历 As Long, y As Long, ByVal blnPrintNO As Boolean, lngEndPage As Long, Optional ByVal bytGridAlign As Byte = 0)
    '功能:根据传入的病历ID从数据中读出数据并传入到数组中,再传给打印
    '分开的目地是为了将来可以共享使用DrawGridArr过程
    Dim arrTmp() As CellInfoType
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long
    On Error GoTo ErrHandle
    
    strSQL = "SELECT * From 病人病历所见单 " & vbCrLf & _
            "Where 病历id =  " & lng病历 & vbCrLf & _
            "ORDER BY 控件类,-行,-列"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "病历打印")
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        '第一行必定是表格的总体说明
        ReDim arrTmp(1 To rsTmp!行, 1 To rsTmp!列) As CellInfoType
        '下面的表格内容
        rsTmp.MoveNext
        If Not rsTmp.EOF Then
            For i = 1 To rsTmp.RecordCount - 1
                arrTmp(-rsTmp!行, -rsTmp!列).FontBackColor = -1
                arrTmp(-rsTmp!行, -rsTmp!列).FontBold = objDraw.Font.Bold
                arrTmp(-rsTmp!行, -rsTmp!列).FontColor = objDraw.ForeColor
                arrTmp(-rsTmp!行, -rsTmp!列).FontItalic = objDraw.Font.Italic
                arrTmp(-rsTmp!行, -rsTmp!列).FontName = objDraw.Font.Name
                arrTmp(-rsTmp!行, -rsTmp!列).FontSize = objDraw.Font.Size
                strSQL = Format(CStr(zlCommFun.Nvl(rsTmp!合并号, "")), "0000000000000000")
                arrTmp(-rsTmp!行, -rsTmp!列).Merge = IIf(strSQL = "0" Or strSQL = "0000000000000000", "", strSQL)
                arrTmp(-rsTmp!行, -rsTmp!列).Text = zlCommFun.Nvl(rsTmp!所见内容)
                arrTmp(-rsTmp!行, -rsTmp!列).Width = zlCommFun.Nvl(rsTmp!宽, 0)
                arrTmp(-rsTmp!行, -rsTmp!列).Height = zlCommFun.Nvl(rsTmp!高, 0)
                '由于对齐方式不同所以需要转换
                Select Case zlCommFun.Nvl(rsTmp!对齐, 1)
                    Case 2: arrTmp(-rsTmp!行, -rsTmp!列).HAlign = 0
                    Case 3: arrTmp(-rsTmp!行, -rsTmp!列).HAlign = 1
                    Case Else: arrTmp(-rsTmp!行, -rsTmp!列).HAlign = 2
                End Select
                arrTmp(-rsTmp!行, -rsTmp!列).VAlign = 1
                rsTmp.MoveNext
            Next
            Call DrawGridArr(objOut, objDraw, arrTmp, y, blnPrintNO, lngEndPage, bytGridAlign)
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub DrawGridArr(objOut As Object, objDraw As Object, arrTmp() As CellInfoType, y As Long, ByVal blnPrintNO As Boolean, lngEndPage As Long, Optional ByVal bytGridAlign As Byte = 0)
    '功能:根据传入的数组进行打印表格
    '参数:lngEndPage    开始的页号码
    '     ObjOut    用来传入预览窗体对象或打印机对象
    '     ObjDraw   用来传入打印输出对象
    '     arrTmp()  传入已经设置了值的单元格信息
    '     Y         设置开始从Y坐标开始画图
    '     bytGridAlign  设置表格整体对齐方式    (左0 中1 右2 )
    Dim i As Long
    Dim j As Long
    Dim m As Long   '合并开始单元格行
    Dim n As Long   '合并开始单元格列
    Dim m1 As Long  '合并终止单元格行
    Dim n1 As Long  '合并终止单元格列
    Dim x As Long
    Dim lngLeft As Long, lngRight As Long, lngTop As Long, lngBottom As Long, lngWidth As Long, lngHeight As Long   '纸张大小边界
    Dim lngGridW As Long, lngGridH As Long  '表格宽高
    Dim lngTmpPageNo As Long    '临时保存页号
    Dim TmpFont As New StdFont
    Dim strMerge As String
    On Error GoTo ErrHandle
    
    '得到纸张的边界与宽高
    lngLeft = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "左边距", OFFSET_LEFT) * 56.7 + Screen.TwipsPerPixelX * 2
    lngRight = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "右边距", OFFSET_RIGHT) * 56.7 - Screen.TwipsPerPixelX * 2
    lngTop = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "上边距", OFFSET_TOP) * 56.7 + Screen.TwipsPerPixelY * 2
    lngBottom = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "下边距", OFFSET_BOTTOM) * 56.7 - Screen.TwipsPerPixelY * 2
    lngWidth = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "宽度", Printer.Width)
    lngHeight = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "高度", Printer.Height)
    
    On Error Resume Next
    Err.Clear
    i = LBound(arrTmp, 1)
    i = LBound(arrTmp, 2)
    If Err.Number <> 0 Then Exit Sub
    On Error GoTo 0
    '得到表格的宽度与高度
    lngGridH = GetCellWH(arrTmp, 1, 1, UBound(arrTmp, 1), UBound(arrTmp, 2), True)
    lngGridW = GetCellWH(arrTmp, 1, 1, UBound(arrTmp, 1), UBound(arrTmp, 2), False)
    y = y + lngGridH
    '判断是否新页
    Set objDraw = Nothing
    If blnPrintNO Then
        Set objDraw = CheckNewPage(objOut, lngEndPage, y, lngGridH)
    Else
        lngTmpPageNo = 0
        Set objDraw = CheckNewPage(objOut, lngTmpPageNo, y, lngGridH)
    End If
    If objDraw Is Nothing Then Exit Sub

    '根据对齐方式设置X坐标
    Select Case bytGridAlign
        Case 1      '中对齐
            lngLeft = lngLeft + (lngWidth - (lngLeft + lngRight) - lngGridW) / 2
        Case 2      '右对齐
            lngLeft = lngWidth - (lngRight + lngGridW)
        Case Else   '左对齐
            lngLeft = lngLeft
    End Select
    '开始画
    For i = LBound(arrTmp, 1) To UBound(arrTmp, 1)
        For j = LBound(arrTmp, 2) To UBound(arrTmp, 2)
            '定位X
            If j = LBound(arrTmp, 2) Then
                x = lngLeft
            Else
                x = x + arrTmp(1, j - 1).Width
            End If
            '读出字体
            TmpFont.Bold = arrTmp(i, j).FontBold
            TmpFont.Italic = arrTmp(i, j).FontItalic
            TmpFont.Size = arrTmp(i, j).FontSize
            TmpFont.Name = arrTmp(i, j).FontName
            '读出合并信息进行解析
            strMerge = arrTmp(i, j).Merge
            If Len(strMerge) = 16 And IsNumeric(strMerge) And strMerge Like "0###0###0###0###" Then
                '例如格式:0006000100060008
                m = CLng(Mid(strMerge, 1, 4))
                n = CLng(Mid(strMerge, 5, 4))
                m1 = CLng(Mid(strMerge, 9, 4))
                n1 = CLng(Mid(strMerge, 13, 4))
                If i = m And j = n Then
                    Call DrawCell(objDraw, arrTmp(m, n).Text, x, y, GetCellWH(arrTmp, m, n, m1, n1, False), GetCellWH(arrTmp, m, n, m1, n1, True), IIf(blnPrintNO, lngEndPage, lngTmpPageNo), , , arrTmp(i, j).LineColor, arrTmp(i, j).FontColor, arrTmp(i, j).FontBackColor, TmpFont, "1111", arrTmp(i, j).HAlign, arrTmp(i, j).VAlign)
                End If
                arrTmp(i, j).Text = arrTmp(m, n).Text
            Else
                Call DrawCell(objDraw, arrTmp(i, j).Text, x, y, arrTmp(1, j).Width, arrTmp(i, 1).Height, IIf(blnPrintNO, lngEndPage, lngTmpPageNo), , , arrTmp(i, j).LineColor, arrTmp(i, j).FontColor, arrTmp(i, j).FontBackColor, TmpFont, "1111", arrTmp(i, j).HAlign, arrTmp(i, j).VAlign)
            End If
        Next
        '定位Y
        y = y + arrTmp(i, 1).Height
    Next
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

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

Public Function DrawCell(Dev As Object, ByVal Data As Variant, ByVal x As Long, ByVal y As Long, ByVal W As Long, ByVal H As Long, lngNowPage As Long, _
    Optional ByVal TW As Long, Optional ByVal TH As Long, Optional BorderColor As Long, _
    Optional ForeColor As Long, Optional BackColor As Long = &HFFFFFF, Optional ByVal Font As StdFont, _
        Optional Border As String = "1111", Optional HAlign As Byte, Optional VAlign As Byte = 1, Optional Warp As Boolean, _
        Optional Ratio As Single = 1) As Boolean
    '功能：在指定设备上按指定格式集输出文字或图象
    '参数：
    '   Dev=输出设备,为Printer或PictureBox对象
    '   Data=输出内容,为线条(x)、字符串("xxx")或图象(stdPicture)。字符串不包含vbCrLf,当Data类型为数字型时,表示输出线条
    '   TW,TH=输出的限定范围,超过这个范围则自动取消或缩小,为0时无效
    '   Border=边框定义,上下左右,"1111"表示全画
    '   Align=文字对齐,0=左,1=中,2=右,分水平对齐及垂直对齐
    '   Warp=当输出内容为字符串时,表示是否自动换行。不自动换行时,超宽部份不输出。
    '   Ratio=输出比例,对字体,坐标都有影响,缺省为1(100%)
    '说明：1.在使用该函数之前,应该没有改变设备的作图初始值
    '      2.输出后定位光标位置在本次输出范围的右上角
    Dim i As Long, Text As String, arrText() As String
    Dim LINE_W As Integer, blnW As Boolean, blnH As Boolean
    
    On Error GoTo errH
    
    DrawCell = True
    
    '范围限定
    If TW > 0 Then
        If x > TW Then Exit Function
        If x + W > TW Then W = TW - x
    End If
    If TH > 0 Then
        If y > TH Then Exit Function
        If y + H > TH Then H = TH - y
    End If
    
    If TypeName(Data) = "Integer" Then
        x = x * Ratio: y = y * Ratio: W = W * Ratio: H = H * Ratio
        If Val(Data) < 0 Then
            Dev.Line (x, y)-(x + W - IIf(W > 0, Screen.TwipsPerPixelX * Ratio, 0), y + H - IIf(H > 0, Screen.TwipsPerPixelY * Ratio, 0)), ForeColor, B '矩形
        Else
            Dev.Line (x, y)-(x + W - IIf(W > 0, Screen.TwipsPerPixelX * Ratio, 0), y + H - IIf(H > 0, Screen.TwipsPerPixelY * Ratio, 0)), ForeColor, BF '实心矩形(线条)
        End If
    ElseIf TypeName(Data) = "String" Then
        '字体
        If Font Is Nothing Then
            Set Font = New StdFont
            Font.Name = "宋体"
            Font.Size = 9
        End If
        '千万不要用Set Dev.Font=Font,不知为何,用的是ByVal
        Dev.Font.Name = Font.Name
        Dev.Font.Size = Font.Size
        Dev.Font.Bold = Font.Bold
        Dev.Font.Underline = Font.Underline
        Dev.Font.Italic = Font.Italic
        SetPrinterFont Dev.Font, Font.Size
        
        '因缩放后可能字体比例不对,判断时以原始大小为准
        If H >= Printer.TextHeight(Replace(Data, vbCrLf, "")) Then blnH = True '高度是否够用(加回车的算一行高度)
        If W >= Printer.TextWidth(Data) Then blnW = True And InStr(Data, vbCrLf) = 0   '宽度是否够用(加回车的为不够用,以便拆行)
        
        '缩变
        LINE_W = 30 * Ratio '边线间隔宽度(输出时用,判断时不用)
        x = -Int(-x * Ratio): y = -Int(-y * Ratio)
        W = -Int(-W * Ratio): H = -Int(-H * Ratio)
        Dev.Font.Size = Font.Size * Ratio
        SetPrinterFont Dev.Font, Font.Size
        
        '背景填充
        If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
        (lngNowPage - mPageNumber >= mPrintBegingPage - 1 And lngNowPage - mPageNumber <= mPrintEndPage) Then
            Dev.Line (x, y)-(x + W, y + H), BackColor, BF
        End If
        
        Dev.ForeColor = ForeColor
        '输出文字(边框之内再隔一线)
        '超出高度范围则不输出
        If blnH Then
            If blnW Then
                Select Case HAlign
                Case 0
                    Dev.CurrentX = x + LINE_W
                Case 1
                    Dev.CurrentX = x + (W - Printer.TextWidth(Data)) / 2
                Case 2
                    Dev.CurrentX = x + W - LINE_W - Printer.TextWidth(Data)
                End Select
                Select Case VAlign
                Case 0
                    Dev.CurrentY = y + LINE_W
                Case 1
                    Dev.CurrentY = y + (H - Printer.TextHeight(Data)) / 2 + LINE_W / 2
                Case 2
                    Dev.CurrentY = y + H - LINE_W - Printer.TextHeight(Data)
                End Select
                Dev.FontTransparent = True
                If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
                (lngNowPage - mPageNumber >= mPrintBegingPage - 1 And lngNowPage - mPageNumber <= mPrintEndPage) Then
                    Dev.Print Data
                End If
                Select Case VAlign
                Case 0
                    Dev.CurrentY = y + LINE_W + Printer.TextHeight(Date)
                Case 1
                    Dev.CurrentY = y + (H - Printer.TextHeight(Data)) / 2 + LINE_W / 2 + Printer.TextHeight(Date)
                Case 2
                    Dev.CurrentY = y + H - LINE_W - Printer.TextHeight(Data) + Printer.TextHeight(Date)
                End Select
            Else
                If Not Warp Then
                    '不自动拆行时超宽部分不输出
                    For i = 1 To Len(Data)
                        If Printer.TextWidth(Text & Mid(Data, i, 1)) > W Then Exit For
                        Text = Text & Mid(Data, i, 1)
                    Next
                    Select Case HAlign
                    Case 0
                        Dev.CurrentX = x + LINE_W
                    Case 1
                        Dev.CurrentX = x + (W - Printer.TextWidth(Text)) / 2
                    Case 2
                        Dev.CurrentX = x + W - LINE_W - Printer.TextWidth(Text)
                    End Select
                    Select Case VAlign
                    Case 0
                        Dev.CurrentY = y + LINE_W
                    Case 1
                        Dev.CurrentY = y + (H - Printer.TextHeight(Text)) / 2 + LINE_W / 2
                    Case 2
                        Dev.CurrentY = y + H - LINE_W - Printer.TextHeight(Text)
                    End Select
                    Dev.FontTransparent = True
                    '输出截取部份
                    If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
                    (lngNowPage - mPageNumber >= mPrintBegingPage - 1 And lngNowPage - mPageNumber <= mPrintEndPage) Then
                        Dev.Print Text
                    End If
                    Select Case VAlign
                    Case 0
                        Dev.CurrentY = y + LINE_W + Printer.TextHeight(Text)
                    Case 1
                        Dev.CurrentY = y + (H - Printer.TextHeight(Text)) / 2 + LINE_W / 2 + Printer.TextHeight(Text)
                    Case 2
                        Dev.CurrentY = y + H - LINE_W - Printer.TextHeight(Text) + Printer.TextHeight(Text)
                    End Select
                Else
                    '拆分文字成多行(在宽高范围内)
                    ReDim arrText(0) '在此,第一行不可能超高
                    Data = Replace(Data, vbCrLf, vbCr)
                    Data = Replace(Data, vbLf, vbCr)
                    For i = 1 To Len(Data)
                        If Mid(Data, i, 1) = vbCr Then
                            '多行超高则退出,超高部份不输出
                            If Printer.TextHeight(Replace(Data, vbCr, "")) * (UBound(arrText) + 2) > H Then Exit For
                            ReDim Preserve arrText(UBound(arrText) + 1)
                        ElseIf Printer.TextWidth(arrText(UBound(arrText)) & Mid(Data, i, 1)) > W Then
                            '多行超高则退出,超高部份不输出
                            If Printer.TextHeight(Replace(Data, vbCr, "")) * (UBound(arrText) + 2) > H Then Exit For
                            ReDim Preserve arrText(UBound(arrText) + 1)
                        End If
                        '有可能一行一个字符宽度都不够
                        If Printer.TextWidth(arrText(UBound(arrText)) & Mid(Data, i, 1)) <= W And Mid(Data, i, 1) <> vbCr Then
                            arrText(UBound(arrText)) = arrText(UBound(arrText)) & Mid(Data, i, 1)
                        End If
                    Next
                    
                    '输出起始坐标
                    Select Case VAlign
                    Case 0
                        Dev.CurrentY = y + LINE_W
                    Case 1
                        Dev.CurrentY = y + (H - Printer.TextHeight(Replace(Data, vbCr, "")) * (UBound(arrText) + 1)) / 2 + LINE_W / 2
                    Case 2
                        Dev.CurrentY = y + H - LINE_W - Printer.TextHeight(Replace(Data, vbCr, "")) * (UBound(arrText) + 1)
                    End Select
                    
                    '输出各行
                    For i = 0 To UBound(arrText)
                        Select Case HAlign
                        Case 0
                            Dev.CurrentX = x + LINE_W
                        Case 1
                            Dev.CurrentX = x + (W - Printer.TextWidth(arrText(i))) / 2
                        Case 2
                            Dev.CurrentX = x + W - LINE_W - Printer.TextWidth(arrText(i))
                        End Select
                        Dev.FontTransparent = True
                        If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
                        (lngNowPage - mPageNumber >= mPrintBegingPage - 1 And lngNowPage - mPageNumber <= mPrintEndPage) Then
                            Dev.Print arrText(i)
                        End If
                    Next
                    If UBound(arrText) > 0 Then
                        Select Case VAlign
                        Case 0
                            Dev.CurrentY = y + LINE_W + Printer.TextHeight(arrText(0))
                        Case 1
                            Dev.CurrentY = y + (H - Printer.TextHeight(Replace(Data, vbCr, "")) * (UBound(arrText) + 1)) / 2 + LINE_W / 2 + Printer.TextHeight(arrText(0))
                        Case 2
                            Dev.CurrentY = y + H - LINE_W - Printer.TextHeight(Replace(Data, vbCr, "")) * (UBound(arrText) + 1) + Printer.TextHeight(arrText(0))
                        End Select
                    End If
                End If
            End If
        End If
    ElseIf Not Data Is Nothing Then
        LINE_W = 30 * Ratio '边线间隔宽度(输出时用,判断时不用)
        x = x * Ratio: y = y * Ratio: W = W * Ratio: H = H * Ratio
        If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
        (lngNowPage - mPageNumber >= mPrintBegingPage - 1 And lngNowPage - mPageNumber <= mPrintEndPage) Then
            '图形(边框之内)
            Dev.PaintPicture Data, x + 15, y + 15, W - LINE_W, H - LINE_W
        End If
    End If
    
    If TypeName(Data) <> "Integer" Then
        '最后处理边框
        If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
        (lngNowPage - mPageNumber >= mPrintBegingPage - 1 And lngNowPage - mPageNumber <= mPrintEndPage) Then
            If Mid(Border, 1, 1) Then Dev.Line (x, y)-(x + W, y), BorderColor
            If Mid(Border, 2, 1) Then Dev.Line (x, y + H)-(x + W, y + H), BorderColor
            If Mid(Border, 3, 1) Then Dev.Line (x, y)-(x, y + H), BorderColor
            If Mid(Border, 4, 1) Then Dev.Line (x + W, y)-(x + W, y + H), BorderColor
        End If
    End If
    Exit Function
errH:
    DrawCell = False
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
    
    If Printers.Count = 0 Then Exit Function
    
    '初始化打印参数
    strPrinter = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "打印机", Printer.DeviceName)
    intPage = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "纸张", Printer.PaperSize)
    lngWidth = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "宽度", Printer.Width)
    lngHeight = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "高度", Printer.Height)
    intOrient = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "纸向", Printer.Orientation)
    intBin = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "进纸", Printer.PaperBin)
    
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
        If AddCustomPaper(objParent.hwnd, lngWidth / 56.7, lngHeight / 56.7) = FORM_NOT_SELECTED Then Exit Function
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
            If .Size.CX = FormSize.CX And .Size.CY = FormSize.CY Then
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
        
        ' Fill the DEVMODE from the printer.
        lngSize = DocumentProperties(lngHwnd, lngHandle, strPrtName, arrDevMode(1), 0&, DM_OUT_BUFFER)
        ' Copy the Public (predefined) portion of the DEVMODE.
        Call CopyMemory(vDevMode, arrDevMode(1), Len(vDevMode))
        
        ' If FormName is "zlBillPaper", we must make sure it exists
        ' before using it. Otherwise, it came from our EnumForms list,
        ' and we do not need to check first. Note that we could have
        ' passed in a Flag instead of checking for a literal name.
        
        ' Use form "zlBillPaper", adding it if necessary.
        ' Set the desired size of the form needed.
        ' Given in thousandths of millimeters
        vFormSize.CX = lngWidth * 1000 ' width
        vFormSize.CY = lngHeight * 1000 ' height
        
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
    Dim sTemp As String * 512, x As Long
    
    x = lstrcpy(sTemp, ByVal Add)
    If (InStr(1, sTemp, Chr(0)) = 0) Then
        PtrCtoVbString = ""
    Else
        PtrCtoVbString = Left(sTemp, InStr(1, sTemp, Chr(0)) - 1)
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
            .CX = vFormSize.CX
            .CY = vFormSize.CY
        End With
        With .ImageableArea
            .Left = 0
            .Top = 0
            .Right = FI1.Size.CX
            .Bottom = FI1.Size.CY
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

Public Sub PopupButtonMenu(ToolBar As Object, Button As Object, objMenu As Object)
    '功能：在下拉式工具按钮中弹出一个菜单
    Dim vRect As RECT, vDot1 As POINTAPI, vDot2 As POINTAPI
    
    Call GetWindowRect(ToolBar.hwnd, vRect)
    vDot1.x = vRect.Left: vDot1.y = vRect.Top
    vDot2.x = vRect.Right: vDot2.y = vRect.Bottom
    
    Call ScreenToClient(ToolBar.Parent.hwnd, vDot1)
    Call ScreenToClient(ToolBar.Parent.hwnd, vDot2)
    
    vDot1.x = vDot1.x * 15: vDot1.y = vDot1.y * 15
    vDot2.x = vDot2.x * 15: vDot2.y = vDot2.y * 15
    ToolBar.Parent.PopupMenu objMenu, 2, vDot1.x + Button.Left, vDot2.y
End Sub

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
                SetWindowPos frmFlash.hwnd, -1, (Screen.Width - frmFlash.Width) / 2 / Screen.TwipsPerPixelX, (Screen.Height - frmFlash.Height) / 2 / Screen.TwipsPerPixelY, 0, 0, 1
                ShowWindow frmFlash.hwnd, 5
            Else
                Err.Clear
                frmFlash.Show , frmParent
                If Err.Number <> 0 Then
                    Err.Clear
                    SetWindowPos frmFlash.hwnd, -1, (Screen.Width - frmFlash.Width) / 2 / Screen.TwipsPerPixelX, (Screen.Height - frmFlash.Height) / 2 / Screen.TwipsPerPixelY, 0, 0, 1
                    ShowWindow frmFlash.hwnd, 5
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

Public Function GetDeptName(lngID As Long) As String
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errH
    
    strSQL = "Select * From 部门表 Where ID=" & lngID
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlPrint")
    If rsTmp.RecordCount > 0 Then GetDeptName = rsTmp!名称
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetLastPrint(ByVal lng病历记录ID As Long, ByRef lngEndY As Long, ByRef lngEndPage As Long) As Boolean
    '功能：读取病人上次病历打印结束位置信息
    '返回：lngEndY=上次打印的结束位置(mm)
    '      intEndPage=上次打印的结束页号
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errH
    
    strSQL = "Select * From 病历打印记录 Where 病历记录ID=" & lng病历记录ID & " Order By 打印时间 Desc"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "病历打印")
    
    If rsTmp.RecordCount > 0 Then
        lngEndY = rsTmp!结束位置
        lngEndPage = rsTmp!结束页号
    Else
        lngEndY = 0
        lngEndPage = 1
    End If
    GetLastPrint = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function NewPrintPage(objOut As Object, intPage As Long, Optional blnNewPage As Boolean = True) As Object
    '功能：打印或预览一页结束时对当前页作结束处理,并产生新页
    '参数：blnNewPage=为False时仅打印页号等,一般打印结束才这样处理,因此不管最后坐标
    '返回：新页对象,可能为打印机或PictureBox
    Dim objDraw As Object, blnPrint As Boolean
    Dim lngWidth As Long, lngHeight As Long, lngOldY As Long
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    Dim strFontName As String, lngFontSize As Long, blnFontBold As Boolean
    Dim blnFontItalic As Boolean, lngFontColor As Long
    Dim x As Long, y As Long, H_9pt As Long, W_9pt As Long
    Dim strText As String
    Dim objPrinter As Object
    On Error GoTo errH
    
    blnPrint = TypeName(objOut) = "Printer"
    
    '边界信息(Twip)
    lngLeft = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "左边距", OFFSET_LEFT) * 56.7
    lngRight = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "右边距", OFFSET_RIGHT) * 56.7
    lngTop = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "上边距", OFFSET_TOP) * 56.7 + Screen.TwipsPerPixelY * 2
    lngBottom = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "下边距", OFFSET_BOTTOM) * 56.7
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
    
    '产生新页
    If blnNewPage Then
        '打印页号(0为不打印)
        If intPage <> 0 Then
            objDraw.ForeColor = 0
            objDraw.Font.Name = "宋体"
            objDraw.Font.Size = 9
            objDraw.Font.Bold = False
            SetPrinterFont objDraw.Font, 9
            objDraw.CurrentY = lngHeight - IIf(lngBottom < 1134, 1134, lngBottom) - (Printer.TextHeight("字") * 2)
            objDraw.CurrentX = lngLeft + (lngWidth - lngLeft - lngRight) * (3 / 4)
            objDraw.FontTransparent = True
            If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
                (intPage - mPageNumber >= mPrintBegingPage - 1 And intPage - mPageNumber <= mPrintEndPage) Then
                objDraw.Print "·第 " & intPage & " 页·"
            End If
        End If
        intPage = intPage + 1
        If intPage - mPageNumber + 1 > mPrintEndPage And mPrintEndPage > 0 Then Set objDraw = Nothing: Exit Function
        If blnPrint Then
            '处理指定页
            If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
                (intPage - mPageNumber >= mPrintBegingPage - 1 And intPage - mPageNumber <= mPrintEndPage) Then
                If (intPage - mPageNumber) >= mPrintBegingPage Then
                    Printer.NewPage
                End If
                Set objDraw = Printer
            Else
'                Printer.KillDoc
'                InitPrint Printer
'                Printer.EndDoc
            End If
        Else
'            If intPage - mPageNumber >= mPrintBegingPage And intPage - mPageNumber <= mPrintEndPage Then
'                Load objOut.picPage(objOut.picPage.UBound + 1)
'            End If
            '预览打印边线
            objDraw.DrawStyle = 2
            objDraw.Line (0, lngTop)-(lngWidth, lngTop), &H808080
            objDraw.Line (0, lngHeight - lngBottom)-(lngWidth, lngHeight - lngBottom), &H808080
            objDraw.Line (lngLeft, 0)-(lngLeft, lngHeight), &H808080
            objDraw.Line (lngWidth - lngRight, 0)-(lngWidth - lngRight, lngHeight), &H808080
            objDraw.DrawStyle = 0
    
            If intPage - mPageNumber >= mPrintBegingPage Then
                Load objOut.picPage(objOut.picPage.UBound + 1)
            End If
            Set objDraw = objOut.picPage(objOut.picPage.UBound)
            objDraw.Width = Printer.Width
            objDraw.Height = Printer.Height
            objDraw.ZOrder
            objDraw.Cls
            objDraw.AutoRedraw = True
        End If
        '新页起点
        objDraw.CurrentX = lngLeft: objDraw.CurrentY = lngTop
        '--曾超修改
'        '预览打印边线
'        objDraw.DrawStyle = 2
'        objDraw.Line (0, lngTop)-(lngWidth, lngTop), &H808080
'        objDraw.Line (0, lngHeight - lngBottom)-(lngWidth, lngHeight - lngBottom), &H808080
'        objDraw.Line (lngLeft, 0)-(lngLeft, lngHeight), &H808080
'        objDraw.Line (lngWidth - lngRight, 0)-(lngWidth - lngRight, lngHeight), &H808080
'        objDraw.DrawStyle = 0
'        objDraw.CurrentY = IIf(mNewPageInit >= lngTop, mNewPageInit, lngTop)
        objDraw.CurrentY = lngTop
        objDraw.Font.Name = "宋体"
        objDraw.Font.Size = 9
        objDraw.Font.Bold = False
        SetPrinterFont objDraw.Font, 9
        H_9pt = Printer.TextHeight("江")
        W_9pt = Printer.TextWidth("江")
        '打印当前病历标题信息
        objDraw.ForeColor = 0
        objDraw.Font.Name = "黑体"
        objDraw.Font.Size = 18
        objDraw.Font.Bold = True
        SetPrinterFont objDraw.Font, 18
        '病历标题
        strText = GetSetting("ZLSOFT", "注册信息", "单位名称", "单位名称")
        '判断是否新页
        y = objDraw.CurrentY + H_9pt + Printer.TextHeight("江")  '起始下移2个字高
        y = y - Printer.TextHeight("江")
        '得到标题XY坐标
        x = lngLeft + (lngWidth - (lngLeft + lngRight) - Printer.TextWidth(strText)) / 2
        objDraw.CurrentX = x: objDraw.CurrentY = y
        objDraw.FontTransparent = True
        If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
            (intPage - mPageNumber >= mPrintBegingPage - 1 And intPage - mPageNumber <= mPrintEndPage) Then
            objDraw.Print strText
        End If
        objDraw.CurrentY = y + Printer.TextHeight(strText)
        
        '打印病历科室,签名,日期
        objDraw.ForeColor = 0
        objDraw.Font.Name = "宋体"
        objDraw.Font.Size = 10.5
        objDraw.Font.Bold = False
        SetPrinterFont objDraw.Font, 10.5
        '判断是否新页
        y = objDraw.CurrentY + H_9pt * 2 + Printer.TextHeight("江") '起始下移2个字高
        y = y - Printer.TextHeight("江")
        strText = mPageHeadDep
        
        x = lngLeft
        objDraw.CurrentX = x: objDraw.CurrentY = y
        objDraw.FontTransparent = True
        If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
            (intPage - mPageNumber >= mPrintBegingPage - 1 And intPage - mPageNumber <= mPrintEndPage) Then
            objDraw.Print strText
        End If
        
        strText = mPageHeadName
        x = lngLeft + (lngWidth - lngLeft - lngRight) * (2 / 9)
        objDraw.CurrentX = x: objDraw.CurrentY = y
        objDraw.FontTransparent = True
        If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
            (intPage - mPageNumber >= mPrintBegingPage - 1 And intPage - mPageNumber <= mPrintEndPage) Then
            objDraw.Print strText
        End If
        
        strText = mPageHeadNo
        x = lngLeft + (lngWidth - lngLeft - lngRight) * (4 / 9)
        objDraw.CurrentX = x: objDraw.CurrentY = y
        objDraw.FontTransparent = True
        If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
            (intPage - mPageNumber >= mPrintBegingPage - 1 And intPage - mPageNumber <= mPrintEndPage) Then
            objDraw.Print strText
        End If
        
        strText = mPageBedNumber
        x = lngLeft + (lngWidth - lngLeft - lngRight) * (7 / 9)
        objDraw.CurrentX = x: objDraw.CurrentY = y
        objDraw.FontTransparent = True
        If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
            (intPage - mPageNumber >= mPrintBegingPage - 1 And intPage - mPageNumber <= mPrintEndPage) Then
            objDraw.Print strText
        End If
        objDraw.CurrentY = y + Printer.TextHeight(strText)
        y = objDraw.CurrentY + H_9pt / 5: x = lngLeft
        If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
            (intPage - mPageNumber >= mPrintBegingPage - 1 And intPage - mPageNumber <= mPrintEndPage) Then
            objDraw.Line (lngLeft, y)-(lngWidth - lngRight, y), 0
        End If
        '--
        objDraw.Font.Name = strFontName
        objDraw.Font.Size = lngFontSize
        objDraw.Font.Bold = blnFontBold
        objDraw.Font.Italic = blnFontItalic
        objDraw.ForeColor = lngFontColor
        SetPrinterFont objDraw.Font, Int(Nvl(lngFontSize, 0))
        
        mNewPageTop = y + 100
        
    Else
        objDraw.CurrentY = lngOldY
        If intPage <> 0 Then
            objDraw.ForeColor = 0
            objDraw.Font.Name = "宋体"
            objDraw.Font.Size = 9
            objDraw.Font.Bold = False
            SetPrinterFont objDraw.Font, 9
            objDraw.CurrentY = lngHeight - IIf(lngBottom < 1134, 1134, lngBottom) - (Printer.TextHeight("字") * 2)
            objDraw.CurrentX = lngLeft + (lngWidth - lngLeft - lngRight) * (3 / 4)
            objDraw.FontTransparent = True
            If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
            (intPage - mPageNumber >= mPrintBegingPage - 1 And intPage - mPageNumber <= mPrintEndPage) Then
                objDraw.Print "·第 " & intPage & " 页·"
            End If
        End If
    End If
    Set NewPrintPage = objDraw
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ReadPatiInfo(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As ADODB.Recordset
    '在病历打印时得到病人的信息
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    If lng主页ID = 0 Then
        strSQL = "Select A.病人ID,A.门诊号,A.姓名,A.性别,A.年龄,A.国籍,A.民族," & _
        " A.婚姻状况,A.职业,A.身份证号,A.工作单位,A.家庭地址" & _
            " From 病人信息 A Where 病人ID=" & lng病人ID
    Else
        strSQL = "Select A.病人ID,A.住院号,A.姓名,A.性别,A.年龄,A.国籍,A.民族," & _
        " A.婚姻状况,A.职业,A.身份证号,A.工作单位,A.家庭地址," & _
            " B.入院日期,C.名称 as 入院科室,B.出院日期,D.名称 as 出院科室,B.出院病床" & _
            " From 病人信息 A,病案主页 B,部门表 C,部门表 D" & _
            " Where B.入院科室ID=C.ID And B.出院科室ID=D.ID And A.病人ID=B.病人ID" & _
            " And A.病人ID=" & lng病人ID & " And B.主页ID=" & lng主页ID
    End If
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "病历打印")
    If rsTmp.RecordCount > 0 Then
        Set ReadPatiInfo = rsTmp
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckNewPage(objOut As Object, lngPage As Long, lngY As Long, Optional lngDefHeight As Long = -1) As Object
    '功能：检测是不是超出边界，并把新开始一页对象返回
    '参数：ObjOut       输出对象
    '       lngPage     页号为0时表示不打印
    '       lngDefY     可打印顶点坐标
    '       lngY        当前打印位置
    Dim lngBottom As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim lngTop As Long
    Dim lngFontHeigh As Long
    Dim lngTmp As Long
    
    
    '边界信息(Twip)
    lngBottom = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "下边距", OFFSET_BOTTOM) * 56.7
    lngTop = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "上边距", OFFSET_TOP) * 56.7
    
    lngWidth = Printer.Width
    lngHeight = Printer.Height
    If lngPage > 0 Then
        lngTmp = Printer.TextHeight("江") * 2
        If UCase(TypeName(objOut)) = UCase("Printer") Then
'            lngTmp = lngTmp - 50
        End If
    End If
    
    If lngY > lngHeight - IIf(lngBottom < 1134, 1134, lngBottom) - lngTmp Then
        Set CheckNewPage = NewPrintPage(objOut, lngPage, True)
'        CheckNewPage.Width = lngWidth: CheckNewPage.Height = lngHeight
        lngTop = mNewPageTop
        lngY = lngTop
    Else
        If lngDefHeight = -1 Then
            lngFontHeigh = Printer.TextHeight("江")
        Else
            lngFontHeigh = lngDefHeight
        End If
        lngY = lngY - lngFontHeigh
        If UCase(TypeName(objOut)) = UCase("Printer") Then
            Set CheckNewPage = Printer
        Else
            Set CheckNewPage = objOut.picPage(objOut.picPage.UBound)
        End If
    End If
End Function

Private Function PrintLineS(objDraw As Object, ByVal strChars As String, ByVal lngLeft As Long, ByVal lngRight As Long, lngNowPage As Long) As String
    '本函数是供病历打印之用:  根据当前位置打印文本并且(根据回画符或边界)返回剩下的字符串以便下次打印
    '
    Dim lngTmp As Long
    Dim strTmp As String
    Dim i As Long
    Dim lngWidth As Long
    Dim y As Long
    On Error GoTo ErrHandle
    
    lngWidth = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "宽度", Printer.Width)
    PrintLineS = ""
    
    '不去掉回车符
    strChars = Replace(strChars, vbCrLf, vbCr)
    strChars = Replace(strChars, vbLf, vbCr)
    strTmp = ""
    lngRight = lngWidth - lngRight
    '循环读字符
    For i = 1 To Len(strChars)
        '读出右位置
        '读出左右位置值
        lngTmp = lngRight - lngLeft
        strTmp = strTmp & Mid(strChars, i, 1)
        '当加上下一个字符的宽度超过边界宽度或者下一个字符是回画符时就开始打印这一行,并将剩下的字符返回
        If (Printer.TextWidth(strTmp) > lngTmp Or Mid(strChars, i, 1) = vbCr) And _
            InStr("！％）、｝〕；：，。》？!%)}];:,.>?", Mid(strChars, i, 1)) = 0 Then
            If i = 1 Then
                PrintLineS = Mid(strChars, i + 1)
            Else
                strTmp = Left(strChars, i - 1)
                If Mid(strChars, i, 1) = vbCr Then
                    If i + 1 > Len(strChars) Then
                        PrintLineS = ""
                    Else
                        PrintLineS = Mid(strChars, i + 1)
                    End If
                Else
                    PrintLineS = Mid(strChars, i)
                End If
            End If
            '此时可以退出循环开始打印
            Exit For
        End If
    Next
    '开始打印这行字符串
    objDraw.CurrentX = lngLeft
    y = objDraw.CurrentY
    If strTmp <> "" Then
        objDraw.FontTransparent = True
        If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
        (lngNowPage - mPageNumber >= mPrintBegingPage - 1 And lngNowPage - mPageNumber <= mPrintEndPage) Then
            objDraw.Print strTmp
        End If
        objDraw.CurrentY = y + Printer.TextHeight(strTmp)
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function PrintOutCase(objParent As Object, objOut As Object, ByVal lng病历种类 As Long, ByVal blnCurCase As Boolean, ByVal lngCurCase As Long, ByVal lng病人ID As Long, _
    ByVal var主页或单据 As Variant, ByVal blnPatiInfo As Boolean, ByVal lngY As Long, Optional ByVal lng页码 As Long = 0, Optional ByVal lng开始页 As Long = 0, Optional ByVal lng结束页 As Long = 0) As Boolean
    '功能：打印所有病历
    '参数：ObjParent        所有者对象
    '       ObjOut          输出对象（是打印机还是预览窗体）
    '       lng病历种类     指定病历种类
    '       blnCurCase      是否为只打印输出当前这页
    '       lngCurCase      指定当前打印输出的那份病历，打印输出时就从那份往后打印输出
    '                       负数时表示病历记录ID
    '       lng病人id       如果是打印病历示范那么,这个病历ID为0,并且 var主页或单据 就为病历记录ID
    '       var主页或单据   如果是住院病人就记录主页ID，如果是门诊病人就记录挂号单，通过参数类型判断是住院还是门诊
    '       blnPatiInfo     是否打印病人信息
    '       lngY            打印开始的Y坐标
    '       lng页码         设置起始的页码,为0时表示不打印页码
    
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim rsNewTmp As New ADODB.Recordset
    Dim i As Long
    Dim lngPrintPageNO As Long
    Dim lngBPage As Long    '开始页
    Dim lngEPage As Long   '结束页
    Dim sngBeginY As Single '开始位置
    Dim sngEndY As Single '结束位置
    Dim lngTop As Long
    Dim lngHeight As Long
    Dim lngNewPage As Boolean           '上次是否新起一页打印
    Dim IntOnePage As Boolean           '是否第一页
    Dim bNowPrint As Boolean            '是否打印第一页
        
    mPrintBegingPage = lng开始页
    mPrintEndPage = lng结束页
    mPageNumber = lng页码
        
    On Error GoTo ErrHandle
    '得到最上边距,为新页做好准备
    lngTop = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "上边距", OFFSET_TOP) * 56.7
    lngHeight = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "高度", Printer.Height)
    
    '初始化打印参数
    If InitPrint(objParent) = False Then
        MsgBox "打印机初始化失败（无任何可用打印机或无法设置打印纸张）！", vbExclamation, gstrSysName
        Exit Function
    End If
    lngHeight = Printer.Height
    
    If lngCurCase < 0 Then
        '负数时表示病历记录ID
        strSQL = "select a.id,a.书写日期,b.新页 from 病人病历记录 a ,病历文件目录 b where a.文件ID = b.ID And a.ID=" & (-1 * lngCurCase)
        blnCurCase = True: lngCurCase = 1: var主页或单据 = ""
    Else
        If lng病人ID > 0 Then
            If VarType(var主页或单据) = vbString Then
                strSQL = "select a.id,a.书写日期,b.新页 from 病人病历记录 a ,病历文件目录 b where a.病历种类 = " & lng病历种类 & " and a.文件ID = b.ID and a.病人ID = " & lng病人ID & " AND a.挂号单='" & var主页或单据 & "' order by b.打印顺序,a.书写日期 "
            ElseIf IsNumeric(var主页或单据) Then
                strSQL = "select a.id,a.书写日期,b.新页 from 病人病历记录 a ,病历文件目录 b where a.病历种类 = " & lng病历种类 & " and a.文件ID = b.ID and a.病人ID = " & lng病人ID & " AND a.主页ID=" & var主页或单据 & "  order  by b.打印顺序,a.书写日期 "
            Else
                MsgBox "打印病人病历时参数不正确，不能继续！", vbExclamation, gstrSysName
                Exit Function
            End If
        Else
            '这里处理那些对病历示范的打印
            If IsNumeric(var主页或单据) Then
                If 1 * var主页或单据 > 0 Then
                    strSQL = "Select * from 病历示范目录 where ID = " & var主页或单据
                Else
                    strSQL = "Select -1*ID As ID,病历记录ID from 病人病历修订记录 where ID = " & -1 * var主页或单据
                End If
            Else
                MsgBox "打印病历示范文件不存在，不能继续！", vbExclamation, gstrSysName
                Exit Function
            End If
        End If
    End If
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "病历打印")
    With rsTmp
        If .RecordCount > 0 Then
            .MoveFirst
            lngPrintPageNO = lng页码
            sngEndY = lngY
            IntOnePage = False
            For i = 1 To .RecordCount  '读每一份病历，将参数传给打印病历的过程，并返回参数来打印下一份病历
                '得到打印前的页码与位置
                If lng病人ID > 0 And IsNumeric(var主页或单据) Then
                    '只有住院病历和护理病历才有在打印病历时新起一页的可能,所有在病历种类加了判断 "病历种类 in (2,3)"
                    strSQL = "select a.病人ID,a.主页ID,a.挂号单,nvl(b.新页,0) 新页 from 病人病历记录 a,病历文件目录 b where a.文件ID=b.id and  a.病历种类 in (2,3) and a.id=" & zlCommFun.Nvl(rsTmp!ID, 0)
                    Call zlDatabase.OpenRecordset(rsNewTmp, strSQL, "病历打印")
                    If rsNewTmp.RecordCount > 0 Then
                        '只有那个病历可以是新页的,并且,不是第一页,并且病历顺序号是大于指定病历顺序号的病历才可以新起一页
                        If (rsNewTmp!新页 = 1 And i > 1 And blnCurCase = False And i > lngCurCase) Or lngNewPage = True Then
                            '要新起一页的
                            sngBeginY = lngTop
                            If PrintOutCase = True Then
                                sngEndY = lngHeight - lngTop
                            End If
                            lngBPage = lngPrintPageNO + 1   '确保在保存时将那份病历的页号新起一页
                            IntOnePage = True
                        Else
                            sngBeginY = sngEndY
                            lngBPage = lngPrintPageNO
                            If rsNewTmp!新页 = 1 And i = 1 And blnCurCase = False Then IntOnePage = True
                        End If
                    Else
                        sngBeginY = sngEndY
                        lngBPage = lngPrintPageNO
                    End If
                    If Not rsNewTmp.EOF Then
                        If rsNewTmp!新页 = 1 And IntOnePage <> False Then
                            lngNewPage = True
                        Else
                            lngNewPage = False
                        End If
                    Else
                        lngNewPage = False
                    End If
                Else
                    '门诊的及示范的都要连续打印
                    sngBeginY = sngEndY
                    lngBPage = lngPrintPageNO
                End If
                '=====================================================================================
                '如果只输出指定的那份病历时在输出后就退出
                If blnCurCase = True Then
                    If i = lngCurCase Then
                        '开始打印指定病历
                        If UCase(TypeName(objParent)) = UCase("FRMCASEPRINT") Then
                            zlCommFun.ShowFlash "共" & .RecordCount & "份病历，正打印第" & i & "份， 请稍候... ..."
                        Else
                            zlCommFun.ShowFlash "共" & .RecordCount & "份病历，正打印第" & i & "份， 请稍候... ...", objParent
                        End If
                        If lng病人ID > 0 Then
                            If PrintOrPreviewCase(objParent, objOut, !ID, blnPatiInfo, lng页码 > 0, lngPrintPageNO, sngEndY) Then
                                '得到打印后的码与位置（位置本身就是打印后的位置）
                                lngEPage = lngPrintPageNO
                                '单个打印病历时不保存病历位置
                                'If UCase(TypeName(objOut)) = "PRINTER" Then
                                '    strSQL = "zl_病历打印记录_insert(" & zlCommFun.NVL(!病人ID, 0) & "," & zlCommFun.NVL(!主页ID, 0) & ",'" & zlCommFun.NVL(!挂号单) & "'," & lngBPage & "," & lngEPage & "," & sngBeginY & "," & sngEndY & ",'" & UserInfo.姓名 & "')"
                                '    Call zlDatabase.ExecuteProcedure(strSQL, "病历打印")
                                'End If
                                PrintOutCase = True
                                bNowPrint = True
                            End If
                        Else
                            If PrintOrPreviewCase(objParent, objOut, !ID, False, lng页码 > 0, lngPrintPageNO, sngEndY, 1 * var主页或单据 > 0) Then
                                '得到打印后的码与位置（位置本身就是打印后的位置）
                                lngEPage = lngPrintPageNO
                                PrintOutCase = True
                                bNowPrint = True
                            End If
                        End If
                        zlCommFun.StopFlash
                        Exit Function
                    End If
                ElseIf i >= lngCurCase Then
                    '开始打印每份病历
                    If UCase(TypeName(objParent)) = UCase("FRMCASEPRINT") Then
                        zlCommFun.ShowFlash "共" & .RecordCount & "份病历，正打印第" & i & "份， 已完成" & Format((i - 1) / .RecordCount, "0.00") * 100 & "%"
                    Else
                        zlCommFun.ShowFlash "共" & .RecordCount & "份病历，正打印第" & i & "份， 已完成" & Format((i - 1) / .RecordCount, "0.00") * 100 & "%", objParent
                    End If
                    If lng病人ID > 0 Then
                        If PrintOrPreviewCase(objParent, objOut, !ID, blnPatiInfo, lng页码 > 0, lngPrintPageNO, sngEndY) Then
                            '得到打印后的码与位置（位置本身就是打印后的位置）
                            lngEPage = lngPrintPageNO
'                            If UCase(TypeName(objOut)) = "PRINTER" Then
                                strSQL = "zl_病历打印记录_insert(" & !ID & "," & lngBPage & "," & lngEPage & "," & sngBeginY & "," & sngEndY & ",'" & UserInfo.姓名 & "')"
                                Call zlDatabase.ExecuteProcedure(strSQL, "病历打印")
'                            End If
                            PrintOutCase = True
                        Else
                            zlCommFun.StopFlash
                            '打印失败直接退出
                            Exit Function
                        End If
                    Else
                        If PrintOrPreviewCase(objParent, objOut, !ID, False, lng页码 > 0, lngPrintPageNO, sngEndY, 1 * var主页或单据 > 0) Then
                            '得到打印后的码与位置（位置本身就是打印后的位置）
                            lngEPage = lngPrintPageNO
                            PrintOutCase = True
                        Else
                            zlCommFun.StopFlash
                            '打印失败直接退出
                            Exit Function
                        End If
                    End If
                End If
                .MoveNext
            Next
        Else
            MsgBox "无任何可打印的病历！", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    zlCommFun.StopFlash
    PrintOutCase = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function PrintOrPreviewCase(objParent As Object, objOut As Object, ByVal lng病历记录ID As Long, ByVal blnPatiInfo As Boolean, _
    ByVal blnPrintNO As Boolean, lngEndPage As Long, sngEndY As Single, Optional ByVal blnDemo As Boolean = False) As Boolean
    '功能:打印指定病历内容
    '参数:  ObjParent       所都对象
    '       objOut          输出对象
    '       lng病历记录ID
    '       blnPatiInfo     是否打印病人信息
    '       blnPrintNO      是否打印页码
    '       lngEndPage      上次的页码,并返回本次打印的最后页码
    '       sngEndY         上次的打印位置,并返回本次的打印位置
    '       blnDemo         用来表示是不是病历示范,如果是那么 lng病历记录ID 就表示病历示范ID
    
    On Error GoTo ErrHandle
    Dim rsTmp As New ADODB.Recordset, rsNewTmp1 As New ADODB.Recordset, rsNewTmp2 As New ADODB.Recordset
    Dim str元素编码 As String
    Dim strSQL As String, i As Long, j As Long, m As Long
    Dim ObjStdPic As New StdPicture, lngStdPicWidth As Long, lngStdPicHeight As Long, dblPic比例 As Double   '取图片的比例
    Dim objDraw As Object   '画图输出对象可能是打印机也可能是图片控件
    Dim blnPrint As Boolean '判断是不是打印机
    Dim x As Long, y As Long, TmpX As Long, TmpY As Long, H_9pt As Long, W_9pt As Long, Tmp_W As Long, Tmp_H As Long
    Dim strFontName As String, strFontSize As String, strFontBold As String, strFontItalic As String
    Dim strTitleFontName As String, strTitleFontSize As String, strTitleFontItalic As String, strTitleFontBold As String, strTitleAlig As String   '标题的对齐方式
    Dim strText As String
    Dim lng病人ID As Long, lng主页ID As Long, blnOutPati As Boolean   '是住院还是门诊
    Dim lngLeft As Long, lngRight As Long, lngTop As Long, lngBottom As Long, lngWidth     As Long, lngHeight  As Long
    Dim lngTmpPageNo As Long    '记录临时的页号
    Dim lngPageTmp  As Long     '用于记录临时的页号用于判断
    Dim tmpPage As Integer, tmpPrintHeight As Long      '临时记录变量
    Dim blOnePrintText As Boolean                       '是否第一次打印文本段
    Dim blnMultiSign As Boolean '是否有多次签名
    
    '得到纸张的边界与宽高
    lngLeft = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "左边距", OFFSET_LEFT) * 56.7 + Screen.TwipsPerPixelX * 2
    lngRight = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "右边距", OFFSET_RIGHT) * 56.7 - Screen.TwipsPerPixelX * 2
    lngTop = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "上边距", OFFSET_TOP) * 56.7 + Screen.TwipsPerPixelY * 2
    lngBottom = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "下边距", OFFSET_BOTTOM) * 56.7 - Screen.TwipsPerPixelY * 2
    lngWidth = Printer.Width ' GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "宽度", Printer.Width)
    lngHeight = Printer.Height ' GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "高度", Printer.Height)
    If blnDemo = False Then
        '判断病历是不是存在，并读出病人ID
        If lng病历记录ID > 0 Then
            strSQL = "select a.*,b.名称 科室名称,c.姓名,c.住院号,c.门诊号,d.出院病床 As 当前床号 from 病人病历记录 a,部门表 b , 病人信息 c,病案主页 d" & _
                " Where a.科室ID=b.id(+) and a.病人id = c.病人id" & _
                " And a.病人ID=d.病人ID(+) And a.主页ID=d.主页ID(+)" & _
                " And a.id =" & lng病历记录ID
        Else
            strSQL = "select a.*,b.名称 科室名称,d.姓名,d.住院号,d.门诊号,e.出院病床 As 当前床号 from 病人病历记录 a,部门表 b,病人病历修订记录 c ,病人信息 d,病案主页 e" & _
                " Where a.科室ID=b.id(+) and a.id=c.病历记录id and a.病人id = d.病人id" & _
                " And a.病人ID=e.病人ID(+) And a.主页ID=e.主页ID(+)" & _
                " And c.id =" & -1 * lng病历记录ID
        End If
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, "病历打印")
        If rsTmp.RecordCount > 0 Then
            lng病人ID = zlCommFun.Nvl(rsTmp!病人id, 0)
            lng主页ID = zlCommFun.Nvl(rsTmp!主页ID, 0)
            '真为门诊
            If lng主页ID = 0 Then: blnOutPati = True: Else blnOutPati = False
        Else
            MsgBox "指定病历不存在！", vbExclamation, gstrSysName: Exit Function
        End If
    Else
        '病历示范
        strSQL = "select * from 病人病历内容 where 病历示范ID=" & lng病历记录ID & " order by 排列序号"
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, "病历打印")
        If rsTmp.RecordCount < 1 Then
            MsgBox "指定病历示范无任何内容！", vbExclamation, gstrSysName: Exit Function
        End If
        strSQL = "select a.名称 病历名称,a.制定人 书写人,a.制定日 书写日期 ,b.名称 科室名称  FROM 病历示范目录 a, 部门表 b where  a.科室ID=b.id(+) and  a.id=" & lng病历记录ID
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, "病历打印")
        If rsTmp.RecordCount < 1 Then
            MsgBox "指定病历示范不存在！", vbExclamation, gstrSysName: Exit Function
        End If
    End If
    '初始化一个新页面
    Set objDraw = Nothing
    If blnPrintNO Then
        Set objDraw = NewPrintPage(objOut, lngEndPage, False)
    Else
        lngTmpPageNo = 0
        Set objDraw = NewPrintPage(objOut, lngEndPage, False)
    End If
    '检查打印机
    If objDraw Is Nothing Then
        MsgBox "打印机出错，退出打印！", vbExclamation, gstrSysName: Exit Function
    End If
    blnPrint = UCase(TypeName(objDraw)) = "PRINTER"
    If blnPrint = False Then
        objDraw.Width = lngWidth: objDraw.Height = lngHeight
    End If
    mNewPageInit = sngEndY
    objDraw.CurrentY = IIf(sngEndY >= lngTop, sngEndY, lngTop)
    
    
    objDraw.Font.Name = "宋体"
    objDraw.Font.Size = 9
    objDraw.Font.Bold = False
    SetPrinterFont objDraw.Font, 9
    H_9pt = Printer.TextHeight("江")
    W_9pt = Printer.TextWidth("江")
    '打印当前病历标题信息
    objDraw.ForeColor = 0
    objDraw.Font.Name = "黑体"
    objDraw.Font.Size = 18
    objDraw.Font.Bold = True
    SetPrinterFont objDraw.Font, 18
    '病历标题
    If blnDemo = False Then
        strText = GetSetting("ZLSOFT", "注册信息", "单位名称", "单位名称")
    Else
        strText = zlCommFun.Nvl(rsTmp!病历名称)
    End If
    '判断是否新页
    y = objDraw.CurrentY + H_9pt * 2 + Printer.TextHeight("江") '起始下移2个字高
    tmpPrintHeight = Printer.TextHeight("江")
    Set objDraw = Nothing
    lngPageTmp = lngEndPage
    If blnPrintNO Then
        Set objDraw = CheckNewPage(objOut, lngEndPage, y, tmpPrintHeight)
    Else
        lngTmpPageNo = 0
        Set objDraw = CheckNewPage(objOut, lngTmpPageNo, y, tmpPrintHeight)
    End If
    If objDraw Is Nothing Then Exit Function
    
    If lngPageTmp = lngEndPage And objDraw.CurrentY = lngTop Then
        '得到标题XY坐标
        x = lngLeft + (lngWidth - (lngLeft + lngRight) - Printer.TextWidth(strText)) / 2
        objDraw.CurrentX = x: objDraw.CurrentY = y
        objDraw.FontTransparent = True
        If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
            (IIf(blnPrintNO, lngEndPage, lngTmpPageNo) - mPageNumber >= mPrintBegingPage - 1 And IIf(blnPrintNO, lngEndPage, lngTmpPageNo) - _
            mPageNumber <= mPrintEndPage) Then
            objDraw.Print strText
        End If
        objDraw.CurrentY = y + Printer.TextHeight(strText)
        
        '打印病历科室,签名,日期
        objDraw.ForeColor = 0
        objDraw.Font.Name = "宋体"
        objDraw.Font.Size = 10.5
        objDraw.Font.Bold = False
        SetPrinterFont objDraw.Font, 10.5
        
        '判断是否新页
        y = objDraw.CurrentY + H_9pt * 2 + Printer.TextHeight("江") '起始下移2个字高
        tmpPrintHeight = Printer.TextHeight("江")
        Set objDraw = Nothing
        If blnPrintNO Then
            Set objDraw = CheckNewPage(objOut, lngEndPage, y, tmpPrintHeight)
        Else
            lngTmpPageNo = 0
            Set objDraw = CheckNewPage(objOut, lngTmpPageNo, y, tmpPrintHeight)
        End If
        If objDraw Is Nothing Then Exit Function
        
        If blnDemo = False Then
            strText = "科室:" & zlCommFun.Nvl(rsTmp!科室名称)
        Else
            strText = "适用科室:" & IIf(Trim(zlCommFun.Nvl(rsTmp!科室名称)) = "", "所有科室", zlCommFun.Nvl(rsTmp!科室名称))
        End If
        x = lngLeft
        objDraw.CurrentX = x: objDraw.CurrentY = y
        objDraw.FontTransparent = True
        If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
            (IIf(blnPrintNO, lngEndPage, lngTmpPageNo) - mPageNumber >= mPrintBegingPage - 1 And IIf(blnPrintNO, lngEndPage, lngTmpPageNo) - _
            mPageNumber <= mPrintEndPage) Then
            objDraw.Print strText
        End If
        mPageHeadDep = strText
        If blnDemo = False Then
            strText = "　姓名:" & zlCommFun.Nvl(rsTmp!姓名)
        Else
            strText = "　制定人:" & zlCommFun.Nvl(rsTmp!书写人)
        End If
        x = lngLeft + (lngWidth - lngLeft - lngRight) * (2 / 9)
        objDraw.CurrentX = x: objDraw.CurrentY = y
        objDraw.FontTransparent = True
        If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
            (IIf(blnPrintNO, lngEndPage, lngTmpPageNo) - mPageNumber >= mPrintBegingPage - 1 And IIf(blnPrintNO, lngEndPage, lngTmpPageNo) - _
            mPageNumber <= mPrintEndPage) Then
            objDraw.Print strText
        End If
        mPageHeadName = strText
        If blnDemo = False Then
            If blnOutPati Then
                strText = "　　门诊号:" & rsTmp!门诊号
            Else
                strText = "    住院号:" & rsTmp!住院号
            End If
        Else
            strText = "制定日期:" & IIf(IsNull(rsTmp!书写日期), "", Format(rsTmp!书写日期, "YYYY年MM月DD日"))
        End If
        x = lngLeft + (lngWidth - lngLeft - lngRight) * (4 / 9)
        objDraw.CurrentX = x: objDraw.CurrentY = y
        objDraw.FontTransparent = True
        If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
            (IIf(blnPrintNO, lngEndPage, lngTmpPageNo) - mPageNumber >= mPrintBegingPage - 1 And IIf(blnPrintNO, lngEndPage, lngTmpPageNo) - _
            mPageNumber <= mPrintEndPage) Then
            objDraw.Print strText
        End If
        mPageHeadNo = strText
        If blnDemo = False Then
            strText = "床号:" & zlCommFun.Nvl(rsTmp!当前床号)
        Else
            strText = "床号:"
        End If
        x = lngLeft + (lngWidth - lngLeft - lngRight) * (7 / 9)
        objDraw.CurrentX = x: objDraw.CurrentY = y
        objDraw.FontTransparent = True
        If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
            (IIf(blnPrintNO, lngEndPage, lngTmpPageNo) - mPageNumber >= mPrintBegingPage - 1 And IIf(blnPrintNO, lngEndPage, lngTmpPageNo) - _
            mPageNumber <= mPrintEndPage) Then
            objDraw.Print strText
        End If
        objDraw.CurrentY = y + Printer.TextHeight(strText)
        mPageBedNumber = strText
        
        y = objDraw.CurrentY + H_9pt / 5: x = lngLeft
        If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
            (IIf(blnPrintNO, lngEndPage, lngTmpPageNo) - mPageNumber >= mPrintBegingPage - 1 And IIf(blnPrintNO, lngEndPage, lngTmpPageNo) - _
            mPageNumber <= mPrintEndPage) Then
            objDraw.Line (lngLeft, y)-(lngWidth - lngRight, y), 0
        End If
    End If
    sngEndY = objDraw.CurrentY
    '该病历首页是否打印病人信息,如果是病历示范那么 lng病人ID 一定为0 此时就不打印病历信息
    If blnPatiInfo And lng病人ID > 0 Then
        '读出病人信息
        Set rsTmp = ReadPatiInfo(lng病人ID, lng主页ID)
        If Not rsTmp Is Nothing Then
            '分开打印门诊和住院病人信息
            objDraw.Font.Name = "宋体"
            objDraw.Font.Size = 10.5
            objDraw.Font.Bold = False
            SetPrinterFont objDraw.Font, 10.5
            '判断是否新页
            y = objDraw.CurrentY + H_9pt / 3 + Printer.TextHeight("江")  '这里行间距为1/3个普通字高
            tmpPrintHeight = Printer.TextHeight("江")
            Set objDraw = Nothing
            If blnPrintNO Then
                Set objDraw = CheckNewPage(objOut, lngEndPage, y, tmpPrintHeight)
            Else
                lngTmpPageNo = 0
                Set objDraw = CheckNewPage(objOut, lngTmpPageNo, y, tmpPrintHeight)
            End If
            If objDraw Is Nothing Then Exit Function
            
            '病人ID,标识号,(床号)
            x = lngLeft
            strText = "　病人ID:" & lng病人ID
            Call DrawCell(objDraw, strText, x, y, (lngWidth - lngLeft - lngRight) / 3, Printer.TextHeight("江"), IIf(blnPrintNO, lngEndPage, lngTmpPageNo), , , , , , objDraw.Font, "0000", 0)
            '住院病人与门诊信息不同
            If blnDemo = False Then
                If blnOutPati Then
                    x = lngLeft + (lngWidth - lngLeft - lngRight) * (2 / 3)
                    strText = "　门诊号:" & rsTmp!门诊号
                    Call DrawCell(objDraw, strText, x, y, (lngWidth - lngLeft - lngRight) / 3, Printer.TextHeight("江"), IIf(blnPrintNO, lngEndPage, lngTmpPageNo), , , , , , objDraw.Font, "0000", 0)
                Else
                    x = lngLeft + (lngWidth - lngLeft - lngRight) * (1 / 3)
                    strText = "　住院号:" & rsTmp!住院号
                    Call DrawCell(objDraw, strText, x, y, (lngWidth - lngLeft - lngRight) / 3, Printer.TextHeight("江"), IIf(blnPrintNO, lngEndPage, lngTmpPageNo), , , , , , objDraw.Font, "0000", 0)
                    
                    x = lngLeft + (lngWidth - lngLeft - lngRight) * (2 / 3)
                    strText = "　病床号:" & zlCommFun.Nvl(rsTmp!出院病床)
                    Call DrawCell(objDraw, strText, x, y, (lngWidth - lngLeft - lngRight) / 3, Printer.TextHeight("江"), IIf(blnPrintNO, lngEndPage, lngTmpPageNo), , , , , , objDraw.Font, "0000", 0)
                End If
            End If
            '判断是否新页
            y = objDraw.CurrentY + H_9pt / 3 + Printer.TextHeight("江")
            tmpPrintHeight = Printer.TextHeight("江")
            Set objDraw = Nothing
            If blnPrintNO Then
                Set objDraw = CheckNewPage(objOut, lngEndPage, y, tmpPrintHeight)
            Else
                lngTmpPageNo = 0
                Set objDraw = CheckNewPage(objOut, lngTmpPageNo, y, tmpPrintHeight)
            End If
            If objDraw Is Nothing Then Exit Function
            
            '姓名,性别,年龄
            x = lngLeft
            strText = "　　姓名:" & zlCommFun.Nvl(rsTmp!姓名)
            Call DrawCell(objDraw, strText, x, y, (lngWidth - lngLeft - lngRight) / 3, Printer.TextHeight("江"), IIf(blnPrintNO, lngEndPage, lngTmpPageNo), , , , , , objDraw.Font, "0000", 0)
            x = lngLeft + (lngWidth - lngLeft - lngRight) * (1 / 3)
            strText = "　　性别:" & zlCommFun.Nvl(rsTmp!性别)
            Call DrawCell(objDraw, strText, x, y, (lngWidth - lngLeft - lngRight) / 3, Printer.TextHeight("江"), IIf(blnPrintNO, lngEndPage, lngTmpPageNo), , , , , , objDraw.Font, "0000", 0)
            x = lngLeft + (lngWidth - lngLeft - lngRight) * (2 / 3)
            strText = "　　年龄:" & zlCommFun.Nvl(rsTmp!年龄)
            Call DrawCell(objDraw, strText, x, y, (lngWidth - lngLeft - lngRight) / 3, Printer.TextHeight("江"), IIf(blnPrintNO, lngEndPage, lngTmpPageNo), , , , , , objDraw.Font, "0000", 0)
            '判断是否新页
            y = objDraw.CurrentY + H_9pt / 3 + Printer.TextHeight("江")
            tmpPrintHeight = Printer.TextHeight("江")
            Set objDraw = Nothing
            If blnPrintNO Then
                Set objDraw = CheckNewPage(objOut, lngEndPage, y, tmpPrintHeight)
            Else
                lngTmpPageNo = 0
                Set objDraw = CheckNewPage(objOut, lngTmpPageNo, y, tmpPrintHeight)
            End If
            If objDraw Is Nothing Then Exit Function
            
            '民族,职业,婚姻状况
            x = lngLeft
            strText = "　　民族:" & zlCommFun.Nvl(rsTmp!民族)
            Call DrawCell(objDraw, strText, x, y, (lngWidth - lngLeft - lngRight) / 3, Printer.TextHeight("江"), IIf(blnPrintNO, lngEndPage, lngTmpPageNo), , , , , , objDraw.Font, "0000", 0)
            x = lngLeft + (lngWidth - lngLeft - lngRight) * (1 / 3)
            strText = "　　职业:" & zlCommFun.Nvl(rsTmp!职业)
            Call DrawCell(objDraw, strText, x, y, (lngWidth - lngLeft - lngRight) / 3, Printer.TextHeight("江"), IIf(blnPrintNO, lngEndPage, lngTmpPageNo), , , , , , objDraw.Font, "0000", 0)
            x = lngLeft + (lngWidth - lngLeft - lngRight) * (2 / 3)
            strText = "婚姻状况:" & zlCommFun.Nvl(rsTmp!婚姻状况)
            Call DrawCell(objDraw, strText, x, y, (lngWidth - lngLeft - lngRight) / 3, Printer.TextHeight("江"), IIf(blnPrintNO, lngEndPage, lngTmpPageNo), , , , , , objDraw.Font, "0000", 0)
            '判断是否新页
            y = objDraw.CurrentY + H_9pt / 3 + Printer.TextHeight("江")
            tmpPrintHeight = Printer.TextHeight("江")
            Set objDraw = Nothing
            If blnPrintNO Then
                Set objDraw = CheckNewPage(objOut, lngEndPage, y, tmpPrintHeight)
            Else
                lngTmpPageNo = 0
                Set objDraw = CheckNewPage(objOut, lngTmpPageNo, y, tmpPrintHeight)
            End If
            If objDraw Is Nothing Then Exit Function
            
            '工作单位
            x = lngLeft
            strText = "工作单位:" & zlCommFun.Nvl(rsTmp!工作单位)
            Call DrawCell(objDraw, strText, x, y, lngWidth - lngLeft - lngRight, Printer.TextHeight("江"), IIf(blnPrintNO, lngEndPage, lngTmpPageNo), , , , , , objDraw.Font, "0000", 0)
            '判断是否新页
            y = objDraw.CurrentY + H_9pt / 3 + Printer.TextHeight("江")
            sngEndY = y
            tmpPrintHeight = Printer.TextHeight("江")
            Set objDraw = Nothing
            If blnPrintNO Then
                Set objDraw = CheckNewPage(objOut, lngEndPage, y, tmpPrintHeight)
            Else
                lngTmpPageNo = 0
                Set objDraw = CheckNewPage(objOut, lngTmpPageNo, y, tmpPrintHeight)
            End If
            If objDraw Is Nothing Then Exit Function
            
            '家庭地址
            x = lngLeft
            strText = "家庭地址:" & zlCommFun.Nvl(rsTmp!家庭地址)
            Call DrawCell(objDraw, strText, x, y, lngWidth - lngLeft - lngRight, Printer.TextHeight("江"), IIf(blnPrintNO, lngEndPage, lngTmpPageNo), , , , , , objDraw.Font, "0000", 0)
            If blnOutPati = False And blnDemo = False Then
                '判断是否新页
                y = objDraw.CurrentY + H_9pt / 3 + Printer.TextHeight("江")
                tmpPrintHeight = Printer.TextHeight("江")
                Set objDraw = Nothing
                If blnPrintNO Then
                    Set objDraw = CheckNewPage(objOut, lngEndPage, y, tmpPrintHeight)
                Else
                    lngTmpPageNo = 0
                    Set objDraw = CheckNewPage(objOut, lngTmpPageNo, y, tmpPrintHeight)
                End If
                If objDraw Is Nothing Then Exit Function
                
                '入院科室,入院日期
                x = lngLeft
                strText = "入院科室:" & zlCommFun.Nvl(rsTmp!入院科室)
                Call DrawCell(objDraw, strText, x, y, (lngWidth - lngLeft - lngRight) / 3, Printer.TextHeight("江"), IIf(blnPrintNO, lngEndPage, lngTmpPageNo), , , , , , objDraw.Font, "0000", 0)
                x = lngLeft + (lngWidth - lngLeft - lngRight) * (2 / 3)
                strText = "入院日期:" & IIf(IsNull(rsTmp!入院日期), "", Format(rsTmp!入院日期, "yyyy年MM月dd日"))
                Call DrawCell(objDraw, strText, x, y, (lngWidth - lngLeft - lngRight) / 3, Printer.TextHeight("江"), IIf(blnPrintNO, lngEndPage, lngTmpPageNo), , , , , , objDraw.Font, "0000", 0)
                '判断是否新页
                y = objDraw.CurrentY + H_9pt / 3 + Printer.TextHeight("江")
                tmpPrintHeight = Printer.TextHeight("江")
                Set objDraw = Nothing
                If blnPrintNO Then
                    Set objDraw = CheckNewPage(objOut, lngEndPage, y, tmpPrintHeight)
                Else
                    lngTmpPageNo = 0
                    Set objDraw = CheckNewPage(objOut, lngTmpPageNo, y, tmpPrintHeight)
                End If
                If objDraw Is Nothing Then Exit Function
                
                '出院科室,出院日期
                x = lngLeft
                strText = "出院科室:" & zlCommFun.Nvl(rsTmp!出院科室)
                Call DrawCell(objDraw, strText, x, y, (lngWidth - lngLeft - lngRight) / 3, Printer.TextHeight("江"), IIf(blnPrintNO, lngEndPage, lngTmpPageNo), , , , , , objDraw.Font, "0000", 0)
                x = lngLeft + (lngWidth - lngLeft - lngRight) * (2 / 3)
                strText = "出院日期:" & IIf(IsNull(rsTmp!出院日期), "", Format(rsTmp!出院日期, "yyyy年MM月dd日"))
                Call DrawCell(objDraw, strText, x, y, (lngWidth - lngLeft - lngRight) / 3, Printer.TextHeight("江"), IIf(blnPrintNO, lngEndPage, lngTmpPageNo), , , , , , objDraw.Font, "0000", 0)
            End If
            y = objDraw.CurrentY + Printer.TextHeight("江") / 5: x = lngLeft
            If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
            (IIf(blnPrintNO, lngEndPage, lngTmpPageNo) - mPageNumber >= mPrintBegingPage - 1 And IIf(blnPrintNO, lngEndPage, lngTmpPageNo) - _
            mPageNumber <= mPrintEndPage) Then
                objDraw.Line (lngLeft, y)-(lngWidth - lngRight, y), 0
            End If
            objDraw.CurrentY = y
            sngEndY = y
        End If
        '完成标题打印及病人信息的打印
    End If
    '开始准备病历的打印
    If blnDemo = False Then
        If lng病历记录ID > 0 Then
            strSQL = "select * from 病人病历内容 where 病历记录id in (select id from 病人病历记录 where 病历种类 not in (-1,-2) and id =" & lng病历记录ID & ") ORDER BY 排列序号"
        Else
            strSQL = "select * from 病人病历内容 where 病历修订id=" & -1 * lng病历记录ID & " ORDER BY 排列序号"
        End If
    Else
        strSQL = "select * from 病人病历内容 where 病历示范ID=" & lng病历记录ID & " order by 排列序号"
    End If
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "病历打印")
    With rsTmp
        '判断是否有多个签名
        .Filter = "元素类型=-1": blnMultiSign = (.RecordCount > 1): .Filter = ""
        
        If .RecordCount > 0 Then
            .MoveFirst
            For i = 0 To .RecordCount - 1
                '读出病历元素编码，以便确定是哪个病历元素
                strSQL = " SELECT * FROM 病历元素目录 where 类型=" & !元素类型 & " and 编码='" & zlCommFun.Nvl(!元素编码) & "'"
                Call zlDatabase.OpenRecordset(rsNewTmp2, strSQL, "病历打印")
                If rsNewTmp2.RecordCount > 0 Then
                    str元素编码 = zlCommFun.Nvl(rsNewTmp2!编码) & "_" & zlCommFun.Nvl(rsNewTmp2!名称)
                Else
                    str元素编码 = ""
                End If
                '首先处理标题的打印(仅处理那些),对于特殊元素在它自己的过程中单独处理
                objDraw.ForeColor = zlCommFun.Nvl(!标题颜色, 0)
                strTitleFontName = zlCommFun.Nvl(!标题字体)
                If Trim(strTitleFontName) = "" Then strTitleFontName = "宋体"
                strTitleFontSize = "10.5": strTitleFontBold = "1": strTitleFontItalic = "0"
                For j = 1 To UBound(Split(strTitleFontName, ","))
                    Select Case j
                    Case 1
                        strTitleFontSize = Val(Split(strTitleFontName, ",")(j))
                    Case 2
                        strTitleFontBold = CLng(Split(strTitleFontName, ",")(j))
                    Case 3
                        strTitleFontItalic = CLng(Split(strTitleFontName, ",")(j))
                    End Select
                Next
                strTitleFontName = Split(strTitleFontName, ",")(0)
                '得到位置临时放在变量里保存
                strTitleAlig = zlCommFun.Nvl(!标题位置, 1)
                If !标题显示 = 1 And (!元素类型 >= 0 Or !元素类型 = -5) Then
                    objDraw.Font.Name = strTitleFontName
                    objDraw.Font.Size = Format(strTitleFontSize)
                    objDraw.Font.Bold = IIf(strTitleFontBold = "1", True, False)
                    objDraw.Font.Italic = IIf(strTitleFontItalic = "1", True, False)
                    SetPrinterFont objDraw.Font, Format(strTitleFontSize)
                    '得到标题文本准备打印
                    strText = zlCommFun.Nvl(!标题文本)
                    strText = IIf(Trim(strText) <> "", strText & "：", strText)
                    '对紧跟的加冒号
                    '                    strText = IIf(CLng(strTitleAlig) = 1 And zlCommFun.NVL(!内容位置, 1) = 1 And zlCommFun.NVL(!嵌入方式, 2) = 1, strText & "：", strText)
                    '根据标题位置求出标题XY坐标
                    Select Case CLng(strTitleAlig)
                    Case 1  '左
                        x = lngLeft
                    Case 2  '中
                        x = lngLeft + (lngWidth - (lngLeft + lngRight) - Printer.TextWidth(strText)) / 2
                    Case 3  '右
                        x = lngWidth - lngRight - Printer.TextWidth(strText)
                    End Select
                    '判断是否新页
                    y = objDraw.CurrentY + H_9pt / 2 + Printer.TextHeight("江")
                    tmpPrintHeight = Printer.TextHeight("江")
                    Set objDraw = Nothing
                    If blnPrintNO Then
                        Set objDraw = CheckNewPage(objOut, lngEndPage, y, tmpPrintHeight)
                    Else
                        lngTmpPageNo = 0
                        Set objDraw = CheckNewPage(objOut, lngTmpPageNo, y, tmpPrintHeight)
                    End If
                    If objDraw Is Nothing Then Exit Function
                    
                    TmpY = y '记下这一位置,以便紧跟标题式打印
                    sngEndY = y
                    If CLng(strTitleAlig) = 1 Then
                        '只有左对齐才这样处理,而右对齐与居中就在此时读出它的字符最后再输出
                        Call DrawCell(objDraw, strText, x, y, lngWidth - lngLeft - lngRight, Printer.TextHeight("江"), IIf(blnPrintNO, lngEndPage, lngTmpPageNo), , , , , , objDraw.Font, "0000", CLng(strTitleAlig) - 1)
                        TmpX = x + Printer.TextWidth(strText) + Printer.TextWidth("江") / 2
                    Else
                        '标题其它对齐方式
                        objDraw.CurrentX = x: objDraw.CurrentY = y: objDraw.FontTransparent = True
                        If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
                        (IIf(blnPrintNO, lngEndPage, lngTmpPageNo) - mPageNumber >= mPrintBegingPage - 1 And IIf(blnPrintNO, lngEndPage, lngTmpPageNo) - _
                        mPageNumber <= mPrintEndPage) Then
                            objDraw.Print strText
                        End If
                        objDraw.CurrentY = y + Printer.TextHeight(strText)
                        TmpX = lngLeft
                        objDraw.CurrentX = TmpX
                    End If
                    sngEndY = objDraw.CurrentY
                End If
                Select Case True
                    '文本段、可转文本的附加表、所见单、不为检验图象报告的专用纸
                Case !元素类型 = 0 Or !元素类型 = -5 _
                    Or (!元素类型 = 1 And IIf(IsNull(!文本转储), 0, !文本转储) = 1) _
                    Or (!元素类型 = 2 And IIf(IsNull(!文本转储), 0, !文本转储) = 1) _
                    Or (!元素类型 = 4 And Trim(zlCommFun.Nvl(!标题文本)) <> "检查图象报告")
                    '实现方法:直接读出字符串,再逐一读出每个字符来取行字符串,
                    '当刚好够一行的长度时就开始打印这一行,再重新取字符,如此
                    '同时在打印前还判断是不是开始新的一页
                    '内容字符控制
                    objDraw.ForeColor = zlCommFun.Nvl(!内容颜色, 0)
                    strFontName = zlCommFun.Nvl(!内容字体)
                    strFontSize = "9"
                    strFontBold = "0"
                    strFontItalic = "0"
                    For j = 1 To UBound(Split(strFontName, ","))
                        Select Case j
                        Case 1
                            strFontSize = Val(Split(strFontName, ",")(j))
                        Case 2
                            strFontBold = CLng(Split(strFontName, ",")(j))
                        Case 3
                            strFontItalic = CLng(Split(strFontName, ",")(j))
                        End Select
                    Next
                    If Trim(strFontName) = "" Then strFontName = "宋体"
                    strFontName = Split(strFontName, ",")(0)
                    objDraw.Font.Name = strFontName
                    objDraw.Font.Size = Format(strFontSize)
                    objDraw.Font.Bold = IIf(strFontBold = "1", True, False)
                    objDraw.Font.Italic = IIf(strFontItalic = "1", True, False)
                    SetPrinterFont objDraw.Font, Format(strFontSize)
                    '求出病历文本段内容并将其保存在文本变量里
                    strText = ""
                    strSQL = "select * from 病人病历文本段 where 病历ID=" & !ID & " order by 行号"
                    Call zlDatabase.OpenRecordset(rsNewTmp1, strSQL, "病历打印")
                    If rsNewTmp1.RecordCount > 0 Then
                        For j = 0 To rsNewTmp1.RecordCount - 1
                            strText = strText & zlCommFun.Nvl(rsNewTmp1!内容)
                            rsNewTmp1.MoveNext
                        Next
                    End If
                    '对于专用纸只打印处理转储过的文本
                    If !元素类型 = 4 Then
                        TmpX = lngLeft
                        objDraw.CurrentY = sngEndY
                        y = objDraw.CurrentY + H_9pt / 2  '下移一个小五字体高
                        objDraw.CurrentY = y
                        sngEndY = y
                    Else
                        '根据嵌入方式分别处理
                        '要紧跟标题，那么标题必须是左对齐的，并标题还应是要显示的，并且还至少能显示了一个字符的宽度，并且内容位置必须是对齐的
                        If CLng(strTitleAlig) = 1 And zlCommFun.Nvl(!内容位置, 1) = 1 And zlCommFun.Nvl(!嵌入方式, 2) = 1 And !标题显示 = 1 And Printer.TextWidth(zlCommFun.Nvl(!标题文本)) < lngWidth - (lngLeft + lngRight) - Printer.TextWidth("江") Then
                            '1 紧跟标题
                            '只有显示标题并且,标题文本的长度小于打印区域的宽度减去一个字的宽度时才可以紧跟标题打印
                            objDraw.CurrentY = TmpY
                            If strText = "" Then '如果该文本段没有内容就下移一行
                                y = objDraw.CurrentY + H_9pt  '下移一个小五字体高
                                objDraw.CurrentY = y
                                sngEndY = y
                            End If
                        Else
                            '2 新起一行, 3 文本环绕
                            '目前其它方式都处理为新起一行来打印
                            'todo:以后完善文本环绕的方式
                            objDraw.CurrentY = sngEndY
                            TmpX = lngLeft
                            y = objDraw.CurrentY + H_9pt / 2 '下移一个小五字体高
                            objDraw.CurrentY = y
                            sngEndY = y
                        End If
                    End If
                    blOnePrintText = True
                    Do While strText <> ""
                        '判断是否新页
                        If blOnePrintText = True Then
                            y = objDraw.CurrentY + Printer.TextHeight("江") + 30
                        Else
                            y = objDraw.CurrentY + H_9pt / 2 + Printer.TextHeight("江")
                        End If
                        blOnePrintText = False
                        tmpPrintHeight = Printer.TextHeight("江")
                        
                        If (Printer.TextWidth("江") * Len(strText) / 0.7) > (lngWidth - (lngRight * 2)) Then
                            Set objDraw = Nothing
                            If blnPrintNO Then
                                Set objDraw = CheckNewPage(objOut, lngEndPage, y, tmpPrintHeight)
                            Else
                                lngTmpPageNo = 0
                                Set objDraw = CheckNewPage(objOut, lngTmpPageNo, y, tmpPrintHeight)
                            End If
                            If objDraw Is Nothing Then Exit Function
                            objDraw.CurrentY = y  '+ H_9pt
                        Else
                            objDraw.CurrentY = y - Printer.TextHeight("江")
                        End If
                        sngEndY = y
                        strText = PrintLineS(objDraw, strText, TmpX, lngRight, IIf(blnPrintNO, lngEndPage, lngTmpPageNo))
                        TmpX = lngLeft
                    Loop
                    sngEndY = objDraw.CurrentY
                    '不可转文本的附加表
                Case !元素类型 = 1 And IIf(IsNull(!文本转储), 0, !文本转储) = 0
                    '判断是否新页
                    y = objDraw.CurrentY + H_9pt / 2 + Printer.TextHeight("江")
                    tmpPrintHeight = Printer.TextHeight("江")
                    Set objDraw = Nothing
                    If blnPrintNO Then
                        Set objDraw = CheckNewPage(objOut, lngEndPage, y, tmpPrintHeight)
                    Else
                        lngTmpPageNo = 0
                        Set objDraw = CheckNewPage(objOut, lngTmpPageNo, y, tmpPrintHeight)
                    End If
                    If objDraw Is Nothing Then Exit Function
                    
                    objDraw.CurrentY = y  '+ H_9pt
                    objDraw.ForeColor = zlCommFun.Nvl(!内容颜色, 0)
                    strFontName = zlCommFun.Nvl(!内容字体)
                    strFontSize = "9"
                    strFontBold = "0"
                    strFontItalic = "0"
                    For j = 1 To UBound(Split(strFontName, ","))
                        Select Case j
                            Case 1
                                strFontSize = Val(Split(strFontName, ",")(j))
                            Case 2
                                strFontBold = CLng(Split(strFontName, ",")(j))
                            Case 3
                                strFontItalic = CLng(Split(strFontName, ",")(j))
                        End Select
                    Next
                    If Trim(strFontName) = "" Then strFontName = "宋体"
                    strFontName = Split(strFontName, ",")(0)
                    objDraw.Font.Name = strFontName
                    objDraw.Font.Size = Format(strFontSize)
                    objDraw.Font.Bold = IIf(strFontBold = "1", True, False)
                    objDraw.Font.Italic = IIf(strFontItalic = "1", True, False)
                    SetPrinterFont objDraw.Font, Format(strFontSize)
                    '从数据中读出表格数据并画图
                    Call GridDraw(objOut, objDraw, !ID, y, blnPrintNO, lngEndPage, zlCommFun.Nvl(!内容位置, 1) - 1)
                    '下移一个小五字体高
                    y = objDraw.CurrentY + H_9pt / 2
                    objDraw.CurrentY = y:      sngEndY = y
                    '不可转文本的所见单
                Case !元素类型 = 2 And IIf(IsNull(!文本转储), 0, !文本转储) = 0
                    '已经处理在指定宽度范围内的折行打印、对齐方式、是否新页等处理
                    '如果当前页能打下就在当前页打印
                    '如果当前页打不下就在下页打印,如果下页还打不下说明是纸张或纸张大小设置的问题就只能在下页打.
                    m = zlCommFun.Nvl(!内容颜色, 0)
                    strFontName = zlCommFun.Nvl(!内容字体)
                    strFontSize = "9"
                    strFontBold = "0"
                    strFontItalic = "0"
                    For j = 1 To UBound(Split(strFontName, ","))
                        Select Case j
                        Case 1
                            strFontSize = Val(Split(strFontName, ",")(j))
                        Case 2
                            strFontBold = CLng(Split(strFontName, ",")(j))
                        Case 3
                            strFontItalic = CLng(Split(strFontName, ",")(j))
                        End Select
                    Next
                    If Trim(strFontName) = "" Then strFontName = "宋体"
                    strFontName = Split(strFontName, ",")(0)
                    '得到所见单的高度与宽度
                    strSQL = "SELECT nvl(MAX(列+宽),0)+30 所见单宽 ,nvl(MAX(行+高),0)+30 所见单高 FROM 病人病历所见单 WHERE 病历id=" & !ID
                    Call zlDatabase.OpenRecordset(rsNewTmp1, strSQL, "病历打印")
                    With rsNewTmp1
                        Tmp_W = rsNewTmp1!所见单宽 + W_9pt * 2
                        Tmp_H = rsNewTmp1!所见单高 + H_9pt * 2
                    End With
                    '对齐位置
                    Select Case zlCommFun.Nvl(!内容位置, 1)
                        Case 1  '左
                            TmpX = lngLeft
                        Case 2  '中
                            TmpX = lngLeft + (lngWidth - (lngLeft + lngRight) - Tmp_W) / 2
                        Case 3  '右
                            TmpX = lngWidth - (lngRight + Tmp_W)
                    End Select
                    y = objDraw.CurrentY + H_9pt / 2 '下移一个小五字体高
                    y = y + Tmp_H '设置新的当前Y位置为当前Y位置加上所见单的高度,在后面将会有判断是不是要新页的
                    '判断是否新页
                    Set objDraw = Nothing
                    If blnPrintNO Then
                        Set objDraw = CheckNewPage(objOut, lngEndPage, y, Tmp_H)
                    Else
                        lngTmpPageNo = 0
                        Set objDraw = CheckNewPage(objOut, lngTmpPageNo, y, Tmp_H)
                    End If
                    If objDraw Is Nothing Then Exit Function
                    
                    sngEndY = y
                    objDraw.ForeColor = m       'm临时用来记录颜色
                    objDraw.Font.Name = strFontName
                    objDraw.Font.Size = Format(strFontSize)
                    objDraw.Font.Bold = IIf(strFontBold = "1", True, False)
                    objDraw.Font.Italic = IIf(strFontItalic = "1", True, False)
                    SetPrinterFont objDraw.Font, Format(strFontSize)
                    strSQL = "Select 标题,所见内容,计量单位,行,列,高,宽 From 病人病历所见单 where 病历ID=" & !ID & " order by 行,列"
                    Call zlDatabase.OpenRecordset(rsNewTmp1, strSQL, "病历打印")
                    With rsNewTmp1
                        Do While Not .EOF
                            x = TmpX + zlCommFun.Nvl(!列, 0)
                            y = sngEndY + zlCommFun.Nvl(!行, 0)
                            Call DrawCell(objDraw, zlCommFun.Nvl(!标题) & " " & zlCommFun.Nvl(!所见内容) & " " & zlCommFun.Nvl(!计量单位), x, y, zlCommFun.Nvl(!宽, 0) + W_9pt, zlCommFun.Nvl(!高, 0), IIf(blnPrintNO, lngEndPage, lngTmpPageNo), , , , m, , objDraw.Font, "0000", 0, 1, True)
                            .MoveNext
                        Loop
                    End With
                    y = objDraw.CurrentY + H_9pt / 2 '下移一个小五字体高
                    objDraw.CurrentY = y:      sngEndY = y
                    '为检验图象报告的专用纸
                Case !元素类型 = 4 And Trim(zlCommFun.Nvl(!标题文本)) = "检查图象报告"
                    objDraw.Font.Name = "宋体": objDraw.Font.Size = 9: objDraw.Font.Bold = False: objDraw.Font.Italic = False
                    SetPrinterFont objDraw.Font, 9
                    TmpX = lngLeft
                    objDraw.CurrentY = sngEndY: objDraw.CurrentX = TmpX
                    y = objDraw.CurrentY + H_9pt / 2 + Printer.TextHeight("江") '下移半个小五字体高
                    tmpPrintHeight = Printer.TextHeight("江")
                    Set objDraw = Nothing
                    If blnPrintNO Then
                        Set objDraw = CheckNewPage(objOut, lngEndPage, y, tmpPrintHeight)
                    Else
                        lngTmpPageNo = 0
                        Set objDraw = CheckNewPage(objOut, lngTmpPageNo, y, tmpPrintHeight)
                    End If
                    If objDraw Is Nothing Then Exit Function
                    
                    objDraw.FontTransparent = True
                    If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
                    (IIf(blnPrintNO, lngEndPage, lngTmpPageNo) - mPageNumber >= mPrintBegingPage - 1 And IIf(blnPrintNO, lngEndPage, lngTmpPageNo) - _
                    mPageNumber <= mPrintEndPage) Then
                        objDraw.Print "    （另附检查图象）"
                    End If
                    objDraw.CurrentY = y + Printer.TextHeight("    （另附检查图象）")
                    y = objDraw.CurrentY + H_9pt    '下移一个小五字体高
                    objDraw.CurrentY = y:   sngEndY = y
                    '标记图
                Case !元素类型 = 3
                    Set ObjStdPic = GetMap(!ID, frmFlash.picTmp)
                    If Not (ObjStdPic Is Nothing Or ObjStdPic = 0) Then
                        TmpX = lngLeft
                        objDraw.CurrentY = sngEndY
                        objDraw.CurrentX = TmpX
                        '得到这张纸的最大高
                        m = lngHeight - lngTop - lngBottom - Screen.TwipsPerPixelY * 6
                        '得到图片的高
                        lngStdPicHeight = objDraw.ScaleY(ObjStdPic.Height, vbHimetric, objDraw.ScaleMode)
                        '得到图片的宽
                        lngStdPicWidth = objDraw.ScaleX(ObjStdPic.Width, vbHimetric, objDraw.ScaleMode)
                        '得到宽与高的比
                        dblPic比例 = ObjStdPic.Width / ObjStdPic.Height
                        '求出最大图片高
                        If lngStdPicHeight > m Then
                            lngStdPicHeight = m
                            '再得到宽
                            lngStdPicWidth = lngStdPicHeight * dblPic比例
                        End If
                        If lngStdPicWidth > lngWidth - lngLeft - lngRight - Screen.TwipsPerPixelX * 3 Then
                            lngStdPicWidth = lngWidth - lngLeft - lngRight - Screen.TwipsPerPixelX * 3
                            lngStdPicHeight = lngStdPicWidth / dblPic比例
                        End If
                        '重新确定X坐标
                        If lngStdPicWidth < lngWidth - (lngLeft + lngRight + Screen.TwipsPerPixelX * 2) Then
                            TmpX = lngLeft + (lngWidth - (lngLeft + lngRight) - lngStdPicWidth) / 2 - Screen.TwipsPerPixelX * 2
                        Else
                            TmpX = lngLeft
                        End If
                        
                        y = objDraw.CurrentY + lngStdPicHeight + H_9pt / 2  '下移半个小五字体高
                        Set objDraw = Nothing
                        If blnPrintNO Then
                            Set objDraw = CheckNewPage(objOut, lngEndPage, y, lngStdPicHeight)
                        Else
                            lngTmpPageNo = 0
                            Set objDraw = CheckNewPage(objOut, lngTmpPageNo, y, lngStdPicHeight)
                        End If
                        If objDraw Is Nothing Then Exit Function
                        
                        objDraw.CurrentY = y + Screen.TwipsPerPixelY * 2
                        y = objDraw.CurrentY
                        '开始打印图片
                        If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
                        (IIf(blnPrintNO, lngEndPage, lngTmpPageNo) - mPageNumber >= mPrintBegingPage - 1 And IIf(blnPrintNO, lngEndPage, lngTmpPageNo) - _
                        mPageNumber <= mPrintEndPage) Then
                            objDraw.PaintPicture ObjStdPic, TmpX, y, lngStdPicWidth, lngStdPicHeight, 0, 0, objDraw.ScaleX(ObjStdPic.Width, vbHimetric, objDraw.ScaleMode), objDraw.ScaleY(ObjStdPic.Height, vbHimetric, objDraw.ScaleMode)  'lngStdPicWidth, lngStdPicHeight
                        End If
                        objDraw.CurrentY = y + lngStdPicHeight
                        y = objDraw.CurrentY
                        sngEndY = y
                    End If
                    '书写者签名、当前日期、当前时间、标题
                Case !元素类型 = -1 Or !元素类型 = -2 Or !元素类型 = -3 Or !元素类型 = -4
                    '对齐方式为2和3的标题在这里打印
'                    If zlCommFun.Nvl(!标题位置, 1) <> 1 Then
                        objDraw.Font.Name = strTitleFontName
                        objDraw.Font.Size = strTitleFontSize
                        objDraw.Font.Bold = IIf(strTitleFontBold = "1", True, False)
                        objDraw.Font.Italic = IIf(strTitleFontItalic = "1", True, False)
                        SetPrinterFont objDraw.Font, Int(strTitleFontSize)
                        If !元素类型 = -4 Then
                            strText = zlCommFun.Nvl(!标题文本)
                        Else
                            strText = IIf(!标题显示 = 1, zlCommFun.Nvl(!标题文本) & "：", "")
                        End If
'                    Else
'                        strText = ""
'                    End If
                    '求出病历文本段内容并将其保存在文本变量里，对于这些这些负类型的元素应该没有文本段，一般应直接跳过
                    strSQL = "select * from 病人病历文本段 where 病历ID=" & !ID & " order by 行号"
                    Call zlDatabase.OpenRecordset(rsNewTmp1, strSQL, "病历打印")
                    If rsNewTmp1.RecordCount > 0 Then
                        For j = 0 To rsNewTmp1.RecordCount - 1
                            If !元素类型 = -1 And Not blnMultiSign Then
                                strText = strText & GetAllName(lng病历记录ID)
                            Else
                                strText = strText & zlCommFun.Nvl(rsNewTmp1!内容)
                            End If
                            rsNewTmp1.MoveNext
                        Next
                    End If
                    If !元素类型 = -4 Or (Printer.TextWidth(zlCommFun.Nvl(!标题文本)) < lngWidth - (lngLeft + lngRight) - Printer.TextWidth("江")) Then
                        If !元素类型 = -4 Then
                            objDraw.CurrentY = sngEndY + H_9pt
                        End If
                        Select Case zlCommFun.Nvl(!标题位置, 1)
                        Case 1  '左
                            TmpX = lngLeft
                        Case 2  '中
                            TmpX = lngLeft + (lngWidth - (lngLeft + lngRight) - Printer.TextWidth(strText)) / 2
                        Case 3  '右
                            TmpX = lngWidth - lngRight - Printer.TextWidth(strText)
                        End Select
                    Else
                        TmpX = lngLeft
                        y = objDraw.CurrentY + H_9pt  '下移一个小五字体高
                        objDraw.CurrentY = y
                    End If
                    Do While strText <> ""
                        '判断是否新页
                        y = objDraw.CurrentY + H_9pt / 3 + Printer.TextHeight("江")
                        tmpPrintHeight = Printer.TextHeight("江")
                        Set objDraw = Nothing
                        If blnPrintNO Then
                            Set objDraw = CheckNewPage(objOut, lngEndPage, y, tmpPrintHeight)
                        Else
                            lngTmpPageNo = 0
                            Set objDraw = CheckNewPage(objOut, lngTmpPageNo, y, tmpPrintHeight)
                        End If
                        If objDraw Is Nothing Then Exit Function
                        
                        objDraw.CurrentY = y  '+ H_9pt
                        sngEndY = y
                        strText = PrintLineS(objDraw, strText, TmpX, lngRight, IIf(blnPrintNO, lngEndPage, lngTmpPageNo))
                        TmpX = lngLeft
                    Loop
                    sngEndY = objDraw.CurrentY
                End Select
                .MoveNext
            Next
        End If
    End With
    If blnPrintNO Then
        Set objDraw = NewPrintPage(objOut, lngEndPage, False)
    Else
        lngTmpPageNo = 0
        Set objDraw = NewPrintPage(objOut, lngTmpPageNo, False)
    End If

    If Not blnPrint And mPrintBegingPage <> 0 Then
        tmpPage = objOut.picPage.UBound + 1 - mPrintEndPage
        If tmpPage > 0 Then
            tmpPage = objOut.picPage.UBound + 1
            For i = tmpPage To mPrintEndPage + 1 Step -1
                Unload objOut.picPage(i - 1)
            Next
        End If
        Set objDraw = objOut.picPage(objOut.picPage.UBound)
    End If
    PrintOrPreviewCase = True
    Set objDraw = Nothing:     Set rsTmp = Nothing:     Set rsNewTmp1 = Nothing:     Set rsNewTmp2 = Nothing:     Set ObjStdPic = Nothing
    Exit Function
ErrHandle:
    If Err.Number = 480 Then
        MsgBox "没有足够的内存或虚拟内存进行预览！", vbInformation, gstrSysName
    Else
        If ErrCenter() = 1 Then
            Resume
        End If
    End If
    Call SaveErrLog
    Set objDraw = Nothing:     Set rsTmp = Nothing:     Set rsNewTmp1 = Nothing:     Set rsNewTmp2 = Nothing:     Set ObjStdPic = Nothing
End Function
Function GetAllName(CaseHistoryID As Long) As String
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能           读出所有的医生的修改记录
    '参数           病历记录ID
    '返回           所有的医生姓名 格式"上级/下级"
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim rsTmp As New ADODB.Recordset
    Dim strTmp As String
    
    
    If CaseHistoryID > 0 Then
        '所有
        strTmp = " 病历记录ID = " & CaseHistoryID & " Order By  版本序号"
    Else
        gstrSql = "select * from 病人病历修订记录 where id = " & Abs(CaseHistoryID)
        zlDatabase.OpenRecordset rsTmp, gstrSql, "医生签名"
        '当前选择的
        strTmp = " 病历记录ID = " & rsTmp("病历记录ID") & "And 版本序号 <= " & rsTmp("版本序号") & " Order By  版本序号"
    End If
    
    On Error GoTo errH
    
    gstrSql = "select * from 病人病历修订记录 where " & strTmp
    
    zlDatabase.OpenRecordset rsTmp, gstrSql, "医生签名"
    Do Until rsTmp.EOF
        If GetAllName = "" Then
            GetAllName = rsTmp("书写人")
        Else
            GetAllName = rsTmp("书写人") & "/" & GetAllName
        End If
        rsTmp.MoveNext
    Loop
    rsTmp.Close
    
    If CaseHistoryID > 0 Then
        gstrSql = "select * from 病人病历记录 where id = " & CaseHistoryID
        zlDatabase.OpenRecordset rsTmp, gstrSql, "医生签名"
        If rsTmp.EOF <> True Then
            If GetAllName <> "" Then
                GetAllName = rsTmp("审阅人") & "/" & GetAllName
            Else
                GetAllName = rsTmp("审阅人")
            End If
        End If
'    Else
'        gstrSql = "select * from 病人病历记录 where 病历修订ID = " & Abs(CaseHistoryID)
    End If
    
    Set rsTmp = Nothing
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
  
End Function

Private Sub SetPrinterFont(ByVal DevFont As StdFont, intFontSize As Integer)
    Printer.Font.Name = DevFont.Name
'    Printer.Font.Size = DevFont.Size
    Printer.Font.Size = intFontSize
    Printer.Font.Bold = DevFont.Bold
    Printer.Font.Underline = DevFont.Underline
    Printer.Font.Italic = DevFont.Italic
    Printer.Font.Strikethrough = DevFont.Strikethrough
    Printer.Font.Weight = DevFont.Weight
End Sub
