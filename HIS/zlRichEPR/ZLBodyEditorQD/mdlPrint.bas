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

Private Const MAXROWS = 46
Private Const ROWHEIGHT = 35
Private Const OPDAYS = 10                           '手术后标记天数
Private Const HOUR_STEP_Twips = 205                 '用来决定第四小时之间的宽度 用于体温表

Private Const INTSTEPTwip = 90  '用来决定5分钟这间的宽度 用于麻醉单
Private Const STRING_WAY As String = "→"
Private msngScale As Single
Private mstrSQL As String
Private mrsTmp As ADODB.Recordset
Private mblnMoved As Boolean
Private mbln呼吸曲线 As Boolean
Private mbln婴儿体温单显示出院 As Boolean
Private mstrChar(2) As String                       '依次为口温,腋温,肛温
Private mstrBreath As String                        '呼吸
Private mstrPulse As String                         '脉搏
Private mint心率应用 As Integer
Private mstr心率符号 As String
Private mlngFirstWidth As Long
Private mintOpDays As Integer
Private mblnStopFlag As Boolean
Private mbyt脉搏() As Byte

Private Enum COLOR
    黑色 = 0
    深灰色 = &H404040
    灰色 = &HE0E0E0
    红色 = 200
End Enum

Private Type GRAPHPOINT
    X As Single
    Y As Single
    符号 As String
    颜色 As Long
    标志 As Byte
End Type

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

Private mBodyFlag As BODYFLAG

Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
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

Public gblnOK As Boolean


    
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
    
Private Const PS_SOLID = 0

Public gblnPrinted As Boolean           '是否打印了体温单
  
Public Sub DrawDcLine(hDC As Long, startpx As Long, startpy As Long, endpx As Long, endpy As Long)
        Dim old     As Long
        Dim p     As Long
        Dim a     As POINTAPI
          
        p = CreatePen(PS_SOLID, 3, vbRed)
        old = SelectObject(hDC, p)
        MoveToEx hDC, startpx, startpy, a
        LineTo hDC, endpx, endpy
        SelectObject hDC, old
        DeleteObject p
End Sub
  
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

Public Function DrawCell(Dev As Object, ByVal Data As Variant, ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long, _
                        Optional ByVal TW As Long, Optional ByVal TH As Long, Optional BorderColor As Long, _
                        Optional ForeColor As Long, Optional BackColor As Long = &HFFFFFF, Optional ByVal Font As StdFont, _
                        Optional Border As String = "1111", Optional HAlign As Byte, Optional VAlign As Byte = 1, Optional Warp As Boolean, _
                        Optional Ratio As Single = 1, Optional ByVal sngScale As Single = 1) As Boolean
                        
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
    Dim sglFontSize As Single
    
    On Error GoTo errH
    
    sglFontSize = Dev.Font.Size
    
    DrawCell = True
    
    '范围限定
    If TW > 0 Then
        If X > TW Then Exit Function
        If X + W > TW Then W = TW - X
    End If
    If TH > 0 Then
        If Y > TH Then Exit Function
        If Y + H > TH Then H = TH - Y
    End If
    
    If TypeName(Data) = "Integer" Then
        X = X * Ratio: Y = Y * Ratio: W = W * Ratio: H = H * Ratio
        If Val(Data) < 0 Then
            Dev.Line (X, Y)-(X + W - IIf(W > 0, Screen.TwipsPerPixelX * Ratio, 0), Y + H - IIf(H > 0, Screen.TwipsPerPixelY * Ratio, 0)), ForeColor, B '矩形
        Else
            Dev.Line (X, Y)-(X + W - IIf(W > 0, Screen.TwipsPerPixelX * Ratio, 0), Y + H - IIf(H > 0, Screen.TwipsPerPixelY * Ratio, 0)), ForeColor, BF '实心矩形(线条)
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
        Dev.Font.Size = Font.Size * sngScale
        Dev.Font.Bold = Font.Bold
        Dev.Font.Underline = Font.Underline
        Dev.Font.Italic = Font.Italic
        
        '因缩放后可能字体比例不对,判断时以原始大小为准
        If H >= Dev.TextHeight(Replace(Data, vbCrLf, "")) Then blnH = True          '高度是否够用(加回车的算一行高度)
        If W >= Dev.TextWidth(Data) Then blnW = True And InStr(Data, vbCrLf) = 0    '宽度是否够用(加回车的为不够用,以便拆行)
        '缩变
        LINE_W = 30 * Ratio '边线间隔宽度(输出时用,判断时不用)
        X = -Int(-X * Ratio): Y = -Int(-Y * Ratio)
        W = -Int(-W * Ratio): H = -Int(-H * Ratio)
        Dev.Font.Size = Font.Size * Ratio
        '背景填充
        Dev.Line (X, Y)-(X + W, Y + H), BackColor, BF
        Dev.ForeColor = ForeColor
        '输出文字(边框之内再隔一线)
        '超出高度范围则不输出
        If blnH Then
            If blnW Then
                Select Case HAlign
                Case 0
                    Dev.CurrentX = X + LINE_W
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
                Dev.FontTransparent = True
                Dev.Print Data
            Else
                If Not Warp Then
                    '不自动拆行时超宽部分不输出
                    
'                    '打不下时，自动缩小字体
'                    If Dev.TextWidth(Data) > W Then
'                        '根据总的宽度计算字体大小
'                        Dev.Font.Size = Dev.Font.Size * (1 - (Dev.TextWidth(Data) - W) / W - 0.05)
'                    End If
'                    Text = Data
                    
                    For i = 1 To Len(Data)
                        If Dev.TextWidth(Text & Mid(Data, i, 1)) > W Then Exit For
                        Text = Text & Mid(Data, i, 1)
                    Next

                    Select Case HAlign
                    Case 0
                        Dev.CurrentX = X + LINE_W
                    Case 1
                        Dev.CurrentX = X + (W - Dev.TextWidth(Text)) / 2
                    Case 2
                        Dev.CurrentX = X + W - LINE_W - Dev.TextWidth(Text)
                    End Select
                    Select Case VAlign
                    Case 0
                        Dev.CurrentY = Y + LINE_W
                    Case 1
                        Dev.CurrentY = Y + (H - Dev.TextHeight(Text)) / 2 + LINE_W / 2
                    Case 2
                        Dev.CurrentY = Y + H - LINE_W - Dev.TextHeight(Text)
                    End Select
                    Dev.FontTransparent = True
                    '输出截取部份
                    Dev.Print Text
                Else
                    '拆分文字成多行(在宽高范围内)
                    ReDim arrText(0) '在此,第一行不可能超高
                    Data = Replace(Data, vbCrLf, vbCr)
                    Data = Replace(Data, vbLf, vbCr)
                    For i = 1 To Len(Data)
                        If Mid(Data, i, 1) = vbCr Then
                            '多行超高则退出,超高部份不输出
                            If Dev.TextHeight(Replace(Data, vbCr, "")) * (UBound(arrText) + 2) > H Then Exit For
                            ReDim Preserve arrText(UBound(arrText) + 1)
                        ElseIf Dev.TextWidth(arrText(UBound(arrText)) & Mid(Data, i, 1)) > W Then
                            '多行超高则退出,超高部份不输出
                            If Dev.TextHeight(Replace(Data, vbCr, "")) * (UBound(arrText) + 2) > H Then Exit For
                            ReDim Preserve arrText(UBound(arrText) + 1)
                        End If
                        '有可能一行一个字符宽度都不够
                        If Dev.TextWidth(arrText(UBound(arrText)) & Mid(Data, i, 1)) <= W And Mid(Data, i, 1) <> vbCr Then
                            arrText(UBound(arrText)) = arrText(UBound(arrText)) & Mid(Data, i, 1)
                        End If
                    Next
                    
                    '输出起始坐标
                    Select Case VAlign
                    Case 0
                        Dev.CurrentY = Y + LINE_W
                    Case 1
                        Dev.CurrentY = Y + (H - Dev.TextHeight(Replace(Data, vbCr, "")) * (UBound(arrText) + 1)) / 2 + LINE_W / 2
                    Case 2
                        Dev.CurrentY = Y + H - LINE_W - Dev.TextHeight(Replace(Data, vbCr, "")) * (UBound(arrText) + 1)
                    End Select
                    
                    '输出各行
                    For i = 0 To UBound(arrText)
                        Select Case HAlign
                        Case 0
                            Dev.CurrentX = X + LINE_W
                        Case 1
                            Dev.CurrentX = X + (W - Dev.TextWidth(arrText(i))) / 2
                        Case 2
                            Dev.CurrentX = X + W - LINE_W - Dev.TextWidth(arrText(i))
                        End Select
                        Dev.FontTransparent = True
                        Dev.Print arrText(i)
                    Next
                End If
            End If
        End If
    ElseIf Not Data Is Nothing Then
        LINE_W = 30 * Ratio '边线间隔宽度(输出时用,判断时不用)
        X = X * Ratio: Y = Y * Ratio: W = W * Ratio: H = H * Ratio
        
        '图形(边框之内)
        Dev.PaintPicture Data, X + 15, Y + 15, W - LINE_W, H - LINE_W
    End If
    If TypeName(Data) <> "Integer" Then
        '最后处理边框,上，下，左，右
        If Mid(Border, 1, 1) Then Dev.Line (X, Y)-(X + W, Y), BorderColor
        If Mid(Border, 2, 1) Then Dev.Line (X, Y + H)-(X + W, Y + H), BorderColor
        If Mid(Border, 3, 1) Then Dev.Line (X, Y)-(X, Y + H), BorderColor
        If Mid(Border, 4, 1) Then Dev.Line (X + W, Y)-(X + W, Y + H), BorderColor
    End If
    
    Dev.Font.Size = sglFontSize
    
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
    If Not ExistsPrinter Then Exit Function
    
    '初始化打印参数
    
    strPrinter = Trim(zlDatabase.GetPara("体温单打印机", glngSys, 1255, Printer.DeviceName))
    intPage = Val(zlDatabase.GetPara("体温单纸张", glngSys, 1255, Printer.PaperSize))
    lngWidth = Val(zlDatabase.GetPara("体温单宽度", glngSys, 1255, Printer.Width))
    lngHeight = Val(zlDatabase.GetPara("体温单高度", glngSys, 1255, Printer.Height))
    intOrient = Val(zlDatabase.GetPara("体温单纸向", glngSys, 1255, Printer.Orientation))
    intBin = Val(zlDatabase.GetPara("体温单进纸", glngSys, 1255, Printer.PaperBin))
    
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
        If AddCustomPaper(objParent.hWnd, lngWidth / 56.7, lngHeight / 56.7) = FORM_NOT_SELECTED Then Exit Function
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
    
    lngLeft = Val(zlDatabase.GetPara("体温单左边距", glngSys, 1255, OFFSET_LEFT)) * 56.7
    lngRight = Val(zlDatabase.GetPara("体温单右边距", glngSys, 1255, OFFSET_RIGHT)) * 56.7
    lngTop = Val(zlDatabase.GetPara("体温单上边距", glngSys, 1255, OFFSET_TOP)) * 56.7
    lngBottom = Val(zlDatabase.GetPara("体温单下边距", glngSys, 1255, OFFSET_BOTTOM)) * 56.7
    
    
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

Public Function DrawPatiInfo(objDraw As Object, X As Long, Y As Long, strInfo As String, ByVal lngPaperWidth As Long) As Boolean
    '功能:打印顶行病人信息
    '参数:strInfo=姓名、住院号、科室、病区、床号、日期
    
    '顺序:
    '   1、DrawPatiInfo
    '   2、DrawBodyInfo
    '   3、DrawBodyScale
    '   4、DrawBodyTopScale
    '   5、DrawBodyPaper
    '   6、DrawBodyGraph
    '   7、DrawBodyRecordItem
    '   8、DrawBodyTips
    '   9、DrawBodyPageNO
    '画图顺序号:=1/9   这是画所有图形的第一步
    Dim strTmp As String
    Dim W_9pt As Long, H_9pt As Long
    Dim lngTmpX As Long, lngTmpY As Long
    
    If Trim(strInfo) = "" Then Exit Function
    If InStr(1, strInfo, "'") < 1 Then Exit Function
    If UBound(Split(strInfo, "'")) < 6 Then Exit Function
    On Error GoTo errHand
    
    lngTmpX = X
    lngTmpY = Y
    
    W_9pt = objDraw.TextWidth("字")
    H_9pt = objDraw.TextHeight("字")
    
    Call DrawText(objDraw, X, Y, "姓名:", 0)
    Call DrawText(objDraw, X + objDraw.TextWidth("姓名:"), Y, Split(strInfo, "'")(0), 16711680)
    X = X + objDraw.TextWidth("姓名:" & Split(strInfo, "'")(0)) + W_9pt / 3
    
    Call DrawText(objDraw, X, Y, "性别:", 0)
    Call DrawText(objDraw, X + objDraw.TextWidth("性别:"), Y, Split(strInfo, "'")(5), 16711680)
    X = X + objDraw.TextWidth("性别:" & Split(strInfo, "'")(5)) + W_9pt / 3
    
    Call DrawText(objDraw, X, Y, "年龄:", 0)
    Call DrawText(objDraw, X + objDraw.TextWidth("年龄:"), Y, Split(strInfo, "'")(6), 16711680)
    X = X + objDraw.TextWidth("年龄:" & Split(strInfo, "'")(6)) + W_9pt / 3
    
    Call DrawText(objDraw, X, Y, "病房:", 0)
    Call DrawText(objDraw, X + objDraw.TextWidth("病房:"), Y, Split(strInfo, "'")(4), 16711680)
    X = X + objDraw.TextWidth("病房:" & Split(strInfo, "'")(4)) + W_9pt / 3
    
    Call DrawText(objDraw, X, Y, "入院日期:", 0)
    Call DrawText(objDraw, X + objDraw.TextWidth("入院日期:"), Y, Split(strInfo, "'")(3), 16711680)
    X = X + objDraw.TextWidth("入院日期:" & Split(strInfo, "'")(3)) + W_9pt / 3
    
    Call DrawText(objDraw, X, Y, "住院号:", 0)
    Call DrawText(objDraw, X + objDraw.TextWidth("住院号:"), Y, Split(strInfo, "'")(1), 16711680)
    X = X + objDraw.TextWidth("住院号:" & Split(strInfo, "'")(1)) + W_9pt / 3

    X = lngTmpX
    Y = lngTmpY + H_9pt + H_9pt \ 2
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub DrawBodyRecordItem(objDraw As Object, X As Long, Y As Long, ByVal colFirstWidth As Long, strValue() As String, ByVal rsArrRecordItemInfo As ADODB.Recordset, _
    ByVal strDays As String, ByVal strOPS As String, Optional sngScale As Single = 1, Optional ByVal intGridRows As Integer = 0)
    '******************************************************************************************************************
    '功能:画底部的录入项目
    '参数:intRows=要画的行数
    '     colFirstWidth=首列的宽度
    '     strValue()=字符列表数组来表示显示的值  说明：strValue()不能为Empty
    
    '顺序:
    '画图顺序号:=7/9   这是画所有图形的第七步
    '******************************************************************************************************************
    
    Dim intRow As Integer, intCol As Integer '循环之用
    Dim H_9pt As Long, W_9pt As Long
    Dim lngTmpX As Long, lngTmpY As Long
    Dim strTmp As String
    Dim intAdd As Integer
    Dim blnExistData As Boolean
    Dim blnGrade As Boolean
    Dim str1 As String
    Dim str2 As String
    Dim str3 As String
    Dim lngColor As Long
    Dim intFactRow As Integer
    Dim intCount As Integer
    Dim bln呼吸 As Boolean
    Dim int呼吸表格输出格式 As Integer
    Dim int呼吸位置 As Integer
    On Error GoTo ErrHead
    '固定从第三行开始,连续五行输出
    
    If UBound(strValue) < 0 Then Exit Sub
    If IsEmpty(strValue) = True Then Exit Sub
    int呼吸位置 = 1
    int呼吸表格输出格式 = zlDatabase.GetPara("呼吸表格输出", glngSys, 1255, 0)
    
    objDraw.FontSize = 9 * sngScale
    H_9pt = objDraw.TextHeight("字")
    W_9pt = objDraw.TextWidth("字")
    objDraw.DrawStyle = 0
    lngTmpX = X
    lngTmpY = Y
    
    intAdd = 1
    If mbln呼吸曲线 Then intAdd = 0
                
    If InStr(1, strValue(0), ";") > 0 Then
        
        intFactRow = LBound(strValue) - 1
        int呼吸位置 = 2
        For intRow = LBound(strValue) To UBound(strValue)
            
            objDraw.FontSize = 9 * sngScale
            
            If intGridRows = 0 Or intGridRows > intFactRow + IIf(bln呼吸, -1, 0) Then

                If Split(strValue(intRow), ";")(0) = "呼吸" Then
                    bln呼吸 = True
                    intFactRow = intFactRow + 2
                    
                    For intCol = 0 To 42
                        If intCol = 0 Then
                        
                            Call DrawCell(objDraw, Split(strValue(intRow), ";")(intCol) & Split(strValue(intRow), ";")(intCol + 2), lngTmpX, Y + intFactRow * (H_9pt + H_9pt \ 2), colFirstWidth, 2 * (H_9pt + H_9pt \ 2), , , , , , objDraw.Font, , 1, 1)
                            Call DrawLine(objDraw, lngTmpX, Y + intFactRow * (H_9pt + H_9pt \ 2), lngTmpX, Y + intFactRow * (H_9pt + H_9pt \ 2) + 2 * (H_9pt + H_9pt \ 2), , , 2)
                            Call DrawLine(objDraw, lngTmpX, Y + intFactRow * (H_9pt + H_9pt \ 2), lngTmpX + colFirstWidth, Y + intFactRow * (H_9pt + H_9pt \ 2), , , 1)
                            Call DrawLine(objDraw, lngTmpX + colFirstWidth, Y + intFactRow * (H_9pt + H_9pt \ 2), lngTmpX + colFirstWidth, Y + intFactRow * (H_9pt + H_9pt \ 2) + 2 * (H_9pt + H_9pt \ 2), , , 1)
                            
                        Else
                            If intCol > UBound(Split(strValue(intRow), ";")) Then
                                strTmp = ""
                            Else
                                strTmp = Split(strValue(intRow), ";")(intCol + 2)
                            End If
                            
                            '打印呼吸值（间隔错开打印）
                            If int呼吸表格输出格式 = 0 Then
                                If intCol Mod 2 = 0 Then
                                    Call DrawCell(objDraw, strTmp, lngTmpX + colFirstWidth + HOUR_STEP_Twips * (intCol - 1), Y + intFactRow * (H_9pt + H_9pt \ 2), HOUR_STEP_Twips, (H_9pt + H_9pt \ 2), , , , COLOR.红色, , objDraw.Font, "1000")
                                Else
                                    Call DrawCell(objDraw, strTmp, lngTmpX + colFirstWidth + HOUR_STEP_Twips * (intCol - 1), Y + intFactRow * (H_9pt + H_9pt \ 2) + (H_9pt + H_9pt \ 2), HOUR_STEP_Twips, (H_9pt + H_9pt \ 2), , , , COLOR.红色, , objDraw.Font, "0000")
                                End If
                            Else
                                If int呼吸位置 = 2 Then
                                    Call DrawCell(objDraw, strTmp, lngTmpX + colFirstWidth + HOUR_STEP_Twips * (intCol - 1), Y + intFactRow * (H_9pt + H_9pt \ 2), HOUR_STEP_Twips, (H_9pt + H_9pt \ 2), , , , COLOR.红色, , objDraw.Font, "1000")
                                Else
                                    Call DrawCell(objDraw, strTmp, lngTmpX + colFirstWidth + HOUR_STEP_Twips * (intCol - 1), Y + intFactRow * (H_9pt + H_9pt \ 2) + (H_9pt + H_9pt \ 2), HOUR_STEP_Twips, (H_9pt + H_9pt \ 2), , , , COLOR.红色, , objDraw.Font, "0000")
                                End If
                                If strTmp <> "" Then
                                    int呼吸位置 = int呼吸位置 + 1
                                    If int呼吸位置 > 2 Then int呼吸位置 = 1
                                End If
                            End If
                        End If
                    Next
                    
                    For intCol = 0 To 42
                        If (intCol) Mod 6 = 0 Then
                            Call DrawLine(objDraw, lngTmpX + colFirstWidth + HOUR_STEP_Twips * intCol, Y + intFactRow * (H_9pt + H_9pt \ 2), lngTmpX + colFirstWidth + HOUR_STEP_Twips * intCol, Y + intFactRow * (H_9pt + H_9pt \ 2) + 2 * (H_9pt + H_9pt \ 2), IIf(intCol = 0, 0, vbRed), , 2)
                        Else
                            Call DrawLine(objDraw, lngTmpX + colFirstWidth + HOUR_STEP_Twips * intCol, Y + intFactRow * (H_9pt + H_9pt \ 2) - 15, lngTmpX + colFirstWidth + HOUR_STEP_Twips * intCol, Y + intFactRow * (H_9pt + H_9pt \ 2) + 2 * (H_9pt + H_9pt \ 2) + 15)
                            Call DrawLine(objDraw, lngTmpX + colFirstWidth + HOUR_STEP_Twips * (intCol + 1), Y + intFactRow * (H_9pt + H_9pt \ 2) - 15, lngTmpX + colFirstWidth + HOUR_STEP_Twips * (intCol + 1), Y + intFactRow * (H_9pt + H_9pt \ 2) + 2 * (H_9pt + H_9pt \ 2) + 15)
                            
                        End If
                    Next
                Else
                    blnExistData = False
                    rsArrRecordItemInfo.Filter = ""
                    rsArrRecordItemInfo.Filter = "序号=" & intRow
                    If rsArrRecordItemInfo.RecordCount > 0 Then
                        If rsArrRecordItemInfo("项目性质").Value = 2 Then
                            '是活动项目，要判断是否为空，如为空值，则不打印此行
                            
                            For intCol = 1 To 14
                                
                                If intCol <= UBound(Split(strValue(intRow), ";")) Then
                                    strTmp = Trim(Split(strValue(intRow), ";")(intCol + 2))
                                    If strTmp <> "" Then
                                        blnExistData = True
                                    End If
                                End If
                                
                            Next
                        Else
                            blnExistData = True
                        End If
                    Else
                        blnExistData = True
                    End If
                    
                    If blnExistData Then
                        
                        intFactRow = intFactRow + 1
                        
                        For intCol = 0 To 14
                            If intCol = 0 Then
                            
                                strTmp = Split(strValue(intRow), ";")(intCol) & Split(strValue(intRow), ";")(intCol + 2)
                                
                                Call DrawCell(objDraw, strTmp, lngTmpX + IIf(intFactRow >= 4 And intFactRow <= 8, 200, 0), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), colFirstWidth, (H_9pt + H_9pt \ 2), , , , , , objDraw.Font, , 1, 1)
                                Call DrawLine(objDraw, lngTmpX, Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), lngTmpX, Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2) + (H_9pt + H_9pt \ 2), , , 2)
                            Else
        
                                Select Case Split(strValue(intRow), ";")(0)
                                Case "血压"
                                    objDraw.FontSize = 7.8 * sngScale
                                Case Else
                                      
                                End Select
                                
                                blnGrade = False
                                
                                If intCol Mod 2 = 1 Then
                                    '一天的上午
                                    If (Split(strValue(intRow), ";")(intCol + 2) <> "" And Split(strValue(intRow), ";")(intCol + 3) <> "") Or Val(Split(strValue(intRow), ";")(1)) = 2 Then
                                        '上下午都有值
                                        
                                        If intCol > UBound(Split(strValue(intRow), ";")) Then
                                            strTmp = ""
                                        Else
                                            strTmp = Split(strValue(intRow), ";")(intCol + 2)
                                        End If
                                        
                                        If Split(strValue(intRow), ";")(0) = "大便次数" And strTmp <> "" Then
                                            '检查是否为分数，如是，则分别计算出整数、分子、分母部份的内容
                                            blnGrade = AnsyGrade(strTmp, str1, str2, str3)
                                        End If
                                
                                        lngColor = GridTextColor(Split(strValue(intRow), ";")(0), strTmp)
                                        Call DrawCell(objDraw, IIf(blnGrade, "", strTmp), lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), HOUR_STEP_Twips * 3, (H_9pt + H_9pt \ 2), , , , lngColor, , objDraw.Font, , 1)
                                        If blnGrade Then
                                            objDraw.FontSize = 7.5 * sngScale
                                            Call DrawGrade(objDraw, lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), HOUR_STEP_Twips * 3, (H_9pt + H_9pt \ 2), str1, str2, str3, 0)
                                            objDraw.FontSize = 9 * sngScale
                                        End If
                                        
                                        If (intCol + 1) Mod 2 = 0 Then
                                            Call DrawLine(objDraw, lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2) + (H_9pt + H_9pt \ 2), , , 2)
                                        End If
                                        
                                    Else
                                        '合并打印
                                        If intCol > UBound(Split(strValue(intRow), ";")) Then
                                            strTmp = ""
                                        Else
                                            If Split(strValue(intRow), ";")(intCol + 2) <> "" Then
                                                strTmp = Split(strValue(intRow), ";")(intCol + 2)
                                            ElseIf Split(strValue(intRow), ";")(intCol + 3) <> "" Then
                                                strTmp = Split(strValue(intRow), ";")(intCol + 3)
                                            Else
                                                strTmp = ""
                                            End If
                                        End If
                                        
                                        If Split(strValue(intRow), ";")(0) = "大便次数" And strTmp <> "" Then
                                            '检查是否为分数，如是，则分别计算出整数、分子、分母部份的内容
                                            blnGrade = AnsyGrade(strTmp, str1, str2, str3)
                                        End If
                                                                        
                                        lngColor = GridTextColor(Split(strValue(intRow), ";")(0), strTmp)
                                        Call DrawCell(objDraw, IIf(blnGrade, "", strTmp), lngTmpX + colFirstWidth + HOUR_STEP_Twips * 6 * Fix(intCol / 2), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), HOUR_STEP_Twips * 6, (H_9pt + H_9pt \ 2), , , , lngColor, , objDraw.Font, , 1)
                                        If blnGrade Then
                                            objDraw.FontSize = 7.5 * sngScale
                                            Call DrawGrade(objDraw, lngTmpX + colFirstWidth + HOUR_STEP_Twips * 6 * Fix(intCol / 2), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), HOUR_STEP_Twips * 6, (H_9pt + H_9pt \ 2), str1, str2, str3, 0)
                                            objDraw.FontSize = 9 * sngScale
                                        End If
                                        
                                        Call DrawLine(objDraw, lngTmpX + colFirstWidth + HOUR_STEP_Twips * 6 * Fix(intCol / 2), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2) + (H_9pt + H_9pt \ 2), IIf(intCol = 1, 0, vbRed), , 2)
                                        
                                        intCol = intCol + 1
                                    End If
                                Else
                                    '一天的下午
                                    If intCol > UBound(Split(strValue(intRow), ";")) Then
                                        strTmp = ""
                                    Else
                                        strTmp = Split(strValue(intRow), ";")(intCol + 2)
                                    End If
            
                                    If Split(strValue(intRow), ";")(0) = "大便次数" And strTmp <> "" Then
                                        '检查是否为分数，如是，则分别计算出整数、分子、分母部份的内容
                                        blnGrade = AnsyGrade(strTmp, str1, str2, str3)
                                    End If
        
                                    lngColor = GridTextColor(Split(strValue(intRow), ";")(0), strTmp)
                                    Call DrawCell(objDraw, IIf(blnGrade, "", strTmp), lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), HOUR_STEP_Twips * 3, (H_9pt + H_9pt \ 2), , , , lngColor, , objDraw.Font, , 1)
                                    If blnGrade Then
                                        objDraw.FontSize = 7.5 * sngScale
                                        Call DrawGrade(objDraw, lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), HOUR_STEP_Twips * 3, (H_9pt + H_9pt \ 2), str1, str2, str3, 0)
                                        objDraw.FontSize = 9 * sngScale
                                    End If
                                    
                                    If (intCol + 1) Mod 2 = 0 Then
                                        Call DrawLine(objDraw, lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2) + (H_9pt + H_9pt \ 2), , , 2)
                                    End If
                                End If
                            End If
                            
                            If intCol = 14 Then
                                Call DrawLine(objDraw, lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1) + HOUR_STEP_Twips * 3, Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1) + HOUR_STEP_Twips * 3, Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2) + (H_9pt + H_9pt \ 2), vbRed, , 2)
                            End If
                        Next
                    End If
                End If
                
    '            If intRow = UBound(strValue) Then
    '                Call DrawLine(objDraw, lngTmpX, y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2) + (H_9pt + H_9pt \ 2), lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1), y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2) + (H_9pt + H_9pt \ 2), , , 2)
    '            End If
    
            End If
        Next

'        '补空行
'        If intGridRows > 0 And intGridRows > intFactRow + IIf(bln呼吸, -1, 0) Then
'            intCount = intGridRows - intFactRow
'            For intRow = 1 To intCount
'                intFactRow = intFactRow + 1
'                Call DrawCell(objDraw, "", lngTmpX, Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), colFirstWidth, (H_9pt + H_9pt \ 2), , , , , , objDraw.Font, , 1, 1)
'                Call DrawLine(objDraw, lngTmpX, Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), lngTmpX, Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2) + (H_9pt + H_9pt \ 2), , , 2)
'
'                For intCol = 1 To 14
'                    Call DrawCell(objDraw, "", lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), HOUR_STEP_Twips * 3, (H_9pt + H_9pt \ 2), , , , lngColor, , objDraw.Font)
'                    If (intCol + 1) Mod 2 = 0 Then
'                        Call DrawLine(objDraw, lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2) + (H_9pt + H_9pt \ 2), IIf(intCol = 1, 0, vbRed), , 2)
'                    End If
'                    If intCol = 14 Then
'                        Call DrawLine(objDraw, lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1) + HOUR_STEP_Twips * 3, Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1) + HOUR_STEP_Twips * 3, Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2) + (H_9pt + H_9pt \ 2), vbRed, , 2)
'                    End If
'                Next
'            Next
'            intFactRow = intGridRows
'        End If
        
'        '手术
        objDraw.FontSize = 9 * sngScale     '血压修改了字体大小,此处还原
        intFactRow = intFactRow + 1
        Call DrawCell(objDraw, "手术后天数", lngTmpX, Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), colFirstWidth, (H_9pt + H_9pt \ 2), , , , , , objDraw.Font, , 1, 1)
        Call DrawLine(objDraw, lngTmpX, Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), lngTmpX, Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2) + (H_9pt + H_9pt \ 2), , , 2)
        For intCol = 1 To 14
            If (intCol + 1) Mod 2 = 0 Then
                Call DrawCell(objDraw, GetSplitStr(strOPS, CInt((intCol - 1) / 2)), lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), HOUR_STEP_Twips * 6, (H_9pt + H_9pt \ 2), , , , lngColor, , objDraw.Font, , 1)
                Call DrawLine(objDraw, lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2) + (H_9pt + H_9pt \ 2), IIf(intCol = 1, 0, vbRed), , 2)
            End If
            If intCol = 14 Then
                Call DrawLine(objDraw, lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1) + HOUR_STEP_Twips * 3, Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1) + HOUR_STEP_Twips * 3, Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2) + (H_9pt + H_9pt \ 2), vbRed, , 2)
            End If
        Next
        
        '住院天数
        Dim lngStartDay As Long
        intFactRow = intFactRow + 1
        lngStartDay = GetSplitStr(strDays, 0)
        Call DrawCell(objDraw, "住院天数", lngTmpX, Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), colFirstWidth, (H_9pt + H_9pt \ 2), , , , , , objDraw.Font, , 1, 1)
        Call DrawLine(objDraw, lngTmpX, Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), lngTmpX, Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2) + (H_9pt + H_9pt \ 2), , , 2)
        For intCol = 1 To 14
            If (intCol + 1) Mod 2 = 0 Then
                Call DrawCell(objDraw, CStr(lngStartDay + CInt((intCol - 1) / 2)), lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), HOUR_STEP_Twips * 6, (H_9pt + H_9pt \ 2), , , , lngColor, , objDraw.Font, , 1)
                Call DrawLine(objDraw, lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2) + (H_9pt + H_9pt \ 2), IIf(intCol = 1, 0, vbRed), , 2)
            End If
            If intCol = 14 Then
                Call DrawLine(objDraw, lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1) + HOUR_STEP_Twips * 3, Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1) + HOUR_STEP_Twips * 3, Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2) + (H_9pt + H_9pt \ 2), vbRed, , 2)
            End If
        Next
        
        Call DrawLine(objDraw, lngTmpX, Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2) + (H_9pt + H_9pt \ 2), lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * 14, Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2) + (H_9pt + H_9pt \ 2), , , 2)
         
        '固定从第三行开始,连续五行输出(由于呼吸占两行,实际从第四行开始),计算中间位置输出,所以从7行开始
        Call DrawRotateText(objDraw, lngTmpX + 30, Y + 7 * (H_9pt + H_9pt \ 2), "出")
        Call DrawRotateText(objDraw, lngTmpX + 30, Y + 7 * (H_9pt + H_9pt \ 2) + 200, "量")
        
        '为后续画图作准备
        X = lngTmpX + W_9pt * 2
        Y = lngTmpY + (intFactRow - LBound(strValue) + 1 + intAdd) * (H_9pt + H_9pt \ 2) + 30
    Else
        '为后续画图作准备
        X = lngTmpX + W_9pt * 2
        Y = lngTmpY + 30
        
    End If
    
    objDraw.FontSize = 9 * sngScale
    Exit Sub
ErrHead:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function GridTextColor(ByVal strItemName As String, strResult As String) As Long
    
    GridTextColor = 0
    
    Select Case strItemName
    Case ""
    Case Else

        If InStr(strItemName, "皮试") > 0 Then
            If InStr(strResult, "+") > 0 Then
                GridTextColor = COLOR.红色
            End If
        End If
        
    End Select

End Function

Private Function AnsyGrade(ByVal strText As String, ByRef str1 As String, ByRef str2 As String, ByRef str3 As String) As Boolean
    
    '功能：分析分数
    
    Dim intPos As Integer
    
    If strText = "" Then Exit Function
    If InStr(strText, "/") = 0 Then Exit Function
    intPos = InStr(strText, "+")
    If intPos > 0 Then
        If intPos = 1 Then Exit Function
        str1 = Mid(strText, 1, intPos - 1)
        strText = Mid(strText, intPos + 1)
    End If
    
    intPos = InStr(strText, "/")
    If intPos > 0 Then
        If intPos = 1 Then Exit Function
        If intPos = Len(strText) Then Exit Function
        
        str2 = Mid(strText, 1, intPos - 1)
        str3 = Mid(strText, intPos + 1)
    End If
    
    AnsyGrade = True
    
End Function

Private Function DrawGrade(objDraw As Object, ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long, ByVal Text1 As String, ByVal Text2 As String, ByVal Text3 As String, Optional ByVal ForeColor As Long = 0) As Boolean
    
    '功能：打印分数
    
    Dim W0 As Long
    Dim H0 As Long
    Dim X0 As Long
    Dim Y0 As Long
    Dim L0 As Long
    
    '计算分数占用的宽度
    W0 = objDraw.TextWidth(Text1)
    H0 = objDraw.TextHeight("A")
    
    If Len(Text2) > Len(Text3) Then
        L0 = objDraw.TextWidth(Text2)
    Else
        L0 = objDraw.TextWidth(Text3)
    End If
    
    W0 = W0 + L0
    
    If Text1 <> "" Then
        X0 = X + (W - W0) / 2
        Y0 = Y + (H - H0) / 2
        Call DrawText(objDraw, X0, Y0, Text1, 0)
    End If
    
    X0 = X + (W - W0) / 2 + objDraw.TextWidth(Text1) + (W0 - objDraw.TextWidth(Text2)) / 2
    Y0 = Y - 15
        
    Call DrawText(objDraw, X0, Y0, Text2, 0)
    
    Y0 = Y0 + H0
    X0 = X + (W - W0) / 2 + objDraw.TextWidth(Text1)
    Call DrawLine(objDraw, X0, Y0, X0 + L0 + objDraw.TextWidth("A"), Y0, , , 1)

    Y0 = Y0 + 0
    X0 = X + (W - W0) / 2 + objDraw.TextWidth(Text1) + (W0 - objDraw.TextWidth(Text3)) / 2
    Call DrawText(objDraw, X0, Y0, Text3, 0)
                                     
End Function

Private Sub DrawBodyTips(objDraw As Object, X As Long, Y As Long, ByVal colFirstWidth As Long, ByVal strTips As String, Optional sngScale As Single = 1)
    '功能：画出最底部的记录符说明
    '参数:strTips=说明字符串
    
    '顺序:
    '画图顺序号:=8/9   这是画所有图形的第八步
    Dim H_9pt As Long
    Dim lngTmpX As Long, lngTmpY As Long
    lngTmpX = X
    lngTmpY = Y
    H_9pt = objDraw.TextHeight("字")
    Call DrawCell(objDraw, strTips, X, Y, colFirstWidth + HOUR_STEP_Twips * 6 * 7, H_9pt + H_9pt \ 2, , , , , , objDraw.Font, "0000", 0)
    X = lngTmpX
    Y = lngTmpY + H_9pt * 2
End Sub

Private Sub DrawBodyPageFooter(objDraw As Object, X As Long, Y As Long, ByVal intPageNo As Integer, ByVal intBeginPage As Integer)
    '******************************************************************************************************************
    '功能：画出最底部说明
    '参数:intNOPage=页码
    '******************************************************************************************************************
    Dim blnWeek As Boolean
    Dim blnPageNo As Boolean
    Dim H_9pt As Long, W_9pt As Long
    Dim strNOPage As String
    Dim lngX As Long
        
    H_9pt = objDraw.TextHeight("字")
    W_9pt = objDraw.TextWidth("字")
    
    blnWeek = (Val(zlDatabase.GetPara("打印周数", glngSys, 1255, "0")) = 1)
    blnPageNo = (Val(zlDatabase.GetPara("打印页号", glngSys, 1255, "1")) = 1)
    
    '打印页码
    '------------------------------------------------------------------------------------------------------------------
    If intPageNo > -1 And blnPageNo Then
        intPageNo = intPageNo + intBeginPage - 1
        strNOPage = "第  -- " & CStr(intPageNo) & " --  页"
    End If
    
    If blnWeek Then
        If strNOPage = "" Then
            strNOPage = "第  -- " & CStr(intBeginPage) & " --  周"
        Else
            strNOPage = strNOPage & "(第 " & CStr(intBeginPage) & " 周)"
        End If
        
    End If
    
    Call DrawCell(objDraw, strNOPage, X + ((6 * W_9pt + HOUR_STEP_Twips * 6 * 7) - objDraw.TextWidth(strNOPage)) / 2, Y, W_9pt * 20, H_9pt + H_9pt \ 2, , , , , , objDraw.Font, "0000", 0, 0)
End Sub

Private Function DrawBodyInfo(objDraw As Object, X As Long, Y As Long, ByVal colFirstWidth As Long, _
    ByVal strDate As String, ByVal strPatiDay As String, ByVal strOPSFate As String, Optional sngScale As Long = 1, Optional ByVal strBeginDate As String) As Boolean
    '功能:画出当前页面的住院天数及日期等
    '参数:objDraw=输出对象
    '     colFirstWidth = 首列的宽度
    '     strDate=本页的住院日期字符列
    '     strPatiDay=住院天数字符列
    '     strOPSFate=手术后字符列
    
    '顺序:
    '画图顺序号:=2/9   这是画所有图形的第二步
    
    Dim H_9pt As Long
    Dim intDay As Integer
    Dim lngLeft As Long
    Dim lngTmpX As Long, lngTmpY As Long
    Dim strTmp As String
    Dim lngStartDay As Long
    
    DrawBodyInfo = True
    
    lngTmpX = X
    lngTmpY = Y
    '参考高度
    objDraw.Font.Name = "宋体"
    objDraw.Font.Size = 9 * sngScale
    objDraw.Font.Bold = False
    H_9pt = objDraw.TextHeight("字")
    If colFirstWidth < H_9pt + H_9pt \ 2 Then DrawBodyInfo = False: Exit Function
    objDraw.DrawStyle = 0
    '画首列
   
    Call DrawCell(objDraw, "日    期", X, Y + (H_9pt + H_9pt \ 2) * 0, colFirstWidth, H_9pt + H_9pt \ 2, 0, 0, , , , objDraw.Font, "1111", 1, 1, False)
'    Call DrawCell(objDraw, "住院天数", X, Y + (H_9pt + H_9pt \ 2) * 1, colFirstWidth, H_9pt + H_9pt \ 2, 0, 0, , , , objDraw.Font, "1111", 1, 1, False)
'    Call DrawCell(objDraw, "手/娩后日数", X, Y + (H_9pt + H_9pt \ 2) * 2, colFirstWidth, H_9pt + H_9pt \ 2, 0, 0, , , , objDraw.Font, "1111", 1, 1, False)
    
    Call DrawLine(objDraw, X, Y, X + colFirstWidth, Y, , , 2)
    Call DrawLine(objDraw, X, Y, X, Y + (H_9pt + H_9pt \ 2) * 2 + H_9pt + H_9pt \ 2, , , 2)
'    Call DrawLine(objDraw, X + colFirstWidth, Y, X + colFirstWidth, Y + (H_9pt + H_9pt \ 2) * 2 + H_9pt + H_9pt \ 2, , , 2)
    
    '画各列
    lngLeft = lngTmpX + colFirstWidth
    lngStartDay = GetSplitStr(strPatiDay, 0)
    
    For intDay = 0 To 6

        strTmp = GetSplitStr(strDate, intDay)
        If Right(strTmp, 5) = "01-01" Then
            '一年的第一天

        ElseIf strTmp = Format(strBeginDate, "yyyy-MM-dd") Then
            '入院第一天，写上年份

        ElseIf intDay = 0 Then
            strTmp = Right(strTmp, 5)
        ElseIf Right(strTmp, 2) = "01" Then
            strTmp = Right(strTmp, 5)
        Else
            strTmp = Right(strTmp, 2)
        End If

        Call DrawCell(objDraw, strTmp, lngLeft + (HOUR_STEP_Twips * 6) * intDay, Y, (HOUR_STEP_Twips * 6), H_9pt + H_9pt \ 2, 0, 0, , &HFF0000, , objDraw.Font, "1111", 1, 1, False)
        
        Call DrawLine(objDraw, lngLeft + (HOUR_STEP_Twips * 6) * intDay, Y, lngLeft + (HOUR_STEP_Twips * 6) * intDay + (HOUR_STEP_Twips * 6), Y, , , 2)
        Call DrawLine(objDraw, lngLeft + (HOUR_STEP_Twips * 6) * intDay, Y, lngLeft + (HOUR_STEP_Twips * 6) * intDay, Y + H_9pt + H_9pt \ 2, IIf(intDay = 0, 0, vbRed), , 2)
        Call DrawLine(objDraw, lngLeft + (HOUR_STEP_Twips * 6) * intDay + (HOUR_STEP_Twips * 6), Y, lngLeft + (HOUR_STEP_Twips * 6) * intDay + (HOUR_STEP_Twips * 6), Y + H_9pt + H_9pt \ 2, vbRed, , 2)

'        Call DrawCell(objDraw, CStr(lngStartDay + intDay), lngLeft + (HOUR_STEP_Twips * 6) * intDay, Y + H_9pt * 1 + H_9pt \ 2, (HOUR_STEP_Twips * 6), H_9pt + H_9pt \ 2, 0, 0, , &HFF0000, , objDraw.Font, "1111", 1, 1, False)
'
'        Call DrawLine(objDraw, lngLeft + (HOUR_STEP_Twips * 6) * intDay, Y + H_9pt * 1 + H_9pt \ 2, lngLeft + (HOUR_STEP_Twips * 6) * intDay, Y + H_9pt * 1 + H_9pt \ 2 + H_9pt + H_9pt \ 2, , , 2)
'        Call DrawLine(objDraw, lngLeft + (HOUR_STEP_Twips * 6) * intDay + (HOUR_STEP_Twips * 6), Y + H_9pt * 1 + H_9pt \ 2, lngLeft + (HOUR_STEP_Twips * 6) * intDay + (HOUR_STEP_Twips * 6), Y + H_9pt * 1 + H_9pt \ 2 + H_9pt + H_9pt \ 2, , , 2)
'
'        If intDay <= UBound(Split(strOPSFate, "'")) + 1 Then
'            Call DrawCell(objDraw, GetSplitStr(strOPSFate, intDay), lngLeft + (HOUR_STEP_Twips * 6) * intDay, Y + (H_9pt + H_9pt \ 2) * 2, (HOUR_STEP_Twips * 6), H_9pt + H_9pt \ 2, 0, 0, , 255, , objDraw.Font, "1111", 1, 1, False)
'        Else
'            Call DrawCell(objDraw, "", lngLeft + (HOUR_STEP_Twips * 6) * intDay, Y + (H_9pt + H_9pt \ 2) * 2, (HOUR_STEP_Twips * 6), H_9pt + H_9pt \ 2, 0, 0, , 255, , objDraw.Font, "1111", 1, 1, False)
'        End If
'        Call DrawLine(objDraw, lngLeft + (HOUR_STEP_Twips * 6) * intDay, Y + (H_9pt + H_9pt \ 2) * 2, lngLeft + (HOUR_STEP_Twips * 6) * intDay, Y + (H_9pt + H_9pt \ 2) * 2 + H_9pt + H_9pt \ 2, , , 2)
'        Call DrawLine(objDraw, lngLeft + (HOUR_STEP_Twips * 6) * intDay + (HOUR_STEP_Twips * 6), Y + (H_9pt + H_9pt \ 2) * 2, lngLeft + (HOUR_STEP_Twips * 6) * intDay + (HOUR_STEP_Twips * 6), Y + (H_9pt + H_9pt \ 2) * 2 + H_9pt + H_9pt \ 2, , , 2)
                  
    Next
    X = lngTmpX
    Y = lngTmpY + (H_9pt + H_9pt \ 2)
End Function

Private Sub DrawBodyTopScale(objDraw As Object, X As Long, Y As Long, Optional sngScale As Long = 1)
    '功能:画出上面的标尺
    '
    
    '顺序:
    '画图顺序号:=4/9   这是画所有图形的第四步
    Dim i As Integer, j As Integer '循环之用
    Dim H_9pt As Long
    Dim lngTmpX As Long, lngTmpY As Long
    
    Dim lngHourBegin As Long
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    lngHourBegin = Val(zlDatabase.GetPara("体温开始时间", glngSys, 1255, 4))
    lngTmpX = X: lngTmpY = Y
    
    '参考高度
    objDraw.Font.Name = "宋体"
    objDraw.Font.Size = 9 * sngScale
    objDraw.Font.Bold = False
    
    objDraw.DrawStyle = 0
    H_9pt = objDraw.TextHeight("字")
'    For i = 1 To 7
'        '画上下午
'        Call DrawCell(objDraw, "上午", X, Y, HOUR_STEP_Twips * 3, (H_9pt * 2), , , , , , , , 1, 1)
'        Call DrawLine(objDraw, X, Y, X, Y + (H_9pt * 2), , , 2)
'
'        X = X + HOUR_STEP_Twips * 3
'
'        Call DrawCell(objDraw, "下午", X, Y, HOUR_STEP_Twips * 3, (H_9pt * 2), , , , , , , , 1, 1)
'        Call DrawLine(objDraw, X + HOUR_STEP_Twips * 3, Y, X + HOUR_STEP_Twips * 3, Y + (H_9pt * 2), , , 2)
'
'        X = X + HOUR_STEP_Twips * 3
'
'    Next
'    Y = lngTmpY + H_9pt * 2
'    X = lngTmpX
    
    Dim intCount As Integer
    Dim lngColor As Long
    
    For i = 1 To 7
        
        intCount = 0
        For j = lngHourBegin To 12 - (4 - lngHourBegin) Step 4
            intCount = intCount + 1
            If intCount < 2 Then
                lngColor = RGB(200, 0, 0)
            Else
                lngColor = 0
            End If
            
            If j = 12 - lngHourBegin Then
'                objDraw.FontSize = 8
                Call DrawCell(objDraw, CStr(j), X, Y, HOUR_STEP_Twips, (H_9pt * 2 + H_9pt \ 2), , , , lngColor, , objDraw.Font, , 1, 1, True)
                X = X + HOUR_STEP_Twips
            Else
'                objDraw.FontSize = 9
                Call DrawCell(objDraw, CStr(j), X, Y, HOUR_STEP_Twips, (H_9pt * 2 + H_9pt \ 2), , , , lngColor, , objDraw.Font, , 1, 1, True)
                X = X + HOUR_STEP_Twips
            End If
        Next
        
        intCount = 0
        For j = lngHourBegin To 12 - (4 - lngHourBegin) Step 4
            intCount = intCount + 1
            If intCount >= 2 Then
                lngColor = RGB(200, 0, 0)
            Else
                lngColor = 0
            End If
            If j = 12 - lngHourBegin Then
'                objDraw.FontSize = 8
                Call DrawCell(objDraw, CStr(j), X, Y, HOUR_STEP_Twips, (H_9pt * 2 + H_9pt \ 2), , , , lngColor, , objDraw.Font, , 1, 1, True)
                X = X + HOUR_STEP_Twips
            Else
'                objDraw.FontSize = 9
                Call DrawCell(objDraw, CStr(j), X, Y, HOUR_STEP_Twips, (H_9pt * 2 + H_9pt \ 2), , , , lngColor, , objDraw.Font, , 1, 1, True)
                X = X + HOUR_STEP_Twips
            End If
        Next
        
    Next
    
    X = lngTmpX
    For i = 1 To 7
        X = X + HOUR_STEP_Twips * 6
        Call DrawLine(objDraw, X, Y, X, Y + (H_9pt * 2 + H_9pt \ 2), vbRed, 1, 2)
    Next
    
    objDraw.FontSize = 9
    X = lngTmpX
    Y = lngTmpY + H_9pt * 2 + H_9pt \ 2
End Sub
'
'Private Function DrawBodyGraph(objDraw As Object, _
'                                x As Long, _
'                                y As Long, _
'                                strArrItemInfo() As String, _
'                                strArrValue() As String, _
'                                strArrValueComment() As String, _
'                                strArrValueInOut() As String, _
'                                Optional strIndex As String = "", _
'                                Optional sngScale As Long = 1) As Boolean
'
'    '功能:在内部表格画上数据点与线
'    '参数:strArrItemInfo()=体温表项目信息  格式：(0)="记录色'最高行'项目名'最大值'最小值'单位值'记录符"
'    '                                                  ... ...
'    '                                            (n)="记录色'最高行'项目名'最大值'最小值'单位值'记录符"
'    '     strArrValue()=值数组  格式：(0)="43'54'65'-2'76'57'657'543 ... ...  65'876'"  <---42(7天*6)个数据，某天某时没有数据为-2
'    '                                 (1)="87'987'09'5'36'-2'865'963 ... ...  -2'53'"   <---42(7天*6)个数据，某天某时没有数据为-2
'    '                                 ......
'    '                                 (n)="453'445'46'-2'33'-2'865'45 ... ...  3'3'"   <---42(7天*6)个数据，某天某时没有数据为-2
'    '                                       上面的参数说明：此与的元素个数应与体温表项目信息的元素个数相同
'    '     strArrValueComment()=说明数组 格式如同值数组
'    '     strIndex=用字符串来来标记要画哪些数据 格式："0'4'2'3"  （要求：列表数总和不能大于“值数组”元素的个数）
'    '画图顺序号:=6/9   这是画所有图形的第六步
'
'    Dim lngRow As Long, lngCol As Long '循环之用
'    Dim lngTmpX As Long, lngTmpY As Long
'    Dim H_9pt As Long
'    Dim W_9pt As Long
'    Dim lngArrXY() As Long
'    Dim intItemCount As Integer  '项目个数
'    Dim lngMax As Long, lngMin As Long '最大最小值
'    Dim lng项目序号 As Long
'    Dim lngColor As Long, lngTopRow As Long, lngStep As Single, strTag As String    '记录色,最高行,单位值,记录符
'    Dim intItemDrawCount As Integer '要画的项目个数
'    Dim lngValue As Double    '当前值
'    Dim lngValue1 As Double    '当前值
'    Dim lngValue2 As Double    '当前值
'    Dim strComment As String  '当前要画的说明
'    Dim lngUpIndex As Long
'    Dim Y2 As Long, i As Long
'    Dim lngCommentColor As Long
'    Dim aryData() As String
'    Dim aryTmp As Variant
'    Dim blnStop As Boolean
'    Dim rsPoint As New ADODB.Recordset
'
'    Dim X1 As Long
'    Dim Y1 As Long
'    Dim dblHeight As Double         '40-42度之间的有效打印高度
'    Dim rsTmp As New ADODB.Recordset
'    On Error GoTo errHand
'
'    DrawBodyGraph = False
'    If IsEmpty(strArrItemInfo) Or IsEmpty(strArrValue) Then Exit Function
'    If UBound(strArrItemInfo) <> UBound(strArrValue) Then Exit Function
'    lngTmpY = y
'    lngTmpX = x
'    objDraw.DrawStyle = 0
'    W_9pt = objDraw.TextWidth("字")
''    H_9pt = objDraw.TextHeight("字")
'    H_9pt = ROWHEIGHT * 10 / 3
'
'    intItemCount = UBound(strArrValue) + 1
'    'ReDim lngArrXY(7 * 6 - 1, 4) '记录坐标
'    ReDim lngArrXY(7 * 6 - 1, 5) '记录坐标      ,最后增加一维,用于记录是否为不升
'
'    Dim mpt脉搏() As POINTAPI
'    Dim mpt心率() As POINTAPI
'    Dim mpt体温() As POINTAPI
'
'    ReDim mpt脉搏(0 To 41)
'    ReDim mpt心率(0 To 41)
'    ReDim mpt体温(0 To 41)
'
'    Dim lngX As Long
'    Dim lngY As Long
'    Dim lngY1 As Long
'    Dim strtmp As String
'    Dim intPointCount As Integer
'    Dim lngYMax As Long
'    Dim intCharNumber As Integer
'    Dim blnPrint As Boolean
'
'    blnPrint = (Val(zlDatabase.GetPara("不打印脉搏短绌图形", glngSys, 1255, "0")) = 0)
'
'    lngYMax = (lngTmpY + 3 * H_9pt / 2) + 40 * 3 * H_9pt / 2 - H_9pt / 2
'    Call PointInit(rsPoint)
'
'    '打印入出转
'    For lngCol = 0 To 41
'
'        lngX = lngTmpX + lngCol * HOUR_STEP_Twips + HOUR_STEP_Twips / 2
'        lngY = (lngTmpY + 3 * H_9pt / 2)
'        lngY1 = (lngTmpY + 3 * H_9pt / 2) + 35 * 3 * H_9pt / 2
'
'        If strArrValueInOut(lngCol) <> "" Then
'
'            strComment = strArrValueInOut(lngCol)
'            aryTmp = Split(strComment, ";")
'
'            Set rsTmp = New ADODB.Recordset
'            '20090926:必须在40-42度间打印,单独一条信息如果超长就缩小字体,有多条信息则延后面一格打印,如果是最后一格就直接全部打印
'            Set rsTmp = New ADODB.Recordset
'            rsTmp.Fields.Append "列号", adVarChar, 30
'            rsTmp.Fields.Append "时间", adVarChar, 30
'            rsTmp.Fields.Append "结果", adVarChar, 50
'            '20090926--
'            rsTmp.Fields.Append "类型", adVarChar, 50       '记录是入出转,手术出院,还是未记说明,上标说明
'            rsTmp.Fields.Append "打印列", adVarChar, 30
'            rsTmp.Fields.Append "坐标", adVarChar, 30
'            rsTmp.Fields.Append "高度", adVarChar, 30       '未记说明及上标说明不用管高度
'            rsTmp.Fields.Append "字体大小", adVarChar, 50
'            '----------
'            rsTmp.Open
'
'            For i = 0 To UBound(aryTmp)
'                If Trim(aryTmp(i)) <> "" Then
'                    rsTmp.AddNew
'                    rsTmp.Fields("时间").Value = Split(aryTmp(i), "'")(1)
'                    rsTmp.Fields("结果").Value = Split(aryTmp(i), "'")(0)
'                End If
'            Next
'            rsTmp.Sort = "时间"
'            If rsTmp.RecordCount > 0 Then
'                rsTmp.MoveFirst
'                strComment = ""
'                Do While Not rsTmp.EOF
'                    strComment = strComment & " " & rsTmp.Fields("结果").Value
'                    rsTmp.MoveNext
'                Loop
'
'                strComment = Trim(strComment)
'
'            End If
'
'            intCharNumber = 0
'            For i = 1 To Len(strComment)
'
'                '在42度下输出，1是固定的纵向格式数
'
'                If lngY < lngYMax Then
'                    strtmp = Mid(strComment, i, 1)
'
'                    If Asc(strtmp) < 0 Then
'                        If intCharNumber Mod 2 = 1 Then lngY = lngY + ROWHEIGHT * 2.5
'                    End If
'
'                    Call DrawRotateText(objDraw, lngX - objDraw.TextWidth(strtmp) / 2 + 15, lngY + 15, strtmp, 255)
'
'                    If Asc(strtmp) < 0 Then
'                        lngY = lngY + ROWHEIGHT * 5
'                        intCharNumber = 0
'                    Else
'                        lngY = lngY + ROWHEIGHT * 2.5
'                        intCharNumber = intCharNumber + 1
'                    End If
'
'
'
'                End If
'            Next
'        End If
'
'        '说明
'        strComment = Split(strArrValueComment(0), "'")(lngCol)
'        If strComment <> "" Then
'            aryTmp = Split(strComment, ";")
'
'            '上标
'            If aryTmp(0) <> "" Then
'                If strArrValueInOut(lngCol) <> "" Then aryTmp(0) = " " & aryTmp(0)
'                intCharNumber = 0
'                For i = 1 To Len(aryTmp(0))
'                    If lngY < lngYMax Then
'                        strtmp = Mid(aryTmp(0), i, 1)
'
'                        If Asc(strtmp) < 0 Then
'                            If intCharNumber Mod 2 = 1 Then lngY = lngY + ROWHEIGHT * 2.5
'                        End If
'
'                        Call DrawRotateText(objDraw, lngX - objDraw.TextWidth(strtmp) / 2 + 15, lngY + 15, strtmp, -2147483635)
'
'                        If Asc(strtmp) < 0 Then
'                            lngY = lngY + ROWHEIGHT * 5
'                            intCharNumber = 0
'                        Else
'                            lngY = lngY + ROWHEIGHT * 2.5
'                            intCharNumber = intCharNumber + 1
'                        End If
'
'                    End If
'                Next
'            End If
'
'            '下标
'            intCharNumber = 0
'            If aryTmp(1) <> "" Then
'                For i = 1 To Len(aryTmp(1))
'                    If lngY1 < lngYMax Then
'                        strtmp = Mid(aryTmp(1), i, 1)
'
'                        If Asc(strtmp) < 0 Then
'                            If intCharNumber Mod 2 = 1 Then lngY = lngY + ROWHEIGHT * 2.5
'                        End If
'
'                        Call DrawRotateText(objDraw, lngX - objDraw.TextWidth(strtmp) / 2 + 15, lngY1 + 15, strtmp, -2147483635)
'
'                        If Asc(strtmp) < 0 Then
'                            lngY1 = lngY1 + ROWHEIGHT * 5
'                            intCharNumber = 0
'                        Else
'                            lngY1 = lngY1 + ROWHEIGHT * 2.5
'                            intCharNumber = intCharNumber + 1
'                        End If
'
'                    End If
'                Next
'            End If
'        End If
'    Next
'
'    Dim int体温 As Integer
'    Dim int脉搏 As Integer
'
'    If Trim(strIndex) = "" Then
'
'        '按项目个数循环
'
'        For intItemDrawCount = 0 To intItemCount - 1
'
'            intPointCount = -1
'            lngTopRow = CInt(Split(strArrItemInfo(intItemDrawCount), "'")(1))
'            lngStep = Val(Split(strArrItemInfo(intItemDrawCount), "'")(5))
'
'            lngMax = CLng(Split(strArrItemInfo(intItemDrawCount), "'")(3))
'            lngMin = CLng(Split(strArrItemInfo(intItemDrawCount), "'")(4))
'
'            lng项目序号 = CLng(Split(strArrItemInfo(intItemDrawCount), "'")(7))
'            Select Case lng项目序号
'            Case 1
'                int体温 = intItemDrawCount
'            Case 2
'                int脉搏 = intItemDrawCount
'            End Select
'            objDraw.ForeColor = lngColor
'
'            'Erase lngArrXY
'            '初始化以便在第一次时不进行连线
'
'            lngUpIndex = -1
'
'            '循环画数据
'            For lngCol = 0 To 7 * 6 - 1 '纵
'
'                lngColor = CLng(Split(strArrItemInfo(intItemDrawCount), "'")(0))
'                strTag = Trim(Split(strArrItemInfo(intItemDrawCount), "'")(6))
'
'                x = lngTmpX + (lngCol + 1) * HOUR_STEP_Twips - HOUR_STEP_Twips / 2
'
'                '1、求出当前值比例值（(当前值/单位值+最高行)*(H_9pt+H_9pt\2)），此时为相对Y坐标位置值
'                '   并将此时的坐标保存在变量里
'                If Val(Split(strArrValue(intItemDrawCount), "'")(lngCol)) = -2 Then
'                    y = lngTmpY
'                    Y2 = lngTmpY
'                    lngArrXY(lngCol, 0) = -1
'                    lngArrXY(lngCol, 1) = x
'                    lngArrXY(lngCol, 2) = y
'                    lngArrXY(lngCol, 3) = Y2
'                    lngArrXY(lngCol, 4) = -1
'                    lngArrXY(lngCol, 5) = 0
'                Else
'                    y = lngTmpY
'                    Y2 = lngTmpY
'
'                    aryData = Split(Split(strArrValue(intItemDrawCount), "'")(lngCol), ",")
'
'                    For i = 0 To UBound(aryData)
'                        lngValue = (lngTopRow - 1) * lngStep + lngMax
'
'                        '超过当前项目的最小值最大值时,直接取最小值或最大值
'                        lngArrXY(lngCol, 5) = IIf(Left(aryData(i), InStr(aryData(i), ";") - 1) = "不升", 1, 0)
'
'                        Select Case Val(Left(aryData(i), InStr(aryData(i), ";") - 1))
'                        Case Is < lngMin
'                            aryData(i) = lngMin & ";" & Mid(aryData(i), InStr(aryData(i), ";") + 1)
'                        Case Is > lngMax
'                            aryData(i) = lngMax & ";" & Mid(aryData(i), InStr(aryData(i), ";") + 1)
'                        End Select
'
'                        If InStr(aryData(i), ";") > 0 Then
'                            lngValue1 = lngValue - Val(Left(aryData(i), InStr(aryData(i), ";") - 1))
'                        Else
'                            lngValue1 = lngValue - Val(aryData(i))
'                        End If
'
'                        If i = 0 Then
'                            y = lngTmpY + (lngValue1 / lngStep) * (H_9pt + H_9pt \ 2)
'                            lngArrXY(lngCol, 0) = 2
'                        Else
'                            Y2 = lngTmpY + (lngValue1 / lngStep) * (H_9pt + H_9pt \ 2)
'                            lngArrXY(lngCol, 4) = 2
'                        End If
'
'                        '因为第3个可能存的是文本（说明体温项目的部位）
'                        If i = 1 Then Exit For
'                    Next
''                    lngArrXY(lngCol, 0) = 2
'                    lngArrXY(lngCol, 1) = x
'                    lngArrXY(lngCol, 2) = y
'                    lngArrXY(lngCol, 3) = Y2
'
'                    '2、在X-W_9pt\2,Y-H_9pt\2 处画一个记录符
'                    lngX = x - W_9pt \ 3
'                    lngY = y - H_9pt \ 3
'
'                    lngX = x
'                    lngY = y
'
'                    Select Case lng项目序号
'                    Case 1
'                        If InStr(aryData(0), ";") > 0 Then
'                            aryTmp = Split(aryData(0), ";")
'                            Select Case aryTmp(5)
'                            Case "口温"
'                                strTag = mstrChar(0)
'                            Case "腋温"
'                                strTag = mstrChar(1)
'                            Case "肛温"
'                                strTag = mstrChar(2)
'                            Case Else
'                                strTag = mstrChar(1)
'                            End Select
'
'                            X1 = lngX
'                            Y1 = lngY
'                            Select Case Val(aryTmp(0))
'                            Case Is <= lngMin
'                                strTag = "·"
'                                Call DrawLine(objDraw, X1, Y1, X1, Y1 + 200, lngColor, , , True)
'                            Case Is >= lngMax
'                                strTag = "·"
'                                Call DrawLine(objDraw, X1, Y1, X1, Y1 - 200, lngColor, , , True)
'                            End Select
'
'                            If Val(aryTmp(2)) = 1 Then
'                                '复试合格
'                                Call DrawText(objDraw, X1 - 50, Y1 - 250, "v", lngColor)
'                            End If
'
'                            mpt体温(lngCol).x = X1
'                            mpt体温(lngCol).y = Y1
'
'                        End If
'                    Case 2
'                        '如果为空表示缺省字符;否则赋为空,在绘图时发现为空,则取资源文件中的位图
'                        aryTmp = Split(aryData(0), ";")
'                        strTag = IIf(aryTmp(5) = "起搏器", "", mstrPulse)
'
'                        mpt脉搏(lngCol).x = lngX
'                        mpt脉搏(lngCol).y = lngY
'
'                        X1 = lngX
'                        Y1 = lngY
'                        If InStr(aryData(0), ";") > 0 Then
'                            aryTmp = Split(aryData(0), ";")
'                            Select Case Val(aryTmp(0))
'                            Case Is <= lngMin
'                                Call DrawLine(objDraw, X1, Y1, X1, Y1 + 200, lngColor, , , True)
'                            Case Is >= lngMax
'                                Call DrawLine(objDraw, X1, Y1, X1, Y1 - 200, lngColor, , , True)
'                            End Select
'                        End If
'
'                    Case -1
'                        mpt心率(lngCol).x = lngX
'                        mpt心率(lngCol).y = lngY
'
'                        X1 = lngX
'                        Y1 = lngY
'
'                        If InStr(aryData(0), ";") > 0 Then
'                            aryTmp = Split(aryData(0), ";")
'                            Select Case Val(aryTmp(0))
'                            Case Is <= lngMin
'                                Call DrawLine(objDraw, X1, Y1, X1, Y1 + 200, lngColor, , , True)
'                            Case Is >= lngMax
'                                Call DrawLine(objDraw, X1, Y1, X1, Y1 - 200, lngColor, , , True)
'                            End Select
'                        End If
'
'                    Case Else
'                        If lng项目序号 = 3 Then
'                            aryTmp = Split(aryData(0), ";")
'                            strTag = IIf(aryTmp(5) = "呼吸机", "", mstrBreath)
'                        End If
'
'                        X1 = lngX
'                        Y1 = lngY
'                        Select Case Val(aryData(0))
'                        Case Is <= lngMin
'                            Call DrawLine(objDraw, X1, Y1, X1, Y1 + 200, lngColor, , , True)
'                        Case Is >= lngMax
'                            Call DrawLine(objDraw, X1, Y1, X1, Y1 - 200, lngColor, , , True)
'                        End Select
'
'                    End Select
'
'                    If lng项目序号 = 1 Then         '体温
'                        Call PointAdd(rsPoint, lngX, lngY, lng项目序号, strTag, lngColor, lngCol, CStr(aryTmp(5)))
'                    ElseIf lng项目序号 = 3 Then     '呼吸
'                        Call PointAdd(rsPoint, lngX, lngY, lng项目序号, strTag, lngColor, lngCol, CStr(aryTmp(5)), IIf(strTag = "", "BREATH", ""))
'                    ElseIf lng项目序号 = 2 Then     '脉搏
'                        Call PointAdd(rsPoint, lngX, lngY, lng项目序号, strTag, lngColor, lngCol, CStr(aryTmp(5)), IIf(strTag = "", "PACEMAKER", ""))
'                    Else
'                        Call PointAdd(rsPoint, lngX, lngY, lng项目序号, strTag, lngColor, lngCol, "")
'                    End If
'
'                    lngColor = CLng(Split(strArrItemInfo(intItemDrawCount), "'")(0))
'
'                    If Y2 <> lngTmpY Then
'
'                        Select Case lng项目序号
'                        Case 2
'
'                            '心率与脉搏共享使用时
''                            If mint心率应用 = 2 Then
'
'                                strTag = mstr心率符号
'                                mpt心率(lngCol).x = lngX
'                                mpt心率(lngCol).y = (Y2 - H_9pt \ 3)
'                                lngColor = 255
'
'                                If InStr(aryData(1), ";") > 0 Then
'                                    X1 = lngX
'                                    Y1 = Y2
'                                    aryTmp = Split(aryData(1), ";")
'                                    Select Case Val(aryTmp(0))
'                                    Case Is <= lngMin
'                                        Call DrawLine(objDraw, X1, Y1, X1, Y1 + 200, lngColor, , , True)
'                                    Case Is >= lngMax
'                                        Call DrawLine(objDraw, X1, Y1, X1, Y1 - 200, lngColor, , , True)
'                                    End Select
'                                End If
'
''                            End If
'                        Case 1
'                            strTag = "○"
'                            lngColor = 255
'                        End Select
'
'                        If lng项目序号 = 1 Then
'                            Call PointAdd(rsPoint, lngX, Y2 - H_9pt \ 3, lng项目序号, strTag, lngColor, lngCol, CStr(aryTmp(5)))
'                        Else
'                            Call PointAdd(rsPoint, lngX, Y2 - H_9pt \ 3, lng项目序号, strTag, lngColor, lngCol, "")
'                        End If
'
'                    End If
'
'                    '3、找到前一个坐标的进行连线。如果当一个坐标的（0）元素为0时表示不连
'                    '--------------------------------------------------------------------------------------------------
'                    If lngCol > 0 Then
'                        lngColor = CLng(Split(strArrItemInfo(intItemDrawCount), "'")(0))
'                        If lngUpIndex > -1 Then
'                            If lngArrXY(lngUpIndex, 0) = 2 Then
'                                If lngArrXY(lngUpIndex, 5) <> 1 And lngArrXY(lngCol, 5) <> 1 Then
'                                    Call DrawLine(objDraw, lngArrXY(lngUpIndex, 1), lngArrXY(lngUpIndex, 2), x, y, lngColor)
'                                End If
'
'                                If Y2 <> lngTmpY Then
'                                    Select Case lng项目序号
'                                    Case 1          '体温项目
'                                        Call DrawLine(objDraw, x, y, x, Y2, 255, IIf(Y2 < y, 0, 2), , True)
'                                    Case 2
'                                        If lngArrXY(lngUpIndex, 4) = 2 Then
'                                            Call DrawLine(objDraw, lngArrXY(lngUpIndex, 1), lngArrXY(lngUpIndex, 3), x, Y2, lngColor)
'                                        End If
'                                    End Select
'
'                                End If
'                            Else
'                                '虽然断点，如果当天有两个点，则先画出来
'                                '--------------------------------------------------------------------------------------
'                                If Y2 <> lngTmpY Then
'                                    Select Case lng项目序号
'                                    Case 1
'                                        Call DrawLine(objDraw, x, y, x, Y2, 255, IIf(Y2 < y, 0, 2), , True)
'                                    End Select
'                                End If
'
'                                '找出当天及昨天上一个有效点,只有隔天无数据才不连线
'                                '--------------------------------------------------------------------------------------
'                                lngUpIndex = lngUpIndex - 1
'                                blnStop = False
'                                Do While lngUpIndex >= (lngCol \ 6) * 6 - 6 And lngUpIndex >= 0
'
'                                    If blnStop = False Then
'                                        If Split(strArrValueComment(0), "'")(lngUpIndex + 1) <> "" Then
'                                            aryTmp = Split(Split(strArrValueComment(0), "'")(lngUpIndex + 1), ";")
'                                        Else
'                                            aryTmp = Split(";;", ";")
'                                        End If
'                                        blnStop = (Val(aryTmp(2)) = 1)
'                                    End If
'                                    If blnStop Then
'                                        Exit Do
'                                    End If
'
'                                    If lngArrXY(lngUpIndex, 0) = 2 And lngArrXY(lngUpIndex, 5) <> 1 And lngArrXY(lngCol, 5) <> 1 Then
'                                        If blnStop = False Then
'                                            blnStop = False
'                                            lngColor = CLng(Split(strArrItemInfo(intItemDrawCount), "'")(0))
'
'                                            Call DrawLine(objDraw, lngArrXY(lngUpIndex, 1), lngArrXY(lngUpIndex, 2), x, y, lngColor)
'
'                                            If Y2 <> lngTmpY Then
'                                                Select Case lng项目序号
'                                                Case 1
'                                                    Call DrawLine(objDraw, x, y, x, Y2, 255, IIf(Y2 < y, 0, 2), , True)
'                                                Case 2                  '脉搏附加心率情况
'                                                    If lngArrXY(lngUpIndex, 4) = 2 Then
'                                                        Call DrawLine(objDraw, lngArrXY(lngUpIndex, 1), lngArrXY(lngUpIndex, 3), x, Y2, lngColor)
'                                                    End If
'                                                End Select
'                                            End If
'                                        End If
'                                        Exit Do
'                                    End If
'                                    lngUpIndex = lngUpIndex - 1
'                                Loop
'
'                                '--------------------------------------------------------------------------------------
'
'                            End If
'                        End If
'                    ElseIf Y2 <> lngTmpY Then
'                        Select Case lng项目序号
'                        Case 1
'                            Call DrawLine(objDraw, x, y, x, Y2, 255, IIf(Y2 < y, 0, 2), , True)
'                        End Select
'                    End If
'                End If
'                lngUpIndex = lngCol
'            Next
'        Next
'
'        '根据脉搏和心率坐标形成多边形，并进行连线和填充
'        '--------------------------------------------------------------------------------------------------------------
'        If blnPrint Then Call DrawPoly(objDraw, mpt脉搏, mpt心率)
'
'    End If
'
'    '画点的字符或图形
'    '--------------------------------------------------------------------------------------------------------------
'    Call DrawPoint(objDraw, rsPoint)
'
'    x = lngTmpX - (UBound(strArrItemInfo) + 1) * (W_9pt + W_9pt \ 2)
'
'    Exit Function
'
'    '------------------------------------------------------------------------------------------------------------------
'errHand:
'    If ErrCenter = 1 Then
'        Resume
'    End If
'End Function

Private Function DrawBodyGraph(objDraw As Object, _
                                X As Long, _
                                Y As Long, _
                                strArrItemInfo() As String, _
                                strArrValue() As String, _
                                strArrValueComment() As String, _
                                strArrValueInOut() As String, _
                                Optional strIndex As String = "", _
                                Optional sngScale As Long = 1) As Boolean
                            
    '功能:在内部表格画上数据点与线
    '参数:strArrItemInfo()=体温表项目信息  格式：(0)="记录色'最高行'项目名'最大值'最小值'单位值'记录符"
    '                                                  ... ...
    '                                            (n)="记录色'最高行'项目名'最大值'最小值'单位值'记录符"
    '     strArrValue()=值数组  格式：(0)="43'54'65'-2'76'57'657'543 ... ...  65'876'"  <---42(7天*6)个数据，某天某时没有数据为-2
    '                                 (1)="87'987'09'5'36'-2'865'963 ... ...  -2'53'"   <---42(7天*6)个数据，某天某时没有数据为-2
    '                                 ......
    '                                 (n)="453'445'46'-2'33'-2'865'45 ... ...  3'3'"   <---42(7天*6)个数据，某天某时没有数据为-2
    '                                       上面的参数说明：此与的元素个数应与体温表项目信息的元素个数相同
    '     strArrValueComment()=说明数组 格式如同值数组
    '     strIndex=用字符串来来标记要画哪些数据 格式："0'4'2'3"  （要求：列表数总和不能大于“值数组”元素的个数）
    '画图顺序号:=6/9   这是画所有图形的第六步
    
    Dim lngRow As Long, lngCol As Long '循环之用
    Dim lngTmpX As Long, lngTmpY As Long
    Dim H_9pt As Long
    Dim W_9pt As Long
    Dim lngArrXY() As Long
    Dim intItemCount As Integer  '项目个数
    Dim lngMax As Long, lngMin As Long '最大最小值
    Dim lng项目序号 As Long
    Dim lngColor As Long, lngTopRow As Long, lngStep As Single, strTag As String    '记录色,最高行,单位值,记录符
    Dim intItemDrawCount As Integer '要画的项目个数
    Dim lngValue As Double    '当前值
    Dim lngValue1 As Double    '当前值
    Dim lngValue2 As Double    '当前值
    Dim strComment As String  '当前要画的说明
    Dim lngUpIndex As Long
    Dim Y2 As Long, i As Long
    Dim lngCommentColor As Long
    Dim aryData() As String
    Dim aryTmp As Variant
    Dim blnStop As Boolean
    Dim rsPoint As New ADODB.Recordset
    
    Dim X1 As Long
    Dim Y1 As Long
    Dim dblHeight As Double         '40-42度之间的有效打印高度
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo errHand
    
    DrawBodyGraph = False
    If IsEmpty(strArrItemInfo) Or IsEmpty(strArrValue) Then Exit Function
    If UBound(strArrItemInfo) <> UBound(strArrValue) Then Exit Function
    lngTmpY = Y
    lngTmpX = X
    objDraw.DrawStyle = 0
    W_9pt = objDraw.TextWidth("字")
'    H_9pt = objDraw.TextHeight("字")
    H_9pt = ROWHEIGHT * 10 / 3
    
    intItemCount = UBound(strArrValue) + 1
    'ReDim lngArrXY(7 * 6 - 1, 4) '记录坐标
    ReDim lngArrXY(7 * 6 - 1, 5) '记录坐标      ,最后增加一维,用于记录是否为不升
    
    Dim mpt脉搏() As POINTAPI
    Dim mpt心率() As POINTAPI
    Dim mpt体温() As POINTAPI
    
    ReDim mpt脉搏(0 To 41)
    ReDim mpt心率(0 To 41)
    ReDim mpt体温(0 To 41)
    
    Dim lngX As Long
    Dim lngY As Long
    Dim lngY1 As Long
    Dim strTmp As String
    Dim intPointCount As Integer
    Dim lngYMax As Long
    Dim intCharNumber As Integer
    Dim blnPrint As Boolean
    
    blnPrint = (Val(zlDatabase.GetPara("不打印脉搏短绌图形", glngSys, 1255, "0")) = 0)
    
    lngYMax = (lngTmpY + 3 * H_9pt / 2) + 44 * 3 * H_9pt / 2 - H_9pt / 2
    Call PointInit(rsPoint)
    
    '打印入出转
            
    Set rsTmp = New ADODB.Recordset
    '20090926:必须在40-42度间打印,单独一条信息如果超长就缩小字体,有多条信息则延后面一格打印,如果是最后一格就直接全部打印
    Set rsTmp = New ADODB.Recordset
    rsTmp.Fields.Append "列号", adDouble, 30
    rsTmp.Fields.Append "时间", adVarChar, 30
    rsTmp.Fields.Append "结果", adVarChar, 50
    '20090926--
    rsTmp.Fields.Append "类型", adVarChar, 50       '记录是入出转,手术出院,还是未记说明,上标说明
    rsTmp.Fields.Append "打印列", adVarChar, 30
    rsTmp.Fields.Append "坐标", adVarChar, 30
    rsTmp.Fields.Append "高度", adVarChar, 30       '未记说明及上标说明不用管高度
    rsTmp.Fields.Append "字体大小", adVarChar, 50
    '----------
    rsTmp.Open
    
    For lngCol = 0 To 41
    
        lngX = lngTmpX + lngCol * HOUR_STEP_Twips + HOUR_STEP_Twips / 2
        lngY = (lngTmpY + 3 * H_9pt / 2)
        lngY1 = (lngTmpY + 3 * H_9pt / 2) + 35 * 3 * H_9pt / 2                          '得到35度的坐标
        dblHeight = (lngTmpY + 3 * H_9pt / 2) + 10 * 3 * H_9pt / 2 - H_9pt / 2          '得到40度的坐标,好像是按格子数来计算
        dblHeight = dblHeight - lngY
        
        If strArrValueInOut(lngCol) <> "" Then
        
            strComment = strArrValueInOut(lngCol)
            aryTmp = Split(strComment, ";")
            
            For i = 0 To UBound(aryTmp)
                If Trim(aryTmp(i)) <> "" Then
                    rsTmp.AddNew
                    rsTmp.Fields("类型").Value = 1                              '打印中未区分入出转手术等的类型(不含未记说明,上标说明)
                    rsTmp.Fields("坐标").Value = lngX & ";" & lngY
                    rsTmp.Fields("列号").Value = lngCol
                    rsTmp.Fields("时间").Value = Split(aryTmp(i), "'")(1)
                    rsTmp.Fields("结果").Value = Split(aryTmp(i), "'")(0)
                End If
            Next
            rsTmp.Sort = "时间"
        End If
        
        '说明
        strComment = Split(strArrValueComment(0), "'")(lngCol)
        If strComment <> "" Then
            aryTmp = Split(strComment, ";")
            
            '上标
            If aryTmp(0) <> "" Then
                rsTmp.AddNew
                rsTmp.Fields("类型").Value = 2                              '指未记说明,上标说明
                rsTmp.Fields("坐标").Value = lngX & ";" & lngY
                rsTmp.Fields("列号").Value = lngCol
                rsTmp.Fields("时间").Value = ""
                rsTmp.Fields("结果").Value = SimplifyString(aryTmp(0))
            End If
            
            '下标
            intCharNumber = 0
            If aryTmp(1) <> "" Then
                For i = 1 To Len(aryTmp(1))
                    If lngY1 < lngYMax Then
                        strTmp = Mid(aryTmp(1), i, 1)
                        
                        If Asc(strTmp) < 0 Then
                            If intCharNumber Mod 2 = 1 Then lngY = lngY + ROWHEIGHT * 2.5
                        End If
                    
                        Call DrawRotateText(objDraw, lngX - objDraw.TextWidth(strTmp) / 2 + 15, lngY1 + 15, strTmp, vbRed)

                        If Asc(strTmp) < 0 Then
                            lngY1 = lngY1 + ROWHEIGHT * 5
                            intCharNumber = 0
                        Else
                            lngY1 = lngY1 + ROWHEIGHT * 2.5
                            intCharNumber = intCharNumber + 1
                        End If
                        
                    End If
                Next
            End If
        End If
    Next
    
    '输出入出转手术,以及未记说明,上标说明
    Call OutputNote(objDraw, dblHeight, rsTmp, lngTmpX, lngTmpY)
    
    Dim int体温 As Integer
    Dim int脉搏 As Integer
    
    If Trim(strIndex) = "" Then
    
        '按项目个数循环
        
        For intItemDrawCount = 0 To intItemCount - 1
        
            intPointCount = -1
            lngTopRow = CInt(Split(strArrItemInfo(intItemDrawCount), "'")(1))
            lngStep = Val(Split(strArrItemInfo(intItemDrawCount), "'")(5))
            
            lngMax = CLng(Split(strArrItemInfo(intItemDrawCount), "'")(3))
            lngMin = CLng(Split(strArrItemInfo(intItemDrawCount), "'")(4))

            lng项目序号 = CLng(Split(strArrItemInfo(intItemDrawCount), "'")(7))
            Select Case lng项目序号
            Case 1
                int体温 = intItemDrawCount
            Case 2
                int脉搏 = intItemDrawCount
            End Select
            objDraw.ForeColor = lngColor
            
            'Erase lngArrXY
            '初始化以便在第一次时不进行连线
            
            lngUpIndex = -1
            
            '循环画数据
            For lngCol = 0 To 7 * 6 - 1 '纵
                
                lngColor = CLng(Split(strArrItemInfo(intItemDrawCount), "'")(0))
                strTag = Trim(Split(strArrItemInfo(intItemDrawCount), "'")(6))

                X = lngTmpX + (lngCol + 1) * HOUR_STEP_Twips - HOUR_STEP_Twips / 2
                    
                '1、求出当前值比例值（(当前值/单位值+最高行)*(H_9pt+H_9pt\2)），此时为相对Y坐标位置值
                '   并将此时的坐标保存在变量里
                If Val(Split(strArrValue(intItemDrawCount), "'")(lngCol)) = -2 Then
                    Y = lngTmpY
                    Y2 = lngTmpY
                    lngArrXY(lngCol, 0) = -1
                    lngArrXY(lngCol, 1) = X
                    lngArrXY(lngCol, 2) = Y
                    lngArrXY(lngCol, 3) = Y2
                    lngArrXY(lngCol, 4) = -1
                    lngArrXY(lngCol, 5) = 0
                Else
                    Y = lngTmpY
                    Y2 = lngTmpY
                    
                    aryData = Split(Split(strArrValue(intItemDrawCount), "'")(lngCol), ",")
                    
                    For i = 0 To UBound(aryData)
                        lngValue = (lngTopRow - 1) * lngStep + lngMax
                        
                        '超过当前项目的最小值最大值时,直接取最小值或最大值
                        lngArrXY(lngCol, 5) = IIf(Left(aryData(i), InStr(aryData(i), ";") - 1) = "不升", 1, 0)
                        
                        Select Case Val(Left(aryData(i), InStr(aryData(i), ";") - 1))
                        Case Is < lngMin
                            aryData(i) = lngMin & ";" & Mid(aryData(i), InStr(aryData(i), ";") + 1)
                        Case Is > lngMax
                            aryData(i) = lngMax & ";" & Mid(aryData(i), InStr(aryData(i), ";") + 1)
                        End Select
                        
                        If InStr(aryData(i), ";") > 0 Then
                            lngValue1 = lngValue - Val(Left(aryData(i), InStr(aryData(i), ";") - 1))
                        Else
                            lngValue1 = lngValue - Val(aryData(i))
                        End If

                        If i = 0 Then
                            Y = lngTmpY + (lngValue1 / lngStep) * (H_9pt + H_9pt \ 2)
                            lngArrXY(lngCol, 0) = 2
                        Else
                            Y2 = lngTmpY + (lngValue1 / lngStep) * (H_9pt + H_9pt \ 2)
                            lngArrXY(lngCol, 4) = 2
                        End If
                        
                        '因为第3个可能存的是文本（说明体温项目的部位）
                        If i = 1 Then Exit For
                    Next
'                    lngArrXY(lngCol, 0) = 2
                    lngArrXY(lngCol, 1) = X
                    lngArrXY(lngCol, 2) = Y
                    lngArrXY(lngCol, 3) = Y2
                    
                    '2、在X-W_9pt\2,Y-H_9pt\2 处画一个记录符
                    lngX = X - W_9pt \ 3
                    lngY = Y - H_9pt \ 3

                    lngX = X
                    lngY = Y
                    
                    Select Case lng项目序号
                    Case 1
                        If InStr(aryData(0), ";") > 0 Then
                            aryTmp = Split(aryData(0), ";")
                            Select Case aryTmp(5)
                            Case "口温"
                                strTag = mstrChar(0)
                            Case "腋温"
                                strTag = mstrChar(1)
                            Case "肛温"
                                strTag = mstrChar(2)
                            Case Else
                                strTag = mstrChar(1)
                            End Select
                            
                            X1 = lngX
                            Y1 = lngY
                            Select Case Val(aryTmp(0))
                            Case Is <= lngMin
                                strTag = ""
                                lngArrXY(lngCol, 5) = 1
                                Call DrawText(objDraw, X1 - 90, Y1, "不")
                                Call DrawText(objDraw, X1 - 90, Y1 + 180, "升")
'                                strTag = "·"
'                                Call DrawLine(objDraw, X1, Y1, X1, Y1 + 200, lngColor, , , True)
                            Case Is >= lngMax
                                strTag = "·"
                                Call DrawLine(objDraw, X1, Y1, X1, Y1 - 200, lngColor, , , True)
                            End Select
                            
                            If Val(aryTmp(2)) = 1 Then
                                '复试合格
                                Call DrawText(objDraw, X1 - 50, Y1 - 250, "√", vbRed)
                            End If
                            
                            mpt体温(lngCol).X = X1
                            mpt体温(lngCol).Y = Y1
            
                        End If
                    Case 2
                        '如果为空表示缺省字符;否则赋为空,在绘图时发现为空,则取资源文件中的位图
                        aryTmp = Split(aryData(0), ";")
                        strTag = IIf(aryTmp(5) = "起搏器", "", mstrPulse)
                        
                        mpt脉搏(lngCol).X = lngX
                        mpt脉搏(lngCol).Y = lngY
                        
                        X1 = lngX
                        Y1 = lngY
                        If InStr(aryData(0), ";") > 0 Then
                            aryTmp = Split(aryData(0), ";")
                            Select Case Val(aryTmp(0))
                            Case Is <= lngMin
                                Call DrawLine(objDraw, X1, Y1, X1, Y1 + 200, lngColor, , , True)
                            Case Is >= lngMax
                                Call DrawLine(objDraw, X1, Y1, X1, Y1 - 200, lngColor, , , True)
                            End Select
                        End If
                        
                    Case -1
                        mpt心率(lngCol).X = lngX
                        mpt心率(lngCol).Y = lngY
                        
                        X1 = lngX
                        Y1 = lngY
                        
                        If InStr(aryData(0), ";") > 0 Then
                            aryTmp = Split(aryData(0), ";")
                            Select Case Val(aryTmp(0))
                            Case Is <= lngMin
                                Call DrawLine(objDraw, X1, Y1, X1, Y1 + 200, lngColor, , , True)
                            Case Is >= lngMax
                                Call DrawLine(objDraw, X1, Y1, X1, Y1 - 200, lngColor, , , True)
                            End Select
                        End If
                        
                    Case Else
                        If lng项目序号 = 3 Then
                            aryTmp = Split(aryData(0), ";")
                            strTag = IIf(aryTmp(5) = "呼吸机", "", mstrBreath)
                        End If
                        
                        X1 = lngX
                        Y1 = lngY
                        Select Case Val(aryData(0))
                        Case Is <= lngMin
                            Call DrawLine(objDraw, X1, Y1, X1, Y1 + 200, lngColor, , , True)
                        Case Is >= lngMax
                            Call DrawLine(objDraw, X1, Y1, X1, Y1 - 200, lngColor, , , True)
                        End Select
                        
                    End Select
                    
                    If lng项目序号 = 1 Then         '体温
                        Call PointAdd(rsPoint, lngX, lngY, lng项目序号, strTag, lngColor, lngCol, CStr(aryTmp(5)))
                    ElseIf lng项目序号 = 3 Then     '呼吸
                        Call PointAdd(rsPoint, lngX, lngY, lng项目序号, strTag, lngColor, lngCol, CStr(aryTmp(5)), IIf(strTag = "", "BREATH", ""))
                    ElseIf lng项目序号 = 2 Then     '脉搏
                        Call PointAdd(rsPoint, lngX, lngY, lng项目序号, strTag, lngColor, lngCol, CStr(aryTmp(5)), IIf(strTag = "", "PACEMAKER", ""))
                    Else
                        Call PointAdd(rsPoint, lngX, lngY, lng项目序号, strTag, lngColor, lngCol, "")
                    End If

                    lngColor = CLng(Split(strArrItemInfo(intItemDrawCount), "'")(0))
                    
                    If Y2 <> lngTmpY Then
                        
                        Select Case lng项目序号
                        Case 2
                        
                            '心率与脉搏共享使用时
'                            If mint心率应用 = 2 Then
                            
                                strTag = mstr心率符号
                                mpt心率(lngCol).X = lngX
                                mpt心率(lngCol).Y = (Y2 - H_9pt \ 3)
                                lngColor = 255

                                If InStr(aryData(1), ";") > 0 Then
                                    X1 = lngX
                                    Y1 = Y2
                                    aryTmp = Split(aryData(1), ";")
                                    Select Case Val(aryTmp(0))
                                    Case Is <= lngMin
                                        Call DrawLine(objDraw, X1, Y1, X1, Y1 + 200, lngColor, , , True)
                                    Case Is >= lngMax
                                        Call DrawLine(objDraw, X1, Y1, X1, Y1 - 200, lngColor, , , True)
                                    End Select
                                End If
                                
'                            End If
                        Case 1
                            strTag = "○"
                            lngColor = 255
                        End Select
                                                                        
                        If lng项目序号 = 1 Then
                            Call PointAdd(rsPoint, lngX, Y2 - H_9pt \ 3, lng项目序号, strTag, lngColor, lngCol, CStr(aryTmp(5)))
                        Else
                            Call PointAdd(rsPoint, lngX, Y2 - H_9pt \ 3, lng项目序号, strTag, lngColor, lngCol, "")
                        End If
                    
                    End If

                    '3、找到前一个坐标的进行连线。如果当一个坐标的（0）元素为0时表示不连
                    '--------------------------------------------------------------------------------------------------
                    If lngCol > 0 Then
                        lngColor = CLng(Split(strArrItemInfo(intItemDrawCount), "'")(0))
                        If lngUpIndex > -1 Then
                            If lngArrXY(lngUpIndex, 0) = 2 Then
                                If lngArrXY(lngUpIndex, 5) <> 1 And lngArrXY(lngCol, 5) <> 1 Then
                                    Call DrawLine(objDraw, lngArrXY(lngUpIndex, 1), lngArrXY(lngUpIndex, 2), X, Y, lngColor)
                                End If
                                
                                If Y2 <> lngTmpY Then
                                    Select Case lng项目序号
                                    Case 1          '体温项目
                                        Call DrawLine(objDraw, X, Y, X, Y2, 255, IIf(Y2 < Y, 0, 2), , True)
                                    Case 2
                                        If lngArrXY(lngUpIndex, 4) = 2 Then
                                            Call DrawLine(objDraw, lngArrXY(lngUpIndex, 1), lngArrXY(lngUpIndex, 3), X, Y2, lngColor)
                                        End If
                                    End Select
                                    
                                End If
                            Else
                                '虽然断点，如果当天有两个点，则先画出来
                                '--------------------------------------------------------------------------------------
                                If Y2 <> lngTmpY Then
                                    Select Case lng项目序号
                                    Case 1
                                        Call DrawLine(objDraw, X, Y, X, Y2, 255, IIf(Y2 < Y, 0, 2), , True)
                                    End Select
                                End If
                                
                                '找出当天及昨天上一个有效点,只有隔天无数据才不连线
                                '--------------------------------------------------------------------------------------
                                lngUpIndex = lngUpIndex - 1
                                blnStop = False
                                Do While lngUpIndex >= (lngCol \ 6) * 6 - 6 And lngUpIndex >= 0
                                    'If lngCol = 9 Then Stop
                                    If blnStop = False Then
                                        If Split(strArrValueComment(0), "'")(lngUpIndex + 1) <> "" Then
                                            aryTmp = Split(Split(strArrValueComment(0), "'")(lngUpIndex + 1), ";")
                                        Else
                                            aryTmp = Split(";;", ";")
                                        End If
                                        blnStop = (Val(aryTmp(2)) = 1)
                                    End If
                                    If lngArrXY(lngUpIndex, 5) = 1 Or lngArrXY(lngCol, 5) = 1 Then Exit Do
                                    If blnStop Then
                                        Exit Do
                                    End If

                                    If lngArrXY(lngUpIndex, 0) = 2 And lngArrXY(lngUpIndex, 5) <> 1 And lngArrXY(lngCol, 5) <> 1 Then
                                        If blnStop = False Then
                                            blnStop = False
                                            lngColor = CLng(Split(strArrItemInfo(intItemDrawCount), "'")(0))
                                            
                                            Call DrawLine(objDraw, lngArrXY(lngUpIndex, 1), lngArrXY(lngUpIndex, 2), X, Y, lngColor)
                                            
                                            If Y2 <> lngTmpY Then
                                                Select Case lng项目序号
                                                Case 1
                                                    Call DrawLine(objDraw, X, Y, X, Y2, 255, IIf(Y2 < Y, 0, 2), , True)
                                                Case 2                  '脉搏附加心率情况
                                                    If lngArrXY(lngUpIndex, 4) = 2 Then
                                                        Call DrawLine(objDraw, lngArrXY(lngUpIndex, 1), lngArrXY(lngUpIndex, 3), X, Y2, lngColor)
                                                    End If
                                                End Select
                                            End If
                                        End If
                                        Exit Do
                                    End If
                                    lngUpIndex = lngUpIndex - 1
                                Loop
                                
                                '--------------------------------------------------------------------------------------
                                
                            End If
                        End If
                    ElseIf Y2 <> lngTmpY Then
                        Select Case lng项目序号
                        Case 1
                            Call DrawLine(objDraw, X, Y, X, Y2, 255, IIf(Y2 < Y, 0, 2), , True)
                        End Select
                    End If
                End If
                lngUpIndex = lngCol
            Next
        Next
        
        '根据脉搏和心率坐标形成多边形，并进行连线和填充
        '--------------------------------------------------------------------------------------------------------------
        Call DrawPoly(objDraw, mpt脉搏, mpt心率, mbyt脉搏)
    End If
    
    '画点的字符或图形
    '--------------------------------------------------------------------------------------------------------------
    Call DrawPoint(objDraw, rsPoint, int体温)
    
    X = lngTmpX - (UBound(strArrItemInfo) + 1) * (W_9pt + W_9pt \ 2)
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function InitOffset(ByRef rsOffset As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Set rsOffset = New ADODB.Recordset
    With rsOffset
        .Fields.Append "Line", adInteger
        .Fields.Append "Column", adInteger
        .Fields.Append "Offset", adVarChar, 30
        .Open
    End With
    
    InitOffset = True
    
End Function


Public Function IsCenterValue(ByRef rsOffset As ADODB.Recordset, ByVal lngLine As Long, ByVal lngColumn As Long, ByVal dtDot As Date, ByVal dtCenter As Date) As Boolean
    '******************************************************************************************************************
    '功能：如果同一列中有多个值，则取最靠中点的值作为本列的值
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngSecond As Long
        
    On Error GoTo errHand
    
    IsCenterValue = False
        
    lngSecond = Abs(DateDiff("s", dtCenter, dtDot))
    
    rsOffset.Filter = ""
    rsOffset.Filter = "Line=" & lngLine & " And Column=" & lngColumn
    
    If rsOffset.RecordCount = 0 Then
    
        rsOffset.AddNew
        rsOffset("Line").Value = lngLine
        rsOffset("Column").Value = lngColumn
        rsOffset("Offset").Value = lngSecond
        
        IsCenterValue = True
        
    ElseIf Val(rsOffset("Offset").Value) >= lngSecond Then
    
        rsOffset("Offset").Value = lngSecond
        IsCenterValue = True
        
    End If
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    
End Function

Public Function PointInit(ByRef rsPoint As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Set rsPoint = New ADODB.Recordset
    With rsPoint
        .Fields.Append "X", adVarChar, 30
        .Fields.Append "Y", adVarChar, 30
        .Fields.Append "Line", adVarChar, 30
        .Fields.Append "Column", adInteger
        .Fields.Append "符号", adVarChar, 30
        .Fields.Append "颜色", adVarChar, 30
        .Fields.Append "标志", adInteger
        .Fields.Append "图形", adVarChar, 200
        .Fields.Append "重叠标识", adVarChar, 50
        .Fields.Append "项目序号", adBigInt
        .Fields.Append "体温部位", adVarChar, 30
        .Open
    End With
    
    PointInit = True
    
End Function

Public Function PointAdd(ByRef rsPoint As ADODB.Recordset, ByVal X1 As Single, ByVal Y1 As Single, ByVal lngLine As Long, ByVal strChar As String, ByVal lngColor As Long, ByVal lngColumn As Long, ByVal str体温部位 As String, Optional ByVal str图形 As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    With rsPoint
        .AddNew
        .Fields("X").Value = X1
        .Fields("Y").Value = Y1
        .Fields("Line").Value = lngLine
        .Fields("Column").Value = lngColumn
        .Fields("符号").Value = strChar
        .Fields("颜色").Value = lngColor
        .Fields("标志").Value = 0
        .Fields("重叠标识").Value = ""
        .Fields("图形").Value = str图形
        If str体温部位 = "" And lngLine = 1 Then str体温部位 = "腋温"
        .Fields("体温部位").Value = str体温部位
    End With
    
    PointAdd = True
End Function

Private Function PointCalc(ByRef rsPoint As ADODB.Recordset, ByVal int体温 As Integer) As Boolean
    '******************************************************************************************************************
    '功能：获取重叠点序列
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strTmp As String
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim rsConverPoint As ADODB.Recordset
    
    On Error GoTo errHand
    
    Call GetConverPoint(rsPoint, rsConverPoint, int体温)
    
    If Not rsConverPoint Is Nothing Then
        If rsConverPoint.RecordCount > 0 Then
            rsConverPoint.MoveFirst
            Do While Not rsConverPoint.EOF
                
                strTmp = rsConverPoint("Lines").Value
                If InStr("," & strTmp & ",", ",1,") > 0 Then

                    strTmp = "0," & strTmp & ",0"
                    strTmp = Replace(strTmp, ",1,", ",")
                    
                    strSQL = "Select a.序号,a.标记符号,a.标记颜色 " & _
                            "From 体温重叠标记 a,(Select 上级序号, Count(1) As 个数 From 体温重叠标记 Where 项目序号 In (" & strTmp & ") Or (项目序号=1 And Nvl(体温部位,'腋温')=[2]) Group By 上级序号) b " & _
                            "Where a.序号 = b.上级序号 And b.个数 = a.重叠数目 And a.重叠数目=[1]"
                    
                Else
                
                    strSQL = "Select a.序号,a.标记符号,a.标记颜色 " & _
                            "From 体温重叠标记 a,(Select 上级序号, Count(1) As 个数 From 体温重叠标记 Where 项目序号 In (" & strTmp & ") Group By 上级序号) b " & _
                            "Where a.序号 = b.上级序号 And b.个数 = a.重叠数目 And a.重叠数目=[1]"
                        
                End If
                Set rs = zlDatabase.OpenSQLRecord(strSQL, "usrBodyEditor", Val(rsConverPoint("重叠数目").Value), CStr(rsConverPoint("体温部位").Value))

                If rs.BOF = False Then
                    rsPoint.Filter = ""
                    rsPoint.Filter = "重叠标识='" & rsConverPoint("重叠标识").Value & "'"
                    If rsPoint.RecordCount > 0 Then
                        rsPoint.MoveFirst
                        rsPoint("标志").Value = 1
                        rsPoint("符号").Value = zlCommFun.NVL(rs("标记符号").Value)
                        rsPoint("颜色").Value = zlCommFun.NVL(rs("标记颜色").Value, 0)
                        
                        '读取标记图形并显示
                        strTmp = App.Path & "\ConverPoint" & zlCommFun.NVL(rs("序号").Value, 0) & ".tmp"
                        If Dir(strTmp) <> "" Then Kill strTmp
                        strTmp = zlBlobRead(9, zlCommFun.NVL(rs("序号").Value, 0), strTmp)
                        If Dir(strTmp) <> "" And strTmp <> "" Then
                            
                            rsPoint("图形").Value = strTmp
                            
                        End If

                    End If
                End If
                rsConverPoint.MoveNext
            Loop
        End If
    End If
    
    PointCalc = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function

Public Function DrawPoint(objDraw As Object, ByRef rsPoint As ADODB.Recordset, ByVal int体温 As Integer) As Boolean
    '******************************************************************************************************************
    '功能：画重叠点字符或图形
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strChar As String
    Dim lngColor As Long
    Dim X1 As Long
    Dim Y1 As Long
    Dim strTmp As String
    
    On Error GoTo errHand
    
    '计算重点
    Call PointCalc(rsPoint, int体温)
    
    '输出点
    rsPoint.Filter = ""
    If rsPoint.RecordCount > 0 Then
        rsPoint.MoveFirst

        Do While Not rsPoint.EOF

            strChar = rsPoint("符号").Value
            lngColor = Val(rsPoint("颜色").Value)

            X1 = Val(rsPoint("X").Value) - objDraw.TextWidth(strChar) / 2
            Y1 = Val(rsPoint("Y").Value) - objDraw.TextHeight(strChar) / 2

            If X1 > 0 Or Y1 > 0 Then
                Select Case rsPoint("标志").Value
                Case 0                              '独立的点
                    '独立的点也可能是图形，如呼吸，当呼吸机辅助呼吸时为图形，而该图形是在资源文件里的，所以需单独处理
                    If rsPoint!图形 <> "" Then  '保存的是资源文件中的ID
                        X1 = Val(rsPoint("X").Value)
                        Y1 = Val(rsPoint("Y").Value)
            
                        Call DrawPicture(objDraw, rsPoint!图形, X1 - 90, Y1 - 90, X1 + 90, Y1 + 90, True)
                    Else
                        Call DrawText(objDraw, X1, Y1, strChar, lngColor)
                    End If
                Case 1                              '重叠的点
                    
                    If strChar <> "" Then
                        Call DrawText(objDraw, X1, Y1, strChar, lngColor)
                    Else
                        strTmp = rsPoint("图形").Value
                        If strTmp <> "" And Dir(strTmp) <> "" Then
                            
                            X1 = Val(rsPoint("X").Value)
                            Y1 = Val(rsPoint("Y").Value)
                
                            Call DrawPicture(objDraw, strTmp, X1 - 90, Y1 - 90, X1 + 90, Y1 + 90)
                        End If
                    End If
                    
                End Select
            End If

            rsPoint.MoveNext
        Loop
    End If
    
    DrawPoint = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function

'Public Function DrawPoly(objDraw As Object, pt脉搏() As POINTAPI, pt心率() As POINTAPI) As Boolean
'    '******************************************************************************************************************
'    '功能：根据脉搏和心率坐标形成多边形，并进行连线和填充
'    '参数：
'    '返回：
'    '******************************************************************************************************************
'    Dim intCount As Integer
'    Dim intStart As Integer
'    Dim intEnd As Integer
'
'    For intCount = 0 To 41
'        If pt脉搏(intCount).X > 0 Then
'            If pt心率(intCount).X = 0 Then
'                If intEnd > 0 Then
'                    intEnd = intCount
'                    Call DrawFillPoly(objDraw, intStart, intEnd, pt脉搏, pt心率)
'                    intEnd = 0
'                    intStart = intCount
'                Else
'                    intStart = intCount
'                End If
'            Else
'                intEnd = intCount
'                If intStart = 0 Then intStart = intCount
'            End If
'        Else
'            If intEnd > 0 And intStart > 0 Then
'                Call DrawFillPoly(objDraw, intStart, intEnd, pt脉搏, pt心率)
'            End If
'            intStart = 0
'            intEnd = 0
'        End If
'    Next
'    If intEnd > 0 Then
'        '有一个多边形了
'        Call DrawFillPoly(objDraw, intStart, intEnd, pt脉搏, pt心率)
'    End If
'
'    DrawPoly = True
'
'End Function

Public Function DrawPoly(objDraw As Object, pt脉搏() As POINTAPI, pt心率() As POINTAPI, byt脉搏() As Byte) As Boolean
    '******************************************************************************************************************
    '功能：根据脉搏和心率坐标形成多边形，并进行连线和填充
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intCount As Integer
    Dim intStart As Integer
    Dim intEnd As Integer
    Dim intDay As Integer
    
    intStart = -1: intEnd = -1
    For intCount = 0 To 41
        If pt脉搏(intCount).X > 0 Then
            If pt心率(intCount).X = 0 Then
                If intEnd >= 0 Then
                    intEnd = intCount
                    Call DrawFillPoly(objDraw, intStart, intEnd, pt脉搏, pt心率)
                    intEnd = -1
                    intStart = intCount
                Else
                    intStart = intCount
                End If
            Else
                intEnd = intCount
                If intStart = -1 Then intStart = intCount
            End If
        Else
            '如果没有数据检查 是否有未记说明或者脉搏数据两点之间数据超过一天
            '有未记说明画脉搏短轴
            '数据超过一天画脉搏短轴

            intDay = intCount - (IIf(intEnd = -1, intStart, intEnd) + ((intCount + 1) Mod 6))

            If byt脉搏(intCount) = 1 Or intDay > 6 Then
                If intEnd >= 0 And intStart >= 0 Then
                    Call DrawFillPoly(objDraw, intStart, intEnd, pt脉搏, pt心率)
                End If
                intStart = -1
                intEnd = -1
            End If
        End If
    Next
    If intEnd >= 0 Then
        '有一个多边形了
        Call DrawFillPoly(objDraw, intStart, intEnd, pt脉搏, pt心率)
    End If

    DrawPoly = True

End Function

Private Function GetConverPoint(ByVal rsPoint As ADODB.Recordset, ByRef rsConverPoint As ADODB.Recordset, ByVal int体温 As Integer) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim X0 As Long
    Dim Y0 As Long
    Dim lngLine As Long
    Dim intMax As Integer

    On Error GoTo errHand

    If rsPoint.RecordCount = 0 Then Exit Function

    Set rsConverPoint = New ADODB.Recordset
    With rsConverPoint
        .Fields.Append "重叠标识", adVarChar, 30
        .Fields.Append "Lines", adVarChar, 30
        .Fields.Append "重叠数目", adInteger
        .Fields.Append "体温部位", adVarChar, 30
        .Open
    End With


    '------------------------------------------------------------------------------------------------------------------
    rsPoint.Sort = "X,Y"
    rsPoint.MoveFirst
    Do While Not rsPoint.EOF
        If rsPoint("X").Value = X0 And rsPoint("Y").Value = Y0 Then
            rsConverPoint.Filter = ""
            rsConverPoint.Filter = "重叠标识='" & X0 & "," & Y0 & "'"
            If rsConverPoint.RecordCount = 0 Then
                rsConverPoint.AddNew
                rsConverPoint("重叠标识").Value = X0 & "," & Y0
                rsConverPoint("Lines").Value = ""
                rsConverPoint("重叠数目").Value = 0
            End If
            If rsConverPoint("Lines").Value = "" Then

                rsPoint.MovePrevious
                rsPoint("重叠标识").Value = X0 & "," & Y0
                rsPoint.MoveNext
                rsConverPoint("重叠数目").Value = 2
                rsConverPoint("Lines").Value = lngLine & "," & rsPoint("Line").Value
            Else
                rsConverPoint("Lines").Value = rsConverPoint("Lines").Value & "," & rsPoint("Line").Value
                rsConverPoint("重叠数目").Value = rsConverPoint("重叠数目").Value + 1
            End If

            If rsPoint("体温部位").Value <> "" Then
                '目前只有体温的部位重叠设置,所以此处仍保存原样
                If InStr(1, rsPoint("体温部位").Value, ";") <> 0 Then
                    '打印中提取的部位是各个曲线的,没有合并在一起
                    rsConverPoint("体温部位").Value = Split(rsPoint("体温部位").Value, ";")(int体温 + 1)
                Else
                    rsConverPoint("体温部位").Value = rsPoint("体温部位").Value
                End If
            End If

            rsPoint("重叠标识").Value = X0 & "," & Y0
            rsPoint("标志").Value = 2

            GetConverPoint = True

        End If

        lngLine = rsPoint("Line").Value
        X0 = rsPoint("X").Value
        Y0 = rsPoint("Y").Value

        rsPoint.MoveNext
    Loop

    rsConverPoint.Filter = ""

    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
'
'Private Function GetConverPoint(ByVal rsPoint As ADODB.Recordset, ByRef rsConverPoint As ADODB.Recordset) As Boolean
'    '******************************************************************************************************************
'    '功能：
'    '参数：
'    '返回：
'    '******************************************************************************************************************
'    Dim X0 As Long
'    Dim Y0 As Long
'    Dim Y1 As Long          '临时使用,用于判断是否重合
'    Dim lngLine As Long
'    Const dblError As Long = 15  '误差值,表示允许+/-15的误差
'
'    On Error GoTo errHand
'    If rsPoint.RecordCount = 0 Then Exit Function
'
'    Set rsConverPoint = New ADODB.Recordset
'    With rsConverPoint
'        .Fields.Append "组号", adVarChar, 30            '每一个组号表示一组重合的点
'        .Fields.Append "Lines", adVarChar, 30           '以前是保存所有重合的曲线序号,现改为保存当前点的序号
'        .Fields.Append "坐标", adInteger                '允许指定范围的误差,以此进行后序点的判断
'        .Fields.Append "实际坐标", adInteger            '实际坐标
'        .Fields.Append "体温部位", adVarChar, 30        '体温部位/呼吸方式/起搏器
'        .Open
'    End With
'
'
'    '------------------------------------------------------------------------------------------------------------------
'    rsPoint.Sort = "X,Y"
'    rsPoint.MoveFirst
'    Do While Not rsPoint.EOF
'
'        '误差设定为30
'        If rsPoint("X").Value = X0 And Abs(rsPoint("Y").Value - Y0) <= dblError Then
'            rsConverPoint.Filter = ""
'            Y1 = IIf(rsPoint!y >= Y0, Y0, rsPoint!y) + dblError
'            rsConverPoint.Filter = "坐标='" & X0 & "," & Y0 & "'"
'            If rsConverPoint.RecordCount = 0 Then
'                rsConverPoint.AddNew
'                rsConverPoint("重叠标识").Value = X0 & "," & Y0
'                rsConverPoint("Lines").Value = ""
'                rsConverPoint("重叠数目").Value = 0
'            End If
'            If rsConverPoint("Lines").Value = "" Then
'
'                rsPoint.MovePrevious
'                rsPoint("重叠标识").Value = X0 & "," & Y0
'                rsPoint.MoveNext
'                rsConverPoint("重叠数目").Value = 2
'                rsConverPoint("Lines").Value = lngLine & "," & rsPoint("Line").Value
'            Else
'                rsConverPoint("Lines").Value = rsConverPoint("Lines").Value & "," & rsPoint("Line").Value
'                rsConverPoint("重叠数目").Value = rsConverPoint("重叠数目").Value + 1
'            End If
'
'            If rsPoint("体温部位").Value <> "" Then
'                rsConverPoint("体温部位").Value = rsPoint("体温部位").Value
'            End If
'
'            rsPoint("重叠标识").Value = X0 & "," & Y0
'            rsPoint("标志").Value = 2
'
'            GetConverPoint = True
'
'        End If
'
'        lngLine = rsPoint("Line").Value
'        X0 = rsPoint("X").Value
'        Y0 = rsPoint("Y").Value
'
'        rsPoint.MoveNext
'    Loop
'
'    rsConverPoint.Filter = ""
'
'    Exit Function
'    '------------------------------------------------------------------------------------------------------------------
'errHand:
'    If ErrCenter = 1 Then
'        Resume
'    End If
'End Function

Public Function GetDataFromHis(ByVal lng病人id As Long, ByVal lng主页id As Long, ByVal lng婴儿 As Long, ByVal dtFrom As Date, ByVal dtTo As Date, Optional ByVal bytMode As Byte = 1) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：从医嘱记录提取手术、分娩数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strSQL As String
    Dim lng诊疗项目id As Long
    Dim rs As New ADODB.Recordset
    
    
    Select Case bytMode
    '------------------------------------------------------------------------------------------------------------------
    Case 1              '从医嘱记录提取手术、分娩数据
        
'        dtFrom = dtFrom - 14
        
        strSQL = _
                "Select 执行时间,内容,次弟" & vbNewLine & _
                "From (Select 执行时间,内容, Rownum As 次弟" & vbNewLine & _
                "       From (Select Distinct C.执行时间,'手术' As 内容 " & vbNewLine & _
                "              From 病人医嘱记录 A, 诊疗项目目录 B, 病人医嘱执行 C" & vbNewLine & _
                "              Where A.病人id = [1] And A.主页id = [2] And Nvl(A.婴儿, 0) = [3] And A.医嘱期效 = 1 And A.诊疗项目id = B.ID And" & vbNewLine & _
                "                    A.诊疗类别 = 'F' And A.医嘱状态 = 8 And C.医嘱id = A.ID And C.执行时间 < =[5] " & vbNewLine & _
                "               Union All Select a.出生时间 As 执行时间,'分娩' As 内容 From 病人新生儿记录 a Where a.病人id=[1] And a.主页id=[2] And a.出生时间 Is Not Null And RowNum<2) " & _
                "       Order By 执行时间)" & vbNewLine & _
                "Where 执行时间 >= [4] And 次弟 <= 12 " & vbNewLine & _
                "Order By 执行时间 "
                
        Set GetDataFromHis = zlDatabase.OpenSQLRecord(strSQL, "体温单", lng病人id, lng主页id, lng婴儿, dtFrom, dtTo)
    '------------------------------------------------------------------------------------------------------------------
    Case 2              '入出转标志(入院,出院,转科,换床)
        
        
        '1-入院；2-入科；3-转科；4-换床；5-床位等级变动；6-护理等级变动；7-经治医师改变；8-责任护士改变,9-留观病人转住院,10-病人预出院,11-主治医师变动,12-主任医师变动,13-病情变动
        
        '有的医院死亡类型有多个
'        strSQL = "Select ID From 诊疗项目目录 Where 类别='Z' And 操作类型='11' "
'        Set rs = zlDatabase.OpenSQLRecord(strSQL, "体温单")
'        If rs.BOF = False Then lng诊疗项目id = zlCommFun.NVL(rs("ID").Value)
        
        strSQL = _
                "   Select b.名称 As 科室,开始时间 As 时间, Decode(开始原因, 2,'入科',3, '转入',4,'换床'||Decode(床号,Null,'','('||床号||')')) As 内容,Decode(开始原因,2,9,3,6,4,7) As 行号 " & vbNewLine & _
                "   From 病人变动记录 A,部门表 b" & vbNewLine & _
                "   Where b.id(+)=a.科室id and a.开始原因 In (2,3,4) And A.病人id = [1] And A.主页id = [2]  And [3]=0 And A.开始时间 Between [4] And [5] " & vbNewLine & _
                "   Union All" & vbNewLine & _
                "   Select '' As 科室,时间,内容,行号 From (Select * From (Select 开始时间 As 时间, '入院' As 内容,5 As 行号 " & vbNewLine & _
                "   From 病人变动记录 A" & vbNewLine & _
                "   Where a.开始原因=1 And A.病人id = [1] And A.主页id = [2] And [3]=0 Order By a.开始时间) Where RowNum=1) Where 时间 Between [4] And [5] " & vbNewLine & _
                "   Union All" & vbNewLine & _
                "   Select '' As 科室,Nvl(b.开始执行时间,a.出院日期) As 时间, Decode(出院方式, '正常', '出院', 出院方式) As 内容,8 As 行号 " & vbNewLine & _
                "   From 病案主页 A,(Select x.病人id,x.主页id,Max(x.开始执行时间) As 开始执行时间 From 病人医嘱记录 x,诊疗项目目录 z Where x.病人id=[1] And x.主页id=[2] " & vbNewLine & _
                "   And x.诊疗项目id+0=z.ID And x.医嘱状态 in (3,8) And z.类别='Z' And z.操作类型='11' Group By x.病人id,x.主页id) B " & vbNewLine & _
                "   Where A.病人id = [1] And A.主页id = [2] And A.出院日期 Between [4] And [5] And a.病人id=b.病人id(+) And a.主页id=b.主页id(+) "
                
        
        strSQL = "Select distinct * From (" & strSQL & ") Order By 时间,行号 "
        Set GetDataFromHis = zlDatabase.OpenSQLRecord(strSQL, "体温单", lng病人id, lng主页id, lng婴儿, dtFrom, dtTo)
    '------------------------------------------------------------------------------------------------------------------
    Case 3              '从新生儿记录中提出生/分娩日期
        
        strSQL = "Select '' As 科室,a.出生时间 As 时间,'出生' As 内容,13 As 行号 From 病人新生儿记录 a Where a.病人id=[1] And a.主页id=[2] And a.序号=[3] And a.出生时间 Is Not Null"
        Set GetDataFromHis = zlDatabase.OpenSQLRecord(strSQL, "体温单", lng病人id, lng主页id, lng婴儿)
        
    End Select
    
End Function

'todo:以下是体温与麻醉用到的过程或函数
Public Sub DrawLine(pic As Object, ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, Optional ByVal ForeColor As Long = 0, Optional ByVal DrawStyle As Byte, Optional ByVal LineWidth As Byte = 1, Optional ByVal blnEndArrow As Boolean)
    
    '在(X1,Y1),(X2,Y2)之间使用ForeColor色画一直线
    Dim lngSaveForeColor As Long
    Dim bytSaveLineWidth As Byte
    Dim lngLoop As Long

    lngSaveForeColor = pic.ForeColor
    bytSaveLineWidth = pic.DrawWidth
    pic.ForeColor = ForeColor
    pic.DrawStyle = DrawStyle
    pic.DrawWidth = LineWidth
    pic.Line (X2, Y2)-(X1, Y1)

    If blnEndArrow Then

        If Y1 < Y2 Then
            For lngLoop = X1 - 40 To X1 + 40
                pic.Line (X2, Y2)-(lngLoop, Y2 - 75)
            Next
        Else

            For lngLoop = X1 - 40 To X1 + 40
                pic.Line (X2, Y2)-(lngLoop, Y2 + 75)
            Next

        End If
    End If

    pic.ForeColor = lngSaveForeColor
    pic.DrawWidth = bytSaveLineWidth

End Sub

Public Sub DrawText(objDraw As Object, ByVal X As Single, ByVal Y As Single, ByVal Text As String, Optional ByVal ForeColor As Long = 0)
    '在(X,Y)处输出Text文本
    Dim lngSaveForeColor As Long
    
    With objDraw
        lngSaveForeColor = .ForeColor
        .ForeColor = ForeColor
        .CurrentX = X
        .CurrentY = Y
        objDraw.FontTransparent = True
        objDraw.Print Text
        .ForeColor = lngSaveForeColor
    End With
End Sub

Public Sub DrawRotateText(objDraw As Object, ByVal X As Single, ByVal Y As Single, ByVal Text As String, Optional ByVal ForeColor As Long = 0, Optional ByVal sglScale As Single = 1.5)
    '在(X,Y)处输出Text文本
    Dim lngSaveForeColor As Long
    Dim lngLoop As Long
    Dim objFont As New clsRotateFont '旋转字体对象

    With objDraw
        lngSaveForeColor = .ForeColor
        .ForeColor = ForeColor
        objDraw.FontTransparent = True

        If Asc(Text) < 0 Then
            '全角
            .CurrentX = X
            .CurrentY = Y
            
            objDraw.Print Text
        Else
            '半角,向左转90度

            .CurrentX = X + objDraw.TextWidth("1") * sglScale
            .CurrentY = Y

            Set objFont = New clsRotateFont
            Set objFont.LogFont = New StdFont
            objFont.LogFont.Name = .FontName
            objFont.LogFont.Size = .FontSize
            objFont.Rotation = -90
            objFont.Output objDraw, .CurrentX, .CurrentY, Text

        End If

        .ForeColor = lngSaveForeColor
    End With
End Sub

Private Function DrawFillPoly(objDraw As Object, ByVal intStart As Integer, ByVal intEnd As Integer, pt脉搏() As POINTAPI, pt心率() As POINTAPI)
    '******************************************************************************************************************
    '
    '
    '
    '******************************************************************************************************************
    
    Dim intCol As Integer
    Dim ptPoly() As POINTAPI
    Dim intLoop1 As Integer
    Dim intLoop2 As Integer
    Dim intLoop As Integer
    Dim lngBrush As Long
    Dim lngRgn As Long
    Dim lngSvrColor As Long
    
    
    lngSvrColor = objDraw.ForeColor
    ReDim ptPoly(0)
    
    intLoop1 = 0
    intLoop2 = 0
    intLoop = 0
    For intCol = intStart To intEnd
        If pt脉搏(intCol).X > 0 Then
            intLoop1 = intLoop1 + 1
            ReDim Preserve ptPoly(intLoop1)
            ptPoly(intLoop1).X = pt脉搏(intCol).X / Screen.TwipsPerPixelX
            ptPoly(intLoop1).Y = pt脉搏(intCol).Y / Screen.TwipsPerPixelY
        End If
    Next
    For intCol = intEnd To intStart Step -1
        If pt心率(intCol).X > 0 Then
            intLoop2 = intLoop2 + 1
            ReDim Preserve ptPoly(intLoop2 + intLoop1)
            ptPoly(intLoop2 + intLoop1).X = pt心率(intCol).X / Screen.TwipsPerPixelX
            ptPoly(intLoop2 + intLoop1).Y = pt心率(intCol).Y / Screen.TwipsPerPixelY
        End If
    Next

    intLoop = intLoop2 + intLoop1 + 1
    ReDim Preserve ptPoly(intLoop)
    ptPoly(intLoop).X = ptPoly(1).X
    ptPoly(intLoop).Y = ptPoly(1).Y

    lngBrush = CreateHatchBrush(3, vbRed)
    lngRgn = CreatePolygonRgn(ptPoly(1), intLoop, ALTERNATE)
    FillRgn objDraw.hDC, lngRgn, lngBrush

    Call DeleteObject(lngRgn)
    Call DeleteObject(lngBrush)

    objDraw.ForeColor = 255
    Polyline objDraw.hDC, ptPoly(intLoop1), intLoop2 + 2

    objDraw.ForeColor = lngSvrColor
End Function

Private Sub DrawBodyPaper(objDraw As Object, X As Long, Y As Long, intGuageRow As Integer, Optional sngScale As Long = 1)
    '******************************************************************************************************************
    '功能:画内部表格的底纹
    '参数:intGuageRow=要画的行数
    '顺序:
    '画图顺序号:=5/9   这是画所有图形的第五步
    '******************************************************************************************************************
    
    Dim intRow As Integer, intCol As Integer
    Dim H_9pt As Long
    Dim lngTmpX As Long, lngTmpY As Long
    Dim X0 As Long, Y0 As Long
    
    objDraw.Font.Name = "宋体"
    objDraw.Font.Size = 9 * sngScale
    objDraw.Font.Bold = False
'    H_9pt = objDraw.TextHeight("字")
    H_9pt = ROWHEIGHT * 10 / 3
    
    lngTmpX = X
    lngTmpY = Y
    
    '------------------------------------------------------------------------------------------------------------------
    '画横向坐标图线
    For intRow = 0 To intGuageRow - 1
        objDraw.DrawStyle = 2
        Y0 = lngTmpY + intRow * (H_9pt + H_9pt \ 2)
        
'        If (intRow + 1) Mod 5 = 0 Then
        If intRow Mod 5 = 0 Then
            If (intRow - 1) = 24 Then
                Call DrawLine(objDraw, lngTmpX + 10, Y0 + ROWHEIGHT * 5, lngTmpX + HOUR_STEP_Twips * 6 * 7, Y0 + ROWHEIGHT * 5, RGB(200, 0, 0), , 2)
            Else
                Call DrawLine(objDraw, lngTmpX + 10, Y0 + ROWHEIGHT * 5, lngTmpX + HOUR_STEP_Twips * 6 * 7, Y0 + ROWHEIGHT * 5, COLOR.黑色, , 2)
            End If
            
        Else
            Call DrawLine(objDraw, lngTmpX + 10, Y0 + ROWHEIGHT * 5, lngTmpX + HOUR_STEP_Twips * 6 * 7, Y0 + ROWHEIGHT * 5, COLOR.黑色)
        End If
    Next
    Y = intGuageRow * (H_9pt + H_9pt \ 2) + lngTmpY  '求出最底下的一根线

    '------------------------------------------------------------------------------------------------------------------
    '画纵向坐标图线
    For intCol = 0 To 6
        objDraw.DrawStyle = 0
        X0 = lngTmpX + intCol * HOUR_STEP_Twips * 6
        objDraw.Line (X0 + HOUR_STEP_Twips * 1, lngTmpY)-(X0 + HOUR_STEP_Twips * 1, Y), COLOR.深灰色
        objDraw.Line (X0 + HOUR_STEP_Twips * 2, lngTmpY)-(X0 + HOUR_STEP_Twips * 2, Y), COLOR.深灰色
        objDraw.Line (X0 + HOUR_STEP_Twips * 3, lngTmpY)-(X0 + HOUR_STEP_Twips * 3, Y), COLOR.深灰色
        objDraw.Line (X0 + HOUR_STEP_Twips * 4, lngTmpY)-(X0 + HOUR_STEP_Twips * 4, Y), COLOR.深灰色
        objDraw.Line (X0 + HOUR_STEP_Twips * 5, lngTmpY)-(X0 + HOUR_STEP_Twips * 5, Y), COLOR.深灰色

        Call DrawLine(objDraw, X0 + HOUR_STEP_Twips * 6, lngTmpY, X0 + HOUR_STEP_Twips * 6, Y, vbRed, , 2)
    Next
    
    '设置下一画图坐标
    X = lngTmpX
    Y = lngTmpY
End Sub

Private Function DrawBodyScale(objDraw As Object, X As Long, Y As Long, intGuageRow As Integer, arrstrItem() As String, Optional sngScale As Long = 1)
    '******************************************************************************************************************
    '功能:画出体温表项目arrstrItem()所决定的当前页面左边项目刻度
    '参数:intGuageRow=刻度行数
    '     arrstrItem=体温的颜色、首先高、名称、最大、最小、间隔、标记（最后这个参数是为了后续画图的方便）
    '注意:在调用时应确定intGuageRow大于等于1(因为只有intGuageRow大于等于1才表示有体温表项目)
    
    '顺序:
    '画图顺序号:=3/9   这是画所有图形的第三步
    '******************************************************************************************************************
    Dim aryItem() As String
    Dim intCountItem As Integer '记录项目个数
    Dim lngColor As Long
    Dim intItemTop As Integer
    Dim strItem As String
    Dim H_9pt As Long, W_9pt As Long
    Dim i As Long, j As Long, l As Long, k As Long '循环之用
    Dim lngTmpY As Long, lngTmpX As Long
    Dim lngTmpVal As Single
    Dim lngY As Long
    Dim strTmp As String
    Dim lngPercW As Long
    Dim intLoop As Integer
    
    '参考高度
    objDraw.Font.Name = "宋体"
    objDraw.Font.Size = 9 * sngScale
    objDraw.Font.Bold = False
    
    W_9pt = objDraw.TextWidth("字")
    H_9pt = ROWHEIGHT * 10 / 3
    
    
    '为空退出
    If IsEmpty(arrstrItem) = True Then Exit Function
    '参数少了退出
    If UBound(arrstrItem) < 0 Then Exit Function
    '行数不够退出
    If intGuageRow < 1 Then Exit Function
    
    DrawBodyScale = True
    objDraw.DrawStyle = 0
    lngTmpY = Y
    lngTmpX = X
    
    intCountItem = UBound(arrstrItem)
    
    ReDim aryItem(intCountItem)
    For i = 0 To intCountItem '纵
        aryItem(i) = arrstrItem(i)
    Next
    
    For i = 0 To intCountItem '纵
        If GetSplitStr(aryItem(i), 2) = "心率" Then
            
            For intLoop = i To intCountItem - 1
                aryItem(intLoop) = aryItem(intLoop + 1)
            Next
            
            intCountItem = intCountItem - 1
            
            Exit For
        
        End If
    Next
    
    For i = 0 To intCountItem '纵
        If GetSplitStr(aryItem(i), 2) = "脉搏" Then
            lngPercW = W_9pt * 3
        Else
            lngPercW = W_9pt * 7
        End If
        
        Y = lngTmpY + 30
        lngColor = CLng(GetSplitStr(aryItem(i), 0))
        intItemTop = CInt(GetSplitStr(aryItem(i), 1))
        l = 0
        For j = 0 To intGuageRow '横
            '如果参数不够就退出
            If UBound(Split(aryItem(i), "'")) < 6 Then: DrawBodyScale = False: Exit Function
            
            lngTmpVal = Val(GetSplitStr(aryItem(i), 3)) - Val(GetSplitStr(aryItem(i), 5)) * l
            
            If j = 0 Then
            
                '首行为名称
'                If GetSplitStr(aryItem(i), 8) <> "" Then
'                    strItem = GetSplitStr(aryItem(i), 2) & "(" & GetSplitStr(aryItem(i), 8) & ")"
'                Else
                    strItem = GetSplitStr(aryItem(i), 2)
'                End If
                    
                lngY = Y
                Call DrawText(objDraw, X + (lngPercW - objDraw.TextWidth(strItem)) / 2, lngY, strItem, lngColor)
                If strItem = "体温" Then
                    Call DrawText(objDraw, X + (lngPercW - objDraw.TextWidth("   F     C   ")) / 2, lngY + 180, "   F     C   ", lngColor)
                End If
'                For k = 1 To Len(strItem)
'
'                    strTmp = Mid(strItem, k, 1)
'
'                    If strTmp = "/" Then
'                        strTmp = "\"
'                    End If
'
'                    Call DrawRotateText(objDraw, X + (lngPercW - objDraw.TextWidth(strTmp)) / 2, lngY, strTmp, lngColor, 1)
'
'                    lngY = lngY + IIf(Asc(strTmp) < 0, objDraw.TextHeight(strTmp), objDraw.TextWidth(strTmp))
'
'                Next

                Y = Y + H_9pt * 2 ' H_9pt * 3  + H_9pt \ 2 + H_9pt + H_9pt \ 2
                
            ElseIf j >= intItemTop And lngTmpVal >= CLng(GetSplitStr(aryItem(i), 4)) - IIf(strItem = "体温", 1, 0) Then
                
                '画递减间隔
                If (lngTmpVal - Fix(lngTmpVal)) = 0 Then
                    
                    Select Case GetSplitStr(aryItem(i), 2)
                    Case "脉搏", "心率", "呼吸"
                        If lngTmpVal Mod 10 = 0 Then
                            Call DrawText(objDraw, X + (lngPercW - objDraw.TextWidth(CStr(lngTmpVal))) / 2, Y + objDraw.TextHeight(CStr(lngTmpVal)) / 2 - 30, CStr(lngTmpVal), lngColor)
                        End If
                    Case Else
                        strTmp = CStr(lngTmpVal)
                        Call GetValue(strTmp)
                        Call DrawText(objDraw, X + (lngPercW - objDraw.TextWidth(strTmp)) / 2, Y + objDraw.TextHeight(CStr(CStr(lngTmpVal))) / 2 - 30, strTmp, lngColor)
                    End Select
                    
                End If
                Y = Y + H_9pt + H_9pt \ 2
                l = l + 1
            Else
                Y = Y + H_9pt + H_9pt \ 2
            End If
        Next
        X = X + lngPercW
    Next
    
    X = lngTmpX
    Y = lngTmpY

    '画线
    Call DrawCell(objDraw, 1, lngTmpX, lngTmpY, (W_9pt + W_9pt \ 2) * (intCountItem + 1), 0)
    For i = 0 To intCountItem + 1
        Call DrawCell(objDraw, 1, X, Y, 0, (intGuageRow) * (H_9pt + H_9pt \ 2) + H_9pt * 7)

        If i = 0 Or i = intCountItem + 1 Then
            Call DrawLine(objDraw, X, Y, X, Y + (intGuageRow) * (H_9pt + H_9pt \ 2) + H_9pt * 7, , , 2)
        End If

        X = X + IIf(i = 0, W_9pt * 3, W_9pt * 7)
    Next
    
    '为了将值传给下一个画图的起始点
    X = lngTmpX + W_9pt * 10
    Y = lngTmpY
End Function

Private Function GetSplitStr(strValue As String, ByVal Index As Long) As String
    '功能：从列表中得到指定字符串
    Dim arrStrTmp As Variant
    If strValue = "" Then Exit Function
    arrStrTmp = Split(strValue, "'")
    If Index >= LBound(arrStrTmp) And Index <= UBound(arrStrTmp) Then
        GetSplitStr = arrStrTmp(Index)
    End If
End Function

Private Sub CloseRs(rs As ADODB.Recordset)
    '功能：关闭Recordset对象
    On Error Resume Next
    If rs.State = ADODB.adStateOpen Then rs.Close
    Set rs = Nothing
End Sub

Private Sub CalcOPS(lngArryOPSDay() As String, ByVal DayCount As Long, ByVal Index As Long, Optional ByVal lng次数 As Long)
    '功能:计算每次手术后的天数
    '参数:lngArryOPSDay=要进行设置的天数数组
    '     DayCount=为住院的总天数,此参数将自动初始化lngArryOPSDay数组的长度
    '     Index=要将某天的索引设置为手术
    Dim i As Long
    Dim strTmp As String
    
    If Index > DayCount - 1 Then Exit Sub
    
    If UBound(lngArryOPSDay) < DayCount - 1 Then
        ReDim Preserve lngArryOPSDay(DayCount - 1)
    End If
    For i = Index To Index + mintOpDays
        
        If i <= UBound(lngArryOPSDay) Then
        Select Case i - Index
        Case 0
        
            strTmp = lng次数
            If mblnStopFlag Then
                lngArryOPSDay(i) = "0"
            Else
                lngArryOPSDay(i) = IIf(lngArryOPSDay(i) <> "" And lngArryOPSDay(i) <> "-1", lngArryOPSDay(i) & "(" & strTmp & ")", "0")
            End If
            
        Case Else
            
            If (i - Index) <= mintOpDays Then
                
                If mblnStopFlag Then
                    lngArryOPSDay(i) = (i - Index)
                Else
                    lngArryOPSDay(i) = IIf(lngArryOPSDay(i) <> "", (i - Index) & "/" & lngArryOPSDay(i), (i - Index))
                End If
                
            End If
        End Select
        End If
    Next
End Sub

Public Function ConvertTimeToChinese(ByVal strTime As String) As String
    '------------------------------------------------------------------------------------------------------------------
    '功能：
    '参数：
    '返回：
    '------------------------------------------------------------------------------------------------------------------
    Dim strTmp1 As String
    Dim strTmp2 As String
    
    If InStr(strTime, ":") <= 0 Then Exit Function
    On Error GoTo errHand
    
    strTmp1 = Left(strTime, InStr(strTime, ":") - 1)
    strTmp2 = Mid(strTime, InStr(strTime, ":") + 1)
    
    strTmp1 = Switch(strTmp1 = "00", "零", strTmp1 = "01", "一", strTmp1 = "02", "二", strTmp1 = "03", "三", _
                    strTmp1 = "04", "四", strTmp1 = "05", "五", strTmp1 = "06", "六", strTmp1 = "07", "七", _
                    strTmp1 = "08", "八", strTmp1 = "09", "九", strTmp1 = "10", "十", strTmp1 = "11", "十一", _
                    strTmp1 = "12", "十二", strTmp1 = "13", "十三", strTmp1 = "14", "十四", strTmp1 = "15", "十五", _
                    strTmp1 = "16", "十六", strTmp1 = "17", "十七", strTmp1 = "18", "十八", strTmp1 = "19", "十九", _
                    strTmp1 = "20", "二十", strTmp1 = "21", "二十一", strTmp1 = "22", "二十二", strTmp1 = "23", "二十三")
    
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
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function PrintOrPreviewBodyState(objOut As Object, _
                                        ByVal lng病人id As Long, _
                                        ByVal lng主页id As Long, _
                                        ByVal intBaby As Integer, _
                                        ByVal lngSectID As Long, _
                                        ByVal lngBeginY As Long, _
                                        ByVal lngLeft As Long, _
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
    
    Dim blnPrint As Boolean
    Dim strInfo As String                                   '是不是打印机 ，显示提示信息
    Dim lngPage As Long                                     '当前页
    Dim intAllOpt As Single
    Dim intCurOpt As Single                                 '总进度，当前进度
    Dim i As Long, j As Long, l As Long, lngRecItemRow As Long, lngRecordCount As Long
    Dim X As Long, Y As Long                                'X坐标，Y坐标（Twip），
    Dim objDraw As Object                                   '进行画图的对像
    Dim intDrawLineRows As Integer                          '求出左边标尺总行数(不含标题行)，最多20行
    Dim intDrawLineCols As Integer                          '求出左边标尺总列数
    Dim intDrawGridRows As Integer                          '求出底部体温表录入项目的行数
    Dim intRepairRows As Integer
    Dim strBeginDate As String, strEndDate As String        '求出病人的开始与终止时间
    Dim strPatiInfo As String                               '病人信息列表:姓名 住院号 科室 病区 床号 日期
    Dim strStateTips As String                              '底部的说明信息
    Dim strArrItemInfo() As String                          '体温表曲线项目列表数组
    Dim strArrItemDataInfo() As String                      '体温表曲线项目数据列表数组
    Dim strArrItemDataComment() As String                   '体表曲线项目数据的说明数组
    Dim strArrRecordItemInfo() As String                    '体温表录入项目列表数组
    Dim rsArrRecordItemInfo As New ADODB.Recordset          '体温表录入项目列表记录
    Dim lngArr手术后天数() As String
    Dim lngOPSDayCount As Long
    Dim blnTag As Boolean '确定是否有曲线数据项目的数据
    Dim blnComment As Boolean '确定是否有曲线数据项目的说明
    Dim lngCountPage As Long '依据病人的总天数求出页数
    Dim rsTemp As New ADODB.Recordset, strSQL As String
    '纸张尺寸信息
    Dim lngTop As Long
    Dim dblSureW As Double, dblSureH As Double
    Dim H_9pt As Long, W_9pt As Long '一个小五字的高度
    Dim H_16pt As Long, W_16pt As Long
    Dim lngTmpX As Long, lngTmpY As Long
    Dim mlngHourBegin As Long
    Dim strTmp As String
    Dim lngCol As Long
    Dim strTime As String
    Dim strSvrTime As String
    Dim intCount As Integer
    Dim lngTmpDay As Long    '临时用来记录当前页某时间点上对应的列数
    Dim strTmpDate As String '临时用来记录当前页面的日期时间段
    Dim strTmpDay As String '临时保存当前页面的开始日期
    Dim strTmpString0 As String, strTmpString1 As String, strTmpString2 As String '临时用
    Dim strNewTmpString As String
    Dim lngNewTmpX As Long, lngNewTmpY As Long
    Dim lngPicPageIndex As Long
    Dim strArrNewTmp() As String, strArrNewTmpComment() As String
    Dim strArrCurLineData(0 To 41) As String
    Dim strArrCurLineComment(0 To 41) As String
    Dim strArrItemDataInOut(0 To 41) As String
    Dim strArrItemTmpInOut() As String
    Dim aryTmp As Variant
    Dim rsTmp As New ADODB.Recordset
    Dim lngValue As Long
    Dim byt未记显示位置 As Byte
    Dim lng偏移量(0 To 41) As Long
    Dim blnAllow As Boolean
    Dim blnShow As Boolean
        
    Dim lngRowCount As Long
    
    ReDim mbyt脉搏(0 To 41) As Byte
    
    '求出项目的属时用到。主要是为了求出 单位个数
    Dim dbl最大 As Double, dbl最小 As Double, dbl单位值 As Double, lng最高行 As Long, lng单位个数 As Long
    
    
    '------------------------------------------------------------------------------------------------------------------
    On Error GoTo ErrPrint
    
    msngScale = 0.85
    
    '读取体温表一天开始时间
    '------------------------------------------------------------------------------------------------------------------
    mlngHourBegin = Val(zlDatabase.GetPara("体温开始时间", glngSys, 1255, 4))

    mintOpDays = Val(zlDatabase.GetPara("手术后标注天数", glngSys, 1255, "10"))
    mblnStopFlag = (Val(zlDatabase.GetPara("再次手术停止前次标注", glngSys, 1255, "0")) = 1)
    byt未记显示位置 = Val(zlDatabase.GetPara("未记说明显示位置", glngSys, 1255, "0"))
    mbln婴儿体温单显示出院 = (zlDatabase.GetPara("婴儿体温单显示出院信息", glngSys, 1255, 1) = 1)
    
    '病人变动标记显示方法
    '------------------------------------------------------------------------------------------------------------------
    strTmp = zlDatabase.GetPara("体温单标记", glngSys, 1255, "1;1;1;1;1;1;1;1")
    If UBound(Split(strTmp, ";")) >= 5 Then
        mBodyFlag.入院 = Val(Split(strTmp, ";")(0))
        mBodyFlag.入科 = Val(Split(strTmp, ";")(1))
        mBodyFlag.转出 = Val(Split(strTmp, ";")(2))
        mBodyFlag.换床 = Val(Split(strTmp, ";")(3))
        mBodyFlag.手术 = Val(Split(strTmp, ";")(4))
        mBodyFlag.出院 = Val(Split(strTmp, ";")(5))
        If UBound(Split(strTmp, ";")) >= 6 Then mBodyFlag.分娩 = Val(Split(strTmp, ";")(6))
        If UBound(Split(strTmp, ";")) >= 7 Then mBodyFlag.出生 = Val(Split(strTmp, ";")(7))
    End If
    
    blnPrint = TypeName(objOut) = "Printer"
    Screen.MousePointer = 11
    intAllOpt = 6
    
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
'                    objOut.txtPage.Text = ""
                Else
                    Unload objOut.picPage(i)
                End If
            Next
            Set objDraw = objOut.picPage(0) 'PictureBox
            objDraw.Width = Printer.Width * sngScale
            objDraw.Height = Printer.Height * sngScale
        Else
            Set objDraw = Printer
        End If
    Else
        If Not blnPrint Then
            i = objOut.picPage.UBound + 1
            Load objOut.picPage(i)
'            objOut.SetPages
            Set objDraw = objOut.picPage(objOut.picPage.UBound)
            objDraw.Width = Printer.Width * sngScale
            objDraw.Height = Printer.Height * sngScale
        Else
            Set objDraw = Printer
        End If
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    strSQL = _
        "   Select Decode(b.出生时间,Null,a.开始,b.出生时间) As 开始,a.终止 From (Select 病人ID,主页id,Min(开始时间) as 开始,Max(Nvl(终止时间,sysdate)) as 终止" & _
        "    From 病人变动记录" & _
        "    Where 开始时间 is Not Null And 病人ID=[1] And 主页ID=[2] Group By 病人ID,主页id) a, " & _
        "   (Select 病人ID,主页id,出生时间 From 病人新生儿记录 Where 病人ID = [1] And 主页ID = [2] And 序号=[3]) b Where a.病人id=b.病人id(+) And a.主页id=b.主页id(+) "

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlPrint", lng病人id, lng主页id, intBaby)
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
        lngCountPage = DateDiff("d", rsTemp!开始, rsTemp!终止) + 1
        lngCountPage = IIf(lngCountPage / 7 = Fix(lngCountPage / 7), lngCountPage / 7, Fix(lngCountPage / 7) + 1)
        strBeginDate = Format(rsTemp!开始, "YYYY-MM-DD HH:MM:SS")
        strEndDate = Format(rsTemp!终止, "YYYY-MM-DD HH:MM:SS")
    Else
        CloseRs rsTemp
        GoTo ErrPrint '无数病人变动信息退出
    End If
    
    ReDim lngArr手术后天数(DateDiff("d", CDate(strBeginDate), CDate(strEndDate)))

    intCurOpt = intCurOpt + 1
    
    Call ShowFlash(strInfo, intCurOpt / intAllOpt, objParent)
    
    '------------------------------------------------------------------------------------------------------------------
    '第１部份：病人的基本信息
    '读取病人基本信息
    Dim varPatiInfo As Variant
    
    '姓名'住院号''入院时间
    strPatiInfo = "''''''"
    varPatiInfo = Split(strPatiInfo, "'")
    
    strSQL = " Select  b.姓名,A.住院号,b.入院时间,b.性别,b.年龄 From 病人信息 B,病案主页 A Where A.病人ID=B.病人ID And A.病人id=[1] And A.主页ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlPrint", lng病人id, lng主页id)
    If rsTemp.BOF = False Then
        varPatiInfo(0) = zlCommFun.NVL(rsTemp("姓名").Value)
        varPatiInfo(1) = zlCommFun.NVL(rsTemp("住院号").Value)
        varPatiInfo(3) = Format(zlCommFun.NVL(rsTemp("入院时间").Value), "yyyy-MM-dd")
        varPatiInfo(5) = zlCommFun.NVL(rsTemp("性别").Value)
        varPatiInfo(6) = zlCommFun.NVL(rsTemp("年龄").Value)
    End If
    
    '入院时间(以入科时间为准)
    mstrSQL = "select 开始时间 from 病人变动记录 where 病人id=[1] And 主页id=[2] and 开始原因=2 order by 开始时间"
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "mdlPrint", lng病人id, lng主页id)
    If rsTemp.BOF = False Then
        varPatiInfo(3) = Format(zlCommFun.NVL(rsTemp("开始时间").Value), "yyyy-MM-dd")
    End If
        
        
    Select Case intBaby
    Case 0
        
    Case Else
        
        varPatiInfo(5) = ""
        varPatiInfo(6) = ""
        
        gstrSQL = "Select Decode(a.婴儿姓名,Null,b.姓名||'之子'||Trim(To_Char(a.序号,'9')),a.婴儿姓名) As 婴儿姓名,婴儿性别,出生时间 From 病人新生儿记录 a,病人信息 b Where a.病人id=[1] And a.主页id=[2] And a.病人id=b.病人id And a.序号=[3]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "mdlPrint", lng病人id, lng主页id, intBaby)
        If rsTemp.BOF = False Then
            varPatiInfo(0) = rsTemp("婴儿姓名").Value
            varPatiInfo(5) = zlCommFun.NVL(rsTemp("婴儿性别").Value)
            varPatiInfo(6) = "新生儿"
            
            If IsNull(rsTemp("出生时间").Value) = False Then varPatiInfo(3) = Format(zlCommFun.NVL(rsTemp("出生时间").Value), "yyyy-MM-dd")
        End If
        
    End Select
        

    intCurOpt = intCurOpt + 1
    Call ShowFlash(strInfo, intCurOpt / intAllOpt, objParent)
    
    '------------------------------------------------------------------------------------------------------------------
    '第２部份：病人的手术信息
    
    '求出病人手术后的天数
    '先求出病人手术在哪些天,并将时间传给数组
    '找出手术最大的日期

    mstrSQL = "Select 时间,项目名称,rownum as 次数 From (SELECT A.发生时间 As 时间,c.项目名称 " & _
                "FROM 病人护理记录 A,病人护理内容 C " & _
                "Where A.ID=C.记录ID " & _
                    "AND A.病人id=[1] And Nvl(a.婴儿,0)=[4] " & _
                    "AND A.主页id=[2] " & _
                    "AND c.记录类型=4 " & _
                    "AND A.发生时间<[3] And c.终止版本 Is Null Order By A.发生时间)"
    If mblnMoved Then
        mstrSQL = Replace(mstrSQL, "病人护理记录", "H病人护理记录")
        mstrSQL = Replace(mstrSQL, "病人护理内容", "H病人护理内容")
    End If
    
    Dim TmplngDayCount As Long
    
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, "mdlPrint", _
                                        lng病人id, _
                                        lng主页id, _
                                        CDate(strEndDate), intBaby)
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        lngOPSDayCount = rsTmp.RecordCount
        
        TmplngDayCount = DateDiff("d", CDate(strBeginDate), CDate(strEndDate)) + 1
        
        For i = 0 To rsTmp.RecordCount - 1
            '按时间调整手术后天数的排列
            If IsNull(rsTmp!时间) = False Then
                Call CalcOPS(lngArr手术后天数, TmplngDayCount, DateDiff("d", CDate(strBeginDate), rsTmp!时间), rsTmp!次数)
            End If
            rsTmp.MoveNext
        Next
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    '第３部份：体温曲线项目
    
    '1、先求初始化项目字符串数组
    '求体温表曲线项目
    
    mint心率应用 = 2
    mstr心率符号 = ""
    mstrSQL = "Select a.应用方式,b.记录符 From 护理记录项目 a,体温记录项目 b Where a.项目序号=-1 And a.项目序号=b.项目序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "mdlPrint")
    If rsTemp.BOF = False Then
        mint心率应用 = zlCommFun.NVL(rsTemp("应用方式").Value, 2)
        mstr心率符号 = zlCommFun.NVL(rsTemp("记录符").Value, "○")
    End If
    
        
    '得到所有曲线项目
    mstrSQL = " Select c.项目序号,A.记录名 as 项目名,A.记录符,To_Char(A.记录色)||''''||To_Char(A.最高行)||''''||A.记录名||''''||To_Char(A.最大值)||''''||To_Char(A.最小值)||''''||To_Char(A.单位值)||''''||A.记录符||''''||To_Char(A.项目序号)||''''||c.项目单位 As 列表 " & _
                " From 体温记录项目 A,护理记录项目 C " & _
                " Where A.项目序号=C.项目序号 And A.记录法=[1] AND C.护理等级>=[2] And Nvl(c.应用方式,0)=1 And Nvl(c.适用病人,0) In (0,[4]) " & _
                " And (c.适用科室=1 Or (c.适用科室=2 And Exists (Select 1 From 护理适用科室 D Where D.项目序号=c.项目序号 And D.科室id=[3]))) " & _
                " Order by A.排列序号"
                
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "mdlPrint", 1, 0, lngSectID, IIf(intBaby = 0, 1, 2))
    If rsTemp.RecordCount < 1 Then
        CloseRs rsTemp
        MsgBox "无任何体温表项目！", vbExclamation, gstrSysName
        GoTo ErrExit    '无数据退出
    End If
    rsTemp.MoveFirst
                        
    ReDim strArrItemInfo(rsTemp.RecordCount - 1)                    '确定项目信息列表元素个数
    ReDim strArrItemDataInfo(rsTemp.RecordCount - 1)                '确定项目数据列表元素个数
    ReDim strArrItemDataComment(rsTemp.RecordCount - 1)             '确定曲线项目各数据列的说明个数
    Dim bytShow As Byte
    
    Dim varTmp As Variant

    
    mbln呼吸曲线 = False
    strStateTips = "说明:"
    For i = 0 To UBound(strArrItemInfo)
        strArrItemInfo(i) = rsTemp!列表
        
        Select Case rsTemp("项目序号").Value
        Case -1
            strTmp = rsTemp!项目名 & "(" & rsTemp!记录符 & ")"
        Case 1
            varTmp = Split(zlCommFun.NVL(rsTemp("记录符").Value, "·,×,○"), ",")
            
            mstrChar(0) = CStr(varTmp(0))
            mstrChar(1) = CStr(varTmp(1))
            mstrChar(2) = CStr(varTmp(2))
    
            strTmp = rsTemp!项目名 & "(口温" & mstrChar(0) & ",腋温" & mstrChar(1) & ",肛温" & mstrChar(2) & ")"
        Case 2
            '有问题:mint心率应用 = 2
            mstrPulse = rsTemp!记录符
            If mint心率应用 = 0 Then
                strTmp = rsTemp!项目名 & "(" & rsTemp!记录符 & ",心率○)"
            Else
                strTmp = rsTemp!项目名 & "(" & rsTemp!记录符 & ")"
            End If
        Case 3
            mstrBreath = rsTemp!记录符
            mbln呼吸曲线 = True
            strTmp = rsTemp!项目名 & "(" & rsTemp!记录符 & ")"
        Case Else
            strTmp = rsTemp!项目名 & "(" & rsTemp!记录符 & ")"
        End Select
        
        If i = 0 Then
            strStateTips = strStateTips & strTmp
        Else
            strStateTips = strStateTips & "、" & strTmp
        End If
        rsTemp.MoveNext
    Next
    intDrawLineCols = UBound(strArrItemInfo) + 1   '求出体温表曲线项目的个数
    intCurOpt = intCurOpt + 1
    Call ShowFlash(strInfo, intCurOpt / intAllOpt, objParent)
    
    '第４部份：体温录入项目
    '------------------------------------------------------------------------------------------------------------------
    '求出所有录入项目
    mstrSQL = " Select RowNum-1 As 序号,A.* From (Select Decode(A.项目序号,4,'血压',A.记录名) as 项目名,C.项目单位 As 单位,A.项目序号,A.记录频次,C.项目性质 " & _
                " From 体温记录项目 A,护理记录项目 C " & _
                " Where A.项目序号=C.项目序号 And A.记录法=[1] AND C.护理等级>=[2] And A.项目序号<>5 And Nvl(c.应用方式,0)=1 And Nvl(c.适用病人,0) In (0,[4]) " & _
                " And (c.适用科室=1 Or (c.适用科室=2 And Exists (Select 1 From 护理适用科室 D Where D.项目序号=c.项目序号 And D.科室id=[3]))) " & _
                " Order by A.排列序号) A"
    Set rsArrRecordItemInfo = zlDatabase.OpenSQLRecord(mstrSQL, "mdlPrint", 2, 0, lngSectID, IIf(intBaby = 0, 1, 2))
    If rsArrRecordItemInfo.RecordCount > 0 Then
        rsArrRecordItemInfo.MoveFirst
        ReDim strArrRecordItemInfo(rsArrRecordItemInfo.RecordCount - 1)
        For i = 0 To UBound(strArrRecordItemInfo)
            
            strArrRecordItemInfo(i) = rsArrRecordItemInfo!项目名 & "'" & IIf(IsNull(rsArrRecordItemInfo!单位), "", rsArrRecordItemInfo!单位) & "'" & zlCommFun.NVL(rsArrRecordItemInfo("记录频次").Value, 2)
            rsArrRecordItemInfo.MoveNext
        Next
        intDrawGridRows = UBound(strArrRecordItemInfo) + 1 '求出体温表录入项目的行数

    Else
        intDrawGridRows = 0 '求出体温表录入项目的行数
    End If
    intCurOpt = intCurOpt + 1
    Call ShowFlash(strInfo, intCurOpt / intAllOpt, objParent)
    
    
    '第５部份：图形数据输出
    '------------------------------------------------------------------------------------------------------------------
    '2、确定X和Y的坐标位置
    '边界信息(Twip)
    lngLeft = Val(zlDatabase.GetPara("体温单左边距", glngSys, 1255, OFFSET_LEFT)) * 56.7 * sngScale
    lngTop = Val(zlDatabase.GetPara("体温单上边距", glngSys, 1255, OFFSET_TOP)) * 56.7 * sngScale
    
    
    '确定起始坐标
    X = lngLeft * sngScale: Y = lngTop * sngScale
    objDraw.CurrentX = X: objDraw.CurrentY = Y
    '3、打印对象的初始化
    '求出参考高度
    objDraw.Font = "宋体"
    objDraw.FontSize = 24 * sngScale
    objDraw.FontBold = True
    H_16pt = objDraw.TextHeight("字")
    W_16pt = objDraw.TextWidth("字")
    
    lngTmpX = X
    lngTmpY = Y
    objDraw.Font = "宋体"
    objDraw.FontSize = 9 * sngScale
    objDraw.FontBold = False
    H_9pt = (ROWHEIGHT * 10 / 3) * sngScale
    W_9pt = objDraw.TextWidth("字")
    
    mlngFirstWidth = W_9pt * 10
    
    '------------------------------------------------------------------------------------------------------------------
    '4、求得首列宽，求左边标尺总共有多少行
    '求得体温表项目的总行数
    strSQL = "Select Max((A.最大值-A.最小值)/Decode(A.单位值,0,1,A.单位值)+A.最高行) as 总高度 From 体温记录项目 A Where A.记录法=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlPrint", 1)
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
        intDrawLineRows = MAXROWS
    Else
        CloseRs rsTemp
        GoTo ErrPrint
    End If
    
    If intDrawLineRows < 1 Then
        CloseRs rsTemp
        GoTo ErrPrint
    End If

'    lngColFirstWidth = 6 * W_9pt
    
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
        If strTmpDay < strBeginDate Then strTmpDay = strBeginDate
        
        strTmpDate = Format(CDate(strTmpDay), "MM-DD") '记录当前期间段
        strTmpDate = strTmpDate & "～" & Format(IIf(CDate(strBeginDate) + 7 * (lngPage - 1) + 6 > CDate(strEndDate), strEndDate, (CDate(strBeginDate) + 7 * (lngPage - 1) + 6)), "MM-DD")
            
        intCurOpt = lngPage / lngCountPage
        strInfo = "正在" & IIf(blnPrint, "打印体温表", "预览") & ",请稍候..."
        Call ShowFlash(strInfo, intCurOpt, objParent)
        
        '按页号打印
        If intBeginPage > 0 Then  '只打印指定页码的
            If lngPage >= intBeginPage And lngPage <= intEndPage Then
                If lngPage > intBeginPage Then  '到第二页时开始初始化纸张或页面
                    If Not blnPrint Then
                        Load objOut.picPage(lngPicPageIndex)
'                        objOut.cboPage.AddItem "第 " & (lngPage - intBeginPage + 1) & " 页"
'                        objOut.txtPage.Text = "当前页" & Space(17) & "共 " & objOut.picPage.UBound + 1 & " 页"
                        Set objDraw = objOut.picPage(lngPicPageIndex) 'PictureBox
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
'                    objOut.cboPage.AddItem "第 " & lngPage & " 页"
'                    objOut.txtPage.Text = "当前页" & Space(17) & "共 " & objOut.picPage.UBound + 1 & " 页"
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
        
        X = lngTmpX
        Y = lngTmpY

        objDraw.Font = "宋体"
        objDraw.FontSize = 9 * sngScale
        objDraw.FontBold = False
        '打印质控号
        
        strTmp = zlDatabase.GetPara("质控号", glngSys, 1255, "")
        
        X = lngTmpX + (6 * W_9pt + HOUR_STEP_Twips * 6 * 7) - objDraw.TextWidth(strTmp)

        Call DrawText(objDraw, X, lngTmpY - objDraw.TextHeight(strTmp), strTmp)
        
        X = lngTmpX
        
        objDraw.Font = "黑体"
        objDraw.FontSize = 18 * sngScale
        objDraw.FontBold = True
        
        '打印医院名称,打印体温单标题
        strTmpString0 = IIf(GetUnitName = "-", "", GetUnitName) & "体温单"
        If strTmpString0 <> "" Then
            Call DrawCell(objDraw, strTmpString0, lngTmpX + (((UBound(strArrItemInfo) + 1) * (W_9pt + W_9pt \ 2) + HOUR_STEP_Twips * 6 * 7) - objDraw.TextWidth(strTmpString0)) / 2, lngTmpY, objDraw.TextWidth(strTmpString0), H_16pt + H_16pt \ 2, , , , , , objDraw.Font, "0000", 1, 1)
        End If

        strTmpString0 = ""
        Y = Y + H_16pt + 2 * H_16pt / 3
        
        objDraw.Font = "宋体"
        objDraw.FontSize = 10 * sngScale
        objDraw.FontBold = True
        

        varPatiInfo(2) = ""
        varPatiInfo(4) = ""
        strTmp = ""
        strTime = ""
        
        strSQL = " Select  c.名称 As 科室,b.名称 As 病区,d.房间号 AS 床号,a.开始原因 " & _
                "From 病人变动记录 a,部门表 b,部门表 c,床位状况记录 D " & _
                "Where a.病人id=[1] And a.主页id=[2] And a.科室id Is Not Null And a.病区id=b.id and a.科室id=c.id " & _
                "And a.床号=d.床号 And a.开始时间-4/24<=[3] And Nvl(a.终止时间,Sysdate)>=[4] " & _
                "Order By a.开始时间 desc"
    
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPrint", lng病人id, lng主页id, CDate(strTmpDay) + 7, CDate(strTmpDay))
        If rsTmp.BOF = False Then
            varPatiInfo(2) = zlCommFun.NVL(rsTmp("科室").Value)
            varPatiInfo(4) = zlCommFun.NVL(rsTmp("床号").Value)
            
'            Do While Not rsTmp.EOF
'
'                If zlCommFun.NVL(rsTmp("科室").Value) <> strTmp And zlCommFun.NVL(rsTmp("科室").Value) <> "" Then
'
'                    strTmp = zlCommFun.NVL(rsTmp("科室").Value)
'
'                    If varPatiInfo(2) = "" Then
'                        varPatiInfo(2) = strTmp
'                    Else
'                        varPatiInfo(2) = varPatiInfo(2) & "->" & strTmp
'                    End If
'
'                End If
'
'                If zlCommFun.NVL(rsTmp("床号").Value) <> strTime And zlCommFun.NVL(rsTmp("床号").Value) <> "" Then
'
'                    strTime = zlCommFun.NVL(rsTmp("床号").Value)
'
'                    If varPatiInfo(4) = "" Then
'                        varPatiInfo(4) = strTime
'                    Else
'                        varPatiInfo(4) = varPatiInfo(4) & "->" & strTime
'                    End If
'
'                End If
'
'                rsTmp.MoveNext
'            Loop
'
'            If Left(varPatiInfo(2), 2) = "->" Then varPatiInfo(2) = Mid(varPatiInfo(2), 3)
'            If Left(varPatiInfo(4), 2) = "->" Then varPatiInfo(4) = Mid(varPatiInfo(4), 3)

        End If
        
        strPatiInfo = Join(varPatiInfo, "'")
    

        mstrSQL = "Select Zl_Replace_Element_Value([1],[2],[3],2) As 最后诊断 From Dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, "体温单", "最后诊断", lng病人id, lng主页id)
        If rsTmp.BOF = False Then
            If intBaby = 0 Then
                strPatiInfo = strPatiInfo & "'" & zlCommFun.NVL(rsTmp("最后诊断").Value)
            Else
                strPatiInfo = strPatiInfo & "'"
            End If
        Else
            strPatiInfo = strPatiInfo & "'"
        End If
    
    
        '在先画顶部三行的病人信息表格（注意：后面还没有加上 {日期} 的显示）
        Call DrawPatiInfo(objDraw, X, Y, strPatiInfo & "'" & strTmpDate, lngLeft + 9726)

        objDraw.Font = "宋体"
        objDraw.FontSize = 9 * sngScale
        objDraw.FontBold = False
                
        '6、求出病人住院日期及手术时间
        '求出当前这个期间段的日期与住院天数
        '------------------------------------------------------------------------------------------------------------------
        strTmpString0 = ""
        strTmpString1 = ""
        strTmpString2 = ""
        
        lngValue = 0
        mstrSQL = "Select zl_CalcInDays([1],[2],[3],[4]) As 开始天数 From Dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, "体温单", lng病人id, lng主页id, intBaby, (Int(CDate(strTmpDay))))
        If rsTmp.BOF = False Then
            lngValue = rsTmp("开始天数").Value
        End If
        
        For i = 0 To 6
            strTmpString0 = strTmpString0 & "'" & Format(CDate(strTmpDay) + i, "YYYY-MM-DD")
            If CDate(Format(strTmpDay, "YYYY-MM-DD")) + i > CDate(Format(strEndDate, "YYYY-MM-DD")) Then
                strTmpString1 = strTmpString1 & "'"
                strTmpString2 = strTmpString2 & "'"
            Else
'                strTmpString1 = strTmpString1 & "'" & Format(Int(CDate(strTmpDay)) + i - Int(CDate(strBeginDate)) + 1)
                strTmpString1 = strTmpString1 & "'" & lngValue + i
                If lngOPSDayCount > 0 Then
                    strTmpString2 = strTmpString2 & "'" & IIf(CStr(lngArr手术后天数((lngPage - 1) * 7 + i)) = "-1", "", CStr(lngArr手术后天数((lngPage - 1) * 7 + i)))
                End If
            End If
        Next
        strTmpString0 = Right(strTmpString0, Len(strTmpString0) - 1)
        strTmpString1 = Right(strTmpString1, Len(strTmpString1) - 1)
        If lngOPSDayCount > 0 Then
            strTmpString2 = Right(strTmpString2, Len(strTmpString2) - 1)
        Else
            strTmpString2 = "'''''''''"
        End If
        lngNewTmpX = X
        
        Call DrawBodyInfo(objDraw, X, Y, mlngFirstWidth, strTmpString0, strTmpString1, strTmpString2, , strBeginDate)
        lngNewTmpY = Y + (intDrawLineRows) * (H_9pt + H_9pt \ 2) + H_9pt * 2 + H_9pt / 2 - 30
        
        Call DrawBodyScale(objDraw, X, Y, intDrawLineRows, strArrItemInfo)
        Call DrawBodyTopScale(objDraw, X, Y)
        Call DrawBodyPaper(objDraw, X, Y, intDrawLineRows)
        
        '入出转标记
        '------------------------------------------------------------------------------------------------------------------
        Set rsTmp = GetDataFromHis(lng病人id, lng主页id, intBaby, CDate(strTmpDay), CDate(strTmpDay) + 8, 2)
        If Not (rsTmp Is Nothing) Then
            If rsTmp.BOF = False Then
                
                Do While Not rsTmp.EOF
                    
                    If zlCommFun.NVL(rsTmp("内容")) <> "" Then
                        
                        bytShow = 0
                        Select Case Val(rsTmp("行号").Value)
                        Case 5
                            bytShow = mBodyFlag.入院
                        Case 6
                            bytShow = mBodyFlag.转出
                        Case 7
                            bytShow = mBodyFlag.换床
                        Case 8
                            bytShow = mBodyFlag.出院
                        Case 9
                            bytShow = mBodyFlag.入科
                        End Select
                    
                        If bytShow > 0 Then
                            blnShow = True
                            If Val(rsTmp("行号").Value) = 8 And intBaby > 0 Then
                                blnShow = mbln婴儿体温单显示出院
                            End If
                            
                            If blnShow Then
                                Select Case bytShow
                                Case 1
                                    strTmp = rsTmp("内容").Value
                                Case 2
                                    strTmp = rsTmp("内容").Value & "--" & ConvertTimeToChinese(Format(rsTmp("时间").Value, "HH:mm"))
                                Case 3
                                    strTmp = rsTmp("内容").Value & rsTmp("科室").Value
                                Case 4
                                    strTmp = rsTmp("内容").Value & rsTmp("科室").Value & "--" & ConvertTimeToChinese(Format(rsTmp("时间").Value, "HH:mm"))
                                End Select
                                strTmp = strTmp & "'" & Format(rsTmp("时间").Value, "HH:mm:ss")
                            Else
                                strTmp = ""
                            End If
                        Else
                            strTmp = ""
                        End If
                        
                        intCount = GetCurveColumn(CDate(rsTmp("时间").Value), CDate(strTmpDay), mlngHourBegin) - 1
                        
                        If intCount >= 0 And intCount <= 41 Then
                            strArrItemDataInOut(intCount) = IIf(Trim(strArrItemDataInOut(intCount)) = "", "", Trim(strArrItemDataInOut(intCount)) & ";") & strTmp
                        End If
                        
                    End If
                    
                    rsTmp.MoveNext
                Loop
            End If
        End If
        
        If intBaby > 0 Then
            
            Set rsTmp = GetDataFromHis(lng病人id, lng主页id, intBaby, CDate(strTmpDay), CDate(strTmpDay) + 8, 3)
            If Not (rsTmp Is Nothing) Then
                If rsTmp.BOF = False Then
                    
                    Do While Not rsTmp.EOF
                        
                        If zlCommFun.NVL(rsTmp("内容")) <> "" Then
                        
                            If mBodyFlag.出生 > 0 Then
                            
                                Select Case mBodyFlag.出生
                                Case 1
                                    strTmp = rsTmp("内容").Value
                                Case 2
                                    strTmp = rsTmp("内容").Value & "--" & ConvertTimeToChinese(Format(rsTmp("时间").Value, "HH:mm"))
                                Case 3
                                    strTmp = rsTmp("内容").Value & rsTmp("科室").Value
                                Case 4
                                    strTmp = rsTmp("内容").Value & rsTmp("科室").Value & "--" & ConvertTimeToChinese(Format(rsTmp("时间").Value, "HH:mm"))
                                End Select
                                strTmp = strTmp & "'" & Format(rsTmp("时间").Value, "HH:mm:ss")
                            Else
                                strTmp = ""
                            End If
                            
                            intCount = GetCurveColumn(CDate(rsTmp("时间").Value), CDate(strTmpDay), mlngHourBegin) - 1
                            
                            If intCount >= 0 And intCount <= 41 Then
                                strArrItemDataInOut(intCount) = IIf(Trim(strArrItemDataInOut(intCount)) = "", "", Trim(strArrItemDataInOut(intCount)) & ";") & strTmp
                            End If
                            
                        End If
                        
                        rsTmp.MoveNext
                    Loop
                End If
            End If
        End If
        
        '提取手术为入出转标识
        '------------------------------------------------------------------------------------------------------------------
        mstrSQL = "SELECT A.发生时间 As 时间,c.项目名称 " & _
                    "FROM 病人护理记录 A,病人护理内容 C " & _
                    "Where A.ID=C.记录ID " & _
                        "AND A.病人id=[1] " & _
                        "AND A.主页id=[2]  And Nvl(a.婴儿,0)=[5] " & _
                        "AND c.记录类型=4 " & _
                        "AND A.发生时间 Between [3] And [4] And c.终止版本 Is Null Order By A.发生时间"
        If mblnMoved Then
            mstrSQL = Replace(mstrSQL, "病人护理记录", "H病人护理记录")
            mstrSQL = Replace(mstrSQL, "病人护理内容", "H病人护理内容")
        End If
        
        Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, "mdlPrint", lng病人id, lng主页id, CDate(strTmpDay), CDate(strTmpDay) + 8, intBaby)
        If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            
            For i = 0 To rsTmp.RecordCount - 1
                '按时间调整手术后天数的排列，
                If IsNull(rsTmp!时间) = False Then
                    lngCol = GetCurveColumn(CDate(rsTmp("时间").Value), CDate(strTmpDay), mlngHourBegin) - 1
                    
                    Select Case rsTmp("项目名称").Value
                    Case "分娩"
                        If lngCol >= 0 And lngCol <= 41 And mBodyFlag.分娩 > 0 Then
                            If mBodyFlag.分娩 = 2 Then
                                strTmp = rsTmp("项目名称").Value & "--" & ConvertTimeToChinese(Format(rsTmp("时间").Value, "HH:mm"))
                            Else
                                strTmp = rsTmp("项目名称").Value
                            End If
                            strTmp = strTmp & "'" & Format(rsTmp("时间").Value, "HH:mm:ss")
                            
                            strArrItemDataInOut(lngCol) = IIf(Trim(strArrItemDataInOut(lngCol)) = "", "", Trim(strArrItemDataInOut(lngCol)) & ";") & strTmp
                        End If
                    
                    Case Else
                        If lngCol >= 0 And lngCol <= 41 And mBodyFlag.手术 > 0 Then
                            If mBodyFlag.手术 = 2 Then
                                strTmp = rsTmp("项目名称").Value & "--" & ConvertTimeToChinese(Format(rsTmp("时间").Value, "HH:mm"))
                            Else
                                strTmp = rsTmp("项目名称").Value
                            End If
                            strTmp = strTmp & "'" & Format(rsTmp("时间").Value, "HH:mm:ss")
                            
                            strArrItemDataInOut(lngCol) = IIf(Trim(strArrItemDataInOut(lngCol)) = "", "", Trim(strArrItemDataInOut(lngCol)) & ";") & strTmp
                        End If
                    End Select
                End If
                rsTmp.MoveNext
            Next
        End If
        
        '7、用For循环，分别求出当前项目
        '求出所有曲线的数据部分 'strArrItemInfo里保存着曲线项目的项目信息
        '将当前页面时间段内（strTmpDay至(strTmpDay+6)）的数据按项目依次读入数组
        '--------------------------------------------------------------------------------------------------------------
        Dim rsOffset As ADODB.Recordset
        
        Call InitOffset(rsOffset)
        
        For l = 0 To UBound(strArrItemInfo)
        
            dbl最大 = Val(Split(strArrItemInfo(l), "'")(3))
            dbl最小 = Val(Split(strArrItemInfo(l), "'")(4))
            dbl单位值 = Val(Split(strArrItemInfo(l), "'")(5))
            lng最高行 = Val(Split(strArrItemInfo(l), "'")(1))
            lng单位个数 = (dbl最大 - dbl最小) / dbl单位值
            lng单位个数 = IIf(lng单位个数 + lng最高行 > (MAXROWS - 1), (MAXROWS - 1) - lng最高行, lng单位个数 + lng最高行)
            
            '----------------------------------------------------------------------------------------------------------
            '读取曲线数据
            mstrSQL = "SELECT A.发生时间 As 时间,C.记录内容 As 数值,c.记录标记,c.体温部位,c.复试合格 " & _
                        "FROM 病人护理记录 A,病人护理内容 C,体温记录项目 D,护理记录项目 E " & _
                        "Where A.ID=C.记录ID " & _
                            "AND A.病人id=[1] " & _
                            "AND A.主页id=[2]  And Nvl(a.婴儿,0)=[7] " & _
                            "AND D.项目序号=C.项目序号 " & _
                            "AND C.记录类型=1 " & _
                            "AND E.项目序号=D.项目序号 And (Nvl(E.应用方式,0)=1 Or ([6]=-1 And E.应用方式=2 And c.记录标记=1)) And Nvl(e.适用病人,0) In (0,[8]) " & _
                            "AND E.护理等级>=0  " & _
                            "AND a.发生时间 BETWEEN [3] And [4] And c.终止版本 Is Null And c.未记说明 Is Null " & _
                            "AND D.记录法=1 AND D.项目序号 In ([5],[6]) And c.记录标记 In ([9],[10]) " & _
                        "Order By a.发生时间,c.记录标记"
                        
            If mblnMoved Then
                mstrSQL = Replace(mstrSQL, "病人护理记录", "H病人护理记录")
                mstrSQL = Replace(mstrSQL, "病人护理内容", "H病人护理内容")
            End If
            
            Select Case Val(Split(strArrItemInfo(l), "'")(7))
            Case 2
                
                If mint心率应用 = 2 Then
                    '脉搏项目还要加上心率项目
                    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "mdlPrint", lng病人id, lng主页id, CDate(Format(strBeginDate, "YYYY-MM-DD")), CDate(Format(strEndDate, "YYYY-MM-DD") & " 23:59:59"), Val(Split(strArrItemInfo(l), "'")(7)), -1, intBaby, IIf(intBaby = 0, 1, 2), 0, 1)
                Else
                    '心率单独应用时，脉搏项目不需要加上心率项目
                    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "mdlPrint", lng病人id, lng主页id, CDate(Format(strBeginDate, "YYYY-MM-DD")), CDate(Format(strEndDate, "YYYY-MM-DD") & " 23:59:59"), Val(Split(strArrItemInfo(l), "'")(7)), 0, intBaby, IIf(intBaby = 0, 1, 2), 0, 0)
                End If
            Case -1
                Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "mdlPrint", lng病人id, lng主页id, CDate(Format(strBeginDate, "YYYY-MM-DD")), CDate(Format(strEndDate, "YYYY-MM-DD") & " 23:59:59"), Val(Split(strArrItemInfo(l), "'")(7)), 2, intBaby, IIf(intBaby = 0, 1, 2), 1, 1)
            Case Else
                Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "mdlPrint", lng病人id, lng主页id, CDate(Format(strBeginDate, "YYYY-MM-DD")), CDate(Format(strEndDate, "YYYY-MM-DD") & " 23:59:59"), Val(Split(strArrItemInfo(l), "'")(7)), 0, intBaby, IIf(intBaby = 0, 1, 2), 0, 1)
            End Select

            '先将数据原始保存到strArrNewTmp数组中
            If rsTemp.RecordCount > 0 Then
            
                rsTemp.MoveFirst
                
                ReDim strArrNewTmp(0)
                
                intCount = -1
                strSvrTime = ""
                For i = 0 To rsTemp.RecordCount - 1

                    lngCol = GetCurveColumn(CDate(rsTemp("时间").Value), CDate(strBeginDate), mlngHourBegin) - 1
                    

                    strTime = Format(Int(CDate(strBeginDate)) + ((lngCol + 1) * 4 - (4 - mlngHourBegin)) / 24, "YYYY-MM-DD hh:mm:ss")
                    
                    If strSvrTime <> strTime Or zlCommFun.NVL(rsTemp("记录标记").Value, 0) = 1 Then
                        strSvrTime = strTime
                        
                        intCount = intCount + 1
                        ReDim Preserve strArrNewTmp(intCount)
                        
                        strArrNewTmp(intCount) = strTime & "'" & Trim(rsTemp!数值) & ";" & Trim(zlCommFun.NVL(rsTemp("记录标记").Value)) & ";" & zlCommFun.NVL(rsTemp("复试合格").Value, 0) & ";" & rsTemp("时间").Value & ";0;" & zlCommFun.NVL(rsTemp("体温部位").Value)
                        
                    ElseIf strSvrTime <> strTime Then
                    
                        intCount = intCount + 1
                        ReDim Preserve strArrNewTmp(intCount)
                        strArrNewTmp(intCount) = strTime & "';"
                        
                    Else
                        
                        intCount = intCount + 1
                        ReDim Preserve strArrNewTmp(intCount)
                        strArrNewTmp(intCount) = strTime & "'" & Trim(rsTemp!数值) & ";" & Trim(zlCommFun.NVL(rsTemp("记录标记").Value)) & ";" & zlCommFun.NVL(rsTemp("复试合格").Value, 0) & ";" & rsTemp("时间").Value & ";1;" & zlCommFun.NVL(rsTemp("体温部位").Value)
                        
                    End If
                    
                    rsTemp.MoveNext
                Next
                lngRowCount = rsTemp.RecordCount
            Else
                lngRowCount = 0
            End If

            '----------------------------------------------------------------------------------------------------------
            '求出所有说明的时间点
            
            mstrSQL = "SELECT c.项目序号 as ItemNO,c.记录类型,A.发生时间 As 时间,Decode(c.记录类型,1,c.未记说明,c.记录内容) As 说明,Decode(c.记录类型,1,1,c.记录标记) As 记录标记 " & _
                        "FROM 病人护理记录 A,病人护理内容 C " & _
                        "Where A.ID=C.记录ID " & _
                            "AND A.病人id=[1] " & _
                            "AND A.主页id=[2]  And Nvl(a.婴儿,0)=[5] " & _
                            "AND (c.记录类型 In (2,6) Or c.记录类型=1 And c.未记说明 Is Not Null) " & _
                            "AND a.发生时间 BETWEEN [3] And [4] And c.终止版本 Is Null " & _
                        "Order By a.发生时间,记录类型"
                        
            If mblnMoved Then
                mstrSQL = Replace(mstrSQL, "病人护理记录", "H病人护理记录")
                mstrSQL = Replace(mstrSQL, "病人护理内容", "H病人护理内容")
            End If
            
            Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "mdlPrint", _
                                                lng病人id, _
                                                lng主页id, _
                                                CDate(Format(strBeginDate, "YYYY-MM-DD")), _
                                                CDate(Format(strEndDate, "YYYY-MM-DD") & " 23:59:59"), intBaby)
            If rsTemp.RecordCount > 0 Then
                '将数据说明保存起来
                rsTemp.MoveFirst
                ReDim strArrNewTmpComment(rsTemp.RecordCount - 1)
                For i = 0 To rsTemp.RecordCount - 1
                
                    lngCol = GetCurveColumn(CDate(rsTemp("时间").Value), CDate(strBeginDate), mlngHourBegin) - 1
                    
                    '未记说明
                    If Val(zlCommFun.NVL(rsTemp!记录类型, 0)) = 1 And Val(zlCommFun.NVL(rsTemp!ItemNO, 0)) = 2 Then
                        '计算当前页的行数
                        If Format(rsTemp("时间").Value, "yyyy-MM-dd") >= strTmpDay And Format(rsTemp("时间").Value, "yyyy-MM-dd") <= Format(DateAdd("d", 6, strTmpDay), "yyyy-MM-dd") Then
                            lngTmpDay = GetCurveColumn(CDate(rsTemp("时间").Value), CDate(strTmpDay), mlngHourBegin) - 1
                            mbyt脉搏(lngTmpDay) = 1
                        End If
                    End If

                    strTime = Format(Int(CDate(strBeginDate)) + ((lngCol + 1) * 4 - (4 - mlngHourBegin)) / 24, "YYYY-MM-DD hh:mm:ss")
                    
                    strArrNewTmpComment(i) = Format(strTime, "YYYY-MM-DD HH:MM:SS") & "'" & Trim(zlCommFun.NVL(rsTemp!说明)) & "'" & zlCommFun.NVL(rsTemp!记录类型, 0) & "'" & zlCommFun.NVL(rsTemp!记录标记, 0)
                    
                    rsTemp.MoveNext
                Next
            End If
               
            For i = 0 To 41
                blnTag = False
                blnComment = False
                
                '计算出当前列的时间点
                strTmp = Format(Int(CDate(strTmpDay)) + ((i + 1) * 4 - (4 - mlngHourBegin)) / 24, "yyyy-MM-dd HH:mm:ss")

                '处理曲线数据
                '------------------------------------------------------------------------------------------------------
                If lngRowCount > 0 Then
                    For j = 0 To UBound(strArrNewTmp)
                    
                        '如果临时数组变量中存在当前时间，就取临时数组变量里的值，否则传入-2
                        '逐个单元格加 4
                        
                        If Trim(strArrNewTmp(j)) <> "" Then
                                                
                            If Format(Split(strArrNewTmp(j), "'")(0), "yyyy-MM-dd HH:mm:ss") = strTmp Then
                                '用来记录当前曲线的数据

                                '如果同一列中有多个值，则取最靠中点的值作为本列的值
                                blnAllow = IsCenterValue(rsOffset, l, i, CDate(Split(Split(strArrNewTmp(j), "'")(1), ";")(3)), CDate(strTmp))
                                                            
                                '记录同一时间两个点的曲线
                                If blnAllow Then
                                    If strArrCurLineData(i) <> "" And strArrCurLineData(i) <> "-2" And Val(Split(Split(strArrNewTmp(j), "'")(1), ";")(4)) = 0 Then
                                        
                                        If Val(Split(strArrItemInfo(l), "'")(7)) = -1 And mint心率应用 = 1 Then
                                            strArrCurLineData(i) = Split(strArrNewTmp(j), "'")(1)
                                        Else
                                            strArrCurLineData(i) = strArrCurLineData(i) & "," & Split(strArrNewTmp(j), "'")(1)
                                        End If
                                        
                                    Else
                                        strArrCurLineData(i) = Split(strArrNewTmp(j), "'")(1)
                                    End If
                                    blnTag = True
                                End If
                            End If
                        End If
                    Next
                End If
                
                '处理说明数据
                '------------------------------------------------------------------------------------------------------
                If rsTemp.RecordCount > 0 Then
                    For j = 0 To UBound(strArrNewTmpComment)
                        
                        If Format(Split(strArrNewTmpComment(j), "'")(0), "yyyy-MM-dd HH:mm:ss") = strTmp Then
                            '用来记录当前曲线数据的说明
                            
                            If strArrCurLineComment(i) = "" Then strArrCurLineComment(i) = ";;"
                            
                            aryTmp = Split(strArrCurLineComment(i), ";")
                            
                            Select Case Val(Split(strArrNewTmpComment(j), "'")(2))
                            Case 1
                                If byt未记显示位置 = 0 Then
                                    If aryTmp(0) = "" Then
                                        aryTmp(0) = Split(strArrNewTmpComment(j), "'")(1)
                                    Else
                                        'If InStr(1, aryTmp(0) & " ", Split(strArrNewTmpComment(j), "'")(1) & " ") = 0 Then
                                            aryTmp(0) = aryTmp(0) & " " & Split(strArrNewTmpComment(j), "'")(1)
                                        'End If
                                    End If
                                Else
                                    If aryTmp(1) = "" Then
                                        aryTmp(1) = Split(strArrNewTmpComment(j), "'")(1)
                                    Else
                                        'If InStr(1, aryTmp(1) & " ", Split(strArrNewTmpComment(j), "'")(1) & " ") = 0 Then
                                            aryTmp(1) = aryTmp(1) & " " & Split(strArrNewTmpComment(j), "'")(1)
                                        'End If
                                    End If
                                End If
                            Case 6

                                If aryTmp(1) = "" Then
                                    aryTmp(1) = Split(strArrNewTmpComment(j), "'")(1)
                                Else
                                    aryTmp(1) = aryTmp(1) & " " & Split(strArrNewTmpComment(j), "'")(1)
                                End If
                                
                            Case Else
                                If aryTmp(0) = "" Then
                                    aryTmp(0) = Split(strArrNewTmpComment(j), "'")(1)
                                Else
                                    aryTmp(0) = aryTmp(0) & " " & Split(strArrNewTmpComment(j), "'")(1)
                                End If
                            End Select
                            
                            If Val(aryTmp(2)) = 0 Then aryTmp(2) = Split(strArrNewTmpComment(j), "'")(3)
                            
                            strArrCurLineComment(i) = Join(aryTmp, ";")
                            
                            blnComment = True
                        End If
                    Next
                End If
                '如果没有就写默认值
                If Not blnTag Then strArrCurLineData(i) = "-2"
                If Not blnComment Then strArrCurLineComment(i) = ""
            Next
            strArrItemDataInfo(l) = Join(strArrCurLineData, "'")
            strArrItemDataComment(l) = Join(strArrCurLineComment, "'")
            
            '清空数据数组
            For j = 0 To UBound(strArrCurLineData)
                strArrCurLineData(j) = ""
            Next
        Next
        '此时已经读取完成所有曲线项目的数据了
        
        '在最后一个参数中加入索引字符串列表表示显示哪些项目
        Call DrawBodyGraph(objDraw, X, Y, strArrItemInfo, strArrItemDataInfo, strArrItemDataComment, strArrItemDataInOut, "")
        
        '8、求出记录项目记录列表数组，重新对 X ，Y 等变量赋值
        X = lngNewTmpX
        Y = lngNewTmpY - H_9pt / 2 - 30    '上午/下午标签行的高度=H_9pt * 2
        
        
        '第６部份：表格数据输出----------------------------------------------------------------------------------------------------------

        ReDim strArrNewTmp(0)
        Dim varNewTmpString As Variant
        Dim intCol As Integer
        Dim intColTmp As Integer
        Dim intColFirst1 As Integer
        Dim intColFirst2 As Integer
        
        intColFirst1 = 0
        intColFirst2 = 0
        
        If intDrawGridRows > 0 Then
            
            ReDim strArrNewTmp(intDrawGridRows - 1) '初始化临时变量准备读取数据
            
            '要体温表录入项目个数进行循环
            For i = LBound(strArrRecordItemInfo) To UBound(strArrRecordItemInfo)
                '此处继续读取数组录入项目数组数据
                
                mstrSQL = "SELECT A.发生时间 As 当时产生的时间,C.记录内容 As 说明,E.保留项目,D.记录名,D.项目序号,D.记录频次 " & _
                            "FROM 病人护理记录 A,病人护理内容 C,体温记录项目 D,护理记录项目 E " & _
                            "Where A.ID=C.记录id " & _
                                "AND A.病人id=[1] " & _
                                "AND A.主页id=[2]  And Nvl(a.婴儿,0)=[7] " & _
                                "AND D.项目序号=C.项目序号 " & _
                                "AND C.记录类型=1 " & _
                                "AND E.项目序号=D.项目序号 And Nvl(E.应用方式,0)=1 And Nvl(e.适用病人,0) In (0,[8]) " & _
                                "AND E.护理等级>=0  " & _
                                "AND A.发生时间 BETWEEN [3] And [4] And c.终止版本 Is Null " & _
                                "AND D.记录法=2 AND D.记录名 In ([5],[6]) " & _
                            "Order By A.发生时间,Decode(D.记录名,'收缩压',0,1)"
                            
                If mblnMoved Then
                    mstrSQL = Replace(mstrSQL, "病人护理记录", "H病人护理记录")
                    mstrSQL = Replace(mstrSQL, "病人护理内容", "H病人护理内容")
                End If
                
                If CStr(Split(strArrRecordItemInfo(i), "'")(0)) = "血压" Then
                    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "mdlPrint", _
                                                        lng病人id, _
                                                        lng主页id, _
                                                        CDate(Format(strBeginDate, "YYYY-MM-DD")), _
                                                        CDate(Format(strEndDate, "YYYY-MM-DD") & " 23:59:59"), _
                                                        "收缩压", "舒张压", intBaby, IIf(intBaby = 0, 1, 2))
                Else
                    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "mdlPrint", _
                                                        lng病人id, _
                                                        lng主页id, _
                                                        CDate(Format(strBeginDate, "YYYY-MM-DD")), _
                                                        CDate(Format(strEndDate, "YYYY-MM-DD") & " 23:59:59"), _
                                                        CStr(Split(strArrRecordItemInfo(i), "'")(0)), "", intBaby, IIf(intBaby = 0, 1, 2))
                End If
                
                If CStr(Split(strArrRecordItemInfo(i), "'")(0)) = "呼吸" Then
                     '格子是42格
                     strNewTmpString = String(42, ";")
                 Else
                     strNewTmpString = String(14, ";")
                 End If
                
                Dim sgl排出量(1 To 14) As Single
                Dim sgl饮入量(1 To 14) As Single
                
                If rsTemp.BOF = False Then
                        
                    varNewTmpString = Split(strNewTmpString, ";")
                    
                    Do While Not rsTemp.EOF
                        
                        If CStr(Split(strArrRecordItemInfo(i), "'")(0)) = "呼吸" Then
                            
                            intCol = GetCurveColumn(CDate(rsTemp("当时产生的时间").Value), CDate(strTmpDay), mlngHourBegin)
                            
                            If intCol >= LBound(varNewTmpString) And intCol <= UBound(varNewTmpString) Then
                                varNewTmpString(intCol) = zlCommFun.NVL(rsTemp!说明)
                            End If
                        Else
                            
                            If Int((rsTemp!当时产生的时间 - Int(CDate(strTmpDay))) * 24) < 0 Then
                                intCol = 0
                            Else
                                intCol = 1 + Int((rsTemp!当时产生的时间 - Int(CDate(strTmpDay))) * 24) \ 12
                            End If
                            
                            If intCol >= 1 And intCol <= 14 Then
                                
                               Select Case rsTemp!项目序号
                                Case 7
                                    
                                    If rsTemp!记录频次 = 1 Then
                                        intColTmp = IIf(intCol Mod 2 = 0, intCol, intCol + 1)
                                    Else
                                        intColTmp = intCol
                                    End If
                                    
                                    sgl饮入量(intColTmp) = sgl饮入量(intColTmp) + Val(zlCommFun.NVL(rsTemp!说明))
                                    
                                    If Right(zlCommFun.NVL(rsTemp!说明), 2) = "/C" Then
                                    
                                        varNewTmpString(intCol) = sgl饮入量(intColTmp) & "/C"
                                        
                                    ElseIf Right(zlCommFun.NVL(rsTemp!说明), 1) = "C" Then
                                        varNewTmpString(intColTmp) = "C"
                                    Else
                                        varNewTmpString(intColTmp) = sgl饮入量(intColTmp)
                                    End If
                                Case 9
                                    
                                    If rsTemp!记录频次 = 1 Then
                                        intColTmp = IIf(intCol Mod 2 = 0, intCol, intCol + 1)
                                    Else
                                        intColTmp = intCol
                                    End If
                                    
                                    sgl排出量(intColTmp) = sgl排出量(intColTmp) + Val(zlCommFun.NVL(rsTemp!说明))
                                    If Right(zlCommFun.NVL(rsTemp!说明), 2) = "/C" Then
                                    
                                        varNewTmpString(intColTmp) = sgl排出量(intColTmp) & "/C"
                                        
                                    ElseIf Right(zlCommFun.NVL(rsTemp!说明), 1) = "C" Then
                                        varNewTmpString(intColTmp) = "C"
                                    Else
                                        varNewTmpString(intColTmp) = sgl排出量(intColTmp)
                                    End If
                                    
                                Case Else
                                    'varPatiInfo(3)
                                    
                                    Select Case rsTemp("记录名").Value
                                    Case "收缩压"
                                        
                                        If Format(rsTemp!当时产生的时间, "yyyy-MM-dd") = Format(varPatiInfo(3), "yyyy-MM-dd") Then
                                            If intColFirst1 = 0 Then
                                                varNewTmpString(intCol) = zlCommFun.NVL(rsTemp!说明)
                                                intColFirst1 = intCol
                                            ElseIf intColFirst1 <> intCol Then
                                                varNewTmpString(intCol) = zlCommFun.NVL(rsTemp!说明)
                                            End If
                                        Else
                                            intColFirst1 = intCol
                                            varNewTmpString(intCol) = zlCommFun.NVL(rsTemp!说明)
                                        End If
                                        
'                                        varNewTmpString(intCol) = rsTemp!说明
                                    Case "舒张压"
                                        
                                        If Format(rsTemp!当时产生的时间, "yyyy-MM-dd") = Format(varPatiInfo(3), "yyyy-MM-dd") Then
                                            
                                            If intColFirst2 = 0 Then
                                                If InStr(varNewTmpString(intCol), "/") > 0 Then
                                                    varNewTmpString(intCol) = varNewTmpString(intCol) & zlCommFun.NVL(rsTemp!说明)
                                                Else
                                                    varNewTmpString(intCol) = varNewTmpString(intCol) & "/" & zlCommFun.NVL(rsTemp!说明)
                                                End If

                                                intColFirst2 = intCol
                                            ElseIf intColFirst2 <> intCol Then
                                                If InStr(varNewTmpString(intCol), "/") > 0 Then
                                                    varNewTmpString(intCol) = varNewTmpString(intCol) & zlCommFun.NVL(rsTemp!说明)
                                                Else
                                                    varNewTmpString(intCol) = varNewTmpString(intCol) & "/" & zlCommFun.NVL(rsTemp!说明)
                                                End If
                                            End If
                                            
                                        Else
                                            intColFirst2 = intCol
                                            If InStr(varNewTmpString(intCol), "/") > 0 Then
                                                varNewTmpString(intCol) = varNewTmpString(intCol) & zlCommFun.NVL(rsTemp!说明)
                                            Else
                                                varNewTmpString(intCol) = varNewTmpString(intCol) & "/" & zlCommFun.NVL(rsTemp!说明)
                                            End If
                                        End If
                                        

                                        If varNewTmpString(intCol) = "/" Then varNewTmpString(intCol) = ""
                                    Case Else
                                        varNewTmpString(intCol) = zlCommFun.NVL(rsTemp!说明)
                                    End Select
                                    
                                End Select
                            End If
                        End If
                        
                        rsTemp.MoveNext
                    Loop
                    
                    strNewTmpString = Join(varNewTmpString, ";")
                    
                End If
                
                strArrNewTmp(i) = strNewTmpString
            Next
            
            
            For j = LBound(strArrRecordItemInfo) To UBound(strArrRecordItemInfo)
                If Split(strArrRecordItemInfo(j), "'")(1) <> "" Then
                    strArrNewTmp(j) = Split(strArrRecordItemInfo(j), "'")(0) & ";" & Split(strArrRecordItemInfo(j), "'")(2) & ";(" & Split(strArrRecordItemInfo(j), "'")(1) & ")" & strArrNewTmp(j)
                Else
                    strArrNewTmp(j) = Split(strArrRecordItemInfo(j), "'")(0) & ";" & Split(strArrRecordItemInfo(j), "'")(2) & ";" & strArrNewTmp(j)
                End If
            Next
            
        End If
        
'        y = y - 90
        
        intRepairRows = zlDatabase.GetPara("体温表格行数", glngSys, 1255, 8)
        
        Call DrawBodyRecordItem(objDraw, X, Y, mlngFirstWidth, strArrNewTmp, rsArrRecordItemInfo, strTmpString1, strTmpString2, sngScale, intRepairRows)
        
        '9、画出记录符说明
        Call DrawBodyTips(objDraw, lngNewTmpX, Y, mlngFirstWidth, strStateTips, sngScale)
        
        '10、最后画出页码，然后进入下一页
        Call DrawBodyPageFooter(objDraw, X, Y, intPageNo, intEndPage)
        
NOPageSub:                                 Next    '结束每一页的循环
        If blnPrint = False Then
            '如果是打印预览,应按打印机的可打印的开始处开始预览
            dblSureW = GetDeviceCaps(Printer.hDC, PHYSICALOFFSETX) / GetDeviceCaps(Printer.hDC, PHYSICALWIDTH)
            dblSureH = GetDeviceCaps(Printer.hDC, PHYSICALOFFSETY) / GetDeviceCaps(Printer.hDC, PHYSICALHEIGHT)
            objDraw.DrawStyle = 2
            On Error Resume Next
            objDraw.Line (objDraw.Width * dblSureW, objDraw.Height * dblSureH)-(objDraw.Width * (1 - dblSureW) - Screen.TwipsPerPixelX * sngScale, objDraw.Height * (1 - dblSureH) - Screen.TwipsPerPixelY * sngScale), &H808080, B
        End If
        Call ShowFlash
        
        Screen.MousePointer = 0
        PrintOrPreviewBodyState = True
        Exit Function
ErrPrint:
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
ErrExit:
        Call ShowFlash
        Screen.MousePointer = 0
        Err.Clear
        PrintOrPreviewBodyState = False
End Function

Public Function Ceil(ByVal dbValue As Double) As Integer
    '******************************************************************************************************************
    '功能： 转换时间为列值
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Ceil = (0 - Int(0 - dbValue))
    
End Function

'Private Function ConvertTimeToCol(ByVal strFrom As String, ByVal dtDate As Date, ByVal lngHourBegin As Long) As Long
'    '******************************************************************************************************************
'    '功能： 转换时间为列值
'    '参数：
'    '返回：
'    '******************************************************************************************************************
'
'    strFrom = Int(CDate(strFrom)) - (4 - lngHourBegin) / 24
'    ConvertTimeToCol = Int(DateDiff("h", CDate(strFrom), dtDate) / 4)
'
'End Function

Public Function InitDateTimeRange(ByRef varTime As Variant, Optional ByVal intHourBegin As Integer = 4) As Boolean
    '******************************************************************************************************************
    '功能：罗列体温单一天的曲线时间范围
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim strTmpBegin As String
    Dim strTmpEnd As String
    
    On Error GoTo errHand
    
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
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
        
        
End Function

Public Function GetCurveColumn(ByVal dtDateTime As Date, ByVal dtBeginDateTime As Date, Optional ByVal intHourBegin As Integer = 4) As Integer
    '******************************************************************************************************************
    '功能： 从时间计算出列
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim varTime As Variant
    Dim strTmp As String
    Dim intDays As Integer
    Dim intLoop As Integer
    
    On Error GoTo errHand
    
    GetCurveColumn = -1
    
    '初始化时间范围划分
    Call InitDateTimeRange(varTime, intHourBegin)

    '计算当前天的时间是在一天的第几格位置上
    strTmp = Format(dtDateTime, "HH:mm:ss")
    For intLoop = 0 To 6
        If strTmp >= Split(varTime(intLoop), ",")(0) And strTmp <= Split(varTime(intLoop), ",")(1) Then
            Exit For
        End If
    Next
    If intLoop < 7 Then
        
        '计算当天在当前体温单页上是第几天（0表示第一天；1表示第二天.....）
        intDays = DateDiff("d", Int(dtBeginDateTime), Int(dtDateTime))
        GetCurveColumn = intDays * 6 + intLoop + 1
    
    End If
    
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
            
End Function

Public Function GetCurveDateTime(ByVal intCol As Integer, ByVal dtBeginDateTime As Date, Optional ByVal intHourBegin As Integer = 4) As String
    '******************************************************************************************************************
    '功能： 从列算出时间
    '参数：
    '返回：返回当前列的时间范围，格式为："开始时间,结束时间"
    '******************************************************************************************************************
    Dim varTime As Variant
    Dim strTmp As String
    Dim intDays As Integer
    Dim strDay As String
    
    On Error GoTo errHand
    
    GetCurveDateTime = ""
    
    '初始化时间范围划分
    Call InitDateTimeRange(varTime, intHourBegin)
        
    intDays = intCol \ 6
    intCol = (intCol Mod 6)
    If intCol = 0 Then
        intCol = 6
        If intDays >= 1 Then intDays = intDays - 1
    End If
    
    '计算所在的日期
    If intCol >= 1 And intCol <= 6 Then
        strDay = Format(DateAdd("d", intDays, Int(dtBeginDateTime)), "yyyy-MM-dd")
        strTmp = strDay & " " & Split(varTime(intCol - 1), ",")(0)
        strTmp = strTmp & "," & strDay & " " & Split(varTime(intCol - 1), ",")(1)
    End If
    
    '加上时间
    GetCurveDateTime = strTmp
    
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Function

Public Function GetEditDateTime(ByVal intCol As Integer, ByVal dtBeginDateTime As Date) As String
    '******************************************************************************************************************
    '功能： 从列算出时间
    '参数：
    '返回：返回当前列的时间范围，格式为："开始时间,结束时间"
    '******************************************************************************************************************
    Dim strTmp As String
    Dim intDays As Integer
    Dim strDay As String
    
    On Error GoTo errHand
    
    GetEditDateTime = ""
        
    intDays = intCol \ 2
    intCol = (intCol Mod 2)
    
    If intCol = 0 Then
        If intDays >= 1 Then intDays = intDays - 1
    End If
    
    strDay = Format(DateAdd("d", intDays, Int(dtBeginDateTime)), "yyyy-MM-dd")
    If intCol = 1 Then
        strTmp = strDay & " 00:00:00"
        strTmp = strTmp & "," & strDay & " 11:59:59"
    ElseIf intCol = 0 Then
        strTmp = strDay & " 12:00:00"
        strTmp = strTmp & "," & strDay & " 23:59:59"
    End If
    
    GetEditDateTime = strTmp

    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Function

'=========================================
Public Sub ClearSpecRowCol(obj As Object, ByVal intRow As Integer, Optional intCol As Variant)
    '功能: 清除指定网格的指定行指定列的数据
    '参数: obj=要操作的网格控件
    '      intRow=要清除的行号
    '      intCol=要清除的列号列表如Array(1,2,3),若所有列则可以表示为Array()
    Dim i As Long
    If UBound(intCol) = -1 Then
        For i = 0 To obj.Cols - 1
            obj.TextMatrix(intRow, i) = ""
        Next
    Else
        For i = 0 To UBound(intCol)
            obj.TextMatrix(intRow, intCol(i)) = ""
        Next
    End If
    obj.RowData(intRow) = 0
End Sub

Public Sub SetColumnText(fgd As Object, intRow As Integer, ByVal varColText As Variant)
    '功能: 设置指定网格控件的列头文本
    '参数: fgd=网格控件
    '      intRow=行号
    '      varColText=列头文本数组
    Dim i As Integer
    For i = 0 To fgd.Cols - 1
        fgd.TextMatrix(intRow, i) = varColText(i)
    Next
End Sub

Public Sub SetColAlignment(fgd As Object, varColAlignment As Variant)
    '功能: 设置指定网格控件的列对齐方式
    '参数: fgd=网格控件
    '      varColAlignment=列对齐方式数组
    Dim i As Long
    For i = 0 To UBound(varColAlignment)
        fgd.ColAlignment(i) = varColAlignment(i)
    Next
End Sub

Public Sub SetColData(fgd As Object, varColData As Variant)
    '功能: 设置指定网格控件的列数据来源方式
    '参数: fgd=网格控件
    '      varColData=列数据来源方式数组
    Dim i As Long
    For i = 0 To UBound(varColData)
        fgd.ColData(i) = varColData(i)
    Next
End Sub

Public Sub SetFixColAlignment(fgd As Object, varFixColAlignment As Variant)
    '功能: 设置指定网格控件的固定列对齐方式
    '参数: fgd=网格控件
    '      varColAlignment=固定列对齐方式数组
    Dim i As Long
    For i = 0 To UBound(varFixColAlignment)
        fgd.ColAlignmentFixed(i) = varFixColAlignment(i)
    Next
End Sub

Public Sub SetColumnWidth(fgd As Object, ByVal varColWidth As Variant)
    '功能: 设置指定网格控件的列宽
    '参数: fgd=网格控件
    '      varColWidth=列宽数组
    Dim i As Integer
    For i = 0 To fgd.Cols - 1
        fgd.ColWidth(i) = varColWidth(i)
    Next
End Sub

Public Sub SetRowForeColor(mshObject As Object, ByVal lngRow As Long, ByVal lngColor As Long)
    Dim i As Integer
    Dim blnPre As Boolean
    Dim intRow As Integer
    Dim intCol As Integer
    
    With mshObject
        blnPre = .Redraw
        intRow = .Row
        intCol = .Col
        .Redraw = False
        .Row = lngRow
        For i = 0 To .Cols - 1
            .Col = i
            .CellForeColor = lngColor
        Next
        
        .Row = intRow
        .Col = intCol
        .Redraw = blnPre
    End With
End Sub

Public Sub CalcXY(objFrm As Object, objMSH As Object, objX As Single, objY As Single, sglX As Single, sglY As Single)
    sglX = objFrm.Left + objX + objMSH.CellLeft + Screen.TwipsPerPixelX
    sglY = objFrm.Top + objFrm.Height - objFrm.ScaleHeight + objY + objMSH.CellTop + objMSH.CellHeight
    If sglX + 5895 > Screen.Width Then
        sglX = Screen.Width - 5895
    End If
    If sglY + 3420 > Screen.Height Then
        sglY = sglY - objMSH.CellHeight - 3420
    End If
End Sub


Public Function Check是否包含(strSource As String, strTarge As String) As Boolean
    '检查strSource中的每一个字符是否在strTarge中
    Dim i As Long
    Check是否包含 = False
    
    Select Case strTarge
    Case "整数"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "小数"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "正整数"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+-_)(*&^%$#@!`~"
    Case "正小数"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/'"":;|\=+-_)(*&^%$#@!`~"
    End Select
    For i = 1 To Len(strSource)
        If InStr(strTarge, Mid(strSource, i, 1)) <= 0 Then Exit Function
    Next
    Check是否包含 = True
End Function

Public Sub SelectRow(mshObject As Object, Optional ByVal BackColor As Long = &H8000000D, Optional ByVal ForeColor As Long = &H8000000E)
    Dim i As Integer
    Dim blnPre As Boolean
    Dim intRow As Integer
    Dim intCol As Integer
    
    With mshObject
        blnPre = .Redraw
        intRow = .Row
        intCol = .Col
        .Redraw = False
        
        For i = 0 To .Cols - 1
            .Col = i
            .CellBackColor = BackColor
            .CellForeColor = ForeColor
        Next
        
        .Row = intRow
        .Col = intCol
        .Redraw = blnPre
    End With
End Sub

Public Sub UnSelectRow(mshObject As Object, Optional lngColorSave As Long = 0)
    Dim i As Integer
    Dim blnPre As Boolean
    
    With mshObject
        blnPre = .Redraw
        .Redraw = False
        
        For i = 0 To .Cols - 1
            .Col = i
            .CellBackColor = .BackColor
            .CellForeColor = lngColorSave
        Next
        .Redraw = blnPre
    End With
End Sub


Public Sub Hook(ByVal hWnd As Long)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    glngPrevWndProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)

    '获取"控制面板"中的滚动行数值

    Call SystemParametersInfo(SPI_GETWHEELSCROLLLINES, 0, WHEEL_SCROLL_LINES, 0)

    If WHEEL_SCROLL_LINES > frmCaseTendBody.BodyEdit.ScrollBarY.Max Then WHEEL_SCROLL_LINES = frmCaseTendBody.BodyEdit.ScrollBarY.Max
End Sub

Public Sub UnHook(ByVal hWnd As Long)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngReturnValue As Long

    lngReturnValue = SetWindowLong(hWnd, GWL_WNDPROC, glngPrevWndProc)
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
    
        ScreenToClient frmCaseTendBody.hWnd, pt

        With frmCaseTendBody.BodyEdit
        
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

Public Function ShowTxtSelDialog(ByVal frmParent As Object, _
                                    ByVal objTXT As Object, _
                                    ByVal strLvw As String, _
                                    ByVal strSavePath As String, _
                                    ByVal strDescrible As String, _
                                    ByVal rsData As ADODB.Recordset, _
                                    ByRef rsResult As ADODB.Recordset, _
                                    Optional ByVal lngCX As Long = 9000, _
                                    Optional ByVal lngCY As Long = 4500, _
                                    Optional blnMuliSel As Boolean = False, _
                                    Optional strInitKey As String = "", _
                                    Optional ByVal WinStyle As Byte = 3, _
                                    Optional ByVal blnSort As Boolean = True) As Boolean
    '******************************************************************************************************************
    '功能:打开树型+列表结构
    '返回:出错返回2;成功返回1;取消返回0
    '******************************************************************************************************************
    
    Dim lngX As Long
    Dim lngY As Long
    Dim objPoint As POINTAPI
    Dim lngObjHeight As Long
    
    On Error GoTo errHand
    
    If rsData.BOF Then Exit Function
    
    If objTXT Is Nothing Then
        '屏幕居中
        
        lngX = (Screen.Width - lngCX) / 2
        lngY = (Screen.Height - lngCY) / 2
        
    Else
        Call ClientToScreen(objTXT.hWnd, objPoint)
                    
        lngX = objPoint.X * Screen.TwipsPerPixelX - Screen.TwipsPerPixelX
        lngY = objTXT.Height + objPoint.Y * Screen.TwipsPerPixelY - Screen.TwipsPerPixelY
        
        lngObjHeight = objTXT.Height
    End If
    
    If frmSelectDialog.ShowSelect(frmParent, WinStyle, rsData, strLvw, strDescrible, lngX, lngY, lngCX, lngCY, lngObjHeight, strInitKey, strSavePath, , False, blnMuliSel, , blnSort) Then
                            
        Set rsResult = rsData
        ShowTxtSelDialog = True
        
    End If
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function SQLRecord(ByRef rs As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Set rs = New ADODB.Recordset
    
    With rs
        
        .Fields.Append "SQL", adVarChar, 300
        .Fields.Append "Trans", adTinyInt                   '1表示开始;2表示结束
        .Fields.Append "Custom", adTinyInt
        .Fields.Append "Parameter", adVarChar, 500
        
        .Open
    End With
    
    SQLRecord = True
    
    Exit Function
    
errHand:
    
End Function

Public Function SQLRecordAdd(ByRef rs As ADODB.Recordset, ByVal strSQL As String, Optional ByVal intTrans As Integer = 0, Optional ByVal intCustom As Integer = 0, Optional ByVal strParameter As String = "") As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    On Error GoTo errHand
    
    rs.AddNew
    rs("SQL").Value = strSQL
    rs("Trans").Value = intTrans
    rs("Custom").Value = intCustom
    rs("Parameter").Value = strParameter
    SQLRecordAdd = True
    
    Exit Function
    
errHand:
End Function

Public Function SQLRecordExecute(ByVal rs As ADODB.Recordset, Optional ByVal strTitle As String, Optional ByVal blnHaveTrans As Boolean = True) As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim blnTran As Boolean
    Dim intLoop As Integer
    Dim strSQL As String
    
    On Error GoTo errHand
    
    If rs.RecordCount > 0 Then
        If Len(strTitle) = 0 Then strTitle = gstrSysName
        blnTran = True
        
        If blnHaveTrans Then gcnOracle.BeginTrans
        
        rs.MoveFirst
    
        For intLoop = 1 To rs.RecordCount
            
            If Val(rs("Custom").Value) = 0 Then
                strSQL = CStr(rs("SQL").Value)
                Call zlDatabase.ExecuteProcedure(strSQL, strTitle)
            End If
            
            rs.MoveNext
        Next
    
        If blnHaveTrans Then gcnOracle.CommitTrans
        blnTran = False
    End If
    
    SQLRecordExecute = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
    If blnTran And blnHaveTrans Then gcnOracle.RollbackTrans
End Function

Public Function SQLRecordSavePicture(ByVal rs As ADODB.Recordset, Optional ByVal strTitle As String, Optional ByVal blnHaveTrans As Boolean = True) As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim blnTran As Boolean
    Dim intLoop As Integer
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim strTmp As String
    Dim aryTmp As Variant
    Dim strFile As String
    
    On Error GoTo errHand
    
    If rs.RecordCount > 0 Then
        If Len(strTitle) = 0 Then strTitle = gstrSysName
        blnTran = True
        
        If blnHaveTrans Then gcnOracle.BeginTrans
        
        rs.MoveFirst
    
        For intLoop = 1 To rs.RecordCount
            
            If Val(rs("Custom").Value) = 100 Then
                
                strTmp = rs("Parameter").Value
                aryTmp = Split(strTmp, ";")
                
                If UBound(aryTmp) >= 2 Then
                    If Dir(CStr(aryTmp(2))) <> "" And CStr(aryTmp(2)) <> "" Then
                        
                        Call zlBlobSave(Val(aryTmp(0)), Val(aryTmp(1)), CStr(aryTmp(2)))

                    End If
                End If
                
            End If
            
            rs.MoveNext
        Next
    
        If blnHaveTrans Then gcnOracle.CommitTrans
        blnTran = False
    End If
    
    SQLRecordSavePicture = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
    If blnTran And blnHaveTrans Then gcnOracle.RollbackTrans
End Function

Public Function DrawPicture(objDraw As Object, ByVal strFile As String, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, Optional ByVal bln资源 As Boolean = False) As Boolean
    '******************************************************************************************************************
    '功能：根据相册大小自动等比例缩放相片文件
    '参数：缩放前的相片文件
    '返回：缩放后的相片文件
    '******************************************************************************************************************
    Dim strTmp As String
    Dim objMap As StdPicture
    Dim W As Single
    Dim H As Single
    Dim sglPerW As Single
    Dim sglPerH As Single
    Dim sglPer As Single
    Dim cx As Long
    Dim cy As Long
    
    On Error GoTo errHand
    
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
'    objDraw.PaintPicture objMap, X1, Y1, W, H, vbSrcAnd
    
    DrawPicture = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If bln资源 Then
        ShowSimpleMsg "未找到该资源(" & strFile & "),可能该资源不存在!"
    Else
        ShowSimpleMsg "不能打开文件(" & strFile & "),该文件可能正在使用或文件不存在!"
    End If
End Function

'######################################################################################################################

Public Function GetGridItem(ByVal byt护理等级 As Byte, ByVal lng科室id As Long, ByVal byt婴儿 As Byte, Optional ByVal byt项目性质 As Byte = 1, Optional ByVal strNotItem As String, Optional ByVal blnBodyItem As Boolean = True) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：
    '参数：strNotItem：序号1,序号2,序号3,......
    '返回：
    '******************************************************************************************************************
    
    If blnBodyItem Then
        mstrSQL = " Select A.排列序号,A.记录频次,A.记录法,c.项目名称,c.项目序号,c.项目类型,c.项目表示,c.项目值域,c.项目单位,A.记录名 as 名称,A.项目序号 as 项目号,A.项目序号 As ID,Nvl(B.ID,0) as 项目ID,c.项目性质,C.保留项目,1 As 末级," & _
                    " C.项目单位 As 单位,记录符,最小值,最大值,记录色,1 as 记录否,单位值,最高行,Nvl(C.项目类型,1) as 存储类型,c.项目长度,c.项目小数,c.分组名 " & _
                    " From 体温记录项目 A,诊治所见项目 B,护理记录项目 C " & _
                    " Where C.项目ID=B.ID(+) And A.项目序号=C.项目序号 And A.记录法=2 And c.项目性质=[4] And Nvl(C.应用方式,0)=1 And C.护理等级>=[1] And Nvl(C.适用病人,0) In (0,[3]) " & _
                    " And (c.适用科室=1 Or (c.适用科室=2 And Exists (Select 1 From 护理适用科室 D Where D.项目序号=c.项目序号 And D.科室id=[2]))) "
                        
        If strNotItem <> "" Then mstrSQL = mstrSQL & " And A.项目序号 Not In (" & strNotItem & ")"
        mstrSQL = mstrSQL & " Order by A.排列序号"
    Else
    
        mstrSQL = " Select c.项目名称,c.项目序号,c.项目类型,c.项目表示,c.项目值域,c.项目单位,c.项目名称 as 名称,c.项目序号 as 项目号,c.项目序号 As ID,Nvl(B.ID,0) as 项目ID,c.项目性质,C.保留项目,1 As 末级," & _
                    " C.项目单位 As 单位,1 as 记录否,Nvl(C.项目类型,1) as 存储类型,c.项目长度,c.项目小数,c.分组名 " & _
                    " From 诊治所见项目 B,护理记录项目 C " & _
                    " Where C.项目ID=B.ID(+) And c.项目性质=[4] And Nvl(C.应用方式,0)=1 And C.护理等级>=[1] And Nvl(C.适用病人,0) In (0,[3]) " & _
                    " And (c.适用科室=1 Or (c.适用科室=2 And Exists (Select 1 From 护理适用科室 D Where D.项目序号=c.项目序号 And D.科室id=[2]))) "
                        
        If strNotItem <> "" Then mstrSQL = mstrSQL & " And c.项目序号 Not In (" & strNotItem & ")"
    
    
    End If
    
    Set GetGridItem = zlDatabase.OpenSQLRecord(mstrSQL, "mdlPrint", byt护理等级, lng科室id, byt婴儿, byt项目性质)
    
End Function

Public Function GetGridDataItem(ByVal byt护理等级 As Byte, _
                                ByVal lng科室id As Long, _
                                ByVal byt适用病人 As Byte, _
                                ByVal lng病人id As Long, _
                                ByVal lng主页id As Long, _
                                ByVal dt开始时间 As Date, _
                                ByVal dt结束时间 As Date, ByVal byt婴儿 As Byte, Optional ByVal blnMoved As Boolean) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：获取有数据的活动护理项目
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strSQL As String
    
    
    mstrSQL = " Select A.排列序号,A.记录频次,A.记录法,A.记录名 as 名称,A.项目序号 as 项目号,A.项目序号 As ID,Nvl(B.ID,0) as 项目ID,c.项目性质,C.保留项目,1 As 末级," & _
                " C.项目单位 As 单位,记录符,最小值,最大值,记录色,1 as 记录否,单位值,最高行,Nvl(C.项目类型,1) as 存储类型,c.项目长度,c.项目小数,c.分组名 " & _
                " From 体温记录项目 A,诊治所见项目 B,护理记录项目 C " & _
                " Where C.项目ID=B.ID(+) And A.项目序号=C.项目序号 And A.记录法=2 And c.项目性质=2 And Nvl(C.应用方式,0)=1 And C.护理等级>=[1] And Nvl(C.适用病人,0) In (0,[3]) " & _
                " And (c.适用科室=1 Or (c.适用科室=2 And Exists (Select 1 From 护理适用科室 D Where D.项目序号=c.项目序号 And D.科室id=[2])))"
                
    strSQL = "Select e.项目序号 " & _
                "FROM 病人护理记录 A,病人护理内容 C,体温记录项目 D,护理记录项目 E " & _
                "Where A.ID = c.记录ID " & _
                    "AND A.病人来源=2 " & _
                    "AND Nvl(a.婴儿,0)=[8] " & _
                    "AND a.病人id=[4] " & _
                    "AND a.主页id=[5] " & _
                    "AND d.项目序号=C.项目序号 " & _
                    "AND c.记录类型=1 And E.项目性质=2 " & _
                    "AND E.项目序号=D.项目序号 " & _
                    "AND E.护理等级>=[1]  " & _
                    "AND a.发生时间 BETWEEN [6] And [7] And c.终止版本 Is Null " & _
                    "AND d.记录法=2"
                                            
    If blnMoved Then
        strSQL = Replace(strSQL, "病人护理记录", "H病人护理记录")
        strSQL = Replace(strSQL, "病人护理内容", "H病人护理内容")
    End If
    
    mstrSQL = mstrSQL & " And c.项目序号 In (" & strSQL & ")"
    mstrSQL = mstrSQL & " Order by A.排列序号"
    
    Set GetGridDataItem = zlDatabase.OpenSQLRecord(mstrSQL, "mdlPrint", byt护理等级, lng科室id, byt适用病人, lng病人id, lng主页id, dt开始时间, dt结束时间, byt婴儿)
    
End Function

Public Function ExistsPrinter() As Boolean
    Dim lngHDc As Long
    
    If Printers.Count = 0 Then Exit Function
    
    On Error Resume Next
    lngHDc = Printer.hDC
    If Err.Number = 0 Then ExistsPrinter = True
    Err.Clear: On Error GoTo 0
End Function

Private Function GetFontSize(ByVal objDraw As Object, ByVal dblHeight As Double, ByVal strText As String, ByRef Y1 As Single) As Single
    Dim sinFontSize As Single
    Dim sinFontSize_Bak As Single
    Dim intCharNumber As Integer
    Dim intCount As Integer
    Dim strChar As String
    '计算合理的字体大小
    
    sinFontSize_Bak = objDraw.FontSize
    For sinFontSize = objDraw.FontSize To 5 Step -1
        Y1 = 0
        intCharNumber = 0
        For intCount = 1 To Len(strText)
            strChar = Mid(strText, intCount, 1)
            
            If Asc(strChar) < 0 Then
                If intCharNumber Mod 2 = 1 Then Y1 = Y1 + ROWHEIGHT * 2.5
            End If
            
            If Asc(strChar) < 0 Then
                intCharNumber = 0
                Y1 = Y1 + ROWHEIGHT * 5
            Else
                Y1 = Y1 + ROWHEIGHT * 2.5
                intCharNumber = intCharNumber + 1
            End If
        Next
        'If Y1 <= dblHeight + 100 Then Exit For
        Exit For            '不管超不超过了,为了得到高度,所以该函数的总体代码还是保留下来
    Next
    
    objDraw.FontSize = sinFontSize_Bak
    GetFontSize = sinFontSize
End Function

Private Sub OutputNote(ByVal objDraw As Object, ByVal dblHeight As Double, ByRef rsNote As ADODB.Recordset, ByVal lngTmpX As Long, ByVal lngTmpY As Long)
    '输出以下信息:入院,入科,转科,出院,手术分娩,未记说明,上标说明及出生
    '未记说明及上标说明,在没有入出转手术分娩及出生的信息时,打印在42-40之间;否则从40开始向下打印
    '除未记说明及上标说明外,入出转等信息当一个刻度发生多个时,依次写入各个刻度中,如其它刻度也有信息,顺移
    Dim intCol As Integer                   '记录当前列号
    Dim intMax As Integer                   '总列数
    Dim intCur As Integer                   '当前记录的位置
    Dim bln上标 As Boolean
    Dim sinX1 As Single, sinY1 As Single, sinHeight As Single, H_9pt As Single, sinMaxY1 As Single
    Dim rsTarget As New ADODB.Recordset

    '输出字符相关变量定义
    Dim sinFontSize As Single
    Dim sinFontSize_Bak As Single
    Dim intCharNumber As Integer
    Dim intCount As Integer
    Dim strChar As String

    intMax = 41
    H_9pt = ROWHEIGHT * 10 / 3
    sinFontSize_Bak = objDraw.FontSize
    Set rsTarget = rsNote.Clone
    With rsNote
        If .RecordCount = 0 Then Exit Sub
        .Sort = "列号,时间"
        intCol = !列号

        '先在入出转手术等中循环
        Do While Not .EOF
            If Trim(NVL(!结果)) <> "" Then
                If !类型 = 1 Then   '入出转手术出生
                    '检查待打印列是否已存在输出,如果存在则校正坐标
                    If intCol > intMax Then intCol = intMax

                    '计算得到合适的字体大小及高度
                    !字体大小 = GetFontSize(objDraw, dblHeight, NVL(!结果), sinY1)
                    !高度 = sinY1
                    !打印列 = IIf(intCol < !列号, !列号, intCol)
                    .Update
                    If intCol <= !列号 Then intCol = !列号
                    intCol = intCol + 1
                Else        '上标说明,未记说明
                    Call GetFontSize(objDraw, dblHeight, NVL(!结果), sinY1)
                    !高度 = sinY1
                    .Update
                End If
            End If

            .MoveNext
        Loop
        .MoveFirst

        '调整入出转等的纵坐标(只有最后一列才存在一格打完的情况)
        sinY1 = (lngTmpY + 3 * H_9pt / 2)       '42度
        .Filter = "打印列='" & intMax & "'"
        .Sort = "列号,时间"
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            '只有入出转手术才更新了打印列
            !坐标 = Split(!坐标, ";")(0) & ";" & sinY1
            .Update
            sinY1 = sinY1 + !高度 + 100

            .MoveNext
        Loop
        .Filter = 0
        .MoveFirst

        '重新校正未记说明以及上标说明的高度(未记说明及上标说明,在没有入出转手术分娩及出生的信息时,打印在42-40之间;否则从40开始向下打印)
        Set rsTarget = .Clone
        intCol = 0
        Do While Not .EOF
            If !类型 = 2 Then       '上标说明,未记说明
                Set rsTarget = .Clone
                rsTarget.Filter = "打印列='" & !列号 & "'"
                If rsTarget.RecordCount <> 0 Then
                    '已存在打印内容的才校正纵坐标
                    sinMaxY1 = Split(rsTarget!坐标, ";")(1)
                    Do While Not rsTarget.EOF
                        If bln上标 = False Then
                            bln上标 = (rsTarget!类型 = 2)
                        End If
                        sinMaxY1 = sinMaxY1 + rsTarget!高度 + 100
                        rsTarget.MoveNext
                    Loop
                    sinY1 = (lngTmpY + 3 * H_9pt / 2) + 10 * 3 * H_9pt / 2 - H_9pt / 2      '40度的坐标
                    If sinY1 < sinMaxY1 Or bln上标 Then sinY1 = sinMaxY1
                    sinHeight = !高度
                    intCol = !列号
                Else
                    sinY1 = (lngTmpY + 3 * H_9pt / 2)       '42度
                    intCol = !列号
                    sinHeight = !高度
                End If
                rsTarget.Filter = 0

                !坐标 = Split(!坐标, ";")(0) & ";" & sinY1
                !打印列 = !列号                                 '此时更新打印列,以便上面的循环过滤
                .Update
            End If
            .MoveNext
        Loop

        '开始按数据输出内容
        .MoveFirst
        Do While Not .EOF
            If Trim(NVL(!结果)) <> "" Then
                'If (!类型 = 2) Then Stop
                sinX1 = lngTmpX + (IIf(!打印列 = "", Val(!列号), Val(!打印列))) * HOUR_STEP_Twips + HOUR_STEP_Twips / 2
                sinY1 = Split(!坐标, ";")(1)
                intCharNumber = 0
                objDraw.FontSize = IIf(!字体大小 = "", 9, !字体大小)

                For intCount = 1 To Len(!结果)
                    strChar = Mid(!结果, intCount, 1)

                    If Asc(strChar) < 0 Then
                        If intCharNumber Mod 2 = 1 Then sinY1 = sinY1 + ROWHEIGHT * 2.5
                    End If
                    Call DrawRotateText(objDraw, sinX1 - objDraw.TextWidth(strChar) / 2, sinY1 + 15, strChar, vbRed)
                    If Asc(strChar) < 0 Then
                        intCharNumber = 0
                        sinY1 = sinY1 + ROWHEIGHT * 5
                    Else
                        sinY1 = sinY1 + ROWHEIGHT * 2.5
                        intCharNumber = intCharNumber + 1
                    End If
                Next
            End If

            .MoveNext
        Loop
    End With
    objDraw.FontSize = sinFontSize_Bak
End Sub

Private Function SimplifyString(ByVal strText As String) As String
    Dim arrData
    Dim strData As String
    Dim intLen As Integer, intActLen As Integer
    Dim intCol As Integer, intCount As Integer
    '精简字符串,去掉重复的内容
    
    arrData = Split(strText, " ")
    intCount = UBound(arrData)
    strData = ""
    For intCol = 0 To intCount
        If InStr(1, " " & strData & " ", " " & arrData(intCol) & " ") = 0 Then
            strData = strData & " " & arrData(intCol)
        End If
    Next
    SimplifyString = Mid(strData, 2)
End Function

Public Sub GetValue(strValue As String)
    Dim str摄氏 As String
    Dim str华氏 As String
    
    str摄氏 = strValue & "°"
    str华氏 = CStr(Val(strValue) * 9 / 5 + 32) & "°"
    strValue = str华氏 & String(10 - Len(str华氏 & str摄氏), " ") & str摄氏
End Sub
