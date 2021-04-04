VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmTendFilePreview 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "preView"
   ClientHeight    =   5130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8400
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form24"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtLength 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   1005
      Left            =   90
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   90
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1995
      LargeChange     =   10
      Left            =   3540
      Max             =   100
      SmallChange     =   2
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   285
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2685
      Left            =   90
      ScaleHeight     =   2655
      ScaleWidth      =   3225
      TabIndex        =   0
      Top             =   630
      Width           =   3255
      Begin VSFlex8Ctl.VSFlexGrid VsfData 
         Height          =   1455
         Left            =   570
         TabIndex        =   5
         Top             =   930
         Width           =   2265
         _cx             =   3995
         _cy             =   2566
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16764057
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   4
         FixedRows       =   3
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   5000
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmTendFilePreview.frx":0000
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   1
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   -1  'True
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   0   'False
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblDownTable 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "表下项可换行"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   270
         TabIndex        =   4
         Top             =   1020
         Width           =   1125
      End
      Begin VB.Label lblUpTable 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "表上项可换行"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   270
         TabIndex        =   3
         Top             =   600
         Width           =   1125
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "标题栏"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1380
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.Line lineRight 
         X1              =   1380
         X2              =   1380
         Y1              =   360
         Y2              =   2220
      End
      Begin VB.Line lineLeft 
         X1              =   720
         X2              =   720
         Y1              =   360
         Y2              =   2220
      End
      Begin VB.Line lineBottom 
         X1              =   630
         X2              =   2790
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line lineTop 
         X1              =   630
         X2              =   2790
         Y1              =   600
         Y2              =   600
      End
   End
End
Attribute VB_Name = "frmTendFilePreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lngRows As Long
Private arrFormat
Private dblTitle As Double      '标题栏的高度
Private dblUpTable As Double    '表上项的高度
Private dblDownTable As Double  '表下项的高度

Private mlngFile As Long                 '病人护理文件.ID
Private mlngFormat As Long               '格式ID
Private mlngRows As Long
Private mstrSQL As String
Private mstrSQL中 As String
Private mstrSQL内 As String
Private mstrSQL列 As String
Private mstrSQL条件 As String
'活动项目SQL
Private mstrSQLActive中 As String
Private mstrSQLActive内 As String
Private mstrSQLActive列 As String
Private mstrSQLActive条件 As String

Private mobjParent As Object

'病历文件格式定义相关
Private mintTabTiers As Integer     '表头层次
Private mintTagFormHour As Integer  '开始时间条件
Private mintTagToHour As Integer    '截止时间条件
Private mobjTagFont As New StdFont  '条件样式字体
Private mlngTagColor As Long        '条件样式颜色
Private mstrPaperSet As String      '格式
Private mstrPageHead As String      '页眉
Private mstrPageFoot As String      '页脚
Private mblnChildForm As Boolean
Private mlngActiveRows As Long      '有效数据行
Private mstrSubhead As String       '表上标签
Private mstrTabHead As String       '表头单元
Private mstrPreHead As String       '需处理的列,文本型项目所属列或绑定多个项目的列
Private Const mlngFixedCOL As Long = 2 '固定绑定的列,目前只绑定了汇总类别和记录ID
Private mstrActivePreHead As String '需处理的活动项目
Private mlngActiveColCount As Long '有效的活动项目列数
Private mstrColWidth As String      '列宽序列串
Private mstrColumns As String       '当前护理文件各列对应的项目
Private lngCurColor As Long, strCurFont As String, objFont As StdFont
Private mrsItems As New ADODB.Recordset


Private Const conLineWide As Integer = 30        '横线所占宽度(单位为缇)占两条线宽度
Private Const conLineHigh As Integer = 30        '竖线所占高度(单位为缇)占两条线高度
Private Const conRatemmToTwip As Single = 56.6857142857143      '毫米与缇的比率
Private Type OSVERSIONINFO 'for GetVersionEx API call
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

'WinNT自定义纸张控制================================================================
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
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal hWnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, pDevModeOutput As Any, pDevModeInput As Any, ByVal fMode As Long) As Long
Private Declare Function ResetDC Lib "gdi32" Alias "ResetDCA" (ByVal hDC As Long, lpInitData As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Private Function zlGetPrinterSet() As Boolean
    '------------------------------------------------
    '功能：读取本系统注册表的打印缺省设置
    '------------------------------------------------
    Dim iCount As Long
    Dim strDeviceName As String
    Dim intPaperSize As Integer
    Dim intPaperBin As Integer
    Dim intOrientation As Long
    
    If Printers.Count = 0 Then
        zlGetPrinterSet = False
        Exit Function
    End If
    
    strDeviceName = GetSetting("ZLSOFT", "公共模块\" & "zl9PrintMode" & "\Default", "DeviceName", Printer.DeviceName)
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
    Printer.PaperBin = GetSetting("ZLSOFT", "公共模块\" & "zl9PrintMode" & "\Default", "PaperBin", Printer.PaperBin)
    Printer.Orientation = arrFormat(1)
    
    intPaperSize = arrFormat(0)
    If intPaperSize = 256 Then
        Dim lngWidth As Long
        Dim lngHeight As Long
        
        lngWidth = arrFormat(3)
        lngHeight = arrFormat(2)
        
        Call SetCustonPager(lngWidth, lngHeight)
    Else
        Printer.PaperSize = intPaperSize
    End If

    zlGetPrinterSet = True
End Function

Private Function SetCustonPager(ByVal lngWidth As Long, ByVal lngHeight As Long) As Integer
'功能：在设置自定义纸张
'参数：是以绨为单位
    If IsWindowsNT Then
        '虽然不能使宽度生效，但能改变PaperSize的属性值
        Printer.Width = lngWidth
        Printer.Height = lngHeight
        SetCustonPager = SetNTPrinterPaper(Me.hWnd, lngWidth / conRatemmToTwip, lngHeight / conRatemmToTwip, Printer.Orientation, Printer.Copies)
    Else
        'Windows98系列还是以通常方法处理
        Printer.PaperSize = 256
        Printer.Width = lngWidth
        Printer.Height = lngHeight
    End If
End Function

Private Function IsWindowsNT() As Boolean
'功能：是否WindowNT操作系统
    Const dwMaskNT = &H2&
    IsWindowsNT = (GetWinPlatform() And dwMaskNT)
End Function

Private Function IsWindows95() As Boolean
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

Private Function SetNTPrinterPaper(ByVal lngHwnd As Long, ByVal intWidth As Integer, ByVal intHeight As Integer, _
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


Private Sub Form_Load()
    Dim lngFixRows As Long                          '固定行数
    Dim dblRowHeight As Double                      '行高
    Dim lngParent As Long
    Dim strUpText As String
    Dim lngHeight As Long, lngWidth As Long         '有效高度，宽度
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    Dim lngOffsetLeft As Long, lngOffsetTop As Long
    Dim rsTemp As New ADODB.Recordset
    Dim arrHead() As String   '表头内容
    Dim arrData() As String, arrColWith() As String
    Dim lngMutilRow1 As Long, lngMutilRow2 As Long, lngMutilRow3 As Long
    Dim i As Integer
    Dim lngFixRowsheight As Long
    On Error GoTo errHand
    'arrFormat(纸张|纸向|高|宽|上边距|下边距|左边距|右边距|行高|固定行数|标题栏字体名|标题栏字体大小|标题文本|表上项字体名|表上项字体大小|表上项文本|表头项目内容)
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = Screen.Height
    Me.Width = Screen.Width
    
    '设置页面格式
    Call zlGetPrinterSet
    
    '获取打印机当前状态
    picDraw.Height = Printer.Height
    picDraw.Width = Printer.Width
    lngOffsetLeft = Printer.ScaleX(GetDeviceCaps(Printer.hDC, PHYSICALOFFSETX), vbPixels, vbTwips)
    lngOffsetTop = Printer.ScaleY(GetDeviceCaps(Printer.hDC, PHYSICALOFFSETY), vbPixels, vbTwips)
    picDraw.ScaleHeight = Printer.Height - lngOffsetTop * 2
    picDraw.ScaleWidth = Printer.Width - lngOffsetLeft * 2
    '页边距
    lngTop = arrFormat(4)
    lngBottom = arrFormat(5)
    lngLeft = arrFormat(6)
    lngRight = arrFormat(7)
    '实际有效高度，宽度
    lngHeight = picDraw.ScaleHeight - lngTop - lngBottom
    lngWidth = picDraw.ScaleWidth - lngLeft - lngRight
    
    '上,下边距(lngTop , lngBottom)
    '左,右边距(lngLeft , lngRight)
    lineTop.X1 = 0
    lineTop.X2 = picDraw.ScaleWidth
    lineTop.Y1 = lngTop
    lineTop.Y2 = lngTop
    lineBottom.X1 = 0
    lineBottom.X2 = picDraw.ScaleWidth
    lineBottom.Y1 = picDraw.ScaleHeight - lngBottom
    lineBottom.Y2 = lineBottom.Y1
    
    lineLeft.X1 = lngLeft
    lineLeft.X2 = lngLeft
    lineLeft.Y1 = 0
    lineLeft.Y2 = picDraw.ScaleHeight
    lineRight.X1 = picDraw.ScaleWidth - lngRight
    lineRight.X2 = lineRight.X1
    lineRight.Y1 = 0
    lineRight.Y2 = picDraw.ScaleHeight
    
    '1、标题栏从上边距开始
    dblRowHeight = arrFormat(8)
    VsfData.RowHeightMin = dblRowHeight
    '固定行数,根据表头内容计算
    '98992,陈刘,2016-12-19
    If UBound(arrFormat) > 15 Then
        lngFixRowsheight = GetFixRowsHeight(arrFormat(16), arrFormat(9))
    Else
        lngFixRowsheight = arrFormat(9)
    End If
    '标题栏的字体设置
    lblTitle.FontName = arrFormat(10)
    lblTitle.FontSize = arrFormat(11)
    lblTitle.Caption = arrFormat(12)
    '表上项的字体设置
    lblUpTable.FontName = arrFormat(13)
    lblUpTable.FontSize = arrFormat(14)
    
    '设置标题栏坐标
    picDraw.FontName = lblTitle.FontName
    picDraw.FontSize = lblTitle.FontSize
    lblTitle.Left = lngLeft
    lblTitle.Top = lngTop + 30
    lblTitle.Width = lngWidth
    lblTitle.Height = picDraw.TextHeight("a")
    
    '2、表上标签从标题栏下开始
    strUpText = arrFormat(11)
    If strUpText <> "" Then
        lblUpTable.Caption = strUpText
        lblUpTable.AutoSize = True
    End If
    '设置表上项坐标
    picDraw.FontName = lblUpTable.FontName
    picDraw.FontSize = lblUpTable.FontSize
    lblUpTable.Left = lngLeft
    lblUpTable.Top = lblTitle.Top + lblTitle.Height + 30
    lblUpTable.Width = picDraw.ScaleWidth
    
    '3、设置表格
    lngHeight = lngHeight - lblUpTable.Height - lblTitle.Height
    VsfData.Top = lblUpTable.Top + lblUpTable.Height + 30
    VsfData.Left = lngLeft
    VsfData.Width = lngWidth
    lngHeight = lngHeight + lngTop - VsfData.Top - lngFixRowsheight
    VsfData.Height = lngHeight
    lngRows = CLng(lngHeight \ dblRowHeight)
    VsfData.Rows = lngFixRows + lngRows
    VsfData.FixedRows = lngFixRows
    VsfData.RowHeightMin = dblRowHeight
    
    Call VScroll1_Change
    
    If mrsItems.State = 0 Then
        '打开现存在的所有护理记录项目
        gstrSQL = " Select 项目序号,项目名称,项目类型,项目性质,项目长度,项目小数,项目表示,项目单位,项目值域,护理等级,应用方式" & _
                  " From 护理记录项目 B" & _
                  " Where B.应用方式<>0 " & _
                  " Order by 项目序号"
        Set mrsItems = zlDatabase.OpenSQLRecord(gstrSQL, "打开现存在的所有护理记录项目")
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mobjParent Is Nothing Then
        Set mobjParent = Nothing
    End If
End Sub

Private Sub VScroll1_Change()
    picDraw.Top = -1 * VScroll1.Value * (picDraw.Height - Me.Height) / 100
End Sub

Public Function ShowMe(ByVal objParent As Object, ByVal strInput As String) As Long
    '读取护理记录单的格式
    lngRows = 0
    arrFormat = Split(strInput, "|")
'    Me.Show 1, objParent   '核对数据正确性时才需要可见窗体
    Unload frmTendFilePreview
    Load frmTendFilePreview
    Unload frmTendFilePreview
    ShowMe = lngRows
End Function

Public Function AnaliseData(ByVal objParent As Object, ByVal lngFormat As Long, ByVal strInput As String) As Boolean
    mlngFormat = lngFormat
    arrFormat = Split(strInput, "|")
    Set mobjParent = objParent
    
    Unload frmTendFilePreview
    Load frmTendFilePreview
    
    If Not ReadStruDef Then
        '没有需要解析的列,因此直接返回解析成功,应该不存在这种情况
        AnaliseData = (mstrPreHead = "")
        Exit Function
    End If
    If Not ReadData Then Exit Function
    Unload frmTendFilePreview
    AnaliseData = True
End Function

Private Function ReadData() As Boolean
    Dim strCaption As String
    Dim rsPati As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim blnTrans As Boolean
    
    On Error GoTo errHand
    
    strCaption = mobjParent.Caption
    '读取所有使用该记录文件的病人列表(已经打印过的护理文件不进行重算数据行)
    gstrSQL = _
        " Select Id, 科室id, 病人id, 主页id, 婴儿" & vbNewLine & _
        " From 病人护理文件 a" & vbNewLine & _
        " Where 格式id = [1] And 归档人 Is Null And Not Exists" & vbNewLine & _
        " (Select 1 From 病人护理打印 b Where b.文件id = a.Id And b.打印人 Is Not Null)"
    
    Set rsPati = zlDatabase.OpenSQLRecord(gstrSQL, "提取使用该护理文件的病人列表", mlngFormat)
    
    gcnOracle.BeginTrans
    blnTrans = True
    Do While Not rsPati.EOF
        mobjParent.Caption = strCaption & Space(2) & "一共有" & rsPati.RecordCount & "份护理文件，正在处理：" & rsPati.AbsolutePosition
        
        '活动项目处理(数据列放在最后)
        Call PreActiveCOL(rsPati!ID)
        '装入数据
        mlngFile = rsPati!ID
        Call SQLCombination
        gstrSQL = mstrSQL
'        If mlngFile = 86 Then Stop
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取护理数据", CLng(rsPati!ID), CLng(rsPati!病人ID), CLng(rsPati!主页ID), CLng(rsPati!婴儿))
        '绑定数据并设置护理记录单的格式,同时实现一行数据分行显示的功能
        Call PreTendFormat(rsTemp)
        '解析每行数据
        If Not ParseData Then
            mobjParent.Caption = strCaption
            gcnOracle.RollbackTrans
            Exit Function
        End If
        
        rsPati.MoveNext
    Loop
    mobjParent.Caption = strCaption
    gcnOracle.CommitTrans
    blnTrans = False
    
    ReadData = True
    Exit Function
errHand:
    If blnTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    mobjParent.Caption = strCaption
End Function

Private Function ParseData() As Boolean
    Dim arrCol, arrData, arrMutilRow
    Dim strTime As String
    Dim lngMutilRow As Long, lngRecord As Long
    Dim lngRow As Long, lngCount As Long
    Dim lngCol As Long, lngMAX As Long
    Dim blnSave As Boolean, i As Long
    Dim strSQLData() As String
    ReDim Preserve strSQLData(1 To 1)
    
    On Error GoTo errHand
    '循环解析所有行数据(一列绑定多个项目,或者项目为文本型)
    
    arrCol = Split(mstrPreHead & mstrActivePreHead, ",")
    lngMAX = UBound(arrCol)
    lngCount = VsfData.Rows - 1
    
    gstrSQL = "ZL_病人护理打印_DELETE(" & mlngFormat & "," & mlngFile & ")"
    zlDatabase.ExecuteProcedure gstrSQL, "清除当前文件的打印数据"
    
    blnSave = False: arrMutilRow = Array()
    For lngRow = 1 To lngCount
        lngMutilRow = 0
        lngRecord = Val(VsfData.TextMatrix(lngRow, VsfData.Cols - 1))
        If lngRecord <> 0 Then
            blnSave = True
            strTime = Format(VsfData.TextMatrix(lngRow, 1), "YYYY-MM-DD HH:mm:ss")
            
            '分类汇总处理:计算数据行时需包含分类明细数据行
            If Val(VsfData.TextMatrix(lngRow, VsfData.Cols - 2)) < 0 Then
                If lngRow + 1 <= lngCount Then
                    If Val(VsfData.TextMatrix(lngRow + 1, VsfData.Cols - 1)) > 0 And Val(VsfData.TextMatrix(lngRow + 1, VsfData.Cols - 2)) < 0 And _
                        strTime = Format(VsfData.TextMatrix(lngRow + 1, 1), "YYYY-MM-DD HH:mm:ss") Then
                        blnSave = False
                    End If
                End If
            End If
                
            For lngCol = 0 To lngMAX
                If VsfData.TextMatrix(lngRow, arrCol(lngCol)) <> "" Then
                    '准备赋值
                    With txtLength
                        .Width = VsfData.ColWidth(arrCol(lngCol))
                        .Text = VsfData.TextMatrix(lngRow, arrCol(lngCol))
                        .FontName = VsfData.FontName
                        .FontSize = VsfData.FontSize
                        .FontBold = VsfData.CellFontBold
                        .FontItalic = VsfData.CellFontItalic
                    End With
                    arrData = GetData(txtLength.Text)
                    If UBound(arrData) > lngMutilRow Then
                        lngMutilRow = UBound(arrData)
                        If Trim(arrData(lngMutilRow)) = "" Then lngMutilRow = lngMutilRow - 1
                    End If
                End If
            Next
            If lngMutilRow < 0 Then lngMutilRow = 0
            ReDim Preserve arrMutilRow(UBound(arrMutilRow) + 1)
            arrMutilRow(UBound(arrMutilRow)) = lngMutilRow + 1
            If blnSave = True Then
                '----此处主要计算分类汇总的行数
                lngMutilRow = 0
                '计算分类明细的数据行数
                For i = 1 To UBound(arrMutilRow)
                    lngMutilRow = lngMutilRow + arrMutilRow(i)
                Next i
                '汇总主数据行数如果大于分类明细数据行数+1(1为默认的总量行数),则以主数据行数为准,否则以明细数据行+1为准
                If lngMutilRow + 1 > Val(arrMutilRow(0)) Then
                    lngMutilRow = lngMutilRow + 1
                Else
                    lngMutilRow = Val(arrMutilRow(0))
                End If
                arrMutilRow = Array()
                
                gstrSQL = "ZL_病人护理打印_UPDATE(" & mlngFile & ",to_date('" & strTime & "','yyyy-MM-dd hh24:mi:ss')" & "," & lngMutilRow & ")"
                strSQLData(ReDimArray(strSQLData)) = gstrSQL
            End If
        End If
    Next
    
    '执行过程
    For i = 1 To UBound(strSQLData)
        If strSQLData(i) <> "" Then
            Call zlDatabase.ExecuteProcedure(strSQLData(i), "产生打印解析数据")
        End If
    Next i
    
    ParseData = True
    Exit Function
errHand:
    MsgBox Err.Description
End Function

Private Sub PreTendFormat(ByVal rsTemp As ADODB.Recordset)
    Dim blnTag As Boolean
    Dim aryItem() As String
    Dim lngRow As Long, lngCol As Long, lngCount As Long, strCell As String
    On Error GoTo errHand
    
    '设置护理记录单的格式
    With VsfData
        .FixedRows = 3
        .Clear
        Set .DataSource = rsTemp
        
        '表头填写
        .MergeCells = flexMergeFixedOnly
        .MergeCellsFixed = flexMergeFree
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeRow(2) = True
        
        '完成活动项目列的处理(要重算的活动项目)，活动项目字段提取在SQL默认绑定在最后，此处同步将此移至到固定添加列的签名,避免能对后面的处理产生影响
        mstrActivePreHead = ""
        If mlngActiveColCount > 0 Then
            '移动活动项目列到固定列的签名
            For lngCol = 1 To mlngFixedCOL
                .ColPosition(.Cols - mlngActiveColCount - mlngFixedCOL) = .Cols - 1
                .ColHidden(.Cols - 1) = True
            Next
            For lngCol = .Cols - mlngActiveColCount - mlngFixedCOL To .Cols - mlngFixedCOL - 1
                mstrActivePreHead = mstrActivePreHead & "," & lngCol
                .ColHidden(lngCol) = True
            Next
        Else
            For lngCol = 1 To mlngFixedCOL
                .ColHidden(.Cols - lngCol) = True
            Next
        End If
        
        '设置列头
        aryItem = Split(mstrTabHead, "|")
        For lngCount = 0 To UBound(aryItem)
            strCell = aryItem(lngCount)
            lngRow = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            lngCol = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            .TextMatrix(lngRow, lngCol + 1) = strCell
        Next
        
        '列宽设置
        Dim blnAlign As Boolean
        aryItem = Split(mstrColWidth, ",")
        For lngCount = 2 To .Cols - 1
            If Not .ColHidden(lngCount) Then
                .ColWidth(lngCount) = Val(Split(aryItem(lngCount - 2), "`")(0))
                If InStr(1, aryItem(lngCount - 2), "`") <> 0 Then
                    blnAlign = True
                    .ColAlignment(lngCount) = Val(Split(aryItem(lngCount - 2), "`")(1))
                End If
            End If
        Next
        
        '固定行格式为居中
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        '再按列合并
        For lngCount = 0 To VsfData.Cols - 1
            VsfData.MergeCol(lngCount) = True
        Next
        .AutoSize 0, .Cols - 1
        
        If blnAlign = False Then
            '改为根据用户的设置显示列对齐方式
            If .FixedRows < .Rows Then .Cell(flexcpAlignment, .FixedRows, 0, .Rows - 1, .Cols - 1) = flexAlignGeneralCenter
        End If
        For lngCount = 0 To .Rows - 1
            If .ROWHEIGHT(lngCount) < .RowHeightMin Then .ROWHEIGHT(lngCount) = .RowHeightMin
        Next
        Select Case mintTabTiers
        Case 1
            .RowHidden(0) = False
            .RowHidden(1) = True
            .RowHidden(2) = True
        Case 2
            .RowHidden(0) = False
            .RowHidden(1) = False
            .RowHidden(2) = True
        Case 3
            .RowHidden(0) = False
            .RowHidden(1) = False
            .RowHidden(2) = False
        End Select
        For lngCount = 0 To .Cols - 1
            .MergeCol(lngCount) = True
        Next
        
    End With
    Exit Sub
errHand:
    MsgBox Err.Description
End Sub

Private Function ReadStruDef() As Boolean
    Dim arrCol
    Dim intCol As Integer, intCount As Integer
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    '读取病历文件格式定义
    gstrSQL = "Select d.对象序号, d.内容文本, d.要素名称" & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表格样式'" & _
        " Order By d.对象序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取病历文件格式定义", mlngFormat)
    With rsTemp
        Do While Not .EOF
            Select Case "" & !要素名称
            Case "表头层数": mintTabTiers = Val("" & !内容文本)
            Case "总列数":  VsfData.Cols = Val("" & !内容文本)
            Case "最小行高": VsfData.RowHeightMin = Val("" & !内容文本)
            Case "文本字体"
                strCurFont = "" & !内容文本
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = Val(Split(strCurFont, ",")(1))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
                End With
                Set VsfData.Font = objFont
                Set lblUpTable.Font = VsfData.Font
                Set Font = lblUpTable.Font
                
            Case "文本颜色": VsfData.ForeColor = Val("" & !内容文本)
            Case "表格颜色": VsfData.GridColor = Val("" & !内容文本): VsfData.GridColorFixed = VsfData.GridColor
            
            Case "标题文本": lblTitle.Caption = "" & !内容文本
            Case "标题字体"
                strCurFont = "" & !内容文本
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = Val(Split(strCurFont, ",")(1))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
                End With
                Set lblTitle.Font = objFont
                lblTitle.AutoSize = False
            
            Case "开始时间": mintTagFormHour = Val("" & !内容文本)
            Case "终止时间": mintTagToHour = Val("" & !内容文本)
            Case "条件字体"
                strCurFont = "" & !内容文本
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = Val(Split(strCurFont, ",")(1))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
                End With
                Set mobjTagFont = objFont
            Case "条件颜色": mlngTagColor = Val("" & !内容文本)
            Case "有效数据行": mlngActiveRows = Val(!内容文本)
            End Select
            .MoveNext
        Loop
    End With
    
    gstrSQL = "Select 格式, 页眉, 页脚,报表 From 病历页面格式 Where 种类 = 3 And 编号 In (Select 页面 From 病历文件列表 Where Id = [1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取病历页面格式", mlngFormat)
    If Not rsTemp.EOF Then
        mstrPaperSet = "" & rsTemp!格式: mstrPageHead = "" & rsTemp!页眉: mstrPageFoot = "" & rsTemp!页脚
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select d.对象序号, d.内容文本, d.要素名称, Nvl(d.是否换行, 0) As 是否换行" & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表上标签'" & _
        " Order By d.对象序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取表上标签定义", mlngFormat)
    With rsTemp
        mstrSubhead = ""
        Do While Not .EOF
            mstrSubhead = mstrSubhead & "|" & IIf(!是否换行 = 0, "", vbCrLf) & !内容文本 & "{" & !要素名称 & "}"
            .MoveNext
        Loop
        If mstrSubhead <> "" Then mstrSubhead = Mid(mstrSubhead, 2)
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select d.对象序号, d.内容行次, d.内容文本" & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表头单元'" & _
        " Order By d.对象序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取表头单元定义", mlngFormat)
    With rsTemp
        mstrTabHead = ""
        Do While Not .EOF
            mstrTabHead = mstrTabHead & "|" & !内容行次 - 1 & "," & !对象序号 & "," & !内容文本
            .MoveNext
        Loop
        If mstrTabHead <> "" Then mstrTabHead = Mid(mstrTabHead, 2)
    End With
    
    '查询语句组织
    '------------------------------------------------------------------------------------------------------------------
    Dim strSql外 As String, str格式 As String
    Dim bln日期 As Boolean, bln时间 As Boolean, bln护士 As Boolean
    Dim bln签名人 As Boolean, bln签名时间 As Boolean, bln签名日期 As Boolean
    Dim lngColumn As Long
    
    gstrSQL = "Select d.对象序号, d.对象属性, d.内容行次, d.内容文本, d.要素名称, d.要素单位,d.要素表示 " & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表列集合'" & _
        " Order By d.对象序号, d.内容行次"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取表列集合定义", mlngFormat)
    With rsTemp
        lngColumn = 0: mstrColumns = "": mstrColWidth = ""
        mstrSQL内 = "": mstrSQL中 = "": strSql外 = "": mstrSQL列 = "": mstrSQL条件 = ""
        bln日期 = False: bln时间 = False: bln护士 = False
        bln签名人 = False: bln签名时间 = False: bln签名日期 = False
        Do While Not .EOF
            
            If lngColumn <> !对象序号 Then
                mstrColumns = mstrColumns & IIf(mstrColumns = "", "", ";1;" & str格式) & "|" & !对象序号 & ";" & !要素名称
                mstrColWidth = mstrColWidth & "," & !对象属性
                str格式 = ""
                If !要素名称 <> "" Then
                    str格式 = "{" & NVL(!内容文本) & "[" & !要素名称 & "]" & NVL(!要素单位) & "}"
                    If strSql外 <> "" Then
                        mstrSQL列 = mstrSQL列 & "," & Mid(strSql外, 3) & " As C" & Format(lngColumn, "00")
                    Else
                        mstrSQL列 = mstrSQL列 & ",'' As C" & Format(lngColumn, "00")
                    End If
                Else
                    If strSql外 <> "" Then
                        mstrSQL列 = mstrSQL列 & "," & Mid(strSql外, 3) & " As C" & Format(lngColumn, "00")
                    Else
                        mstrSQL列 = mstrSQL列 & ",'' As C" & Format(lngColumn, "00")
                    End If
                End If
                strSql外 = ""
                lngColumn = !对象序号
            Else
                mstrColumns = mstrColumns & "'" & !要素名称
                str格式 = str格式 & "{" & NVL(!内容文本) & "[" & !要素名称 & "]" & NVL(!要素单位) & "}"
            End If
            
            Select Case !要素名称
            Case "日期"
                bln日期 = True
                mstrSQL中 = mstrSQL中 & ",日期"
                mstrSQL内 = mstrSQL内 & ",To_Char(l.发生时间, 'yyyy-mm-dd') As 日期"
                strSql外 = strSql外 & "||" & !要素名称
            Case "时间"
                bln时间 = True
                mstrSQL中 = mstrSQL中 & ",时间"
                mstrSQL内 = mstrSQL内 & ",To_Char(l.发生时间, 'hh24:mi') As 时间"
                strSql外 = strSql外 & "||" & !要素名称
                
            Case "签名人"
                bln签名人 = True
                mstrSQL中 = mstrSQL中 & ",签名人"
                mstrSQL内 = mstrSQL内 & ",l.签名人 As 签名人"
                strSql外 = strSql外 & "||" & !要素名称
                
            Case "签名时间"
                bln签名时间 = True
                mstrSQL中 = mstrSQL中 & ",签名时间"
                mstrSQL内 = mstrSQL内 & ",Decode(a.项目名称,Null,Null,Substr(a.项目名称,12,5)) As 签名时间"
                strSql外 = strSql外 & "||" & !要素名称
                
            Case "签名日期"
                bln签名日期 = True
                mstrSQL中 = mstrSQL中 & ",签名日期"
                mstrSQL内 = mstrSQL内 & ",Decode(a.项目名称,Null,Null,Substr(a.项目名称, 1,11)) As 签名日期"
                strSql外 = strSql外 & "||" & !要素名称
                
            Case "护士"
                bln护士 = True
                mstrSQL中 = mstrSQL中 & ",护士"
                mstrSQL内 = mstrSQL内 & ",l.保存人 As 护士"
                strSql外 = strSql外 & "||" & !要素名称
            Case Else
                If !要素名称 <> "" Then
                    mstrSQL中 = mstrSQL中 & ",Max(""" & !要素名称 & """) As """ & !要素名称 & """"
                    mstrSQL条件 = mstrSQL条件 & " Or """ & !要素名称 & """ Is Not Null"
                    strSql外 = strSql外 & "||""" & !要素名称 & """"
                    
                    If Trim("" & !内容文本) = "" And Trim("" & !要素单位) = "" Then
                        mstrSQL内 = mstrSQL内 & ", Decode(c.项目名称, '" & !要素名称 & "', Nvl(c.未记说明,c.记录内容), '') As """ & !要素名称 & """"
                    Else
                        mstrSQL内 = mstrSQL内 & ", Decode(c.项目名称, '" & !要素名称 & "', Nvl(c.未记说明,Decode(c.记录内容,Null,Null,'" & !内容文本 & "'||c.记录内容||'" & !要素单位 & "')), '') As """ & !要素名称 & """"
                    End If
                End If
            End Select
            .MoveNext
        Loop
        
        mstrColWidth = Mid(mstrColWidth, 2)
        '加入最后一列的格式
        mstrColumns = mstrColumns & IIf(mstrColumns = "", "", ";1;" & str格式) '& "|" & !对象序号 & ";" & !要素名称
        mstrColumns = Mid(mstrColumns, 2)     '格式如:列号;项目名称1,项目名称2|列号...,实例;1;体温|2;脉搏|3...
        If Mid(strSql外, 3) <> "" Then
            mstrSQL列 = mstrSQL列 & "," & Mid(strSql外, 3) & " As C" & Format(lngColumn, "00")
        Else
            mstrSQL列 = mstrSQL列 & ",'' As C" & Format(lngColumn, "00")
        End If
        
        If mstrSQL条件 <> "" Then mstrSQL条件 = "(" & Mid(mstrSQL条件, 5) & ")"
        
        '如果没有出现日期，时间，护士，则内层需要补充，以保证中层分组的正常：
        If bln日期 = False Then mstrSQL内 = mstrSQL内 & ",To_Char(l.发生时间, 'yyyy-mm-dd') As 日期"
        If bln时间 = False Then mstrSQL内 = mstrSQL内 & ",To_Char(l.发生时间, 'hh24:mi') As 时间"
        If bln护士 = False Then mstrSQL内 = mstrSQL内 & ",l.保存人 As 护士"
        
        If bln签名人 = False Then mstrSQL内 = mstrSQL内 & ",l.签名人 As 签名人"
        If bln签名日期 = False Then mstrSQL内 = mstrSQL内 & ",Decode(a.项目名称,Null,Null,Substr(a.项目名称,1,11)) As 签名日期"
        If bln签名时间 = False Then mstrSQL内 = mstrSQL内 & ",Decode(a.项目名称,Null,Null,Substr(a.项目名称,12,5)) As 签名时间"
        
        If Mid(mstrSQL中, 2) = "" Then
            MsgBox "对不起，您没有定义当前护理单的显示列信息，请在病历文件管理中定义！"
            Exit Function
        End If
        
        '程序内部控制增加固定列
        mstrSQL中 = mstrSQL中 & ",MAX(汇总类别) AS 汇总类别,MAX(记录ID) AS 记录ID"
        mstrSQL内 = mstrSQL内 & ",NVL(L.汇总类别,0) AS 汇总类别,C.记录ID"
        mstrSQL列 = mstrSQL列 & ",汇总类别,记录ID"
        
        '分析哪些列的数据需要进行打印解析处理
        Dim arrData
        Dim strtodo As String
        Dim intto As Integer, intDo As Integer
        mstrPreHead = ""
        arrCol = Split(mstrColumns, "|")
        intCount = UBound(arrCol)
        For intCol = 0 To intCount
            'If UBound(Split(Split(arrCol(intCol), ";")(3), "]}{[")) > 0 Then
            If UBound(Split(Split(arrCol(intCol), ";")(1), "'")) > 0 Then
                '只要有一个不是数字型则作为文本型处理
                
'                strtodo = Split(arrCol(intCol), ";")(3)
'                strtodo = Replace(strtodo, "]}{[", "||")
'                strtodo = Replace(Replace(strtodo, "{[", ""), "]}", "")
'                arrData = Split(strtodo, "||")
                strtodo = Split(arrCol(intCol), ";")(1)
                arrData = Split(strtodo, "'")
                intDo = UBound(arrData)
                For intto = 0 To intDo
                    mrsItems.Filter = "项目名称='" & arrData(intto) & "'"
                    If mrsItems.RecordCount <> 0 Then
                        '如果用户设置项目时都是设置成文本型,那么长度在20及以上的项目才检查,用户设置将数字型的设置成数字型才正确
'                        If mrsItems!项目类型 = 1 And mrsItems!项目表示 = 0 And mrsItems!项目长度 >= 10 Then
                            mstrPreHead = mstrPreHead & "," & Val(Split(arrCol(intCol), ";")(0)) + 1    '有两列固定的列，而列序号从0开始，因此+1
                            Exit For
'                        End If
                    End If
                Next
            Else
                '检查是否为文本项
                'mrsItems.Filter = "项目名称='" & Replace(Replace(Split(arrCol(intCol), ";")(3), "{[", ""), "]}", "") & "'"
                mrsItems.Filter = "项目名称='" & Split(arrCol(intCol), ";")(1) & "'"
                If mrsItems.RecordCount <> 0 Then
                    '如果用户设置项目时都是设置成文本型,那么长度在20及以上的项目才检查,用户设置将数字型的设置成数字型才正确
'                    If mrsItems!项目类型 = 1 And mrsItems!项目表示 = 0 And mrsItems!项目长度 >= 10 Then
                        mstrPreHead = mstrPreHead & "," & Val(Split(arrCol(intCol), ";")(0)) + 1    '有两列固定的列，而列序号从0开始，因此+1
'                    End If
                End If
            End If
        Next
        
        mrsItems.Filter = 0
        If mstrPreHead = "" Then Exit Function
        mstrPreHead = Mid(mstrPreHead, 2)
    End With
    
    ReadStruDef = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub SQLCombination()
    mstrSQL = "Select 备用,发生时间," & Mid(mstrSQL列, 12) & mstrSQLActive列 & vbCrLf & _
                " From (Select 记录组号,时间 as 备用,发生时间," & Mid(mstrSQL中, 2) & mstrSQLActive中 & vbCrLf & _
                "        From (Select NVL(c.记录组号,0) 记录组号,l.发生时间," & Mid(mstrSQL内, 2) & mstrSQLActive内 & vbCrLf & _
                "               From 病人护理数据 l, 病人护理明细 c,病人护理明细 a,病人护理文件 f " & vbCrLf & _
                "               Where l.Id = c.记录id And l.文件ID+0=f.ID " & _
                "               And a.记录id(+)=l.ID And a.记录类型(+)=5 And Nvl(a.终止版本,0)=0 And c.终止版本 Is Null And c.记录类型<>5  " & _
                "               And f.id=[1] And f.病人id = [2] And f.主页id = [3] And Nvl(f.婴儿,0)=[4] )" & vbCrLf & _
                IIf(mstrSQL条件 <> "", "Where " & mstrSQL条件 & mstrSQLActive条件, "") & _
                "       Group By 日期, 时间, 发生时间,记录组号,护士,签名人,签名日期,签名时间" & _
                                "       Order By 日期, 时间, 发生时间,记录组号,护士,签名人,签名日期,签名时间)"
End Sub

Private Sub PreActiveCOL(ByVal lngFileID As Long)
'功能：获取指定文件绑定的活动项目,并绑定到数据提取SQL中
    Dim rsTemp As New ADODB.Recordset
    Dim strCOLActive As String, StrKey As String, strCOL As String
    Dim i As Integer, j As Integer, blnAdd As Boolean, intMax As Integer, intCol As Integer
    Dim arrCol, arrAc
    Dim strColFormat As String, strCOLNames As String, strCOLPart As String, strCOLCOND As String, strCOLDEF As String, strCOLMID As String, strCOLIN As String
    On Error GoTo errHand
    
    mlngActiveColCount = 0
    mstrActivePreHead = ""
    mstrSQLActive列 = ""
    mstrSQLActive条件 = ""
    mstrSQLActive中 = ""
    mstrSQLActive内 = ""
    '1：获取绑定的活动项目信息
    gstrSQL = " Select   A.列号,A.页号,A.列头名称,A.序号,A.项目序号,A.部位 From 病人护理活动项目 A " & _
              " Where A.文件ID=[1]" & _
              " Order by A.页号,A.列号,A.序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取出所有自定义的活动项目", lngFileID)
    If rsTemp.RecordCount <> 0 Then
        Do While Not rsTemp.EOF
            If StrKey <> rsTemp!页号 & "_" & rsTemp!列号 Then
                StrKey = rsTemp!页号 & "_" & rsTemp!列号
                strCOLActive = strCOLActive & "||" & StrKey & "|" & rsTemp!项目序号 & "," & NVL(rsTemp!部位)
            Else
                strCOLActive = strCOLActive & ";" & rsTemp!项目序号 & "," & NVL(rsTemp!部位)
            End If
            rsTemp.MoveNext
        Loop
    End If
    If strCOLActive <> "" Then strCOLActive = Mid(strCOLActive, 3)
    If strCOLActive = "" Then Exit Sub
    '2:活动项目取重处理
    arrCol = Split(strCOLActive, "||")
    arrAc = Array()
    For i = 0 To UBound(arrCol)
        blnAdd = True
        StrKey = CStr(arrCol(i))
        StrKey = Mid(StrKey, InStr(1, StrKey, "|") + 1)
        For j = 0 To UBound(arrAc)
            If CStr(arrAc(j)) = StrKey Then
                blnAdd = False
                Exit For
            End If
        Next j
        If blnAdd = True Then
            ReDim Preserve arrAc(UBound(arrAc) + 1)
            arrAc(UBound(arrAc)) = StrKey
        End If
    Next i
    '3:开始进行活动项目数据提取SQL组装
    For i = 0 To UBound(arrAc)
        intCol = i + 1
        arrCol = Split(arrAc(i), ";") '每一列绑定的项目
        intMax = UBound(arrCol)
        '处理列表示(每列最多绑定两个项目)
        strCOLPart = "": strCOLNames = "": strColFormat = "": strCOLCOND = "": strCOLMID = "": strCOLIN = "": strCOLDEF = ""
        For j = 0 To intMax
            strCOLPart = Split(arrCol(j), ",")(1)
            mrsItems.Filter = "项目序号=" & Val(Split(arrCol(j), ",")(0))
            If mrsItems.RecordCount > 0 Then
                strCOLNames = strCOLNames & "," & mrsItems!项目名称
                strCOLCOND = strCOLCOND & " OR """ & strCOLPart & mrsItems!项目名称 & """ IS NOT NULL"
                strCOLMID = strCOLMID & ",Max(""" & strCOLPart & mrsItems!项目名称 & """) As """ & strCOLPart & mrsItems!项目名称 & """"
                If j = 0 Then
                    strCOLIN = strCOLIN & ", Decode(" & IIf(strCOLPart = "", "", "c.体温部位||") & "c.项目名称, '" & strCOLPart & mrsItems!项目名称 & "', Nvl(c.未记说明,c.记录内容), '') As """ & strCOLPart & mrsItems!项目名称 & """"
                Else
                    strCOLIN = strCOLIN & ", Decode(" & IIf(strCOLPart = "", "", "c.体温部位||") & "c.项目名称, '" & strCOLPart & mrsItems!项目名称 & "', Nvl(c.未记说明,Decode(c.记录内容,Null,'/','/'||c.记录内容||'')), '') As """ & strCOLPart & mrsItems!项目名称 & """"
                End If
                If j = 0 Then
                    If intMax = 0 Then
                        strCOLDEF = strCOLDEF & " """ & strCOLPart & mrsItems!项目名称 & """ AS A" & Format(intCol, "00")
                    Else
                        strCOLDEF = strCOLDEF & " """ & strCOLPart & mrsItems!项目名称 & """||"
                    End If
                Else
                    strCOLDEF = strCOLDEF & "NVL(""" & strCOLPart & mrsItems!项目名称 & """,'/')"
                    If j = intMax Then
                        strCOLDEF = "Decode(" & strCOLDEF & ",'" & String(intMax, "/") & "',''," & strCOLDEF & ") As A" & Format(intCol, "00")
                    End If
                End If
            End If
        Next j
        '将活动项目列加入要计算的列中
        If strCOLNames <> "" Then
            mlngActiveColCount = mlngActiveColCount + 1
        End If
        '组装活动项目SQL
        mstrSQLActive列 = mstrSQLActive列 & "," & strCOLDEF
        mstrSQLActive条件 = mstrSQLActive条件 & strCOLCOND
        mstrSQLActive中 = mstrSQLActive中 & strCOLMID
        mstrSQLActive内 = mstrSQLActive内 & strCOLIN
    Next i
    mrsItems.Filter = ""
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

'######################################################################################################################
'**********************************************************************************************************************
'以#分隔的区域内的代码都与分行相关,没事别动
Private Function GetData(ByVal strInput As String) As Variant
    Dim arrData
    Dim strData As String
    Dim strLine(256) As Byte
    Dim lngRow As Long, lngRows As Long, lngLen As Long
    
    GetData = ""
    lngRows = SendMessage(txtLength.hWnd, EM_GETLINECOUNT, 0&, 0&)
    For lngRow = 1 To lngRows
        Call ClearArray(strLine)
        lngLen = SendMessage(txtLength.hWnd, EM_GETLINE, lngRow - 1, strLine(0))
        Call ClearArray(strLine, lngLen)
        strData = StrConv(strLine, vbUnicode)
        strData = TruncZero(strData)
        GetData = GetData & IIf(GetData = "", "", "|ZYB.ZLSOFT|") & strData
    Next
    GetData = Split(GetData, "|ZYB.ZLSOFT|")
End Function

Private Sub ClearArray(strLine() As Byte, Optional ByVal lngPos As Long = 0)
    Dim intDo As Integer, intMax As Integer
    intMax = UBound(strLine)
    For intDo = lngPos To intMax
        strLine(intDo) = 0
        If lngPos > 0 Then Exit Sub '不为零,表示仅设置字符串结束符
    Next
    strLine(1) = 1
End Sub

Private Function TrimStr(ByVal str As String) As String
'功能：去掉字符串中\0以后的字符，并且去掉两端的空格

    If InStr(str, Chr(0)) > 0 Then
        TrimStr = Trim(Left(str, InStr(str, Chr(0)) - 1))
    Else
        TrimStr = Trim(str)
    End If
End Function

Private Function TruncZero(ByVal strInput As String) As String
'功能：去掉字符串中\0以后的字符
    Dim lngPos As Long
    
    lngPos = InStr(strInput, Chr(0))
    If lngPos > 0 Then
        TruncZero = Mid(strInput, 1, lngPos - 1)
    Else
        TruncZero = strInput
    End If
End Function

Private Function GetFixRowsHeight(ByVal strTabHead As String, ByVal lngFixRow As Long) As Long
    Dim aryItem() As String
    Dim arrTemp() As String
    Dim strCell As String, StrText As String
    Dim lngCellWith As Long
    Dim lngRow As Long, lngCol As Long
    Dim lngCount As Long

        aryItem = Split(strTabHead, "'")
        VsfData.Cols = (UBound(aryItem) + 1) / lngFixRow
        VsfData.FixedRows = 3
        With VsfData
            .MergeCells = flexMergeRestrictRows
            .MergeCellsFixed = flexMergeFree
            For lngCount = 0 To UBound(aryItem)
                strCell = aryItem(lngCount)
                lngCol = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
                lngRow = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
                StrText = Left(strCell, InStr(1, strCell, ",") - 1)
                lngCellWith = Mid(strCell, InStr(1, strCell, ",") + 1)
                .TextMatrix(lngRow, lngCol) = StrText
                '列宽设置
                    
                .ColWidth(lngCol) = lngCellWith
                
            Next
            '再按列合并
            For lngCount = 0 To .Cols - 1
                .MergeCol(lngCount) = True
            Next
            .MergeRow(-1) = True
            .AutoResize = True
            .WordWrap = True
            .AutoSizeMode = flexAutoSizeRowHeight
            .AutoSize 0, .Cols - 1
            .AutoResize = False
            '调整行高
            For lngCount = 0 To .Rows - 1
                If .ROWHEIGHT(lngCount) < .RowHeightMin Then .ROWHEIGHT(lngCount) = .RowHeightMin
            Next
        End With
        GetFixRowsHeight = VsfData.ROWHEIGHT(0) + VsfData.ROWHEIGHT(1) + VsfData.ROWHEIGHT(2)
End Function
