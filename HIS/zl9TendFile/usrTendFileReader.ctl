VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.UserControl usrTendFileReader 
   AutoRedraw      =   -1  'True
   ClientHeight    =   5550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8565
   KeyPreview      =   -1  'True
   ScaleHeight     =   5550
   ScaleWidth      =   8565
   Begin MSComctlLib.ImageList imgRow 
      Left            =   6150
      Top             =   510
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrTendFileReader.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrTendFileReader.ctx":039A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtLength 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   1005
      Left            =   3930
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   3090
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   450
      ScaleHeight     =   3825
      ScaleWidth      =   7485
      TabIndex        =   5
      Top             =   810
      Width           =   7515
      Begin VB.ComboBox cbo页码 
         Height          =   300
         Left            =   3405
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   3435
         Width           =   1320
      End
      Begin VB.OptionButton optPageAlign 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1380
         Picture         =   "usrTendFileReader.ctx":0734
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   3420
         Width           =   345
      End
      Begin VB.OptionButton optPageAlign 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   1710
         Picture         =   "usrTendFileReader.ctx":0ABA
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   3420
         Width           =   345
      End
      Begin VB.OptionButton optPageAlign 
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   2040
         Picture         =   "usrTendFileReader.ctx":0E4A
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   3420
         Width           =   345
      End
      Begin VB.CheckBox chk页码 
         Caption         =   "打印页码"
         Height          =   195
         Left            =   150
         TabIndex        =   8
         Top             =   3480
         Width           =   1155
      End
      Begin RichTextLib.RichTextBox rtbHead 
         Height          =   1380
         Left            =   0
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Width           =   6810
         _ExtentX        =   12012
         _ExtentY        =   2434
         _Version        =   393217
         BorderStyle     =   0
         ScrollBars      =   2
         OLEDragMode     =   0
         OLEDropMode     =   0
         TextRTF         =   $"usrTendFileReader.ctx":11A3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox rtbFoot 
         Height          =   1380
         Left            =   0
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1950
         Width           =   6810
         _ExtentX        =   12012
         _ExtentY        =   2434
         _Version        =   393217
         BorderStyle     =   0
         ScrollBars      =   2
         OLEDragMode     =   0
         OLEDropMode     =   0
         TextRTF         =   $"usrTendFileReader.ctx":1240
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lbl页码 
         AutoSize        =   -1  'True
         Caption         =   "页码位置"
         Height          =   180
         Left            =   2610
         TabIndex        =   10
         Top             =   3495
         Width           =   720
      End
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4515
      Left            =   60
      ScaleHeight     =   4515
      ScaleWidth      =   8385
      TabIndex        =   1
      Top             =   510
      Width           =   8385
      Begin VSFlex8Ctl.VSFlexGrid VsfData 
         Height          =   2655
         Left            =   1590
         TabIndex        =   0
         Top             =   930
         Width           =   4305
         _cx             =   7594
         _cy             =   4683
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
         SelectionMode   =   1
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
         FormatString    =   $"usrTendFileReader.ctx":12DD
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
         WordWrap        =   0   'False
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
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "一般护理记录单"
         Height          =   180
         Left            =   3450
         TabIndex        =   3
         Top             =   30
         Width           =   1275
      End
      Begin VB.Label lblSubhead 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名:##"
         Height          =   180
         Left            =   390
         TabIndex        =   2
         Top             =   540
         Width           =   720
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "usrTendFileReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'基础条件:
'1.护理记录同一时点只可能存在一条记录
'2.护理记录中不需要像体温单那样 , 记录病人是否外出, 拒测的数据, 测试了的数据才记录
'3.录入护理记录数据时,如果所录入的数据存在体温数据, 则提取过来
'4.护理记录单中不需要录入物理降温及脉搏短拙，如确需要可录入在护理摘要等文字型的列中
'#实现原理:
'1.对于用户修改过的数据,由于提供编辑状态页面切换的功能,对用户修改过的页数据进行整页复制,减少程序实现难度
'2.增加记录集记录哪些页哪些单元格被用户修改过
'3.任何编辑(粘贴,清空数据),都需要重新计算每行数据的占用行

Public mblnEditable As Boolean
'Public objFileSys As New FileSystemObject
'Public objStream As TextStream
Private mfrmParent As Object
Private mblnInit As Boolean
Private mblnShow As Boolean                 '是否显示录入框
Private mintPreDays As Long
Private mstrMaxDate As String
Private mintNORule As Integer               '0-统一编号;1-按文件格式顺序编号

Private mArrPage                           '记录单打印的页码数组:格式：页码;打印标识(1-续打,2-正常打印)
Private mlngMinIndex As Long, mlngMaxIndex As Long '数组最小和最大索引
Private mlng当前页码 As Long
Private mint当前起始页 As Integer           '当前文件的起始页(考虑已打印部分,以及预览从已打印页开始预览)
Private mint结束页 As Integer
Private mint页码 As Integer
Private mlng当前文件ID As Long
Private mstrMergeID As String  '合并文件
Private mlng格式ID As Long
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mlng科室ID As Long
Private mlng病区ID As Long
Private mint婴儿 As Integer
Private mbln心率 As Boolean                 '是否需要录入心率
Private mstrPrivs As String

Private mintSymbol As Integer               '当前控件索引
Private mstrSymbol As String                '特殊字符
Private mblnClear As Boolean                '如果为真,连续重打mrsDataMap记录集;当换页时应传假,保留用户修改的数据以备显示、保存使用
Private mstrCollectItems As String          '汇总项目集合
Private mstrColCollect As String            '汇总项目列集合:col;1|col;4,5
Private mstrColCorrelative As String        '汇总项目关联列集合:COl,3;COl,4|COl,5;COl,6(名称列号,项目序号;汇总列,项目序号),主要针对分类汇总
Private mstrCOLNothing As String            '未绑定的列集合+活动项目列(不管活动项目列是否绑定)
Private mstrCOLActive As String             '活动列集合
Private mstrCatercorner As String           '列对角线集合
Private mblnEditAssistant As Boolean        '当前选择的项目是否允许进行词句选择
Private mlngPageRows As Long                '此文件格式一页所显示的数据行
Private mlngOverrunRows As Long             '超出数据行
Private mlngRowCount As Long                '当前记录总行数
Private mlngRowCurrent As Long              '当前记录在本页的实际行数
Private mlngStartSpread As Long             '判断数据是否在打印起始页首行跨页：1-是，其它-否(实际开始行号)
Private mlngDate As Long                    '日期
Private mlngTime As Long                    '时间
Private mlngOperator As Long                '护士
Private mlngSignLevel As Long               '签名级别
Private mlngSigner As Long                  '签名信息
Private mlngSignName As Long                '签名人
Private mlngSignTime As Long                '签名时间
Private mlngJoinSignName As Long            '交班签名人
Private mlngRecord As Long                  '记录ID
Private mlngFileID As Long                  '文件ID：主要对于合并文件使用
Private mlngNoEditor As Long                '禁止编辑列,存在护士列则以护士列为准,不存在护士列则以签名列为准
Private mlngCollectType As Long             '汇总类别
Private mlngCollectText As Long             '汇总文本
Private mlngCollectStyle As Long            '汇总标记
Private mlngCollectDay As Long              '汇总日期:0-昨天;1-今天
Private mlngPrintedPage As Long             '打印页号
Private mlngPrintedRow As Long              '打印行号
Private mlngPrintedEndPage As Long          '打印结束页号,主要记录跨页数据当前打印到那一页
Private mlngCollectValue As Long
Private mlngPrintedTag As Long                '打印标识,记录上次打印是否采用未满页打印空白行
Private mbln日期时间合并 As Boolean         '日期与时间合并
Private mbln时间列隐藏 As Boolean           '隐藏时间列(如：血糖监测单只需要显示日期)
Private Const mlngDemo As Long = 0          '备用
Private mlngSingerType As Long              '护士、签名人显示模式（是首行显示还是首尾显示等）
Private mblnSignPic As Boolean            '签名人显示方式
Private mblnPrintRow As Boolean           '记录单预览、打印时，数据未满页空白部分是否输出表格
Private mblnFullPagePrint As Boolean      '记录单预览、打印时,数据满页才进行打印
Private mblnOddEvenPagePrint As Boolean   '记录单打印时，数据页奇偶输出
Private mblnDateModel As Boolean          '日期显示方式：相同日期当天显示一次；每一条记录都显示
Private mlngCollectColor As Long            '小结标识颜色
Private mblnShowNullCollet As Boolean       '小结是否在空值汇总下画横线

Private mblnSign As Boolean                 '是否签名
Private mblnArchive As Boolean              '是否归档
Private mintType As Integer                 '记录当前的编辑模式
Private mblnDateAd As Boolean               '日期缩写?
Private mstr开始时间 As String              '当前文件的开始时间
Private mstr结束时间 As String              '当前文件的结束时间
Private CellRect As RECT

Private mrsTemp As New ADODB.Recordset
Private mrsItems As New ADODB.Recordset             '所有护理记录项目清单
Private mrsElement As New ADODB.Recordset           '适用于记录单的标签要素
Private mrsSelItems As New ADODB.Recordset          '当前录入的护理记录项目清单
Private mrsDataMap As New ADODB.Recordset           '当前数据镜像

Private Enum ColIcon
    签名 = 1
    审签 = 2
End Enum
Private Enum SignLevel
    正高 = 1
    副高 = 2
    中级 = 3
    师级 = 4
    员士 = 5
    未定义 = 9
End Enum

Private Const WS_MAXIMIZE = &H1000000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_CAPTION = &HC00000
Private Const WS_SYSMENU = &H80000
Private Const WS_THICKFRAME = &H40000
Private Const WS_CHILD = &H40000000
Private Const WS_POPUP = &H80000000
Private Const SWP_NOZORDER = &H4
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

Public Event AfterDataChanged(ByVal blnChange As Boolean)
Public Event AfterRefresh()
Public Event AfterRowColChange(ByVal strInfo As String, ByVal blnImportant As Boolean, ByVal blnSign As Boolean, ByVal blnArchive As Boolean)

Dim strFields As String
Dim strValues As String
Dim blnScroll As Boolean

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
Private mstrSubhead As String       '表上标签
Private mstrTabHead As String       '表头单元
Private mstrColWidth As String      '列宽序列串
Private mstrColumns As String       '当前护理文件各列对应的项目
Private lngCurColor As Long, strCurFont As String, objFont As StdFont
'保存打开护理记录文件的SQL，在其它地方也有使用，不能修改
Private mstrSQL内 As String
Private mstrSQL中 As String
Private mstrSQL列 As String
Private mstrSQL条件 As String
Private mstrSQL As String

'##############################################################################################
'页眉页脚打印相关
Private Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type
'矩形
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'包含用于格式化指定设备的相关信息
Private Type FORMATRANGE
    hDC As Long             '渲染设备
    hdcTarget As Long       '目标设备
    rc As RECT              '渲染区域，单位：缇。
    rcPage As RECT          '渲染设备的整体区域，单位：缇。
    chrg As CHARRANGE       '用于格式化的文本范围。
End Type

Private Type PageInfo
    PageNumber As Long      '页码
    Start As Long           '字符起始位置
    End As Long             '字符终止位置
    ActualHeight As Long    '本页实际打印高度
End Type
Private AllPages() As PageInfo   '页信息
Private Const WM_PASTE = &H302&              '粘贴
Private Const WM_USER = &H400                '通常用 WM_USER + X 来自定义消息
Private Const EM_FORMATRANGE = (WM_USER + 57)    '为某一设备格式化指定范围的文本。
Private Const EM_SETTARGETDEVICE = (WM_USER + 72) '设置用于所见即所得的目标设备和行宽。
Private Const EM_HIDESELECTION = (WM_USER + 63)  '显示/隐藏文本。
Private Const PHYSICALOFFSETX = 112  '对于打印设备而言，表示从物理页的左边缘到可打印区域的左边缘的距离，采用设备单位。
Private Const PHYSICALOFFSETY = 113  '对于打印设备而言，表示从物理页的上边缘到可打印区域的上边缘的距离，采用设备单位。
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long '获取中英文混合字符串长度


'######################################################################################################################
'**********************************************************************************************************************
'以#分隔的区域内的代码都与绘图相关,没事别动
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Const WHITE_BRUSH = 0    '白色画笔
Private Const cdblWidth As Double = 6          '一个英文字符的宽度
Private Const cHideCols = 3         '前缀隐藏列:备用,时间,日期时间合并时显示日期列
Private Const cControlFields = 2    '记录集控制列:页号,行号

Private Function GetRBGFromOLEColor(ByVal dwOleColour As Long) As Long
    '将VB的颜色转换为RGB表示
    Dim clrref As Long
    Dim r As Long, g As Long, b As Long
    
    OleTranslateColor dwOleColour, 0, clrref
    
    b = (clrref \ 65536) And &HFF
    g = (clrref \ 256) And &HFF
    r = clrref And &HFF
    
    GetRBGFromOLEColor = RGB(r, g, b)
End Function

Private Function GetSymbolWidth(ByVal strPara As String) As Double
    '缺省是宋体9号,按字体大小同比放大
    Dim sinFontSize As Single
    Dim i As Integer, j As Integer
    
    j = Len(strPara)
    sinFontSize = VsfData.FontSize
    For i = 1 To j
        GetSymbolWidth = GetSymbolWidth + IIf(Asc(Mid(strPara, i, 1)) > 0, 1, 2) * cdblWidth * sinFontSize / 9
    Next
End Function

Private Sub DrawCell(ByVal hDC As Long, ByVal ROW As Long, ByVal COL As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Dim strText As String
    Dim strLeft As String
    Dim strRight As String
    Dim lngLeft As Long
    Dim lngRight As Long
    Dim dblWidth As Double
    Dim lngBackColor As Long
    Dim lngForeColor As Long
    Dim blnDraw As Boolean
    '绘图相关
    Dim lngPen As Long
    Dim lngOldPen As Long
    Dim lngBrush As Long
    Dim lngOldBrush As Long
    Dim lpPoint As POINTAPI
    Dim t_ClientRect As RECT
    On Error GoTo ErrHand
    '******************************************
    '在此事件中不能对单元格的任何属性赋值,包括Celldata,否则会引起该事件的死循环,导致工具栏或计时器无法正常工作。
    '******************************************
    '使用匹配的背景色，前景色与字体进行文本输出。
    If Not mblnInit Then Exit Sub
    If VsfData.RowHidden(ROW) Then Exit Sub
    Done = False
    
    strText = VsfData.TextMatrix(ROW, COL)
'    If IsDiagonal(Col) And InStr(1, strText, "/") <> 0 Then
    If InStr(1, strText, "/") <> 0 Then
        blnDraw = True
        '赋初值
        strLeft = Split(strText, "/")(0)
        strRight = Mid(strText, InStr(1, strText, "/") + 1)
        lngLeft = LenB(StrConv(strLeft, vbFromUnicode))
        lngRight = LenB(StrConv(strRight, vbFromUnicode))
        '取字符宽度
        dblWidth = GetSymbolWidth(strRight)
        '设定客户区域大小
        With t_ClientRect
            .Left = Left + 1
            .Top = Top + 1
            .Right = Right - 1
            .Bottom = Bottom - 1
        End With
        
        '1、清空内容
        '创建与背景色相同的刷子
        If ROW < VsfData.FixedRows Then
            lngBackColor = GetRBGFromOLEColor(VsfData.BackColorFixed)
            lngForeColor = GetRBGFromOLEColor(VsfData.ForeColorFixed)
        Else
            If ROW = VsfData.RowSel Then
                lngBackColor = GetRBGFromOLEColor(VsfData.BackColorSel)
                lngForeColor = RGB(0, 0, 0)
            Else
                lngBackColor = RGB(255, 255, 255)
                lngForeColor = GetRBGFromOLEColor(VsfData.Cell(flexcpForeColor, ROW, COL))
            End If

        End If
        lngBrush = CreateSolidBrush(lngBackColor)
        '使用该刷子填充背景色
        lngOldBrush = SelectObject(hDC, lngBrush)
        Call FillRect(hDC, t_ClientRect, lngBrush)
        '立即销毁临时使用的刷子并还原刷子
        Call SelectObject(hDC, lngOldBrush)
        Call DeleteObject(lngBrush)
        
        '2、准备画线
        '创建新画笔
        Call SetTextColor(hDC, lngForeColor)
        lngPen = CreatePen(0, 1, lngForeColor)
        lngOldPen = SelectObject(hDC, lngPen)
        '画线
        Call MoveToEx(hDC, Left, Bottom - 2, lpPoint)
        Call LineTo(hDC, Right, Top)
        '输出文本
        Call TextOut(hDC, Left, Top, strLeft, lngLeft)
        Call TextOut(hDC, IIf(Right - dblWidth >= Left, Right - dblWidth, Left), Bottom - 16, strRight, lngRight)
        
        '还原画笔并销毁
        Call SelectObject(hDC, lngOldPen)
        Call DeleteObject(lngPen)
        
        '已完成作图
        Done = True
    End If
    
    '3、如果是汇总行，则进行特殊处理
    If Val(VsfData.TextMatrix(ROW, mlngCollectType)) < 0 And Val(VsfData.TextMatrix(ROW, mlngCollectStyle)) > 0 _
        And (COL >= mlngDate And COL < mlngNoEditor) Then
        Call DrawCollectCell(hDC, ROW, COL, Left, Top, Right, Bottom)
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub DrawCollectCell(ByVal hDC As Long, ByVal ROW As Long, ByVal COL As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
    Dim lngPen As Long, lngOldPen As Long
    Dim lpPoint As POINTAPI
    
    '创建新画笔
    lngPen = CreatePen(0, 1, mlngCollectColor)
    lngOldPen = SelectObject(hDC, lngPen)
    
    Select Case Val(VsfData.TextMatrix(ROW, mlngCollectStyle))
    Case 1 '上下划横线
        '画线
        Call MoveToEx(hDC, Left, Top, lpPoint)
        Call LineTo(hDC, Right, Top)
        Call MoveToEx(hDC, Left, Bottom - 2, lpPoint)
        Call LineTo(hDC, Right, Bottom - 2)
    Case 2  '汇总项下双横线
        If IIf(mblnShowNullCollet, True, VsfData.TextMatrix(ROW, COL) <> "") Then
            '画线
            Call MoveToEx(hDC, Left, Bottom - 4, lpPoint)
            Call LineTo(hDC, Right, Bottom - 4)
            Call MoveToEx(hDC, Left, Bottom - 2, lpPoint)
            Call LineTo(hDC, Right, Bottom - 2)
        End If
    Case 3  '上横线
        '画线
        Call MoveToEx(hDC, Left, Top, lpPoint)
        Call LineTo(hDC, Right, Top)
    Case 4 '汇总项下单横线
        If InStr(1, "|" & mstrColCollect & ";", "|" & COL - (cHideCols + VsfData.FixedCols - 1) & ";") <> 0 Then 'And Val(VsfData.TextMatrix(ROW, COL)) <> 0 Then
             If IIf(mblnShowNullCollet, True, VsfData.TextMatrix(ROW, COL) <> "") Then
                '画线
                Call MoveToEx(hDC, Left, Bottom - 2, lpPoint)
                Call LineTo(hDC, Right, Bottom - 2)
            End If
        End If
    End Select
    
    '还原画笔并销毁
    Call SelectObject(hDC, lngOldPen)
    Call DeleteObject(lngPen)
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
        GetData = GetData & IIf(GetData = "", "", "|ZYB.ZLSOFT|") & strData & IIf(lngRow < lngRows, vbCrLf, "")
    Next
    GetData = Split(GetData, "|ZYB.ZLSOFT|")
End Function

Private Sub ClearArray(strLine() As Byte, Optional ByVal lngPos As Long = 0)
    Dim intDo As Integer, intMax As Integer
    intMax = UBound(strLine)
    For intDo = lngPos To intMax
        strLine(intDo) = 0
        If lngPos > 0 Then Exit Sub     '不为零,表示仅设置字符串结束符
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

'**********************************************************************************************************************
'######################################################################################################################


Private Function GetPeriod() As String
    Dim rs As New ADODB.Recordset
    Dim strPeriod As String
    On Error GoTo ErrHand
    
    '53588:刘鹏飞,2013-4-25,修改数据的时间小于病人入院时间，床号，病区不能显示问题
    '如：病人入科时间为2013-03-13 11:23:34 文件开始时间和入科相同，此时录入数据时间为 2013-03-13 11:23
    '就会导致无法提取床号，应为保存的数据时间为2013-03-13 11:23:00 小于了病人入科时间导致无法提取到数据
    '获取病人的入科时间
    If mint婴儿 = 0 Then
        gstrSQL = "Select 开始时间, Sysdate As 结束时间" & vbNewLine & _
            " From 病人变动记录" & vbNewLine & _
            " Where 病人id = [1] And 主页id = [2] And 开始原因 = 2" & vbNewLine & _
            " Union All" & vbNewLine & _
            " Select 开始时间, Sysdate As 结束时间" & vbNewLine & _
            " From 病人变动记录 a" & vbNewLine & _
            " Where a.病人id = [1] And a.主页id = [2] And a.开始原因 = 1 And Not Exists" & vbNewLine & _
            " (Select 1 From 病人变动记录 Where 病人id = a.病人id And 主页id = a.主页id And 开始原因 = 2)"
    Else
        gstrSQL = " Select   出生时间 AS 开始时间,sysdate AS 结束时间 From 病人新生儿记录 Where 病人ID=[1] And 主页ID=[2] And 序号=[3]"
    End If
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "取入院日期或出生日期", mlng病人ID, mlng主页ID, mint婴儿)
    
    '获取指定页码的数据发生时间范围
    gstrSQL = " Select  MIN(发生时间) 开始时间,MAX(发生时间) AS 结束时间 From 病人护理打印 Where 文件ID=[1] And (开始页号=[2] OR 结束页号=[2])"
    Call SQLDIY(gstrSQL)
    Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取指定页码的数据发生时间范围", mlng当前文件ID, mint页码)
    If NVL(mrsTemp!开始时间) = "" Then
        strPeriod = Format(rs!开始时间, "yyyy-MM-dd HH:mm:ss") & "～" & Format(rs!结束时间, "yyyy-MM-dd HH:mm:ss")
    Else
        If Format(mrsTemp!开始时间, "yyyy-MM-dd HH:mm:ss") < Format(rs!开始时间, "yyyy-MM-dd HH:mm:ss") Then
            strPeriod = Format(rs!开始时间, "yyyy-MM-dd HH:mm:ss") & "～" & Format(mrsTemp!结束时间, "yyyy-MM-dd HH:mm:ss")
        Else
            strPeriod = Format(mrsTemp!开始时间, "yyyy-MM-dd HH:mm:ss") & "～" & Format(mrsTemp!结束时间, "yyyy-MM-dd HH:mm:ss")
        End If
    End If
    GetPeriod = strPeriod
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ReadStruDef() As Boolean
    Dim lngCol As Long
    On Error GoTo ErrHand
    
    '读取文件属性
    mblnDateAd = False
    Call GetFileProperty
    
    '提取活动项目并加入列定义(格式：列号;表头名称|项目序号,部位;项目序号,部位||列号;表头名称...)
    mbln日期时间合并 = False
    mbln时间列隐藏 = False
    mstrCOLActive = ""
    mstrCOLNothing = ""
    mstrCollectItems = ""
    mstrColCollect = ""
    mstrColCorrelative = ""
    gstrSQL = " Select   A.列号,A.列头名称,A.序号,A.项目序号,A.部位 From 病人护理活动项目 A " & _
              " Where A.文件ID=[1] And A.页号=[2] " & _
              " Order by A.列号,A.序号"
    Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取出所有自定义的活动项目", mlng当前文件ID, mint页码)
    If mrsTemp.RecordCount <> 0 Then
        Do While Not mrsTemp.EOF
            If lngCol <> mrsTemp!列号 Then
                lngCol = mrsTemp!列号
                mstrCOLActive = mstrCOLActive & "||" & mrsTemp!列号 & ";" & mrsTemp!列头名称 & "|" & mrsTemp!项目序号 & "," & NVL(mrsTemp!部位)
            Else
                mstrCOLActive = mstrCOLActive & ";" & mrsTemp!项目序号 & "," & NVL(mrsTemp!部位)
            End If
            mrsTemp.MoveNext
        Loop
    End If
    If mstrCOLActive <> "" Then mstrCOLActive = Mid(mstrCOLActive, 3)
    
    '读取病历文件格式定义
    gstrSQL = "Select   d.对象序号, d.内容文本, d.要素名称" & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表格样式'" & _
        " Order By d.对象序号"
    Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取病历文件格式定义", mlng格式ID)
    With mrsTemp
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
                Set lblSubhead.Font = VsfData.Font
                Set Font = lblSubhead.Font
                
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
            Case "有效数据行"
                mlngOverrunRows = 0
                mlngPageRows = Val("" & !内容文本)
            Case "日期时间合并"
                mbln日期时间合并 = (Val("" & !内容文本) = 1)
            '65502:刘鹏飞,2013-11-12
            Case "时间列隐藏"
                mbln时间列隐藏 = (Val("" & !内容文本) = 1)
            End Select
            .MoveNext
        Loop
    End With
    
    If mbln时间列隐藏 = True Then mbln日期时间合并 = False
    
    gstrSQL = "Select   格式,页眉 ,页脚, 种类||'-'||编号 AS KEY From 病历页面格式 Where 种类 = 3 And 编号 In (Select 页面 From 病历文件列表 Where Id = [1])"
    Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取病历页面格式", mlng格式ID)
    If Not mrsTemp.EOF Then
        mstrPaperSet = "" & mrsTemp!格式
        If picHead.Tag = "" Then
            '考虑到医院内护理文件页眉页脚格式统一，此处只读取一次
            Call ReadPageHead(rtbHead, mrsTemp!Key)
            Call ReadPageFoot(rtbFoot, mrsTemp!Key)
            picHead.Tag = mrsTemp!Key
            chk页码.Value = IIf(Val(NVL(mrsTemp!页脚, 0)) > 0, 1, 0)
            If chk页码.Value = 1 Then
                optPageAlign(Val(NVL(mrsTemp!页脚, 0)) - 1).Value = True
                '46251,刘鹏飞,2012-09-11,装载页码输出位置
                If CInt(Val(NVL(mrsTemp!页眉, 0))) > 0 And CInt(Val(NVL(mrsTemp!页眉, 0))) < 5 Then
                    Call zlControl.CboSetIndex(cbo页码.hWnd, CInt(Val(NVL(mrsTemp!页眉, 0))) - 1)
                End If
            End If
        End If
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select   d.对象序号, d.内容文本, d.要素名称, Nvl(d.是否换行, 0) As 是否换行" & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表上标签'" & _
        " Order By d.对象序号"
    Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取表上标签定义", mlng格式ID)
    With mrsTemp
        mstrSubhead = ""
        Do While Not .EOF
            mstrSubhead = mstrSubhead & "|" & IIf(!是否换行 = 0, "", vbCrLf) & !内容文本 & "{" & !要素名称 & "}"
            .MoveNext
        Loop
        If mstrSubhead <> "" Then mstrSubhead = Mid(mstrSubhead, 2)
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select   d.对象序号, d.内容行次, d.内容文本" & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表头单元'" & _
        " Order By d.对象序号"
    Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取表头单元定义", mlng格式ID)
    With mrsTemp
        mstrTabHead = ""
        Do While Not .EOF
            mstrTabHead = mstrTabHead & "|" & !内容行次 - 1 & "," & !对象序号 & "," & !内容文本
            .MoveNext
        Loop
        If mstrTabHead <> "" Then mstrTabHead = Mid(mstrTabHead, 2)
    End With
    
    '查询语句组织
    '------------------------------------------------------------------------------------------------------------------
    Dim strSql外 As String, str格式 As String, strSqlNull As String
    Dim bln日期 As Boolean, bln时间 As Boolean, bln护士 As Boolean
    Dim bln签名人 As Boolean, bln签名时间 As Boolean, bln签名日期 As Boolean
    Dim bln对角线 As Boolean, bln选择项 As Boolean          '如果上一列是对角线且选择项,则直接提取各项数据,拼列头时在数值间加上/
    Dim lngColumn As Long, blnAddCollect As Boolean
    Dim strColCorrelative As String
    Dim str汇总值 As String
    
    gstrSQL = "Select   d.对象序号,d.对象标记, d.对象属性, d.内容行次, d.内容文本, upper(d.要素名称) AS 要素名称, d.要素单位,d.要素表示 " & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表列集合'" & _
        " Order By d.对象序号, d.内容行次"
        
    Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取表列集合定义", mlng格式ID)
    With mrsTemp
        lngColumn = 0: mstrColumns = "": mstrColWidth = "": mstrCatercorner = "": strColCorrelative = ""
        mstrSQL内 = "": mstrSQL中 = "": strSql外 = "": mstrSQL列 = "": mstrSQL条件 = "": strSqlNull = ""
        bln日期 = False: bln时间 = False: bln护士 = False
        bln签名人 = False: bln签名时间 = False: bln签名日期 = False
        Do While Not .EOF
            If lngColumn <> !对象序号 Then
                blnAddCollect = False
                If strColCorrelative <> "" Then
                    mstrColCorrelative = mstrColCorrelative & "|" & strColCorrelative
                End If
                strColCorrelative = ""
                
                mstrColumns = mstrColumns & IIf(mstrColumns = "", "", "'1'" & str格式) & "|" & !对象序号 & "'" & !要素名称
                mstrColWidth = mstrColWidth & "," & !对象属性 & "`" & !对象序号 & "`" & !要素表示
                If !要素表示 = 1 Then mstrCatercorner = mstrCatercorner & "," & !对象序号
                str格式 = ""
                If !要素名称 <> "" Then str格式 = "{" & NVL(!内容文本) & "[" & !要素名称 & "]" & NVL(!要素单位) & "}"
                    
                If Mid(strSqlNull, 3) = "" Then
                    strSqlNull = "''"
                Else
                    strSqlNull = Mid(strSqlNull, 3)
                End If
                mstrSQL列 = mstrSQL列 & "," & IIf(Mid(strSql外, 3) = "", "''", "Decode(" & Mid(strSql外, 3) & "," & strSqlNull & ",''," & Mid(strSql外, 3) & ")") & " As C" & Format(lngColumn, "00")
                
                strSql外 = ""
                strSqlNull = ""
                lngColumn = !对象序号
                bln对角线 = (NVL(!要素表示, 0) = 1)
                bln选择项 = False
                mrsItems.Filter = "项目名称='" & NVL(!要素名称) & "'"
                If mrsItems.RecordCount <> 0 Then
                    bln选择项 = (mrsItems!项目表示 = 5)
                    If mrsItems!项目表示 = 4 Then   '汇总项目
                        blnAddCollect = True
                        mstrCollectItems = mstrCollectItems & "," & mrsItems!项目序号
                        mstrColCollect = mstrColCollect & "|" & !对象序号 & ";" & mrsItems!项目序号
                        If Val(NVL(!对象标记)) > 0 And Val(NVL(!对象序号)) <> Val(NVL(!对象标记)) Then
                            strColCorrelative = Val(NVL(!对象标记)) & ";" & !对象序号 & "," & mrsItems!项目序号
                        End If
                    End If
                End If
                mrsItems.Filter = 0
            Else
                mstrColumns = mstrColumns & "," & !要素名称
                str格式 = str格式 & "{" & NVL(!内容文本) & "[" & !要素名称 & "]" & NVL(!要素单位) & "}"
                mrsItems.Filter = "项目名称='" & NVL(!要素名称) & "'"
                If mrsItems.RecordCount <> 0 Then
                    If mrsItems!项目表示 = 4 Then   '汇总项目
                        mstrCollectItems = mstrCollectItems & "," & mrsItems!项目序号
                        If blnAddCollect Then
                            strColCorrelative = ""
                            mstrColCollect = mstrColCollect & "," & mrsItems!项目序号
                        Else    '有可能一列绑定两个项目,第一个项目不是汇总项目,第二个项目才是汇总项目,因此,下面的代码保证加上列序号
                            blnAddCollect = True
                            mstrColCollect = mstrColCollect & "|" & !对象序号 & ";" & mrsItems!项目序号
                            If Val(NVL(!对象标记)) > 0 And Val(NVL(!对象序号)) <> Val(NVL(!对象标记)) Then
                                strColCorrelative = Val(NVL(!对象标记)) & ";" & !对象序号 & "," & mrsItems!项目序号
                            End If
                        End If
                    End If
                End If
                mrsItems.Filter = 0
            End If
            
            Select Case !要素名称
            Case "日期"
                bln日期 = True
                mblnDateAd = (NVL(!要素表示, 0) = 1)
                mstrSQL中 = mstrSQL中 & ",日期"
                mstrSQL内 = mstrSQL内 & ",To_Char(l.发生时间, " & IIf(mblnDateAd, "'dd/MM'", "'yyyy-mm-dd'") & ") As 日期"
                strSql外 = strSql外 & "||" & !要素名称
            Case "时间"
                bln时间 = True
                mstrSQL中 = mstrSQL中 & ",时间"
                mstrSQL内 = mstrSQL内 & ",To_Char(l.发生时间, 'hh24:mi') As 时间"
                strSql外 = strSql外 & "||" & !要素名称
                
            Case "签名人"
                bln签名人 = True
                mstrSQL中 = mstrSQL中 & ",签名人"
                '51589:刘鹏飞,2013-03-01,添加交班签名
                'mstrSQL内 = mstrSQL内 & ",l.签名人"
                mstrSQL内 = mstrSQL内 & ",DECODE(TRIM(NVL(L.签名人,'')),'',TRIM(L.签名人),DECODE(TRIM(NVL(L.交班签名人,'')),'',TRIM(L.签名人), TRIM(L.签名人) || '/' || TRIM(L.交班签名人))) 签名人"
                strSql外 = strSql外 & "||" & !要素名称
                
            Case "签名时间"
                bln签名时间 = True
                mstrSQL中 = mstrSQL中 & ",签名时间"
                mstrSQL内 = mstrSQL内 & ",l.签名时间"
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
                    
                    strSql外 = strSql外 & "||'" & !内容文本 & "'||""" & !要素名称 & """||'" & !要素单位 & "'"
                    strSqlNull = strSqlNull & "||" & "'" & !内容文本 & "'||'" & !要素单位 & "'"
                    mstrSQL内 = mstrSQL内 & ", Decode(c.项目名称, '" & !要素名称 & "', c.记录内容, '') As """ & !要素名称 & """"
                    
''                    If bln对角线 And bln选择项 Then
''                        If strSql外 <> "" Then
''                            '第二项
''                            strSql外 = strSql外 & "||'/'||""" & !要素名称 & """"
''                        Else
''                            '第一项
''                            strSql外 = strSql外 & "||""" & !要素名称 & """"
''                        End If
''                    Else
''                        strSql外 = strSql外 & "||""" & !要素名称 & """"
''                        strSqlNull = strSqlNull & "||" & "'" & !内容文本 & "'||'" & !要素单位 & "'"
''                    End If
''
''                    If (Trim("" & !内容文本) = "" And Trim("" & !要素单位) = "") Or (bln对角线 And bln选择项) Then
''                        mstrSQL内 = mstrSQL内 & ", Decode(c.项目名称, '" & !要素名称 & "', Nvl(c.未记说明,c.记录内容), '') As """ & !要素名称 & """"
''                    Else
''                        'mstrSQL内 = mstrSQL内 & ", Decode(c.项目名称, '" & !要素名称 & "', Nvl(c.未记说明,Decode(c.记录内容,Null,'" & !内容文本 & "'||'" & !要素单位 & "','" & !内容文本 & "'||c.记录内容||'" & !要素单位 & "')), '') As """ & !要素名称 & """"
''                        mstrSQL内 = mstrSQL内 & ", Decode(c.项目名称, '" & !要素名称 & "', Nvl(c.未记说明,Decode(c.记录内容,Null,'" & !内容文本 & "'||'" & !要素单位 & "','" & !内容文本 & "'||c.记录内容||'" & !要素单位 & "')),  '" & !内容文本 & "'||'" & !要素单位 & "') As """ & !要素名称 & """"
''                    End If
                Else
                    '为空表示未绑定列,强制加,后面进行替换
                    mstrCOLNothing = mstrCOLNothing & "," & Val(Format(!对象序号, "00"))
                    mstrSQL中 = mstrSQL中 & ",Max(""" & "C" & Format(!对象序号, "00") & """) As C" & Format(!对象序号, "00")
                    mstrSQL条件 = mstrSQL条件 & " Or """ & "C" & Format(!对象序号, "00") & """ Is Not Null"
                    mstrSQL内 = mstrSQL内 & ", C" & Format(!对象序号, "00") & " AS C" & Format(!对象序号, "00")
                End If
            End Select
            .MoveNext
        Loop
        
        If mstrCollectItems <> "" Then
            mstrCollectItems = Mid(mstrCollectItems, 2)
            mstrColCollect = Mid(mstrColCollect, 2)
        End If
        '在InitRecords中需要给汇总项目关列的名称列明添加项目序号
        If Left(mstrColCorrelative, 1) = "|" Then mstrColCorrelative = Mid(mstrColCorrelative, 2)
        mstrCOLNothing = Mid(mstrCOLNothing, 2)
        mstrCatercorner = Mid(mstrCatercorner, 2)
        mstrColWidth = Mid(mstrColWidth, 2)
        '加入最后一列的格式
        mstrColumns = mstrColumns & IIf(mstrColumns = "", "", "'1'" & str格式) '& "|" & !对象序号 & "'" & !要素名称
        mstrColumns = Mid(mstrColumns, 2)     '格式如:列号;项目名称1,项目名称2|列号...,实例;1;体温|2;脉搏|3...

        If Mid(strSqlNull, 3) = "" Then
            strSqlNull = "''"
        Else
            strSqlNull = Mid(strSqlNull, 3)
        End If
        mstrSQL列 = mstrSQL列 & "," & IIf(Mid(strSql外, 3) = "", "''", "Decode(" & Mid(strSql外, 3) & "," & strSqlNull & ",''," & Mid(strSql外, 3) & ")") & " As C" & Format(lngColumn, "00")
                
                
        
        If mstrSQL条件 <> "" Then mstrSQL条件 = "(" & Mid(mstrSQL条件, 5) & ")"
        
        '如果没有出现日期，时间，护士，则内层需要补充，以保证中层分组的正常：
        If bln日期 = False Then mstrSQL内 = mstrSQL内 & ",To_Char(l.发生时间, 'yyyy-mm-dd') As 日期"
        If bln时间 = False Then mstrSQL内 = mstrSQL内 & ",To_Char(l.发生时间, 'hh24:mi') As 时间"
        If bln护士 = False Then mstrSQL内 = mstrSQL内 & ",l.保存人 As 护士"
        
        '51589:刘鹏飞,2013-03-01,添加交班签名
        'If bln签名人 = False Then mstrSQL内 = mstrSQL内 & ",l.签名人 As 签名人"
        If bln签名人 = False Then mstrSQL内 = mstrSQL内 & ",DECODE(TRIM(NVL(L.签名人,'')),'',TRIM(L.签名人),DECODE(TRIM(NVL(L.交班签名人,'')),'',TRIM(L.签名人), TRIM(L.签名人) || '/' || TRIM(L.交班签名人))) 签名人"
        If bln签名时间 = False Then mstrSQL内 = mstrSQL内 & ",l.签名时间"
        
        If Mid(mstrSQL中, 2) = "" Then
            MsgBox "对不起，您没有定义当前护理单的显示列信息，请在病历文件管理中定义！", vbInformation, gstrSysName
            Exit Function
        End If
        '50503:刘鹏飞,2012-09-12,数据从某一页第一行就开始跨页,添加开始行号
        '56134:刘鹏飞,2012-12-19,病人护理打印添加打印标识
        '46506:刘鹏飞,2012-12-27,病人护理打印添加打印结束页号，用于标识跨页数据打印
        '程序内部控制增加固定列
        '说明如果要加列请在“打印标识”列之前添加，否则请修改zlPrintMdl
        str汇总值 = " Decode(Instr('|" & mstrColCollect & "|', ';' || C.项目序号 || '|'), 0," & _
                  " Decode(Instr('|" & mstrColCollect & "|', ',' || C.项目序号 || '|'), 0," & _
                  " Decode(Instr('|" & mstrColCollect & "|', ';' || C.项目序号 || ','), 0, Null, " & _
                  " Substr(Substr('|" & mstrColCollect & "|', 1, Instr('|" & mstrColCollect & "|', ';' || C.项目序号 || ',') - 1) || ';' || C.项目序号," & _
                  " Instr(Substr('|" & mstrColCollect & "|', 1, Instr('|" & mstrColCollect & "|', ';' || C.项目序号 || ',') - 1) || ';' || C.项目序号, '|', -1) + 1)), " & _
                  " Substr(Substr('|" & mstrColCollect & "|', 1, Instr('|" & mstrColCollect & "|', ',' || C.项目序号 || '|') - 1) || ',' || C.项目序号, " & _
                  " Instr(Substr('|" & mstrColCollect & "|', 1, Instr('|" & mstrColCollect & "|', ',' || C.项目序号 || '|') - 1) || ',' || " & _
                  " C.项目序号, '|', -1) + 1)),Substr(Substr('|" & mstrColCollect & "|', 1, Instr('|" & mstrColCollect & "|', ';' || C.项目序号 || '|') - 1) || ';' || C.项目序号, " & _
                  " Instr(Substr('|" & mstrColCollect & "|', 1, Instr('|" & mstrColCollect & "|', ';' || C.项目序号 || '|') - 1) || ';' || C.项目序号, '|', -1) + 1)) 汇总值 "
                  
        mstrSQL中 = UCase(mstrSQL中 & ",MAX(签名级别) AS 签名级别,MAX(签名信息) AS 签名信息,MAX(交班签名人) AS 交班签名人,MAX(文件ID) AS 文件ID,MAX(记录ID) AS 记录ID,MAX(行数) AS 行数,MAX(实际行数) AS 实际行数,MAX(开始行号) AS 开始行号,MAX(打印结束页号) as 打印结束页号,f_List2str(Cast(Collect(汇总值) As t_Strlist), '|') 汇总值,MAX(打印标识) AS 打印标识,MAX(汇总类别) AS 汇总类别,MAX(汇总文本) AS 汇总文本,MAX(汇总标记) AS 汇总标记,MAX(汇总日期) AS 汇总日期,MAX(打印页号) AS 打印页号,MAX(打印行号) AS 打印行号")
        mstrSQL内 = UCase(mstrSQL内 & ",l.签名级别,l.签名人 AS 签名信息,l.交班签名人,l.文件ID,C.记录ID,P.行数||'' AS 行数,DECODE(SIGN(P.结束页号-P.开始页号),1,DECODE(SIGN([5]-P.开始页号),1, P.结束行号,P.行数-P.结束行号 ),P.行数) AS 实际行数,DECODE(SIGN(P.结束页号-P.开始页号),1,DECODE(SIGN([5]-P.开始页号),1,P.开始行号+P.行数-P.结束行号,P.开始行号),P.开始行号) 开始行号,P.打印结束页号," & str汇总值 & ",P.打印标识,NVL(L.汇总类别,0) AS 汇总类别,L.汇总文本,L.汇总标记,to_char(L.发生时间,'yyyy-MM-dd hh24:mi:ss')||'' AS 汇总日期,p.打印页号,p.打印行号")
        mstrSQL列 = UCase(mstrSQL列 & ",签名级别,签名信息,交班签名人,文件ID,记录ID,行数,实际行数,开始行号,打印结束页号,汇总值,打印标识,汇总类别,汇总文本,汇总标记,汇总日期,打印页号,打印行号")
        
        
        '将活动项目加入到SQL中
        Call DelActiveNoUsed
        Call PreActiveCOL
        'Call SQLCombination
    End With
    
    ReadStruDef = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub PreActiveHead()
    Dim arrData
    Dim intCol As Integer
    Dim strName As String
    Dim intDo As Integer, intCount As Integer
    '更新表头
    
    arrData = Split(mstrCOLActive, "||")
    intCount = UBound(arrData)
    For intDo = 0 To intCount
        intCol = Split(Split(arrData(intDo), "|")(0), ";")(0)
        strName = Split(Split(arrData(intDo), "|")(0), ";")(1)
        VsfData.TextMatrix(mintTabTiers - 1, intCol + cHideCols + VsfData.FixedCols - 1) = strName
        If mintTabTiers = 3 And VsfData.TextMatrix(1, intCol + cHideCols + VsfData.FixedCols - 1) = "" Then VsfData.TextMatrix(1, intCol + cHideCols + VsfData.FixedCols - 1) = strName
        If mintTabTiers = 2 And VsfData.TextMatrix(0, intCol + cHideCols + VsfData.FixedCols - 1) = "" Then VsfData.TextMatrix(0, intCol + cHideCols + VsfData.FixedCols - 1) = strName
    Next
End Sub

Private Function DelActiveNoUsed() As Boolean
'------------------------------------------------
'功能:删除绑定在非空列上的活动项目列信息
'编制:刘鹏飞,2013-07-16
'问题号:63401
'------------------------------------------------
    Dim arrData, arrActive, arrCol
    Dim strSQL As String
    Dim lngCol As Long, intDo As Integer, intCount As Integer
    Dim blnTran As Boolean
    
    If mstrCOLNothing = "" Then DelActiveNoUsed = True: Exit Function
    arrActive = Array()
    arrCol = Array()
    arrData = Split(mstrCOLActive, "||")
    intCount = UBound(arrData)
    For intDo = 0 To intCount
        lngCol = Val(Split(Split(arrData(intDo), "|")(0), ";")(0))
        If InStr(1, "," & mstrCOLNothing & ",", "," & lngCol & ",") <> 0 Then
            '记录现有正常空列上的活动项目设置信息
            ReDim Preserve arrActive(UBound(arrActive) + 1)
            arrActive(UBound(arrActive)) = CStr(arrData(intDo))
        Else
            '记录将要移除的活动项目列号
            ReDim Preserve arrCol(UBound(arrCol) + 1)
            arrCol(UBound(arrCol)) = lngCol
        End If
    Next
    
    On Error GoTo ErrHand
    
    '删除不需要的活动项目信息(主要是修正之前错误的数据,发生情况较少)
    If UBound(arrCol) > 1 Then
        gcnOracle.BeginTrans
        blnTran = True
    End If
    
    For intDo = 0 To UBound(arrCol)
        If CStr(arrCol(intDo)) <> "" Then
            strSQL = "ZL_病人护理页面_UPDATE(" & mlng当前文件ID & "," & mint页码 & "," & Val(arrCol(intDo)) & ",NULL,'" & gstrUserName & "')"
            Call zlDatabase.ExecuteProcedure(strSQL, "保存活动项目绑定数据")
        End If
    Next
    If blnTran = True Then gcnOracle.CommitTrans
    
    '重新更新提取的活动项目列信息
    If UBound(arrActive) = -1 Then
        mstrCOLActive = ""
    Else
        mstrCOLActive = Join(arrActive, "||")
    End If
    
    DelActiveNoUsed = True
    Exit Function
ErrHand:
    If blnTran = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub PreActiveCOL()
    Dim arrData
    Dim arrCol
    Dim intCol As Integer, strName As String
    Dim strColFormat As String, strCOLNames As String, strCOLPart As String, strCOLCOND As String, strCOLDEF As String, strCOLMID As String, strCOLIN As String
    Dim intDo As Integer, intCount As Integer
    Dim intIn As Integer, intMax As Integer
    '将活动项目加入到查询SQL中，格式：列号;表头名称|项目序号,部位;项目序号,部位||列号;表头名称...
    '绑定多个项目，该列就自动转为对角线列
    
    arrData = Split(mstrCOLActive, "||")
    intCount = UBound(arrData)
    For intDo = 0 To intCount
        intCol = Split(Split(arrData(intDo), "|")(0), ";")(0)
        strName = Split(Split(arrData(intDo), "|")(0), ";")(1)
        arrCol = Split(Split(arrData(intDo), "|")(1), ";")
        
        '处理列表示(每列最多绑定两个项目)
        strCOLPart = ""
        strCOLNames = ""
        strColFormat = ""
        strCOLCOND = ""
        strCOLMID = ""
        strCOLIN = ""
        strCOLDEF = ""
        intMax = UBound(arrCol)
        For intIn = 0 To intMax
            strCOLPart = Split(arrCol(intIn), ",")(1)
            mrsItems.Filter = "项目序号=" & Val(Split(arrCol(intIn), ",")(0))
            strCOLNames = strCOLNames & "," & mrsItems!项目名称
            strCOLCOND = strCOLCOND & " OR """ & strCOLPart & mrsItems!项目名称 & """ IS NOT NULL"
            strCOLMID = strCOLMID & ",Max(""" & strCOLPart & mrsItems!项目名称 & """) As """ & strCOLPart & mrsItems!项目名称 & """"
            If intIn = 0 Then
                strCOLIN = strCOLIN & ", Decode(" & IIf(strCOLPart = "", "", "c.体温部位||") & "c.项目名称, '" & strCOLPart & mrsItems!项目名称 & "', Nvl(c.未记说明,c.记录内容), '') As """ & strCOLPart & mrsItems!项目名称 & """"
            Else
                strCOLIN = strCOLIN & ", Decode(" & IIf(strCOLPart = "", "", "c.体温部位||") & "c.项目名称, '" & strCOLPart & mrsItems!项目名称 & "', Nvl(c.未记说明,Decode(c.记录内容,Null,'/','/'||c.记录内容||'')), '') As """ & strCOLPart & mrsItems!项目名称 & """"
            End If
            If intIn = 0 Then
                If intMax = 0 Then
                    strCOLDEF = strCOLDEF & " """ & strCOLPart & mrsItems!项目名称 & """ AS C" & Format(intCol, "00")
                Else
                    strCOLDEF = strCOLDEF & " """ & strCOLPart & mrsItems!项目名称 & """||"
                End If
            Else
                strCOLDEF = strCOLDEF & "NVL(""" & strCOLPart & mrsItems!项目名称 & """,'/')"
                If intIn = intMax Then
                    strCOLDEF = "Decode(" & strCOLDEF & ",'" & String(intMax, "/") & "',''," & strCOLDEF & ") As C" & Format(intCol, "00")
                End If
            End If
            
            strColFormat = strColFormat & "{[" & strCOLPart & mrsItems!项目名称 & "]" & IIf(intMax > 0 And intIn < intMax, "/", "") & "}"
        Next
        If strCOLPart <> "" Then
            strCOLPart = Mid(strCOLPart, 2)
        End If
        strCOLNames = Mid(strCOLNames, 2)
        
        '对角线
        If intMax > 0 Then
            mstrCatercorner = mstrCatercorner & IIf(mstrCatercorner = "", "", ",") & intCol
        End If
        '列格式:15'护士'1'{[护士]}
        '77476:LPF:活动列替换intcol前添加"|"字符,避免第3列和第13列都为活动项目时项目替换错误
        mstrColumns = Replace(mstrColumns, "|" & intCol & "''1'", "|" & intCol & "'" & strCOLNames & "'1'" & strColFormat)
        '列
        mstrSQL列 = Replace(mstrSQL列, "'' AS C" & Format(intCol, "00"), strCOLDEF)
        '条件
         '53893:刘鹏飞,2012-09-21,处理活动项目绑定在时间后面的情况
        'mstrSQL条件 = Replace(UCase(mstrSQL条件), " OR """ & "C" & Format(intCol, "00") & """ IS NOT NULL", strCOLCOND)
        mstrSQL条件 = Replace(UCase(Replace(UCase(mstrSQL条件), " OR """ & "C" & Format(intCol, "00") & """ IS NOT NULL", strCOLCOND)), """" & "C" & Format(intCol, "00") & """ IS NOT NULL", Mid(strCOLCOND, 5))
        '中
        mstrSQL中 = Replace(mstrSQL中, ",MAX(""" & "C" & Format(intCol, "00") & """) AS C" & Format(intCol, "00"), strCOLMID)
        '内
        mstrSQL内 = Replace(mstrSQL内, ", C" & Format(intCol, "00") & " AS C" & Format(intCol, "00"), strCOLIN)
    Next
    mrsItems.Filter = 0
    
    '将未绑定的列的SQL部分连续重打
    If mstrCOLNothing = "" Then Exit Sub
    arrData = Split(mstrCOLNothing, ",")
    intCount = UBound(arrData)
    For intDo = 0 To intCount
        '列(必须要保留)
'        mstrSQL列 = Replace(mstrSQL列, ",'' AS C" & arrData(intDo), "")
        '条件
        'mstrSQL条件 = Replace(UCase(mstrSQL条件), " OR """ & "C" & Format(arrData(intDo), "00") & """ IS NOT NULL", "")
        mstrSQL条件 = Replace(UCase(Replace(UCase(mstrSQL条件), " OR """ & "C" & Format(arrData(intDo), "00") & """ IS NOT NULL", "")), """" & "C" & Format(arrData(intDo), "00") & """ IS NOT NULL OR ", "")
        mstrSQL条件 = Replace(UCase(mstrSQL条件), "(""" & "C" & Format(arrData(intDo), "00") & """ IS NOT NULL)", "")

        '中
        mstrSQL中 = Replace(mstrSQL中, ",MAX(""" & "C" & Format(arrData(intDo), "00") & """) AS C" & Format(arrData(intDo), "00"), "")
        '内
        mstrSQL内 = Replace(mstrSQL内, ", C" & Format(arrData(intDo), "00") & " AS C" & Format(arrData(intDo), "00"), "")
    Next
End Sub

Private Sub SQLCombination(ByVal str条件 As String)
    mstrSQL = "Select  '' as 备用,发生时间,发生时间 发生时间1," & Mid(mstrSQL列, 12) & vbCrLf & _
                " From (Select 记录组号,时间 as 备用,发生时间," & Mid(mstrSQL中, 2) & vbCrLf & _
                "        From (Select nvl(c.记录组号,0) 记录组号,to_char(l.发生时间,'yyyy-MM-dd hh24:mi:ss') AS 发生时间," & Mid(mstrSQL内, 2) & vbCrLf & _
                "               From 病人护理数据 l, 病人护理明细 c,病人护理文件 f,病人护理打印 p " & vbCrLf & _
                "               Where l.ID=p.记录ID And l.Id = c.记录id And l.文件ID+0=f.ID+0 And f.ID=p.文件ID " & _
                "               And c.终止版本 Is Null And c.记录类型<>5  " & _
                "               And f.id=[1] And f.病人id = [2] And f.主页id = [3] And Nvl(f.婴儿,0)=[4] " & str条件 & ")" & vbCrLf & _
                IIf(mstrSQL条件 <> "", "Where " & mstrSQL条件, "") & _
                "       Group By 日期, 时间, 发生时间,记录组号,护士,签名人,签名时间" & _
                                "       Order By 发生时间,记录组号,护士,签名人,签名时间)"
End Sub

Private Sub zlReadTip(aryPeriod)
    Dim aryRow() As String, aryItem() As String
    Dim strPrefix As String, strItemName As String
    Dim lngRow As Long, lngCol As Long, lngCount As Long, strCell As String, strBed As String
    Dim strTmpSQL As String
    Dim strTmp As String
    Dim blnReplace As Boolean
    
    Err = 0: On Error GoTo ErrHand
    
    '表上标签获取
    lblSubhead.Caption = ""
    lblSubhead.Tag = ""
    
    '87057,病人10:30:20转科入住建立护理文件时间为10:30:20,此时录入首条数据为10:30(记录单无法录入秒),导致无法显示新的科室
    aryPeriod(0) = Format(aryPeriod(0), "YYYY-MM-DD HH:mm") & ":59"
    
    '获取当前页之前的最后科室ID
    gstrSQL = "Select 科室ID From 病人变动记录 " & _
        "   Where  病人ID=[1] And 主页ID=[2] And [3]>=开始时间 " & _
        " And 开始时间 IS NOT NULL And 科室id IS NOT NULL And NVL(附加床位,0)=0 Order by 开始时间 DESC"
    Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取当前页之前的最后科室ID", mlng病人ID, mlng主页ID, CDate(aryPeriod(0)))
    If mrsTemp.RecordCount > 0 Then mlng科室ID = Val(mrsTemp!科室ID)
    
    gstrSQL = "Select [1] || Zl_Replace_Element_Value([2],[3],[4],2,NULL,[5],[6]) as 信息 From Dual"
    aryItem = Split(mstrSubhead, "|")
        
    For lngCount = 0 To UBound(aryItem)
        strPrefix = Left(aryItem(lngCount), InStr(1, aryItem(lngCount), "{") - 1)
        strItemName = Mid(aryItem(lngCount), InStr(1, aryItem(lngCount), "{") + 1, InStr(1, aryItem(lngCount), "}") - InStr(1, aryItem(lngCount), "{") - 1)
        
        strTmp = strPrefix
        strCell = ""
        '68336
        blnReplace = True
        mrsElement.Filter = "中文名='" & strItemName & "'"
        If mrsElement.RecordCount > 0 Then
            blnReplace = Val(NVL(mrsElement!替换域, 0)) = 1
        End If
        Select Case strItemName
        Case "当前病区"
        
            strTmpSQL = "Select   b.名称" & vbNewLine & _
                        "From (Select 病区id, 开始时间, Nvl(终止时间, To_Date('3000-01-01', 'yyyy-mm-dd')) As 终止时间" & vbNewLine & _
                        "            From 病人变动记录" & vbNewLine & _
                        "            Where 病人id = [1] And 主页id = [2] And 科室id = [3]　And NVL(附加床位,0)=0 And 开始时间 IS NOT NULL) a,部门表 b " & vbNewLine & _
                        "Where ([4] Between a.开始时间 And a.终止时间 Or [4] >= a.开始时间) And a.病区id Is Not Null And b.ID=a.病区id" & vbNewLine & _
                        "Order By a.开始时间"
                        
            Set mrsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "当前病区", mlng病人ID, mlng主页ID, mlng科室ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            If mrsTemp.BOF = False Then mrsTemp.MoveLast
            
        Case "当前床号"

            strTmpSQL = "Select   a.床号" & vbNewLine & _
                        "From (Select 床号, 开始时间, Nvl(终止时间, To_Date('3000-01-01', 'yyyy-mm-dd')) As 终止时间" & vbNewLine & _
                        "            From 病人变动记录" & vbNewLine & _
                        "            Where 病人id = [1] And 主页id = [2] And 科室id = [3]　And NVL(附加床位,0)=0 And 开始时间 IS NOT NULL) a" & vbNewLine & _
                        "Where ([4] Between a.开始时间 And a.终止时间 Or [4] >= a.开始时间) And a.床号 Is Not Null" & vbNewLine & _
                        "Order By a.开始时间"

            Set mrsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "当前床号", mlng病人ID, mlng主页ID, mlng科室ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            If mrsTemp.BOF = False Then mrsTemp.MoveLast
        Case "床位变动"
            strTmpSQL = "Select   a.床号" & vbNewLine & _
                        "From (Select 床号, 开始时间, Nvl(终止时间, To_Date('3000-01-01', 'yyyy-mm-dd')) As 终止时间" & vbNewLine & _
                        "            From 病人变动记录" & vbNewLine & _
                        "            Where 病人id = [1] And 主页id = [2] And 科室id = [3] And NVL(附加床位,0)=0 And 开始时间 IS NOT NULL) a" & vbNewLine & _
                        "Where (a.终止时间>=[4] And a.开始时间<=[5]) And a.床号 Is Not Null" & vbNewLine & _
                        "Order By a.开始时间"

            Set mrsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "当前床号", mlng病人ID, mlng主页ID, mlng科室ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            strCell = "": strBed = ""
            Do While Not mrsTemp.EOF
                If strBed <> mrsTemp.Fields(0).Value Then
                    strBed = mrsTemp.Fields(0).Value
                    strCell = strCell & "->" & mrsTemp.Fields(0).Value
                End If
            mrsTemp.MoveNext
            Loop
            strCell = Mid(strCell, 3)
            If mrsTemp.RecordCount > 0 Then mrsTemp.MoveFirst
        Case "当前科室"
        
            strTmpSQL = "Select   名称 From 部门表 a Where a.ID=[1]"
            Set mrsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "当前科室", mlng科室ID)
            
        Case "住院医师"
            strTmpSQL = "Select   a.经治医师" & vbNewLine & _
                        "From (Select 经治医师, 开始时间, Nvl(终止时间, To_Date('3000-01-01', 'yyyy-mm-dd')) As 终止时间" & vbNewLine & _
                        "            From 病人变动记录" & vbNewLine & _
                        "            Where 病人id = [1] And 主页id = [2] And 科室id = [3]　And NVL(附加床位,0)=0 And 开始时间 IS NOT NULL) a" & vbNewLine & _
                        "Where ([4] Between a.开始时间 And a.终止时间 Or [4] >= a.开始时间) And a.经治医师 Is Not Null" & vbNewLine & _
                        "Order By a.开始时间"
            Set mrsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "住院医师", mlng病人ID, mlng主页ID, mlng科室ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            If mrsTemp.BOF = False Then mrsTemp.MoveLast
        Case "责任护士"
        
            strTmpSQL = "Select   a.责任护士" & vbNewLine & _
                        "From (Select 责任护士, 开始时间, Nvl(终止时间, To_Date('3000-01-01', 'yyyy-mm-dd')) As 终止时间" & vbNewLine & _
                        "            From 病人变动记录" & vbNewLine & _
                        "            Where 病人id = [1] And 主页id = [2] And 科室id = [3]　And NVL(附加床位,0)=0 And 开始时间 IS NOT NULL) a" & vbNewLine & _
                        "Where ([4] Between a.开始时间 And a.终止时间 Or [4] >= a.开始时间) And a.责任护士 Is Not Null" & vbNewLine & _
                        "Order By a.开始时间"
            Set mrsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "责任护士", mlng病人ID, mlng主页ID, mlng科室ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            If mrsTemp.BOF = False Then mrsTemp.MoveLast
            
        Case "护理等级"
            strTmpSQL = "Select   b.名称" & vbNewLine & _
                        "From (Select 护理等级ID, 开始时间, Nvl(终止时间, To_Date('3000-01-01', 'yyyy-mm-dd')) As 终止时间" & vbNewLine & _
                        "            From 病人变动记录" & vbNewLine & _
                        "            Where 病人id = [1] And 主页id = [2] And 科室id = [3]　And NVL(附加床位,0)=0 And 开始时间 IS NOT NULL) a,护理等级 b" & vbNewLine & _
                        "Where ([4] Between a.开始时间 And a.终止时间 Or [4] >= a.开始时间) And a.护理等级ID Is Not Null And b.序号=a.护理等级ID" & vbNewLine & _
                        "Order By a.开始时间"
            Set mrsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "护理等级", mlng病人ID, mlng主页ID, mlng科室ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            If mrsTemp.BOF = False Then mrsTemp.MoveLast
            
        Case "最后诊断"
            strTmp = strPrefix
            gstrSQL = " Select f_List2str(Cast(Collect(Rownum || '、' || 诊断内容) As t_Strlist), ' ') As 诊断内容 from (Select 诊断内容 From ( Select  诊断内容 || Decode(Nvl(是否疑诊, 0), 0, '', ' (？)') 诊断内容 ,Mod(诊断类型, 10) 诊断类型, 标记时间 " & vbNewLine & _
                    "                   From 病人护理诊断 C" & vbNewLine & _
                    "                      Where 病人id = [1] And 主页id = [2] And 文件id = [3] And c.标记时间 Between [4] And [5])" & vbNewLine & _
                    "                       Group By 诊断内容 Order By Min(标记时间), Min(诊断类型)) "
            Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取设置的诊断", mlng病人ID, mlng主页ID, mlng当前文件ID, CDate(Format(aryPeriod(0), "YYYY-MM-DD hh:mm")), CDate(aryPeriod(1)))
            If NVL(mrsTemp!诊断内容) = "" Then
                gstrSQL = " Select f_List2str(Cast(Collect(Rownum || '、' || 诊断内容) As t_Strlist), ' ') As 诊断内容 from (Select 诊断内容 From ( Select 诊断内容 || Decode(Nvl(是否疑诊, 0), 0, '', ' (？)') 诊断内容, 诊断类型,标记时间" & vbNewLine & _
                    "                                                  From (Select  Distinct 诊断内容,诊断类型, 标记时间, 是否疑诊," & vbNewLine & _
                    "                                                                Rank() Over(Partition By 文件id Order By 标记时间 Desc) As Top" & vbNewLine & _
                    "                                                         From (Select 文件id,诊断内容, 标记时间, 诊断类型, 是否疑诊" & vbNewLine & _
                    "                                                                From 病人护理诊断 C" & vbNewLine & _
                    "                                                                Where 病人id = [1] And 主页id = [2] And 文件ID = [3] And C.标记时间 < [4] ))" & vbNewLine & _
                    "                                                  Where Top = 1) Order By 诊断类型, 标记时间) "
                    
                Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取设置的诊断", mlng病人ID, mlng主页ID, mlng当前文件ID, CDate(Format(aryPeriod(0), "YYYY-MM-DD hh:mm")), CDate(aryPeriod(1)))
            End If
            
            If NVL(mrsTemp!诊断内容) = "" Then
                strTmp = ""
                gstrSQL = "Select [1] || Zl_Replace_Element_Value([2],[3],[4],2,NULL,[5],[6]) as 信息 From Dual"
                Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取要素", strPrefix, strItemName, mlng病人ID, mlng主页ID, mint婴儿, CDate(aryPeriod(0)))
            End If
        Case Else
            If blnReplace = True Then
                strTmp = ""
                gstrSQL = "Select [1] || Zl_Replace_Element_Value([2],[3],[4],2,NULL,[5],[6]) as 信息 From Dual"
                Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取要素", strPrefix, strItemName, mlng病人ID, mlng主页ID, mint婴儿, CDate(aryPeriod(0)))
            Else
                strTmp = strPrefix
                gstrSQL = "Select 内容 From 病人护理要素内容 Where 文件ID=[1] And 页号=[2] And 名称=[3]"
                Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取要素", mlng当前文件ID, mint页码, strItemName)
            End If
        End Select
        
        If mrsTemp.BOF = False Then
            If strCell = "" Then
                If strTmp <> "" Then
                    lblSubhead.Tag = lblSubhead.Tag & " " & strTmp & mrsTemp.Fields(0).Value
                Else
                    lblSubhead.Tag = lblSubhead.Tag & " " & mrsTemp.Fields(0).Value
                End If
            Else
                lblSubhead.Tag = lblSubhead.Tag & " " & strTmp & strCell
            End If
        End If
    Next
    lblSubhead.Tag = Trim(lblSubhead.Tag)
    
    '表上标签分散处理
    Call zlLableBruit
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub zlRefresh(ByVal str条件 As String)
    Dim aryRow() As String, aryItem() As String
    Dim strPrefix As String, strItemName As String
    Dim lngRow As Long, lngCol As Long, lngCount As Long, strCell As String
    Dim strTmpSQL As String
    Dim strTmp As String
    
    Err = 0: On Error GoTo ErrHand
    
    '装入数据
    Call SQLCombination(str条件)
    gstrSQL = mstrSQL
    If gblnMoved Then
        gstrSQL = Replace(gstrSQL, "病人护理数据", "H病人护理数据")
        gstrSQL = Replace(gstrSQL, "病人护理明细", "H病人护理明细")
    End If
    Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取护理数据", mlng当前文件ID, mlng病人ID, mlng主页ID, mint婴儿, mint页码, mint结束页)
    '数据记录集,用于快速恢复
    Set mrsDataMap = CopyNewRec(mrsTemp, mrsDataMap)
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function CopyNewRec(ByVal rsSource As ADODB.Recordset, rsTarget As ADODB.Recordset) As ADODB.Recordset
    '只拷贝记录集的结构,同时增加页号,行号字段
    Dim intFields As Integer
    
    With rsTarget
        If .Fields.Count = 0 Then
            For intFields = 0 To rsSource.Fields.Count - 1
                If rsSource.Fields(intFields).Name = "汇总日期" Then
                    .Fields.Append rsSource.Fields(intFields).Name, adLongVarChar, 50, adFldIsNullable      '0:表示新增
                ElseIf rsSource.Fields(intFields).Type = 200 Then       '日期型处理为字符型
                    .Fields.Append rsSource.Fields(intFields).Name, adLongVarChar, rsSource.Fields(intFields).DefinedSize, adFldIsNullable     '0:表示新增
                Else
                    .Fields.Append rsSource.Fields(intFields).Name, IIf(rsSource.Fields(intFields).Type = adNumeric, adDouble, rsSource.Fields(intFields).Type), rsSource.Fields(intFields).DefinedSize, adFldIsNullable    '0:表示新增
                End If
            Next
            
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Open
        End If
        
        If rsSource.RecordCount <> 0 Then rsSource.MoveFirst
        Do While Not rsSource.EOF
            .AddNew
            For intFields = 0 To rsSource.Fields.Count - 1
                .Fields(intFields) = rsSource.Fields(intFields).Value
            Next
            .Update
            rsSource.MoveNext
        Loop
    End With
    
    Set CopyNewRec = rsTarget
End Function

Private Sub PreTendMutilRows()
    Dim arrData
    Dim intData As Integer, intDatas As Integer
    Dim lngRowCount As Long, lngRowCurrent As Long  '当前记录总行数,当前记录在本页的实际行数
    Dim lngCol As Long, lngMax As Long
    Dim lngRow As Long, lngStart As Long, lngPrintedRow As Long, lngLastRow As Long
    Dim str发生时间 As String, str发生时间_L As String
    Dim blnDelete As Boolean
    Dim strSignName As String
    Dim blnClear As Boolean
    Dim blnCollectType As Boolean  '记录正常数据行的上一行是否是汇总行
    Dim lngCurrRow As Long, lngCollectMutilRows As Long '汇总数据当前行、汇总列数据的行数
    Dim i As Integer, j As Integer, arrItem, arrCorrelative, arrLastRow, arrMutilRows '分类汇总项目数组
    
    On Error GoTo ErrHand
    '如果一行显示不完则分行显示(根据当前数据占用行数先添加空白行并处理行坐标,然后再依次处理当前行的数据)
    '每页只显示实际的数据行,把'@处取消注释即可
    '重新调整所有数据的实际行
    arrItem = Split(mstrColCorrelative, "|")
    lngRow = VsfData.FixedRows
    Do While True
        If lngRow > VsfData.Rows - 1 Then Exit Do
        If InStr(1, VsfData.TextMatrix(lngRow, mlngRowCount), "|") <> 0 Then Exit Do
        
        lngRowCount = Val(VsfData.TextMatrix(lngRow, mlngRowCount))
        lngRowCurrent = Val(VsfData.TextMatrix(lngRow, mlngRowCurrent))
        
        str发生时间 = Format(VsfData.TextMatrix(lngRow, 1), "YYYY-MM-DD HH:mm:ss")
        If Val(VsfData.TextMatrix(lngRow, mlngCollectType)) < 0 Then
            If blnCollectType = False Then str发生时间_L = "": blnCollectType = True
            '分类汇总子类明细数据的处理(数据保存方式是一条护理数据对应多条明细,明细中的记录组号不同)
            If str发生时间_L <> "" And str发生时间_L = str发生时间 Then
                If UBound(arrItem) < 0 Then '如果当前没有设置汇总关系,但之前数据存在分类汇总的情况，按子分类条数循环处理
                    lngCurrRow = lngLastRow + lngCollectMutilRows '确定每一条子数据输出的起始位置
                    lngCollectMutilRows = 1
                    If lngCurrRow < lngRow Then
                        VsfData.TextMatrix(lngCurrRow, mlngDate) = ""
                        VsfData.TextMatrix(lngCurrRow, mlngTime) = ""
                        
                        For lngCol = mlngTime + 1 To mlngNoEditor - 1
                            If (lngCol <> mlngSignTime And VsfData.ColHidden(lngCol) = False) Then
                                '准备赋值
                                With txtLength
                                    .Width = VsfData.ColWidth(lngCol)
                                    '这里需要注意一点：提取子类数据的行数应该是lngRow而不是lngCurrRow，因为在处理汇总总量记录时会导致子类数据的行位置发生变化 (行数 = 主记录开始行号 + 数据行数)
                                    .Text = Replace(Replace(Replace(VsfData.TextMatrix(lngRow, lngCol), Chr(10), ""), Chr(13), ""), Chr(1), "")
                                    .FontName = VsfData.CellFontName
                                    .FontSize = VsfData.CellFontSize
                                    .FontBold = VsfData.CellFontBold
                                    .FontItalic = VsfData.CellFontItalic
                                End With
                                arrData = GetData(txtLength.Text)
                                intDatas = UBound(arrData)
                                
                                If intDatas >= 0 Then
                                    '循环赋值
                                    If intDatas + 1 > lngRow - lngCurrRow Then intDatas = lngRow - lngCurrRow - 1
                                    If lngCollectMutilRows < intDatas + 1 Then lngCollectMutilRows = intDatas + 1
                                    For intData = 0 To intDatas
                                        VsfData.TextMatrix(lngCurrRow + intData, lngCol) = Replace(Replace(Replace(arrData(intData), Chr(10), ""), Chr(13), ""), Chr(1), "")
                                    Next
                                End If
                            End If
                        Next lngCol
                    End If
                    lngLastRow = lngCurrRow
                Else
                    '设置了分类汇总关系，按照每个汇总项目依次展示数据
                    For i = 0 To UBound(arrItem)
                        lngCurrRow = Val(arrLastRow(i)) + Val(arrMutilRows(i)) '按项目分类,确定每条数据输出的起始位置
                        lngCollectMutilRows = 1
                        arrMutilRows(i) = lngCollectMutilRows
                        If lngCurrRow < lngRow Then
                            arrCorrelative = Split(arrItem(i), ";")
                            For j = 0 To 1
                                '准备赋值
                                    lngCol = Split(arrCorrelative(j), ",")(0) + cHideCols + VsfData.FixedCols - 1
                                    With txtLength
                                        .Width = VsfData.ColWidth(lngCol)
                                        '这里需要注意一点：提取子类数据的行数应该是lngRow而不是lngCurrRow，因为在处理汇总总量记录时会导致子类数据的行位置发生变化 (行数 = 主记录开始行号 + 数据行数)
                                        .Text = Replace(Replace(Replace(VsfData.TextMatrix(lngRow, lngCol), Chr(10), ""), Chr(13), ""), Chr(1), "")
                                        .FontName = VsfData.CellFontName
                                        .FontSize = VsfData.CellFontSize
                                        .FontBold = VsfData.CellFontBold
                                        .FontItalic = VsfData.CellFontItalic
                                    End With
                                    arrData = GetData(txtLength.Text)
                                    intDatas = UBound(arrData)
                                    
                                    If intDatas >= 0 Then
                                        If intDatas + 1 > lngRow - lngCurrRow Then intDatas = lngRow - lngCurrRow - 1
                                        If lngCollectMutilRows < intDatas + 1 Then lngCollectMutilRows = intDatas + 1
                                        arrMutilRows(i) = lngCollectMutilRows
                                        For intData = 0 To intDatas
                                            VsfData.TextMatrix(lngCurrRow + intData, lngCol) = Replace(Replace(Replace(arrData(intData), Chr(10), ""), Chr(13), ""), Chr(1), "")
                                        Next intData
                                    End If
                            Next j
                        End If
                        arrLastRow(i) = lngCurrRow
                    Next i
                End If
                '赋值完成后移除原有子类行
                VsfData.RowPosition(lngRow) = VsfData.Rows - 1
                VsfData.RemoveItem VsfData.Rows - 1
                GoTo NextData
            Else
                '总量行默认为一行(只是针对汇总列的数据)
                lngCollectMutilRows = 1
                lngLastRow = lngRow '记录分类汇总总量行的位置
                '确定分类汇总首条子分类数据每个汇总项目的起始位置
                arrLastRow = Array(): arrMutilRows = Array()
                For i = 0 To UBound(arrItem)
                    ReDim Preserve arrLastRow(UBound(arrLastRow) + 1)
                    arrLastRow(UBound(arrLastRow)) = lngLastRow
                    ReDim Preserve arrMutilRows(UBound(arrMutilRows) + 1)
                    arrMutilRows(UBound(arrMutilRows)) = lngCollectMutilRows
                Next i
            End If
        Else
            If blnCollectType = True Then str发生时间_L = "": blnCollectType = False
            If str发生时间_L <> "" And Mid(str发生时间_L, 1, 16) = Mid(str发生时间, 1, 16) And str发生时间_L <> str发生时间 Then
                '日期相同，秒数不同，且不是汇总数据行，则说明这些数据是一组，更新lngDemo列
                VsfData.TextMatrix(lngRow, mlngDate) = ""
                VsfData.TextMatrix(lngRow, mlngTime) = ""
                VsfData.TextMatrix(lngRow, mlngDemo) = lngRow - lngLastRow + 1
                If lngRow - lngLastRow = Val(VsfData.TextMatrix(lngLastRow, mlngRowCount)) Then
                    VsfData.TextMatrix(lngLastRow, mlngDemo) = 1
                End If
            Else
                lngLastRow = lngRow
            End If
        End If
        
        If lngRowCount > 1 Then
            '先增加空行
            VsfData.Rows = VsfData.Rows + lngRowCount - 1
            '从当前行的下一行开始，每行的位置+所增加的空白行数，保证新增的空白行从当前行的下一行开始
            For intData = VsfData.Rows - lngRowCount To lngRow + 1 Step -1
                VsfData.RowPosition(intData) = intData + lngRowCount - 1
            Next
            
            '循环处理当前行数据
            For lngCol = 0 To VsfData.Cols - 1
                If VsfData.ColHidden(lngCol) And lngCol <> mlngRowCount And lngCol <> mlngDemo Then
                    '循环赋值
                    For intData = 2 To lngRowCount
                        VsfData.TextMatrix(lngRow + intData - 1, lngCol) = VsfData.TextMatrix(lngRow, lngCol)
                        '46506:刘鹏飞,2012-12-27,满页打印
                        '连续打印的情况才存在打印跨页数据后半部分内容
                        '打印结束页号为空说明之前未使用满页打印，跨页数据已经全部打印
                        '打印跨页数据后半部分内容为:打印页号+(当前行+打印行号-1)\本页数据行>打印结束页号
                        If lngCol = mlngPrintedPage And gintPrintState = 1 And Val(VsfData.TextMatrix(lngRow, mlngRowCount)) <> Val(VsfData.TextMatrix(lngRow, mlngRowCurrent)) Then
                            If Val(VsfData.TextMatrix(lngRow, mlngPrintedEndPage)) <> 0 And Val(VsfData.TextMatrix(lngRow, mlngPrintedPage)) <> 0 And Val(VsfData.TextMatrix(lngRow, mlngPrintedEndPage)) >= Val(VsfData.TextMatrix(lngRow, mlngPrintedPage)) Then
                                If Val(VsfData.TextMatrix(lngRow, mlngPrintedPage)) + (intData + lngPrintedRow - 2) \ mlngPageRows > Val(VsfData.TextMatrix(lngRow, mlngPrintedEndPage)) Then
                                    VsfData.TextMatrix(lngRow + intData - 1, lngCol) = ""
                                End If
                            ElseIf Val(VsfData.TextMatrix(lngRow, mlngPrintedEndPage)) <> 0 And Val(VsfData.TextMatrix(lngRow, mlngPrintedPage)) = 0 Then
                                '上次只打印了跨页数据跨页部分的
                                '例如：第3页数据跨页到第4页，之前只打印了第4跨页的数据。再次续打分两种情况：打印第3页，打印第4页
                                If Val(VsfData.TextMatrix(lngRow, mlngPrintedEndPage)) > mint页码 Then
                                    If Val(VsfData.TextMatrix(lngRow, mlngRowCurrent)) > intData Then
                                        VsfData.TextMatrix(lngRow + intData - 1, lngCol) = Val(VsfData.TextMatrix(lngRow, mlngPrintedEndPage))
                                    End If
                                Else
                                    If Val(VsfData.TextMatrix(lngRow, mlngRowCount)) - Val(VsfData.TextMatrix(lngRow, mlngRowCurrent)) < intData Then
                                        VsfData.TextMatrix(lngRow + intData - 1, lngCol) = Val(VsfData.TextMatrix(lngRow, mlngPrintedEndPage))
                                    End If
                                End If
                            End If
                        End If
                    Next
                ElseIf (lngCol < mlngNoEditor And lngCol <> mlngDate And lngCol <> mlngTime) Then
                    '准备赋值
                    With txtLength
                        .Width = VsfData.ColWidth(lngCol)
                        .Text = Replace(Replace(Replace(VsfData.TextMatrix(lngRow, lngCol), Chr(10), ""), Chr(13), ""), Chr(1), "")
                        .FontName = VsfData.CellFontName
                        .FontSize = VsfData.CellFontSize
                        .FontBold = VsfData.CellFontBold
                        .FontItalic = VsfData.CellFontItalic
                    End With
                    arrData = GetData(txtLength.Text)
                    intDatas = UBound(arrData)
                    
                    If intDatas > 0 Then
                        '循环赋值
                        If intDatas + 1 > lngRowCount Then intDatas = lngRowCount - 1
                        For intData = 0 To intDatas
                            If VsfData.Rows <= lngRow + intData Then VsfData.Rows = VsfData.Rows + 1
                            VsfData.TextMatrix(lngRow + intData, lngCol) = Replace(Replace(Replace(arrData(intData), Chr(10), ""), Chr(13), ""), Chr(1), "")
                        Next
                    End If
                ElseIf lngCol = mlngNoEditor Then
                    '将行值改为从1开始,比如有4行数据,就是4|1
                    For intData = 1 To lngRowCount
                        VsfData.TextMatrix(lngRow + intData - 1, mlngRowCount) = lngRowCount & "|" & intData
                    Next
                    '最后一行需要填写封闭签名
                    If mlngSignName > 0 Then VsfData.TextMatrix(lngRow + lngRowCount - 1, mlngSignName) = VsfData.TextMatrix(lngRow, mlngSignName)
                    If mlngSignTime > 0 Then VsfData.TextMatrix(lngRow + lngRowCount - 1, mlngSignTime) = VsfData.TextMatrix(lngRow, mlngSignTime)
                    '--58414,刘鹏飞,2013-04-10,添加护士、签名列显示模式
                    Call SingerShowType(VsfData, lngRow, lngRow + lngRowCount - 1)
                Else
                
                End If
            Next
            lngRow = lngRow + lngRowCount - 1
        Else
            VsfData.TextMatrix(lngRow, mlngRowCount) = "1|1"
        End If
        lngRow = lngRow + 1
NextData:
        str发生时间_L = str发生时间
    Loop
    
    '填充每页的启动行
    lngRow = VsfData.FixedRows
    
    Do While True
        '固定复制显示日期时间与签名列
        lngStart = GetStartRow(lngRow)
        
        '特殊处理第一行(第一行可能存在跨页数据)
        '50503:刘鹏飞,2012-09-12,只有开始行号<>1的才进行处理，避免处理掉了数据从记录单某页第一行就跨页的数据
        If lngRow = VsfData.FixedRows And Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0)) <> Val(VsfData.TextMatrix(lngRow, mlngRowCurrent)) And Val(VsfData.TextMatrix(lngRow, mlngStartSpread)) > 1 Then
            blnDelete = True
            lngRow = lngRow + Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0)) - Val(VsfData.TextMatrix(lngRow, mlngRowCurrent))
        End If
        
        If lngStart <> lngRow Or (Val(VsfData.TextMatrix(lngStart, mlngDemo)) > 1 And lngStart = lngRow) Then
            '--58414,刘鹏飞,2013-04-10,添加护士、签名列显示模式
            If Val(VsfData.TextMatrix(lngStart, mlngDemo)) > 1 Then
                For lngRowCount = lngStart To VsfData.FixedRows Step -1
                    If Val(VsfData.TextMatrix(lngRowCount, mlngDemo)) = 1 Then
                        If mlngDate > -1 Then VsfData.TextMatrix(lngRow, mlngDate) = VsfData.TextMatrix(lngRowCount, mlngDate)
                        If mlngTime > -1 Then VsfData.TextMatrix(lngRow, mlngTime) = VsfData.TextMatrix(lngRowCount, mlngTime)
                        Exit For
                    End If
                Next
            Else
                If mlngDate > -1 Then VsfData.TextMatrix(lngRow, mlngDate) = VsfData.TextMatrix(lngStart, mlngDate)
                If mlngTime > -1 Then VsfData.TextMatrix(lngRow, mlngTime) = VsfData.TextMatrix(lngStart, mlngTime)
            End If
            If lngStart <> lngRow Then
                '65994:刘鹏飞,2013-09-26,处理尾行签名,如果数据跨页则会导致两页都没有签名或只有第二页有签名
                If mlngSingerType <> 3 Then '非尾行签名，填写跨页数据在第二页数据的起始行
                    If mlngOperator <> -1 Then VsfData.TextMatrix(lngRow, mlngOperator) = VsfData.TextMatrix(lngStart, mlngOperator)
                    If mlngSignName <> -1 Then VsfData.TextMatrix(lngRow, mlngSignName) = VsfData.TextMatrix(lngStart, mlngSignName)
                    If mlngSignTime <> -1 Then VsfData.TextMatrix(lngRow, mlngSignTime) = VsfData.TextMatrix(lngStart, mlngSignTime)
                Else '尾行签名,填写跨页数据起始页的起始行(应为开始添加数据时,起始行已经清空)
                    If mlngOperator <> -1 Then VsfData.TextMatrix(lngStart, mlngOperator) = VsfData.TextMatrix(lngStart + Val(VsfData.TextMatrix(lngStart, mlngRowCount)) - 1, mlngOperator)
                    If mlngSignName <> -1 Then VsfData.TextMatrix(lngStart, mlngSignName) = VsfData.TextMatrix(lngStart + Val(VsfData.TextMatrix(lngStart, mlngRowCount)) - 1, mlngSignName)
                    If mlngSignTime <> -1 Then VsfData.TextMatrix(lngStart, mlngSignTime) = VsfData.TextMatrix(lngStart + Val(VsfData.TextMatrix(lngStart, mlngRowCount)) - 1, mlngSignTime)
                End If
                '更新当前页最后一行，跨页数据的签名人和签名时间
                If lngStart <> lngRow - 1 Then
                    '--58414,刘鹏飞,2013-04-10,添加护士、签名列显示模式
                    If mlngOperator <> -1 Then VsfData.TextMatrix(lngRow - 1, mlngOperator) = VsfData.TextMatrix(lngStart + Val(VsfData.TextMatrix(lngStart, mlngRowCount)) - 1, mlngOperator)
                    If mlngSignName <> -1 Then VsfData.TextMatrix(lngRow - 1, mlngSignName) = VsfData.TextMatrix(lngStart + Val(VsfData.TextMatrix(lngStart, mlngRowCount)) - 1, mlngSignName)
                    If mlngSignTime <> -1 Then VsfData.TextMatrix(lngRow - 1, mlngSignTime) = VsfData.TextMatrix(lngStart + Val(VsfData.TextMatrix(lngStart, mlngRowCount)) - 1, mlngSignTime)
                    Call SingerShowType(VsfData, lngStart, lngRow - 1)
                End If
            End If
        End If
        
        If blnDelete Then
            '89208:除之前重新调整数据行数后，在进行删除(以便处理后面63760问题中的签名人显示)
            lngRowCount = Val(VsfData.TextMatrix(lngStart, mlngRowCount))
            lngRowCount = lngRowCount - (lngRow - lngStart)
            If lngRowCount > 0 Then
                VsfData.TextMatrix(lngRow, mlngDemo) = VsfData.TextMatrix(lngStart, mlngDemo)
                For lngCol = 0 To lngRowCount - 1
                    VsfData.TextMatrix(lngCol + lngRow, mlngRowCount) = lngRowCount & "|" & lngCol + 1
                    VsfData.TextMatrix(lngCol + lngRow, mlngRowCurrent) = lngRowCount
                Next
            End If
            
            For lngCol = lngStart To lngRow - 1
                VsfData.RemoveItem lngStart
            Next
            
            blnDelete = False
            lngRow = VsfData.FixedRows  '只处理第一行记录删除的情况,所以固定设置为固定行为启始行
        End If
        
        lngRow = lngRow + mlngPageRows
        If lngRow > VsfData.Rows - 1 Then Exit Do
    Loop
    
    '63760:刘鹏飞,分组数据护士、签名人、签名时间的处理（同一个签名人始终显示一次）
    If mlngSingerType > 0 And VsfData.FixedRows <= VsfData.Rows - 1 Then
        lngPrintedRow = IIf(mlngOperator <> -1, mlngOperator, mlngSignName)
        lngRow = VsfData.FixedRows
        Do While True
            lngStart = GetStartRow(lngRow)
            lngRowCount = Val(VsfData.TextMatrix(lngStart, mlngRowCount))
            If lngRowCount <= 0 Then Exit Do
            
            If mlngSingerType = 3 Then '尾行签名
                strSignName = VsfData.TextMatrix(lngStart + lngRowCount - 1, lngPrintedRow)
            Else '首行签名或首尾签名
                strSignName = VsfData.TextMatrix(lngStart, lngPrintedRow)
            End If
            strSignName = FormatValue(strSignName)
            '检查是否是分组数据，从分组起始行开始处理
            If Val(VsfData.TextMatrix(lngStart, mlngDemo)) = 1 And lngStart = lngRow And strSignName <> "" Then
                For lngRow = lngStart + lngRowCount To VsfData.Rows - 1
                    If lngRow = lngStart + lngRowCount Then
                    
                        If Val(VsfData.TextMatrix(lngRow, mlngDemo)) <= 1 Then Exit For
                        
                        '检查同一分组不同数据之间的护士、签名人是否相同，并作相应的处理
                        lngRowCount = Val(VsfData.TextMatrix(lngRow, mlngRowCount))
                        If lngRowCount = 0 Then Exit For
                        
                        If mlngSingerType = 3 Then '尾行签名
                            '护士、签名人相同，只在本分组最后一条数据最后一行显示护士、签名人
                            If strSignName = FormatValue(VsfData.TextMatrix(lngRow + lngRowCount - 1, lngPrintedRow)) Then
                                '如果分组的本条数据刚好是某一页的首行则不清除上一页本分组最后一条数据的护士、签名人
                                If (lngRow - VsfData.FixedRows) Mod mlngPageRows > 0 And lngStart <= lngRow - 1 Then
                                    If mlngOperator <> -1 Then VsfData.TextMatrix(lngRow - 1, mlngOperator) = ""
                                    If mlngSignName <> -1 Then VsfData.TextMatrix(lngRow - 1, mlngSignName) = ""
                                    If mlngSignTime <> -1 Then VsfData.TextMatrix(lngRow - 1, mlngSignTime) = ""
                                End If
                            Else
                                If FormatValue(VsfData.TextMatrix(lngRow + lngRowCount - 1, lngPrintedRow)) <> "" Then
                                    strSignName = FormatValue(VsfData.TextMatrix(lngRow + lngRowCount - 1, lngPrintedRow))
                                End If
                            End If
                        Else '首行签名或首尾签名
                            If strSignName = FormatValue(VsfData.TextMatrix(lngRow, lngPrintedRow)) Then
                                '如果分组的本条数据刚好是某一页的首行则不清除上一页本分组最后一条数据的护士、签名人
                                If (lngRow - VsfData.FixedRows) Mod mlngPageRows > 0 Then
                                    blnClear = True
                                    '首尾签名需要注意：如果分组某条数据的首行(非起始行数据)在某页的最后一行，则不取消护士签名人的显示
                                    If mlngSingerType = 2 And lngRowCount = 1 Then
                                        If lngRow + lngRowCount < VsfData.Rows Then
                                            If Val(VsfData.TextMatrix(lngRow + lngRowCount, mlngDemo)) <= 1 Then
                                                blnClear = False
                                            End If
                                        Else
                                            blnClear = False
                                        End If
                                    End If
                                    
                                    If blnClear = True Then
                                        If mlngSingerType = 1 Or (mlngSingerType = 2 And (lngRow + 1 - VsfData.FixedRows) Mod mlngPageRows > 0) Then
                                            blnClear = False
                                            If mlngOperator <> -1 Then VsfData.TextMatrix(lngRow, mlngOperator) = ""
                                            If mlngSignName <> -1 Then VsfData.TextMatrix(lngRow, mlngSignName) = ""
                                            If mlngSignTime <> -1 Then VsfData.TextMatrix(lngRow, mlngSignTime) = ""
                                        End If
                                    End If
                                    
                                    If mlngSingerType = 2 And lngStart < lngRow - 1 Then '首尾签名还应该去掉上一条数据的尾行(上一行数据行数需要>1)
                                        '如果上一行数据跨页，并且在当前页只有一行则不进行清除本条数据的最后一行签名、护士
                                        If (lngRow - 1 - VsfData.FixedRows) Mod mlngPageRows > 0 Then
                                            If mlngOperator <> -1 Then VsfData.TextMatrix(lngRow - 1, mlngOperator) = ""
                                            If mlngSignName <> -1 Then VsfData.TextMatrix(lngRow - 1, mlngSignName) = ""
                                            If mlngSignTime <> -1 Then VsfData.TextMatrix(lngRow - 1, mlngSignTime) = ""
                                        End If
                                    End If
                                End If
                            Else
                                If FormatValue(VsfData.TextMatrix(lngRow, lngPrintedRow)) <> "" Then
                                    strSignName = FormatValue(VsfData.TextMatrix(lngRow, lngPrintedRow))
                                End If
                            End If
                        End If
                        
                        lngStart = lngRow
                    End If
                Next lngRow
            Else
                lngRow = lngStart + lngRowCount
            End If
            
            If lngRow > VsfData.Rows - 1 Then Exit Do
        Loop
    End If
    
    '如果是重打,将超出页有效数据行的部分删掉
    If gintPrintState = 2 Then
        If VsfData.Rows > VsfData.FixedRows + mlngPageRows Then
            VsfData.Rows = VsfData.FixedRows + mlngPageRows
        End If
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub PreTendFormat()
    Dim aryItem() As String
    Dim lngRow As Long, lngCol As Long, lngCount As Long, strCell As String
    Dim blnAlign As Boolean
    
    On Error GoTo ErrHand
    
    '设置护理记录单的格式
    With VsfData
        .Redraw = flexRDNone
        .Clear
        Set .DataSource = mrsDataMap
        
        '表头填写
        .MergeCells = flexMergeFixedOnly ' = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFree
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeRow(2) = True
        '先清空行合并属性，.Clear不能清除之前的合并信息
        For lngCount = .FixedRows To .Rows - 1
            .MergeRow(lngCount) = False
        Next
        '程序内部控制列隐藏
        .ColHidden(0) = True
        .ColHidden(1) = True
        .ColHidden(2) = True
        .ColHidden(mlngRowCount) = True
        .ColHidden(mlngRowCurrent) = True
        .ColHidden(mlngStartSpread) = True
        '51589:刘鹏飞,2013-03-01,添加交班签名
        .ColHidden(mlngJoinSignName) = True
        .ColHidden(mlngFileID) = True
        .ColHidden(mlngRecord) = True
        .ColHidden(mlngSigner) = True
        .ColHidden(mlngSignLevel) = True
        .ColHidden(mlngCollectStyle) = True
        .ColHidden(mlngCollectText) = True
        .ColHidden(mlngCollectType) = True
        .ColHidden(mlngCollectDay) = True
        .ColHidden(mlngPrintedPage) = True
        .ColHidden(mlngPrintedRow) = True
        .ColHidden(mlngPrintedTag) = True
        .ColHidden(mlngPrintedEndPage) = True
        .ColHidden(mlngCollectValue) = True
        '设置列头
        Dim strCOL As String
        Dim dblWidth As Double
        
        aryItem = Split(mstrTabHead, "|")
        For lngCount = 0 To UBound(aryItem)
            strCell = aryItem(lngCount)
            lngRow = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            lngCol = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            .TextMatrix(lngRow, lngCol + cHideCols + .FixedCols - 1) = strCell
        Next
        Call PreActiveHead
        
        '列宽设置
        blnAlign = False
        aryItem = Split(mstrColWidth, ",")
        If mbln日期时间合并 Then strCOL = "," & mlngDate & "," & mlngTime & ","
        For lngCount = cHideCols + .FixedCols To .Cols - 1
            If Not .ColHidden(lngCount) Then
                .ColWidth(lngCount) = Val(Split(aryItem(lngCount - cHideCols - .FixedCols), "`")(0))
                If mbln日期时间合并 And InStr(1, strCOL, "," & lngCount & ",") > 0 Then
                    dblWidth = dblWidth + .ColWidth(lngCount)
                End If
                If InStr(1, aryItem(lngCount - cHideCols - .FixedCols), "`") <> 0 Then
                    blnAlign = True
                    .ColAlignment(lngCount) = Val(Split(aryItem(lngCount - cHideCols - .FixedCols), "`")(1))
                End If
            End If
        Next
        '将发生时间列显示出来,列宽为日期与时间列的总宽度
        If mbln日期时间合并 Then
            .ColWidth(2) = IIf(dblWidth < 1600, 1600, dblWidth)
            .TextMatrix(0, 2) = "发生时间"
            If mintTabTiers >= 2 Then .TextMatrix(1, 2) = "发生时间"
            If mintTabTiers >= 3 Then .TextMatrix(2, 2) = "发生时间"
            .ColAlignment(2) = .ColAlignment(mlngDate)
        End If
        
        '固定行格式为居中
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        '再按列合并
        For lngCount = 0 To .Cols - 1
            .MergeCol(lngCount) = True
        Next
        
        If blnAlign = False Then
            '改为根据用户的设置显示列对齐方式
            If .FixedRows < .Rows Then .Cell(flexcpAlignment, .FixedRows, 0, .Rows - 1, .Cols - 1) = flexAlignGeneralCenter
        End If
        For lngCount = 0 To .Rows - 1
            If .RowHeight(lngCount) <> .RowHeightMin Then .RowHeight(lngCount) = .RowHeightMin
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
        
        Call PreTendMutilRows
        If mbln日期时间合并 Then
            '在PreTendMutilRows()中要处理数据,所以必须将列的隐藏属性放在这里设置
            .ColHidden(mlngDate) = True
            .ColHidden(mlngTime) = True
            .ColHidden(2) = False
        End If
        
        If mbln时间列隐藏 = True Then .ColHidden(mlngTime) = True
        
        Call WriteColor
        
        '可能固定行的行高不正确需要自动调整下
        .AutoResize = True
        .WordWrap = True
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 0, .Cols - 1
        .AutoResize = False
        '将非固定行的行高设置为最小行高
        For lngCount = 0 To .FixedRows - 1
            If .RowHeight(lngCount) < .RowHeightMin Then .RowHeight(lngCount) = .RowHeightMin
        Next
        For lngCount = .FixedRows To .Rows - 1
            .RowHeight(lngCount) = .RowHeightMin
        Next
        
        .Redraw = flexRDDirect
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function PrintHead() As Boolean
    Dim lngPage As Long
    On Error GoTo ErrHand
    
    lngPage = mint页码
    mlng当前页码 = lngPage
    PrintHead = PrintRTBData(rtbHead, True, lngPage)
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function PrintFoot() As Boolean
    Dim lngPage As Long
    On Error GoTo ErrHand
    
    lngPage = mint页码
    mlng当前页码 = lngPage
    PrintFoot = PrintRTBData(rtbFoot, False, lngPage)
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function PrintRTBData(ByVal objRTB As RichTextBox, ByVal blnHead As Boolean, Optional ByVal lngPage As Long = 0) As Boolean
    Dim fr As FORMATRANGE           '格式化的文本范围
    Dim rcDrawTo As RECT            '目标文字区域
    Dim rcPage As RECT              '目标页面区域
    Dim rcHeand As RECT
    Dim rcFoot As RECT
    Dim gTargetDC As Long
    Dim lngFoot As Long
    Dim lngOffsetLeft As Long
    Dim lngOffsetTop As Long
    Dim lngNextPos As Long, lngLen As Long, lngTmp As Long, lngPageCount As Long
    Dim mrsTemp As New ADODB.Recordset
    Dim lngPageIndex As Long, lngPrintTextY As Long
    
    lngLen = lstrlen(objRTB.Text)
    lngOffsetLeft = gobjOutTo.ScaleX(GetDeviceCaps(gobjOutTo.hDC, PHYSICALOFFSETX), vbPixels, vbTwips)
    lngOffsetTop = gobjOutTo.ScaleY(GetDeviceCaps(gobjOutTo.hDC, PHYSICALOFFSETY), vbPixels, vbTwips)
    
    '46251,刘鹏飞,2012-09-11,装载页码输出位置
    lngPageIndex = Val(cbo页码.ItemData(cbo页码.ListIndex))
    If lngPageIndex <= 0 Or lngPageIndex > 4 Then lngPageIndex = 4
    If blnHead Then
        If chk页码.Value = 1 And (lngPageIndex = 1 Or lngPageIndex = 2) Then
            lngFoot = gobjOutTo.TextHeight("第")
            With rcHeand
                .Left = lngOffsetLeft
                .Right = gobjOutTo.Width - lngOffsetLeft
                If lngPageIndex = 1 Then
                    lngPrintTextY = lngOffsetTop + 30
                    .Top = lngOffsetTop + lngFoot + 60
                    .Bottom = gobjOutTo.ScaleX(gobjSend.EmptyUp, vbMillimeters, vbTwips)
                Else
                    .Top = lngOffsetTop
                    .Bottom = gobjOutTo.ScaleX(gobjSend.EmptyUp, vbMillimeters, vbTwips) - lngFoot - 60
                    lngPrintTextY = .Bottom + 30
                End If
                If lngPrintTextY < lngOffsetTop + 30 Then lngPrintTextY = lngOffsetTop + 30
            End With
        Else
            With rcHeand
                .Left = lngOffsetLeft
                .Top = lngOffsetTop
                .Right = gobjOutTo.Width - lngOffsetLeft
                .Bottom = gobjOutTo.ScaleX(gobjSend.EmptyUp, vbMillimeters, vbTwips) - 30
            End With
            gobjOutTo.Print ""
        End If
    Else
        '62436:刘鹏飞,2013-06-20,修改页码输出坐标，保证在打印机可打印区域内。
        If chk页码.Value = 1 And (lngPageIndex = 3 Or lngPageIndex = 4) Then
            If lngPageIndex = 3 Then
                lngFoot = gobjOutTo.TextHeight("第") + 60
                lngPrintTextY = gobjOutTo.Height - gobjOutTo.ScaleY(gobjSend.EmptyDown, vbMillimeters, vbTwips) - lngOffsetTop * 2
                rcFoot.Bottom = gobjOutTo.Height
            Else
                lngFoot = gobjOutTo.TextHeight("第") + 60
                lngPrintTextY = gobjOutTo.Height - lngOffsetTop * 2 - lngFoot
                rcFoot.Bottom = lngPrintTextY
            End If
            If lngPrintTextY + lngFoot > gobjOutTo.Height - lngOffsetTop * 2 Then lngPrintTextY = gobjOutTo.Height - lngOffsetTop * 2 - lngFoot
            If lngPageIndex = 4 Then lngFoot = 0
        Else
            gobjOutTo.Print ""
            lngFoot = 0
            rcFoot.Bottom = gobjOutTo.Height
        End If
        With rcFoot
            .Left = lngOffsetLeft
            .Top = gobjOutTo.Height - lngOffsetTop * 2 - gobjOutTo.ScaleY(gobjSend.EmptyDown, vbMillimeters, vbTwips) + lngFoot
            .Right = gobjOutTo.Width - lngOffsetLeft
        End With
    End If
    
    gTargetDC = hDC
    With rcPage
        .Left = 0
        .Top = 0
        .Right = gobjOutTo.Width
        .Bottom = gobjOutTo.Height
    End With
    With rcDrawTo
        If blnHead Then
            .Left = rcHeand.Left
            .Top = rcHeand.Top
            .Right = rcHeand.Right
            .Bottom = rcHeand.Bottom
        Else
            .Left = rcFoot.Left
            .Top = rcFoot.Top
            .Right = rcFoot.Right
            .Bottom = rcFoot.Bottom
        End If
    End With
    With fr
        .hDC = gobjOutTo.hDC
        .hdcTarget = gTargetDC
        .rc = rcDrawTo
        .rcPage = rcPage
        .chrg.cpMin = 0
        .chrg.cpMax = -1
    End With
    
    Do
        lngNextPos = SendMessage(objRTB.hWnd, EM_FORMATRANGE, 0, fr)
        
        lngPageCount = lngPageCount + 1             ' 页数＋1
        '记录分页信息
        ReDim Preserve AllPages(1 To lngPageCount) As PageInfo
        AllPages(lngPageCount).PageNumber = lngPageCount
        AllPages(lngPageCount).ActualHeight = fr.rc.Bottom - fr.rc.Top          '实际打印高度
        AllPages(lngPageCount).Start = lngTmp
        AllPages(lngPageCount).End = lngNextPos
        
        fr.chrg.cpMin = lngNextPos
        If lngNextPos <= lngTmp Or lngNextPos >= lngLen Then Exit Do      ' 完成所有页面的分页
        lngTmp = lngNextPos
    Loop
    Call SendMessage(objRTB.hWnd, EM_FORMATRANGE, 0, ByVal CLng(0))
    
    For lngLen = 1 To lngPageCount
        If lngLen > 1 Then Exit For
        With fr
            .hDC = gobjOutTo.hDC
            .hdcTarget = gTargetDC
            .rc = rcDrawTo
            .rcPage = rcPage
            .chrg.cpMin = AllPages(lngLen).Start
            .chrg.cpMax = AllPages(lngLen).End
        End With
        Call SendMessage(objRTB.hWnd, EM_FORMATRANGE, 1, fr)
        Call SendMessage(objRTB.hWnd, EM_FORMATRANGE, 0, ByVal CLng(0))
    Next
    
    '最后在输出页码
    If chk页码.Value = 1 And ((blnHead = True And (lngPageIndex = 1 Or lngPageIndex = 2)) Or (blnHead = False And (lngPageIndex = 3 Or lngPageIndex = 4))) Then
        gobjOutTo.CurrentY = lngPrintTextY
        If optPageAlign(0).Value Then
            gobjOutTo.CurrentX = gobjOutTo.ScaleX(gobjSend.EmptyLeft, vbMillimeters, vbTwips) - 30
        ElseIf optPageAlign(1).Value Then
            gobjOutTo.CurrentX = (gobjOutTo.Width - 90 * LenB(StrConv("第 " & lngPage & " 页", vbFromUnicode))) / 2
        Else
            gobjOutTo.CurrentX = gobjOutTo.Width - gobjOutTo.ScaleX(gobjSend.EmptyRight, vbMillimeters, vbTwips) - 90 * LenB(StrConv("页码:" & mint页码, vbFromUnicode))
        End If
        gobjOutTo.Print "第 " & lngPage & " 页"
    End If
End Function

Public Function PrintPage(Optional blnOddEvenPrint As Boolean = False, Optional ArrSQL As Variant) As Boolean
    Dim strSQL() As String
    Dim blnTrans As Boolean
    Dim blnSave As Boolean          '已打印的数据不保存
    Dim strTime As String
    Dim strCurrDate As String
    Dim lngRow As Long, lngRows As Long
    Dim intMax As Integer, intPos As Integer
    Dim lngCurRow As Long, lngDataLines As Long
    Dim intTag As Integer
    Dim int结束页码 As Integer
    Dim lngFileID As Long
    
    ReDim Preserve strSQL(1 To 1)
    On Error GoTo ErrHand
    
    '56134:刘鹏飞,2012-12-19,病人护理打印添加打印标识
    If mblnPrintRow = True Then
        intTag = 1
    Else
        intTag = 0
        If gintPrintState = 1 And glngPrintRow > 0 And Val(VsfData.TextMatrix(glngPrintRow, VsfData.Cols - 7)) > 0 Then
            intTag = 1
        End If
    End If
    
    '对显示行进行处理
    lngRows = VsfData.Rows - 1
    For lngRow = VsfData.FixedRows To lngRows
        If Not VsfData.RowHidden(lngRow) Then
            If lngCurRow = 0 Then lngCurRow = 1
            If FormatValue(VsfData.TextMatrix(lngRow, mlngRowCount)) Like "*|1" Then
                lngDataLines = Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0)
                blnSave = (Val(VsfData.TextMatrix(lngRow, mlngPrintedPage)) = 0) Or gintPrintState > 1
                '重打的话保持原有的结束页码
                int结束页码 = IIf(mlng当前页码 = 0, mint页码, mlng当前页码)
                If blnSave Then
                    strTime = VsfData.TextMatrix(lngRow, 1)
                    lngFileID = Val(VsfData.TextMatrix(lngRow, mlngFileID))
                    gstrSQL = "ZL_病人护理打印_PRINT(" & lngFileID & ",to_date('" & strTime & "','yyyy-MM-dd hh24:mi:ss'),'" & gstrUserName & "'," & _
                        IIf(mlng当前页码 = 0, mint页码, mlng当前页码) & "," & lngCurRow & "," & intTag & "," & int结束页码 & ")"
                    'Debug.Print gstrSQL
                    strSQL(ReDimArray(strSQL)) = gstrSQL
                End If
            End If
            '46506:刘鹏飞,2012-12-28,记录单满页打印
            '每一页的首行如果不是本条数据的起始行,就说明是跨页数据
            If lngCurRow = 1 And Not FormatValue(VsfData.TextMatrix(lngRow, mlngRowCount)) Like "*|1" Then
                strTime = VsfData.TextMatrix(lngRow, 1)
                lngFileID = Val(VsfData.TextMatrix(lngRow, mlngFileID))
                gstrSQL = "ZL_病人护理打印_PRINT(" & lngFileID & ",to_date('" & strTime & "','yyyy-MM-dd hh24:mi:ss'),'" & gstrUserName & "'," & _
                         "NULL,NULL," & intTag & "," & IIf(mlng当前页码 = 0, mint页码, mlng当前页码) & ",1)"
                strSQL(ReDimArray(strSQL)) = gstrSQL
            End If
            lngCurRow = lngCurRow + 1
        End If
    Next
    
    '如果是奇偶打印，并且打印页数不为1,就返回数据SQL
    If blnOddEvenPrint = True Then
        ArrSQL = strSQL
        PrintPage = True
        Exit Function
    End If
    
    On Error Resume Next
    intMax = UBound(strSQL)

    gcnOracle.BeginTrans
    blnTrans = True

    On Error GoTo ErrHand
    If intMax > 0 Then
        For intPos = 1 To intMax
            If strSQL(intPos) <> "" Then
                gcnOracle.Execute strSQL(intPos), , adCmdStoredProc
            End If
        Next
    End If

    gcnOracle.CommitTrans
    blnTrans = False
    PrintPage = True
    Exit Function
ErrHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub InitRecords()
    Dim i As Integer, j As Integer, k As Integer, l As Integer
    Dim lngCol As Long, lngOrder As Long, strName As String, intImmovable As Integer, strFormat As String
    Dim arrColumn, arrItem, arrCorrelative(), strColumns As String
    Dim blnSet As Boolean
    
    On Error GoTo ErrHand
    
    strColumns = mstrColumns
    If Not mblnInit Then
        '初始化内存记录集(未对应项目的列为活动项目,其它列均为固定项)
        strFields = "列," & adDouble & ",18|序号," & adDouble & ",2|项目序号," & adDouble & ",18|项目名称," & adLongVarChar & ",20|固定," & adDouble & ",2|格式," & adLongVarChar & ",2000"
        Call Record_Init(mrsSelItems, strFields)
        strFields = "列|序号|项目序号|项目名称|固定|格式"
    End If
    
    '加入列定义
    If Not mblnInit Then
        arrColumn = Split(strColumns, "|")
        j = UBound(arrColumn)
        For i = 0 To j
            lngCol = Split(arrColumn(i), "'")(0)
            arrItem = Split(Split(arrColumn(i), "'")(1), ",")
            blnSet = False   '如果已设置以传入值为准'否则找不到项目就是活动项目
            If UBound(Split(arrColumn(i), "'")) > 1 Then
                blnSet = True
                intImmovable = Split(arrColumn(i), "'")(2)
            End If
            If UBound(Split(arrColumn(i), "'")) > 2 Then
                strFormat = Split(arrColumn(i), "'")(3)
            End If
            
            k = UBound(arrItem)
            For l = 0 To k
                strName = arrItem(l)
                mrsItems.Filter = "项目名称='" & strName & "'"
                If mrsItems.RecordCount <> 0 Then
                    lngOrder = mrsItems!项目序号
                    If Not blnSet Then intImmovable = 1   '固定不允许修改
                Else
                    lngOrder = 0
                    If Not blnSet Then intImmovable = 0
                    
                    '记录特殊列
                    Select Case strName
                    Case "日期"
                        mlngDate = i + cHideCols + VsfData.FixedCols
                    Case "时间"
                        mlngTime = i + cHideCols + VsfData.FixedCols
                    Case "护士"
                        mlngOperator = i + cHideCols + VsfData.FixedCols
                    Case "签名人"
                        mlngSignName = i + cHideCols + VsfData.FixedCols
                    Case "签名时间"
                        mlngSignTime = i + cHideCols + VsfData.FixedCols
                    End Select
                End If
                strValues = lngCol & "|" & l + 1 & "|" & lngOrder & "|" & strName & "|" & intImmovable & "|" & strFormat
                Call Record_Add(mrsSelItems, strFields, strValues)
            Next
        Next
        
        '整理分类汇总关联列信息
        arrCorrelative = Array()
        arrColumn = Split(mstrColCorrelative, "|")
        For i = 0 To UBound(arrColumn)
            arrItem = Split(arrColumn(i), ";")
            If UBound(arrItem) = 1 Then
                mrsSelItems.Filter = "列=" & Val(arrItem(0))
                If mrsSelItems.RecordCount = 1 Then
                    ReDim Preserve arrCorrelative(UBound(arrCorrelative) + 1)
                    arrCorrelative(UBound(arrCorrelative)) = Val(arrItem(0)) & "," & mrsSelItems!项目序号 & ";" & CStr(arrItem(1))
                End If
            End If
        Next i
        If UBound(arrCorrelative) = -1 Then
            mstrColCorrelative = ""
        Else
            mstrColCorrelative = Join(arrCorrelative, "|")
        End If
        mrsSelItems.Filter = ""
        
        'Call OutputRsData(mrsSelItems)
        
        '加入程序内部控制列(列是在读取数据后绑定时增加的,此时只有预处理下)
        mlngSignLevel = VsfData.Cols + cHideCols + VsfData.FixedCols '加上隐藏列
        mlngSigner = mlngSignLevel + 1
        '51589:刘鹏飞,2013-03-01,添加交班签名
        mlngJoinSignName = mlngSigner + 1
        mlngFileID = mlngJoinSignName + 1
        mlngRecord = mlngFileID + 1
        mlngRowCount = mlngRecord + 1
        mlngRowCurrent = mlngRowCount + 1
        mlngStartSpread = mlngRowCurrent + 1 '50503:刘鹏飞,2012-09-12
        mlngPrintedEndPage = mlngStartSpread + 1 '46506:刘鹏飞,2012-12-27
        mlngCollectValue = mlngPrintedEndPage + 1  '陈刘,105302
        mlngPrintedTag = mlngCollectValue + 1 '56134:刘鹏飞,2012-12-19
        mlngCollectType = mlngPrintedTag + 1
        mlngCollectText = mlngCollectType + 1
        mlngCollectStyle = mlngCollectText + 1
        mlngCollectDay = mlngCollectStyle + 1
        mlngPrintedPage = mlngCollectDay + 1
        mlngPrintedRow = mlngPrintedPage + 1
        
        
        If mlngOperator <> -1 And mlngSignName <> -1 Then
            mlngNoEditor = IIf(mlngOperator < mlngSignName, mlngOperator, mlngSignName)
        Else
            mlngNoEditor = IIf(mlngOperator <> -1, mlngOperator, mlngSignName)
        End If
    End If
    
    mrsItems.Filter = 0
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function ShowPage(Optional ByVal intPage As Integer = 0) As Boolean
    '显示指定页面数据并更新打印对象
    Dim aryPeriod
    Dim strBegin As String, strEnd As String
    Dim lngRow As Long, lngRows As Long, lngStart As Long
    Dim lngOffsetLeft As Long, lngScaleWidth As Long
    Dim lngShows As Long
    '活动项目相关变量
    Dim mrsTemp As New ADODB.Recordset
    Dim blnPrintRow As Boolean
    On Error GoTo ErrHand
    
    If intPage <> 0 Then mlngMinIndex = intPage - Val(mArrPage(0))
    If mlngMinIndex > mlngMaxIndex Then mlngMinIndex = mlngMaxIndex
    
    mint页码 = Val(mArrPage(mlngMinIndex))
    If InStr(1, CStr(mArrPage(mlngMinIndex)), ";") <> 0 Then
        gintPrintState = Val(Split(CStr(mArrPage(mlngMinIndex)), ";")(1))
    Else
        gintPrintState = 2
    End If
    
    lngOffsetLeft = Printer.ScaleX(GetDeviceCaps(Printer.hDC, PHYSICALOFFSETX), vbPixels, vbTwips)
    
    Call LoadPageData '加载相应页打印数据
    
    With VsfData
        '小于页面有效数据行说明只有一页数据
        If VsfData.Rows - VsfData.FixedRows > mlngPageRows Then
            lngRows = .Rows - 1
            For lngRow = .FixedRows To lngRows
                .RowHidden(lngRow) = True
            Next
        End If
        
        '小于页面有效数据行说明只有一页数据
        If VsfData.Rows - VsfData.FixedRows > mlngPageRows Then
            lngRow = 3 + mlngPageRows * (mint页码 - mint当前起始页)
            lngRows = 3 + mlngPageRows * (mint页码 - mint当前起始页 + 1) - 1
        Else
            lngRow = 3
            lngRows = .Rows - 1
        End If
        If lngRows > .Rows - 1 Then lngRows = .Rows - 1
        '获取指定页的时间范围
        If lngRow > lngRows Then
            Exit Function
        End If
        strBegin = Format(.TextMatrix(lngRow, 1), "YYYY-MM-DD HH:mm:ss")
        lngStart = lngRows
        lngStart = GetStartRow(lngStart)
        strEnd = .TextMatrix(lngStart, 1)
        If Not IsDate(strEnd) And lngStart <> lngRow Then
            lngStart = lngRow
            strEnd = .TextMatrix(lngStart, 1)
        End If
        strEnd = Format(strEnd, "YYYY-MM-DD HH:mm") & ":59"
        '53588:刘鹏飞,2013-4-25,修改数据的时间小于病人入院时间，床号，病区不能显示问题
        '如：病人入科时间为2013-03-13 11:23:34 文件开始时间和入科相同，此时录入数据时间为 2013-03-13 11:23
        '就会导致无法提取床号，应为保存的数据时间为2013-03-13 11:23:00 小于了病人入科时间导致无法提取到数据
        '获取病人的入院时间
        If mint婴儿 = 0 Then
            gstrSQL = "Select 开始时间, Sysdate As 结束时间" & vbNewLine & _
                " From 病人变动记录" & vbNewLine & _
                " Where 病人id = [1] And 主页id = [2] And 开始原因 = 2" & vbNewLine & _
                " Union All" & vbNewLine & _
                " Select 开始时间, Sysdate As 结束时间" & vbNewLine & _
                " From 病人变动记录 a" & vbNewLine & _
                " Where a.病人id = [1] And a.主页id = [2] And a.开始原因 = 1 And Not Exists" & vbNewLine & _
                " (Select 1 From 病人变动记录 Where 病人id = a.病人id And 主页id = a.主页id And 开始原因 = 2)"
        Else
            gstrSQL = " Select   出生时间 AS 开始时间 From 病人新生儿记录 Where 病人ID=[1] And 主页ID=[2] And 序号=[3]"
        End If
        Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取入院日期或出生日期", mlng病人ID, mlng主页ID, mint婴儿)
        If Format(strBegin, "yyyy-MM-dd HH:mm:ss") < Format(mrsTemp!开始时间, "yyyy-MM-dd HH:mm:ss") Then
            strBegin = Format(mrsTemp!开始时间, "yyyy-MM-dd HH:mm:ss")
        End If
        aryPeriod = Split(strBegin & "||" & strEnd, "||")
        
        lngStart = lngRow
        '小于页面有效数据行说明只有一页数据
        If VsfData.Rows - VsfData.FixedRows > mlngPageRows Then
            '显示数据行
            For lngRow = lngRow To lngRows
                .RowHidden(lngRow) = False
                lngShows = lngShows + 1
            Next
        End If
        
        '56134:刘鹏飞,2012-12-19
        '不足一页的输出剩余表格(只有最后一页才会有此情况)
        '重打和连续重打的情况，只要mblnPrintRow=True，并且数据不足一页，都需要输出表格
        '连续打印时mblnPrintRow=True，并且数据不足一页，如果本页未打印就输出表格，如果已经打印过说明之前已经输出过表格了。
        '预览的情况，只要mblnPrintRow=True都输出表格
        If mblnPrintRow = True And lngRows - lngStart + 1 < mlngPageRows Then
            blnPrintRow = False
            If gblnPrintMode = False Then '预览
                blnPrintRow = True
            Else '打印
                If gintPrintState = 1 Then '连续打印
                    '如果本页之前没有打印过，就输入表格
                    If glngPrintRow >= lngStart And glngPrintRow <= lngRows Then
                        blnPrintRow = (Val(.TextMatrix(glngPrintRow, mlngPrintedTag)) = 0)
                    Else
                        blnPrintRow = True
                    End If
                Else '重打或连续重打
                    blnPrintRow = True
                End If
            End If
            If blnPrintRow = True Then
                VsfData.Rows = VsfData.Rows + mlngPageRows - (lngRows - lngStart + 1)
                For lngRow = lngRows To VsfData.Rows - 1
                    VsfData.RowHeight(lngRow) = VsfData.RowHeightMin
                Next
            End If
        End If
        
        ShowPage = True
        Call zlReadTip(aryPeriod)
    End With
    
    '设置打印相关内容
    Dim objPrint As New zlTFPrintTends, objAppRow As zlTFTabAppRow
    Dim strLable As String, strAppRow As String, lngSpaces As Long
    Dim lngPos As Long, lngMax As Long, lngNumber As Long, blnNumber As Boolean, lngASC As Long
    
    '设置打印格式
    If UBound(Split(mstrPaperSet, ";")) >= 0 Then SaveSetting "ZLSOFT", "公共模块\zl9TendFile\Default", "PaperSize", Val(Split(mstrPaperSet, ";")(0))
    If UBound(Split(mstrPaperSet, ";")) >= 1 Then SaveSetting "ZLSOFT", "公共模块\zl9TendFile\Default", "Orientation", Val(Split(mstrPaperSet, ";")(1))
    If UBound(Split(mstrPaperSet, ";")) >= 2 Then SaveSetting "ZLSOFT", "公共模块\zl9TendFile\Default", "Height", Val(Split(mstrPaperSet, ";")(2))
    If UBound(Split(mstrPaperSet, ";")) >= 3 Then SaveSetting "ZLSOFT", "公共模块\zl9TendFile\Default", "Width", Val(Split(mstrPaperSet, ";")(3))
    If UBound(Split(mstrPaperSet, ";")) >= 4 Then objPrint.EmptyLeft = Round(ScaleY(Val(Split(mstrPaperSet, ";")(4)), vbTwips, vbMillimeters), 2)
    If UBound(Split(mstrPaperSet, ";")) >= 5 Then objPrint.EmptyRight = Round(ScaleY(Val(Split(mstrPaperSet, ";")(5)), vbTwips, vbMillimeters), 2)
    If UBound(Split(mstrPaperSet, ";")) >= 6 Then objPrint.EmptyUp = Round(ScaleX(Val(Split(mstrPaperSet, ";")(6)), vbTwips, vbMillimeters), 2)
    If UBound(Split(mstrPaperSet, ";")) >= 7 Then objPrint.EmptyDown = Round(ScaleX(Val(Split(mstrPaperSet, ";")(7)), vbTwips, vbMillimeters), 2)
    
    On Error Resume Next
    Printer.PaperSize = Val(Split(mstrPaperSet, ";")(0))
    Printer.Orientation = Val(Split(mstrPaperSet, ";")(1))
    
    If Printer.PaperSize = 256 Then
        Call SetCustonPager(Val(Split(mstrPaperSet, ";")(3)), Val(Split(mstrPaperSet, ";")(2)))
    End If
    
    On Error GoTo ErrHand
    Set objPrint.Body = VsfData
    objPrint.Title.Text = lblTitle.Caption
    Set objPrint.Title.Font = lblTitle.Font
    Set objPrint.AppFont = lblSubhead.Font
    
    lngSpaces = lblSubhead.Height / 210
    strLable = lblSubhead.Caption
    '60333:刘鹏飞,2013-10-14,修改文件编号导致打印信息丢失
    If UBound(Split(mstrPaperSet, ";")) >= 4 Then
        lngScaleWidth = Printer.Width - (lngOffsetLeft + Val(Split(mstrPaperSet, ";")(4))) * 2
    Else
        lngScaleWidth = Printer.Width - (lngOffsetLeft) * 2
    End If
    lngMax = Len(strLable)
    lngNumber = 0
    lngStart = 1
    For lngPos = 1 To lngMax
        '如果数学超长,则把数字移到下一行显示
        lngASC = Asc(Mid(strLable, lngPos, 1))

        '检查是否超宽(长度超过行宽,或者遇到回车换行符)
        If TextWidth(Mid(strLable, lngStart, lngPos - lngStart + 1) & "测") > lngScaleWidth Or lngPos = lngMax Or lngASC = 10 Then
            If lngPos = lngMax Or lngASC = 10 Then
                strAppRow = Mid(strLable, lngStart, lngPos - lngStart + 1)
            Else
                strAppRow = Mid(strLable, lngStart, lngPos - lngStart - 1) & "…"
            End If
            lngStart = lngPos + 1
            
            '输出表上项
            Set objAppRow = New zlTFTabAppRow
            Call objAppRow.Add(strAppRow)
            Call objPrint.UnderAppRows.Add(objAppRow)
            
            If lngPos = lngMax Or lngASC = 10 Then
            Else
                Exit For        '护理记录单表上标签超长也只打一行，表上标签的行数变化会影响表格打印行，所以固定不允许添加
            End If
        End If
    Next
    '60333:刘鹏飞,2013-10-14,修改文件编号导致打印信息丢失
    If UBound(Split(mstrPaperSet, ";")) >= 3 Then
        lngMax = Val(Split(mstrPaperSet, ";")(3))
    Else
        lngMax = Printer.Width
    End If
'    If mstrPageHead <> "" Then objPrint.Header = mstrPageHead
'    If mstrPageFoot <> "" Then
'        mstrPageFoot = Replace(mstrPageFoot, "{打印时间}", Now)
'        mstrPageFoot = Replace(mstrPageFoot, "{页码}", mint页码 + mint合并起始页 - 1)
'        mstrPageFoot = Replace(mstrPageFoot, "{打印人}", gstrUserName)
'        objPrint.Footer = LeftB(mstrPageFoot & Space(lngMax), lngMax - objPrint.EmptyLeft - objPrint.EmptyRight)
'    End If
    
    Set gobjSend = objPrint

    '保存标题的属性
    gstrTabTitle = gobjSend.Title.Text
    gstrTitleFName = gobjSend.Title.Font.Name
    gintTitleFSize = gobjSend.Title.Font.Size
    gblnTitleFItalic = gobjSend.Title.Font.Italic
    gblnTitleFBold = gobjSend.Title.Font.Bold
    glngTitleColor = gobjSend.Title.Color
    '保存表上项目与表下项目的属性
    gstrAppRowFName = gobjSend.AppFont.Name
    gintAppRowFSize = gobjSend.AppFont.Size
    gblnAppRowFItalic = gobjSend.AppFont.Italic
    gblnAppRowFBold = gobjSend.AppFont.Bold
    glngAppRowColor = gobjSend.AppColor
    gintUpAppRow = gobjSend.UnderAppRows.Count
    gintDownAppRow = gobjSend.BelowAppRows.Count
    
    If gobjSend.FixRow = 0 Then gobjSend.FixRow = gobjSend.Body.FixedRows
    gintFixRow = gobjSend.FixRow
    gintFixCol = gobjSend.FixCol
'    gintRowTotal = gobjSend.Rows
'    gintColTotal = gobjSend.Cols
    gintGroups = 1
    
    gsngDown = gobjSend.EmptyDown
    gsngLeft = gobjSend.EmptyLeft
    gsngRight = gobjSend.EmptyRight
    gsngUp = gobjSend.EmptyUp
    gsngHeader = gobjSend.PageHeader
    gsngFooter = gobjSend.PageFooter
    
    gstrHeader = gobjSend.Header
    gstrHeader = IIf(gstrHeader = "", ";;", gstrHeader)
    gstrFooter = gobjSend.Footer
    gstrFooter = IIf(gstrFooter = "", ";;", gstrFooter)
    
    '最多不过一列就是一页面
    Call GetPrinterSet
    Call CalculateHeight
    Call CalculateRC
    gstr对角线 = GetDiagonal
    glngHideCols = cHideCols
    glngSignName = IIf(mblnSignPic = True, mlngSignName, -1)
    '64583:刘鹏飞,2013-09-22,打印时同一个日期是否重复显示
    glngDate = IIf(mblnDateModel = True, mlngDate, -1)
    glngCollectColor = mlngCollectColor
    
    ShowPage = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetFixedProperty(ByVal strName As String) As Variant
'根据名称获取固定属性
    Dim varProperty As Variant
    Select Case strName
        Case "有效数据行"
            varProperty = mlngPageRows
        Case "汇总颜色"
            varProperty = mlngCollectColor
    End Select
    GetFixedProperty = varProperty
End Function

Public Function GetFixedCol(ByVal strName As String) As Long
'根据名称获取固定列信息
    Dim lngCol As Long
    Select Case strName
        Case "日期"
            lngCol = mlngDate
        Case "时间"
            lngCol = mlngTime
        Case "护士"
            lngCol = mlngOperator
        Case "签名人"
            lngCol = mlngSignName
        Case "签名时间"
            lngCol = mlngSignTime
        Case "签名级别"
            lngCol = mlngSignLevel
        Case "签名信息"
            lngCol = mlngSigner
        Case "交班签名人"
            lngCol = mlngJoinSignName
        Case "文件ID"
            lngCol = mlngFileID
        Case "记录ID"
            lngCol = mlngRecord
        Case "行数"
            lngCol = mlngRowCount
        Case "实际行数"
            lngCol = mlngRowCurrent
        Case "开始行号"
            lngCol = mlngStartSpread
        Case "打印结束页号"
            lngCol = mlngPrintedEndPage
        Case "汇总值"
            lngCol = mlngCollectValue
        Case "打印标识"
            lngCol = mlngPrintedTag
        Case "汇总类别"
            lngCol = mlngCollectType
        Case "汇总文本"
            lngCol = mlngCollectText
        Case "汇总标记"
            lngCol = mlngCollectStyle
        Case "汇总日期"
            lngCol = mlngCollectDay
        Case "打印页号"
            lngCol = mlngPrintedPage
        Case "打印行号"
            lngCol = mlngPrintedRow
        Case "禁止编辑"
            lngCol = mlngNoEditor
    End Select
    GetFixedCol = lngCol
End Function

Public Function GetStartPage() As Integer
    If UBound(mArrPage) < 0 Then
        GetStartPage = 1
    Else
        GetStartPage = Val(mArrPage(0))
    End If
End Function

Public Function GetCollectCols(ByVal lngRaw As Long) As String
    GetCollectCols = VsfData.TextMatrix(lngRaw, mlngCollectValue)
End Function

Public Function GetPages() As Integer
    If UBound(mArrPage) < 0 Then
        GetPages = 1
    Else
        GetPages = mlngMaxIndex + Val(mArrPage(0))
    End If
End Function

Public Function isEndPage() As Boolean
    isEndPage = (mlngMinIndex = mlngMaxIndex)
End Function

Public Sub PrevPage()
    If mlngMinIndex > 0 Then
        mlngMinIndex = mlngMinIndex - 1
        If mlngMinIndex <= UBound(mArrPage) Then
            Call ShowPage
        End If
    End If
End Sub

Public Function NextPage() As Boolean
    If mlngMinIndex < mlngMaxIndex Then
        mlngMinIndex = mlngMinIndex + 1
        If mlngMinIndex <= UBound(mArrPage) Then
            NextPage = ShowPage
        End If
    End If
End Function

Public Function AppointPage(ByVal intPage As Integer) As Boolean
    If UBound(mArrPage) >= 0 Then
        If intPage <= mlngMaxIndex + Val(mArrPage(0)) Then
            mlngMinIndex = intPage - Val(mArrPage(0))
            AppointPage = ShowPage
        End If
    End If
End Function

Public Function GetFileName() As String
    GetFileName = lblTitle.Caption
End Function

Public Function blnOddEvenPagePrint() As Boolean
    blnOddEvenPagePrint = mblnOddEvenPagePrint
End Function

Public Function blnShowNullCollet() As Boolean
    blnShowNullCollet = mblnShowNullCollet
End Function

Private Sub WriteColor()
    Dim blnTag As Boolean
    Dim lngCount As Long
    Dim lngRow As Long, lngCol As Long
    On Error GoTo ErrHand
    '晚班以红色显示,打印过的页号字体设置为灰色
    
    glngPrintRow = 0
    With VsfData
        For lngCount = .FixedRows To .Rows - 1
            If .TextMatrix(lngCount, 1) <> "" Then
                If .TextMatrix(lngCount, mlngPrintedPage) <> "" And gintPrintState = 1 Then
                    .Cell(flexcpForeColor, lngCount, 0, lngCount, .Cols - 1) = &HE0E0E0
                    glngPrintRow = lngCount                 '记录下该坐标,从此开始向下打印
                Else
                    '以第一条未打印的数据为当前显示页
                    If lngRow = 0 And gintPrintState = 1 Then
                        lngRow = lngCount
                        mint页码 = (lngCount - VsfData.FixedRows) \ mlngPageRows + mint当前起始页 - 1
                        If (lngCount - VsfData.FixedRows) > mlngPageRows Then
                            If (lngCount - VsfData.FixedRows) Mod mlngPageRows <> 0 Then mint页码 = mint页码 + 1
                        End If
                        If mint页码 < mint当前起始页 Then mint页码 = mint当前起始页
                    End If
                    
                    If Val(.TextMatrix(lngCount, mlngCollectType)) = 0 Then
                        '晚班以红色显示
                        blnTag = False
                        If mintTagFormHour < mintTagToHour Then
                            blnTag = (Hour(.TextMatrix(lngCount, 1)) >= mintTagFormHour And Hour(.TextMatrix(lngCount, 1)) < mintTagToHour)
                        Else
                            blnTag = (Hour(.TextMatrix(lngCount, 1)) >= mintTagFormHour Or Hour(.TextMatrix(lngCount, 1)) < mintTagToHour)
                        End If
                        If blnTag Then
                            Set .Cell(flexcpFont, lngCount, 0, lngCount, .Cols - 1) = mobjTagFont
                            .Cell(flexcpForeColor, lngCount, 0, lngCount, .Cols - 1) = mlngTagColor
                        End If
                    End If
                End If
                '处理小结的显示
                '65889:刘鹏飞,2013-11-1,处理小结行跨页的情况，保证下一页小结首行也能正确显示小结名称，而不是数据发生时间
                '添加 (lngCount - .FixedRows + 1) Mod mlngPageRows = 1
                If FormatValue(VsfData.TextMatrix(lngCount, mlngRowCount)) Like "*|1" Or (lngCount - .FixedRows + 1) Mod mlngPageRows = 1 Then
                    If Val(VsfData.TextMatrix(lngCount, mlngCollectType)) <> 0 Then
                        VsfData.TextMatrix(lngCount, mlngDate) = VsfData.TextMatrix(lngCount, mlngCollectText)
                        If mbln时间列隐藏 = False Then
                            VsfData.TextMatrix(lngCount, mlngTime) = VsfData.TextMatrix(lngCount, mlngCollectText)
                        End If
            
                        '88967:护士、签名列同时存在，且属于同一操作员，则应避免合并(打印签名人输出签名图片需注意签名人有回车的情况)
                        For lngCol = mlngTime + 1 To IIf(mlngNoEditor < mlngSignName, mlngSignName, mlngNoEditor)
                            '52953,刘鹏飞,2012-08-24,汇总数据为0也要显示,关联问题:60792
                            'If .TextMatrix(lngCount, lngCOL) = "0" Then .TextMatrix(lngCount, lngCOL) = ""
                            If Trim(.TextMatrix(lngCount, lngCol)) <> "" And .ColHidden(lngCol) = False Then
                                '66085:刘鹏飞,2012-09-26,避免相邻汇总列合并,将原来的列内容+空格同一改成在列后面在chr(13)
                                '避免因加空格后列宽不够导致内容显示不完全(主要针对右对其)
'                                Select Case .ColAlignment(lngCol)
'                                    Case 6, 7, 8
'                                        .TextMatrix(lngCount, lngCol) = IIf(lngCol Mod 2 = 1, " ", "") & .TextMatrix(lngCount, lngCol)
'                                    Case 3, 4, 5
'                                        .TextMatrix(lngCount, lngCol) = IIf(lngCol Mod 2 = 1, " ", "") & .TextMatrix(lngCount, lngCol) & IIf(lngCol Mod 2 = 1, " ", "")
'                                    Case 0, 1, 2
'                                        .TextMatrix(lngCount, lngCol) = .TextMatrix(lngCount, lngCol) & IIf(lngCol Mod 2 = 1, " ", "")
'                                    Case Else
'                                        .TextMatrix(lngCount, lngCol) = IIf(lngCol Mod 2 = 1, " ", "") & .TextMatrix(lngCount, lngCol)
'                                End Select
                                .TextMatrix(lngCount, lngCol) = .TextMatrix(lngCount, lngCol) & IIf(lngCol Mod 2 = 1, Chr(13), "")
                            End If
                        Next
                        .MergeRow(lngCount) = True
                    End If
                End If
            End If
        Next
        
        '如果未赋值,取最后一页
        If (lngRow = 0 And gintPrintState = 1) Then
            mint页码 = (.Rows - VsfData.FixedRows) \ mlngPageRows + mint当前起始页 - 1
            If (.Rows - VsfData.FixedRows) > mlngPageRows Then
                If (.Rows - VsfData.FixedRows) Mod mlngPageRows <> 0 Then mint页码 = mint页码 + 1
            End If
            If mint页码 < mint当前起始页 Then mint页码 = mint当前起始页
        End If
        If mint页码 = 0 Then mint页码 = 1
                        
        '如果页码>当前起始页,说明起始页无效
        If gintPrintState = 1 Then
            '如果当前页大于起始页,删除无效页数据
            If mint页码 > mint当前起始页 Then
                For lngRow = 1 To mlngPageRows
                    VsfData.RemoveItem VsfData.FixedRows
                Next
                mint当前起始页 = mint页码
                glngPrintRow = glngPrintRow - mlngPageRows
                If glngPrintRow < VsfData.FixedRows Then glngPrintRow = 0
            End If
            '如果起始行超过一页,删除无效页数据
            If lngRow >= VsfData.FixedRows + mlngPageRows Then
                For lngRow = 1 To mlngPageRows
                    VsfData.RemoveItem VsfData.FixedRows
                Next
                glngPrintRow = glngPrintRow - mlngPageRows
                If glngPrintRow < VsfData.FixedRows Then glngPrintRow = 0
                mint当前起始页 = mint当前起始页 + 1
                mint页码 = mint页码 + 1
            End If
            If mint结束页 > mint页码 And VsfData.Rows - VsfData.FixedRows <= mlngPageRows Then
                mint结束页 = mint页码
            End If
        End If
        
        '将日期为空的行的发生时间也设置为空(分组数据)
        If mbln日期时间合并 Then
            lngRow = VsfData.FixedRows
            Do While True
                If lngRow > VsfData.Rows - 1 Then Exit Do
                If VsfData.TextMatrix(lngRow, mlngDate) = "" Then
                    VsfData.TextMatrix(lngRow, 1) = ""
                    VsfData.TextMatrix(lngRow, 2) = ""
                Else
                    If Val(VsfData.TextMatrix(lngRow, mlngCollectType)) <> 0 Then
                        VsfData.TextMatrix(lngRow, 2) = VsfData.TextMatrix(lngRow, mlngDate)
                    Else
                        VsfData.TextMatrix(lngRow, 1) = Format(VsfData.TextMatrix(lngRow, 1), "yyyy-MM-dd HH:mm")
                        VsfData.TextMatrix(lngRow, 2) = Format(VsfData.TextMatrix(lngRow, 1), "yyyy-MM-dd HH:mm")
                    End If
                End If
                lngRow = lngRow + 1
            Loop
        End If
        
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub zlLableBruit()
    Dim lngScaleLeft As Long, lngScaleTop As Long, lngScaleRight As Long, lngScaleBottom As Long
    
    lblSubhead.Top = lblTitle.Top + lblTitle.Height + 120
    lblSubhead.Width = VsfData.Width
    lblSubhead.Caption = lblSubhead.Tag
    VsfData.Move lngScaleLeft + 210, lblSubhead.Top + lblSubhead.Height + 45, ScaleWidth - lngScaleLeft - 210 * 2
    VsfData.Height = picMain.Height - VsfData.Top
End Sub

Private Sub GetFileProperty()
    '提取文件属性
    On Error GoTo ErrHand
    
    gstrSQL = " Select   开始时间,结束时间,格式ID,科室ID,归档人 From 病人护理文件 " & _
              " Where 病人ID=[1] And 主页ID=[2] And 婴儿=[3] And ID=[4] And Rownum<2"
    If gblnMoved Then
        gstrSQL = Replace(gstrSQL, "病人护理文件", "H病人护理文件")
    End If
    Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取护理文件数据", mlng病人ID, mlng主页ID, mint婴儿, mlng当前文件ID)
    If mrsTemp.RecordCount <> 0 Then
        mlng格式ID = mrsTemp!格式ID
        mlng科室ID = mrsTemp!科室ID
        mstr开始时间 = Format(mrsTemp!开始时间, "yyyy-MM-dd HH:mm:ss")
        mstr结束时间 = Format(mrsTemp!结束时间, "yyyy-MM-dd HH:mm:ss")
    End If
    
    RaiseEvent AfterRowColChange("", False, mblnSign, mblnArchive)
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitEnv()
    On Error GoTo ErrHand
    
     '46251,刘鹏飞,2012-09-11,装载页码输出位置
    With cbo页码
        .Clear
        .AddItem "页眉上方": .ItemData(.NewIndex) = 1
        .AddItem "页眉下方": .ItemData(.NewIndex) = 2
        .AddItem "页脚上方": .ItemData(.NewIndex) = 3
        .AddItem "页脚下方": .ItemData(.NewIndex) = 4
        cbo页码.Tag = 3
        Call zlControl.CboSetIndex(cbo页码.hWnd, 2)
    End With
    
    '打开现存在的所有护理记录项目
    gstrSQL = " Select   项目序号,upper(项目名称) AS 项目名称,项目类型,项目性质,项目长度,项目小数,项目表示,项目单位,项目值域,护理等级,应用方式" & _
              " From 护理记录项目 B" & _
              " Order by 项目序号"
    Set mrsItems = zlDatabase.OpenSQLRecord(gstrSQL, "打开现存在的所有护理记录项目")
    
    '提取适用于记录单的诊治所见项目
    gstrSQL = _
        " Select i.分类id, i.编码, i.中文名, nvl(i.替换域,0) 替换域,i.类型,i.长度,i.小数,i.单位,i.表示法,i.数值域,i.必填" & vbNewLine & _
        " From 诊治所见项目 i, 诊治所见分类 k" & vbNewLine & _
        " Where k.Id = i.分类id And ((k.编码 In ('02', '05', '06') And i.替换域 = 1) Or (k.性质 = 2 And k.编码 = '06' And NVL(i.替换域,0) = 0))" & vbNewLine & _
        " Order By k.性质, k.编码, i.编码"
    Set mrsElement = zlDatabase.OpenSQLRecord(gstrSQL, "提取适用于记录单的诊治所见项目")
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function ShowMe(ByVal frmParent As Form, ByVal lngFileID As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal intBaby As Integer, Optional ByVal strPages As String = "") As Boolean
    '******************************************************************************************************************
    '功能： 显示护理记录文件内容
    '参数： frmParent           上级窗体对象
    '       lngFileID           文件ID
    '       lngPatiID           病人id
    '       lngPageID           主页id
    '       intBaby             婴儿标志
    '       strPage             为空说明从第一页开始进行打印,不为空格式为：页码;标识(续打或正常打印),页码;标识......
    '返回： 无
    '******************************************************************************************************************
    Dim mrsTemp As New ADODB.Recordset
    Dim i As Long
    Dim arrTemp() As String
    On Error GoTo ErrHand
    Err = 0
    
    mArrPage = Array(): mlngMinIndex = -1: mlngMaxIndex = -1
    mblnInit = False
    mlng当前文件ID = lngFileID
    mlng病人ID = lngPatiID
    mlng主页ID = lngPageId
    mint婴儿 = intBaby
    mlngPageRows = frmAsk.mintPageRows
    Set mfrmParent = frmParent
    mintNORule = Val(zlDatabase.GetPara("护理文件页码规则", glngSys, 1255, 0))
    mblnSignPic = (Val(zlDatabase.GetPara("记录单签名人显示方式", glngSys, 1255, 0)) = 1)
    '56134:刘鹏飞,2012-12-19,记录单打印时,数据未满页空白部分输出表格
    mblnPrintRow = (Val(zlDatabase.GetPara("记录单未满页打印表格", glngSys, 1255, 0)) = 1)
    '46506:刘鹏飞,2012-12-19,记录单打印时，数据满页才进行输出(文件为结束时有效)
    mblnFullPagePrint = (Val(zlDatabase.GetPara("记录单满页打印", glngSys, 1255, 0)) = 1)
    '49753:刘鹏飞,2012-12-19,记录单打印时，数据页奇偶输出
    mblnOddEvenPagePrint = (Val(zlDatabase.GetPara("记录单奇偶打印", glngSys, 1255, 0)) = 1)
    '--58414,刘鹏飞,2013-04-10,添加护士、签名列显示模式
    mlngSingerType = Val(zlDatabase.GetPara("护士、签名列显示模式", glngSys, 1255, "2"))
    If InStr(1, ",0,1,2,3,", "," & mlngSingerType & ",") = 0 Then mlngSingerType = 2
    '64583:刘鹏飞,2013-09-22,预览、打印时同一页相同日期显示方式:多次;一次
    mblnDateModel = (Val(zlDatabase.GetPara("记录单日期显示方式", glngSys, 1255, 0)) = 1)
    '68739:刘鹏飞,2014-1-2,添加"小结标识颜色"
    mlngCollectColor = Val(zlDatabase.GetPara("小结标识颜色", glngSys, 1255, "255"))
    
    arrTemp = Split(zlDatabase.GetPara("小结缺省格式", glngSys, 1255), ";")
    If UBound(arrTemp) > 0 Then
        mblnShowNullCollet = arrTemp(1) = 0
    Else
        mblnShowNullCollet = True
    End If
    '判断病人是否转出
    gstrSQL = "Select 数据转出 From 病案主页 Where 病人ID=[1] And 主页ID=[2]"
    Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断数据是否转出", mlng病人ID, mlng主页ID)
    gblnMoved = NVL(mrsTemp!数据转出, 0) <> 0
    
    '如果时批量打印，则进行全部打印
    If gblnBatch = False Then
        '判断当前文件是否已经结束
        gstrSQL = " Select  结束时间 From 病人护理文件 " & _
                  " Where 病人ID=[1] And 主页ID=[2] And 婴儿=[3] And ID=[4] And Rownum<2"
        Call SQLDIY(gstrSQL)
        Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取护理文件数据", mlng病人ID, mlng主页ID, mint婴儿, mlng当前文件ID)
        If mrsTemp.RecordCount > 0 Then
            '如果文件已经结束,不管记录单是否满页都进行打印
            If Trim(NVL(mrsTemp!结束时间)) <> "" Then mblnFullPagePrint = False
        End If
    Else
        mblnFullPagePrint = False
        mblnOddEvenPagePrint = False
    End If
    If mblnFullPagePrint = True Then mblnPrintRow = False
    
    If mrsItems.State = 0 Then
        Call InitEnv            '初始化环境
    End If
    Call InitVariable
    
    mstrMergeID = ""
    gstrSQL = " Select MIN(开始页号) AS 开始页号,MAX(结束页号) AS 结束页号 From 病人护理打印 Where 文件ID=[1]"
    Call SQLDIY(gstrSQL)
    Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取文件最小和最大页码", mlng当前文件ID)
    mint结束页 = Val(NVL(mrsTemp!结束页号, 0))
    mint页码 = Val(NVL(mrsTemp!开始页号, 0))
    If mint结束页 = 0 Then Exit Function
    
    If strPages <> "" Then
        For i = 0 To UBound(Split(strPages, ","))
            If Val(Split(strPages, ",")(i)) >= mint页码 And Val(Split(strPages, ",")(i)) <= mint结束页 Then
                ReDim Preserve mArrPage(UBound(mArrPage) + 1)
                mArrPage(UBound(mArrPage)) = Split(strPages, ",")(i)
            End If
        Next i
    Else
        For i = mint页码 To mint结束页
            ReDim Preserve mArrPage(UBound(mArrPage) + 1)
            mArrPage(UBound(mArrPage)) = i & ";2"
        Next i
    End If
   
    '如果是满页打印页，并且选择的打印最后一页等于文件最后页号，则检查是否满页
    If mblnFullPagePrint = True And mint结束页 = Val(mArrPage(UBound(mArrPage))) Then
        gstrSQL = "Select Max(结束行号) 结束行号 From 病人护理打印 where 文件ID=[1] And 结束页号=[2] "
        Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取文件最后一页的结束行号", mlng当前文件ID, mint结束页)
        If mrsTemp!结束行号 < mlngPageRows Then
            If UBound(mArrPage) <= 0 Then
                mArrPage = Array()
            Else
                ReDim Preserve mArrPage(UBound(mArrPage) - 1)
            End If
        End If
    End If
    If UBound(mArrPage) < 0 Then
        If gblnBatch = False Then
            MsgBox "没有满足条件或可打印的护理记录单数据！", vbInformation, gstrSysName
        End If
        Exit Function
    End If
    
    '奇偶打印时需检查页码是否连续，不连续则不进行奇偶打印
    If mblnOddEvenPagePrint = True Then
        mint页码 = Val(mArrPage(0)) - 1
        For i = 0 To UBound(mArrPage)
            If mint页码 + 1 <> Val(mArrPage(i)) Then
                mblnOddEvenPagePrint = False
                Exit For
            End If
            mint页码 = Val(mArrPage(i))
        Next i
    End If
    
    mlngMinIndex = 0: mlngMaxIndex = UBound(mArrPage)
    mint页码 = Val(mArrPage(mlngMinIndex))
    mint结束页 = Val(mArrPage(UBound(mArrPage)))
    
    mstrMergeID = ""
    '提取合并文件信息
    gstrSQL = _
        "Select Id From (With 病人护理文件_F1 As" & vbNewLine & _
        " (Select a.Id, a.续打id From 病人护理文件 a Where a.病人id = [1] And a.主页id = [2] And Nvl(a.婴儿, 0) = [3])" & vbNewLine & _
        "Select Id From 病人护理文件_F1 Start With 续打id = [4] Connect By Prior Id = 续打id Order By Level Desc)"
    Call SQLDIY(gstrSQL)
    Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查当前文件是否与其它文件设置为合并打印", mlng病人ID, mlng主页ID, mint婴儿, mlng当前文件ID)
    Do While Not mrsTemp.EOF
        mstrMergeID = mstrMergeID & "," & mrsTemp!ID
    mrsTemp.MoveNext
    Loop
    mstrMergeID = Mid(mstrMergeID, 2)
    Call ShowPage
    mblnInit = True
    mblnEditable = False
    ShowMe = True
'    Call OutputRsData(mrsSelItems)
    Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function LoadPageData() As Boolean
    Dim str条件 As String, lng结束页 As Long, lngFileID As Long
    Dim blnInitRec As Boolean
    Dim arrCode, i As Integer
    On Error GoTo ErrHand
    
    mint当前起始页 = mint页码
    Set mrsDataMap = New ADODB.Recordset
    lngFileID = mlng当前文件ID
    '合并文件处理
    blnInitRec = False
    arrCode = Split(mstrMergeID, ",")
    If UBound(arrCode) >= 0 And mint页码 = Val(mArrPage(0)) Then
        For i = 0 To UBound(arrCode)
            gstrSQL = "Select MAX(结束页号) 结束页 From 病人护理打印 Where 文件ID=[1]"
            Call SQLDIY(gstrSQL)
            Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取合并文件最大结束页号", Val(arrCode(i)))
            lng结束页 = NVL(mrsTemp!结束页, 0)
            If lng结束页 = mint页码 Then
                mlng当前文件ID = Val(arrCode(i))
                Call ReadStruDef
                If Not blnInitRec Then
                    Call InitRecords
                    blnInitRec = True
                End If
                str条件 = " And P.结束页号=[5]"
                Call zlRefresh(str条件)
            End If
        Next i
    End If
    mlng当前文件ID = lngFileID
    '要打印的文件处理
    Call ReadStruDef
    If Not blnInitRec Then
        Call InitRecords
        blnInitRec = True
    End If
    str条件 = " AND (P.开始页号=[5] OR (P.结束页号=[5])) "
    Call zlRefresh(str条件)
    Call PreTendFormat
    LoadPageData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub lblTitle_Click()
    Call NextPage
'    Call PrevPage
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    
    With picMain
        .Top = 0
        .Left = 0
        .Width = ScaleWidth
        .Height = ScaleHeight
    End With
    
    Call zlLableBruit
End Sub

Private Sub vsfData_DrawCell(ByVal hDC As Long, ByVal ROW As Long, ByVal COL As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Call DrawCell(hDC, ROW, COL, Left, Top, Right, Bottom, Done)
End Sub

Private Sub InitVariable()
    '连续重打常量
    mlngDate = -1
    mlngTime = -1
    mlngOperator = -1
    mlngSigner = -1
    mlngSignTime = -1
    mlngSignName = -1
    mlngFileID = -1
    mlngRecord = -1
    mlngNoEditor = -1
    mlngPrintedEndPage = -1
    mlngCollectValue = -1
    
    mblnShow = False
    mblnSign = False
    mblnArchive = False
    mblnEditAssistant = False
    
    Set mrsDataMap = New ADODB.Recordset
End Sub

Private Function GetStartRow(ByVal lngRow As Long) As Long
    Dim lngStart As Long
    Dim lngCurRows As Long, lngRows As Long
    '提取数据起始行,超出本页则返回0
    '如果本页未显示全,则说明超出本页,也返回0
    '不允许在连续的数据行中插入新行
    
    If VsfData.TextMatrix(lngRow, mlngRowCount) = "" Then VsfData.TextMatrix(lngRow, mlngRowCount) = "1|1"
    lngRows = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0))    '总行数
    lngCurRows = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(1)) '当前行
    If lngCurRows = 1 Then
        GetStartRow = lngRow
        Exit Function
    End If
    
    '寻找起始行
    For lngRow = lngRow To 3 Step -1
        If FormatValue(VsfData.TextMatrix(lngRow, mlngRowCount)) = lngRows & "|1" Then
            lngStart = lngRow
            Exit For
        End If
    Next
    
    GetStartRow = lngStart
End Function

Public Function GetDiagonal() As String
    GetDiagonal = "," & mstrCatercorner & "," '& mstrCOLNothing & ","
End Function

Private Function IsDiagonal(ByVal intCol As Integer) As Boolean
    Dim arrCol, arrData
    Dim intDo As Integer, intCount As Integer
    '判断指定列是否设置了列对角线（mstrColWidth的格式：765`11`1`1,765`11`2`1,...，对象属性`对象序号`列对角线）
    
    IsDiagonal = (InStr(1, "," & mstrCatercorner & "," & mstrCOLNothing & ",", "," & intCol - (cHideCols + VsfData.FixedCols - 1) & ",") <> 0)
End Function


'######################################################################################################################
'**********************************************************************************************************************
'以下是基础函数或过程

Private Sub picMain_Resize()
    On Error Resume Next
    picMain.Left = 0
    
    lblTitle.Left = 0
    lblTitle.Width = picMain.Width
    
    VsfData.Width = picMain.Width - VsfData.Left * 2
    VsfData.Height = picMain.Height - VsfData.Top
End Sub

Private Sub UserControl_GotFocus()
    On Error Resume Next
    VsfData.SetFocus
End Sub

Private Sub UserControl_Initialize()
    mblnShow = False
    mblnInit = False
    
'    Set objStream = objFileSys.OpenTextFile("C:\WORKLOG.txt", ForAppending, True)
End Sub

Private Sub UserControl_Terminate()
'    objStream.Close
    If Not mrsTemp Is Nothing Then
        If mrsTemp.State = adStateOpen Then mrsTemp.Close
        Set mrsTemp = Nothing
    End If
    If Not mrsItems Is Nothing Then
        If mrsItems.State = adStateOpen Then mrsItems.Close
        Set mrsItems = Nothing
    End If
    If Not mrsElement Is Nothing Then
        If mrsElement.State = adStateOpen Then mrsElement.Close
        Set mrsElement = Nothing
    End If
    If Not mrsSelItems Is Nothing Then
        If mrsSelItems.State = adStateOpen Then mrsSelItems.Close
        Set mrsSelItems = Nothing
    End If
    If Not mrsDataMap Is Nothing Then
        If mrsDataMap.State = adStateOpen Then mrsDataMap.Close
        Set mrsDataMap = Nothing
    End If
    If Not mfrmParent Is Nothing Then Set mfrmParent = Nothing
    If Not mobjTagFont Is Nothing Then Set mobjTagFont = Nothing
End Sub

Private Function ReDimArray(ByRef strArray() As String) As Long
    '----------------------------------------------------------------------
    '功能：重新定义数组
    '----------------------------------------------------------------------
    Dim lngCount As Long
    Dim strTmp As String
    
    On Error GoTo InitHand
    
    strTmp = strArray(1)
    
    lngCount = UBound(strArray) + 1
    
    GoTo OkHand
    
InitHand:
    
    lngCount = 1
    
OkHand:
    
    ReDim Preserve strArray(1 To lngCount)
            
    ReDimArray = lngCount
End Function

Private Sub SingerShowType(ByVal vsfObj As VSFlexGrid, ByVal lngStartRow As Long, ByVal lngEndRow As Long)
'-------------------------------------------------
'功能：护士签名人显示方式
''--58414,刘鹏飞,2013-04-10,添加护士、签名列显示模式
'-------------------------------------------------
    Dim lngRow As Integer
    
    Select Case mlngSingerType
        Case 0 '所有行显示
            For lngRow = lngStartRow To lngEndRow
                If mlngOperator > 0 Then vsfObj.TextMatrix(lngRow, mlngOperator) = vsfObj.TextMatrix(lngStartRow, mlngOperator)
                If mlngSignName > 0 Then vsfObj.TextMatrix(lngRow, mlngSignName) = vsfObj.TextMatrix(lngStartRow, mlngSignName)
                If mlngSignTime > 0 Then vsfObj.TextMatrix(lngRow, mlngSignTime) = vsfObj.TextMatrix(lngStartRow, mlngSignTime)
            Next
        Case 1 '首行显示
            For lngRow = lngStartRow To lngEndRow
                If lngRow = lngStartRow Then
                    If mlngOperator > 0 Then vsfObj.TextMatrix(lngRow, mlngOperator) = vsfObj.TextMatrix(lngStartRow, mlngOperator)
                    If mlngSignName > 0 Then vsfObj.TextMatrix(lngRow, mlngSignName) = vsfObj.TextMatrix(lngStartRow, mlngSignName)
                    If mlngSignTime > 0 Then vsfObj.TextMatrix(lngRow, mlngSignTime) = vsfObj.TextMatrix(lngStartRow, mlngSignTime)
                Else
                    If mlngOperator > 0 Then vsfObj.TextMatrix(lngRow, mlngOperator) = ""
                    If mlngSignName > 0 Then vsfObj.TextMatrix(lngRow, mlngSignName) = ""
                    If mlngSignTime > 0 Then vsfObj.TextMatrix(lngRow, mlngSignTime) = ""
                End If
            Next
        Case 3 '尾行显示
            If mlngOperator > 0 Then
                If vsfObj.TextMatrix(lngStartRow, mlngOperator) = "" Then vsfObj.TextMatrix(lngStartRow, mlngOperator) = vsfObj.TextMatrix(lngEndRow, mlngOperator)
            End If
            If mlngSignName > 0 Then
                If vsfObj.TextMatrix(lngStartRow, mlngSignName) = "" Then vsfObj.TextMatrix(lngStartRow, mlngSignName) = vsfObj.TextMatrix(lngEndRow, mlngSignName)
            End If
            If mlngSignTime > 0 Then
                If vsfObj.TextMatrix(lngStartRow, mlngSignTime) = "" Then vsfObj.TextMatrix(lngStartRow, mlngSignTime) = vsfObj.TextMatrix(lngEndRow, mlngSignTime)
            End If
            For lngRow = lngEndRow To lngStartRow Step -1
                If lngRow = lngEndRow Then
                    If mlngOperator > 0 Then vsfObj.TextMatrix(lngRow, mlngOperator) = vsfObj.TextMatrix(lngStartRow, mlngOperator)
                    If mlngSignName > 0 Then vsfObj.TextMatrix(lngRow, mlngSignName) = vsfObj.TextMatrix(lngStartRow, mlngSignName)
                    If mlngSignTime > 0 Then vsfObj.TextMatrix(lngRow, mlngSignTime) = vsfObj.TextMatrix(lngStartRow, mlngSignTime)
                Else
                    If mlngOperator > 0 Then vsfObj.TextMatrix(lngRow, mlngOperator) = ""
                    If mlngSignName > 0 Then vsfObj.TextMatrix(lngRow, mlngSignName) = ""
                    If mlngSignTime > 0 Then vsfObj.TextMatrix(lngRow, mlngSignTime) = ""
                End If
            Next
        Case Else '首尾显示
            '最后一行需要填写封闭签名
            For lngRow = lngStartRow To lngEndRow
                If lngRow = lngStartRow Or lngRow = lngEndRow Then
                    If mlngOperator > 0 Then vsfObj.TextMatrix(lngRow, mlngOperator) = vsfObj.TextMatrix(lngStartRow, mlngOperator)
                    If mlngSignName > 0 Then vsfObj.TextMatrix(lngRow, mlngSignName) = vsfObj.TextMatrix(lngStartRow, mlngSignName)
                    If mlngSignTime > 0 Then vsfObj.TextMatrix(lngRow, mlngSignTime) = vsfObj.TextMatrix(lngStartRow, mlngSignTime)
                Else
                    If mlngOperator > 0 Then vsfObj.TextMatrix(lngRow, mlngOperator) = ""
                    If mlngSignName > 0 Then vsfObj.TextMatrix(lngRow, mlngSignName) = ""
                    If mlngSignTime > 0 Then vsfObj.TextMatrix(lngRow, mlngSignTime) = ""
                End If
            Next
    End Select
End Sub

