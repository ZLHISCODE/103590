VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.UserControl usrTendFileReader 
   AutoRedraw      =   -1  'True
   ClientHeight    =   5550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8565
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
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
      Begin VB.OptionButton optPageAlign 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1380
         Picture         =   "usrTendFileReader.ctx":0734
         Style           =   1  'Graphical
         TabIndex        =   11
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
         TabIndex        =   10
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
         Enabled         =   -1  'True
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
         Enabled         =   -1  'True
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
Private mintNORule As Integer             '0-按文件格式编号;1-统一编号

Private mlng当前页码 As Long
Private mlng开始页码 As Long
Private mint当前起始页 As Integer           '当前文件的起始页(考虑已打印部分,以及预览从已打印页开始预览)
Private mint结束页 As Integer
Private mint页码 As Integer
Private mlng当前文件ID As Long
Private mlng合并文件ID As Long
Private mlng打印页 As Long
Private mlng格式ID As Long
Private mlng病人id As Long
Private mlng主页id As Long
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
Private mstrCOLNothing As String            '未绑定的列集合+活动项目列(不管活动项目列是否绑定)
Private mstrCOLActive As String             '活动列集合
Private mstrCatercorner As String           '列对角线集合
Private mblnEditAssistant As Boolean        '当前选择的项目是否允许进行词句选择
Private mlngPageRows As Long                '此文件格式一页所显示的数据行
Private mlngOverrunRows As Long             '超出数据行
Private mlngRowCount As Long                '当前记录总行数
Private mlngRowCurrent As Long              '当前记录在本页的实际行数
Private mlngDate As Long                    '日期
Private mlngTime As Long                    '时间
Private mlngOperator As Long                '护士
Private mlngSignLevel As Long               '签名级别
Private mlngSigner As Long                  '签名信息
Private mlngSignName As Long                '签名人
Private mlngSignTime As Long                '签名时间
Private mlngRecord As Long                  '记录ID
Private mlngNoEditor As Long                '禁止编辑列,存在护士列则以护士列为准,不存在护士列则以签名列为准
Private mlngCollectType As Long             '汇总类别
Private mlngCollectText As Long             '汇总文本
Private mlngCollectStyle As Long            '汇总标记
Private mlngCollectDay As Long              '汇总日期:0-昨天;1-今天
Private mlngPrintedPage As Long             '打印页号
Private mlngPrintedRow As Long              '打印行号

Private mblnSign As Boolean                 '是否签名
Private mblnArchive As Boolean              '是否归档
Private mintType As Integer                 '记录当前的编辑模式
Private mblnDateAd As Boolean               '日期缩写?
Private mstr开始时间 As String              '当前文件的开始时间
Private mstr结束时间 As String              '当前文件的结束时间
Private CellRect As RECT

Private rsTemp As New ADODB.Recordset
Private mrsItems As New ADODB.Recordset             '所有护理记录项目清单
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

Private Const madLongVarCharDefault As Integer = 10          '字符型字段缺省长度
Private Const madDoubleDefault As Integer = 18               '数字型字段缺省长度
Private Const madDbDateDefault As Integer = 20               '日期型字段缺省长度

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
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
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
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long

Private Type POINTAPI
        x As Long
        Y As Long
End Type

Private Const WHITE_BRUSH = 0    '白色画笔
Private Const cdblWidth As Double = 6          '一个英文字符的宽度
Private Const cHideCols = 2         '前缀隐藏列:备用,时间
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
    On Error GoTo errHand
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
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub DrawCollectCell(ByVal hDC As Long, ByVal ROW As Long, ByVal COL As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
    Dim lngPen As Long, lngOldPen As Long
    Dim lpPoint As POINTAPI
    
    '创建新画笔
    lngPen = CreatePen(0, 1, vbRed)
    lngOldPen = SelectObject(hDC, lngPen)
    
    If Val(VsfData.TextMatrix(ROW, mlngCollectStyle)) = 1 Then  '上下划红线
        '画线
        Call MoveToEx(hDC, Left, Top, lpPoint)
        Call LineTo(hDC, Right, Top)
        Call MoveToEx(hDC, Left, Bottom - 2, lpPoint)
        Call LineTo(hDC, Right, Bottom - 2)
    Else                                                        '汇总项下双红线
        If InStr(1, "|" & mstrColCollect & ";", "|" & COL - (cHideCols + VsfData.FixedCols - 1) & ";") <> 0 Then 'And Val(VsfData.TextMatrix(ROW, COL)) <> 0 Then
            '画线
            Call MoveToEx(hDC, Left, Bottom - 4, lpPoint)
            Call LineTo(hDC, Right, Bottom - 4)
            Call MoveToEx(hDC, Left, Bottom - 2, lpPoint)
            Call LineTo(hDC, Right, Bottom - 2)
        End If
    End If
    
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
    Dim lngRow As Long, lngROWS As Long

    GetData = ""
    lngROWS = SendMessage(txtLength.Hwnd, EM_GETLINECOUNT, 0&, 0&)
    For lngRow = 1 To lngROWS
        Call ClearArray(strLine)
        Call SendMessage(txtLength.Hwnd, EM_GETLINE, lngRow - 1, strLine(0))
        strData = StrConv(strLine, vbUnicode)
        strData = TruncZero(strData)
        GetData = GetData & IIf(GetData = "", "", "|ZYB.ZLSOFT|") & strData & IIf(lngRow < lngROWS, vbCrLf, "")
    Next
    GetData = Split(GetData, "|ZYB.ZLSOFT|")
End Function

Private Sub ClearArray(strLine() As Byte)
    Dim intDo As Integer, intMax As Integer
    intMax = UBound(strLine)
    For intDo = 0 To intMax
        strLine(intDo) = 0
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
    On Error GoTo errHand
    
    '获取指定页码的数据发生时间范围
    gstrSQL = " Select /*+ RULE */ MIN(发生时间) 开始时间,MAX(发生时间) AS 结束时间 From 病人护理打印 Where 文件ID=[1] And (开始页号=[2] OR 结束页号=[2])"
    Set rsTemp = OpenSQLRecord(gstrSQL, "获取指定页码的数据发生时间范围", mlng当前文件ID, mint页码)
    If NVL(rsTemp!开始时间) = "" Then
        If mint婴儿 = 0 Then
            gstrSQL = " Select  /*+ RULE */ 入院日期 AS 开始时间,sysdate AS 结束时间 From 病案主页 Where 病人ID=[1] And 主页ID=[2]"
        Else
            gstrSQL = " Select  /*+ RULE */ 出生时间 AS 开始时间,sysdate AS 结束时间 From 病人新生儿记录 Where 病人ID=[1] And 主页ID=[2] And 序号=[3]"
        End If
        Set rsTemp = OpenSQLRecord(gstrSQL, "取入院日期或出生日期", mlng病人id, mlng主页id, mint婴儿)
    End If
    GetPeriod = Format(rsTemp!开始时间, "yyyy-MM-dd HH:mm:ss") & "～" & Format(rsTemp!结束时间, "yyyy-MM-dd HH:mm:ss")
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub ReadStruDef()
    Dim lngCOL As Long
    On Error GoTo errHand
    
    '读取文件属性
    mblnDateAd = False
    Call GetFileProperty
    
    '提取活动项目并加入列定义(格式：列号;表头名称|项目序号,部位;项目序号,部位||列号;表头名称...)
    mstrCOLActive = ""
    mstrCOLNothing = ""
    mstrCollectItems = ""
    mstrColCollect = ""
    gstrSQL = " Select  /*+ RULE */ A.列号,A.列头名称,A.序号,A.项目序号,A.部位 From 病人护理页面_活动项目 A " & _
              " Where A.文件ID=[1] And A.页号=[2] " & _
              " Order by A.列号,A.序号"
    Set rsTemp = OpenSQLRecord(gstrSQL, "提取出所有自定义的活动项目", mlng当前文件ID, mint页码)
    If rsTemp.RecordCount <> 0 Then
        Do While Not rsTemp.EOF
            If lngCOL <> rsTemp!列号 Then
                lngCOL = rsTemp!列号
                mstrCOLActive = mstrCOLActive & "||" & rsTemp!列号 & ";" & rsTemp!列头名称 & "|" & rsTemp!项目序号 & "," & NVL(rsTemp!部位)
            Else
                mstrCOLActive = mstrCOLActive & ";" & rsTemp!项目序号 & "," & NVL(rsTemp!部位)
            End If
            rsTemp.MoveNext
        Loop
    End If
    If mstrCOLActive <> "" Then mstrCOLActive = Mid(mstrCOLActive, 3)
    
    '读取病历文件格式定义
    gstrSQL = "Select  /*+ RULE */ d.对象序号, d.内容文本, d.要素名称" & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表格样式'" & _
        " Order By d.对象序号"
    Set rsTemp = OpenSQLRecord(gstrSQL, "读取病历文件格式定义", mlng格式ID)
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
                mlngPageRows = Val(!内容文本)
            End Select
            .MoveNext
        Loop
    End With
    
    gstrSQL = "Select  /*+ RULE */ 格式, 页脚, 种类||'-'||编号 AS KEY From 病历页面格式 Where 种类 = 3 And 编号 In (Select 页面 From 病历文件列表 Where Id = [1])"
    Set rsTemp = OpenSQLRecord(gstrSQL, "读取病历页面格式", mlng格式ID)
    If Not rsTemp.EOF Then
        mstrPaperSet = "" & rsTemp!格式
        If picHead.Tag = "" Then
            '考虑到医院内护理文件页眉页脚格式统一，此处只读取一次
            Call ReadPageHead(rtbHead, rsTemp!Key)
            Call ReadPageFoot(rtbFoot, rsTemp!Key)
            picHead.Tag = rsTemp!Key
            chk页码.Value = IIf(Val(NVL(rsTemp!页脚, 0)) > 0, 1, 0)
            If chk页码.Value = 1 Then optPageAlign(Val(NVL(rsTemp!页脚, 0)) - 1).Value = True
        End If
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select  /*+ RULE */ d.对象序号, d.内容文本, d.要素名称, Nvl(d.是否换行, 0) As 是否换行" & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表上标签'" & _
        " Order By d.对象序号"
    Set rsTemp = OpenSQLRecord(gstrSQL, "读取表上标签定义", mlng格式ID)
    With rsTemp
        mstrSubhead = ""
        Do While Not .EOF
            mstrSubhead = mstrSubhead & "|" & IIf(!是否换行 = 0, "", vbCrLf) & !内容文本 & "{" & !要素名称 & "}"
            .MoveNext
        Loop
        If mstrSubhead <> "" Then mstrSubhead = Mid(mstrSubhead, 2)
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select  /*+ RULE */ d.对象序号, d.内容行次, d.内容文本" & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表头单元'" & _
        " Order By d.对象序号"
    Set rsTemp = OpenSQLRecord(gstrSQL, "读取表头单元定义", mlng格式ID)
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
    Dim bln对角线 As Boolean, bln选择项 As Boolean          '如果上一列是对角线且选择项,则直接提取各项数据,拼列头时在数值间加上/
    Dim lngColumn As Long, blnAddCollect As Boolean
    
    gstrSQL = "Select  /*+ RULE */ d.对象序号, d.对象属性, d.内容行次, d.内容文本, d.要素名称, d.要素单位,d.要素表示 " & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表列集合'" & _
        " Order By d.对象序号, d.内容行次"
    Set rsTemp = OpenSQLRecord(gstrSQL, "读取表列集合定义", mlng格式ID)
    With rsTemp
        lngColumn = 0: mstrColumns = "": mstrColWidth = "": mstrCatercorner = ""
        mstrSQL内 = "": mstrSQL中 = "": strSql外 = "": mstrSQL列 = "": mstrSQL条件 = ""
        bln日期 = False: bln时间 = False: bln护士 = False
        bln签名人 = False: bln签名时间 = False: bln签名日期 = False
        Do While Not .EOF
            If lngColumn <> !对象序号 Then
                blnAddCollect = False
                mstrColumns = mstrColumns & IIf(mstrColumns = "", "", "'1'" & str格式) & "|" & !对象序号 & "'" & !要素名称
                mstrColWidth = mstrColWidth & "," & !对象属性 & "`" & !对象序号 & "`" & !要素表示
                If !要素表示 = 1 Then mstrCatercorner = mstrCatercorner & "," & !对象序号
                str格式 = ""
                If !要素名称 <> "" Then
                    str格式 = "{" & NVL(!内容文本) & "[" & !要素名称 & "]" & NVL(!要素单位) & "}"
                    mstrSQL列 = mstrSQL列 & "," & IIf(Mid(strSql外, 3) = "", "''", Mid(strSql外, 3)) & " As C" & Format(lngColumn, "00")
                Else
                    If strSql外 <> "" Then
                        mstrSQL列 = mstrSQL列 & "," & Mid(strSql外, 3) & " As C" & Format(lngColumn, "00")
                    Else
                        mstrSQL列 = mstrSQL列 & ",'' As C" & Format(lngColumn, "00")
                    End If
                End If
                strSql外 = ""
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
                            mstrColCollect = mstrColCollect & "," & mrsItems!项目序号
                        Else    '有可能一列绑定两个项目,第一个项目不是汇总项目,第二个项目才是汇总项目,因此,下面的代码保证加上列序号
                            mstrColCollect = mstrColCollect & "|" & !对象序号 & ";" & mrsItems!项目序号
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
                mstrSQL内 = mstrSQL内 & ",l.签名人"
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
                    
                    If bln对角线 And bln选择项 Then
                        If strSql外 <> "" Then
                            '第二项
                            strSql外 = strSql外 & "||'/'||""" & !要素名称 & """"
                        Else
                            '第一项
                            strSql外 = strSql外 & "||""" & !要素名称 & """"
                        End If
                    Else
                        strSql外 = strSql外 & "||""" & !要素名称 & """"
                    End If
                    
                    If (Trim("" & !内容文本) = "" And Trim("" & !要素单位) = "") Or (bln对角线 And bln选择项) Then
                        mstrSQL内 = mstrSQL内 & ", Decode(c.项目名称, '" & !要素名称 & "', Nvl(c.未记说明,c.记录内容), '') As """ & !要素名称 & """"
                    Else
                        mstrSQL内 = mstrSQL内 & ", Decode(c.项目名称, '" & !要素名称 & "', Nvl(c.未记说明,Decode(c.记录内容,Null,'" & !内容文本 & "'||'" & !要素单位 & "','" & !内容文本 & "'||c.记录内容||'" & !要素单位 & "')), '') As """ & !要素名称 & """"
                    End If
                Else
                    '为空表示未绑定列,强制加,后面进行替换
                    mstrCOLNothing = mstrCOLNothing & "," & Format(!对象序号, "00")
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
        mstrCOLNothing = Mid(mstrCOLNothing, 2)
        mstrCatercorner = Mid(mstrCatercorner, 2)
        mstrColWidth = Mid(mstrColWidth, 2)
        '加入最后一列的格式
        mstrColumns = mstrColumns & IIf(mstrColumns = "", "", "'1'" & str格式) '& "|" & !对象序号 & "'" & !要素名称
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
        If bln签名时间 = False Then mstrSQL内 = mstrSQL内 & ",l.签名时间"
        
        If Mid(mstrSQL中, 2) = "" Then
            MsgBox "对不起，您没有定义当前护理单的显示列信息，请在病历文件管理中定义！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '程序内部控制增加固定列
        mstrSQL中 = UCase(mstrSQL中 & ",MAX(签名级别) AS 签名级别,MAX(签名信息) AS 签名信息,MAX(记录ID) AS 记录ID,MAX(行数) AS 行数,MAX(实际行数) AS 实际行数,MAX(汇总类别) AS 汇总类别,MAX(汇总文本) AS 汇总文本,MAX(汇总标记) AS 汇总标记,MAX(汇总日期) AS 汇总日期,MAX(打印页号) AS 打印页号,MAX(打印行号) AS 打印行号")
        mstrSQL内 = UCase(mstrSQL内 & ",l.签名级别,l.签名人 AS 签名信息,C.记录ID,P.行数||'' AS 行数,DECODE(SIGN(P.结束页号-P.开始页号),1,DECODE(SIGN([5]-P.开始页号),1, P.结束行号,P.行数-P.结束行号 ),P.行数) AS 实际行数,NVL(L.汇总类别,0) AS 汇总类别,L.汇总文本,L.汇总标记,to_char(L.发生时间,'yyyy-MM-dd hh24:mi:ss')||'' AS 汇总日期,p.打印页号,p.打印行号")
        mstrSQL列 = UCase(mstrSQL列 & ",签名级别,签名信息,记录ID,行数,实际行数,汇总类别,汇总文本,汇总标记,汇总日期,打印页号,打印行号")
        
        '将活动项目加入到SQL中
        Call PreActiveCOL
        'Call SQLCombination
    End With
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

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
                    strCOLDEF = strCOLDEF & " """ & strCOLPart & mrsItems!项目名称 & """ AS C" & intCol
                Else
                    strCOLDEF = strCOLDEF & " """ & strCOLPart & mrsItems!项目名称 & """||"
                End If
            Else
                strCOLDEF = strCOLDEF & "NVL(" & strCOLPart & mrsItems!项目名称 & ",'/') AS C" & intCol
            End If
            
            strColFormat = strColFormat & "{[" & strCOLPart & mrsItems!项目名称 & "]" & IIf(intMax > 0 And intIn = 0, "/", "") & "}"
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
        mstrColumns = Replace(mstrColumns, intCol & "''1'", intCol & "'" & strCOLNames & "'1'" & strColFormat)
        '列
        mstrSQL列 = Replace(mstrSQL列, "'' AS C" & Format(intCol, "00"), strCOLDEF)
        '条件
        mstrSQL条件 = Replace(UCase(mstrSQL条件), " OR """ & "C" & Format(intCol, "00") & """ IS NOT NULL", strCOLCOND)
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
        mstrSQL条件 = Replace(UCase(mstrSQL条件), " OR """ & "C" & Format(arrData(intDo), "00") & """ IS NOT NULL", "")
        '中
        mstrSQL中 = Replace(mstrSQL中, ",MAX(""" & "C" & Format(arrData(intDo), "00") & """) AS C" & Format(arrData(intDo), "00"), "")
        '内
        mstrSQL内 = Replace(mstrSQL内, ", C" & Format(arrData(intDo), "00") & " AS C" & Format(arrData(intDo), "00"), "")
    Next
End Sub

Private Sub SQLCombination(ByVal str条件 As String)
    mstrSQL = "Select  /*+ RULE */ 备用,发生时间," & Mid(mstrSQL列, 12) & vbCrLf & _
                " From (Select 记录组号,时间 as 备用,发生时间," & Mid(mstrSQL中, 2) & vbCrLf & _
                "        From (Select c.记录组号,to_char(l.发生时间,'yyyy-MM-dd hh24:mi:ss') AS 发生时间," & Mid(mstrSQL内, 2) & vbCrLf & _
                "               From 病人护理数据 l, 病人护理明细 c,病人护理文件 f,病人护理打印 p " & vbCrLf & _
                "               Where l.ID=p.记录ID And l.Id = c.记录id And l.文件ID=f.ID And f.ID=p.文件ID " & _
                "               And c.终止版本 Is Null And c.记录类型<>5  " & _
                "               And f.id=[1] And f.病人id = [2] And f.主页id = [3] And Nvl(f.婴儿,0)=[4] " & str条件 & ")" & vbCrLf & _
                IIf(mstrSQL条件 <> "", "Where " & mstrSQL条件, "") & _
                "       Group By 日期, 时间, 发生时间,记录组号,护士,签名人,签名时间" & _
                                "       Order By 发生时间,记录组号,护士,签名人,签名时间)"
End Sub

Private Sub zlReadTip(aryPeriod)
    Dim aryRow() As String, aryItem() As String
    Dim strPrefix As String, strItemName As String
    Dim lngRow As Long, lngCOL As Long, lngCount As Long, strCell As String
    Dim strTmpSQL As String
    Dim strTmp As String
    
    Err = 0: On Error GoTo errHand
    
    '表上标签获取
    lblSubhead.Caption = ""
    lblSubhead.Tag = ""
    gstrSQL = "Select [1] || Zl_Replace_Element_Value([2],[3],[4],2,NULL,[5]) as 信息 From Dual"
    aryItem = Split(mstrSubhead, "|")
        
    For lngCount = 0 To UBound(aryItem)
        strPrefix = Left(aryItem(lngCount), InStr(1, aryItem(lngCount), "{") - 1)
        strItemName = Mid(aryItem(lngCount), InStr(1, aryItem(lngCount), "{") + 1, InStr(1, aryItem(lngCount), "}") - InStr(1, aryItem(lngCount), "{") - 1)
        
        strTmp = strPrefix
        Select Case strItemName
        Case "当前病区"
        
            strTmpSQL = "Select  /*+ RULE */ b.名称" & vbNewLine & _
                        "From (Select 病区id, 开始时间, Nvl(终止时间, To_Date('3000-01-01', 'yyyy-mm-dd')) As 终止时间" & vbNewLine & _
                        "            From 病人变动记录" & vbNewLine & _
                        "            Where 病人id = [1] And 主页id = [2] And 科室id = [3]) a,部门表 b " & vbNewLine & _
                        "Where ([4] Between a.开始时间 And a.终止时间 Or [4] >= a.开始时间) And a.病区id Is Not Null And b.ID=a.病区id" & vbNewLine & _
                        "Order By a.开始时间"
                        
            Set rsTemp = OpenSQLRecord(strTmpSQL, "当前病区", mlng病人id, mlng主页id, mlng科室ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            
        Case "当前床号"
        
            strTmpSQL = "Select  /*+ RULE */ a.床号" & vbNewLine & _
                        "From (Select 床号, 开始时间, Nvl(终止时间, To_Date('3000-01-01', 'yyyy-mm-dd')) As 终止时间" & vbNewLine & _
                        "            From 病人变动记录" & vbNewLine & _
                        "            Where 病人id = [1] And 主页id = [2] And 科室id = [3]) a" & vbNewLine & _
                        "Where ([4] Between a.开始时间 And a.终止时间 Or [4] >= a.开始时间) And a.床号 Is Not Null" & vbNewLine & _
                        "Order By a.开始时间"

            Set rsTemp = OpenSQLRecord(strTmpSQL, "当前床号", mlng病人id, mlng主页id, mlng科室ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            If rsTemp.BOF = False Then rsTemp.MoveLast
            
        Case "当前科室"
        
            strTmpSQL = "Select  /*+ RULE */ 名称 From 部门表 a Where a.ID=[1]"
            Set rsTemp = OpenSQLRecord(strTmpSQL, "当前科室", mlng科室ID)
            
        Case "住院医师"
            strTmpSQL = "Select  /*+ RULE */ a.经治医师" & vbNewLine & _
                        "From (Select 经治医师, 开始时间, Nvl(终止时间, To_Date('3000-01-01', 'yyyy-mm-dd')) As 终止时间" & vbNewLine & _
                        "            From 病人变动记录" & vbNewLine & _
                        "            Where 病人id = [1] And 主页id = [2] And 科室id = [3]) a" & vbNewLine & _
                        "Where ([4] Between a.开始时间 And a.终止时间 Or [4] >= a.开始时间) And a.经治医师 Is Not Null" & vbNewLine & _
                        "Order By a.开始时间"
            Set rsTemp = OpenSQLRecord(strTmpSQL, "住院医师", mlng病人id, mlng主页id, mlng科室ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            If rsTemp.BOF = False Then rsTemp.MoveLast
        Case "责任护士"
        
            strTmpSQL = "Select  /*+ RULE */ a.责任护士" & vbNewLine & _
                        "From (Select 责任护士, 开始时间, Nvl(终止时间, To_Date('3000-01-01', 'yyyy-mm-dd')) As 终止时间" & vbNewLine & _
                        "            From 病人变动记录" & vbNewLine & _
                        "            Where 病人id = [1] And 主页id = [2] And 科室id = [3]) a" & vbNewLine & _
                        "Where ([4] Between a.开始时间 And a.终止时间 Or [4] >= a.开始时间) And a.责任护士 Is Not Null" & vbNewLine & _
                        "Order By a.开始时间"
            Set rsTemp = OpenSQLRecord(strTmpSQL, "责任护士", mlng病人id, mlng主页id, mlng科室ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            If rsTemp.BOF = False Then rsTemp.MoveLast
            
        Case "护理等级"
            strTmpSQL = "Select  /*+ RULE */ b.名称" & vbNewLine & _
                        "From (Select 护理等级ID, 开始时间, Nvl(终止时间, To_Date('3000-01-01', 'yyyy-mm-dd')) As 终止时间" & vbNewLine & _
                        "            From 病人变动记录" & vbNewLine & _
                        "            Where 病人id = [1] And 主页id = [2] And 科室id = [3]) a,护理等级 b" & vbNewLine & _
                        "Where ([4] Between a.开始时间 And a.终止时间 Or [4] >= a.开始时间) And a.护理等级ID Is Not Null And b.序号=a.护理等级ID" & vbNewLine & _
                        "Order By a.开始时间"
            Set rsTemp = OpenSQLRecord(strTmpSQL, "护理等级", mlng病人id, mlng主页id, mlng科室ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            If rsTemp.BOF = False Then rsTemp.MoveLast
            
        Case Else
            strTmp = ""
            Set rsTemp = OpenSQLRecord(gstrSQL, "取要素", strPrefix, strItemName, mlng病人id, mlng主页id, mint婴儿)
        End Select
        
        If rsTemp.BOF = False Then
            If strTmp <> "" Then
                lblSubhead.Tag = lblSubhead.Tag & " " & strTmp & rsTemp.Fields(0).Value
            Else
                lblSubhead.Tag = lblSubhead.Tag & " " & rsTemp.Fields(0).Value
            End If
        End If
    Next
    lblSubhead.Tag = Trim(lblSubhead.Tag)
    
    '表上标签分散处理
    Call zlLableBruit
    Exit Sub

errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub zlRefresh(ByVal str条件 As String)
    Dim aryRow() As String, aryItem() As String
    Dim strPrefix As String, strItemName As String
    Dim lngRow As Long, lngCOL As Long, lngCount As Long, strCell As String
    Dim strTmpSQL As String
    Dim strTmp As String
    
    Err = 0: On Error GoTo errHand
    
    '装入数据
    Call SQLCombination(str条件)
    gstrSQL = mstrSQL
    If gblnMoved_HL Then
        gstrSQL = Replace(gstrSQL, "病人护理数据", "H病人护理数据")
        gstrSQL = Replace(gstrSQL, "病人护理明细", "H病人护理明细")
    End If
    Set rsTemp = OpenSQLRecord(gstrSQL, "提取护理数据", mlng当前文件ID, mlng病人id, mlng主页id, mint婴儿, mint页码, mlngPageRows)
    '数据记录集,用于快速恢复
    Set mrsDataMap = CopyNewRec(rsTemp, mrsDataMap)
    
    Exit Sub

errHand:
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
    Dim lngRowCount As Long, lngRowCurrent As Long  '当前记录总行数,当前记录在本页的实际行数
    Dim lngCOL As Long, lngMax As Long
    Dim lngRow As Long, lngStart As Long, lngPrintedRow As Long
    Dim blnDelete As Boolean
    On Error GoTo errHand
    
    Dim arrData
    Dim intData As Integer, intDatas As Integer
    '如果一行显示不完则分行显示(根据当前数据占用行数先添加空白行并处理行坐标,然后再依次处理当前行的数据)
    '每页只显示实际的数据行,把'@处取消注释即可
    '重新调整所有数据的实际行
    
    lngRow = VsfData.FixedRows
    Do While True
        If lngRow > VsfData.Rows - 1 Then Exit Do
        If InStr(1, VsfData.TextMatrix(lngRow, mlngRowCount), "|") <> 0 Then Exit Do
        lngRowCount = Val(VsfData.TextMatrix(lngRow, mlngRowCount))
'        @实际数据行
        lngPrintedRow = Val(VsfData.TextMatrix(lngRow, mlngPrintedRow))
        If lngPrintedRow = 0 Then
            lngRowCurrent = VsfData.TextMatrix(lngRow, mlngRowCurrent)
        Else
            If mlngPageRows < (lngRowCount + lngPrintedRow - 1) Then
                '始终当作首行，计算跨页数据的跨页行有多少
                VsfData.TextMatrix(lngRow, mlngRowCurrent) = (lngRowCount + lngPrintedRow - mlngPageRows - 1)
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
            For lngCOL = 0 To VsfData.Cols - 1
                If VsfData.ColHidden(lngCOL) And lngCOL <> mlngRowCount Then
                    '循环赋值
                    For intData = 2 To lngRowCount
                        VsfData.TextMatrix(lngRow + intData - 1, lngCOL) = VsfData.TextMatrix(lngRow, lngCOL)
                    Next
                ElseIf (lngCOL < mlngNoEditor And lngCOL <> mlngDate And lngCOL <> mlngTime) Then
                    '准备赋值
                    With txtLength
                        .Width = VsfData.ColWidth(lngCOL)
                        .Text = VsfData.TextMatrix(lngRow, lngCOL)
                        .FontName = VsfData.CellFontName
                        .FontSize = VsfData.CellFontSize
                    End With
                    arrData = GetData(txtLength.Text)
                    intDatas = UBound(arrData)
                    
                    If intDatas > 0 Then
                        '循环赋值
                        For intData = 0 To intDatas
                            If VsfData.Rows <= lngRow + intData Then VsfData.Rows = VsfData.Rows + 1
                            VsfData.TextMatrix(lngRow + intData, lngCOL) = arrData(intData)
                        Next
                    End If
                ElseIf lngCOL = mlngNoEditor Then
                        '将行值改为从1开始,比如有4行数据,就是4|1
                        For intData = 1 To lngRowCount
                            VsfData.TextMatrix(lngRow + intData - 1, mlngRowCount) = lngRowCount & "|" & intData
                        Next
                    Else
                End If
            Next
            lngRow = lngRow + lngRowCount - 1
        Else
            VsfData.TextMatrix(lngRow, mlngRowCount) = "1|1"
        End If
        lngRow = lngRow + 1
    Loop
    
    '填充每页的启动行
    lngRow = VsfData.FixedRows
    Do While True
        '固定复制显示日期时间与签名列
        lngStart = GetStartRow(lngRow)
        
        '特殊处理第一行(第一行可能存在跨页数据)
        If lngRow = VsfData.FixedRows And Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0)) <> Val(VsfData.TextMatrix(lngRow, mlngRowCurrent)) Then
            blnDelete = True
            lngRow = lngRow + Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0)) - Val(VsfData.TextMatrix(lngRow, mlngRowCurrent))
        End If
        
        If lngStart <> lngRow Then
            If mlngDate > -1 Then VsfData.TextMatrix(lngRow, mlngDate) = VsfData.TextMatrix(lngStart, mlngDate)
            If mlngTime > -1 Then VsfData.TextMatrix(lngRow, mlngTime) = VsfData.TextMatrix(lngStart, mlngTime)
            If mlngOperator <> -1 Then VsfData.TextMatrix(lngRow, mlngOperator) = VsfData.TextMatrix(lngStart, mlngOperator)
            If mlngSignName <> -1 Then VsfData.TextMatrix(lngRow, mlngSignName) = VsfData.TextMatrix(lngStart, mlngSignName)
            If mlngSignTime <> -1 Then VsfData.TextMatrix(lngRow, mlngSignTime) = VsfData.TextMatrix(lngStart, mlngSignTime)
        End If
        
        If blnDelete Then
            For lngCOL = lngStart To lngRow - 1
                VsfData.RemoveItem lngStart
            Next
            blnDelete = False
            lngRow = VsfData.FixedRows  '只处理第一行记录删除的情况,所以固定设置为固定行为启始行
        End If
        
        lngRow = lngRow + mlngPageRows
        If lngRow > VsfData.Rows - 1 Then Exit Do
    Loop
    
    '如果是重打,将超出页有效数据行的部分删掉
    If gintPrintState = 2 Then
        If VsfData.Rows > VsfData.FixedRows + mlngPageRows Then
            VsfData.Rows = VsfData.FixedRows + mlngPageRows
        End If
    End If
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub PreTendFormat()
    Dim aryItem() As String
    Dim lngRow As Long, lngCOL As Long, lngCount As Long, strCell As String
    On Error GoTo errHand
    
    '设置护理记录单的格式
    With VsfData
        .Redraw = flexRDNone
        .Clear
        Set .DataSource = mrsDataMap
        
        '表头填写
        .MergeCells = flexMergeFixedOnly
        .MergeCellsFixed = flexMergeFree
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeRow(2) = True
        
        '程序内部控制列隐藏
        .ColHidden(0) = True
        .ColHidden(1) = True
        .ColHidden(mlngRowCount) = True
        .ColHidden(mlngRowCurrent) = True
        .ColHidden(mlngRecord) = True
        .ColHidden(mlngSigner) = True
        .ColHidden(mlngSignLevel) = True
        .ColHidden(mlngCollectStyle) = True
        .ColHidden(mlngCollectText) = True
        .ColHidden(mlngCollectType) = True
        .ColHidden(mlngCollectDay) = True
        .ColHidden(mlngPrintedPage) = True
        .ColHidden(mlngPrintedRow) = True
        
        '设置列头
        aryItem = Split(mstrTabHead, "|")
        For lngCount = 0 To UBound(aryItem)
            strCell = aryItem(lngCount)
            lngRow = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            lngCOL = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            .TextMatrix(lngRow, lngCOL + cHideCols + .FixedCols - 1) = strCell
        Next
        Call PreActiveHead
        
        '列宽设置
        Dim blnAlign As Boolean
        aryItem = Split(mstrColWidth, ",")
        For lngCount = cHideCols + .FixedCols To .Cols - 1
            If Not .ColHidden(lngCount) Then
                .ColWidth(lngCount) = Val(Split(aryItem(lngCount - cHideCols - .FixedCols), "`")(0))
                If InStr(1, aryItem(lngCount - cHideCols - .FixedCols), "`") <> 0 Then
                    blnAlign = True
                    .ColAlignment(lngCount) = Val(Split(aryItem(lngCount - cHideCols - .FixedCols), "`")(1))
                End If
            End If
        Next
        
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
        
        If gintPrintState <> 2 Then
            mint结束页 = (VsfData.Rows - VsfData.FixedRows) \ mlngPageRows
            'If (VsfData.Rows - VsfData.FixedRows) Mod mlngPageRows <> 0 Then mint结束页 = mint结束页 + 1
            mint结束页 = mint结束页 + mint当前起始页
            If gintPrintState = 1 Then mint页码 = mint结束页
        End If
        
        Call WriteColor
        Call ShowPage
        .Redraw = flexRDDirect
    End With
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function PrintHead() As Boolean
    PrintHead = PrintRTBData(rtbHead, True)
End Function

Public Function PrintFoot() As Boolean
    Dim lngPage As Long
    On Error GoTo errHand
    '如果要打印页码则先打印页码,再打印页脚
    
    If mintNORule = 1 Then
        If gintPrintState = 1 Then
            '取当前文件的最后页,如果未打印完,第一页就取最后页的码
            lngPage = mlng开始页码 + mint页码 - mint当前起始页
        Else
            lngPage = mint页码
        End If
    Else
        lngPage = mint页码
    End If
    mlng当前页码 = lngPage
    
    PrintFoot = PrintRTBData(rtbFoot, False, lngPage)
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function PrintRTBData(ByVal objRTB As RichTextBox, ByVal blnHead As Boolean, Optional ByVal lngPage As Long = 0) As Boolean
    Dim fr As FORMATRANGE           '格式化的文本范围
    Dim rcDrawTo As RECT            '目标文字区域
    Dim rcPage As RECT              '目标页面区域
    Dim gTargetDC As Long
    Dim lngFoot As Long
    Dim lngOffsetLeft As Long
    Dim lngOffsetTop As Long
    Dim lngNextPos As Long, lngLen As Long, lngTmp As Long, lngPageCount As Long
    Dim rsTemp As New ADODB.Recordset
    
    lngLen = lstrlen(objRTB.Text)
    lngOffsetLeft = gobjOutTo.ScaleX(GetDeviceCaps(gobjOutTo.hDC, PHYSICALOFFSETX), vbPixels, vbTwips)
    lngOffsetTop = gobjOutTo.ScaleY(GetDeviceCaps(gobjOutTo.hDC, PHYSICALOFFSETY), vbPixels, vbTwips)
    
    If blnHead Then
        gobjOutTo.Print ""
    Else
        If chk页码.Value = 1 Then
            lngFoot = 180
            gobjOutTo.CurrentY = gobjOutTo.Height - gobjOutTo.ScaleX(gobjSend.EmptyDown, vbMillimeters, vbTwips) - 200
            If optPageAlign(0).Value Then
                gobjOutTo.CurrentX = gobjOutTo.ScaleX(gobjSend.EmptyLeft, vbMillimeters, vbTwips) - 30
            ElseIf optPageAlign(1).Value Then
                gobjOutTo.CurrentX = (gobjOutTo.Width - 90 * LenB(StrConv("页码:" & mint页码, vbFromUnicode))) / 2
            Else
                gobjOutTo.CurrentX = gobjOutTo.Width - gobjOutTo.ScaleX(gobjSend.EmptyRight, vbMillimeters, vbTwips) - 90 * LenB(StrConv("页码:" & mint页码, vbFromUnicode))
            End If
            gobjOutTo.Print "页码:" & lngPage
        Else
            gobjOutTo.Print ""
        End If
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
            .Left = lngOffsetLeft
            .Top = lngOffsetTop
            .Right = gobjOutTo.Width - lngOffsetLeft
            .Bottom = gobjOutTo.ScaleX(gobjSend.EmptyUp, vbMillimeters, vbTwips) - 30
        Else
            .Left = lngOffsetLeft
            .Top = gobjOutTo.Height - gobjOutTo.ScaleX(gobjSend.EmptyDown, vbMillimeters, vbTwips) + lngFoot
            .Right = gobjOutTo.Width - lngOffsetLeft
            .Bottom = gobjOutTo.Height
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
        lngNextPos = SendMessage(objRTB.Hwnd, EM_FORMATRANGE, 0, fr)
        
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
    Call SendMessage(objRTB.Hwnd, EM_FORMATRANGE, 0, ByVal CLng(0))
    
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
        Call SendMessage(objRTB.Hwnd, EM_FORMATRANGE, 1, fr)
        Call SendMessage(objRTB.Hwnd, EM_FORMATRANGE, 0, ByVal CLng(0))
    Next
    
End Function

Public Function PrintPage() As Boolean
    Dim strSQL() As String
    Dim blnTrans As Boolean
    Dim blnSave As Boolean          '已打印的数据不保存
    Dim strTime As String
    Dim strCurrDate As String
    Dim lngRow As Long, lngROWS As Long
    Dim intMax As Integer, intPos As Integer
    Dim lngCurRow As Long, lngDataLines As Long
    On Error GoTo errHand
    
    '对显示行进行处理
    lngROWS = VsfData.Rows - 1
    For lngRow = VsfData.FixedRows To lngROWS
        If Not VsfData.RowHidden(lngRow) Then
            If lngCurRow = 0 Then lngCurRow = 1
            If VsfData.TextMatrix(lngRow, mlngRowCount) Like "*|1" Then
                lngDataLines = Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0)
                blnSave = (Val(VsfData.TextMatrix(lngRow, mlngPrintedPage)) = 0) Or gintPrintState > 1
                
                If blnSave Then
                    strTime = VsfData.TextMatrix(lngRow, 1)
                    gstrSQL = "ZL_病人护理打印_PRINT(" & mlng当前文件ID & ",to_date('" & strTime & "','yyyy-MM-dd hh24:mi:ss'),'" & gstrUserName & "'," & IIf(mlng当前页码 = 0, mint页码, mlng当前页码) & "," & lngCurRow & ")"
                    Debug.Print gstrSQL
                    strSQL(ReDimArray(strSQL)) = gstrSQL
                End If
            End If
            lngCurRow = lngCurRow + 1
        End If
    Next
    
    On Error Resume Next
    intMax = UBound(strSQL)

    gcnOracle.BeginTrans
    blnTrans = True

    On Error GoTo errHand
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
errHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub InitRecords()
    Dim i As Integer, j As Integer, k As Integer, l As Integer
    Dim lngCOL As Long, lngOrder As Long, strName As String, intImmovable As Integer, strFormat As String
    Dim arrColumn, arrItem, strColumns As String
    Dim blnSet As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
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
            lngCOL = Split(arrColumn(i), "'")(0)
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
                strValues = lngCOL & "|" & l + 1 & "|" & lngOrder & "|" & strName & "|" & intImmovable & "|" & strFormat
                Call Record_Add(mrsSelItems, strFields, strValues)
            Next
        Next
        
        'Call OutputRsData(mrsSelItems)
        
        '加入程序内部控制列(列是在读取数据后绑定时增加的,此时只有预处理下)
        mlngSignLevel = VsfData.Cols + cHideCols + VsfData.FixedCols '加上隐藏列
        mlngSigner = mlngSignLevel + 1
        mlngRecord = mlngSigner + 1
        mlngRowCount = mlngRecord + 1
        mlngRowCurrent = mlngRowCount + 1
        mlngCollectType = mlngRowCurrent + 1
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
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function ShowPage(Optional ByVal intPage As Integer = 0) As Boolean
    '显示指定页面数据并更新打印对象
    Dim aryPeriod
    Dim strBegin As String, strEnd As String
    Dim lngRow As Long, lngROWS As Long, lngStart As Long
    Dim lngShows As Long
    On Error GoTo errHand
    
    If intPage <> 0 Then mint页码 = intPage
    With VsfData
        '小于页面有效数据行说明只有一页数据
        If VsfData.Rows - VsfData.FixedRows > mlngPageRows Then
            lngROWS = .Rows - 1
            For lngRow = .FixedRows To lngROWS
                .RowHidden(lngRow) = True
            Next
        End If
        
        '小于页面有效数据行说明只有一页数据
        If VsfData.Rows - VsfData.FixedRows > mlngPageRows Then
            lngRow = 3 + mlngPageRows * (mint页码 - mint当前起始页)
            lngROWS = 3 + mlngPageRows * (mint页码 - mint当前起始页 + 1) - 1
        Else
            lngRow = 3
            lngROWS = .Rows - 1
        End If
        If lngROWS > .Rows - 1 Then lngROWS = .Rows - 1
        '获取指定页的时间范围
        If lngRow > lngROWS Then
            Exit Function
        End If
        strBegin = .TextMatrix(lngRow, 1)
        lngStart = lngROWS
        lngStart = GetStartRow(lngStart)
        strEnd = .TextMatrix(lngStart, 1)
        aryPeriod = Split(strBegin & "||" & strEnd, "||")
        
        '小于页面有效数据行说明只有一页数据
        If VsfData.Rows - VsfData.FixedRows > mlngPageRows Then
            '显示数据行
            For lngRow = lngRow To lngROWS
                .RowHidden(lngRow) = False
                lngShows = lngShows + 1
            Next
        End If
        
        ShowPage = True
        Call zlReadTip(aryPeriod)
    End With
    
    '设置打印相关内容
    Dim objPrint As New zlPrintTends, objAppRow As zlTabAppRow
    Dim strLable As String, strAppRow As String, lngSpaces As Long
    Dim lngPos As Long, lngMax As Long, lngNumber As Long, blnNumber As Boolean, lngASC As Long
    
    '设置打印格式
    If UBound(Split(mstrPaperSet, ";")) >= 0 Then SaveSetting "ZLSOFT", "公共模块\zl9TendFilePrint\Default", "PaperSize", Val(Split(mstrPaperSet, ";")(0))
    If UBound(Split(mstrPaperSet, ";")) >= 1 Then SaveSetting "ZLSOFT", "公共模块\zl9TendFilePrint\Default", "Orientation", Val(Split(mstrPaperSet, ";")(1))
    If UBound(Split(mstrPaperSet, ";")) >= 2 Then SaveSetting "ZLSOFT", "公共模块\zl9TendFilePrint\Default", "Height", Val(Split(mstrPaperSet, ";")(2))
    If UBound(Split(mstrPaperSet, ";")) >= 3 Then SaveSetting "ZLSOFT", "公共模块\zl9TendFilePrint\Default", "Width", Val(Split(mstrPaperSet, ";")(3))
    If UBound(Split(mstrPaperSet, ";")) >= 4 Then objPrint.EmptyLeft = Round(ScaleY(Val(Split(mstrPaperSet, ";")(4)), vbTwips, vbMillimeters), 2)
    If UBound(Split(mstrPaperSet, ";")) >= 5 Then objPrint.EmptyRight = Round(ScaleY(Val(Split(mstrPaperSet, ";")(5)), vbTwips, vbMillimeters), 2)
    If UBound(Split(mstrPaperSet, ";")) >= 6 Then objPrint.EmptyUp = Round(ScaleX(Val(Split(mstrPaperSet, ";")(6)), vbTwips, vbMillimeters), 2)
    If UBound(Split(mstrPaperSet, ";")) >= 7 Then objPrint.EmptyDown = Round(ScaleX(Val(Split(mstrPaperSet, ";")(7)), vbTwips, vbMillimeters), 2)
    
    Set objPrint.Body = VsfData
    objPrint.Title.Text = lblTitle.Caption
    Set objPrint.Title.Font = lblTitle.Font
    Set objPrint.AppFont = lblSubhead.Font
    
    lngSpaces = lblSubhead.Height / 210
    strLable = lblSubhead.Caption
    lngMax = Len(strLable)
    lngNumber = 0
    lngStart = 1
    For lngPos = 1 To lngMax
        '如果数学超长,则把数字移到下一行显示
        lngASC = Asc(Mid(strLable, lngPos, 1))

        '检查是否超宽(长度超过行宽,或者遇到回车换行符)
        If TextWidth(Mid(strLable, lngStart, lngPos - lngStart + 1) & "测") > (Val(Split(mstrPaperSet, ";")(3)) - Val(Split(mstrPaperSet, ";")(4)) - Val(Split(mstrPaperSet, ";")(5)) - 500) Or lngPos = lngMax Or lngASC = 10 Then

            strAppRow = Mid(strLable, lngStart, lngPos - lngStart + 1)
            
            lngStart = lngPos + 1
            
            '输出表上项
            Set objAppRow = New zlTabAppRow
            Call objAppRow.Add(strAppRow)
            Call objPrint.UnderAppRows.Add(objAppRow)
        End If
    Next

    lngMax = Val(Split(mstrPaperSet, ";")(3))
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
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetStartPage() As Integer
    GetStartPage = mint当前起始页
End Function

Public Function GetCollectCols() As String
    GetCollectCols = mstrColCollect
End Function

Public Function GetPages() As Integer
    GetPages = mint结束页
End Function

Public Function isEndPage() As Boolean
    isEndPage = (mint页码 = mint结束页)
End Function

Public Sub PrevPage()
    If mint页码 > 1 Then
        mint页码 = mint页码 - 1
        Call ShowPage
    End If
End Sub

Public Function NextPage() As Boolean
    If mint页码 < mint结束页 Then
        mint页码 = mint页码 + 1
        NextPage = ShowPage
    End If
End Function

Private Sub WriteColor()
    Dim blnTag As Boolean
    Dim lngCount As Long
    Dim lngRow As Long
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
                '处理小结的显示
                If VsfData.TextMatrix(lngCount, mlngRowCount) Like "*|1" Then
                    If Val(VsfData.TextMatrix(lngCount, mlngCollectType)) <> 0 Then
                        VsfData.TextMatrix(lngCount, mlngDate) = VsfData.TextMatrix(lngCount, mlngCollectText)
                        VsfData.TextMatrix(lngCount, mlngTime) = VsfData.TextMatrix(lngCount, mlngCollectText)
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
    End With
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
    On Error GoTo errHand
    
    gstrSQL = "Select 数据转出 From 病案主页 Where 病人ID=[1] And 主页ID=[2]"
    Set rsTemp = OpenSQLRecord(gstrSQL, "判断数据是否转出", mlng病人id, mlng主页id)
    gblnMoved_HL = NVL(rsTemp!数据转出, 0) <> 0
    
    gstrSQL = " Select  /*+ RULE */ 开始时间,结束时间,格式ID,科室ID,归档人 From 病人护理文件 " & _
              " Where 病人ID=[1] And 主页ID=[2] And 婴儿=[3] And ID=[4] And Rownum<2"
    If gblnMoved_HL Then
        gstrSQL = Replace(gstrSQL, "病人护理文件", "H病人护理文件")
    End If
    Set rsTemp = OpenSQLRecord(gstrSQL, "提取护理文件数据", mlng病人id, mlng主页id, mint婴儿, mlng当前文件ID)
    If rsTemp.RecordCount <> 0 Then
        mlng格式ID = rsTemp!格式ID
        mlng科室ID = rsTemp!科室ID
        mstr开始时间 = Format(rsTemp!开始时间, "yyyy-MM-dd HH:mm:ss")
        mstr结束时间 = Format(rsTemp!结束时间, "yyyy-MM-dd HH:mm:ss")
    End If
    
    RaiseEvent AfterRowColChange("", False, mblnSign, mblnArchive)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitEnv()
    On Error GoTo errHand
    
    '打开现存在的所有护理记录项目
    gstrSQL = " Select  /*+ RULE */ 项目序号,项目名称,项目类型,项目性质,项目长度,项目小数,项目表示,项目单位,项目值域,护理等级,应用方式" & _
              " From 护理记录项目 B" & _
              " Where B.应用方式<>0 " & _
              " Order by 项目序号"
    Set mrsItems = OpenSQLRecord(gstrSQL, "打开现存在的所有护理记录项目")
    
    '提取护理文件编号规则
    gstrSQL = " Select NVL(参数值,0) AS 参数值 From zlparameters Where 模块=1255 and 参数名='护理文件页码规则'"
    Set rsTemp = OpenSQLRecord(gstrSQL, "提取护理文件编号规则")
    mintNORule = 0
    If rsTemp.RecordCount <> 0 Then
        mintNORule = rsTemp!参数值
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function ShowMe(ByVal frmParent As Form, ByVal lngFileID As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal intBaby As Integer, Optional ByVal intPage As Integer = 0) As Boolean
    '******************************************************************************************************************
    '功能： 显示护理记录文件内容
    '参数： frmParent           上级窗体对象
    '       lngPatiID           病人id
    '       lngPageID           主页id
    '       lngDeptID           要显示护理记录的科室
    '       intBaby             婴儿标志
    '       blnEditable         如果为假,说明是做为查询子窗体在使用,取消与编辑相关的功能
    '       blnClear            如果为真,连续重打mrsDataMap记录集;当换页时应传假,保留用户修改的数据以备显示、保存使用
    '返回： 无
    '******************************************************************************************************************
    Dim str文件名 As String, lng结束页 As Long, lng结束行 As Long
    Dim rsTemp As New ADODB.Recordset
    Dim str条件 As String
    Dim blnInitRec As Boolean, blnPrint As Boolean
    On Error GoTo errHand
    Err = 0
    
    mblnInit = False
    mlng当前文件ID = lngFileID
    mint页码 = intPage
    mlng病人id = lngPatiID
    mlng主页id = lngPageId
    mint婴儿 = intBaby
    mlngPageRows = frmAsk.intPageRows
    Set mfrmParent = frmParent
    
    If mrsItems.State = 0 Then
        Call InitEnv            '初始化环境
    End If
    Call InitVariable
    
    If mintNORule = 1 Then
        '肯定是一份文件打完了才打印下一份文件,所以,先取当前文件
        '取出当前文件最大页码
        gstrSQL = " Select MAX(B.打印页号) AS 页号" & _
                  " From 病人护理打印 B" & _
                  " Where B.文件ID=[1]"
        Set rsTemp = OpenSQLRecord(gstrSQL, "取最大打印页码", mlng当前文件ID)
        mlng开始页码 = NVL(rsTemp!页号, 0)
        
        If mlng开始页码 = 0 Then
            '取出本次住院所有文件的最大页码
            gstrSQL = " Select MAX(B.打印页号) AS 页号" & _
                      " From 病人护理文件 A,病人护理打印 B" & _
                      " Where A.ID=B.文件ID And A.病人ID=[1] And A.主页ID=[2] And A.婴儿=[3]"
            Set rsTemp = OpenSQLRecord(gstrSQL, "取最大打印页码", mlng病人id, mlng主页id, mint婴儿)
            mlng开始页码 = NVL(rsTemp!页号, 0) + 1
        Else
            '取当前文件最后一页的最后一条数据的打印行号,如果超过一页则+1
            gstrSQL = " Select MAX(B.打印行号) AS 页号" & _
                      " From 病人护理打印 B" & _
                      " Where B.文件ID=[1] AND B.打印页号=[2]"
            gstrSQL = " Select 行数,打印行号 From 病人护理打印 Where 文件ID=[1] And 打印页号=[2] And 打印行号=(" & gstrSQL & ")"
            Set rsTemp = OpenSQLRecord(gstrSQL, "取最大打印页码", mlng当前文件ID, mlng开始页码)
            If rsTemp!行数 + rsTemp!打印行号 - 1 > mlngPageRows Then mlng开始页码 = mlng开始页码 + 1
        End If
    End If
    
    mlng合并文件ID = 0
    gstrSQL = " Select  /*+ RULE */ MAX(打印页号) AS 打印页号,MAX(结束页号) AS 页码 From 病人护理打印 Where 文件ID=[1]"
    Set rsTemp = OpenSQLRecord(gstrSQL, "获取指定页码的数据发生时间范围", mlng当前文件ID)
    mlng打印页 = NVL(rsTemp!打印页号, 1)
    mint结束页 = mlng打印页
    If mint结束页 < NVL(rsTemp!页码, 1) Then mint结束页 = NVL(rsTemp!页码, 1)
    
    '从最后一次未打印完处接着打印
    'If mint页码 = 0 Or (mint页码 > 0 And gintPrintState <> 3) Then
    If intPage = 0 Then
        gstrSQL = " SELECT /*+ RULE */ 结束页号,结束行号 FROM 病人护理打印 " & vbNewLine & _
                  " WHERE 文件ID=[1] AND 打印人 IS NOT NULL" & vbNewLine & _
                  "       AND 发生时间=(SELECT MAX(发生时间) FROM 病人护理打印 WHERE 文件ID=[1] AND 打印人 IS NOT NULL)"
        Set rsTemp = OpenSQLRecord(gstrSQL, "从最后一次未打印完处接着打印", mlng当前文件ID)
        If rsTemp.RecordCount = 0 Then
            intPage = 1
            blnPrint = False
        Else
            intPage = rsTemp!结束页号
            If rsTemp!结束行号 = mlngPageRows Then intPage = intPage + 1
            blnPrint = True
        End If
    Else
        blnPrint = True
    End If
    mint当前起始页 = IIf(intPage > mlng打印页, mlng打印页, intPage)
    
    '第一页，且连续打印模式下
'    If intPage = 1 And (gintPrintState = 1 Or gintPrintState = 3) Then
        '检查当前文件是否与其它文件设置为合并打印
        gstrSQL = " Select A.ID,A.文件名称" & vbNewLine & _
                  " From 病人护理文件 A" & vbNewLine & _
                  " Where A.病人ID=[1] And A.主页ID=[2] And A.婴儿=[3] And A.续打ID=[4]"
        Set rsTemp = OpenSQLRecord(gstrSQL, "检查当前文件是否与其它文件设置为合并打印", mlng病人id, mlng主页id, mint婴儿, mlng当前文件ID)
        If rsTemp.RecordCount <> 0 Then
            mlng合并文件ID = rsTemp!Id
            str文件名 = rsTemp!文件名称
            '读出该文件的最后一页打印数据
            gstrSQL = " SELECT MAX(打印页号) AS 打印页号,MAX(打印行号) AS 打印行号 FROM 病人护理打印" & vbNewLine & _
                      " Where 文件ID=[1] And 打印人 Is Not NULL AND 打印页号=" & vbNewLine & _
                      "     (SELECT MAX(打印页号) AS 打印页号 FROM 病人护理打印 WHERE 文件ID=[1] AND 打印人 IS NOT NULL)"
            Set rsTemp = OpenSQLRecord(gstrSQL, "读出该文件的最后一页打印数据", mlng合并文件ID)
            lng结束页 = NVL(rsTemp!打印页号, 0)
            lng结束行 = NVL(rsTemp!打印行号, 0)
            If mlng合并文件ID <> 0 And lng结束页 = 0 Then
                MsgBox "当前文件与“" & str文件名 & "”设置为合并打印，而" & str文件名 & "还未打印！", vbInformation, gstrSysName
                Exit Function
            End If
            If rsTemp!打印行号 = mlngPageRows Then lng结束页 = lng结束页 + 1
            If mint当前起始页 < lng结束页 Then mint当前起始页 = lng结束页
            If mint结束页 < lng结束页 Then mint结束页 = mint当前起始页 + mint结束页
            mint页码 = mint当前起始页
        End If
'    End If
    If gintPrintState = 2 Then mint结束页 = mint当前起始页
    If gintPrintState = 1 And mlng合并文件ID <> 0 And blnPrint Then mint当前起始页 = mint结束页: intPage = mint结束页: mint页码 = mint结束页
    
    If mlng合并文件ID <> 0 And mint页码 = lng结束页 Then
        mlng当前文件ID = mlng合并文件ID
        mint页码 = lng结束页
        Call ReadStruDef
        Call InitRecords
        blnInitRec = True
        str条件 = " AND (P.打印页号=[5] OR (P.打印页号=[5]-1 AND P.打印行号+P.行数-1>[6]))"
        Call zlRefresh(str条件)
    End If
    
    mlng当前文件ID = lngFileID
    mint页码 = IIf(intPage > mlng打印页, mlng打印页, intPage) 'IIf(gintPrintState = 1, 1, intPage)
    Do While True
        'mint页码：没打印前是1，打印后是实际的页号，所以需要处理下，不然我重打第3页，始终不会显示当前文件的数据了
        If blnPrint Then
            If mint页码 > mint结束页 Then Exit Do
        Else
            If mint页码 > mint结束页 - mint当前起始页 + 1 Then Exit Do
        End If
        Call ReadStruDef
        If Not blnInitRec Then
            Call InitRecords
            blnInitRec = True
        End If
        
        Select Case gintPrintState
        Case 1  '续打
            If mlng合并文件ID = 0 Then
                If mint页码 = IIf(intPage > mlng打印页, mlng打印页, intPage) Then
                    str条件 = " AND (P.开始页号=[5] OR (P.结束页号=[5])) "
                Else
                    str条件 = " AND (P.开始页号=[5]) "
                End If
            Else
                If Not blnPrint Then
                    str条件 = " AND P.开始页号=[5]"
                Else
                    str条件 = " AND ((P.打印页号=[5] OR (P.打印页号=[5]-1 AND P.打印行号+P.行数-1>[6]))"
                    str条件 = str条件 & " OR (P.打印页号 Is NULL))"
                End If
            End If
        Case 2  '重打指定页
            str条件 = " AND (P.打印页号=[5] OR (P.打印页号=[5]-1 AND P.打印行号+P.行数-1>[6]))"
        Case 3  '从指定页开始连续重打
            str条件 = " AND (P.打印页号>=[5] OR (P.打印页号=[5]-1 AND P.打印行号+P.行数-1>[6]))"
        End Select
        
        Call zlRefresh(str条件)
        If gintPrintState > 1 Then Exit Do '重打读取指定页后直接退出
        mint页码 = mint页码 + 1
    Loop
    
    '绑定数据并设置护理记录单的格式,同时实现一行数据分行显示的功能
    mint页码 = IIf(intPage > mlng打印页, mlng打印页, intPage)
    Call PreTendFormat
    
    mblnInit = True
    mblnEditable = False
    ShowMe = True
'    Call OutputRsData(mrsSelItems)
    Exit Function
    
errHand:
    If ErrCenter() = 1 Then
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
    mlngRecord = -1
    mlngNoEditor = -1
    
    mblnShow = False
    mblnSign = False
    mblnArchive = False
    mblnEditAssistant = False
    
    Set mrsDataMap = New ADODB.Recordset
End Sub

Private Function GetStartRow(ByVal lngRow As Long) As Long
    Dim lngStart As Long
    Dim lngCurRows As Long, lngROWS As Long
    '提取数据起始行,超出本页则返回0
    '如果本页未显示全,则说明超出本页,也返回0
    '不允许在连续的数据行中插入新行
    
    lngROWS = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0))    '总行数
    lngCurRows = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(1)) '当前行
    If lngCurRows = 1 Then
        GetStartRow = lngRow
        Exit Function
    End If
    
    '寻找起始行
    For lngRow = lngRow To 3 Step -1
        If VsfData.TextMatrix(lngRow, mlngRowCount) = lngROWS & "|1" Then
            lngStart = lngRow
            Exit For
        End If
    Next
    
    GetStartRow = lngStart
End Function

Public Function GetDiagonal() As String
    GetDiagonal = "," & mstrCatercorner & "," & mstrCOLNothing & ","
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
End Sub

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

Private Sub OutputRsData(ByVal rsObj As ADODB.Recordset)
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
            Debug.Print Mid(strValues, 2)
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
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
