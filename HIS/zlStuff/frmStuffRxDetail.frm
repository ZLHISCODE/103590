VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStuffRxDetail 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame fraLine 
      Height          =   15
      Left            =   0
      TabIndex        =   6
      Top             =   480
      Width           =   9135
   End
   Begin VB.ComboBox cboPre 
      Height          =   300
      Left            =   6360
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   97
      Width           =   1455
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   0
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   39
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":0000
            Key             =   "打印11"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":039A
            Key             =   "当前"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":6BFC
            Key             =   "指示器"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":D45E
            Key             =   "附件"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":D9F8
            Key             =   "报告"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":DD92
            Key             =   "标志"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":E12C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":E4C6
            Key             =   "图标"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":E860
            Key             =   "选择"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":F272
            Key             =   "Person"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":15AD4
            Key             =   "未检"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":1C336
            Key             =   "在检"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":22B98
            Key             =   "已检"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":293FA
            Key             =   "Family"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":2FC5C
            Key             =   "分类"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":2FFF6
            Key             =   "分类_选中"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":30390
            Key             =   "套餐"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":36BF2
            Key             =   "类型"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":3D454
            Key             =   "照片"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":43CB6
            Key             =   "参数"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":4A518
            Key             =   "指标"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":50D7A
            Key             =   "体检"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":575DC
            Key             =   "病历样式"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":5DE3E
            Key             =   "病历文件"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":646A0
            Key             =   "规则"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":6AF02
            Key             =   "收费"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":6B914
            Key             =   "诊断"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":72176
            Key             =   "创建"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":789D8
            Key             =   "确认"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":7F23A
            Key             =   "开始"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":85A9C
            Key             =   "结束"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":8C2FE
            Key             =   "部份"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":8C698
            Key             =   "全部"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":8CA32
            Key             =   "部份总检"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":8CDCC
            Key             =   "全部总检"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":8D166
            Key             =   "总检"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":8D500
            Key             =   "打印"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":8DF12
            Key             =   "已经打印"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":8E924
            Key             =   "呼叫"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFDetail 
      Height          =   5640
      Left            =   30
      TabIndex        =   7
      Top             =   600
      Width           =   8760
      _cx             =   15452
      _cy             =   9948
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
      BackColorSel    =   16769992
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   10329501
      GridColorFixed  =   10329501
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   315
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmStuffRxDetail.frx":8EEBE
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   2
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
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
      AutoSizeMouse   =   -1  'True
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
   Begin VB.Label lblPre 
      Caption         =   "配料人"
      Height          =   255
      Left            =   5640
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblDept 
      Caption         =   "科室："
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblAge 
      Caption         =   "年龄："
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblSex 
      Caption         =   "性别："
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblName 
      Caption         =   "姓名："
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmStuffRxDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'定义表格列变量
'材料名称，规格，产地，批号，付数，数量，单价，金额
Private Const mconIntCol列数 = 32
Private mIntCol当前行 As Integer
Private mIntCol顺序号 As Integer
Private mIntCol材料名称 As Integer
Private mIntCol其它名 As Integer
Private mIntCol英文名 As Integer
Private mintcol序号 As Integer
Private mintcol规格 As Integer
Private mintcol批号 As Integer
Private mIntColId As Integer
Private mintcol材料id As Integer
Private mintcol批次 As Integer
Private mintcol单位 As Integer
Private mIntCol单价 As Integer
Private mintcol数量 As Integer
Private mIntCol金额 As Integer
Private mIntCol库存数 As Integer
Private mIntCol货位 As Integer
Private mIntCol已退数 As Integer
Private mIntCol准退数 As Integer
Private mIntCol退药数 As Integer
Private mIntCol分批 As Integer
Private mIntCol新批号 As Integer
Private mIntCol新效期 As Integer
Private mIntCol新产地 As Integer
Private mIntCol备注 As Integer
Private mIntCol医嘱id As Integer
Private mIntCol实际数量 As Integer
Private mIntCol包装 As Integer
Private mIntCol单据 As Integer
Private mIntColNO As Integer
Private mIntCol门诊标志 As Integer
Private mIntCol记录性质 As Integer
Private mIntCol可操作 As Integer

Private mintType As Integer   '当前页面
Private mstrUnallowSetColHide  As String   '不能设置隐藏的列
Private mstrUnallowShow As String     '不能显示的列

Private mrsWork As Recordset '当前操作的数据集
Private mstrVBMoneyForamt As String
Private mintMoneyDigit As Integer
Private mFMT As g_FmtString
Private mintUnit As Integer

Private Enum mListType
    待发料 = 0
    退料 = 1
End Enum

Private Sub Form_Load()
    '获取数量金额保留的小数位数
    mintUnit = Val(zlDatabase.GetPara("卫材单位", glngSys, glngModul, "0"))
    With mFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价)
        .FM_金额 = GetFmtString(mintUnit, g_金额)
        .FM_零售价 = GetFmtString(mintUnit, g_售价)
        .FM_数量 = GetFmtString(mintUnit, g_数量)
    End With
    
    mintMoneyDigit = GetDigit
    Call GetMoneyFormat
    Call LoadPrepare
    Call InitVSFDetail(mintType)
End Sub

Private Sub Form_Resize()
    With Me.fraLine
        .Left = 0
        .Width = Me.Width
    End With
    
    Me.VSFDetail.Move 80, VSFDetail.Top, Me.Width - 2 * Me.lblName.Left, Me.Height - VSFDetail.Top
End Sub


Public Sub InitVSFDetail(ByVal intType As Integer)
    Dim i As Integer
    Dim n As Integer
    Dim str列设置 As String
    Dim arr列设置 As Variant
    
    '''初始化列顺序
    '默认列顺序
    mIntCol当前行 = 0
    mIntCol顺序号 = 1
    mIntCol材料名称 = 2
    mIntCol其它名 = 3
    mIntCol英文名 = 4
    mintcol序号 = 5
    mintcol规格 = 6
    mintcol批号 = 7
    mIntColId = 8
    mintcol材料id = 9
    mintcol批次 = 10
    mintcol单位 = 11
    mIntCol单价 = 12
    mintcol数量 = 13
    mIntCol金额 = 14
    mIntCol货位 = 15
    mIntCol已退数 = 16
    mIntCol准退数 = 17
    mIntCol退药数 = 18
    mIntCol分批 = 19
    mIntCol新批号 = 20
    mIntCol新效期 = 21
    mIntCol新产地 = 22
    mIntCol备注 = 23
    mIntCol医嘱id = 24
    mIntCol实际数量 = 25
    mIntCol包装 = 26
    mIntCol单据 = 27
    mIntColNO = 28
    mIntCol门诊标志 = 29
    mIntCol记录性质 = 30
    mIntCol可操作 = 31
    
    '恢复用户自定义列顺序
    str列设置 = LoadListColState
    If str列设置 <> "" Then
        arr列设置 = Split(str列设置, "|")
        If UBound(arr列设置) + 1 <> mconIntCol列数 Then
            str列设置 = ""
        Else
            For n = 0 To UBound(arr列设置)
                SetColumnValue Split(arr列设置(n), ",")(0), n
            Next
        End If
    End If
     
    '初始化未发药清单
    With VSFDetail
        .Redraw = flexRDNone
        
        .Rows = 1
        .Rows = 2
        .Cols = mconIntCol列数
        
        .Cell(flexcpPicture, 1, mIntCol当前行, 1, mIntCol当前行) = Me.imgList.ListImages(2).Picture
        .Cell(flexcpPictureAlignment, 1, mIntCol当前行, .Rows - 1, mIntCol当前行) = flexPicAlignRightCenter
        
        VsfGridColFormat VSFDetail, mIntCol当前行, "", 250, flexAlignCenterCenter, "当前行"

        VsfGridColFormat VSFDetail, mIntCol顺序号, "顺序号", 0, flexAlignRightCenter, "顺序号"
        VsfGridColFormat VSFDetail, mIntCol材料名称, "材料名称", 2000, flexAlignLeftCenter, "材料名称"
        VsfGridColFormat VSFDetail, mIntCol其它名, "其它名", 1000, flexAlignLeftCenter, "其它名"
        VsfGridColFormat VSFDetail, mIntCol英文名, "英文名", 0, flexAlignCenterCenter, "英文名"
        VsfGridColFormat VSFDetail, mintcol序号, "序号", 0, flexAlignCenterCenter, "序号"
        VsfGridColFormat VSFDetail, mintcol规格, "规格", 1200, flexAlignLeftCenter, "规格"
        VsfGridColFormat VSFDetail, mintcol批号, "批号", 800, flexAlignLeftCenter, "批号"
        
        VsfGridColFormat VSFDetail, mIntColId, "Id", 0, flexAlignRightCenter, "Id"
        VsfGridColFormat VSFDetail, mintcol材料id, "材料id", 0, flexAlignLeftCenter, "材料id"
        VsfGridColFormat VSFDetail, mintcol批次, "批次", 0, flexAlignCenterCenter, "批次"
        VsfGridColFormat VSFDetail, mintcol单位, "单位", 800, flexAlignLeftCenter, "单位"
        VsfGridColFormat VSFDetail, mIntCol单价, "单价", 1000, flexAlignRightCenter, "单价"
        VsfGridColFormat VSFDetail, mintcol数量, "数量", 1000, flexAlignRightCenter, "数量"
        VsfGridColFormat VSFDetail, mIntCol金额, "金额", 1600, flexAlignRightCenter, "金额"
        VsfGridColFormat VSFDetail, mIntCol货位, "货位", 0, flexAlignCenterCenter, "货位"
        
        
        VsfGridColFormat VSFDetail, mIntCol已退数, "已退数", IIf(intType = 1, 1000, 0), flexAlignRightCenter, "已退数"
        VsfGridColFormat VSFDetail, mIntCol准退数, "准退数", IIf(intType = 1, 1000, 0), flexAlignRightCenter, "准退数"
        
        VsfGridColFormat VSFDetail, mIntCol退药数, "退药数", IIf(intType = 1, 1000, 0), flexAlignRightCenter, "退药数"
        VsfGridColFormat VSFDetail, mIntCol分批, "分批", 0, flexAlignCenterCenter, "分批"
        VsfGridColFormat VSFDetail, mIntCol新批号, "新批号", 0, flexAlignCenterCenter, "新批号"
        VsfGridColFormat VSFDetail, mIntCol新效期, "新效期", 0, flexAlignCenterCenter, "新效期"
        VsfGridColFormat VSFDetail, mIntCol新产地, "新产地", 0, flexAlignCenterCenter, "新产地"
        VsfGridColFormat VSFDetail, mIntCol备注, "备注", 0, flexAlignCenterCenter, "备注"
        VsfGridColFormat VSFDetail, mIntCol医嘱id, "医嘱id", 0, flexAlignCenterCenter, "医嘱id"
        VsfGridColFormat VSFDetail, mIntCol实际数量, "实际数量", 0, flexAlignRightCenter, "实际数量"
        VsfGridColFormat VSFDetail, mIntCol包装, "包装", 0, flexAlignCenterCenter, "包装"
        VsfGridColFormat VSFDetail, mIntCol单据, "单据", 0, flexAlignCenterCenter, "单据"
        VsfGridColFormat VSFDetail, mIntColNO, "NO", 0, flexAlignCenterCenter, "NO"
        VsfGridColFormat VSFDetail, mIntCol门诊标志, "门诊标志", 0, flexAlignCenterCenter, "门诊标志"
        VsfGridColFormat VSFDetail, mIntCol记录性质, "记录性质", 0, flexAlignCenterCenter, "记录性质"
        VsfGridColFormat VSFDetail, mIntCol可操作, "可操作", 0, flexAlignCenterCenter, "可操作"
        
        mstrUnallowSetColHide = "材料名称;规格;单位;单价;数量;金额"
        mstrUnallowShow = "库房ID;记录性质;门诊标志;NO;单据;包装;材料id;id;批次;医嘱id"
        If intType = 1 Then mstrUnallowShow = mstrUnallowShow & ";退药数;已退数;准退数"

        
        '恢复自定义列宽（不包括不允许显示的列）
        If str列设置 <> "" Then
            arr列设置 = Split(str列设置, "|")
            For n = 0 To UBound(arr列设置)
                If InStr(mstrUnallowShow, Split(arr列设置(n), ",")(0)) > 0 Then
                    For i = 0 To VSFDetail.Cols - 1
                        If Split(arr列设置(n), ",")(0) = VSFDetail.ColKey(i) Then
                            VSFDetail.ColWidth(i) = Val(Split(arr列设置(n), ",")(1))
                        End If
                    Next
                End If
            Next
        End If
        
        '重新生成网格
        .Select 0, 0, .Rows - 1, .Cols - 1
        .CellBorder &H9D9D9D, 1, 1, 1, 1, 1, 1
        
        .RowSel = 1
        .Redraw = flexRDBuffered
    End With
End Sub

Private Sub SetColumnValue(ByVal str列名 As String, ByVal intValue As Integer)
    Select Case str列名
        Case "材料名称"
            mIntCol材料名称 = intValue
        Case "其它名"
            mIntCol其它名 = intValue
        Case "英文名"
            mIntCol英文名 = intValue
            
        Case "序号"
            mintcol序号 = intValue
        Case "规格"
            mintcol规格 = intValue

        Case "批号"
            mintcol批号 = intValue
        Case "单位"
            mintcol单位 = intValue
        Case "单价"
            mIntCol单价 = intValue
    
        Case "数量"
            mintcol数量 = intValue
    End Select
                
End Sub


Public Sub WriteSendList(ByVal intType As Integer, ByVal rsTemp As Recordset, ByVal int可操作 As Integer)
    Dim i As Integer
    Dim intCount As Integer
    Dim dblMoney As Double
    Dim str合计 As String
    
    Set mrsWork = rsTemp
    
    With mrsWork
        Call InitVSFDetail(intType)
        
        If .RecordCount = 0 Then Exit Sub
        
        VSFDetail.Redraw = flexRDNone
        
        VSFDetail.Rows = .RecordCount + 1
        
        '填充病人基本信息
        Me.lblName.Caption = "姓名：" & Nvl(!姓名)
        Me.lblAge.Caption = "年龄：" & Nvl(!年龄)
        Me.lblSex.Caption = "性别：" & Nvl(!性别)
        Me.lblDept.Caption = "科室：" & Nvl(!科室)
        
        For i = 1 To .RecordCount
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("顺序号")) = i
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("材料名称")) = Nvl(!材料名称)
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("其它名")) = Nvl(!其它名)
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("序号")) = Nvl(!序号)
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("规格")) = Nvl(!规格)
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("批号")) = Nvl(!批号)
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("Id")) = !Id
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("材料id")) = !材料ID
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("批次")) = Nvl(!批次)
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("单位")) = Nvl(!单位)
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("单价")) = Format(!单价, mFMT.FM_零售价)
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("数量")) = Format(!数量, mFMT.FM_数量)
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("金额")) = Format(!金额, mFMT.FM_金额)
    
            If intType = 1 Then
                VSFDetail.TextMatrix(i, VSFDetail.ColIndex("已退数")) = Format(!已退数量, mFMT.FM_数量)
                VSFDetail.TextMatrix(i, VSFDetail.ColIndex("准退数")) = Format(!准退数, mFMT.FM_数量)
                VSFDetail.TextMatrix(i, VSFDetail.ColIndex("退药数")) = Format(!准退数, mFMT.FM_数量)
            End If
            
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("分批")) = !分批
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("备注")) = Nvl(!说明)
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("医嘱id")) = Nvl(!医嘱id)
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("实际数量")) = Format(!数量, mFMT.FM_数量)
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("单据")) = !单据
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("NO")) = Nvl(!NO)
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("门诊标志")) = !门诊标志
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("记录性质")) = !记录性质
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("可操作")) = int可操作

            dblMoney = dblMoney + Val(!金额)

            .MoveNext
        Next
        
         '添加合计行
        '最后空白行显示金额合计
        str合计 = zlStr.ChineseMoney(dblMoney)

        intCount = i
        VSFDetail.Rows = intCount + 1
        
        str合计 = "金额合计：" & Format(dblMoney, mstrVBMoneyForamt) & "  大写：" & str合计
        
        VSFDetail.Cell(flexcpText, intCount, 1, intCount, VSFDetail.Cols - 1) = str合计
        VSFDetail.Cell(flexcpAlignment, intCount, mIntCol顺序号, intCount, VSFDetail.Cols - 1) = flexAlignLeftCenter
        VSFDetail.Cell(flexcpFontBold, intCount, mIntCol顺序号, intCount, VSFDetail.Cols - 1) = True
        
        VSFDetail.MergeCells = flexMergeRestrictRows
        VSFDetail.MergeRow(VSFDetail.Rows - 1) = True
        
        '重新生成网格
        VSFDetail.Select 0, 0, VSFDetail.Rows - 1, VSFDetail.Cols - 1
        VSFDetail.CellBorder &H9D9D9D, 1, 1, 1, 1, 1, 1

        VSFDetail.Redraw = flexRDBuffered
        VSFDetail.Refresh
        
        VSFDetail.Row = VSFDetail.Rows - 1
    End With
End Sub

Private Function LoadListColState() As String
    Dim strType As String
    
    If Val(zlDatabase.GetPara("使用个性化风格")) = 0 Then Exit Function
    
    Select Case mintType
        Case mListType.待发料
            strType = "待发料"
        Case mListType.退料
            strType = "退料"
    End Select
    
    LoadListColState = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & Me.Name & "\" & TypeName(VSFDetail), strType, "")
End Function

Public Function GetWorkRs(ByVal intType As Integer, ByRef str退药数量 As String) As Recordset
'获取当前操作的数据集
'参数：intType，0-发料，1-退料,str退药数量,退药时为退药数量，发药为配药人
    Dim i As Integer
    
    Set GetWorkRs = mrsWork
    
    '获取退药数量
    If intType = 1 Then
        With VSFDetail
            For i = 1 To .Rows - 1
                str退药数量 = str退药数量 & "," & .TextMatrix(i, mIntColId) & "," & .TextMatrix(i, mIntCol退药数) & "|"
            Next
        End With
    Else
        str退药数量 = cboPre.Text
    End If
    
End Function


Public Sub SetFontSize(ByVal intFont As Integer)
    Me.VSFDetail.FontSize = intFont
End Sub

Private Sub LoadPrepare()
    Dim rsTemp As Recordset
    Dim lng发料部门ID As Long
    
    Set rsTemp = LoadPerson(UserInfo.Id)
    
    '装入发料部门数据
    With cboPre
        .Clear
        Do While Not rsTemp.EOF
            .AddItem rsTemp!名称
            .ItemData(.NewIndex) = rsTemp!Id
            If rsTemp!Id = UserInfo.Id Then
                .ListIndex = .NewIndex
            End If
            rsTemp.MoveNext
        Loop
        If .ListIndex = -1 Then .ListIndex = 0
    End With
End Sub

Private Sub VSFDetail_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = mIntCol退药数 Then
        With VSFDetail
            If Val(.TextMatrix(Row, Col)) > Val(.TextMatrix(Row, mIntCol准退数)) Then
               .TextMatrix(Row, Col) = .TextMatrix(Row, mIntCol准退数)
            ElseIf Val(.TextMatrix(Row, Col)) < 0 Then
                .TextMatrix(Row, Col) = 0
            End If
        End With
        
    End If
End Sub

Private Sub VSFDetail_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> mIntCol退药数 Or Row = 0 Then
        Cancel = True
        Exit Sub
    End If
    If Val(VSFDetail.TextMatrix(Row, Col)) <> 1 Then Cancel = True
End Sub

Private Sub VSFDetail_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    If Col = mIntCol材料名称 Then
        Position = mIntCol材料名称
    End If
End Sub

Private Sub vsfDetail_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = mIntCol当前行 Then Cancel = True
End Sub

Private Sub VSFDetail_EnterCell()
    With VSFDetail
        If .Row = 0 Then Exit Sub
        
        .Cell(flexcpPicture, 1, 0, .Rows - 1, 0) = Nothing
        .Cell(flexcpPicture, .Row, 0, .Row, 0) = Me.imgList.ListImages(2).Picture
    End With
End Sub

Private Sub VSFDetail_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With VSFDetail
        If Col = mIntCol退药数 Then
            If InStr("1234567890" + Chr(46) + Chr(8) + Chr(13), Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            ElseIf KeyAscii = Asc(".") Then
                If InStr(.EditText, ".") <> 0 Then     '只能存在一个小数点
                    KeyAscii = 0
                End If
            End If
        End If
    End With
End Sub


Private Sub GetMoneyFormat()
    Dim n As Integer
    Dim strOracleTmp As String
    Dim strVbTmp As String
    
    strOracleTmp = "999999990."
    strVbTmp = "########0."
    For n = 1 To mintMoneyDigit
        strOracleTmp = strOracleTmp & "0"
        strVbTmp = strVbTmp & "0"
    Next
    
    mstrVBMoneyForamt = strVbTmp
    
End Sub
