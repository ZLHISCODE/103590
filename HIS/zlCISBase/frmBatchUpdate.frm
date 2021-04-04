VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmBatchUpdate 
   BackColor       =   &H80000005&
   Caption         =   "批量修改规格"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7560
   Icon            =   "frmBatchUpdate.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5070
   ScaleWidth      =   7560
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtFind 
      Height          =   300
      Left            =   3000
      TabIndex        =   9
      Top             =   4680
      Width           =   1455
   End
   Begin VB.PictureBox picList 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   2520
      ScaleHeight     =   2415
      ScaleWidth      =   4815
      TabIndex        =   4
      Top             =   2160
      Width           =   4815
      Begin VB.CheckBox chk显示所有已修改的药品 
         BackColor       =   &H00FFFFFF&
         Caption         =   "显示所有已修改的药品"
         Height          =   375
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   2295
      End
      Begin VB.ComboBox cbo中药形态 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   37
         Width           =   1335
      End
      Begin VB.Frame fraSplit 
         Height          =   50
         Left            =   -120
         TabIndex        =   6
         Top             =   1440
         Width           =   3855
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfOtherName 
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   3615
         _cx             =   6376
         _cy             =   873
         Appearance      =   1
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
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
      Begin VSFlex8Ctl.VSFlexGrid vsfDetails 
         Height          =   975
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   3375
         _cx             =   5953
         _cy             =   1720
         Appearance      =   1
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
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
      Begin VB.Label lbl中药形态 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "中药形态"
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   2400
         TabIndex        =   12
         Top             =   97
         Width           =   720
      End
   End
   Begin VB.PictureBox picDetails 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   2520
      ScaleHeight     =   1815
      ScaleWidth      =   3495
      TabIndex        =   2
      Top             =   120
      Width           =   3495
      Begin XtremeSuiteControls.TabControl tbcDetails 
         Height          =   975
         Left            =   720
         TabIndex        =   3
         Top             =   360
         Width           =   1935
         _Version        =   589884
         _ExtentX        =   3413
         _ExtentY        =   1720
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picClass 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   120
      ScaleHeight     =   4215
      ScaleWidth      =   2175
      TabIndex        =   0
      Top             =   840
      Width           =   2175
      Begin VB.CheckBox chkAllDetails 
         Caption         =   "显示所有下级药品"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   1815
      End
      Begin MSComctlLib.TreeView tvwDetails 
         Height          =   4800
         Left            =   0
         TabIndex        =   8
         Tag             =   "1000"
         Top             =   600
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   8467
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "ImgTvw"
         Appearance      =   0
      End
   End
   Begin MSComctlLib.ImageList ImgTvw 
      Left            =   1680
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBatchUpdate.frx":6852
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBatchUpdate.frx":6DEC
            Key             =   "Edit"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBatchUpdate.frx":D64E
            Key             =   "Flag"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBatchUpdate.frx":13EB0
            Key             =   "规格U"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfAdditional 
      Height          =   1335
      Left            =   3000
      TabIndex        =   10
      Top             =   6120
      Visible         =   0   'False
      Width           =   2175
      _cx             =   3836
      _cy             =   2355
      Appearance      =   1
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager imgTool 
      Bindings        =   "frmBatchUpdate.frx":1444A
      Left            =   1320
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmBatchUpdate.frx":1445E
   End
   Begin XtremeDockingPane.DockingPane dkpPanel 
      Left            =   720
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmBatchUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
Private mint状态 As Integer         '记录是品种修改还是规格修改 1-品种 2-规格
Private mint次数 As Integer         '记录是不是首次加载 1-首次 2-不是
Private mblnData As Boolean  '用来判断是否在窗体加载时在树中有值
Private mstr上次节点 As String  '用来保存上次所选中的节点
Private mintRow As Integer        '用来记录上次所选中的行号
Private mintRow上次 As Integer
Private mintCol上次 As Integer
Private mbln库存 As Boolean        '用来记录是否有库存 true-有库存 flase-无库存
Private mbln药库分批 As Boolean    '药库分批 true-分批 false-不分批
Private mbln药房分批 As Boolean    '药房分批 true-分批 false-不分批
Private mint是否变价 As Integer     '定价还是时价 0-定价 1-时价
Private mstr类别 As String         '用来记录是什么分类 中草药，西成药、中成药
Private mstrNode As String         '记录被点击的节点的值
Private mint配置中心 As Integer
Private mstrPrivs As String        '记录用户有哪些权限
Private mrsRecord As ADODB.Recordset '用来记录选中节点查询出来的数据，为以后恢复数据做准备
Private mstrOtherName As String    '记录别名
Private mintOtherRow As Integer
Private mintExit As Integer         '用来记录退出时是否点击了保存按钮 1
Private mintLen住院单位 As Integer          '记录住院单位的长度
Private mintLen住院系数 As Integer
Private mstrChangedCell As String   '记录已改变的单元格位置
Private mintPos As Integer          '记录定位次数

Private mstr药品种类 As String     '记录药品种类
Private mstr存储库房 As String     '记录存储库房
Private mstr存储库房ID As String     '记录存储库房ID
Private mstr库房科室 As String     '记录库房科室的字符串
Private mstr库房科室ID As String     '记录库房科室ID的字符串
Private mint行号 As Integer     '记录行号
Private mint列号 As Integer     '记录列号
Private mrsMyRecords As ADODB.Recordset

'从参数表中取药品价格小数位数
Private mintCostDigit As Integer        '成本价小数位数
Private mintPriceDigit As Integer       '售价小数位数
Private mintSaleCostDigit As Integer
Private mintSalePriceDigit As Integer

Private mstrFind As String           '用来记录要要查询的值
Private mlngFind As Long
Private mlngFindFirst As Long
Private mrsFindName As ADODB.Recordset
Private mstrValue As String         '用来记录查找框中的值

Private mstrMatch As String         '匹配方式
Private mstrOldValue As String      '记录原来的单元格中的值
Private mblnClick As Boolean
Private mblnSetKey As Boolean       '判断是否设置了
Private mint当前单位 As Integer      '用来系统参数中设置的显示单位
Private mbln自管药 As Boolean        '用来记录是否是通过自管药设置方式打开的窗体

Private Const mcon应用于本列 As Integer = 101
Private Const mcon默认值 As Integer = 102
Private Const mcon保存 As Integer = 103
Private Const mcon帮助 As Integer = 104
Private Const mcon退出 As Integer = 105
Private Const mcon查找 As Integer = 106
Private Const mconFind As Integer = 107
Private Const mcon过滤 As Integer = 108
Private Const mcon定位已修改项目 As Integer = 109

Private Const cstcolor_backcolor = &H80000005   '窗口背景
Private Const CSTCOLOR_UNMODIFY = &HC0C0FF      '粉红 选项页颜色
Private Const CSTCOLOR_NORECORDS = &HFFFFFF     '白色 选项页颜色
Private Const mlngColor As Long = &H8000000F    '不能修改的列将背景颜色改成灰色
Private Const mlngApplyColor As Long = &H8000&  '单元格内容改变后为暗绿色
Private Const mlngBorderColor As Long = &H0&    '选中行边框颜色

Private mobjPopup As CommandBar
Private mobjControl As CommandBarControl
Private mcbrToolBar As CommandBar


'品种类别
Private Enum mVariList
    基本信息 = 0
    品种属性 = 1
    临床应用 = 2
End Enum
'品种列
Private Enum mVaricolumn
    品种_序号 = 0
    品种_id = 1
    品种_分类id = 2
    品种_药品分类
    品种_药品编码
    品种_通用名称
    品种_英文名称
    品种_拼音码
    品种_五笔码
    '品种属性
    品种_毒理分类
    品种_价值分类
    品种_货源情况
    品种_用药梯次
    品种_药品类型
    品种_剂型
    品种_原研药
    品种_专利药
    品种_单独定价
    品种_急救药
    品种_新药
    品种_原料药
    品种_单味使用
    品种_辅助用药
    品种_肿瘤药
    品种_溶媒
    品种_ATCCODE
    '临床应用
    品种_参考项目
    品种_处方职务
    品种_医保职务
    品种_处方限量
    品种_适用性别
    品种_剂量单位
    品种_皮试
    品种_抗生素
    品种_品种下长期医嘱
    品种_参考项目ID
    品种_Count
End Enum

'规格类别
Private Enum mSpecList
    基本信息 = 0
    商品信息 = 1
    包装单位 = 2
    价格信息 = 3
    药价属性 = 4
    分批管理 = 5
    临床应用 = 6
    配药属性 = 7
    存储库房 = 8
End Enum

'规格列
Private Enum mSpecColumn
    '基本信息
    规格_序号 = 0
    规格_id = 1
    规格_药名id = 2
    规格_通用名称
    规格_规格编码
    规格_药品规格
    规格_本位码
    规格_数字码
    规格_标识码
    规格_备选码
    规格_容量
    '商品信息
    规格_商品名称
    规格_生产商
    规格_原产地
    规格_来源分类
    规格_拼音码
    规格_五笔码
    规格_合同单位
    规格_批准文号
    规格_注册商标
    规格_GMP认证
    规格_易跌倒
    规格_带量采购
    规格_非常备药
    '包装单位
    规格_售价单位
    规格_剂量系数
    规格_剂量单位
    规格_住院单位
    规格_住院系数
    规格_门诊单位
    规格_门诊系数
    规格_药库单位
    规格_药库系数
    规格_送货单位
    规格_送货包装
    规格_申领单位
    规格_申领阀值
    规格_中药形态
    '价格信息
    规格_药价属性
    规格_零差价管理
    规格_成本价格
    规格_当前售价
    规格_采购限价
    规格_采购扣率
    规格_结算价
    规格_指导售价
    规格_加成率
    '药价属性
    规格_收入项目
    规格_病案费目
    规格_药价级别
    规格_屏蔽费别
    规格_医保类型
    '分批管理
    规格_药库分批
    规格_药房分批
    规格_保质期
    '临床应用
    规格_标识说明
    规格_发药类型
    规格_站点编号
    规格_DDD值
    规格_服务对象
    规格_住院分零使用
    规格_住院动态分零
    规格_门诊分零使用
    规格_是否摆药
    规格_基本药物
    规格_高危药品
    '配药属性
    规格_存储温度
    规格_存储条件
    规格_配药类型
    规格_不予调配
    规格_输液注意事项
    '存储库房
    规格_存储库房
    规格_存储库房id
    规格_服务科室
    规格_库房科室id
    规格_服务科室未改
    
    规格_招标药品
    规格_合同单位id
    规格_收入项目id
    规格_count
End Enum

Private Sub CheckValue(ByVal intRow As Integer, ByVal lng药品ID As Long)
    Dim rsTemp As ADODB.Recordset
    Dim dblTemp As Double

    gstrSql = ""
    On Error GoTo ErrHandle
    With vsfDetails
        If .TextMatrix(intRow, mSpecColumn.规格_药库分批) = "" Then
            .Cell(flexcpBackColor, intRow, mSpecColumn.规格_药房分批, intRow) = mlngColor: .TextMatrix(intRow, mSpecColumn.规格_药房分批) = ""
            .Cell(flexcpBackColor, intRow, mSpecColumn.规格_保质期, intRow) = mlngColor: .TextMatrix(intRow, mSpecColumn.规格_保质期) = 0
        Else
            If Val(.TextMatrix(intRow, mSpecColumn.规格_保质期)) = 0 Then
                .Cell(flexcpBackColor, intRow, mSpecColumn.规格_保质期, intRow) = mlngColor
            End If
        End If

        '提取显示当前售价
        If Mid(.TextMatrix(intRow, mSpecColumn.规格_药价属性), 1, 1) <> 0 Then
            '时价药品，取库存金额/库存数量做为其价格，无库存时取价表定价 非时价药品调价，取其价格记录中的价格
            gstrSql = "select Decode(K.库存数量,0,P.现价,K.库存金额/Nvl(K.库存数量,1)) as 现价,P.收入项目id" & _
                    " from 收费价目 P," & _
                    "     (Select nvl(Sum(实际金额),0) as 库存金额,nvl(Sum(实际数量),0) as 库存数量" & _
                    "      From 药品库存 Where 药品ID=[1]) K" & _
                    " where P.收费细目id=[1] and (P.终止日期 is null or Sysdate Between P.执行日期 And P.终止日期)" & _
                    GetPriceClassString("P")
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lng药品ID)
        End If

        If gstrSql <> "" Then
            If rsTemp.RecordCount > 0 Then
                If Val(mint当前单位) <> 0 Then
                    .TextMatrix(intRow, mSpecColumn.规格_当前售价) = FormatEx(rsTemp!现价 * Val(.TextMatrix(intRow, mSpecColumn.规格_药库系数)), mintPriceDigit)
                Else
                    .TextMatrix(intRow, mSpecColumn.规格_当前售价) = FormatEx(rsTemp!现价, mintPriceDigit)
                End If
                .TextMatrix(intRow, mSpecColumn.规格_收入项目id) = rsTemp!收入项目id
            End If
        End If

        '根据是否有发生，确定：药价属性、成本价格、零售价格可修改否
        gstrSql = " Select nvl(Count(*),0) From 药品收发记录 Where 药品ID=[1] And rownum<2"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lng药品ID)

        If rsTemp.Fields(0).Value > 0 Then
'            If Mid(.TextMatrix(intRow, mSpecColumn.规格_药价属性), 1, 1) <> 0 Then .Cell(flexcpBackColor, intRow, mSpecColumn.规格_药价属性, intRow) = mlngColor
            .Cell(flexcpBackColor, intRow, mSpecColumn.规格_成本价格, intRow) = mlngColor
            .Cell(flexcpBackColor, intRow, mSpecColumn.规格_当前售价, intRow) = mlngColor
            .Cell(flexcpBackColor, intRow, mSpecColumn.规格_收入项目, intRow) = mlngColor
            .Cell(flexcpBackColor, intRow, mSpecColumn.规格_住院系数, intRow) = mlngColor
            .Cell(flexcpBackColor, intRow, mSpecColumn.规格_门诊系数, intRow) = mlngColor
            .Cell(flexcpBackColor, intRow, mSpecColumn.规格_药库系数, intRow) = mlngColor
        End If

        '根据是否存在医嘱记录，确定剂量系数是否能够修改
        gstrSql = "Select 1 From 病人医嘱记录 Where 收费细目ID=[1] And Rownum=1"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lng药品ID)
        If rsTemp.RecordCount > 0 Then
            .Cell(flexcpBackColor, intRow, mSpecColumn.规格_剂量系数, intRow) = mlngColor
        End If

        '根据是否有库存，确定：分批特性可修改否
        gstrSql = " Select nvl(Count(*),0) From 药品库存 A,部门性质说明 B" & _
                 " Where A.药品ID=[1] And A.库房ID=B.部门ID And B.工作性质 Like '%药库'"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lng药品ID)

        If rsTemp.Fields(0).Value > 0 Then
            .Cell(flexcpBackColor, intRow, mSpecColumn.规格_药库分批, intRow) = mlngColor
            .Cell(flexcpBackColor, intRow, mSpecColumn.规格_保质期, intRow) = mlngColor
        End If
        If .TextMatrix(intRow, mSpecColumn.规格_药库分批) <> "" Then
            gstrSql = " Select nvl(Count(*),0) From 药品库存 A,部门性质说明 B" & _
                     " Where A.药品ID=[1] And A.库房ID=B.部门ID And (B.工作性质 Like '%药房' Or B.工作性质 Like '%制剂室')"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lng药品ID)

            If rsTemp.Fields(0).Value > 0 Then
                .Cell(flexcpBackColor, intRow, mSpecColumn.规格_药房分批, intRow) = mlngColor
                If .Cell(flexcpBackColor, intRow, mSpecColumn.规格_药库分批) <> mlngColor Then
                    .Cell(flexcpBackColor, intRow, mSpecColumn.规格_药库分批, intRow) = IIf(.TextMatrix(intRow, mSpecColumn.规格_药房分批) = "", cstcolor_backcolor, mlngColor)
                End If
            End If
        End If
            .Cell(flexcpBackColor, intRow, mSpecColumn.规格_结算价, intRow) = mlngColor
            If Val(Mid(.TextMatrix(intRow, mSpecColumn.规格_住院分零使用), 1, 1)) = 0 Then
                .Cell(flexcpBackColor, intRow, mSpecColumn.规格_住院动态分零, intRow) = mlngColor
            End If
            If .TextMatrix(intRow, mSpecColumn.规格_中药形态) = "散装" And mstrNode Like "中草药*" Then
                .Cell(flexcpBackColor, intRow, mSpecColumn.规格_住院分零使用, intRow) = mlngColor
                .Cell(flexcpBackColor, intRow, mSpecColumn.规格_门诊分零使用, intRow) = mlngColor
            End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub ShowMe(ByVal int状态 As Integer, ByVal strPrivs As String, ByVal bln自管药 As Boolean)
    '提供其他窗体访问本窗体的公用方法
    mint状态 = int状态
    mstrPrivs = strPrivs

    mbln自管药 = bln自管药
    Me.Show vbModal, frmMediLists
End Sub

Private Sub InitTreeView()
    With tvwDetails
        .LabelEdit = 1  '设置treeview为不可编辑状态
    End With
End Sub

Private Sub InitComandBars()
    '初始化工具栏，弹出菜单等
    Dim cbrControlMain As CommandBarControl
    Dim ctrCustom As CommandBarControlCustom
    Dim intCount As Integer

    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto

    Me.cbsMain.VisualTheme = xtpThemeOffice2003 + xtpThemeOfficeXP

    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16

    End With

    Me.cbsMain.EnableCustomization False
    Me.cbsMain.Icons = imgTool.Icons

    '工具栏定义
    Set mcbrToolBar = Me.cbsMain.Add("工具栏", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.EnableDocking xtpFlagStretched Or xtpFlagAlignAny

    With mcbrToolBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mcon应用于本列, "应用于本列")
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        Set cbrControlMain = .Add(xtpControlButton, mcon默认值, "恢复默认值")
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        cbrControlMain.Enabled = False
        
'        Set cbrControlMain = .Add(xtpControlButton, mcon过滤, "过滤")
'        cbrControlMain.BeginGroup = True
'        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
'        cbrControlMain.Enabled = False
'        cbrControlMain.ToolTipText = "过滤已修改的药品"

        Set cbrControlMain = .Add(xtpControlButton, mcon定位已修改项目, "定位已修改项目")
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        cbrControlMain.Enabled = False
        cbrControlMain.ToolTipText = "定位到修改过的行列"
        
        Set cbrControlMain = .Add(xtpControlButton, mcon保存, "保存")
        cbrControlMain.BeginGroup = True
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        cbrControlMain.Enabled = False
'        Set cbrControlMain = .Add(xtpControlButton, mcon帮助, "帮助")
'        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        cbrControlMain.BeginGroup = True
        Set cbrControlMain = .Add(xtpControlButton, mcon退出, "退出")
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字

        Set cbrControlMain = .Add(xtpControlLabel, mcon查找, "查找")
        cbrControlMain.Flags = xtpFlagRightAlign    '靠右对齐

        Set ctrCustom = mcbrToolBar.Controls.Add(xtpControlCustom, mconFind, "查询")
        ctrCustom.Handle = txtFind.hwnd
        ctrCustom.Flags = xtpFlagRightAlign
    End With

    cbsMain.Item(1).Delete

    '右键菜单
    Set mobjPopup = cbsMain.Add("Popup", xtpBarPopup)
    With mobjPopup.Controls
        Set mobjControl = .Add(xtpControlButton, mcon应用于本列, "应用于本列")
        Set mobjControl = .Add(xtpControlButton, mcon默认值, "恢复默认值")
'        Set mobjControl = .Add(xtpControlButton, mcon过滤, "过滤已修改项")
    End With

    '快键绑定
    With Me.cbsMain.KeyBindings
        .Add 0, VK_F3, mconFind
        .Add 0, VK_F2, mcon定位已修改项目
    End With
End Sub

Private Sub initPanel()
    '初始化分栏控件
    'DockingPane
    '-----------------------------------------------------
    Dim objPaneCon As Pane
    Dim objPaneDetail As Pane

    Me.dkpPanel.SetCommandBars Me.cbsMain
    Me.dkpPanel.Options.UseSplitterTracker = False '实时拖动
    Me.dkpPanel.Options.ThemedFloatingFrames = True
    Me.dkpPanel.Options.AlphaDockingContext = True

    Set objPaneCon = Me.dkpPanel.CreatePane(1, 200, 0, DockLeftOf, Nothing)
    objPaneCon.Options = PaneNoCloseable Or PaneNoFloatable
    objPaneCon.Title = "分类"
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strTemp As String

    Select Case Control.ID
        Case mcon应用于本列
            Call SetBatch
        Case mcon默认值
'            mrsRecord.MoveFirst
'            Call showColumn(mrsRecord, mstrNode)
            Call tvwDetails_NodeClick(tvwDetails.Nodes(tvwDetails.SelectedItem.Index))
        Case mcon保存
            Call Save
        Case mconFind
        
'        Case mcon过滤
'            Call FindChange
        Case mcon定位已修改项目
            Call FindChangeCell
        Case mcon退出
            Call ExitFrom
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long

    On Error Resume Next

    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    Me.picDetails.Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - lngTop

    Call initControl
End Sub

Private Sub chkAllDetails_Click()
    If mint状态 = 1 Then
        With vsfDetails
            If chkAllDetails.Value = 1 Then
                .ColWidth(mVaricolumn.品种_药品分类) = 2000
                .ColHidden(mVaricolumn.品种_药品分类) = False
            Else
                .ColHidden(mVaricolumn.品种_药品分类) = True
            End If
        End With
    End If
    Call tvwDetails_NodeClick(tvwDetails.SelectedItem)
End Sub

Private Sub dkpPanel_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
        Case 1
            Item.Handle = picClass.hwnd '将控件加入到dockingpanel控件中
    End Select
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim j As Integer

    Me.Width = 14000    '第一次加载时，窗体大小
    Me.Height = 9000

    Call RestoreWinState(Me, App.ProductName, Me.Caption)
    If GetSetting("ZLSOFT", "私有全局\" & gstrDBUser, "使用个性化风格", "1") = "1" Then
        chkAllDetails = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "是否显示下级", 0)
    End If

    mint次数 = 1
    Call InitTreeView   '初始化树
    Call InitComandBars '初始化菜单和工具栏
    Call initPanel  '初始化面板
    Call InitTabControl '向TabControl控件中加入窗体
    Call initControl    '初始化控件

    If mint状态 = 1 Then
        Call initColumn_品种信息    '初始化品种列
        mint次数 = 2
    ElseIf mint状态 = 2 Then
        Call initColumn_规格信息
        mint次数 = 2
        cbo中药形态.AddItem "全部形态"
        cbo中药形态.AddItem "0-散装"
        cbo中药形态.AddItem "1-中药饮片"
        cbo中药形态.AddItem "2-免煎剂"
        cbo中药形态.ListIndex = 0
    End If

'    mstrNode = "西成药"
    mblnData = ReadAndSendDataToTvw(mint状态)     '往树中填充值
    Call setColumn(0)    '初始化vsflexgrid控件列
    Call Set权限判断 '权限判断

    mstrMatch = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = 0, "%", "")  '匹配方式
    mint当前单位 = Val(zlDatabase.GetPara(29, glngSys))  '记录当前设置的显示单位

    mintCostDigit = GetDigit(1, 1, IIf(mint当前单位 = 0, 1, 4))
    mintPriceDigit = GetDigit(1, 2, IIf(mint当前单位 = 0, 1, 4))

    mintSaleCostDigit = GetDigit(1, 1, 1)
    mintSalePriceDigit = GetDigit(1, 2, 1)

    If tvwDetails.Nodes.Count > 0 Then
        If chkAllDetails = 1 And Not tvwDetails.Nodes(tvwDetails.SelectedItem.Index) Is Nothing Then
            Call tvwDetails_NodeClick(tvwDetails.Nodes(tvwDetails.SelectedItem.Index))
        End If
    End If
End Sub

Private Sub initControl()
    '重新布局控件位置
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long

    On Error Resume Next

    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    chkAllDetails.Move 0, 0, picClass.Width
    chk显示所有已修改的药品.Move 0, 0
    cbo中药形态.Move chk显示所有已修改的药品.Width + lbl中药形态.Width + 80, 40
    lbl中药形态.Move chk显示所有已修改的药品.Width + 40
    
    tvwDetails.Move 0, chkAllDetails.Height + chkAllDetails.Top, picClass.ScaleWidth, lngBottom - lngTop - chkAllDetails.Height - 385
    tbcDetails.Move 0, 0, picDetails.ScaleWidth, picDetails.ScaleHeight

    If mint状态 = 1 Then    '品种才有别名
        frmBatchUpdate.Caption = "品种批量修改"
        vsfDetails.Move 0, 400, picList.ScaleWidth, picList.ScaleHeight - 400
        fraSplit.Visible = False
        vsfOtherName.Visible = False
    Else    '规格无别名
        frmBatchUpdate.Caption = "规格批量修改"

        vsfDetails.Move 0, 400, picList.ScaleWidth, picList.ScaleHeight - 400
        fraSplit.Visible = False
        vsfOtherName.Visible = False
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Call initControl
End Sub

Private Sub InitTabControl()
    '初始化Tabcontrol控件
    With Me.tbcDetails
        .Icons = imgTool.Icons
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With

        If mint状态 = 1 Then    '品种
            .InsertItem(mVariList.基本信息, "基本信息", picList.hwnd, 0).Tag = "基本信息_"
            .InsertItem(mVariList.品种属性, "品种属性", picList.hwnd, 0).Tag = "品种属性_"
            .InsertItem(mVariList.临床应用, "临床应用", picList.hwnd, 0).Tag = "临床应用_"

            .Item(mVariList.品种属性).Selected = True
            .Item(mVariList.基本信息).Selected = True

        Else    '规格
            mint配置中心 = Val(zlDatabase.GetPara("配置中心", glngSys, , 0)) '0或空不启用，>0为对应的部门id

            .InsertItem(mSpecList.基本信息, "基本信息", picList.hwnd, 0).Tag = "基本信息_"
            .InsertItem(mSpecList.商品信息, "商品信息", picList.hwnd, 0).Tag = "商品信息_"
            .InsertItem(mSpecList.包装单位, "包装单位", picList.hwnd, 0).Tag = "包装单位_"
            .InsertItem(mSpecList.价格信息, "价格信息", picList.hwnd, 0).Tag = "价格信息_"
            .InsertItem(mSpecList.药价属性, "药价属性", picList.hwnd, 0).Tag = "药价属性_"
            .InsertItem(mSpecList.分批管理, "分批管理", picList.hwnd, 0).Tag = "分批管理_"
            .InsertItem(mSpecList.临床应用, "临床应用", picList.hwnd, 0).Tag = "临床应用_"
            .InsertItem(mSpecList.配药属性, "配药属性", picList.hwnd, 0).Tag = "配药属性_"
            .InsertItem(mSpecList.存储库房, "存储库房", picList.hwnd, 0).Tag = "存储库房_"
            If mint配置中心 = 0 Then '如果没有启用输液配置中心则不显示配药属性页面
                .Item(mSpecList.配药属性).Visible = False
            End If
            .Item(mSpecList.商品信息).Selected = True
            .Item(mSpecList.基本信息).Selected = True
        End If
    End With
    Call setTabControlColor(tbcDetails)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Recover
    mblnSetKey = False
    mintExit = 0
    Call SaveWinState(Me, App.ProductName, Me.Caption)

    If GetSetting("ZLSOFT", "私有全局\" & gstrDBUser, "使用个性化风格", "1") = "1" Then
        Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "是否显示下级", chkAllDetails.Value)
    End If
    Unload Me
End Sub
Private Sub fraSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 And vsfDetails.Height + y > 100 And fraSplit.Height + fraSplit.Top + y < Me.ScaleHeight - 1000 Then
        vsfDetails.Move 0, 0, picList.ScaleWidth, vsfDetails.Height + y
        fraSplit.Move 0, fraSplit.Top + y, picList.ScaleWidth, 50
        vsfOtherName.Move 0, fraSplit.Top + fraSplit.Height, picList.ScaleWidth, vsfOtherName.Height - y
    End If
End Sub

Private Sub tbcDetails_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'只有在药品品种基本页面才会有别名供用户修改

    If mint状态 = 1 Then    '品种
        fraSplit.Visible = False
        vsfOtherName.Visible = False
        vsfDetails.Move 0, 400, picList.ScaleWidth, picList.ScaleHeight - 400
    Else
        vsfDetails.Move 0, 400, picList.ScaleWidth, picList.ScaleHeight - 400
    End If

    Call setTabControlColor(tbcDetails)

    If mint次数 = 2 Then    '只有在列初始化后才能进行列设置
        Call setColumn(Item.Index)  '列隐藏显示设置
        Call SetBorder '设置行选中边框
    End If
End Sub

Private Sub setTabControlColor(ByVal objtbc As TabControl)
    '对Tabcontrol控件进行颜色判断
    Dim i As Integer

    With objtbc
        For i = 0 To .ItemCount - 1
            If .Item(i).Selected = True Then
                .Item(i).Color = CSTCOLOR_UNMODIFY
            Else
                .Item(i).Color = CSTCOLOR_NORECORDS
            End If
        Next
    End With
End Sub

Private Sub setColumn(ByVal intPageItem As Integer)
    '列显示与隐藏设置
    With vsfDetails
        .Editable = flexEDKbdMouse
        .MergeCells = flexMergeRestrictColumns
        .GridLineWidth = 2
        .GridLines = flexGridInset
        .GridColor = &H0&
        .ExplorerBar = flexExSortShowAndMove
        .AllowUserResizing = flexResizeColumns
        .AllowSelection = False '不能多选单元格
    End With

    With vsfDetails
        If mint状态 = 1 Then '品种
            cbo中药形态.Visible = False
            lbl中药形态.Visible = False
        
            vsfDetails.MergeCol(mVaricolumn.品种_药品分类) = True   '与上面的.MergeCells属性结合使用不同行同列内容相同的合并
            '基本信息
            .ColHidden(mVaricolumn.品种_序号) = True
            .ColHidden(mVaricolumn.品种_id) = True
            .ColHidden(mVaricolumn.品种_分类id) = True
            .ColHidden(mVaricolumn.品种_参考项目ID) = True

            .ColWidth(mVaricolumn.品种_通用名称) = 2000 '不隐藏该列
            .ColHidden(mVaricolumn.品种_药品编码) = IIf(intPageItem = mVariList.基本信息, False, True)
            .ColHidden(mVaricolumn.品种_英文名称) = IIf(intPageItem = mVariList.基本信息, False, True)
            .ColHidden(mVaricolumn.品种_拼音码) = IIf(intPageItem = mVariList.基本信息, False, True)
            .ColHidden(mVaricolumn.品种_五笔码) = IIf(intPageItem = mVariList.基本信息, False, True)

            '品种属性
            .ColHidden(mVaricolumn.品种_毒理分类) = IIf(intPageItem = mVariList.品种属性, False, True)
            .ColHidden(mVaricolumn.品种_价值分类) = IIf(intPageItem = mVariList.品种属性, False, True)
            .ColHidden(mVaricolumn.品种_货源情况) = IIf(intPageItem = mVariList.品种属性, False, True)
            .ColHidden(mVaricolumn.品种_用药梯次) = IIf(intPageItem = mVariList.品种属性, False, True)
            .ColHidden(mVaricolumn.品种_药品类型) = IIf(intPageItem = mVariList.品种属性, False, True)
            .ColHidden(mVaricolumn.品种_剂型) = IIf(intPageItem = mVariList.品种属性, False, True)
            .ColHidden(mVaricolumn.品种_原研药) = IIf(intPageItem = mVariList.品种属性, False, True)
            .ColHidden(mVaricolumn.品种_专利药) = IIf(intPageItem = mVariList.品种属性, False, True)
            .ColHidden(mVaricolumn.品种_单独定价) = IIf(intPageItem = mVariList.品种属性, False, True)
            .ColHidden(mVaricolumn.品种_急救药) = IIf(intPageItem = mVariList.品种属性, False, True)
            .ColHidden(mVaricolumn.品种_新药) = IIf(intPageItem = mVariList.品种属性, False, True)
            .ColHidden(mVaricolumn.品种_原料药) = IIf(intPageItem = mVariList.品种属性, False, True)
            .ColHidden(mVaricolumn.品种_辅助用药) = IIf(intPageItem = mVariList.品种属性, False, True)
            .ColHidden(mVaricolumn.品种_单味使用) = IIf(intPageItem = mVariList.品种属性, False, True)
            .ColHidden(mVaricolumn.品种_肿瘤药) = IIf(intPageItem = mVariList.品种属性, False, True)
            .ColHidden(mVaricolumn.品种_溶媒) = IIf(intPageItem = mVariList.品种属性, False, True)
            .ColHidden(mVaricolumn.品种_ATCCODE) = IIf(intPageItem = mVariList.品种属性, False, True)

            '临床应用
            .ColHidden(mVaricolumn.品种_参考项目) = IIf(intPageItem = mVariList.临床应用, False, True)
            .ColHidden(mVaricolumn.品种_处方职务) = IIf(intPageItem = mVariList.临床应用, False, True)
            .ColHidden(mVaricolumn.品种_医保职务) = IIf(intPageItem = mVariList.临床应用, False, True)
            .ColHidden(mVaricolumn.品种_处方限量) = IIf(intPageItem = mVariList.临床应用, False, True)
            .ColHidden(mVaricolumn.品种_适用性别) = IIf(intPageItem = mVariList.临床应用, False, True)
            .ColHidden(mVaricolumn.品种_剂量单位) = IIf(intPageItem = mVariList.临床应用, False, True)
            .ColHidden(mVaricolumn.品种_皮试) = IIf(intPageItem = mVariList.临床应用, False, True)
            .ColHidden(mVaricolumn.品种_抗生素) = IIf(intPageItem = mVariList.临床应用, False, True)
            .ColHidden(mVaricolumn.品种_品种下长期医嘱) = IIf(intPageItem = mVariList.临床应用, False, True)

            If mstrNode Like "中草药*" And intPageItem = mVariList.临床应用 Then
                .ColHidden(mVaricolumn.品种_皮试) = True
                .ColHidden(mVaricolumn.品种_抗生素) = True
                .ColHidden(mVaricolumn.品种_品种下长期医嘱) = True
            Else
                If intPageItem = mVariList.临床应用 Then
                    .ColHidden(mVaricolumn.品种_皮试) = False
                    .ColHidden(mVaricolumn.品种_抗生素) = False
                    .ColHidden(mVaricolumn.品种_品种下长期医嘱) = False
                End If
            End If

            If mstrNode Like "中草药*" Then
                If intPageItem = mVariList.品种属性 Then
                    .ColHidden(mVaricolumn.品种_单味使用) = False
                    .ColHidden(mVaricolumn.品种_原料药) = False
                End If
                .ColHidden(mVaricolumn.品种_剂型) = True
                .ColHidden(mVaricolumn.品种_原研药) = True
                .ColHidden(mVaricolumn.品种_专利药) = True
                .ColHidden(mVaricolumn.品种_单独定价) = True
                .ColHidden(mVaricolumn.品种_急救药) = True
                .ColHidden(mVaricolumn.品种_新药) = True
                .ColHidden(mVaricolumn.品种_肿瘤药) = True
                .ColHidden(mVaricolumn.品种_溶媒) = True
                .ColHidden(mVaricolumn.品种_ATCCODE) = True
            Else
                .ColHidden(mVaricolumn.品种_单味使用) = True
                If intPageItem = mVariList.品种属性 Then
                    .ColHidden(mVaricolumn.品种_剂型) = False
                    .ColHidden(mVaricolumn.品种_原研药) = False
                    .ColHidden(mVaricolumn.品种_专利药) = False
                    .ColHidden(mVaricolumn.品种_单独定价) = False
                    .ColHidden(mVaricolumn.品种_急救药) = False
                    .ColHidden(mVaricolumn.品种_新药) = False
                    .ColHidden(mVaricolumn.品种_原料药) = False
                    .ColHidden(mVaricolumn.品种_肿瘤药) = False
                    .ColHidden(mVaricolumn.品种_溶媒) = False
                    .ColHidden(mVaricolumn.品种_ATCCODE) = False
                End If
            End If

            If chkAllDetails.Value = 1 Then
                .ColHidden(mVaricolumn.品种_药品分类) = False
            Else
                .ColHidden(mVaricolumn.品种_药品分类) = True
            End If
        Else    '规格
            vsfDetails.MergeCol(mSpecColumn.规格_通用名称) = True    '设置合并

            .ColHidden(mSpecColumn.规格_序号) = True
            .ColWidth(mSpecColumn.规格_通用名称) = 1800
            .ColWidth(mSpecColumn.规格_药品规格) = 1500
            .ColHidden(mSpecColumn.规格_id) = True
            .ColHidden(mSpecColumn.规格_药名id) = True
            .ColHidden(mSpecColumn.规格_招标药品) = True
            .ColHidden(mSpecColumn.规格_合同单位id) = True
            .ColHidden(mSpecColumn.规格_收入项目id) = True
            '基本信息
            .ColHidden(mSpecColumn.规格_规格编码) = IIf(intPageItem = mSpecList.基本信息, False, True)
            .ColHidden(mSpecColumn.规格_本位码) = IIf(intPageItem = mSpecList.基本信息, False, True)
            .ColHidden(mSpecColumn.规格_数字码) = IIf(intPageItem = mSpecList.基本信息, False, True)
            .ColHidden(mSpecColumn.规格_标识码) = IIf(intPageItem = mSpecList.基本信息, False, True)
            .ColHidden(mSpecColumn.规格_备选码) = IIf(intPageItem = mSpecList.基本信息, False, True)

            If mstrNode Like "中草药*" Then
                .ColHidden(mSpecColumn.规格_容量) = True
                cbo中药形态.Visible = True
                lbl中药形态.Visible = True
            Else
                .ColHidden(mSpecColumn.规格_容量) = IIf(intPageItem = mSpecList.基本信息, False, True)
                cbo中药形态.Visible = False
                lbl中药形态.Visible = False
            End If
            '商品信息
            .ColHidden(mSpecColumn.规格_商品名称) = IIf(intPageItem = mSpecList.商品信息, False, True)
            .ColHidden(mSpecColumn.规格_生产商) = IIf(intPageItem = mSpecList.商品信息, False, True)
            .ColHidden(mSpecColumn.规格_原产地) = IIf(intPageItem = mSpecList.商品信息, False, True)
            .ColHidden(mSpecColumn.规格_来源分类) = IIf(intPageItem = mSpecList.商品信息, False, True)
            .ColHidden(mSpecColumn.规格_合同单位) = IIf(intPageItem = mSpecList.商品信息, False, True)
            .ColHidden(mSpecColumn.规格_批准文号) = IIf(intPageItem = mSpecList.商品信息, False, True)
            .ColHidden(mSpecColumn.规格_注册商标) = IIf(intPageItem = mSpecList.商品信息, False, True)
            .ColHidden(mSpecColumn.规格_拼音码) = IIf(intPageItem = mSpecList.商品信息, False, True)
            .ColHidden(mSpecColumn.规格_五笔码) = IIf(intPageItem = mSpecList.商品信息, False, True)
            .ColHidden(mSpecColumn.规格_GMP认证) = IIf(intPageItem = mSpecList.商品信息, False, True)
            .ColHidden(mSpecColumn.规格_非常备药) = IIf(intPageItem = mSpecList.商品信息, False, True)
            .ColHidden(mSpecColumn.规格_易跌倒) = IIf(intPageItem = mSpecList.商品信息, False, True)
            .ColHidden(mSpecColumn.规格_带量采购) = IIf(intPageItem = mSpecList.商品信息, False, True)
            If mstrNode Like "中草药*" Then
                .ColHidden(mSpecColumn.规格_拼音码) = True
                .ColHidden(mSpecColumn.规格_五笔码) = True
                .ColHidden(mSpecColumn.规格_GMP认证) = True
                .ColHidden(mSpecColumn.规格_商品名称) = True
                .ColHidden(mSpecColumn.规格_易跌倒) = True
                .ColHidden(mSpecColumn.规格_带量采购) = True
            End If
            If mstrNode Like "*成药*" Then
                .ColHidden(mSpecColumn.规格_原产地) = True
            End If
            
            '包装单位
            .ColHidden(mSpecColumn.规格_售价单位) = IIf(intPageItem = mSpecList.包装单位, False, True)
            .ColHidden(mSpecColumn.规格_剂量系数) = IIf(intPageItem = mSpecList.包装单位, False, True)
            .ColHidden(mSpecColumn.规格_剂量单位) = IIf(intPageItem = mSpecList.包装单位, False, True)
            .ColHidden(mSpecColumn.规格_住院单位) = IIf(intPageItem = mSpecList.包装单位, False, True)
            .ColHidden(mSpecColumn.规格_住院系数) = IIf(intPageItem = mSpecList.包装单位, False, True)
            .ColHidden(mSpecColumn.规格_门诊单位) = IIf(intPageItem = mSpecList.包装单位, False, True)
            .ColHidden(mSpecColumn.规格_门诊系数) = IIf(intPageItem = mSpecList.包装单位, False, True)
            .ColHidden(mSpecColumn.规格_药库单位) = IIf(intPageItem = mSpecList.包装单位, False, True)
            .ColHidden(mSpecColumn.规格_药库系数) = IIf(intPageItem = mSpecList.包装单位, False, True)
            .ColHidden(mSpecColumn.规格_送货单位) = IIf(intPageItem = mSpecList.包装单位, False, True)
            .ColHidden(mSpecColumn.规格_送货包装) = IIf(intPageItem = mSpecList.包装单位, False, True)
            .ColHidden(mSpecColumn.规格_申领单位) = IIf(intPageItem = mSpecList.包装单位, False, True)
            .ColHidden(mSpecColumn.规格_申领阀值) = IIf(intPageItem = mSpecList.包装单位, False, True)
            .ColHidden(mSpecColumn.规格_中药形态) = IIf(intPageItem = mSpecList.包装单位, False, True)

            If mstrNode Like "中草药*" Then
                If intPageItem = mSpecList.包装单位 Then
                    .ColHidden(mSpecColumn.规格_中药形态) = False
                    .ColHidden(mSpecColumn.规格_门诊单位) = True
                    .ColHidden(mSpecColumn.规格_门诊系数) = True
                    .ColHidden(mSpecColumn.规格_送货单位) = True
                    .ColHidden(mSpecColumn.规格_送货包装) = True
                    VsfGridColFormat vsfDetails, mSpecColumn.规格_住院单位, "药房单位", 1000, flexAlignLeftCenter, "药房单位"
                    VsfGridColFormat vsfDetails, mSpecColumn.规格_住院系数, "药房系数", 1000, flexAlignRightCenter, "药房系数"
                End If
            Else
                VsfGridColFormat vsfDetails, mSpecColumn.规格_住院单位, "住院单位", 1000, flexAlignLeftCenter, "住院单位"
                VsfGridColFormat vsfDetails, mSpecColumn.规格_住院系数, "住院系数", 1000, flexAlignRightCenter, "住院系数"
                .ColHidden(mSpecColumn.规格_中药形态) = True
            End If
            '价格信息
            .ColHidden(mSpecColumn.规格_药价属性) = IIf(intPageItem = mSpecList.价格信息, False, True)
            .ColHidden(mSpecColumn.规格_零差价管理) = IIf(intPageItem = mSpecList.价格信息, False, True)
            .ColHidden(mSpecColumn.规格_采购限价) = IIf(intPageItem = mSpecList.价格信息, False, True)
            .ColHidden(mSpecColumn.规格_采购扣率) = IIf(intPageItem = mSpecList.价格信息, False, True)
            .ColHidden(mSpecColumn.规格_结算价) = IIf(intPageItem = mSpecList.价格信息, False, True)
            .ColHidden(mSpecColumn.规格_指导售价) = IIf(intPageItem = mSpecList.价格信息, False, True)
            .ColHidden(mSpecColumn.规格_加成率) = IIf(intPageItem = mSpecList.价格信息, False, True)
            .ColHidden(mSpecColumn.规格_成本价格) = IIf(intPageItem = mSpecList.价格信息, False, True)
            .ColHidden(mSpecColumn.规格_当前售价) = IIf(intPageItem = mSpecList.价格信息, False, True)
            '药价属性
            .ColHidden(mSpecColumn.规格_收入项目) = IIf(intPageItem = mSpecList.药价属性, False, True)
            .ColHidden(mSpecColumn.规格_病案费目) = IIf(intPageItem = mSpecList.药价属性, False, True)
            .ColHidden(mSpecColumn.规格_药价级别) = IIf(intPageItem = mSpecList.药价属性, False, True)
            .ColHidden(mSpecColumn.规格_屏蔽费别) = IIf(intPageItem = mSpecList.药价属性, False, True)
            .ColHidden(mSpecColumn.规格_医保类型) = IIf(intPageItem = mSpecList.药价属性, False, True)
            '分批管理
            .ColHidden(mSpecColumn.规格_药库分批) = IIf(intPageItem = mSpecList.分批管理, False, True)
            .ColHidden(mSpecColumn.规格_药房分批) = IIf(intPageItem = mSpecList.分批管理, False, True)
            .ColHidden(mSpecColumn.规格_保质期) = IIf(intPageItem = mSpecList.分批管理, False, True)

            If mstrNode Like "中草药*" Then
                If intPageItem = mSpecList.分批管理 Then
                    .ColHidden(mSpecColumn.规格_保质期) = True
                End If
            Else
                If intPageItem = mSpecList.分批管理 Then
                    .ColHidden(mSpecColumn.规格_保质期) = False
                End If
            End If

            '临床应用
            .ColHidden(mSpecColumn.规格_标识说明) = IIf(intPageItem = mSpecList.临床应用, False, True)
            .ColHidden(mSpecColumn.规格_发药类型) = IIf(intPageItem = mSpecList.临床应用, False, True)
            .ColHidden(mSpecColumn.规格_站点编号) = IIf(intPageItem = mSpecList.临床应用, False, True)
            .ColHidden(mSpecColumn.规格_DDD值) = IIf(intPageItem = mSpecList.临床应用, False, True)
            .ColHidden(mSpecColumn.规格_服务对象) = IIf(intPageItem = mSpecList.临床应用, False, True)
            .ColHidden(mSpecColumn.规格_住院分零使用) = IIf(intPageItem = mSpecList.临床应用, False, True)
            .ColHidden(mSpecColumn.规格_门诊分零使用) = IIf(intPageItem = mSpecList.临床应用, False, True)
            .ColHidden(mSpecColumn.规格_基本药物) = IIf(intPageItem = mSpecList.临床应用, False, True)
            .ColHidden(mSpecColumn.规格_住院动态分零) = IIf(intPageItem = mSpecList.临床应用, False, True)
            .ColHidden(mSpecColumn.规格_高危药品) = IIf(intPageItem = mSpecList.临床应用, False, True)
            .ColHidden(mSpecColumn.规格_是否摆药) = IIf(intPageItem = mSpecList.临床应用, False, True)
            If mstrNode Like "中草药*" Then
                .ColHidden(mSpecColumn.规格_基本药物) = True
                .ColHidden(mSpecColumn.规格_住院动态分零) = True
                .ColHidden(mSpecColumn.规格_高危药品) = True
                .ColHidden(mSpecColumn.规格_DDD值) = True
            Else
                If intPageItem = mSpecList.临床应用 Then
                    .ColHidden(mSpecColumn.规格_基本药物) = False
                    .ColHidden(mSpecColumn.规格_住院动态分零) = False
                    .ColHidden(mSpecColumn.规格_高危药品) = False
                    .ColHidden(mSpecColumn.规格_DDD值) = False
                End If
            End If

            '配药属性
            .ColHidden(mSpecColumn.规格_存储温度) = IIf(intPageItem = mSpecList.配药属性, False, True)
            .ColHidden(mSpecColumn.规格_存储条件) = IIf(intPageItem = mSpecList.配药属性, False, True)
            .ColHidden(mSpecColumn.规格_配药类型) = IIf(intPageItem = mSpecList.配药属性, False, True)
            .ColHidden(mSpecColumn.规格_不予调配) = IIf(intPageItem = mSpecList.配药属性, False, True)
            .ColHidden(mSpecColumn.规格_输液注意事项) = IIf(intPageItem = mSpecList.配药属性, False, True)
            If mstrNode Like "中草药*" Then
                If intPageItem = mSpecList.配药属性 Then
                    tbcDetails.Item(mSpecList.基本信息).Selected = True
                End If
                tbcDetails.Item(mSpecList.配药属性).Visible = False
                .ColHidden(mSpecColumn.规格_存储温度) = True
                .ColHidden(mSpecColumn.规格_存储条件) = True
                .ColHidden(mSpecColumn.规格_配药类型) = True
                .ColHidden(mSpecColumn.规格_不予调配) = True
                .ColHidden(mSpecColumn.规格_输液注意事项) = True
            Else
                If mint配置中心 <> 0 Then
                    tbcDetails.Item(mSpecList.配药属性).Visible = True
                    If intPageItem = mSpecList.配药属性 Then
                        .ColHidden(mSpecColumn.规格_存储温度) = False
                        .ColHidden(mSpecColumn.规格_存储条件) = False
                        .ColHidden(mSpecColumn.规格_配药类型) = False
                        .ColHidden(mSpecColumn.规格_不予调配) = False
                        .ColHidden(mSpecColumn.规格_输液注意事项) = False
                    Else
                        .ColHidden(mSpecColumn.规格_存储温度) = True
                        .ColHidden(mSpecColumn.规格_存储条件) = True
                        .ColHidden(mSpecColumn.规格_配药类型) = True
                        .ColHidden(mSpecColumn.规格_不予调配) = True
                        .ColHidden(mSpecColumn.规格_输液注意事项) = True
                    End If
                End If
            End If
            
            '存储库房
            .ColHidden(mSpecColumn.规格_存储库房) = IIf(intPageItem = mSpecList.存储库房, False, True)
            .ColHidden(mSpecColumn.规格_存储库房id) = True
            .ColHidden(mSpecColumn.规格_服务科室) = IIf(intPageItem = mSpecList.存储库房, False, True)
            .ColHidden(mSpecColumn.规格_服务科室未改) = True
            .ColHidden(mSpecColumn.规格_库房科室id) = True
            
            If InStr(1, mstrPrivs, "存储库房") = 0 Then
                tbcDetails.Item(mSpecList.存储库房).Visible = False
            End If
        End If
    End With
End Sub

Private Sub initColumn_品种信息()
    '初始化基本信息页面
    Dim rsRecord As ADODB.Recordset
    Dim strTemp As String
    Dim i As Integer

    With vsfDetails
        .Cols = mVaricolumn.品种_Count
        .Rows = 1
        '基本信息
        VsfGridColFormat vsfDetails, mVaricolumn.品种_序号, "序号", 500, flexAlignCenterCenter, "序号"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_id, "id", 300, flexAlignCenterCenter, "id"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_分类id, "分类id", 300, flexAlignCenterCenter, "分类id"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_药品分类, "药品分类", 2000, flexAlignLeftCenter, "药品分类"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_药品编码, "药品编码", 1000, flexAlignLeftCenter, "药品编码"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_通用名称, "通用名称", 1000, flexAlignLeftCenter, "通用名称"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_英文名称, "英文名称", 1000, flexAlignLeftCenter, "英文名称"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_拼音码, "拼音码", 1000, flexAlignLeftCenter, "拼音码"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_五笔码, "五笔码", 1000, flexAlignLeftCenter, "五笔码"
        '品种属性
        VsfGridColFormat vsfDetails, mVaricolumn.品种_毒理分类, "毒理分类", 1000, flexAlignLeftCenter, "毒理分类"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_价值分类, "价值分类", 1000, flexAlignLeftCenter, "价值分类"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_货源情况, "货源情况", 1000, flexAlignLeftCenter, "货源情况"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_用药梯次, "用药梯次", 1000, flexAlignLeftCenter, "用药梯次"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_药品类型, "药品类型", 1000, flexAlignLeftCenter, "药品类型"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_剂型, "剂型", 2000, flexAlignLeftCenter, "剂型"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_原研药, "原研药", 800, flexAlignCenterCenter, "原研药"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_专利药, "专利药", 800, flexAlignCenterCenter, "专利药"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_单独定价, "单独定价", 1000, flexAlignCenterCenter, "单独定价"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_急救药, "急救药", 800, flexAlignCenterCenter, "急救药"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_新药, "新药", 800, flexAlignCenterCenter, "新药"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_原料药, "原料药", 1000, flexAlignCenterCenter, "原料药"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_辅助用药, "辅助用药", 1000, flexAlignCenterCenter, "辅助用药"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_单味使用, "单味使用", 1000, flexAlignCenterCenter, "单味使用"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_肿瘤药, "肿瘤药", 1000, flexAlignCenterCenter, "肿瘤药"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_溶媒, "溶媒", 800, flexAlignCenterCenter, "溶媒"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_ATCCODE, "ATCCODE", 1000, flexAlignRightCenter, "ATCCODE"
        '临床应用
        VsfGridColFormat vsfDetails, mVaricolumn.品种_参考项目, "参考项目", 1000, flexAlignLeftCenter, "参考项目"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_处方职务, "处方职务", 1000, flexAlignLeftCenter, "处方职务"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_医保职务, "医保职务", 1000, flexAlignLeftCenter, "医保职务"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_处方限量, "处方限量", 1000, flexAlignRightCenter, "处方限量"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_适用性别, "适用性别", 1500, flexAlignLeftCenter, "适用性别"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_剂量单位, "剂量单位", 1000, flexAlignLeftCenter, "剂量单位"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_皮试, "皮试", 800, flexAlignCenterCenter, "皮试"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_抗生素, "抗生素", 1500, flexAlignLeftCenter, "抗生素"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_品种下长期医嘱, "品种下长期医嘱", 1500, flexAlignCenterCenter, "品种下长期医嘱"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_参考项目ID, "参考项目ID", 10, flexAlignLeftCenter, "参考项目ID"

        If chkAllDetails.Value = 1 Then
            .ColWidth(mVaricolumn.品种_药品分类) = 2000
        Else
            .ColHidden(mVaricolumn.品种_药品分类) = True
        End If
    End With

    With vsfDetails
        '抗生素
        .ColComboList(mVaricolumn.品种_抗生素) = "0-非抗生素|1-非限制使用|2-限制使用|3-特殊使用"
        '剂量单位
        gstrSql = "select distinct 计算单位 from 诊疗项目目录 where 类别  in ('5','6','7')"
        Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "initColumn_品种信息")
        If Not rsRecord.EOF Then
            For i = 1 To rsRecord.RecordCount
                strTemp = strTemp & "|" & rsRecord!计算单位
                rsRecord.MoveNext
            Next
        End If
        .ColComboList(mVaricolumn.品种_剂量单位) = strTemp
        '剂型
        gstrSql = "select 编码||'-'|| 名称 as 剂型 from 药品剂型"
        Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "initColumn_品种信息")
        .ColComboList(mVaricolumn.品种_剂型) = vsfDetails.BuildComboList(rsRecord, "剂型")
        '参考项目
        .ColComboList(mVaricolumn.品种_参考项目) = "|..."
        '毒理分类
        gstrSql = "select 编码||'-'|| 名称 as 毒理分类 from 药品毒理分类"
        Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "initColumn_品种信息")
        .ColComboList(mVaricolumn.品种_毒理分类) = vsfDetails.BuildComboList(rsRecord, "毒理分类")
        '价值分类
        gstrSql = "select 编码||'-'|| 名称 as 价值分类 from 药品价值分类"
        Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "initColumn_品种信息")
        .ColComboList(mVaricolumn.品种_价值分类) = vsfDetails.BuildComboList(rsRecord, "价值分类")
        '货源情况
        gstrSql = "select 编码||'-'|| 名称 as 货源情况 from 药品货源情况"
        Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "initColumn_品种信息")
        .ColComboList(mVaricolumn.品种_货源情况) = vsfDetails.BuildComboList(rsRecord, "货源情况")
        '用药梯次
        gstrSql = "select 编码||'-'|| 名称 as 用药梯次 from 药品用药梯次"
        Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "initColumn_品种信息")
        .ColComboList(mVaricolumn.品种_用药梯次) = vsfDetails.BuildComboList(rsRecord, "用药梯次")
        '药品类型
        .ColComboList(mVaricolumn.品种_药品类型) = "0-未设定|1-处方药|2-甲类非处方药|3-乙类非处方药|4-非处方药|5-其它用药"
        '处方职务
        .ColComboList(mVaricolumn.品种_处方职务) = "0-不限|1-正高|2-副高|3-中级|4-助理/师级|5-员/士|9-待聘"
        '医保职务
        .ColComboList(mVaricolumn.品种_医保职务) = "0-不限|1-正高|2-副高|3-中级|4-助理/师级|5-员/士|9-待聘"
        '适用性别
        .ColComboList(mVaricolumn.品种_适用性别) = "0-无性别区分|1-男性|2-女性"
    End With
End Sub

Private Sub initColumn_规格信息()
    Dim rsRecord As ADODB.Recordset
    Dim strTemp As String
    Dim i As Integer

    '初始化规格列
    On Error GoTo ErrHandle
    With vsfDetails
        .Cols = mSpecColumn.规格_count
        .Rows = 1
        '基本信息
        VsfGridColFormat vsfDetails, mSpecColumn.规格_序号, "序号", 500, flexAlignCenterCenter, "序号"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_id, "id", 300, flexAlignLeftCenter, "id"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_药名id, "药名id", 600, flexAlignCenterCenter, "药名id"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_通用名称, "通用名称", 1000, flexAlignLeftCenter, "通用名称"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_规格编码, "规格编码", 1000, flexAlignLeftCenter, "规格编码"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_药品规格, "药品规格", 1500, flexAlignLeftCenter, "药品规格"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_本位码, "本位码", 2500, flexAlignLeftCenter, "本位码"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_数字码, "数字码", 1000, flexAlignLeftCenter, "数字码"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_标识码, "标识码", 1000, flexAlignLeftCenter, "标识码"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_备选码, "备选码", 1000, flexAlignLeftCenter, "备选码"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_容量, "容量", 800, flexAlignRightCenter, "容量"
        '商品信息
        VsfGridColFormat vsfDetails, mSpecColumn.规格_商品名称, "商品名称", 1500, flexAlignLeftCenter, "商品名称"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_生产商, "生产商", 1500, flexAlignLeftCenter, "生产商"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_原产地, "原产地", 1500, flexAlignLeftCenter, "原产地"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_来源分类, "来源分类", 1000, flexAlignLeftCenter, "来源分类"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_拼音码, "拼音码", 1000, flexAlignLeftCenter, "拼音码"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_五笔码, "五笔码", 1000, flexAlignLeftCenter, "五笔码"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_合同单位, "合同单位", 1000, flexAlignLeftCenter, "合同单位"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_批准文号, "批准文号", 1000, flexAlignLeftCenter, "批准文号"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_注册商标, "注册商标", 1000, flexAlignLeftCenter, "注册商标"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_GMP认证, "GMP认证", 800, flexAlignCenterCenter, "GMP认证"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_非常备药, "非常备药", 1000, flexAlignCenterCenter, "非常备药"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_易跌倒, "易跌倒", 800, flexAlignCenterCenter, "易跌倒"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_带量采购, "带量采购", 1000, flexAlignCenterCenter, "带量采购"
        '包装单位
        VsfGridColFormat vsfDetails, mSpecColumn.规格_售价单位, "售价单位", 1000, flexAlignLeftCenter, "售价单位"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_剂量系数, "剂量系数", 1000, flexAlignRightCenter, "剂量系数"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_剂量单位, "剂量单位", 1000, flexAlignRightCenter, "剂量单位"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_住院单位, "住院单位", 1000, flexAlignLeftCenter, "住院单位"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_住院系数, "住院系数", 1000, flexAlignRightCenter, "住院系数"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_门诊单位, "门诊单位", 1000, flexAlignLeftCenter, "门诊单位"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_门诊系数, "门诊系数", 1000, flexAlignRightCenter, "门诊系数"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_药库单位, "药库单位", 1000, flexAlignLeftCenter, "药库单位"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_药库系数, "药库系数", 1000, flexAlignRightCenter, "药库系数"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_送货单位, "送货单位", 1000, flexAlignLeftCenter, "送货单位"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_送货包装, "送货包装", 1000, flexAlignRightCenter, "送货包装"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_申领单位, "申领单位", 1500, flexAlignLeftCenter, "申领单位"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_申领阀值, "申领阀值", 1000, flexAlignRightCenter, "申领阀值"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_中药形态, "中药形态", 1500, flexAlignRightCenter, "中药形态"
        '价格信息
        VsfGridColFormat vsfDetails, mSpecColumn.规格_药价属性, "药价属性", 900, flexAlignLeftCenter, "药价属性"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_零差价管理, "零差价管理", 1200, flexAlignCenterCenter, "零差价管理"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_采购限价, "采购限价", 1000, flexAlignRightCenter, "采购限价"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_采购扣率, "采购扣率", 1000, flexAlignRightCenter, "采购扣率"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_结算价, "结算价", 1000, flexAlignRightCenter, "结算价"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_指导售价, "指导售价", 1000, flexAlignRightCenter, "指导售价"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_加成率, "加成率", 1000, flexAlignRightCenter, "加成率"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_成本价格, "成本价格", 1000, flexAlignRightCenter, "成本价格"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_当前售价, "当前售价", 1000, flexAlignRightCenter, "当前售价"
        '药价属性
        VsfGridColFormat vsfDetails, mSpecColumn.规格_收入项目, "收入项目", 1500, flexAlignLeftCenter, "收入项目"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_病案费目, "病案费目", 1000, flexAlignLeftCenter, "病案费目"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_药价级别, "药价级别", 1000, flexAlignLeftCenter, "药价级别"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_屏蔽费别, "屏蔽费别", 900, flexAlignCenterCenter, "屏蔽费别"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_医保类型, "医保类型", 1000, flexAlignLeftCenter, "医保类型"
        '分批管理
        VsfGridColFormat vsfDetails, mSpecColumn.规格_药库分批, "药库分批", 800, flexAlignCenterCenter, "药库分批"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_药房分批, "药房分批", 800, flexAlignCenterCenter, "药房分批"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_保质期, "保质期", 1000, flexAlignRightCenter, "保质期"
        '临床应用
        VsfGridColFormat vsfDetails, mSpecColumn.规格_标识说明, "标识说明", 1000, flexAlignLeftCenter, "标识说明"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_发药类型, "发药类型", 900, flexAlignLeftCenter, "发药类型"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_站点编号, "站点编号", 900, flexAlignLeftCenter, "站点编号"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_DDD值, "DDD值", 900, flexAlignLeftCenter, "DDD值"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_服务对象, "服务对象", 1500, flexAlignLeftCenter, "服务对象"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_住院分零使用, "住院分零使用", 1300, flexAlignLeftCenter, "住院分零使用"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_门诊分零使用, "门诊分零使用", 1300, flexAlignLeftCenter, "门诊分零使用"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_住院动态分零, "住院动态分零", 1300, flexAlignCenterCenter, "住院动态分零"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_是否摆药, "是否摆药", 900, flexAlignCenterCenter, "是否摆药"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_基本药物, "基本药物", 1400, flexAlignLeftCenter, "基本药物"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_高危药品, "高危药品", 1000, flexAlignLeftCenter, "高危药品"
        '配药属性
        VsfGridColFormat vsfDetails, mSpecColumn.规格_存储温度, "存储温度", 1000, flexAlignLeftCenter, "存储温度"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_存储条件, "存储条件", 1000, flexAlignCenterCenter, "存储条件"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_配药类型, "配药类型", 1000, flexAlignLeftCenter, "配药类型"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_不予调配, "不予调配", 1000, flexAlignCenterCenter, "不予调配"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_输液注意事项, "输液注意事项", 5000, flexAlignLeftCenter, "输液注意事项"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_招标药品, "招标药品", 1000, flexAlignLeftCenter, "招标药品"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_合同单位id, "合同单位id", 1000, flexAlignLeftCenter, "合同单位id"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_收入项目id, "收入项目id", 1000, flexAlignLeftCenter, "收入项目id"
        '存储库房
        VsfGridColFormat vsfDetails, mSpecColumn.规格_存储库房, "存储库房", 4000, flexAlignLeftCenter, "存储库房"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_存储库房id, "存储库房id", 3000, flexAlignLeftCenter, "存储库房id"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_服务科室, "服务科室", 4000, flexAlignLeftCenter, "服务科室"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_服务科室未改, "服务科室未改", 3000, flexAlignLeftCenter, "服务科室未改"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_库房科室id, "库房科室id", 3000, flexAlignLeftCenter, "库房科室id"
    End With

    With vsfDetails
        '生产商
        .ColComboList(mSpecColumn.规格_生产商) = "|..."
        '原产地
        .ColComboList(mSpecColumn.规格_原产地) = "|..."
        '来源分类
        gstrSql = "select 编码||'-'|| 名称 as 来源分类 from 药品来源分类"
        Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "initColumn_品种信息")
        .ColComboList(mSpecColumn.规格_来源分类) = vsfDetails.BuildComboList(rsRecord, "来源分类")
        '合同单位
        .ColComboList(mSpecColumn.规格_合同单位) = "|..."
        '发药类型
        gstrSql = "select 名称 as 发药类型 from 发药类型"
        Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "initColumn_品种信息")
        .ColComboList(mSpecColumn.规格_发药类型) = vsfDetails.BuildComboList(rsRecord, "发药类型")
        '站点编号
        gstrSql = "select 编号||'-'||名称 as 站点编号 from zlnodelist"
        Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "initColumn_品种信息")
        .ColComboList(mSpecColumn.规格_站点编号) = vsfDetails.BuildComboList(rsRecord, "站点编号")
        '申领单位
        .ColComboList(mSpecColumn.规格_申领单位) = "1-售价单位|2-住院单位|3-门诊单位|4-药库单位"
        '药价属性
        .ColComboList(mSpecColumn.规格_药价属性) = "0-定价|1-时价"
        '基本药物
        gstrSql = "Select 名称 as 基本药物  From 基本药物说明  Order By 编码"
        Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "initColumn_品种信息")
        For i = 0 To rsRecord.RecordCount - 1
            strTemp = strTemp & rsRecord!基本药物 & "|"
            rsRecord.MoveNext
        Next
        .ColComboList(mSpecColumn.规格_基本药物) = "| |" & strTemp
        '收入项目
        gstrSql = "Select ID, '[' || 编码 || ']' || 名称 As 收入项目" & _
                  "  From 收入项目" & _
                  "  Where 末级 = 1 And (撤档时间 Is Null Or 撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))" & _
                  "  Order By 编码"
        Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "initColumn_品种信息")
        .ColComboList(mSpecColumn.规格_收入项目) = vsfDetails.BuildComboList(rsRecord, "收入项目")
        '病案费目
        .ColComboList(mSpecColumn.规格_病案费目) = "..."
        '药价管理级别
        gstrSql = "select 编码||'-'||名称 as 管理级别 from 药价管理级别"
        Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "initColumn_品种信息")
        .ColComboList(mSpecColumn.规格_药价级别) = vsfDetails.BuildComboList(rsRecord, "管理级别")
        '医保类型
        gstrSql = "Select 编码||'-'||名称 as 医保类型 From 费用类型 where 性质=1 Order By 编码"
        Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "initColumn_品种信息")
        .ColComboList(mSpecColumn.规格_医保类型) = vsfDetails.BuildComboList(rsRecord, "医保类型")
        '服务对象
        .ColComboList(mSpecColumn.规格_服务对象) = "0-不应用于病人|1-门诊|2-住院|3-门诊和住院"
        '住院/门诊分零使用
        .ColComboList(mSpecColumn.规格_住院分零使用) = "0-可以分零|1-不可分零|2-一次性使用|3-分零后一天内有效|4-分零后两天内有效|5-分零后三天内有效"
        .ColComboList(mSpecColumn.规格_门诊分零使用) = "0-可以分零|1-不可分零|2-一次性使用|3-分零后一天内有效|4-分零后两天内有效|5-分零后三天内有效"
        '存储温度
        .ColComboList(mSpecColumn.规格_存储温度) = " |1-常温(0-30℃)|2-阴凉(20℃以下)|3-冷藏(2-8℃)"
        '配药类型
        .ColComboList(mSpecColumn.规格_配药类型) = " |1-抗生素|2-细胞毒|3-营养(普通)"
        '中药形态
        .ColComboList(mSpecColumn.规格_中药形态) = "0-散装|1-中药饮片|2-免煎剂"
        '高危药品
        .ColComboList(mSpecColumn.规格_高危药品) = " |1-A级|2-B级|3-C级"
        '是否摆药
        .ColComboList(mSpecColumn.规格_是否摆药) = "是|否"
        '存储库房
        .ColComboList(mSpecColumn.规格_存储库房) = "..."
        '服务科室
        .ColComboList(mSpecColumn.规格_服务科室) = "..."
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function ReadAndSendDataToTvw(ByVal int状态 As Integer) As Boolean
'功能：用来向树中填充节点
'参数 int状态 用来判断界面加载时是品种修改还是规格修改

    Dim NodeThis As Node
    Dim Int末级 As Integer
    Dim lng库房ID As Long
    Dim rs材质分类 As ADODB.Recordset
    Dim recdata As ADODB.Recordset

    '药品用途分类是否有数据
    ReadAndSendDataToTvw = False
    On Error GoTo ErrHandle
    gstrSql = " Select 编码,名称 From 诊疗项目类别 " & _
              " Where Instr([1],编码,1) > 0 " & _
              " Order by 编码"
    Set rs材质分类 = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, "567")

    If rs材质分类 Is Nothing Then
        Exit Function
    End If

'    Set rs材质分类 = GetFilter分类(rs材质分类)
    With tvwDetails
        .Nodes.Clear
        Do While Not rs材质分类.EOF
            .Nodes.Add , , "Root" & rs材质分类!名称, rs材质分类!名称, 1, 1
            .Nodes("Root" & rs材质分类!名称).Tag = rs材质分类!编码
            rs材质分类.MoveNext
        Loop
    End With

    gstrSql = "Select ID, 上级id, 编码, 名称, Decode(类型, 1, '西成药', 2, '中成药', 3, '中草药') 分类, '分类' As 类别" & _
            " From 诊疗分类目录" & _
            " Where 类型 In ('1', '2', '3') And Nvl(To_Char(撤档时间, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01' " & _
            " Start With 上级id Is Null" & _
            " Connect By Prior ID = 上级id"

    Set recdata = zlDatabase.OpenSQLRecord(gstrSql, "ReadAndSendDataToTvw")

    If recdata.EOF Then
        MsgBox "请初始化药品用途分类（药品用途分类）！", vbInformation, gstrSysName
        Exit Function
    End If

    With recdata
        Do While Not .EOF
            If IsNull(!上级ID) Then
                Set NodeThis = tvwDetails.Nodes.Add("Root" & !分类, 4, "K_" & !ID, !名称, 1, 1)
            Else
                Set NodeThis = tvwDetails.Nodes.Add("K_" & !上级ID, 4, "K_" & !ID, !名称, 1, 1)
            End If
            NodeThis.Tag = !分类 & "-" & !类别  '存放分类类型:1-西成药,2-中成药,3-中草药
            .MoveNext
        Loop
    End With

    If int状态 <> 1 Then '品种修改
        gstrSql = "Select ID, 分类id, 编码, 名称, Decode(类别, 5, '西成药', 6, '中成药', 7, '中草药') 分类, '品种' As 类别" & _
                  "  From 诊疗项目目录" & _
                  "  Where 分类id In (Select ID" & _
                                   " From 诊疗分类目录" & _
                                   " Where 类型 In ('1', '2', '3') And Nvl(To_Char(撤档时间, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01'" & _
                                   " Start With 上级id Is Null" & _
                                   " Connect By Prior ID = 上级id)"
        Set recdata = zlDatabase.OpenSQLRecord(gstrSql, "品种")

        With recdata
            Do While Not .EOF
                Set NodeThis = tvwDetails.Nodes.Add("K_" & !分类id, 4, !类别 & "K_" & !ID, !名称, 1, 1)
                NodeThis.Tag = !分类 & "-" & !类别  '存放分类类型:1-西成药,2-中成药,3-中草药
                .MoveNext
            Loop
        End With
    End If

    Call GetFilter权限  '根据用户所具有的权限来过滤数据

    With tvwDetails
        If .Nodes.Count <> 0 Then
            .Nodes(1).Selected = True
            If .Nodes(1).Children <> 0 Then
                Int末级 = 1
                .Nodes(Int末级).Child.Selected = True
                .SelectedItem.Selected = True
            ElseIf .Nodes(2).Children <> 0 Then
                Int末级 = 2
                .Nodes(Int末级).Child.Selected = True
                .SelectedItem.Selected = True
            ElseIf .Nodes(3).Children <> 0 Then
                Int末级 = 3
                .Nodes(Int末级).Child.Selected = True
                .SelectedItem.Selected = True
            Else
                Int末级 = 0
                .Nodes(1).Selected = True
                .SelectedItem.Selected = True
            End If
            If Int末级 <> 0 Then .Nodes(Int末级).Expanded = True
        End If
    End With

    ReadAndSendDataToTvw = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub GetFilter权限()
    Dim strTemp As String

    With tvwDetails
        If mint状态 = 1 Then
            If InStr(1, mstrPrivs, "管理西成药品种") = 0 Then
                .Nodes.Remove (.Nodes("Root西成药").Index)
            End If
            If InStr(1, mstrPrivs, "管理中成药品种") = 0 Then
                .Nodes.Remove (.Nodes("Root中成药").Index)
            End If
            If InStr(1, mstrPrivs, "管理中草药品种") = 0 Then
                .Nodes.Remove (.Nodes("Root中草药").Index)
            End If
        Else
            If InStr(1, mstrPrivs, "管理西成药规格") = 0 Then
                .Nodes.Remove (.Nodes("Root西成药").Index)
            End If
            If InStr(1, mstrPrivs, "管理中成药规格") = 0 Then
                .Nodes.Remove (.Nodes("Root中成药").Index)
            End If
            If InStr(1, mstrPrivs, "管理中草药规格") = 0 Then
                .Nodes.Remove (.Nodes("Root中草药").Index)
            End If
        End If
    End With
End Sub

Private Sub SetParentNode(ByVal Node As MSComctlLib.Node)
    If Not Node.Parent Is Nothing Then
        If Node.Parent.Parent Is Nothing Then
            mstr药品种类 = Node.Parent
        Else
            Set Node = Node.Parent
            SetParentNode Node
        End If
    End If
End Sub

Private Sub tvwDetails_NodeClick(ByVal Node As MSComctlLib.Node)
    '节点点击事件
    Dim rsRecord As ADODB.Recordset
'    Dim cbrControl1 As CommandBarControl
'    Dim cbrControl2 As CommandBarControl
    Dim lngkey As Long  '用来保存所选中的key值
    Dim str分类 As String   '药品规格修改中用来判断选中的节点是品种还是分类
    Dim intupdate As Integer
    Dim i As Integer
    Dim j As Integer
    Dim bln修改 As Boolean  '用来记录是否有值被修改了

    If Node Is Nothing Then
        Exit Sub
    End If
    mstrNode = Node.Tag '记录节点中的值
    
    Call SetParentNode(Node) '取得药品分类
    
    mblnClick = False
    chk显示所有已修改的药品.Value = 0

    On Error GoTo ErrHandle
    If Node.Tag Like "中草药*" And mint状态 = 2 Then
        vsfDetails.ColComboList(mSpecColumn.规格_住院分零使用) = "0-可以分零|1-不可分零"
        vsfDetails.ColComboList(mSpecColumn.规格_门诊分零使用) = "0-可以分零|1-不可分零"
    ElseIf mint状态 = 2 Then
        vsfDetails.ColComboList(mSpecColumn.规格_住院分零使用) = "0-可以分零|1-不可分零|2-一次性使用|3-分零后一天内有效|4-分零后两天内有效|5-分零后三天内有效"
        vsfDetails.ColComboList(mSpecColumn.规格_门诊分零使用) = "0-可以分零|1-不可分零|2-一次性使用|3-分零后一天内有效|4-分零后两天内有效|5-分零后三天内有效"
    End If
    If Node.Key Like "Root*" Then Exit Sub  '如果选择的节点时最顶级节点则退出

    '判断界面中是否有值刚被修改了
    bln修改 = Check修改

    If bln修改 = True Then
        intupdate = MsgBox("已修改内容还未保存，执行该操作后数据将恢复原状，" & vbCrLf & vbCrLf & "是否继续？", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName)
        If intupdate = vbNo Then Exit Sub
    End If
    
    Call FS.ShowFlash("正在加载数据,请稍候 ...", Me)
    Me.MousePointer = vbHourglass

    If mint状态 = 1 Then    '品种
        gstrSql = "Select Distinct a.id as ID,c.id as 类别id,a.参考目录ID, '['||c.编码||']'||c.名称 as 名称, a.编码, a.名称 As 通用名称, d.英文名称, d.拼音码, d.五笔码, e.毒理分类, e.价值分类, e.货源情况, e.用药梯次, nvl(e.药品类型,0) as 药品类型, e.药品剂型 as 剂型, nvl(e.急救药否,0)  as 急救药," & _
                        "  e.是否肿瘤药 as 肿瘤药, e.溶媒, e.ATCCODE, e.是否原研药 as 原研药, e.是否专利药 as 专利药, e.是否单独定价 as 单独定价, nvl(e.是否新药,0) as 新药, nvl(e.是否原料,0) as 原料药,Nvl(e.是否辅助用药, 0) as 辅助用药, f.名称 As 参考项目, nvl(e.处方职务,'00') as 处方职务, nvl(e.处方限量,0) as 处方限量, Nvl(a.适用性别,0) AS 适用性别, a.计算单位 As 剂量单位, nvl(e.是否皮试,0) as 皮试, nvl(e.抗生素,0) as 抗生素 , nvl(e.品种医嘱,0) as 品种下长期医嘱,a.单独应用 as 单味使用" & _
                        "  From 诊疗项目目录 A, 诊疗项目别名 B, 诊疗分类目录 C," & _
                    " (Select n.诊疗项目id, n.名称, n.拼音码, m.五笔码, p.英文名称" & _
                    "  From (Select 诊疗项目id, 名称, 简码 As 拼音码 From 诊疗项目别名 Where 性质 = 1 And 码类 = 1) N," & _
                    "       (Select 诊疗项目id, 名称, 简码 As 五笔码 From 诊疗项目别名 Where 性质 = 1 And 码类 = 2) M," & _
                    "       (Select 诊疗项目id, 名称 As 英文名称 From 诊疗项目别名 Where 性质 = 2) P" & _
                    "  Where n.诊疗项目id = m.诊疗项目id And n.诊疗项目id = p.诊疗项目id) D, 药品特性 E, 诊疗参考目录 F " & _
                    "   Where a.Id = b.诊疗项目id(+) And a.分类id = c.Id And a.Id = d.诊疗项目id(+) And a.Id = e.药名id And a.参考目录ID = f.Id(+) And " & _
                    " a.撤档时间 = To_Date('3000-1-1', 'yyyy-MM-DD') "
                    
        If mbln自管药 = True Then
            gstrSql = gstrSql & " and e.临床自管药=1"
        Else
            gstrSql = gstrSql & " and e.临床自管药 is null"
        End If
        
        If chkAllDetails.Value = 1 Then '当选择了显示所有节点中的数据时
            gstrSql = gstrSql & " and a.分类id in (Select ID From 诊疗分类目录 Where 类型 In (1, 2, 3) Start With ID = [1] Connect By Prior ID = 上级id) order by id"
        Else
            gstrSql = gstrSql & " and a.分类id=[1] order by id"
        End If
    Else    '规格
        str分类 = Node.Tag
        If str分类 Like "*品种" Then '选中的是品种节点
            gstrSql = "Select a.Id as ID, c.药名id, a.编码 As 规格编码, a.规格 as 药品规格 , j.编码 As 品种编码, j.名称 As 通用名称, m.数字码, c.标识码, a.备选码," & _
                              " Decode(n.商品名, Null, p.商品名, n.商品名) 商品名称, a.产地 As 生产商, n.拼音码, p.五笔码, c.药品来源 As 来源分类, d.名称 As 合同单位, c.批准文号, c.注册商标," & _
                              " c.Gmp认证, c.是否常备 as 非常备药, c.是否带量采购 as 带量采购, c.是否易至跌倒 as 易跌倒,a.计算单位 As 售价单位, c.剂量系数, j.计算单位 as 剂量单位, c.住院单位, c.住院包装 as 住院系数, c.门诊单位, c.门诊包装 as 门诊系数, c.药库单位, c.药库包装 as 药库系数, c.高危药品, c.送货单位, c.送货包装, c.申领单位, c.申领阀值," & _
                              " c.中药形态, a.是否变价 As 药价属性, c.指导批发价 As 采购限价, c.扣率 As 采购扣率, c.指导零售价 As 指导售价, c.加成率, c.成本价 as 成本价格," & _
                              " e.现价 As 当前售价, f.名称 As 收入项目,a.病案费目,c.容量, c.药价级别, a.屏蔽费别, a.费用类型 As 医保类型, c.药库分批, c.药房分批, c.招标药品, c.合同单位id," & _
                              " e.收入项目id, c.最大效期 As 保质期, a.说明 As 标识说明, c.发药类型, a.服务对象, c.住院可否分零 as 住院分零使用, c.动态分零 as 住院动态分零, c.门诊可否分零 as 门诊分零使用, c.基本药物, a.站点 As 站点编号,C.ddd值, i.存储温度, i.存储条件,c.是否摆药," & _
                              " i.配药类型, i.是否不予配置 As 不予调配, i.输液注意事项,c.是否零差价管理 as 零差价管理 ,C.本位码,C.原产地" & _
                       " From 收费项目目录 A, (Select 收费细目id, 简码 As 数字码 From 收费项目别名 Where 码类 = 3 And 性质 = 1) M," & _
                            " (Select 收费细目id, 简码 As 拼音码, 名称 As 商品名 From 收费项目别名 Where 码类 = 1 And 性质 = 3) N," & _
                            " (Select 收费细目id, 简码 As 五笔码, 名称 As 商品名 From 收费项目别名 Where 码类 = 2 And 性质 = 3) P, 药品规格 C, 诊疗项目目录 J, 供应商 D, 收费价目 E," & _
                            " 收入项目 F, 输液药品属性 I, 药品特性 B" & _
                       " Where c.药名id = j.Id And j.Id = [1] And a.撤档时间 = To_Date('3000-1-1', 'yyyy-MM-DD') And a.Id = c.药品id And" & _
                             " c.合同单位id = d.Id(+) And e.收费细目id = a.Id And e.收入项目id = f.Id And a.Id = i.药品id(+) And a.Id = m.收费细目id(+) And" & _
                             " a.Id = n.收费细目id(+) And a.Id = p.收费细目id(+)  and (e.终止日期 is null or Sysdate Between e.执行日期 And e.终止日期)" & _
                             GetPriceClassString("E")
            
            If mbln自管药 = True Then
                gstrSql = gstrSql & " and b.药名id=c.药名id and b.临床自管药=1 Order By a.Id"
            Else
                gstrSql = gstrSql & " and b.药名id=c.药名id and b.临床自管药 is null Order By a.Id"
            End If
            
        Else    '选中的是分类节点
            gstrSql = " Select a.Id as ID, c.药名id, a.编码 As 规格编码, a.规格 as 药品规格, j.编码 As 品种编码, j.名称 As 通用名称, m.数字码, c.标识码, a.备选码," & _
                              " Decode(n.商品名, Null, p.商品名, n.商品名) 商品名称, a.产地 As 生产商, n.拼音码, p.五笔码, c.药品来源 As 来源分类, d.名称 As 合同单位, c.批准文号, c.注册商标, " & _
                              " c.Gmp认证, c.是否常备 as 非常备药, c.是否带量采购 as 带量采购, c.是否易至跌倒 as 易跌倒,a.计算单位 As 售价单位, c.剂量系数, j.计算单位 as 剂量单位, c.住院单位, c.住院包装 as 住院系数, c.门诊单位, c.门诊包装 as 门诊系数, c.药库单位, c.药库包装 as 药库系数, c.高危药品, c.送货单位, c.送货包装, c.申领单位, c.申领阀值," & _
                              " c.中药形态, a.是否变价 As 药价属性, c.指导批发价 As 采购限价, c.扣率 As 采购扣率, c.指导零售价 As 指导售价, c.加成率, c.成本价 as 成本价格," & _
                              " e.现价 As 当前售价, f.名称 As 收入项目,a.病案费目, c.容量,c.药价级别, a.屏蔽费别, a.费用类型 As 医保类型, c.药库分批, c.药房分批, c.招标药品, 合同单位id," & _
                              " e.收入项目id, c.最大效期 As 保质期, a.说明 As 标识说明, c.发药类型, a.服务对象, c.住院可否分零  as 住院分零使用, c.动态分零 as 住院动态分零,c.门诊可否分零 as 门诊分零使用, c.基本药物, a.站点 As 站点编号,c.DDD值, i.存储温度, i.存储条件,c.是否摆药," & _
                              " i.配药类型, i.是否不予配置 As 不予调配, i.输液注意事项,c.是否零差价管理 as 零差价管理 ,C.本位码,C.原产地" & _
                       " From 收费项目目录 A, (Select 收费细目id, 简码 As 数字码 From 收费项目别名 Where 码类 = 3 And 性质 = 1) M," & _
                            " (Select 收费细目id, 简码 As 拼音码, 名称 As 商品名 From 收费项目别名 Where 码类 = 1 And 性质 = 3) N," & _
                            " (Select 收费细目id, 简码 As 五笔码, 名称 As 商品名 From 收费项目别名 Where 码类 = 2 And 性质 = 3) P, 药品规格 C, 供应商 D, 收费价目 E, 收入项目 F," & _
                            " 输液药品属性 I, 诊疗项目目录 J, 药品特性 B" & _
                       " Where a.Id In" & _
                            "  (Select 药品id" & _
                              " From 药品规格" & _
                              " Where 药名id In " & _
                                    " (Select ID " & _
                                    "  From 诊疗项目目录 " & _
                                     " Where 分类id In " & _
                                          "  (Select ID From 诊疗分类目录 Where 类型 In (1, 2, 3) Start With ID = [1] Connect By Prior ID = 上级id))) And" & _
                             " a.撤档时间 = To_Date('3000-1-1', 'yyyy-MM-DD')" & _
                             " And a.Id = c.药品id And c.合同单位id = d.Id(+) And e.收费细目id = a.Id And e.收入项目id = f.Id And a.Id = i.药品id(+) And" & _
                             " c.药名id = j.Id And a.Id = m.收费细目id(+) And a.Id = n.收费细目id(+) And a.Id = p.收费细目id(+) and (e.终止日期 is null or Sysdate Between e.执行日期 And e.终止日期)" & _
                             GetPriceClassString("E")
            
            If mbln自管药 = True Then
                gstrSql = gstrSql & " and b.药名id=c.药名id and b.临床自管药=1 Order By j.名称,a.Id"
            Else
                gstrSql = gstrSql & " and b.药名id=c.药名id and b.临床自管药 is null Order By j.名称,a.Id"
            End If
            
        End If
        Call setColumn(tbcDetails.Selected.Index)
        If chkAllDetails.Value = 0 Then '不能获取到下级节点
            If Node.Tag Like "*分类" Then
                vsfDetails.Rows = 1
                Me.MousePointer = vbDefault
                Call FS.StopFlash
                Exit Sub
            End If
        End If
    End If

    If mint状态 = 2 Then '规格
        If Node.Tag Like "中草药*" Then  '是否显示配药属性
            tbcDetails.Item(mSpecList.配药属性).Visible = False

            With vsfDetails
                .ColHidden(mSpecColumn.规格_存储温度) = True
                .ColHidden(mSpecColumn.规格_存储条件) = True
                .ColHidden(mSpecColumn.规格_配药类型) = True
                .ColHidden(mSpecColumn.规格_不予调配) = True
                .ColHidden(mSpecColumn.规格_输液注意事项) = True
'                If tbcDetails.Selected.Index = tbcDetails.ItemCount - 1 Then
'                    tbcDetails.Item(mSpecList.基本信息).Selected = True
'                End If
            End With
        Else
            If mint配置中心 > 0 Then
                tbcDetails.Item(mSpecList.配药属性).Visible = True
                With vsfDetails
                    If tbcDetails.Item(mSpecList.配药属性).Selected = True Then
                        .ColHidden(mSpecColumn.规格_存储温度) = False
                        .ColHidden(mSpecColumn.规格_存储条件) = False
                        .ColHidden(mSpecColumn.规格_配药类型) = False
                        .ColHidden(mSpecColumn.规格_不予调配) = False
                        .ColHidden(mSpecColumn.规格_输液注意事项) = False
                    Else
                        .ColHidden(mSpecColumn.规格_存储温度) = True
                        .ColHidden(mSpecColumn.规格_存储条件) = True
                        .ColHidden(mSpecColumn.规格_配药类型) = True
                        .ColHidden(mSpecColumn.规格_不予调配) = True
                        .ColHidden(mSpecColumn.规格_输液注意事项) = True
                    End If
                End With
            End If
        End If
    End If
    '获取key值
    lngkey = Mid(Node.Key, InStr(1, Node.Key, "_") + 1, Len(Node.Key) - InStr(1, Node.Key, "_"))
    Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "节点点击", lngkey)

    vsfDetails.Rows = 1
    If rsRecord.EOF Then
        Call setColumn(tbcDetails.Selected.Index)
        Me.MousePointer = vbDefault
        Call FS.StopFlash
        Exit Sub
    End If
    Set mrsRecord = rsRecord.Clone  '克隆
    
    mstrChangedCell = ""
    mintPos = 0
'    Set cbrControl1 = cbsMain.FindControl(, mcon过滤)
'    cbrControl1.Enabled = False
'    Set cbrControl2 = cbsMain.FindControl(, mcon定位)
'    cbrControl2.Enabled = False
    
    Set mrsMyRecords = New ADODB.Recordset
    
    Call showColumn(rsRecord, Node.Tag)   '将值绑定到vsflexgrid控件中
    Call setColumn(tbcDetails.Selected.Index)
    Call GetDefineSize(rsRecord)
    With vsfDetails
        If .Rows > 1 Then
'            cbrControl1.Enabled = True
'            cbrControl2.Enabled = True
            .Row = 1
            .Col = 5
        End If
    End With
    Call Set权限判断
    
    Me.MousePointer = vbDefault
    Call FS.StopFlash
    
    Exit Sub
ErrHandle:
    Me.MousePointer = vbDefault
    Call FS.StopFlash
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub getNewData()
    Dim rsTemp As Recordset
    Dim str分类  As String
    Dim lngkey As Long
    
    On Error GoTo ErrHandle
    
    If mint状态 = 1 Then    '品种
        gstrSql = "Select Distinct a.id as ID,c.id as 类别id,a.参考目录ID, '['||c.编码||']'||c.名称 as 名称, a.编码, a.名称 As 通用名称, d.英文名称, d.拼音码, d.五笔码, e.毒理分类, e.价值分类, e.货源情况, e.用药梯次, e.药品类型, e.药品剂型 as 剂型, e.急救药否 as 急救药," & _
                        "  e.是否肿瘤药 as 肿瘤药, e.溶媒, e.ATCCODE, e.是否原研药 as 原研药, e.是否专利药 as 专利药, e.是否单独定价 as 单独定价, e.是否新药 as 新药, e.是否原料 as 原料药, e.是否辅助用药 as 辅助用药, f.名称 As 参考项目, e.处方职务, e.处方限量, a.适用性别, a.计算单位 As 剂量单位, e.是否皮试 as 皮试, e.抗生素, e.品种医嘱 as 品种下长期医嘱,a.单独应用 as 单味使用" & _
                        "  From 诊疗项目目录 A, 诊疗项目别名 B, 诊疗分类目录 C," & _
                    " (Select n.诊疗项目id, n.名称, n.拼音码, m.五笔码, p.英文名称" & _
                    "  From (Select 诊疗项目id, 名称, 简码 As 拼音码 From 诊疗项目别名 Where 性质 = 1 And 码类 = 1) N," & _
                    "       (Select 诊疗项目id, 名称, 简码 As 五笔码 From 诊疗项目别名 Where 性质 = 1 And 码类 = 2) M," & _
                    "       (Select 诊疗项目id, 名称 As 英文名称 From 诊疗项目别名 Where 性质 = 2) P" & _
                    "  Where n.诊疗项目id = m.诊疗项目id And n.诊疗项目id = p.诊疗项目id) D, 药品特性 E, 诊疗参考目录 F " & _
                    "   Where a.Id = b.诊疗项目id(+) And a.分类id = c.Id And a.Id = d.诊疗项目id(+) And a.Id = e.药名id And a.参考目录ID = f.Id(+) And " & _
                    " a.撤档时间 = To_Date('3000-1-1', 'yyyy-MM-DD') "
                    
        If mbln自管药 = True Then
            gstrSql = gstrSql & " and e.临床自管药=1"
        Else
            gstrSql = gstrSql & " and e.临床自管药 is null"
        End If
        
        If chkAllDetails.Value = 1 Then '当选择了显示所有节点中的数据时
            gstrSql = gstrSql & " and a.分类id in (Select ID From 诊疗分类目录 Where 类型 In (1, 2, 3) Start With ID = [1] Connect By Prior ID = 上级id) order by id"
        Else
            gstrSql = gstrSql & " and a.分类id=[1] order by id"
        End If
    Else    '规格
        str分类 = tvwDetails.SelectedItem.Tag
        If str分类 Like "*品种" Then '选中的是品种节点
            gstrSql = "Select a.Id as ID, c.药名id, a.编码 As 规格编码, a.规格 as 药品规格 , j.编码 As 品种编码, j.名称 As 通用名称, m.数字码, c.标识码, a.备选码," & _
                              " Decode(n.商品名, Null, p.商品名, n.商品名) 商品名称, a.产地 As 生产商, n.拼音码, p.五笔码, c.药品来源 As 来源分类, d.名称 As 合同单位, c.批准文号, c.注册商标," & _
                              " c.Gmp认证, c.是否常备  as 非常备药, c.是否带量采购 as 带量采购, c.是否易至跌倒 as 易跌倒,a.计算单位 As 售价单位, c.剂量系数, j.计算单位 as 剂量单位, c.住院单位, c.住院包装 as 住院系数, c.门诊单位, c.门诊包装 as 门诊系数, c.药库单位, c.药库包装 as 药库系数, c.高危药品, c.送货单位, c.送货包装, c.申领单位, c.申领阀值,c.是否摆药," & _
                              " c.中药形态, a.是否变价 As 药价属性, c.指导批发价 As 采购限价, c.扣率 As 采购扣率, c.指导零售价 As 指导售价, c.加成率, c.成本价  as 成本价格," & _
                              " e.现价 As 当前售价, f.名称 As 收入项目,a.病案费目,c.容量, c.药价级别, a.屏蔽费别, a.费用类型 As 医保类型, c.药库分批, c.药房分批, c.招标药品, c.合同单位id," & _
                              " e.收入项目id, c.最大效期 As 保质期, a.说明 As 标识说明, c.发药类型, a.服务对象, c.住院可否分零 as 住院分零使用, c.动态分零 as 住院动态分零, c.门诊可否分零 as 门诊分零使用, c.基本药物, a.站点 As 站点编号,C.ddd值, i.存储温度, i.存储条件," & _
                              " i.配药类型, i.是否不予配置 As 不予调配, i.输液注意事项,c.是否零差价管理 as 零差价管理,C.本位码 " & _
                       " From 收费项目目录 A, (Select 收费细目id, 简码 As 数字码 From 收费项目别名 Where 码类 = 3 And 性质 = 1) M," & _
                            " (Select 收费细目id, 简码 As 拼音码, 名称 As 商品名 From 收费项目别名 Where 码类 = 1 And 性质 = 3) N," & _
                            " (Select 收费细目id, 简码 As 五笔码, 名称 As 商品名 From 收费项目别名 Where 码类 = 2 And 性质 = 3) P, 药品规格 C, 诊疗项目目录 J, 供应商 D, 收费价目 E," & _
                            " 收入项目 F, 输液药品属性 I, 药品特性 B" & _
                       " Where c.药名id = j.Id And j.Id = [1] And a.撤档时间 = To_Date('3000-1-1', 'yyyy-MM-DD') And a.Id = c.药品id And" & _
                             " c.合同单位id = d.Id(+) And e.收费细目id = a.Id And e.收入项目id = f.Id And a.Id = i.药品id(+) And a.Id = m.收费细目id(+) And" & _
                             " a.Id = n.收费细目id(+) And a.Id = p.收费细目id(+)  and (e.终止日期 is null or Sysdate Between e.执行日期 And e.终止日期)" & _
                             GetPriceClassString("E")
            
            If mbln自管药 = True Then
                gstrSql = gstrSql & " and b.药名id=c.药名id and b.临床自管药=1 Order By a.Id"
            Else
                gstrSql = gstrSql & " and b.药名id=c.药名id and b.临床自管药 is null Order By a.Id"
            End If
            
        Else    '选中的是分类节点
            gstrSql = " Select a.Id as ID, c.药名id, a.编码 As 规格编码, a.规格 as 药品规格, j.编码 As 品种编码, j.名称 As 通用名称, m.数字码, c.标识码, a.备选码," & _
                              " Decode(n.商品名, Null, p.商品名, n.商品名) 商品名称, a.产地 As 生产商, n.拼音码, p.五笔码, c.药品来源 As 来源分类, d.名称 As 合同单位, c.批准文号, c.注册商标, " & _
                              " c.Gmp认证, c.是否常备  as 非常备药, c.是否带量采购 as 带量采购, c.是否易至跌倒 as 易跌倒,a.计算单位 As 售价单位, c.剂量系数, j.计算单位 as 剂量单位, c.住院单位, c.住院包装 as 住院系数, c.门诊单位, c.门诊包装 as 门诊系数, c.药库单位, c.药库包装 as 药库系数, c.高危药品, c.送货单位, c.送货包装, c.申领单位, c.申领阀值,c.是否摆药," & _
                              " c.中药形态, a.是否变价 As 药价属性, c.指导批发价 As 采购限价, c.扣率 As 采购扣率, c.指导零售价 As 指导售价, c.加成率, c.成本价  as 成本价格," & _
                              " e.现价 As 当前售价, f.名称 As 收入项目,a.病案费目, c.容量,c.药价级别, a.屏蔽费别, a.费用类型 As 医保类型, c.药库分批, c.药房分批, c.招标药品, 合同单位id," & _
                              " e.收入项目id, c.最大效期 As 保质期, a.说明 As 标识说明, c.发药类型, a.服务对象, c.住院可否分零 as 住院分零使用, c.动态分零 as 住院动态分零, c.门诊可否分零 as 门诊分零使用,c.基本药物, a.站点 As 站点编号,c.DDD值, i.存储温度, i.存储条件," & _
                              " i.配药类型, i.是否不予配置 As 不予调配, i.输液注意事项,c.是否零差价管理 as 零差价管理 ,C.本位码" & _
                       " From 收费项目目录 A, (Select 收费细目id, 简码 As 数字码 From 收费项目别名 Where 码类 = 3 And 性质 = 1) M," & _
                            " (Select 收费细目id, 简码 As 拼音码, 名称 As 商品名 From 收费项目别名 Where 码类 = 1 And 性质 = 3) N," & _
                            " (Select 收费细目id, 简码 As 五笔码, 名称 As 商品名 From 收费项目别名 Where 码类 = 2 And 性质 = 3) P, 药品规格 C, 供应商 D, 收费价目 E, 收入项目 F," & _
                            " 输液药品属性 I, 诊疗项目目录 J, 药品特性 B" & _
                       " Where a.Id In" & _
                            "  (Select 药品id" & _
                              " From 药品规格" & _
                              " Where 药名id In " & _
                                    " (Select ID " & _
                                    "  From 诊疗项目目录 " & _
                                     " Where 分类id In " & _
                                          "  (Select ID From 诊疗分类目录 Where 类型 In (1, 2, 3) Start With ID = [1] Connect By Prior ID = 上级id))) And" & _
                             " a.撤档时间 = To_Date('3000-1-1', 'yyyy-MM-DD')" & _
                             " And a.Id = c.药品id And c.合同单位id = d.Id(+) And e.收费细目id = a.Id And e.收入项目id = f.Id And a.Id = i.药品id(+) And" & _
                             " c.药名id = j.Id And a.Id = m.收费细目id(+) And a.Id = n.收费细目id(+) And a.Id = p.收费细目id(+) and (e.终止日期 is null or Sysdate Between e.执行日期 And e.终止日期)" & _
                             GetPriceClassString("E")
            
            If mbln自管药 = True Then
                gstrSql = gstrSql & " and b.药名id=c.药名id and b.临床自管药=1 Order By j.名称,a.Id"
            Else
                gstrSql = gstrSql & " and b.药名id=c.药名id and b.临床自管药 is null Order By j.名称,a.Id"
            End If

        End If
    End If
    
    '获取key值
    lngkey = Mid(tvwDetails.SelectedItem.Key, InStr(1, tvwDetails.SelectedItem.Key, "_") + 1, Len(tvwDetails.SelectedItem.Key) - InStr(1, tvwDetails.SelectedItem.Key, "_"))
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "节点点击", lngkey)
    Set mrsRecord = rsTemp
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub showColumn(ByVal rsRecord As ADODB.Recordset, ByVal str分类 As String)
    '当点击树节点时，将值绑定到vsflexgrid控件中
    Dim i As Integer
    Dim j As Integer
    Dim strTemp As String
    Dim intTemp As Integer
    Dim bln剂量系数 As Boolean

    vsfDetails.Rows = rsRecord.RecordCount + 1 '根据查询出来的值的数量来确定列表行数
    
    vsfDetails.Select 1, 1
    If mint状态 = 1 Then    '品种
        For i = 1 To rsRecord.RecordCount
            With vsfDetails
                .TextMatrix(i, mVaricolumn.品种_序号) = i
                .TextMatrix(i, mVaricolumn.品种_id) = IIf(IsNull(rsRecord!ID), "", rsRecord!ID)
                .TextMatrix(i, mVaricolumn.品种_分类id) = IIf(IsNull(rsRecord!类别id), "", rsRecord!类别id)
                .TextMatrix(i, mVaricolumn.品种_药品分类) = IIf(IsNull(rsRecord!名称), "", rsRecord!名称)
                .TextMatrix(i, mVaricolumn.品种_药品编码) = IIf(IsNull(rsRecord!编码), "", rsRecord!编码)
                .TextMatrix(i, mVaricolumn.品种_通用名称) = IIf(IsNull(rsRecord!通用名称), "", rsRecord!通用名称)
                .TextMatrix(i, mVaricolumn.品种_英文名称) = IIf(IsNull(rsRecord!英文名称), "", rsRecord!英文名称)
                .TextMatrix(i, mVaricolumn.品种_拼音码) = IIf(IsNull(rsRecord!拼音码), "", rsRecord!拼音码)
                .TextMatrix(i, mVaricolumn.品种_五笔码) = IIf(IsNull(rsRecord!五笔码), "", rsRecord!五笔码)

                If .TextMatrix(i, mVaricolumn.品种_拼音码) = "" Then
                    .TextMatrix(i, mVaricolumn.品种_拼音码) = zlGetSymbol(.TextMatrix(i, mVaricolumn.品种_通用名称), 0, 30)
                End If

                If .TextMatrix(i, mVaricolumn.品种_五笔码) = "" Then
                    .TextMatrix(i, mVaricolumn.品种_五笔码) = zlGetSymbol(.TextMatrix(i, mVaricolumn.品种_通用名称), 1, 30)
                End If

                .TextMatrix(i, mVaricolumn.品种_毒理分类) = ShowValue(.ColComboList(mVaricolumn.品种_毒理分类), IIf(IsNull(rsRecord!毒理分类), "", rsRecord!毒理分类))
                .TextMatrix(i, mVaricolumn.品种_价值分类) = ShowValue(.ColComboList(mVaricolumn.品种_价值分类), IIf(IsNull(rsRecord!价值分类), "", rsRecord!价值分类))
                .TextMatrix(i, mVaricolumn.品种_货源情况) = ShowValue(.ColComboList(mVaricolumn.品种_货源情况), IIf(IsNull(rsRecord!货源情况), "", rsRecord!货源情况))
                .TextMatrix(i, mVaricolumn.品种_用药梯次) = ShowValue(.ColComboList(mVaricolumn.品种_用药梯次), IIf(IsNull(rsRecord!用药梯次), "", rsRecord!用药梯次))
                .TextMatrix(i, mVaricolumn.品种_药品类型) = ShowValue(.ColComboList(mVaricolumn.品种_药品类型), IIf(IsNull(rsRecord!药品类型), "", rsRecord!药品类型))
                .TextMatrix(i, mVaricolumn.品种_剂型) = ShowValue(.ColComboList(mVaricolumn.品种_剂型), IIf(IsNull(rsRecord!剂型), "", rsRecord!剂型))
                .TextMatrix(i, mVaricolumn.品种_原研药) = IIf(Nvl(rsRecord!原研药, 0) = 0, "", "√")
                .TextMatrix(i, mVaricolumn.品种_专利药) = IIf(Nvl(rsRecord!专利药, 0) = 0, "", "√")
                .TextMatrix(i, mVaricolumn.品种_单独定价) = IIf(Nvl(rsRecord!单独定价, 0) = 0, "", "√")
                .TextMatrix(i, mVaricolumn.品种_急救药) = IIf(Nvl(rsRecord!急救药, 0) = 0, "", "√")
                .TextMatrix(i, mVaricolumn.品种_新药) = IIf(Nvl(rsRecord!新药, 0) = 0, "", "√")
                .TextMatrix(i, mVaricolumn.品种_原料药) = IIf(Nvl(rsRecord!原料药, 0) = 0, "", "√")
                .TextMatrix(i, mVaricolumn.品种_辅助用药) = IIf(Nvl(rsRecord!辅助用药, 0) = 0, "", "√")
                .TextMatrix(i, mVaricolumn.品种_单味使用) = IIf(Nvl(rsRecord!单味使用, 0) = 0, "", "√")
                .TextMatrix(i, mVaricolumn.品种_肿瘤药) = IIf(Nvl(rsRecord!肿瘤药, 0) = 0, "", "√")
                .TextMatrix(i, mVaricolumn.品种_溶媒) = IIf(Nvl(rsRecord!溶媒, 0) = 0, "", "√")
                .TextMatrix(i, mVaricolumn.品种_ATCCODE) = IIf(IsNull(rsRecord!ATCCODE), "", rsRecord!ATCCODE)
                .TextMatrix(i, mVaricolumn.品种_参考项目) = IIf(IsNull(rsRecord!参考项目), "", rsRecord!参考项目)
                .TextMatrix(i, mVaricolumn.品种_处方职务) = ShowValue(.ColComboList(mVaricolumn.品种_处方职务), IIf(IsNull(Mid(rsRecord!处方职务, 1, 1)), "", Mid(rsRecord!处方职务, 1, 1)))
                .TextMatrix(i, mVaricolumn.品种_医保职务) = ShowValue(.ColComboList(mVaricolumn.品种_医保职务), IIf(IsNull(Mid(rsRecord!处方职务, 2, 1)), "", Mid(rsRecord!处方职务, 2, 1)))
                .TextMatrix(i, mVaricolumn.品种_处方限量) = FormatEx(IIf(IsNull(rsRecord!处方限量), "", rsRecord!处方限量), 5)
                .TextMatrix(i, mVaricolumn.品种_适用性别) = ShowValue(.ColComboList(mVaricolumn.品种_适用性别), IIf(IsNull(rsRecord!适用性别), "0", rsRecord!适用性别))
                .TextMatrix(i, mVaricolumn.品种_剂量单位) = IIf(IsNull(rsRecord!剂量单位), "", rsRecord!剂量单位)
                .TextMatrix(i, mVaricolumn.品种_皮试) = IIf(Nvl(rsRecord!皮试, 0) = 0, "", "√")
                .TextMatrix(i, mVaricolumn.品种_抗生素) = ShowValue(.ColComboList(mVaricolumn.品种_抗生素), IIf(IsNull(rsRecord!抗生素), "", rsRecord!抗生素))
                .TextMatrix(i, mVaricolumn.品种_品种下长期医嘱) = IIf(Nvl(rsRecord!品种下长期医嘱, 0) = 0, "", "√")
                .TextMatrix(i, mVaricolumn.品种_参考项目ID) = IIf(IsNull(rsRecord!参考目录ID), "", rsRecord!参考目录ID)
            End With
            rsRecord.MoveNext
        Next
        vsfDetails.Cell(flexcpBackColor, 1, mVaricolumn.品种_药品编码, vsfDetails.Rows - 1) = mlngColor    '设置不可编辑列的背景颜色为灰色
        vsfDetails.Cell(flexcpBackColor, 1, mVaricolumn.品种_药品分类, vsfDetails.Rows - 1) = mlngColor     '设置不可编辑列的背景颜色为灰色
        
        vsfDetails.MergeCol(mVaricolumn.品种_药品分类) = True  '相同列中 药品分类相同合并
    Else    '规格
        For i = 1 To rsRecord.RecordCount
            With vsfDetails
                .TextMatrix(i, mSpecColumn.规格_序号) = i
                .TextMatrix(i, mSpecColumn.规格_id) = IIf(IsNull(rsRecord!ID), "", rsRecord!ID)
                .TextMatrix(i, mSpecColumn.规格_药名id) = IIf(IsNull(rsRecord!药名ID), "", rsRecord!药名ID)
                .TextMatrix(i, mSpecColumn.规格_通用名称) = IIf(IsNull(rsRecord!通用名称), "", rsRecord!通用名称)
                .TextMatrix(i, mSpecColumn.规格_规格编码) = IIf(IsNull(rsRecord!规格编码), "", rsRecord!规格编码)
                .TextMatrix(i, mSpecColumn.规格_药品规格) = IIf(IsNull(rsRecord!药品规格), "", rsRecord!药品规格)
                .TextMatrix(i, mSpecColumn.规格_本位码) = IIf(IsNull(rsRecord!本位码), "", rsRecord!本位码)
                .TextMatrix(i, mSpecColumn.规格_数字码) = IIf(IsNull(rsRecord!数字码), "", rsRecord!数字码)

                If .TextMatrix(i, mSpecColumn.规格_数字码) = "" And .TextMatrix(i, mSpecColumn.规格_药品规格) <> "" Then
                    .TextMatrix(i, mSpecColumn.规格_数字码) = zlGetDigitSign(rsRecord!药名ID, rsRecord!药品规格)
                End If

                .TextMatrix(i, mSpecColumn.规格_标识码) = IIf(IsNull(rsRecord!标识码), "", rsRecord!标识码)
                .TextMatrix(i, mSpecColumn.规格_备选码) = IIf(IsNull(rsRecord!备选码), "", rsRecord!备选码)
                .TextMatrix(i, mSpecColumn.规格_容量) = FormatEx(IIf(IsNull(rsRecord!容量), "", rsRecord!容量), 5)
                .TextMatrix(i, mSpecColumn.规格_商品名称) = IIf(IsNull(rsRecord!商品名称), "", rsRecord!商品名称)
                .TextMatrix(i, mSpecColumn.规格_生产商) = IIf(IsNull(rsRecord!生产商), "", rsRecord!生产商)
                .TextMatrix(i, mSpecColumn.规格_原产地) = IIf(IsNull(rsRecord!原产地), "", rsRecord!原产地)
                .TextMatrix(i, mSpecColumn.规格_来源分类) = ShowValue(.ColComboList(mSpecColumn.规格_来源分类), IIf(IsNull(rsRecord!来源分类), "", rsRecord!来源分类))
                .TextMatrix(i, mSpecColumn.规格_拼音码) = IIf(IsNull(rsRecord!拼音码), "", rsRecord!拼音码)
                .TextMatrix(i, mSpecColumn.规格_五笔码) = IIf(IsNull(rsRecord!五笔码), "", rsRecord!五笔码)

                If .TextMatrix(i, mSpecColumn.规格_商品名称) <> "" And .TextMatrix(i, mSpecColumn.规格_拼音码) = "" Then
                    .TextMatrix(i, mSpecColumn.规格_拼音码) = zlGetSymbol(.TextMatrix(i, mSpecColumn.规格_通用名称), 0, 30)
                End If

                If .TextMatrix(i, mSpecColumn.规格_商品名称) <> "" And .TextMatrix(i, mSpecColumn.规格_五笔码) = "" Then
                    .TextMatrix(i, mSpecColumn.规格_五笔码) = zlGetSymbol(.TextMatrix(i, mSpecColumn.规格_通用名称), 1, 30)
                End If

                .TextMatrix(i, mSpecColumn.规格_合同单位) = IIf(IsNull(rsRecord!合同单位), "", rsRecord!合同单位)
                .TextMatrix(i, mSpecColumn.规格_批准文号) = IIf(IsNull(rsRecord!批准文号), "", rsRecord!批准文号)

                .TextMatrix(i, mSpecColumn.规格_注册商标) = IIf(IsNull(rsRecord!注册商标), "", rsRecord!注册商标)
                .TextMatrix(i, mSpecColumn.规格_GMP认证) = IIf(Nvl(rsRecord!GMP认证, 0) = 0, "", "√")
                .TextMatrix(i, mSpecColumn.规格_非常备药) = IIf(Nvl(rsRecord!非常备药, 0) = 0, "", "√")
                .TextMatrix(i, mSpecColumn.规格_带量采购) = IIf(Nvl(rsRecord!带量采购, 0) = 0, "", "√")
                .TextMatrix(i, mSpecColumn.规格_易跌倒) = IIf(Nvl(rsRecord!易跌倒, 0) = 0, "", "√")
                .TextMatrix(i, mSpecColumn.规格_售价单位) = IIf(IsNull(rsRecord!售价单位), "", rsRecord!售价单位)
                .TextMatrix(i, mSpecColumn.规格_剂量系数) = FormatEx(IIf(IsNull(rsRecord!剂量系数), "", rsRecord!剂量系数), 5)
                .TextMatrix(i, mSpecColumn.规格_剂量单位) = IIf(IsNull(rsRecord!剂量单位), "", rsRecord!剂量单位)
                .TextMatrix(i, mSpecColumn.规格_住院单位) = IIf(IsNull(rsRecord!住院单位), "", rsRecord!住院单位)
                .TextMatrix(i, mSpecColumn.规格_住院系数) = FormatEx(IIf(IsNull(rsRecord!住院系数), "", rsRecord!住院系数), 5)
                .TextMatrix(i, mSpecColumn.规格_门诊单位) = IIf(IsNull(rsRecord!门诊单位), "", rsRecord!门诊单位)
                .TextMatrix(i, mSpecColumn.规格_门诊系数) = FormatEx(IIf(IsNull(rsRecord!门诊系数), "", rsRecord!门诊系数), 5)
                .TextMatrix(i, mSpecColumn.规格_药库单位) = IIf(IsNull(rsRecord!药库单位), "", rsRecord!药库单位)

                .TextMatrix(i, mSpecColumn.规格_药价属性) = ShowValue(.ColComboList(mSpecColumn.规格_药价属性), IIf(IsNull(rsRecord!药价属性), "", rsRecord!药价属性))
                .TextMatrix(i, mSpecColumn.规格_零差价管理) = IIf(Nvl(rsRecord!零差价管理, 0) = 0, "", "√")
                
                .TextMatrix(i, mSpecColumn.规格_药库系数) = FormatEx(IIf(IsNull(rsRecord!药库系数), "", rsRecord!药库系数), 5)
                .TextMatrix(i, mSpecColumn.规格_送货单位) = IIf(IsNull(rsRecord!送货单位), "", rsRecord!送货单位)
                .TextMatrix(i, mSpecColumn.规格_送货包装) = FormatEx(IIf(IsNull(rsRecord!送货包装), "", rsRecord!送货包装), 5)
                Select Case rsRecord!中药形态
                    Case "0"
                        strTemp = "0-散装"
                    Case "1"
                        strTemp = "1-中药饮片"
                    Case Else
                        strTemp = "2-免煎剂"
                End Select

                .TextMatrix(i, mSpecColumn.规格_中药形态) = strTemp

                Select Case rsRecord!申领单位
                    Case "1"
                        strTemp = "1-售价单位"
                    Case "2"
                        strTemp = "2-住院单位"
                    Case "3"
                        strTemp = "3-门诊单位"
                    Case "4"
                        strTemp = "4-药库单位"
                    Case Else
                        strTemp = "1-售价单位"
                End Select
                .TextMatrix(i, mSpecColumn.规格_申领单位) = strTemp

                Select Case Nvl(rsRecord!申领单位, 1)
                    Case 1 '零售
                        .TextMatrix(i, mSpecColumn.规格_申领阀值) = Format(Nvl(rsRecord!申领阀值, 0), "#0.00;-#0.00; ;")
                    Case 2 '住院
                        .TextMatrix(i, mSpecColumn.规格_申领阀值) = Format(Nvl(rsRecord!申领阀值, 0) / Nvl(rsRecord!住院系数, 1), "#0.00;-#0.00; ;")
                    Case 3 '门诊
                        .TextMatrix(i, mSpecColumn.规格_申领阀值) = Format(Nvl(rsRecord!申领阀值, 0) / Nvl(rsRecord!门诊系数, 1), "#0.00;-#0.00; ;")
                    Case 4 '药库
                        .TextMatrix(i, mSpecColumn.规格_申领阀值) = Format(Nvl(rsRecord!申领阀值, 0) / Nvl(rsRecord!药库系数, 1), "#0.00;-#0.00; ;")
                End Select

                If mint当前单位 <> 0 Then
                    .TextMatrix(i, mSpecColumn.规格_采购限价) = FormatEx(IIf(IsNull(rsRecord!采购限价), 0, rsRecord!采购限价) * .TextMatrix(i, mSpecColumn.规格_药库系数), mintCostDigit)
                    .TextMatrix(i, mSpecColumn.规格_指导售价) = FormatEx(IIf(IsNull(rsRecord!指导售价), 0, rsRecord!指导售价) * .TextMatrix(i, mSpecColumn.规格_药库系数), mintPriceDigit)
                    .TextMatrix(i, mSpecColumn.规格_成本价格) = FormatEx(IIf(IsNull(rsRecord!成本价格), "", rsRecord!成本价格) * .TextMatrix(i, mSpecColumn.规格_药库系数), mintCostDigit)
                Else
                    .TextMatrix(i, mSpecColumn.规格_采购限价) = FormatEx(IIf(IsNull(rsRecord!采购限价), 0, rsRecord!采购限价), mintCostDigit)
                    .TextMatrix(i, mSpecColumn.规格_指导售价) = FormatEx(IIf(IsNull(rsRecord!指导售价), 0, rsRecord!指导售价), mintPriceDigit)
                    .TextMatrix(i, mSpecColumn.规格_成本价格) = FormatEx(IIf(IsNull(rsRecord!成本价格), "", rsRecord!成本价格), mintCostDigit)
                End If

                .TextMatrix(i, mSpecColumn.规格_采购扣率) = FormatEx(IIf(IsNull(rsRecord!采购扣率), "", rsRecord!采购扣率), 5)
                .TextMatrix(i, mSpecColumn.规格_结算价) = FormatEx(.TextMatrix(i, mSpecColumn.规格_采购限价) * (.TextMatrix(i, mSpecColumn.规格_采购扣率) / 100), mintCostDigit)
                .TextMatrix(i, mSpecColumn.规格_加成率) = FormatEx(IIf(IsNull(rsRecord!加成率), "", rsRecord!加成率), 5)

                If mint当前单位 <> 0 Then
                    .TextMatrix(i, mSpecColumn.规格_当前售价) = FormatEx(IIf(IsNull(rsRecord!当前售价), 0, rsRecord!当前售价) * .TextMatrix(i, mSpecColumn.规格_药库系数), mintPriceDigit)
                Else
                    .TextMatrix(i, mSpecColumn.规格_当前售价) = FormatEx(IIf(IsNull(rsRecord!当前售价), 0, rsRecord!当前售价), mintPriceDigit)
                End If
                .TextMatrix(i, mSpecColumn.规格_收入项目) = ShowValue(.ColComboList(mSpecColumn.规格_收入项目), rsRecord!收入项目)
                .TextMatrix(i, mSpecColumn.规格_病案费目) = IIf(IsNull(rsRecord!病案费目), "", rsRecord!病案费目)
                .TextMatrix(i, mSpecColumn.规格_药价级别) = ShowValue(.ColComboList(mSpecColumn.规格_药价级别), IIf(IsNull(rsRecord!药价级别), "", rsRecord!药价级别))
                .TextMatrix(i, mSpecColumn.规格_屏蔽费别) = IIf(Nvl(rsRecord!屏蔽费别, 0) = 0, "", "√")
                .TextMatrix(i, mSpecColumn.规格_医保类型) = ShowValue(.ColComboList(mSpecColumn.规格_医保类型), IIf(IsNull(rsRecord!医保类型), "", rsRecord!医保类型))
                .TextMatrix(i, mSpecColumn.规格_药库分批) = IIf(Nvl(rsRecord!药库分批, 0) = 0, "", "√")
                .TextMatrix(i, mSpecColumn.规格_药房分批) = IIf(Nvl(rsRecord!药房分批, 0) = 0, "", "√")
                .TextMatrix(i, mSpecColumn.规格_保质期) = FormatEx(IIf(Nvl(rsRecord!保质期, 0) = 0, 0, rsRecord!保质期), 5)
                
                .TextMatrix(i, mSpecColumn.规格_标识说明) = IIf(IsNull(rsRecord!标识说明), "", rsRecord!标识说明)
                .TextMatrix(i, mSpecColumn.规格_发药类型) = ShowValue(.ColComboList(mSpecColumn.规格_发药类型), IIf(IsNull(rsRecord!发药类型), "", rsRecord!发药类型))
                .TextMatrix(i, mSpecColumn.规格_站点编号) = ShowValue(.ColComboList(mSpecColumn.规格_站点编号), IIf(IsNull(rsRecord!站点编号), "", rsRecord!站点编号))
                .TextMatrix(i, mSpecColumn.规格_DDD值) = FormatEx(IIf(IsNull(rsRecord!DDD值), "", rsRecord!DDD值), 5)
                .TextMatrix(i, mSpecColumn.规格_服务对象) = ShowValue(.ColComboList(mSpecColumn.规格_服务对象), IIf(IsNull(rsRecord!服务对象), "", rsRecord!服务对象))
                .TextMatrix(i, mSpecColumn.规格_高危药品) = ShowValue(.ColComboList(mSpecColumn.规格_高危药品), IIf(IsNull(rsRecord!高危药品), "", rsRecord!高危药品))
                
                If str分类 Like "中草药*" Then
                    If IsNull(rsRecord!住院分零使用) Or rsRecord!住院分零使用 = 0 Then
                        .TextMatrix(i, mSpecColumn.规格_住院分零使用) = "0-可以分零"
                    Else
                        .TextMatrix(i, mSpecColumn.规格_住院分零使用) = "1-不可分零"
                    End If
                    If IsNull(rsRecord!门诊分零使用) Or rsRecord!门诊分零使用 = 0 Then
                        .TextMatrix(i, mSpecColumn.规格_门诊分零使用) = "0-可以分零"
                    Else
                        .TextMatrix(i, mSpecColumn.规格_门诊分零使用) = "1-不可分零"
                    End If

                    If .TextMatrix(i, mSpecColumn.规格_中药形态) = "0-散装" Then
                        .TextMatrix(i, mSpecColumn.规格_住院分零使用) = "0-可以分零"
                        .Cell(flexcpBackColor, i, mSpecColumn.规格_住院分零使用) = mlngColor
                        .TextMatrix(i, mSpecColumn.规格_门诊分零使用) = "0-可以分零"
                        .Cell(flexcpBackColor, i, mSpecColumn.规格_门诊分零使用) = mlngColor
                    Else
                        .Cell(flexcpBackColor, i, mSpecColumn.规格_住院分零使用) = mlngApplyColor
                        .Cell(flexcpBackColor, i, mSpecColumn.规格_门诊分零使用) = mlngApplyColor
                    End If
                Else
                    If IsNull(rsRecord!住院分零使用) Or rsRecord!住院分零使用 = 0 Then
                        intTemp = 0
                    ElseIf rsRecord!住院分零使用 = 1 Then
                        intTemp = 1
                    ElseIf rsRecord!住院分零使用 = 2 Then
                        intTemp = 2
                    ElseIf rsRecord!住院分零使用 = -1 Then
                        intTemp = 3
                    ElseIf rsRecord!住院分零使用 = -2 Then
                        intTemp = 4
                    ElseIf rsRecord!住院分零使用 = -3 Then
                        intTemp = 5
                    End If
                    .TextMatrix(i, mSpecColumn.规格_住院分零使用) = ShowValue(.ColComboList(mSpecColumn.规格_住院分零使用), IIf(IsNull(rsRecord!住院分零使用), "", intTemp))

                    If IsNull(rsRecord!门诊分零使用) Or rsRecord!门诊分零使用 = 0 Then
                        intTemp = 0
                    ElseIf rsRecord!门诊分零使用 = 1 Then
                        intTemp = 1
                    ElseIf rsRecord!门诊分零使用 = 2 Then
                        intTemp = 2
                    ElseIf rsRecord!门诊分零使用 = -1 Then
                        intTemp = 3
                    ElseIf rsRecord!门诊分零使用 = -2 Then
                        intTemp = 4
                    ElseIf rsRecord!门诊分零使用 = -3 Then
                        intTemp = 5
                    End If
                    .TextMatrix(i, mSpecColumn.规格_门诊分零使用) = ShowValue(.ColComboList(mSpecColumn.规格_门诊分零使用), IIf(IsNull(rsRecord!门诊分零使用), "", intTemp))
                End If
                .TextMatrix(i, mSpecColumn.规格_基本药物) = ShowValue(.ColComboList(mSpecColumn.规格_基本药物), IIf(IsNull(rsRecord!基本药物), "", rsRecord!基本药物))
                .TextMatrix(i, mSpecColumn.规格_住院动态分零) = IIf(Nvl(rsRecord!住院动态分零, 0) = 0, "", "√")
                .TextMatrix(i, mSpecColumn.规格_存储温度) = ShowValue(.ColComboList(mSpecColumn.规格_存储温度), IIf(IsNull(rsRecord!存储温度), "", rsRecord!存储温度))
                .TextMatrix(i, mSpecColumn.规格_存储条件) = IIf(Nvl(rsRecord!存储条件, 0) = 0, "", "√")
                .TextMatrix(i, mSpecColumn.规格_是否摆药) = IIf(Nvl(rsRecord!是否摆药, 0) = 0, "否", "是")

                .TextMatrix(i, mSpecColumn.规格_配药类型) = ShowValue(.ColComboList(mSpecColumn.规格_配药类型), IIf(IsNull(rsRecord!配药类型), "", rsRecord!配药类型))
                .TextMatrix(i, mSpecColumn.规格_不予调配) = IIf(Nvl(rsRecord!不予调配, 0), "", "√")
                .TextMatrix(i, mSpecColumn.规格_输液注意事项) = IIf(IsNull(rsRecord!输液注意事项), "", rsRecord!输液注意事项)
                .TextMatrix(i, mSpecColumn.规格_招标药品) = IIf(IsNull(rsRecord!招标药品), 0, rsRecord!招标药品)
                .TextMatrix(i, mSpecColumn.规格_合同单位id) = IIf(IsNull(rsRecord!合同单位id), "", rsRecord!合同单位id)
                .TextMatrix(i, mSpecColumn.规格_收入项目id) = IIf(IsNull(rsRecord!收入项目id), "", rsRecord!收入项目id)
                
                Call ShowDisplay(rsRecord, i)
                Call CheckValue(i, rsRecord!ID)
            End With
            rsRecord.MoveNext
        Next
        vsfDetails.MergeCol(mSpecColumn.规格_通用名称) = True   '合并通用名称
        With vsfDetails
            .Cell(flexcpBackColor, 1, mSpecColumn.规格_规格编码, .Rows - 1) = mlngColor
            .Cell(flexcpBackColor, 1, mSpecColumn.规格_通用名称, .Rows - 1) = mlngColor
            .Cell(flexcpBackColor, 1, mSpecColumn.规格_剂量单位, .Rows - 1) = mlngColor
        End With
    End If

    Call Recover    '将修改了的颜色改变回来

    '调整行高
    With vsfDetails
        For i = 1 To .Rows - 1
            .RowHeight(i) = 350
        Next
    End With
End Sub

Private Function ShowValue(ByVal strValue As String, ByVal strBiJiao As String) As String
    '功能 ：通过传入的值比较返回所获取的值
    '参数 strvalue 原字符串
    'strBiJiao 需要比较的字符串
    Dim arr As Variant
    Dim i As Integer

    If strValue = "" Then Exit Function
    ReDim arr(UBound(Split(strValue, "|"))) As String   '重新定义数组长度

    '将值分解开来保存到数组中
    For i = 0 To UBound(Split(strValue, "|"))
        arr(i) = Split(strValue, "|")(i)
    Next
    If strBiJiao = "" Then
        ShowValue = ""
        Exit Function
    End If

    '循环比较
    For i = 0 To UBound(Split(strValue, "|"))
        If InStr(1, arr(i), strBiJiao) > 0 Then
            ShowValue = arr(i)
            Exit Function
        End If
    Next
End Function

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call FindGridRow(UCase(txtFind))
        txtFind.SetFocus
    End If
End Sub

Private Sub vsfDetails_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim i As Integer
    Dim j As Integer
    Dim rsRecord As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim strPro As String
    Dim strSql As String
    Dim intAttr As Integer
    Dim strSQLItem As String
    Dim intupdate As Integer
    Dim dblLeft As Double
    Dim dblTop As Double
    Dim strTemp As String

    On Error GoTo ErrHandle
    With vsfDetails
        If .Cell(flexcpBackColor, NewRow, NewCol) = mlngColor Then
            .FocusRect = flexFocusLight
        Else
            .FocusRect = flexFocusSolid
        End If

        If .Row < OldRow Then
            OldRow = 1
        End If

        If .Rows = 1 Then
           OldRow = 0
        End If
        .TextMatrix(OldRow, OldCol) = Trim(.TextMatrix(OldRow, OldCol))
    End With

    '控制菜单中应用于所有列显示与否
    With vsfDetails
        If mint状态 = 1 Then '品种
            Select Case NewCol
                Case mVaricolumn.品种_通用名称, mVaricolumn.品种_英文名称, mVaricolumn.品种_拼音码, mVaricolumn.品种_五笔码
                    mcbrToolBar.Controls(1).Enabled = False
                    mobjPopup.Controls(1).Enabled = False
                Case Else
                    If .Cell(flexcpBackColor, NewRow, NewCol, NewRow, NewCol) = mlngColor Then
                        mcbrToolBar.Controls(1).Enabled = False
                        mobjPopup.Controls(1).Enabled = False
                    Else
                        mcbrToolBar.Controls(1).Enabled = True
                        mobjPopup.Controls(1).Enabled = True
                    End If
            End Select

            Select Case OldCol
                Case mVaricolumn.品种_通用名称, mVaricolumn.品种_剂量单位
                    If vsfDetails.TextMatrix(OldRow, OldCol) = "" Then
                        MsgBox "该单元格内容不能为空，请输入！", vbInformation, gstrSysName
                        vsfDetails.Select OldRow, OldCol
                    End If
                Case mVaricolumn.品种_参考项目
                    Dim iAttr As Integer

                    If .TextMatrix(OldRow, OldCol) = "" Or .TextMatrix(OldRow, OldCol) = "参考项目" Then Exit Sub
                    
                    vRect = zlControl.GetControlRect(vsfDetails.hwnd) '获取位置
                    dblLeft = vRect.Left + vsfDetails.CellLeft
                    dblTop = vRect.Top + vsfDetails.CellTop + vsfDetails.CellHeight + 3200

                    strSql = "Select 类型 From 诊疗分类目录 Where ID=[1]"
                    Set rsRecord = zlDatabase.OpenSQLRecord(strSql, Me.Caption, .TextMatrix(OldRow, mVaricolumn.品种_分类id))

                    If rsRecord.EOF Then
                        iAttr = -1
                    Else
                        iAttr = rsRecord(0)
                    End If
                    strSql = "Select Distinct a.Id, a.分类id, a.编码, a.名称, a.说明" & _
                             "   From 诊疗参考目录 A, 诊疗参考别名 B " & _
                             "   Where a.ID = b.参考目录ID And a.类型 = [1] And (Upper(a.编码) Like [2] Or Upper(a.名称) Like [3] Or Upper(b.名称) Like [3] Or Upper(b.简码) Like [3]) " & _
                             "   Order By 编码 "
                    Set rsRecord = zlDatabase.ShowSQLSelect(Me, strSql, 0, "诊疗项目", False, "", "", False, False, _
                    True, dblLeft, dblTop, vsfDetails.Height, blnCancel, False, True, iAttr, UCase(.TextMatrix(OldRow, OldCol)) & "%", mstrMatch & UCase(.TextMatrix(OldRow, OldCol)) & "%")

                    If rsRecord Is Nothing Then
                        .TextMatrix(OldRow, OldCol) = ""
                        .TextMatrix(OldRow, mVaricolumn.品种_参考项目ID) = ""
                        Exit Sub
                    End If
                    .EditText = rsRecord!名称
                    .TextMatrix(OldRow, mVaricolumn.品种_参考项目) = rsRecord!名称
                    .TextMatrix(OldRow, mVaricolumn.品种_参考项目ID) = rsRecord!ID
                Case mVaricolumn.品种_处方限量
                    .TextMatrix(OldRow, mVaricolumn.品种_处方限量) = FormatEx(.TextMatrix(OldRow, mVaricolumn.品种_处方限量), 5)
            End Select
        Else    '规格
            Select Case NewCol
                Case mSpecColumn.规格_商品名称, mSpecColumn.规格_拼音码, mSpecColumn.规格_五笔码, mSpecColumn.规格_药品规格, mSpecColumn.规格_备选码, mSpecColumn.规格_标识码, mSpecColumn.规格_数字码, mSpecColumn.规格_本位码
                    mcbrToolBar.Controls(1).Enabled = False
                    mobjPopup.Controls(1).Enabled = False
                Case Else
                    If .Cell(flexcpBackColor, NewRow, NewCol, NewRow, NewCol) = mlngColor Then
                        mcbrToolBar.Controls(1).Enabled = False
                        mobjPopup.Controls(1).Enabled = False
                    Else
                        mcbrToolBar.Controls(1).Enabled = True
                        mobjPopup.Controls(1).Enabled = True
                    End If
            End Select

            Select Case OldCol
                Case mSpecColumn.规格_剂量系数, mSpecColumn.规格_住院系数, mSpecColumn.规格_门诊系数, mSpecColumn.规格_药库系数, mSpecColumn.规格_送货包装, mSpecColumn.规格_申领阀值, mSpecColumn.规格_采购扣率, mSpecColumn.规格_加成率, mSpecColumn.规格_容量, mSpecColumn.规格_保质期, mSpecColumn.规格_DDD值
                    If IsNumeric(.TextMatrix(OldRow, OldCol)) Then
                        .TextMatrix(OldRow, OldCol) = FormatEx(.TextMatrix(OldRow, OldCol), 5)
                    End If
                Case mSpecColumn.规格_成本价格, mSpecColumn.规格_采购限价, mSpecColumn.规格_结算价
                    If IsNumeric(.TextMatrix(OldRow, OldCol)) Then
                        .TextMatrix(OldRow, OldCol) = FormatEx(.TextMatrix(OldRow, OldCol), mintCostDigit)
                    End If
                Case mSpecColumn.规格_当前售价, mSpecColumn.规格_指导售价
                    If IsNumeric(.TextMatrix(OldRow, OldCol)) Then
                        .TextMatrix(OldRow, OldCol) = FormatEx(.TextMatrix(OldRow, OldCol), mintPriceDigit)
                    End If
                Case mSpecColumn.规格_收入项目
                    If OldRow <> 0 Then
                        If .TextMatrix(OldRow, OldCol) <> "" Then
                            strTemp = Mid(.TextMatrix(OldRow, OldCol), 2, InStr(1, .TextMatrix(OldRow, OldCol), "]") - 2)
                        End If
                        gstrSql = "Select ID" & _
                                  "  From 收入项目" & _
                                  "  Where 编码=[1] and 末级 = 1 And (撤档时间 Is Null Or 撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))"

                        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, "收入项目查询", strTemp)
                        If rsTmp.RecordCount > 0 Then
                            .TextMatrix(OldRow, mSpecColumn.规格_收入项目id) = rsTmp!ID
                        End If
                    End If
            End Select
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsfDetails_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '控制哪些列可以编辑，那些列不可以编辑,当背景颜色为灰色的列都不允许修改
    With vsfDetails
        If .CellBackColor <> mlngColor And mblnClick = True And Row = mintRow And .Rows <> 1 Then
            mstrOldValue = ""
            mrsRecord.Filter = ""
            mrsRecord.Filter = "ID=" & Val(.TextMatrix(Row, 1))
            mrsMyRecords.Filter = ""
            mrsMyRecords.Filter = "ID=" & Val(.TextMatrix(Row, 1))
            
            If Not mrsRecord.EOF Then
                If mint状态 = 1 Then '品种
                    If Col = mVaricolumn.品种_处方职务 Then
                        mstrOldValue = Mid(mrsRecord.Fields(.TextMatrix(0, mVaricolumn.品种_处方职务)), 1, 1)
                    ElseIf Col = mVaricolumn.品种_医保职务 Then
                        mstrOldValue = Mid(mrsRecord.Fields(.TextMatrix(0, mVaricolumn.品种_处方职务)), 2, 1)
                    ElseIf Col = mVaricolumn.品种_处方限量 Then
                        mstrOldValue = FormatEx(mrsRecord.Fields(.TextMatrix(0, Col)), 7)
                    Else
                        mstrOldValue = IIf(IsNull(mrsRecord.Fields(.TextMatrix(0, Col))), "", mrsRecord.Fields(.TextMatrix(0, Col)))
                    End If
                Else '规格
                    If Col = mSpecColumn.规格_申领单位 Then
                        mstrOldValue = IIf(IsNull(mrsRecord.Fields(.TextMatrix(0, Col))), 1, mrsRecord.Fields(.TextMatrix(0, Col)))
                    ElseIf Col = mSpecColumn.规格_中药形态 Then
                        mstrOldValue = Nvl(mrsRecord.Fields(.TextMatrix(0, Col)), 0) & "," & IIf(Nvl(mrsRecord.Fields(.TextMatrix(0, mSpecColumn.规格_住院分零使用)), 0) <> 0, 1, 0) & "," & IIf(Nvl(mrsRecord.Fields(.TextMatrix(0, mSpecColumn.规格_门诊分零使用)), 0) <> 0, 1, 0)
                    ElseIf Col = mSpecColumn.规格_住院单位 Then
                        mstrOldValue = IIf(IsNull(mrsRecord.Fields("住院单位")), 1, mrsRecord.Fields("住院单位"))
                    ElseIf Col = mSpecColumn.规格_住院系数 Then
                        mstrOldValue = FormatEx(IIf(IsNumeric(mrsRecord.Fields("住院系数")), mrsRecord.Fields("住院系数"), ""), 7)
                    ElseIf Col = mSpecColumn.规格_住院分零使用 Or Col = mSpecColumn.规格_门诊分零使用 Then
                        If tvwDetails.SelectedItem.Tag Like "中草药*" Then
                            mstrOldValue = IIf(Nvl(mrsRecord.Fields(.TextMatrix(0, Col)), 0) <> 0, 1, 0)
                        Else
                            mstrOldValue = Nvl(mrsRecord.Fields(.TextMatrix(0, Col)), 0)
                        End If
                    ElseIf Col = mSpecColumn.规格_药库分批 Then
                        mstrOldValue = Nvl(mrsRecord.Fields(.TextMatrix(0, mSpecColumn.规格_保质期)), 0)
                    ElseIf Col = mSpecColumn.规格_高危药品 Then
                        mstrOldValue = Nvl(mrsRecord.Fields(.TextMatrix(0, Col)), 0)
                    ElseIf Col = mSpecColumn.规格_存储库房 Then
                        mstrOldValue = IIf(IsNull(mrsMyRecords.Fields(.TextMatrix(0, Col))), "", mrsMyRecords.Fields(.TextMatrix(0, Col)))
                    ElseIf Col = mSpecColumn.规格_服务科室 Then
                        mstrOldValue = IIf(IsNull(mrsMyRecords.Fields(.TextMatrix(0, Col))), "", mrsMyRecords.Fields(.TextMatrix(0, Col)))
                    Else
                        If IsNumeric(mrsRecord.Fields(.TextMatrix(0, Col))) Then
                            mstrOldValue = FormatEx(mrsRecord.Fields(.TextMatrix(0, Col)), 7)
                        Else
                            mstrOldValue = IIf(IsNull(mrsRecord.Fields(.TextMatrix(0, Col))), "", mrsRecord.Fields(.TextMatrix(0, Col)))
                        End If
                    End If
                End If
            End If
        End If
        
        If .Cell(flexcpBackColor, Row, Col) = mlngColor Then
            Cancel = True
        End If
        
        If mint状态 = 1 Then '品种
            Select Case .Col
                Case mVaricolumn.品种_肿瘤药, mVaricolumn.品种_溶媒, mVaricolumn.品种_原研药, mVaricolumn.品种_专利药, mVaricolumn.品种_单独定价, mVaricolumn.品种_急救药, mVaricolumn.品种_新药, mVaricolumn.品种_原料药, mVaricolumn.品种_辅助用药, mVaricolumn.品种_品种下长期医嘱, mVaricolumn.品种_皮试, mVaricolumn.品种_单味使用, mVaricolumn.品种_拼音码, mVaricolumn.品种_五笔码
                    Cancel = True
            End Select
        Else
            Select Case .Col
                Case mSpecColumn.规格_屏蔽费别, mSpecColumn.规格_住院动态分零, mSpecColumn.规格_GMP认证, mSpecColumn.规格_非常备药, mSpecColumn.规格_带量采购, mSpecColumn.规格_易跌倒, mSpecColumn.规格_药库分批, _
                    mSpecColumn.规格_药房分批, mSpecColumn.规格_存储条件, mSpecColumn.规格_不予调配, mSpecColumn.规格_拼音码, mSpecColumn.规格_五笔码, _
                    mSpecColumn.规格_零差价管理
                    Cancel = True
            End Select
        End If
    End With
End Sub

Private Sub vsfDetails_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsRecord As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim strPro As String
    Dim strSql As String
    Dim intAttr As Integer
    Dim strSQLItem As String
    Dim dblLeft As Double
    Dim dblTop As Double

    vRect = zlControl.GetControlRect(vsfDetails.hwnd) '获取位置
    dblLeft = vRect.Left + vsfDetails.CellLeft
    dblTop = vRect.Top + vsfDetails.CellTop + vsfDetails.CellHeight + 3200
    On Error GoTo ErrHandle
    With vsfDetails
        If mint状态 = 1 Then    '品种
            If Col = mVaricolumn.品种_参考项目 Then
                strSql = "Select 类型 From 诊疗分类目录 Where ID=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, .TextMatrix(.Row, mVaricolumn.品种_分类id))

                If rsTmp.EOF Then
                    intAttr = -1
                Else
                    intAttr = rsTmp!类型
                End If

                strSql = " Select ID,分类ID,编码,名称,说明 From 诊疗参考目录 a Where 类型=[1] Order By 编码"

                Set rsRecord = zlDatabase.ShowSQLSelect(Me, strSql, 0, "诊疗项目", False, "", "", False, False, _
                True, dblLeft, dblTop, vsfDetails.Height, blnCancel, False, True, intAttr)

                If rsRecord Is Nothing Then
                    Exit Sub
                End If
                .TextMatrix(.Row, mVaricolumn.品种_参考项目) = rsRecord!名称
                .TextMatrix(.Row, mVaricolumn.品种_参考项目ID) = rsRecord!ID
            End If
        Else    '规格
            Select Case Col
                Case mSpecColumn.规格_生产商
                    strSql = "Select 编码 as id,名称,简码 From 药品生产商 Order By 编码 "

                    Set rsRecord = zlDatabase.ShowSQLSelect(Me, strSql, 0, "诊疗项目", False, "", "", False, False, _
                        True, dblLeft, dblTop, vsfDetails.Height, blnCancel, False, True)

                    If rsRecord Is Nothing Then
                        Exit Sub
                    Else
                        .TextMatrix(.Row, mSpecColumn.规格_生产商) = rsRecord!名称
                    End If
                Case mSpecColumn.规格_原产地
                    strSql = "Select 编码 as id,名称,简码 From 药品生产商 Order By 编码 "

                    Set rsRecord = zlDatabase.ShowSQLSelect(Me, strSql, 0, "诊疗项目", False, "", "", False, False, _
                        True, dblLeft, dblTop, vsfDetails.Height, blnCancel, False, True)

                    If rsRecord Is Nothing Then
                        Exit Sub
                    Else
                        .TextMatrix(.Row, mSpecColumn.规格_原产地) = rsRecord!名称
                    End If
                Case mSpecColumn.规格_合同单位
                    strSql = "Select id,编码,名称,简码" & _
                                " From 供应商" & _
                                " where 末级=1 And substr(类型,1,1) = '1' And (撤档时间 is null or 撤档时间=to_date('3000-01-01','YYYY-MM-DD')) " & _
                                " Order By 编码 "
                    Set rsRecord = zlDatabase.ShowSQLSelect(Me, strSql, 0, "诊疗项目", False, "", "", False, False, _
                        True, dblLeft, dblTop, vsfDetails.Height, blnCancel, False, True)

                    If rsRecord Is Nothing Then
                        Exit Sub
                    Else
                        .TextMatrix(.Row, mSpecColumn.规格_合同单位) = rsRecord!名称
                        .TextMatrix(.Row, mSpecColumn.规格_合同单位id) = rsRecord!ID
                    End If
                Case mSpecColumn.规格_病案费目
                    Dim blnRe As Boolean
                    Dim str名称 As String
                    Dim strID As String

                    gstrSql = "Select 编码 as id,上级 as 上级id, 名称, 简码, 末级 From 病案费目 Start With 上级 Is Null Connect By Prior 编码 = 上级"
                    blnRe = frmTreeLeafSel.ShowTree(gstrSql, strID, str名称, "病案费目")
                    '成功返回
                    If blnRe Then
                        .TextMatrix(.Row, mSpecColumn.规格_病案费目) = str名称
                    End If
                Case mSpecColumn.规格_存储库房
                    mint行号 = Row
                    mint列号 = Col

                    Call frmServiceRoom.ShowMe(Me, mstr药品种类, vsfDetails.TextMatrix(Row, Col), mstrPrivs)
                    Call InitDepartment(vsfDetails.TextMatrix(Row, Col), vsfDetails.TextMatrix(Row, Col + 1), vsfDetails.TextMatrix(Row, Col + 2), vsfDetails.TextMatrix(Row, Col + 3))
                Case mSpecColumn.规格_服务科室
                    mint行号 = Row
                    mint列号 = Col
                    Call frmServiceDepartment.ShowMe(Me, vsfDetails.TextMatrix(Row, Col - 2), vsfDetails.TextMatrix(Row, Col - 1), vsfDetails.TextMatrix(Row, Col), vsfDetails.TextMatrix(Row, Col + 1))
            End Select
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsfDetails_CellChanged(ByVal Row As Long, ByVal Col As Long)
'    Dim cbrControl As CommandBarControl
    Dim cbrControlSave As CommandBarControl
    Dim cbrControlLocation As CommandBarControl
    Dim cbrControl恢复默认 As CommandBarControl
    Dim strText As String
    Dim intOld As Integer
    Dim intNum As Integer
    Dim i As Integer
    Dim j As Integer
    
    Set cbrControl恢复默认 = cbsMain.FindControl(, mcon默认值)
    Set cbrControlSave = cbsMain.FindControl(, mcon保存)
    Set cbrControlLocation = cbsMain.FindControl(, mcon定位已修改项目)
    
    With vsfDetails
        If .CellBackColor <> mlngColor And mblnClick = True And Row = mintRow And .Rows <> 1 And Row > 0 Then
            If mint状态 = 1 Then '品种
                Select Case Col
                    Case mVaricolumn.品种_拼音码, mVaricolumn.品种_五笔码
                        strText = .EditText
                    Case mVaricolumn.品种_毒理分类, mVaricolumn.品种_价值分类, mVaricolumn.品种_货源情况, mVaricolumn.品种_用药梯次, mVaricolumn.品种_剂型
                        strText = Mid(.TextMatrix(Row, Col), InStr(1, .TextMatrix(Row, Col), "-") + 1)
                    Case mVaricolumn.品种_药品类型, mVaricolumn.品种_抗生素, mVaricolumn.品种_适用性别, mVaricolumn.品种_处方职务, mVaricolumn.品种_医保职务
                        strText = Mid(.TextMatrix(Row, Col), 1, InStr(1, .TextMatrix(Row, Col), "-") - 1)
                    Case mVaricolumn.品种_肿瘤药, mVaricolumn.品种_溶媒, mVaricolumn.品种_原研药, mVaricolumn.品种_专利药, mVaricolumn.品种_单独定价, mVaricolumn.品种_急救药, mVaricolumn.品种_新药, mVaricolumn.品种_原料药, mVaricolumn.品种_辅助用药, mVaricolumn.品种_品种下长期医嘱, mVaricolumn.品种_皮试, mVaricolumn.品种_单味使用
                        strText = "√"
                    Case mVaricolumn.品种_参考项目ID
                        strText = mstrOldValue
                    Case mVaricolumn.品种_处方限量
                        strText = FormatEx(.TextMatrix(Row, Col), 7)
                    Case Else
                        strText = .TextMatrix(Row, Col)
                End Select
            Else '规格
                Select Case Col
                    Case mSpecColumn.规格_拼音码, mSpecColumn.规格_五笔码
                        strText = .EditText
                    Case mSpecColumn.规格_来源分类, mSpecColumn.规格_药价级别, mSpecColumn.规格_医保类型
                        strText = Mid(.TextMatrix(Row, Col), InStr(1, .TextMatrix(Row, Col), "-") + 1)
                    Case mSpecColumn.规格_GMP认证, mSpecColumn.规格_非常备药, mSpecColumn.规格_带量采购, mSpecColumn.规格_易跌倒, mSpecColumn.规格_屏蔽费别, mSpecColumn.规格_住院动态分零, mSpecColumn.规格_药库分批, mSpecColumn.规格_药房分批, mSpecColumn.规格_存储条件, mSpecColumn.规格_不予调配, mSpecColumn.规格_零差价管理
                        strText = "√"
                        
                        If Col = mSpecColumn.规格_零差价管理 And .TextMatrix(Row, Col) = "√" Then
                            If Val(zlDatabase.GetPara(275, glngSys, , 0)) > 0 Then
                                If .Cell(flexcpBackColor, Row, mSpecColumn.规格_成本价格, Row, mSpecColumn.规格_成本价格) = mlngColor Then
                                    '当前价格不能调整时，表示有库存
                                    If CheckPriceAdjust(Val(.TextMatrix(Row, mSpecColumn.规格_id)), 0, -1, True) = False Then
                                        MsgBox "该药品已启用零差价管理，但售价和成本价不一致，请注意调价！", vbInformation, gstrSysName
                                    End If
                                Else
                                    '能调整价格时，直接比较价格
                                    If Val(.TextMatrix(Row, mSpecColumn.规格_成本价格)) <> Val(.TextMatrix(Row, mSpecColumn.规格_当前售价)) Then
                                        MsgBox "该药品已启用零差价管理，但售价和成本价不一致，请重新录入价格！", vbInformation, gstrSysName
                                    End If
                                End If
                            End If
                        End If
                    Case mSpecColumn.规格_合同单位id
                        strText = mstrOldValue
                    Case mSpecColumn.规格_申领单位, mSpecColumn.规格_药价属性, mSpecColumn.规格_站点编号, mSpecColumn.规格_服务对象
                        strText = Mid(.TextMatrix(Row, Col), 1, InStr(1, .TextMatrix(Row, Col), "-") - 1)
                    Case mSpecColumn.规格_中药形态
                        strText = Mid(.TextMatrix(Row, Col), 1, InStr(1, .TextMatrix(Row, Col), "-") - 1)
                        If strText <> 0 Then
                            MsgBox "你修改了“中药形态”，系统将强制设定“临床应用”页中分零使用为“不可分零”！", vbInformation, gstrSysName
                            .Cell(flexcpBackColor, .Row, mSpecColumn.规格_住院分零使用) = cstcolor_backcolor
                            .TextMatrix(.Row, mSpecColumn.规格_住院分零使用) = "1-不可分零"
                            .Cell(flexcpBackColor, .Row, mSpecColumn.规格_门诊分零使用) = cstcolor_backcolor
                            .TextMatrix(.Row, mSpecColumn.规格_门诊分零使用) = "1-不可分零"
                        Else
                            .Cell(flexcpBackColor, .Row, mSpecColumn.规格_住院分零使用) = mlngColor
                            .Cell(flexcpBackColor, .Row, mSpecColumn.规格_门诊分零使用) = mlngColor
                            .TextMatrix(.Row, mSpecColumn.规格_住院分零使用) = "0-可以分零"
                            .TextMatrix(.Row, mSpecColumn.规格_门诊分零使用) = "0-可以分零"
                        End If
                        If Mid(.TextMatrix(Row, mSpecColumn.规格_住院分零使用), 1, InStr(1, .TextMatrix(Row, mSpecColumn.规格_住院分零使用), "-") - 1) = Split(mstrOldValue, ",")(1) Then
                            .Cell(flexcpForeColor, Row, mSpecColumn.规格_住院分零使用) = vbBack: .Cell(flexcpFontSize, Row, mSpecColumn.规格_住院分零使用) = 9: .Cell(flexcpFontBold, Row, mSpecColumn.规格_住院分零使用) = False
                        Else
                            .Cell(flexcpForeColor, Row, mSpecColumn.规格_住院分零使用) = mlngApplyColor: .Cell(flexcpFontSize, Row, mSpecColumn.规格_住院分零使用) = 10: .Cell(flexcpFontBold, Row, mSpecColumn.规格_住院分零使用) = True
                        End If
                        If Mid(.TextMatrix(Row, mSpecColumn.规格_门诊分零使用), 1, InStr(1, .TextMatrix(Row, mSpecColumn.规格_门诊分零使用), "-") - 1) = Split(mstrOldValue, ",")(2) Then
                            .Cell(flexcpForeColor, Row, mSpecColumn.规格_门诊分零使用) = vbBack: .Cell(flexcpFontSize, Row, mSpecColumn.规格_门诊分零使用) = 9: .Cell(flexcpFontBold, Row, mSpecColumn.规格_门诊分零使用) = False
                        Else
                            .Cell(flexcpForeColor, Row, mSpecColumn.规格_门诊分零使用) = mlngApplyColor: .Cell(flexcpFontSize, Row, mSpecColumn.规格_门诊分零使用) = 10: .Cell(flexcpFontBold, Row, mSpecColumn.规格_门诊分零使用) = True
                        End If
                        mstrOldValue = Split(mstrOldValue, ",")(0)
                    Case mSpecColumn.规格_住院分零使用, mSpecColumn.规格_门诊分零使用
                        strText = Mid(.TextMatrix(Row, Col), 1, InStr(1, .TextMatrix(Row, Col), "-") - 1)
                        If Not tvwDetails.SelectedItem.Tag Like "中草药*" Then
                            strText = Switch(strText = 0, 0, strText = 1, 1, strText = 2, 2, strText = 3, -1, strText = 4, -2, strText = 5, -3)
                        End If
                    Case mSpecColumn.规格_采购限价
                        strText = FormatEx(.TextMatrix(Row, Col), mintCostDigit)
                        If mstrOldValue = strText Then
                            .Cell(flexcpForeColor, Row, mSpecColumn.规格_结算价) = vbBack: .Cell(flexcpFontSize, Row, mSpecColumn.规格_结算价) = 9: .Cell(flexcpFontBold, Row, mSpecColumn.规格_结算价) = False
                        End If
                    Case mSpecColumn.规格_收入项目
                        strText = Mid(.TextMatrix(Row, Col), InStr(1, .TextMatrix(Row, Col), "]") + 1)
                    Case mSpecColumn.规格_高危药品
                        strText = IIf(Trim(.TextMatrix(Row, Col)) = "", "0", .TextMatrix(Row, Col))
                        If strText <> "0" Then
                            strText = Mid(.TextMatrix(Row, Col), 1, InStr(1, .TextMatrix(Row, Col), "-") - 1)
                        End If
                    Case mSpecColumn.规格_剂量系数, mSpecColumn.规格_住院系数, mSpecColumn.规格_门诊系数, mSpecColumn.规格_药库系数, mSpecColumn.规格_送货包装, mSpecColumn.规格_申领阀值, mSpecColumn.规格_采购扣率, mSpecColumn.规格_加成率, mSpecColumn.规格_容量, mSpecColumn.规格_保质期, mSpecColumn.规格_DDD值
                        strText = .TextMatrix(Row, Col)
                        If IsNumeric(strText) Then
                            strText = FormatEx(strText, 5)
                        End If
                    Case mSpecColumn.规格_成本价格, mSpecColumn.规格_结算价
                        strText = .TextMatrix(Row, Col)
                        If IsNumeric(strText) Then
                            strText = FormatEx(strText, mintCostDigit)
                        End If
                    Case mSpecColumn.规格_当前售价, mSpecColumn.规格_指导售价
                        strText = .TextMatrix(Row, Col)
                        If IsNumeric(strText) Then
                            strText = FormatEx(strText, mintPriceDigit)
                        End If
                    Case Else
                        strText = .TextMatrix(Row, Col)
                End Select
            End If
            
            If Trim(mstrOldValue) = Trim(strText) Then
                If .Cell(flexcpForeColor, Row, Col) <> vbBack Then .Cell(flexcpForeColor, Row, Col) = vbBack
                If .Cell(flexcpFontSize, Row, Col) <> 9 Then .Cell(flexcpFontSize, Row, Col) = 9
                If .Cell(flexcpFontBold, Row, Col) = True Then .Cell(flexcpFontBold, Row, Col) = False
            Else
                If mint状态 = 1 Then '品种
                    Select Case Col
                        Case mVaricolumn.品种_肿瘤药, mVaricolumn.品种_溶媒, mVaricolumn.品种_原研药, mVaricolumn.品种_专利药, mVaricolumn.品种_单独定价, mVaricolumn.品种_急救药, mVaricolumn.品种_新药, mVaricolumn.品种_原料药, mVaricolumn.品种_辅助用药, mVaricolumn.品种_品种下长期医嘱, mVaricolumn.品种_皮试, mVaricolumn.品种_单味使用
                            If .Cell(flexcpBackColor, Row, Col) <> mlngApplyColor Then
                                .Cell(flexcpBackColor, Row, Col) = mlngApplyColor
                            Else
                                .Cell(flexcpBackColor, Row, Col) = cstcolor_backcolor
                            End If
                        Case Else
                            .Cell(flexcpForeColor, Row, Col) = mlngApplyColor: .Cell(flexcpFontSize, Row, Col) = 10: .Cell(flexcpFontBold, Row, Col) = True
                    End Select
                Else '规格
                    Select Case Col
                        Case mSpecColumn.规格_药库分批
                            If .TextMatrix(Row, mSpecColumn.规格_药库分批) = "√" Then
                                .Cell(flexcpBackColor, Row, mSpecColumn.规格_药房分批) = cstcolor_backcolor
                                If Not mstrNode Like "中草药*" And .TextMatrix(Row, mSpecColumn.规格_保质期) = 0 Then
                                    .Cell(flexcpBackColor, Row, mSpecColumn.规格_保质期) = cstcolor_backcolor
                                    .TextMatrix(Row, mSpecColumn.规格_保质期) = 24
                                End If
                            Else
                                .Cell(flexcpBackColor, Row, mSpecColumn.规格_药房分批) = mlngColor
                                .TextMatrix(Row, mSpecColumn.规格_药房分批) = ""
                                If Not mstrNode Like "中草药*" Then
                                    .Cell(flexcpBackColor, Row, mSpecColumn.规格_保质期) = mlngColor
                                    .TextMatrix(Row, mSpecColumn.规格_保质期) = 0
                                End If
                            End If
                            If mstrOldValue <> .TextMatrix(Row, mSpecColumn.规格_保质期) Then
                                .Cell(flexcpForeColor, Row, mSpecColumn.规格_保质期) = mlngApplyColor
                                .Cell(flexcpFontBold, Row, mSpecColumn.规格_保质期) = True
                                .Cell(flexcpFontSize, Row, mSpecColumn.规格_保质期) = 10
                            Else
                                .Cell(flexcpForeColor, Row, mSpecColumn.规格_保质期) = vbBlack
                                .Cell(flexcpFontSize, Row, mSpecColumn.规格_保质期) = 9
                                .Cell(flexcpFontBold, Row, mSpecColumn.规格_保质期) = False
                            End If
                            If .Cell(flexcpBackColor, Row, Col) <> mlngApplyColor Then
                                .Cell(flexcpBackColor, Row, Col) = mlngApplyColor
                            Else
                                .Cell(flexcpBackColor, Row, Col) = cstcolor_backcolor
                            End If
                        Case mSpecColumn.规格_屏蔽费别, mSpecColumn.规格_住院动态分零, mSpecColumn.规格_GMP认证, mSpecColumn.规格_非常备药, mSpecColumn.规格_带量采购, mSpecColumn.规格_易跌倒, mSpecColumn.规格_药房分批, mSpecColumn.规格_存储条件, mSpecColumn.规格_不予调配
                            If .Cell(flexcpBackColor, Row, Col) <> mlngApplyColor Then
                                .Cell(flexcpBackColor, Row, Col) = mlngApplyColor
                            Else
                                .Cell(flexcpBackColor, Row, Col) = cstcolor_backcolor
                            End If
                        Case mSpecColumn.规格_存储库房id, mSpecColumn.规格_库房科室id
                        Case Else
                            .Cell(flexcpForeColor, Row, Col) = mlngApplyColor: .Cell(flexcpFontSize, Row, Col) = 10: .Cell(flexcpFontBold, Row, Col) = True
                    End Select
                End If
            End If
            
            cbrControl恢复默认.Enabled = False
            cbrControlSave.Enabled = False
            cbrControlLocation.Enabled = False
            
            For intNum = 0 To tbcDetails.ItemCount - 1
                tbcDetails.Item(intNum).Image = 0
            Next
                
            For i = 1 To .Rows - 1
                For j = 1 To .Cols - 1
                    If .Cell(flexcpForeColor, i, j) = mlngApplyColor Or .Cell(flexcpFontSize, i, j) = 10 Or .Cell(flexcpFontBold, i, j) = True Or .Cell(flexcpBackColor, i, j) = mlngApplyColor Then
                        cbrControl恢复默认.Enabled = True
                        cbrControlSave.Enabled = True
                        cbrControlLocation.Enabled = True
                        If mint状态 = 1 Then '品种
                            Select Case j
                                Case mVaricolumn.品种_英文名称 To mVaricolumn.品种_五笔码
                                    tbcDetails.Item(0).Image = 116
                                Case mVaricolumn.品种_毒理分类 To mVaricolumn.品种_ATCCODE
                                    tbcDetails.Item(1).Image = 116
                                Case mVaricolumn.品种_参考项目 To mVaricolumn.品种_参考项目ID
                                    tbcDetails.Item(2).Image = 116
                                Case mVaricolumn.品种_通用名称
                                    tbcDetails.Item(0).Image = 116
                                    tbcDetails.Item(1).Image = 116
                                    tbcDetails.Item(2).Image = 116
                            End Select
                        Else  '规格
                            Select Case j
                                Case mSpecColumn.规格_本位码 To mSpecColumn.规格_容量
                                    tbcDetails.Item(0).Image = 116
                                Case mSpecColumn.规格_商品名称 To mSpecColumn.规格_非常备药
                                    tbcDetails.Item(1).Image = 116
                                Case mSpecColumn.规格_售价单位 To mSpecColumn.规格_中药形态
                                    tbcDetails.Item(2).Image = 116
                                Case mSpecColumn.规格_药价属性 To mSpecColumn.规格_加成率
                                    tbcDetails.Item(3).Image = 116
                                Case mSpecColumn.规格_收入项目 To mSpecColumn.规格_医保类型
                                    tbcDetails.Item(4).Image = 116
                                Case mSpecColumn.规格_药库分批 To mSpecColumn.规格_保质期
                                    tbcDetails.Item(5).Image = 116
                                Case mSpecColumn.规格_标识说明 To mSpecColumn.规格_高危药品
                                    tbcDetails.Item(6).Image = 116
                                Case mSpecColumn.规格_存储温度 To mSpecColumn.规格_输液注意事项
                                    tbcDetails.Item(7).Image = 116
                                Case mSpecColumn.规格_存储库房 To mSpecColumn.规格_服务科室
                                    tbcDetails.Item(8).Image = 116
                                Case mSpecColumn.规格_药品规格
                                    tbcDetails.Item(0).Image = 116
                                    tbcDetails.Item(1).Image = 116
                                    tbcDetails.Item(2).Image = 116
                                    tbcDetails.Item(3).Image = 116
                                    tbcDetails.Item(4).Image = 116
                                    tbcDetails.Item(5).Image = 116
                                    tbcDetails.Item(6).Image = 116
                                    tbcDetails.Item(7).Image = 116
                                    tbcDetails.Item(8).Image = 116
                            End Select
                        End If
                    End If
                Next
            Next
        End If
        
        If .CellBackColor <> mlngColor And mblnClick = True And Row <> mintRow And .Rows <> 1 And Row > 0 Then
            cbrControl恢复默认.Enabled = True
            cbrControlSave.Enabled = True
            cbrControlLocation.Enabled = True
        End If
        
    End With
End Sub

Private Sub vsfDetails_ChangeEdit()
    Dim lngId As Long
    Dim strTemp As String
    
    mstrChangedCell = ""
    mintPos = 0
    With vsfDetails
        If mint状态 = 1 Then '品种
            Select Case .Col
                Case mVaricolumn.品种_通用名称
                    .TextMatrix(.Row, mVaricolumn.品种_拼音码) = zlGetSymbol(.EditText, 0, 30)
                    .TextMatrix(.Row, mVaricolumn.品种_五笔码) = zlGetSymbol(.EditText, 1, 30)
            End Select
        Else    '规格
            Select Case .Col
                Case mSpecColumn.规格_商品名称
                    .TextMatrix(.Row, mSpecColumn.规格_拼音码) = zlGetSymbol(.EditText, 0, 30)
                    .TextMatrix(.Row, mSpecColumn.规格_五笔码) = zlGetSymbol(.EditText, 1, 30)
                Case mSpecColumn.规格_药品规格
                    lngId = .TextMatrix(.Row, mSpecColumn.规格_id)
                    .TextMatrix(.Row, mSpecColumn.规格_数字码) = zlGetDigitSign(lngId, .EditText)
                    If mstrOldValue <> .EditText Then
                        .Cell(flexcpForeColor, .Row, mSpecColumn.规格_数字码) = mlngApplyColor: .Cell(flexcpFontSize, .Row, mSpecColumn.规格_数字码) = 10: .Cell(flexcpFontBold, .Row, mSpecColumn.规格_数字码) = True
                    End If
                Case mSpecColumn.规格_采购限价
                    .TextMatrix(.Row, mSpecColumn.规格_结算价) = FormatEx(Val(.EditText) * (.TextMatrix(.Row, mSpecColumn.规格_采购扣率) / 100), mintPriceDigit)
                    .Cell(flexcpForeColor, .Row, mSpecColumn.规格_结算价) = mlngApplyColor: .Cell(flexcpFontSize, .Row, mSpecColumn.规格_结算价) = 10: .Cell(flexcpFontBold, .Row, mSpecColumn.规格_结算价) = True
                Case mSpecColumn.规格_住院分零使用
                    If Val(Mid(.EditText, 1, 1)) = 0 Then
                        .Cell(flexcpBackColor, .Row, mSpecColumn.规格_住院动态分零) = mlngColor
                    Else
                        .Cell(flexcpBackColor, .Row, mSpecColumn.规格_住院动态分零) = cstcolor_backcolor
                    End If
            End Select
        End If
    End With
End Sub

Private Sub vsfDetails_Click()
    mblnClick = True
End Sub

Private Sub vsfDetails_DblClick()
    With vsfDetails
        If .Cell(flexcpBackColor, .Row, .Col) <> mlngColor Then
            If mint状态 = 1 Then '品种
                Select Case .Col
                    Case mVaricolumn.品种_肿瘤药, mVaricolumn.品种_溶媒, mVaricolumn.品种_原研药, mVaricolumn.品种_专利药, mVaricolumn.品种_单独定价, mVaricolumn.品种_急救药, mVaricolumn.品种_新药, mVaricolumn.品种_原料药, mVaricolumn.品种_辅助用药, mVaricolumn.品种_品种下长期医嘱, mVaricolumn.品种_皮试, mVaricolumn.品种_单味使用
                        If .TextMatrix(.Row, .Col) = "" Then
                            .TextMatrix(.Row, .Col) = "√"
                        Else
                            .TextMatrix(.Row, .Col) = ""
                        End If
                End Select
            Else
                Select Case .Col
                    Case mSpecColumn.规格_屏蔽费别, mSpecColumn.规格_住院动态分零, mSpecColumn.规格_GMP认证, mSpecColumn.规格_非常备药, mSpecColumn.规格_带量采购, mSpecColumn.规格_易跌倒, mSpecColumn.规格_药库分批, mSpecColumn.规格_药房分批, mSpecColumn.规格_存储条件, mSpecColumn.规格_不予调配, mSpecColumn.规格_零差价管理
                        If .TextMatrix(.Row, .Col) = "" Then
                            .TextMatrix(.Row, .Col) = "√"
                        Else
                            .TextMatrix(.Row, .Col) = ""
                        End If
                End Select
            End If
        End If
    End With
End Sub

Private Sub vsfDetails_EnterCell()
    Dim cbrControl As CommandBarControl
    Dim rsRecord As ADODB.Recordset
    Dim strkey As String
    Dim i As Integer
    Dim j As Integer
    
    With vsfDetails
        If mintRow上次 > .Rows - 1 Then
            mintRow上次 = 1
        End If
    
        If .Rows <> 1 Then
            .Cell(flexcpPicture, mintRow上次, 0, mintRow上次, 0) = Nothing    '设置图片
            For i = 1 To .Rows - 1    '当在切换选项页+排序时会出现多个图片 在这种情况下先将多余的一个清除掉
                If Not .Cell(flexcpPicture, i, 0, i, 0) Is Nothing Then
                    .Cell(flexcpPicture, i, 0, i, 0) = Nothing
                    Exit For
                End If
            Next
            .Cell(flexcpPicture, .Row, 0, .Row, 0) = Me.ImgTvw.ListImages(3).Picture
            
            Call SetBorder '设置行选中边框
        End If
        
        If .Row = mintRow Then Exit Sub
        mintRow = .Row '记录当前行
        strkey = .TextMatrix(.Row, mVaricolumn.品种_id)
        
    End With
End Sub

Private Sub SetBorder()
    '设置行选中边框
    Dim intCol As Integer
    Dim intRow As Integer
    
    With vsfDetails
        If .Rows <> 1 Then
            For intRow = 1 To .Rows - 1
                .CellBorderRange intRow, 0, intRow, .Cols - 1, &HE0E0E0, 0, 0, 0, 0, 0, 0
            Next
            If mint状态 = 1 Then  '品种
                Select Case tbcDetails.Selected.Index
                    Case 0
                        .CellBorderRange .Row, mVaricolumn.品种_药品编码, .Row, mVaricolumn.品种_五笔码, mlngBorderColor, 0, 2, 0, 2, 0, 2
                        .CellBorderRange .Row, mVaricolumn.品种_药品编码, .Row, mVaricolumn.品种_药品编码, mlngBorderColor, 2, 2, 0, 2, 2, 2
                        .CellBorderRange .Row, mVaricolumn.品种_五笔码, .Row, mVaricolumn.品种_五笔码, mlngBorderColor, 0, 2, 2, 2, 2, 2
                    Case 1
                        If mstrNode Like "中草药*" Then
                            intCol = mVaricolumn.品种_辅助用药
                        Else
                            intCol = mVaricolumn.品种_ATCCODE
                        End If
                        .CellBorderRange .Row, mVaricolumn.品种_通用名称, .Row, intCol, mlngBorderColor, 0, 2, 0, 2, 0, 2
                        .CellBorderRange .Row, mVaricolumn.品种_通用名称, .Row, mVaricolumn.品种_通用名称, mlngBorderColor, 2, 2, 0, 2, 2, 2
                        .CellBorderRange .Row, intCol, .Row, intCol, mlngBorderColor, 0, 2, 2, 2, 2, 2
                    Case 2
                        If mstrNode Like "中草药*" Then
                            intCol = mVaricolumn.品种_剂量单位
                        Else
                            intCol = mVaricolumn.品种_品种下长期医嘱
                        End If
                        .CellBorderRange .Row, mVaricolumn.品种_通用名称, .Row, intCol, mlngBorderColor, 0, 2, 0, 2, 0, 2
                        .CellBorderRange .Row, mVaricolumn.品种_通用名称, .Row, mVaricolumn.品种_通用名称, mlngBorderColor, 2, 2, 0, 2, 2, 2
                        .CellBorderRange .Row, intCol, .Row, intCol, mlngBorderColor, 0, 2, 2, 2, 2, 2
                End Select
            Else  '规格
                Select Case tbcDetails.Selected.Index
                    Case 0
                        If mstrNode Like "中草药*" Then
                            intCol = mSpecColumn.规格_备选码
                        Else
                            intCol = mSpecColumn.规格_容量
                        End If
                        .CellBorderRange .Row, mSpecColumn.规格_规格编码, .Row, intCol, mlngBorderColor, 0, 2, 0, 2, 0, 2
                        .CellBorderRange .Row, mSpecColumn.规格_规格编码, .Row, mSpecColumn.规格_规格编码, mlngBorderColor, 2, 2, 0, 2, 2, 2
                        .CellBorderRange .Row, intCol, .Row, intCol, mlngBorderColor, 0, 2, 2, 2, 2, 2
                    Case 1
                        .CellBorderRange .Row, mSpecColumn.规格_药品规格, .Row, mSpecColumn.规格_非常备药, mlngBorderColor, 0, 2, 0, 2, 0, 2
                        .CellBorderRange .Row, mSpecColumn.规格_药品规格, .Row, mSpecColumn.规格_药品规格, mlngBorderColor, 2, 2, 0, 2, 2, 2
                        .CellBorderRange .Row, mSpecColumn.规格_非常备药, .Row, mSpecColumn.规格_非常备药, mlngBorderColor, 0, 2, 2, 2, 2, 2
                    Case 2
                        If mstrNode Like "中草药*" Then
                            intCol = mSpecColumn.规格_中药形态
                        Else
                            intCol = mSpecColumn.规格_申领阀值
                        End If
                        .CellBorderRange .Row, mSpecColumn.规格_药品规格, .Row, intCol, mlngBorderColor, 0, 2, 0, 2, 0, 2
                        .CellBorderRange .Row, mSpecColumn.规格_药品规格, .Row, mSpecColumn.规格_药品规格, mlngBorderColor, 2, 2, 0, 2, 2, 2
                        .CellBorderRange .Row, intCol, .Row, intCol, mlngBorderColor, 0, 2, 2, 2, 2, 2
                    Case 3
                        .CellBorderRange .Row, mSpecColumn.规格_药品规格, .Row, mSpecColumn.规格_加成率, mlngBorderColor, 0, 2, 0, 2, 0, 2
                        .CellBorderRange .Row, mSpecColumn.规格_药品规格, .Row, mSpecColumn.规格_药品规格, mlngBorderColor, 2, 2, 0, 2, 2, 2
                        .CellBorderRange .Row, mSpecColumn.规格_加成率, .Row, mSpecColumn.规格_加成率, mlngBorderColor, 0, 2, 2, 2, 2, 2
                    Case 4
                        .CellBorderRange .Row, mSpecColumn.规格_药品规格, .Row, mSpecColumn.规格_医保类型, mlngBorderColor, 0, 2, 0, 2, 0, 2
                        .CellBorderRange .Row, mSpecColumn.规格_药品规格, .Row, mSpecColumn.规格_药品规格, mlngBorderColor, 2, 2, 0, 2, 2, 2
                        .CellBorderRange .Row, mSpecColumn.规格_医保类型, .Row, mSpecColumn.规格_医保类型, mlngBorderColor, 0, 2, 2, 2, 2, 2
                    Case 5
                        If mstrNode Like "中草药*" Then
                            intCol = mSpecColumn.规格_药房分批
                        Else
                            intCol = mSpecColumn.规格_保质期
                        End If
                        .CellBorderRange .Row, mSpecColumn.规格_药品规格, .Row, intCol, mlngBorderColor, 0, 2, 0, 2, 0, 2
                        .CellBorderRange .Row, mSpecColumn.规格_药品规格, .Row, mSpecColumn.规格_药品规格, mlngBorderColor, 2, 2, 0, 2, 2, 2
                        .CellBorderRange .Row, intCol, .Row, intCol, mlngBorderColor, 0, 2, 2, 2, 2, 2
                    Case 6
                        If mstrNode Like "中草药*" Then
                            intCol = mSpecColumn.规格_门诊分零使用
                        Else
                            intCol = mSpecColumn.规格_高危药品
                        End If
                        .CellBorderRange .Row, mSpecColumn.规格_药品规格, .Row, intCol, mlngBorderColor, 0, 2, 0, 2, 0, 2
                        .CellBorderRange .Row, mSpecColumn.规格_药品规格, .Row, mSpecColumn.规格_药品规格, mlngBorderColor, 2, 2, 0, 2, 2, 2
                        .CellBorderRange .Row, intCol, .Row, intCol, mlngBorderColor, 0, 2, 2, 2, 2, 2
                    Case 7
                        If mint配置中心 <> 0 And Not mstrNode Like "中草药*" Then
                            .CellBorderRange .Row, mSpecColumn.规格_药品规格, .Row, mSpecColumn.规格_输液注意事项, mlngBorderColor, 0, 2, 0, 2, 0, 2
                            .CellBorderRange .Row, mSpecColumn.规格_药品规格, .Row, mSpecColumn.规格_药品规格, mlngBorderColor, 2, 2, 0, 2, 2, 2
                            .CellBorderRange .Row, mSpecColumn.规格_输液注意事项, .Row, mSpecColumn.规格_输液注意事项, mlngBorderColor, 0, 2, 2, 2, 2, 2
                        End If
                End Select
            End If
        End If
    End With
End Sub

Private Sub vsfDetails_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call MoveRowCol
    End If
End Sub

Private Sub vsfDetails_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strSql As String, strSQLItem As String
    Dim rsRecord As ADODB.Recordset
    Dim iAttr As Integer
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    Dim intAllCol As Integer

    On Error GoTo ErrHandle
    If KeyAscii = vbKeyReturn Then
        If vsfDetails.EditText = "" Then
            Call MoveRowCol
            Exit Sub
        End If

        If mint状态 = 1 Then '品种
            vRect = zlControl.GetControlRect(vsfDetails.hwnd) '获取位置
            dblLeft = vRect.Left + vsfDetails.CellLeft
            dblTop = vRect.Top + vsfDetails.CellTop + vsfDetails.CellHeight + 3200
            With vsfDetails
                If .Col = mVaricolumn.品种_参考项目 Then
                    strSql = "Select 类型 From 诊疗分类目录 Where ID=[1]"
                    Set rsRecord = zlDatabase.OpenSQLRecord(strSql, Me.Caption, .TextMatrix(.Row, mVaricolumn.品种_分类id))

                    If rsRecord.EOF Then
                        iAttr = -1
                    Else
                        iAttr = rsRecord(0)
                    End If
                    If .EditText = "" Then
                        strSql = " Select ID,分类ID,编码,名称,说明 From 诊疗参考目录 a Where 类型=" & iAttr & " Order By 编码"
                    Else
                        strSQLItem = " From 诊疗参考目录 A,诊疗参考别名 B" & _
                            " Where A.ID=B.参考目录ID And A.类型=[1]" & _
                            " And (Upper(A.编码) Like [2] " & _
                            " Or Upper(A.名称) Like [3] " & _
                            " Or Upper(B.名称) Like [3] " & _
                            " Or Upper(B.简码) Like [3] " & ")"

                        strSql = " Select DISTINCT A.ID,A.分类ID,A.编码,A.名称,A.说明 " & strSQLItem & " Order By 编码"
                        Set rsRecord = zlDatabase.ShowSQLSelect(Me, strSql, 0, "诊疗项目", False, "", "", False, False, _
                        True, dblLeft, dblTop, vsfDetails.Height, blnCancel, False, True, iAttr, UCase(.EditText) & "%", mstrMatch & UCase(.EditText) & "%")

                        If rsRecord Is Nothing Then
                            Exit Sub
                        End If
                        .EditText = rsRecord!名称
                        .TextMatrix(.Row, mVaricolumn.品种_参考项目) = rsRecord!名称
                        .TextMatrix(.Row, mVaricolumn.品种_参考项目ID) = rsRecord!ID
                        End If
                End If
            End With
        Else    '规格
            Dim str As String
            vRect = zlControl.GetControlRect(vsfDetails.hwnd) '获取位置
            dblLeft = vRect.Left + vsfDetails.CellLeft
            dblTop = vRect.Top + vsfDetails.CellTop + vsfDetails.CellHeight + 3200
            With vsfDetails
                If .EditText = "" Then Exit Sub
                Select Case Col
                    Case mSpecColumn.规格_生产商
                        str = UCase(.EditText)
                        If .Col = mSpecColumn.规格_生产商 Then
                            strSql = "Select 编码 as id,名称,简码" & _
                                        " From 药品生产商" & _
                                        " where 编码 Like [1] " & _
                                        "       Or 名称 Like [2] " & _
                                        "       Or 简码 Like [2] Order By 编码 "
                            Set rsRecord = zlDatabase.ShowSQLSelect(Me, strSql, 0, "诊疗项目", False, "", "", False, False, _
                                True, dblLeft, dblTop, vsfDetails.Height, blnCancel, False, True, str & "%", mstrMatch & str & "%")
                            If rsRecord Is Nothing Then
                                .EditText = ""
                                Exit Sub
                            Else
                                .EditText = Nvl(rsRecord!名称)
                                .TextMatrix(.Row, mSpecColumn.规格_生产商) = Nvl(rsRecord!名称)
                            End If
                        End If
                    Case mSpecColumn.规格_原产地
                        str = UCase(.EditText)
                        If .Col = mSpecColumn.规格_原产地 Then
                            strSql = "Select 编码 as id,名称,简码" & _
                                        " From 药品生产商" & _
                                        " where 编码 Like [1] " & _
                                        "       Or 名称 Like [2] " & _
                                        "       Or 简码 Like [2] Order By 编码 "
                            Set rsRecord = zlDatabase.ShowSQLSelect(Me, strSql, 0, "诊疗项目", False, "", "", False, False, _
                                True, dblLeft, dblTop, vsfDetails.Height, blnCancel, False, True, str & "%", mstrMatch & str & "%")
                            If rsRecord Is Nothing Then
                                .EditText = ""
                                Exit Sub
                            Else
                                .EditText = Nvl(rsRecord!名称)
                                .TextMatrix(.Row, mSpecColumn.规格_原产地) = Nvl(rsRecord!名称)
                            End If
                        End If
                    Case mSpecColumn.规格_合同单位
                        strSql = "Select 编码,名称,简码,id" & _
                                    " From 供应商" & _
                                    " where (编码 Like [1] " & _
                                    "       Or 名称 Like [2] " & _
                                    "       Or 简码 Like [2])" & _
                                    " And 末级=1 And substr(类型,1,1) = '1' And (撤档时间 is null or 撤档时间=to_date('3000-01-01','YYYY-MM-DD')) " & _
                                    " Order By 编码 "
                        Set rsRecord = zlDatabase.ShowSQLSelect(Me, strSql, 0, "诊疗项目", False, "", "", False, False, _
                            True, dblLeft, dblTop, vsfDetails.Height, blnCancel, False, True, UCase(.EditText) & "%", mstrMatch & UCase(.EditText) & "%")

                        If rsRecord Is Nothing Then
                            MsgBox "没有找到匹配的供应商，请在供应商管理中增加供应商！", vbInformation, gstrSysName
                            .TextMatrix(.Row, mSpecColumn.规格_合同单位) = ""
                            .TextMatrix(.Row, mSpecColumn.规格_合同单位id) = ""
                            Exit Sub
                        Else
                            .EditText = rsRecord!名称
                            .TextMatrix(.Row, mSpecColumn.规格_合同单位) = rsRecord!名称
                            .TextMatrix(.Row, mSpecColumn.规格_合同单位id) = rsRecord!ID
                        End If
                End Select
            End With
        End If

        Call MoveRowCol
    End If

    If KeyAscii <> vbKeyBack Then
        With vsfDetails
            If mint状态 = 1 Then    '品种
                Select Case Col
                    Case mVaricolumn.品种_通用名称
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mVaricolumn.品种_通用名称)) Then
                            KeyAscii = 0
                        End If
                    Case mVaricolumn.品种_英文名称
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mVaricolumn.品种_英文名称)) Then
                            KeyAscii = 0
                        End If
                    Case mVaricolumn.品种_拼音码
                        If Not (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Or KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mVaricolumn.品种_拼音码)) Then
                            KeyAscii = 0
                        End If
                    Case mVaricolumn.品种_五笔码
                        If Not (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Or KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mVaricolumn.品种_五笔码)) Then
                            KeyAscii = 0
                        End If
                    Case mVaricolumn.品种_处方限量
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mVaricolumn.品种_处方限量)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mVaricolumn.品种_剂量单位
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mVaricolumn.品种_处方限量)) Then
                            KeyAscii = 0
                        End If
                    Case mVaricolumn.品种_ATCCODE
                        If KeyAscii <> vbKeyDelete Then
                            If Not (KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mVaricolumn.品种_ATCCODE)) Then
                                KeyAscii = 0
                            Else
                                KeyAscii = Asc(UCase(Chr(KeyAscii)))
                            End If
                        End If
                End Select
            Else    '规格
                Select Case Col
                    Case mSpecColumn.规格_药品规格
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_药品规格)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.规格_本位码
                        If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= 20 Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.规格_数字码
                        If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= 7 Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.规格_标识码
                        If Not (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Or KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_标识码)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.规格_备选码
                        If Not (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Or KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_备选码)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.规格_容量
                        If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or Val(.EditText) >= 999999999999999# Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.规格_商品名称
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_商品名称)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.规格_生产商
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_生产商)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.规格_原产地
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_原产地)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.规格_拼音码
                        If Not (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Or KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_拼音码)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.规格_五笔码
                        If Not (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Or KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_五笔码)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.规格_合同单位
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_合同单位)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.规格_批准文号
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_批准文号)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.规格_注册商标

                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_注册商标)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.规格_售价单位
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_售价单位)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.规格_剂量系数
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_剂量系数)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.规格_住院单位
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= mintLen住院单位 Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.规格_住院系数
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= mintLen住院系数 Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.规格_门诊单位
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_门诊单位)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.规格_门诊系数
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_门诊系数)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.规格_药库单位
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_药库单位)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.规格_药库系数
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_药库系数)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.规格_送货单位
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_送货单位)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.规格_送货包装
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_送货包装)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.规格_申领阀值
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_申领阀值)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.规格_采购限价
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_采购限价)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.规格_采购扣率
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_采购扣率)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.规格_指导售价
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_指导售价)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.规格_加成率
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= 19 Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.规格_成本价格
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_成本价格)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.规格_当前售价
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_当前售价)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.规格_保质期
                        If KeyAscii = vbKeyDelete Then
                            KeyAscii = 0
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_保质期)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.规格_标识说明
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_标识说明)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.规格_输液注意事项
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_输液注意事项)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.规格_存储库房
                            KeyAscii = 0
                    Case mSpecColumn.规格_服务科室
                            KeyAscii = 0
                End Select
            End If
        End With
    Else
        If Col = mSpecColumn.规格_存储库房 Or Col = mSpecColumn.规格_服务科室 Then
            KeyAscii = 0
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsfDetails_LeaveCell()
    Dim i As Integer
    Dim j As Integer
    
    With vsfDetails
        mintRow上次 = .Row
        mintCol上次 = .Col
    End With
End Sub


Private Sub vsfDetails_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        mobjPopup.ShowPopup
    End If
End Sub

Private Sub Set权限判断()
'权限判断过程
    With vsfDetails
        If .Rows > 1 Then
            If mint状态 = 1 Then    '品种
                If InStr(1, mstrPrivs, "医保用药目录") = 0 Then
                    .Cell(flexcpBackColor, 1, mVaricolumn.品种_医保职务, .Rows - 1, mVaricolumn.品种_医保职务) = mlngColor
                End If
            Else    '规格
                If InStr(1, mstrPrivs, "医保用药目录") = 0 Then
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_医保类型, .Rows - 1, mSpecColumn.规格_医保类型) = mlngColor
                End If
                If InStr(1, mstrPrivs, "管理扣率") = 0 Then
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_采购扣率, .Rows - 1, mSpecColumn.规格_采购扣率) = mlngColor
                End If
                If InStr(1, mstrPrivs, "指导价格管理") = 0 Then
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_加成率, .Rows - 1, mSpecColumn.规格_加成率) = mlngColor
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_采购限价, .Rows - 1, mSpecColumn.规格_采购限价) = mlngColor
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_指导售价, .Rows - 1, mSpecColumn.规格_指导售价) = mlngColor
                End If
                If InStr(1, mstrPrivs, "售价管理") = 0 Then
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_药价属性, .Rows - 1, mSpecColumn.规格_药价属性) = mlngColor
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_收入项目, .Rows - 1, mSpecColumn.规格_收入项目) = mlngColor
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_当前售价, .Rows - 1, mSpecColumn.规格_当前售价) = mlngColor
                End If
                If InStr(1, mstrPrivs, "药价级别") = 0 Then
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_药价级别, .Rows - 1, mSpecColumn.规格_药价级别) = mlngColor
                End If
                If InStr(1, mstrPrivs, "成本价管理") = 0 Then
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_成本价格, .Rows - 1, mSpecColumn.规格_成本价格) = mlngColor
                End If
                If InStr(1, mstrPrivs, "调整服务对象") = 0 Then
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_服务对象, .Rows - 1, mSpecColumn.规格_服务对象) = mlngColor
                End If
                If InStr(1, mstrPrivs, "药品单位管理") = 0 Then
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_售价单位, .Rows - 1, mSpecColumn.规格_售价单位) = mlngColor
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_住院单位, .Rows - 1, mSpecColumn.规格_住院单位) = mlngColor
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_门诊单位, .Rows - 1, mSpecColumn.规格_门诊单位) = mlngColor
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_药库单位, .Rows - 1, mSpecColumn.规格_药库单位) = mlngColor
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_剂量系数, .Rows - 1, mSpecColumn.规格_剂量系数) = mlngColor
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_住院系数, .Rows - 1, mSpecColumn.规格_住院系数) = mlngColor
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_门诊系数, .Rows - 1, mSpecColumn.规格_门诊系数) = mlngColor
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_药库系数, .Rows - 1, mSpecColumn.规格_药库系数) = mlngColor
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_送货单位, .Rows - 1, mSpecColumn.规格_送货单位) = mlngColor
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_送货包装, .Rows - 1, mSpecColumn.规格_送货包装) = mlngColor
                End If

                If InStr(1, mstrPrivs, "存储库房") = 0 Then
                    tbcDetails.Item(mSpecList.存储库房).Visible = False
                End If
                
                If mstrNode Like "中草药*" Then
                    If InStr(1, mstrPrivs, "草药分包管理") = 0 Then
                        .Cell(flexcpBackColor, 1, mSpecColumn.规格_中药形态, .Rows - 1, mSpecColumn.规格_中药形态) = mlngColor
                    End If
                End If
                
                If Val(zlDatabase.GetPara(275, glngSys, , 0)) = 0 Then
                     .Cell(flexcpBackColor, 1, mSpecColumn.规格_零差价管理, .Rows - 1, mSpecColumn.规格_零差价管理) = mlngColor
                End If
            End If
        End If
    End With
End Sub

Private Sub Save()
    '数据保存方法
    Dim i As Integer
    Dim strTemp As String
    Dim j As Integer
    Dim m As Integer
    Dim n As Integer
    Dim intupdate As Integer
    Dim rsRecord As ADODB.Recordset
    Dim str别名 As String
    Dim intCount As Integer
    Dim bln修改 As Boolean
    Dim lng保存 As Long
    Dim lngSave As Long
    Dim intTemp As Integer
    Dim blnShowMsg As Boolean
    Dim blnTrans As Boolean
    Dim arrSql() As Variant     '纪录存储过程的数组
    Dim rsOther As ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    Dim strPara As String
    Dim str其他库房ID As String
    Dim strIdArr As Variant
    Dim str科室ID As String
    Dim intRow As Integer
    Dim dbl所有库房 As Boolean
    Dim str药品分类 As String
                
    bln修改 = Check修改
    On Error GoTo ErrHandle
    
    arrSql = Array()
    If bln修改 = False Then '没有修改的话直接退出不进行保存
        Exit Sub
    End If

    If mintExit <> 2 Then
        lngSave = MsgBox("是否批量保存已修改的记录？", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName)
        If lngSave = vbNo Then
            Exit Sub
        End If
        mintExit = 0
    End If
    With vsfDetails
        If mint状态 = 1 Then    '品种
            If .TextMatrix(1, mVaricolumn.品种_id) = "" Then Exit Sub
            '检查数据的合法性
            If CheckData = False Then Exit Sub

            If mstrNode Like "中草药*" Then '中草药
                For i = 1 To .Rows - 1
                    gstrSql = ""
                    strTemp = ""
                    gstrSql = "Zl_草药品种_Update (" & .TextMatrix(i, mVaricolumn.品种_分类id) & ","
                    gstrSql = gstrSql + .TextMatrix(i, mVaricolumn.品种_id) & ","
                    gstrSql = gstrSql + "'" + .TextMatrix(i, mVaricolumn.品种_药品编码) + "'" & ","
                    gstrSql = gstrSql + "'" + .TextMatrix(i, mVaricolumn.品种_通用名称) + "'" & ","
                    gstrSql = gstrSql + "'" + .TextMatrix(i, mVaricolumn.品种_拼音码) + "'" & ","
                    gstrSql = gstrSql + "'" + .TextMatrix(i, mVaricolumn.品种_五笔码) + "'" & ","
                    gstrSql = gstrSql + "'" + .TextMatrix(i, mVaricolumn.品种_英文名称) + "'" & ","
                    gstrSql = gstrSql + "'" + .TextMatrix(i, mVaricolumn.品种_剂量单位) + "'" & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.品种_毒理分类), InStr(1, .TextMatrix(i, mVaricolumn.品种_毒理分类), "-") + 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.品种_价值分类), InStr(1, .TextMatrix(i, mVaricolumn.品种_价值分类), "-") + 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.品种_货源情况), InStr(1, .TextMatrix(i, mVaricolumn.品种_货源情况), "-") + 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.品种_用药梯次), InStr(1, .TextMatrix(i, mVaricolumn.品种_用药梯次), "-") + 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.品种_药品类型), 1, 1)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.品种_处方职务), 1, 1) + Mid(.TextMatrix(i, mVaricolumn.品种_医保职务), 1, 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    gstrSql = gstrSql + .TextMatrix(i, mVaricolumn.品种_处方限量) & ","
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.品种_单味使用) = "√", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.品种_原料药) = "√", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.品种_适用性别), 1, 1)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.品种_参考项目) = "", "Null", .TextMatrix(i, mVaricolumn.品种_参考项目ID))
                    gstrSql = gstrSql + strTemp & ","

                    str别名 = "Select distinct n.名称 as 药品名称, p.简码 As 拼音, w.简码 As 五笔" & _
                              "  From (Select Distinct 诊疗项目id,名称 From 诊疗项目别名 Where  性质 = 9) N," & _
                                    " (Select 名称, 简码 From 诊疗项目别名 Where  性质 = 9 And 码类 = 1) P," & _
                                    " (Select 名称, 简码 From 诊疗项目别名 Where  性质 = 9 And 码类 = 2) W" & _
                               " Where n.名称 = p.名称(+) And n.名称 = w.名称(+) and n.诊疗项目id = [1]"
                    Set rsRecord = zlDatabase.OpenSQLRecord(str别名, "品种保存", .TextMatrix(i, mVaricolumn.品种_id))
                    
                    strTemp = ""
                    If Not rsRecord.EOF Then
                        Do While Not rsRecord.EOF
                            strTemp = strTemp & "|" & rsRecord!药品名称 & "^" & rsRecord!拼音 & "^" & rsRecord!五笔
                            rsRecord.MoveNext
                        Loop
                    End If

                    If strTemp <> "" Then
                        strTemp = Mid(strTemp, 2)
                        gstrSql = gstrSql + "'" + strTemp + "'"
                    Else
                        strTemp = "Null"
                        gstrSql = gstrSql + strTemp
                    End If
 
                    strTemp = ",NULL," & IIf(.TextMatrix(i, mVaricolumn.品种_辅助用药) = "√", 1, 0)
                    gstrSql = gstrSql + strTemp & ")"
                    
                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = gstrSql
'                    zlDatabase.ExecuteProcedure gstrSql, "保存"
                Next
                '其他别名
            Else    '西成药、中成药
                For i = 1 To vsfDetails.Rows - 1
                    gstrSql = ""
                    strTemp = ""

                    gstrSql = "Zl_成药品种_Update (" & .TextMatrix(i, mVaricolumn.品种_分类id) & ","
                    gstrSql = gstrSql + .TextMatrix(i, mVaricolumn.品种_id) & ","
                    gstrSql = gstrSql + "'" + .TextMatrix(i, mVaricolumn.品种_药品编码) + "'" & ","
                    gstrSql = gstrSql + "'" + .TextMatrix(i, mVaricolumn.品种_通用名称) + "'" & ","
                    gstrSql = gstrSql + "'" + .TextMatrix(i, mVaricolumn.品种_拼音码) + "'" & ","
                    gstrSql = gstrSql + "'" + .TextMatrix(i, mVaricolumn.品种_五笔码) + "'" & ","
                    gstrSql = gstrSql + "'" + .TextMatrix(i, mVaricolumn.品种_英文名称) + "'" & ","
                    gstrSql = gstrSql + "'" + .TextMatrix(i, mVaricolumn.品种_剂量单位) + "'" & ","

                    strTemp = Mid(.TextMatrix(i, mVaricolumn.品种_剂型), InStr(1, .TextMatrix(i, mVaricolumn.品种_剂型), "-") + 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.品种_毒理分类), InStr(1, .TextMatrix(i, mVaricolumn.品种_毒理分类), "-") + 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.品种_价值分类), InStr(1, .TextMatrix(i, mVaricolumn.品种_价值分类), "-") + 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.品种_货源情况), InStr(1, .TextMatrix(i, mVaricolumn.品种_货源情况), "-") + 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.品种_用药梯次), InStr(1, .TextMatrix(i, mVaricolumn.品种_用药梯次), "-") + 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.品种_药品类型), 1, 1)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.品种_处方职务), 1, 1) + Mid(.TextMatrix(i, mVaricolumn.品种_医保职务), 1, 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    gstrSql = gstrSql + .TextMatrix(i, mVaricolumn.品种_处方限量) & ","

                    strTemp = IIf(.TextMatrix(i, mVaricolumn.品种_急救药) = "√", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.品种_新药) = "√", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.品种_原料药) = "√", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.品种_皮试) = "√", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.品种_抗生素), 1, 1)
                    gstrSql = gstrSql + strTemp & ","

                    '参考目录ID
                    '''''''''''''''''''''
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.品种_参考项目) = "", "Null", .TextMatrix(i, mVaricolumn.品种_参考项目ID))
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.品种_品种下长期医嘱) = "√", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.品种_适用性别), 1, 1)
                    gstrSql = gstrSql + strTemp & ","

                    '别名
                    str别名 = "Select distinct n.名称 as 药品名称, p.简码 As 拼音, w.简码 As 五笔" & _
                              "  From (Select Distinct 诊疗项目id,名称 From 诊疗项目别名 Where  性质 = 9) N," & _
                                    " (Select 名称, 简码 From 诊疗项目别名 Where  性质 = 9 And 码类 = 1) P," & _
                                    " (Select 名称, 简码 From 诊疗项目别名 Where  性质 = 9 And 码类 = 2) W" & _
                               " Where n.名称 = p.名称(+) And n.名称 = w.名称(+) and n.诊疗项目id = [1]"
                    Set rsRecord = zlDatabase.OpenSQLRecord(str别名, "品种保存", .TextMatrix(i, mVaricolumn.品种_id))
                    
                    strTemp = ""
                    If Not rsRecord.EOF Then
                        Do While Not rsRecord.EOF
                            strTemp = strTemp & "|" & rsRecord!药品名称 & "^" & rsRecord!拼音 & "^" & rsRecord!五笔
                            rsRecord.MoveNext
                        Loop
                    End If
                    If strTemp <> "" Then
                        strTemp = Mid(strTemp, 2)
                        gstrSql = gstrSql + "'" & strTemp & "',"
                    Else
                        strTemp = "Null"
                        gstrSql = gstrSql + strTemp & ","
                    End If
                    gstrSql = gstrSql + "Null,"
                    gstrSql = gstrSql + "'" & .TextMatrix(i, mVaricolumn.品种_ATCCODE) & "',"
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.品种_肿瘤药) = "√", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.品种_溶媒) = "√", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.品种_原研药) = "√", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.品种_专利药) = "√", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.品种_单独定价) = "√", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.品种_辅助用药) = "√", 1, 0)
                    gstrSql = gstrSql + strTemp & ")"
                    
                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = gstrSql
'                    zlDatabase.ExecuteProcedure gstrSql, "保存"
                Next
            End If
        Else    '规格
            If .TextMatrix(1, mSpecColumn.规格_id) = "" Then Exit Sub
            '检查数据的合法性
            If CheckData = False Then Exit Sub

            For i = 1 To vsfDetails.Rows - 1
                If .TextMatrix(i, mSpecColumn.规格_药品规格) = "" Then
                    MsgBox "第" & i & "行药品规格为空，请收入药品规格！", vbExclamation, gstrSysName
                    Exit Sub
                End If
            Next

            If mstrNode Like "中草药*" Then '中草药
                For i = 1 To vsfDetails.Rows - 1
                    gstrSql = "zl_草药规格_Update(" & .TextMatrix(i, mSpecColumn.规格_id) & ","
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_规格编码) & "',"
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_药品规格) & "',"
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_生产商) & "',"

                    If .TextMatrix(i, mSpecColumn.规格_商品名称) <> "" Then
                        strTemp = "'" & .TextMatrix(i, mSpecColumn.规格_商品名称) & "'"
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql & strTemp & ","

                    strTemp = "Null"    '拼音码
                    gstrSql = gstrSql & strTemp & ","

                    strTemp = "Null"    '五笔码
                    gstrSql = gstrSql & strTemp & ","

                    If .TextMatrix(i, mSpecColumn.规格_数字码) <> "" Then
                        strTemp = "'" & .TextMatrix(i, mSpecColumn.规格_数字码) & "'"
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql & strTemp & ","

                    If .TextMatrix(i, mSpecColumn.规格_标识码) <> "" Then
                        strTemp = "'" & .TextMatrix(i, mSpecColumn.规格_标识码) & "'"
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql & strTemp & ","

                    If .TextMatrix(i, mSpecColumn.规格_来源分类) <> "" Then
                        strTemp = "'" & Mid(.TextMatrix(i, mSpecColumn.规格_来源分类), InStr(1, .TextMatrix(i, mSpecColumn.规格_来源分类), "-") + 1) & "'"
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql + strTemp + ","

                    If .TextMatrix(i, mSpecColumn.规格_批准文号) <> "" Then
                        strTemp = "'" & .TextMatrix(i, mSpecColumn.规格_批准文号) & "'"
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql & strTemp & ","

                    If .TextMatrix(i, mSpecColumn.规格_注册商标) <> "" Then
                        strTemp = "'" & .TextMatrix(i, mSpecColumn.规格_注册商标) & "'"
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql & strTemp & ","

                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_售价单位) & "',"
                    gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_剂量系数) & ","
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_门诊单位) & "',"
                    gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_门诊系数) & ","
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_药库单位) & "',"
                    gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_药库系数) & ","

                    Select Case .TextMatrix(i, mSpecColumn.规格_申领单位)
                        Case "1-售价单位"
                            strTemp = 1
                        Case "2-住院单位"
                            strTemp = 2
                        Case "3-门诊单位"
                            strTemp = 3
                        Case "4-药库单位"
                            strTemp = 4
                    End Select
                    gstrSql = gstrSql & strTemp & ","

                    If Trim(.TextMatrix(i, mSpecColumn.规格_申领阀值)) <> "" Then
                        strTemp = .TextMatrix(i, mSpecColumn.规格_申领阀值)
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql & strTemp & ","
                    strTemp = Mid(.TextMatrix(i, mSpecColumn.规格_药价属性), 1, 1)
                    gstrSql = gstrSql & strTemp & ","
                    gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_采购限价) & ","
                    gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_采购扣率) & ","

                    If mint当前单位 <> 0 Then
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_指导售价) / Nvl(.TextMatrix(i, mSpecColumn.规格_药库系数), 1) & ","
                    Else
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_指导售价) & ","
                    End If

                    gstrSql = gstrSql & IIf(.TextMatrix(i, mSpecColumn.规格_加成率) = "", "Null", .TextMatrix(i, mSpecColumn.规格_加成率)) & ","
                    '管理费比例
                    strTemp = "0"
                    gstrSql = gstrSql & strTemp & ","

                    strTemp = Mid(.TextMatrix(i, mSpecColumn.规格_药价级别), InStr(1, .TextMatrix(i, mSpecColumn.规格_药价级别), "-") + 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","

                    strTemp = Mid(.TextMatrix(i, mSpecColumn.规格_医保类型), InStr(1, .TextMatrix(i, mSpecColumn.规格_医保类型), "-") + 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","

                    strTemp = Mid(.TextMatrix(i, mSpecColumn.规格_服务对象), 1, 1)
                    gstrSql = gstrSql & strTemp & ","

                    strTemp = IIf(.TextMatrix(i, mSpecColumn.规格_GMP认证) = "√", 1, 0)
                    gstrSql = gstrSql + strTemp & ","

                    gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_招标药品) & ","
                    strTemp = IIf(.TextMatrix(i, mSpecColumn.规格_屏蔽费别) = "√", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = Mid(.TextMatrix(i, mSpecColumn.规格_住院分零使用), 1, 1)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = IIf(.TextMatrix(i, mSpecColumn.规格_药库分批) = "√", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = IIf(.TextMatrix(i, mSpecColumn.规格_药房分批) = "√", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_保质期) & ","
                    '差价让利
                    strTemp = "100"
                    gstrSql = gstrSql & strTemp & ","

                    If mint当前单位 <> 0 Then
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_成本价格) / Nvl(.TextMatrix(i, mSpecColumn.规格_药库系数), 1) & ","
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_当前售价) / Nvl(.TextMatrix(i, mSpecColumn.规格_药库系数), 1) & ","
                    Else
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_成本价格) & ","
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_当前售价) & ","
                    End If


                    strTemp = .TextMatrix(i, mSpecColumn.规格_收入项目id)
                    gstrSql = gstrSql & strTemp & ","

                    gstrSql = gstrSql & IIf(.TextMatrix(i, mSpecColumn.规格_合同单位id) = "", "Null", .TextMatrix(i, mSpecColumn.规格_合同单位id)) & ","
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_标识说明) & "',"
                    strTemp = IIf(.TextMatrix(i, mSpecColumn.规格_住院动态分零) = "√", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_发药类型) & "',"
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_备选码) & "',"

                    strTemp = "0"
                    gstrSql = gstrSql & strTemp & ","

                    gstrSql = gstrSql & "'',"

                    Select Case .TextMatrix(i, mSpecColumn.规格_中药形态)
                        Case "0-散装"
                            strTemp = 0
                        Case "1-中药饮片"
                            strTemp = 1
                        Case "2-免煎剂"
                            strTemp = 2
                    End Select
                    gstrSql = gstrSql + strTemp & ","

                    If .TextMatrix(i, mSpecColumn.规格_站点编号) <> "" Then
                        strTemp = Mid(.TextMatrix(i, mSpecColumn.规格_站点编号), 1, 1)
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql & strTemp & ","

                    strTemp = IIf(.TextMatrix(i, mSpecColumn.规格_非常备药) = "√", 1, 0)
                    gstrSql = gstrSql + strTemp & ","

                    If .TextMatrix(i, mSpecColumn.规格_病案费目) <> "" Then
                        strTemp = "'" & .TextMatrix(i, mSpecColumn.规格_病案费目) & "'"
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql + strTemp & ","

                    strTemp = Mid(.TextMatrix(i, mSpecColumn.规格_门诊分零使用), 1, 1)
                    
                    gstrSql = gstrSql + strTemp & ","
                    
                    gstrSql = gstrSql + "'" & Trim(.TextMatrix(i, mSpecColumn.规格_送货单位)) & "',"
                    strTemp = IIf(Trim(.TextMatrix(i, mSpecColumn.规格_送货包装)) = "", "Null", Trim(.TextMatrix(i, mSpecColumn.规格_送货包装)))
                    gstrSql = gstrSql + strTemp & ","
                    
                    strTemp = IIf(.TextMatrix(i, mSpecColumn.规格_是否摆药) = "是", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                                        
                    strTemp = IIf(.TextMatrix(i, mSpecColumn.规格_零差价管理) = "√", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    
                    strTemp = "'" & .TextMatrix(i, mSpecColumn.规格_本位码) & "',"
                    gstrSql = gstrSql + strTemp
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_原产地) & "'" & ")"

                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = gstrSql
'                    zlDatabase.ExecuteProcedure gstrSql, "草药规格保存"
                Next
            Else    '西成药、中成药
                For i = 1 To vsfDetails.Rows - 1
                    gstrSql = ""
                    gstrSql = "zl_成药规格_Update(" & .TextMatrix(i, mSpecColumn.规格_id) & ","
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_规格编码) & "',"
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_药品规格) & "',"
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_生产商) & "',"
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_商品名称) & "',"
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_拼音码) & "',"
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_五笔码) & "',"
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_数字码) & "',"

                    If Trim(.TextMatrix(i, mSpecColumn.规格_标识码)) = "" Then
                        strTemp = "Null"
                    Else
                        strTemp = "'" & .TextMatrix(i, mSpecColumn.规格_标识码) & "'"
                    End If
                    gstrSql = gstrSql & strTemp & ","

                    strTemp = Mid(.TextMatrix(i, mSpecColumn.规格_来源分类), InStr(1, .TextMatrix(i, mSpecColumn.规格_来源分类), "-") + 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_批准文号) & "',"
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_注册商标) & "',"
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_售价单位) & "',"
                    gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_剂量系数) & ","
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_门诊单位) & "',"
                    gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_门诊系数) & ","
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_住院单位) & "',"
                    gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_住院系数) & ","
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_药库单位) & "',"
                    gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_药库系数) & ","

                    Select Case .TextMatrix(i, mSpecColumn.规格_申领单位)
                        Case "1-售价单位"
                            strTemp = 1
                        Case "2-住院单位"
                            strTemp = 2
                        Case "3-门诊单位"
                            strTemp = 3
                        Case "4-药库单位"
                            strTemp = 4
                    End Select
                    gstrSql = gstrSql & strTemp & ","

                    If Trim(.TextMatrix(i, mSpecColumn.规格_申领阀值)) = "" Then
                        strTemp = "Null"
                    Else
                        strTemp = .TextMatrix(i, mSpecColumn.规格_申领阀值)
                    End If
                    gstrSql = gstrSql & strTemp & ","

                    strTemp = Mid(.TextMatrix(i, mSpecColumn.规格_药价属性), 1, 1)
                    gstrSql = gstrSql & strTemp & ","

                    If mint当前单位 <> 0 Then
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_采购限价) / Nvl(.TextMatrix(i, mSpecColumn.规格_药库系数), 1) & ","
                    Else
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_采购限价) & ","
                    End If

                    gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_采购扣率) & ","

                    If mint当前单位 <> 0 Then
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_指导售价) / Nvl(.TextMatrix(i, mSpecColumn.规格_药库系数), 1) & ","
                    Else
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_指导售价) & ","
                    End If
                    
                    If .TextMatrix(i, mSpecColumn.规格_加成率) = "" Then
                        strTemp = "null"
                    Else
                        strTemp = .TextMatrix(i, mSpecColumn.规格_加成率)
                    End If
                    gstrSql = gstrSql & strTemp & ","

                    '管理费比例
                    strTemp = "0"
                    gstrSql = gstrSql & strTemp & ","

                    strTemp = Mid(.TextMatrix(i, mSpecColumn.规格_药价级别), InStr(1, .TextMatrix(i, mSpecColumn.规格_药价级别), "-") + 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","

                    strTemp = Mid(.TextMatrix(i, mSpecColumn.规格_医保类型), InStr(1, .TextMatrix(i, mSpecColumn.规格_医保类型), "-") + 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","

                    strTemp = Mid(.TextMatrix(i, mSpecColumn.规格_服务对象), 1, 1)
                    gstrSql = gstrSql & strTemp & ","

                    strTemp = IIf(.TextMatrix(i, mSpecColumn.规格_GMP认证) = "√", 1, 0)
                    gstrSql = gstrSql + strTemp & ","

                    gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_招标药品) & ","
                    strTemp = IIf(.TextMatrix(i, mSpecColumn.规格_屏蔽费别) = "√", 1, 0)
                    gstrSql = gstrSql + strTemp & ","

                    If .TextMatrix(i, mSpecColumn.规格_住院分零使用) <> "" Then
                        strTemp = Mid(.TextMatrix(i, mSpecColumn.规格_住院分零使用), 1, 1)
                        If strTemp = 0 Then
                            strTemp = "0"
                        ElseIf strTemp = 1 Then
                            strTemp = "1"
                        ElseIf strTemp = 2 Then
                            strTemp = "2"
                        ElseIf strTemp = 3 Then
                            strTemp = "-1"
                        ElseIf strTemp = 4 Then
                            strTemp = "-2"
                        ElseIf strTemp = 5 Then
                            strTemp = "-3"
                        End If
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql + strTemp & ","

                    strTemp = IIf(.TextMatrix(i, mSpecColumn.规格_药库分批) = "√", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = IIf(.TextMatrix(i, mSpecColumn.规格_药房分批) = "√", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_保质期) & ","
                    '差价让利比
                    gstrSql = gstrSql & 100 & ","

                    If mint当前单位 <> 0 Then
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_成本价格) / Nvl(.TextMatrix(i, mSpecColumn.规格_药库系数), 1) & ","
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_当前售价) / Nvl(.TextMatrix(i, mSpecColumn.规格_药库系数), 1) & ","
                    Else
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_成本价格) & ","
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_当前售价) & ","
                    End If

                    strTemp = .TextMatrix(i, mSpecColumn.规格_收入项目id)
                    gstrSql = gstrSql & strTemp & ","

                    gstrSql = gstrSql & IIf(.TextMatrix(i, mSpecColumn.规格_合同单位id) = "", "Null", .TextMatrix(i, mSpecColumn.规格_合同单位id)) & ","
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_标识说明) & "',"
                    strTemp = IIf(.TextMatrix(i, mSpecColumn.规格_住院动态分零) = "√", 1, 0)
                    gstrSql = gstrSql + strTemp & ","

                    If .TextMatrix(i, mSpecColumn.规格_发药类型) = "" Then
                        strTemp = "Null"
                    Else
                        strTemp = "'" & .TextMatrix(i, mSpecColumn.规格_发药类型) & "'"
                    End If
                    gstrSql = gstrSql & strTemp & ","

                    If .TextMatrix(i, mSpecColumn.规格_备选码) = "" Then
                        strTemp = "Null"
                    Else
                        strTemp = "'" & .TextMatrix(i, mSpecColumn.规格_备选码) & "'"
                    End If
                    gstrSql = gstrSql & strTemp & ","
                    '增值税率
                    strTemp = "0"
                    gstrSql = gstrSql & strTemp & ","

                    strTemp = "'" & .TextMatrix(i, mSpecColumn.规格_基本药物) & "'"
                    gstrSql = gstrSql & strTemp & ","

                    If .TextMatrix(i, mSpecColumn.规格_站点编号) <> "" Then
                        strTemp = Mid(.TextMatrix(i, mSpecColumn.规格_站点编号), 1, 1)
                    Else
                       strTemp = "Null"
                    End If
                    gstrSql = gstrSql & strTemp & ","

                    strTemp = IIf(.TextMatrix(i, mSpecColumn.规格_非常备药) = "√", 1, 0)
                    gstrSql = gstrSql + strTemp & ","

                    If Trim(.TextMatrix(i, mSpecColumn.规格_存储温度)) <> "" Then
                        strTemp = Mid(.TextMatrix(i, mSpecColumn.规格_存储温度), 1, 1)
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql + strTemp & ","

                    strTemp = IIf(.TextMatrix(i, mSpecColumn.规格_存储条件) = "√", 1, 0)
                    gstrSql = gstrSql + strTemp & ","

                    strTemp = "'" & Trim(.TextMatrix(i, mSpecColumn.规格_配药类型)) & "'"
                    gstrSql = gstrSql + strTemp & ","

                    strTemp = IIf(.TextMatrix(i, mSpecColumn.规格_不予调配) = "√", 1, 0)
                    gstrSql = gstrSql + strTemp & ","

                    If Trim(.TextMatrix(i, mSpecColumn.规格_容量)) <> "" Then
                        strTemp = .TextMatrix(i, mSpecColumn.规格_容量)
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql + strTemp & ","

                    If .TextMatrix(i, mSpecColumn.规格_病案费目) <> "" Then
                        strTemp = "'" & .TextMatrix(i, mSpecColumn.规格_病案费目) & "'"
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql + strTemp & ","

                    If .TextMatrix(i, mSpecColumn.规格_门诊分零使用) <> "" Then
                        strTemp = Mid(.TextMatrix(i, mSpecColumn.规格_门诊分零使用), 1, 1)
                        If strTemp = 0 Then
                            strTemp = "0"
                        ElseIf strTemp = 1 Then
                            strTemp = "1"
                        ElseIf strTemp = 2 Then
                            strTemp = "2"
                        ElseIf strTemp = 3 Then
                            strTemp = "-1"
                        ElseIf strTemp = 4 Then
                            strTemp = "-2"
                        ElseIf strTemp = 5 Then
                            strTemp = "-3"
                        End If
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql + strTemp & ","

                    If Trim(.TextMatrix(i, mSpecColumn.规格_DDD值)) <> "" Then
                        strTemp = .TextMatrix(i, mSpecColumn.规格_DDD值)
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql + strTemp & ","
                    
                    strTemp = IIf(Trim(.TextMatrix(i, mSpecColumn.规格_高危药品)) = "", 0, Mid(Trim(.TextMatrix(i, mSpecColumn.规格_高危药品)), 1, 1))
                    gstrSql = gstrSql + strTemp & ","
                    
                    gstrSql = gstrSql + "'" & Trim(.TextMatrix(i, mSpecColumn.规格_送货单位)) & "',"
                    strTemp = IIf(Trim(.TextMatrix(i, mSpecColumn.规格_送货包装)) = "", "Null", Trim(.TextMatrix(i, mSpecColumn.规格_送货包装)))
                    gstrSql = gstrSql + strTemp & ","
                    gstrSql = gstrSql + "'" & Trim(.TextMatrix(i, mSpecColumn.规格_输液注意事项)) & "',"
                    
                    strTemp = IIf(.TextMatrix(i, mSpecColumn.规格_是否摆药) = "是", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                                        
                    strTemp = IIf(.TextMatrix(i, mSpecColumn.规格_零差价管理) = "√", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    
                     strTemp = "'" & .TextMatrix(i, mSpecColumn.规格_本位码) & "'"
                    gstrSql = gstrSql + strTemp & ","

                    strTemp = IIf(.TextMatrix(i, mSpecColumn.规格_易跌倒) = "√", 1, 0)
                    gstrSql = gstrSql + strTemp & ","

                    strTemp = IIf(.TextMatrix(i, mSpecColumn.规格_带量采购) = "√", 1, 0)
                    gstrSql = gstrSql + strTemp & ")"
                    
                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = gstrSql
'                    zlDatabase.ExecuteProcedure gstrSql, "规格保存"
                Next
            End If
                        
            '存储库房，服务科室保存
            For i = 1 To vsfDetails.Rows - 1
               
                If mstr药品种类 = "西成药" Then
                    str药品分类 = "'西药%'"
                ElseIf mstr药品种类 = "中成药" Then
                    str药品分类 = "'成药%'"
                Else
                    str药品分类 = "'中药%'"
                End If
                
                strPara = ""
                If InStr(1, ";" & mstrPrivs & ";", ";所有库房;") > 0 Then dbl所有库房 = True

                If Not dbl所有库房 Then
                    '先取其他库房
                    gstrSql = "Select ID, 编码, 名称" & vbNewLine & _
                                    "From 部门表" & vbNewLine & _
                                    "Where ID In (Select Distinct 部门id From 部门性质说明 Where 工作性质 Like " & str药品分类 & " Or 工作性质 = '制剂室') And" & vbNewLine & _
                                    "      (撤档时间 Is Null Or 撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And" & vbNewLine & _
                                    "      ID Not In (Select 部门id From 部门人员 Where 人员id = [1])"

                    Set rsOther = zlDatabase.OpenSQLRecord(gstrSql, "根据药品的用途分类提取所允许存储的库房(其他库房)", UserInfo.ID)
                    
                    str其他库房ID = ""
                    Do While Not rsOther.EOF
                        str其他库房ID = str其他库房ID & "," & rsOther!ID
                        rsOther.MoveNext
                    Loop
                    If str其他库房ID <> "" Then
                        str其他库房ID = Mid(str其他库房ID, 2)
                    End If
                                
                End If
                
                If Not dbl所有库房 And str其他库房ID <> "" Then
                    gstrSql = " Select DISTINCT 开单科室ID,执行科室ID From 收费执行科室 " & _
                          " Where 收费细目ID=[1] And Instr([2],','||执行科室ID||',') > 0 " & _
                          " Order by 执行科室ID"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "提取已设置的收费执行科室数据", .TextMatrix(i, mSpecColumn.规格_id), "," & str其他库房ID & ",")
                                    
                    strIdArr = Split(str其他库房ID, ",")
                    For intRow = 0 To UBound(strIdArr)
                        str科室ID = ""
                        rsTemp.Filter = "执行科室ID=" & strIdArr(intRow)
                        Do While Not rsTemp.EOF
                            str科室ID = str科室ID & "," & Nvl(rsTemp!开单科室ID, 0)
                            rsTemp.MoveNext
                        Loop
                        If str科室ID <> "" Then
                            str科室ID = Mid(str科室ID, 2)
                            If str科室ID = "0" Then str科室ID = ""
                            strPara = strPara & "!!" & CStr(strIdArr(intRow))
                            strPara = strPara & "|" & str科室ID
                        End If
                    Next
                End If
            
            
                If vsfDetails.TextMatrix(i, mSpecColumn.规格_库房科室id) <> "" Then
                    If vsfDetails.TextMatrix(i, mSpecColumn.规格_库房科室id) <> vsfDetails.TextMatrix(i, mSpecColumn.规格_服务科室未改) Then
                        gstrSql = "Zl_药品存储库房_Update(" & vsfDetails.TextMatrix(i, mSpecColumn.规格_id) & ","
                        gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_库房科室id) & strPara & "',1)"
                        
                        ReDim Preserve arrSql(UBound(arrSql) + 1)
                        arrSql(UBound(arrSql)) = gstrSql
                    End If
                Else
                        If .TextMatrix(i, mSpecColumn.规格_存储库房id) = "" Then
                            strPara = Mid(strPara, 3)
                        End If
                        gstrSql = "Zl_药品存储库房_Update(" & vsfDetails.TextMatrix(i, mSpecColumn.规格_id) & ","
                        gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_存储库房id) & strPara & "',1)"
                        
                        ReDim Preserve arrSql(UBound(arrSql) + 1)
                        arrSql(UBound(arrSql)) = gstrSql
                End If
            Next
                        
        End If
    End With
    
    gcnOracle.BeginTrans: blnTrans = True          '开启事务
    For i = 0 To UBound(arrSql)
        Call zlDatabase.ExecuteProcedure(CStr(arrSql(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTrans = False     '提交事物
    
    Call Recover    '保存后刷新界面
    Call getNewData '保存后获取新数据
    Call RecoverData '保存后刷新界面数据
    
    With vsfDetails
        If mint状态 = 2 Then    '规格
            For i = 1 To vsfDetails.Rows - 1
                If .TextMatrix(i, mSpecColumn.规格_零差价管理) = "√" And blnShowMsg = False Then
                    If Val(zlDatabase.GetPara(275, glngSys, , 0)) > 0 Then
                        If CheckPriceAdjust(Val(.TextMatrix(i, mSpecColumn.规格_id)), 0, -1) = False Then
                            blnShowMsg = True
                            MsgBox "部分药品已启用零差价管理，但售价和成本价不一致，请注意调价！", vbInformation, gstrSysName
                        End If
                    End If
                End If
            Next
        End If
    End With
    
    Exit Sub
ErrHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub RecoverData()
    Dim Node As MSComctlLib.Node
    Set Node = tvwDetails.SelectedItem
    Call tvwDetails_NodeClick(Node)
End Sub

Private Sub Recover()
    '使窗体中改变的颜色或者字体还原
    Dim cbrControl As CommandBarControl
    Dim i As Integer
    Dim j As Integer

    With vsfDetails
        For i = 1 To .Rows - 1
            For j = 1 To .Cols - 1
               If .Cell(flexcpBackColor, i, j) <> mlngColor Then
                    .Cell(flexcpBackColor, i, j) = cstcolor_backcolor
                    .Cell(flexcpForeColor, i, j) = vbBlack
                    .Cell(flexcpFontSize, i, j) = 9
                    .Cell(flexcpFontBold, i, j) = False
                End If
                If j = mSpecColumn.规格_保质期 And mint状态 = 2 Then
                    .Cell(flexcpForeColor, i, j) = vbBlack
                    .Cell(flexcpFontSize, i, j) = 9
                    .Cell(flexcpFontBold, i, j) = False
                End If
                If .Cell(flexcpForeColor, i, j) = mlngApplyColor Then
                    .Cell(flexcpForeColor, i, j) = vbBlack
                    .Cell(flexcpFontSize, i, j) = 9
                    .Cell(flexcpFontBold, i, j) = False
                End If
            Next
        Next
    End With
    
    With tbcDetails
        For i = 0 To .ItemCount - 1
            .Item(i).Image = 0
        Next
    End With
    
    mintPos = 0
    mstrChangedCell = ""
    
    Set cbrControl = cbsMain.FindControl(, mcon默认值)
    cbrControl.Enabled = False
    
    Set cbrControl = cbsMain.FindControl(, mcon保存)
    cbrControl.Enabled = False
    
    Set cbrControl = cbsMain.FindControl(, mcon定位已修改项目)
    cbrControl.Enabled = False
End Sub

Private Sub SetBatch()
    '批量设置每一列的值
    Dim i As Integer

    With vsfDetails
        If .Col = mSpecColumn.规格_服务科室 Then
             If MsgBox("因为设置服务科室必须对应存储库房" & vbCrLf & "所以该设置会同时应用于存储库房列！" & vbCrLf & "是否确定应用?", vbYesNo + vbExclamation + vbDefaultButton2, "提示") = 6 Then
                For i = 1 To .Rows - 1
                    .TextMatrix(i, .Col) = .TextMatrix(.Row, .Col)
                    .TextMatrix(i, .Col + 1) = .TextMatrix(.Row, .Col + 1)
                    .TextMatrix(i, .Col - 2) = .TextMatrix(.Row, .Col - 2)
                    .TextMatrix(i, .Col - 1) = .TextMatrix(.Row, .Col - 1)
                    
                    If mint状态 <> 1 Then   '规格
                        If .Col = mSpecColumn.规格_收入项目 Then
                            .TextMatrix(i, mSpecColumn.规格_收入项目id) = .TextMatrix(.Row, mSpecColumn.规格_收入项目id)
                        End If
                    End If
                
                    .Cell(flexcpForeColor, i, .Col) = mlngApplyColor
                    .Cell(flexcpFontSize, i, .Col) = 10
                    .Cell(flexcpFontBold, i, .Col) = True
                    .Cell(flexcpForeColor, i, .Col + 1) = mlngApplyColor
                    .Cell(flexcpFontSize, i, .Col + 1) = 10
                    .Cell(flexcpFontBold, i, .Col + 1) = True
                    .Cell(flexcpForeColor, i, .Col - 2) = mlngApplyColor
                    .Cell(flexcpFontSize, i, .Col - 2) = 10
                    .Cell(flexcpFontBold, i, .Col - 2) = True
                    .Cell(flexcpForeColor, i, .Col - 1) = mlngApplyColor
                    .Cell(flexcpFontSize, i, .Col - 1) = 10
                    .Cell(flexcpFontBold, i, .Col - 1) = True
                Next
            Else
                Exit Sub
            End If
                    
        Else
            For i = 1 To .Rows - 1
                If .Cell(flexcpBackColor, i) <> mlngColor Then '只有在背景颜色不是灰色的情况下才能进行设置
    
                    If .Col = mSpecColumn.规格_存储库房 Then
                        .TextMatrix(i, .Col) = .TextMatrix(.Row, .Col)
                        .TextMatrix(i, .Col + 1) = .TextMatrix(.Row, .Col + 1)
                        
                        mint行号 = i
                        mint列号 = vsfDetails.Col
                        Call InitDepartment(vsfDetails.TextMatrix(i, vsfDetails.Col), vsfDetails.TextMatrix(i, vsfDetails.Col + 1), vsfDetails.TextMatrix(i, vsfDetails.Col + 2), vsfDetails.TextMatrix(i, vsfDetails.Col + 3))
                                                     
                        If mint状态 <> 1 Then   '规格
                            If .Col = mSpecColumn.规格_收入项目 Then
                                .TextMatrix(i, mSpecColumn.规格_收入项目id) = .TextMatrix(.Row, mSpecColumn.规格_收入项目id)
                            End If
                        End If
                    
                        .Cell(flexcpForeColor, i, .Col) = mlngApplyColor
                        .Cell(flexcpFontSize, i, .Col) = 10
                        .Cell(flexcpFontBold, i, .Col) = True
                        .Cell(flexcpForeColor, i, .Col + 1) = mlngApplyColor
                        .Cell(flexcpFontSize, i, .Col + 1) = 10
                        .Cell(flexcpFontBold, i, .Col + 1) = True
                        
                    Else
                        .TextMatrix(i, .Col) = .TextMatrix(.Row, .Col)
                        
                        If mint状态 <> 1 Then   '规格
                            If .Col = mSpecColumn.规格_收入项目 Then
                                .TextMatrix(i, mSpecColumn.规格_收入项目id) = .TextMatrix(.Row, mSpecColumn.规格_收入项目id)
                            End If
                        End If
                        
                        .Cell(flexcpForeColor, i, .Col) = mlngApplyColor
                        .Cell(flexcpFontSize, i, .Col) = 10
                        .Cell(flexcpFontBold, i, .Col) = True
                    End If
                End If
            Next
        End If
    End With
    
End Sub

Public Sub VsfGridColFormat(ByVal objGrid As VSFlexGrid, ByVal intCol As Integer, ByVal strColName As String, _
    ByVal lngColWidth As Long, ByVal intColAlignment As Integer, _
    Optional ByVal strColKey As String = "", Optional ByVal intFixedColAlignment As Integer = 4)
    'vsf列设置：列名，列宽，列对齐方式，固定列对齐方式（默认为居中对齐）

    With objGrid
        .TextMatrix(0, intCol) = strColName
        .ColWidth(intCol) = lngColWidth
        .ColAlignment(intCol) = intColAlignment
        .ColKey(intCol) = strColKey
        .FixedAlignment(intCol) = intFixedColAlignment
    End With
End Sub

Private Sub GetDefineSize(ByVal rsRecord As ADODB.Recordset)
    '功能：得到数据库的表字段的长度
    If mblnSetKey = False Then
        mblnSetKey = True
        With vsfDetails
            If mint状态 = 1 Then
                .ColKey(mVaricolumn.品种_通用名称) = rsRecord.Fields("通用名称").DefinedSize
                .ColKey(mVaricolumn.品种_英文名称) = rsRecord.Fields("英文名称").DefinedSize
                .ColKey(mVaricolumn.品种_拼音码) = rsRecord.Fields("拼音码").DefinedSize
                .ColKey(mVaricolumn.品种_五笔码) = rsRecord.Fields("五笔码").DefinedSize
                .ColKey(mVaricolumn.品种_处方限量) = rsRecord.Fields("处方限量").DefinedSize
                .ColKey(mVaricolumn.品种_剂量单位) = rsRecord.Fields("剂量单位").DefinedSize
                .ColKey(mVaricolumn.品种_ATCCODE) = rsRecord.Fields("ATCCODE").DefinedSize
            Else
                .ColKey(mSpecColumn.规格_药品规格) = rsRecord.Fields("药品规格").DefinedSize
                .ColKey(mSpecColumn.规格_本位码) = rsRecord.Fields("本位码").DefinedSize
                .ColKey(mSpecColumn.规格_数字码) = rsRecord.Fields("数字码").DefinedSize
                .ColKey(mSpecColumn.规格_标识码) = rsRecord.Fields("标识码").DefinedSize
                .ColKey(mSpecColumn.规格_备选码) = rsRecord.Fields("备选码").DefinedSize
                .ColKey(mSpecColumn.规格_容量) = rsRecord.Fields("容量").DefinedSize
                .ColKey(mSpecColumn.规格_商品名称) = rsRecord.Fields("商品名称").DefinedSize
                .ColKey(mSpecColumn.规格_生产商) = rsRecord.Fields("生产商").DefinedSize
                .ColKey(mSpecColumn.规格_原产地) = rsRecord.Fields("原产地").DefinedSize
                .ColKey(mSpecColumn.规格_拼音码) = rsRecord.Fields("拼音码").DefinedSize
                .ColKey(mSpecColumn.规格_五笔码) = rsRecord.Fields("五笔码").DefinedSize
                .ColKey(mSpecColumn.规格_合同单位) = rsRecord.Fields("合同单位").DefinedSize
                .ColKey(mSpecColumn.规格_批准文号) = rsRecord.Fields("批准文号").DefinedSize
                .ColKey(mSpecColumn.规格_注册商标) = rsRecord.Fields("注册商标").DefinedSize
                .ColKey(mSpecColumn.规格_售价单位) = rsRecord.Fields("售价单位").DefinedSize
                .ColKey(mSpecColumn.规格_剂量系数) = rsRecord.Fields("剂量系数").DefinedSize
                .ColKey(mSpecColumn.规格_住院单位) = rsRecord.Fields("住院单位").DefinedSize
                mintLen住院单位 = Val(rsRecord.Fields("住院单位").DefinedSize)
                .ColKey(mSpecColumn.规格_住院系数) = rsRecord.Fields("住院系数").DefinedSize
                mintLen住院系数 = Val(rsRecord.Fields("住院系数").DefinedSize)
                .ColKey(mSpecColumn.规格_门诊单位) = rsRecord.Fields("门诊单位").DefinedSize
                .ColKey(mSpecColumn.规格_门诊系数) = rsRecord.Fields("门诊系数").DefinedSize
                .ColKey(mSpecColumn.规格_药库单位) = rsRecord.Fields("药库单位").DefinedSize
                .ColKey(mSpecColumn.规格_药库系数) = rsRecord.Fields("药库系数").DefinedSize
                .ColKey(mSpecColumn.规格_送货单位) = rsRecord.Fields("送货单位").DefinedSize
                .ColKey(mSpecColumn.规格_送货包装) = rsRecord.Fields("送货包装").DefinedSize
                .ColKey(mSpecColumn.规格_申领阀值) = rsRecord.Fields("申领阀值").DefinedSize
                .ColKey(mSpecColumn.规格_采购限价) = rsRecord.Fields("采购限价").DefinedSize
                .ColKey(mSpecColumn.规格_采购扣率) = rsRecord.Fields("采购扣率").DefinedSize
                .ColKey(mSpecColumn.规格_指导售价) = rsRecord.Fields("指导售价").DefinedSize
                .ColKey(mSpecColumn.规格_成本价格) = rsRecord.Fields("成本价格").DefinedSize
                .ColKey(mSpecColumn.规格_当前售价) = rsRecord.Fields("当前售价").DefinedSize
                .ColKey(mSpecColumn.规格_保质期) = rsRecord.Fields("保质期").DefinedSize
                .ColKey(mSpecColumn.规格_标识说明) = rsRecord.Fields("标识说明").DefinedSize
                .ColKey(mSpecColumn.规格_输液注意事项) = rsRecord.Fields("输液注意事项").DefinedSize
            End If
        End With
   End If
End Sub

Public Function zlGetSymbol(strInput As String, Optional bytIsWB As Byte, Optional intOutNum As Integer = 10) As String
    '----------------------------------
    '功能：生成字符串的简码
    '入参：strInput-输入字符串；bytIsWB-是否五笔(否则为拼音)
    '出参：正确返回字符串；错误返回"-"
    '----------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String

    If bytIsWB Then
        strSql = "select zlWBcode('" & strInput & "'," & intOutNum & ") from dual"
    Else
        strSql = "select zlSpellcode('" & strInput & "'," & intOutNum & ") from dual"
    End If
    On Error GoTo ErrHand
    With rsTmp
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, "mdlCISBase", strSql)
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "zlGetSymbol")
'        Call SQLTest
        zlGetSymbol = IIf(IsNull(rsTmp.Fields(0).Value), "", rsTmp.Fields(0).Value)
    End With
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlGetSymbol = "-"
End Function

Private Sub FindGridRow(ByVal strInput As String)
    '在控件中查询指定的品种和规格

    Dim lngStart As Long, lngRows As Long
    Dim str编码 As String, str名称 As String, str简码 As String
    Dim str其他名称 As String
    Dim n As Integer
    Dim blnEnd As Boolean
    Dim lngFindRow As Long
    Dim strFindStyle As String
    Dim strTmp As String

    If strInput = "" Then Exit Sub
    '查找药品
    If strInput = mstrFind Then
        '表示查找下一条记录
        If mlngFind >= vsfDetails.Rows - 1 Then
            lngStart = 0
        Else
            lngStart = mlngFind
        End If
    Else
        '表示新的查找
        lngStart = 0
        mlngFindFirst = 0
        mstrFind = strInput

        strFindStyle = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = "0", "%", "")

        Set mrsFindName = New ADODB.Recordset

        If mint状态 = 1 Then    '品种
            gstrSql = "Select Distinct a.Id, a.编码" & _
                      "  From 诊疗项目目录 A, 诊疗项目别名 B " & _
                      " Where a.Id = b.诊疗项目id And a.类别 = [1] "
        Else    '规格
            gstrSql = "Select Distinct A.Id,A.编码 From 收费项目目录 A,收费项目别名 B" & _
                 " Where A.Id =B.收费细目id And A.类别=[1] "
        End If

        If IsNumeric(Replace(strInput, "-", "")) Then       '输入全是数字（或包含一个"-"）时只匹配编码
            gstrSql = gstrSql & " And A.编码 Like [2] Or B.简码 Like [2] And B.码类=3 "
        ElseIf zlStr.IsCharAlpha(strInput) Then          '输入全是字母时只匹配简码
        
            gstrSql = gstrSql & " And B.简码 Like [3] "
        ElseIf zlStr.IsCharChinese(strInput) Then        '输入全是汉字时只匹配名称
            gstrSql = gstrSql & " And B.名称 Like [3] "
        Else
            gstrSql = gstrSql & " And (A.编码 Like [2] Or B.名称 Like [3] Or B.简码 Like [3] )"
        End If

        gstrSql = gstrSql & " Order By A.编码 "

        If mstrNode Like "西成药*" Then
            strTmp = "5"
        ElseIf mstrNode Like "中成药*" Then
            strTmp = "6"
        Else
            strTmp = "7"
        End If

        Set mrsFindName = zlDatabase.OpenSQLRecord(gstrSql, "取匹配的药品ID", strTmp, strInput & "%", strFindStyle & strInput & "%")

        If mrsFindName.RecordCount = 0 Then Exit Sub
    End If

    '开始查找
    If mrsFindName.State <> adStateOpen Then Exit Sub
    If mrsFindName.RecordCount = 0 Then Exit Sub

    lngStart = lngStart + 1
    lngRows = vsfDetails.Rows - 1

    With mrsFindName
        If .EOF Then .MoveFirst

        Do While Not .EOF
            If mint状态 = 1 Then    '品种
                lngFindRow = vsfDetails.FindRow(!编码, lngStart, mVaricolumn.品种_药品编码, True, True)
            Else    '规格
                lngFindRow = vsfDetails.FindRow(!编码, lngStart, mSpecColumn.规格_规格编码, True, True)
            End If

            If lngFindRow > 0 Then
                vsfDetails.SetFocus
                vsfDetails.TopRow = lngFindRow
                vsfDetails.Row = lngFindRow

                mlngFind = lngFindRow

                '记录找到的第1条记录
                If mlngFindFirst = 0 Then mlngFindFirst = mlngFind

                mrsFindName.MoveNext
                Exit Do
            End If
            mrsFindName.MoveNext

            '如果到底了，则返回第1条记录
            If .EOF And lngFindRow = -1 Then
                mlngFind = mlngFindFirst
                If vsfDetails.Rows > 1 Then
                    vsfDetails.Row = 1
                End If
            End If
        Loop
    End With
End Sub

Public Function zlGetDigitSign(ByVal lngMediId As Long, ByVal strSpec As String) As String
    '-------------------------------------------------------------
    '功能：根据药品通用名称、剂型的数字标记码和规格前三位数值，产生返回药品七位码
    '入参：strSpellcode-通用名称的拼音码；strDoseCode:剂型的数字标记码, strSpec：规格数值
    '返回：药品简码
    '-------------------------------------------------------------
    Dim rsThis As New ADODB.Recordset
    Dim strSpellcode As String, strDoseCode As String
    Dim strChange As String
    Dim intLocate As Integer
    Dim strTemp As String
    Dim intCount As Integer

    gstrSql = "Select 简码 From 诊疗项目别名 where 诊疗项目id=[1] and 性质=1 and 码类=1"
    Set rsThis = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngMediId)

    If rsThis.RecordCount > 0 Then
        strSpellcode = IIf(IsNull(rsThis!简码), "", rsThis!简码)
    Else
        strSpellcode = ""
    End If

    gstrSql = "select P.标记码 from 药品特性 T,药品剂型 P where T.药品剂型=P.名称(+) and 药名id=[1]"
    Set rsThis = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngMediId)

    If rsThis.RecordCount > 0 Then
        strDoseCode = IIf(IsNull(rsThis!标记码), "", rsThis!标记码)
    Else
        strDoseCode = ""
    End If

    strChange = "AOEYUVBP MF DT NL GKHJQXZCSRW "

    strTemp = ""
    strSpellcode = Mid(strSpellcode, 1, 3)
    For intCount = 1 To Len(strSpellcode)
        intLocate = InStr(1, strChange, Mid(strSpellcode, intCount, 1))
        If intLocate Mod 3 = 0 Then
            intLocate = (intLocate \ 3) - 1
        Else
            intLocate = intLocate \ 3
        End If
        If intLocate <> -1 Then strTemp = strTemp & CStr(intLocate)
    Next
    strTemp = strTemp & strDoseCode & Format(Val(Mid(strSpec, 1, 3)), "000")
    zlGetDigitSign = strTemp
End Function

Private Sub ExitFrom()
    '退出时过程
    '判断界面中是否有值刚被修改了
    Dim i As Integer
    Dim j As Integer
    Dim intupdate As Integer
    Dim bln修改 As Boolean

    bln修改 = Check修改
    mintExit = 0

    If bln修改 = True Then
        intupdate = MsgBox("刚有内容被修改了，退出之前是否保存？", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName)
        If intupdate = vbYes Then
            mintExit = 2
            Call Save
            Unload Me
        Else
            Unload Me
        End If
    Else
        Unload Me
    End If
End Sub

Private Function CheckData() As Boolean
    '检查数据的合法性和完整性
    Dim i As Integer
    Dim j As Integer
    Dim intupdate As Integer
    Dim blnShowMsg As Boolean
    
    With vsfDetails
        If mint状态 = 1 Then '品种
            For i = 1 To .Rows - 1
                If .TextMatrix(i, mVaricolumn.品种_通用名称) = "" Then
                    MsgBox "基本信息页第" & i & "行通用名称不能为空，请输入！", vbInformation, gstrSysName
                    tbcDetails.Item(mVariList.基本信息).Selected = True
                    .Select i, mVaricolumn.品种_通用名称
                    Exit Function
                End If
                If .TextMatrix(i, mVaricolumn.品种_剂量单位) = "" Then
                    MsgBox "临床应用页第" & i & "行剂量单位不能为空，请输入！", vbInformation, gstrSysName
                    tbcDetails.Item(mVariList.临床应用).Selected = True
                    .Select i, mVaricolumn.品种_剂量单位
                    Exit Function
                End If
                For j = 2 To .Rows - 1
                    If .TextMatrix(i, mVaricolumn.品种_通用名称) = .TextMatrix(j, mVaricolumn.品种_通用名称) And i <> j Then
                        MsgBox "基本信息页第" & i & "行通用名称与第" & j & "行通用名称相同了！", vbExclamation, gstrSysName
                        tbcDetails.Item(mVariList.基本信息).Selected = True
                        .Select i, mVaricolumn.品种_通用名称
                        Exit Function
                    End If
                Next
            Next
        Else    '规格
            For i = 1 To .Rows - 1
                If .TextMatrix(i, mSpecColumn.规格_药品规格) = "" Then
                    MsgBox "基本信息页第" & i & "行药品规格不能为空，请输入！", vbInformation, gstrSysName
                    tbcDetails.Item(0).Selected = True
                    .Select i, mSpecColumn.规格_药品规格
                    Exit Function
                End If
                If LenB(StrConv(.TextMatrix(i, mSpecColumn.规格_生产商), vbFromUnicode)) > Val(.ColKey(mSpecColumn.规格_生产商)) Then
                    MsgBox "商品信息页第" & i & "行药品生产商最多个" & Val(.ColKey(mSpecColumn.规格_生产商)) & "字符或" & Int(Val(.ColKey(mSpecColumn.规格_生产商)) / 2) & "个汉字！", vbInformation, gstrSysName
                    tbcDetails.Item(1).Selected = True
                    .Select i, mSpecColumn.规格_生产商
                    Exit Function
                End If
                If LenB(StrConv(.TextMatrix(i, mSpecColumn.规格_原产地), vbFromUnicode)) > Val(.ColKey(mSpecColumn.规格_原产地)) Then
                    MsgBox "商品信息页第" & i & "行药品原产地最多个" & Val(.ColKey(mSpecColumn.规格_原产地)) & "字符或" & Int(Val(.ColKey(mSpecColumn.规格_原产地)) / 2) & "个汉字！", vbInformation, gstrSysName
                    tbcDetails.Item(1).Selected = True
                    .Select i, mSpecColumn.规格_原产地
                    Exit Function
                End If
                If Val(.TextMatrix(i, mSpecColumn.规格_剂量系数)) > 100000 Then
                    MsgBox "包装单位页第" & i & "行剂量系数过大，请重新输入！", vbInformation, gstrSysName
                    tbcDetails.Item(2).Selected = True
                    .Select i, mSpecColumn.规格_剂量系数
                    Exit Function
                End If
                If Val(.TextMatrix(i, mSpecColumn.规格_门诊系数)) > 100000 Then
                    MsgBox "包装单位页第" & i & "行门诊系数过大，请重新输入！", vbInformation, gstrSysName
                    tbcDetails.Item(2).Selected = True
                    .Select i, mSpecColumn.规格_门诊系数
                    Exit Function
                End If
                If Val(.TextMatrix(i, mSpecColumn.规格_住院系数)) > 100000 Then
                    MsgBox "包装单位页第" & i & "行住院系数过大，请重新输入！", vbInformation, gstrSysName
                    tbcDetails.Item(2).Selected = True
                    .Select i, mSpecColumn.规格_住院系数
                    Exit Function
                End If
                If Val(.TextMatrix(i, mSpecColumn.规格_药库系数)) > 100000 Then
                    MsgBox "包装单位页第" & i & "行药库系数过大，请重新输入！", vbInformation, gstrSysName
                    tbcDetails.Item(2).Selected = True
                    .Select i, mSpecColumn.规格_药库系数
                    Exit Function
                End If
                If Val(.TextMatrix(i, mSpecColumn.规格_申领阀值)) > 100000 Then
                    MsgBox "包装单位页第" & i & "行申领阀值过大，请重新输入！", vbInformation, gstrSysName
                    tbcDetails.Item(2).Selected = True
                    .Select i, mSpecColumn.规格_申领阀值
                    Exit Function
                End If
                If Val(.TextMatrix(i, mSpecColumn.规格_采购限价)) > 100000 Then
                    MsgBox "价格信息页第" & i & "行采购限价过大，请重新输入！", vbInformation, gstrSysName
                    tbcDetails.Item(3).Selected = True
                    .Select i, mSpecColumn.规格_采购限价
                    Exit Function
                End If
                If Val(.TextMatrix(i, mSpecColumn.规格_指导售价)) > 100000 Then
                    MsgBox "价格信息页第" & i & "行指导售价过大，请重新输入！", vbInformation, gstrSysName
                    tbcDetails.Item(3).Selected = True
                    .Select i, mSpecColumn.规格_指导售价
                    Exit Function
                End If
                If .TextMatrix(i, mSpecColumn.规格_采购限价) = "" Then
                    MsgBox "价格信息页第" & i & "行采购限价不能为空，请输入！", vbInformation, gstrSysName
                    tbcDetails.Item(3).Selected = True
                    .Select i, mSpecColumn.规格_采购限价
                    Exit Function
                End If
                If .TextMatrix(i, mSpecColumn.规格_指导售价) = "" Then
                    MsgBox "价格信息页第" & i & "行指导售价不能为空，请输入！", vbInformation, gstrSysName
                    tbcDetails.Item(3).Selected = True
                    .Select i, mSpecColumn.规格_指导售价
                    Exit Function
                End If
                If Val(.TextMatrix(i, mSpecColumn.规格_加成率)) > 100 Then
                    MsgBox "价格信息页第" & i & "行加成率超过了最大值，请重新输入！", vbInformation, gstrSysName
                    tbcDetails.Item(3).Selected = True
                    .Select i, mSpecColumn.规格_加成率
                    Exit Function
                End If
                If CheckUnit(i) = False Then
                    Exit Function
                End If
'                If .TextMatrix(i, mSpecColumn.规格_零差价管理) = "√" And blnShowMsg = False Then
'                    If Val(zlDatabase.GetPara(275, glngSys, , 0)) > 0 Then
'                        If CheckPriceAdjust(Val(.TextMatrix(i, mSpecColumn.规格_id)), 0, -1, True) = False Then
'                            blnShowMsg = True
''                            MsgBox "[" & .TextMatrix(i, mSpecColumn.规格_规格编码) & "]" & .TextMatrix(i, mSpecColumn.规格_通用名称) & "已启用零差价管理，但售价和成本价不一致，请注意调价！", vbInformation, gstrSysName
'                            MsgBox "部分药品已启用零差价管理，但售价和成本价不一致，请注意调价！", vbInformation, gstrSysName
'                        End If
'                    End If
'                End If
            Next
        End If
    End With
    CheckData = True
End Function

Private Function CheckUnit(ByVal intRow As Integer) As Boolean
    Dim intOut As Integer, intIN As Integer
    Dim arr单位, arr系数
    Dim str单位 As String, str系数 As String
    Dim str单位_Tmp As String, str系数_Tmp As String
    Dim int位置 As Integer
    Dim strTemp As String

    With vsfDetails
        '检查是否存在单位名称一样，但系数不一致的情况
        '检查是否存在系数一样，但单位名称不一样的情况
        If mstrNode Like "中草药*" Then
            str单位 = .TextMatrix(intRow, mSpecColumn.规格_售价单位) & "|" & .TextMatrix(intRow, mSpecColumn.规格_住院单位) & "|" & .TextMatrix(intRow, mSpecColumn.规格_药库单位)
            str系数 = .TextMatrix(intRow, mSpecColumn.规格_剂量系数) & "|" & .TextMatrix(intRow, mSpecColumn.规格_住院系数) & "|" & .TextMatrix(intRow, mSpecColumn.规格_药库系数)
        Else
            str单位 = .TextMatrix(intRow, mSpecColumn.规格_售价单位) & "|" & .TextMatrix(intRow, mSpecColumn.规格_住院单位) & "|" & .TextMatrix(intRow, mSpecColumn.规格_门诊单位) & "|" & .TextMatrix(intRow, mSpecColumn.规格_药库单位)
            str系数 = .TextMatrix(intRow, mSpecColumn.规格_剂量系数) & "|" & .TextMatrix(intRow, mSpecColumn.规格_住院系数) & "|" & .TextMatrix(intRow, mSpecColumn.规格_门诊系数) & "|" & .TextMatrix(intRow, mSpecColumn.规格_药库系数)
        End If

        '考虑到其他单位可能与售价单位一致，但系数肯定不一致，所以必须分开判断
        '除售价单位外的检查
        For intOut = 2 To IIf(mstrNode Like "中草药*" = True, 3, 4)
            If mstrNode Like "中草药*" Then
                str单位_Tmp = IIf(intOut = 2, .TextMatrix(intRow, mSpecColumn.规格_住院单位), .TextMatrix(intRow, mSpecColumn.规格_药库单位))
                str系数_Tmp = Val(IIf(intOut = 2, .TextMatrix(intRow, mSpecColumn.规格_住院系数), .TextMatrix(intRow, mSpecColumn.规格_药库系数)))
            Else
                str单位_Tmp = IIf(intOut = 1, .TextMatrix(intRow, mSpecColumn.规格_售价单位), IIf(intOut = 2, .TextMatrix(intRow, mSpecColumn.规格_住院单位), IIf(intOut = 3, .TextMatrix(intRow, mSpecColumn.规格_门诊单位), .TextMatrix(intRow, mSpecColumn.规格_药库单位))))
                str系数_Tmp = Val(IIf(intOut = 1, .TextMatrix(intRow, mSpecColumn.规格_剂量系数), IIf(intOut = 2, .TextMatrix(intRow, mSpecColumn.规格_住院系数), IIf(intOut = 3, .TextMatrix(intRow, mSpecColumn.规格_门诊系数), .TextMatrix(intRow, mSpecColumn.规格_药库系数)))))
            End If
            arr单位 = Split(str单位, "|")
            arr系数 = Split(str系数, "|")
            For intIN = 2 To IIf(mstrNode Like "中草药*" = True, 3, 4)
                If intIN <> intOut Then
                    '单位相同系数不同
                    If str单位_Tmp = arr单位(intIN - 1) And (Val(str系数_Tmp) <> Val(arr系数(intIN - 1))) Then
                        If mstrNode Like "中草药*" Then
                            strTemp = IIf(intOut = 2, "药房", "药库") & "单位与" & IIf(intIN = 2, "药房", "药库") & "单位一致，但其系数却不相同，请检查！"
                        Else
                            strTemp = IIf(intOut = 2, "住院", IIf(intOut = 3, "门诊", "药库")) & "单位与" & IIf(intIN = 2, "住院", IIf(intIN = 3, "门诊", "药库")) & "单位一致，但其系数却不相同，请检查！"
                        End If

                        MsgBox strTemp, vbInformation, gstrSysName
                        tbcDetails.Item(2).Selected = True
                        If InStr(1, strTemp, "单位与住院") > 0 Then
                            int位置 = mSpecColumn.规格_住院单位
                        ElseIf InStr(1, strTemp, "单位与门诊") > 0 Then
                            int位置 = mSpecColumn.规格_门诊单位
                        ElseIf InStr(1, strTemp, "单位与药库") > 0 Then
                            int位置 = mSpecColumn.规格_药库单位
                        ElseIf InStr(1, strTemp, "药房单位一致") > 0 Then
                            int位置 = mSpecColumn.规格_住院单位
                        ElseIf InStr(1, strTemp, "药库单位一致") > 0 Then
                            int位置 = mSpecColumn.规格_药库单位
                        End If

                        .Select intRow, int位置
                        Exit Function
                    End If
                    If str单位_Tmp <> arr单位(intIN - 1) And (Val(str系数_Tmp) = Val(arr系数(intIN - 1))) Then
                        If mstrNode Like "中草药*" Then
                            strTemp = IIf(intOut = 2, "药房", "药库") & "包装与" & IIf(intIN = 2, "药房", "药库") & "包装一致，但其单位却不相同，请检查！"
                        Else
                            strTemp = IIf(intOut = 2, "住院", IIf(intOut = 3, "门诊", "药库")) & "包装与" & IIf(intIN = 2, "住院", IIf(intIN = 3, "门诊", "药库")) & "包装一致，但其单位却不相同，请检查！"
                        End If

                        MsgBox strTemp, vbInformation, gstrSysName
                        tbcDetails.Item(2).Selected = True

                        If InStr(1, strTemp, "包装与住院") > 0 Then
                            int位置 = mSpecColumn.规格_住院单位
                        ElseIf InStr(1, strTemp, "包装与门诊") > 0 Then
                            int位置 = mSpecColumn.规格_门诊单位
                        ElseIf InStr(1, strTemp, "包装与药库") > 0 Then
                            int位置 = mSpecColumn.规格_药库单位
                        ElseIf InStr(1, strTemp, "药房包装一致") > 0 Then
                            int位置 = mSpecColumn.规格_住院单位
                        ElseIf InStr(1, strTemp, "药库包装一致") > 0 Then
                            int位置 = mSpecColumn.规格_药库单位
                        End If
                        .Select intRow, int位置
                        Exit Function
                    End If
                End If
            Next
        Next

        '避免其它单位与售价单位相同，但系数不为1的情况
        '各单位与售价单位进行检查
        For intOut = 2 To IIf(mstrNode Like "中草药*" = True, 3, 4)
            If mstrNode Like "中草药*" Then
                str单位_Tmp = IIf(intOut = 2, .TextMatrix(intRow, mSpecColumn.规格_住院单位), .TextMatrix(intRow, mSpecColumn.规格_药库单位))
                str系数_Tmp = Val(IIf(intOut = 2, .TextMatrix(intRow, mSpecColumn.规格_住院系数), .TextMatrix(intRow, mSpecColumn.规格_药库系数)))
            Else
                str单位_Tmp = IIf(intOut = 1, .TextMatrix(intRow, mSpecColumn.规格_售价单位), IIf(intOut = 2, .TextMatrix(intRow, mSpecColumn.规格_住院单位), IIf(intOut = 3, .TextMatrix(intRow, mSpecColumn.规格_门诊单位), .TextMatrix(intRow, mSpecColumn.规格_药库单位))))
                str系数_Tmp = Val(IIf(intOut = 1, .TextMatrix(intRow, mSpecColumn.规格_剂量系数), IIf(intOut = 2, .TextMatrix(intRow, mSpecColumn.规格_住院系数), IIf(intOut = 3, .TextMatrix(intRow, mSpecColumn.规格_门诊系数), .TextMatrix(intRow, mSpecColumn.规格_药库系数)))))
            End If

            If str单位_Tmp = .TextMatrix(intRow, mSpecColumn.规格_售价单位) And Val(str系数_Tmp) <> 1 Then
                If mstrNode Like "中草药*" Then
                    strTemp = IIf(intOut = 2, "药房", "药库") & "单位与售价单位一致，" & IIf(intOut = 2, "药房", "药库") & "系数应该为1"
                Else
                    strTemp = IIf(intOut = 2, "住院", IIf(intOut = 3, "门诊", "药库")) & "单位与售价单位一致，" & IIf(intOut = 2, "住院", IIf(intOut = 3, "门诊", "药库")) & "系数应该为1"
                End If
                MsgBox strTemp, vbInformation, gstrSysName
                tbcDetails.Item(2).Selected = True

                If InStr(1, strTemp, "住院系数") > 0 Then
                    int位置 = mSpecColumn.规格_住院单位
                ElseIf InStr(1, strTemp, "门诊系数") > 0 Then
                    int位置 = mSpecColumn.规格_门诊单位
                ElseIf InStr(1, strTemp, "药库系数") > 0 Then
                    int位置 = mSpecColumn.规格_药库单位
                ElseIf InStr(1, strTemp, "药房系数") > 0 Then
                    int位置 = mSpecColumn.规格_住院单位
                ElseIf InStr(1, strTemp, "药库系数") > 0 Then
                    int位置 = mSpecColumn.规格_药库单位
                End If
                .Select intRow, int位置
                Exit Function
            End If
        Next

    End With
    CheckUnit = True
End Function

'Private Sub ShowPercent(sngPercent As Single)
''功能:在状态条上根据百分比显示当前处理进度(█)
'    Dim intAll As Integer
'    intAll = stbThis.Panels(2).Width / TextWidth("█") - 4
'    stbThis.Panels(2).Text = Format(sngPercent, "0% ") & String(intAll * sngPercent, "█")
'End Sub

Private Function Check修改() As Boolean
    '判断界面中是否有值刚被修改了
    '返回值为true 已经修改了 否者未修改
    Dim i As Integer
    Dim j As Integer

    With vsfDetails
        Check修改 = False
        For i = 1 To .Rows - 1
            For j = 1 To vsfDetails.Cols - 1
                If .Cell(flexcpForeColor, i, j) = mlngApplyColor Or .Cell(flexcpFontSize, i, j) = 10 Or .Cell(flexcpFontBold, i, j) = True Or .Cell(flexcpBackColor, i, j) = mlngApplyColor Then
                    Check修改 = True
                    Exit Function
                End If
            Next
        Next
    End With
End Function

Private Sub MoveRowCol()
    '行列移动方法
    With vsfDetails
        If mint状态 = 1 Then    '品种
            If mstrNode Like "中草药*" Then
                If tbcDetails.Selected.Index = mVariList.基本信息 Then    '基本页面
                    If .Col = mVaricolumn.品种_五笔码 Then
                        tbcDetails.Item(mVariList.品种属性).Selected = True
                        .SetFocus
                        .Col = mVaricolumn.品种_毒理分类
                    Else
                        .Col = .Col + 1
                    End If
                ElseIf tbcDetails.Selected.Index = mVariList.品种属性 Then    '品种属性
                    If .Col = mVaricolumn.品种_辅助用药 Then
                        tbcDetails.Item(mVariList.临床应用).Selected = True
                        .SetFocus
                        .Col = mVaricolumn.品种_处方职务
                    ElseIf .Col = mVaricolumn.品种_通用名称 Then
                        .Col = mVaricolumn.品种_毒理分类
                    ElseIf .Col = mVaricolumn.品种_药品类型 Then
                        .Col = mVaricolumn.品种_原料药
                    Else
                        .Col = .Col + 1
                    End If
                ElseIf tbcDetails.Selected.Index = mVariList.临床应用 Then    '临床应用
                    If .Col = mVaricolumn.品种_剂量单位 And .Row <> .Rows - 1 Then
                        tbcDetails.Item(mVariList.基本信息).Selected = True
                        .SetFocus
                        .Row = .Row + 1
                        .Col = mVaricolumn.品种_通用名称
                    Else
                        If .Col = mVaricolumn.品种_通用名称 Then
                            .Col = mVaricolumn.品种_参考项目
                        Else
                            If .Col <> mVaricolumn.品种_剂量单位 Then
                                .Col = .Col + 1
                            End If
                        End If
                    End If
                End If
            Else    '西成药、中成药
                If tbcDetails.Selected.Index = mVariList.基本信息 Then    '基本页面
                    If .Col = mVaricolumn.品种_五笔码 Then
                        tbcDetails.Item(mVariList.品种属性).Selected = True
                        .SetFocus
                        .Col = mVaricolumn.品种_毒理分类
                    Else
                        .Col = .Col + 1
                    End If
                ElseIf tbcDetails.Selected.Index = mVariList.品种属性 Then    '品种属性
                    If .Col = mVaricolumn.品种_ATCCODE Then
                        tbcDetails.Item(mVariList.临床应用).Selected = True
                        .SetFocus
                        .Col = mVaricolumn.品种_处方职务
                    ElseIf .Col = mVaricolumn.品种_通用名称 Then
                        .Col = mVaricolumn.品种_毒理分类
                    ElseIf .Col = mVaricolumn.品种_原料药 Then
                        .Col = mVaricolumn.品种_辅助用药
                    Else
                        .Col = .Col + 1
                    End If
                ElseIf tbcDetails.Selected.Index = mVariList.临床应用 Then    '临床应用
                    If .Col = mVaricolumn.品种_品种下长期医嘱 And .Row <> .Rows - 1 Then
                        tbcDetails.Item(mVariList.基本信息).Selected = True
                        .SetFocus
                        .Row = .Row + 1
                        .Col = mVaricolumn.品种_通用名称
                    Else
                        If .Col = mVaricolumn.品种_通用名称 Then
                            .Col = mVaricolumn.品种_参考项目
                        Else
                            If .Col <> mVaricolumn.品种_品种下长期医嘱 Then
                                .Col = .Col + 1
                            End If
                        End If
                    End If
                End If
            End If
        Else    '规格
            If mstrNode Like "中草药*" Then '中草药
                If tbcDetails.Selected.Index = mSpecList.基本信息 Then
                    If .Col = mSpecColumn.规格_备选码 Then
                        tbcDetails.Item(mSpecList.商品信息).Selected = True
                        .SetFocus
                        .Col = mSpecColumn.规格_生产商
                    Else
                        .Col = .Col + 1
                    End If
                ElseIf tbcDetails.Selected.Index = mSpecList.商品信息 Then
                    If .Col = mSpecColumn.规格_非常备药 Then
                        tbcDetails.Item(mSpecList.包装单位).Selected = True
                        .SetFocus
                        .Col = mSpecColumn.规格_售价单位
                    ElseIf .Col = mSpecColumn.规格_药品规格 Then
                        .Col = mSpecColumn.规格_生产商
                    ElseIf .Col = mSpecColumn.规格_注册商标 Then
                        .Col = mSpecColumn.规格_非常备药
                    ElseIf .Col = mSpecColumn.规格_来源分类 Then
                        .Col = mSpecColumn.规格_合同单位
                    Else
                        .Col = .Col + 1
                    End If
                ElseIf tbcDetails.Selected.Index = mSpecList.包装单位 Then
                    If .Col = mSpecColumn.规格_中药形态 Then
                        tbcDetails.Item(mSpecList.价格信息).Selected = True
                        .SetFocus
                        .Col = mSpecColumn.规格_药价属性
                    ElseIf .Col = mSpecColumn.规格_药品规格 Then
                        .Col = mSpecColumn.规格_售价单位
                    ElseIf .Col = mSpecColumn.规格_住院系数 Then
                        .Col = mSpecColumn.规格_药库单位
                    ElseIf .Col = mSpecColumn.规格_药库系数 Then
                        .Col = mSpecColumn.规格_申领单位
                    Else
                        .Col = .Col + 1
                    End If
                ElseIf tbcDetails.Selected.Index = mSpecList.价格信息 Then
                    If .Col = mSpecColumn.规格_加成率 Then
                        tbcDetails.Item(mSpecList.药价属性).Selected = True
                        .SetFocus
                        .Col = mSpecColumn.规格_收入项目
                    ElseIf .Col = mSpecColumn.规格_药品规格 Then
                        .Col = mSpecColumn.规格_药价属性
                    Else
                        .Col = .Col + 1
                    End If
                ElseIf tbcDetails.Selected.Index = mSpecList.药价属性 Then
                    If .Col = mSpecColumn.规格_医保类型 Then
                        tbcDetails.Item(mSpecList.分批管理).Selected = True
                        .SetFocus
                        .Col = mSpecColumn.规格_药库分批
                    ElseIf .Col = mSpecColumn.规格_药品规格 Then
                        .Col = mSpecColumn.规格_收入项目
                    Else
                        .Col = .Col + 1
                    End If
                ElseIf tbcDetails.Selected.Index = mSpecList.分批管理 Then
                    If .Col = mSpecColumn.规格_药房分批 Then
                        tbcDetails.Item(mSpecList.临床应用).Selected = True
                        .SetFocus
                        .Col = mSpecColumn.规格_标识说明
                    ElseIf .Col = mSpecColumn.规格_药品规格 Then
                        .Col = mSpecColumn.规格_药库分批
                    Else
                        .Col = .Col + 1
                    End If
                ElseIf tbcDetails.Selected.Index = mSpecList.临床应用 Then
                    If .Col <> mSpecColumn.规格_是否摆药 Then
                        If .Col = mSpecColumn.规格_药品规格 Then
                            .Col = mSpecColumn.规格_标识说明
                        ElseIf .Col = mSpecColumn.规格_住院分零使用 Then
                            .Col = mSpecColumn.规格_门诊分零使用
                        Else
                            .Col = .Col + 1
                        End If
                    Else
                        If .Row <> .Rows - 1 Then
                            tbcDetails.Item(mSpecList.基本信息).Selected = True
                            .SetFocus
                            .Row = .Row + 1
                            .Col = mSpecColumn.规格_药品规格
                        End If
                    End If
                End If
            Else    '西成药，中成药
                If tbcDetails.Selected.Index = mSpecList.基本信息 Then
                    If .Col = mSpecColumn.规格_容量 Then
                        tbcDetails.Item(mSpecList.商品信息).Selected = True
                        .SetFocus
                        .Col = mSpecColumn.规格_商品名称
                    Else
                        .Col = .Col + 1
                    End If
                ElseIf tbcDetails.Selected.Index = mSpecList.商品信息 Then
                    If .Col = mSpecColumn.规格_非常备药 Then
                        tbcDetails.Item(mSpecList.包装单位).Selected = True
                        .SetFocus
                        .Col = mSpecColumn.规格_售价单位
                    ElseIf .Col = mSpecColumn.规格_药品规格 Then
                        .Col = mSpecColumn.规格_商品名称
                    Else
                        If .Col = mSpecColumn.规格_生产商 Then
                            .Col = .Col + 2
                        Else
                            .Col = .Col + 1
                        End If
                    End If
                ElseIf tbcDetails.Selected.Index = mSpecList.包装单位 Then
                    If .Col = mSpecColumn.规格_申领阀值 Then
                        tbcDetails.Item(mSpecList.价格信息).Selected = True
                        .SetFocus
                        .Col = mSpecColumn.规格_药价属性
                    ElseIf .Col = mSpecColumn.规格_药品规格 Then
                        .Col = mSpecColumn.规格_售价单位
                    Else
                        .Col = .Col + 1
                    End If
                ElseIf tbcDetails.Selected.Index = mSpecList.价格信息 Then
                    If .Col = mSpecColumn.规格_加成率 Then
                        tbcDetails.Item(mSpecList.药价属性).Selected = True
                        .SetFocus
                        .Col = mSpecColumn.规格_收入项目
                    ElseIf .Col = mSpecColumn.规格_药品规格 Then
                            .Col = mSpecColumn.规格_药价属性
                    Else
                        .Col = .Col + 1
                    End If
                ElseIf tbcDetails.Selected.Index = mSpecList.药价属性 Then
                    If .Col = mSpecColumn.规格_医保类型 Then
                        tbcDetails.Item(mSpecList.分批管理).Selected = True
                        .SetFocus
                        .Col = mSpecColumn.规格_药库分批
                    ElseIf .Col = mSpecColumn.规格_药品规格 Then
                            .Col = mSpecColumn.规格_收入项目
                    Else
                        .Col = .Col + 1
                    End If
                ElseIf tbcDetails.Selected.Index = mSpecList.分批管理 Then
                    If .Col = mSpecColumn.规格_保质期 Then
                        tbcDetails.Item(mSpecList.临床应用).Selected = True
                        .SetFocus
                        .Col = mSpecColumn.规格_标识说明
                    ElseIf .Col = mSpecColumn.规格_药品规格 Then
                            .Col = mSpecColumn.规格_药库分批
                    Else
                        .Col = .Col + 1
                    End If
                ElseIf tbcDetails.Selected.Index = mSpecList.临床应用 Then
                    If .Col = mSpecColumn.规格_药品规格 Then
                        .Col = mSpecColumn.规格_标识说明
                        Exit Sub
                    End If
                    If .Col <> mSpecColumn.规格_高危药品 Then
                        .Col = .Col + 1
                    Else
                        If mint配置中心 <> 0 Then   '启用了输液配置中心
                            tbcDetails.Item(mSpecList.配药属性).Selected = True
                            .SetFocus
                            .Col = mSpecColumn.规格_存储温度
                        Else
                            If .Row <> .Rows - 1 Then
                                tbcDetails.Item(mSpecList.基本信息).Selected = True
                                .SetFocus
                                .Row = .Row + 1
                                .Col = mSpecColumn.规格_药品规格
                            End If
                        End If
                    End If
                ElseIf tbcDetails.Selected.Index = mSpecList.配药属性 Then
                    If .Col = mSpecColumn.规格_药品规格 Then
                        .Col = mSpecColumn.规格_存储温度
                        Exit Sub
                    End If
                    If .Col <> mSpecColumn.规格_输液注意事项 Then
                        .Col = .Col + 1
                    Else
                        If .Row <> .Rows - 1 Then
                            tbcDetails.Item(mSpecList.基本信息).Selected = True
                            .SetFocus
                            .Row = .Row + 1
                            .Col = mSpecColumn.规格_药品规格
                        End If
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub FindChange()
    '过滤品种已修改的药品信息
    Dim i As Integer
    Dim j As Integer
    Dim blnChange As Boolean
    
    For i = 1 To vsfDetails.Rows - 1
         vsfDetails.RowHidden(i) = False
    Next
    If chk显示所有已修改的药品.Value = 1 Then
        With vsfDetails
            For i = 1 To .Rows - 1
                blnChange = False
                For j = 1 To .Cols - 1
                    If .Cell(flexcpForeColor, i, j) = mlngApplyColor Or .Cell(flexcpFontSize, i, j) = 10 Or .Cell(flexcpFontBold, i, j) = True Or .Cell(flexcpBackColor, i, j) = mlngApplyColor Then
                        blnChange = True
                        Exit For
                    End If
                Next
                If blnChange = False Then .RowHidden(i) = True
            Next
            .SetFocus
        End With
    End If
End Sub

Private Sub FindChangeCell()
    '查找定位到已修改单元格
    Dim arrPos As Variant
    Dim intRow As Integer
    Dim intCol As Integer
    Dim i As Integer
    
    With vsfDetails
        If mstrChangedCell = "" Then
            For intRow = 1 To .Rows - 1
                For intCol = 1 To vsfDetails.Cols - 1
                    If .Cell(flexcpForeColor, intRow, intCol) = mlngApplyColor Or .Cell(flexcpFontSize, intRow, intCol) = 10 Or .Cell(flexcpFontBold, intRow, intCol) = True Or .Cell(flexcpBackColor, intRow, intCol) = mlngApplyColor Then
                        mstrChangedCell = mstrChangedCell & intRow & "," & intCol & "|"
                    End If
                Next
            Next
        End If
        
        If mstrChangedCell = "" Then Exit Sub
        arrPos = Split(mstrChangedCell, "|")
        i = mintPos
        
        If mint状态 = 1 Then '品种
            Select Case Split(arrPos(i), ",")(1)
                Case mVaricolumn.品种_英文名称 To mVaricolumn.品种_五笔码
                    tbcDetails.Item(0).Selected = True
                Case mVaricolumn.品种_毒理分类 To mVaricolumn.品种_ATCCODE
                    tbcDetails.Item(1).Selected = True
                Case mVaricolumn.品种_参考项目 To mVaricolumn.品种_参考项目ID
                    tbcDetails.Item(2).Selected = True
                Case mVaricolumn.品种_通用名称
                    tbcDetails.Item(0).Selected = True
            End Select
        Else  '规格
            Select Case Split(arrPos(i), ",")(1)
                Case mSpecColumn.规格_本位码 To mSpecColumn.规格_容量
                    tbcDetails.Item(0).Selected = True
                Case mSpecColumn.规格_商品名称 To mSpecColumn.规格_非常备药
                    tbcDetails.Item(1).Selected = True
                Case mSpecColumn.规格_售价单位 To mSpecColumn.规格_中药形态
                    tbcDetails.Item(2).Selected = True
                Case mSpecColumn.规格_药价属性 To mSpecColumn.规格_加成率
                    tbcDetails.Item(3).Selected = True
                Case mSpecColumn.规格_收入项目 To mSpecColumn.规格_医保类型
                    tbcDetails.Item(4).Selected = True
                Case mSpecColumn.规格_药库分批 To mSpecColumn.规格_保质期
                    tbcDetails.Item(5).Selected = True
                Case mSpecColumn.规格_标识说明 To mSpecColumn.规格_高危药品
                    tbcDetails.Item(6).Selected = True
                Case mSpecColumn.规格_存储温度 To mSpecColumn.规格_输液注意事项
                    tbcDetails.Item(7).Selected = True
                Case mSpecColumn.规格_存储库房 To mSpecColumn.规格_服务科室
                    tbcDetails.Item(8).Selected = True
                Case mSpecColumn.规格_药品规格
                    tbcDetails.Item(0).Selected = True
            End Select
        End If
        
        .SetFocus
        .FocusRect = flexFocusLight
        .TopRow = Split(arrPos(i), ",")(0)
        .LeftCol = Split(arrPos(i), ",")(1)
        .Row = Split(arrPos(i), ",")(0)
        .Col = Split(arrPos(i), ",")(1)
        mintPos = i + 1
        
        '如果到底了，则返回第1条记录
        If mintPos = UBound(arrPos) Then
            mintPos = 0
            mstrChangedCell = ""
        End If
        
    End With
End Sub

Private Sub chk显示所有已修改的药品_Click()
    If mint状态 = 2 Then
        Call FilterDrugs
    Else
        Call FindChange
    End If
End Sub

Private Sub cbo中药形态_Click()
    Call FilterDrugs
End Sub

Private Sub FilterDrugs()
    '过滤已修改的药品信息
    Dim i As Integer
    Dim j As Integer
    Dim blnChange As Boolean
    
    With vsfDetails
        
        For i = 1 To .Rows - 1
             .RowHidden(i) = False
        Next
        
        If chk显示所有已修改的药品.Value = 1 Then
            For i = 1 To .Rows - 1
                blnChange = False
                For j = 1 To .Cols - 1
                    If .Cell(flexcpForeColor, i, j) = mlngApplyColor Or .Cell(flexcpFontSize, i, j) = 10 Or .Cell(flexcpFontBold, i, j) = True Or .Cell(flexcpBackColor, i, j) = mlngApplyColor Then
                        blnChange = True
                        Exit For
                    End If
                Next
                If blnChange = False Then .RowHidden(i) = True
            Next
            
            If cbo中药形态.Text <> "全部形态" Then
                For i = 1 To .Rows - 1
                    If .Cell(flexcpText, i, 规格_中药形态) = cbo中药形态.Text And .RowHidden(i) = False Then
                        .RowHidden(i) = False
                    Else
                        .RowHidden(i) = True
                    End If
                Next
            End If
        Else
            If cbo中药形态.Text <> "全部形态" Then
                For i = 1 To .Rows - 1
                    If .Cell(flexcpText, i, 规格_中药形态) = cbo中药形态.Text And .RowHidden(i) = False Then
                        .RowHidden(i) = False
                    Else
                        .RowHidden(i) = True
                    End If
                Next
            End If
        End If
        
'        tbcDetails.Item(0).Selected = True
        If .Rows <> 1 Then .SetFocus
    End With
End Sub

Public Sub ShowDepartment(ByVal str库房科室 As String, ByVal str库房科室ID As String, ByVal int参数 As Integer)

    mstr库房科室 = str库房科室
    mstr库房科室ID = str库房科室ID
    
    If int参数 = 1 Then
        vsfDetails.TextMatrix(mint行号, mint列号 + 2) = mstr库房科室
        vsfDetails.TextMatrix(mint行号, mint列号 + 3) = mstr库房科室ID
    Else
        vsfDetails.TextMatrix(mint行号, mint列号) = mstr库房科室
        vsfDetails.TextMatrix(mint行号, mint列号 + 1) = mstr库房科室ID
    End If
End Sub

Public Sub ShowRoom(ByVal str存储库房 As String, ByVal str存储库房ID As String)
    mstr存储库房 = str存储库房
    mstr存储库房ID = str存储库房ID

    vsfDetails.TextMatrix(mint行号, mint列号) = mstr存储库房
    vsfDetails.TextMatrix(mint行号, mint列号 + 1) = mstr存储库房ID
    
End Sub

Private Sub MyAppend()
'创建动态纪录集
    On Error GoTo ErrHand
    If mrsMyRecords.State <> 1 Then
        With mrsMyRecords
            Call .Fields.Append("id", adDouble, 20, adFldIsNullable)
            Call .Fields.Append("存储库房", adLongVarChar, 500, adFldIsNullable)
            Call .Fields.Append("存储库房id", adLongVarChar, 500, adFldIsNullable)
            Call .Fields.Append("服务科室", adLongVarChar, 500, adFldIsNullable)
            Call .Fields.Append("库房科室id", adLongVarChar, 500, adFldIsNullable)

            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Open '打开纪录集
        End With
    End If
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ShowDisplay(ByVal rsRecord As ADODB.Recordset, ByVal intRow As Integer)
    Dim rsRoom As ADODB.Recordset
    Dim rsDepar As ADODB.Recordset
    Dim j As Integer
    Dim str库房 As String
    Dim str存储库房ID As String
    Dim str库房科室 As String
    Dim str库房科室ID As String
    Dim dbl所有库房 As Boolean
    
    On Error GoTo ErrHandle
    
    Call MyAppend
    
    If InStr(1, ";" & mstrPrivs & ";", ";所有库房;") > 0 Then dbl所有库房 = True
    
    If dbl所有库房 Then
        gstrSql = "Select  c.名称,a.执行科室id, a.开单科室id as 服务科室,a.开单科室id" & vbNewLine & _
                    "From 收费执行科室 a, 收费项目目录 b,部门表 c" & vbNewLine & _
                    "Where a.收费细目id = b.Id and a.执行科室id=c.id And b.Id = [1]" & vbNewLine & _
                    "order by a.执行科室id,a.开单科室id"
    Else
        gstrSql = "Select c.名称, a.执行科室id, a.开单科室id As 服务科室, a.开单科室id" & vbNewLine & _
                    "From 收费执行科室 A, 收费项目目录 B, 部门表 C" & vbNewLine & _
                    "Where a.收费细目id = b.Id And a.执行科室id = c.Id And b.Id = [1] And" & vbNewLine & _
                    "      c.Id In(Select 部门id From 部门人员 Where 人员id = [2])" & vbNewLine & _
                    "Order By a.执行科室id, a.开单科室id"
    End If
    
    Set rsRoom = zlDatabase.OpenSQLRecord(gstrSql, "", rsRecord!ID, UserInfo.ID)
    vsfAdditional.Clear
    Set vsfAdditional.DataSource = rsRoom
    
    gstrSql = "Select c.名称 as 服务科室 From 部门表 c Where c.Id = [1]"
    
    For j = 1 To vsfAdditional.Rows - 1
        Set rsDepar = zlDatabase.OpenSQLRecord(gstrSql, "", vsfAdditional.TextMatrix(j, 2))
    
        If Not rsDepar.EOF Then
            vsfAdditional.TextMatrix(j, 2) = rsDepar!服务科室
        End If
    Next
    
    
    str库房 = ""
    str存储库房ID = ""
    str库房科室 = ""
    str库房科室ID = ""
    If vsfAdditional.Rows = 2 Then
        str库房 = "|" & vsfAdditional.TextMatrix(1, 0)
        str存储库房ID = "!!" & vsfAdditional.TextMatrix(1, 1) & "|"
        
        If vsfAdditional.TextMatrix(1, 2) <> "" Then
            str库房科室 = "；" & vsfAdditional.TextMatrix(1, 0) & "：" & vsfAdditional.TextMatrix(1, 2)
        End If
        
        If vsfAdditional.TextMatrix(1, 2) = "" Then
            str库房科室ID = "!!" & vsfAdditional.TextMatrix(1, 1) & "|"
        Else
            str库房科室ID = "!!" & vsfAdditional.TextMatrix(1, 1) & "|" & vsfAdditional.TextMatrix(1, 3)
        End If
    Else
        For j = 2 To vsfAdditional.Rows - 1
            If j = 2 Then
                 str库房 = "|" & vsfAdditional.TextMatrix(j - 1, 0)
                 str存储库房ID = "!!" & vsfAdditional.TextMatrix(j - 1, 1) & "|"
             End If
    
             If vsfAdditional.TextMatrix(j - 1, 1) <> vsfAdditional.TextMatrix(j, 1) Then
                 str库房 = str库房 & "|" & vsfAdditional.TextMatrix(j, 0)
                 str存储库房ID = str存储库房ID & "!!" & vsfAdditional.TextMatrix(j, 1) & "|"
             End If
        Next
    
        For j = 2 To vsfAdditional.Rows - 1
             If j = 2 Then
                If vsfAdditional.TextMatrix(j - 1, 2) <> "" Then
                    str库房科室 = "；" & vsfAdditional.TextMatrix(j - 1, 0) & "：" & vsfAdditional.TextMatrix(j - 1, 2)
                End If
                
                If vsfAdditional.TextMatrix(j - 1, 2) = "" Then
                    str库房科室ID = "!!" & vsfAdditional.TextMatrix(j - 1, 1) & "|"
                Else
                    str库房科室ID = "!!" & vsfAdditional.TextMatrix(j - 1, 1) & "|" & vsfAdditional.TextMatrix(j - 1, 3)
                End If
             End If
    
             If vsfAdditional.TextMatrix(j - 1, 1) <> vsfAdditional.TextMatrix(j, 1) Then
                 If vsfAdditional.TextMatrix(j, 2) <> "" Then
                    str库房科室 = str库房科室 & "；" & vsfAdditional.TextMatrix(j, 0) & "：" & vsfAdditional.TextMatrix(j, 2)
                 End If
                 
                 If vsfAdditional.TextMatrix(j, 2) = "" Then
                     str库房科室ID = str库房科室ID & "!!" & vsfAdditional.TextMatrix(j, 1) & "|"
                 Else
                     str库房科室ID = str库房科室ID & "!!" & vsfAdditional.TextMatrix(j, 1) & "|" & vsfAdditional.TextMatrix(j, 3)
                 End If
             Else
                 str库房科室 = str库房科室 & "," & vsfAdditional.TextMatrix(j, 2)
                 If vsfAdditional.TextMatrix(j, 2) <> "" Then
                     str库房科室ID = str库房科室ID & "," & vsfAdditional.TextMatrix(j, 3)
                 End If
             End If
        Next
    End If
    
    vsfDetails.TextMatrix(intRow, mSpecColumn.规格_存储库房) = Mid(str库房, 2)
    vsfDetails.TextMatrix(intRow, mSpecColumn.规格_存储库房id) = Mid(str存储库房ID, 3)
    vsfDetails.TextMatrix(intRow, mSpecColumn.规格_服务科室) = Mid(str库房科室, 2)
    vsfDetails.TextMatrix(intRow, mSpecColumn.规格_服务科室未改) = Mid(str库房科室ID, 3)
    vsfDetails.TextMatrix(intRow, mSpecColumn.规格_库房科室id) = Mid(str库房科室ID, 3)
    
    With mrsMyRecords
        .AddNew
        .Fields!ID = vsfDetails.TextMatrix(intRow, mSpecColumn.规格_id)
        .Fields!存储库房 = vsfDetails.TextMatrix(intRow, mSpecColumn.规格_存储库房)
        .Fields!存储库房ID = vsfDetails.TextMatrix(intRow, mSpecColumn.规格_存储库房id)
        .Fields!服务科室 = vsfDetails.TextMatrix(intRow, mSpecColumn.规格_服务科室)
        .Fields!库房科室id = vsfDetails.TextMatrix(intRow, mSpecColumn.规格_库房科室id)
        .UpDate
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub InitDepartment(ByVal str存储库房 As String, ByVal str存储库房ID As String, ByVal str库房科室 As String, ByVal str库房科室ID As String)
    Dim i As Integer, j As Integer
    Dim rsRoom As New ADODB.Recordset
    Dim strArr存储库房() As String
    Dim strArr存储库房ID() As String
    Dim strArr库房科室ID() As String

    vsfAdditional.Clear
    vsfAdditional.Rows = 1
    vsfAdditional.Cols = 4
    
    VsfGridColFormat vsfAdditional, 0, "存储库房", 1000, flexAlignLeftCenter, "存储库房"
    VsfGridColFormat vsfAdditional, 1, "存储库房ID", 1000, flexAlignCenterCenter, "存储库房ID"
    VsfGridColFormat vsfAdditional, 2, "服务科室", 4000, flexAlignLeftCenter, "服务科室"
    VsfGridColFormat vsfAdditional, 3, "服务科室id", 4000, flexAlignCenterCenter, "服务科室id"
    
    strArr存储库房 = Split(str存储库房, "|")
    strArr存储库房ID = Split(str存储库房ID, "!!")
    strArr库房科室ID = Split(str库房科室ID, "!!")
    
    For i = LBound(strArr存储库房) To UBound(strArr存储库房)
        vsfAdditional.Rows = vsfAdditional.Rows + 1
        vsfAdditional.RowHeight(i + 1) = 400
        vsfAdditional.TextMatrix(i + 1, 0) = strArr存储库房(i)
    Next
     
    For i = LBound(strArr存储库房ID) To UBound(strArr存储库房ID)
        vsfAdditional.TextMatrix(i + 1, 1) = Split(strArr存储库房ID(i), "|")(0)
    Next
    
    For i = LBound(strArr库房科室ID) To UBound(strArr库房科室ID)
        For j = 1 To vsfAdditional.Rows - 1
            If Split(strArr库房科室ID(i), "|")(0) = vsfAdditional.TextMatrix(j, 1) Then
                    vsfAdditional.TextMatrix(j, 3) = Split(strArr库房科室ID(i), "|")(1)
                    
                    gstrSql = "select a.名称 from 部门表 a where a.id in(Select Column_Value From Table(f_num2list([1])))"
                    Set rsRoom = zlDatabase.OpenSQLRecord(gstrSql, "", vsfAdditional.TextMatrix(j, 3))
                    
                    Do While Not rsRoom.EOF
                        vsfAdditional.TextMatrix(j, 2) = vsfAdditional.TextMatrix(j, 2) & "," & rsRoom!名称
                        rsRoom.MoveNext
                    Loop
                    If vsfAdditional.TextMatrix(j, 2) <> "" Then
                        vsfAdditional.TextMatrix(j, 2) = Mid(vsfAdditional.TextMatrix(j, 2), 2)
                    End If
                    
                Exit For
            End If
        Next
    Next
        
    str库房科室 = ""
    str库房科室ID = ""
    
    For i = 1 To vsfAdditional.Rows - 1
        If vsfAdditional.TextMatrix(i, 2) = "" Then
            str库房科室ID = str库房科室ID & "!!" & vsfAdditional.TextMatrix(i, 1) & "|"
        End If
        
        If vsfAdditional.TextMatrix(i, 2) <> "" Then
            str库房科室 = str库房科室 & "；" & vsfAdditional.TextMatrix(i, 0) & "：" & vsfAdditional.TextMatrix(i, 2)
            str库房科室ID = str库房科室ID & "!!" & vsfAdditional.TextMatrix(i, 1) & "|" & vsfAdditional.TextMatrix(i, 3)
        End If
    Next
     
     str库房科室 = Mid(str库房科室, 2)
     str库房科室ID = Mid(str库房科室ID, 3)
     Call ShowDepartment(str库房科室, str库房科室ID, 1)
End Sub


