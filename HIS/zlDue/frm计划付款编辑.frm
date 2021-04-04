VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frm计划付款编辑 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "计划付款单"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9960
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   9960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.ImageList img32 
      Left            =   2595
      Top             =   5955
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm计划付款编辑.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   1875
      Top             =   5895
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm计划付款编辑.frx":064A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "下一步(&N)"
      Height          =   350
      Left            =   6570
      TabIndex        =   13
      Top             =   5895
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   7695
      TabIndex        =   14
      Top             =   5895
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   8805
      TabIndex        =   15
      Top             =   5895
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   150
      TabIndex        =   16
      Top             =   5910
      Width           =   1100
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   35
      Top             =   6375
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm计划付款编辑.frx":0C94
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12965
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "上一步(&B)"
      Height          =   350
      Left            =   5415
      TabIndex        =   12
      Top             =   5895
      Width           =   1100
   End
   Begin VB.PictureBox Pic 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   5685
      Index           =   0
      Left            =   0
      ScaleHeight     =   5685
      ScaleWidth      =   9930
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   105
      Width           =   9930
      Begin VB.CommandButton cmd条件 
         Caption         =   "条件重置(&R)"
         Height          =   330
         Left            =   60
         TabIndex        =   0
         Top             =   0
         Width           =   1320
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
         Height          =   2565
         Left            =   6570
         TabIndex        =   9
         ToolTipText     =   "预付款清单"
         Top             =   3090
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   4524
         _Version        =   393216
         FixedCols       =   0
         AllowBigSelection=   0   'False
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSComctlLib.ListView lvwMain 
         Height          =   2430
         Left            =   -15
         TabIndex        =   2
         Top             =   375
         Width           =   9960
         _ExtentX        =   17568
         _ExtentY        =   4286
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "img32"
         SmallIcons      =   "img16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   14
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "名称"
            Object.Tag             =   "名称"
            Text            =   "名称"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "许可证号"
            Object.Tag             =   "许可证号"
            Text            =   "许可证号"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "许可证效期"
            Object.Tag             =   "许可证效期"
            Text            =   "许可证效期"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "执照号"
            Object.Tag             =   "执照号"
            Text            =   "执照号"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Key             =   "执照效期"
            Object.Tag             =   "执照效期"
            Text            =   "执照效期"
            Object.Width           =   2893
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Key             =   "税务登记号"
            Object.Tag             =   "税务登记号"
            Text            =   "税务登记号"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Key             =   "地址"
            Object.Tag             =   "地址"
            Text            =   "地址"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Key             =   "电话"
            Object.Tag             =   "电话"
            Text            =   "电话"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Key             =   "开户银行"
            Object.Tag             =   "开户银行"
            Text            =   "开户银行"
            Object.Width           =   3598
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Key             =   "帐号"
            Object.Tag             =   "帐号"
            Text            =   "帐号"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Key             =   "联系人"
            Object.Tag             =   "联系人"
            Text            =   "联系人"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Key             =   "类型"
            Object.Tag             =   "类型"
            Text            =   "类型"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Object.Tag             =   "信用期"
            Text            =   "信用期"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Object.Tag             =   "信用额"
            Text            =   "信用额"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshMain 
         Height          =   2565
         Left            =   0
         TabIndex        =   8
         ToolTipText     =   "未付款清单"
         Top             =   3090
         Width           =   6540
         _ExtentX        =   11536
         _ExtentY        =   4524
         _Version        =   393216
         FixedCols       =   0
         AllowBigSelection=   0   'False
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lblDATE 
         AutoSize        =   -1  'True
         Caption         =   "日期范围:"
         Height          =   180
         Left            =   1500
         TabIndex        =   1
         Top             =   75
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label lbl金额 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "本次应付:"
         ForeColor       =   &H00000040&
         Height          =   180
         Index           =   5
         Left            =   4125
         TabIndex        =   7
         Top             =   2850
         Width           =   810
      End
      Begin VB.Label lbl金额 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "冲预交:"
         ForeColor       =   &H00000040&
         Height          =   180
         Index           =   4
         Left            =   8385
         TabIndex        =   6
         Top             =   2865
         Width           =   630
      End
      Begin VB.Label lbl金额 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "预交累计:"
         ForeColor       =   &H00000040&
         Height          =   180
         Index           =   3
         Left            =   6570
         TabIndex        =   4
         Top             =   2850
         Width           =   810
      End
      Begin VB.Label lbl金额 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "累计应付:"
         ForeColor       =   &H00000040&
         Height          =   180
         Index           =   1
         Left            =   15
         TabIndex        =   3
         Top             =   2850
         Width           =   810
      End
      Begin VB.Label lbl金额 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "付款金额:"
         ForeColor       =   &H00000040&
         Height          =   180
         Index           =   2
         Left            =   2100
         TabIndex        =   5
         Top             =   2850
         Width           =   810
      End
      Begin VB.Label lbl 
         BackColor       =   &H80000010&
         Height          =   270
         Index           =   0
         Left            =   0
         TabIndex        =   40
         Top             =   2805
         Width           =   9945
      End
   End
   Begin VB.PictureBox Pic 
      AutoRedraw      =   -1  'True
      Enabled         =   0   'False
      Height          =   5685
      Index           =   1
      Left            =   0
      ScaleHeight     =   5625
      ScaleWidth      =   9855
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   105
      Visible         =   0   'False
      Width           =   9915
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh预付 
         Height          =   2625
         Left            =   5205
         TabIndex        =   38
         Top             =   1500
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   4630
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   -2147483628
         FocusRect       =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   0
         Left            =   975
         TabIndex        =   11
         Top             =   4500
         Width           =   8820
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   1
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   4875
         Width           =   3240
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   2
         Left            =   6555
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   4875
         Width           =   3240
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   3
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   5235
         Width           =   3240
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   4
         Left            =   6555
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   5250
         Width           =   3240
      End
      Begin ZL9BillEdit.BillEdit mshEdit 
         Height          =   2610
         Left            =   165
         TabIndex        =   10
         Top             =   1500
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   4604
         Appearance      =   0
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483643
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483643
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "付款通知单"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   30
         TabIndex        =   24
         Top             =   90
         Width           =   9780
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "合计:"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   3
         Left            =   165
         TabIndex        =   39
         Top             =   4095
         Width           =   5055
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "本次冲预付款:"
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   2
         Left            =   5205
         TabIndex        =   37
         Top             =   4110
         Width           =   4605
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "单位名称:"
         Height          =   180
         Index           =   9
         Left            =   390
         TabIndex        =   34
         Top             =   600
         Width           =   810
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "地址电话:"
         Height          =   180
         Index           =   1
         Left            =   390
         TabIndex        =   33
         Top             =   825
         Width           =   810
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "开户银行:"
         Height          =   180
         Index           =   2
         Left            =   390
         TabIndex        =   32
         Top             =   1050
         Width           =   810
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "税务登记号:"
         Height          =   180
         Index           =   3
         Left            =   210
         TabIndex        =   31
         Top             =   1290
         Width           =   990
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "付款说明"
         Height          =   180
         Index           =   4
         Left            =   165
         TabIndex        =   30
         Top             =   4560
         Width           =   750
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "填制人"
         Height          =   180
         Index           =   5
         Left            =   345
         TabIndex        =   29
         Top             =   4935
         Width           =   570
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "填制日期"
         Height          =   180
         Index           =   6
         Left            =   5745
         TabIndex        =   28
         Top             =   4935
         Width           =   750
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "审核人"
         Height          =   180
         Index           =   7
         Left            =   345
         TabIndex        =   27
         Top             =   5310
         Width           =   570
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "审核日期"
         Height          =   180
         Index           =   8
         Left            =   5745
         TabIndex        =   26
         Top             =   5310
         Width           =   750
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "NO."
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Index           =   10
         Left            =   8055
         TabIndex        =   25
         Top             =   450
         Width           =   315
      End
      Begin VB.Label txtNo 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   8355
         TabIndex        =   23
         Top             =   390
         Width           =   1425
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "本次付款:"
         Height          =   180
         Index           =   1
         Left            =   7950
         TabIndex        =   36
         Top             =   1260
         Width           =   810
      End
   End
   Begin VB.Menu mnuIco 
      Caption         =   "弹出菜单(&P)"
      Visible         =   0   'False
      Begin VB.Menu mnuViewIcon 
         Caption         =   "大图标(&G)"
         Index           =   0
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "小图标(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "列表(&L)"
         Index           =   2
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "详细资料(&D)"
         Checked         =   -1  'True
         Index           =   3
      End
   End
   Begin VB.Menu mnuHandle 
      Caption         =   "表格操作"
      Visible         =   0   'False
      Begin VB.Menu mnuSelect 
         Caption         =   "选择(&S)"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "取消选择(&D)"
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "全部选择(&A)"
      End
      Begin VB.Menu mnuClearAll 
         Caption         =   "全部取消(&C)"
      End
   End
End
Attribute VB_Name = "frm计划付款编辑"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private msngDownY As Single, msngDownX As Single

Private mintStep As Integer

Private mstrNo As String                                '单据号
Private mlng单位ID As Long
Private mblnFirst As Boolean
Private mblnChange As Boolean
Private mblnSave As Boolean
Private mlngID As Long                                  '单据ID
Private mfrmMain  As Object

Private mEditType As gEditType
Private mint记录状态 As RecBillStatus                   '1:正常记录;2-冲销记录;3-已经冲销的原记录
Private mErrBillStatusInfor As ErrBillStatusInfor       '对于新增后单据并行执行的处理： 1、代表正常情况；2、已经删除的记录；3、已经审核的记录
Private mblnEdit As Boolean                             '编辑状态
Private mblnSuccess As Boolean                          '是否有单据保存成功
Private mstrPrivs  As String
Private mbln付款单 As Boolean                           '是否是普通的付款单,为False 是计划付款
Private mlng付款序号 As Long                            '付款序号

Private mdbl累计应付 As Long
Private mdbl本次应付 As Double
Private mdbl本次预交 As Double
Private mdbl累计预交 As Double
Private mstrSelectTag As String
Private mstrStartDate As String
Private mstrEndDate As String
Private mintPreCol As Integer
Private mintsort As Integer

Private Enum PayHeadCol
        付款方式 = 0
        付款金额
        结算号码
End Enum
Private Const mlngModule = 1322

'检查数据依赖性
Private Function GetDepend() As Boolean
    Dim strSQL As String
    Dim rsTemp As New Recordset

    GetDepend = False
    '读取结算方式
    'by lesfeng 2009-12-2 性能优化
    strSQL = "Select 应用场合,结算方式,缺省标志 From 结算方式应用 Where 应用场合='付货款' Order by 缺省标志 desc"
    Err = 0
    On Error GoTo ErrHand:
    
    zlDatabase.OpenRecordset rsTemp, strSQL, Me.Caption
    If rsTemp.RecordCount = 0 Then
        ShowMsgbox "结算方式应用信息不全,请在结算方式管理中进行设置！"
        Exit Function
    End If
    
    '初始化数据
    With rsTemp
        mshEdit.Clear
        Do While Not .EOF
                mshEdit.AddItem !结算方式
            .MoveNext
        Loop
        'mshEdit.ListIndex = 0
        .Close
    End With
    GetDepend = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub initPayGrd()
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:初始付款单表头信息
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    With mshEdit
        .Cols = 3
        .TextMatrix(0, PayHeadCol.付款方式) = "付款方式"
        .TextMatrix(0, PayHeadCol.付款金额) = "付款金额"
        .TextMatrix(0, PayHeadCol.结算号码) = "结算号码"
                
        If Not RestoreFlexState(mshEdit, Me.Caption) Then
            .ColWidth(PayHeadCol.付款方式) = 1600
            .ColWidth(PayHeadCol.付款金额) = 1200
            .ColWidth(PayHeadCol.结算号码) = 1000
        End If
        
        .ColAlignment(PayHeadCol.付款方式) = 1
        .ColAlignment(PayHeadCol.付款金额) = 7
        .ColAlignment(PayHeadCol.结算号码) = 1
        
        .ColData(PayHeadCol.付款方式) = 3
        .ColData(PayHeadCol.付款金额) = 4
        .ColData(PayHeadCol.结算号码) = 4
        .LocateCol = PayHeadCol.付款方式
        .PrimaryCol = PayHeadCol.付款方式
        .Active = True
    End With
End Sub

Private Sub Set预交列头()
    '初始冲预付情况
    With msh预付
        .Cols = 4
        .TextMatrix(0, 0) = "ID"
        .TextMatrix(0, 1) = "付款方式"
        .TextMatrix(0, 2) = "付款金额"
        .TextMatrix(0, 3) = "结算号码"
                
        If Not RestoreFlexState(msh预付, Me.Caption) Then
            .ColWidth(0) = 0
            .ColWidth(1) = 1400
            .ColWidth(2) = 1200
            .ColWidth(3) = 1000
        End If
        
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 7
        .ColAlignment(3) = 1
    End With
End Sub

Private Sub initCard()
    Dim i As Integer
    Dim strSQL As String
    Dim rsTemp As New Recordset
    Dim lngLoop As Long
    Dim itmTemp As ListItem
    Dim strTmp As String
    Dim str类型 As String
    Dim intR As Integer
    '初始表格
    Call initPayGrd
    On Error GoTo errHandle
    Select Case mEditType
        Case g新增
                txtInfo(1).Text = UserInfo.姓名
                txtInfo(2).Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd hh:mm:ss")
                txtInfo(3).Text = ""
                txtInfo(4).Text = ""
                lblDATE.Visible = True
                cmd条件.Enabled = True
        Case g审核, g修改, g查看, g取消
            lblDATE.Visible = False
            '读取付款序号
            'by lesfeng 2009-12-2 性能优化  取消 select * from 增加绑定变量
            strSQL = "Select ID,记录状态,NO,序号,预付款,单位ID,金额,结算方式,结算号码,摘要," & _
                     "       填制人,填制日期,审核人,审核日期,付款序号 " & _
                     "  From 付款记录 Where NO=[1] And 记录状态=[2] order by 序号"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrNo, mint记录状态)
            
            If rsTemp.EOF Then
                mErrBillStatusInfor = 已经删除
                Exit Sub
            End If
            mlng付款序号 = Nvl(rsTemp!付款序号, 0)
            mlng单位ID = Nvl(rsTemp!单位ID, 0)
            
            txtInfo(0).Text = Nvl(rsTemp!摘要)
            txtInfo(1).Text = Nvl(rsTemp!填制人)
            txtInfo(2).Text = Format(rsTemp!填制日期, "yyyy-MM-dd hh:mm:ss")
            txtInfo(3).Text = Nvl(rsTemp!审核人)
            txtInfo(4).Text = Format(rsTemp!审核日期, "yyyy-MM-dd hh:mm:ss")
            txtNo = Nvl(rsTemp!NO)
            If mEditType = g审核 Or mEditType = g取消 Then
                txtInfo(3).Text = UserInfo.姓名
                txtInfo(4).Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd hh:mm:ss")
            End If
                        
            With mshEdit
                .Rows = rsTemp.RecordCount + 1
                lngLoop = 1
                Do While Not rsTemp.EOF
                    .TextMatrix(lngLoop, 0) = Nvl(rsTemp!结算方式)
                    .TextMatrix(lngLoop, 1) = Format(rsTemp!金额, "###0.00;-###0.00; ;")
                    .TextMatrix(lngLoop, 2) = Nvl(rsTemp!结算号码)
                    lngLoop = lngLoop + 1
                    rsTemp.MoveNext
                Loop
            End With
            cmd条件.Enabled = False
    End Select
    
    If mlng单位ID <> 0 Then
        '如果提供了供应商ID则读取该供应商信息
        'by lesfeng 2009-12-2 性能优化  取消 select * from 绑定变量
        strSQL = "Select ID,上级ID,编码,名称,简码,末级,许可证号,许可证效期,执照号,执照效期,税务登记号,地址,电话,开户银行," & _
                  "       帐号,联系人,建档时间,撤档时间,类型,信用期,信用额,销售委托人,销售委托日期,质量认证号,质量认证日期," & _
                  "       药监局备案号,药监局备案日期,授权号,授权期,站点" & _
                  "  From 供应商 where id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng单位ID)
        
        If Not rsTemp.EOF Then
            With rsTemp
                Set itmTemp = Me.lvwMain.ListItems.Add(, "K" & !ID, Nvl(!编码) & "--" & Nvl(!名称), 1, 1)
                  i = 1
                  itmTemp.SubItems(i) = Nvl(!许可证号)
                  i = i + 1
                  itmTemp.SubItems(i) = Format(!许可证效期, "yyyy-mm-dd")
                  i = i + 1
                  itmTemp.SubItems(i) = Nvl(!执照号)
                  i = i + 1
                  itmTemp.SubItems(i) = Format(!执照效期, "yyyy-mm-dd")
                  i = i + 1
                  itmTemp.SubItems(i) = Nvl(!税务登记号)
                  i = i + 1
                  itmTemp.SubItems(i) = Nvl(!地址)
                  i = i + 1
                  itmTemp.SubItems(i) = Nvl(!电话)
                  i = i + 1
                  itmTemp.SubItems(i) = Nvl(!开户银行)
                  i = i + 1
                  itmTemp.SubItems(i) = Nvl(!帐号)
                  i = i + 1
                  itmTemp.SubItems(i) = Nvl(!联系人)
                  i = i + 1
                  strTmp = Nvl(!类型)
                  str类型 = ""
                  For intR = 1 To Len(strTmp)
                      If Mid(Nvl(!类型), intR, 1) = 1 Then
                          Select Case intR
                              Case 1
                                  str类型 = str类型 & " " & "药品"
                              Case 2
                                  str类型 = str类型 & " " & "物资"
                              Case 3
                                  str类型 = str类型 & " " & "设备"
                              Case 4
                                  str类型 = str类型 & " " & "其他"
                          End Select
                      End If
                  Next
                  itmTemp.SubItems(i) = str类型
                  i = i + 1
                  itmTemp.SubItems(i) = IIf(Nvl(!信用期, 0) = 0, "", Nvl(!信用期) & "个月")
                  i = i + 1
                  itmTemp.SubItems(i) = Format(Nvl(!信用额, 0), "####0.00;-####0.00; ;")
                  If lvwMain.SelectedItem Is Nothing Then
                      itmTemp.Selected = True
                  End If
            End With
            
            lblInfo(9).Caption = "单位名称:" & rsTemp!名称
            lblInfo(1).Caption = "地址电话:" & IIf(IsNull(rsTemp!地址), "", rsTemp!地址) & IIf(IsNull(rsTemp!地址), "", "  TEL:") & IIf(IsNull(rsTemp!电话), "", rsTemp!电话)
            lblInfo(2).Caption = "开户银行:" & IIf(IsNull(rsTemp!开户银行), "", rsTemp!开户银行)
            lblInfo(3).Caption = "税务登记号:" & IIf(IsNull(rsTemp!税务登记号), "", rsTemp!税务登记号)
        End If
    End If
    If mbln付款单 Then
        '加载数据
        Call LoadPayMoney
    Else
        '加载计划付款数据
        Call GetPlanPayMoney
    End If
    cmdBack.Enabled = False
    SetEditPro
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub ShowCard(FrmMain As Form, ByVal bln付款单 As Boolean, _
    ByVal int编辑状态 As gEditType, ByVal strPrivs As String, _
    Optional strNO As String = "", _
    Optional lng单位ID As Long = 0, _
    Optional int记录状态 As RecBillStatus = 1, _
    Optional blnSuccess As Boolean = False)
    
    mstrNo = strNO
    mbln付款单 = bln付款单
    mblnSave = False
    mblnSuccess = False
    mEditType = int编辑状态
    mint记录状态 = int记录状态
    mstrPrivs = strPrivs

    mlng单位ID = lng单位ID
    
    mblnChange = False
    mErrBillStatusInfor = 正常情况
    Set mfrmMain = FrmMain
        
    '检查数据依赖关系
    If Not GetDepend Then
        mblnChange = False
        Unload Me
        Exit Sub
    End If
     
    If mEditType = g新增 Then
        mblnEdit = True
    ElseIf mEditType = g修改 Then
        mblnEdit = True
    ElseIf mEditType = g审核 Then
        mblnEdit = False
        cmdOK.Caption = "审核(&V)"
    ElseIf mEditType = g取消 Then
        mblnEdit = False
        cmdOK.Caption = "冲销(&O)"
    ElseIf mEditType = g查看 Then
        mblnEdit = False
        cmdOK.Caption = "打印(&P)"
        If InStr(mstrPrivs, ";付款通知单;") = 0 Then
            cmdOK.Visible = False
        Else
            cmdOK.Visible = True
        End If
    End If
    lblTitle.Caption = GetUnitName & lblTitle.Caption
     Me.Show vbModal, FrmMain
    blnSuccess = mblnSuccess
End Sub

Private Sub LoadPayMoney()
    '--------------------------------------------------------------
    '功能：填充供选择的应付款数据
    '参数：
    '返回：
    '说明：
    '--------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String, strWhere As String
    Dim lngLoop As Long, lngJLoop As Long
    Dim sngAllCount As Single, sngCount As Single
    Dim lng付款序号 As Long
    
    '标志,发票号,入库单号,品名,规格,单位,数量,发票金额
    Call zlcommfun.ShowFlash("正在搜索付款记录,请稍候 ...", Me)
    
    mshMain.Redraw = False
    DoEvents
    Screen.MousePointer = vbHourglass
    
    '根据操作类型设定记录读取条件
    'by lesfeng 2009-12-2 性能优化  修改绑定变量
    lng付款序号 = mlng付款序号
    If IsNull(lng付款序号) Then lng付款序号 = 0
    If mEditType = g新增 Then
        '新增时读取付款序号为空的应付款供选择
        strWhere = " and 付款序号 Is Null"
    ElseIf mEditType = g修改 Then
        '编辑时读取付款序号为空或当前编辑的付款序号所对应的应付款
        strWhere = " And (付款序号 Is Null Or 付款序号=[2])"
    Else
        '查看或审核时仅读取当前编辑的付款单所对应的应付款
        strWhere = " And 付款序号=[2]"
    End If
    '读取应付款记录
    'lblTemp(0).Caption = "未付款发票清单"
    strSQL = "" & _
        "   Select Decode(付款序号,Null,'','√') As 标志,ID,计划日期,发票号,入库单据号," & _
        "           品名,规格,计量单位,to_char(数量,'99999999999.9999') as 数量,to_char(发票金额,'99999999999.99') as 发票金额 " & _
        "   From 应付记录 " & _
        "   Where 计划日期 Is Null AND 记录状态=1 and 审核日期 is not null And 记录性质<>-1 And 单位ID=[1]" & strWhere & _
        "   Order By 发票号"
    On Error GoTo errHandle
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng单位ID, lng付款序号)
    '初始化并填充数据
    With mshMain
        .Clear
        If rsTemp.EOF Then
            Set .Recordset = Nothing
            .Rows = 2
        Else
            Set .Recordset = rsTemp
        End If
    
        .FormatString = "^标志|||^发票号|^入库单号|^品名|^规格|^单位|^数量|^发票金额"
        .ColAlignment(0) = 4
        .ColWidth(1) = 0: .ColWidth(2) = 0
        .ColWidth(3) = 1000: .ColAlignment(3) = 1
        .ColWidth(4) = 1000: .ColAlignment(4) = 4
        .ColWidth(5) = 1800: .ColAlignment(5) = 1
        .ColWidth(6) = 800: .ColAlignment(6) = 1
        .ColWidth(7) = 800: .ColAlignment(7) = 4
        .ColWidth(8) = 1200: .ColAlignment(8) = 7
        .ColWidth(9) = 1200: .ColAlignment(9) = 7
        
        mdbl本次应付 = 0
        mdbl累计应付 = 0
        
        sngCount = 0: sngAllCount = 0
        For lngLoop = 1 To .Rows - 1
            mdbl累计应付 = mdbl累计应付 + Val(.TextMatrix(lngLoop, .Cols - 1))
            If .TextMatrix(lngLoop, 0) <> "" Then
                mdbl本次应付 = mdbl本次应付 + Val(.TextMatrix(lngLoop, .Cols - 1))
            End If
        Next
        .Row = 1: .Col = 1
        .Redraw = True
    End With
    
    Call zlcommfun.StopFlash
    Screen.MousePointer = vbDefault
    Call SetMoneyLbl
    Call Get预付数据             '读取预付款
    Call SetCmdEn
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then
        Resume
    Else
        zlcommfun.StopFlash
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub GetPlanPayMoney()
    '--------------------------------------------------------------
    '功能：按计划付款时读取并填充计划付款记录供选择
    '参数：
    '返回：
    '说明：
    '--------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim strWhere As String
    Dim lngLoop As Long, lngJLoop As Long
    Dim sngAllCount As Single, sngCount As Single
    Dim lng付款序号 As Long
    Dim strStartDate As String
    Dim strEndDate As String
    
    '标志，计划日期，计划金额
'    lblTemp(0).Caption = "付款计划清单"
    Call zlcommfun.ShowFlash("正在搜索付款记录,请稍候 ...", Me)
    
    mshMain.Redraw = False
    Screen.MousePointer = vbHourglass
    
    '根据操作类型设定记录读取条件
    'by lesfeng 2009-12-2 性能优化  修改绑定变量
    lng付款序号 = mlng付款序号
    If IsNull(lng付款序号) Then lng付款序号 = 0
    If mEditType = g新增 Then
        '新增时读取付款序号为空的应付款计划供选择
        strWhere = " And 付款序号 Is Null and ID in (Select ID From 应付记录 where (记录状态=1   or 记录状态=3) and 记录性质<>-1 and 计划日期 is not null) "
        strWhere = strWhere & "  and 计划日期  between [3] and [4]" '+1-1/24/60/60
    ElseIf mEditType = g修改 Then
        '编辑时读取付款序号为空或当前编辑的付款序号所对应的应付款计划
        strWhere = "  and ID in (Select ID From 应付记录 where (记录状态=1  or 记录状态=3) and 记录性质<>-1 and 计划日期 is not null) and (付款序号 Is Null Or 付款序号=[2])"
    Else
        '查看或审核时仅读取当前编辑的付款单所对应的应付款计划
        strWhere = " and 付款序号=[2]"
    End If
    '问题29231 by lesfeng 2010-04-23
    strStartDate = mstrStartDate & " 00:00:00"
    strEndDate = mstrEndDate & " 23:59:59"
    
    '读取应付款计划数据
    strSQL = "" & _
        "   Select  Decode(付款序号,Null,'','√') As 标志,ID,计划序号," & _
        "           TO_CHAR(计划日期,'yyyy-MM-dd') As 计划日期,to_char(计划金额,'99999999999.99') as 计划金额,发票号,入库单据号," & _
        "           品名,规格,计量单位,to_char(数量,'999999999999.9999') as 数量,摘要 " & _
        "   From 应付记录 " & _
        "   Where 记录性质=-1   And 单位ID=[1]" & strWhere & _
        "   Order By 发票号"
    On Error GoTo errHandle
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng单位ID, lng付款序号, CDate(strStartDate), CDate(strEndDate))
    
    With mshMain
        .Clear
        If rsTemp.EOF Then
            Set .Recordset = Nothing
            .Rows = 2
        Else
            Set .Recordset = rsTemp
        End If
    
        .FormatString = "^标志|||^计划日期|^计划金额|^发票号|^入库单号|^品名|^规格|^单位|^数量|^摘要"
        
        .ColAlignment(0) = 4
        .ColWidth(1) = 0: .ColWidth(2) = 0
        .ColWidth(3) = 1000: .ColAlignment(3) = 4
        .ColWidth(4) = 1200: .ColAlignment(4) = 7
        .ColWidth(5) = 1000: .ColAlignment(5) = 4
        .ColWidth(6) = 1000: .ColAlignment(6) = 4
        .ColWidth(7) = 1800: .ColAlignment(7) = 1
        .ColWidth(8) = 800: .ColAlignment(8) = 1
        .ColWidth(9) = 800: .ColAlignment(9) = 4
        .ColWidth(10) = 1200: .ColAlignment(10) = 7
        .ColWidth(11) = 2000: .ColAlignment(11) = 1
        
        mdbl累计应付 = 0
        mdbl本次应付 = 0
        sngCount = 0: sngAllCount = 0
        For lngLoop = 1 To .Rows - 1
            mdbl累计应付 = mdbl累计应付 + Val(.TextMatrix(lngLoop, 4))
            If Trim(.TextMatrix(lngLoop, 0)) <> "" Then
                mdbl本次应付 = mdbl本次应付 + Val(.TextMatrix(lngLoop, 4))
            End If
        Next
        .Row = 1: .Col = 1
        .Redraw = True
    End With
    Call SetMoneyLbl
    Call SetCmdEn
    
    Call zlcommfun.StopFlash
    Screen.MousePointer = vbDefault
    Get预付数据         '读取预付款
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then
        Resume
    Else
        zlcommfun.StopFlash
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub Get预付数据()
    '--------------------------------------------------------------
    '功能：读取并填充预付款记录供选择
    '参数：
    '返回：
    '说明：
    '--------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim strWhere As String
    Dim lngLoop As Long
    Dim lng付款序号 As Long
    
    '标志,结算方式,结算金额,结算号码
    'by lesfeng 2009-12-2 性能优化  修改绑定变量
    lng付款序号 = mlng付款序号
    If IsNull(lng付款序号) Then lng付款序号 = 0
    Call zlcommfun.ShowFlash("正在搜索预付款记录,请稍候 ...", Me)
    Screen.MousePointer = vbHourglass
    
    If mEditType = g新增 Then
        strWhere = " And 付款序号 Is Null"
    ElseIf mEditType = g修改 Then
        strWhere = " and (付款序号 Is Null Or 付款序号=[2])"
    Else
        strWhere = " And 付款序号=[2]"
    End If
    
    strSQL = "" & _
        "   Select Decode(付款序号,Null,'','√') As 标志,ID,结算方式,金额,结算号码 " & _
        "   From 付款记录 " & _
        "   Where 审核日期 Is not  Null And ( 记录状态=1 and 预付款=1)  And 单位ID=[1]" & strWhere & _
        "   Order By ID"
    On Error GoTo errHandle
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng单位ID, lng付款序号)
    
    mshList.Redraw = False
    mshList.Clear
    mshList.Tag = 0
    
    If rsTemp.EOF Then
        Set mshList.Recordset = Nothing
        mshList.Rows = 2
    Else
        Set mshList.Recordset = rsTemp
        mshList.Row = 1: mshList.Col = 1
    End If
    
    With mshList
        .FormatString = "^标志||^结算方式|^结算金额|^结算号码"
        .ColAlignment(0) = 4
        .ColWidth(1) = 0
        .ColWidth(2) = 1000: .ColAlignment(2) = 4
        .ColWidth(3) = 1200: .ColAlignment(3) = 7
        .ColWidth(4) = 1000: .ColAlignment(4) = 1
        mdbl累计预交 = 0
        mdbl本次预交 = 0
        For lngLoop = 1 To .Rows - 1
            mdbl累计预交 = mdbl累计预交 + Val(.TextMatrix(lngLoop, 3))
            If Trim(.TextMatrix(lngLoop, 0)) = "√" Then
                mdbl本次预交 = mdbl本次预交 + Val(.TextMatrix(lngLoop, 3))
            End If
            
            If Val(.TextMatrix(lngLoop, 3)) < 0 Then
                    Call SetMshRowColor(mshList, lngLoop, vbRed)
            Else
                    Call SetMshRowColor(mshList, lngLoop, &H0&)
            End If
            
        Next
    End With
    mshList.Redraw = True
    
    Call SetMoneyLbl
    SetCmdEn
    Call zlcommfun.StopFlash
    Screen.MousePointer = vbDefault
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then
        Resume
    Else
        zlcommfun.StopFlash
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub Full预付()
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:填充本次预付
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim lngRow As Long
    With msh预付
        .Clear
        Call Set预交列头
        .Rows = 2
        lngRow = 1
        For lngLoop = 1 To mshList.Rows - 1
            If Trim(mshList.TextMatrix(lngLoop, 0)) = "√" Then
                .TextMatrix(lngRow, 0) = mshList.TextMatrix(lngLoop, 1)
                .TextMatrix(lngRow, 1) = mshList.TextMatrix(lngLoop, 2)
                .TextMatrix(lngRow, 2) = mshList.TextMatrix(lngLoop, 3)
                .TextMatrix(lngRow, 3) = mshList.TextMatrix(lngLoop, 4)
                If Val(.TextMatrix(lngRow, 2)) < 0 Then
                    Call SetMshRowColor(msh预付, lngRow, vbRed)
                Else
                    Call SetMshRowColor(msh预付, lngRow, &H0&)
                End If
                lngRow = lngRow + 1
                .Rows = .Rows + 1
            End If
        Next
    End With
End Sub

Private Sub SetMshRowColor(ByVal mshGrid As MSHFlexGrid, ByVal lngRow As Long, ByVal oleColor As OLE_COLOR)
    '功能:设置指定行的颜色
    Dim lngOldRow As Long, lngoldCol As Long
    Dim i As Long
    With mshGrid
        lngOldRow = .Row: lngoldCol = .Col
        .Row = lngRow
        For i = 0 To .Cols - 1
            .Col = i
            .CellForeColor = oleColor
        Next
        .Row = lngOldRow: .Col = lngoldCol
    End With
End Sub

Private Sub cmdBack_Click()
    ChangeMode 1
    cmdDown.Enabled = True
    cmdBack.Enabled = False
    SetCmdEn
    mshMain.SetFocus
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDown_Click()
    Dim dblCount As Double
    Dim lngRow As Long, i As Long, j As Long
    If mEditType = g新增 Or mEditType = g修改 Then
        If mdbl本次预交 < 0 Then
            MsgBox "本次冲预付款总额不能小于零", vbInformation + vbDefaultButton1, gstrSysName
            Exit Sub
        End If
        '检查各结算方式的预付款总额的累计是否为负数
        Dim str结算方式 As String
        Dim dbl金额 As Double
        str结算方式 = ","
        With mshList
            For i = 1 To .Rows - 1
                dbl金额 = 0
                
                If InStr(1, str结算方式, "," & .TextMatrix(i, 2) & ",") = 0 And Trim(.TextMatrix(i, 0)) = "√" Then
                    For j = 1 To .Rows - 1
                        If .TextMatrix(i, 2) = .TextMatrix(j, 2) And Trim(.TextMatrix(j, 0)) = "√" Then
                            dbl金额 = dbl金额 + Val(.TextMatrix(j, 3))
                        End If
                    Next
                    If dbl金额 < 0 Then
                        MsgBox "结算方式为:" & .TextMatrix(i, 2) & "的总额不能为负数!", vbInformation + vbDefaultButton1, gstrSysName
                        Exit Sub
                    End If
                    str结算方式 = str结算方式 & .TextMatrix(i, 2) & ","
                End If
            Next
        End With
    
        With mshEdit
            If .Rows <= 2 And Trim(.TextMatrix(1, 0)) = "" Then
                .Rows = 2
                .PrimaryCol = 0
                If .ListIndex < 0 Then
                    .ListIndex = 0
                End If
                If Trim(.TextMatrix(1, 0)) = "" Then
                    .TextMatrix(1, 0) = .CboText
                End If
            End If
            If .Rows <= 2 Then
                If Val(.TextMatrix(1, 2)) = 0 Then
                    .TextMatrix(1, 1) = mdbl本次应付 - mdbl本次预交
                End If
            End If
            .Active = True
        End With
    End If
    '落列本次冲预付的数据
    Call Full预付
    
    Call 合计
    ChangeMode 2
    If mshEdit.Enabled And mshEdit.Visible Then mshEdit.SetFocus
    cmdDown.Enabled = False
    cmdBack.Enabled = True
    SetCmdEn
    If mshEdit.Enabled Then mshEdit.SetFocus
End Sub

Private Sub cmdHelp_Click()
       ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim strReg As String
    Dim blnSuccess As Boolean
    
    If mEditType = g查看 Then    '查看
        '打印
        printbill
        Unload Me
        Exit Sub
    End If
    
    If mEditType = g审核 Then        '审核
        If SaveCheck = True Then
            If IIf(Val(zlDatabase.GetPara("审核打印", glngSys, mlngModule)) = 1, 1, 0) = 1 Then
                '打印
                If InStr(mstrPrivs, ";付款通知单;") <> 0 Then
                    printbill
                End If
            End If
            mblnChange = False
            mblnSuccess = True
            Unload Me
        End If
        Exit Sub
    End If
    
    If ValidData = False Then Exit Sub
    
    If mEditType = g取消 Then
        If SaveStrike() = True Then
            mblnChange = False
            mblnSuccess = True
            Unload Me
        End If
        Exit Sub
    End If
    
    blnSuccess = SaveCard
    mblnChange = False
    If blnSuccess = True Then
        If IIf(Val(zlDatabase.GetPara("存盘打印", glngSys, mlngModule)) = 1, 1, 0) = 1 Then
            '打印
            If InStr(mstrPrivs, ";付款通知单;") <> 0 Then
                printbill
            End If
        End If
        mblnSuccess = True
        If mEditType = g修改 Then    '修改
            Unload Me
            Exit Sub
        End If
        
        GetPrivoder
    Else
        Exit Sub
    End If
    
    txtInfo(0).Text = ""
    Me.Tag = "-1"
    
    mshEdit.ClearBill
    msh预付.Clear
    msh预付.Rows = 2
    Set预交列头
    
    ChangeMode 1
    FillDeptDue
      
      
    mblnSave = False
    mblnEdit = True
    cmdBack.Enabled = False
    cmdDown.Enabled = True
End Sub

Private Sub cmd条件_Click()
        Dim blnOk As Boolean
        
        If frmTimeSel.GetTimeScope(mstrStartDate, mstrEndDate, Me) = False Then Exit Sub
        
        lblDATE.Caption = "日期范围:" & mstrStartDate & " 至 " & mstrEndDate
        
        '确定相关的供应商
        Call GetPrivoder
End Sub

Private Function GetPrivoder() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:加载供应商
    '--入参数:
    '--出参数:
    '--返  回:成功返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim itmTemp As ListItem
    Dim intR As Integer
    Dim strWhere  As String
    Dim lng付款序号 As Long
    Dim strStartDate As String
    Dim strEndDate As String
    
    'by lesfeng 2009-12-2 性能优化  修改绑定变量
    lng付款序号 = mlng付款序号
    If IsNull(lng付款序号) Then lng付款序号 = 0
    '根据操作类型设定记录读取条件
    If mEditType = g新增 Then
        '新增时读取付款序号为空的应付款计划供选择
        strWhere = " And 付款序号 Is Null  and ID in (Select ID From 应付记录 where (记录状态=1 or 记录状态=3) and 记录性质<>-1 and 计划日期 is not null) "
        strWhere = strWhere & " and 计划日期 between [3] and [4]" '+1-1/24/60/60
        
    ElseIf mEditType = g修改 Then
        '编辑时读取付款序号为空或当前编辑的付款序号所对应的应付款计划
        strWhere = " and (付款序号 Is Null Or 付款序号=[2]) And 单位id=[1]"
    Else
        '查看或审核时仅读取当前编辑的付款单所对应的应付款计划
        strWhere = " and 付款序号=[2] And 单位id=[1]"
    End If
    
    Dim str权限 As String
    '问题29231 by lesfeng 2010-04-23
    strStartDate = mstrStartDate & " 00:00:00"
    strEndDate = mstrEndDate & " 23:59:59"
    
    str权限 = " and " & Get分类权限(gstrPrivs)
    On Error GoTo errHandle
    strSQL = "Select ID,上级ID,编码,名称,简码,末级,许可证号,许可证效期,执照号,执照效期,税务登记号,地址,电话,开户银行," & _
                  "       帐号,联系人,建档时间,撤档时间,类型,信用期,信用额,销售委托人,销售委托日期,质量认证号,质量认证日期," & _
                  "       药监局备案号,药监局备案日期,授权号,授权期,站点" & _
                  "  From 供应商 where id in (Select distinct 单位id  from 应付记录 where 记录性质=-1 " & strWhere & ") " & str权限
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng单位ID, lng付款序号, CDate(strStartDate), CDate(strEndDate))
    
    Dim i As Long
    Dim strTmp As String
    Dim str类型 As String
    With rsTemp
        Me.lvwMain.ListItems.Clear
        Do While Not .EOF
            Set itmTemp = Me.lvwMain.ListItems.Add(, "K" & !ID, Nvl(!编码) & "--" & Nvl(!名称), 1, 1)
            i = 1
            itmTemp.SubItems(i) = Nvl(!许可证号)
            i = i + 1
            itmTemp.SubItems(i) = Format(!许可证效期, "yyyy-mm-dd")
            i = i + 1
            itmTemp.SubItems(i) = Nvl(!执照号)
            i = i + 1
            itmTemp.SubItems(i) = Format(!执照效期, "yyyy-mm-dd")
            i = i + 1
            itmTemp.SubItems(i) = Nvl(!税务登记号)
            i = i + 1
            itmTemp.SubItems(i) = Nvl(!地址)
            i = i + 1
            itmTemp.SubItems(i) = Nvl(!电话)
            i = i + 1
            itmTemp.SubItems(i) = Nvl(!开户银行)
            i = i + 1
            itmTemp.SubItems(i) = Nvl(!帐号)
            i = i + 1
            itmTemp.SubItems(i) = Nvl(!联系人)
            i = i + 1
            strTmp = Nvl(!类型)
            str类型 = ""
            For intR = 1 To Len(strTmp)
                If Mid(Nvl(!类型), intR, 1) = 1 Then
                    Select Case intR
                        Case 1
                            str类型 = str类型 & " " & "药品"
                        Case 2
                            str类型 = str类型 & " " & "物资"
                        Case 3
                            str类型 = str类型 & " " & "设备"
                        Case 4
                            str类型 = str类型 & " " & "其他"
                    End Select
                End If
            Next
            itmTemp.SubItems(i) = str类型
            i = i + 1
            itmTemp.SubItems(i) = IIf(Nvl(!信用期, 0) = 0, "", Nvl(!信用期) & "个月")
            i = i + 1
            itmTemp.SubItems(i) = Format(Nvl(!信用额, 0), "####0.00;-####0.00; ;")
            
            If lvwMain.SelectedItem Is Nothing Then
                itmTemp.Selected = True
            End If
            .MoveNext
        Loop
    End With
    
    '获取相关数据
    If Me.lvwMain.SelectedItem Is Nothing Then
        mlng单位ID = 0
    Else
        mlng单位ID = Val(Mid(lvwMain.SelectedItem.Key, 2))
    End If
    Call FillDeptDue
    GetPlanPayMoney
    Exit Function
    
errHandle:
    If ErrCenter = 1 Then Resume
End Function

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    SetEditPro
'    If mEditType = g新增 Or mEditType = g修改 Then
'        If txtDept.Enabled And txtDept.Visible Then txtDept.SetFocus
'    End If
  SetCmdEn
  mblnChange = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
        mblnFirst = True
        mintStep = 0
        mstrStartDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
        mstrEndDate = mstrStartDate
        lblDATE.Caption = "日期范围:" & mstrStartDate & " 至 " & mstrEndDate
        Call initCard
End Sub

Private Sub Form_Resize()
'    If Me.WindowState = 1 Then Exit Sub
'
'    cmdHelp.Move 90, Me.ScaleHeight - cmdHelp.Height - 90
'    cmdCancel.Move Me.ScaleWidth - cmdCancel.Width - 90, cmdHelp.Top
'    fraTemp.Move -150, cmdHelp.Top - 90, Me.ScaleWidth + 300
'    cmdOK.Move cmdCancel.Left - 1100, cmdHelp.Top
'    cmdBack.Move cmdOK.Left - 1100, cmdHelp.Top
'    cmdDown.Move cmdOK.Left, cmdOK.Top
'    Pic_Resize 0
End Sub

Private Sub ChangeMode(intMode As Integer)

    If intMode = mintStep Then Exit Sub
    
    mintStep = intMode
    
    If mintStep = 1 Then
        Pic(1).Enabled = False
        Pic(1).Visible = False
        Pic(0).Visible = True
        Pic(0).Enabled = True
    ElseIf mintStep = 2 Then
        Pic(0).Enabled = False
        Pic(0).Visible = False
        Pic(1).Visible = True
        Pic(1).Enabled = True
    ElseIf mintStep = 3 Then
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Dim blnYes As Boolean
    If mblnChange = False Then Exit Sub
    ShowMsgbox "你已经更改了单据信息,你这样退出的话," & vbCrLf & "所更改的数据将不能保存,真的要退出吗?", True, blnYes
    If blnYes = True Then Exit Sub
    Cancel = 1
End Sub

Private Sub lblInfo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    msngDownY = Y
    msngDownX = X
End Sub '

Private Sub lvwMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim blnYes As Boolean
    If mlng单位ID = Val(Mid(Item.Key, 2)) Then Exit Sub
    If mblnChange Then
        ShowMsgbox "你已经修改了当前数据,如果选择了其他单位," & vbCrLf & "则会清除你所设置的内容,真的要改变单位吗?", True, blnYes
        If blnYes = False Then
            Err = 0
            On Error GoTo ErrHand:
            lvwMain.ListItems("K" & mlng单位ID).Selected = True
            Exit Sub
        End If
        mblnChange = False
    Else
        mlng单位ID = Val(Mid(Item.Key, 2))
    End If
ErrHand:
    Call FillDeptDue
    '加数据
    Call GetPlanPayMoney
End Sub

Private Sub lvwMain_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlcommfun.PressKey vbKeyTab
    End If
End Sub

Private Sub lvwMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    PopupMenu mnuIco
End Sub

Private Sub mnuClear_Click()
    If Me.ActiveControl Is mshMain Then
        mshMain_DblClick
    Else
        mshList_DblClick
    End If
End Sub

Private Sub mnuClearAll_Click()
    Dim lngLoop As Long
    Dim objTemp As Object
    
    If Not (Me.ActiveControl Is mshMain) And Not (Me.ActiveControl Is mshList) Then Exit Sub
    
    Set objTemp = Me.ActiveControl
    For lngLoop = 1 To objTemp.Rows - 1
        objTemp.TextMatrix(lngLoop, 0) = ""
    Next
    If objTemp Is mshMain Then
        mdbl本次应付 = 0
    Else
        mdbl本次预交 = 0
    End If
    
    Call SetMoneyLbl
    Call SetCmdEn
End Sub

Private Sub mnuSelect_Click()
    If Me.ActiveControl Is mshMain Then
        mshMain_DblClick
    Else
        mshList_DblClick
    End If
End Sub

Private Sub mnuSelectAll_Click()
    Dim lngLoop As Long
    Dim objTemp As Object
    
    If Not (Me.ActiveControl Is mshMain) And Not (Me.ActiveControl Is mshList) Then Exit Sub
    
    Set objTemp = Me.ActiveControl
    For lngLoop = 1 To objTemp.Rows - 1
        objTemp.TextMatrix(lngLoop, 0) = "√"
    Next
    If objTemp Is mshList Then
        mdbl本次预交 = mdbl累计预交
    Else
        mdbl本次应付 = mdbl累计应付
    End If
    Call SetMoneyLbl
    Call SetCmdEn
End Sub

Private Sub mnuViewIcon_Click(Index As Integer)
    Dim intTemp As Integer
    For intTemp = 0 To 3
        mnuViewIcon(intTemp).Checked = False
    Next
    
    mnuViewIcon(Index).Checked = True
    lvwMain.View = Index
    lvwMain.Refresh
End Sub

Private Sub mshEdit_AfterDeleteRow()
    Dim Cur余额 As Currency
    Dim intLop As Integer
    
    Cur余额 = 0
    
    For intLop = 1 To mshEdit.Rows - 1
        If intLop <> mshEdit.Row Then
            Cur余额 = Cur余额 + Val(mshEdit.TextMatrix(intLop, 1))
        End If
    Next
    
    Cur余额 = (mdbl本次应付 - mdbl本次预交) - Cur余额
    
    If Cur余额 <> 0 Then
        mshEdit.TextMatrix(mshEdit.Row, 1) = Format(Cur余额, "#####0.00;-#####0.00; ;")
        mshEdit.TextMatrix(mshEdit.Row, 0) = mshEdit.CboText
    End If
    Call 合计
End Sub

Private Sub mshEdit_cboClick(ListIndex As Long)
    With mshEdit
        If .Col <> 0 Then Exit Sub
        .TextMatrix(.Row, .Col) = .CboText
    End With
End Sub

Private Sub mshEdit_cboKeyDown(KeyCode As Integer, Shift As Integer)
    With mshEdit
        .TextMatrix(.Row, .Col) = .CboText
    End With
End Sub

Private Sub mshEdit_EditChange(curText As String)
    mblnChange = True
    SetCmdEn
End Sub

Private Sub mshEdit_EnterCell(Row As Long, Col As Long)
    With mshEdit
        Select Case Col
            Case 1
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
            Case 2
                .TxtCheck = True
                .MaxLength = 10
        End Select
    End With
End Sub

Private Sub mshEdit_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim intLop As Integer, Cur余额 As Currency, curImprest As Currency
    If mEditType <> g新增 And mEditType <> g修改 Then
        If KeyCode = vbKeyReturn And mshEdit.Row = mshEdit.Rows - 1 Then
            zlcommfun.PressKey vbKeyTab
        End If
        Exit Sub
    End If
      
    With mshEdit
        If mEditType = g新增 Or mEditType = g修改 Then mblnChange = True
        If .Col = 2 Then
            If KeyCode <> vbKeyReturn Then
                .ColData(2) = 4
                .TxtCheck = False
            Else
                .ColData(2) = 0
                .TxtCheck = True
                .TextLen = 10
            End If
        End If
        If .Col = 1 And .Row = .Rows - 1 And KeyCode = vbKeyReturn Then
            If txtInfo(0).Enabled And txtInfo(0).Visible Then txtInfo(0).SetFocus
        End If
        
        If KeyCode <> vbKeyReturn Then Exit Sub
        If .TxtVisible = False Then Exit Sub
        
        If .Col = 1 Then
            Cur余额 = 0
            For intLop = 1 To .Rows - 1
                If intLop <> .Row Then
                    Cur余额 = Cur余额 + Val(.TextMatrix(intLop, 1))
                End If
            Next
            
            Cur余额 = (mdbl本次应付 - mdbl本次预交) - Cur余额
            
            If Val(.Text) = 0 And Cur余额 > 0 Then
                MsgBox "付款金额不能为空!", vbInformation, gstrSysName
                Cancel = True
                .TxtSetFocus
                Exit Sub
            End If
            If Not IsNumeric(.Text) And Trim(.Text) <> "" Then
                MsgBox "付款金额中含有非法字符!", vbInformation, gstrSysName
                Cancel = True
                .TxtSetFocus
                Exit Sub
            End If
            If Val(.Text) < 0 Then
                MsgBox "付款分录金额不能为负数!", vbInformation, gstrSysName
                Cancel = True
                .TxtSetFocus
                Exit Sub
            End If
            If Val(.Text) >= 10 ^ 14 - 1 Then
                MsgBox "付款金额必须小于" & (10 ^ 14 - 1), vbInformation + vbOKOnly, gstrSysName
                Cancel = True
                .TxtSetFocus
                Exit Sub
            End If
            If Trim(.Text) = "" Then Exit Sub
            
            Cur余额 = Cur余额 - IIf(Trim(.Text) = "", 0, .Text)
            If Cur余额 < 0 Then
                MsgBox "付款金额超出总额!", vbInformation, gstrSysName
                Cancel = True
                .TxtSetFocus
                Exit Sub
            End If
            If .Row >= .Rows - 1 And Cur余额 > 0 Then
                .Rows = .Rows + 1
            End If
                    
            .Text = GetFormat(.Text, 2)
            .TextMatrix(.Row, .Col) = .Text
            If Cur余额 > 0 Then
                .TextMatrix(.Row + 1, 1) = GetFormat(Cur余额, 2)
                .TextMatrix(.Row + 1, 0) = .CboText
            End If
            Call 合计
        End If
    End With
End Sub

Private Sub mshList_DblClick()
    If mEditType <> g新增 And mEditType <> g修改 Then Exit Sub
    
    With mshList
        If .Recordset Is Nothing Then Exit Sub
        
        .TextMatrix(.Row, 0) = IIf(Trim(.TextMatrix(.Row, 0)) = "", "√", "")
        If Trim(.TextMatrix(.Row, 0)) = "" Then
            mdbl本次预交 = mdbl本次预交 - Val(.TextMatrix(.Row, 3))
        Else
            mdbl本次预交 = mdbl本次预交 + Val(.TextMatrix(.Row, 3))
        End If
    End With
    Call SetMoneyLbl
    
    Call SetCmdEn
End Sub

Private Sub SetMoneyLbl()
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:设置标签金额
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    lbl(2).Caption = "冲预付款合计：" & Format(mdbl本次预交, "###0.00;-###0.00;0;0") & "元"
    lbl(1).Caption = "本次付款：" & Format(mdbl本次应付, "###0.00;-###0.00;0;0") & "元"
    lbl金额(1).Caption = "累计应付:" & Format(mdbl累计应付, "###0.00;-###0.00;0.00;0.00") & ""
    lbl金额(2).Caption = "付款金额:" & Format(mdbl本次应付, "###0.00;-###0.00;0.00;0.00") & ""
    lbl金额(3).Caption = "预交累计:" & Format(mdbl累计预交, "###0.00;-###0.00;0.00;0.00") & ""
    lbl金额(4).Caption = "冲预付:" & Format(mdbl本次预交, "###0.00;-###0.00;0.00;0.00") & ""
    lbl金额(5).Caption = "本次应付:" & Format(mdbl本次应付 - mdbl本次预交, "###0.00;-###0.00;0.00;0.00") & ""
End Sub

Private Sub 合计()
    Dim lngRow As Long
    Dim dblCount As Double
   '获取合计数
    With mshEdit
        For lngRow = 1 To .Rows - 1
            dblCount = dblCount + Val(.TextMatrix(lngRow, 1))
        Next
    End With
    lbl(3).Caption = "结算合计:" & Format(dblCount, "###0.00;-###0.00;0;0") & "元"
End Sub

Private Sub mshList_GotFocus()
    '
    Err = 0
    On Error Resume Next
    mshList.Col = 0
    mshList.ColSel = mshList.Cols - 1
    Err = 0
End Sub

Private Sub mshList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        mshList_DblClick
    ElseIf KeyCode = vbKeyReturn Then
        zlcommfun.PressKey vbKeyTab
    End If
End Sub

Private Sub mshList_LostFocus()
    Err = 0
    On Error Resume Next
    mshList.Col = 0
    mshList.ColSel = 0
    Err = 0
End Sub

Private Sub mshList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button <> 2 Or mshList.Recordset Is Nothing Then Exit Sub
    If mEditType <> g新增 And mEditType <> g修改 Then Exit Sub

    SetEnabled 1
    Me.PopupMenu mnuHandle
    If mshList.Enabled Then mshList.SetFocus
End Sub

Private Sub mshMain_DblClick()
    Dim intCol As Integer
    With mshMain
        If .Recordset Is Nothing Then Exit Sub
        If mEditType <> g新增 And mEditType <> g修改 Then Exit Sub
        
        .TextMatrix(.Row, 0) = IIf(.TextMatrix(.Row, 0) = "", "√", "")
        intCol = IIf(mbln付款单, .Cols - 1, 4)
        
        If Trim(.TextMatrix(.Row, 0)) = "" Then
            mdbl本次应付 = mdbl本次应付 - Val(.TextMatrix(.Row, intCol))
        Else
            mdbl本次应付 = mdbl本次应付 + Val(.TextMatrix(.Row, intCol))
        End If
        mblnChange = True
    End With
    Call SetMoneyLbl
    Call SetCmdEn
End Sub

Private Sub SetCmdEn()
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:设置控件属性
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
   ' cmdDown.Enabled = mdbl本次应付 <> 0 And mdbl本次应付 - mdbl本次预交 > 0
    If mEditType = g审核 Or mEditType = g取消 Then
        cmdOK.Enabled = Me.cmdBack.Enabled
    ElseIf mEditType = g查看 Then
        cmdOK.Enabled = True
    Else
        cmdOK.Enabled = Me.cmdBack.Enabled And mblnChange
    End If
End Sub

Private Sub mshMain_GotFocus()
    Err = 0
    On Error Resume Next
    mshMain.Col = 0
    mshMain.ColSel = mshMain.Cols - 1
    Err = 0
    
End Sub

Private Sub mshMain_LostFocus()
    Err = 0
    On Error Resume Next
    mshMain.Col = 0
    mshMain.ColSel = 0
    Err = 0
End Sub

Private Sub mshMain_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then        '
        mshMain_DblClick
    ElseIf KeyCode = vbKeyReturn Then
        zlcommfun.PressKey vbKeyTab
    End If
End Sub

Private Sub mshMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    If mEditType <> g新增 And mEditType <> g修改 Then Exit Sub
    If mshMain.Recordset Is Nothing Then Exit Sub
    If mshMain.Enabled Then mshMain.SetFocus

    SetEnabled 0
    Me.PopupMenu mnuHandle
End Sub

Private Sub msh预付_KeyDown(KeyCode As Integer, Shift As Integer)
       If KeyCode = vbKeyReturn Then
            zlcommfun.PressKey vbKeyTab
       End If
End Sub

Private Sub SetEnabled(iControl As Integer)
    If iControl = 1 Then
        If mshList.TextMatrix(mshList.Row, 0) = "" Then
            mnuSelect.Enabled = True
            mnuClear.Enabled = False
        Else
            mnuSelect.Enabled = False
            mnuClear.Enabled = True
        End If
    Else
        If mshMain.TextMatrix(mshMain.Row, 0) = "" Then
            mnuSelect.Enabled = True
            mnuClear.Enabled = False
        Else
            mnuSelect.Enabled = False
            mnuClear.Enabled = True
        End If
    End If
End Sub

Private Sub FillDeptDue()
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:加载部门数据
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHandle
    strSQL = "Select ID,上级ID,编码,名称,简码,末级,许可证号,许可证效期,执照号,执照效期,税务登记号,地址,电话,开户银行," & _
                  "       帐号,联系人,建档时间,撤档时间,类型,信用期,信用额,销售委托人,销售委托日期,质量认证号,质量认证日期," & _
                  "       药监局备案号,药监局备案日期,授权号,授权期,站点" & _
                  "  From 供应商 where id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng单位ID)
    
    If Not rsTemp.EOF Then
        lblInfo(9).Caption = "单位名称:" & rsTemp!名称
        lblInfo(1).Caption = "地址电话:" & IIf(IsNull(rsTemp!地址), "", rsTemp!地址) & IIf(IsNull(rsTemp!地址), "", "  TEL:") & IIf(IsNull(rsTemp!电话), "", rsTemp!电话)
        lblInfo(2).Caption = "开户银行:" & IIf(IsNull(rsTemp!开户银行), "", rsTemp!开户银行)
        lblInfo(3).Caption = "税务登记号:" & IIf(IsNull(rsTemp!税务登记号), "", rsTemp!税务登记号)
    End If
    If mshMain.Enabled And mshMain.Visible Then mshMain.SetFocus
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub txtDept_LostFocus()
    ImeLanguage False
End Sub

Private Sub txtInfo_Change(Index As Integer)
    mblnChange = True
    SetCmdEn
End Sub

Private Sub txtInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Index = 4 Then
            cmdOK.SetFocus
        Else
            txtInfo(Index + 1).SetFocus
        End If
    End If
End Sub

Private Function ValidData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:验证数据的合法性
    '--入参数:
    '--出参数:
    '--返  回:验证合法,返回True,否则=false
    '-----------------------------------------------------------------------------------------------------------
    Dim intIndex As Integer
    Dim lngRow As Long
    Dim strTemp As String
    Dim dblCount As Double
    If mlng单位ID = 0 Then
         ShowMsgbox "供应商选择有误,请重新选择!"
         Call cmdBack_Click
         Exit Function
    End If
        
    With mshEdit
        For lngRow = 1 To .Rows - 1
            If Trim(.TextMatrix(lngRow, 0)) <> "" And Trim(.TextMatrix(lngRow, 1)) <> "" Then
                strTemp = Trim(.TextMatrix(lngRow, 1))
                If strTemp = "" Then
                    ShowMsgbox "结算金额必需输入!"
                    .Row = lngRow
                    .Col = 1
                    If mshEdit.Enabled Then mshEdit.SetFocus
                    Exit Function
                End If
                
                If Not IsNumeric(strTemp) Then
                    ShowMsgbox "结算金额不是数据型,请重输!"
                    .Row = lngRow
                    .Col = 1
                    If mshEdit.Enabled Then mshEdit.SetFocus
                    Exit Function
                End If
                If Val(strTemp) < 0 Then
                    ShowMsgbox "结算金额不能小于零,请重输!"
                    .Row = lngRow
                    .Col = 1
                    If mshEdit.Enabled Then mshEdit.SetFocus
                    Exit Function
                End If
                If Val(strTemp) > 999999999.99 Then
                    ShowMsgbox "结算金额不能大于999999999.99,请重输!"
                    .Row = lngRow
                    .Col = 1
                    If mshEdit.Enabled Then mshEdit.SetFocus
                    Exit Function
                End If
                dblCount = dblCount + Val(strTemp)
                strTemp = Trim(.TextMatrix(lngRow, 2))
                If strTemp <> "" Then
                    If LenB(StrConv(strTemp, vbFromUnicode)) > 10 Then
                        ShowMsgbox "结算号码超长,最多能输入5个汉字或10个字符!"
                        .Row = lngRow
                        .Col = 2
                        If mshEdit.Enabled Then mshEdit.SetFocus
                        Exit Function
                    End If
                    If InStr(1, strTemp, "'") <> 0 Then
                        ShowMsgbox "结算号码不能输入单引号!"
                        .Row = lngRow
                        .Col = 2
                        If mshEdit.Enabled Then mshEdit.SetFocus
                        Exit Function
                    End If
                End If
            End If
        Next
    End With
    If CCur(mdbl本次应付 - (dblCount + mdbl本次预交)) <> 0 Then
        ShowMsgbox "付款金额不平,请检查付款金额与入库单" & vbCrLf & "发票金额和预付款之差是否相同!"
        If mshEdit.Enabled Then mshEdit.SetFocus
        Exit Function
    End If
    If mdbl本次应付 = 0 Then
        ShowMsgbox "本次不存在任何应付记录,请检查!"
        Exit Function
    End If
    If LenB(StrConv(txtInfo(0).Text, vbFromUnicode)) > 50 Then
        ShowMsgbox "付款说明的长度超长!(最多为50个字符或25个汉字)"
        txtInfo(0).SetFocus
        Exit Function
    End If
    
    ValidData = True
End Function

Private Function SaveCard() As Boolean
    Dim strNO_IN As String
    Dim int序号_IN As Integer
    Dim dbl金额_IN As Double
    Dim str结算方式_IN As String
    Dim str结算号码_IN As String
    Dim intCol   As Integer
    Dim str填制人_IN As String
    Dim str填制日期_IN As String
    Dim lng付款序号_IN As Long
    Dim str摘要_IN As String
    Dim lngRow As Long
    
    SaveCard = False
    
    'txtNo = NextNo(31)
    strNO_IN = txtNo
    str填制人_IN = UserInfo.姓名
    str填制日期_IN = Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
    
    str摘要_IN = txtInfo(0).Text

    
    On Error GoTo errHandle:
    
    '开始事务
    gcnOracle.BeginTrans
    
    If mEditType = g新增 Then
        strNO_IN = NextNo(31)
        lng付款序号_IN = zlDatabase.GetNextId("付款记录")
    Else
        lng付款序号_IN = mlng付款序号
        gstrSQL = "zl_付款记录_DELETE('" & strNO_IN & "')"
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    End If
       
     Dim blnData As Boolean
     blnData = False
    '循环保存每行数据
    With mshEdit
        'zl_付款管理_INSERT( /*strNO_IN*/, /*int序号_IN*/, /*int预付款_IN*/, /*lng单位ID_IN*/,
            '/*dbl金额_IN*/, /*str结算方式_IN*/, /*str结算号码_IN*/, /*str填制人_IN*/, /*str填制日期_IN*/,
            '/*lng付款序号_IN*/, /*str摘要_IN*/ );
        For lngRow = 1 To .Rows - 1
            If Val(.TextMatrix(lngRow, 1)) <> 0 And Trim(.TextMatrix(lngRow, 0)) <> "" Then
                blnData = True
                dbl金额_IN = .TextMatrix(lngRow, 1)
                str结算方式_IN = .TextMatrix(lngRow, 0)
                str结算号码_IN = .TextMatrix(lngRow, 2)
                
                gstrSQL = "" & _
                    "   zl_付款管理_INSERT('" & _
                    strNO_IN & "'," & _
                    lngRow & "," & _
                    0 & "," & _
                    mlng单位ID & "," & _
                    dbl金额_IN & ",'" & _
                    str结算方式_IN & "','" & _
                    str结算号码_IN & "','" & _
                    str填制人_IN & "',to_date('" & _
                    str填制日期_IN & "','yyyy-mm-dd HH24:MI:SS')," & _
                    lng付款序号_IN & ",'" & _
                    str摘要_IN & "')"
                    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
            End If
        Next
    End With
    If blnData = False Then
            gstrSQL = "" & _
                "   zl_付款管理_INSERT('" & _
                strNO_IN & "'," & _
                lngRow & "," & _
                0 & "," & _
                mlng单位ID & "," & _
                dbl金额_IN & ",'" & _
                "" & "','" & _
                "" & "','" & _
                str填制人_IN & "',to_date('" & _
                str填制日期_IN & "','yyyy-mm-dd HH24:MI:SS')," & _
                lng付款序号_IN & ",'" & _
                str摘要_IN & "')"
         zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
        
    End If
    Dim strIdin As String
    Dim str计划IN As String
    strIdin = ""
    str计划IN = ""
    
    '对应采购清单
    With mshMain
        For lngRow = 1 To .Rows - 1
            If .TextMatrix(lngRow, 0) <> "" Then
            
                '    Id_In       In Varchar2 := Null,
                '    计划序号_In In Varchar2 := Null, --以0,1,2,3方式传入
                '    付款序号_In In 付款记录.付款序号%Type := Null,
                '    预付款_In   In 付款记录.预付款%Type := 0,
                '    金额_In     In 应付记录.发票金额%Type := 0
  
                 intCol = IIf(mbln付款单, .Cols - 1, 4)
                gstrSQL = "zl_付款序号_UPDATE(" & _
                    "'" & Val(.TextMatrix(lngRow, 1)) & "'," & _
                    "'" & Val(.TextMatrix(lngRow, 2)) & "'," & _
                    lng付款序号_IN & "," & _
                    "0," & _
                    "" & Val(.TextMatrix(lngRow, intCol)) & ")"
                zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
            End If
        Next
    End With

    strIdin = ""
    '保存预付款
    With mshList
        For lngRow = 1 To .Rows - 1
            If .TextMatrix(lngRow, 0) <> "" Then
                'strIdin = strIdin & "," & Val(.TextMatrix(lngRow, 1))
                gstrSQL = "zl_付款序号_UPDATE(" & _
                    "'" & Val(.TextMatrix(lngRow, 1)) & "'," & _
                    "NULL" & "," & _
                    lng付款序号_IN & "," & _
                    "1," & _
                    Val(.TextMatrix(lngRow, 3)) & "" & _
                    ")"
                zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
            End If
        Next
    End With
    
    '提交事务
    gcnOracle.CommitTrans
    Me.stbThis.Panels(2).Text = "上张单据号为:" & strNO_IN
    SaveCard = True
    Exit Function
errHandle:
    
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Function

Private Sub SetEditPro()
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:设置编辑属性
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim intIndex As Integer
    For intIndex = 0 To 4
        txtInfo(intIndex).Enabled = mblnEdit
    Next
    mshEdit.Active = mblnEdit
    cmdOK.Enabled = (Not mblnEdit) And mEditType <> g查看
End Sub

Private Function SaveCheck() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:审核单据
    '--入参数:
    '--出参数:
    '--返  回:成功,返回True,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim strNO_IN As String
    SaveCheck = False
    
    strNO_IN = txtNo
    On Error GoTo errHandle:
    '   zl_付款管理_VERIFY(NO_IN);
    gstrSQL = "zl_付款管理_VERIFY('" & _
        strNO_IN & "')"
    
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    SaveCheck = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SaveStrike() As Boolean
 '-----------------------------------------------------------------------------------------------------------
    '--功  能:冲销单据
    '--入参数:
    '--出参数:
    '--返  回:成功,返回True,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim strNO_IN As String
    
    SaveStrike = False
    
    strNO_IN = txtNo
    On Error GoTo errHandle:
    '   zl_付款管理_VERIFY(NO_IN);
    gstrSQL = "zl_付款管理_strike('" & _
        strNO_IN & "')"
    
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    SaveStrike = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

'打印单据
Private Sub printbill()
    ReportOpen gcnOracle, glngSys, "ZL1_BILL_1323_1", Me, "单据编号=" & txtNo, "记录状态=" & mint记录状态
End Sub

