VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmMediPrice 
   Caption         =   "药品调价单"
   ClientHeight    =   9195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14700
   Icon            =   "frmMediPrice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9195
   ScaleWidth      =   14700
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picItem 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   5040
      ScaleHeight     =   2415
      ScaleWidth      =   5175
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CommandButton cmdExit 
         Caption         =   "退出(&E)"
         Height          =   350
         Left            =   3720
         Picture         =   "frmMediPrice.frx":058A
         TabIndex        =   30
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "添加(&A)"
         Height          =   350
         Left            =   2520
         Picture         =   "frmMediPrice.frx":06D4
         TabIndex        =   29
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton CmdSelecter 
         Caption         =   "…"
         Height          =   300
         Left            =   2450
         TabIndex        =   28
         Top             =   55
         Width           =   255
      End
      Begin VB.CheckBox ChkSelect 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "全选"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3000
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   120
         Width           =   675
      End
      Begin VB.TextBox txtItem 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   960
         TabIndex        =   25
         Top             =   60
         Width           =   1485
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfSpec 
         Height          =   1200
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   4800
         _cx             =   8467
         _cy             =   2117
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
         BackColorSel    =   16761024
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   0
         GridColorFixed  =   0
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   16
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmMediPrice.frx":081E
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
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "草药品种"
         Height          =   180
         Left            =   120
         TabIndex        =   24
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdItem 
      Caption         =   "选择品种(&I)"
      Height          =   350
      Left            =   11430
      Picture         =   "frmMediPrice.frx":0A3D
      TabIndex        =   22
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdPstor 
      Caption         =   "打印库存变动表(&S)…"
      Height          =   350
      Left            =   8400
      Picture         =   "frmMediPrice.frx":0B87
      TabIndex        =   5
      Top             =   4200
      Width           =   1965
   End
   Begin TabDlg.SSTab sstabDetail 
      Height          =   4095
      Left            =   0
      TabIndex        =   9
      Top             =   4320
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   7223
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "库存变动表(&S)"
      TabPicture(0)   =   "frmMediPrice.frx":0CD1
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblTitle"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "BillStore"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "应付款变动表(&P)"
      TabPicture(1)   =   "frmMediPrice.frx":0CED
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "BillPay"
      Tab(1).ControlCount=   1
      Begin ZL9BillEdit.BillEdit BillStore 
         Height          =   3615
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   6376
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
         BackColorBkg    =   14737632
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483643
      End
      Begin ZL9BillEdit.BillEdit BillPay 
         Height          =   3555
         Left            =   -74880
         TabIndex        =   12
         Top             =   360
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   6271
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
         BackColorBkg    =   14737632
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
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "库存变动表："
         Height          =   180
         Left            =   3240
         TabIndex        =   10
         Top             =   120
         Width           =   1080
      End
   End
   Begin MSComctlLib.ListView lvwItem 
      Height          =   2505
      Left            =   2400
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   4419
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin ZL9BillEdit.BillEdit BillPrice 
      Height          =   2595
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   4577
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
   Begin VB.CommandButton cmdPrint 
      Caption         =   "打印(&P)…"
      Height          =   350
      Left            =   11430
      Picture         =   "frmMediPrice.frx":0D09
      TabIndex        =   3
      Top             =   1161
      Width           =   1215
   End
   Begin VB.CommandButton cmdCanc 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   11430
      Picture         =   "frmMediPrice.frx":0E53
      TabIndex        =   2
      Top             =   663
      Width           =   1215
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Left            =   -435
      TabIndex        =   7
      Top             =   4060
      Width           =   16815
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   11430
      Picture         =   "frmMediPrice.frx":0F9D
      TabIndex        =   4
      Top             =   1659
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   11430
      Picture         =   "frmMediPrice.frx":10E7
      TabIndex        =   1
      Top             =   165
      Width           =   1215
   End
   Begin VB.Frame fraCondition 
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   0
      TabIndex        =   13
      Top             =   3000
      Width           =   14535
      Begin VB.ComboBox cbo售价计算方式 
         Height          =   300
         Left            =   11880
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   900
         Width           =   2415
      End
      Begin VB.OptionButton opt时间 
         Caption         =   "指定日期执行"
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   36
         Top             =   503
         Width           =   1695
      End
      Begin VB.OptionButton opt时间 
         Caption         =   "立即执行"
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   35
         Top             =   503
         Width           =   1215
      End
      Begin VB.CheckBox chk草药批量调价 
         Caption         =   "同品种药品价格一致(按剂量换算时)"
         Height          =   210
         Left            =   10440
         TabIndex        =   31
         Top             =   525
         Width           =   3210
      End
      Begin VB.CheckBox chk自动调成本价 
         Caption         =   "调售价时自动按加成率调整成本价"
         Height          =   210
         Left            =   4680
         TabIndex        =   21
         Top             =   960
         Width           =   3015
      End
      Begin VB.CheckBox chk自动计算应付款变动 
         Caption         =   "自动计算应付款变动"
         Height          =   210
         Left            =   2520
         TabIndex        =   14
         Top             =   960
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chk按批次 
         Caption         =   "成本价按库房批次调整"
         Height          =   210
         Left            =   90
         TabIndex        =   15
         Top             =   960
         Width           =   2175
      End
      Begin VB.CheckBox Chk定价 
         Caption         =   "时价药品改为定价"
         Enabled         =   0   'False
         Height          =   210
         Left            =   8520
         TabIndex        =   16
         Top             =   525
         Width           =   1770
      End
      Begin VB.TextBox txtSummary 
         Height          =   300
         Left            =   960
         TabIndex        =   18
         Top             =   60
         Width           =   6765
      End
      Begin VB.TextBox txtValuer 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   300
         Left            =   8805
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   60
         Width           =   2445
      End
      Begin MSComCtl2.DTPicker dtpRunDate 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
         Height          =   300
         Left            =   5880
         TabIndex        =   33
         Top             =   480
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
         Format          =   184745987
         CurrentDate     =   36846.5833333333
      End
      Begin VB.Label lbl调价方式 
         AutoSize        =   -1  'True
         Caption         =   "售价计算方式"
         Height          =   180
         Left            =   10680
         TabIndex        =   38
         Top             =   960
         Width           =   1080
      End
      Begin VB.Label lbl执行时间 
         Caption         =   "执行时间"
         Height          =   180
         Left            =   120
         TabIndex        =   34
         Top             =   540
         Width           =   855
      End
      Begin VB.Label lblInfo 
         Caption         =   "无调价权限不能调价！"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   8040
         TabIndex        =   32
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lblSummary 
         AutoSize        =   -1  'True
         Caption         =   "调价说明"
         Height          =   180
         Left            =   90
         TabIndex        =   20
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lblValuer 
         AutoSize        =   -1  'True
         Caption         =   "调价人"
         Height          =   180
         Left            =   8175
         TabIndex        =   19
         Top             =   120
         Width           =   540
      End
   End
   Begin VB.Label lblHelp 
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   2640
      Width           =   12600
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMediPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lngBillId As Long                '功能类型:0-调价处理;其他-显示lngBillId确定的历史调价单
Public lngMediId As Long                '进入类型:0-未指定调价药品;其他-进入时直接显示lngMediId的原价格情况
Public lngItemID As Long                '进入类型:>0时按品种提取所有规格

Private blnModify As Boolean
Private blnFirst As Boolean
Private intDrugType As Integer          '1-成药（西成药、中成药）;2-中草药
Private mstrPrivs As String
Private mstrAdjMsg As String            '存在未执行调价记录的药品的提示信息
Private mblnAllUnAdj As Boolean         '品种对应的规格都存在未执行价格
Private Const mlngColUpdate As Long = &H8000000F '不能被修改的背景颜色
Private mstr所有记录 As String          '记录界面中所有的数据，看数据是否进行了修改
Private mrs分段加成 As ADODB.Recordset    '记录设置了哪些加成率段
Private mdbl分段加成率 As Double
Private mdbl成本价 As Double            '记录修改之前的成本价

'--------调价单列（售价调价）--------------
Private Enum 售价列表
    药品id = 0
    品名 = 1
    规格 = 2
    产地 = 3
    单位 = 4
    类型 = 5
    上次日期 = 6
    原成本价 = 7
    现成本价 = 8
    原价 = 9
    现价 = 10
    现收入ID = 11
    原收入ID = 12
    收入名称 = 13
    原采购限价 = 14
    现采购限价 = 15
    原指导售价 = 16
    现指导售价 = 17
    是否有库存 = 18
    剂量系数 = 19
    药名ID = 20
    包装系数 = 21
    差价让利比 = 22
    加成率 = 23
    列数 = 24
End Enum

'--------库存变动列（时价药品按批次调售价、成本价调价）--------------
Private Enum 库存列表
    库房 = 0
    供应商 = 1
    药品 = 2
    规格 = 3
    单位 = 4
    批号 = 5
    效期 = 6
    产地 = 7
    数量 = 8
    原价 = 9
    现价 = 10
    调整金额 = 11
    加成率 = 12
    原成本价 = 13
    现成本价 = 14
    差价差 = 15
    批次 = 16
    变价 = 17
    药品id = 18
    库房id = 19
    供应商ID = 20

    列数 = 21
End Enum

'--------应付款列（成本价调价需要产生应付记录时）--------------
Private Enum 应付款列
    药品id = 0
    品名 = 1
    发票号 = 2
    发票日期 = 3
    发票金额 = 4
    
    列数 = 5
End Enum

Dim rsTemp As New ADODB.Recordset
Dim intCount As Integer
Dim objItem As ListItem
Dim objNode As Node
Dim dtToday As Date
Dim int药库单位 As Integer      '是否以药库单位显示
Dim mstrNo As String            '调价单No

Private mbln时价药品调价 As Boolean         '时价药品调价是否按批次执行
Private mbln限价提示 As Boolean             '新售价超过限价时是否提示
Private mstr药品 As String
Private mlng批次 As Long
Private mlng药品ID As Long
Private mintCurRow As Integer
Private mintCurCol As Integer

'从参数表中取药品价格小数位数
Private mintCostDigit As Integer        '成本价小数位数
Private mintPriceDigit As Integer       '售价小数位数
Private mintNumberDigit As Integer      '数量小数位数
Private mintMoneyDigit As Integer       '金额小数位数
Private mstrMoneyFormat As String

Private mintSalePriceDigit As Integer

'调价导航中传入的参数
Private mint调价 As Integer             '0-调售价;1-调成本价;2-调售价及成本价;3-仅调整收入项目
Private mlng供应商ID As Long
Private mdbl加成率 As Double
Private mbln应付记录 As Boolean         'False-不产生应付记录;True-产生应付记录
Private Sub BatchAdjustPriceByItem(ByVal lngRow As Long, ByVal dblPrice As Double)
    '按品种调价：相同品种的规格的售价保持一致（按剂量系数换算），暂时仅支持中草药
    Dim lng药名id As Long
    Dim dbl剂量系数 As Double
    Dim dbl包装系数 As Double
    Dim n As Long
    Dim dbl单价 As Double
    Dim dbl现价 As Double

    If chk草药批量调价.Visible = False Then Exit Sub
    If chk草药批量调价.Value <> 1 Then Exit Sub
    
    With BillPrice
        lng药名id = Val(.TextMatrix(lngRow, 售价列表.药名ID))
        dbl剂量系数 = Val(.TextMatrix(lngRow, 售价列表.剂量系数))
        dbl包装系数 = Val(.TextMatrix(lngRow, 售价列表.包装系数))
        dbl单价 = dblPrice / dbl包装系数 / dbl剂量系数
        
        For n = 1 To .Rows - 1
            If Val(.TextMatrix(n, 售价列表.药品id)) > 0 Then
                If Val(.TextMatrix(n, 售价列表.药名ID)) = lng药名id And n <> lngRow Then
                    dbl现价 = dbl单价 * Val(.TextMatrix(n, 售价列表.包装系数)) * Val(.TextMatrix(n, 售价列表.剂量系数))
                    
                    '现价大于指导售价时，提示是否继续
                    If mbln限价提示 = True Then
                        If .TextMatrix(n, 售价列表.类型) = "定价" And dbl现价 > Val(BillPrice.TextMatrix(n, 售价列表.现指导售价)) Then
                           MsgBox .TextMatrix(n, 售价列表.品名) & "现价高于指导零售价" & Val(BillPrice.TextMatrix(n, 售价列表.现指导售价)) & "，采购限价将和采购价一致！", vbInformation, gstrSysName
                        End If
                    End If
            
                    .TextMatrix(n, 售价列表.现价) = dbl现价
                    If dbl现价 > Val(BillPrice.TextMatrix(n, 售价列表.现指导售价)) Then
                        .TextMatrix(.Row, 售价列表.现指导售价) = FormatEx(dbl现价, mintPriceDigit)
                    End If
                    
                    Call ChangeDrugStore(n, Val(.TextMatrix(n, 售价列表.药品id)), dbl现价)
                End If
            End If
        Next
    End With
End Sub

Private Sub BatchAdjustCostByItem(ByVal lngRow As Long, ByVal dblCost As Double)
    '按品种调价：相同品种的规格的成本价保持一致（按剂量系数换算），暂时仅支持中草药
    Dim lng药名id As Long
    Dim dbl剂量系数 As Double
    Dim dbl包装系数 As Double
    Dim n As Long
    Dim dbl单价 As Double
    Dim dbl现价 As Double

    If chk草药批量调价.Visible = False Then Exit Sub
    If chk草药批量调价.Value <> 1 Then Exit Sub
    
    With BillPrice
        lng药名id = Val(.TextMatrix(lngRow, 售价列表.药名ID))
        dbl剂量系数 = Val(.TextMatrix(lngRow, 售价列表.剂量系数))
        dbl包装系数 = Val(.TextMatrix(lngRow, 售价列表.包装系数))
        dbl单价 = dblCost / dbl包装系数 / dbl剂量系数
        
        For n = 1 To .Rows - 1
            If Val(.TextMatrix(n, 售价列表.药品id)) > 0 Then
                If Val(.TextMatrix(n, 售价列表.药名ID)) = lng药名id And n <> lngRow Then
                    dbl现价 = dbl单价 * Val(.TextMatrix(n, 售价列表.包装系数)) * Val(.TextMatrix(n, 售价列表.剂量系数))
                    
                    '现价大于指导售价时，提示是否继续
                    If mbln限价提示 = True Then
                        If dbl现价 > Val(BillPrice.TextMatrix(n, 售价列表.现采购限价)) Then
                            MsgBox .TextMatrix(n, 售价列表.品名) & "现成本价高于指导采购限价" & Val(BillPrice.TextMatrix(n, 售价列表.现采购限价)) & "，采购限价将和采购价一致！", vbInformation, gstrSysName
                        End If
                    End If
            
                    .TextMatrix(n, 售价列表.现成本价) = dbl现价
                    
                    If dbl现价 > Val(BillPrice.TextMatrix(n, 售价列表.现采购限价)) Then
                        .TextMatrix(.Row, 售价列表.现采购限价) = FormatEx(dbl现价, mintPriceDigit)
                    End If
                    
                    Call CaculateCost(Val(.TextMatrix(n, 售价列表.药品id)), dbl现价)
                End If
            End If
        Next
    End With
End Sub

Private Sub CaculateCost(ByVal lng药品ID As Long, ByVal dbl现成本价 As Double)
    Dim n As Integer
    Dim dbl发票金额 As Double
    
    With BillStore
        For n = 1 To .Rows - 1
            If .TextMatrix(n, 库存列表.药品id) <> "" Then
                If Val(.TextMatrix(n, 库存列表.药品id)) = lng药品ID Then
                    .TextMatrix(n, 库存列表.现成本价) = FormatEx(dbl现成本价, mintCostDigit)
                    If dbl现成本价 <> 0 Then
                        .TextMatrix(n, 库存列表.加成率) = FormatEx((Val(.TextMatrix(n, 库存列表.现价)) / dbl现成本价 - 1) * 100, 5)
                    End If
                    If cbo售价计算方式 = "售价按分段加成计算" Then
                        .TextMatrix(n, 库存列表.加成率) = FormatEx(mdbl分段加成率 * 100, 5)
                    End If
                    
                    .TextMatrix(n, 库存列表.差价差) = Format((dbl现成本价 - .TextMatrix(n, 库存列表.原成本价)) * Val(.TextMatrix(n, 库存列表.数量)), mstrMoneyFormat)
                        
                    dbl发票金额 = dbl发票金额 + (dbl现成本价 - .TextMatrix(n, 库存列表.原成本价)) * Val(.TextMatrix(n, 库存列表.数量))
                     
                    If (cbo售价计算方式 = "售价按分段加成计算" Or cbo售价计算方式 = "售价按固定比例计算") And BillPrice.TextMatrix(BillPrice.Row, 售价列表.类型) = "时价" And mint调价 = 2 Then
                        .TextMatrix(n, 库存列表.现价) = BillPrice.TextMatrix(BillPrice.Row, 售价列表.现价)
                    End If
                End If
            End If
        Next
    End With
    
    If chk自动计算应付款变动.Value = 1 Then
        For n = 1 To BillPay.Rows - 1
            If BillPay.TextMatrix(1, 0) <> "" Then
                If Val(BillPay.TextMatrix(n, 应付款列.药品id)) = lng药品ID Then
                    BillPay.TextMatrix(n, 应付款列.发票金额) = Format(dbl发票金额, mstrMoneyFormat)
                End If
            End If
        Next
    End If
End Sub

Private Sub CaluateAverCost(ByVal lng药品ID As Long)
    '计算平均成本价
    Dim i As Integer
    Dim dblSumCost As Double
    Dim dblSumNumber As Double
    
    With BillStore
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 库存列表.药品id) <> "" Then
                If Val(.TextMatrix(i, 库存列表.药品id)) = lng药品ID Then
                    dblSumCost = dblSumCost + Val(.TextMatrix(i, 库存列表.现成本价)) * Val(.TextMatrix(i, 库存列表.数量))
                    dblSumNumber = dblSumNumber + Val(.TextMatrix(i, 库存列表.数量))
                End If
            End If
        Next
    End With
    
    With BillPrice
        If dblSumNumber > 0 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 售价列表.药品id) <> "" Then
                    If Val(.TextMatrix(i, 售价列表.药品id)) = lng药品ID Then
                        .TextMatrix(i, 售价列表.现成本价) = FormatEx(dblSumCost / dblSumNumber, mintCostDigit)
                        Exit For
                    End If
                End If
            Next
        End If
    End With
End Sub

Private Sub ChangeDrugStore(ByVal intRow As Integer, ByVal lngDrugId As Long, ByVal dblNewPrice As Double)
    Dim dblOldPrice As Double
    Dim dblOldCost As Double
    Dim dblNewCost As Double
    Dim dblNum As Double
    Dim dbl包装 As Double
    Dim n As Integer
    Dim dbl发票金额 As Double
    
    If intRow = 0 Or mint调价 = 1 Then Exit Sub
    
    dblOldPrice = Val(BillPrice.TextMatrix(intRow, 售价列表.原价))
    dbl包装 = GetModulus(lngDrugId)
    
    With BillStore
        For n = 1 To .Rows - 1
            If .TextMatrix(n, 0) <> "" Then
                If Val(.TextMatrix(n, 库存列表.药品id)) = lngDrugId Then
                    dblNum = Val(.TextMatrix(n, 库存列表.数量))
                    
                    .TextMatrix(n, 库存列表.现价) = FormatEx(dblNewPrice, mintPriceDigit)
                    .TextMatrix(n, 库存列表.调整金额) = Format(Val(.TextMatrix(n, 库存列表.数量)) * (dblNewPrice - dblOldPrice), mstrMoneyFormat)
                    
                    If mint调价 = 2 And chk自动调成本价.Value = 1 Then
                        dblOldCost = .TextMatrix(n, 库存列表.原成本价)
                        dblNewCost = dblNewPrice / (1 + Round(Val(.TextMatrix(n, 库存列表.加成率)) / 100, 7))
                        .TextMatrix(n, 库存列表.现成本价) = FormatEx(dblNewCost, mintCostDigit)
                        .TextMatrix(n, 库存列表.差价差) = Format((dblNewCost - dblOldCost) * dblNum, mstrMoneyFormat)
                        dbl发票金额 = dbl发票金额 + (dblNewCost - dblOldCost) * dblNum
                    End If
                End If
            End If
        Next
    End With
    
    If chk自动计算应付款变动.Value = 1 Then
        With BillPay
            For n = 1 To .Rows - 1
                If .TextMatrix(1, 0) <> "" Then
                    If Val(.TextMatrix(n, 应付款列.药品id)) = lngDrugId Then
                        .TextMatrix(n, 应付款列.发票金额) = FormatEx(dbl发票金额, 2)
                    End If
                End If
            Next
        End With
    End If
    
    CaluateAverCost lngDrugId
End Sub

Private Function CheckUnVerify(ByVal lng药品ID As Long) As Boolean
    '检查药品是否存在未审核单据
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSql = "Select 1 From 药品收发记录 Where 药品id = [1] And Rownum = 1 And 审核日期 Is Null"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "检查药品是否存在未审核单据", lng药品ID)
    
    If rsTemp.RecordCount > 0 Then
        CheckUnVerify = True
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetBatchData(ByVal BlnAll As Boolean)
    Dim lngRow As Long
    Dim n As Long
    Dim blnRepeat As Boolean
    
    For lngRow = 1 To vsfSpec.Rows - 1
        blnRepeat = False
        
        If Val(vsfSpec.TextMatrix(lngRow, vsfSpec.ColIndex("药品id"))) > 0 Then
            If Val(vsfSpec.TextMatrix(lngRow, vsfSpec.ColIndex("选择"))) <> 0 Or BlnAll = True Then
                For n = 1 To BillPrice.Rows - 1
                    If BillPrice.TextMatrix(n, 售价列表.药品id) <> "" Then
                        If Val(BillPrice.TextMatrix(n, 售价列表.药品id)) = Val(vsfSpec.TextMatrix(lngRow, vsfSpec.ColIndex("药品id"))) Then
                            blnRepeat = True
                            Exit For
                        End If
                    End If
                Next
                
                '不重复则增加
                If blnRepeat = False Then


                    With BillPrice
                        If .TextMatrix(.Rows - 1, 售价列表.药品id) <> "" Then
                            .Rows = .Rows + 1
                        End If
                        .TextMatrix(.Rows - 1, 售价列表.药品id) = vsfSpec.TextMatrix(lngRow, vsfSpec.ColIndex("药品id"))
                        .TextMatrix(.Rows - 1, 售价列表.品名) = vsfSpec.TextMatrix(lngRow, vsfSpec.ColIndex("药品"))
                        .TextMatrix(.Rows - 1, 售价列表.规格) = vsfSpec.TextMatrix(lngRow, vsfSpec.ColIndex("规格"))
                        .TextMatrix(.Rows - 1, 售价列表.产地) = vsfSpec.TextMatrix(lngRow, vsfSpec.ColIndex("产地"))
                        .TextMatrix(.Rows - 1, 售价列表.单位) = vsfSpec.TextMatrix(lngRow, vsfSpec.ColIndex("单位"))
                        .TextMatrix(.Rows - 1, 售价列表.类型) = vsfSpec.TextMatrix(lngRow, vsfSpec.ColIndex("类型"))
                        .TextMatrix(.Rows - 1, 售价列表.原成本价) = vsfSpec.TextMatrix(lngRow, vsfSpec.ColIndex("成本价"))
                        .TextMatrix(.Rows - 1, 售价列表.现成本价) = vsfSpec.TextMatrix(lngRow, vsfSpec.ColIndex("成本价"))
                        .TextMatrix(.Rows - 1, 售价列表.原采购限价) = vsfSpec.TextMatrix(lngRow, vsfSpec.ColIndex("采购限价"))
                        .TextMatrix(.Rows - 1, 售价列表.现采购限价) = vsfSpec.TextMatrix(lngRow, vsfSpec.ColIndex("采购限价"))
                        .TextMatrix(.Rows - 1, 售价列表.原指导售价) = vsfSpec.TextMatrix(lngRow, vsfSpec.ColIndex("指导售价"))
                        .TextMatrix(.Rows - 1, 售价列表.现指导售价) = vsfSpec.TextMatrix(lngRow, vsfSpec.ColIndex("指导售价"))
                        .TextMatrix(.Rows - 1, 售价列表.剂量系数) = vsfSpec.TextMatrix(lngRow, vsfSpec.ColIndex("剂量系数"))
                        .TextMatrix(.Rows - 1, 售价列表.药名ID) = vsfSpec.TextMatrix(lngRow, vsfSpec.ColIndex("药名ID"))
                        .TextMatrix(.Rows - 1, 售价列表.包装系数) = vsfSpec.TextMatrix(lngRow, vsfSpec.ColIndex("包装系数"))
                        .TextMatrix(.Rows - 1, 售价列表.差价让利比) = vsfSpec.TextMatrix(lngRow, vsfSpec.ColIndex("差价让利比"))
                        .TextMatrix(.Rows - 1, 售价列表.加成率) = vsfSpec.TextMatrix(lngRow, vsfSpec.ColIndex("加成率"))
                        
                        Call zlGetPrice(.Rows - 1, .TextMatrix(.Rows - 1, 售价列表.药品id), IIf(.TextMatrix(.Rows - 1, 售价列表.类型) = "时价", True, False))
                        
                        DoEvents
                        
                        Call GetDrugStore(.Rows - 1, Val(.TextMatrix(.Rows - 1, 售价列表.药品id)))
                    End With
                    
                    DoEvents
                End If
            End If
        End If
    Next
End Sub

Private Sub GetDrugStore(ByVal intRow As Integer, ByVal lngAddDrugId As Long, Optional ByVal lngDelDrugId As Long = 0)
    Dim n As Integer
    Dim intRows As Integer
    Dim dbl包装 As Double
    Dim dblOldPrice As Double
    Dim dblNewPrice As Double
    Dim strSql供应商ID As String
    Dim dbl加成率 As Double
    Dim dblOldCost As Double
    Dim dblNewCost As Double
    Dim str药品名称 As String
    Dim dbl发票金额 As Double
    
    On Error GoTo errHandle
    If lngDelDrugId > 0 Then
        With BillStore
            For n = .Rows - 1 To 1 Step -1
                If Val(.TextMatrix(n, 库存列表.药品id)) = lngDelDrugId Then
                    .MsfObj.RemoveItem n
                End If
            Next
        End With
        
        If mint调价 = 1 Or mint调价 = 2 Then
            With BillPay
                For n = .Rows - 1 To 1 Step -1
                    If Val(.TextMatrix(n, 应付款列.药品id)) = lngDelDrugId Then
                       .MsfObj.RemoveItem n
                    End If
                Next
            End With
        End If
    End If
    
    If lngAddDrugId = 0 Then Exit Sub
    
    With BillStore
        .Active = True
        dbl包装 = GetModulus(lngAddDrugId)
        dblOldPrice = Val(BillPrice.TextMatrix(intRow, 售价列表.原价))
        dblNewPrice = Val(BillPrice.TextMatrix(intRow, 售价列表.现价))
        
        If mint调价 = 1 Or mint调价 = 2 Then
            strSql供应商ID = IIf(mlng供应商ID = 0, "", " And S.上次供应商ID=[2] ")
        End If
            
        gstrSql = "select S.库房ID,D.名称 as 库房,'['||M.编码||']'||M.名称 as 药品,M.规格,M.产地,M.计算单位 售价单位,p.药库单位,S.批号,S.数量,S.批次, Nvl(M.是否变价, 0) 变价, M.ID, S.时价售价,P.指导差价率 As 差价率,S.成本价,S.上次供应商ID, N.名称 As 供应商,S.效期,S.产地 " & _
            " from (select S.库房ID,S.药品ID,S.上次供应商ID,S.上次批号 批号,S.效期,S.上次产地 As 产地,S.实际数量 as 数量,S.批次, Decode(Nvl(S.批次,0),0,Nvl(S.实际金额,0) / S.实际数量,Nvl(S.零售价,Nvl(S.实际金额,0) / S.实际数量)) 时价售价, s.平均成本价 As 成本价" & _
            "       from 药品库存 S" & _
            "       where S.性质=1 and S.实际数量<>0 and S.药品id=[1] ) S, " & _
            "      部门表 D,收费项目目录 M,药品规格 P, 供应商 N " & _
            " where D.id=S.库房id and S.药品ID=M.ID And M.ID=P.药品ID And Nvl(S.上次供应商id, 0) = N.ID(+) " & _
            " order by 库房,S.批号"
        
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngAddDrugId, mlng供应商ID)
            
'        If rsTemp.RecordCount = 0 And mint调价 = 1 Then
'            MsgBox "该药品无库存，不能调整成本价!", vbInformation, gstrSysName
'            If BillPrice.Rows = 2 Then
'                BillPrice.Rows = BillPrice.Rows + 1
'            End If
'            BillPrice.MsfObj.RemoveItem intRow
'            Exit Sub
'        End If
        
        intRows = .Rows - 1
        
        BillPrice.TextMatrix(intRow, 售价列表.是否有库存) = IIf(rsTemp.EOF, 0, 1)
        
        If mlng供应商ID > 0 Then
            rsTemp.Filter = "上次供应商ID=" & mlng供应商ID
        End If
        
        .Rows = .Rows + rsTemp.RecordCount
        
        Do While Not rsTemp.EOF
            .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, 库存列表.库房) = rsTemp!库房
            .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, 库存列表.供应商) = NVL(rsTemp!供应商)
            .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, 库存列表.药品) = rsTemp!药品
            str药品名称 = rsTemp!药品
            .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, 库存列表.规格) = IIf(IsNull(rsTemp!规格), "", rsTemp!规格)
            If int药库单位 = 0 Then
                .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, 库存列表.单位) = IIf(IsNull(rsTemp!售价单位), "", rsTemp!售价单位)
            Else
                .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, 库存列表.单位) = IIf(IsNull(rsTemp!药库单位), "", rsTemp!药库单位)
            End If
            .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, 库存列表.数量) = FormatEx(rsTemp!数量 / dbl包装, mintNumberDigit)
            .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, 库存列表.批号) = IIf(IsNull(rsTemp!批号), "", rsTemp!批号)
            .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, 库存列表.效期) = IIf(IsNull(rsTemp!效期), "", Format(rsTemp!效期, "yyyy-mm-dd"))
            .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, 库存列表.产地) = IIf(IsNull(rsTemp!产地), "", rsTemp!产地)
            .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, 库存列表.现价) = FormatEx(dblNewPrice, mintPriceDigit)
            .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, 库存列表.批次) = NVL(rsTemp!批次, 0)
            .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, 库存列表.变价) = rsTemp!变价
            .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, 库存列表.药品id) = rsTemp!ID
            .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, 库存列表.库房id) = rsTemp!库房id
            .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, 库存列表.供应商ID) = IIf(mlng供应商ID > 0, mlng供应商ID, NVL(rsTemp!上次供应商ID))
            If mint调价 = 1 Or mint调价 = 2 Then
                dblOldCost = FormatEx(rsTemp!成本价 * dbl包装, mintCostDigit)
               
                If mdbl加成率 > 0 Then
                    dbl加成率 = Round(mdbl加成率 / 100, 7)
                ElseIf dblOldCost > 0 Then
                    dbl加成率 = Round(dblOldPrice / dblOldCost - 1, 7)
                Else
                    dbl加成率 = Round(1 / (1 - rsTemp!差价率 / 100) - 1, 7)
                End If
                
'                If dblOldPrice = dblNewPrice Then
'                    dblNewCost = dblOldCost
'                Else
                    dblNewCost = dblNewPrice / (1 + dbl加成率)
'                End If
                
                .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, 库存列表.原价) = FormatEx(IIf(rsTemp!变价 = 1, rsTemp!时价售价 * dbl包装, dblOldPrice), mintPriceDigit)
                .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, 库存列表.调整金额) = Format(rsTemp!数量 / dbl包装 * (dblNewPrice - IIf(rsTemp!变价 = 1, rsTemp!时价售价 * dbl包装, dblOldPrice)), mstrMoneyFormat)
                .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, 库存列表.加成率) = dbl加成率 * 100
                .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, 库存列表.原成本价) = FormatEx(dblOldCost, mintCostDigit)
                .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, 库存列表.现成本价) = FormatEx(dblNewCost, mintCostDigit)
                .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, 库存列表.差价差) = Format((dblNewCost - dblOldCost) * Val(.TextMatrix(intRows + rsTemp.AbsolutePosition - 1, 库存列表.数量)), mstrMoneyFormat)
                dbl发票金额 = dbl发票金额 + (dblNewCost - dblOldCost) * Val(.TextMatrix(intRows + rsTemp.AbsolutePosition - 1, 库存列表.数量))
            Else
                .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, 库存列表.原价) = FormatEx(IIf(rsTemp!变价 = 1, rsTemp!时价售价 * dbl包装, dblOldPrice), mintPriceDigit)
                .TextMatrix(intRows + rsTemp.AbsolutePosition - 1, 库存列表.调整金额) = Format(rsTemp!数量 / dbl包装 * (dblNewPrice - IIf(rsTemp!变价 = 1, rsTemp!时价售价 * dbl包装, dblOldPrice)), mstrMoneyFormat)
            End If
            
            rsTemp.MoveNext
        Loop
    
    End With
    
    If mint调价 = 1 Or mint调价 = 2 Then
        With BillPay
            .Active = True
            .TextMatrix(.Rows - 1, 应付款列.药品id) = lngAddDrugId
            .TextMatrix(.Rows - 1, 应付款列.品名) = str药品名称
            .TextMatrix(.Rows - 1, 应付款列.发票金额) = FormatEx(dbl发票金额, 2)
            .Rows = .Rows + 1
        End With
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetItem(ByVal strKey As String)
    Dim vRect As RECT
    Dim strReturn As String
    Dim sngX As Single
    Dim sngY As Single
    Dim sngH As Single
    Dim blnCancel As Boolean
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    vRect = zlControl.GetControlRect(txtItem.hWnd)
    sngX = picItem.Left + vRect.Left - 100
    sngY = picItem.Top + vRect.Top + txtItem.Height + 175
    sngH = picItem.Height - vsfSpec.Top
    
    If strKey = "" Then
        gstrSql = "Select Distinct I.ID, '[' || I.编码 || ']' || I.名称 As 药品, I.计算单位 " & _
            " From 诊疗项目目录 I, 诊疗项目别名 N " & _
            " Where I.ID = N.诊疗项目id And I.类别 = '7' And (撤档时间 Is Null Or 撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd')) " & _
            " Order By '[' || I.编码 || ']' || I.名称"
        Set rsTemp = zldatabase.ShowSQLSelect(Me, gstrSql, 0, "药品选择器", False, "", "选择药品", False, False, True, sngX, sngY, sngH, blnCancel, False, False)
    Else
        gstrSql = "Select Distinct I.ID, '[' || I.编码 || ']' || I.名称 As 药品, I.计算单位 " & _
            " From 诊疗项目目录 I, 诊疗项目别名 N " & _
            " Where I.ID = N.诊疗项目id And I.类别 = '7' And (撤档时间 Is Null Or 撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd')) And (I.编码 Like [1] Or N.名称 Like [2] Or N.简码 Like [2]) " & _
            " Order By '[' || I.编码 || ']' || I.名称"
         Set rsTemp = zldatabase.ShowSQLSelect(Me, gstrSql, 0, "药品选择器", False, "", "选择药品", False, False, True, sngX, sngY, sngH, blnCancel, False, False, UCase(strKey) & "%", "%" & UCase(strKey) & "%")
    End If
    
    If blnCancel = True Then Exit Sub

    If Not rsTemp Is Nothing Then
        Call GetSpec(Val(rsTemp!ID), 2)
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetSpec(ByVal lngItem As Long, ByVal intType As Integer)
    Dim rsTemp As ADODB.Recordset
    Dim lngRow As Long
    Dim dbl包装 As Double
    
    On Error GoTo errHandle
    gstrSql = "Select Distinct I.ID, I.编码, I.名称, I.规格, I.产地, I.计算单位, P.药库单位, Decode(I.是否变价, 1, '时价', '定价') 类型, Nvl(P.成本价, 0) 成本价," & _
        " P.指导批发价 , P.指导零售价, Z.名称 As 品种, P.剂量系数, P.药名ID,p.差价让利比,1/(1-p.指导差价率/100)-1  加成率" & _
        " From 收费项目目录 I, 收费项目别名 N, 药品规格 P, 诊疗项目目录 Z " & _
        " Where I.ID = N.收费细目id And I.类别 In (" & IIf(intType = 1, "'5','6'", "'7'") & ") And I.ID = P.药品id And P.药名id = Z.ID And " & _
        " (I.撤档时间 Is Null Or I.撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd')) And P.药名id = [1] " & _
        " Order By I.编码 "
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "取药品规格", lngItem)
    
    With vsfSpec
        .Redraw = flexRDNone
'        .Rows = 1
'        .Rows = 2
        lngRow = .Rows - 1
        
        mblnAllUnAdj = True
        
        If rsTemp.RecordCount > 0 Then
            Do While Not rsTemp.EOF
                If Check存在未执行价格(Val(rsTemp!ID)) = False Then
                    mblnAllUnAdj = False
                    
                    dbl包装 = GetModulus(Val(rsTemp!ID))
                    
                    .TextMatrix(lngRow, .ColIndex("药品id")) = rsTemp!ID
                    .TextMatrix(lngRow, .ColIndex("品种")) = rsTemp!品种
                    .TextMatrix(lngRow, .ColIndex("药品")) = "[" & rsTemp!编码 & "]" & rsTemp!名称
                    .TextMatrix(lngRow, .ColIndex("规格")) = IIf(IsNull(rsTemp!规格), "", rsTemp!规格)
                    .TextMatrix(lngRow, .ColIndex("产地")) = IIf(IsNull(rsTemp!产地), "", rsTemp!产地)
                    
                    If int药库单位 = 0 Then
                        .TextMatrix(lngRow, .ColIndex("单位")) = IIf(IsNull(rsTemp!计算单位), "", rsTemp!计算单位)
                    Else
                        .TextMatrix(lngRow, .ColIndex("单位")) = IIf(IsNull(rsTemp!药库单位), "", rsTemp!药库单位)
                    End If
                    
                    .TextMatrix(lngRow, .ColIndex("类型")) = IIf(IsNull(rsTemp!类型), "", rsTemp!类型)
                    .TextMatrix(lngRow, .ColIndex("成本价")) = FormatEx(Val(IIf(IsNull(rsTemp!成本价), "", rsTemp!成本价)) * dbl包装, mintCostDigit)
                    .TextMatrix(lngRow, .ColIndex("采购限价")) = FormatEx(Val(IIf(IsNull(rsTemp!指导批发价), "", rsTemp!指导批发价)) * dbl包装, mintCostDigit)
                    .TextMatrix(lngRow, .ColIndex("指导售价")) = FormatEx(Val(IIf(IsNull(rsTemp!指导零售价), "", rsTemp!指导零售价)) * dbl包装, mintPriceDigit)
                    .TextMatrix(lngRow, .ColIndex("剂量系数")) = rsTemp!剂量系数
                    .TextMatrix(lngRow, .ColIndex("药名ID")) = rsTemp!药名ID
                    .TextMatrix(lngRow, .ColIndex("包装系数")) = dbl包装
                    .TextMatrix(lngRow, .ColIndex("差价让利比")) = IIf(IsNull(rsTemp!差价让利比), 0, rsTemp!差价让利比)
                    .TextMatrix(lngRow, .ColIndex("加成率")) = rsTemp!加成率
                                                        
                    .Rows = .Rows + 1
                    lngRow = .Rows - 1
                Else
                    mstrAdjMsg = IIf(mstrAdjMsg = "", "", mstrAdjMsg & vbCrLf) & "[" & rsTemp!编码 & "]" & rsTemp!名称
                End If
                
                rsTemp.MoveNext
            Loop
        End If
        
        .Redraw = flexRDDirect
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub IniBatchData()
    Dim strToday As String
    
    '进入调价编辑状态
    Me.BillPrice.Active = True
    
    strToday = Format(Sys.Currentdate(), "yyyy-MM-dd hh:mm:ss")
    
    Me.lblTitle.Caption = "库存变动表：(由于调价未保存，反映的库存可能不准确)"
    Me.dtpRunDate.MinDate = DateAdd("s", 1, CDate(strToday))
    Me.dtpRunDate.Value = DateAdd("d", 1, CDate(strToday))
    Me.txtValuer.Text = gstrUserName
    
    Call GetSpec(lngItemID, intDrugType)
    
    If mstrAdjMsg <> "" Then
        MsgBox "以下药品存在未执行价格，不能再进行调价操作：" & vbCrLf & mstrAdjMsg, vbInformation, gstrSysName
    End If
    
    If mblnAllUnAdj = True Then
        '如果所有规格都存在未执行价格，则退出
        Unload Me
    Else
        Call GetBatchData(True)
    End If
End Sub

Private Sub IniData()
    Dim strToday As String
    Dim dbl包装 As Double
    
    On Error GoTo errHandle
    strToday = Format(Sys.Currentdate(), "yyyy-MM-dd hh:mm:ss")
    If lngBillId = 0 Then
        '进入调价编辑状态
        Me.BillPrice.Active = True
        
        Me.lblTitle.Caption = "库存变动表：(由于调价未保存，反映的库存可能不准确)"
        Me.dtpRunDate.MinDate = DateAdd("s", 1, CDate(strToday))
        Me.dtpRunDate.Value = DateAdd("d", 1, CDate(strToday))
        Me.txtValuer.Text = gstrUserName
        
        If lngMediId = 0 Then Exit Sub
        
        If Check存在未执行价格(lngMediId) = True Then
            MsgBox "该药品存在未执行价格，不能再进行调价操作!", vbInformation, gstrSysName
            Unload Me
            Exit Sub
        End If
        
        dbl包装 = GetModulus(lngMediId)
        
        '如果指定首先调价的药品，则直接将该药品调入
        gstrSql = "select P.药名ID,P.剂量系数,I.ID,I.编码,I.名称,I.规格,I.产地,I.计算单位,P.药库单位,decode(I.是否变价,1,'时价','定价') 类型,Nvl(P.成本价,0) 成本价,P.指导批发价,P.指导零售价,p.差价让利比,1/(1-p.指导差价率/100)-1 加成率" & _
                 " from 收费项目目录 I,药品规格 P" & _
                 " where I.ID=[1] And I.ID=P.药品ID"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngMediId)
        
        With rsTemp
            If .BOF Or .EOF Then Exit Sub
            Me.BillPrice.TextMatrix(.AbsolutePosition, 售价列表.药品id) = !ID
            Me.BillPrice.TextMatrix(.AbsolutePosition, 售价列表.品名) = "[" & !编码 & "]" & !名称
            Me.BillPrice.TextMatrix(.AbsolutePosition, 售价列表.规格) = IIf(IsNull(!规格), "", !规格)
            Me.BillPrice.TextMatrix(.AbsolutePosition, 售价列表.产地) = IIf(IsNull(!产地), "", !产地)
            Me.BillPrice.TextMatrix(.AbsolutePosition, 售价列表.原成本价) = FormatEx(Val(IIf(IsNull(!成本价), "", !成本价)) * dbl包装, mintCostDigit)
            Me.BillPrice.TextMatrix(.AbsolutePosition, 售价列表.现成本价) = FormatEx(Val(IIf(IsNull(!成本价), "", !成本价)) * dbl包装, mintCostDigit)
            Me.BillPrice.TextMatrix(.AbsolutePosition, 售价列表.原采购限价) = FormatEx(Val(IIf(IsNull(!指导批发价), "", !指导批发价)) * dbl包装, mintPriceDigit)
            Me.BillPrice.TextMatrix(.AbsolutePosition, 售价列表.现采购限价) = FormatEx(Val(IIf(IsNull(!指导批发价), "", !指导批发价)) * dbl包装, mintPriceDigit)
            Me.BillPrice.TextMatrix(.AbsolutePosition, 售价列表.原指导售价) = FormatEx(Val(IIf(IsNull(!指导零售价), "", !指导零售价)) * dbl包装, mintPriceDigit)
            Me.BillPrice.TextMatrix(.AbsolutePosition, 售价列表.现指导售价) = FormatEx(Val(IIf(IsNull(!指导零售价), "", !指导零售价)) * dbl包装, mintPriceDigit)
            Me.BillPrice.TextMatrix(.AbsolutePosition, 售价列表.药名ID) = !药名ID
            Me.BillPrice.TextMatrix(.AbsolutePosition, 售价列表.剂量系数) = !剂量系数
            Me.BillPrice.TextMatrix(.AbsolutePosition, 售价列表.包装系数) = dbl包装
            Me.BillPrice.TextMatrix(.AbsolutePosition, 售价列表.差价让利比) = IIf(IsNull(!差价让利比), 0, !差价让利比)
            Me.BillPrice.TextMatrix(.AbsolutePosition, 售价列表.加成率) = !加成率
            
            If int药库单位 = 0 Then
                Me.BillPrice.TextMatrix(.AbsolutePosition, 售价列表.单位) = IIf(IsNull(!计算单位), "", !计算单位)
            Else
                Me.BillPrice.TextMatrix(.AbsolutePosition, 售价列表.单位) = IIf(IsNull(!药库单位), "", !药库单位)
            End If
            Me.BillPrice.TextMatrix(.AbsolutePosition, 售价列表.类型) = IIf(IsNull(!类型), "", !类型)
            
            Call zlGetPrice(.AbsolutePosition, lngMediId, IIf(!类型 = "时价", True, False))
            
            If mint调价 = 0 Then
                Me.BillPrice.Col = 售价列表.现价
            ElseIf mint调价 = 1 Or mint调价 = 2 Then
                Me.BillPrice.Col = 售价列表.现成本价
            ElseIf mint调价 = 3 Then
                Me.BillPrice.Col = 售价列表.收入名称
            End If
        
            Call GetDrugStore(1, lngMediId)
            
'            If mint调价 = 1 Or mint调价 = 2 Then
'                If BillStore.Rows = 1 Then
'                    Me.BillPrice.TextMatrix(.AbsolutePosition, 售价列表.是否有库存) = 0
'                ElseIf BillStore.TextMatrix(1, 0) = "" Then
'                    Me.BillPrice.TextMatrix(.AbsolutePosition, 售价列表.是否有库存) = 0
'                Else
'                    Me.BillPrice.TextMatrix(.AbsolutePosition, 售价列表.是否有库存) = 1
'                End If
'            End If
        End With
    Else
        '进入调价显示状态
        Me.BillPrice.Active = False
        Me.BillStore.Active = False
        Me.BillPay.Active = False
        Me.cmdOk.Visible = False
        Me.cmdCanc.Caption = "返回(&C)"
        Me.cmdCanc.Top = Me.cmdOk.Top
        Me.txtSummary.Enabled = False
        opt时间(1).Value = True
        opt时间(0).Enabled = False
        opt时间(1).Enabled = False
        Me.dtpRunDate.Enabled = False
        Me.chk自动计算应付款变动.Enabled = False
        Me.chk按批次.Enabled = False
        
        Dim strBills As String
        strBills = ""
        
        gstrSql = "select P.ID,M.id as 药品id,'['||M.编码||']'||M.名称 as 品名,M.规格,M.产地,M.计算单位 as 单位,P.药库单位," & _
            "        P.原价,P.现价,P.收入项目id,I.名称 as 收入名称," & _
            "        To_Char(P.执行日期,'yyyy-MM-dd hh24:mi:ss') 执行日期,P.变动原因,P.调价说明,P.调价人,p.差价让利比,1/(1-p.指导差价率/100)-1 加成率," & _
            " from 收费价目 P,收费项目目录 M,收入项目 I,药品规格 P" & _
            " where P.收费细目id=M.id and P.收入项目id=I.id And M.ID=P.药品ID and P.ID=[1] " & _
            GetPriceClassString("P") & _
            " order by P.id"                            '因调价ID取的是价格记录ID的上一个ID
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngBillId)
        
        With rsTemp
            Me.BillPrice.Rows = .RecordCount + 1
            Do While Not .EOF
                dbl包装 = GetModulus(Val(!药品id))
                
                strBills = strBills & "," & !ID
                Me.BillPrice.TextMatrix(.AbsolutePosition, 售价列表.药品id) = !药品id
                Me.BillPrice.TextMatrix(.AbsolutePosition, 售价列表.品名) = !品名
                Me.BillPrice.TextMatrix(.AbsolutePosition, 售价列表.规格) = IIf(IsNull(!规格), "", !规格)
                Me.BillPrice.TextMatrix(.AbsolutePosition, 售价列表.产地) = IIf(IsNull(!产地), "", !产地)
                If int药库单位 = 0 Then
                    Me.BillPrice.TextMatrix(.AbsolutePosition, 售价列表.单位) = IIf(IsNull(!单位), "", !单位)
                Else
                    Me.BillPrice.TextMatrix(.AbsolutePosition, 售价列表.单位) = IIf(IsNull(!药库单位), "", !药库单位)
                End If
                Me.BillPrice.TextMatrix(.AbsolutePosition, 售价列表.原价) = FormatEx(!原价 * dbl包装, mintPriceDigit)
                Me.BillPrice.TextMatrix(.AbsolutePosition, 售价列表.现价) = FormatEx(!现价 * dbl包装, mintPriceDigit)
                Me.BillPrice.TextMatrix(.AbsolutePosition, 售价列表.现收入ID) = !收入项目id
                Me.BillPrice.TextMatrix(.AbsolutePosition, 售价列表.收入名称) = !收入名称
                Me.BillPrice.TextMatrix(.AbsolutePosition, 售价列表.差价让利比) = IIf(IsNull(!差价让利比), 0, !差价让利比)
                Me.BillPrice.TextMatrix(.AbsolutePosition, 售价列表.加成率) = !加成率
                
                Me.txtSummary = IIf(IsNull(!调价说明), "", !调价说明)
                Me.txtValuer.Text = IIf(IsNull(!调价人), "", !调价人)
                Me.dtpRunDate.Value = !执行日期
                
                If !执行日期 <= strToday And !变动原因 = 0 Then        '未进行调价计算,则执行计算
                    gstrSql = "zl_药品收发记录_Adjust(" & !ID & ")"
                    Call zldatabase.ExecuteProcedure(gstrSql, Me.Caption)
                End If
                .MoveNext
            Loop
            If .RecordCount <> 0 Then .MoveFirst
            
            If !执行日期 > strToday Then
                '如果执行时间未到，则只能模拟显示库存变动
                Me.lblTitle.Caption = "库存变动表：(由于执行时间未到，反映的库存可能不准确)"
            Else
                '执行时间已到，肯定也进行了调价计算，直接从收发记录提取调价变动情况
                Me.lblTitle.Caption = "库存变动表："
                gstrSql = "select S.ID,S.药品ID,D.名称 as 库房,'['||M.编码||']'||M.名称 as 药品,M.规格,M.产地,M.计算单位 as 单位,P.药库单位,S.批号,S.数量,S.原价,S.现价,S.调整金额" & _
                        " from (select ID,库房ID,药品ID,批号,填写数量 as 数量,成本价 as 原价,零售价 as 现价,零售金额 as 调整金额" & _
                        "       from (select P.ID,N.库房ID,N.药品ID,N.批号,N.填写数量,N.成本价,N.零售价,N.零售金额" & _
                        "            from 药品收发记录 N, (select ID,收费细目ID,执行日期,终止日期 from 收费价目 where ID=[1]" & _
                        GetPriceClassString("") & ") P" & _
                        "       where N.药品ID=P.收费细目ID and 单据=13 and N.费用ID is null " & _
                        "             and N.审核日期 Between P.执行日期 and nvl(P.终止日期,sysdate))) S," & _
                        "       部门表 D,收费项目目录 M,药品规格 P" & _
                        " where S.库房id+0=D.id and S.药品ID=M.ID And M.ID=P.药品ID" & _
                        " order by M.编码,S.批号"
                Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Mid(strBills, 2)))
                    
                With rsTemp
                    If .RecordCount > 0 Then Me.BillStore.Rows = .RecordCount + 1
                    Do While Not .EOF
                        Me.BillStore.TextMatrix(.AbsolutePosition, 库存列表.库房) = !库房
                        Me.BillStore.TextMatrix(.AbsolutePosition, 库存列表.药品) = !药品
                        Me.BillStore.TextMatrix(.AbsolutePosition, 库存列表.规格) = IIf(IsNull(!规格), "", !规格)
                        If int药库单位 = 0 Then
                            Me.BillStore.TextMatrix(.AbsolutePosition, 库存列表.单位) = IIf(IsNull(!单位), "", !单位)
                        Else
                            Me.BillStore.TextMatrix(.AbsolutePosition, 库存列表.单位) = IIf(IsNull(!药库单位), "", !药库单位)
                        End If
                        Me.BillStore.TextMatrix(.AbsolutePosition, 库存列表.批号) = IIf(IsNull(!批号), "", !批号)
                        Me.BillStore.TextMatrix(.AbsolutePosition, 库存列表.数量) = Format(!数量 / dbl包装, "0.00000")
                        Me.BillStore.TextMatrix(.AbsolutePosition, 库存列表.原价) = FormatEx(!原价 * dbl包装, mintPriceDigit)
                        Me.BillStore.TextMatrix(.AbsolutePosition, 库存列表.现价) = FormatEx(!现价 * dbl包装, mintPriceDigit)
                        Me.BillStore.TextMatrix(.AbsolutePosition, 库存列表.调整金额) = Format(!调整金额, mstrMoneyFormat)
                        .MoveNext
                    Loop
                End With
            
            End If
            
        End With
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub IniGrid()
    With Me.BillPrice
        .Cols = 售价列表.列数
        .MsfObj.FixedCols = 0

        If intDrugType = 1 Then   '西药、成药
            .TextMatrix(0, 售价列表.药品id) = "药品id"
            .TextMatrix(0, 售价列表.品名) = "品名"
            .TextMatrix(0, 售价列表.规格) = "规格"
            .TextMatrix(0, 售价列表.产地) = "厂牌"
            .TextMatrix(0, 售价列表.单位) = "单位"
            .TextMatrix(0, 售价列表.类型) = "类型"
            .TextMatrix(0, 售价列表.上次日期) = "上次日期"
            .TextMatrix(0, 售价列表.原价) = "原零售价"
            .TextMatrix(0, 售价列表.现价) = "现零售价"
            .TextMatrix(0, 售价列表.现收入ID) = "收入id"
            .TextMatrix(0, 售价列表.原收入ID) = "原收入id"
            .TextMatrix(0, 售价列表.收入名称) = "收入项目"
            .TextMatrix(0, 售价列表.原成本价) = IIf(mint调价 = 1 Or mint调价 = 2, "原采购价", "成本价")
            .TextMatrix(0, 售价列表.现成本价) = "现采购价"
            .TextMatrix(0, 售价列表.原采购限价) = IIf(InStr(1, mstrPrivs, "指导价格管理") = 0, "采购限价", "原采购限价")
            .TextMatrix(0, 售价列表.现采购限价) = "现采购限价"
            .TextMatrix(0, 售价列表.原指导售价) = IIf(InStr(1, mstrPrivs, "指导价格管理") = 0, "指导售价", "原指导售价")
            .TextMatrix(0, 售价列表.现指导售价) = "现指导售价"
            .TextMatrix(0, 售价列表.是否有库存) = "是否有库存"
            .TextMatrix(0, 售价列表.剂量系数) = "剂量系数"
            .TextMatrix(0, 售价列表.药名ID) = "药名ID"
            .TextMatrix(0, 售价列表.包装系数) = "包装系数"
            .TextMatrix(0, 售价列表.差价让利比) = "差价让利比"
            .TextMatrix(0, 售价列表.加成率) = "加成率"
            
            .ColWidth(售价列表.药品id) = 0
            .ColWidth(售价列表.品名) = 2600
            .ColWidth(售价列表.规格) = 1200
            .ColWidth(售价列表.产地) = 1000
            .ColWidth(售价列表.单位) = 600
            .ColWidth(售价列表.类型) = 0
            .ColWidth(售价列表.上次日期) = 0
            .ColWidth(售价列表.原价) = 900
            .ColWidth(售价列表.现价) = 900
            .ColWidth(售价列表.现收入ID) = 0
            .ColWidth(售价列表.原收入ID) = 0
            .ColWidth(售价列表.收入名称) = 900
            .ColWidth(售价列表.原成本价) = 975
            .ColWidth(售价列表.现成本价) = IIf(mint调价 = 1 Or mint调价 = 2, 975, 0)
            .ColWidth(售价列表.原采购限价) = 0
            .ColWidth(售价列表.现采购限价) = 0 'IIf(InStr(1, mstrPrivs, "指导价格管理") = 0, 0, 1000)
            .ColWidth(售价列表.原指导售价) = 0
            .ColWidth(售价列表.现指导售价) = 0 'IIf(InStr(1, mstrPrivs, "指导价格管理") = 0, 0, 1000)
            .ColWidth(售价列表.是否有库存) = 0
            .ColWidth(售价列表.剂量系数) = 0
            .ColWidth(售价列表.药名ID) = 0
            .ColWidth(售价列表.包装系数) = 0
            .ColWidth(售价列表.差价让利比) = 0
            .ColWidth(售价列表.加成率) = 0
            
        Else    '中草药
            .TextMatrix(0, 售价列表.药品id) = "药品id"
            .TextMatrix(0, 售价列表.品名) = "品名"
            .TextMatrix(0, 售价列表.规格) = "规格"
            .TextMatrix(0, 售价列表.产地) = "产地"
            .TextMatrix(0, 售价列表.单位) = "单位"
            .TextMatrix(0, 售价列表.类型) = "类型"
            .TextMatrix(0, 售价列表.上次日期) = "上次日期"
            .TextMatrix(0, 售价列表.原价) = "原零售价"
            .TextMatrix(0, 售价列表.现价) = "现零售价"
            .TextMatrix(0, 售价列表.现收入ID) = "收入id"
            .TextMatrix(0, 售价列表.原收入ID) = "原收入id"
            .TextMatrix(0, 售价列表.收入名称) = "收入项目"
            .TextMatrix(0, 售价列表.原成本价) = IIf(mint调价 = 1 Or mint调价 = 2, "原采购价", "成本价")
            .TextMatrix(0, 售价列表.现成本价) = "现采购价"
            .TextMatrix(0, 售价列表.原采购限价) = "原采购限价"
            .TextMatrix(0, 售价列表.现采购限价) = "现采购限价"
            .TextMatrix(0, 售价列表.原指导售价) = "原指导售价"
            .TextMatrix(0, 售价列表.现指导售价) = "现指导售价"
            .TextMatrix(0, 售价列表.是否有库存) = "是否有库存"
            .TextMatrix(0, 售价列表.剂量系数) = "剂量系数"
            .TextMatrix(0, 售价列表.药名ID) = "药名ID"
            .TextMatrix(0, 售价列表.包装系数) = "包装系数"
            .TextMatrix(0, 售价列表.差价让利比) = "差价让利比"
            .TextMatrix(0, 售价列表.加成率) = "加成率"
                        
            .ColWidth(售价列表.药品id) = 0
            .ColWidth(售价列表.品名) = 2800
            .ColWidth(售价列表.规格) = 1200
            .ColWidth(售价列表.产地) = 1000
            .ColWidth(售价列表.单位) = 600
            .ColWidth(售价列表.类型) = 0
            .ColWidth(售价列表.上次日期) = 0
            .ColWidth(售价列表.原价) = 1200
            .ColWidth(售价列表.现价) = 1200
            .ColWidth(售价列表.现收入ID) = 0
            .ColWidth(售价列表.原收入ID) = 0
            .ColWidth(售价列表.收入名称) = 1200
            .ColWidth(售价列表.原成本价) = 975
            .ColWidth(售价列表.现成本价) = IIf(mint调价 = 1 Or mint调价 = 2, 975, 0)
            .ColWidth(售价列表.原采购限价) = 0
            .ColWidth(售价列表.现采购限价) = 0 'IIf(InStr(1, mstrPrivs, "指导价格管理") = 0, 0, 1000)
            .ColWidth(售价列表.原指导售价) = 0
            .ColWidth(售价列表.现指导售价) = 0 'IIf(InStr(1, mstrPrivs, "指导价格管理") = 0, 0, 1000)
            .ColWidth(售价列表.是否有库存) = 0
            .ColWidth(售价列表.剂量系数) = 0
            .ColWidth(售价列表.药名ID) = 0
            .ColWidth(售价列表.包装系数) = 0
            .ColWidth(售价列表.差价让利比) = 0
            .ColWidth(售价列表.加成率) = 0
        End If
        
        If lngBillId <> 0 Then
            .ColWidth(售价列表.原成本价) = 0
            .ColWidth(售价列表.现成本价) = 0
        End If
        
        .ColData(售价列表.药品id) = 5
        .ColData(售价列表.品名) = 1
        .ColData(售价列表.规格) = 5
        .ColData(售价列表.产地) = 5
        .ColData(售价列表.单位) = 5
        .ColData(售价列表.类型) = 5
        .ColData(售价列表.上次日期) = 5
        .ColData(售价列表.原价) = 5
        .ColData(售价列表.现价) = IIf(mint调价 = 3, 5, IIf(mint调价 = 1, 0, 4))
        .ColData(售价列表.现收入ID) = 5
        .ColData(售价列表.原收入ID) = 5
        .ColData(售价列表.收入名称) = 1
        .ColData(售价列表.原成本价) = 5
        .ColData(售价列表.现成本价) = IIf(mint调价 = 1 Or mint调价 = 2, 4, 0)
        .ColData(售价列表.原采购限价) = 5
        .ColData(售价列表.现采购限价) = IIf(mint调价 = 3, 5, IIf(InStr(1, mstrPrivs, "指导价格管理") = 0, 5, 4))
        .ColData(售价列表.原指导售价) = 5
        .ColData(售价列表.现指导售价) = IIf(mint调价 = 3, 5, IIf(InStr(1, mstrPrivs, "指导价格管理") = 0, 5, 4))

        .ColAlignment(售价列表.药品id) = 1
        .ColAlignment(售价列表.品名) = 1
        .ColAlignment(售价列表.规格) = 1
        .ColAlignment(售价列表.产地) = 1
        .ColAlignment(售价列表.单位) = 4
        .ColAlignment(售价列表.类型) = 1
        .ColAlignment(售价列表.上次日期) = 1
        .ColAlignment(售价列表.原价) = 7
        .ColAlignment(售价列表.现价) = 7
        .ColAlignment(售价列表.现收入ID) = 1
        .ColAlignment(售价列表.原收入ID) = 1
        .ColAlignment(售价列表.收入名称) = 1
        .ColAlignment(售价列表.原成本价) = 7
        .ColAlignment(售价列表.现成本价) = 7
        .ColAlignment(售价列表.原采购限价) = 7
        .ColAlignment(售价列表.现采购限价) = 7
        .ColAlignment(售价列表.原指导售价) = 7
        .ColAlignment(售价列表.现指导售价) = 7
        
        .PrimaryCol = 售价列表.品名
        .LocateCol = 售价列表.品名
    End With
    
    With Me.BillStore
        .Rows = 2
        .MsfObj.FixedCols = 0
        .Cols = 库存列表.列数
        .TextMatrix(0, 库存列表.库房) = "库房"
        .TextMatrix(0, 库存列表.供应商) = "供应商"
        .TextMatrix(0, 库存列表.药品) = "药品"
        .TextMatrix(0, 库存列表.规格) = "规格"
        .TextMatrix(0, 库存列表.单位) = "单位"
        .TextMatrix(0, 库存列表.批号) = "批号"
        .TextMatrix(0, 库存列表.效期) = "效期"
        .TextMatrix(0, 库存列表.产地) = "产地"
        .TextMatrix(0, 库存列表.数量) = "数量"
        .TextMatrix(0, 库存列表.原价) = "原零售价"
        .TextMatrix(0, 库存列表.现价) = "现零售价"
        .TextMatrix(0, 库存列表.调整金额) = "调整金额"
        .TextMatrix(0, 库存列表.加成率) = "加成率(%)"
        .TextMatrix(0, 库存列表.原成本价) = "原采购价"
        .TextMatrix(0, 库存列表.现成本价) = "现采购价"
        .TextMatrix(0, 库存列表.差价差) = "差价差"
        .TextMatrix(0, 库存列表.批次) = "批次"
        .TextMatrix(0, 库存列表.变价) = "变价"
        .TextMatrix(0, 库存列表.药品id) = "药品ID"
        .TextMatrix(0, 库存列表.库房id) = "库房ID"
        .TextMatrix(0, 库存列表.供应商ID) = "供应商ID"
        
        .ColData(库存列表.库房) = 5
        .ColData(库存列表.供应商) = 5
        .ColData(库存列表.药品) = 5
        .ColData(库存列表.规格) = 5
        .ColData(库存列表.单位) = 5
        .ColData(库存列表.批号) = 5
        .ColData(库存列表.效期) = 5
        .ColData(库存列表.产地) = 5
        .ColData(库存列表.数量) = 5
        .ColData(库存列表.原价) = 5
        .ColData(库存列表.现价) = 0
        .ColData(库存列表.调整金额) = 5
        .ColData(库存列表.加成率) = 4
        .ColData(库存列表.原成本价) = 5
        .ColData(库存列表.现成本价) = 4
        .ColData(库存列表.差价差) = 5
        .ColData(库存列表.批次) = 5
        .ColData(库存列表.变价) = 5
        .ColData(库存列表.药品id) = 5
        .ColData(库存列表.库房id) = 5
        .ColData(库存列表.供应商ID) = 5
        
        .ColWidth(库存列表.库房) = 1000
        .ColWidth(库存列表.供应商) = 1500
        .ColWidth(库存列表.药品) = 2800
        .ColWidth(库存列表.规格) = 1350
        .ColWidth(库存列表.单位) = 600
        .ColWidth(库存列表.批号) = 800
        .ColWidth(库存列表.效期) = 1000
        .ColWidth(库存列表.产地) = 1000
        .ColWidth(库存列表.数量) = 1000
        .ColWidth(库存列表.原价) = 900
        .ColWidth(库存列表.现价) = 900
        .ColWidth(库存列表.调整金额) = 1050
        .ColWidth(库存列表.批次) = 0
        .ColWidth(库存列表.变价) = 0
        .ColWidth(库存列表.药品id) = 0
        .ColWidth(库存列表.库房id) = 0
        .ColWidth(库存列表.供应商ID) = 0
        
        If mint调价 = 0 Then
            .ColWidth(库存列表.加成率) = 0
            .ColWidth(库存列表.原成本价) = 0
            .ColWidth(库存列表.现成本价) = 0
            .ColWidth(库存列表.差价差) = 0
            .ColWidth(库存列表.原价) = 900
            .ColWidth(库存列表.现价) = 900
        ElseIf mint调价 = 1 Then
            .ColWidth(库存列表.加成率) = 900
            .ColWidth(库存列表.原成本价) = 900
            .ColWidth(库存列表.现成本价) = 900
            .ColWidth(库存列表.差价差) = 900
            .ColWidth(库存列表.原价) = 0
            .ColWidth(库存列表.现价) = 0
        Else
            .ColWidth(库存列表.加成率) = 900
            .ColWidth(库存列表.原成本价) = 900
            .ColWidth(库存列表.现成本价) = 900
            .ColWidth(库存列表.差价差) = 900
            .ColWidth(库存列表.原价) = 900
            .ColWidth(库存列表.现价) = 900
        End If
        
        .ColAlignment(库存列表.库房) = 1
        .ColAlignment(库存列表.供应商) = 1
        .ColAlignment(库存列表.药品) = 1
        .ColAlignment(库存列表.规格) = 1
        .ColAlignment(库存列表.单位) = 4
        .ColAlignment(库存列表.批号) = 1
        .ColAlignment(库存列表.效期) = 1
        .ColAlignment(库存列表.产地) = 1
        .ColAlignment(库存列表.数量) = 7
        .ColAlignment(库存列表.原价) = 7
        .ColAlignment(库存列表.现价) = 7
        .ColAlignment(库存列表.调整金额) = 7
        .ColAlignment(库存列表.加成率) = 7
        .ColAlignment(库存列表.原成本价) = 7
        .ColAlignment(库存列表.现成本价) = 7
        .ColAlignment(库存列表.差价差) = 7
        .ColAlignment(库存列表.批次) = 7
        .ColAlignment(库存列表.变价) = 7
        .ColAlignment(库存列表.药品id) = 7
        
        .PrimaryCol = 库存列表.库房
        .LocateCol = 库存列表.库房
    End With
    
    
    With BillPay
        .Rows = 2
        .Cols = 应付款列.列数
        .MsfObj.FixedCols = 0
        
        .TextMatrix(0, 应付款列.药品id) = "药品id"
        .TextMatrix(0, 应付款列.品名) = "品名"
        .TextMatrix(0, 应付款列.发票号) = "发票号"
        .TextMatrix(0, 应付款列.发票日期) = "发票日期"
        .TextMatrix(0, 应付款列.发票金额) = "发票金额"
        
        .ColWidth(应付款列.药品id) = 0
        .ColWidth(应付款列.品名) = 3000
        .ColWidth(应付款列.发票号) = 1000
        .ColWidth(应付款列.发票日期) = 2000
        .ColWidth(应付款列.发票金额) = 1000
        
        .ColData(应付款列.药品id) = 5
        .ColData(应付款列.品名) = 5
        .ColData(应付款列.发票号) = 4
        .ColData(应付款列.发票日期) = 2
        .ColData(应付款列.发票金额) = 4

        .ColAlignment(应付款列.药品id) = 1
        .ColAlignment(应付款列.品名) = 1
        .ColAlignment(应付款列.发票号) = 1
        .ColAlignment(应付款列.发票日期) = 4
        .ColAlignment(应付款列.发票金额) = 7
        
'        .PrimaryCol = 应付款列.品名
'        .LocateCol = 应付款列.品名
    End With
End Sub

    
Private Sub BillPay_EnterCell(Row As Long, Col As Long)
    With BillPay
        Select Case Col
            Case 应付款列.发票号
                .TxtCheck = False
                .MaxLength = 20
            Case 应付款列.发票金额
                .TxtCheck = True
                .MaxLength = 14
                .TextMask = "-.1234567890"
            Case 应付款列.发票日期
                .TxtCheck = True
                .TextMask = "1234567890-"
                .Value = Sys.Currentdate
                .MaxLength = 10
        End Select
   End With
End Sub


Private Sub BillPay_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    
    If KeyCode <> 13 Then Exit Sub
    
    With BillPay
        .Text = UCase(Trim(.Text))
        strKey = UCase(Trim(.Text))
        Select Case .Col
            Case 应付款列.发票金额
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "对不起，发票金额必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Abs(Val(strKey)) < 0.001 Then
                        MsgBox "对不起，发票金额必须大于0.001,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(strKey) >= 10 ^ 14 - 1 Then
                        MsgBox "发票金额必须小于" & (10 ^ 14 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                End If
                
                If strKey <> "" Then
                    strKey = FormatEx(strKey, 2)
                    .Text = strKey
                ElseIf .TxtVisible = True Then
                    .Text = " "
                ElseIf .TxtVisible = False Then
                    If .TextMatrix(.Row, .Col) = "" Then
                        .Text = " "
                    Else
                        .Text = .TextMatrix(.Row, .Col)
                    End If
                    
                End If
            Case 应付款列.发票日期
                If strKey <> "" Then
                    If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
                        strKey = TranNumToDate(strKey)
                        
                        If strKey = "" Then
                            MsgBox "对不起，效期必须为日期型！", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            Exit Sub
                        End If
                        .Text = strKey
                        Exit Sub
                    End If
                    If Not IsDate(strKey) Then
                        MsgBox "对不起，发票日期必须为日期型如(2000-10-10) 或 （20001010）,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                End If
        End Select
    End With
End Sub

'转换数值为日期
Public Function TranNumToDate(ByVal strNum As String) As String
    Dim strYear As String
    Dim strMonth As String
    Dim strDay As String
    Dim StrDate As String
    
    TranNumToDate = ""
    strYear = Mid(strNum, 1, 4)
    strMonth = Mid(strNum, 5, 2)
    strDay = Mid(strNum, 7, 2)
        
    If strYear < 1000 Or strYear > 5000 Then Exit Function
    If strMonth = "" Then strMonth = "01"
    If strDay = "" Then strDay = "01"
    
    If strMonth > 12 Or strMonth < 1 Then Exit Function
    StrDate = strYear & "-" & strMonth & "-" & strDay
        
    If Not IsDate(StrDate) Then Exit Function
    
    StrDate = Format(StrDate, "yyyy-mm-dd")
    TranNumToDate = StrDate
End Function

Private Sub BillPrice_AfterAddRow(Row As Long)
    Call SetColor
End Sub

Private Sub BillPrice_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    Call GetDrugStore(Row, 0, Val(BillPrice.TextMatrix(Row, 售价列表.药品id)))
End Sub

Private Sub BillPrice_CommandClick()
    Dim strSqlType As String
    Dim dbl包装 As Double
    
    On Error GoTo errHandle
    If intDrugType = 1 Then
        If InStr(1, mstrPrivs, "管理西成药") > 0 And InStr(1, mstrPrivs, "管理中成药") > 0 Then
            strSqlType = "In('5','6')"
        ElseIf InStr(1, mstrPrivs, "管理西成药") > 0 Then
            strSqlType = "='5'"
        ElseIf InStr(1, mstrPrivs, "管理中成药") > 0 Then
            strSqlType = "='6'"
        End If
    Else
        strSqlType = "='7'"
    End If
    
    Select Case Me.BillPrice.Col
    Case 售价列表.品名
        gstrSql = "select I.ID,I.编码,I.名称,I.规格,I.产地,I.计算单位,P.药库单位,decode(I.是否变价,1,'时价','定价') 类型,Nvl(P.成本价,0) 成本价,P.指导批发价,P.指导零售价,P.剂量系数,P.药名ID " & _
                 " from 收费项目目录 I,药品规格 P" & _
                 " where I.类别 " & strSqlType & " And I.ID=P.药品ID" & _
                 "       and (I.撤档时间 Is Null Or I.撤档时间=To_Date('3000-01-01','yyyy-MM-dd'))"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption)
        
        With rsTemp
            If .BOF Or .EOF Then
                MsgBox "未建立药品！", vbInformation, gstrSysName: Exit Sub
            End If
            
            Me.lvwItem.Tag = 售价列表.品名

            With Me.lvwItem.ColumnHeaders
                .Clear
                .Add , "编码", "编码", 900
                .Add , "名称", "名称", 2000
                .Add , "规格", "规格", 1200
                .Add , "产地", "产地", 1200
                .Add , "单位", "单位", 500
                .Add , "类型", "类型", 600
                .Add , "成本价", "成本价", 600
                .Add , "采购限价", "采购限价", 0
                .Add , "指导售价", "指导售价", 0
                .Add , "剂量系数", "剂量系数", 0
                .Add , "药名ID", "药名ID", 0
            End With
            Me.lvwItem.Width = 6500
            
            Me.lvwItem.ListItems.Clear
            Do While Not .EOF
                dbl包装 = GetModulus(Val(!ID))
                
                Set objItem = Me.lvwItem.ListItems.Add(, "_" & !ID, !编码)
                objItem.SubItems(Me.lvwItem.ColumnHeaders("名称").Index - 1) = !名称
                objItem.SubItems(Me.lvwItem.ColumnHeaders("规格").Index - 1) = IIf(IsNull(!规格), "", !规格)
                objItem.SubItems(Me.lvwItem.ColumnHeaders("产地").Index - 1) = IIf(IsNull(!产地), "", !产地)
                If int药库单位 = 0 Then
                    objItem.SubItems(Me.lvwItem.ColumnHeaders("单位").Index - 1) = IIf(IsNull(!计算单位), "", !计算单位)
                Else
                    objItem.SubItems(Me.lvwItem.ColumnHeaders("单位").Index - 1) = IIf(IsNull(!药库单位), "", !药库单位)
                End If
                objItem.SubItems(Me.lvwItem.ColumnHeaders("类型").Index - 1) = IIf(IsNull(!类型), "", !类型)
                objItem.SubItems(Me.lvwItem.ColumnHeaders("成本价").Index - 1) = FormatEx(Val(IIf(IsNull(!成本价), "", !成本价)) * dbl包装, mintCostDigit)
                
                objItem.SubItems(Me.lvwItem.ColumnHeaders("采购限价").Index - 1) = FormatEx(Val(IIf(IsNull(!指导批发价), "", !指导批发价)) * dbl包装, mintCostDigit)
                objItem.SubItems(Me.lvwItem.ColumnHeaders("指导售价").Index - 1) = FormatEx(Val(IIf(IsNull(!指导零售价), "", !指导零售价)) * dbl包装, mintPriceDigit)
                objItem.SubItems(Me.lvwItem.ColumnHeaders("剂量系数").Index - 1) = !剂量系数
                objItem.SubItems(Me.lvwItem.ColumnHeaders("药名ID").Index - 1) = !药名ID
                .MoveNext
            Loop
            Me.lvwItem.ListItems(1).Selected = True
            If Me.lvwItem.ListItems.Count = 1 Then
                Call lvwItem_DblClick: Exit Sub
            End If
        End With
        With Me.lvwItem
            .Left = Me.BillPrice.Left
            .Top = Me.BillPrice.Top + Me.BillPrice.CellTop + Me.BillPrice.RowHeight(1)
            If Me.ScaleHeight - .Top < 3000 Then
                .Height = Me.ScaleHeight - .Top
            Else
                .Height = 3000
            End If
            .Visible = True
            .ZOrder 0
            .SetFocus
        End With
    Case 售价列表.收入名称
        
        gstrSql = "select id,编码,名称" & _
                " from 收入项目" & _
                " where (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','yyyy-MM-dd')) and 末级=1"
'            If .State = adStateOpen Then .Close
'            Call SQLTest(App.Title, Me.Caption, gstrSql)
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "BillPrice_CommandClick")
'            Call SQLTest
        With rsTemp
            If .BOF Or .EOF Then
                MsgBox "没有设置好收入项目", vbExclamation, gstrSysName: Exit Sub
            End If
            
            Me.lvwItem.Tag = 售价列表.收入名称
            With Me.lvwItem.ColumnHeaders
                .Clear
                .Add , "编码", "编码", 600
                .Add , "名称", "名称", 1000
            End With
            Me.lvwItem.Width = 1800
            Me.lvwItem.ListItems.Clear
            Do While Not .EOF
                Set objItem = Me.lvwItem.ListItems.Add(, "_" & !ID, !编码)
                objItem.SubItems(Me.lvwItem.ColumnHeaders("名称").Index - 1) = !名称
                If Me.lvwItem.SelectedItem Is Nothing Then
                    objItem.Selected = True
                End If
                .MoveNext
            Loop
            Me.lvwItem.ListItems(1).Selected = True
            If Me.lvwItem.ListItems.Count = 1 Then
                Call lvwItem_DblClick: Exit Sub
            End If
        End With
        
        With Me.lvwItem
            .Left = BillPrice.Left + BillPrice.MsfObj.CellLeft
            .Top = Me.BillPrice.Top + Me.BillPrice.CellTop + Me.BillPrice.RowHeight(1)
            If Me.ScaleHeight - .Top < 2000 Then
                .Height = Me.ScaleHeight - .Top
            Else
                .Height = 2000
            End If
            .Visible = True
            .ZOrder 0
            .SetFocus
        End With
    End Select
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function CheckDrugRepeat(ByVal lng药品ID As Long) As Boolean
    Dim n As Integer
    
    With BillPrice
        For n = 1 To .Rows - 1
            If .TextMatrix(n, 售价列表.药品id) <> "" Then
                If Val(.TextMatrix(n, 售价列表.药品id)) = lng药品ID Then
                    MsgBox "对不起，已有该药品，不能重复输入！", vbOKOnly, gstrSysName
                    Exit Function
                End If
            End If
        Next
    End With
    CheckDrugRepeat = True
End Function

Private Sub BillPrice_EditKeyPress(KeyAscii As Integer)
    With BillPrice
        If .Col = 售价列表.现成本价 Then
            mdbl成本价 = Val(.TextMatrix(.Row, .Col))
        End If
    End With
End Sub

Private Sub BillPrice_EnterCell(Row As Long, Col As Long)
    Dim n As Integer
    
    Select Case Col
    Case 售价列表.现采购限价, 售价列表.现指导售价
'        BillPrice.TxtCheck = True
        BillPrice.MaxLength = 11
        BillPrice.TextMask = ".1234567890"
    Case 售价列表.品名
        Me.lblHelp.Caption = "提示：输入药品编码、简码选择调价药品"
    Case 售价列表.现价
        Me.lblHelp.Caption = "提示：F3进入药价辅助计算，根据成本价计算产生新的售价"
        
        If mint调价 <> 3 Then
            If mint调价 = 1 Or (Me.BillPrice.TextMatrix(Row, 售价列表.类型) = "时价" And mbln时价药品调价) Then
                Me.BillPrice.ColData(售价列表.现价) = 0
            Else
                Me.BillPrice.ColData(售价列表.现价) = 4
                BillPrice.MaxLength = 11
                BillPrice.TextMask = ".1234567890"
            End If
        End If
    Case 售价列表.现成本价
        Me.BillPrice.ColData(售价列表.现成本价) = 0
        If mint调价 = 1 Or mint调价 = 2 Then
            Me.BillPrice.ColData(售价列表.现成本价) = 4
            BillPrice.MaxLength = 11
            BillPrice.TextMask = ".1234567890"
        End If
    Case 售价列表.收入名称
        Me.BillPrice.TextMatrix(Row, 售价列表.现价) = FormatEx(Me.BillPrice.TextMatrix(Row, 售价列表.现价), mintPriceDigit)
        Me.lblHelp.Caption = "提示：正确设置药品的收入项目，以便有效完成财务科目核算"

    Case Else
        Me.lblHelp.Caption = ""
    End Select
    
    If BillStore.Rows > 1 Then
        If Trim(BillStore.TextMatrix(1, 0)) <> "" Then
            For n = 1 To BillStore.Rows - 1
                If Val(Me.BillPrice.TextMatrix(Row, 售价列表.药品id)) = Val(BillStore.TextMatrix(n, 库存列表.药品id)) Then
                    BillStore.MsfObj.TopRow = n
                    Exit For
                End If
            Next
        End If
    End If
End Sub

Private Sub BillPrice_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strInput As String
    Dim strSqlType As String
    Dim lng药品ID As Long
    Dim dbl包装 As Double
    Dim dblSalePrice As Double
    
    If KeyCode = 13 And Not BillPrice.Active Then
        Cancel = True: Call OS.PressKey(vbKeyTab): Exit Sub
    End If
    
    If KeyCode <> 13 Then Exit Sub
    
    On Error GoTo errHandle
    If intDrugType = 1 Then
        If InStr(1, mstrPrivs, "管理西成药") > 0 And InStr(1, mstrPrivs, "管理中成药") > 0 Then
            strSqlType = "In('5','6')"
        ElseIf InStr(1, mstrPrivs, "管理西成药") > 0 Then
            strSqlType = "='5'"
        ElseIf InStr(1, mstrPrivs, "管理中成药") > 0 Then
            strSqlType = "='6'"
        End If
    Else
        strSqlType = "='7'"
    End If
    
    Select Case Me.BillPrice.Col
    Case 售价列表.品名
        If Trim(Me.BillPrice.Text) = "" Then Exit Sub
        If Me.BillPrice.TextMatrix(BillPrice.Row, 售价列表.品名) = UCase(Trim(Me.BillPrice.Text)) Then Exit Sub
        strInput = UCase(Trim(Me.BillPrice.Text))
        
        gstrSql = "select distinct I.ID,I.编码,I.名称,I.规格,I.产地,I.计算单位,P.药库单位,decode(I.是否变价,1,'时价','定价') 类型,Nvl(P.成本价,0) 成本价,P.指导批发价,P.指导零售价,P.剂量系数,P.药名ID " & _
                 " from 收费项目目录 I,收费项目别名 N,药品规格 P" & _
                 " where I.ID=N.收费细目ID and I.类别 " & strSqlType & " And I.ID=P.药品ID " & _
                 "       and (I.编码 like [1] " & _
                 "            or N.简码 Like [2] " & _
                 "            or N.名称 Like [2])" & _
                 "       and (I.撤档时间 Is Null Or I.撤档时间=To_Date('3000-01-01','yyyy-MM-dd'))"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, strInput & "%", gstrMatch & strInput & "%")
        
        With rsTemp
            If .BOF Or .EOF Then
                MsgBox "未找到相关药品，请重新输入！", vbInformation, gstrSysName
'                Cancel = True
                Exit Sub
            End If
            
            Me.lvwItem.Tag = 售价列表.品名
            Me.lvwItem.Tag = 售价列表.品名
            With Me.lvwItem.ColumnHeaders
                .Clear
                .Add , "编码", "编码", 900
                .Add , "名称", "名称", 2000
                .Add , "规格", "规格", 1200
                .Add , "产地", "产地", 1200
                .Add , "单位", "单位", 500
                .Add , "类型", "类型", 600
                .Add , "成本价", "成本价", 800
                .Add , "采购限价", "采购限价", 0
                .Add , "指导售价", "指导售价", 0
                .Add , "剂量系数", "剂量系数", 0
                .Add , "药名ID", "药名ID", 0
            End With
            Me.lvwItem.Width = 6500
            
            Me.lvwItem.ListItems.Clear
            Do While Not .EOF
                dbl包装 = GetModulus(Val(!ID))
                
                Set objItem = Me.lvwItem.ListItems.Add(, "_" & !ID, !编码)
                objItem.SubItems(Me.lvwItem.ColumnHeaders("名称").Index - 1) = !名称
                objItem.SubItems(Me.lvwItem.ColumnHeaders("规格").Index - 1) = IIf(IsNull(!规格), "", !规格)
                objItem.SubItems(Me.lvwItem.ColumnHeaders("产地").Index - 1) = IIf(IsNull(!产地), "", !产地)
                If int药库单位 = 0 Then
                    objItem.SubItems(Me.lvwItem.ColumnHeaders("单位").Index - 1) = IIf(IsNull(!计算单位), "", !计算单位)
                Else
                    objItem.SubItems(Me.lvwItem.ColumnHeaders("单位").Index - 1) = IIf(IsNull(!药库单位), "", !药库单位)
                End If
                objItem.SubItems(Me.lvwItem.ColumnHeaders("类型").Index - 1) = IIf(IsNull(!类型), "", !类型)
                objItem.SubItems(Me.lvwItem.ColumnHeaders("成本价").Index - 1) = FormatEx(Val(IIf(IsNull(!成本价), "", !成本价)) * dbl包装, mintCostDigit)
                objItem.SubItems(Me.lvwItem.ColumnHeaders("采购限价").Index - 1) = FormatEx(Val(IIf(IsNull(!指导批发价), "", !指导批发价)) * dbl包装, mintCostDigit)
                objItem.SubItems(Me.lvwItem.ColumnHeaders("指导售价").Index - 1) = FormatEx(Val(IIf(IsNull(!指导零售价), "", !指导零售价)) * dbl包装, mintPriceDigit)
                objItem.SubItems(Me.lvwItem.ColumnHeaders("剂量系数").Index - 1) = !剂量系数
                objItem.SubItems(Me.lvwItem.ColumnHeaders("药名ID").Index - 1) = !药名ID
                
                .MoveNext
            Loop
            Me.lvwItem.ListItems(1).Selected = True
            If Me.lvwItem.ListItems.Count = 1 Then
                Call lvwItem_DblClick: Cancel = True: Exit Sub
            End If
        End With
        With Me.lvwItem
            .Left = Me.BillPrice.Left
            .Top = Me.BillPrice.Top + Me.BillPrice.CellTop + Me.BillPrice.RowHeight(1)
            If Me.ScaleHeight - .Top < 3000 Then
                .Height = Me.ScaleHeight - .Top
            Else
                .Height = 3000
            End If
            .Visible = True
            .ZOrder 0
            .SetFocus
        End With
        Cancel = True
    
    Case 售价列表.收入名称
        If Trim(Me.BillPrice.Text) = "" Then Exit Sub
        strInput = UCase(Me.BillPrice.Text)
        
        gstrSql = "select id,编码,名称" & _
                " from 收入项目" & _
                " where (编码 like [1] or 简码 like [2] or 名称 like [2])" & _
                "       and (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','yyyy-MM-dd')) and 末级=1"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, strInput & "%", gstrMatch & strInput & "%")
        
        With rsTemp
            If .BOF Or .EOF Then
                MsgBox "该项目不存在", vbExclamation, gstrSysName: Cancel = True: Exit Sub
            End If
            
            Me.lvwItem.Tag = 售价列表.收入名称
            With Me.lvwItem.ColumnHeaders
                .Clear
                .Add , "编码", "编码", 600
                .Add , "名称", "名称", 1000
            End With
            Me.lvwItem.Width = 1800
            Me.lvwItem.ListItems.Clear
            Do While Not .EOF
                Set objItem = Me.lvwItem.ListItems.Add(, "_" & !ID, !编码)
                objItem.SubItems(Me.lvwItem.ColumnHeaders("名称").Index - 1) = !名称
                If Me.lvwItem.SelectedItem Is Nothing Then
                    objItem.Selected = True
                End If
                .MoveNext
            Loop
            Me.lvwItem.ListItems(1).Selected = True
            If Me.lvwItem.ListItems.Count = 1 Then
                Call lvwItem_DblClick: Cancel = True: Exit Sub
            End If
        End With
        
        With Me.lvwItem
            .Left = BillPrice.Left + BillPrice.MsfObj.CellLeft
            .Top = Me.BillPrice.Top + Me.BillPrice.CellTop + Me.BillPrice.RowHeight(1)
            If Me.ScaleHeight - .Top < 2000 Then
                .Height = Me.ScaleHeight - .Top
            Else
                .Height = 2000
            End If
            .Visible = True
            .ZOrder 0
            .SetFocus
        End With
        Cancel = True
    
    Case 售价列表.现价
        With BillPrice
            If .Text = "" Then Exit Sub
            
            lng药品ID = Val(BillPrice.TextMatrix(BillPrice.Row, 售价列表.药品id))
            If lng药品ID = 0 Then Exit Sub
            
            '现价大于指导售价时，提示是否继续
            If mbln限价提示 = True Then
                If BillPrice.TextMatrix(BillPrice.Row, 售价列表.类型) = "定价" And Val(.Text) > Val(BillPrice.TextMatrix(BillPrice.Row, 售价列表.现指导售价)) Then
                    MsgBox "现价高于指导零售价" & Val(BillPrice.TextMatrix(BillPrice.Row, 售价列表.现指导售价)) & "，指导价格将和售价一致！", vbInformation, gstrSysName
                End If
            End If
            
            If Val(.Text) < 0 Then
                MsgBox "售价不能为负数！", vbExclamation, gstrSysName
                Cancel = True
                .TxtSetFocus
            End If
            
            .TextMatrix(BillPrice.Row, 售价列表.现价) = .Text
            If Val(.Text) > Val(BillPrice.TextMatrix(BillPrice.Row, 售价列表.现指导售价)) Then
                BillPrice.TextMatrix(BillPrice.Row, 售价列表.现指导售价) = .Text
            End If
            
            '调整库存数据
            Call ChangeDrugStore(BillPrice.Row, lng药品ID, Val(.Text))
            
            '中草药按规格批量调整现价
            Call BatchAdjustPriceByItem(BillPrice.Row, Val(.Text))
            
            blnModify = True
        End With
        
    Case 售价列表.现采购限价
        With BillPrice
            If Val(BillPrice.TextMatrix(BillPrice.Row, 售价列表.药品id)) = 0 Then Exit Sub
            
            If .Text = "" Then Exit Sub
            
            If Val(.Text) < 0 Then
                MsgBox "价格不能为负数！", vbExclamation, gstrSysName
                Cancel = True
                .TxtSetFocus
            End If
            
            If mbln限价提示 = True Then
                If Val(.Text) < Val(BillPrice.TextMatrix(BillPrice.Row, 售价列表.现成本价)) Then
                    If MsgBox("现指导采购限价低于现价" & Val(BillPrice.TextMatrix(BillPrice.Row, 售价列表.现成本价)) & "。" & vbCrLf & "继续吗？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                        Cancel = True
                        .TxtSetFocus
                    End If
                End If
            End If
            
            .TextMatrix(BillPrice.Row, 售价列表.现采购限价) = .Text
            
            blnModify = True
        End With
    Case 售价列表.现指导售价
        With BillPrice
            If Val(BillPrice.TextMatrix(BillPrice.Row, 售价列表.药品id)) = 0 Then Exit Sub
            
            If .Text = "" Then Exit Sub
            
            '现指导售价小于指导售价时，提示是否继续
            If mbln限价提示 = True Then
                If BillPrice.TextMatrix(BillPrice.Row, 售价列表.类型) = "定价" And Val(.Text) < Val(BillPrice.TextMatrix(BillPrice.Row, 售价列表.现价)) Then
                    If MsgBox("现指导零售价低于现价" & Val(BillPrice.TextMatrix(BillPrice.Row, 售价列表.现价)) & "。" & vbCrLf & "继续吗？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                        Cancel = True
                        .TxtSetFocus
                    End If
                End If
            End If
            
            If Val(.Text) < 0 Then
                MsgBox "价格不能为负数！", vbExclamation, gstrSysName
                Cancel = True
                .TxtSetFocus
            End If
            
            .TextMatrix(BillPrice.Row, 售价列表.现指导售价) = .Text
            
            blnModify = True
        End With
    Case 售价列表.现成本价
        With BillPrice
            If .Text = "" Then Exit Sub
            
            If Val(.Text) < 0 Then
                MsgBox "价格不能为负数！", vbExclamation, gstrSysName
                Cancel = True
                .TxtSetFocus
            End If
            
            If mbln限价提示 = True Then
                If Val(.Text) > Val(BillPrice.TextMatrix(BillPrice.Row, 售价列表.现采购限价)) Then
                    MsgBox "现成本价高于指导采购限价" & Val(BillPrice.TextMatrix(BillPrice.Row, 售价列表.现采购限价)) & "，采购限价将和采购价一致！", vbInformation, gstrSysName
                End If
            End If
            
            .TextMatrix(BillPrice.Row, 售价列表.现成本价) = .Text
            If Val(.Text) > Val(BillPrice.TextMatrix(BillPrice.Row, 售价列表.现采购限价)) Then
                BillPrice.TextMatrix(BillPrice.Row, 售价列表.现采购限价) = .Text
            End If
            
            If cbo售价计算方式 = "售价按分段加成计算" And .TextMatrix(.Row, 售价列表.类型) = "时价" And mint调价 = 2 Then
                Call get分段加成售价(Val(.TextMatrix(BillPrice.Row, 售价列表.现成本价)), dblSalePrice)
                If dblSalePrice = 0 Then
                    .Text = mdbl成本价
                    .TextMatrix(BillPrice.Row, 售价列表.现成本价) = .Text
                    .TxtSetFocus
                    Exit Sub
                End If
                dblSalePrice = dblSalePrice + (Val(.TextMatrix(.Row, 售价列表.原指导售价)) - dblSalePrice) * (1 - Val(.TextMatrix(.Row, 售价列表.差价让利比)) / 100)
'                If dblSalePrice > Val(.TextMatrix(.Row, 售价列表.原指导售价)) Then dblSalePrice = Val(.TextMatrix(.Row, 售价列表.原指导售价))
                .TextMatrix(.Row, 售价列表.现价) = FormatEx(dblSalePrice, mintPriceDigit)
            ElseIf cbo售价计算方式 = "售价按固定比例计算" And .TextMatrix(.Row, 售价列表.类型) = "时价" And mint调价 = 2 Then
                dblSalePrice = Val(.Text) * (1 + Val(.TextMatrix(.Row, 售价列表.加成率)))
                If dblSalePrice > Val(.TextMatrix(.Row, 售价列表.原指导售价)) Then dblSalePrice = Val(.TextMatrix(.Row, 售价列表.原指导售价))
                .TextMatrix(.Row, 售价列表.现价) = FormatEx(dblSalePrice, mintPriceDigit)
            End If
            
            CaculateCost Val(BillPrice.TextMatrix(BillPrice.Row, 售价列表.药品id)), Val(.Text)
            
            
            '中草药按规格批量调整现价
            Call BatchAdjustCostByItem(BillPrice.Row, Val(.Text))
            
            blnModify = True
        End With
    End Select
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub BillPrice_LostFocus()
    Me.lblHelp.Caption = ""
End Sub

Private Sub BillStore_EnterCell(Row As Long, Col As Long)
    Dim i As Integer
    
    With BillStore
        If Row = 0 Then Exit Sub
        If .TextMatrix(Row, 0) = "" Or .TextMatrix(Row, 库存列表.药品) = "" Then Exit Sub
        If mint调价 = 3 Then
            .ColData(库存列表.现价) = 0
            .ColData(库存列表.现成本价) = 0
            .ColData(库存列表.加成率) = 0
            Exit Sub
        End If
        Select Case Col
            Case 库存列表.现价
                .TxtCheck = True
                .MaxLength = 11
                .TextMask = ".1234567890"
                
                If Val(.TextMatrix(Row, 库存列表.变价)) = 1 And mbln时价药品调价 And mint调价 <> 1 Then
                    .ColData(库存列表.现价) = 4
                Else
                    .ColData(库存列表.现价) = 0
                End If
                
                If BillPrice.Rows = 1 Then Exit Sub
                If BillPrice.TextMatrix(1, 售价列表.药品id) = "" Then Exit Sub
                
                For i = 1 To BillPrice.Rows - 1
                    If Val(BillPrice.TextMatrix(i, 售价列表.药品id)) = Val(.TextMatrix(Row, 库存列表.药品id)) Then
                        BillPrice.Row = i
                        Exit For
                    End If
                Next
                
            Case 库存列表.现成本价
                .TxtCheck = True
                .MaxLength = 11
                .TextMask = ".1234567890"
            Case 库存列表.加成率
                .TxtCheck = True
                .MaxLength = 8
                .TextMask = ".1234567890"
        End Select
    End With
End Sub


Private Sub BillStore_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strInput As String
    Dim n As Integer
    Dim intRow As Integer
    Dim dbl发票金额 As Double
    Dim dbl数量 As Double
    Dim dbl金额 As Double
    Dim dbl现成本价 As Double
    
    If KeyCode <> 13 Then Exit Sub
    
    With BillStore
        If .Text = "" Then Exit Sub
        intRow = .Row
        Select Case .Col
            Case 库存列表.现价
                If Not IsNumeric(.Text) Then
                    MsgBox "请输入新的售价。", vbInformation, gstrSysName
                    Exit Sub
                Else
                    .Text = FormatEx(.Text, mintPriceDigit)
                End If
                .TextMatrix(intRow, 库存列表.调整金额) = Format(Val(.TextMatrix(intRow, 库存列表.数量)) * (Val(.Text) - Val(.TextMatrix(intRow, 库存列表.原价))), mstrMoneyFormat)
                .TextMatrix(intRow, 库存列表.现价) = FormatEx(Val(.Text), mintPriceDigit)
                .TextMatrix(intRow, 库存列表.现成本价) = FormatEx(Val(.TextMatrix(intRow, 库存列表.现价)) / (1 + Val(.TextMatrix(intRow, 库存列表.加成率)) / 100), mintCostDigit)
                .TextMatrix(intRow, 库存列表.差价差) = Format((Val(.TextMatrix(intRow, 库存列表.现成本价)) - Val(.TextMatrix(intRow, 库存列表.原成本价))) * Val(.TextMatrix(intRow, 库存列表.数量)), mstrMoneyFormat)
                
                For n = 1 To .Rows - 1
                    If BillPrice.TextMatrix(BillPrice.Row, 售价列表.药品id) = .TextMatrix(n, 库存列表.药品id) Then
                        If Val(.TextMatrix(intRow, 库存列表.批次)) <> 0 And Val(.TextMatrix(intRow, 库存列表.批次)) = Val(.TextMatrix(n, 库存列表.批次)) Then
                            .TextMatrix(n, 库存列表.现价) = .TextMatrix(intRow, 库存列表.现价)
                            .TextMatrix(n, 库存列表.调整金额) = Format(Val(.TextMatrix(n, 库存列表.数量)) * (Val(.Text) - Val(.TextMatrix(n, 库存列表.原价))), mstrMoneyFormat)
                            .TextMatrix(n, 库存列表.现成本价) = FormatEx(Val(.TextMatrix(n, 库存列表.现价)) / (1 + Val(.TextMatrix(n, 库存列表.加成率)) / 100), mintCostDigit)
                            .TextMatrix(n, 库存列表.差价差) = Format((Val(.TextMatrix(n, 库存列表.现成本价)) - Val(.TextMatrix(n, 库存列表.原成本价))) * Val(.TextMatrix(n, 库存列表.数量)), mstrMoneyFormat)
                        End If
                        dbl数量 = dbl数量 + .TextMatrix(n, 库存列表.数量)
                        dbl金额 = dbl金额 + .TextMatrix(n, 库存列表.数量) * Val(.TextMatrix(n, 库存列表.现价))
                    End If
                Next
                
                BillPrice.TextMatrix(BillPrice.Row, 售价列表.现价) = FormatEx(dbl金额 / dbl数量, mintPriceDigit)
                
                If mint调价 > 0 Then
                    For n = 1 To .Rows - 1
                        If .TextMatrix(n, 库存列表.药品id) <> "" Then
                            If Val(.TextMatrix(n, 库存列表.药品id)) = Val(.TextMatrix(intRow, 库存列表.药品id)) Then
                                dbl发票金额 = dbl发票金额 + (Val(.TextMatrix(n, 库存列表.现成本价)) - Val(.TextMatrix(n, 库存列表.原成本价))) * Val(.TextMatrix(n, 库存列表.数量))
                            End If
                        End If
                    Next
    
                    If chk自动计算应付款变动.Value = 1 Then
                        For n = 1 To BillPay.Rows - 1
                            If BillPay.TextMatrix(1, 0) <> "" Then
                                If Val(BillPay.TextMatrix(n, 应付款列.药品id)) = Val(BillStore.TextMatrix(intRow, 库存列表.药品id)) Then
                                    BillPay.TextMatrix(n, 应付款列.发票金额) = FormatEx(dbl发票金额, 2)
                                End If
                            End If
                        Next
                    End If
                End If
            Case 库存列表.加成率
                If Val(.Text) < 0 Then Exit Sub
                
                .TextMatrix(intRow, 库存列表.加成率) = FormatEx(Val(.Text), 5)
                .TextMatrix(intRow, 库存列表.现成本价) = FormatEx(Val(.TextMatrix(intRow, 库存列表.现价)) / (1 + Val(.TextMatrix(intRow, 库存列表.加成率)) / 100), mintCostDigit)
                .TextMatrix(intRow, 库存列表.差价差) = Format((Val(.TextMatrix(intRow, 库存列表.现成本价)) - .TextMatrix(intRow, 库存列表.原成本价)) * Val(.TextMatrix(intRow, 库存列表.数量)), mstrMoneyFormat)
                dbl发票金额 = (Val(.TextMatrix(intRow, 库存列表.现成本价)) - .TextMatrix(intRow, 库存列表.原成本价)) * Val(.TextMatrix(intRow, 库存列表.数量))
                
                For n = 1 To .Rows - 1
                    If .TextMatrix(n, 库存列表.药品id) <> "" Then
                        If Val(.TextMatrix(n, 库存列表.药品id)) = Val(.TextMatrix(intRow, 库存列表.药品id)) And n <> intRow Then
                            If chk按批次.Value = 0 Or (Val(.TextMatrix(intRow, 库存列表.批次)) <> 0 And Val(.TextMatrix(intRow, 库存列表.批次)) = Val(.TextMatrix(n, 库存列表.批次))) Then
                                .TextMatrix(n, 库存列表.加成率) = FormatEx(.TextMatrix(intRow, 库存列表.加成率), 5)
                                .TextMatrix(n, 库存列表.现成本价) = .TextMatrix(intRow, 库存列表.现成本价)
                                .TextMatrix(n, 库存列表.差价差) = Format((Val(.TextMatrix(n, 库存列表.现成本价)) - .TextMatrix(n, 库存列表.原成本价)) * Val(.TextMatrix(n, 库存列表.数量)), mstrMoneyFormat)
                            End If
                        End If
                        dbl发票金额 = dbl发票金额 + (Val(.TextMatrix(n, 库存列表.现成本价)) - .TextMatrix(n, 库存列表.原成本价)) * Val(.TextMatrix(n, 库存列表.数量))
                    End If
                Next

                If chk自动计算应付款变动.Value = 1 Then
                    For n = 1 To BillPay.Rows - 1
                        If BillPay.TextMatrix(1, 0) <> "" Then
                            If Val(BillPay.TextMatrix(n, 应付款列.药品id)) = Val(BillStore.TextMatrix(intRow, 库存列表.药品id)) Then
                                BillPay.TextMatrix(n, 应付款列.发票金额) = FormatEx(dbl发票金额, 2)
                            End If
                        End If
                    Next
                End If
            Case 库存列表.现成本价
                If Val(.Text) > Val(.TextMatrix(.Row, 库存列表.现价)) Then
                    MsgBox "注意，新成本价大于了新售价！", vbExclamation, gstrSysName
                End If
                
                If Val(.Text) < 0 Then
                    MsgBox "成本价不能为负数！", vbExclamation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                End If
                
                .TextMatrix(intRow, 库存列表.现成本价) = FormatEx(Val(.Text), mintCostDigit)
                If Val(.Text) <> 0 Then
                    .TextMatrix(intRow, 库存列表.加成率) = FormatEx((Val(.TextMatrix(intRow, 库存列表.现价)) / Val(.Text) - 1) * 100, 5)
                End If
                .TextMatrix(intRow, 库存列表.差价差) = Format((Val(.Text) - .TextMatrix(intRow, 库存列表.原成本价)) * Val(.TextMatrix(intRow, 库存列表.数量)), mstrMoneyFormat)
                dbl发票金额 = (Val(.Text) - .TextMatrix(intRow, 库存列表.原成本价)) * Val(.TextMatrix(intRow, 库存列表.数量))
                
                For n = 1 To .Rows - 1
                    If .TextMatrix(n, 库存列表.药品id) <> "" Then
                        If Val(.TextMatrix(n, 库存列表.药品id)) = Val(.TextMatrix(intRow, 库存列表.药品id)) And n <> intRow Then
                            If chk按批次.Value = 0 Or (Val(.TextMatrix(intRow, 库存列表.批次)) <> 0 And Val(.TextMatrix(intRow, 库存列表.批次)) = Val(.TextMatrix(n, 库存列表.批次))) Then
                                dbl现成本价 = Val(.Text)
                                .TextMatrix(n, 库存列表.现成本价) = FormatEx(dbl现成本价, mintCostDigit)
                                If dbl现成本价 <> 0 Then
                                    .TextMatrix(n, 库存列表.加成率) = FormatEx((Val(.TextMatrix(n, 库存列表.现价)) / dbl现成本价 - 1) * 100, 5)
                                End If
                                .TextMatrix(n, 库存列表.差价差) = Format((dbl现成本价 - .TextMatrix(n, 库存列表.原成本价)) * Val(.TextMatrix(n, 库存列表.数量)), mstrMoneyFormat)
                            Else
                                dbl现成本价 = Val(.TextMatrix(n, 库存列表.现成本价))
                            End If
                            dbl发票金额 = dbl发票金额 + (dbl现成本价 - .TextMatrix(n, 库存列表.原成本价)) * Val(.TextMatrix(n, 库存列表.数量))
                        End If
                    End If
                Next

                If chk自动计算应付款变动.Value = 1 Then
                    For n = 1 To BillPay.Rows - 1
                        If BillPay.TextMatrix(1, 0) <> "" Then
                            If Val(BillPay.TextMatrix(n, 应付款列.药品id)) = Val(BillStore.TextMatrix(intRow, 库存列表.药品id)) Then
                                BillPay.TextMatrix(n, 应付款列.发票金额) = Format(dbl发票金额, mstrMoneyFormat)
                            End If
                        End If
                    Next
                End If
                
                If chk按批次.Value = 0 Then
                    For n = 1 To BillPrice.Rows - 1
                        If Val(.TextMatrix(intRow, 库存列表.药品id)) = Val(BillPrice.TextMatrix(n, 售价列表.药品id)) Then
                            BillPrice.TextMatrix(n, 售价列表.现成本价) = .TextMatrix(intRow, 库存列表.现成本价)
                            Exit For
                        End If
                    Next
                Else
                    CaluateAverCost Val(.TextMatrix(intRow, 库存列表.药品id))
                End If
        End Select
    End With
End Sub

Private Sub cbo售价计算方式_Click()
    Set mrs分段加成 = Nothing
    If cbo售价计算方式.Text = "售价按分段加成计算" Then
        gstrSql = "select 序号, 最低价, 最高价, 加成率, 差价额, 说明 from 药品加成方案 order by 序号"
        Set mrs分段加成 = zldatabase.OpenSQLRecord(gstrSql, "药品加成方案")
    End If
End Sub

Private Sub get分段加成售价(ByVal dbl采购价 As Double, ByRef dbl售价 As Double)
'功能：通过成本价按分段加成方式计算售价
'参数：成本价,售价
    Dim dbl差价额 As Double
    
    mdbl分段加成率 = 0
    If mrs分段加成.EOF Then
        dbl售价 = 0!
        MsgBox "没有设置金额段为：" & dbl采购价 & "  的加成率，请在药品目录管理（分段加成率）中设置！"
        Exit Sub
    End If
    mrs分段加成.MoveFirst
    Do Until mrs分段加成.EOF
        If dbl采购价 > mrs分段加成!最低价 And dbl采购价 <= mrs分段加成!最高价 Then
            mdbl分段加成率 = mrs分段加成!加成率 / 100
            dbl差价额 = IIf(IsNull(mrs分段加成!差价额), 0, mrs分段加成!差价额)
            Exit Do
        End If
        mrs分段加成.MoveNext
    Loop
    If mdbl分段加成率 = 0 Then
        MsgBox "没有设置金额段为：" & dbl采购价 & "  的加成率，请在药品目录管理（分段加成率）中设置！"
        dbl售价 = 0
        Exit Sub
    Else
        If dbl采购价 <= 2000 Then
            dbl售价 = dbl采购价 * (1 + mdbl分段加成率) + dbl差价额
        Else
            dbl售价 = dbl采购价 + dbl差价额
        End If
    End If
End Sub

'Private Sub cbo执行时间_Click()
'    If cbo执行时间.Text = "立即生效" Then
'       cbo执行时间.ListIndex = IIf(Check存在未执行价格, 1, 0)
'    End If
'
'    If Me.cbo执行时间.Text = "立即生效" Then
'        Me.dtpRunDate.Enabled = False
'    Else
'        Me.dtpRunDate.Enabled = True
'    End If
'
'    On Error Resume Next
'    Me.BillPrice.SetFocus
'End Sub

Private Sub ChkSelect_Click()
    Dim lngRow As Long
    
    With vsfSpec
        For lngRow = 1 To .Rows - 1
            If Val(.TextMatrix(lngRow, .ColIndex("药品ID"))) > 0 Then
                .TextMatrix(lngRow, .ColIndex("选择")) = IIf(ChkSelect.Value = 1, 1, 0)
            End If
        Next
    End With
End Sub

Private Sub cmdAdd_Click()
    Call GetBatchData(False)
End Sub

Private Sub cmdCanc_Click()
    Dim strTemp As String
    Dim i As Integer
    Dim j As Integer
    
    With BillPrice
        For i = 1 To .Rows - 1
            For j = 1 To .Cols - 1
                strTemp = strTemp & .TextMatrix(i, j) & "|"
            Next
        Next
    End With
    strTemp = strTemp & "|" & txtSummary.Text & "|" & txtValuer.Text & "|" & opt时间(0).Value & "|" & opt时间(1).Value & "|" & dtpRunDate.Value & "|" & Chk定价.Value & "|" & chk草药批量调价.Value & "|" & _
                    chk按批次 & "|" & chk自动计算应付款变动.Value & "|" & chk自动调成本价.Value
    
    If strTemp <> mstr所有记录 Then
        If MsgBox("有数据被修改了，是否退出？", vbYesNo, gstrSysName) = vbYes Then
            lngBillId = 0
            lngMediId = 0
            lngItemID = 0
            Unload Me
        Else
            Exit Sub
        End If
    Else
        lngBillId = 0
        lngMediId = 0
        lngItemID = 0
        Unload Me
    End If
End Sub

Private Sub CmdExit_Click()
    txtItem.Text = ""
    vsfSpec.Rows = 1
    vsfSpec.Rows = 2
    ChkSelect.Value = 0
    picItem.Visible = False
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub cmdItem_Click()
    picItem.Visible = True
    
    picItem.Left = fraCondition.Left + lblSummary.Left
    picItem.Top = fraCondition.Top + lblSummary.Top
    picItem.Width = fraCondition.Left + txtSummary.Left + txtSummary.Width
    picItem.Height = (Me.Height - fraCondition.Top) * 2 / 3
    
    txtItem.SetFocus
End Sub

Private Sub cmdOK_Click()
    Dim strID As String, LngCurID As Long
    Dim ArrayID
    Dim lngAdjId As Long
    Dim strOldId As String
    Dim strNewId As String
    
    Dim Array批次价格
    Dim str批次价格 As String
    Dim lngCurrBatch As Long
    Dim strTmp As String
    Dim str时价分批 As String
    Dim n As Integer
    Dim i As Integer
    
    Dim lng库房ID As Long
    Dim lng供应商ID As Long
    Dim lng药品ID As Long
    Dim lng批次 As Long
    Dim str批号 As String
    Dim str效期 As String
    Dim str产地 As String
    Dim dblOldCost As Double
    Dim dblNewCost As Double
    Dim str发票号 As String
    Dim str发票日期 As String
    Dim dbl发票金额 As Double
    
    Dim dbl包装 As Double
    Dim strUpdate As String
    Dim rsTmp As ADODB.Recordset
    
    Dim blnPrint As Boolean
    Dim blnIgnore As Boolean
    Dim inProc As Integer
    Dim blnOne As Boolean
    Dim blnCancel As Boolean
    
    If Me.BillPrice.Rows = 1 Then Exit Sub
    If Me.BillPrice.TextMatrix(0, 售价列表.药品id) = "" Then Exit Sub
    
    If Me.BillPrice.Text <> "" Then
        Call BillPrice_KeyDown(13, 0, blnCancel)
        If blnCancel = True Then Exit Sub
    End If

    '检测相关输入合法性
    If CheckPrice = False Then Exit Sub
    
    '如果是仅调整收入项目，那么只执行这个
    
    Err = 0: On Error GoTo ErrHand
    If mint调价 = 3 Then
        gcnOracle.BeginTrans
        With Me.BillPrice
            For intCount = 1 To IIf(Trim(.TextMatrix(.Rows - 1, 0)) = "", .Rows - 2, .Rows - 1)
                If Val(.TextMatrix(intCount, 售价列表.原收入ID)) <> Val(.TextMatrix(intCount, 售价列表.现收入ID)) Then
                    gstrSql = "Select 收费细目id, 收入项目id, 原价, 现价, 附术收费率, 加班加价率, 调价说明, 调价id, 缺省价格 " & _
                        " From 收费价目 " & _
                        " Where 收费细目id = [1] And Decode(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD'), Null, 终止日期) Is Null" & _
                        GetPriceClassString("")
                        
                    Set rsTmp = zldatabase.OpenSQLRecord(gstrSql, "取价目信息", Val(.TextMatrix(intCount, 售价列表.药品id)))
                    
                    If Not rsTmp.EOF Then
                        gstrSql = "zl_收费价目_update("
                        '收费细目id_In
                        gstrSql = gstrSql & Val(.TextMatrix(intCount, 售价列表.药品id))
                        '收入项目id_In
                        gstrSql = gstrSql & "," & Val(.TextMatrix(intCount, 售价列表.现收入ID))
                        '原价_In
                        gstrSql = gstrSql & "," & IIf(IsNull(rsTmp!原价), "Null", rsTmp!原价)
                        '现价_In
                        gstrSql = gstrSql & "," & IIf(IsNull(rsTmp!现价), "Null", rsTmp!现价)
                        '附术收费率_In
                        gstrSql = gstrSql & "," & IIf(IsNull(rsTmp!附术收费率), "Null", rsTmp!附术收费率)
                        '加班加价率_In
                        gstrSql = gstrSql & "," & IIf(IsNull(rsTmp!加班加价率), "Null", rsTmp!加班加价率)
                        '调价说明_In
                        gstrSql = gstrSql & "," & IIf(IsNull(rsTmp!调价说明), "Null", "'" & rsTmp!调价说明 & "'")
                        '调价id_In
                        gstrSql = gstrSql & "," & IIf(IsNull(rsTmp!调价id), "Null", rsTmp!调价id)
                        '调价人_In
                        gstrSql = gstrSql & ",'" & gstrUserName & "'"
                        '缺省价格_In
                        gstrSql = gstrSql & "," & IIf(IsNull(rsTmp!缺省价格), "Null", rsTmp!缺省价格)
                        gstrSql = gstrSql & ")"
                        
                        Call zldatabase.ExecuteProcedure(gstrSql, Me.Caption)
                    End If
                End If
            Next
        End With
        gcnOracle.CommitTrans
        
        lngBillId = 0
        lngMediId = 0
        lngItemID = 0
        
        blnModify = False
        Unload Me
        Exit Sub
    End If
    
    dtToday = Sys.Currentdate()

    gstrSql = "select 收费价目_ID.nextval from dual"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "取收费价目序号")
        
    lngAdjId = rsTemp.Fields(0).Value
    
    '再次检查是否存在未执行价格，防止并发
'    If chkImmediately.Value = 1 Then
        If Check存在未执行价格 Then
            Exit Sub
        End If
'    End If
    
    mstrNo = Sys.GetNextNo(9)
    
    gcnOracle.BeginTrans
    With Me.BillPrice
        strOldId = ""
        strNewId = ""
        strID = ""
        For intCount = 1 To IIf(Trim(.TextMatrix(.Rows - 1, 0)) = "", .Rows - 2, .Rows - 1)
            If Val(.TextMatrix(intCount, 售价列表.原收入ID)) <> Val(.TextMatrix(intCount, 售价列表.现收入ID)) Or _
                Val(.TextMatrix(intCount, 售价列表.现价)) <> Val(.TextMatrix(intCount, 售价列表.原价)) Then
                    
                LngCurID = Sys.NextId("收费价目")
                strID = strID & IIf(strID = "", "", ",") & LngCurID
                
                dbl包装 = GetModulus(Val(.TextMatrix(intCount, 售价列表.药品id)))
                
                If .TextMatrix(intCount, 售价列表.类型) = "时价" And mbln时价药品调价 And mint调价 <> 1 Then
                    strTmp = ""
                    lngCurrBatch = -1
                    For n = 1 To BillStore.Rows - 1
                        If Val(.TextMatrix(intCount, 售价列表.药品id)) = Val(BillStore.TextMatrix(n, 库存列表.药品id)) Then
                            If InStr(1, "|" & strTmp, "|" & BillStore.TextMatrix(n, 库存列表.批次) & ",") = 0 Then
                                lngCurrBatch = BillStore.TextMatrix(n, 库存列表.批次)
                                strTmp = strTmp & IIf(strTmp = "", "", "|") & BillStore.TextMatrix(n, 库存列表.批次) & "," & BillStore.TextMatrix(n, 库存列表.现价) / dbl包装
                            End If
                        End If
                    Next
                    str批次价格 = str批次价格 & strTmp
                End If
                str批次价格 = str批次价格 & ";"
                
                If CLng(.RowData(intCount)) <> 0 Then
                    If .RowData(intCount) <> -1 And InStr(1, strOldId & ",", "," & .RowData(intCount) & ",") > 0 Then
                        MsgBox "在一次调价中不能对相同品种(" & .TextMatrix(intCount, 售价列表.品名) & ")重复调价", vbExclamation, gstrSysName
                        gcnOracle.RollbackTrans: .SetFocus: Exit Sub
                    End If
                    If .RowData(intCount) = -1 And InStr(1, strNewId & ",", "," & .TextMatrix(intCount, 售价列表.药品id) & ",") > 0 Then
                        MsgBox "不能对相同品种(" & .TextMatrix(intCount, 售价列表.品名) & ")重复设置价格", vbExclamation, gstrSysName
                        gcnOracle.RollbackTrans: .SetFocus: Exit Sub
                    End If
                    If .RowData(intCount) <> -1 Then
                        strOldId = strOldId & "," & .RowData(intCount)
                    Else
                        strNewId = strNewId & "," & .TextMatrix(intCount, 售价列表.药品id)
                    End If
                    
                    '设置上一次的价格记录终止执行
                    gstrSql = "zl_收费价目_stop(" & .TextMatrix(intCount, 售价列表.药品id) & ","
                    If opt时间(0).Value = True Then
                        gstrSql = gstrSql & "to_date('" & Format(DateAdd("s", -1, dtToday), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    Else
                        gstrSql = gstrSql & "to_date('" & Format(DateAdd("s", -1, Me.dtpRunDate.Value), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    End If
                    gstrSql = gstrSql & ")"
                    Call zldatabase.ExecuteProcedure(gstrSql, Me.Caption)
                    
                    '产生价格记录
                    gstrSql = "zl_收费价目_Insert(" & LngCurID & "," & IIf(.RowData(intCount) = -1, "NUll", .RowData(intCount)) & _
                              "," & .TextMatrix(intCount, 售价列表.药品id) & "," & Val(.TextMatrix(intCount, 售价列表.现收入ID)) & "," & _
                              Round(Val(.TextMatrix(intCount, 售价列表.原价)) / dbl包装, gtype_MaxDigits.dig_零售价) & "," & _
                              Round(Val(.TextMatrix(intCount, 售价列表.现价)) / dbl包装, gtype_MaxDigits.dig_零售价) & _
                              ",NULL,NULL,'" & Me.txtSummary.Text & "'," & lngAdjId & ",'" & Trim(Me.txtValuer.Text) & "',"
                    If opt时间(0).Value = True Then
                        gstrSql = gstrSql & "to_date('" & Format(dtToday, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    Else
                        gstrSql = gstrSql & "to_date('" & Format(Me.dtpRunDate.Value, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    End If
                    gstrSql = gstrSql & ",0,'" & mstrNo & "'," & intCount & ")"
                    Call zldatabase.ExecuteProcedure(gstrSql, Me.Caption)
                    blnPrint = True
                End If
            End If
        Next
    End With
    
    '成本价调价处理
    If mint调价 = 1 Or mint调价 = 2 Then
        If BillStore.TextMatrix(1, 0) <> "" Then
            If BillPrice.Rows = 2 Then
                blnOne = True
            ElseIf BillPrice.Rows = 3 Then
                If BillPrice.TextMatrix(BillPrice.Rows - 1, 0) = "" Then
                    blnOne = True
                End If
            End If
            
            For n = 1 To BillStore.Rows - 1
                If BillStore.TextMatrix(n, 0) = "" Then Exit For
                
                '检查未审核单据
                If blnOne = True Then
                    If CheckUnVerify(Val(BillStore.TextMatrix(n, 库存列表.药品id))) = True Then
                        If MsgBox(BillStore.TextMatrix(n, 库存列表.药品) & "存在未审核单据，调整成本价可能会造成差价误差。" & _
                            vbCrLf & Space(4) & "建议先处理未审核单据。是否还继续调价？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            gcnOracle.RollbackTrans
                            Exit Sub
                        End If
                    End If
                Else
                    If blnIgnore = False Then
                        If CheckUnVerify(Val(BillStore.TextMatrix(n, 库存列表.药品id))) = True Then
                            inProc = frmMsgBox.ShowMsgBox(BillStore.TextMatrix(n, 库存列表.药品) & "存在未审核单据，调整成本价可能会造成差价误差。" & _
                                vbCrLf & Space(4) & "建议先处理未审核单据。是否还继续调价？", Me)
                            
                            If inProc = vbNo Or inProc = vbCancel Then
                                gcnOracle.RollbackTrans
                                Exit Sub
                            ElseIf inProc = vbIgnore Then
                                blnIgnore = True
                            End If
                        End If
                    End If
                End If
                
                For i = 1 To BillPay.Rows - 1
                    If BillPay.TextMatrix(i, 0) = "" Then Exit For
                    If Val(BillStore.TextMatrix(n, 库存列表.药品id)) = Val(BillPay.TextMatrix(i, 应付款列.药品id)) Then
                        lng库房ID = Val(BillStore.TextMatrix(n, 库存列表.库房id))
                        lng供应商ID = Val(BillStore.TextMatrix(n, 库存列表.供应商ID))
                        lng药品ID = Val(BillStore.TextMatrix(n, 库存列表.药品id))
                        lng批次 = Val(BillStore.TextMatrix(n, 库存列表.批次))
                        str批号 = BillStore.TextMatrix(n, 库存列表.批号)
                        str效期 = IIf(Trim(BillStore.TextMatrix(n, 库存列表.效期)) = "", "", BillStore.TextMatrix(n, 库存列表.效期))
                        str产地 = BillStore.TextMatrix(n, 库存列表.产地)
                        dblOldCost = FormatEx(Val(BillStore.TextMatrix(n, 库存列表.原成本价)) / GetModulus(lng药品ID), gtype_MaxDigits.dig_成本价)
                        dblNewCost = FormatEx(Val(BillStore.TextMatrix(n, 库存列表.现成本价)) / GetModulus(lng药品ID), gtype_MaxDigits.dig_成本价)
                        str发票号 = BillPay.TextMatrix(i, 应付款列.发票号)
                        str发票日期 = Format(BillPay.TextMatrix(i, 应付款列.发票日期), "yyyy-mm-dd")
                        dbl发票金额 = Val(BillPay.TextMatrix(i, 应付款列.发票金额))
                                                
                        gstrSql = "Zl_成本价调价信息_Insert(" & IIf(lng供应商ID = 0, "Null", lng供应商ID) & "," & lng库房ID & "," & lng药品ID & "," & lng批次 & ",'" & str批号 & "'" & _
                                "," & IIf(str效期 = "", "Null", "to_date('" & Format(str效期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ",'" & str产地 & "',Null," & dblOldCost & ", " & dblNewCost & "," & IIf(str发票号 <> "", "'" & str发票号 & "'", "NULL") & "," & IIf(str发票日期 = "", "Null", "to_date('" & Format(str发票日期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ", " & dbl发票金额 & "," & IIf(mbln应付记录 = True, 1, 0) & ")"
                        Call zldatabase.ExecuteProcedure(gstrSql, Me.Caption)
                    End If
                Next
            Next
        End If
    End If
    
    '无库存时调整成本价
    If mint调价 = 1 Or mint调价 = 2 Then
        With Me.BillPrice
            For intCount = 1 To IIf(Trim(.TextMatrix(.Rows - 1, 0)) = "", .Rows - 2, .Rows - 1)
                If .TextMatrix(intCount, 售价列表.是否有库存) = "0" And Val(.TextMatrix(intCount, 售价列表.原成本价)) <> Val(.TextMatrix(intCount, 售价列表.现成本价)) Then
                    dbl包装 = GetModulus(Val(.TextMatrix(intCount, 售价列表.药品id)))

                    lng药品ID = Val(.TextMatrix(intCount, 售价列表.药品id))
                    dblOldCost = Val(Round(Val(.TextMatrix(intCount, 售价列表.原成本价)) / dbl包装, gtype_MaxDigits.dig_成本价))
                    dblNewCost = Val(Round(Val(.TextMatrix(intCount, 售价列表.现成本价)) / dbl包装, gtype_MaxDigits.dig_成本价))
                    
                    gstrSql = "Zl_成本价调价信息_Insert(Null,Null," & lng药品ID & ",0,Null,Null,Null,Null," & dblOldCost & ", " & dblNewCost & ",NULL,Null,0,0)"
                    Call zldatabase.ExecuteProcedure(gstrSql, Me.Caption)
                End If
            Next
        End With
    End If
    
    '立即执行
    If mint调价 = 1 Then
        '单独成本价调价时
        If opt时间(0).Value = True Then
            With Me.BillPrice
                For intCount = 1 To IIf(Trim(.TextMatrix(.Rows - 1, 0)) = "", .Rows - 2, .Rows - 1)
                    gstrSql = "zl_药品收发记录_Adjust(0,0,Null," & Val(.TextMatrix(intCount, 售价列表.药品id)) & ")"
                    Call zldatabase.ExecuteProcedure(gstrSql, Me.Caption)
                Next
            End With
        End If
    Else
        '调售价
        ArrayID = Split(strID, ",")
        Array批次价格 = Split(str批次价格, ";")
        For intCount = 0 To UBound(ArrayID)
            If opt时间(0).Value = True Or BillPrice.RowData(intCount + 1) = -1 Then
                gstrSql = "zl_药品收发记录_Adjust(" & ArrayID(intCount) & "," & Me.Chk定价.Value & ",'" & Array批次价格(intCount) & "')"
                Call zldatabase.ExecuteProcedure(gstrSql, Me.Caption)
            End If
        Next
    End If
    
    '调整指导价格
    With Me.BillPrice
        For intCount = 1 To IIf(Trim(.TextMatrix(.Rows - 1, 0)) = "", .Rows - 2, .Rows - 1)
            dbl包装 = GetModulus(Val(.TextMatrix(intCount, 售价列表.药品id)))
            
            '更新指导零售价
            If Val(.TextMatrix(intCount, 售价列表.原指导售价)) <> Val(.TextMatrix(intCount, 售价列表.现指导售价)) And Val(.TextMatrix(intCount, 售价列表.现指导售价)) <> 0 Then
                strUpdate = Val(Round(Val(.TextMatrix(intCount, 售价列表.现指导售价)) / dbl包装, mintSalePriceDigit))
                
                gstrSql = "zl_药品目录_UpdateCustom(" & Val(.TextMatrix(intCount, 售价列表.药品id)) & ",'指导零售价=" & strUpdate & "')"
                Call zldatabase.ExecuteProcedure(gstrSql, Me.Caption)
            End If
            
            '更新采购限价
            If Val(.TextMatrix(intCount, 售价列表.原采购限价)) <> Val(.TextMatrix(intCount, 售价列表.现采购限价)) And Val(.TextMatrix(intCount, 售价列表.现采购限价)) <> 0 Then
                strUpdate = Val(Round(Val(.TextMatrix(intCount, 售价列表.现采购限价)) / dbl包装, mintSalePriceDigit))
                                
                gstrSql = "zl_药品目录_UpdateCustom(" & Val(.TextMatrix(intCount, 售价列表.药品id)) & ",'指导批发价=" & strUpdate & "')"
                Call zldatabase.ExecuteProcedure(gstrSql, Me.Caption)
            End If
        Next
    End With
    
    gcnOracle.CommitTrans
    
    If blnPrint = True Then
        If MsgBox("你需要打印调价通知单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1023_3", Me, "NO=" & mstrNo, "包装单位=" & int药库单位, 2)
        End If
    End If
                
    lngBillId = 0
    lngMediId = 0
    lngItemID = 0
    
    blnModify = False
    
    BillPrice.ClearBill
    BillStore.ClearBill
    BillPay.ClearBill
    
    BillPrice.SetFocus
    Exit Sub
    
ErrHand:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
    Me.BillPrice.SetFocus
End Sub

Private Sub cmdPrint_Click()
   Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1023_3", Me, "NO=" & mstrNo, "包装单位=" & int药库单位, 1)
End Sub

Private Sub cmdPstor_Click()
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    
    If Trim(Me.BillStore.TextMatrix(1, 库存列表.库房)) = "" Then Exit Sub
    
    objPrint.Title.Text = "调价库存变动表"
    
    Set objRow = New zlTabAppRow
    objRow.Add "调价说明:" & Me.txtSummary.Text
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "执行时间:" & Format(IIf(opt时间(0).Value = True, Sys.Currentdate, Me.dtpRunDate.Value), "yyyy年MM月DD日 HH:mm:ss")
    objRow.Add "调价人:" & Me.txtValuer.Text
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人:" & gstrUserName
    objRow.Add "打印时间:" & Format(Sys.Currentdate, "yyyy年MM月DD日 HH:mm:ss")
    objPrint.BelowAppRows.Add objRow
    
    Set objPrint.Body = Me.BillStore.MsfObj
    objPrint.PageFooter = 2
     
    Select Case zlPrintAsk(objPrint)
    Case 1
         zlPrintOrView1Grd objPrint, 1
    Case 2
        zlPrintOrView1Grd objPrint, 2
    Case 3
        zlPrintOrView1Grd objPrint, 3
    End Select
    Set objPrint = Nothing

End Sub

Private Sub CmdSelecter_Click()
    Call GetItem("")
End Sub

Private Sub dtpRunDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Me.cmdOk.SetFocus
End Sub

Private Sub Form_Activate()
    Dim i As Integer
    Dim j As Integer
    Dim strBillPrice As String
    
    If Not blnFirst Then Exit Sub
    blnFirst = False
    
    '-----------------界面显示调整---------------------------------
    Select Case Me.Tag
    Case "5", "6"
        intDrugType = 1
        Me.Caption = "成药调价"
        cmdItem.Visible = False
        chk草药批量调价.Visible = False
    Case "7"
        intDrugType = 2
        Me.Caption = "中草药调价"
        cmdItem.Visible = True
        chk草药批量调价.Visible = True
    End Select
    
'    With cbo执行时间
'        .AddItem "立即生效"
'        .AddItem "指定日期生效"
'    End With
    '-----------------------------------------------------------
    If lngBillId = 0 Then
        If InStr(1, mstrPrivs, "仅调整收入项目") > 0 Then
            '先判断这个，有这个权限就不管其他权限了
            mint调价 = 3
            opt时间(0).Value = True
            opt时间(0).Enabled = False
            opt时间(1).Value = False
            chk按批次.Enabled = False
            dtpRunDate.Enabled = False
        ElseIf InStr(1, mstrPrivs, "成本价管理") = 0 Then
            mint调价 = 0
        Else
            If frmMediPriceNavigation.GetCondition(Me, mstrPrivs, mint调价, mlng供应商ID, mdbl加成率, mbln应付记录) = False Then
                Unload Me
                Exit Sub
            End If
        End If
    End If
    
    Call GetMaxDigit    '获取最大精度

    With cbo售价计算方式
        .AddItem "售价与成本价不关联计算"
        .AddItem "售价按固定比例计算"
        .AddItem "售价按分段加成计算"
        .ListIndex = 0
    End With
    
    Call IniGrid
    
    If lngItemID > 0 Then
        Call IniBatchData
    Else
        Call IniData
    End If
    
    If mint调价 = 0 Then
        sstabDetail.TabVisible(1) = False
        chk按批次.Visible = False
    ElseIf mbln应付记录 = False Then
        sstabDetail.TabVisible(1) = False
    End If
    
    If mint调价 = 1 Then
        opt时间(0).Value = True
        opt时间(0).Enabled = False
        opt时间(1).Enabled = False
    End If
    
    chk自动计算应付款变动.Visible = sstabDetail.TabVisible(1)
    
    If mint调价 = 2 Then
        chk自动调成本价.Left = IIf(chk自动计算应付款变动.Visible, chk自动计算应付款变动.Left + chk自动计算应付款变动.Width + 1000, dtpRunDate.Left)
    Else
        chk自动调成本价.Visible = False
    End If
    
    If mint调价 = 0 Then
        fraCondition.Height = 800
        fraCondition.Top = (fraLine.Top - fraCondition.Height - lblHelp.Height) + 80
        BillPrice.Height = fraCondition.Top - 250
    End If
    If InStr(1, mstrPrivs, "仅调整收入项目") > 0 Then
        If gstrDBUser = "ZLHIS" Then
            lblInfo.Visible = True
'            If chk自动调成本价.Visible = False Then
'                lblInfo.Left = chk按批次.Left + chk按批次.Width + 1000
'            Else
                lblInfo.Left = dtpRunDate.Left
'            End If
        Else
            lblInfo.Visible = False
        End If
    Else
        lblInfo.Visible = False
'        If chk自动调成本价.Visible = False Then
'            lblInfo.Left = chk按批次.Left + chk按批次.Width + 1000
'        Else
            lblInfo.Left = dtpRunDate.Left
'        End If
    End If
    
    lbl调价方式.Left = IIf(lblInfo.Visible = True, lblInfo.Left + lblInfo.Width + 410, chk自动调成本价.Left + chk自动调成本价.Width + 410)
    lbl调价方式.Top = chk自动调成本价.Top
    cbo售价计算方式.Left = lbl调价方式.Left + lbl调价方式.Width + 50
    cbo售价计算方式.Top = lbl调价方式.Top - 50
    If mint调价 = 2 Then
        lbl调价方式.Visible = True
        cbo售价计算方式.Visible = True
    Else
        lbl调价方式.Visible = False
        cbo售价计算方式.Visible = False
    End If
    
    If opt时间(0).Value <> True Then
        opt时间(1).Value = True
    End If
    With BillPrice
        For i = 1 To .Rows - 1
            For j = 1 To .Cols - 1
                strBillPrice = strBillPrice & .TextMatrix(i, j) & "|"
            Next
        Next
    End With
    mstr所有记录 = ""
    mstr所有记录 = strBillPrice & "|" & txtSummary.Text & "|" & txtValuer.Text & "|" & opt时间(0).Value & "|" & opt时间(1).Value & "|" & dtpRunDate.Value & "|" & Chk定价.Value & "|" & chk草药批量调价.Value & "|" & _
                    chk按批次 & "|" & chk自动计算应付款变动.Value & "|" & chk自动调成本价.Value
    
    Call SetColor
    Call RestoreWinState(Me)
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim cur零售价 As Currency
    
    On Error Resume Next
    If KeyCode = vbKeyEscape Then
        If Me.ActiveControl.Name = "lvwItem" Then
            lvwItem.Visible = False
            BillPrice.SetFocus
        Else
            cmdCanc_Click
        End If
    ElseIf KeyCode = vbKeyF3 Then
        If BillPrice.Col <> 售价列表.现价 Then Exit Sub
        cur零售价 = frmMediPriceCpt.ShowMe(int药库单位, Val(BillPrice.TextMatrix(BillPrice.Row, 售价列表.药品id)))
        If cur零售价 <> 0 Then
            BillPrice.TextMatrix(BillPrice.Row, 售价列表.现价) = FormatEx(cur零售价, mintPriceDigit)
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    mbln时价药品调价 = (zldatabase.GetPara("时价药品按批次调价", glngSys, 1023, 0) = 1)
    mbln限价提示 = (zldatabase.GetPara("限价提示", glngSys, 1023, 1) = 1)
    
    mstrPrivs = ";" & GetPrivFunc(glngSys, 1023) & ";"
    
    '判断是否以药库单位显示
    int药库单位 = Val(zldatabase.GetPara(29, glngSys))
    
    mintCostDigit = GetDigit(1, 1, IIf(int药库单位 = 0, 1, 4))
    mintPriceDigit = GetDigit(1, 2, IIf(int药库单位 = 0, 1, 4))
    mintNumberDigit = GetDigit(1, 3, IIf(int药库单位 = 0, 1, 4))
    mintMoneyDigit = GetDigit(1, 4)
    mstrMoneyFormat = "0." & String(mintMoneyDigit, "0")
    
    mintSalePriceDigit = GetDigit(1, 2, 1)
    
    blnFirst = True
End Sub

Private Sub SetColor()
    '控制界面中表格的颜色
    Dim i As Long
    
    For i = 1 To BillPrice.Cols - 1
        If BillPrice.ColData(i) = 5 Or BillPrice.ColData(i) = 0 Then
            BillPrice.SetColColor i, &HE7CFBA
        Else
            BillPrice.SetColColor i, vbWhite
        End If
    Next
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    If Me.Height < 8100 Then
        Me.Height = 8100
    End If
    If Me.Width < 12075 Then
        Me.Width = 12075
    End If
    
    Me.cmdOk.Left = Me.ScaleWidth - Me.cmdOk.Width - 150
    Me.cmdCanc.Left = Me.cmdOk.Left
    Me.cmdPrint.Left = Me.cmdOk.Left
    Me.cmdHelp.Left = Me.cmdOk.Left
    Me.cmdItem.Left = Me.cmdOk.Left
    
    Me.BillPrice.Width = Me.cmdOk.Left - 150
    lblHelp.Left = BillPrice.Left + 80
    lblHelp.Top = BillPrice.Top + BillPrice.Height + 80
    lblHelp.Height = 450
    lbl执行时间.Left = lblHelp.Left
    opt时间(0).Left = txtSummary.Left
    opt时间(1).Left = opt时间(0).Left + 1000 + opt时间(0).Width
    Me.fraCondition.Width = Me.BillPrice.Width
    Me.fraLine.Left = Me.BillPrice.Left
    Me.fraLine.Width = Me.Width
    fraLine.Top = fraCondition.Top + fraCondition.Height + 50
    Me.txtValuer.Left = Me.fraCondition.Width - Me.txtValuer.Width
    Me.lblValuer.Left = txtValuer.Left - lblValuer.Width - 50
    Me.txtSummary.Width = lblValuer.Left - txtSummary.Left - 300
    
    Me.chk按批次.Left = Me.lblSummary.Left
        
    Me.dtpRunDate.Left = opt时间(1).Left + opt时间(1).Width + 1000
    
    If dtpRunDate.Visible = True Then
        Me.Chk定价.Left = dtpRunDate.Left + dtpRunDate.Width + 1000
    Else
        Me.Chk定价.Left = opt时间(1).Left + opt时间(1).Width + 1000
    End If
    Me.chk草药批量调价.Left = Chk定价.Left + Chk定价.Width + 1000

    Me.chk自动计算应付款变动.Left = IIf(chk自动计算应付款变动.Visible = True, chk自动计算应付款变动.Left, Chk定价.Left)
'    If Me.chk自动计算应付款变动.Left < chk按批次.Left + chk按批次.Width + 100 Then
'        Me.chk自动计算应付款变动.Left = chk按批次.Left + chk按批次.Width + 100
'    End If
     
    Me.cmdPstor.Left = Me.cmdOk.Left + Me.cmdOk.Width - Me.cmdPstor.Width
    cmdPstor.Top = fraLine.Top + fraLine.Height + 10
    sstabDetail.Top = fraLine.Top + fraLine.Height + 50
    Me.sstabDetail.Width = Me.ScaleWidth - 50
    Me.sstabDetail.Height = Me.ScaleHeight - Me.sstabDetail.Top - 50
    
    Me.BillStore.Width = sstabDetail.Width - 200
    Me.BillStore.Height = sstabDetail.Height - 500
    
    Me.BillPay.Width = Me.BillStore.Width
    Me.BillPay.Height = Me.BillStore.Height
    lbl调价方式.Left = IIf(lblInfo.Visible = True, lblInfo.Left + lblInfo.Width + 410, chk自动调成本价.Left + chk自动调成本价.Width + 410)
    lbl调价方式.Top = chk自动调成本价.Top
    cbo售价计算方式.Left = lbl调价方式.Left + lbl调价方式.Width + 50
    cbo售价计算方式.Top = lbl调价方式.Top - 50
        
'    lblHelp.Top = IIf(Me.cmdItem.Visible = True, Me.cmdItem.Top + Me.cmdItem.Height + 50, Me.CmdHelp.Top + Me.CmdHelp.Height + 50)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If blnModify Then If MsgBox("你确定要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Cancel = 1: Exit Sub
    
    mstrAdjMsg = ""
    
    SaveWinState Me
End Sub

Private Sub lvwItem_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvwItem
        .Sorted = False
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIf(.SortOrder = lvwDescending, lvwAscending, lvwDescending)
        .Sorted = True
    End With
End Sub

Private Sub lvwItem_DblClick()
    Dim lngOldDrugId As Long
    
    If Me.lvwItem.SelectedItem Is Nothing Then Exit Sub
    Set objItem = Me.lvwItem.SelectedItem
    If Me.lvwItem.Tag = 售价列表.品名 Then
        If CheckDrugRepeat(Val(Mid(objItem.Key, 2))) = False Then Exit Sub
        
        With Me.BillPrice
            lngOldDrugId = Val(.TextMatrix(.Row, 售价列表.药品id))
            .TextMatrix(.Row, 售价列表.药品id) = Mid(objItem.Key, 2)
            .TextMatrix(.Row, 售价列表.品名) = "[" & objItem.Text & "]" & objItem.SubItems(Me.lvwItem.ColumnHeaders("名称").Index - 1)
            .TextMatrix(.Row, 售价列表.规格) = objItem.SubItems(Me.lvwItem.ColumnHeaders("规格").Index - 1)
            .TextMatrix(.Row, 售价列表.产地) = objItem.SubItems(Me.lvwItem.ColumnHeaders("产地").Index - 1)
            .TextMatrix(.Row, 售价列表.单位) = objItem.SubItems(Me.lvwItem.ColumnHeaders("单位").Index - 1)
            .TextMatrix(.Row, 售价列表.类型) = objItem.SubItems(Me.lvwItem.ColumnHeaders("类型").Index - 1)
            .TextMatrix(.Row, 售价列表.原成本价) = objItem.SubItems(Me.lvwItem.ColumnHeaders("成本价").Index - 1)
            .TextMatrix(.Row, 售价列表.现成本价) = objItem.SubItems(Me.lvwItem.ColumnHeaders("成本价").Index - 1)
            .TextMatrix(.Row, 售价列表.原采购限价) = objItem.SubItems(Me.lvwItem.ColumnHeaders("采购限价").Index - 1)
            .TextMatrix(.Row, 售价列表.现采购限价) = objItem.SubItems(Me.lvwItem.ColumnHeaders("采购限价").Index - 1)
            .TextMatrix(.Row, 售价列表.原指导售价) = objItem.SubItems(Me.lvwItem.ColumnHeaders("指导售价").Index - 1)
            .TextMatrix(.Row, 售价列表.现指导售价) = objItem.SubItems(Me.lvwItem.ColumnHeaders("指导售价").Index - 1)
            .TextMatrix(.Row, 售价列表.剂量系数) = objItem.SubItems(Me.lvwItem.ColumnHeaders("剂量系数").Index - 1)
            .TextMatrix(.Row, 售价列表.药名ID) = objItem.SubItems(Me.lvwItem.ColumnHeaders("药名ID").Index - 1)
            
            Call zlGetPrice(.Row, .TextMatrix(.Row, 售价列表.药品id), IIf(.TextMatrix(.Row, 售价列表.类型) = "时价", True, False))
            .CmdVisible = False
            
            If mint调价 = 0 Then
                .Col = 售价列表.现价
            ElseIf mint调价 = 1 Or mint调价 = 2 Then
                .Col = 售价列表.现成本价
            ElseIf mint调价 = 3 Then
                .Col = 售价列表.收入名称
            End If
            
            Call GetDrugStore(.Row, Val(Mid(objItem.Key, 2)), lngOldDrugId)
        End With
    Else
        With Me.BillPrice
            .TextMatrix(.Row, 售价列表.现收入ID) = Mid(objItem.Key, 2)
            .TextMatrix(.Row, 售价列表.收入名称) = objItem.SubItems(Me.lvwItem.ColumnHeaders("名称").Index - 1)
            .CmdVisible = False
            .Col = 售价列表.收入名称
        End With
    End If
    Me.lvwItem.Visible = False
    BillPrice.SetFocus
    blnModify = True
End Sub

Private Sub lvwItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> 13 Then Exit Sub
    lvwItem_DblClick
End Sub

Private Sub lvwItem_LostFocus()
    Me.lvwItem.Visible = False
End Sub

Private Sub zlGetPrice(ByVal lngRow As Long, ByVal lngMediId As Long, ByVal blnSeason As Boolean)
    '----------------------------------------------------
    '功能：填写指定药品id的对应价格信息
    '入参：lngMediId药品ID，blnSeason是否时价药品
    '----------------------------------------------------
    On Error GoTo errHandle
    If blnSeason Then
        Me.Chk定价.Enabled = True
        '表示时价药品调价，取库存金额/库存数量做为其价格
        gstrSql = "select P.id,Decode(Nvl(K.库存数量,0),0,P.现价,K.库存金额/Nvl(K.库存数量,1)) 现价,P.执行日期,P.收入项目id,I.名称 as 收入名称" & _
                " from 收费价目 P,收入项目 I," & _
                "   (Select Sum(实际金额) 库存金额,Sum(实际数量) 库存数量" & _
                "    From 药品库存 Where 性质=1 and 药品ID=[1]) K" & _
                " where P.收入项目id=I.id and P.收费细目id=[1] " & _
                "       and (P.终止日期 is null or SYSDATE BETWEEN P.执行日期 AND P.终止日期)" & _
                GetPriceClassString("P")
    Else
        '非时价药品调价，取其价格记录中的价格
        gstrSql = "select P.id,P.现价,P.执行日期,P.收入项目id,I.名称 as 收入名称" & _
                " from 收费价目 P,收入项目 I" & _
                " where P.收入项目id=I.id and P.收费细目id=[1] " & _
                "       and (P.终止日期 is null or SYSDATE BETWEEN P.执行日期 AND P.终止日期)" & _
                GetPriceClassString("P")
    End If
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngMediId)
    
    With rsTemp
        If .RecordCount > 0 Then
            Me.BillPrice.RowData(lngRow) = !ID
            Me.BillPrice.TextMatrix(lngRow, 售价列表.上次日期) = Format(!执行日期, "YYYY-MM-DD HH:MM:SS")
            Me.BillPrice.TextMatrix(lngRow, 售价列表.原价) = FormatEx(!现价 * GetModulus(lngMediId), mintPriceDigit)
            Me.BillPrice.TextMatrix(lngRow, 售价列表.现价) = FormatEx(!现价 * GetModulus(lngMediId), mintPriceDigit)
            Me.BillPrice.TextMatrix(lngRow, 售价列表.现收入ID) = !收入项目id
            Me.BillPrice.TextMatrix(lngRow, 售价列表.原收入ID) = !收入项目id
            Me.BillPrice.TextMatrix(lngRow, 售价列表.收入名称) = !收入名称
        Else
            Me.BillPrice.RowData(lngRow) = -1
            Me.BillPrice.TextMatrix(lngRow, 售价列表.上次日期) = Format(!执行日期, "YYYY-MM-DD HH:MM:SS")
            Me.BillPrice.TextMatrix(lngRow, 售价列表.原价) = FormatEx(0, mintPriceDigit)
            Me.BillPrice.TextMatrix(lngRow, 售价列表.现价) = FormatEx(0, mintPriceDigit)
            If lngRow > 1 Then
                Me.BillPrice.TextMatrix(lngRow, 售价列表.现收入ID) = Me.BillPrice.TextMatrix(lngRow - 1, 售价列表.现收入ID)
                Me.BillPrice.TextMatrix(lngRow, 售价列表.原收入ID) = Me.BillPrice.TextMatrix(lngRow - 1, 售价列表.现收入ID)
                Me.BillPrice.TextMatrix(lngRow, 售价列表.收入名称) = Me.BillPrice.TextMatrix(lngRow - 1, 售价列表.收入名称)
            Else
            
            End If
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub opt时间_Click(Index As Integer)
    Dim bln未执行 As Boolean
    
    If opt时间(0).Value = True Then
        bln未执行 = Check存在未执行价格
        If bln未执行 = True Then
            opt时间(0).Value = False
            opt时间(1).Value = True
        Else
            opt时间(0).Value = True
            opt时间(1).Value = False
        End If
    End If
    
    If opt时间(0).Value = True Then
        Me.dtpRunDate.Enabled = False
    Else
        Me.dtpRunDate.Enabled = True
    End If
    
    On Error Resume Next
    Me.BillPrice.SetFocus
End Sub

Private Sub picItem_Resize()
    On Error Resume Next
    
    With CmdExit
        .Top = picItem.Height - .Height - 50
        .Left = picItem.Width - .Width - 100
    End With
    
    With cmdAdd
        .Top = CmdExit.Top
        .Left = CmdExit.Left - .Width - 50
    End With
    
    With vsfSpec
        .Left = lblItem.Left
        .Top = lblItem.Top + lblItem.Height + 100
        .Width = picItem.Width - .Left - 100
        .Height = CmdExit.Top - .Top - 50
    End With
End Sub


Private Sub txtItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txtItem.Text) = "" Then Exit Sub
    Call GetItem(Trim(txtItem.Text))
End Sub


Private Sub txtSummary_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Me.dtpRunDate.Enabled Then Me.dtpRunDate.SetFocus
End Sub

Private Function CheckPrice() As Boolean
    Dim IntCheck As Integer
    Dim n As Integer
    Dim bln无库存 As Boolean
    
    '检测各执行价格是否正确
    '以及收入项目相同的情况下现价是否与原价相同
    CheckPrice = False
    With BillPrice
        For IntCheck = 1 To .Rows - 1
            If Val(.TextMatrix(IntCheck, 售价列表.药品id)) <> 0 Then
                If Not IsNumeric(Trim(.TextMatrix(IntCheck, 售价列表.现价))) Then
                    MsgBox "第" & IntCheck & "行的药品现价中含有非法字符！", vbInformation, gstrSysName
                    Exit Function
                End If
'                If Val(.TextMatrix(IntCheck, 售价列表.现价)) = 0 Then
'                    MsgBox "第" & IntCheck & "行的药品现价不能为空！", vbInformation, gstrSysName
'                    Exit Function
'                End If
                
                If mint调价 <> 1 Then
                    If Val(.TextMatrix(IntCheck, 售价列表.原收入ID)) = Val(.TextMatrix(IntCheck, 售价列表.现收入ID)) Then
                        If Val(.TextMatrix(IntCheck, 售价列表.现价)) = Val(.TextMatrix(IntCheck, 售价列表.原价)) Then
                            MsgBox "第" & IntCheck & "行的药品现价与原价相同，不能执行调价！", vbInformation, gstrSysName
                            Exit Function
                        End If
                    End If
                End If
                If .TextMatrix(IntCheck, 售价列表.类型) = "时价" And opt时间(0).Value <> True And mint调价 <> 1 Then
                    MsgBox "第" & IntCheck & "行为时价药品，必须设置为立即执行！", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        Next
    End With
    
    CheckPrice = True
End Function

Private Function Check存在未执行价格(Optional ByVal lngDrugId As Long = 0) As Boolean
    Dim RecCheck As New ADODB.Recordset
    Dim LngmediIDThis As Long, IntCheck As Integer
    
    Err = 0
    On Error GoTo ErrHand
    
    If lngDrugId = 0 Then
        '循环判断所有药品
        For IntCheck = 1 To BillPrice.Rows - 1
            LngmediIDThis = Val(BillPrice.TextMatrix(IntCheck, 售价列表.药品id))
            If LngmediIDThis <> 0 Then
                If mint调价 = 0 Or mint调价 = 2 Then
                    '判断是否有未执行的历史价格
                    gstrSql = " Select Count(*) Records From 收费价目 Where 变动原因=0 And 执行日期 > Sysdate And 收费细目ID=[1]" & _
                            GetPriceClassString("")
                    
                    Set RecCheck = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, LngmediIDThis)
                    
                    With RecCheck
                        If Not .EOF Then
                            If Not IsNull(!Records) Then
                                If !Records <> 0 Then
                                    MsgBox "药品" & BillPrice.TextMatrix(IntCheck, 售价列表.品名) & "存在未执行价格，不能设置本次调价！", vbInformation, gstrSysName
                                    Check存在未执行价格 = True
                                    Exit Function
                                End If
                            End If
                        End If
                    End With
                End If
                
                If mint调价 = 1 Or mint调价 = 2 Then
                    '检查是否还有未执行的成本价调价计划
                    gstrSql = "Select 1 From 成本价调价信息 Where 药品id = [1] And 执行日期 Is Null And Rownum = 1 "
                    Set RecCheck = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, LngmediIDThis)
                    
                    If RecCheck.RecordCount > 0 Then
                        MsgBox "药品" & BillPrice.TextMatrix(IntCheck, 售价列表.品名) & "存在未执行成本价，不能设置本次调价！", vbInformation, gstrSysName
                        Check存在未执行价格 = True
                        Exit Function
                    End If
                End If
            End If
        Next
    Else
        If mint调价 = 0 Or mint调价 = 2 Then
            '判断是否有未执行的历史价格
            gstrSql = " Select Count(*) Records From 收费价目 Where 变动原因=0 And 执行日期 > Sysdate And 收费细目ID=[1]" & _
                    GetPriceClassString("")
            
            Set RecCheck = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngDrugId)
            
            With RecCheck
                If Not .EOF Then
                    If Not IsNull(!Records) Then
                        If !Records <> 0 Then
                            Check存在未执行价格 = True
                            Exit Function
                        End If
                    End If
                End If
            End With
        End If
        
        If mint调价 = 1 Or mint调价 = 2 Then
            '检查是否还有未执行的成本价调价计划
            gstrSql = "Select 1 From 成本价调价信息 Where 药品id = [1] And 执行日期 Is Null And Rownum = 1 "
            Set RecCheck = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngDrugId)
            
            If RecCheck.RecordCount > 0 Then
                Check存在未执行价格 = True
                Exit Function
            End If
        End If
    End If
    
   
    Check存在未执行价格 = False
    Exit Function
ErrHand:
    Call ErrCenter
    Call SaveErrLog
    Me.BillPrice.SetFocus

End Function

Private Function GetModulus(ByVal lng药品ID As Long) As Double
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    '返回指定药品的单位系数
    If int药库单位 = 0 Then GetModulus = 1: Exit Function
    
    '提取药库包装系数
    gstrSql = "Select Nvl(药库包装,1) 系数 From 药品规格 Where 药品ID=[1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lng药品ID)
    
    If Not rsTemp.EOF Then GetModulus = rsTemp!系数
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub vsfSpec_EnterCell()
    With vsfSpec
        .Editable = flexEDNone
        If .Col = .ColIndex("选择") Then
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub


