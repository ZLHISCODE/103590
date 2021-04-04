VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmMediPriceDiffCard 
   Caption         =   "零差价药品调价"
   ClientHeight    =   7995
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14010
   Icon            =   "frmMediPriceDiffCard.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7995
   ScaleWidth      =   14010
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picCondition 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   575
      Left            =   120
      ScaleHeight     =   570
      ScaleWidth      =   13575
      TabIndex        =   6
      Top             =   600
      Width           =   13575
      Begin VB.Image imgNote 
         Height          =   480
         Left            =   0
         Picture         =   "frmMediPriceDiffCard.frx":058A
         Top             =   0
         Width           =   480
      End
      Begin VB.Label lblComment 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "说明：1.请选择提取零差价管理药品 2.定价药品全院只有一个价格（售价和成本价相同） 3.时价药品同库房批次的售价和成本价要一致"
         Height          =   180
         Left            =   600
         TabIndex        =   7
         Top             =   150
         Width           =   10800
      End
   End
   Begin VB.PictureBox picInfo 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   240
      ScaleHeight     =   495
      ScaleWidth      =   13335
      TabIndex        =   1
      Top             =   7200
      Width           =   13335
      Begin VB.PictureBox picAdjustTime 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   7800
         ScaleHeight     =   375
         ScaleWidth      =   5535
         TabIndex        =   8
         Top             =   120
         Width           =   5535
         Begin VB.OptionButton opt时间 
            BackColor       =   &H80000003&
            Caption         =   "指定日期"
            Height          =   255
            Index           =   1
            Left            =   1920
            TabIndex        =   10
            Top             =   15
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton opt时间 
            BackColor       =   &H80000003&
            Caption         =   "立即执行"
            Height          =   255
            Index           =   0
            Left            =   840
            TabIndex        =   9
            Top             =   15
            Width           =   1095
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
            Left            =   3000
            TabIndex        =   11
            Top             =   0
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
            Format          =   145031171
            CurrentDate     =   36846.5833333333
         End
         Begin VB.Label lbl执行时间 
            BackColor       =   &H80000003&
            Caption         =   "执行时间"
            Height          =   180
            Left            =   0
            TabIndex        =   12
            Top             =   45
            Width           =   855
         End
      End
      Begin VB.TextBox txtSummary 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   4800
         MaxLength       =   100
         TabIndex        =   2
         Top             =   120
         Width           =   2805
      End
      Begin VB.TextBox txtValuer 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   300
         Left            =   2550
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   120
         Width           =   1125
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   600
         MaxLength       =   100
         TabIndex        =   15
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label lbl查找 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "查找"
         Height          =   180
         Left            =   120
         TabIndex        =   14
         Top             =   180
         Width           =   360
      End
      Begin VB.Label lblValuer 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "调价人"
         Height          =   180
         Left            =   1920
         TabIndex        =   5
         Top             =   180
         Width           =   540
      End
      Begin VB.Label lblSummary 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "调价说明"
         Height          =   180
         Left            =   3960
         TabIndex        =   4
         Top             =   180
         Width           =   720
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfPrice 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   13575
      _cx             =   23945
      _cy             =   8281
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
      BackColorSel    =   15191994
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   10329501
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   22
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmMediPriceDiffCard.frx":0E54
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
   Begin VSFlex8Ctl.VSFlexGrid vsfPrint 
      Height          =   1695
      Left            =   1080
      TabIndex        =   13
      Top             =   8160
      Visible         =   0   'False
      Width           =   11175
      _cx             =   19711
      _cy             =   2990
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
      GridColor       =   0
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
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmMediPriceDiffCard.frx":1199
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
   Begin XtremeCommandBars.ImageManager imgList 
      Bindings        =   "frmMediPriceDiffCard.frx":1312
      Left            =   480
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmMediPriceDiffCard.frx":1326
   End
End
Attribute VB_Name = "frmMediPriceDiffCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'功能按钮
Private Const mconMenu_Save = 100 '确定(&A)
Private Const mconMenu_Quit = 101 '取消(&Q)
Private Const mconMenu_PrintStore = 102 '打印库存变动单(&P)
Private Const mconMenu_ClearAll = 103 '清空列表(&C)
Private Const mconMenu_ClearAllPrice = 109 '清空现价格
Private Const mconMenu_ClearAllDate = 110 '清空界面数据
Private Const mconMenu_Adjust = 104 '自动调价方式
Private Const mconMenu_AdjustByCost = 105 '调价方式：以成本价为准调整售价
Private Const mconMenu_AdjustByPrice = 106 '调价方式：以售价为准调整成本价
Private Const mconMenu_AllDrug = 107 '选择设置了零差价管理的药品
Private Const mconMenu_AllDiff = 108 '提取零差价售价、成本价不想等的药品
Private Const mconMenu_BatchExtraction = 111 '批量提取零差价药品
Private Const mconMenu_Location = 112 '快速定位到一个未调整价格的记录行上
Private Const mconMenu_Find = 113 '查找

Private mintCostDigit As Integer        '成本价小数位数
Private mintPriceDigit As Integer       '售价小数位数
'颜色方案
'Private Const mconlngColor As Long = &HFFFFFF        '不能修改列颜色为白色
Private Const mconlngCanColColor As Long = &HE7CFBA    '能修改列颜色为淡蓝色
Private Const mlngBorderColor As Long = &H0&    '选中行边框颜色
Private Const mlngNoneBorderColor As Long = &HE0E0E0    ' 没选中行边框颜色

Private mstr纪录流水号 As String '纪录调价模式
Private marrSql() As Variant     '纪录存储过程的数组

Private mintUnit As Integer     '用来记录启用的是什么单位
Private mstr药品ID As String
Private mintRow As Integer    '没有填写价格的具体行号
Private mrsFindName As ADODB.Recordset '记录查询数据集
Private mlngFindCurrRow As Long             '查询到的当前行
Private Const MStrCaption As String = "零差价药品调价"

Private Sub GetPartPriceDiff(Optional bln提示 As Boolean = True)
    '提取所有已设置了零差价管理但价格不一致的药品
    Dim rsData As ADODB.Recordset
    Dim bln存在未执行价格 As Boolean
    Dim int序号 As Integer
    
    On Error GoTo errHandle
       
    Call setNOtExcetePrice
    
    gstrSQL = "Select 药品id, 通用名, 规格, 0 As 库房id, '' As 库房, 生产商, '' As 批号, 批次, 单位, 包装系数, 售价, Sum(成本价 * 实际数量) / Sum(实际数量) As 成本价, 是否时价," & vbNewLine & _
                    "       有库存, 价格id, 收入项目id, Null As 上次供应商id, Null As 效期" & vbNewLine & _
                    "From (Select a.药品id, '[' || c.编码 || ']' || c.名称 As 通用名, c.规格, c.产地 As 生产商, 0 As 批次," & vbNewLine & _
                    "              Decode([1], 0, a.药库单位, 2, a.住院单位, 1, a.门诊单位, c.计算单位) As 单位," & vbNewLine & _
                    "              Decode([1], 0, a.药库包装, 2, a.住院包装, 1, a.门诊包装, 1) As 包装系数, b.现价 As 售价, decode(d.平均成本价,null,a.成本价,d.平均成本价) As 成本价, 0 As 是否时价, d.实际数量," & vbNewLine & _
                    "              1 As 有库存, b.Id As 价格id, b.收入项目id" & vbNewLine & _
                    "       From 药品规格 A, 收费价目 B, 收费项目目录 C, 药品库存 D" & vbNewLine & _
                    "       Where a.药品id = b.收费细目id And a.药品id = c.Id And a.药品id = d.药品id And d.性质 = 1 And (Sysdate Between b.执行日期 And b.终止日期) And" & vbNewLine & _
                    "             (c.撤档时间 Is Null Or c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And c.是否变价 = 0 And Nvl(a.是否零差价管理, 0) = 1 And" & vbNewLine & _
                    "             b.现价 <> decode(d.平均成本价,null,a.成本价,d.平均成本价) " & vbNewLine & _
                    "  And Not (zl_fun_getbatchpro(d.库房id,d.药品id)=1 And Nvl(d.批次,0) = 0 And d.可用数量 < 0 And d.实际数量 = 0 And d.实际金额 = 0 And d.实际差价 = 0)) " & vbNewLine & _
                    "Group By 药品id, 通用名, 规格, 生产商, 批次, 单位, 包装系数, 售价, 价格id, 收入项目id, 是否时价, 有库存 " & vbNewLine & _
                    "Union All "

    gstrSQL = gstrSQL & " Select a.药品id, '[' || c.编码 || ']' || c.名称 As 通用名, c.规格, d.库房id, e.名称 As 库房, d.上次产地 As 生产商, d.上次批号 As 批号, d.批次," & vbNewLine & _
                    "       Decode([1], 0, a.药库单位, 2, a.住院单位, 1, a.门诊单位, c.计算单位) As 单位," & vbNewLine & _
                    "       Decode([1], 0, a.药库包装, 2, a.住院包装, 1, a.门诊包装, 1) As 包装系数, d.零售价 As 售价, decode(d.平均成本价,null,a.成本价,d.平均成本价) As 成本价, 1 As 是否时价, 1 As 有库存," & vbNewLine & _
                    "       b.Id As 价格id, b.收入项目id, nvl(d.上次供应商id,0) As 上次供应商id, d.效期" & vbNewLine & _
                    "From 药品规格 A, 收费价目 B, 收费项目目录 C, 药品库存 D, 部门表 E" & vbNewLine & _
                    "Where a.药品id = b.收费细目id And a.药品id = c.Id And a.药品id = d.药品id And d.库房id = e.Id And d.性质 = 1 And" & vbNewLine & _
                    "      (Sysdate Between b.执行日期 And b.终止日期) And c.是否变价 = 1 And" & vbNewLine & _
                    "      (c.撤档时间 Is Null Or c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And Nvl(a.是否零差价管理, 0) = 1 And d.零售价 <> decode(d.平均成本价,null,a.成本价,d.平均成本价)  " & vbNewLine & _
                    "  And Not (zl_fun_getbatchpro(d.库房id,d.药品id)=1 And Nvl(d.批次,0) = 0 And d.可用数量 < 0 And d.实际数量 = 0 And d.实际金额 = 0 And d.实际差价 = 0) " & vbNewLine & _
                    "Union All" & vbNewLine & _
                    "Select a.药品id, '[' || c.编码 || ']' || c.名称 As 通用名, c.规格, 0 As 库房id, '' As 库房, '' As 生产商, '' As 批号, 0 As 批次," & vbNewLine & _
                    "       Decode([1], 0, a.药库单位, 2, a.住院单位, 1, a.门诊单位, c.计算单位) As 单位," & vbNewLine & _
                    "       Decode([1], 0, a.药库包装, 2, a.住院包装, 1, a.门诊包装, 1) As 包装系数, b.现价 As 售价, a.成本价, c.是否变价 As 是否时价, 0 As 有库存," & vbNewLine & _
                    "       b.Id As 价格id, b.收入项目id, Null As 上次供应商id, Null As 效期" & vbNewLine & _
                    "From 药品规格 A, 收费价目 B, 收费项目目录 C" & vbNewLine & _
                    "Where a.药品id = c.Id And a.药品id = b.收费细目id And (c.撤档时间 Is Null Or c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And" & vbNewLine & _
                    "      Nvl(a.是否零差价管理, 0) = 1 And b.现价 <> a.成本价 And (Sysdate Between b.执行日期 And b.终止日期) And Not Exists" & vbNewLine & _
                    " (Select 1 From 药品库存 D Where d.药品id = a.药品id And d.性质 = 1)" & vbNewLine & _
                    "Order By 药品id, 库房id, 批号,批次"

    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "GetPartPriceDiff", mintUnit)
    
    With vsfPrice
        .Redraw = flexRDNone
        .MergeCells = flexMergeNever
        .rows = 1
        
        If rsData.RecordCount = 0 Then
            .rows = 2
            .Cell(flexcpText, 1, 1, 1, .Cols - 1) = "没有找到不满足零差价管理模式的药品......"
            .MergeCells = flexMergeRestrictRows
            .MergeRow(1) = True
        Else
            Do While Not rsData.EOF
                '检查是否存在未执行价格，如果存在就不取数据
                If CheckExistExecutePrice(Val(rsData!药品id)) = False Then
                    .rows = .rows + 1
                    
                    .TextMatrix(.rows - 1, .ColIndex("序号")) = int序号 + 1
                    .TextMatrix(.rows - 1, .ColIndex("药品id")) = rsData!药品id
                    .TextMatrix(.rows - 1, .ColIndex("药价属性")) = IIf(rsData!是否时价 = 1, "时价", "定价")
                    .TextMatrix(.rows - 1, .ColIndex("品名")) = rsData!通用名
                    .TextMatrix(.rows - 1, .ColIndex("规格")) = rsData!规格
                    .TextMatrix(.rows - 1, .ColIndex("生产商")) = Nvl(rsData!生产商, "")
                    .TextMatrix(.rows - 1, .ColIndex("库房id")) = rsData!库房id
                    .TextMatrix(.rows - 1, .ColIndex("库房")) = Nvl(rsData!库房, "")
                    .TextMatrix(.rows - 1, .ColIndex("批号")) = Nvl(rsData!批号, "")
                    .TextMatrix(.rows - 1, .ColIndex("单位")) = rsData!单位
                    .TextMatrix(.rows - 1, .ColIndex("包装系数")) = rsData!包装系数
                    .TextMatrix(.rows - 1, .ColIndex("原售价")) = zlStr.FormatEx(rsData!售价 * rsData!包装系数, mintPriceDigit, , True)
                    .TextMatrix(.rows - 1, .ColIndex("原成本价")) = zlStr.FormatEx(rsData!成本价 * rsData!包装系数, mintCostDigit, , True)
                    .TextMatrix(.rows - 1, .ColIndex("原始售价")) = rsData!售价
                    .TextMatrix(.rows - 1, .ColIndex("原始成本价")) = rsData!成本价
                    .TextMatrix(.rows - 1, .ColIndex("有库存")) = rsData!有库存
                    .TextMatrix(.rows - 1, .ColIndex("价格id")) = rsData!价格id
                    .TextMatrix(.rows - 1, .ColIndex("收入项目id")) = rsData!收入项目ID
                    .TextMatrix(.rows - 1, .ColIndex("批次")) = Nvl(rsData!批次, 0)
                    .TextMatrix(.rows - 1, .ColIndex("上次供应商ID")) = Nvl(rsData!上次供应商ID)
                    .TextMatrix(.rows - 1, .ColIndex("效期")) = Nvl(rsData!效期)
                    
                    .Cell(flexcpForeColor, .rows - 1, .ColIndex("药价属性"), .rows - 1, .ColIndex("药价属性")) = IIf(rsData!是否时价 = 1, vbRed, vbBlack)
                    int序号 = int序号 + 1
                Else
                    bln存在未执行价格 = True
                End If
                
                rsData.MoveNext
            Loop
            
            If .rows >= 2 Then
                .Cell(flexcpBackColor, 1, .ColIndex("现价格"), .rows - 1, .ColIndex("现价格")) = mconlngCanColColor
                .Cell(flexcpForeColor, 1, .ColIndex("现价格"), .rows - 1, .ColIndex("现价格")) = vbBlue
                .Cell(flexcpFontBold, 1, .ColIndex("现价格"), .rows - 1, .ColIndex("现价格")) = True
            End If
            
            .rows = .rows + 1
            .RowHidden(.rows - 1) = True
        End If
        
        .Redraw = flexRDDirect
    End With
    
    txtValuer.Text = UserInfo.用户姓名
    txtSummary.Text = "零差价调价"
    
    txtValuer.Tag = "部分零差价药品"
    
    If bln存在未执行价格 = True Then
        If bln提示 Then
            MsgBox "部分零差价管理药品还存在未执行的预调价记录，本次零差价批量调价列表不在显示这些药品，请注意查看！", vbInformation, gstrSysName
        Else
            MsgBox "本次零差价管理药品批量调价中有部分药品未进行调价，请注意查看！", vbInformation, gstrSysName
        End If
    End If
        
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetBorder()
    '设置行选中边框
    Dim intRow As Integer
    
    With vsfPrice
        If .rows <> 1 Then
            For intRow = 1 To .rows - 2
                .CellBorderRange intRow, 0, intRow, .Cols - 1, mlngNoneBorderColor, 0, 0, 0, 0, 0, 0
            Next
            
            .CellBorderRange .Row, .ColIndex("序号"), .Row, .ColIndex("现价格"), mlngBorderColor, 0, 2, 0, 2, 0, 2
        End If
    End With
End Sub


Private Function CheckExistExecutePrice(ByVal lngDrugID As Long) As Boolean
    '功能 ：检查是否存在未执行的价格
    '返回：true-存在未执行价格；false-不存在未执行价格
    Dim RecCheck As New ADODB.Recordset
    
    On Error GoTo errHandle

    '判断是否有未执行的历史价格
    gstrSQL = " Select 1 Records From 收费价目 Where 变动原因=0 And 执行日期 > Sysdate And 收费细目ID=[1]"
    Set RecCheck = zlDatabase.OpenSQLRecord(gstrSQL, "CheckExistExecutePrice", lngDrugID)
    
    If Not RecCheck.EOF Then CheckExistExecutePrice = True: Exit Function
    
    '检查是否还有未执行的成本价调价计划
    gstrSQL = "Select 1 From 药品价格记录 Where 药品id = [1] And 记录状态 = 0 And Rownum < 2 "
    Set RecCheck = zlDatabase.OpenSQLRecord(gstrSQL, "CheckExistExecutePrice", lngDrugID)
    
    If Not RecCheck.EOF Then CheckExistExecutePrice = True: Exit Function
    
    CheckExistExecutePrice = False
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub initCommandBars()
    Dim cbrToolBar As CommandBar
    Dim cbrControl As CommandBarControl
    Dim cbrControlPopu As CommandBarControl
    Dim lngCount As Integer
    
    With CommandBarsGlobalSettings
        .App = App
        .CompanyName = "重庆中联信息产业有限责任公司" '公司名称
        .ResourceFile = .OcxPath & "\XTPResourceZhCn.dll" '设置中文语言资源文件
        .ColorManager.SystemTheme = xtpSystemThemeAuto  '控件整体的颜色方案
    End With

    With cbsMain.Options
        .ShowExpandButtonAlways = False '总是在工具栏右侧显示选项按钮,即使窗体宽度足够。
        .ToolBarAccelTips = True '显示按钮提示
        .AlwaysShowFullMenus = False '不常用的菜单项先隐藏
        .UseFadedIcons = True '图标显示为褪色效果
        .IconsWithShadow = True '鼠标指向的命令图标显示阴影效果
        .UseDisabledIcons = True '工具栏按钮禁用时图标显示为禁用样式
        .LargeIcons = True '工具栏显示为大图标
        .SetIconSize True, 24, 24 '设置大图标的尺寸
        .SetIconSize False, 16, 16 '设置小图标的尺寸
    End With

    With cbsMain
        .VisualTheme = xtpThemeOffice2003 '设置控件显示风格
        .EnableCustomization False '是否允许自定义设置
        Set .Icons = imgList.Icons '设置关联的图标控件
        .ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap '窗体变化时，如果显示不完菜单也不换行
        .ActiveMenuBar.Title = "菜单"
    End With
    
    '删除现在的工具栏及顶级菜单项
    For lngCount = cbsMain.ActiveMenuBar.Controls.count To 1 Step -1
        cbsMain.ActiveMenuBar.Controls(lngCount).Delete
    Next
    For lngCount = cbsMain.count To 1 Step -1
        cbsMain(lngCount).Delete
    Next
    
    '创建工具栏
    Set cbrToolBar = cbsMain.Add("工具栏", xtpBarTop)
    cbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    cbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    cbrToolBar.ContextMenuPresent = False

    With cbrToolBar
        Set cbrControl = .Controls.Add(xtpControlButton, mconMenu_PrintStore, "打印库存变动单")
        
        Set cbrControl = .Controls.Add(xtpControlPopup, mconMenu_ClearAll, "清空")
        cbrControl.BeginGroup = True
        cbrControl.Id = mconMenu_ClearAll
        cbrControl.IconId = mconMenu_ClearAll
        Set cbrControlPopu = cbrControl.CommandBar.Controls.Add(xtpControlButton, mconMenu_ClearAllPrice, "清空现价格(&A)", -1, False)
        Set cbrControlPopu = cbrControl.CommandBar.Controls.Add(xtpControlButton, mconMenu_ClearAllDate, "清空界面数据(&Q)", -1, False)

        Set cbrControl = .Controls.Add(xtpControlPopup, mconMenu_Adjust, "调价方式")
        cbrControl.BeginGroup = True
        cbrControl.Id = mconMenu_Adjust
        cbrControl.IconId = mconMenu_Adjust
        Set cbrControlPopu = cbrControl.CommandBar.Controls.Add(xtpControlButton, mconMenu_AdjustByCost, "按成本价调整售价(&C)", -1, False)
        Set cbrControlPopu = cbrControl.CommandBar.Controls.Add(xtpControlButton, mconMenu_AdjustByPrice, "按售价调整成本价(&P)", -1, False)
        
        Set cbrControl = .Controls.Add(xtpControlPopup, mconMenu_BatchExtraction, "提取零差价药品")
        cbrControl.BeginGroup = True
        cbrControl.Id = mconMenu_BatchExtraction
        cbrControl.IconId = mconMenu_BatchExtraction
        Set cbrControlPopu = cbrControl.CommandBar.Controls.Add(xtpControlButton, mconMenu_AllDrug, "批量提取零差价药品(&E)", -1, False)
        Set cbrControlPopu = cbrControl.CommandBar.Controls.Add(xtpControlButton, mconMenu_AllDiff, "只提取售价和成本价不一致的药品(&R)", -1, False)

        Set cbrControl = .Controls.Add(xtpControlButton, mconMenu_Location, "定位到未调整价格的行")
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Controls.Add(xtpControlButton, mconMenu_Save, "确定")
        cbrControl.BeginGroup = True
        Set cbrControl = .Controls.Add(xtpControlButton, mconMenu_Quit, "退出")
                
        Set cbrControl = .Controls.Add(xtpControlButton, mconMenu_Find, "查找")
        cbrControl.Visible = False
    End With

    For Each cbrControl In cbrToolBar.Controls  '让工具栏中按钮同时显示图标和文字
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    With Me.cbsMain.KeyBindings
        .Add 0, VK_F3, mconMenu_Find
    End With

End Sub

Private Sub SavePriceAdjust()
    '保存或执行调价
    Dim int修改记录 As Integer
    Dim i As Integer
    Dim Array流水号 As Variant
    Dim blnTrans As Boolean
  
    On Error GoTo ErrHand
    
    marrSql = Array()
    Array流水号 = Array()
    mstr纪录流水号 = ""
    
    If vsfPrice.rows <= 1 Then Exit Sub
    
    With vsfPrice
        '检查价格是否全为空
        For i = 1 To .rows - 2
            If .TextMatrix(i, .ColIndex("现价格")) = "" Then
                int修改记录 = int修改记录 + 1
                If int修改记录 = .rows - 2 Then
                    MsgBox "所有行的药品现价格都为空，不能执行调价！", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        Next
        
        '检查现价格是否与原售价、原成本价都相等
        For i = 1 To .rows - 2
            If .TextMatrix(i, .ColIndex("现价格")) = .TextMatrix(i, .ColIndex("原售价")) And .TextMatrix(i, .ColIndex("现价格")) = .TextMatrix(i, .ColIndex("原成本价")) Then
                MsgBox "第【" & i & "】行现价格与原售价和原成本价相等，不能执行调价！", vbInformation, gstrSysName
                .Select i, .ColIndex("现价格")
                Exit Sub
            End If
        Next
        
        '检查现价是否太大
        For i = 1 To .rows - 2
            If Val(.TextMatrix(i, .ColIndex("现价格"))) > 100000 Then
                MsgBox "第【" & i & "】行输入的价格过大，请重新输入！", vbInformation, gstrSysName
                .Select i, .ColIndex("现价格")
                Exit Sub
            End If
        Next

    End With
    
    Call ModifyCostPrice          '调成本价
    Call ModifyRetailPrice        '调售价
    Call ModifyAllPrice             '成本价售价一起调
             
    Array流水号 = Split(Mid(mstr纪录流水号, 2), ";")
    
    For i = 0 To UBound(Array流水号)
        '流水号
        gstrSQL = "Zl_调价汇总记录_Insert(" & Split(Array流水号(i), "|")(0) & ","
        '类型
        gstrSQL = gstrSQL & Split(Array流水号(i), "|")(1) & ","
        '执行日期
        If opt时间(0).Value = True Then
            gstrSQL = gstrSQL & "sysdate" & ","
        Else
            gstrSQL = gstrSQL & "to_date('" & Format(Me.dtpRunDate.Value, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
        End If
        '说明
        gstrSQL = gstrSQL & "'" & MoveSpecialChar(txtSummary.Text) & "',"
        '分类、填置人
        gstrSQL = gstrSQL & "0,'" & MoveSpecialChar(txtValuer.Text) & "')"
        
        ReDim Preserve marrSql(UBound(marrSql) + 1)
        marrSql(UBound(marrSql)) = gstrSQL
    Next
                   
    gcnOracle.BeginTrans: blnTrans = True          '开启事务
    For i = 0 To UBound(marrSql)
        Call zlDatabase.ExecuteProcedure(CStr(marrSql(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTrans = False     '提交事物
    
    If int修改记录 = 0 Then
        Unload Me
    ElseIf int修改记录 <> 0 Then
        If txtValuer.Tag = "全部零差价药品" Then
            Call GetAllPriceDiff(False)
        Else
            Call GetPartPriceDiff(False)
        End If
    End If
        
    Exit Sub
ErrHand:
    If blnTrans Then gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Sub SetAdjust(ByVal intAdjustType As Integer)
    '批量设置调价
    'intAdjustType：0-以原成本价为准调整售价；1-以原售价为准调整成本价
    Dim i As Integer
    
    With vsfPrice
        If .rows <= 1 Then Exit Sub
        If Val(.TextMatrix(1, .ColIndex("药品id"))) = 0 Then Exit Sub
        
        If intAdjustType = 0 Then
            If MsgBox("本操作将以原成本价作为现价格，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Else
            If MsgBox("本操作将以原售价作为现价格，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        
        .Redraw = flexRDNone
        
        For i = 1 To .rows - 2
            .TextMatrix(i, .ColIndex("现价格")) = IIf(intAdjustType = 0, .TextMatrix(i, .ColIndex("原成本价")), .TextMatrix(i, .ColIndex("原售价")))
        Next
        
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
        Case mconMenu_Save  '执行调价
            Call SavePriceAdjust
        Case mconMenu_PrintStore    '打印库存变动单
            Call PrintPrice            '打印
        Case mconMenu_AdjustByCost  '调价方式：以成本价为准调整售价
            Call SetAdjust(0)
        Case mconMenu_AdjustByPrice  '调价方式：以售价为准调整成本价
            Call SetAdjust(1)
        Case mconMenu_ClearAllPrice  '清空所有现价格
            Call ClearAllPrice
        Case mconMenu_ClearAllDate  '清空界面数据
            Call ClearAllDate
            mintRow = 1
        Case mconMenu_AllDrug  '选择零差价药品
            If txtValuer.Tag = "部分零差价药品" And vsfPrice.rows > 1 Then
                 If MsgBox("该操作会清除空面数据并批量提取已选择的零差价管理药品，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
            Call AllDrug
            mintRow = 1
        Case mconMenu_AllDiff  '提取价格不等零差价药品
            If vsfPrice.rows > 1 Then
                If MsgBox("该操作会清空界面数据并只提取售价和成本价不一致的药品，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
            Call GetPartPriceDiff
            mstr药品ID = ""
            mintRow = 1
        Case mconMenu_Location '快速定位到一个未调整价格的记录行上
            Call FindLocation
        Case mconMenu_Find '查找
            txtCode.SetFocus
            If Trim(txtCode.Text) <> "" Then Call FindGridRow(txtCode.Text)
        Case mconMenu_Quit  '取消
            If vsfPrice.rows > 1 Then
                If Val(vsfPrice.TextMatrix(1, vsfPrice.ColIndex("药品id"))) > 0 Then
                    If MsgBox("未保存本次调价，是否退出？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                    
                    Unload Me
                Else
                    Unload Me
                End If
            Else
                Unload Me
            End If
            
            mintRow = 1
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    On Error Resume Next
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    Me.picCondition.Move lngLeft, lngTop, lngRight - lngLeft
    
    Me.picInfo.Move lngLeft, Me.ScaleHeight - Me.picInfo.Height, lngRight - lngLeft
    
    Me.vsfPrice.Move lngLeft, picCondition.Top + picCondition.Height + 50, lngRight - lngLeft, Me.picInfo.Top - Me.picCondition.Top - Me.picCondition.Height - 100
End Sub

Private Sub Form_Load()
    Dim intUnitTemp As Integer
    
    '获取设置的单位
    mintUnit = Val(zlDatabase.GetPara("药品单位", glngSys, 1333, "1"))
    
    Select Case mintUnit
        Case 0 '药库
            intUnitTemp = 4
        Case 1 '门诊
            intUnitTemp = 2
        Case 2 '住院
            intUnitTemp = 3
        Case 3 '售价
            intUnitTemp = 1
    End Select
    '获取各级单位精度
    mintCostDigit = GetDigitTiaoJia(1, 1, intUnitTemp)
    mintPriceDigit = GetDigitTiaoJia(1, 2, intUnitTemp)
 
    '初始化时间为当前时间+1天
    dtpRunDate.Value = DateAdd("d", 1, CDate(Format(Sys.Currentdate(), "yyyy-MM-dd hh:mm:ss")))
    
    Call initCommandBars
    Call RestoreWinState(Me, App.ProductName, MStrCaption)
    
    vsfPrice.rows = 1
    mlngFindCurrRow = 1
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstr药品ID = ""
    mintRow = 1
End Sub

Private Sub opt时间_Click(Index As Integer)
    If Index = 0 Then
        dtpRunDate.Enabled = False
    Else
        dtpRunDate.Enabled = True
    End If
End Sub

Private Sub picCondition_Resize()
    On Error Resume Next
    
    With lblComment
        .Left = 50
        .Height = picCondition.Width - 50
    End With
End Sub

Private Sub picInfo_Resize()
    On Error Resume Next
    
    With picAdjustTime
        .Left = picInfo.Width - .Width - 100
    End With

    With txtSummary
        .Width = picAdjustTime.Left - .Left - 100
    End With
    
End Sub

Private Sub vsfPrice_EnterCell()
    With vsfPrice
        .Editable = flexEDNone
        If .Col = .ColIndex("现价格") Then
            .FocusRect = flexFocusSolid
            .Editable = flexEDKbdMouse
        Else
            .FocusRect = flexFocusLight
        End If
        
        Call SetBorder '设置行选中边框
    End With
End Sub

Private Sub vsfPrice_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strkey As String
    Dim intDigit As Integer

    With vsfPrice
        strkey = .EditText
        If Col = .ColIndex("现价格") Then
            If KeyAscii = vbKeyReturn Then
                .EditText = zlStr.FormatEx(.EditText, mintPriceDigit, , True)
                If Row <> .rows - 2 Then
                    .Row = Row + 1
                    .Col = Col
                End If
                Exit Sub
            End If
            
            If KeyAscii <> vbKeyBack Then
                If KeyAscii = vbKeyDelete Then
                    If InStr(1, strkey, ".") > 0 Then
                        KeyAscii = 0
                    End If
                ElseIf KeyAscii = Asc(".") Or (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
                    If vsfPrice.EditSelLength = Len(strkey) Then Exit Sub
                    If InStr(strkey, ".") <> 0 And Chr(KeyAscii) = "." Then   '只能存在一个小数点
                        KeyAscii = 0
                        Exit Sub
                    End If
'                    If Len(Mid(strkey, InStr(1, strkey, ".") + 1)) > mintPriceDigit And strkey Like "*.*" Then
'                        KeyAscii = 0
'                        Exit Sub
'                    Else
'                        Exit Sub
'                    End If
                Else
                    KeyAscii = 0
                End If
            End If
        End If
    End With
End Sub

Private Sub vsfPrice_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim dbl现价格 As Double
    Dim dbl原始售价 As Double
    Dim dbl原始成本价 As Double
    Dim intRow As Integer
    
    With vsfPrice
        If Col = .ColIndex("现价格") Then
            If Trim(.EditText) = "" Then Exit Sub
            
            .EditText = zlStr.FormatEx(Val(.EditText), mintPriceDigit, , True)
            .TextMatrix(Row, .ColIndex("现价格")) = .EditText
            
            dbl现价格 = Val(zlStr.FormatEx(Val(.TextMatrix(Row, .ColIndex("现价格"))) / Val(.TextMatrix(Row, .ColIndex("包装系数"))), gtype_UserDrugDigits.Digit_零售价, , True))
            dbl原始售价 = Val(.TextMatrix(Row, .ColIndex("原始售价")))
            dbl原始成本价 = Val(.TextMatrix(Row, .ColIndex("原始成本价")))
            
            If dbl现价格 = dbl原始售价 And dbl现价格 = dbl原始成本价 Then
                MsgBox "注意：现售价和原价一样了，请重新录入！", vbInformation, gstrSysName
                Cancel = True
                Exit Sub
            End If
        End If
    End With
End Sub

Private Sub ClearAllPrice()
    '清空所有现价格
    Dim i As Integer
    
    If vsfPrice.rows <= 1 Then Exit Sub
    If MsgBox("未保存本次调价，是否清空所有价格？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub

    For i = 1 To vsfPrice.rows - 2
        vsfPrice.TextMatrix(i, vsfPrice.ColIndex("现价格")) = ""
    Next
End Sub

Private Sub ModifyCostPrice()
    '调成本价
    Dim i As Integer
    Dim bln是否调过成本价 As Boolean
    Dim strCost流水号 As String
    Dim str调过成本价ID As String
    Dim Array调成本价ID As Variant
    Dim rsTemp As ADODB.Recordset
    Dim dbl包装 As Double
    Dim dtToday As Date
    
    Array调成本价ID = Array()
    On Error GoTo ErrHand
    
    With vsfPrice
        For i = 1 To .rows - 2
            If Val(.TextMatrix(i, .ColIndex("原成本价"))) <> Val(.TextMatrix(i, .ColIndex("现价格"))) And Val(.TextMatrix(i, .ColIndex("原售价"))) = Val(.TextMatrix(i, .ColIndex("现价格"))) And .TextMatrix(i, .ColIndex("现价格")) <> "" Then
                bln是否调过成本价 = True
                If strCost流水号 = "" Then
                    gstrSQL = "select nextno(135) as 流水号 from dual"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "调价流水号")
                    If rsTemp.RecordCount = 0 Then
                        MsgBox "调价流水号未能初始化成功，请与管理员联系！", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    strCost流水号 = rsTemp!流水号

                    mstr纪录流水号 = mstr纪录流水号 & ";" & strCost流水号 & "|" & 1
                    dtToday = Sys.Currentdate() - 1 / 24 / 60 / 60
                End If

                If InStr(str调过成本价ID & ";", ";" & Val(.TextMatrix(i, .ColIndex("药品ID"))) & ";") = 0 Then
                    str调过成本价ID = str调过成本价ID & ";" & Val(.TextMatrix(i, .ColIndex("药品ID")))
                End If

                dbl包装 = Val(vsfPrice.TextMatrix(i, vsfPrice.ColIndex("包装系数")))
                
                If .TextMatrix(i, .ColIndex("有库存")) = 1 And .TextMatrix(i, .ColIndex("药价属性")) = "定价" Then
                    gstrSQL = "Select s.库房id, s.药品id, d.名称 As 库房, '[' || m.编码 || ']' || m.名称 As 药品, m.规格,s.上次产地 as 产地," & vbNewLine & _
                                    "       Decode([2], 0, p.药库单位, 2, p.住院单位, 1, p.门诊单位, m.计算单位) As 单位," & vbNewLine & _
                                    "       Decode([2], 0, p.药库包装, 2, p.住院包装, 1, p.门诊包装, 1) As 包装系数," & vbNewLine & _
                                    "       s.上次批号 As 批号, Nvl(s.实际数量, 0) As 数量, Nvl(s.批次,0) as 批次," & vbNewLine & _
                                    "       Nvl(m.是否变价, 0) 变价, m.Id, Decode(Nvl(m.是否变价, 0), 0, e.现价, Decode(Nvl(s.零售价, 0),0,s.实际金额/s.实际数量,s.零售价)) As 时价售价, p.加成率, Decode(Nvl(s.平均成本价, 0), 0, p.成本价, s.平均成本价) As 成本价, nvl(s.上次供应商id,0) As 上次供应商id," & vbNewLine & _
                                    "       n.名称 As 供应商, s.效期" & vbNewLine & _
                                    "From 药品库存 S, 部门表 D, 收费项目目录 M, 药品规格 P, 供应商 N, 收费价目 E" & vbNewLine & _
                                    "Where d.Id = s.库房id And s.药品id = m.Id And m.Id = p.药品id And Nvl(s.上次供应商id, 0) = n.Id(+) And m.Id = e.收费细目id And" & vbNewLine & _
                                    "      s.性质 = 1 And s.药品id = [1] And Sysdate Between e.执行日期 And e.终止日期  And e.价格等级 Is Null" & vbNewLine & _
                                    "Order By s.药品id,s.库房id, s.上次批号,s.批次 "

                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取定价药品信息", Val(.TextMatrix(i, .ColIndex("药品ID"))), mintUnit)
                    
                    With rsTemp
                        Do While Not .EOF
                            gstrSQL = "Zl_药品价格记录_Stop("
                            '价格类型_In
                            gstrSQL = gstrSQL & 2
                            '库房id_In
                            gstrSQL = gstrSQL & "," & !库房id
                            '药品id_In
                            gstrSQL = gstrSQL & "," & !药品id
                            '批次_In
                            gstrSQL = gstrSQL & "," & Nvl(!批次, 0)
                            '终止日期_In
                            If opt时间(0).Value = True Then
                                gstrSQL = gstrSQL & "," & "to_date('" & Format(DateAdd("s", -1, dtToday), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            Else
                                gstrSQL = gstrSQL & "," & "to_date('" & Format(DateAdd("s", -1, Me.dtpRunDate.Value), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            End If
                            gstrSQL = gstrSQL & ")"
                            
                            ReDim Preserve marrSql(UBound(marrSql) + 1)
                            marrSql(UBound(marrSql)) = gstrSQL
                            
                            gstrSQL = "Zl_药品价格记录_Insert("
                            '调价类型_In
                            gstrSQL = gstrSQL & 1
                            '价格类型_In
                            gstrSQL = gstrSQL & ",2"
                            '库房id_In
                            gstrSQL = gstrSQL & "," & !库房id
                            '药品id_In
                            gstrSQL = gstrSQL & "," & !药品id
                            '批次_In
                            gstrSQL = gstrSQL & "," & Nvl(!批次, 0)
                            '原价_In
                            gstrSQL = gstrSQL & "," & zlStr.FormatEx(!成本价, gtype_UserDrugDigits.Digit_零售价, , True)
                            '现价_In
                            gstrSQL = gstrSQL & "," & zlStr.FormatEx(Val(vsfPrice.TextMatrix(i, vsfPrice.ColIndex("现价格"))) / dbl包装, gtype_UserDrugDigits.Digit_零售价, , True)
                            '执行日期_In
                            If opt时间(0).Value = True Then
                                gstrSQL = gstrSQL & "," & "to_date('" & Format(dtToday, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            Else
                                gstrSQL = gstrSQL & "," & "to_date('" & Format(Me.dtpRunDate.Value, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            End If
                            '调价说明_In
                            gstrSQL = gstrSQL & ",'" & Me.txtSummary.Text & "'"
                            '调价人_In
                            gstrSQL = gstrSQL & ",'" & Trim(Me.txtValuer.Text) & "'"
                            '调价汇总号_In
                            gstrSQL = gstrSQL & ",'" & strCost流水号 & "'"
                            '供药单位id_In
                            gstrSQL = gstrSQL & "," & IIf(!上次供应商ID = 0, "Null", !上次供应商ID)
                            '批号_In
                            gstrSQL = gstrSQL & ",'" & Nvl(!批号) & "'"
                            '效期_In
                            gstrSQL = gstrSQL & "," & "to_date('" & Format(IIf(IsNull(!效期), "", !效期), "YYYY-MM-DD") & "','yyyy-mm-dd')"
                            '产地_In
                            gstrSQL = gstrSQL & ",'" & Nvl(!产地) & "'"
                            gstrSQL = gstrSQL & ")"
                            
                            ReDim Preserve marrSql(UBound(marrSql) + 1)
                            marrSql(UBound(marrSql)) = gstrSQL
                            
                        .MoveNext
                        Loop
                    End With
                End If
                            
                If .TextMatrix(i, .ColIndex("有库存")) = 1 And .TextMatrix(i, .ColIndex("药价属性")) = "时价" Then
                    '有库存
                    gstrSQL = "Zl_药品价格记录_Stop("
                    '价格类型_In
                    gstrSQL = gstrSQL & 2
                    '库房id_In
                    gstrSQL = gstrSQL & "," & Val(.TextMatrix(i, .ColIndex("库房id")))
                    '药品id_In
                    gstrSQL = gstrSQL & "," & Val(.TextMatrix(i, .ColIndex("药品id")))
                    '批次_In
                    gstrSQL = gstrSQL & "," & Nvl(.TextMatrix(i, .ColIndex("批次")), 0)
                    '终止日期_In
                    If opt时间(0).Value = True Then
                        gstrSQL = gstrSQL & "," & "to_date('" & Format(DateAdd("s", -1, dtToday), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    Else
                        gstrSQL = gstrSQL & "," & "to_date('" & Format(DateAdd("s", -1, Me.dtpRunDate.Value), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    End If
                    gstrSQL = gstrSQL & ")"
                    
                    ReDim Preserve marrSql(UBound(marrSql) + 1)
                    marrSql(UBound(marrSql)) = gstrSQL
                    
                    gstrSQL = "Zl_药品价格记录_Insert("
                    '调价类型_In
                    gstrSQL = gstrSQL & 1
                    '价格类型_In
                    gstrSQL = gstrSQL & ",2"
                    '库房id_In
                    gstrSQL = gstrSQL & "," & Val(.TextMatrix(i, .ColIndex("库房id")))
                    '药品id_In
                    gstrSQL = gstrSQL & "," & Val(.TextMatrix(i, .ColIndex("药品id")))
                    '批次_In
                    gstrSQL = gstrSQL & "," & Nvl(.TextMatrix(i, .ColIndex("批次")), 0)
                    '原价_In
                    gstrSQL = gstrSQL & "," & zlStr.FormatEx(Val(.TextMatrix(i, .ColIndex("原成本价"))) / dbl包装, gtype_UserDrugDigits.Digit_零售价, , True)
                    '现价_In
                    gstrSQL = gstrSQL & "," & zlStr.FormatEx(Val(.TextMatrix(i, .ColIndex("现价格"))) / dbl包装, gtype_UserDrugDigits.Digit_零售价, , True)
                    '执行日期_In
                    If opt时间(0).Value = True Then
                        gstrSQL = gstrSQL & "," & "to_date('" & Format(dtToday, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    Else
                        gstrSQL = gstrSQL & "," & "to_date('" & Format(Me.dtpRunDate.Value, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    End If
                    '调价说明_In
                    gstrSQL = gstrSQL & ",'" & Me.txtSummary.Text & "'"
                    '调价人_In
                    gstrSQL = gstrSQL & ",'" & Trim(Me.txtValuer.Text) & "'"
                    '调价汇总号_In
                    gstrSQL = gstrSQL & ",'" & strCost流水号 & "'"
                    '供药单位id_In
                    gstrSQL = gstrSQL & "," & IIf(Val(.TextMatrix(i, .ColIndex("上次供应商ID"))) = 0, "Null", Val(.TextMatrix(i, .ColIndex("上次供应商ID"))))
                    '批号_In
                    gstrSQL = gstrSQL & ",'" & Nvl(.TextMatrix(i, .ColIndex("批号"))) & "'"
                    '效期_In
                    gstrSQL = gstrSQL & "," & "to_date('" & Format(IIf(IsNull(.TextMatrix(i, .ColIndex("效期"))), "", .TextMatrix(i, .ColIndex("效期"))), "YYYY-MM-DD") & "','yyyy-mm-dd')"
                    '产地_In
                    gstrSQL = gstrSQL & ",'" & Nvl(.TextMatrix(i, .ColIndex("生产商"))) & "'"
                    gstrSQL = gstrSQL & ")"
                    
                    ReDim Preserve marrSql(UBound(marrSql) + 1)
                    marrSql(UBound(marrSql)) = gstrSQL
                End If
                    
                If .TextMatrix(i, .ColIndex("有库存")) = 0 Then
                    '无库存
                    gstrSQL = "Zl_药品价格记录_Insert("
                    '调价类型_In
                    gstrSQL = gstrSQL & 1
                    '价格类型_In
                    gstrSQL = gstrSQL & ",2"
                    '库房id_In
                    gstrSQL = gstrSQL & ",Null"
                    '药品id_In
                    gstrSQL = gstrSQL & "," & Val(.TextMatrix(i, .ColIndex("药品id")))
                    '批次_In
                    gstrSQL = gstrSQL & ",0"
                    '原价_In
                    gstrSQL = gstrSQL & "," & zlStr.FormatEx(Val(.TextMatrix(i, .ColIndex("原成本价"))) / dbl包装, gtype_UserDrugDigits.Digit_零售价, , True)
                    '现价_In
                    gstrSQL = gstrSQL & "," & zlStr.FormatEx(Val(.TextMatrix(i, .ColIndex("现价格"))) / dbl包装, gtype_UserDrugDigits.Digit_零售价, , True)
                    '执行日期_In
                    If opt时间(0).Value = True Then
                        gstrSQL = gstrSQL & "," & "to_date('" & Format(dtToday, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    Else
                        gstrSQL = gstrSQL & "," & "to_date('" & Format(Me.dtpRunDate.Value, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    End If
                    '调价说明_In
                    gstrSQL = gstrSQL & ",'" & Me.txtSummary.Text & "'"
                    '调价人_In
                    gstrSQL = gstrSQL & ",'" & Trim(Me.txtValuer.Text) & "'"
                    '调价汇总号_In
                    gstrSQL = gstrSQL & ",'" & strCost流水号 & "'"
                    gstrSQL = gstrSQL & ")"
                    
                    ReDim Preserve marrSql(UBound(marrSql) + 1)
                    marrSql(UBound(marrSql)) = gstrSQL
                End If
  
            End If
        Next
    End With
    
    If opt时间(0).Value = True Then
        Array调成本价ID = Split(Mid(str调过成本价ID, 2), ";")
        If bln是否调过成本价 Then
            For i = 0 To UBound(Array调成本价ID)
                gstrSQL = "zl_药品收发记录_Adjust(" & Array调成本价ID(i) & ",2)"
                ReDim Preserve marrSql(UBound(marrSql) + 1)
                marrSql(UBound(marrSql)) = gstrSQL
            Next
            bln是否调过成本价 = False
        End If
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ModifyRetailPrice()
    '调售价
    Dim i As Integer
    Dim n As Integer
    Dim int同药品ID数 As Integer
    Dim int收费价目序号 As Integer
    Dim bln是否调过售价 As Boolean
    Dim dbl现售价 As Double
    Dim strRetail流水号 As String
    Dim lngAdjId As Long
    Dim dtToday As Date
    Dim rsTemp As ADODB.Recordset
    Dim strNo As String
    Dim LngCurID As Long
    Dim dbl包装 As Double
    
    On Error GoTo ErrHand
    
     With vsfPrice
        For i = 1 To .rows - 2
            If Val(.TextMatrix(i, .ColIndex("药品id"))) = Val(.TextMatrix(i + 1, .ColIndex("药品id"))) Then
                int同药品ID数 = int同药品ID数 + 1
            Else
                For n = i - int同药品ID数 To i
                    dbl现售价 = dbl现售价 + Val(.TextMatrix(n, .ColIndex("现价格")))
                    If Val(.TextMatrix(n, .ColIndex("原售价"))) <> Val(.TextMatrix(n, .ColIndex("现价格"))) And Val(.TextMatrix(n, .ColIndex("原成本价"))) = Val(.TextMatrix(n, .ColIndex("现价格"))) And .TextMatrix(i, .ColIndex("现价格")) <> "" Then
                        bln是否调过售价 = True
                    End If
                Next

                If bln是否调过售价 Then
                    If strRetail流水号 = "" Then
                        strNo = Sys.GetNextNo(9)
                        gstrSQL = "select nextno(135) as 流水号 from dual"
                        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "调价流水号")
                        If rsTemp.RecordCount = 0 Then
                            MsgBox "调价流水号未能初始化成功，请与管理员联系！", vbInformation, gstrSysName
                            Exit Sub
                        End If
                        strRetail流水号 = rsTemp!流水号

                        gstrSQL = "select 收费价目_ID.nextval from dual"
                        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取收费价目序号")
                        lngAdjId = rsTemp.Fields(0).Value

                        mstr纪录流水号 = mstr纪录流水号 & ";" & strRetail流水号 & "|" & 0
                        dtToday = Sys.Currentdate() - 1 / 24 / 60 / 60
                    End If
                    
                    int收费价目序号 = int收费价目序号 + 1
                    dbl现售价 = Round(dbl现售价 / (int同药品ID数 + 1), 2)
                    LngCurID = Sys.NextId("收费价目")
                    dbl包装 = Val(.TextMatrix(i, .ColIndex("包装系数")))
            
                    If CLng(.TextMatrix(i, .ColIndex("价格ID"))) <> 0 Then
                        '设置上一次的价格记录终止执行
                        gstrSQL = "zl_收费价目_stop(" & .TextMatrix(i, .ColIndex("药品id")) & ","
                        If opt时间(0).Value = True Then
                            gstrSQL = gstrSQL & "to_date('" & Format(DateAdd("s", -1, dtToday), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                        Else
                            gstrSQL = gstrSQL & "to_date('" & Format(DateAdd("s", -2, Me.dtpRunDate.Value), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                        End If
                        gstrSQL = gstrSQL & ")"
                        
                        ReDim Preserve marrSql(UBound(marrSql) + 1)
                        marrSql(UBound(marrSql)) = gstrSQL
            
                        '产生价格记录
                        'ID
                        gstrSQL = "zl_收费价目_Insert(" & LngCurID & ","
                        '原价ID
                        gstrSQL = gstrSQL & Val(.TextMatrix(i, .ColIndex("价格ID"))) & ","
                        '收费细目ID
                        gstrSQL = gstrSQL & Val(.TextMatrix(i, .ColIndex("药品id"))) & ","
                        '收入项目ID
                        gstrSQL = gstrSQL & Val(.TextMatrix(i, .ColIndex("收入项目ID"))) & ","
                        '原价
                        gstrSQL = gstrSQL & zlStr.FormatEx(Val(.TextMatrix(i, .ColIndex("原售价"))) / dbl包装, gtype_UserDrugDigits.Digit_零售价, , True) & ","
                        '现价
                        gstrSQL = gstrSQL & zlStr.FormatEx(dbl现售价 / dbl包装, gtype_UserDrugDigits.Digit_零售价, , True) & ","
                        '附术收费率、加班加价率、调价说明
                        gstrSQL = gstrSQL & "NULL,NULL,'" & MoveSpecialChar(txtSummary.Text) & "',"
                        '调价id、调价人
                        gstrSQL = gstrSQL & lngAdjId & ",'" & MoveSpecialChar(txtValuer.Text) & "',"
                        '执行日期
                        If opt时间(0).Value = True Then
                            gstrSQL = gstrSQL & "to_date('" & Format(dtToday, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
                        Else
                            gstrSQL = gstrSQL & "to_date('" & Format(DateAdd("s", -1, Me.dtpRunDate.Value), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
                        End If
                        '变动原因
                        gstrSQL = gstrSQL & "0,"
                        'NO
                        gstrSQL = gstrSQL & "'" & strNo & "',"
                        '序号、缺省价格
                        gstrSQL = gstrSQL & int收费价目序号 & ",Null,"
                        '调价汇总号
                        gstrSQL = gstrSQL & strRetail流水号 & ")"
                        
                        ReDim Preserve marrSql(UBound(marrSql) + 1)
                        marrSql(UBound(marrSql)) = gstrSQL
                    End If
                    
                    For n = i - int同药品ID数 To i
                        If Val(.TextMatrix(n, .ColIndex("原售价"))) <> Val(.TextMatrix(n, .ColIndex("现价格"))) And Val(.TextMatrix(n, .ColIndex("原成本价"))) = Val(.TextMatrix(n, .ColIndex("现价格"))) And .TextMatrix(i, .ColIndex("现价格")) <> "" And .TextMatrix(n, .ColIndex("药价属性")) = "时价" And .TextMatrix(n, .ColIndex("有库存")) = 1 Then
                            '时价药品有库存调价
                            gstrSQL = "Zl_药品价格记录_Stop("
                            '价格类型_In
                            gstrSQL = gstrSQL & 1
                            '库房id_In
                            gstrSQL = gstrSQL & "," & Val(.TextMatrix(n, .ColIndex("库房ID")))
                            '药品id_In
                            gstrSQL = gstrSQL & "," & Val(.TextMatrix(n, .ColIndex("药品ID")))
                            '批次_In
                            gstrSQL = gstrSQL & "," & Val(.TextMatrix(n, .ColIndex("批次")))
                            '终止日期_In
                            If opt时间(0).Value = True Then
                                gstrSQL = gstrSQL & "," & "to_date('" & Format(DateAdd("s", -1, dtToday), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            Else
                                gstrSQL = gstrSQL & "," & "to_date('" & Format(DateAdd("s", -1, Me.dtpRunDate.Value), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            End If
                            gstrSQL = gstrSQL & ")"
                            
                            ReDim Preserve marrSql(UBound(marrSql) + 1)
                            marrSql(UBound(marrSql)) = gstrSQL
                            
                            gstrSQL = "Zl_药品价格记录_Insert("
                            '调价类型_In
                            gstrSQL = gstrSQL & 1
                            '价格类型_In
                            gstrSQL = gstrSQL & ",1"
                            '库房id_In
                            gstrSQL = gstrSQL & "," & Val(.TextMatrix(n, .ColIndex("库房ID")))
                            '药品id_In
                            gstrSQL = gstrSQL & "," & Val(.TextMatrix(n, .ColIndex("药品ID")))
                            '批次_In
                            gstrSQL = gstrSQL & "," & Val(.TextMatrix(n, .ColIndex("批次")))
                            '原价_In
                            gstrSQL = gstrSQL & "," & zlStr.FormatEx(Val(.TextMatrix(n, .ColIndex("原售价"))) / dbl包装, gtype_UserDrugDigits.Digit_零售价, , True)
                            '现价_In
                            gstrSQL = gstrSQL & "," & zlStr.FormatEx(Val(.TextMatrix(n, .ColIndex("现价格"))) / dbl包装, gtype_UserDrugDigits.Digit_零售价, , True)
                            '执行日期_In
                            If opt时间(0).Value = True Then
                                gstrSQL = gstrSQL & "," & "to_date('" & Format(dtToday, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            Else
                                gstrSQL = gstrSQL & "," & "to_date('" & Format(Me.dtpRunDate.Value, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            End If
                            '调价说明_In
                            gstrSQL = gstrSQL & ",'" & Me.txtSummary.Text & "'"
                            '调价人_In
                            gstrSQL = gstrSQL & ",'" & Trim(Me.txtValuer.Text) & "'"
                            '调价汇总号_In
                            gstrSQL = gstrSQL & ",'" & strRetail流水号 & "'"
                            '供药单位id_In
                            gstrSQL = gstrSQL & "," & IIf(Val(.TextMatrix(n, .ColIndex("上次供应商ID"))) = 0, "null", Val(.TextMatrix(n, .ColIndex("上次供应商ID"))))
                            '批号_In
                            gstrSQL = gstrSQL & ",'" & Nvl(.TextMatrix(n, .ColIndex("批号"))) & "'"
                            '效期_In
                            gstrSQL = gstrSQL & "," & "to_date('" & Format(Nvl(.TextMatrix(n, .ColIndex("效期"))), "YYYY-MM-DD") & "','yyyy-mm-dd')"
                            '产地_In
                            gstrSQL = gstrSQL & ",'" & Nvl(.TextMatrix(n, .ColIndex("生产商"))) & "'"
                            gstrSQL = gstrSQL & ")"
                            
                            ReDim Preserve marrSql(UBound(marrSql) + 1)
                            marrSql(UBound(marrSql)) = gstrSQL
                        End If
                        
                        If Val(.TextMatrix(n, .ColIndex("原售价"))) <> Val(.TextMatrix(n, .ColIndex("现价格"))) And Val(.TextMatrix(n, .ColIndex("原成本价"))) = Val(.TextMatrix(n, .ColIndex("现价格"))) And .TextMatrix(i, .ColIndex("现价格")) <> "" And .TextMatrix(n, .ColIndex("药价属性")) = "时价" And .TextMatrix(n, .ColIndex("有库存")) = 0 Then
                            '时价药品无库存调价
                            gstrSQL = "Zl_药品价格记录_Insert("
                            '调价类型_In
                            gstrSQL = gstrSQL & 1
                            '价格类型_In
                            gstrSQL = gstrSQL & ",1"
                            '库房id_In
                            gstrSQL = gstrSQL & ",Null"
                            '药品id_In
                            gstrSQL = gstrSQL & "," & Val(.TextMatrix(n, .ColIndex("药品ID")))
                            '批次_In
                            gstrSQL = gstrSQL & ",0"
                            '原价_In
                            gstrSQL = gstrSQL & "," & zlStr.FormatEx(Val(.TextMatrix(n, .ColIndex("原售价"))) / dbl包装, gtype_UserDrugDigits.Digit_零售价, , True)
                            '现价_In
                            gstrSQL = gstrSQL & "," & zlStr.FormatEx(Val(.TextMatrix(n, .ColIndex("现价格"))) / dbl包装, gtype_UserDrugDigits.Digit_零售价, , True)
                            '执行日期_In
                            If opt时间(0).Value = True Then
                                gstrSQL = gstrSQL & "," & "to_date('" & Format(dtToday, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            Else
                                gstrSQL = gstrSQL & "," & "to_date('" & Format(Me.dtpRunDate.Value, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            End If
                            '调价说明_In
                            gstrSQL = gstrSQL & ",'" & Me.txtSummary.Text & "'"
                            '调价人_In
                            gstrSQL = gstrSQL & ",'" & Trim(Me.txtValuer.Text) & "'"
                            '调价汇总号_In
                            gstrSQL = gstrSQL & ",'" & strRetail流水号 & "'"
                            gstrSQL = gstrSQL & ")"

                            ReDim Preserve marrSql(UBound(marrSql) + 1)
                            marrSql(UBound(marrSql)) = gstrSQL
                        End If
                    Next
                    
                    If opt时间(0).Value = True Then
                        gstrSQL = "zl_药品收发记录_Adjust(" & Val(.TextMatrix(i, .ColIndex("药品id"))) & ",1)"
                        
                        ReDim Preserve marrSql(UBound(marrSql) + 1)
                        marrSql(UBound(marrSql)) = gstrSQL
                    End If
                End If
                
                bln是否调过售价 = False
                dbl现售价 = 0
                int同药品ID数 = 0
            End If
        Next
    End With
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ModifyAllPrice()
    '成本价、售价一起调
    Dim strAll流水号 As String
    Dim rsTemp As ADODB.Recordset
    Dim dbl包装 As Double
    Dim i As Integer
    Dim n As Integer
    Dim int同药品ID数 As Integer
    Dim int收费价目序号 As Integer
    Dim bln是否调过售价 As Boolean
    Dim dbl现售价 As Double
    Dim lngAdjId As Long
    Dim dtToday As Date
    Dim strNo As String
    Dim LngCurID As Long

    On Error GoTo ErrHand
    
    With vsfPrice
        For i = 1 To .rows - 2
            If Val(.TextMatrix(i, .ColIndex("原售价"))) <> Val(.TextMatrix(i, .ColIndex("现价格"))) And Val(.TextMatrix(i, .ColIndex("原成本价"))) <> Val(.TextMatrix(i, .ColIndex("现价格"))) And .TextMatrix(i, .ColIndex("现价格")) <> "" Then
                '先处理成本价
                If strAll流水号 = "" Then
                    gstrSQL = "select nextno(135) as 流水号 from dual"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "调价流水号")
                    If rsTemp.RecordCount = 0 Then
                        MsgBox "调价流水号未能初始化成功，请与管理员联系！", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    strAll流水号 = rsTemp!流水号

                    mstr纪录流水号 = mstr纪录流水号 & ";" & strAll流水号 & "|" & 2
                    dtToday = Sys.Currentdate()
                End If
                              
                dbl包装 = Val(vsfPrice.TextMatrix(i, vsfPrice.ColIndex("包装系数")))
                
                If .TextMatrix(i, .ColIndex("有库存")) = 1 And .TextMatrix(i, .ColIndex("药价属性")) = "定价" Then
                    gstrSQL = "Select s.库房id, s.药品id, d.名称 As 库房, '[' || m.编码 || ']' || m.名称 As 药品, m.规格,s.上次产地 as 产地," & vbNewLine & _
                                    "       Decode([2], 0, p.药库单位, 2, p.住院单位, 1, p.门诊单位, m.计算单位) As 单位," & vbNewLine & _
                                    "       Decode([2], 0, p.药库包装, 2, p.住院包装, 1, p.门诊包装, 1) As 包装系数," & vbNewLine & _
                                    "       s.上次批号 As 批号, Nvl(s.实际数量, 0) As 数量, Nvl(s.批次,0) as 批次," & vbNewLine & _
                                    "       Nvl(m.是否变价, 0) 变价, m.Id, Decode(Nvl(m.是否变价, 0), 0, e.现价, Decode(Nvl(s.零售价, 0),0,s.实际金额/s.实际数量,s.零售价)) As 时价售价, p.加成率, Decode(Nvl(s.平均成本价, 0), 0, p.成本价, s.平均成本价) As 成本价, nvl(s.上次供应商id,0) As 上次供应商id," & vbNewLine & _
                                    "       n.名称 As 供应商, s.效期" & vbNewLine & _
                                    "From 药品库存 S, 部门表 D, 收费项目目录 M, 药品规格 P, 供应商 N, 收费价目 E" & vbNewLine & _
                                    "Where d.Id = s.库房id And s.药品id = m.Id And m.Id = p.药品id And Nvl(s.上次供应商id, 0) = n.Id(+) And m.Id = e.收费细目id And" & vbNewLine & _
                                    "      s.性质 = 1 And s.药品id = [1] And Sysdate Between e.执行日期 And e.终止日期   And e.价格等级 Is Null" & vbNewLine & _
                                    "Order By s.药品id,s.库房id, s.上次批号,s.批次 "

                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取定价药品信息", Val(.TextMatrix(i, .ColIndex("药品ID"))), mintUnit)
                    
                    With rsTemp
                        Do While Not .EOF
                            gstrSQL = "Zl_药品价格记录_Stop("
                            '价格类型_In
                            gstrSQL = gstrSQL & 2
                            '库房id_In
                            gstrSQL = gstrSQL & "," & !库房id
                            '药品id_In
                            gstrSQL = gstrSQL & "," & !药品id
                            '批次_In
                            gstrSQL = gstrSQL & "," & Nvl(!批次, 0)
                            '终止日期_In
                            If opt时间(0).Value = True Then
                                gstrSQL = gstrSQL & "," & "to_date('" & Format(DateAdd("s", -1, dtToday), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            Else
                                gstrSQL = gstrSQL & "," & "to_date('" & Format(DateAdd("s", -1, Me.dtpRunDate.Value), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            End If
                            gstrSQL = gstrSQL & ")"
                            
                            ReDim Preserve marrSql(UBound(marrSql) + 1)
                            marrSql(UBound(marrSql)) = gstrSQL
                            
                            gstrSQL = "Zl_药品价格记录_Insert("
                            '调价类型_In
                            gstrSQL = gstrSQL & 1
                            '价格类型_In
                            gstrSQL = gstrSQL & ",2"
                            '库房id_In
                            gstrSQL = gstrSQL & "," & !库房id
                            '药品id_In
                            gstrSQL = gstrSQL & "," & !药品id
                            '批次_In
                            gstrSQL = gstrSQL & "," & Nvl(!批次, 0)
                            '原价_In
                            gstrSQL = gstrSQL & "," & zlStr.FormatEx(!成本价, gtype_UserDrugDigits.Digit_零售价, , True)
                            '现价_In
                            gstrSQL = gstrSQL & "," & zlStr.FormatEx(Val(vsfPrice.TextMatrix(i, vsfPrice.ColIndex("现价格"))) / dbl包装, gtype_UserDrugDigits.Digit_零售价, , True)
                            '执行日期_In
                            If opt时间(0).Value = True Then
                                gstrSQL = gstrSQL & "," & "to_date('" & Format(dtToday, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            Else
                                gstrSQL = gstrSQL & "," & "to_date('" & Format(Me.dtpRunDate.Value, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            End If
                            '调价说明_In
                            gstrSQL = gstrSQL & ",'" & Me.txtSummary.Text & "'"
                            '调价人_In
                            gstrSQL = gstrSQL & ",'" & Trim(Me.txtValuer.Text) & "'"
                            '调价汇总号_In
                            gstrSQL = gstrSQL & ",'" & strAll流水号 & "'"
                            '供药单位id_In
                            gstrSQL = gstrSQL & "," & IIf(!上次供应商ID = 0, "null", !上次供应商ID)
                            '批号_In
                            gstrSQL = gstrSQL & ",'" & Nvl(!批号) & "'"
                            '效期_In
                            gstrSQL = gstrSQL & "," & "to_date('" & Format(IIf(IsNull(!效期), "", !效期), "YYYY-MM-DD") & "','yyyy-mm-dd')"
                            '产地_In
                            gstrSQL = gstrSQL & ",'" & Nvl(!产地) & "'"
                            gstrSQL = gstrSQL & ")"
                            
                            ReDim Preserve marrSql(UBound(marrSql) + 1)
                            marrSql(UBound(marrSql)) = gstrSQL
                            
                        .MoveNext
                        Loop
                    End With
                End If
            
                If .TextMatrix(i, .ColIndex("有库存")) = 1 And .TextMatrix(i, .ColIndex("药价属性")) = "时价" Then
                    '有库存
                    gstrSQL = "Zl_药品价格记录_Stop("
                    '价格类型_In
                    gstrSQL = gstrSQL & 2
                    '库房id_In
                    gstrSQL = gstrSQL & "," & Val(.TextMatrix(i, .ColIndex("库房id")))
                    '药品id_In
                    gstrSQL = gstrSQL & "," & Val(.TextMatrix(i, .ColIndex("药品id")))
                    '批次_In
                    gstrSQL = gstrSQL & "," & Nvl(.TextMatrix(i, .ColIndex("批次")), 0)
                    '终止日期_In
                    If opt时间(0).Value = True Then
                        gstrSQL = gstrSQL & "," & "to_date('" & Format(DateAdd("s", -1, dtToday), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    Else
                        gstrSQL = gstrSQL & "," & "to_date('" & Format(DateAdd("s", -1, Me.dtpRunDate.Value), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    End If
                    gstrSQL = gstrSQL & ")"
                    
                    ReDim Preserve marrSql(UBound(marrSql) + 1)
                    marrSql(UBound(marrSql)) = gstrSQL
                    
                    gstrSQL = "Zl_药品价格记录_Insert("
                    '调价类型_In
                    gstrSQL = gstrSQL & 1
                    '价格类型_In
                    gstrSQL = gstrSQL & ",2"
                    '库房id_In
                    gstrSQL = gstrSQL & "," & Val(.TextMatrix(i, .ColIndex("库房id")))
                    '药品id_In
                    gstrSQL = gstrSQL & "," & Val(.TextMatrix(i, .ColIndex("药品id")))
                    '批次_In
                    gstrSQL = gstrSQL & "," & Nvl(.TextMatrix(i, .ColIndex("批次")), 0)
                    '原价_In
                    gstrSQL = gstrSQL & "," & zlStr.FormatEx(Val(.TextMatrix(i, .ColIndex("原成本价"))) / dbl包装, gtype_UserDrugDigits.Digit_零售价, , True)
                    '现价_In
                    gstrSQL = gstrSQL & "," & zlStr.FormatEx(Val(.TextMatrix(i, .ColIndex("现价格"))) / dbl包装, gtype_UserDrugDigits.Digit_零售价, , True)
                    '执行日期_In
                    If opt时间(0).Value = True Then
                        gstrSQL = gstrSQL & "," & "to_date('" & Format(dtToday, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    Else
                        gstrSQL = gstrSQL & "," & "to_date('" & Format(Me.dtpRunDate.Value, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    End If
                    '调价说明_In
                    gstrSQL = gstrSQL & ",'" & Me.txtSummary.Text & "'"
                    '调价人_In
                    gstrSQL = gstrSQL & ",'" & Trim(Me.txtValuer.Text) & "'"
                    '调价汇总号_In
                    gstrSQL = gstrSQL & ",'" & strAll流水号 & "'"
                    '供药单位id_In
                    gstrSQL = gstrSQL & "," & IIf(Val(.TextMatrix(i, .ColIndex("上次供应商ID"))) = 0, "null", Val(.TextMatrix(i, .ColIndex("上次供应商ID"))))
                    '批号_In
                    gstrSQL = gstrSQL & ",'" & Nvl(.TextMatrix(i, .ColIndex("批号"))) & "'"
                    '效期_In
                    gstrSQL = gstrSQL & "," & "to_date('" & Format(IIf(IsNull(.TextMatrix(i, .ColIndex("效期"))), "", .TextMatrix(i, .ColIndex("效期"))), "YYYY-MM-DD") & "','yyyy-mm-dd')"
                    '产地_In
                    gstrSQL = gstrSQL & ",'" & Nvl(.TextMatrix(i, .ColIndex("生产商"))) & "'"
                    gstrSQL = gstrSQL & ")"
                    
                    ReDim Preserve marrSql(UBound(marrSql) + 1)
                    marrSql(UBound(marrSql)) = gstrSQL
                End If
                
                If .TextMatrix(i, .ColIndex("有库存")) = 0 Then
                    '无库存
                    gstrSQL = "Zl_药品价格记录_Insert("
                    '调价类型_In
                    gstrSQL = gstrSQL & 1
                    '价格类型_In
                    gstrSQL = gstrSQL & ",2"
                    '库房id_In
                    gstrSQL = gstrSQL & ",Null"
                    '药品id_In
                    gstrSQL = gstrSQL & "," & Val(.TextMatrix(i, .ColIndex("药品id")))
                    '批次_In
                    gstrSQL = gstrSQL & ",0"
                    '原价_In
                    gstrSQL = gstrSQL & "," & zlStr.FormatEx(Val(.TextMatrix(i, .ColIndex("原成本价"))) / dbl包装, gtype_UserDrugDigits.Digit_零售价, , True)
                    '现价_In
                    gstrSQL = gstrSQL & "," & zlStr.FormatEx(Val(.TextMatrix(i, .ColIndex("现价格"))) / dbl包装, gtype_UserDrugDigits.Digit_零售价, , True)
                    '执行日期_In
                    If opt时间(0).Value = True Then
                        gstrSQL = gstrSQL & "," & "to_date('" & Format(dtToday, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    Else
                        gstrSQL = gstrSQL & "," & "to_date('" & Format(Me.dtpRunDate.Value, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    End If
                    '调价说明_In
                    gstrSQL = gstrSQL & ",'" & Me.txtSummary.Text & "'"
                    '调价人_In
                    gstrSQL = gstrSQL & ",'" & Trim(Me.txtValuer.Text) & "'"
                    '调价汇总号_In
                    gstrSQL = gstrSQL & ",'" & strAll流水号 & "'"
                    gstrSQL = gstrSQL & ")"

                    ReDim Preserve marrSql(UBound(marrSql) + 1)
                    marrSql(UBound(marrSql)) = gstrSQL
                End If
            End If
        Next

        For i = 1 To .rows - 2
            '再处理售价
            If Val(.TextMatrix(i, .ColIndex("药品id"))) = Val(.TextMatrix(i + 1, .ColIndex("药品id"))) Then
                int同药品ID数 = int同药品ID数 + 1
            Else
                For n = i - int同药品ID数 To i
                    dbl现售价 = dbl现售价 + Val(.TextMatrix(n, .ColIndex("现价格")))
                    If Val(.TextMatrix(n, .ColIndex("原售价"))) <> Val(.TextMatrix(n, .ColIndex("现价格"))) And Val(.TextMatrix(n, .ColIndex("原成本价"))) <> Val(.TextMatrix(n, .ColIndex("现价格"))) And .TextMatrix(n, .ColIndex("现价格")) <> "" Then
                        bln是否调过售价 = True
                    End If
                Next

                If bln是否调过售价 Then
                    If lngAdjId = 0 Then
                        strNo = Sys.GetNextNo(9)

                        gstrSQL = "select 收费价目_ID.nextval from dual"
                        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取收费价目序号")
                        lngAdjId = rsTemp.Fields(0).Value

                        dtToday = Sys.Currentdate()
                    End If

                    int收费价目序号 = int收费价目序号 + 1
                    dbl现售价 = Round(dbl现售价 / (int同药品ID数 + 1), 2)
                    LngCurID = Sys.NextId("收费价目")
                    dbl包装 = Val(.TextMatrix(i, .ColIndex("包装系数")))
            
                    If CLng(.TextMatrix(i, .ColIndex("价格ID"))) <> 0 Then
                        '设置上一次的价格记录终止执行
                        gstrSQL = "zl_收费价目_stop(" & .TextMatrix(i, .ColIndex("药品id")) & ","
                        If opt时间(0).Value = True Then
                            gstrSQL = gstrSQL & "to_date('" & Format(DateAdd("s", -1, dtToday), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                        Else
                            gstrSQL = gstrSQL & "to_date('" & Format(DateAdd("s", -1, Me.dtpRunDate.Value), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                        End If
                        gstrSQL = gstrSQL & ")"
                        
                        ReDim Preserve marrSql(UBound(marrSql) + 1)
                        marrSql(UBound(marrSql)) = gstrSQL
            
                        '产生价格记录
                        'ID
                        gstrSQL = "zl_收费价目_Insert(" & LngCurID & ","
                        '原价ID
                        gstrSQL = gstrSQL & Val(.TextMatrix(i, .ColIndex("价格ID"))) & ","
                        '收费细目ID
                        gstrSQL = gstrSQL & Val(.TextMatrix(i, .ColIndex("药品id"))) & ","
                        '收入项目ID
                        gstrSQL = gstrSQL & Val(.TextMatrix(i, .ColIndex("收入项目ID"))) & ","
                        '原价
                        gstrSQL = gstrSQL & zlStr.FormatEx(Val(.TextMatrix(i, .ColIndex("原售价"))) / dbl包装, gtype_UserDrugDigits.Digit_零售价, , True) & ","
                        '现价
                        gstrSQL = gstrSQL & zlStr.FormatEx(dbl现售价 / dbl包装, gtype_UserDrugDigits.Digit_零售价, , True) & ","
                        '附术收费率、加班加价率、调价说明
                        gstrSQL = gstrSQL & "NULL,NULL,'" & MoveSpecialChar(txtSummary.Text) & "',"
                        '调价id、调价人
                        gstrSQL = gstrSQL & lngAdjId & ",'" & MoveSpecialChar(txtValuer.Text) & "',"
                        '执行日期
                        If opt时间(0).Value = True Then
                            gstrSQL = gstrSQL & "to_date('" & Format(dtToday, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
                        Else
                            gstrSQL = gstrSQL & "to_date('" & Format(Me.dtpRunDate.Value, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
                        End If
                        '变动原因
                        gstrSQL = gstrSQL & "0,"
                        'NO
                        gstrSQL = gstrSQL & "'" & strNo & "',"
                        '序号、缺省价格
                        gstrSQL = gstrSQL & int收费价目序号 & ",Null,"
                        '调价汇总号
                        gstrSQL = gstrSQL & strAll流水号 & ")"
                        
                        ReDim Preserve marrSql(UBound(marrSql) + 1)
                        marrSql(UBound(marrSql)) = gstrSQL
                    End If
                
                    For n = i - int同药品ID数 To i
                        If Val(.TextMatrix(n, .ColIndex("原售价"))) <> Val(.TextMatrix(n, .ColIndex("现价格"))) And Val(.TextMatrix(n, .ColIndex("原成本价"))) <> Val(.TextMatrix(n, .ColIndex("现价格"))) And .TextMatrix(n, .ColIndex("现价格")) <> "" And .TextMatrix(n, .ColIndex("药价属性")) = "时价" And .TextMatrix(n, .ColIndex("有库存")) = 1 Then
                            gstrSQL = "Zl_药品价格记录_Stop("
                            '价格类型_In
                            gstrSQL = gstrSQL & 1
                            '库房id_In
                            gstrSQL = gstrSQL & "," & Val(.TextMatrix(n, .ColIndex("库房ID")))
                            '药品id_In
                            gstrSQL = gstrSQL & "," & Val(.TextMatrix(n, .ColIndex("药品ID")))
                            '批次_In
                            gstrSQL = gstrSQL & "," & Val(.TextMatrix(n, .ColIndex("批次")))
                            '终止日期_In
                            If opt时间(0).Value = True Then
                                gstrSQL = gstrSQL & "," & "to_date('" & Format(DateAdd("s", -1, dtToday), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            Else
                                gstrSQL = gstrSQL & "," & "to_date('" & Format(DateAdd("s", -1, Me.dtpRunDate.Value), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            End If
                            gstrSQL = gstrSQL & ")"
                            
                            ReDim Preserve marrSql(UBound(marrSql) + 1)
                            marrSql(UBound(marrSql)) = gstrSQL
                            
                            gstrSQL = "Zl_药品价格记录_Insert("
                            '调价类型_In
                            gstrSQL = gstrSQL & 1
                            '价格类型_In
                            gstrSQL = gstrSQL & ",1"
                            '库房id_In
                            gstrSQL = gstrSQL & "," & Val(.TextMatrix(n, .ColIndex("库房ID")))
                            '药品id_In
                            gstrSQL = gstrSQL & "," & Val(.TextMatrix(n, .ColIndex("药品ID")))
                            '批次_In
                            gstrSQL = gstrSQL & "," & Val(.TextMatrix(n, .ColIndex("批次")))
                            '原价_In
                            gstrSQL = gstrSQL & "," & zlStr.FormatEx(Val(.TextMatrix(n, .ColIndex("原售价"))) / dbl包装, gtype_UserDrugDigits.Digit_零售价, , True)
                            '现价_In
                            gstrSQL = gstrSQL & "," & zlStr.FormatEx(Val(.TextMatrix(n, .ColIndex("现价格"))) / dbl包装, gtype_UserDrugDigits.Digit_零售价, , True)
                            '执行日期_In
                            If opt时间(0).Value = True Then
                                gstrSQL = gstrSQL & "," & "to_date('" & Format(dtToday, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            Else
                                gstrSQL = gstrSQL & "," & "to_date('" & Format(Me.dtpRunDate.Value, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            End If
                            '调价说明_In
                            gstrSQL = gstrSQL & ",'" & Me.txtSummary.Text & "'"
                            '调价人_In
                            gstrSQL = gstrSQL & ",'" & Trim(Me.txtValuer.Text) & "'"
                            '调价汇总号_In
                            gstrSQL = gstrSQL & ",'" & strAll流水号 & "'"
                            '供药单位id_In
                            gstrSQL = gstrSQL & "," & IIf(Val(.TextMatrix(n, .ColIndex("上次供应商ID"))) = 0, "null", Val(.TextMatrix(n, .ColIndex("上次供应商ID"))))
                            '批号_In
                            gstrSQL = gstrSQL & ",'" & Nvl(.TextMatrix(n, .ColIndex("批号"))) & "'"
                            '效期_In
                            gstrSQL = gstrSQL & "," & "to_date('" & Format(Nvl(.TextMatrix(n, .ColIndex("效期"))), "YYYY-MM-DD") & "','yyyy-mm-dd')"
                            '产地_In
                            gstrSQL = gstrSQL & ",'" & Nvl(.TextMatrix(n, .ColIndex("生产商"))) & "'"
                            gstrSQL = gstrSQL & ")"
                            
                            ReDim Preserve marrSql(UBound(marrSql) + 1)
                            marrSql(UBound(marrSql)) = gstrSQL
                        End If
                        
                        If Val(.TextMatrix(n, .ColIndex("原售价"))) <> Val(.TextMatrix(n, .ColIndex("现价格"))) And Val(.TextMatrix(n, .ColIndex("原成本价"))) <> Val(.TextMatrix(n, .ColIndex("现价格"))) And .TextMatrix(n, .ColIndex("现价格")) <> "" And .TextMatrix(n, .ColIndex("药价属性")) = "时价" And .TextMatrix(n, .ColIndex("有库存")) = 0 Then
                            '时价药品无库存调价
                            gstrSQL = "Zl_药品价格记录_Insert("
                            '调价类型_In
                            gstrSQL = gstrSQL & 1
                            '价格类型_In
                            gstrSQL = gstrSQL & ",1"
                            '库房id_In
                            gstrSQL = gstrSQL & ",Null"
                            '药品id_In
                            gstrSQL = gstrSQL & "," & Val(.TextMatrix(n, .ColIndex("药品ID")))
                            '批次_In
                            gstrSQL = gstrSQL & ",0"
                            '原价_In
                            gstrSQL = gstrSQL & "," & zlStr.FormatEx(Val(.TextMatrix(n, .ColIndex("原售价"))) / dbl包装, gtype_UserDrugDigits.Digit_零售价, , True)
                            '现价_In
                            gstrSQL = gstrSQL & "," & zlStr.FormatEx(Val(.TextMatrix(n, .ColIndex("现价格"))) / dbl包装, gtype_UserDrugDigits.Digit_零售价, , True)
                            '执行日期_In
                            If opt时间(0).Value = True Then
                                gstrSQL = gstrSQL & "," & "to_date('" & Format(dtToday, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            Else
                                gstrSQL = gstrSQL & "," & "to_date('" & Format(Me.dtpRunDate.Value, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            End If
                            '调价说明_In
                            gstrSQL = gstrSQL & ",'" & Me.txtSummary.Text & "'"
                            '调价人_In
                            gstrSQL = gstrSQL & ",'" & Trim(Me.txtValuer.Text) & "'"
                            '调价汇总号_In
                            gstrSQL = gstrSQL & ",'" & strAll流水号 & "'"
                            gstrSQL = gstrSQL & ")"

                            ReDim Preserve marrSql(UBound(marrSql) + 1)
                            marrSql(UBound(marrSql)) = gstrSQL
                        End If
                    Next
                    
                    If opt时间(0).Value = True Then
                        gstrSQL = "zl_药品收发记录_Adjust(" & Val(.TextMatrix(i, .ColIndex("药品id"))) & ")"
                        
                        ReDim Preserve marrSql(UBound(marrSql) + 1)
                        marrSql(UBound(marrSql)) = gstrSQL
                    End If
                End If
                bln是否调过售价 = False
                dbl现售价 = 0
                int同药品ID数 = 0
            End If
        Next
    End With
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub setNOtExcetePrice()
    '将到时间还未执行调价药品执行调价
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    Dim blnTrans As Boolean
    Dim arrSql() As Variant
    
    arrSql = Array()
    On Error GoTo errHandle
    
    gstrSQL = "Select Distinct i.Id As 药品id " & _
               " From 收费项目目录 I, 收费价目 N, 药品规格 P" & _
               " Where i.Id = n.收费细目id And i.Id = p.药品id And (i.撤档时间 Is Null Or i.撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd')) And" & _
                   " n.变动原因 = 0 And Sysdate>n.执行日期" & GetPriceClassString("N") & _
               " Union " & _
               " Select Distinct a.药品id From 药品价格记录 A Where a.记录状态 = 0 And a.执行日期 <= Sysdate " & _
               " Order By 药品id "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "执行调价")
    
    If rsTemp.RecordCount = 0 Then Exit Sub
    
    For i = 0 To rsTemp.RecordCount - 1
        gstrSQL = "Zl_药品收发记录_Adjust(" & rsTemp!药品id & ")"
        
        ReDim Preserve arrSql(UBound(arrSql) + 1)
        arrSql(UBound(arrSql)) = gstrSQL
        rsTemp.MoveNext
    Next
                   
    gcnOracle.BeginTrans: blnTrans = True          '开启事务
    For i = 0 To UBound(arrSql)
        Call zlDatabase.ExecuteProcedure(CStr(arrSql(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTrans = False     '提交事物
    
    Exit Sub
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub PrintPrice()
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    
    Call Summary
    
    If vsfPrint.rows = 1 Then
        MsgBox "没有库存变动记录！", vbInformation, gstrSysName
        Exit Sub
    End If

    objPrint.Title.Text = "调价库存变动表"

    Set objRow = New zlTabAppRow
    objRow.Add "调价说明:" & Me.txtSummary.Text
    objPrint.UnderAppRows.Add objRow

    Set objRow = New zlTabAppRow
    objRow.Add "执行时间:" & Format(IIf(opt时间(0).Value = True, Sys.Currentdate, Me.dtpRunDate.Value), "yyyy年MM月DD日 HH:mm:ss")
    objRow.Add "调价人:" & Me.txtValuer.Text
    objPrint.UnderAppRows.Add objRow

    Set objRow = New zlTabAppRow
    objRow.Add "打印人:" & UserInfo.用户姓名
    objRow.Add "打印时间:" & Format(Sys.Currentdate, "yyyy年MM月DD日 HH:mm:ss")
    objPrint.BelowAppRows.Add objRow

    Set objPrint.Body = Me.vsfPrint.Object
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

Private Sub Summary()
    '汇总库存变动表
    Dim i As Integer
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    vsfPrint.rows = 1
    
    For i = 1 To vsfPrice.rows - 2
        If vsfPrice.TextMatrix(i, vsfPrice.ColIndex("有库存")) = 1 And vsfPrice.TextMatrix(i, vsfPrice.ColIndex("药价属性")) = "时价" And vsfPrice.TextMatrix(i, vsfPrice.ColIndex("现价格")) <> "" Then
            vsfPrint.rows = vsfPrint.rows + 1
            vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("药价属性")) = vsfPrice.TextMatrix(i, vsfPrice.ColIndex("药价属性"))
            vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("库房")) = vsfPrice.TextMatrix(i, vsfPrice.ColIndex("库房"))
            vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("品名")) = vsfPrice.TextMatrix(i, vsfPrice.ColIndex("品名"))
            vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("规格")) = vsfPrice.TextMatrix(i, vsfPrice.ColIndex("规格"))
            vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("生产商")) = vsfPrice.TextMatrix(i, vsfPrice.ColIndex("生产商"))
            vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("批号")) = vsfPrice.TextMatrix(i, vsfPrice.ColIndex("批号"))
            vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("单位")) = vsfPrice.TextMatrix(i, vsfPrice.ColIndex("单位"))
            vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("原售价")) = vsfPrice.TextMatrix(i, vsfPrice.ColIndex("原售价"))
            vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("原成本价")) = vsfPrice.TextMatrix(i, vsfPrice.ColIndex("原成本价"))
            vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("现售价")) = vsfPrice.TextMatrix(i, vsfPrice.ColIndex("现价格"))
            vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("现成本价")) = vsfPrice.TextMatrix(i, vsfPrice.ColIndex("现价格"))
        End If

        If vsfPrice.TextMatrix(i, vsfPrice.ColIndex("有库存")) = 1 And vsfPrice.TextMatrix(i, vsfPrice.ColIndex("药价属性")) = "定价" And vsfPrice.TextMatrix(i, vsfPrice.ColIndex("现价格")) <> "" Then
            
            gstrSQL = "Select s.库房id, s.药品id, d.名称 As 库房, '[' || m.编码 || ']' || m.名称 As 药品, m.规格,s.上次产地 as 产地," & vbNewLine & _
                            "       Decode([2], 0, p.药库单位, 2, p.住院单位, 1, p.门诊单位, m.计算单位) As 单位," & vbNewLine & _
                            "       Decode([2], 0, p.药库包装, 2, p.住院包装, 1, p.门诊包装, 1) As 包装系数," & vbNewLine & _
                            "       s.上次批号 As 批号, Nvl(s.实际数量, 0) As 数量, s.批次," & vbNewLine & _
                            "       Nvl(m.是否变价, 0) 变价, m.Id, Decode(Nvl(m.是否变价, 0), 0, e.现价, Decode(Nvl(s.零售价, 0),0,s.实际金额/s.实际数量,s.零售价)) As 时价售价, p.加成率, Decode(Nvl(s.平均成本价, 0), 0, p.成本价, s.平均成本价) As 成本价, nvl(s.上次供应商id,0) As 上次供应商id," & vbNewLine & _
                            "       n.名称 As 供应商, s.效期" & vbNewLine & _
                            "From 药品库存 S, 部门表 D, 收费项目目录 M, 药品规格 P, 供应商 N, 收费价目 E" & vbNewLine & _
                            "Where d.Id = s.库房id And s.药品id = m.Id And m.Id = p.药品id And Nvl(s.上次供应商id, 0) = n.Id(+) And m.Id = e.收费细目id And" & vbNewLine & _
                            "      s.性质 = 1 And s.药品id = [1] And Sysdate Between e.执行日期 And e.终止日期 And e.价格等级 Is Null" & vbNewLine & _
                            "Order By s.药品id,s.库房id, s.上次批号,s.批次 "

                            
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取定价药品信息", Val(vsfPrice.TextMatrix(i, vsfPrice.ColIndex("药品ID"))), mintUnit)
            
            With rsTemp
                Do While Not .EOF
                    vsfPrint.rows = vsfPrint.rows + 1
                    vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("药价属性")) = "定价"
                    vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("库房")) = !库房
                    vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("品名")) = !药品
                    vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("规格")) = !规格
                    vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("生产商")) = Nvl(!产地)
                    vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("批号")) = Nvl(!批号)
                    vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("单位")) = !单位
                    vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("原售价")) = zlStr.FormatEx(!时价售价 * !包装系数, mintPriceDigit, , True)
                    vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("原成本价")) = zlStr.FormatEx(!成本价 * !包装系数, mintCostDigit, , True)
                    vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("现售价")) = vsfPrice.TextMatrix(i, vsfPrice.ColIndex("现价格"))
                    vsfPrint.TextMatrix(vsfPrint.rows - 1, vsfPrint.ColIndex("现成本价")) = vsfPrice.TextMatrix(i, vsfPrice.ColIndex("现价格"))
                    .MoveNext
                Loop
            End With
        End If
    Next
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ClearAllDate()
    '清空界面数据
    Dim i As Integer
    If vsfPrice.rows <= 1 Then Exit Sub
    If MsgBox("未保存本次调价，是否清空所有数据？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    mstr药品ID = ""
    vsfPrice.rows = 1
End Sub

Private Sub AllDrug()
    Dim intRow As Integer
    Dim rsReturn As ADODB.Recordset
    Dim blnOK As Boolean
    
    frmBatchSelect.ShowME Me, rsReturn, blnOK, 1

    On Error GoTo errHandle
    If blnOK = False Then Exit Sub
    If rsReturn.RecordCount = 0 Then Exit Sub
    
'    If txtValuer.Tag = "部分零差价药品" And vsfPrice.rows > 1 Then
'         If MsgBox("该操作会清除空面数据并批量提取已选择的零差价管理药品，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
'    End If
    
    If mstr药品ID <> "" Then
        Select Case MsgBox("是否清除界面数据？", vbQuestion + vbYesNoCancel + vbDefaultButton2, gstrSysName)
            Case vbYes
                mstr药品ID = ""
            Case vbCancel
                Exit Sub
        End Select
    End If
        
    rsReturn.MoveFirst
    Do While Not rsReturn.EOF
        If InStr(mstr药品ID & ",", "," & rsReturn!药品id & ",") = 0 Then
            mstr药品ID = mstr药品ID & "," & rsReturn!药品id
        End If
        rsReturn.MoveNext
    Loop

    Call GetAllPriceDiff
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetAllPriceDiff(Optional bln提示 As Boolean = True)
    '提取所有已设置了零差价管理的药品包括价格一致的药品
    Dim rsData As ADODB.Recordset
    Dim bln存在未执行价格 As Boolean
    Dim int序号 As Integer
    
    On Error GoTo errHandle
       
    Call setNOtExcetePrice
    
    gstrSQL = "select * from (" & vbNewLine & _
                    "Select 药品id, 通用名, 规格, 0 As 库房id, '' As 库房, 生产商, '' As 批号, 批次, 单位, 包装系数, 售价, Sum(成本价 * 实际数量) / Sum(实际数量) As 成本价, 是否时价," & vbNewLine & _
                    "       有库存, 价格id, 收入项目id, Null As 上次供应商id, Null As 效期" & vbNewLine & _
                    "From (Select a.药名id,a.药品id, '[' || c.编码 || ']' || c.名称 As 通用名, c.规格, c.产地 As 生产商, 0 As 批次," & vbNewLine & _
                    "              Decode([1], 0, a.药库单位, 2, a.住院单位, 1, a.门诊单位, c.计算单位) As 单位," & vbNewLine & _
                    "              Decode([1], 0, a.药库包装, 2, a.住院包装, 1, a.门诊包装, 1) As 包装系数, b.现价 As 售价, decode(d.平均成本价,null,a.成本价,d.平均成本价) As 成本价, 0 As 是否时价, d.实际数量," & vbNewLine & _
                    "              1 As 有库存, b.Id As 价格id, b.收入项目id" & vbNewLine & _
                    "       From 药品规格 A, 收费价目 B, 收费项目目录 C, 药品库存 D" & vbNewLine & _
                    "       Where a.药品id = b.收费细目id And a.药品id = c.Id And a.药品id = d.药品id And d.性质 = 1 And (Sysdate Between b.执行日期 And b.终止日期) And" & vbNewLine & _
                    "             (c.撤档时间 Is Null Or c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And c.是否变价 = 0 And Nvl(a.是否零差价管理, 0) = 1 " & vbNewLine & _
                    " And Not (zl_fun_getbatchpro(d.库房id,d.药品id)=1 And Nvl(d.批次,0) = 0 And d.可用数量 < 0 And d.实际数量 = 0 And d.实际金额 = 0 And d.实际差价 = 0)) " & vbNewLine & _
                    "Group By 药品id, 通用名, 规格, 生产商, 批次, 单位, 包装系数, 售价, 价格id, 收入项目id, 是否时价, 有库存 " & vbNewLine & _
                    "Union All "

    gstrSQL = gstrSQL & "Select a.药品id, '[' || c.编码 || ']' || c.名称 As 通用名, c.规格, d.库房id, e.名称 As 库房, d.上次产地 As 生产商, d.上次批号 As 批号, d.批次," & vbNewLine & _
                    "       Decode([1], 0, a.药库单位, 2, a.住院单位, 1, a.门诊单位, c.计算单位) As 单位," & vbNewLine & _
                    "       Decode([1], 0, a.药库包装, 2, a.住院包装, 1, a.门诊包装, 1) As 包装系数, d.零售价 As 售价, decode(d.平均成本价,null,a.成本价,d.平均成本价) As 成本价, 1 As 是否时价, 1 As 有库存," & vbNewLine & _
                    "       b.Id As 价格id, b.收入项目id, nvl(d.上次供应商id,0) As 上次供应商id, d.效期" & vbNewLine & _
                    "From 药品规格 A, 收费价目 B, 收费项目目录 C, 药品库存 D, 部门表 E" & vbNewLine & _
                    "Where a.药品id = b.收费细目id And a.药品id = c.Id And a.药品id = d.药品id And d.库房id = e.Id And d.性质 = 1 And" & vbNewLine & _
                    "      (Sysdate Between b.执行日期 And b.终止日期) And c.是否变价 = 1 And" & vbNewLine & _
                    "      (c.撤档时间 Is Null Or c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And Nvl(a.是否零差价管理, 0) = 1  " & vbNewLine & _
                    " And Not (zl_fun_getbatchpro(d.库房id,d.药品id)=1 And Nvl(d.批次,0) = 0 And d.可用数量 < 0 And d.实际数量 = 0 And d.实际金额 = 0 And d.实际差价 = 0) " & vbNewLine & _
                    "Union All" & vbNewLine & _
                    "Select a.药品id, '[' || c.编码 || ']' || c.名称 As 通用名, c.规格, 0 As 库房id, '' As 库房, '' As 生产商, '' As 批号, 0 As 批次," & vbNewLine & _
                    "       Decode([1], 0, a.药库单位, 2, a.住院单位, 1, a.门诊单位, c.计算单位) As 单位," & vbNewLine & _
                    "       Decode([1], 0, a.药库包装, 2, a.住院包装, 1, a.门诊包装, 1) As 包装系数, b.现价 As 售价, a.成本价, c.是否变价 As 是否时价, 0 As 有库存, b.Id As 价格id," & vbNewLine & _
                    "       b.收入项目id, Null As 上次供应商id, Null As 效期" & vbNewLine & _
                    "From 药品规格 A, 收费价目 B, 收费项目目录 C" & vbNewLine & _
                    "Where a.药品id = c.Id And a.药品id = b.收费细目id And (c.撤档时间 Is Null Or c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And" & vbNewLine & _
                    "      Nvl(a.是否零差价管理, 0) = 1 And (Sysdate Between b.执行日期 And b.终止日期) And Not Exists" & vbNewLine & _
                    " (Select 1 From 药品库存 D Where d.药品id = a.药品id And d.性质 = 1)" & vbNewLine & _
                    "Order By 药品id, 库房id, 批号,批次) m" & vbNewLine & _
                    "where m.药品id In (Select Column_Value From Table(f_num2list([2]))) Order By 药品id, 库房id, 批号,批次 "

    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "GetAllPriceDiff", mintUnit, Mid(mstr药品ID, 2))
    
    With vsfPrice
        .Redraw = flexRDNone
        .MergeCells = flexMergeNever
        .rows = 1
    
        Do While Not rsData.EOF
            '检查是否存在未执行价格，如果存在就不取数据
            If CheckExistExecutePrice(Val(rsData!药品id)) = False Then
                .rows = .rows + 1
                
                .TextMatrix(.rows - 1, .ColIndex("序号")) = int序号 + 1
                .TextMatrix(.rows - 1, .ColIndex("药品id")) = rsData!药品id
                .TextMatrix(.rows - 1, .ColIndex("药价属性")) = IIf(rsData!是否时价 = 1, "时价", "定价")
                .TextMatrix(.rows - 1, .ColIndex("品名")) = rsData!通用名
                .TextMatrix(.rows - 1, .ColIndex("规格")) = rsData!规格
                .TextMatrix(.rows - 1, .ColIndex("生产商")) = Nvl(rsData!生产商, "")
                .TextMatrix(.rows - 1, .ColIndex("库房id")) = rsData!库房id
                .TextMatrix(.rows - 1, .ColIndex("库房")) = Nvl(rsData!库房, "")
                .TextMatrix(.rows - 1, .ColIndex("批号")) = Nvl(rsData!批号, "")
                .TextMatrix(.rows - 1, .ColIndex("单位")) = rsData!单位
                .TextMatrix(.rows - 1, .ColIndex("包装系数")) = rsData!包装系数
                .TextMatrix(.rows - 1, .ColIndex("原售价")) = zlStr.FormatEx(rsData!售价 * rsData!包装系数, mintPriceDigit, , True)
                .TextMatrix(.rows - 1, .ColIndex("原成本价")) = zlStr.FormatEx(rsData!成本价 * rsData!包装系数, mintCostDigit, , True)
                .TextMatrix(.rows - 1, .ColIndex("原始售价")) = rsData!售价
                .TextMatrix(.rows - 1, .ColIndex("原始成本价")) = rsData!成本价
                .TextMatrix(.rows - 1, .ColIndex("有库存")) = rsData!有库存
                .TextMatrix(.rows - 1, .ColIndex("价格id")) = rsData!价格id
                .TextMatrix(.rows - 1, .ColIndex("收入项目id")) = rsData!收入项目ID
                .TextMatrix(.rows - 1, .ColIndex("批次")) = Nvl(rsData!批次, 0)
                .TextMatrix(.rows - 1, .ColIndex("上次供应商ID")) = Nvl(rsData!上次供应商ID)
                .TextMatrix(.rows - 1, .ColIndex("效期")) = Nvl(rsData!效期)
                
                .Cell(flexcpForeColor, .rows - 1, .ColIndex("药价属性"), .rows - 1, .ColIndex("药价属性")) = IIf(rsData!是否时价 = 1, vbRed, vbBlack)
                int序号 = int序号 + 1
            Else
                bln存在未执行价格 = True
            End If
            
            rsData.MoveNext
        Loop
            
        If .rows >= 2 Then
            .Cell(flexcpBackColor, 1, .ColIndex("现价格"), .rows - 1, .ColIndex("现价格")) = mconlngCanColColor
            .Cell(flexcpForeColor, 1, .ColIndex("现价格"), .rows - 1, .ColIndex("现价格")) = vbBlue
            .Cell(flexcpFontBold, 1, .ColIndex("现价格"), .rows - 1, .ColIndex("现价格")) = True
        End If
        
        .rows = .rows + 1
        .RowHidden(.rows - 1) = True
        .Redraw = flexRDDirect
    End With
    
    txtValuer.Text = UserInfo.用户姓名
    txtSummary.Text = "零差价调价"
    
    txtValuer.Tag = "全部零差价药品"
    If bln存在未执行价格 = True Then
        If bln提示 Then
            MsgBox "部分零差价管理药品还存在未执行的预调价记录，本次零差价批量调价列表不在显示这些药品，请注意查看！", vbInformation, gstrSysName
        Else
            MsgBox "本次零差价管理药品批量调价中有部分药品未进行调价，请注意查看！", vbInformation, gstrSysName
        End If
    End If
        
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FindLocation()
    Dim i As Integer
    
    With vsfPrice
        If .rows > 1 Then
            For i = mintRow To .rows - 2
                If .TextMatrix(mintRow, .ColIndex("现价格")) = "" Then
                    .TopRow = mintRow
                    .Row = mintRow
                    .Col = .ColIndex("现价格")
                    mintRow = mintRow + 1
                    Exit For
                Else
                    mintRow = mintRow + 1
                End If
            Next

            If mintRow = .rows - 1 And .TextMatrix(mintRow - 1, .ColIndex("现价格")) <> "" Then
                mintRow = 1
                For i = mintRow To .rows - 2
                    If .TextMatrix(mintRow, .ColIndex("现价格")) = "" Then
                        .TopRow = mintRow
                        .Row = mintRow
                        .Col = .ColIndex("现价格")
                        mintRow = mintRow + 1
                        Exit For
                    Else
                        mintRow = mintRow + 1
                    End If
                Next
            End If
            
            If mintRow = .rows - 1 Then
                mintRow = 1
            End If
        End If
    End With
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
    If KeyAscii = 13 And Trim(txtCode.Text) <> "" Then
        Call FindGridRow(txtCode.Text)
    End If
End Sub

Private Sub FindGridRow(ByVal strInput As String)
    Dim n As Integer
    Dim lngFindRow As Long
    Dim str药名 As String
    Dim lngRow As Long
    
    '查找药品
    On Error GoTo errHandle
    If strInput <> txtCode.Tag Then
        '表示新的查找
        txtCode.Tag = strInput
        
        gstrSQL = "Select Distinct A.Id,'[' || A.编码 || ']' As 药品编码, A.名称 As 通用名, B.名称 As 商品名 " & _
                  "From 收费项目目录 A,收费项目别名 B " & _
                  "Where (A.站点 = [3] Or A.站点 is Null) And A.Id =B.收费细目id And A.类别 In ('5','6','7') " & _
                  "  And (A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2] ) " & _
                  "Order By 药品编码 "
        Set mrsFindName = zlDatabase.OpenSQLRecord(gstrSQL, "取匹配的药品ID", strInput & "%", "%" & strInput & "%", gstrNodeNo)
        
        If mrsFindName.RecordCount = 0 Then Exit Sub
        mrsFindName.MoveFirst
    End If
    
    '开始查找
    If mrsFindName.State <> adStateOpen Then Exit Sub
    If mrsFindName.RecordCount = 0 Then Exit Sub
    
    For n = 1 To mrsFindName.RecordCount
        '如果到底了，则返回第1条记录
        If mrsFindName.EOF Then mrsFindName.MoveFirst
        If mrsFindName.RecordCount = 1 Then mlngFindCurrRow = 1
        
        If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
            str药名 = mrsFindName!药品编码 & mrsFindName!通用名
        Else
            str药名 = mrsFindName!药品编码 & IIf(IsNull(mrsFindName!商品名), mrsFindName!通用名, mrsFindName!商品名)
        End If
        lngFindRow = vsfPrice.FindRow(str药名, mlngFindCurrRow, CLng(vsfPrice.ColIndex("品名")), True, True)
        
        If lngFindRow > 0 Then '查询到数据后就移动下到下一行，继续检查下一行是否有相同的药品
'            vsfPrice.Select lngFindRow, 1, lngFindRow, vsfPrice.Cols - 1
            vsfPrice.TopRow = lngFindRow
            vsfPrice.Row = lngFindRow
            vsfPrice.Col = vsfPrice.ColIndex("现价格")
                        
            If lngFindRow < vsfPrice.rows - 2 Then
                mlngFindCurrRow = lngFindRow + 1
            Else
                mlngFindCurrRow = 1
                mrsFindName.MoveNext '未查询到数据则移动到下一条数据集继续查询
            End If
            Exit For
        Else
            mrsFindName.MoveNext '未查询到数据则移动到下一条数据集继续查询
            mlngFindCurrRow = 1 '继续从第一行开始比较其他药品
        End If
    Next
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

