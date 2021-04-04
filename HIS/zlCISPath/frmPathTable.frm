VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPathTable 
   BorderStyle     =   0  'None
   Caption         =   "临床路径表"
   ClientHeight    =   7245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9255
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame fraTop 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   15015
      Begin VB.Frame fraPath 
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   380
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   10815
         Begin VB.ComboBox cboPath 
            Height          =   300
            Left            =   870
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   30
            Width           =   2415
         End
         Begin VB.Label lblInPep 
            BackColor       =   &H8000000E&
            Caption         =   "导入人：梁唐彬"
            Height          =   255
            Left            =   3360
            TabIndex        =   18
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label lblInDate 
            BackColor       =   &H8000000E&
            Caption         =   "导入时间：2011-01-01 24:24"
            Height          =   255
            Left            =   5160
            TabIndex        =   17
            Top             =   120
            Width           =   2415
         End
         Begin VB.Label lblOutDate 
            BackColor       =   &H8000000E&
            Caption         =   "结束时间：2011-01-01 01:01"
            Height          =   255
            Left            =   7620
            TabIndex        =   16
            Top             =   120
            Width           =   2415
         End
         Begin VB.Label lblInDiag 
            BackColor       =   &H8000000E&
            Caption         =   "导入诊断："
            Height          =   255
            Left            =   10080
            TabIndex        =   15
            Top             =   120
            Width           =   4995
         End
         Begin VB.Label Label1 
            BackColor       =   &H8000000E&
            Caption         =   "路径名称"
            Height          =   255
            Left            =   60
            TabIndex        =   14
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.Frame fraSendor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   380
         Left            =   11880
         TabIndex        =   7
         Top             =   0
         Width           =   3015
         Begin VB.OptionButton optSelect 
            BackColor       =   &H00FFFFFF&
            Caption         =   "全部"
            Height          =   180
            Index           =   0
            Left            =   840
            TabIndex        =   10
            Top             =   120
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton optSelect 
            BackColor       =   &H00FFFFFF&
            Caption         =   "医生"
            Height          =   180
            Index           =   1
            Left            =   1560
            TabIndex        =   9
            Top             =   120
            Width           =   735
         End
         Begin VB.OptionButton optSelect 
            BackColor       =   &H00FFFFFF&
            Caption         =   "护士"
            Height          =   180
            Index           =   2
            Left            =   2280
            TabIndex        =   8
            Top             =   120
            Width           =   735
         End
         Begin VB.Label lblSendNote 
            BackColor       =   &H00FFFFFF&
            Caption         =   "生成者："
            Height          =   180
            Left            =   120
            TabIndex        =   11
            Top             =   120
            Width           =   735
         End
      End
   End
   Begin zlCISPath.UCAdviceList UCAdvice 
      Height          =   1455
      Left            =   0
      TabIndex        =   4
      Top             =   5640
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   2566
   End
   Begin VB.Frame fraline 
      Height          =   30
      Left            =   0
      MousePointer    =   7  'Size N S
      TabIndex        =   3
      Top             =   5520
      Width           =   8175
   End
   Begin MSComctlLib.ImageList imgCharacter 
      Left            =   8280
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTable.frx":0000
            Key             =   "已经执行"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTable.frx":059A
            Key             =   "尚未执行"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTable.frx":0B34
            Key             =   "取消执行"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTable.frx":10CE
            Key             =   "部分执行"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTable.frx":1668
            Key             =   "提前执行"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTable.frx":1C02
            Key             =   "延后执行"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraMore 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   8400
      TabIndex        =   2
      Top             =   2280
      Visible         =   0   'False
      Width           =   225
      Begin VB.Image imgMore 
         Height          =   225
         Left            =   0
         Picture         =   "frmPathTable.frx":219C
         Top             =   0
         Width           =   225
      End
   End
   Begin MSComctlLib.ImageList imgFlow 
      Left            =   8280
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTable.frx":259D
            Key             =   "node"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTable.frx":26E4
            Key             =   "currnode"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTable.frx":2833
            Key             =   "multnode"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTable.frx":29B5
            Key             =   "currmultnode"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTable.frx":2B7B
            Key             =   "arrow"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTable.frx":2FFE
            Key             =   "arrowlate"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTable.frx":3479
            Key             =   "arrow_Branch"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTable.frx":3899
            Key             =   "arrowlate_Branch"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPath 
      Height          =   3105
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "双击查看路径项目定义"
      Top             =   2280
      Width           =   8175
      _cx             =   14420
      _cy             =   5477
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
      BackColorFixed  =   15597549
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   32768
      GridColorFixed  =   32768
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   3
      FixedRows       =   3
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPathTable.frx":3CBD
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
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
      OwnerDraw       =   0
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   14811105
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8Ctl.VSFlexGrid vsFlow 
      Height          =   1920
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "双击查看路径阶段定义"
      Top             =   360
      Width           =   8175
      _cx             =   14420
      _cy             =   3387
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483634
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   16777215
      GridColor       =   0
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   1800
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPathTable.frx":3DFA
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   101
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
      OwnerDraw       =   0
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   16777215
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin VB.Label lblPrinted 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "已打印患者版路径表"
         ForeColor       =   &H0000C0C0&
         Height          =   225
         Left            =   0
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   1725
         WordWrap        =   -1  'True
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPathPrint 
      Height          =   3105
      Index           =   0
      Left            =   0
      TabIndex        =   19
      ToolTipText     =   "双击查看路径项目定义"
      Top             =   -99999
      Visible         =   0   'False
      Width           =   8175
      _cx             =   14420
      _cy             =   5477
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
      BackColorFixed  =   15597549
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   32768
      GridColorFixed  =   32768
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   3
      FixedRows       =   3
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPathTable.frx":3E67
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
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
      OwnerDraw       =   0
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   14811105
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin XtremeCommandBars.CommandBars cbsSub 
      Left            =   8880
      Top             =   360
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPathTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnMoved As Long
Private mint场合 As Integer                 '0-医生站调用,1-护士站调用
Private mbln启用执行环节 As Boolean         '是否启用路径执行环节
Private mstr执行场合 As String                 '"10"-医生,"01"-护士，"11"—医生护士
Private mbln启用不评估 As Boolean         '允许前一天不评估就生成今天的路径项目 F-评估检查,T-忽略评估检查
Private mbln启用提前生成 As Boolean         '允许提前生成明天的路径项目
Private mbytPrintWay As Byte                '0-表格打印;1-报表打印
Private mblnInsideTools As Boolean          '内部工具条模式
Private mfrmParent As Object, mcbsMain As Object
Attribute mcbsMain.VB_VarUserMemId = 1073938436
Private mclsMipModule As zl9ComLib.clsMipModule '消息平台对象
Private mobjPublicPACS As Object
Public Event ViewEPRReport(ByVal 报告ID As Long, ByVal CanPrint As Boolean)    '要求查看报告
Public Event Activate()    '自已激活时
Public Event RequestRefresh(ByVal lngPathState As Long)  '要求主窗体刷新
Attribute RequestRefresh.VB_UserMemId = 3
Public Event StatusTextUpdate(ByVal Text As String)    '要求更新主窗体状态栏文字
Attribute StatusTextUpdate.VB_UserMemId = 4
Private Const C_Exe = "√"  '■
Private Const C_UnExe = "□"
Private Enum EFixedRow
    R0阶段名 = 0
    R1天数 = 1
    R2日期 = 2
End Enum
'生成者下标值
Private Enum CONST_IX_SENDOR
    IX_ALL = 0
    IX_医生 = 1
    IX_护士 = 2
End Enum

Private mblnUnChange As Boolean    '不调用单元格变化事件，刷新单元格内容
Attribute mblnUnChange.VB_VarUserMemId = 1073938439

Private mPP As TYPE_PATH_Pati
Attribute mPP.VB_VarUserMemId = 1073938440
Private mPati As TYPE_Pati
Attribute mPati.VB_VarUserMemId = 1073938441
Private mcolReason As Collection
Attribute mcolReason.VB_VarUserMemId = 1073938442
Private mblnInOverScope As Boolean    '病人当前执行天数是否在标准住院日范围（允许结束路径）
Attribute mblnInOverScope.VB_VarUserMemId = 1073938443
Private mlng病案状态 As Long    '病案审查状态
Attribute mlng病案状态.VB_VarUserMemId = 1073938444
Private mlng医护科室ID As Long
Private mlng婴儿科室ID As Long
Private mlng婴儿病区ID As Long
Private mlngState As Long      '病人变动状态  =5为转出病人
Attribute mlngState.VB_VarUserMemId = 1073938445
Private mrsPlugInBar As ADODB.Recordset '菜单样式
Private mlngPlugInID As Long '自动执行的插件功能ID

'刷新数据时传入的病人状态
Public Enum TYPE_PATI_State
    ps在院 = 0
    ps预出 = 1
    ps出院 = 2
    ps待诊 = 3          '医生站:待会诊病人(在院)
    ps已诊 = 4          '医生站:已会诊病人
    ps最近转出 = 5      '医护站:最近转科或转病区的病人(在院)
    ps待转入 = 6        '医护站:入科待入住或转病区待入往病人
End Enum

Private mlngFontSize As Long
Attribute mlngFontSize.VB_VarUserMemId = 1073938446
Private mlngPathCount As Long   '当次住院的路径数
Attribute mlngPathCount.VB_VarUserMemId = 1073938447

Private Const CON_SmallFontSize As Long = 9     '小字体
Private Const CON_BigFontSize As Long = 12     '大字体
Private Const CON_PathOutItemColor As Long = &HC0FFFF        '路径外项目，浅黄色
Private Const CON_PathOutItemColorBlue As Long = &HFAEADA    '暂存路径外项目,浅蓝色标识

Private Sub SetUnImport()
'功能：设置未导入时的状态和信息
    With vsFlow
        .Clear
        .Rows = 1
        .Cols = 1
        .ColWidth(0) = 5000
        .ForeColorSel = vbBlack
        .TextMatrix(0, 0) = "  该病人未导入临床路径。"
    End With
    Call ClearPathItem
End Sub

Private Sub SetImportFalse()
'功能：设置当病人导入临床路径失败时的状态和信息
    With vsFlow
        .Clear
        .Rows = 1
        .Cols = 1
        .ColWidth(0) = 5000
        .TextMatrix(0, 0) = "  该病人不符合路径导入条件。" & vbCrLf & "  原因：" & mPP.未导入原因
        .AutoSize 0
        .ForeColorSel = &HC0&
        If .Visible And .Enabled Then .SetFocus
    End With
    Call ClearPathItem
End Sub

Private Sub ClearPathItem(Optional blnImported As Boolean)
'功能：当病人没有可用的临床路径时清除路径表项目
    With vsPath
        .FixedCols = 0
        .FixedRows = 0
        .Rows = 0
        .Cols = 0
        
        If blnImported Then
            .Rows = 1
            .Cols = 1
            .TextMatrix(0, 0) = vbCrLf & "  该病人还没有生成路径项目。"
            .Select 0, 0
            .CellAlignment = flexAlignLeftTop
        End If
        fraSendor.Tag = "隐藏"
    End With
End Sub


Private Sub cboPath_Click()
    If cboPath.ListIndex >= 0 Then
        Call zlRefresh(mPati.病人ID, mPati.主页ID, mPati.病区ID, mPati.科室ID, mPati.病人状态, , , , Val(cboPath.ItemData(cboPath.ListIndex)))
    End If
End Sub

Private Sub cbsSub_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call zlExecuteCommandBars(Control)
End Sub

Private Sub cbsSub_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    Call zlPopupCommandBars(CommandBar)
End Sub

Private Sub cbsSub_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    On Error Resume Next
    Call Me.cbsSub.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    If fraSendor.Visible Or fraPath.Visible Then
        fraTop.Top = lngTop
        fraTop.Left = lngLeft
        fraTop.Width = lngRight - lngLeft
        fraSendor.Left = fraTop.Width - fraSendor.Width
        fraPath.Width = fraSendor.Left
        lngTop = fraTop.Top + fraTop.Height
    End If
    vsFlow.Left = lngLeft
    vsFlow.Top = lngTop
    vsFlow.Width = lngRight - lngLeft
    vsFlow.Height = lngBottom - lngTop
        
    If vsPath.FixedRows = 0 And vsPath.Rows = 0 Then  '没有导入路径时
        vsFlow.Height = vsFlow.Height + vsPath.Height
        vsPath.Visible = False
        UCAdvice.Visible = False
        fraline.Visible = False
    Else
        If vsPath.Visible = False Then vsPath.Visible = True
        If UCAdvice.Visible = False Then UCAdvice.Visible = True
        If fraline.Visible = False Then fraline.Visible = True
        
        If Grid.HScrollVisible(vsFlow) = False Then
            vsFlow.Height = 1140 + IIf(mPP.合并路径个数 > 2, (mPP.合并路径个数 - 2) * 180, 0)
        Else
            vsFlow.Height = 1440 + IIf(mPP.合并路径个数 > 2, (mPP.合并路径个数 - 2) * 180, 0)
        End If
        With vsPath
            .Top = lngTop + vsFlow.Height
            .Width = lngRight - lngLeft
            If lngBottom - lngTop - vsFlow.Height - IIf(UCAdvice.Visible, UCAdvice.Height + fraline.Height, 0) - 30 > 0 Then
                .Height = lngBottom - lngTop - vsFlow.Height - IIf(UCAdvice.Visible, UCAdvice.Height + fraline.Height, 0) - 30
            Else
                .Height = lngBottom - lngTop - vsFlow.Height
            End If
        
            If .FixedRows = 0 And .Rows = 1 Then             '没有生成项目
                .ColWidth(0) = .Width - 30
                .RowHeight(0) = .Height
            End If
            fraline.Top = .Top + .Height
            fraline.Width = .Width
            
            UCAdvice.Top = fraline.Top + fraline.Height
            UCAdvice.Width = .Width
        End With
    End If
    
    If fraMore.Visible Then fraMore.Visible = False
End Sub

Private Sub fraline_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        On Error Resume Next
        If vsPath.Height + Y < 1000 Or vsPath.Height - Y < 500 Then Exit Sub
        If UCAdvice.Height + Y < 250 Or UCAdvice.Height - Y < 500 Then Exit Sub
                
        If fraMore.Visible Then fraMore.Visible = False
        
        fraline.Top = fraline.Top + Y
        vsPath.Height = vsPath.Height + Y
        UCAdvice.Top = UCAdvice.Top + Y
        UCAdvice.Height = UCAdvice.Height - Y
    End If
End Sub

Private Sub cbsSub_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call zlUpdateCommandBars(Control)
End Sub

Private Sub Form_Load()
    Set mrsPlugInBar = Nothing
    Call RestoreWinState(Me, App.ProductName)
    Call InitCbsSubBar
    '初始化LIS对象
    Call InitObjLis(P临床路径应用)
End Sub

Private Sub Form_Resize()
    Call cbsSub_Resize
    lblPrinted.Top = vsFlow.RowPos(0)
    lblPrinted.Left = vsFlow.ColPos(0)
End Sub

Private Sub LoadPathFlow()
'功能：根据病人导入的路径表加载路径基本信息和流程
    Dim strSql As String, i As Long, j As Long, lngCurCol As Long
    Dim rsTmp As ADODB.Recordset, lngDayMin As Long, lngDayMax As Long
    Dim lng理论天数 As Long
    Dim lng序号 As Long
    Dim str标准住院日 As String
    Dim rsBranch As Recordset
    Dim rsMerge As Recordset

    With vsFlow
        .Clear
        .Rows = 1: .Cols = 1
        .ForeColorSel = vbBlack
        mblnInOverScope = False
        On Error GoTo errH
        '已打印患者版路径表
        strSql = "Select 1 From 电子病历打印 Where 文件id = [1] And 种类 = 12"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID)
        If rsTmp.RecordCount > 0 Then
            lblPrinted.Caption = "已打印患者版路径表"
            lblPrinted.ForeColor = &HC0&
            lblPrinted.Visible = True
        Else
            lblPrinted.Caption = ""
            lblPrinted.Visible = False
        End If

        If mPP.当前阶段分支ID <> 0 Then
            strSql = "Select NVL(c.序号,b.序号) as 序号 From 临床路径分支 A,临床路径阶段 B,临床路径阶段 C Where a.前一阶段ID=b.ID And b.父ID=c.id(+) And a.ID=[1]"
            '先求出当前分支的前一阶段的序号
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.当前阶段分支ID)
            lng序号 = Val(rsTmp!序号 & "")
        End If

        If mPP.当前阶段分支ID = 0 Then
            strSql = "Select a.ID,a.名称 阶段名, Decode(a.结束天数, Null, 0, 1) 多天,b.分类,b.名称 路径名,b.最新版本,c.标准住院日 ,a.分支id" & _
                     " From 临床路径阶段 a,临床路径目录 b,临床路径版本 c " & _
                     " Where a.路径id = [1] And a.版本号 = [2] And a.路径id=b.id And a.父ID is null And b.id = c.路径id And a.版本号 = c.版本号 " & _
                     " And a.分支ID is Null" & _
                     " Order by a.序号"
        Else
            strSql = "Select a.ID,a.名称 阶段名, Decode(a.结束天数, Null, 0, 1) 多天,b.分类,b.名称 路径名,b.最新版本,c.标准住院日 ,a.分支id" & _
                     " From 临床路径阶段 a,临床路径目录 b,临床路径版本 c,临床路径分支 D,临床路径阶段 E,临床路径阶段 F,临床路径阶段 G " & _
                     " Where a.路径id = [1] And a.版本号 = [2] And a.路径id=b.id And a.父ID is null And b.id = c.路径id And a.版本号 = c.版本号 " & _
                     " And a.分支ID=d.ID(+) And a.父ID=e.id(+) And d.前一阶段ID=f.id(+) And f.父ID=g.id(+) And (a.分支ID=[3] Or NVL(e.序号,a.序号)<=[4] and a.分支ID is null )" & _
                     " Order by Decode(a.分支ID,Null,NVL(e.序号,a.序号),NVL(e.序号,a.序号)+NVL(g.序号,f.序号))"
        End If

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.路径ID, mPP.版本号, mPP.当前阶段分支ID, lng序号)
        If mPP.当前阶段分支ID = 0 Then
            str标准住院日 = rsTmp!标准住院日 & ""
        Else
            strSql = "Select 标准住院日 From 临床路径分支 Where ID=[1]"
            Set rsBranch = zlDatabase.OpenSQLRecord(strSql, "获取分支标准住院日", mPP.当前阶段分支ID)
            str标准住院日 = rsBranch!标准住院日 & ""
        End If
        If rsTmp.RecordCount > 0 Then
            .Rows = 1
            .Cols = rsTmp.RecordCount * 2    '第一列为路径名，箭头为阶段数-1
            .Select 0, 0
            .RowHeight(0) = 1100 + IIf(mPP.合并路径个数 > 2, (mPP.合并路径个数 - 2) * 180, 0)

            '第一列显示路径名称
            .ColWidth(0) = 2800

            If mPP.病人路径状态 > 0 Then
                strSql = "Select b.名称,a.结束时间 From 病人合并路径 A,临床路径目录 B Where a.路径ID=b.ID And a.首要路径记录ID=[1]"
                Set rsMerge = zlDatabase.OpenSQLRecord(strSql, "合并路径", mPP.病人路径ID)
                .TextMatrix(0, 0) = rsTmp!路径名 & ""
                Do While Not rsMerge.EOF
                    .TextMatrix(0, 0) = .TextMatrix(0, 0) & vbCrLf & "(合并)" & rsMerge!名称 & IIf(IsNull(rsMerge!结束时间), "", "(完成)")
                    rsMerge.MoveNext
                Loop


                If mPP.病人路径状态 = 3 Then
                    .Cell(flexcpForeColor, 0, 0) = vbRed
                End If
            Else
                .TextMatrix(0, 0) = rsTmp!路径名
            End If
            If mPP.当前天数 > 0 And mPP.病人路径状态 = 1 Then
                If InStr(str标准住院日, "-") > 0 Then
                    j = Split(str标准住院日, "-")(1)
                    lngDayMin = Val(Split(str标准住院日, "-")(0))
                    lngDayMax = j
                Else
                    j = Val(str标准住院日)   '小于等于n天的情况
                    lngDayMin = 1
                    lngDayMax = j
                End If
                lng理论天数 = GetMustDay(mPP.病人路径ID, mPP.当前天数)

                i = Format(lng理论天数 / j * 100, "0")
                If i = 100 And lng理论天数 <> j Then i = 99
                .TextMatrix(0, 0) = .TextMatrix(0, 0) & vbCrLf & "进度：" & i & "%"


                If lng理论天数 > lngDayMax Then
                    mblnInOverScope = True
                Else
                    mblnInOverScope = Between(lng理论天数, lngDayMin, lngDayMax)
                End If
            End If
            If mPP.病人路径状态 > 0 Then
                .TextMatrix(0, 0) = .TextMatrix(0, 0) & vbCrLf & "状态：" & IIf(mPP.病人路径状态 = 1, "执行中", IIf(mPP.病人路径状态 = 2, "完成", "变异退出"))
            End If
            .Cell(flexcpTextStyle, 0, 0) = 3

            For i = 1 To .Cols Step 2
                .TextMatrix(0, i) = " " & rsTmp!阶段名 & " "    '设置边距
                .ColAlignment(i) = flexAlignCenterCenter

                .ColWidth(i) = 1750
                .Col = i
                .PicturesOver = True
                .CellPictureAlignment = flexPicAlignLeftCenter
                If mPP.当前阶段ID = rsTmp!ID Or mPP.阶段父ID = rsTmp!ID Or (mPP.当前阶段ID = 0 And i = 1 And mPP.病人路径状态 = 1) Then
                    lngCurCol = i
                    .CellPicture = imgFlow.ListImages(IIf(rsTmp!多天 = 1, "currmultnode", "currnode")).Picture
                    Call .ShowCell(0, i)
                Else
                    .CellPicture = imgFlow.ListImages(IIf(rsTmp!多天 = 1, "multnode", "node")).Picture
                End If
                .ColData(i) = Val(rsTmp!ID)

                rsTmp.MoveNext

                '箭头
                If i < .Cols - 1 Then
                    .ColWidth(i + 1) = 550
                    .Col = i + 1
                    .CellPictureAlignment = flexPicAlignCenterCenter
                    .CellPicture = imgFlow.ListImages(IIf(i + 1 > lngCurCol And lngCurCol <> 0 Or mPP.病人路径状态 > 1, "arrowlate", "arrow") & IIf(rsTmp!分支ID & "" <> "", "_Branch", "")).Picture
                End If
            Next
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Get执行结果性质图标(ByVal lng执行结果性质 As Long) As Long
'功能：根据执行结果性质返回对应的图标序号
'1-已经执行，2-尚未执行，3-取消执行，4-部分执行，5-提前执行，6-延后执行
    Dim lngIdx As Long
    
    Select Case lng执行结果性质
        Case 1
            lngIdx = imgCharacter.ListImages("已经执行").Index
        Case 2
            lngIdx = imgCharacter.ListImages("尚未执行").Index
        Case 3
            lngIdx = imgCharacter.ListImages("取消执行").Index
        Case 4
            lngIdx = imgCharacter.ListImages("部分执行").Index
        Case 5
            lngIdx = imgCharacter.ListImages("提前执行").Index
        Case 6
            lngIdx = imgCharacter.ListImages("延后执行").Index
    End Select
    Get执行结果性质图标 = lngIdx
End Function

Private Sub LoadPathItem()
'功能：加载病人已生成的路径项目
    Dim strSql As String, strOldType As String, str评估结果 As String
    Dim lngRow As Long, lngCol As Long, i As Long, j As Long, arrtmp As Variant, strTmp As String
    Dim rsTmp As ADODB.Recordset
    Dim CPos As New Collection  '每个分类的起始行
    Dim lngPreRow As Long, lngPreCol As Long, lngPrePathID As Long, lngDayRow As Long
    Dim lngBranchID As Long  '分支ID
    Dim rsSort As Recordset
    Dim str变异原因 As String
    
    With vsPath
        lngPreRow = -1
        lngPreCol = -1
        If .Row >= .FixedRows Then lngPreRow = .Row
        If .Col >= .FixedCols Then lngPreCol = .Col

        '1)分类部分
        .Redraw = flexRDNone
        mblnUnChange = True
        .Clear
        .Rows = 3: .FixedRows = 3
        .Cols = 1: .FixedCols = 1
        mblnUnChange = False
        .MergeCol(0) = True
        .MergeRow(0) = True

        .ColWidth(0) = 1200
        .FixedAlignment(0) = flexAlignCenterCenter
        .Cell(flexcpText, 0, 0, .FixedRows - 1, 0) = "时间阶段"
        On Error GoTo errH

        '如果存在路径跳转，不同阶段的项目可能是不同的路径表的
        strSql = _
        "Select 分类, Max(个数) As 个数,100 as 序号" & vbNewLine & _
                 "From (Select Count(a.Id) As 个数, a.分类, a.阶段id, a.日期" & vbNewLine & _
                 "       From 病人路径执行 A" & vbNewLine & _
                 "       Where a.路径记录id = [1]" & vbNewLine & _
                 "       Group By a.分类, a.日期, a.阶段id)" & vbNewLine & _
                 "Group By 分类"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID)
        Set rsTmp = zlDatabase.CopyNewRec(rsTmp)

        '读取序号(规则：主路径优先、然后是分支路径、最后是合并路径，如果有交叉的情况也是这个规则)
        strSql = _
        "Select 分类, 序号" & vbNewLine & _
                 "From (Select 分类, 序号," & vbNewLine & _
                 "              Row_Number() Over(Partition By 分类 Order By Decode(合并路径记录id, Null, Decode(分支id, Null, 1, 2), Decode(分支id, Null, 3, 4))) As Top" & vbNewLine & _
                 "       From (Select a.序号, a.名称 As 分类, c.分支id, b.合并路径记录id" & vbNewLine & _
                 "              From 临床路径分类 A, 病人路径执行 B, 临床路径项目 C" & vbNewLine & _
                 "              Where a.名称 = c.分类 And b.路径记录id = [1] And b.项目id = c.Id And c.路径id = a.路径id And c.版本号 = a.版本号 And" & vbNewLine & _
                 "                    Nvl(c.分支id, 0) = Nvl(a.分支id, 0)" & vbNewLine & _
                 "              Union" & vbNewLine & _
                 "              Select a.序号, a.名称 As 分类, c.分支id, b.合并路径记录id" & vbNewLine & _
                 "              From 临床路径分类 A, 病人路径执行 B, 临床路径阶段 C" & vbNewLine & _
                 "              Where a.名称 = b.分类 And b.阶段id+0 = c.Id And b.路径记录id = [1] And b.项目id Is Null And a.路径id = c.路径id And" & vbNewLine & _
                 "                    a.版本号 = c.版本号 And Nvl(c.分支id, 0) = Nvl(a.分支id, 0)))" & vbNewLine & _
                 "Where Top = 1" & vbNewLine & _
                 "Order By 序号"

        Set rsSort = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID)
        '排序
        If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            Do While Not rsTmp.EOF
                rsSort.Filter = "分类='" & rsTmp!分类 & "'"
                If rsSort.RecordCount > 0 Then
                    rsTmp!序号 = Val(rsSort!序号 & "")
                    rsTmp.Update
                End If
                rsTmp.MoveNext
            Loop
            rsTmp.Sort = "序号"
            rsTmp.MoveFirst
        End If
        For i = 1 To rsTmp.RecordCount
            CPos.Add .Rows, "T" & rsTmp!分类
            .Rows = .Rows + rsTmp!个数
            For j = 1 To rsTmp!个数
                .TextMatrix(.Rows - j, .FixedCols - 1) = rsTmp!分类
            Next
            rsTmp.MoveNext
        Next


        '2)时间阶段部分
        '阶段排序时用 NVL(c.序号,b.序号) 是为了处理备用分支序列排序的问题，取值b.序号 是因为界面上需要显示是第几个分支。（取分支路径的序号时，取其其一阶段的序号加上分支路径的序号）
        If mPP.当前阶段分支ID = 0 Then
            strSql = _
            "Select a.阶段id, a.天数, To_Char(a.日期, 'yyyy-mm-dd') 日期, To_Char(a.日期, 'day') 星期, b.名称 As 阶段名, b.序号, b.说明, b.父id,Decode(g.路径id,b.路径id,1,0) as 排序" & vbNewLine & _
                     "From (Select a.阶段id, a.天数, a.日期,a.路径记录id" & vbNewLine & _
                     "       From 病人路径执行 A" & vbNewLine & _
                     "       Where a.路径记录id = [1]" & vbNewLine & _
                     "       Group By a.阶段id, a.天数, a.日期,a.路径记录id) A, 临床路径阶段 B,临床路径阶段 C,病人临床路径 G" & vbNewLine & _
                     "Where a.阶段id = b.Id And b.父id=c.id(+) And g.id=A.路径记录ID " & vbNewLine & _
                     "Order By 日期,排序, NVL(c.序号,b.序号)"
        Else
            strSql = _
            "Select a.阶段id, a.天数, To_Char(a.日期, 'yyyy-mm-dd') 日期, To_Char(a.日期, 'day') 星期, b.名称 As 阶段名, b.序号, b.说明, b.父id,Decode(g.路径id,b.路径id,1,0) as 排序" & vbNewLine & _
                     "From (Select a.阶段id, a.天数, a.日期,a.路径记录id" & vbNewLine & _
                     "       From 病人路径执行 A" & vbNewLine & _
                     "       Where a.路径记录id = [1]" & vbNewLine & _
                     "       Group By a.阶段id, a.天数, a.日期,a.路径记录id) A, 临床路径阶段 B,临床路径阶段 C,临床路径分支 D,临床路径阶段 E,临床路径阶段 F,病人临床路径 G" & vbNewLine & _
                     "Where a.阶段id = b.Id And b.父id=c.id(+) And b.分支id=d.id(+) and d.前一阶段id=e.id(+) And e.父id=f.id(+)  And g.id=A.路径记录ID " & vbNewLine & _
                     "Order By 日期,排序, Decode(b.分支ID,Null,NVL(c.序号,b.序号),NVL(c.序号,b.序号)+NVL(f.序号,e.序号))"
        End If

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID)
        .AutoSizeMode = flexAutoSizeRowHeight
        .Cols = .Cols + rsTmp.RecordCount
        For i = 1 To rsTmp.RecordCount
            .ColWidth(i) = 2800
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColData(i) = Val("" & rsTmp!阶段ID)
            If IsNull(rsTmp!父ID) Then
                .TextMatrix(EFixedRow.R0阶段名, i) = Replace(rsTmp!阶段名, vbLf, vbCrLf)    '为了打印时正常换行(vbLf时换行显示有问题)
            Else
                .TextMatrix(EFixedRow.R0阶段名, i) = Replace(rsTmp!阶段名, vbLf, vbCrLf) & ",分支:" & Nvl(rsTmp!说明, rsTmp!序号)
            End If
            .TextMatrix(EFixedRow.R1天数, i) = "第" & rsTmp!天数 & "天"
            .Cell(flexcpData, EFixedRow.R1天数, i) = rsTmp!天数
            .TextMatrix(EFixedRow.R2日期, i) = rsTmp!日期 & "(" & rsTmp!星期 & ")"
            .Cell(flexcpData, EFixedRow.R2日期, i) = rsTmp!日期 & ""
            If rsTmp!天数 = mPP.当前天数 Then mPP.当前日期 = rsTmp!日期
            rsTmp.MoveNext
        Next
        
        For i = 1 To mcolReason.count
            mcolReason.Remove 1 '删除局部变量数据(上下移动后,重新加载时需要清空变异原因)
        Next i
        strSql = _
        "Select a.Id, Nvl(b.图标id, a.图标id) 图标id, a.分类, To_Char(a.日期, 'yyyy-mm-dd') 日期, a.天数, a.阶段id, Nvl(a.项目序号, b.项目序号) As 项目序号," & vbNewLine & _
                 "Nvl(b.项目内容, a.项目内容) 项目内容, a.项目id, Decode(a.执行人, Null, 0, 1) 执行状态, Nvl(b.执行方式, 1) 执行方式, a.添加原因,NVl(a.生成时间性质,0) as 生成时间性质, c.名称 As 变异原因," & vbNewLine & _
                 "Nvl(b.项目结果, a.项目结果) As 项目结果, a.执行结果, d.路径id, d.分支id,NVL(NVL(A.生成者,B.生成者),1) as 生成者" & vbNewLine & _
                 "From 病人路径执行 A, 临床路径项目 B, 变异常见原因 C, 临床路径阶段 D" & vbNewLine & _
                 "Where a.路径记录id = [1] And a.项目id = b.Id(+) And a.变异原因 = c.编码(+) And a.阶段id + 0 = d.Id" & vbNewLine & _
                 Decode(Val(optSelect(IX_ALL).Tag), 0, " ", 1, " And Decode(a.项目id, Null, Nvl(a.生成者, 1), Nvl(b.生成者, 1)) = 1 ", 2, " And Decode(a.项目id, Null, Nvl(a.生成者, 1), Nvl(b.生成者, 1)) = 2") & vbNewLine & _
                 "Order By a.日期, 分类, 项目序号"
        'Nvl(a.生成者,1)是为了兼容以前版本,不存在生成者时默认为医生。
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID)
        For lngCol = .FixedCols To .Cols - 1
            rsTmp.Filter = "阶段ID='" & .ColData(lngCol) & "' And 天数=" & Val(Replace(.TextMatrix(EFixedRow.R1天数, lngCol), "第", ""))
            strOldType = ""
            If rsTmp.RecordCount > 0 Then
                If lngPrePathID <> Val(rsTmp!路径ID & "") Or lngBranchID <> Val(rsTmp!分支ID & "") Then
                    '路径跳转，画竖分隔线
                    If lngPrePathID <> 0 Or Val(rsTmp!分支ID & "") <> 0 Then Call .CellBorderRange(0, lngCol, .Rows - 1, lngCol, vbBlack, 1, 0, 0, 0, 0, 1)
                    lngPrePathID = rsTmp!路径ID
                    lngBranchID = Val(rsTmp!分支ID & "")
                End If
            End If

            Do While Not rsTmp.EOF
                If strOldType <> rsTmp!分类 Then
                    lngRow = CPos("T" & rsTmp!分类)
                    strOldType = rsTmp!分类
                End If

                If mbln启用执行环节 Then
                    .TextMatrix(lngRow, lngCol) = IIf(rsTmp!执行方式 = 0, "", IIf(rsTmp!执行状态 = 0, C_UnExe, C_Exe)) & rsTmp!项目内容
                Else
                    .TextMatrix(lngRow, lngCol) = "" & rsTmp!项目内容   '医嘱界面添加后，还未弹出路径外项目前刷新，项目内容为空
                End If
                '附加数据组织形式 ID|项目ID|项目序号|生成者|生成时间性质
                '路径外项目项目id为空
                '生成者 1-医生,2-护士
                .Cell(flexcpData, lngRow, lngCol) = Val(rsTmp!ID) & "|" & Val("" & rsTmp!项目ID) & "|" & Val("" & rsTmp!项目序号) & "|" & rsTmp!生成者 & "|" & rsTmp!生成时间性质
                
                If IsNull(rsTmp!项目ID) Then
                    .Cell(flexcpBackColor, lngRow, lngCol) = CON_PathOutItemColor         '路径外项目，浅黄色
                    If Val(rsTmp!生成时间性质 & "") = 2 Then
                        .Cell(flexcpBackColor, lngRow, lngCol) = CON_PathOutItemColorBlue     '暂存外项目,浅蓝色
                    End If
                    mcolReason.Add "变异说明：" & rsTmp!添加原因 & vbCrLf & "变异原因：" & rsTmp!变异原因, "C" & rsTmp!ID
                    If rsTmp!变异原因 & "" <> "" Or rsTmp!添加原因 & "" <> "" Then
                        .TextMatrix(lngRow, lngCol) = .TextMatrix(lngRow, lngCol) & vbCrLf & "变异原因：" & rsTmp!变异原因 & vbCrLf & "变异说明：" & rsTmp!添加原因
                    End If
                ElseIf InStr("124", CStr(rsTmp!执行方式)) > 0 Then '必须生成的，未生成
                    If Not IsNull(rsTmp!变异原因) Then
                        .Cell(flexcpBackColor, lngRow, lngCol) = &HE0EFED    '浅灰色
                        .TextMatrix(lngRow, lngCol) = .TextMatrix(lngRow, lngCol) & vbCrLf & "变异原因：" & rsTmp!变异原因
                    End If
                ElseIf rsTmp!执行方式 = 3 Then                          '可选项，深蓝色
                    .Cell(flexcpForeColor, lngRow, lngCol) = &HC00000
                    If Not IsNull(rsTmp!变异原因) Then '93648 中药路径项目的变异原因
                        .Cell(flexcpBackColor, lngRow, lngCol) = &HE0EFED    '浅灰色
                        .TextMatrix(lngRow, lngCol) = .TextMatrix(lngRow, lngCol) & vbCrLf & "变异原因：" & rsTmp!变异原因
                    End If
                End If

                If InStr(rsTmp!项目结果, "|") > 0 And Not IsNull(rsTmp!执行结果) Then
                    i = Val(Mid(rsTmp!项目结果, InStr(rsTmp!项目结果, rsTmp!执行结果) + Len(rsTmp!执行结果) + 1, 1))
                    If i > 0 Then i = Get执行结果性质图标(i)
                Else
                    i = 0
                End If

                If Not IsNull(rsTmp!图标ID) Or i > 0 Then
                    .Cell(flexcpPictureAlignment, lngRow, lngCol) = flexPicAlignRightCenter    ' flexPicAlignLeftCenter
                    If i > 0 Then
                        .Cell(flexcpPicture, lngRow, lngCol) = imgCharacter.ListImages(i).Picture
                    Else
                        .Cell(flexcpPicture, lngRow, lngCol) = GetPathIcon(rsTmp!图标ID)
                    End If
                End If

                lngRow = lngRow + 1
                rsTmp.MoveNext
            Loop
        Next

        '4)显示评估信息

        If .Rows = .FixedRows And .Cols = .FixedCols Then
            Call ClearPathItem(True)
            .BackColorSel = vbWhite
            .ForeColorSel = vbBlack
        Else
            If Val(optSelect(IX_ALL).Tag) <> IX_护士 Then '目前未考虑护士参与评估,故选中护士不显示评估信息
                .BackColorSel = &H8000000D
                .ForeColorSel = &H8000000E
                .Rows = .Rows + 1
                lngRow = .Rows - 1
                lngDayRow = .FixedRows - 2
                .TextMatrix(lngRow, .FixedCols - 1) = "评估情况"
                .Cell(flexcpBackColor, lngRow, 0) = .BackColorFixed  '&HEFF0E0      '&HD0EFFF
                Call .CellBorderRange(.Rows - 1, 0, .Rows - 1, .Cols - 1, vbBlack, 0, 1, 0, 0, 0, 0)
    
                strSql = "Select a.阶段id, a.天数, a.评估结果, a.评估说明, a.评估人,a.评估时间, c.名称 As 变异原因, a.变异审核人, Nvl(a.时间进度, 0) 时间进度, a.跳转审核人, a.原路径id" & vbNewLine & _
                        "From 病人路径评估 A, 病人路径变异 B, 变异常见原因 C" & vbNewLine & _
                        "Where a.路径记录id = b.路径记录id(+) And a.阶段ID=B.阶段ID(+) And a.日期=b.日期(+) And a.路径记录id = [1] And b.变异原因 = c.编码(+)"

                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID)
                For lngCol = .FixedCols To .Cols - 1
                    .Cell(flexcpBackColor, lngRow, lngCol) = &HEDF8FF   '&HD0EFFF
    
                    rsTmp.Filter = "阶段ID='" & .ColData(lngCol) & "' And 天数=" & Val(Replace(.TextMatrix(EFixedRow.R1天数, lngCol), "第", ""))
                    str变异原因 = ""
                    For j = 1 To rsTmp.RecordCount
                        '读取多个变异原因
                        str变异原因 = str变异原因 & rsTmp!变异原因 & "、"
                        If j = rsTmp.RecordCount Then
                            str变异原因 = Mid(str变异原因, 1, Len(str变异原因) - 1)
                            If InStr(rsTmp!评估说明, vbCrLf) = 0 Or IsNull(rsTmp!评估说明) Then
                                strTmp = "" & rsTmp!评估说明
                            Else
                                arrtmp = Split(rsTmp!评估说明, vbCrLf)
                                strTmp = ""
                                For i = 0 To UBound(arrtmp)
                                    strTmp = strTmp & vbCrLf & Space(4) & (i + 1) & "." & arrtmp(i)
                                Next
                            End If
                            strTmp = strTmp & vbCrLf & "评 估 人：" & rsTmp!评估人
                            If rsTmp!评估结果 = 1 Then
                                str评估结果 = "正常"
                            ElseIf mPP.病人路径状态 = 3 And lngCol = .Cols - 1 Then
                                str评估结果 = "变异后退出" & vbCrLf & "变异原因：" & str变异原因 & vbCrLf & "审 核 人：" & rsTmp!变异审核人
        
                            ElseIf mPP.病人路径状态 = 2 And lngCol = .Cols - 1 Then
                                str评估结果 = "变异后完成" & vbCrLf & "变异原因：" & str变异原因
                            Else
                                str评估结果 = "变异后继续" & vbCrLf & "变异原因：" & str变异原因
                                If Not IsNull(rsTmp!变异审核人) Then str评估结果 = str评估结果 & vbCrLf & "审 核 人：" & rsTmp!变异审核人
                            End If
        
                            .TextMatrix(lngRow, lngCol) = "评估结果：" & str评估结果 & vbCrLf & "评估说明：" & strTmp
                            If rsTmp!评估结果 = -1 Then
                                .Cell(flexcpForeColor, lngRow, lngCol) = vbRed     '变异用红色表示
                            End If
        
                            If rsTmp!时间进度 = 1 Or rsTmp!时间进度 = 2 Then
                                '提前
                                .TextMatrix(lngDayRow, lngCol) = .TextMatrix(lngDayRow, lngCol) & "←"
                                .Cell(flexcpForeColor, lngDayRow, lngCol) = &H80FF&
                            ElseIf rsTmp!时间进度 = -1 Then    '延后
                                .TextMatrix(lngDayRow, lngCol) = .TextMatrix(lngDayRow, lngCol) & "→"
                                .Cell(flexcpForeColor, lngDayRow, lngCol) = &H80FF&
                            End If
                            '未审核的跳转阶段
                            If rsTmp!原路径ID & "" <> "" And rsTmp!跳转审核人 & "" = "" Then
                                .TextMatrix(lngDayRow, lngCol) = .TextMatrix(lngDayRow, lngCol) & "(未审核)"
                                .Cell(flexcpForeColor, lngDayRow, lngCol) = &H80FF&
                            End If
                        End If
                        
                        rsTmp.MoveNext
                    Next
                    If rsTmp.RecordCount = 0 Then
                        .TextMatrix(lngRow, lngCol) = ""
                    End If
                Next
            End If
        End If
        .Redraw = True
        If .FixedRows > 1 Then .AutoSize .FixedCols, .Cols - 1, , 45    '在要Draw之后才生效

        If lngPreRow <> -1 And lngPreCol <> -1 And lngPreRow <= .Rows - 1 And lngPreCol <= .Cols - 1 Then
            .Select lngPreRow, lngPreCol
        Else
            .Select .FixedRows, .FixedCols
        End If
    End With

    Exit Sub
errH:
    vsPath.Redraw = True
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function CheckPathIsTurnAduit() As Boolean
'功能：检查是否存在未审核的跳转阶段。true为存在
     Dim strSql As String, rsTmp As Recordset
     
     strSql = "Select 1 From 病人路径评估 Where 原路径id is not null And 跳转审核人 is null And 路径记录ID=[1]"
     
     On Error GoTo errH
     Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "跳转审核", mPP.病人路径ID)
     
     CheckPathIsTurnAduit = rsTmp.RecordCount > 0
     Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub Get病人路径信息(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lng科室ID As Long, Optional ByVal lng路径记录ID As Long, Optional ByVal blnReadOnly As Boolean)
'功能：获取病人的临床路径信息
'参数：lng路径记录ID=当一个病人有多条路径时，刷新指定路径记录ID的路径表
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    '一次住院只支持一个路径，不管科室
    ' And (科室ID = [3] Or Exists(Select 1 From 部门性质说明 B Where (a.id = b.部门id or b.部门id = [3]) and b.工作性质='ICU'))
    '当前阶段为0表示还未生成过路径
    strSql = "Select a.ID,a.路径ID,c.路径ID as 原路径ID,a.版本号,a.状态,a.当前阶段ID,a.当前天数,b.名称 as 未导入原因,c.父ID,c.分支ID,d.分支ID as 前一阶段分支ID,e.结束路径控制,a.合并路径个数,e.名称 as 路径名称,a.导入人,a.导入时间,a.结束时间" & _
            " From 病人临床路径 A,变异常见原因 B,临床路径阶段 C,临床路径阶段 D,临床路径目录 E" & _
            " Where a.病人ID = [1] And a.主页ID = [2] And a.路径ID=e.id And a.未导入原因 = b.编码(+) And a.当前阶段ID = c.ID(+) And a.前一阶段ID=d.id(+)" & _
            IIf(lng路径记录ID <> 0, " And a.ID=[4] ", "") & _
            " Order By a.导入时间 Desc"  '取最后一次导入的路径（支持一次住院多个路径2012-10-25）
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "Get病人路径信息", lng病人ID, lng主页ID, lng科室ID, lng路径记录ID)
    If rsTmp.RecordCount > 0 Then
        mPP.原路径ID = Val("" & rsTmp!原路径ID)
        mPP.路径ID = rsTmp!路径ID
        mPP.版本号 = rsTmp!版本号
        mPP.病人路径ID = rsTmp!ID
        mPP.病人路径状态 = rsTmp!状态
        mPP.当前阶段ID = Val("" & rsTmp!当前阶段ID)
        mPP.阶段父ID = Val("" & rsTmp!父ID)
        mPP.当前天数 = Val("" & rsTmp!当前天数)
        mPP.导入时间 = CDate(Nvl(rsTmp!导入时间, 0))
        mPP.当前日期 = "0" '在LoadPathItem中赋值
        mPP.未导入原因 = "" & rsTmp!未导入原因
        mPP.当前阶段分支ID = Val("" & rsTmp!分支ID)
        '由于路径结束后会清空当前阶段ID，所以取前一阶段的分支ID
        If mPP.病人路径状态 = 2 Or mPP.病人路径状态 = 3 Then
            mPP.当前阶段分支ID = Val(rsTmp!前一阶段分支ID & "")
        End If
        mPP.结束路径控制 = Val(rsTmp!结束路径控制 & "")
        mPP.合并路径个数 = Val(rsTmp!合并路径个数 & "")
        If lng路径记录ID = 0 Then mlngPathCount = rsTmp.RecordCount
    Else
        mPP.原路径ID = 0
        mPP.路径ID = 0
        mPP.版本号 = 0
        mPP.病人路径ID = 0
        mPP.病人路径状态 = -1
        mPP.当前阶段ID = 0
        mPP.阶段父ID = 0
        mPP.当前天数 = 0
        mPP.当前日期 = "0"
        mPP.未导入原因 = ""
        mPP.当前阶段分支ID = 0
        mPP.结束路径控制 = 0
        mPP.合并路径个数 = 0
        mPP.导入时间 = CDate(0)
        mlngPathCount = 0
    End If
    
        If blnReadOnly Then Exit Sub
    If mlngPathCount > 1 Then
        fraPath.Visible = True
        lblInDiag.Caption = "导入诊断:" & Get导入诊断(mPP.病人路径ID)
        lblInPep.Caption = "导入人:" & rsTmp!导入人
        lblInDate.Caption = "导入时间:" & Format(rsTmp!导入时间, "YYYY-MM-DD HH:mm")
        lblOutDate.Caption = "结束时间:" & Format(rsTmp!结束时间, "YYYY-MM-DD HH:mm")
        If lng路径记录ID = 0 Then
            cboPath.Clear
            Do While Not rsTmp.EOF
                cboPath.AddItem rsTmp!路径名称 & ""
                cboPath.ItemData(cboPath.NewIndex) = rsTmp!ID & ""
                rsTmp.MoveNext
            Loop
            zlControl.CboSetIndex cboPath.Hwnd, 0
        End If
    Else
        fraPath.Visible = False
        cboPath.Clear
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Get导入诊断(ByVal lng路径记录ID As Long) As String
'功能：取导入诊断的名称
    Dim strSql As String, rsTmp As Recordset
    
    On Error GoTo errH
    strSql = "Select B.诊断描述 From 病人临床路径 A,病人诊断记录 B Where " & _
            " a.病人id = b.病人id And a.主页id = b.主页id  and a.诊断类型 = b.诊断类型 And a.诊断来源 = b.记录来源 And NVL(a.疾病id,0) = NVL(b.疾病id,0) And NVL(a.诊断id,0) = NVL(b.诊断id,0) And A.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "Get导入诊断", lng路径记录ID)
    If rsTmp.RecordCount > 0 Then Get导入诊断 = rsTmp!诊断描述 & ""
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub zlPrintOutPut(ByVal bytStyle As Byte, Optional ByVal blnIsSetup As Boolean, Optional ByVal strPDFFile As String, Optional ByVal strDeviceName As String)
'功能：临床路径表单打印
'参数：bytStyle=1-打印,2-预览,3-输出到Excel,4-输出到PDF
'     blnIsSetup-表示批量打印，不进行打印前设置
'     当bytStyle=4时，需要传入strPDFFile=PDF输出默认路径,包含文件名、后缀
    Call FuncPathTableOutput(bytStyle, blnIsSetup, strPDFFile, strDeviceName)
End Sub

Public Function zlRefresh(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lng病区ID As Long, ByVal lng科室ID As Long, _
        ByVal int病人状态 As Integer, Optional ByVal blnMoved As Boolean, Optional ByVal blnForceRefresh As Boolean = True, Optional ByVal lngState As Long, _
        Optional ByVal lng路径记录ID As Long, Optional ByVal lng医护科室ID As Long, Optional ByRef objMip As Object, Optional ByVal blnReadOnly As Boolean) As Long
'参数：lng路径记录ID=当一个病人有多条路径时，刷新指定路径记录ID的路径表
'      blnForceRefresh=True 未切换病人刷新时也进行刷新，否则不刷新
'      objMip 消息对象
'      blnReadOnly=只读取病人信息，不操作界面
    
    Dim objControl As CommandBarControl
    Dim strPrePati As String
    
    strPrePati = mPati.病人ID & "_" & mPati.主页ID
    If strPrePati = lng病人ID & "_" & lng主页ID And lng病人ID <> 0 And Not blnForceRefresh Then Exit Function       '保持之前单元格位置不变
    
    
    If mPati.病人ID & "_" & mPati.主页ID = lng病人ID & "_" & lng主页ID And lng病人ID <> 0 And Not blnForceRefresh Then Exit Function       '保持之前单元格位置不变
    
    mPati.病人ID = lng病人ID
    mPati.主页ID = lng主页ID
    mPati.病区ID = lng病区ID
    mPati.科室ID = lng科室ID
    mPati.病人状态 = int病人状态
    mlngState = lngState
    mlng医护科室ID = lng医护科室ID
    If Not (objMip Is Nothing) Then Set mclsMipModule = objMip
    mlng病案状态 = Get病人病案状态(lng病人ID, lng主页ID, mlng婴儿科室ID, mlng婴儿病区ID)
    mblnMoved = blnMoved

    Set mcolReason = New Collection

    Call Get病人路径信息(lng病人ID, lng主页ID, lng科室ID, lng路径记录ID, blnReadOnly)
    
    '新版病历调用时不传入主窗体对象，只传入句柄，本句会导致主窗体隐藏
    If blnReadOnly = True Then Exit Function
    lblPrinted.Visible = False '默认不显示，LoadPathFlow中设置是否显示
    fraSendor.Tag = ""
    
    If mPP.病人路径ID = 0 Then
        Call SetUnImport
    Else
        If mPP.病人路径状态 = 0 Then
            Call SetImportFalse
        Else
            Call LoadPathFlow
            Call LoadPathItem
        End If
    End If
    fraTop.Visible = True  '若fraTop被隐藏则下面这条语句赋值无效，始终是false
    fraSendor.Visible = (mPP.病人路径状态 > 0 And fraSendor.Tag <> "隐藏")
    fraTop.Visible = fraPath.Visible Or fraSendor.Visible
    Call Form_Resize    '根据路径流程表是否有滚动条来调整高度
    If strPrePati <> lng病人ID & "_" & lng主页ID And lng病人ID <> 0 And mlngPlugInID <> 0 Then
        If mblnInsideTools Then
            Set objControl = cbsSub.FindControl(, mlngPlugInID, , True)
        Else
            Set objControl = mcbsMain.FindControl(, mlngPlugInID, , True)
        End If
        If Not objControl Is Nothing Then
            objControl.Execute
        End If
    End If
End Function


Private Sub InitCbsSubBar()
    Dim objBar As CommandBar
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsSub.VisualTheme = xtpThemeOffice2003
    With Me.cbsSub.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = True
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize True, 16, 16
        .UseSharedImageList = False 'ImageList方式时,因同一App中共享,在AddImageList之前设置为False
    End With
    Set cbsSub.Icons = zlCommFun.GetPubIcons
    cbsSub.EnableCustomization False
    cbsSub.ActiveMenuBar.Visible = False
    
    
    Set objBar = cbsSub.Add("内部工具栏", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False
    objBar.SetIconSize 24, 24
    objBar.Visible = False  '只有内部调用时才显示(zlDefCommandBars)
    
End Sub

Public Sub zlDefCommandBars(ByVal frmParent As Object, ByVal cbsMain As Object, ByVal int场合 As Integer, Optional ByVal blnInsideTools As Boolean = False)
    Dim objControl As CommandBarControl
    Dim objPopup As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim lngStart As Long, i As Long
 
    mint场合 = int场合
    mbln启用执行环节 = Val(zlDatabase.GetPara("是否启用路径执行环节", glngSys, P临床路径应用, 1))
    If mbln启用执行环节 Then
        mstr执行场合 = zlDatabase.GetPara("路径执行环节启用场合", glngSys, P临床路径应用, 11)
    End If
    mbln启用不评估 = Val(zlDatabase.GetPara("允许前一天不评估就生成今天的路径项目", glngSys, P临床路径应用, 1))
    mbln启用提前生成 = Val(zlDatabase.GetPara("允许提前生成明天的路径项目", glngSys, P临床路径应用, 1))
    mbytPrintWay = Val(zlDatabase.GetPara("路径表单打印方式", glngSys, P临床路径应用, "0"))
    mblnInsideTools = blnInsideTools

    Set mfrmParent = frmParent

    If cbsMain Is Nothing Then Exit Sub
    If mrsPlugInBar Is Nothing Then
        Call GetPlugInBar(P临床路径应用, mint场合, mrsPlugInBar)
    End If
    
    Set mcbsMain = cbsMain
    Set cbsMain.Icons = zlCommFun.GetPubIcons

    '文件菜单
    '-----------------------------------------------------
    Set objPopup = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    With objPopup.CommandBar.Controls
        Set objControl = .Find(, conMenu_File_Excel)
        objControl.Caption = "输出到&Excel(医师版)…"
        Set objControl = .Find(, conMenu_File_Print)
        objControl.Caption = "打印路径表(医师版)(&P)"
        Set objControl = .Add(xtpControlButton, conMenu_File_Print_PatiPath, "打印路径表(患者版)(&Q)", objControl.Index + 1)
        objControl.IconId = conMenu_File_Print
    End With

    '编辑菜单:放在管理菜单(主窗体可能没有)、文件菜单后面
    '-----------------------------------------------------
    Set objPopup = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If objPopup Is Nothing Then
        Set objPopup = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    Else
        Call DefCommandPlugInPopup(objPopup.CommandBar.Controls, mrsPlugInBar)
    End If
    Set objPopup = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "路径(&E)", objPopup.Index + 1, False)
    objPopup.ID = conMenu_EditPopup
    With objPopup.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Import, "导入路径(&I)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "取消导入")

        Set objControl = .Add(xtpControlButton, conMenu_Edit_ImportMerge, "导入合并路径")
        objControl.IconId = conMenu_Edit_Import
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_UnImportMerge, "取消合并路径")
        objControl.IconId = conMenu_Edit_Untread
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ViewMergeImport, "查看合并路径导入评估")
        objControl.IconId = conMenu_Edit_Select

        Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "生成路径项目(&C)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Append, "补充生成项目(&S)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_SendBack, "重新生成医嘱")

        Set objControl = .Add(xtpControlButton, conMenu_Edit_Surplus, "添加路径外项目")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改路径外项目")

        Set objControl = .Add(xtpControlButton, conMenu_Edit_Blankoff, "取消本次生成(&X)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "取消当前项目(&V)")

        Set objControl = .Add(xtpControlButton, conMenu_Edit_Archive, "项目执行登记(&E)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_UnArchive, "取消执行登记(&Z)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Merge, "批量执行登记(&B)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_DeleteParent, "批量取消执行(&F)")


        Set objControl = .Add(xtpControlButton, conMenu_Edit_Audit, "评估(&D)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "修改评估")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Clear, "取消评估")


        Set objControl = .Add(xtpControlButton, conMenu_Edit_Stop, "完成路径(&O)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ClearUp, "取消完成")

        Set objControl = .Add(xtpControlButton, conMenu_Edit_OutLogModi, "修改出径登记表")


        Set objControl = .Add(xtpControlButton, conMenu_Edit_Up, "上移")
        objControl.IconId = conMenu_Manage_Up
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Down, "下移")
        objControl.IconId = conMenu_Manage_Down
        '外挂菜单
        Call DefCommandPlugInPopup(objPopup.CommandBar.Controls, mrsPlugInBar)
    End With

    '查看菜单
    '-----------------------------------------------------
    Set objPopup = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    With objPopup.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_View_StPath, "标准路径参考")
        objControl.BeginGroup = True
        objPopup.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Select, "查看导入评估")
        objPopup.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_OutLogView, "查看出径登记表")
        'Set objControl = .Find(, conMenu_View_Refresh) '刷新项前(多个时注意反序)
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Compend, "报告(&R)", objControl.Index + 1)
        objPopup.BeginGroup = True
        objPopup.IconId = conMenu_Manage_Report
        '
        '        Set objControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "观片处理(&V)")
        '        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReportView, "查看检验结果(&R)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_View, "查看项目定义(&A)")
    End With

    '工具菜单:主窗体可能没有,放在帮助菜单前面
    '-----------------------------------------------------
    Set objPopup = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ToolPopup)
    If objPopup Is Nothing Then
        Set objPopup = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_HelpPopup)
        Set objPopup = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "工具(&T)", objPopup.Index, False)
        objPopup.ID = conMenu_ToolPopup
    End If

    '工具栏定义
    '-----------------------------------------------------
    lngStart = 0
    If blnInsideTools Then
        Set cbrToolBar = cbsSub(2)
        Set objControl = cbrToolBar.FindControl(, conMenu_Edit_Import)
        If objControl Is Nothing Then lngStart = 1: cbrToolBar.Visible = True
        For i = cbrToolBar.Controls.count To 1 Step -1
            If cbrToolBar.Controls(i).ID > conMenu_Tool_PlugIn_Item And cbrToolBar.Controls(i).ID < conMenu_Tool_PlugIn_Item + 100 Or cbrToolBar.Controls(i).ID = conMenu_Tool_PlugIn Then
                cbrToolBar.Controls(i).Delete
            End If
        Next i
    Else
        Set cbrToolBar = cbsMain(2)
        For Each objControl In cbrToolBar.Controls    '先求出前面的最后一个Control
            If Val(Left(objControl.ID, 1)) <> conMenu_FilePopup And Val(Left(objControl.ID, 1)) <> conMenu_ManagePopup Then
                Set objControl = cbrToolBar.Controls(objControl.Index - 1): Exit For
            End If
        Next
        lngStart = objControl.Index + 1
    End If

    If lngStart <> 0 Then
        With cbrToolBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Import, "导入", lngStart)
            objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "取消", objControl.Index + 1)

            Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "生成", objControl.Index + 1)
            objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Append, "补充", objControl.Index + 1)
            objControl.ToolTipText = "补充生成可选生成的路径项目"

            Set objControl = .Add(xtpControlButton, conMenu_Edit_Merge, "执行", objControl.Index + 1)
            objControl.BeginGroup = True
            objControl.ToolTipText = "批量执行路径项目"
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Audit, "评估", objControl.Index + 1)
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Stop, "完成", objControl.Index + 1)

            Set objControl = .Add(xtpControlButton, conMenu_Edit_Up, "上移", objControl.Index + 1)
            objControl.IconId = conMenu_Manage_Up
            objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Down, "下移", objControl.Index + 1)
            objControl.IconId = conMenu_Manage_Down
       
            Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Edit_Compend, "报告", objControl.Index + 1)
            objPopup.BeginGroup = True
            objPopup.IconId = conMenu_Manage_Report
            objPopup.ToolTipText = "查阅报告"
        End With

        If blnInsideTools Then
            For Each objControl In cbrToolBar.Controls
                If objControl.Type <> xtpControlLabel Then
                    objControl.Style = xtpButtonIconAndCaption
                End If
            Next
        End If
    End If


    '命令的快键绑定
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        '.Add FCONTROL, Asc("O"), conMenu_File_Open
'        .Add 0, vbKeyF11, conMenu_Tool_Option    '路径选项
    End With
    
    '外挂程序命令加载
    Call DefCommandPlugIn(cbsMain, mrsPlugInBar)
End Sub

Private Sub DefCommandPlugIn(ByVal cbsMain As Object, ByRef rsBar As ADODB.Recordset)
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim strFuncName As String, lngFuncID As Long
    Dim strFunc As String, i As Long
    
    Dim blnGroup As Boolean
    Dim lngTmp  As Long
    
    If rsBar Is Nothing Then Exit Sub
    rsBar.Filter = 0
    If rsBar.RecordCount = 0 Then Exit Sub
 
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ToolPopup)
    '独立按钮
    rsBar.Filter = "IsInTool=1 and BarType=1"
    If Not rsBar.EOF Then
        rsBar.Sort = "序号"
        If Not objMenu Is Nothing Then
            With objMenu.CommandBar.Controls
                For i = 1 To rsBar.RecordCount
                    Set objControl = .Add(xtpControlButton, rsBar!功能ID, rsBar!功能名)
                        objControl.IconId = rsBar!图标ID
                        objControl.Parameter = rsBar!功能名
                        objControl.Style = xtpButtonIconAndCaption
                    If Val(rsBar!IsGroup) = 1 Then
                        objControl.BeginGroup = True
                        blnGroup = True
                    End If
                    rsBar.MoveNext
                Next
            End With
        End If
    End If
    
    '下拉按钮，如果只有一个按钮，也当作独立按钮
    rsBar.Filter = "IsInTool=0 and BarType=1"
    If Not rsBar.EOF Then
        rsBar.Sort = "序号"
        If Not objMenu Is Nothing Then
            With objMenu.CommandBar.Controls
                If rsBar.RecordCount = 1 Then
                    Set objControl = .Add(xtpControlButton, rsBar!功能ID, rsBar!菜单名)
                        objControl.IconId = rsBar!图标ID
                        objControl.Parameter = rsBar!功能名
                        objControl.Style = xtpButtonIconAndCaption
                    If Val(rsBar!IsGroup) = 1 Then
                        objControl.BeginGroup = True
                        blnGroup = True
                    End If
                Else
                    Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_PlugIn, "扩展功能", , False)
                        objPopup.BeginGroup = True
                    With objPopup.CommandBar.Controls
                        For i = 1 To rsBar.RecordCount
                            Set objControl = .Add(xtpControlButton, rsBar!功能ID, rsBar!菜单名)
                            objControl.IconId = rsBar!图标ID
                            objControl.Parameter = rsBar!功能名
                            objControl.Style = xtpButtonIconAndCaption
                            If Val(rsBar!IsGroup) = 1 Then
                                objControl.BeginGroup = True
                                blnGroup = True
                            End If
                            rsBar.MoveNext
                        Next
                    End With
                End If
            End With
        End If
    End If
    
    '工具栏按钮
    If mblnInsideTools Then
        Set objBar = cbsSub(2)
    Else
        Set objBar = cbsMain(2)
    End If
    Set objControl = objBar.FindControl(, conMenu_Help_Help)
    If Not objControl Is Nothing Then
        objControl.BeginGroup = True
        lngTmp = objControl.Index - 1
    Else
        lngTmp = -1
    End If
    rsBar.Filter = "IsInTool=1 and BarType=2"
    If Not rsBar.EOF Then
        With objBar.Controls
            For i = 1 To rsBar.RecordCount
                Set objControl = .Add(xtpControlButton, rsBar!功能ID, rsBar!功能名, lngTmp + 1)
                    objControl.IconId = rsBar!图标ID
                    objControl.Parameter = rsBar!功能名
                    objControl.Style = xtpButtonIconAndCaption
                lngTmp = objControl.Index
                If Val(rsBar!IsGroup) = 1 Then objControl.BeginGroup = True
                rsBar.MoveNext
            Next
            objControl.BeginGroup = True
        End With
    End If
    
    rsBar.Filter = "IsInTool=0 and BarType=2"
    If Not rsBar.EOF Then
        rsBar.Sort = "序号"
        Set objPopup = objBar.Controls.Add(xtpControlPopup, conMenu_Tool_PlugIn, "扩展功能", lngTmp + 1, False)
            objPopup.ID = conMenu_Tool_PlugIn
            objPopup.IconId = conMenu_Tool_PlugIn
            objPopup.BeginGroup = True
            objPopup.Style = xtpButtonIconAndCaption
        lngTmp = objPopup.Index
        With objPopup.CommandBar.Controls
            For i = 1 To rsBar.RecordCount
                Set objControl = .Add(xtpControlButton, rsBar!功能ID, rsBar!菜单名, lngTmp + 1)
                objControl.IconId = rsBar!图标ID
                objControl.Parameter = rsBar!功能名
                If Val(rsBar!IsGroup) = 1 Then objControl.BeginGroup = True
                lngTmp = objPopup.Index
                rsBar.MoveNext
            Next
        End With
    End If
    
    '自动执行的功能
    rsBar.Filter = "IsAuto=1"
    If Not rsBar.EOF Then mlngPlugInID = rsBar!功能ID
End Sub

Private Sub FuncPatiPathPrint()
'功能：输出患者版临床路径
    Dim WordApp As Object   'Word.Application
    Dim WordDoc As Object     'Word.Document
    Dim strSql As String
    Dim rsTmp As Recordset
    Dim strFileName As String, strFilePath As String
    Dim lngRetu As Long, strInfo As String

    If vsPath.FixedRows < 3 Then
        MsgBox "该病人还未生成临床路径项目。", vbInformation, gstrSysName
        Exit Sub
    End If

    On Error GoTo errH
    '获得路径
    Screen.MousePointer = 11
    strSql = "Select 文件名 from 临床路径文件 where 路径ID=[1] And 类别=1 "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.路径ID)
    If rsTmp.RecordCount > 0 Then
        strFileName = rsTmp!文件名 & ""
        strFilePath = gobjFile.GetSpecialFolder(TemporaryFolder) & "\" & strFileName
        If gobjFile.FileExists(strFilePath) Then gobjFile.DeleteFile strFilePath, True
        '将数据库中BLOB数据读到本地临时文件目录下
        strFilePath = Sys.ReadLob(glngSys, 10, mPP.路径ID & "," & strFileName, strFilePath)
        If Not gobjFile.FileExists(strFilePath) Then
            MsgBox "文件内容读取失败！", vbInformation, gstrSysName:
            Screen.MousePointer = 0: Exit Sub
        End If
    Else
        Screen.MousePointer = 0
        MsgBox "该路径表没有设置对应的临床路径表(患者版),请到临床路径管理中设置。", vbInformation, Me.Caption
        Exit Sub
    End If

    Set WordApp = CreateObject("Word.Application")
    If WordApp Is Nothing Then
        MsgBox "请安装Microsoft Office Word。", vbInformation, gstrSysName
        Exit Sub
    End If

    Set WordDoc = WordApp.Documents.Open(strFilePath)      '打开RTF文档
    WordDoc.PrintPreview
    WordApp.Visible = True
    WordApp.ScreenUpdating = True
    WordApp.Activate
    Screen.MousePointer = 0
    
    '记录打印信息
    Call zlDatabase.ExecuteProcedure("Zl_电子病历打印_Insert(" & mPP.病人路径ID & ",12," & mPati.病人ID & "," & mPati.主页ID & ",'" & UserInfo.姓名 & "')", "打印患者版路径表")
    '打印后强制重新加载提示信息，更新提示信息
    Call LoadPathFlow
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncPathTableOutput(bytStyle As Byte, Optional ByVal blnIsSetup As Boolean, Optional ByVal strPDFFile As String, Optional ByVal strDeviceName As String)
'功能：输出临床路径表
'参数：bytStyle=1-打印,2-预览,3-输出到Excel,4-输出到PDF
'     blnIsSetup-批量打印不进行打印前设置
'     strPDFFile=PDF输出默认路径
'     strDeviceName=指定打印机名称
    Dim rsTmp As ADODB.Recordset
    Dim vsBody As VSFlexGrid
    Dim objOut As New zlPrint1Grd
    Dim objRow As zlTabAppRow
    Dim lngColor As Long, bytR As Byte
    Dim strSql As String
    Dim rsSQLTmp As ADODB.Recordset
    Dim strDisease As String        '诊断描述
    Dim strStandardDate As String   '标准住院日
    Dim i As Long, j As Long
    Dim strTitle As String
    Dim strTmp As String
    Dim lngDefDay As Long
    If mbytPrintWay = 1 Then
    
        If bytStyle = 1 Then
            bytStyle = 2
        ElseIf bytStyle = 2 Then
            bytStyle = 1
        End If
        Call FuncPathTableReport(bytStyle)
    Else
        strSql = "Select a.病人id, a.主页id, b.疾病id, b.诊断id, b.诊断描述, c.标准住院日" & vbNewLine & _
                 "From 病人临床路径 A, 病人诊断记录 B, 临床路径版本 C" & vbNewLine & _
                 "Where a.病人id = b.病人id And a.主页id = b.主页id And a.诊断类型 = b.诊断类型" & vbNewLine & _
                 "      And a.诊断来源 = b.记录来源 And c.路径id = a.路径id And c.版本号 = a.版本号 And" & vbNewLine & _
                 "      b.诊断次序 = 1 And a.病人id = [1] And a.主页id = [2] And a.ID=[3]"
    
        mblnUnChange = True
        If vsPath.FixedRows < 3 Then
            '输出PDF，如果不是路径病人，则直接退出不提示
            If bytStyle = 4 Then Exit Sub
            '批量打印不提示
            If blnIsSetup Then Exit Sub
            MsgBox "该病人还未生成临床路径项目。", vbInformation, gstrSysName
            Exit Sub
        End If
        On Error GoTo errH
        Set rsTmp = GetPatiInfo(mPati.病人ID, mPati.主页ID)
        Set rsSQLTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPati.病人ID, mPati.主页ID, mPP.病人路径ID)
    
        If rsSQLTmp.RecordCount > 0 Then
            strDisease = rsSQLTmp!诊断描述 & ""
            strStandardDate = rsSQLTmp!标准住院日 & ""
        Else
            strDisease = ""
            strStandardDate = ""
        End If
        '表头
        If InStr(vsFlow.TextMatrix(0, 0), vbCrLf) > 0 Then
            strTitle = Mid(vsFlow.TextMatrix(0, 0), 1, InStr(vsFlow.TextMatrix(0, 0), vbCrLf) - 1)
        Else
            strTitle = vsFlow.TextMatrix(0, 0)
        End If
        objOut.Title.Text = strTitle & vbCrLf & "临床路径表"
        objOut.Title.Font.Name = "楷体_GB2312"
        objOut.Title.Font.Size = 20
        objOut.Title.Font.Bold = True
    
        '表上
        strSql = "Select a.诊断描述" & vbNewLine & _
                 "From 病人诊断记录 A" & vbNewLine & _
                 "Where a.病人id = [1] And a.主页id = [2] And a.记录来源 = 3 And a.诊断类型 In (2, 12) Order By a.诊断次序"
        Set rsSQLTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPati.病人ID, mPati.主页ID)
        If rsSQLTmp.RecordCount > 0 Then
            strTmp = rsSQLTmp!诊断描述 & ""
            strTmp = Mid(strTmp, InStr(strTmp, ")") + 1) & Mid(strTmp, 1, InStr(strTmp, ")"))
        Else
            strTmp = ""
        End If
        strSql = "Select a.已行手术 || Decode(Nvl(a.诊疗项目id, 0), 0, '(ICD9CM-3:' || b.编码 || ')', '(诊疗项目:' || c.编码 || ')') As 已行手术" & vbNewLine & _
                 "From 病人手麻记录 A, 疾病编码目录 B, 诊疗项目目录 C" & vbNewLine & _
                 "Where a.病人id = [1] And a.主页id = [2] And a.记录来源 = 3 And a.手术操作id = b.Id(+) And a.诊疗项目id = c.Id(+)" & vbNewLine & _
                 "Order By a.手术开始时间"
        Set rsSQLTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPati.病人ID, mPati.主页ID)
    
        If rsSQLTmp.RecordCount > 0 Then
            strTmp = strTmp & " 行 " & rsSQLTmp!已行手术
        End If
        Set objRow = New zlTabAppRow
        objRow.Add "适用对象：第一诊断为 " & strTmp
        objOut.UnderAppRows.Add objRow
    
        Set objRow = New zlTabAppRow
        objRow.Add "患者姓名：" & rsTmp!姓名 & " 性别：" & rsTmp!性别 & " 年龄：" & rsTmp!年龄 & " 住院号：" & rsTmp!住院号 & " 门诊号：" & rsTmp!门诊号 & ""
        objOut.UnderAppRows.Add objRow
    
        Set objRow = New zlTabAppRow
        objRow.Add "住院日期:" & Format(rsTmp!入院日期, "yyyy年MM月dd日")
        objRow.Add "出院日期:" & Format(rsTmp!出院日期, "yyyy年MM月dd日")
        objRow.Add "标准住院日：" & IIf(InStr(strStandardDate, "-") > 0, "", "≤") & strStandardDate & "天"
        objOut.UnderAppRows.Add objRow
        objOut.AppFont.Size = 12
        '表下
        Set objRow = New zlTabAppRow
        objRow.Add "打印人：" & UserInfo.姓名
        objRow.Add "打印日期：" & Format(zlDatabase.Currentdate(), "yyyy年MM月dd日")
        objOut.BelowAppRows.Add objRow
    
        '页脚
        objOut.Footer = ";第[页码]页，共[页数]页;"
        objOut.PageFooter = 5
    
        '表体
        strTmp = zlDatabase.GetPara("路径表单打印规则", glngSys, P临床路径应用, "0")
        If strTmp = "1" Then
            Set vsBody = FuncConvertPathTable
        Else
            Set vsBody = vsPath
        End If
        
        '输出
        With vsBody
            .Redraw = flexRDNone
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = "护士签名"
            .RowHeight(.Rows - 1) = 440
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = "医生签名"
            .RowHeight(.Rows - 1) = 440
            
            '默认打印天数
            lngDefDay = Val(zlDatabase.GetPara("路径表单每页打印的天数", glngSys, P临床路径应用, "2"))
            objOut.PageCols = lngDefDay + .FixedCols
            '表格列数不够时，补充空列
            If (.Cols - 1) Mod lngDefDay <> 0 Then
               .Cols = .Cols + (lngDefDay - ((.Cols - 1) Mod lngDefDay))
            End If
            '打印表格转换
            Call FuncPathTableChange(vsBody, lngDefDay)
           
            '破坏合并列首行,打印部件中对合并列有单独处理
            For i = .FixedCols To .Cols - 1
                If i Mod 2 = 0 Then
                    .TextMatrix(R0阶段名, i) = .TextMatrix(R0阶段名, i) & vbTab
                End If
            Next
            .Redraw = flexRDDirect
            '行宽自适应
            If .FixedRows > 1 Then .AutoSize .FixedCols, .Cols - 1, , 45    '在要Draw之后才生效
    
            objOut.FixCol = vsBody.FixedCols
            objOut.FixRow = vsBody.FixedRows
            Set objOut.Body = vsBody
    
            '指定打印机
            If strDeviceName <> "" Then SaveSetting "ZLSOFT", "公共模块\zl9PrintMode\Default", "DeviceName", strDeviceName
            If bytStyle = 1 Or bytStyle = 4 Then
                If bytStyle = 4 Then
                    bytR = 4
                    objOut.Privileged = True '电子病案查阅 用于公共部件Zl9PrintMode内部跳过打印权限检查
                Else
                    If Not blnIsSetup Then
                        bytR = zlPrintAsk(objOut)
                    Else
                        bytR = 1
                    End If
                End If
                Me.Refresh
                
                If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR, strPDFFile
                '打印完了产生打印记录
                strSql = "zl_电子病历打印_insert(" & mPP.病人路径ID & ",11," & mPati.病人ID & "," & mPati.主页ID & ",'" & UserInfo.姓名 & "')"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            Else
                zlPrintOrView1Grd objOut, bytStyle
            End If
            mblnUnChange = False
            '恢复到初始状态
            Call zlRefresh(mPati.病人ID, mPati.主页ID, mPati.病区ID, mPati.科室ID, mPati.病人状态)    'vsPath变动后重新加载
            
            If vsPathPrint.UBound = 1 Then Unload vsPathPrint(1)
        End With
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncPathTableReport(ByVal bytType As Byte)
'功能:用报表打印临床路径表单
'bytType:0=缺省值,可不传,表示正常(含报表及预览),1=直接到预览,2=直接打印,3-输出到Excel,4-输出到PDF
    Dim arrSQL As Variant
    Dim i As Long, j As Long
    Dim strTmp As String
    
    On Error GoTo errH
    arrSQL = Array()
    With vsPath
        For i = .FixedCols To .Cols - 1
            For j = 0 To .Rows - 1
                strTmp = ""
                If TypeName(.Cell(flexcpData, j, i)) = "String" Then
                    If .Cell(flexcpData, j, i) & "" <> "" Then strTmp = Split(.Cell(flexcpData, j, i), "|")(0) & ""
                End If
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_路径打印记录_Insert(" & ZVal(Val(.ColData(i) & "")) & ",'" & .TextMatrix(j, 0) & "'," & i & "," & j & ",'" & .TextMatrix(j, i) & "'," & ZVal(Val(strTmp)) & ")"
            Next
        Next
    End With
    gcnOracle.BeginTrans
    For i = LBound(arrSQL) To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1256", Me, "病人ID=" & mPati.病人ID, "主页ID=" & mPati.主页ID, "病人路径ID=" & mPP.病人路径ID, bytType)
    gcnOracle.CommitTrans
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetControlVisible(ByVal Control As CommandBarControl)
'功能：根据权限设置菜单和工具栏的可见状态
    Dim blnVisible As Boolean

    '权限只需判断一次,已经判断过的命令不用再判断
    If Control.Category = "已判断" Then Exit Sub

    blnVisible = True
    Select Case Control.ID
    Case conMenu_Edit_Import, conMenu_Edit_Untread, conMenu_Edit_ImportMerge, conMenu_Edit_UnImportMerge
        If InStr(GetInsidePrivs(P临床路径应用), ";导入路径;") = 0 Then blnVisible = False
    Case conMenu_Edit_Send, conMenu_Edit_Append, conMenu_Edit_Delete, conMenu_Edit_Blankoff, conMenu_Edit_SendBack
        If InStr(GetInsidePrivs(P临床路径应用), ";生成路径;") = 0 Then blnVisible = False
        If Control.ID = conMenu_Edit_SendBack And blnVisible Then
            blnVisible = Not InStr(GetInsidePrivs(p住院医嘱下达), ";医嘱下达;") = 0
        End If
    Case conMenu_Edit_Surplus, conMenu_Edit_Modify, conMenu_Edit_Up, conMenu_Edit_Down
        If InStr(GetInsidePrivs(P临床路径应用), ";路径外项目;") = 0 Then blnVisible = False

    Case conMenu_Edit_Archive, conMenu_Edit_UnArchive, conMenu_Edit_Merge, conMenu_Edit_DeleteParent
        If InStr(GetInsidePrivs(P临床路径应用), ";执行路径;") = 0 Or mbln启用执行环节 = False Then blnVisible = False
        '启用路径执行环节时，启用场合和当前场合不一致时,隐藏菜单按钮
        If blnVisible Then
            If Mid(mstr执行场合, mint场合 + 1, 1) = "0" Then
                blnVisible = False
            End If
        End If
    Case conMenu_Edit_Audit, conMenu_Edit_Reuse, conMenu_Edit_Clear
        If InStr(GetInsidePrivs(P临床路径应用), ";阶段评估;") = 0 Then blnVisible = False

    Case conMenu_Edit_Stop, conMenu_Edit_ClearUp
        If InStr(GetInsidePrivs(P临床路径应用), ";结束路径;") = 0 Then blnVisible = False

    Case conMenu_Edit_OutLogModi, conMenu_Edit_OutLogView
        If Control.ID = conMenu_Edit_OutLogModi Then
            If InStr(GetInsidePrivs(P临床路径应用), ";结束路径;") = 0 Then blnVisible = False
        End If
        If blnVisible Then blnVisible = CheckPathOutLog
    Case conMenu_Edit_Compend
        '报告弹出(含打印),查阅报告
        If InStr(GetInsidePrivs(p住院医嘱下达), ";报告查阅;") = 0 Then blnVisible = False
    End Select

    Control.Visible = blnVisible
    Control.Category = "已判断"
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnHaveItem As Boolean
    Dim lng项目ID As Long

    If vsPath.Redraw = flexRDNone Then Exit Sub

    '根据权限设置按钮可见状态
    Call SetControlVisible(Control)
    If Not Control.Visible Then Exit Sub

    With vsPath
        blnHaveItem = .Row > .FixedRows - 1 And .FixedRows <> 0 And .Col > .FixedCols - 1   '.FixedRows=0时，只有一行提示信息
    End With
    Select Case Control.ID
        '0.输出
    Case conMenu_File_PrintSet, conMenu_File_Print, conMenu_File_Preview, conMenu_File_Excel, conMenu_File_Print_PatiPath
        Control.Enabled = mPP.病人路径ID <> 0

        '1.导入
        '-----------------------------------------
    Case conMenu_Edit_Import    '导入路径
        Control.Enabled = mlngState <> ps最近转出 And mlngState <> ps出院 And mPati.病人状态 = 0 And (mPP.病人路径ID = 0 Or mPP.病人路径状态 <> 1) And mPati.病人ID <> 0 And cboPath.ListIndex <= 0

    Case conMenu_Edit_Untread   '取消导入(仅在第一次生成时可取消导入)
        Control.Enabled = mPati.病人状态 = 0 And mPP.病人路径ID <> 0 And (mPP.病人路径状态 = 0 Or mPP.病人路径状态 = 1) And vsPath.Cols <= vsPath.FixedCols + 1
    Case conMenu_Edit_Select      '查看导入评估
        Control.Enabled = mPati.病人状态 = 0 And mPP.病人路径ID <> 0
    Case conMenu_Edit_ImportMerge  '导入合并路径
        Control.Enabled = mPati.病人状态 = 0 And mPP.病人路径ID <> 0 And cboPath.ListIndex <= 0
    Case conMenu_Edit_UnImportMerge    '取消导入合并路径
        Control.Enabled = mPati.病人状态 = 0 And mPP.病人路径ID <> 0 And cboPath.ListIndex <= 0 And mPP.合并路径个数 > 0
    Case conMenu_Edit_ViewMergeImport    '查看合并路径导入评估
        Control.Enabled = mPati.病人状态 = 0 And mPP.病人路径ID <> 0 And mPP.合并路径个数 > 0
        '2.生成
        '-----------------------------------------
    Case conMenu_Edit_Send      '生成路径
        Control.Enabled = mPati.病人状态 = 0 And mPP.病人路径ID <> 0 And mPP.病人路径状态 = 1

    Case conMenu_Edit_Append    '补充生成
        Control.Enabled = mPP.当前阶段ID <> 0 And mPP.病人路径状态 = 1
    Case conMenu_Edit_Blankoff  '取消本次生成
        Control.Enabled = mPP.当前阶段ID <> 0 And mPP.病人路径状态 = 1
    Case conMenu_Edit_Delete, conMenu_Edit_SendBack   '取消路径项目,重新生成医嘱
        Control.Enabled = mPP.当前阶段ID <> 0 And mPP.病人路径状态 = 1
        If Control.Enabled Then
            With vsPath
                If .TextMatrix(.Row, .Col) <> "" And .Row <> .Rows - 1 And .Col > 0 Then
                    Control.Enabled = ((mint场合 = 0 And .ColData(.Col) = mPP.当前阶段ID And .Col = .Cols - 1) _
                        Or (mint场合 = 1 And conMenu_Edit_Delete = Control.ID) _
                        Or (conMenu_Edit_Delete = Control.ID And Split(.Cell(flexcpData, .Row, .Col), "|")(4) = 1))
                Else
                    Control.Enabled = False
                End If
            End With
        End If
    Case conMenu_Edit_Surplus   '添加路径外项目
        Control.Enabled = mPP.当前阶段ID <> 0 And mPP.病人路径状态 = 1

    Case conMenu_Edit_Modify, conMenu_Edit_Up, conMenu_Edit_Down     '修改路径外项目
        Control.Enabled = mPP.当前阶段ID <> 0 And mPP.病人路径状态 = 1
        If Control.Enabled Then
            If vsPath.Cell(flexcpData, vsPath.Row, vsPath.Col) <> "" And vsPath.Row <> vsPath.Rows - 1 Then
                lng项目ID = Split(vsPath.Cell(flexcpData, vsPath.Row, vsPath.Col), "|")(1)    '路径外项目为0
                Control.Enabled = lng项目ID = 0
            Else
                Control.Enabled = False
            End If
        End If

    Case conMenu_Edit_View      '查看项目定义
        Control.Enabled = blnHaveItem
        If Control.Enabled Then
            If vsPath.Cell(flexcpData, vsPath.Row, vsPath.Col) <> "" Then
                lng项目ID = Split(vsPath.Cell(flexcpData, vsPath.Row, vsPath.Col), "|")(1)    '路径外项目为0
                Control.Enabled = lng项目ID <> 0
            End If
        End If


        '3.执行
        '-----------------------------------------
    Case conMenu_Edit_Archive, conMenu_Edit_UnArchive   '单个项目执行(仅最后一次的列才能) '取消执行
        Control.Enabled = mPP.当前阶段ID <> 0 And mPP.病人路径状态 = 1
        If Control.Enabled Then
            With vsPath
                If .TextMatrix(.Row, .Col) <> "" And .Row <> .Rows - 1 And .Row >= .FixedRows Then
                    Control.Enabled = (.ColData(.Col) = mPP.当前阶段ID And mint场合 = 0 And .Col = .Cols - 1) Or mint场合 = 1
                Else
                    Control.Enabled = False
                End If
            End With
        End If
    Case conMenu_Edit_Merge, conMenu_Edit_DeleteParent    '批量执行,批量取消执行
        Control.Enabled = mPP.当前阶段ID <> 0 And mPP.病人路径状态 = 1


        '4.评估
        '-----------------------------------------
    Case conMenu_Edit_Audit     '阶段评估
        Control.Enabled = mPP.当前阶段ID <> 0 And mPP.病人路径状态 = 1 And mint场合 = 0
    Case conMenu_Edit_Reuse     '修改评估
        Control.Enabled = mPP.当前阶段ID <> 0 And mPP.病人路径状态 = 1 And mint场合 = 0
    Case conMenu_Edit_Clear     '取消评估
        Control.Enabled = mPP.当前阶段ID <> 0 And mPP.病人路径状态 = 1 And mint场合 = 0


        '5.完成
        '-----------------------------------------
    Case conMenu_Edit_Stop      '完成路径
        Control.Enabled = mPP.当前阶段ID <> 0 And mPP.病人路径状态 = 1
        If Control.Enabled Then    '当前天数到达标准住院日范围且已评估才允许完成
            Control.Enabled = mblnInOverScope And vsPath.TextMatrix(vsPath.Rows - 1, vsPath.Cols - 1) <> ""
        End If

    Case conMenu_Edit_ClearUp   '取消完成
        If mPP.病人路径状态 = 3 Then
            Control.Caption = "取消退出"
        Else
            Control.Caption = "取消完成"
        End If
        Control.Enabled = (mPP.病人路径状态 = 2 Or mPP.病人路径状态 = 3) And cboPath.ListIndex <= 0    '2-正常完成，3-变异完成

    Case conMenu_Edit_OutLogModi, conMenu_Edit_OutLogView   '出径登记表
        Control.Enabled = (mPP.病人路径状态 = 2 Or mPP.病人路径状态 = 3)     '2-正常完成，3-变异完成
        If Control.ID = conMenu_Edit_OutLogModi And Control.Enabled Then
            Control.Enabled = mlng病案状态 = 0  '提交审核后就不允许修改
        End If

        '6.其它
        '-----------------------------------------
    Case conMenu_Edit_Compend    '查看报告
        With vsPath
            Control.Enabled = blnHaveItem
            If Control.Enabled Then Control.Enabled = .Cell(flexcpData, .Row, .Col) <> ""
        End With
    End Select
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim rsTmp As ADODB.Recordset, str疾病编码 As String
    Dim blnDo As Boolean
    Dim strTmp As String
    
    If mlng婴儿病区ID <> 0 Then
        If mlng婴儿病区ID = mlng医护科室ID Or mlng婴儿科室ID = mlng医护科室ID Then
            MsgBox "该病人已经转出本科室了，只有婴儿留在本科室，不允许操作路径。", vbInformation, Me.Caption
            Exit Sub
        End If
    End If

    Select Case Control.ID
        '0.输出
    Case conMenu_File_PrintSet
        Call zlPrintSet
    Case conMenu_File_Print
        Call FuncPathTableOutput(1)
    Case conMenu_File_Preview
        Call FuncPathTableOutput(2)
    Case conMenu_File_Excel
        Call FuncPathTableOutput(3)
    Case conMenu_File_Print_PatiPath
        '打印患者版路径表
        Call FuncPatiPathPrint
        '1.导入
        '-----------------------------------------
    Case conMenu_Edit_Import    '导入路径
        Call FuncImport
    Case conMenu_Edit_Untread   '取消导入
        Call FuncUnImport
    Case conMenu_Edit_ImportMerge  '导入合并路径
        Call FuncImportMerge
    Case conMenu_Edit_UnImportMerge  '取消导入合并路径
        Call FuncUnImportMerge
    Case conMenu_Edit_Select      '查看导入评估
        Call frmEvaluate.ShowMe(mfrmParent, 0, 0, mPati, mPP)
    Case conMenu_Edit_ViewMergeImport      '查看合并路径导入评估
        Call ViewMergeImport

        '2.生成
        '-----------------------------------------
    Case conMenu_Edit_Send      '生成路径
        Call FuncSendItem
    Case conMenu_Edit_Append    '补充生成
        Call FuncSendItemApend
    Case conMenu_Edit_Delete    '取消已生成的项目
        Call FuncDelItem

    Case conMenu_Edit_Blankoff  '取消本次生成
        Call FuncDelAllItem
    Case conMenu_Edit_SendBack  '重新生成医嘱
        Call FuncReSendItem
    Case conMenu_Edit_Surplus   '添加路径外项目
        Call FuncAppendItem(0)
    Case conMenu_Edit_Modify    '修改路径外项目
        Call FuncAppendItemModify

        '3.执行
        '-----------------------------------------
    Case conMenu_Edit_Archive   '执行路径
        Call FuncExecuteItem
    Case conMenu_Edit_Merge     '批量执行
        Call FuncExecuteAll
    Case conMenu_Edit_UnArchive     '取消执行
        Call FuncExecuteItemCancel
    Case conMenu_Edit_DeleteParent  '批量取消执行
        Call FuncExecuteAllCancel

        '4.评估
        '-----------------------------------------
    Case conMenu_Edit_Audit     '评估
        Call FuncEvaluate
    Case conMenu_Edit_Reuse     '修改评估
        Call FuncReEvaluate
    Case conMenu_Edit_Clear     '取消评估
        Call FuncEvaluateCancel


        '5.完成
        '-----------------------------------------
    Case conMenu_Edit_Stop      '完成路径
        Call FuncOver
    Case conMenu_Edit_ClearUp   '取消完成
        Call FuncOverCancel
    Case conMenu_Edit_OutLogModi    '修改出径登记表
        Call OutLogModi
    Case conMenu_Edit_OutLogView   '查看出径登记表
        Call frmPathOutLog.ShowMe(mfrmParent, mPati.病人ID, mPati.主页ID, 1, Nothing, mPP.路径ID, mPP.病人路径ID)
        '6.上移，下移
        '-----------------------------------------
    Case conMenu_Edit_Up    '1-上移
        Call MovePathItem(1)
    Case conMenu_Edit_Down    '-1-下移
        Call MovePathItem(-1)

        '7.其它
        '-----------------------------------------
    Case conMenu_Edit_Compend * 10# + 1 To conMenu_Edit_Compend * 10# + 10  '限制最多10个报告
        If InStr(Control.Parameter, ":") > 0 Then
            Call FuncViewReport(Split(Control.Parameter, ":")(0), Split(Control.Parameter, ":")(1))
        End If
    Case conMenu_Edit_MarkMap

        'Case conMenu_Manage_ReportView  '查看检验结果

    Case conMenu_Edit_View    '显示路径项目定义的信息
        Call vsPath_DblClick
    Case conMenu_View_StPath    '查看标准路径参考
        Set rsTmp = GetPatiDiagnose(mPati.病人ID, mPati.主页ID, 2)  '获取首要诊断
        If rsTmp.RecordCount <> 0 Then
            str疾病编码 = rsTmp!编码
        End If
        Call frmStPathList.ShowMe(mfrmParent, str疾病编码)
'    Case conMenu_Tool_Option    '路径选项
'        Dim objControl As CommandBarControl
'
'        If InStr(GetInsidePrivs(p临床路径应用), ";参数设置;") = 0 Then
'            MsgBox "你没有参数设置的权限。", vbInformation, gstrSysName
'        Else
'            frmPathSetup.mbytFun = 0
'            frmPathSetup.Show 1, mfrmParent
'
'            strTmp = zlDatabase.GetPara("路径执行环节启用场合", glngSys, p临床路径应用, "11")
'            If strTmp <> mstr执行场合 Then
'                mstr执行场合 = strTmp
'                blnDo = True
'            End If
'
'            If mbln启用执行环节 <> CBool(Val(zlDatabase.GetPara("是否启用路径执行环节", glngSys, p临床路径应用, 0))) Or blnDo Then
'                mbln启用执行环节 = Val(zlDatabase.GetPara("是否启用路径执行环节", glngSys, p临床路径应用, 1))
'                If mblnInsideTools Then
'                    Set objControl = cbsSub.FindControl(, conMenu_Edit_Archive, , True): objControl.Category = ""
'                    Set objControl = cbsSub.FindControl(, conMenu_Edit_Merge, , True): objControl.Category = ""
'                    Set objControl = cbsSub.FindControl(, conMenu_Edit_UnArchive, , True): objControl.Category = ""
'                    Set objControl = cbsSub.FindControl(, conMenu_Edit_DeleteParent, , True): objControl.Category = ""
'                Else
'                    Set objControl = mcbsMain.ActiveMenuBar.FindControl(, conMenu_Edit_Merge, , True): objControl.Category = ""
'                    Set objControl = mcbsMain.FindControl(, conMenu_Edit_Archive, , True): objControl.Category = ""
'                    Set objControl = mcbsMain.FindControl(, conMenu_Edit_Merge, , True): objControl.Category = ""
'                    Set objControl = mcbsMain.FindControl(, conMenu_Edit_UnArchive, , True): objControl.Category = ""
'                    Set objControl = mcbsMain.FindControl(, conMenu_Edit_DeleteParent, , True): objControl.Category = ""
'                End If
'            End If
'            End If
        Case conMenu_Tool_PlugIn_Item + 1 To conMenu_Tool_PlugIn_Item + 99 '外挂功能执行
            If CreatePlugInOK(P临床路径应用) Then
                Call gobjPlugIn.ExecuteFunc(glngSys, P临床路径应用, Control.Parameter, mPati.病人ID, mPati.主页ID, mPP.路径ID, , mint场合)
            End If
    End Select
End Sub

Public Sub zlPopupCommandBars(ByVal CommandBar As CommandBar)
'功能：定义菜单项的弹出菜单
    Dim objControl As CommandBarControl
    Dim rsTmp As ADODB.Recordset, i As Long, j As Long
    Dim rsTmpPacs As Recordset
    
    If CommandBar.Parent Is Nothing Then Exit Sub
    Select Case CommandBar.Parent.ID
        Case conMenu_Edit_Compend
             With CommandBar.Controls
                .DeleteAll
                With vsPath
                    Set rsTmp = GetReportOfPath(Val(Split(.Cell(flexcpData, .Row, .Col), "|")(0)))
                    Set rsTmpPacs = GetPACSReportOfPath(Val(Split(.Cell(flexcpData, .Row, .Col), "|")(0)))
                End With
                
                If rsTmp.RecordCount = 0 And rsTmpPacs.RecordCount = 0 Then
                     .Add xtpControlButton, conMenu_Edit_Compend * 10 + 1, "无报告或未书写"
                Else
                    For i = 1 To rsTmp.RecordCount
                        Set objControl = .Add(xtpControlButton, conMenu_Edit_Compend * 10 + i, rsTmp!病历名称 & "(&" & i & ")")
                        objControl.Parameter = rsTmp!ID & ":" & rsTmp!医嘱id
                        rsTmp.MoveNext
                    Next
                    i = rsTmp.RecordCount
                    For j = 1 To rsTmpPacs.RecordCount
                        Set objControl = .Add(xtpControlButton, conMenu_Edit_Compend * 10 + i + j, rsTmpPacs!文档标题 & "(&" & i + j & ")")
                        objControl.Parameter = rsTmpPacs!报告ID & ":" & rsTmpPacs!医嘱id
                        rsTmpPacs.MoveNext
                    Next
                End If
                
                
            End With
    End Select
End Sub

Private Function GetReportOfPath(ByVal lng路径执行ID As Long) As ADODB.Recordset
'功能：获取路径对应的报告名称
    Dim strSql As String
 
    strSql = "Select d.id, d.病历名称,c.医嘱Id" & vbNewLine & _
            "From 病人路径执行 A, 病人路径医嘱 B, 病人医嘱报告 C, 电子病历记录 D" & vbNewLine & _
            "Where a.Id = [1] And a.Id = b.路径执行id And b.病人医嘱id = c.医嘱Id And c.病历id = d.Id"
    On Error GoTo errH
    Set GetReportOfPath = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng路径执行ID)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetPACSReportOfPath(ByVal lng路径执行ID As Long) As ADODB.Recordset
'功能：获取路径对应的报告名称
    Dim strSql As String
    Dim strIDs As String
    Dim rsTmp As Recordset
 
    strSql = "Select b.病人医嘱id" & vbNewLine & _
            "From 病人路径执行 A, 病人路径医嘱 B" & vbNewLine & _
            "Where a.Id = [1] And a.Id = b.路径执行id "
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng路径执行ID)
    Do While Not rsTmp.EOF
        strIDs = strIDs & "," & rsTmp!病人医嘱id
        rsTmp.MoveNext
    Loop
    strIDs = Mid(strIDs, 2)
    If strIDs <> "" Then
        Call CreateObjectPacs(mobjPublicPACS)
        Set GetPACSReportOfPath = mobjPublicPACS.zlDocGetListWithAdvice(strIDs)
    Else
        Set rsTmp = New Recordset
        rsTmp.Fields.Append "ID", adInteger, 1
        rsTmp.CursorLocation = adUseClient
        rsTmp.LockType = adLockOptimistic
        rsTmp.CursorType = adOpenStatic
        rsTmp.Open
        Set GetPACSReportOfPath = rsTmp
    End If

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CreateObjectPacs(objPublicPACS As Object) As Boolean
    If objPublicPACS Is Nothing Then
        On Error Resume Next
        Set objPublicPACS = CreateObject("zlPublicPACS.clsPublicPACS")
        Err.Clear: On Error GoTo 0
        If Not objPublicPACS Is Nothing Then
            Call objPublicPACS.InitInterface(gcnOracle, UserInfo.姓名)
        End If
        If objPublicPACS Is Nothing Then
            MsgBox "PACS公共部件未创建成功！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    CreateObjectPacs = True
End Function

Private Sub FuncOver()
'功能：完成路径
    Dim strSql As String, blnOK As Boolean, lngValue As Long
    Dim colSQL As New Collection, blnTrans As Boolean, i As Long
    Dim str审核人 As String
    Dim rsTmp As ADODB.Recordset
    Dim lngPPStatus As Long
    
    On Error GoTo errH
    '最后检查护士生成的项目是否存在未执行登记的项目(因为护士没有评估环节)
    If mbln启用执行环节 Then
        If Mid(mstr执行场合, 2, 1) = "1" Then  '护士场合启用执行环节
            strSql = "Select a.执行时间" & vbNewLine & _
                    "From 病人路径执行 A, 临床路径项目 B" & vbNewLine & _
                    "Where a.项目id = b.Id(+) And a.路径记录id = [1] and NVl(NVl(a.执行者,b.执行者),1)=2 and a.执行时间 is null and rownum <2"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID)
            If rsTmp.RecordCount > 0 Then
                MsgBox "存在护士未完成执行登记的路径项目，请先完成执行登记后再完成路径。"
                Exit Sub
            End If
        End If
    End If
    
    '先判断该路径是否允许诊断不同完成路径，不允许则检查出院诊断是否和导入诊断相同
    If mPP.结束路径控制 = 0 Then
        If Not CheckPathOutDiag(mPP.病人路径ID, mPati.病人ID, mPati.主页ID) Then
            MsgBox "出院诊断不在适用病种范围内，不允许正常完成路径，只能变异退出路径。", vbInformation, gstrSysName
            Call FuncReEvaluate
            Exit Sub
        End If
    End If
    '判断是否有未审核的阶段
    If InStr(GetInsidePrivs(P临床路径应用), ";跳转审核;") = 0 Then
        If CheckPathIsTurnAduit Then
            str审核人 = zlDatabase.UserIdentify(Me, "前面阶段存在未审核的路径跳转，必须审核后才允许完成。", glngSys, P临床路径应用, "跳转审核")
            If str审核人 = "" Then Exit Sub
        End If
    Else
        str审核人 = UserInfo.姓名
    End If
            
    If MsgBox("你确定要完成当前病人的临床路径吗?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
        Exit Sub
    End If
    If CheckPathOutLog Then
        blnOK = frmPathOutLog.ShowMe(mfrmParent, mPati.病人ID, mPati.主页ID, 0, colSQL, mPP.路径ID, mPP.病人路径ID)
        If blnOK = False Then
            lngValue = Val(zlDatabase.GetPara("必须填写出径登记表", glngSys, P临床路径应用, "0"))
            If lngValue = 1 Then
                MsgBox "由于完成路径前必须填写出径登记表，你取消了填写，路径完成操作未执行。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    End If
    
    lngPPStatus = mPP.病人路径状态
    
    strSql = "Zl_病人路径结束_Update(" & mPP.病人路径ID & ",'" & str审核人 & "')"
    gcnOracle.BeginTrans: blnTrans = True
        For i = 1 To colSQL.count
            '执行出径登记表的SQL
            Call zlDatabase.ExecuteProcedure(colSQL("C" & i), "出径登记表")
        Next
        Call zlDatabase.ExecuteProcedure(strSql, "取消路径完成")
    gcnOracle.CommitTrans: blnTrans = False
    
    Call zlRefresh(mPati.病人ID, mPati.主页ID, mPati.病区ID, mPati.科室ID, mPati.病人状态)
    
    '当前病人路径状态发生变化时更新Lis病人路径状态
    If lngPPStatus <> mPP.病人路径状态 Then
        If Not gobjLIS Is Nothing Then
           Call gobjLIS.ModifyPathState(mPati.病人ID, mPati.主页ID, mPP.病人路径状态)
        End If
    End If
    
    RaiseEvent RequestRefresh(mPP.病人路径状态)
    
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncOverCancel()
'功能：取消路径的完成
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim lngPPStatus As Long
    
    On Error GoTo errH
    '完成后没有新增可用的医嘱才允许取消
    strSql = "Select Null" & vbNewLine & _
            "From 病人临床路径 A, 病人医嘱记录 B" & vbNewLine & _
            "Where a.Id = [1] And a.病人id = b.病人id And a.主页id = b.主页id And b.开嘱时间 > Trunc(a.结束时间, 'MI') And b.医嘱状态 Not In (-1, 4) And Nvl(b.婴儿,0)=0 And Rownum < 2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID)
    If rsTmp.RecordCount > 0 Then
        MsgBox "路径完成后已产生了新的医嘱，请删除或作废后再进行取消操作。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If mPP.病人路径状态 = 3 Then
        If MsgBox("当前路径是变异后自动完成的，取消后变异评估结果将同时删除，并且取消评估，你确定要继续吗?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    
    lngPPStatus = mPP.病人路径状态
    
    strSql = "Zl_病人路径结束_Delete(" & mPP.病人路径ID & "," & mPP.病人路径状态 & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "取消路径完成")
    Call zlRefresh(mPati.病人ID, mPati.主页ID, mPati.病区ID, mPati.科室ID, mPati.病人状态)
    '当前病人路径状态发生变化时更新Lis病人路径状态
    If lngPPStatus <> mPP.病人路径状态 Then
        If Not gobjLIS Is Nothing Then
           Call gobjLIS.ModifyPathState(mPati.病人ID, mPati.主页ID, mPP.病人路径状态)
        End If
    End If
    
    RaiseEvent RequestRefresh(mPP.病人路径状态)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncExecuteItem()
'功能：执行路径项目
    Dim lng执行ID As Long, lng项目ID As Long
    Dim rsTmp As ADODB.Recordset, strSql As String
    
    With vsPath
        lng执行ID = Split(.Cell(flexcpData, .Row, .Col), "|")(0)
        lng项目ID = Split(.Cell(flexcpData, .Row, .Col), "|")(1)    '路径外项目为0
    End With
        
    
    strSql = "Select 1 From 病人路径执行 Where ID = [1] And 执行时间 is Not Null"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng执行ID)
    If rsTmp.RecordCount > 0 Then
        MsgBox "该项目已执行。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '身份检查
    If lng项目ID <> 0 Then
        strSql = "Select 执行者 From 临床路径项目 Where ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng项目ID)
    Else
        strSql = "Select 执行者 From 病人路径执行 Where ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng执行ID)
    End If
    '根据场合判断,不根据人员性质
    If (mint场合 = 0 And rsTmp!执行者 = 2) Or (mint场合 = 1 And rsTmp!执行者 = 1) Then
        MsgBox "该项目只能由" & IIf(rsTmp!执行者 = 1, "医生", "护士") & "执行。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If frmPathExecute.ShowMe(mfrmParent, 1, mPati, mPP, lng执行ID, mint场合) Then
        Call zlRefresh(mPati.病人ID, mPati.主页ID, mPati.病区ID, mPati.科室ID, mPati.病人状态)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncExecuteAll()
'功能：批量执行
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    If mbln启用不评估 Then
'        strSQL = "Select * From (" & _
'            "Select Distinct a.阶段id, a.日期, a.天数, a.登记时间" & vbNewLine & _
'            "From 病人路径执行 A, 临床路径项目 B" & vbNewLine & _
'            "Where a.项目id = b.Id(+) And a.路径记录id = [1] And Nvl(a.生成时间性质, 0) = 0 And Nvl(Nvl(a.执行者, b.执行者), 0) = " & IIf(mint场合 = 0, 1, 2) & " And" & vbNewLine & _
'            "      a.执行时间 Is Null " & vbNewLine & _
'            "Order By a.登记时间) where Rownum <2 "
'        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mPP.病人路径ID)
'        If rsTmp.RecordCount > 0 Then
'            mPP.当前阶段ID = rsTmp!阶段ID
'            mPP.当前日期 = rsTmp!日期
'            mPP.当前天数 = rsTmp!天数
'        End If
        GetPathCurrPhase 1, mPP.当前阶段ID, mPP.当前天数, mPP.当前日期
    End If
    
'    If mint场合 = 1 Then
'        Call GetPhaseInNurse(0, mPP.当前阶段ID, mPP.当前天数, mPP.当前日期)
'    End If
    
    If frmPathExecute.ShowMe(mfrmParent, 0, mPati, mPP, 0, mint场合) Then
        Call zlRefresh(mPati.病人ID, mPati.主页ID, mPati.病区ID, mPati.科室ID, mPati.病人状态)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncExecuteItemCancel()
'功能：取消路径项目的执行
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim lng执行ID As Long
    Dim blnTip As Boolean
    
    With vsPath
        lng执行ID = Split(.Cell(flexcpData, .Row, .Col), "|")(0)
    End With
        
    strSql = "Select 1 From 病人路径执行 Where ID = [1] And 执行时间 is Null"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng执行ID)
    If rsTmp.RecordCount > 0 Then
        MsgBox "该项目还未执行。", vbInformation, gstrSysName
        Exit Sub
    End If
    '医生执行的项目只能医生取消,护士执行的只能护士取消
    strSql = "Select 1 From 病人路径执行 A,临床路径项目 B Where A.ID=[1] And A.项目ID=B.ID(+) " & _
            "And NVL(NVL(A.执行者,B.执行者),1)=" & IIf(mint场合 = 0, 1, 2) & " And A.执行时间 is Not Null"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng执行ID)
    If rsTmp.RecordCount = 0 Then
        MsgBox "该项目由" & IIf(mint场合 = 0, "护士", "医生") & "执行登记,您不能取消。", vbInformation, gstrSysName
        Exit Sub
    End If

    strSql = "Select 1 From 病人路径评估 Where 路径记录ID = [1] And 阶段ID = [2] And 日期 = [3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, Val(vsPath.ColData(vsPath.Col)), CDate(vsPath.Cell(flexcpData, EFixedRow.R2日期, vsPath.Col)))
    If rsTmp.RecordCount > 0 Then
        '强制取消评估，不检查权限
        If mint场合 = 0 Then
            If MsgBox("该病人在" & mPP.当前日期 & "已进行了评估，必须取消评估后才能取消执行。" & vbCrLf & vbCrLf & "你现在要取消评估吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                Call FuncEvaluateCancel(False, False)
            Else
                Exit Sub
            End If
        ElseIf mint场合 = 1 Then
            '检查路径生成者
            If CheckPathSendByNurse(2, lng执行ID) Then
                blnTip = True
            Else
                MsgBox "该项目是医生生成的项目在" & mPP.当前日期 & "已进行了评估。" & vbCrLf & vbCrLf & "您不能取消执行。", vbInformation + vbOKOnly, gstrSysName
                Exit Sub
            End If
        End If
    Else
        blnTip = True
    End If
    
    If blnTip Then
        If MsgBox("你确定要取消[" & vsPath.TextMatrix(vsPath.Row, vsPath.Col) & "]的执行吗？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    
    strSql = "Zl_病人路径执行_Delete(" & lng执行ID & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "取消路径项目")
    Call zlRefresh(mPati.病人ID, mPati.主页ID, mPati.病区ID, mPati.科室ID, mPati.病人状态)

    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function FuncExecuteAllCancel(Optional blnRefresh As Boolean = True) As Boolean
'功能：批量取消路径项目的执行
'说明：护士生成的项目可以不检查评估环节。（医生站评估的时候会检查医生生成者的执行登记情况）
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim blnDo As Boolean
    Dim blnNurse As Boolean    'Ture -存在护士生成的项目 ;False-不不存在护士生成

    On Error GoTo errH
    
    
    '是否存在已经执行登记的项目,取消本次生成时,如果执行登记场合只启用护士时,医生强制取消本次生成时,会因为当前不存在医生执行登记的项目而禁止退出
    If blnRefresh = True Then
        If mbln启用不评估 Then
            GetPathCurrPhase 2, mPP.当前阶段ID, mPP.当前天数, mPP.当前日期
        End If
        strSql = "Select 1 From 病人路径执行 A,临床路径项目 B Where A.路径记录ID = [1] And A.阶段ID = [2] And A.天数 = [3] And A.项目ID=B.ID(+) " & _
                "And NVL(NVL(A.执行者,B.执行者),1)=" & IIf(mint场合 = 0, 1, 2) & " And A.执行时间 is Not Null And Rownum<2"
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, mPP.当前阶段ID, mPP.当前天数)
        If rsTmp.RecordCount = 0 Then
            MsgBox "当前不存在由" & IIf(mint场合 = 0, "医生", "护士") & "执行登记的任何项目。", vbInformation, gstrSysName
            FuncExecuteAllCancel = True
            Exit Function
        End If
    End If
    
    '评估环节检查
    strSql = "Select 1 From 病人路径评估 Where 路径记录ID = [1] And 阶段ID = [2] And 日期 = [3]"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, mPP.当前阶段ID, CDate(mPP.当前日期))
    If rsTmp.RecordCount > 0 Then
        If mint场合 = 1 Then
            strSql = "Select 1 " & vbNewLine & _
                    "From 病人路径执行 A, 临床路径项目 B" & vbNewLine & _
                    "Where a.路径记录id = [1] And a.阶段id = [2] And a.日期 = [3] And a.项目id = b.Id(+) And Nvl(Nvl(a.生成者, b.生成者), 1) = 2 and rownum<2 "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, mPP.当前阶段ID, CDate(mPP.当前日期))
            If CheckPathSendByNurse(1, mPP.病人路径ID, mPP.当前阶段ID, CDate(mPP.当前日期)) Then
                '存在护士生成的路径项目
                blnNurse = True
            Else
                MsgBox "该病人在" & mPP.当前日期 & "已进行了评估,您不能取消由医生生成的项目。", vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
        ElseIf mint场合 = 0 Then
            '强制取消评估，不检查权限
            If MsgBox("该病人在" & mPP.当前日期 & "已进行了评估，必须取消评估后才能取消执行。" & vbCrLf & vbCrLf & "你现在要取消评估吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                Call FuncEvaluateCancel(False, True)
            Else
                Exit Function
            End If
        End If
    End If
 
    blnDo = frmPathExecute.ShowMe(mfrmParent, 2, mPati, mPP, 0, mint场合, blnNurse)
    If blnDo And blnRefresh Then
        Call zlRefresh(mPati.病人ID, mPati.主页ID, mPati.病区ID, mPati.科室ID, mPati.病人状态)
    End If
    
    FuncExecuteAllCancel = blnDo
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Function

Private Function CheckSameDayOfPhaseTurn() As Boolean
'功能：检查当前路径是否刚刚跳转，是则检查当天是否有可用的阶段
    Dim strSql As String, rsTmp As Recordset
    
    strSql = "select 原路径ID,原路径版本 from 病人路径评估  where 路径记录id=[1] and 阶段ID=[2] and 天数=[3]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, mPP.当前阶段ID, mPP.当前天数)
    If rsTmp.RecordCount > 0 Then
        If rsTmp!原路径ID & "" <> "" Then
            strSql = "Select 1 From 临床路径阶段 Where 路径ID=[1] And 版本号=[2] and [3] Between 开始天数 And Nvl(结束天数, 开始天数)  And rownum<2 and 分支ID is null And 父ID is null"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.路径ID, mPP.版本号, mPP.当前天数)
            CheckSameDayOfPhaseTurn = rsTmp.RecordCount > 0
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function FuncSendItem(Optional ByRef blnIsCancel As Boolean, Optional ByVal lngType As Long) As Boolean
'功能：执行生成路径(通过clsDockPath中的接口开放给医嘱部件调用)
'参数：blnIsCancel，没有路径可生成时，用户是否取消了评估。true=取消
'     lngType:1-医嘱编辑界面调用，则评估后不继续生成，因为医嘱编辑界面不能再调用医嘱编辑。
    Dim rsTmp As ADODB.Recordset
    '-------
    Dim lng天数 As Long, lng时间进度 As Long, lng理论天数 As Long
    Dim lng阶段ID As Long
    Dim lngPPStatus As Long
    Dim i As Long
    '-------
    Dim strTmp As String
    Dim strSql As String
    Dim strDate As String
    Dim strPhase As String
    Dim strMsg As String
    '-------
    Dim blnDo As Boolean
    Dim blnIsNext As Boolean
    Dim blnEvaluate As Boolean
    Dim blnRefresh As Boolean
    Dim blnTrans As Boolean
    
    Dim DatCurr As Date
    Dim colSQL As Collection
    
    On Error GoTo errH
    
LineBegin:
    If mint场合 = 1 Then
        '护士场合生成前强制刷新，避免医生生成或取消生成操作时，护士站未能同步更新的情况
        Call zlRefresh(mPati.病人ID, mPati.主页ID, mPati.病区ID, mPati.科室ID, mPati.病人状态)
    End If

    If mPP.当前天数 = 0 Then '第一天
        '护士场合
        If mint场合 = 1 Then
            MsgBox "医生还没有生成任何路径项目,护士不能提前生成。", vbInformation, gstrSysName
            Exit Function
        End If
            
        strSql = "Select To_number(Trunc(Sysdate)-Trunc(a.开始时间)+1) as 入院天数,Nvl(b.确诊天数,0) as 确诊天数,a.开始时间 as 入院时间" & _
                " From 病人临床路径 a,临床路径目录 b Where a.ID = [1] And a.路径id = b.id"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID)
        DatCurr = zlDatabase.Currentdate
        If rsTmp!确诊天数 > 0 And DatCurr > Format(DateAdd("d", Val(rsTmp!确诊天数), rsTmp!入院时间), "yyyy-MM-DD HH:mm:ss") Then
            MsgBox "该病人已入院" & rsTmp!入院天数 & "天，超过了规定的确诊天数(" & rsTmp!确诊天数 & "天)，不允许生成路径。", vbInformation, gstrSysName
            Exit Function
        End If
        If mPP.导入时间 <> CDate(0) Then
            '导入了路径后首次生成且生成日期大于导入日期
            DatCurr = zlDatabase.Currentdate
            If Int(DatCurr) - Int(mPP.导入时间) >= 1 Then
                Set colSQL = New Collection
                Call CreatePathItem(DatCurr, mPP.导入时间, mPati, mPP, mPP.病人路径ID, colSQL)
                If colSQL.count > 0 Then
                    gcnOracle.BeginTrans: blnTrans = True
                    For i = 1 To colSQL.count
                        Call zlDatabase.ExecuteProcedure(colSQL("C" & i), "路径生成")
                    Next
                    gcnOracle.CommitTrans: blnTrans = False
                    '强制刷新
                    Call zlRefresh(mPati.病人ID, mPati.主页ID, mPati.病区ID, mPati.科室ID, mPati.病人状态)
                    GoTo LineBegin
                End If
            End If
        End If
        lng天数 = rsTmp!入院天数
        lng时间进度 = 0
    Else
        If mint场合 = 1 Then
            '护士场合生成路径项目检查
            '1)当前路径医生未生成任何路径项目,护士不允许生成
            '2)医生还没有生成下一阶段,护士没有路径项目可以生成
            '获取护士生成最后阶段及天数
            strSql = "Select 1 from 病人路径执行 A where A.路径记录ID=[1] And A.阶段ID=[2] And A.天数=[3] And NVL(a.生成者,1) =2 And RowNum<2 "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "生成路径项目", mPP.病人路径ID, mPP.当前阶段ID, mPP.当前天数)
            
            If rsTmp.RecordCount > 0 Then
                MsgBox "该病人在当天的路径已生成。", vbInformation, gstrSysName
                Exit Function
            Else
                Call GetPhaseInNurse(1, mPP.当前阶段ID, mPP.当前天数, mPP.当前日期, , lng天数, strPhase)
            End If
            lng时间进度 = 0
            
            If Not CheckPathIsExecuted(blnRefresh) Then
                '强制刷新
                If blnRefresh Then
                    Call zlRefresh(mPati.病人ID, mPati.主页ID, mPati.病区ID, mPati.科室ID, mPati.病人状态)
                End If
                Exit Function
            End If
        End If
        
        If mint场合 = 0 Then
            '2.当前未评估，不允许生成新的;护士场合没有评估环节
            strSql = "Select 时间进度 From 病人路径评估 Where 路径记录ID = [1] And 阶段ID = [2] And 日期 = [3]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, mPP.当前阶段ID, CDate(mPP.当前日期))
            If Not mbln启用不评估 Then
                If rsTmp.RecordCount = 0 Then
                    If InStr(GetInsidePrivs(P临床路径应用), ";阶段评估;") = 0 Then
                        MsgBox "该病人在" & mPP.当前日期 & "还没有进行评估，不能进行后续操作。", vbInformation, gstrSysName
                        Exit Function
                    Else
                        If MsgBox("该病人在" & mPP.当前日期 & "还没有进行评估，必须先评估。" & vbCrLf & "你现在要进行评估操作吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                            '评估前需先检查执行登记情况
                            If Not CheckPathIsExecuted() Then
                                Exit Function
                            End If
                            '
                            If frmEvaluate.ShowMe(mfrmParent, 1, 1, mPati, mPP) = False Then
                                Exit Function
                            Else
                                lngPPStatus = mPP.病人路径状态
                                Call zlRefresh(mPati.病人ID, mPati.主页ID, mPati.病区ID, mPati.科室ID, mPati.病人状态)
                                '当前病人路径状态发生变化时更新Lis病人路径状态
                                If lngPPStatus <> mPP.病人路径状态 Then
                                    If Not gobjLIS Is Nothing Then
                                       Call gobjLIS.ModifyPathState(mPati.病人ID, mPati.主页ID, mPP.病人路径状态)
                                    End If
                                End If
                                '评估后，可能结束或退出路径，所以根据评估有的状态进行判断是否要继续生成,退出或完成则不继续生成
                                If mPP.病人路径状态 <> 1 Or lngType = 1 Then
                                    Exit Function
                                End If

                                Call zlRefresh(mPati.病人ID, mPati.主页ID, mPati.病区ID, mPati.科室ID, mPati.病人状态)
                                strSql = "Select 时间进度 From 病人路径评估 Where 路径记录ID = [1] And 阶段ID = [2] And 日期 = [3]"
                                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, mPP.当前阶段ID, CDate(mPP.当前日期))
                                If rsTmp.RecordCount <> 0 Then
                                    lng时间进度 = Val("" & rsTmp!时间进度): blnEvaluate = True
                                Else
                                    Exit Function
                                End If
                                
                                blnIsNext = True
                            End If
                        Else
                            Exit Function
                        End If
                    End If
                Else
                    lng时间进度 = Val("" & rsTmp!时间进度): blnEvaluate = True
                End If
            Else
                '当未启用必须评估时,如果用户已经评估了,就按评估的时间进度（提前/延后/正常）继续下一阶段;若未评估，则缺省时间进度 0-正常
                If rsTmp.RecordCount <> 0 Then
                    lng时间进度 = Val("" & rsTmp!时间进度): blnEvaluate = True
                End If
            End If
            strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
            If lng时间进度 = 0 Then
                If mPP.当前日期 = strDate Then
                    lng理论天数 = GetMustDay(mPP.病人路径ID, mPP.当前天数)
                    'a.如果当天还有其它阶段，允许生成其他阶段，但天数仍是当天
                    If CheckSameDayOfPhase(mPP.当前阶段ID, lng理论天数) Then
                        lng天数 = mPP.当前天数
                    Else
                        '检查当前路径是否刚刚跳转，是则检查当天是否有可用的阶段
                        If CheckSameDayOfPhaseTurn Then
                            lng天数 = mPP.当前天数
                        Else
                            blnDo = False
                            If mbln启用提前生成 Then
                                'c.提前生成后续阶段
                                If MsgBox("该病人在当天的路径项目已生成，你现在要提前生成下一天的路径项目吗？", vbYesNo + vbDefaultButton1 + vbQuestion, gstrSysName) = vbYes Then
                                    If CheckSendOfBefore() Then
                                        lng天数 = mPP.当前天数 + 1: blnDo = True
                                    Else
                                        Exit Function
                                    End If
                                Else
                                    Exit Function
                                End If
                            Else
                                MsgBox "该病人当天没有其他可用的阶段可以生成。", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
                                Exit Function
                            End If
                            
                            If Not blnDo And blnEvaluate Then
                                '如果没有后续阶段了，用户又是刚刚评估，则直接退出
                                If blnIsNext Then Exit Function
                                If MsgBox("该病人在今天已进行了评估，要补充医嘱需先取消评估，你现在要取消评估吗？", vbYesNo + vbDefaultButton1 + vbQuestion, "是否取消评估？") = vbYes Then
                                    Call FuncEvaluateCancel(False, True)
                                    blnIsCancel = True
                                End If
                                Exit Function
                            End If
                        End If
                    End If
                ElseIf mPP.当前日期 < strDate Then
                    'b.之前的天数没有生成，则补充生成
                    lng天数 = mPP.当前天数 + 1
                Else 'c.提前生成后续阶段
                    If mbln启用提前生成 Then
                        If CheckSendOfBefore() Then
                            lng天数 = mPP.当前天数 + 1: blnDo = True
                        Else
                            Exit Function
                        End If
                    Else
                        Exit Function
                    End If
                End If
            ElseIf lng时间进度 = 1 Then '下一阶段提前至今天(时间不变，同一天生成多个阶段的内容)
                lng天数 = mPP.当前天数
            ElseIf lng时间进度 = 2 Then '下一阶段提前至明天
                If mPP.当前日期 = strDate Then
                    MsgBox "上一阶段评估为“下一阶段提前至明天”,请明天再生成下一阶段的路径项目。", vbOKOnly + vbInformation, gstrSysName
                    Exit Function
                End If
                lng天数 = mPP.当前天数 + 1
            Else    '下一阶段延后(继续当前阶段)
                If mPP.当前日期 = Format(zlDatabase.Currentdate, "yyyy-MM-dd") Then
                    MsgBox "该病人在今天的路径已生成。", vbInformation, gstrSysName
                    Exit Function
                End If
                lng天数 = mPP.当前天数 + 1
            End If
        End If
    End If

    If frmPathSend.ShowMe(mfrmParent, 0, mint场合, mPati, mPP, mPP.当前阶段ID, lng天数, 0, 0, lng时间进度, mclsMipModule, blnDo, strPhase) Then
        FuncSendItem = True
        Call zlRefresh(mPati.病人ID, mPati.主页ID, mPati.病区ID, mPati.科室ID, mPati.病人状态)
    End If
    
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub FuncSendItemApend()
'功能：补充生成路径
'      当天的路径已生成时才允许补充
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim i As Long
    Dim strTmp As String
    Dim strDate As String
    
    On Error GoTo errH
    If mint场合 = 1 Then
        '护士场合根据选择阶段补充生成
        Call GetPhaseInNurse(0, mPP.当前阶段ID, mPP.当前天数)
    End If
    
    strSql = "Select Max(ID) as ID From 病人路径执行 Where 路径记录ID = [1] And 阶段ID = [2] And 天数 = [3]"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, mPP.当前阶段ID, mPP.当前天数)
    If IsNull(rsTmp!ID) Then
        MsgBox "该病人在今天的路径还没有生成。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If mint场合 = 0 Then '医生才对评估环节做判断，护士忽略评估环节
        strSql = "Select 1 From 病人路径评估 Where 路径记录ID = [1] And 阶段ID = [2] And 日期 = [3]"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, mPP.当前阶段ID, CDate(mPP.当前日期))
        If rsTmp.RecordCount > 0 Then
            If InStr(GetInsidePrivs(P临床路径应用), ";阶段评估;") = 0 Then
                MsgBox "该病人在" & mPP.当前日期 & "已进行了评估，不能再补充生成项目。", vbInformation, gstrSysName
                Exit Sub
            Else
                '取消评估
                If MsgBox("该病人在" & mPP.当前日期 & "已进行了评估，必须取消评估后才能补充生成项目。" & vbCrLf & vbCrLf & "你现在要取消评估吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    Call FuncEvaluateCancel(False, False)
                Else
                    Exit Sub
                End If
            End If
        End If
    End If
    If frmPathSend.ShowMe(mfrmParent, 1, mint场合, mPati, mPP, mPP.当前阶段ID, mPP.当前天数, , , , mclsMipModule) Then
        Call zlRefresh(mPati.病人ID, mPati.主页ID, mPati.病区ID, mPati.科室ID, mPati.病人状态)
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncReSendItem()
'功能：重新生成路径项目的医嘱
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim lng执行ID As Long, lng项目ID As Long, blnMust As Boolean, lng天数 As Long
            
    With vsPath
        lng执行ID = Split(.Cell(flexcpData, .Row, .Col), "|")(0)
        lng项目ID = Val(Split(.Cell(flexcpData, .Row, .Col), "|")(1))
    End With
    If lng项目ID = 0 Then
        MsgBox "要重新生成路径外项目，请取消该项目的生成后重新添加。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    On Error GoTo errH
    
    '1.已经执行的不允许重新生成
    strSql = "Select a.执行时间,c.内容要求  From 病人路径执行 a,临床路径医嘱 b,临床路径项目 C Where a.ID = [1] And a.项目id = b.路径项目id And a.项目ID = c.ID And rownum<2"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng执行ID)
    If rsTmp.RecordCount > 0 Then
        If Not IsNull(rsTmp!执行时间) And mbln启用执行环节 Then
            If rsTmp.RecordCount > 0 Then
                MsgBox "该项目已执行，不能重新生成。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        If Val("" & rsTmp!内容要求) = 0 Then
            strSql = "Select 1" & vbNewLine & _
                    "From 病人路径医嘱 A, 病人路径医嘱 B" & vbNewLine & _
                    "Where a.路径执行id = [1] And a.病人医嘱id = b.病人医嘱id And b.路径执行id <> a.路径执行id  And rownum<2"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "检查医嘱", lng执行ID)
            If rsTmp.RecordCount > 0 Then
                MsgBox "该项目对应的医嘱是根据上次的长嘱生成的，但不是可重选生成的，不能执行重新生成。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    Else
        MsgBox "该项目不是医嘱类项目，不能重新生成。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '2.检查医嘱
    If mint场合 = 1 Then
        '对于已经过审核的医嘱，不允许修改删除。
        strSql = "Select 1 From 病人路径医嘱 B, 病人医嘱记录 C Where b.路径执行id = [1] And b.病人医嘱id = c.Id And c.开嘱医生 Like '%/%' And rownum<2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "检查医嘱", lng执行ID)
        If rsTmp.RecordCount > 0 Then
            MsgBox "该项目对应的医嘱已经过医生审核，不能执行此操作。", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    If frmPathSend.ShowMe(mfrmParent, 3, mint场合, mPati, mPP, mPP.当前阶段ID, mPP.当前天数, lng项目ID, lng执行ID, , mclsMipModule) Then
        Call zlRefresh(mPati.病人ID, mPati.主页ID, mPati.病区ID, mPati.科室ID, mPati.病人状态)
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncDelPhaseItem()
'功能：强制删除最后一天所有的执行项目(用于测试时清除数据)
    Dim strSql As String
    Dim lng执行ID As Long
    Dim i As Long
        
    On Error GoTo errH
    With vsPath
        For i = .FixedRows To .Rows - 2     '最后一行是评估
            If .TextMatrix(i, .Cols - 1) <> "" Then
                lng执行ID = Split(.Cell(flexcpData, i, .Cols - 1), "|")(0)
                strSql = "Zl_病人路径生成_Delete(" & lng执行ID & ")"
                Call zlDatabase.ExecuteProcedure(strSql, "取消路径项目")
            End If
        Next
    End With
    Call zlRefresh(mPati.病人ID, mPati.主页ID, mPati.病区ID, mPati.科室ID, mPati.病人状态)
    Exit Sub
 Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function FuncDelAllItem(Optional ByVal blnRefresh As Boolean = True, Optional ByVal blnPrompt As Boolean = True) As Boolean
'功能：整体取消本次生成的所有路径项目
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim i As Long, strIDs As String, strIDSQL As String, blnTrans As Boolean
    Dim strNewIDs As String
    Dim blnExecuted As Boolean
    Dim dat导入时间 As Date
    Dim lng天数 As Long
    
    If blnPrompt Then
        If mint场合 = 0 Then
            If MsgBox("取消生成将删除路径项目对应的医嘱和病历文件。" & vbCrLf & "你确实要取消本次生成的所有路径项目吗?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                Exit Function
            End If
        Else
            If MsgBox("取消生成将删除路径所有护理类项目。" & vbCrLf & "你确实要取消本次生成的所有护理项目吗?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
    If mint场合 = 1 Then
        Call GetPhaseInNurse(0, mPP.当前阶段ID, mPP.当前天数)
    End If
    
    On Error GoTo errH
    
    strSql = "Select A.ID,A.执行时间,NVL(NVL(A.执行者,B.执行者),1) as 执行者,NVL(NVl(A.生成者,B.生成者),1) as 生成者 From 病人路径执行 A,临床路径项目 B Where A.路径记录ID = [1] And A.阶段ID = [2] And A.天数 = [3] and A.项目ID=B.ID(+) "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, mPP.当前阶段ID, mPP.当前天数)
    If mint场合 = 0 Then
'   医生站取消本次生成时,不考虑护士是否已经生成项目或医生生成项目由护士执行登记的情况。（原因：降低医护之间的关联）
'        rsTmp.Filter = "生成者 =2"
'        If rsTmp.RecordCount > 0 Then
'            MsgBox "护士已生成护理类项目，请先通知护士取消生成。", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
'            Exit Function
'        End If
'        医生生成的项目不管护士是否登记 都允许医生强制取消
'        rsTmp.Filter = "执行者=2 and 生成者 =1"
'        If rsTmp.RecordCount > 0 Then
'            If Not IsNull(rsTmp!执行时间) Then
'                MsgBox "存在医生生成的项目被护士执行登记，请先通知护士取消执行登记。", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
'                Exit Function
'            End If
'        End If
    Else
        rsTmp.Filter = "生成者 = 2"
        If rsTmp.RecordCount = 0 Then
            MsgBox "当前阶段没有生成过护理类项目。", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
            Exit Function
        End If
    End If
    
    rsTmp.Filter = IIf(mint场合 = 0, "", "生成者 = 2") '医生站时
    
    Do While Not rsTmp.EOF
        If blnExecuted = False Then
            If Not IsNull(rsTmp!执行时间) And ((Mid(mstr执行场合, 1, 1) = "1" And mint场合 = 0 And Val(rsTmp!执行者) = 1) Or mint场合 = 1) Then
                blnExecuted = True
            End If
        End If
        strIDs = strIDs & "," & rsTmp!ID
        rsTmp.MoveNext
    Loop
    strIDs = Mid(strIDs, 2)
    If blnExecuted Then
        '不判断权限，不提示，强制取消
        If FuncExecuteAllCancel(False) = False Then
            Exit Function
        End If
    End If
    
    If mint场合 = 0 Then
        strSql = "Select 导入时间 from 病人临床路径 Where ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID)
        dat导入时间 = Format(rsTmp!导入时间 & "", "yyyy-MM-dd HH:mm:ss")
        '检查是否已评估
        If mbln启用执行环节 = False Or Not blnExecuted Then
            strSql = "Select 1 From 病人路径评估 Where 路径记录ID = [1] And 阶段ID = [2] And 日期 = [3]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, mPP.当前阶段ID, CDate(mPP.当前日期))
            If rsTmp.RecordCount > 0 Then
                '强制取消评估，不检查权限
                MsgBox "本次生成的项目已评估，取消生成之前将自动取消评估。", vbInformation, gstrSysName
                Call FuncEvaluateCancel(False, False)
            End If
        End If
        strIDSQL = "(Select Column_value From Table(f_Str2List([1])))"
        '2.检查医嘱
        '不是当天生成的长嘱，允许取消路径项目，不管是否发送；
        '是当天生成的长嘱，已校对但未作废，不允许取消，未校对的，取消时自动删除对应的医嘱。

        strSql = "Select /*+ Rule*/ distinct A.路径执行id" & vbNewLine & _
                 "From 病人路径医嘱 A, 病人路径医嘱 B" & vbNewLine & _
                 "Where a.路径执行id In " & strIDSQL & " And a.病人医嘱id = b.病人医嘱id And b.路径执行id <> a.路径执行id"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strIDs)
        If rsTmp.RecordCount = 0 Then
            strNewIDs = strIDs
            '没有非当日的长嘱
        Else
            '把以前生成了长嘱的那部分去掉，只检查当天的
            strNewIDs = "," & strIDs & ","
            For i = 1 To rsTmp.RecordCount
                If InStr(strNewIDs, "," & rsTmp!路径执行id & ",") > 0 Then
                    strNewIDs = Replace(strNewIDs, "," & rsTmp!路径执行id & ",", ",")
                End If
                rsTmp.MoveNext
            Next
            If strNewIDs = "," Then
                strNewIDs = ""
            Else
                strNewIDs = Mid(strNewIDs, 2, Len(strNewIDs) - 2)
            End If
        End If
        
        If strNewIDs <> "" Then
            '即使已停止的医嘱也不允许删除，加59秒是因为开嘱时间未精确到秒
            strSql = "Select /*+ Rule*/ C.医嘱内容 From 病人路径医嘱 B, 病人医嘱记录 C Where b.路径执行id In " & strIDSQL & _
                     " And b.病人医嘱id = c.Id And c.医嘱状态 > 1 And c.医嘱状态 <> 4 And rownum<2 And to_date(to_char(c.开嘱时间 +59/24/60/60,'yyyy-mm-dd hh24:mi:ss'),'yyyy-mm-dd hh24:mi:ss') >[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNewIDs, dat导入时间)
            If rsTmp.RecordCount > 0 Then
                strIDs = ""
                For i = 1 To rsTmp.RecordCount
                    If i > 10 Then strIDs = strIDs & "......": Exit For
                    strIDs = strIDs & vbNewLine & rsTmp!医嘱内容
                    rsTmp.MoveNext
                Next
                MsgBox "当前生成的项目存在已校对但未作废的医嘱：" & strIDs & vbNewLine & "请先作废医嘱后再执行取消。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    
'    If mint场合 = 1 Then
'        '对于已经过审核的医嘱，不允许修改删除。
'        strSql = "Select /*+ Rule*/ 1 From 病人路径医嘱 B, 病人医嘱记录 C Where b.路径执行id In " & strIDSQL & _
'                 " And b.病人医嘱id = c.Id And c.开嘱医生 Like '%/%' And rownum<2  And to_date(to_char(c.开嘱时间 +59/24/60/60,'yyyy-mm-dd hh24:mi:ss'),'yyyy-mm-dd hh24:mi:ss') >[2]"
'        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strIDs, dat导入时间)
'        If rsTmp.RecordCount > 0 Then
'            MsgBox "当前生成的项目对应的医嘱已经过医生审核，不能整体取消。", vbInformation, gstrSysName
'            Exit Function
'        End If
'    End If

        '3.检查病历
        strSql = "Select /*+ Rule*/ 1 From 电子病历记录 Where 路径执行id In " & strIDSQL & _
                 " And (完成时间 is not null or 打印人 is not null) And rownum<2  And 创建时间 >[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strIDs, dat导入时间)
        If rsTmp.RecordCount > 0 Then
            MsgBox "当前生成的项目对应的病历已签名或已打印，不能整体取消。", vbInformation, gstrSysName
            Exit Function
        End If
        
        '检查新版电子病历
        If Not CheckDelNewEMR(strIDs, 1, rsTmp) Then  '返回需要删除的电子病历任务ID
            Exit Function
        Else
            '删除
            If Not gobjEmr Is Nothing Then
                If Not gobjEmr.IsInited Or gobjEmr.IsOffline Then Set gobjEmr = Nothing
                If Not gobjEmr Is Nothing Then
                    For i = 1 To rsTmp.RecordCount
                        strSql = "<parameter><taskid>" & rsTmp!任务ID & "</taskid></parameter>"
                        On Error Resume Next
                        Call gobjEmr.DeleteTask(strSql)
                        On Error GoTo 0
                        rsTmp.MoveNext
                    Next
                End If
            End If
        End If
    End If
    
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(Split(strIDs, ","))
        strSql = "Zl_病人路径生成_Delete(" & Split(strIDs, ",")(i) & ",0," & mint场合 & ")"
        Call zlDatabase.ExecuteProcedure(strSql, "取消路径项目")
    Next
    gcnOracle.CommitTrans: blnTrans = False
    FuncDelAllItem = True

    If blnRefresh Then
        Call zlRefresh(mPati.病人ID, mPati.主页ID, mPati.病区ID, mPati.科室ID, mPati.病人状态)
    End If
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub FuncDelItem()
'功能：取消生成当前选择的未执行的路径项目
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim lng执行ID As Long, lng项目ID As Long, blnMust As Boolean, lng天数 As Long
    Dim blnCancel As Boolean, strReason As String, blnTrans As Boolean
    Dim vPoint As POINTAPI
    Dim i As Long
    
    With vsPath

        If mint场合 = 0 And Split(.Cell(flexcpData, .Row, .Col), "|")(3) = 2 Then '生成者 1-医生,2-护士
            MsgBox "当前项目是护士生成的,你不能删除。", vbInformation, Me.Caption
            Exit Sub
        ElseIf mint场合 = 1 And Split(.Cell(flexcpData, .Row, .Col), "|")(3) = 1 Then
            MsgBox "当前项目是医生生成的，你不能删除。", vbInformation, Me.Caption
            Exit Sub
        End If
         
        If .Cell(flexcpBackColor, .Row, .Col) = &HE0EFED Then
            MsgBox "该项目为必须生成但没有生成的项目，不用取消生成。", vbInformation, Me.Caption
            Exit Sub
        End If
        lng执行ID = Split(.Cell(flexcpData, .Row, .Col), "|")(0)
        lng项目ID = Split(.Cell(flexcpData, .Row, .Col), "|")(1)
    End With
    
    If mbln启用执行环节 Then
        '已经执行的不允许取消
        strSql = "Select 1 " & vbNewLine & _
                "From 病人路径执行 A, 临床路径项目 B" & vbNewLine & _
                "Where a.项目id = b.Id(+) And a.Id = [1] And Nvl(a.生成时间性质,0)<>1 And a.执行时间 Is Not Null"

        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng执行ID)
        If rsTmp.RecordCount > 0 Then
            MsgBox "该项目已执行，不能取消。", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    
    '1.检查路径项目
    strSql = "Select b.执行方式,a.天数 From 病人路径执行 a, 临床路径项目 b Where a.项目ID = b.ID And a.ID = [1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng执行ID)
    If rsTmp.RecordCount > 0 Then '临时项目，可以取消
        lng天数 = Val("" & rsTmp!天数)
        If rsTmp!执行方式 = 1 Then
            blnMust = True
        ElseIf rsTmp!执行方式 = 2 Or rsTmp!执行方式 = 4 Then  '至少一次或必须一次
            strSql = "Select 开始天数,结束天数 From 临床路径阶段 Where ID = [1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.当前阶段ID)
            If Not IsNull(rsTmp!开始天数) Then
                If Not IsNull(rsTmp!结束天数) Then
                    blnMust = (lng天数 = Val("" & rsTmp!结束天数))    '是否最后一天
                    If blnMust Then '判断该项目之前有没有执行过(路径外项目除外)
                    
                        strSql = "Select 1 From 病人路径执行 Where 路径记录ID = [1] And 阶段ID = [2] And 项目ID = [3] And 天数<[4] And rownum<2"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, mPP.当前阶段ID, lng项目ID, lng天数)
                        If rsTmp.RecordCount > 0 Then blnMust = False
                    End If
                Else
                    blnMust = True  '单天
                End If
            End If
        End If
        
    End If
    
    '2.检查医嘱
    If CheckDelPathItem(lng执行ID, mint场合) = False Then Exit Sub
    '3.必须生成的项目填写变异原因
    If blnMust Then
        '取消必须生成的项目时选择变异原因
        strSql = "Select b.名称 as 分类,a.编码 as ID,a.编码,a.名称,a.简码 From 变异常见原因 a,变异常见原因 b" & _
                " Where a.性质=1 And a.末级=1 And a.上级=b.编码 And b.末级=0 " & _
                " Order by 分类,a.编码"
        vPoint = zlControl.GetCoordPos(vsPath.Hwnd, vsPath.CellLeft, vsPath.CellTop)
        Set rsTmp = zlDatabase.ShowSelect(Me, strSql, 0, "变异常见原因", True, , , True, True, True, _
                 vPoint.X, vPoint.Y, vsPath.RowHeight(vsPath.Row), blnCancel, False, True)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "系统没有初始变异常见原因，请与系统管理员联系。", vbInformation, gstrSysName
            End If
            Exit Sub
        Else
            strReason = rsTmp!ID
        End If
    End If
    '4.检查病历
    strSql = "Select 1 From 电子病历记录 Where 路径执行id = [1] And (完成时间 is not null or 打印人 is not null) And rownum<2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng执行ID)
    If rsTmp.RecordCount > 0 Then
        MsgBox "该项目对应的病历已签名或已打印，不能取消。", vbInformation, gstrSysName
        Exit Sub
    End If
    '检查新版电子病历
    If Not CheckDelNewEMR(lng执行ID & "", 0, rsTmp) Then
        Exit Sub
    Else
        '删除
        On Error Resume Next
        For i = 1 To rsTmp.RecordCount
            strSql = "<parameter><taskid>" & rsTmp!任务ID & "</taskid></parameter>"
            Call gobjEmr.DeleteTask(strSql)
            rsTmp.MoveNext
        Next
        Err.Clear: On Error GoTo 0
    End If
    
    With vsPath
        If MsgBox("确实要取消路径项目""" & .TextMatrix(.Row, .Col) & """吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Sub
    End With
    If Not mbln启用执行环节 Then
        '判断是否已经评估
        strSql = "Select 1 From 病人路径评估 Where 路径记录ID = [1] And 阶段ID = [2] And 日期 = [3]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, mPP.当前阶段ID, CDate(mPP.当前日期))
        If rsTmp.RecordCount > 0 Then
            '强制取消评估，不检查权限
            MsgBox "本次生成的项目已评估，取消生成之前将自动取消评估。", vbInformation, gstrSysName
            Call FuncEvaluateCancel(False, False)
        End If
    End If
    gcnOracle.BeginTrans: blnTrans = True
    If strReason <> "" Then
        strSql = "Zl_病人路径生成_Update(" & lng执行ID & ",'" & vsPath.TextMatrix(vsPath.Row, 0) & "',Null,NULL,NULL,NULL,NULL,'" & strReason & "')"
        Call zlDatabase.ExecuteProcedure(strSql, "修改路径项目")
    End If
    strSql = "Zl_病人路径生成_Delete(" & lng执行ID & "," & IIf(strReason <> "", "2", "0") & "," & mint场合 & ")"
    
    Call zlDatabase.ExecuteProcedure(strSql, "取消路径项目")
    gcnOracle.CommitTrans: blnTrans = False
    Call zlRefresh(mPati.病人ID, mPati.主页ID, mPati.病区ID, mPati.科室ID, mPati.病人状态)

    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncAppendItemModify()
'功能：修改路径外项目
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim lng执行ID As Long
            
    With vsPath
        lng执行ID = Split(.Cell(flexcpData, .Row, .Col), "|")(0)
    End With
    
    strSql = "Select 1 From 病人路径评估 Where 路径记录ID = [1] And 阶段ID = [2] And 日期 = [3]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, mPP.当前阶段ID, CDate(mPP.当前日期))
    If rsTmp.RecordCount > 0 Then
        If InStr(GetInsidePrivs(P临床路径应用), ";阶段评估;") = 0 Then
            MsgBox "该病人在" & mPP.当前日期 & "已进行了评估，不能再修改路径外项目。", vbInformation, gstrSysName
            Exit Sub
        Else
            '取消评估
            If MsgBox("该病人在" & mPP.当前日期 & "已进行了评估，必须取消评估后才能修改路径外项目。" & vbCrLf & vbCrLf & "你现在要取消评估吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                Call FuncEvaluateCancel(False, False)
            Else
                Exit Sub
            End If
        End If
    End If
    
    If frmPathAppend.ShowMe(mfrmParent, mint场合, mPati, mPP, "", 2, "", lng执行ID, mclsMipModule) Then
        Call zlRefresh(mPati.病人ID, mPati.主页ID, mPati.病区ID, mPati.科室ID, mPati.病人状态)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function FuncAppendItem(ByVal bytUseType As Byte, Optional ByVal strItemType As String, Optional ByVal strAdviceIDs As String, _
                                Optional ByVal lng执行ID As Long, Optional ByVal datDate As Date) As Boolean
'功能：添加路径外项目(通过clsDockPath中的接口开放给医嘱部件调用)
'参数：bytUseType=0-直接添加,1-医嘱新开时添加
'       strItemType=医嘱接口调用时传入（当天最后一个项目的分类）
'       strAdviceIDs=医嘱接口调用时传入,医嘱序号
'       datDate =医嘱的开始执行日期（同一批路径外医嘱）
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim DatCur As Date
    Dim blnRefresh As Boolean
    
    If mint场合 = 0 Then '医生才有评估环节，护士没有评估环节
        strSql = "Select 1 From 病人路径评估 Where 路径记录ID = [1] And 阶段ID = [2] And 日期 = [3]"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, mPP.当前阶段ID, CDate(mPP.当前日期))
        If rsTmp.RecordCount > 0 Then
            If InStr(GetInsidePrivs(P临床路径应用), ";阶段评估;") = 0 Then
                MsgBox "该病人在" & mPP.当前日期 & "已进行了评估，不能再添加路径外项目。", vbInformation, gstrSysName
                Exit Function
            Else
                '取消评估
                If MsgBox("该病人在" & mPP.当前日期 & "已进行了评估，必须取消评估后才能添加路径外项目。" & vbCrLf & vbCrLf & "你现在要取消评估吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    Call FuncEvaluateCancel(False, False)
                Else
                    Exit Function
                End If
            End If
        Else
            '未评估则检查是否是当天未评估，不是的话则提示是否要添加到最近一个阶段
            If bytUseType = 0 Then
                DatCur = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
                If DatCur <> Format(mPP.当前日期, "yyyy-MM-dd") Then
                    If MsgBox("你要添加路径外项目到""" & mPP.当前日期 & """?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Call FuncSendItem
                        Exit Function
                    End If
                End If
            End If
        End If
    ElseIf mint场合 = 1 Then
        If bytUseType = 0 Then
            Call GetPhaseInNurse(0, mPP.当前阶段ID, mPP.当前天数, mPP.当前日期)
            blnRefresh = True
        End If
    End If
    
    If bytUseType = 0 Then
        With vsPath
            If .Row > 0 And .Row < .Rows - 2 Then strItemType = .TextMatrix(.Row, .FixedCols - 1) '最后一行是"路径评估"
        End With
    End If
    If frmPathAppend.ShowMe(mfrmParent, mint场合, mPati, mPP, strItemType, bytUseType, strAdviceIDs, lng执行ID, mclsMipModule, datDate) Or blnRefresh Then
        FuncAppendItem = True
        Call zlRefresh(mPati.病人ID, mPati.主页ID, mPati.病区ID, mPati.科室ID, mPati.病人状态)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function FuncImport(Optional ByVal lngHwnd As Long) As Boolean
'功能：导入路径
'参数：lngHwnd=新版病历传入父窗体句柄，默认为0,新版病历不提示导入不成功的原因
    Dim rsTmp As ADODB.Recordset
    '----
    Dim strSql As String
    '----
    Dim lngPPStatus As Long
    Dim t_pp As TYPE_PATH_Pati
    Dim str名称 As String, lngDiagnosisType As Long, lngDiagnosisSorce As Long
    Dim lng疾病ID As Long, lng诊断ID As Long
    
    '1.检查该病人当前是否存在正在执行的路径，包括非本次住院和他科的,以前并发产生的
    strSql = "Select b.名称 From 病人临床路径 a,部门表 b Where a.科室id = b.id And a.病人ID = [1] And a.状态 = 1"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "FuncImport", mPati.病人ID)
    If rsTmp.RecordCount > 0 Then
        If lngHwnd = 0 Then MsgBox "该病人在[" & rsTmp!名称 & "]还有正在执行的临床路径，不允许导入新的路径。", vbInformation, gstrSysName
        Exit Function
    End If
    
    lngPPStatus = mPP.病人路径状态
    
    FuncImport = frmPathImport.ShowMe(mfrmParent, mPati, 0, t_pp, , , , , , , , lngHwnd, str名称, lngDiagnosisType, lngDiagnosisSorce, lng疾病ID, lng诊断ID)
    If lngHwnd <> 0 And FuncImport = True And t_pp.路径ID <> 0 Then
        '新版病历不能在窗体中再去调用评估窗体,否则会被最小化
        FuncImport = frmEvaluate.ShowMe(mfrmParent, 0, 1, mPati, t_pp, str名称, lngDiagnosisType, lngDiagnosisSorce, lng疾病ID, lng诊断ID, 0)
    End If
    If lngHwnd <> 0 Then Exit Function
    Call zlRefresh(mPati.病人ID, mPati.主页ID, mPati.病区ID, mPati.科室ID, mPati.病人状态)
    
    If lngPPStatus <> mPP.病人路径状态 Then
        '当前病人路径状态发生变化时更新Lis病人路径状态
        If Not gobjLIS Is Nothing Then
           Call gobjLIS.ModifyPathState(mPati.病人ID, mPati.主页ID, mPP.病人路径状态)
        End If
    End If
    
    If mPP.病人路径状态 = 1 Then
        '如果导入成功，则检查是否需要继续导入合并路径
        Call frmPathImport.ShowMe(mfrmParent, mPati, 2, t_pp, , , , True, mPP.病人路径ID, , , lngHwnd)
    End If
    Call zlRefresh(mPati.病人ID, mPati.主页ID, mPati.病区ID, mPati.科室ID, mPati.病人状态)
    RaiseEvent RequestRefresh(mPP.病人路径状态)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function FuncImportMerge() As Boolean
'功能：导入合并路径
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim t_pp As TYPE_PATH_Pati
    
    '1.判断当前是否评估
    If Val(mPP.当前阶段ID & "") <> 0 Then
        strSql = "Select 时间进度 From 病人路径评估 Where 路径记录ID = [1] And 阶段ID = [2] And 日期 = [3]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, mPP.当前阶段ID, CDate(mPP.当前日期))
        If rsTmp.RecordCount = 0 Then
            If InStr(GetInsidePrivs(P临床路径应用), ";阶段评估;") = 0 Then
                MsgBox "该病人在" & mPP.当前日期 & "还没有进行评估，不能进行后续操作。", vbInformation, gstrSysName
                Exit Function
            Else
                If MsgBox("该病人在" & mPP.当前日期 & "还没有进行评估，必须先评估。" & vbCrLf & "你现在要进行评估操作吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    If frmEvaluate.ShowMe(mfrmParent, 1, 1, mPati, mPP) = False Then
                        Exit Function
                    Else
                        Call zlRefresh(mPati.病人ID, mPati.主页ID, mPati.病区ID, mPati.科室ID, mPati.病人状态)
                        '评估后，可能结束或退出路径，所以根据评估有的状态进行判断是否要继续生成,退出或完成则不继续生成
                        If mPP.病人路径状态 <> 1 Then
                            Exit Function
                        End If
                    End If
                Else
                    Exit Function
                End If
            End If
        End If
    
    End If
     '2.检查合并路径个数不超过5个
    strSql = "Select 合并路径个数 From 病人临床路径 Where ID = [1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID)
    If Val(rsTmp!合并路径个数 & "") >= 5 Then
        MsgBox "该病人在已经导入了5个合并路径，不允许再导入新的合并路径了。", vbInformation, gstrSysName
        Exit Function
    End If
        
    FuncImportMerge = frmPathImport.ShowMe(mfrmParent, mPati, 2, t_pp, , , , False, mPP.病人路径ID)
    Call zlRefresh(mPati.病人ID, mPati.主页ID, mPati.病区ID, mPati.科室ID, mPati.病人状态)
    RaiseEvent RequestRefresh(mPP.病人路径状态)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub FuncUnImport(Optional ByVal blnPrompt As Boolean = True)
'功能：取消导入,未生成路径时可取消导入
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset, blnTrans As Boolean
    Dim str审核人 As String
    Dim lngPPStatus As Long
    
    '先检查是否有取消路径的权限
    If InStr(GetInsidePrivs(P临床路径应用), ";取消导入;") = 0 Then
        str审核人 = zlDatabase.UserIdentify(Me, "没有取消导入权限需要审核。", glngSys, P临床路径应用, "取消导入")
        If str审核人 = "" Then Exit Sub
    Else
        str审核人 = UserInfo.姓名
    End If
    strSql = "Select 1 From 病人路径执行 Where 路径记录ID = [1] And rownum<2"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID)
    If rsTmp.RecordCount > 0 Then
        
        If MsgBox("当前阶段的路径项目已生成，首先将进行取消生成操作。" & vbCrLf & "你确实要取消该病人已导入的临床路径吗?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        End If
        
        '要刷新，以便重取路径信息(当前阶段等)
        If FuncDelAllItem(True, False) Then
            Call FuncUnImport(False)    '重新调用，再次检查
        End If
        Exit Sub
    ElseIf blnPrompt Then
        If MsgBox("你确实要取消该病人已导入的临床路径吗?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    
    lngPPStatus = mPP.病人路径状态
    
    gcnOracle.BeginTrans: blnTrans = True
    strSql = "Zl_病人路径导入_Delete(" & mPP.病人路径ID & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "取消导入")
    '插入取消导入记录
    strSql = "Zl_病人路径取消_Insert(" & mPati.病人ID & "," & mPati.主页ID & ",'" & UserInfo.姓名 & "','" & str审核人 & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "取消导入")
    gcnOracle.CommitTrans: blnTrans = False
    Call zlRefresh(mPati.病人ID, mPati.主页ID, mPati.病区ID, mPati.科室ID, mPati.病人状态)
    
    '当前病人路径状态发生变化时更新Lis病人路径状态
    If lngPPStatus <> mPP.病人路径状态 Then
        If Not gobjLIS Is Nothing Then
           Call gobjLIS.ModifyPathState(mPati.病人ID, mPati.主页ID, mPP.病人路径状态)
        End If
    End If
    
    RaiseEvent RequestRefresh(mPP.病人路径状态)
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncUnImportMerge(Optional ByVal blnPrompt As Boolean = True)
'功能：取消导入合并路径
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset, blnTrans As Boolean
    Dim str审核人 As String
    Dim colSQL As New Collection
    Dim strIDs As String, i As Long
    Dim t_pp As TYPE_PATH_Pati
    
    '查找已经导入的合并路径
    strSql = "Select a.Id, a.路径id, b.名称, b.编码, b.说明, NVL(Sign(Max(c.Id)),0) As 是否执行" & vbNewLine & _
            "From 病人合并路径 A, 临床路径目录 B, 病人路径执行 C" & vbNewLine & _
            "Where a.路径id = b.Id And c.合并路径记录id(+) = a.Id And a.首要路径记录id = [1]" & vbNewLine & _
            "Group By a.Id, a.路径id, b.名称, b.编码, b.说明"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID)
    '如果只有一条合并路径，则直接提示否则弹出选择
    If rsTmp.RecordCount = 0 Then
        MsgBox "该病人未导入任何合并路径，不能取消导入。", vbInformation, gstrSysName
        Exit Sub
    ElseIf rsTmp.RecordCount = 1 Then
        If Val(rsTmp!是否执行 & "") = 1 Then
            MsgBox "该病人的合并路径:" & rsTmp!名称 & "已经生成了项目，请取消合并路径的项目后再取消导入。", vbInformation, gstrSysName
            Exit Sub
        Else
            If MsgBox("你确定要取消导入合并路径：" & rsTmp!名称 & "?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                Exit Sub
            End If
        End If
        strIDs = rsTmp!ID & ""
    Else
        If Not frmPathImport.ShowMe(mfrmParent, mPati, 3, t_pp, , , , , , rsTmp, True) Then Exit Sub
        Unload frmPathImport
        If rsTmp.RecordCount = 0 Then Exit Sub
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            strIDs = strIDs & "," & rsTmp!ID
            rsTmp.MoveNext
        Loop
        strIDs = Mid(strIDs, 2)
    End If
    If strIDs <> "" Then
        For i = 0 To UBound(Split(strIDs, ","))
            strSql = "Zl_病人路径导入_Delete(" & mPP.病人路径ID & "," & Val(Split(strIDs, ",")(i)) & ")"
            colSQL.Add strSql, "C" & colSQL.count + 1
        Next
    End If
    
    gcnOracle.BeginTrans: blnTrans = True
    For i = 1 To colSQL.count
        Call zlDatabase.ExecuteProcedure(colSQL("C" & i), "取消合并路径")
    Next
    gcnOracle.CommitTrans: blnTrans = False
    Call zlRefresh(mPati.病人ID, mPati.主页ID, mPati.病区ID, mPati.科室ID, mPati.病人状态)
    RaiseEvent RequestRefresh(mPP.病人路径状态)
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ViewMergeImport()
'功能：查看合并路径导入评估
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim t_pp As TYPE_PATH_Pati
    
    '查找已经导入的合并路径
    strSql = "Select a.Id, a.路径id, b.名称, b.编码, b.说明" & vbNewLine & _
            "From 病人合并路径 A, 临床路径目录 B" & vbNewLine & _
            "Where a.路径id = b.Id  And a.首要路径记录id = [1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID)
    If rsTmp.RecordCount = 0 Then
        MsgBox "当前未导入任何合并路径。", vbInformation, gstrSysName
        Exit Sub
    End If
    If Not frmPathImport.ShowMe(mfrmParent, mPati, 3, t_pp, , , , , , rsTmp) Then Exit Sub
    Unload frmPathImport
    If rsTmp.RecordCount = 0 Then Exit Sub
    Call frmEvaluate.ShowMe(mfrmParent, 0, 0, mPati, mPP, , , , , , 1, , Val(rsTmp!ID & ""))
    
    Call zlRefresh(mPati.病人ID, mPati.主页ID, mPati.病区ID, mPati.科室ID, mPati.病人状态)
    RaiseEvent RequestRefresh(mPP.病人路径状态)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function FuncEvaluateCancel(Optional ByVal blnPrompt As Boolean = True, Optional ByVal blnRefresh As Boolean = True) As Boolean
'功能：取消评估,未变异时才能取消（变异后自动结束，只能取消结束）
'参数：blnPrompt=是否弹出询问提示
    Dim strSql As String, rsTmp As ADODB.Recordset, i As Long
    Dim lngPPStatus As Long
    
    On Error GoTo errH
    
    If Not mbln启用不评估 Then
        strSql = "Select 1 From 病人路径评估 Where 路径记录ID = [1] And 阶段ID = [2] And 日期 = [3]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, mPP.当前阶段ID, CDate(mPP.当前日期))
        If rsTmp.RecordCount = 0 Then
            MsgBox "该病人在" & mPP.当前日期 & "还没有进行评估。", vbInformation, gstrSysName
            Exit Function
        End If
    Else
        strSql = "Select A.阶段ID,A.日期,A.天数 From (Select t.阶段id, t.日期, t.天数 From 病人路径评估 T Where t.路径记录id = [1] Order By t.登记时间 Desc, t.天数 Desc) A where rownum<2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID)
        If rsTmp.RecordCount > 0 Then
            If Format(rsTmp!日期, "YYYY-MM-DD") < mPP.当前日期 Or (Format(rsTmp!日期, "YYYY-MM-DD") = mPP.当前日期 And Val(rsTmp!阶段ID & "") <> mPP.当前阶段ID) Then
                mPP.当前阶段ID = rsTmp!阶段ID
                mPP.当前日期 = Format(rsTmp!日期, "YYYY-MM-DD")
                mPP.当前天数 = rsTmp!天数
             End If
         Else
            MsgBox "该病人不存在任何评估记录。", vbInformation, gstrSysName
            Exit Function
         End If
    End If
        
    If blnPrompt Then
        strSql = "Select 1 From 病人合并路径 Where 首要路径记录ID = [1] And 首要路径阶段ID = [2] And 首要路径天数=[3]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, mPP.当前阶段ID, mPP.当前天数)
        If rsTmp.RecordCount > 0 Then
            If MsgBox("当前阶段已经导入了合并路径，取消评估将同时取消导入合并路径，是否继续？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Function
        Else
            If MsgBox("你确定要取消第" & mPP.当前天数 & "天的评估吗？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Function
        End If
    End If
    lngPPStatus = mPP.病人路径状态
    
    strSql = "Zl_病人路径评估_Delete(" & mPP.病人路径ID & ", " & mPP.当前阶段ID & ",To_Date('" & mPP.当前日期 & "','YYYY-MM-DD HH24:MI:SS'))"
    Call zlDatabase.ExecuteProcedure(strSql, "取消评估")
    FuncEvaluateCancel = True
    If blnRefresh Then
        Call zlRefresh(mPati.病人ID, mPati.主页ID, mPati.病区ID, mPati.科室ID, mPati.病人状态)
    End If
                 
    '当前病人路径状态发生变化时更新Lis病人路径状态
    If lngPPStatus <> mPP.病人路径状态 Then
        If Not gobjLIS Is Nothing Then
           Call gobjLIS.ModifyPathState(mPati.病人ID, mPati.主页ID, mPP.病人路径状态)
        End If
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub FuncEvaluate()
'功能：阶段评估,只能对当前阶段的最后一次且执行了的进行
    Dim strSql As String, rsTmp As ADODB.Recordset, i As Long
    Dim strTmp As String
    Dim bln补录 As Boolean
    Dim blnRefresh As Boolean
    Dim strDate  As String
    Dim lngPPStatus As Long
    
    '1.已评估的不能再评估 '已结束的不能再评估(禁用了菜单项)
    '2.只能对最后一次执行的记录进行评估(凡定义了评估指标的，必须评估后才能生成次日路径，没有定义的则不评估就执行生成)，因为评估为变异则可能结束路径
    '3.必须该阶段的所有项目（医生生成的项目）都执行后才能评估

    On Error GoTo errH
  
    If Not mbln启用不评估 Then
        strSql = "Select 1 From 病人路径评估 Where 路径记录ID = [1] And 阶段ID = [2] And 日期 = [3]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, mPP.当前阶段ID, CDate(mPP.当前日期))
        If rsTmp.RecordCount > 0 Then
            MsgBox "该病人在" & mPP.当前日期 & "已进行了评估。", vbInformation, gstrSysName
            Exit Sub
        End If
    Else
        '查询结果缺省是按照路径执行ID排降序,能够取到第一个未评估的阶段
        strSql = "Select * " & vbNewLine & _
                "From (Select a.阶段id, a.日期, a.天数" & vbNewLine & _
                "       From 病人路径执行 A, 病人路径评估 B" & vbNewLine & _
                "       Where a.路径记录id = [1] And a.路径记录id = b.路径记录id(+) And a.阶段id = b.阶段id(+) And a.日期 = b.日期(+) And b.阶段id Is Null" & vbNewLine & _
                "       Order By a.Id) A" & vbNewLine & _
                "Where Rownum < 2"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID)
        '读取当日日期备用
        strDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD")
        If rsTmp.RecordCount > 0 Then
            If Format(rsTmp!日期, "YYYY-MM-DD") <= strDate Then  '允许评估的
                If mPP.当前日期 & "_" & mPP.当前阶段ID = Format(rsTmp!日期, "YYYY-MM-DD") & "_" & Val(rsTmp!阶段ID & "") Then
                    '非补录评估,没有提前生成的路径阶段
                    bln补录 = False
                ElseIf Format(rsTmp!日期, "YYYY-MM-DD") <= mPP.当前日期 Then
                    mPP.当前阶段ID = rsTmp!阶段ID
                    mPP.当前日期 = rsTmp!日期 & ""
                    mPP.当前天数 = rsTmp!天数
                    bln补录 = True
                End If
            ElseIf Format(rsTmp!日期, "YYYY-MM-DD") > strDate Then
                MsgBox "你现在要评估的阶段日期：【" & Format(rsTmp!日期, "YYYY-MM-DD") & "】" & vbCrLf & "超过了当前日期：【" & strDate & "】" & vbCrLf & "不能进行评估操作。", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
                Exit Sub
            End If
        Else
            MsgBox "该病人所有阶段都已经完成评估。", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    '执行登记检查
    If Not CheckPathIsExecuted(blnRefresh) Then
        '强制刷新
        If blnRefresh Then
            Call zlRefresh(mPati.病人ID, mPati.主页ID, mPati.病区ID, mPati.科室ID, mPati.病人状态)
        End If
        Exit Sub
    End If
    
    lngPPStatus = mPP.病人路径状态
    
    If frmEvaluate.ShowMe(mfrmParent, 1, 1, mPati, mPP, , , , , , , , , bln补录) Or bln补录 Then
        Call zlRefresh(mPati.病人ID, mPati.主页ID, mPati.病区ID, mPati.科室ID, mPati.病人状态)
    End If
    
    '当前病人路径状态发生变化时更新Lis病人路径状态
    If lngPPStatus <> mPP.病人路径状态 Then
        If Not gobjLIS Is Nothing Then
           Call gobjLIS.ModifyPathState(mPati.病人ID, mPati.主页ID, mPP.病人路径状态)
        End If
    End If
    
    If mPP.病人路径状态 = 2 Or mPP.病人路径状态 = 3 Then RaiseEvent RequestRefresh(mPP.病人路径状态)
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncReEvaluate()
'功能：修改评估，如果产生了后续阶段的项目，不能修改评估结果为变异来结束评估，在保存的存储过程中判断。
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim bln补录 As Boolean
    Dim lng阶段ID As Long
    Dim strSysDate As String
    Dim lng天数 As Long
    Dim lngPPStatus As Long

    On Error GoTo errH
    
    If Not mbln启用不评估 Then
        strSql = "Select 原路径ID From 病人路径评估 Where 路径记录ID = [1] And 阶段ID = [2] And 日期 = [3]"
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, mPP.当前阶段ID, CDate(mPP.当前日期))
        If rsTmp.RecordCount = 0 Then
            MsgBox "该病人在当前阶段还没有进行评估。", vbInformation, gstrSysName
            Exit Sub
        ElseIf Val("" & rsTmp!原路径ID) <> 0 Then
            MsgBox "该病人已跳转到其他路径，如果要修改评估，请取消评估后重新评估。", vbInformation, gstrSysName
            Exit Sub
        End If
    Else
        strSql = "Select A.阶段ID,A.日期,A.天数 From (Select t.阶段id, t.日期, t.天数 From 病人路径评估 T Where t.路径记录id = [1] Order By t.登记时间 Desc, t.天数 Desc) A where rownum<2"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID)
        If rsTmp.RecordCount > 0 Then
            '等于的情况不更新阶段ID，会查不到数据报错
            If Format(rsTmp!日期, "YYYY-MM-DD") <= mPP.当前日期 Then
                mPP.当前阶段ID = rsTmp!阶段ID
                mPP.当前日期 = Format(rsTmp!日期, "YYYY-MM-DD")
                mPP.当前天数 = rsTmp!天数
                bln补录 = True
             End If
         Else
            MsgBox "该病人在当前阶段还没有进行评估。", vbInformation, gstrSysName
            Exit Sub
         End If
    End If
    
    lngPPStatus = mPP.病人路径状态
    
    If frmEvaluate.ShowMe(mfrmParent, 1, 2, mPati, mPP, , , , , , , , , bln补录) Or bln补录 Then
        Call zlRefresh(mPati.病人ID, mPati.主页ID, mPati.病区ID, mPati.科室ID, mPati.病人状态)
    End If
    
    '当前病人路径状态发生变化时更新Lis病人路径状态
    If lngPPStatus <> mPP.病人路径状态 Then
        If Not gobjLIS Is Nothing Then
           Call gobjLIS.ModifyPathState(mPati.病人ID, mPati.主页ID, mPP.病人路径状态)
        End If
    End If
    
    If mPP.病人路径状态 = 2 Or mPP.病人路径状态 = 3 Then RaiseEvent RequestRefresh(mPP.病人路径状态)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncViewReport(ByVal str报告ID As String, ByVal lng医嘱ID As Long)
'功能：查阅报告
        
    '先判断是否可以继续操作
    If IsNumeric(str报告ID) Then
        If CheckEPRReport(Val(str报告ID), lng医嘱ID) = 2 Then
            If InStr(GetInsidePrivs(p住院医嘱下达), "查阅未完成报告") > 0 Then
                MsgBox "注意：该医嘱的报告还没有正式签名！", vbInformation, gstrSysName
            Else
                MsgBox "该医嘱的报告还没有完成(没有正式签名或完成执行)，你没有权限操作！", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        RaiseEvent ViewEPRReport(Val(str报告ID), False)
    Else
        Call CreateObjectPacs(mobjPublicPACS)
        Call mobjPublicPACS.zlDocShowReport(0, str报告ID, Val(zlDatabase.GetPara("自动标记报告查阅状态", glngSys, p住院医嘱下达, "1")) = 1, mfrmParent)
    End If
    
End Sub


Public Function CheckEPRReport(ByVal lng报告ID As Long, ByVal lng医嘱ID As Long) As Integer
'功能：检查对应项目的报告填写情况
'参数：lng路径执行ID=病人路径执行记录中的ID
'      lng报告ID=返回报告病历ID
'返回：
'      1-报告已填写完成(已签名,包括修订后签名,或已执行完成)
'      2-报告未填写完成(未签名,或修订后未签名,且未执行完成)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
        
    '检查报告执行过程(5-审核;6-报告完成)和状态(1-完成)
    '检验报告是关联到采集方式上面的，但采集方式可能为叮嘱未产生发送记录
    strSql = _
        " Select 2 as 排序,医嘱ID,执行过程,执行状态,发送时间 From 病人医嘱发送 Where 医嘱ID=[1]" & _
        " Union ALL" & _
        " Select 排序,医嘱ID,执行过程,执行状态,发送时间" & _
        " From (" & _
            " Select 1 as 排序,B.医嘱ID,B.执行过程,B.执行状态,B.发送时间 From 病人医嘱记录 A,病人医嘱发送 B" & _
            " Where A.ID=B.医嘱ID And A.相关ID=(" & _
                " Select A.ID From 病人医嘱记录 A,诊疗项目目录 B Where A.ID=[1] And A.诊疗项目ID=B.ID And A.诊疗类别='E' And B.操作类型='6')" & _
            " Order by A.序号" & _
        " ) Where Rownum=1" & _
        " Order by 排序,发送时间 Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "CheckEPRReport", lng医嘱ID)
    If Nvl(rsTmp!执行过程, 0) >= 5 Or Nvl(rsTmp!执行状态, 0) = 1 Then
        CheckEPRReport = 1
    Else
        CheckEPRReport = 2
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set mcolReason = Nothing
    Set mclsMipModule = Nothing
    SaveWinState Me, App.ProductName
    Set mobjPublicPACS = Nothing
End Sub

Private Sub imgMore_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strInfo As String, lngId As Long, i As Long
    Dim strSql As String, rsTmp As ADODB.Recordset
        
    lngId = fraMore.Tag
    If lngId = 0 Then
        Call zlCommFun.ShowTipInfo(0, strInfo)
    Else
        strSql = "Select  decode(NVL(Nvl(a.生成者, b.生成者),1),1,'医生',2,'护士') As 生成者," & IIf(mbln启用执行环节, "A.执行结果,A.执行说明,A.执行人,to_char(A.执行时间,'yyyy-mm-dd hh24:mi') as 执行时间,", "") & _
                " A.登记人,to_char(A.登记时间,'yyyy-mm-dd hh24:mi') as 登记时间 From 病人路径执行 A,临床路径项目 B Where A.项目ID=B.ID(+) And A.ID = [1]"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngId)
        If rsTmp.RecordCount > 0 Then
            With rsTmp
                For i = 0 To .Fields.count - 1
                    strInfo = strInfo & .Fields(i).Name & "：" & .Fields(i).Value & vbCrLf
                Next
            End With
            Call zlCommFun.ShowTipInfo(fraMore.Hwnd, strInfo, True)
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub optSelect_Click(Index As Integer)
    If Me.Visible Then
        optSelect(IX_ALL).Tag = Index  '标记当前选择项
        Call LoadPathItem '刷新
    End If
End Sub

Private Sub vsFlow_DblClick()
    Dim lngPhaseID As Long
    If mPP.病人路径状态 = 0 And mPP.病人路径ID <> 0 Then   '导入失败
        Call frmEvaluate.ShowMe(mfrmParent, 0, 0, mPati, mPP)
    Else
        lngPhaseID = Val(vsFlow.ColData(vsFlow.Col))
        If lngPhaseID <> 0 Then
            Call frmPathSend.ShowMe(mfrmParent, 2, mint场合, mPati, mPP, lngPhaseID, 0, , , , mclsMipModule)
        ElseIf vsFlow.Col = 0 And mPP.路径ID <> 0 Then
            Call frmPathDefinition.ShowMe(mfrmParent, mPP.路径ID)
        Else
            If vsFlow.Col = vsFlow.Cols - 2 And gstrDBUser = "ZLHIS" Then
                vsFlow.Editable = flexEDKbdMouse
            End If
        End If
    End If
End Sub

Private Sub vsFlow_LostFocus()
    If Not (mPP.病人路径状态 = 0 And mPP.病人路径ID <> 0) Then
        vsFlow.ForeColorSel = vsFlow.CellForeColor
    End If
End Sub

Private Sub vsFlow_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 '功能：用于测试时强制删除最后一天的项目(选中最后一个箭头后，输入DELA)
    Dim strPass As String, i As Long
        
    If vsFlow.Col = vsFlow.Cols - 2 Then
        strPass = UCase(vsFlow.EditText)
        vsFlow.EditText = ""
        If strPass = "DELA" Then
            If MsgBox("您确定要删除最后一天的所有项目吗？", vbInformation + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Sub
            Call FuncDelPhaseItem
        End If
        vsFlow.Editable = flexEDNone
    End If
End Sub

Private Sub vsPath_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow <> NewRow Or OldCol <> NewCol Then
        If fraMore.Visible Then fraMore.Visible = False
        
        If NewRow <> -1 And NewCol <> -1 And mblnUnChange = False Then
            '显示路径项目生成的医嘱清单
            Dim strTmp As String
            
            strTmp = vsPath.Cell(flexcpData, NewRow, NewCol)
            If InStr(strTmp, "|") > 0 Then
                Call UCAdvice.ShowAdvice(1, "", Val("" & Split(strTmp, "|")(0)))
            Else
                Call UCAdvice.ShowAdvice(1, "", 0)
            End If
        Else
            Call UCAdvice.ShowAdvice(1, "", 0)
        End If
    End If
End Sub

Private Sub vsPath_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    vsPath.AutoSize vsPath.FixedCols, vsPath.Cols - 1, , 45
End Sub

Private Sub vsPath_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)
    
    If fraMore.Visible Then fraMore.Visible = False
End Sub

Private Sub vsPath_DblClick()
    Dim lng项目ID As Long

    With vsPath
        If Trim(.TextMatrix(.Row, .Col)) <> "" And .Cell(flexcpData, .Row, .Col) <> "" And .Row <> .Rows - 1 Then
            lng项目ID = Split(.Cell(flexcpData, .Row, .Col), "|")(1)
            If lng项目ID <> 0 Then
                Call frmPathItemEdit.ShowView(mfrmParent, lng项目ID)
            End If
        End If
    End With
End Sub

Private Sub vsPath_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With vsPath
        If .MouseCol >= .FixedCols And .MouseRow >= .FixedRows Then
            Dim lngId As Long, lngRow As Long, lngCol As Long, lngItemID As Long
            
            lngRow = .MouseRow: lngCol = .MouseCol
            If .Cell(flexcpData, lngRow, lngCol) <> "" And lngRow <> .Rows - 1 Then
                lngId = Split(.Cell(flexcpData, lngRow, lngCol), "|")(0)
                lngItemID = Split(.Cell(flexcpData, lngRow, lngCol), "|")(1)
                If lngItemID = 0 Then
                    .ToolTipText = ""
                    Call zlCommFun.ShowTipInfo(.Hwnd, mcolReason("C" & lngId), True)      '路径外项目的添加原因
                Else
                    If .ToolTipText = "" Then .ToolTipText = "双击查看路径项目定义"
                    Call zlCommFun.ShowTipInfo(.Hwnd, "")
                End If
            Else
                .ToolTipText = ""
            End If
            
            If lngId = 0 Then
                If imgMore.Visible Then fraMore.Visible = False
                fraMore.Tag = ""
            Else
                If lngRow = .Row And lngCol = .Col Then
                    fraMore.BackColor = .BackColorSel
                Else
                    fraMore.BackColor = .BackColor
                End If
            
                fraMore.Tag = lngId
                If fraMore.Visible = False Then fraMore.Visible = True
                fraMore.Top = .Top + .RowPos(lngRow) + .RowHeight(lngRow) - imgMore.Height - 30
                fraMore.Left = .Left + .ColPos(lngCol) + .ColWidth(lngCol) - imgMore.Width - 30
            End If
        Else
            If fraMore.Visible Then fraMore.Visible = False
        End If
    End With
End Sub

Private Sub vsPath_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBarPopup
    Dim lng项目ID As Long

    '显示编辑菜单下面的内容
    If Button = 2 Then
        If mcbsMain Is Nothing Then Exit Sub
        Set objPopup = mcbsMain.ActiveMenuBar.FindControl(, conMenu_EditPopup)
        If Not objPopup Is Nothing Then
            objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub

Private Sub OutLogModi()
    Dim colSQL As New Collection, i As Long, blnTrans As Boolean

    Call frmPathOutLog.ShowMe(mfrmParent, mPati.病人ID, mPati.主页ID, 2, colSQL, mPP.路径ID, mPP.病人路径ID)

    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    '执行出径登记表的SQL
    For i = 1 To colSQL.count
        Call zlDatabase.ExecuteProcedure(colSQL("C" & i), "修改出径登记表")
    Next
    gcnOracle.CommitTrans: blnTrans = False

    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub SetFontSize(ByVal bytSize As Byte)
'功能:设置路径表流程图和清单的字体大小
'入参:bytSize：0-小(缺省)，1-大
    mlngFontSize = IIf(bytSize = 0, CON_SmallFontSize, CON_BigFontSize)
    
    vsFlow.Font.Size = mlngFontSize
    vsFlow.Redraw = flexRDDirect
    
    Call Grid.SetFontSize(vsPath, mlngFontSize)
    If vsPath.FixedRows > 1 Then vsPath.AutoSize vsPath.FixedCols, vsPath.Cols - 1, , 45 '在要Draw之后才生效
    
    Call UCAdvice.SetVsAdviceFontSize(mlngFontSize)
End Sub

Private Sub MovePathItem(ByVal lngWay As Long)
'功能:当前单元格选中路径外项目时，可在路径外项目中上下移动
'参数:lngWay=1上移一行,-1下移一行(相当于下一行上移一行)
    Dim lngId       As Long
    Dim lngItemNum  As Long
    Dim arrSQL()    As Variant
    Dim i           As Integer
    Dim blnTran     As Boolean
    Dim blnDo As Boolean, blnFind As Boolean
    Dim lngRow As Long, lngCol As Long

    blnDo = True: blnFind = False

    With vsPath
        Do While blnDo

            If .TextMatrix(.Row, .FixedCols - 1) <> .TextMatrix(.Row - lngWay, .FixedCols - 1) Or .Cell(flexcpData, .Row - lngWay, .Col) = "" Then
                MsgBox "项目内容:" & .TextMatrix(.Row, .Col) & vbCrLf & _
                       "已处于【" & .TextMatrix(.Row, .FixedCols - 1) & "】分类的" & IIf(lngWay > 0, "第一项", "最后一项"), vbInformation, gstrSysName
                blnDo = False: blnFind = False: Exit Do
            Else
                lngRow = .Row - lngWay: lngCol = .Col
                blnFind = True: Exit Do
            End If
        Loop
        '交换项目序号
        If blnFind Then
            arrSQL = Array()

            lngId = Split(.Cell(flexcpData, .Row, .Col), "|")(0)
            lngItemNum = Split(.Cell(flexcpData, .Row - lngWay, .Col), "|")(2)
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "zl_病人路径序号_update(" & lngId & "," & lngItemNum & ")"

            lngId = Split(.Cell(flexcpData, .Row - lngWay, .Col), "|")(0)
            lngItemNum = Split(.Cell(flexcpData, .Row, .Col), "|")(2)
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "zl_病人路径序号_update(" & lngId & "," & lngItemNum & ")"

            On Error GoTo errH
            gcnOracle.BeginTrans: blnTran = True
            For i = LBound(arrSQL) To UBound(arrSQL)
                zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
            Next i
            gcnOracle.CommitTrans: blnTran = False

            Call ClearPathItem(True)
            Call LoadPathItem

            '焦点移动
            .Row = lngRow: .Col = lngCol
            .ShowCell lngRow, lngCol

        End If
    End With
    Exit Sub
errH:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function CheckSendOfBefore() As Boolean
'功能:提前生成前检查
'返回: T-允许提前生成,F-不允许提前生成
    Dim strTmp As String
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim blnReturn As Boolean
    
    On Error GoTo errH
    strTmp = "," & UserInfo.性质 & ","
    If InStr(strTmp, ",医生,") > 0 Then
        '允许提前生成
        blnReturn = True
    Else
        '护士提前生成时，需检查医生首先生成过没有
        strSql = "Select 1" & vbNewLine & _
            "From 病人路径执行 A, 临床路径项目 B" & vbNewLine & _
            "Where a.路径记录id = [1] And a.阶段ID=[2] And a.天数=[3] And a.项目id = b.Id(+) And Nvl(Nvl(a.生成者, b.生成者),1) = 1 And Rownum < 2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, mPP.当前阶段ID, mPP.当前天数)
        If rsTmp.RecordCount > 0 Then
            blnReturn = True
        Else
            MsgBox "医生还没有提前生成下一天的路径项目，护士不能提前生成。", vbInformation + vbOKOnly
            blnReturn = False
        End If
    End If
    CheckSendOfBefore = blnReturn
    Exit Function
errH:
   If ErrCenter() = 1 Then
        Resume
   End If
   Call SaveErrLog

End Function

Private Function CheckDelNewEMR(ByVal str路径执行IDs As String, ByVal bytMode As Byte, ByRef rsTmp As ADODB.Recordset) As Boolean
'功能:删除新版电子病历检查
'参数:
'   str路径执行IDs
'   bytMode 0-单个路径项目,1-多个路径项目
'   str任务IDs -任务ID串 ID,ID,ID....
    Dim strSql As String
    Dim rsTask As ADODB.Recordset
    Dim i As Long
    
    If bytMode = 0 Then
        strSql = "Select 任务ID from 病人路径病历 where 路径执行Id=[1] "
    Else
        strSql = "Select 任务ID from 病人路径病历 where 路径执行Id in (Select Column_value From Table(f_Num2List([2]))) "
    End If
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(str路径执行IDs), str路径执行IDs)
    If rsTmp.RecordCount > 0 Then
        If Not gobjEmr Is Nothing Then
            If Not gobjEmr.IsInited Or gobjEmr.IsOffline Then Set gobjEmr = Nothing
            If Not gobjEmr Is Nothing Then
                For i = 1 To rsTmp.RecordCount
                    strSql = "<parameter><taskid>" & rsTmp!任务ID & "</taskid></parameter>"
                    On Error Resume Next
                    Set rsTask = gobjEmr.GetTaskStatus(strSql)
                    Err.Clear: On Error GoTo 0
                    '记录集返回0行（数据异常）或1行数据；
                    '记录集包含字段：ID，创建人，创建时间，完成人，完成时间，审核人，审核时间，最近打印人，最近打印时间；其中ID为任务ID；除ID，创建人，创建时间外，其余字段都可能为空；
                    '当完成人为非空时，表示该病历文件已完成。
                    If rsTask.State <> adStateClosed Then
                        If rsTask.RecordCount = 1 Then
                            If rsTask!完成人 <> "" Then
                                If bytMode = 0 Then
                                    MsgBox "该项目对应的电子病历文件已完成,不能取消。", vbInformation, gstrSysName
                                Else
                                    MsgBox "本次生成中存在对应的电子病历文件已完成,不能取消本次生成。", vbInformation, gstrSysName
                                End If
                                Exit Function
                            End If
                        End If
                    End If
                    rsTmp.MoveNext
                Next
            End If
        End If
    End If
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    CheckDelNewEMR = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetPhaseInNurse(ByVal bytType As Byte, ByRef lng阶段ID As Long, ByRef lng天数 As Long, Optional ByRef strDate As String, _
            Optional ByRef lng下一阶段 As Long, Optional ByRef lng下一天 As Long, _
             Optional ByRef strPhase As String = "-1")
'功能:获取护士场合阶段和天数
'
'参数:
'    bytType:0-护士补充生成、添加\取消路径外项目、取消生成时缺省的阶段和天数;（最后生成的阶段和天数）
'            1-护士生成时缺省的阶段和天数
'返回:
'    出参：
'    lng阶段ID:
'    lng天数: 天数用于区同一阶段,生成多天的情况
'    strPhase:阶段信息SQL =-1不返回SQL
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim lngRowNUM As Long
    
    On Error GoTo errH
    '取护士场合最后生成的阶段及天数,登记时间取最小，是为了保证取的是生成时的执行记录,不包含补充生成、暂存路径项目
    If mPP.当前阶段分支ID = 0 Then
        strSql = "Select 阶段ID,天数,日期,生成者,RowNum as 排序 " & vbNewLine & _
        "From (Select a.阶段id, a.天数, To_Char(a.日期, 'yyyy-mm-dd') as 日期,生成者 " & vbNewLine & _
                 "From (Select a.阶段id, a.天数, a.日期,a.路径记录id,NVl(生成者,1) as 生成者" & vbNewLine & _
                 "       From 病人路径执行 A" & vbNewLine & _
                 "       Where a.路径记录id = [1] " & vbNewLine & _
                 "       Group By a.阶段id, a.天数, a.日期,a.路径记录id,NVl(生成者,1)) A, 临床路径阶段 B,临床路径阶段 C,病人临床路径 G" & vbNewLine & _
                 "Where a.阶段id = b.Id And b.父id=c.id(+) And g.id=A.路径记录ID " & vbNewLine & _
                 "Order By 日期,Decode(g.路径id,b.路径id,1,0), NVL(c.序号,b.序号))"
    Else
        strSql = "Select 阶段ID,天数,日期,生成者,RowNum as 排序 " & vbNewLine & _
            "From (Select a.阶段id, a.天数, To_Char(a.日期, 'yyyy-mm-dd') as 日期,生成者 " & vbNewLine & _
                 "From (Select a.阶段id, a.天数, a.日期,a.路径记录id,NVl(生成者,1) as 生成者" & vbNewLine & _
                 "       From 病人路径执行 A" & vbNewLine & _
                 "       Where a.路径记录id = [1] " & vbNewLine & _
                 "       Group By a.阶段id, a.天数, a.日期,a.路径记录id,NVl(生成者,1)) A, 临床路径阶段 B,临床路径阶段 C,临床路径分支 D,临床路径阶段 E,临床路径阶段 F,病人临床路径 G" & vbNewLine & _
                 "Where a.阶段id = b.Id And b.父id=c.id(+) And b.分支id=d.id(+) and d.前一阶段id=e.id(+) And e.父id=f.id(+)  And g.id=A.路径记录ID " & vbNewLine & _
                 "Order By 日期,Decode(g.路径id,b.路径id,1,0), Decode(b.分支ID,Null,NVL(c.序号,b.序号),NVL(c.序号,b.序号)+NVL(f.序号,e.序号)))"
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "生成路径项目", mPP.病人路径ID)
    rsTmp.Filter = "生成者=2"  '=2护士生成 =1医生生成
    If rsTmp.RecordCount = 0 Then
        '护士场合还没有生成过任何路径项目,医生场合生成的第一个阶段和天数
        rsTmp.Filter = "生成者=1"
        If rsTmp.RecordCount > 0 Then
            '护士下一阶段及天数
            If bytType = 1 Then
                lng下一阶段 = Val(rsTmp!阶段ID & "")
                lng下一天 = Val(rsTmp!天数 & "")
                strDate = rsTmp!日期 & ""
                If strPhase <> "-1" Then
                    strPhase = Rec.ToSQL(rsTmp)
                End If
                Exit Sub
            End If
        End If
    Else
        '护士最后阶段及天数
        rsTmp.Sort = "排序 DESC"
        lng阶段ID = Val(rsTmp!阶段ID & "")
        lng天数 = Val(rsTmp!天数 & "")
        strDate = rsTmp!日期 & ""
        If bytType = 1 Then
            rsTmp.Sort = ""
            rsTmp.Filter = "生成者=1"
            Do
                If rsTmp!阶段ID & "_" & rsTmp!天数 & "_" & rsTmp!日期 = lng阶段ID & "_" & lng天数 & "_" & strDate Then
                    '找到护士下一阶段及天数
                    lngRowNUM = Val(rsTmp!排序 & "")
                    rsTmp.Filter = "生成者=1 and 排序 >" & lngRowNUM
                    lng下一阶段 = Val(rsTmp!阶段ID & "")
                    lng下一天 = Val(rsTmp!天数 & "")
                    If strPhase <> "-1" Then
                        strPhase = Rec.ToSQL(rsTmp)
                    End If
                    Exit Do
                End If
                rsTmp.MoveNext
            Loop While Not rsTmp.EOF
       End If
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function FuncConvertPathTable() As VSFlexGrid
'功能:转换临床路径表单,用于应对特殊的打印需求 78233
'返回
'   -转换后的路径表单
    Dim lngFirstSameCol As Long
    Dim lngLastRow As Long
    Dim i As Long, j As Long, k As Long
    
    Grid.CopyTo vsPath, vsPathPrint(0)
    vsPath.Redraw = flexRDNone
    vsPathPrint(0).MergeCol(0) = True
    vsPathPrint(0).MergeRow(0) = True
    
    With vsPathPrint(0)
        '一个阶段打印一列，不打印天数和日期
        For i = 2 To .Cols - 1
            If .TextMatrix(R0阶段名, i) = .TextMatrix(R0阶段名, i - 1) Then
            '相邻两列为同一阶段要合并
                For j = 1 To i - 1
                    If .TextMatrix(R0阶段名, i) = .TextMatrix(R0阶段名, j) Then
                        lngFirstSameCol = j '找到与当前列在同一阶段的首列
                        Exit For
                    End If
                Next
                'j-当前行,i-当前列
                For j = R2日期 + 1 To .Rows - 1
                    If .TextMatrix(j, i) <> "" And .TextMatrix(j, 0) <> "评估情况" Then
                        k = 0
                        For k = R2日期 + 1 To .Rows - 1
                            If .TextMatrix(j, i) = .TextMatrix(k, lngFirstSameCol) Then
                                Exit For
                            End If
                        Next
                        If k = .Rows Then
                            '新增
                            k = 0
                            lngLastRow = 0
                            For k = R2日期 + 1 To .Rows - 1
                                If .TextMatrix(j, 0) = .TextMatrix(k, 0) Then
                                    If .TextMatrix(k, lngFirstSameCol) = "" Then
                                        '同分类下空行新增
                                        .TextMatrix(k, lngFirstSameCol) = .TextMatrix(j, i)
                                        Exit For
                                    End If
                                    lngLastRow = k
                                End If
                            Next
                            If k = .Rows Then
                                '同分类最后一行新增一行
                                .AddItem "", lngLastRow + 1
                                .TextMatrix(lngLastRow + 1, lngFirstSameCol) = .TextMatrix(j, i)
                                .TextMatrix(lngLastRow + 1, 0) = .TextMatrix(j, 0)
                                .RowHeight(lngLastRow + 1) = .RowHeight(j)
                            End If
                        Else
                            '前一列存在，所以不处理
                        End If
                        
                    End If
                Next
                '标记删除列
                .ColHidden(i) = True
                .ColWidth(i) = 0
            End If
        Next
        '删除列
        For i = .Cols - 1 To 0 Step -1
            If .ColHidden(i) = True And .ColWidth(i) = 0 Then
                '最后一行直接删除
                If i = .Cols - 1 Then
                    .Cols = .Cols - 1
                Else
                    '后面往前移
                    For k = i + 1 To .Cols - 1
                        For j = 0 To .Rows - 1
                            .TextMatrix(j, k - 1) = .TextMatrix(j, k)
                        Next
                    Next
                    .Cols = .Cols - 1
                End If
            End If
        Next
        '隐藏日期和天数
        .RowHidden(R1天数) = True: .RowHidden(R2日期) = True
        .RowHeight(R1天数) = 0: .RowHeight(R2日期) = 0
        .Redraw = flexRDDirect
        If .FixedRows > 1 Then .AutoSize .FixedCols, .Cols - 1, , 45    '在要Draw之后才生效
    End With
    Set FuncConvertPathTable = vsPathPrint(0)
End Function

Private Function CheckPathIsExecuted(Optional ByRef blnRefresh As Boolean) As Boolean
'-------------------------------------------------------------------------------------------
'功能:检查当前阶段是否存在未执行的路径项目
'参数：=1 生成环节调用,=2 评估时候调用
'返回：F-存在未完成执行登记的路径项目,不允许生成或评估
'     T-不存在未完成执行登记的路径项目\不检查执行登记情况，允许生成或评估
'说明：1.护士生成时,当前未执行时不允许生成新的阶段,医生评估时,当前未执行不允许生成新的阶段
'      2.考虑医生要提前生成后续阶段 mbln启用不评估=true时,医生站生成时不检查是否完成执行登记，放在评估环节检查
'      3.护士站没有评估环节,需要每次生成都要检查前一次的执行登记情况
'-------------------------------------------------------------------------------------------
    Dim blnHave As Boolean       '不启用执行环节的检查
    Dim blnReturn As Boolean
    Dim blnExePath As Boolean
    Dim blnUnExe As Boolean      '用于标记没有执行路径权限且存在操作员执行的路径项目时,需要给予用户提示
    Dim strSubSQL As String
    Dim strSql As String
    Dim strTmp As String
    
    Dim strMsg As String
    Dim rsTmp As ADODB.Recordset
    
    Dim i As Long
    
    On Error GoTo errH
    
    blnHave = True '默认检查执行登记情况
    blnExePath = InStr(GetInsidePrivs(P临床路径应用), ";执行路径;") > 0
    blnReturn = True
    
    If mbln启用执行环节 And mstr执行场合 <> "00" Then
        If mint场合 = 0 Then
            '医生站评估
            If mstr执行场合 = "11" Then
                strSubSQL = "And NVL(NVL(a.生成者,b.生成者),1)=1"
            ElseIf mstr执行场合 = "10" Then
                strSubSQL = "And NVL(NVL(a.生成者,b.生成者),1)=1 And Nvl(Nvl(a.执行者,b.执行者),1)=1 "
            ElseIf mstr执行场合 = "01" Then
                strSubSQL = "And NVL(NVL(a.生成者,b.生成者),1)=1 And Nvl(Nvl(a.执行者,b.执行者),1)=2 "
            End If
        ElseIf mint场合 = 1 Then
            '护士站（没有评估环节）
            If mstr执行场合 = "11" Or mstr执行场合 = "01" Then
                strSubSQL = "And Nvl(Nvl(a.执行者,b.执行者),1)=2 "
            ElseIf mstr执行场合 = "10" Then
                blnHave = False
            End If
        End If
    Else
        blnHave = False
    End If
    
    If blnHave Then
      
        strSql = "Select Nvl(b.项目内容,a.项目内容) 项目内容,NVl(Nvl(a.执行者,b.执行者),1) as 执行者 From 病人路径执行 a,临床路径项目 b " & vbNewLine & _
                        "Where a.项目id=b.id(+) And a.路径记录ID = [1] And a.阶段ID = [2] And a.日期 = [3] And Nvl(a.生成时间性质,0)<>2 And a.执行时间 Is null " & strSubSQL

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, mPP.当前阶段ID, CDate(mPP.当前日期))
        If rsTmp.RecordCount > 0 Then
            If mint场合 = 0 Then
                '医生场合评估环节检查,生成时不检查
                If mstr执行场合 = "11" Then
                    rsTmp.Filter = " 执行者 = 1"
                    If rsTmp.RecordCount > 0 Then
                        Call FuncGetRSTipInfo(rsTmp, "项目内容", strTmp)
                        If blnExePath Then
                            If MsgBox("该病人还有未执行的项目:" & vbCrLf & strTmp & vbCrLf & "必须先执行。你现在要进行执行操作吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                                If frmPathExecute.ShowMe(mfrmParent, 0, mPati, mPP, 0, mint场合) Then
                                    blnRefresh = True
                                    If Not CheckPathIsExecuted() Then
                                        blnReturn = False
                                        '再次检查，有可能存在护士未执行的项目
                                    End If
                                Else
                                    blnReturn = False
                                End If
                            Else
                                blnReturn = False
                            End If
                        Else
                            blnUnExe = True: blnReturn = False
                        End If
                    Else
                        rsTmp.Filter = " 执行者 = 2"
                        Call FuncGetRSTipInfo(rsTmp, "项目内容", strTmp)
                        strMsg = "该病人还有护士未执行的项目:" & vbCrLf & strTmp & vbCrLf & "必须执行后才能继续。"
                        blnReturn = False
                    End If
                ElseIf mstr执行场合 = "01" Then
                    '只检查生成者是医生且执行者是护士的路径项目
                    Call FuncGetRSTipInfo(rsTmp, "项目内容", strTmp)
                    strMsg = "该病人还有护士未执行的项目:" & vbCrLf & strTmp & vbCrLf & "必须执行后才能继续。"
                    blnReturn = False
                End If
            End If
            
            If mint场合 = 1 Or (mint场合 = 0 And mstr执行场合 = "10") Then
                '护士生成
                Call FuncGetRSTipInfo(rsTmp, "项目内容", strTmp)
                If (mint场合 = 1 And (mstr执行场合 = "11" Or mstr执行场合 = "01")) Or (mint场合 = 0 And mstr执行场合 = "10") Then
                     If blnExePath Then
                        If MsgBox("该病人还有未执行的项目:" & vbCrLf & strTmp & vbCrLf & "必须先执行。你现在要进行执行操作吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                            Call frmPathExecute.ShowMe(mfrmParent, 0, mPati, mPP, 0, mint场合) '批量执行登记
                            blnRefresh = True
                        Else
                            blnReturn = False
                        End If
                     Else
                        blnUnExe = True: blnReturn = False
                     End If
                End If
            End If
            '
            If blnUnExe Then
                '没有执行路径权限且存在操作员执行的路径项目时 , 需要给予用户提示
                strMsg = "该病人还有未执行的项目：" & vbCrLf & strTmp & vbCrLf & "必须执行后才能继续。"
            End If
            
            If strMsg <> "" Then
                MsgBox strMsg, vbInformation, gstrSysName
            End If
        End If
    End If
    
    CheckPathIsExecuted = blnReturn
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub FuncGetRSTipInfo(ByVal rsTmp As ADODB.Recordset, ByVal strFieldName As String, ByRef strTipInfo As String)
'------------------------------------------------------------------------------------------
'功能:循环读取记录集中简要信息
'-----------------------------------------------------------------------------------------
    Dim i As Long
    
    strTipInfo = ""
    For i = 1 To rsTmp.RecordCount
        strTipInfo = IIf(i = 1, "", strTipInfo & vbCrLf) & rsTmp.Fields(strFieldName)
        If Len(strTipInfo) > 200 Then strTipInfo = strTipInfo & "…": Exit For
        rsTmp.MoveNext
    Next
End Sub

Private Function CheckPathSendByNurse(ByVal bytFunc As Byte, ByVal lng路径记录ID As Long, Optional ByVal lng阶段ID As Long, Optional ByVal dat日期 As Date) As Boolean
'--------------------------------------------
'功能：检查当前阶段的项目中是否存在护士生成的项目
'参数: bytFunc=1 批量取消执行登记   =2单个取消执行登记
'      bytFunc=1时 lng路径记录ID 路径记录ID
'      bytFunc=2时 lng路径记录ID 路径执行ID
'返回:T-存在,F-不存在
'--------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim blnRet As Boolean
    
    On Error GoTo errH
    If bytFunc = 1 Then
        strSql = "Select 1 " & vbNewLine & _
                "From 病人路径执行 A, 临床路径项目 B" & vbNewLine & _
                "Where a.路径记录id = [1] And a.阶段id = [2] And a.日期 = [3] And a.项目id = b.Id(+) And Nvl(Nvl(a.生成者, b.生成者), 1) = 2 and rownum<2 "
    Else
        strSql = "Select 1 " & vbNewLine & _
               "From 病人路径执行 A, 临床路径项目 B" & vbNewLine & _
               "Where a.ID=[1] And a.项目id = b.Id(+) And Nvl(Nvl(a.生成者, b.生成者), 1) = 2 and rownum<2 "
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng路径记录ID, lng阶段ID, dat日期)
    
    If rsTmp.RecordCount > 0 Then
        blnRet = True
    End If
    
    CheckPathSendByNurse = blnRet
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub GetPathCurrPhase(ByVal bytType As Byte, ByRef lng阶段ID As Long, ByRef lng天数 As Long, Optional ByRef strDate As String)
'--------------------------------------------------
'功能:获取批量执行登记或批量取消执行登记的当前阶段
'参数:bytType =1 批量执行,=2批量取消执行
'-------------------------------------------------
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    If bytType = 1 Then
        strSql = "Select *" & vbNewLine & _
            "From (Select Distinct a.阶段id, a.日期, a.天数, Min(a.登记时间) As 登记时间" & vbNewLine & _
            "       From 病人路径执行 A, 临床路径项目 B" & vbNewLine & _
            "       Where a.项目id = b.Id(+) And a.路径记录id = [1] And Nvl(a.生成时间性质, 0) = 0 And Nvl(Nvl(a.执行者, b.执行者), 0) = " & IIf(mint场合 = 0, 1, 2) & _
            "             And a.执行时间 Is Null" & vbNewLine & _
            "       Group By a.阶段id, a.日期, a.天数" & vbNewLine & _
            "       Order By Min(a.登记时间))" & vbNewLine & _
            "Where Rownum < 2"
    Else
        strSql = "Select *" & vbNewLine & _
            "From (Select Distinct a.阶段id, a.日期, a.天数, Min(a.登记时间) As 登记时间" & vbNewLine & _
            "       From 病人路径执行 A, 临床路径项目 B" & vbNewLine & _
            "       Where a.项目id = b.Id(+) And a.路径记录id = [1] And Nvl(a.生成时间性质, 0) = 0 And Nvl(Nvl(a.执行者, b.执行者), 0) = " & IIf(mint场合 = 0, 1, 2) & _
            "             And a.执行时间 Is Not Null" & vbNewLine & _
            "       Group By a.阶段id, a.日期, a.天数" & vbNewLine & _
            "       Order By Min(a.登记时间) Desc )  " & vbNewLine & _
            "Where Rownum < 2"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID)
    If rsTmp.RecordCount > 0 Then
        lng阶段ID = rsTmp!阶段ID
        strDate = rsTmp!日期 & ""
        lng天数 = rsTmp!天数
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub DefCommandPlugInPopup(ByVal objBar As Object, ByRef rsBar As ADODB.Recordset)
'功能：在医嘱卡右键弹出菜单
    Dim i As Long
    Dim objControl As CommandBarControl
    Dim objCtl As CommandBarControl
    Dim objPopup As CommandBarPopup
    
    If rsBar Is Nothing Then Exit Sub
    rsBar.Filter = 0
    If rsBar.RecordCount = 0 Then Exit Sub
    
    '独立按钮
    rsBar.Filter = "IsInTool=1 and BarType=3"
    If Not rsBar.EOF Then
        rsBar.Sort = "序号"
        For i = 1 To rsBar.RecordCount
            Set objControl = objBar.Add(xtpControlButton, rsBar!功能ID, rsBar!功能名)
            objControl.IconId = rsBar!图标ID
            objControl.Parameter = rsBar!功能名
            objControl.Style = xtpButtonIconAndCaption
            If Val(rsBar!IsGroup) = 1 Then
                objControl.BeginGroup = True
            End If
            rsBar.MoveNext
        Next
    End If
    
    rsBar.Filter = "IsInTool=0 and BarType=3"
    If Not rsBar.EOF Then
        rsBar.Sort = "序号"
        Set objPopup = objBar.Add(xtpControlButtonPopup, conMenu_Tool_PlugIn, "扩展功能")
            objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            For i = 1 To rsBar.RecordCount
                Set objControl = .Add(xtpControlButton, rsBar!功能ID, rsBar!菜单名)
                objControl.IconId = rsBar!图标ID
                objControl.Parameter = rsBar!功能名
                objControl.Style = xtpButtonIconAndCaption
                If Val(rsBar!IsGroup) = 1 Then
                    objControl.BeginGroup = True
                End If
                rsBar.MoveNext
            Next
        End With
    End If
End Sub

Private Function GetPlugInBar(ByVal lng模块 As Long, ByVal int场合 As Integer, rsBar As ADODB.Recordset) As String
'功能：组织外挂部件的菜单样按钮
    Dim strFunc As String
    Dim strXML As String
    Call CreatePlugInOK(lng模块, int场合)
    If gobjPlugIn Is Nothing Then Exit Function
    On Error Resume Next
    strFunc = gobjPlugIn.GetFuncNames(glngSys, lng模块, int场合, strXML)
    Call zlPlugInErrH(Err, "GetFuncNames")
    Err.Clear: On Error GoTo 0
    Call MakePlugInBar(strFunc, strXML, rsBar)
    GetPlugInBar = strFunc
End Function

Private Sub MakePlugInBar(ByVal strFunc As String, ByVal strXML As String, rsBar As ADODB.Recordset)
'功能：组织菜单到本地记录集中，注意对老版本的兼容处理
'参数：strFunc 老版本功能列串，strXML含配置信息的功能串
    Dim strM As String
    Dim strB As String
    Dim strP As String
    Dim strTag As String
    Dim i As Long
    Dim strTmp As String
    Dim lngS As Long, lngE As Long
    Dim rsBarFuncID As ADODB.Recordset
    
    If strXML = "" And strFunc = "" Then Exit Sub
    If strXML = "" And strFunc <> "" Then
        '兼容以前老版本的方式
        Call InitPlugInRsBar(rsBar)
        Call AddPlugInBarRs(rsBar, strFunc, 1)
        Call AddPlugInBarRs(rsBar, strFunc, 2)
        Call AddPlugInBarRs(rsBar, strFunc, 3)
        Call SetPlugInBar(rsBar, 1)
        Exit Sub
    End If
    
    On Error GoTo errH
    strXML = Trim(strXML)
    '暂定为200个扩展功能插件，防止死循环
    For i = 0 To 200
        lngS = InStr(strXML, "<")
        lngE = InStr(strXML, ">")
        strTag = Mid(strXML, lngS + 1, lngE - lngS - 1)
        If strTag = "menubar" Then
            lngS = lngE
            lngE = InStr(strXML, "</menubar>")
            strTmp = Mid(strXML, lngS + 1, lngE - lngS - 1)
            If strTmp <> "" Then strM = strM & "," & strTmp
            strXML = Mid(strXML, lngE + 10)
        ElseIf strTag = "toolbar" Then
            lngS = lngE
            lngE = InStr(strXML, "</toolbar>")
            strTmp = Mid(strXML, lngS + 1, lngE - lngS - 1)
            If strTmp <> "" Then strB = strB & "," & strTmp
            strXML = Mid(strXML, lngE + 10)
        ElseIf strTag = "popbar" Then
            lngS = lngE
            lngE = InStr(strXML, "</popbar>")
            strTmp = Mid(strXML, lngS + 1, lngE - lngS - 1)
            If strTmp <> "" Then strP = strP & "," & strTmp
            strXML = Mid(strXML, lngE + 9)
        End If
        If strXML = "" Then
            Exit For
        End If
    Next
    If strM = "" Then Exit Sub
    strM = Mid(strM, 2)
    strB = Mid(strB, 2)
    strP = Mid(strP, 2)

    Call InitPlugInRsBar(rsBar)
    Call AddPlugInBarRs(rsBar, strM, 1)
    Call AddPlugInBarRs(rsBar, strB, 2)
    Call AddPlugInBarRs(rsBar, strP, 3)
    Call SetPlugInBar(rsBar, 2)
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub AddPlugInBarRs(ByRef rsBar As ADODB.Recordset, ByVal strFunc As String, ByVal intType As Integer)
'功能：将功能串转换为记录集方式
'参数：strFunc 功能串，intType 功能按钮属于那一栏 1-菜单栏，2-工具栏，3-左键栏
    Dim varFunc As Variant
    Dim i As Long
    Dim strFuncName As String
    Dim blnFirstTool As Boolean
    If strFunc = "" Then Exit Sub
    varFunc = Split(strFunc, ",")
    With rsBar
        For i = 0 To UBound(varFunc)
            strFuncName = varFunc(i)
            .AddNew
            !BarType = intType
            If InStr(strFuncName, "Auto:") > 0 Then
                !IsAuto = 1
                strFuncName = Replace(strFuncName, "Auto:", "")
            Else
                !IsAuto = 0
            End If
            
            If InStr(strFuncName, "InTool:") > 0 Then
                !IsInTool = 1
                strFuncName = Replace(strFuncName, "InTool:", "")
            Else
                !IsInTool = 0
            End If
            If InStr(strFuncName, "|:") > 0 Then
                !IsGroup = 1
                strFuncName = Replace(strFuncName, "|:", "")
            Else
                !IsGroup = 0
                If Not blnFirstTool And !IsInTool = 1 Then
                    '第一个独立按钮显示分割线
                    blnFirstTool = True
                    !IsGroup = 1
                End If
            End If
            !功能名 = strFuncName
            !菜单名 = strFuncName
            .Update
        Next
    End With
End Sub

Private Function SetPlugInBar(ByRef rsBar As ADODB.Recordset, ByVal lngV As Long) As String
'功能：分配功能ID，加菜单快键
'参数：lngV 版本，1-老版，2-新版
'返回：字符串，以前低版本方式的功能串
    Dim i As Long
    '分配功能ID，图标ID
    With rsBar
        .Filter = 0
        If .EOF Then Exit Function
        .MoveFirst
        For i = 1 To .RecordCount
            !序号 = i
            !功能ID = conMenu_Tool_PlugIn_Item + i
            !图标ID = conMenu_Tool_PlugIn_Item
            If lngV = 1 Then
                !IsInTool = 0
                !IsGroup = 0
            End If
            .Update
            .MoveNext
        Next
    End With
    Call SetPlugInBarKey(rsBar, 1, lngV)
    Call SetPlugInBarKey(rsBar, 2, lngV)
    Call SetPlugInBarKey(rsBar, 3, lngV)
    rsBar.Filter = 0
End Function

Private Sub SetPlugInBarKey(rsBar As ADODB.Recordset, ByVal intType As Integer, ByVal lngV As Long)
'功能：设定快键
'参数：lngV 版本，1-老版，2-新版 intType 功能按钮属于那一栏 1-菜单栏，2-工具栏，3-左键栏
    Dim i As Long
    With rsBar
        .Filter = "IsInTool=0 and BarType=" & intType
        If .RecordCount = 1 And lngV = 2 Then
            '如果只有一个，也归为独立按钮
            !IsInTool = 1
            .Update
        Else
            For i = 1 To .RecordCount
                If i <= 35 Then
                    If i <= 9 Then
                        !菜单名 = !菜单名 & "(&" & i & ")"
                    Else
                        !菜单名 = !菜单名 & "(&" & Chr(55 + i) & ")"
                    End If
                    .Update
                    .MoveNext
                Else
                    Exit For
                End If
            Next
        End If
        
        .Filter = "IsInTool=1 and BarType=" & intType
        For i = 1 To .RecordCount
            If i <= 35 Then
                If i <= 9 Then
                    !菜单名 = !菜单名 & "(&" & i & ")"
                Else
                    !菜单名 = !菜单名 & "(&" & Chr(55 + i) & ")"
                End If
                .Update
                .MoveNext
            Else
                Exit For
            End If
        Next
    End With
End Sub

Private Sub InitPlugInRsBar(rsBar As ADODB.Recordset)
    Set rsBar = New ADODB.Recordset
    rsBar.Fields.Append "序号", adBigInt '用于排序
    rsBar.Fields.Append "功能ID", adBigInt '菜单按钮 Control.ID
    rsBar.Fields.Append "图标ID", adBigInt
    rsBar.Fields.Append "功能名", adVarChar, 1000 '去掉关键字之后的 名称 即工具栏上的按钮名称
    rsBar.Fields.Append "菜单名", adVarChar, 1000 '菜单栏/右键菜单 名称
    rsBar.Fields.Append "IsAuto", adInteger '是否自动执行功能
    rsBar.Fields.Append "IsGroup", adInteger '是否分割线
    rsBar.Fields.Append "IsInTool", adInteger '是否独立显示
    rsBar.Fields.Append "BarType", adInteger '1-菜单栏，2－工具栏，3－弹出栏
    rsBar.CursorLocation = adUseClient
    rsBar.LockType = adLockOptimistic
    rsBar.CursorType = adOpenStatic
    rsBar.Open
End Sub

Private Sub FuncPathTableChange(ByRef vsBody As VSFlexGrid, ByVal lngPageCOL As Long, Optional vsHead As VSFlexGrid)
'功能:将打印表单转换成固定列,便于打印输出：
'主要解决问题:89612-当阶段行高超过打印有效范围时要求下一页继续补打当前阶段剩余行
'            80442-每一阶段的字体自动缩放行间距,剔除空白行。
'参数:
'出参:vsBody打印表体
'入参:lngPageCOL 打印列数(不含固定列)
    Dim lngRow As Long
    Dim lngCol As Long
    
    On Error Resume Next
    Load vsPathPrint(1)
    Err.Clear: On Error GoTo 0
    
    With vsPathPrint(1)
        '清空
        .Rows = 0
        .Cols = 0
        
        If lngPageCOL = 0 Then Exit Sub
        If (vsBody.Cols - vsBody.FixedCols) Mod lngPageCOL <> 0 Then Exit Sub
        '
        .Rows = ((vsBody.Cols - vsBody.FixedCols) / lngPageCOL) * vsBody.Rows
        .Cols = vsBody.FixedCols + lngPageCOL
        .FixedCols = vsBody.FixedCols
        .FixedRows = vsBody.FixedRows
        
        '从vsBody将数据复制到vsPathPrint(1)
        '固定列
        For lngCol = 0 To .FixedCols
            lngRow = 0
            Do
                '（原本是固定行转换成非固定行时需要特殊标记便于打印部件识别）
                If lngRow Mod vsBody.Rows < vsBody.FixedRows And lngRow >= vsBody.FixedRows And lngCol = 0 Then
                    .RowData(lngRow) = UCase("FIXEDROW")
                End If
                Call FuncPathCellCopy(vsBody, vsPathPrint(1), lngRow Mod vsBody.Rows, lngCol, lngRow, lngCol)
                lngRow = lngRow + 1
            Loop While lngRow <> .Rows
        Next
        '非固定列
        For lngCol = .FixedCols To (.FixedCols + lngPageCOL) - 1
            lngRow = 0
            Do
                Call FuncPathCellCopy(vsBody, vsPathPrint(1), lngRow Mod vsBody.Rows, (lngPageCOL * (lngRow \ vsBody.Rows)) + lngCol, lngRow, lngCol)
                lngRow = lngRow + 1
            Loop While lngRow <> .Rows
        Next
        
        '清空多列都是空白的行
        For lngRow = 0 To .Rows - 1
            For lngCol = 1 To lngPageCOL
                If .RowData(lngRow) = UCase("FIXEDROW") Then
                    .Cell(flexcpAlignment, lngRow, lngCol, lngRow, .Cols - 1) = flexAlignCenterCenter
                    Exit For
                ElseIf .TextMatrix(lngRow, 0) = "护士签名" Or .TextMatrix(lngRow, 0) = "医生签名" Then
                    Exit For
                ElseIf .TextMatrix(lngRow, lngCol) <> "" Then
                    Exit For
                ElseIf lngCol = lngPageCOL Then
                    '记录下要删除的空白行
                   .RemoveItem lngRow
                   lngRow = lngRow - 1  '删除一行,下一行向上填充
                End If
            Next
            If lngRow = .Rows - 1 Then Exit For
        Next
        '显示风格处理
        .MergeCol(0) = True
        '设定字体，宽度
        .FontSize = IIf(mlngFontSize = 0, CON_SmallFontSize, mlngFontSize) '路径跟踪批量打印mlngFontSize=0
        '显示风格
        .Cell(flexcpAlignment, 0, .FixedCols, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        .Cell(flexcpForeColor, 0, 0, .Rows - 1, .Cols - 1) = vbBlack
    End With
    
    
    Set vsBody = vsPathPrint(1)
End Sub

Private Sub FuncPathCellCopy(ByRef vsSource As VSFlexGrid, ByRef vsCopy As VSFlexGrid, _
        ByVal lngSourRow As Long, ByVal lngSourCol As Long, ByVal lngCopyRow As Long, ByVal lngCopyCol As Long)
'功能:复制单元格
'参数：vsSource-被Copy的表单
'    vsCopy-copy后的表单
'    lngSourRow ,lngSourCol 被Copy的表单对应行和列
'    lngCopyRow，lngCopyCol  Copy后表单对应行和列
    With vsCopy
        .Cell(flexcpText, lngCopyRow, lngCopyCol) = vsSource.Cell(flexcpText, lngSourRow, lngSourCol)
        .Cell(flexcpAlignment, lngCopyRow, lngCopyCol) = vsSource.Cell(flexcpAlignment, lngSourRow, lngSourCol)
        .Cell(flexcpBackColor, lngCopyRow, lngCopyCol) = vsSource.Cell(flexcpBackColor, lngSourRow, lngSourCol)
        .Cell(flexcpForeColor, lngCopyRow, lngCopyCol) = vsSource.Cell(flexcpForeColor, lngSourRow, lngSourCol)
        .Cell(flexcpPicture, lngCopyRow, lngCopyCol) = vsSource.Cell(flexcpPicture, lngSourRow, lngSourCol)
    End With
End Sub

Public Function GetFormOperation() As String
'功能：获取窗体操作选择，该接口会在窗体卸载前调用，新版护士站 病人事务窗口
'返回：记录当前界面中控件选择状态
 
    Dim strXML As String
    Dim lngIdx As Long
     
    If optSelect(IX_ALL).Value Then
        lngIdx = IX_ALL
    ElseIf optSelect(IX_医生).Value Then
        lngIdx = IX_医生
    ElseIf optSelect(IX_护士).Value Then
        lngIdx = IX_护士
    End If
    strXML = "<root><scz>" & lngIdx & "</scz></root>"  '生成者
    GetFormOperation = strXML
End Function

Public Function RestoreFormOperation(ByVal strValue As String)
'功能：恢复窗体操作选择
'参数：strValue 前界面中控件选择状态

    Dim objXML As New zl9ComLib.clsXML
    Dim strTmp As String
     Dim lngIdx As Long
    
    On Error Resume Next
    
    Call objXML.OpenXMLDocument(strValue)
    
    Call objXML.GetSingleNodeValue("scz", strTmp) '婴儿
    lngIdx = Val(strTmp)

    Set mcolReason = New Collection
    optSelect(lngIdx).Value = True
End Function