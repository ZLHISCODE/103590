VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCHRecipe 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "中药配方"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCHRecipe.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   400
      Left            =   10635
      TabIndex        =   18
      ToolTipText     =   "热键：F2"
      Top             =   240
      Width           =   1170
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   400
      Left            =   10635
      TabIndex        =   19
      Top             =   750
      Width           =   1170
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   400
      Left            =   10635
      TabIndex        =   20
      Top             =   7725
      Width           =   1170
   End
   Begin VB.Frame fraInfo 
      Height          =   600
      Left            =   15
      TabIndex        =   22
      Top             =   15
      Width           =   10455
      Begin VB.Frame fraAuto 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   8940
         TabIndex        =   24
         Top             =   465
         Width           =   330
      End
      Begin VB.TextBox txtAuto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         IMEMode         =   3  'DISABLE
         Left            =   8970
         MaxLength       =   2
         TabIndex        =   1
         Text            =   "5"
         Top             =   225
         Width           =   285
      End
      Begin VB.CheckBox chkAuto 
         Caption         =   "自动识别   位输入码"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   7680
         TabIndex        =   0
         Top             =   240
         Width           =   2595
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请选择中药形态后依次输入每味中草药及剂量。选择中草药请按 * 键。"
         Height          =   240
         Left            =   105
         TabIndex        =   23
         Top             =   255
         Width           =   7560
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vs中药规格 
      Height          =   1365
      Left            =   30
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4785
      Width           =   10410
      _cx             =   18362
      _cy             =   2408
      Appearance      =   1
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   4210752
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmCHRecipe.frx":058A
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   1
      MergeCompare    =   0
      AutoResize      =   0   'False
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
      Begin VB.CommandButton cmd形态 
         Caption         =   "换为散装"
         Height          =   330
         Left            =   6675
         TabIndex        =   25
         Top             =   630
         Width           =   1245
      End
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   21
      Top             =   8175
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2619
            MinWidth        =   882
            Picture         =   "frmCHRecipe.frx":061C
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15637
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   370
            MinWidth        =   88
            Object.Tag             =   "中药味数"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
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
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   15
      TabIndex        =   26
      Top             =   495
      Width           =   10455
      Begin VB.PictureBox pic形态 
         BorderStyle     =   0  'None
         Height          =   390
         Left            =   45
         ScaleHeight     =   390
         ScaleWidth      =   10260
         TabIndex        =   27
         Top             =   165
         Width           =   10260
         Begin VB.OptionButton opt形态 
            Caption         =   "散装(&0)"
            Height          =   420
            Index           =   0
            Left            =   750
            TabIndex        =   3
            Top             =   -15
            Width           =   1245
         End
         Begin VB.OptionButton opt形态 
            Caption         =   "饮片(&1)"
            Height          =   420
            Index           =   1
            Left            =   1980
            TabIndex        =   4
            Top             =   -15
            Width           =   1245
         End
         Begin VB.OptionButton opt形态 
            Caption         =   "免煎剂(&2)"
            Height          =   420
            Index           =   2
            Left            =   3225
            TabIndex        =   5
            Top             =   0
            Value           =   -1  'True
            Width           =   1410
         End
         Begin VB.ComboBox cbo药房 
            Height          =   360
            Left            =   7635
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   30
            Width           =   2625
         End
         Begin VB.TextBox txt付数 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            IMEMode         =   3  'DISABLE
            Left            =   6315
            MaxLength       =   3
            TabIndex        =   7
            Text            =   "1"
            Top             =   45
            Width           =   495
         End
         Begin VB.Label lbl形态 
            AutoSize        =   -1  'True
            Caption         =   "形态"
            Height          =   240
            Left            =   90
            TabIndex        =   2
            Top             =   60
            Width           =   480
         End
         Begin VB.Label lbl药房 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "药房"
            Height          =   240
            Left            =   7035
            TabIndex        =   8
            Top             =   90
            Width           =   480
         End
         Begin VB.Label lbl付数 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "付数"
            Height          =   240
            Left            =   5760
            TabIndex        =   6
            Top             =   90
            Width           =   480
         End
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsBill 
      Height          =   3525
      Left            =   30
      TabIndex        =   10
      Top             =   1215
      Width           =   10425
      _cx             =   18389
      _cy             =   6218
      Appearance      =   1
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   4210752
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   3
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   11
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmCHRecipe.frx":0EB0
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
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
   Begin VB.Frame fraSplit 
      Height          =   705
      Left            =   45
      TabIndex        =   28
      Top             =   6120
      Width           =   10410
      Begin VB.TextBox txt应收 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   6915
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   225
         Width           =   1395
      End
      Begin VB.TextBox txt实收 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   8955
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   225
         Width           =   1395
      End
      Begin VB.ComboBox cbo煎法 
         Height          =   360
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   240
         Width           =   1800
      End
      Begin VB.Label lbl应收 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "应收"
         Height          =   240
         Left            =   6360
         TabIndex        =   14
         Top             =   300
         Width           =   480
      End
      Begin VB.Label lbl实收 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "实收"
         Height          =   240
         Left            =   8400
         TabIndex        =   16
         Top             =   300
         Width           =   480
      End
      Begin VB.Label lbl煎法 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "煎法"
         Height          =   240
         Left            =   60
         TabIndex        =   12
         Top             =   300
         Width           =   480
      End
   End
   Begin VB.Frame fra规格 
      Height          =   1380
      Left            =   45
      TabIndex        =   29
      Top             =   6765
      Width           =   10410
      Begin VSFlex8Ctl.VSFlexGrid vsSpecShow 
         Height          =   1095
         Left            =   60
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   195
         Width           =   10275
         _cx             =   18124
         _cy             =   1931
         Appearance      =   0
         BorderStyle     =   0
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
         BackColor       =   -2147483644
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483644
         BackColorAlternate=   -2147483644
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483644
         FloodColor      =   192
         SheetBorder     =   -2147483644
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   7
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmCHRecipe.frx":0F0B
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
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   8430
      Left            =   10515
      TabIndex        =   31
      Top             =   -150
      Width           =   45
   End
End
Attribute VB_Name = "frmCHRecipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'入口参数：
Private mbytFun As Byte  '0-记帐,1-划价
Private mstrPrivs As String
Private mstrPrivsOpt As String '记帐操作1150模块的授权功能
Private mlng病人ID As Long
Private mint病人来源 As Integer '从调用方传入
Private mlng病人科室ID As Long '病人科室
Private mlng开单科室ID As Long '开单科室ID
Private mlng中药房 As Long '中药房ID
Private mobjDetails As BillDetails
Private mstr费别 As String
Private mint险类 As Integer '如果是医保病人，则为病人险类
Private mbln加班 As Boolean
Private mcolStock As Collection '存放各个药品库房的出库检查方式
Private mrsPati As ADODB.Recordset
Private mrsWarn As ADODB.Recordset
Private mstrWarn As String
Private mblnFirst As Boolean

Private mblnReturn As Boolean
Private mblnChange As Boolean
Private mblnOK As Boolean

Private mcurModiMoney As Currency
Private mcur非中药金额 As Currency      '进入配方之前,单据的金额,用来报警时计算当前单据金额

Public mstr煎法 As String   'out
Private mcll规格  As Collection  '以品种ID为主键的数据:规格1,数量;规格2,数量|未分配数量
Private mcllInput规则摘要 As Collection  '以药品ID为主键记录曾经录入过的摘要
Private mint中药形态 As Integer
Private Const mlngModul = 1150
Private Const MIPTS = 4 '配方分栏数
Private Const MCOLS = 3 '每一栏列数
Private Const MROWS = 12 '界面可见行数
Private Const STR_HEAD = "中草药,1280,1;剂量,700,7;,400,1"
Private Enum COL_BILL
    col中药 = 0
    col剂量 = 1
    col单位 = 2
End Enum
'--自动分配规格及分中药形态(散装,饮片,免煎剂)的问题是:31867
Public Function ShowMe(frmParent As Object, ByVal strPrivs As String, ByVal bytFun As Byte, ByVal curModiMoney As Currency, _
    ByVal lng病人ID As Long, ByVal int病人来源, ByVal lng病人科室ID As Long, ByVal lng开单科室ID As Long, ByVal lng中药房 As Long, _
    ByVal objDetails As BillDetails, ByVal str费别 As String, _
    ByVal int险类 As Integer, ByVal bln加班 As Boolean, ByVal str煎法 As String, rsWarn As ADODB.Recordset, colStock As Collection, _
    Optional int中药形态 As Integer = -1) As BillDetails
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示中药配方编辑界面(程序入口)
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2010-02-02 14:37:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    
    mbytFun = bytFun  '0-记帐,1-划价  暂时未用到此变量,为将来预留
    mstrPrivs = strPrivs
    mstrPrivsOpt = GetInsidePrivs(Enum_Inside_Program.p记帐操作)
    
    mcurModiMoney = curModiMoney
    mlng病人ID = lng病人ID
    mint病人来源 = int病人来源
    mlng病人科室ID = lng病人科室ID
    mlng开单科室ID = lng开单科室ID
    mlng中药房 = lng中药房
    mstr费别 = str费别
    mint险类 = int险类
    mbln加班 = bln加班
    mstr煎法 = str煎法
    Set mrsWarn = rsWarn
    Set mcolStock = colStock
    mint中药形态 = int中药形态

    mcur非中药金额 = 0
    
    '加入传入的单据明细内容的中草药行
    Set mobjDetails = New BillDetails
    For i = 1 To objDetails.Count
        With objDetails(i)
            If .收费类别 = "7" Then
                 Call mobjDetails.Add(.Detail, .收费细目ID, .序号, .从属父号, .病人ID, .主页ID, .病区ID, .科室ID, _
                 .姓名, .性别, .年龄, .住院号, .床号, .费别, .病人性质, .收费类别, .计算单位, .发药窗口, .付数, .数次, _
                 .附加标志, .执行部门ID, .InComes, .就诊卡号, "", .担保额, .医疗付款, .保险项目否, .保险大类ID, .保险编码, .摘要)
            Else
                For j = 1 To .InComes.Count
                    mcur非中药金额 = mcur非中药金额 + .InComes(j).实收金额
                Next
            End If
        End With
    Next
    On Error Resume Next
    Me.Show 1, frmParent
    On Error GoTo 0
    If mblnOK Then
        Set ShowMe = mobjDetails
    End If
End Function

Private Sub cbo煎法_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo药房_Click()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    
    If Visible = False Then Exit Sub
    If cbo药房.ListIndex < 0 Then Exit Sub
    
    
    If Val(cbo药房.Tag) <> cbo药房.ItemData(cbo药房.ListIndex) Then
        Call 重新刷新所有中药规格
         cbo药房.Tag = cbo药房.ItemData(cbo药房.ListIndex)
        Call ReCalc应收合计
        mblnChange = True
        Call ShowSpecs(Val(vsBill.Cell(flexcpData, vsBill.Row, (vsBill.Col \ MCOLS) * MCOLS + 2)))
    End If
End Sub

Private Sub cbo药房_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub chkAuto_Click()
    txtAuto.Enabled = chkAuto.Value = 1
    If txtAuto.Enabled And Visible Then txtAuto.SetFocus
End Sub

Private Sub chkAuto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim blnCancel As Boolean, i As Long, j As Long, strStock As String
    
    If mobjDetails.Count = 0 Then
        MsgBox "请在配方中至少输入一味中草药。", vbInformation, gstrSysName
        vsBill.Row = vsBill.FixedRows
        vsBill.Col = vsBill.FixedCols
        vsBill.SetFocus: Exit Sub
    End If
    If cbo药房.Visible And cbo药房.ListIndex = -1 Then
        MsgBox "请确定中药配方的发药药房。", vbInformation, gstrSysName
        cbo药房.SetFocus: Exit Sub
    End If
    
     '记录所选中药煎法
    mstr煎法 = Mid(cbo煎法.Text, InStr(1, cbo煎法.Text, "-") + 1)
    
    '强行使输入付数生效
    If Me.ActiveControl Is txt付数 Then
        Call txt付数_Validate(blnCancel)
        If blnCancel Then Exit Sub
    End If
    
    '库存检查:不必在cbo药房的Click中检查
    Dim lng药名ID As Long
    For i = 1 To mobjDetails.Count
        With mobjDetails(i)
            If InStr(1, mstrPrivsOpt, ";显示库存;") > 0 Then
                strStock = FormatEx(.Detail.库存, 5) & IIf(gbln住院单位, .Detail.住院单位, .计算单位)
            End If
            lng药名ID = mobjDetails(i).Detail.药名ID
            If .付数 * .数次 > .Detail.库存 Then
                If Not gbln分离发药 Then
                    If .Detail.分批 Or .Detail.变价 Then
                        MsgBox """" & .Detail.名称 & """为分批或时价药品，当前库存" & strStock & ",不足输入数量。", vbInformation, gstrSysName
                        Exit For
                    ElseIf mcolStock("_" & .执行部门ID) <> 0 Then
                        If mcolStock("_" & .执行部门ID) = 1 Then
                            If MsgBox("""" & .Detail.名称 & """的当前库存" & strStock & ",不足输入数量，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                Exit For
                            End If
                        ElseIf mcolStock("_" & .执行部门ID) = 2 Then
                            MsgBox """" & .Detail.名称 & """的当前库存" & strStock & ",不足输入数量。", vbInformation, gstrSysName
                            Exit For
                        End If
                    End If
                ElseIf gblnStock And gstr中药房 <> "" Then
                    MsgBox "[" & .Detail.名称 & "]的当前库存" & strStock & ",不足输入数量。", vbInformation, gstrSysName
                    Exit For
                End If
            End If
        End With
    Next
    
    '出错了,需要移动到出错列
    If i <= mobjDetails.Count Then
        With vsBill
            For i = 1 To .Rows - 1
                For j = 0 To .Cols - 1 Step MCOLS
                    If lng药名ID = Val(.Cell(flexcpData, i, j + 2)) Then
                        .Row = i: .Col = j
                        If vsBill.Editable And vsBill.Visible Then vsBill.SetFocus
                        vsBill.ShowCell .Row, .Col
                        Exit Sub
                    End If
                Next
            Next
        End With
        Exit Sub
    End If
    
    '重新计算数据,由于输入的过程经常进行删改,因此,明细数据可能不一致,因此,需要重新整理
    Dim ObjBillDetails As BillDetails
    Set ObjBillDetails = New BillDetails
    Dim q As Integer, intRow As Integer
    
    With vsBill
        intRow = 1
        For i = 1 To .Rows - 1
            For j = 0 To .Cols - 1 Step MCOLS
                lng药名ID = Val(.Cell(flexcpData, i, j + 2))
                If opt形态(0).Value = False And lng药名ID <> 0 Then
                
                    '非散装,需要检查是否分摊完成
                    If InStr(1, mcll规格("_" & lng药名ID), "|") > 0 Or mcll规格("_" & lng药名ID) = "" Then
                            ShowMsgbox "药名为" & .TextMatrix(i, j) & "的草药未分配完成,不能继续!"
                            .Row = i: .Col = j
                            If vsBill.Enabled Then vsBill.SetFocus
                            Exit Sub
                    End If
                End If
                If lng药名ID <> 0 Then
                For q = 1 To mobjDetails.Count
                    If lng药名ID = mobjDetails(q).Detail.药名ID Then
                        '重新赋值
                        With mobjDetails(q)
                            ObjBillDetails.Add .Detail, .收费细目ID, intRow, .从属父号, .病人ID, .主页ID, .病区ID, .科室ID, .姓名, .性别, .年龄, _
                                .住院号, .床号, .费别, .病人性质, .收费类别, .计算单位, .发药窗口, .付数, .数次, .附加标志, .执行部门ID, .InComes, _
                                .就诊卡号, .Key, .担保额, .医疗付款, .保险项目否, .保险大类ID, .保险编码, .摘要, .原始数量, .原始执行部门ID, .婴儿费
                        End With
                        intRow = intRow + 1
                    End If
                Next
            End If
            Next
        Next
    End With
    Set mobjDetails = ObjBillDetails
    
    mblnOK = True
    Unload Me
End Sub

Private Sub cmd形态_Click()
    '功能：单味药，非散装分配不完时，换成散装规格
    Dim lng药名ID As Long, dbl数量 As Double, lng药品ID As Long
    Dim objBillDetail As BillDetail, objDetail As Detail
    With vsBill
        lng药名ID = Val(.Cell(flexcpData, .Row, (.Col \ MCOLS) * MCOLS + 2))
        dbl数量 = Val(.TextMatrix(.Row, (.Col \ MCOLS) * MCOLS + 1))
        lng药品ID = Val(cmd形态.Tag)    '缺省规格
        If zlGetDetail(lng药品ID, dbl数量, objDetail) = False Then
            Exit Sub
        End If
        dbl数量 = FormatEx(dbl数量 / IIf(objDetail.剂量系数 = 0, 1, objDetail.剂量系数), 5)
        If CheckStock(lng药品ID, dbl数量, objDetail) = False Then
            Exit Sub
        End If
        
        '删除药名为
        Call DeleteDetails(lng药名ID)
        Call mcll规格.Remove("_" & lng药名ID)
        mcll规格.Add lng药品ID & "," & dbl数量, "_" & lng药名ID
        
         '设置明细
         If SetBillDetail(lng药品ID, dbl数量, 1, Nothing, objBillDetail) = False Then
            '分解失败
            Call DeleteDetails(lng药名ID)
             Call ReCalc应收合计
         Else
            '设置收费项目数据
            Call zlCalcMoney(objBillDetail, True)
         End If
        Call ReCalc应收合计
        Call Show中药规格(lng药名ID, dbl数量, 0)
        .SetFocus
    End With
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If mobjDetails Is Nothing Then Exit Sub
    If mobjDetails.Count <> 0 Then
        vsBill.SetFocus
    Else
        If opt形态(0).Enabled Then opt形态(0).SetFocus
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        If cmdOK.Enabled And cmdOK.Visible Then Call cmdOK_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
    mblnFirst = True
    mstrWarn = ""
    
    mblnOK = False
    mblnChange = False
                        
                       
    chkAuto.Value = IIf(zlDatabase.GetPara("中药自动输入", glngSys, mlngModul) = "1", 1, 0)
    txtAuto.Text = Val(zlDatabase.GetPara("中药自动输入长度", glngSys, mlngModul, 5))
            
    
    '初始化数据
    If Not InitData Then Unload Me: Exit Sub
                        
    '显示单据内容
    Call ShowDetails
    Call vsBill_GotFocus
End Sub

Private Function InitData() As Boolean
'功能：初始化相应的数据
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long, cur病人余额 As Currency
    
    On Error GoTo errH
    '读取病人信息,门诊费用多了这个条件mbytFun = 2 And
    If mlng病人ID <> 0 Then
        Set rsTmp = GetMoneyInfo(mlng病人ID, IIf(gbytBilling = 0, mcurModiMoney, 0), True, 2)
    
        If Not rsTmp Is Nothing And Not mrsWarn Is Nothing Then
            cur病人余额 = Val("" & rsTmp!预交余额) - Val("" & rsTmp!费用余额)
            If gbln报警包含划价费用 Then cur病人余额 = cur病人余额 - GetPriceMoneyTotal(1, mlng病人ID) + IIf(gbytBilling = 1, mcurModiMoney, 0)   '修改划价单且报警要算划价单时,加当前单据金额
            
            strSQL = "," & Val("" & rsTmp!预交余额) & " as 预交余额," & (Val("" & rsTmp!预交余额) - cur病人余额) & " as 费用余额," & cur病人余额 & " as 病人余额"
        Else
            strSQL = ",0 as 预交余额,0 as 费用余额,0 as 病人余额"
        End If
        '76451,冉俊明,2014-8-19
        strSQL = "Select A.姓名,A.住院号,A.当前床号 As 床号,A.病人ID,A.主页ID 主页Id,Nvl(A.当前病区ID,0) as 病区ID,Zl_Patiwarnscheme(A.病人id, A.主页ID) As 适用病人," & _
            " Decode(A.担保额,null,A.担保额,Zl_Patientsurety(A.病人ID,A.主页ID)) 担保额,zl_PatiDayCharge(A.病人ID) as 当日额" & _
            strSQL & _
            " From 病人信息 A Where A.病人ID=[1]"
        Set mrsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID)
    End If
    
    '读取中药房
    If gbln分离发药 Then
        lbl药房.Visible = False
        cbo药房.Visible = False
    Else
        Set rsTmp = GetDepartments("中药房", mint病人来源 & ",3")
        For i = 1 To rsTmp.RecordCount
            cbo药房.AddItem IIf(zlIsShowDeptCode, rsTmp!编码 & "-", "") & rsTmp!名称
            cbo药房.ItemData(cbo药房.NewIndex) = rsTmp!ID
            If rsTmp!ID = mlng中药房 Then cbo药房.ListIndex = cbo药房.NewIndex
            rsTmp.MoveNext
        Next
    End If
    
     '读取中药煎法
    strSQL = "select ID,rownum||'-'||名称 as 名称 from 诊疗项目目录 where 类别='E' and 操作类型='3' order by 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    cbo煎法.Clear
    cbo煎法.AddItem ""
    
    Do While Not rsTmp.EOF
        cbo煎法.AddItem rsTmp!名称
        rsTmp.MoveNext
    Loop
    
    If mstr煎法 <> "" Then  '单据未保存之前重新进入
        For i = 0 To cbo煎法.ListCount
            If Mid(cbo煎法.List(i), InStr(1, cbo煎法.List(i), "-") + 1) = mstr煎法 Then
                cbo煎法.ListIndex = i
                Exit For
            End If
        Next
        If i > cbo煎法.ListCount Then
            cbo煎法.AddItem mstr煎法
            cbo煎法.ListIndex = cbo煎法.NewIndex
        End If
    Else
        If cbo煎法.ListCount = 0 Then cbo煎法.Enabled = False
        '默认为不选煎法
    End If
    
    InitData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Function

Private Sub ShowDetails()
    '功能：全部刷新显示当前配方内容
    Dim cur应收 As Currency, cur实收 As Currency
    Dim lngRow As Long, lngCol As Long
    Dim i As Long, j As Long, k As Long, intIndex As Integer
    Dim str药名ID As String
    Dim varData As Variant, dbl数次 As Double
    Dim str规格数量 As String
    Dim lng药名ID As Long, dblTemp As Double
    
    Set mcll规格 = New Collection
    Set mcllInput规则摘要 = New Collection
    Call InitFace
    
    str药名ID = ""
    For i = 1 To mobjDetails.Count
        If InStr(1, str药名ID & ",", "," & mobjDetails(i).Detail.药名ID & ",") = 0 Then
            str药名ID = str药名ID & "," & mobjDetails(i).Detail.药名ID
        End If
        If i = 1 Then
            txt付数.Text = mobjDetails(i).付数
            intIndex = Decode(mobjDetails(i).Detail.中药形态, 0, 0, 1, 1, 2)
            opt形态(intIndex).Value = True
        End If
        If mobjDetails(i).摘要 <> "" Then
            mcllInput规则摘要.Add mobjDetails(i).摘要, "K" & mobjDetails(i).Detail.ID
        End If
        
        '累计费用
        For j = 1 To mobjDetails(i).InComes.Count
            cur应收 = cur应收 + mobjDetails(i).InComes(j).应收金额
            cur实收 = cur实收 + mobjDetails(i).InComes(j).实收金额
        Next
    Next
    
    varData = Split(str药名ID, ",")
    Dim str诊疗名称 As String, str计算单位 As String, dbl数量 As Double
    With vsBill
        .Redraw = flexRDNone
        For i = 1 To UBound(varData)
            lngRow = ((i - 1) \ MIPTS) + 1
            lngCol = ((i - 1) Mod MIPTS) * MCOLS
            If i = 1 Then lng药名ID = Val(varData(i))
            If lngRow > .Rows - 1 Then
                .AddItem ""
                Call SetSplitLine
            End If
            '药品id,数量;...|剩余数量
            dbl数量 = 0: str规格数量 = ""
            For j = 1 To mobjDetails.Count
                If Val(varData(i)) = mobjDetails(j).Detail.药名ID Then
                    dblTemp = mobjDetails(j).数次 * mobjDetails(j).Detail.剂量系数
                    If gbln住院单位 Then    '52722
                        dblTemp = dblTemp * mobjDetails(j).Detail.住院包装
                    End If
                    dbl数量 = dbl数量 + dblTemp
                    str诊疗名称 = mobjDetails(j).Detail.诊疗名称
                    str计算单位 = mobjDetails(j).Detail.剂量单位
                    str规格数量 = str规格数量 & ";" & mobjDetails(j).Detail.ID & "," & dblTemp
                End If
            Next
            
            If str规格数量 <> "" Then str规格数量 = Mid(str规格数量, 2)
            mcll规格.Add str规格数量, "_" & Val(varData(i))
            .TextMatrix(lngRow, lngCol) = str诊疗名称
            .TextMatrix(lngRow, lngCol + 1) = FormatEx(dbl数量, 5)
            .TextMatrix(lngRow, lngCol + 2) = str计算单位
            .Cell(flexcpData, lngRow, lngCol) = .TextMatrix(lngRow, lngCol)
            .Cell(flexcpData, lngRow, lngCol + 1) = .TextMatrix(lngRow, lngCol + 1)
            .Cell(flexcpData, lngRow, lngCol + 2) = Val(varData(i))
            If i = UBound(varData) Then
                '定位到最后一列的剂量列中
                .Row = lngRow
                If dbl数量 <> 0 Then
                    If lngCol + MCOLS > .Cols - 1 Then
                        If .Rows - 1 > .Row Then
                            .Row = .Row + 1
                        Else
                            .Rows = .Rows + 1
                            .Row = .Rows - 1
                        End If
                        .Col = .FixedCols
                    Else
                        .Col = lngCol + MCOLS
                    End If
                Else
                         .Col = lngCol + 1
                End If
            End If
        Next
        .Redraw = flexRDDirect
    End With
    
    txt应收.Text = Format(cur应收, gstrDec)
    txt实收.Text = Format(cur实收, gstrDec)
    If mobjDetails.Count > 0 Then
        If Not gbln分离发药 Then
            cbo药房.ListIndex = cbo.FindIndex(cbo药房, mobjDetails(1).执行部门ID)
        End If
        If cbo药房.ListIndex < 0 And cbo药房.ListCount > 0 Then cbo药房.ListIndex = 0
        Show中药规格 lng药名ID, Val(vsBill.TextMatrix(vsBill.Row, GetBillCol(1, vsBill.Col)))
    End If

End Sub

Private Sub InitFace()
'功能：初始化中药配方表格格式及数据
'参数：mstrExtData=包含每味中药信息及煎法信息的串,为空时表示新输入中药配方
    Dim arrCols As Variant
    Dim blnPre As Boolean, i As Integer
    
    arrCols = Split(STR_HEAD, ";")
    
    With vsBill
        blnPre = .Redraw
        .Redraw = flexRDNone
        .Rows = 0: .Cols = 0
        .Rows = MROWS: .Cols = (UBound(arrCols) + 1) * MIPTS
        .FixedCols = 0: .FixedRows = 1
        .RowHidden(0) = True
        
        For i = 0 To .Cols - 1
            .TextMatrix(0, i) = Split(arrCols(i Mod 3), ",")(0)
            .ColWidth(i) = Split(arrCols(i Mod 3), ",")(1)
            .ColAlignment(i) = Split(arrCols(i Mod 3), ",")(2)
        Next
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = 4
        .GridColor = .BackColor
        .GridColorFixed = .BackColorFixed
        
        .Editable = flexEDKbdMouse
        .Row = .FixedRows: .Col = .FixedCols
                
        Call SetSplitLine
        
        .Redraw = blnPre
    End With
    
    txt应收.Text = gstrDec
    txt实收.Text = gstrDec
    txt付数.Text = 1: txt付数.Tag = 1

End Sub

Private Sub SetSplitLine()
'功能：设置中药配方输入的三列分隔线
    Dim lngRow As Long, lngCol As Long
    Dim blnPre As Boolean, i As Long
    
    With vsBill
        blnPre = .Redraw
        lngRow = .Row: lngCol = .Col
        
        .Redraw = flexRDNone
        For i = 0 To .Cols - 1 Step MCOLS
            .Select .FixedRows, i + MCOLS - 1, .Rows - 1, i + MCOLS - 1
            .CellBorder &H808080, 0, 0, 1, 0, 0, 0
        Next
        
        .Row = lngRow: .Col = lngCol
        .Redraw = blnPre
    End With
End Sub

Private Function GetRow(ByVal lngRow As Long, ByVal lngCol As Long) As Long
'功能：获取当前单元对应的费用行号
    GetRow = (lngRow - 1) * MIPTS + lngCol \ MCOLS + 1
End Function

Private Sub opt形态_Click(Index As Integer)
    Dim lng药名ID As Long, lng药房ID As Long, str规格数量 As String, lng药品ID As Long
    Dim rsTemp As ADODB.Recordset, lngTemp As Long
    
    If Not Me.Visible Then Exit Sub
    With vsBill
        lng药名ID = Val(.Cell(flexcpData, .FixedRows, .FixedCols + 2))
        If gblnStock = False Or lng药名ID = 0 Then
            Call 重新刷新所有中药规格
            Exit Sub  '不限定库存时,退出
        End If
        
        '限定库存时,第一味约的缺省规格可能变化,变化后,可用药房就变了
         str规格数量 = mcll规格("_" & lng药名ID)
         lng药房ID = mlng中药房
         If cbo药房.ListIndex >= 0 Then lng药房ID = cbo药房.ItemData(cbo药房.ListIndex)
         If str规格数量 <> "" Then
                '确定药品ID
                Set rsTemp = Get中药规格(lng药名ID, Index)
                If rsTemp.RecordCount > 0 Then
                    lng药品ID = Val(Split(str规格数量, ",")(0))
                    If lng药品ID <> Val(rsTemp!药品ID) Then
                        If mlng病人ID <> 0 Then
                            lngTemp = Get收费执行科室ID("7", Val(NVL(rsTemp!药品ID)), NVL(rsTemp!执行科室, 0), mlng病人科室ID, mlng开单科室ID, mint病人来源, mlng中药房, mrsPati!病区ID)
                        Else
                            lngTemp = Get收费执行科室ID("7", Val(NVL(rsTemp!药品ID)), NVL(rsTemp!执行科室, 0), mlng病人科室ID, mlng开单科室ID, mint病人来源, mlng中药房)
                        End If
                       '设置中药房
                        If Not gbln分离发药 Then
                            If lngTemp <> lng药房ID And lngTemp <> 0 Then
                                cbo药房.ListIndex = cbo.FindIndex(cbo药房, lngTemp)
                            End If
                        End If
                    End If
                End If
         End If
    End With
    '形态变了，要重新分配规格和数量
    Call 重新刷新所有中药规格
End Sub

Private Sub opt形态_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtAuto_GotFocus()
    Call zlControl.TxtSelAll(txtAuto)
End Sub

Private Sub txtAuto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtAuto_Validate(Cancel As Boolean)
    If Not IsNumeric(txtAuto.Text) Then txtAuto.Text = 5
    If Val(txtAuto.Text) > 20 Then txtAuto.Text = 20
    If Val(txtAuto.Text) < 2 Then txtAuto.Text = 2
End Sub

Private Sub txt付数_GotFocus()
    Call zlControl.TxtSelAll(txt付数)
End Sub

Private Sub txt付数_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf InStr("1234567890" & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub
Private Sub txt付数_Validate(Cancel As Boolean)
    Dim cur应收 As Currency, cur实收 As Currency
    Dim i As Integer, strStock As String
    '检查输入
    If Not IsNumeric(txt付数.Text) Then
        MsgBox "请输入一个有效的数值。", vbInformation, gstrSysName
        Cancel = True: Exit Sub
    End If
    If Val(txt付数.Text) <> Int(txt付数.Text) Then
        MsgBox "中药付数应该是整数数值。", vbInformation, gstrSysName
        Cancel = True: Exit Sub
    End If
    
    If Val(txt付数.Text) = 0 Then
        MsgBox "请输入一个非零的付数。", vbInformation, gstrSysName
        Call zlControl.TxtSelAll(txt付数)
        Cancel = True: Exit Sub
    End If
    If Val(txt付数.Tag) = Val(txt付数.Text) Then Exit Sub
    
    If Get中药形态 = 0 Then
        '散装形态的,先要检查库存
        For i = 1 To mobjDetails.Count
            If CheckStock(mobjDetails(i).收费细目ID, mobjDetails(i).数次, mobjDetails(i).Detail) = False Then
                Cancel = True: Exit Sub
            End If
        Next
        '为库存充足时，设置付数
        For i = 1 To mobjDetails.Count
            mobjDetails(i).付数 = Val(txt付数.Text)
            Call zlCalcMoney(mobjDetails(i), True)
        Next
        '重算应收合计
        Call ReCalc应收合计
        txt付数.Tag = Val(txt付数.Text)
        Exit Sub
    End If
    '非散装形态的,需要重新刷新规格
    Call 重新刷新所有中药规格
    txt付数.Tag = Val(txt付数.Text)
End Sub

Private Sub vsBill_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim cur单价 As Currency, cur金额 As Currency
    Dim i As Long, strStock As String
    
    If NewRow <= 0 Or NewCol = -1 Then Exit Sub
    
    Call vsBill.ShowCell(NewRow, vsBill.LeftCol)
     
     If OldRow <> NewRow Or (OldCol \ MCOLS) <> (NewCol \ MCOLS) Then   '换行或换到另一药品列
        If vsBill.Cell(flexcpData, NewRow, (NewCol \ MCOLS) * MCOLS + 2) <> 0 Then
            Call Show中药规格(Val(vsBill.Cell(flexcpData, NewRow, (NewCol \ MCOLS) * MCOLS + 2)), Val(vsBill.TextMatrix(NewRow, (NewCol \ MCOLS) * MCOLS + 1)))
        Else
            vs中药规格.Rows = vs中药规格.FixedRows
            cmd形态.Visible = False
        End If
        Call ShowSpecs(Val(vsBill.Cell(flexcpData, NewRow, (NewCol \ MCOLS) * MCOLS + 2)))
    End If
End Sub

Private Sub vsBill_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    '单位列鼠标不可进入
    If Button = 1 And (vsBill.MouseCol Mod MCOLS) = col单位 Then Cancel = True
End Sub

Private Sub vsBill_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    '单位列按键不可进入
    If Not Visible Or vsBill.Redraw = flexRDNone Then Exit Sub
    If (NewCol Mod MCOLS) = col单位 Then
        Cancel = True
        If OldCol > NewCol Then '按键移动时跳过
            vsBill.Col = NewCol - 1
        Else
            If NewCol + 1 <= vsBill.Cols - 1 Then
                vsBill.Col = NewCol + 1
            Else
                vsBill.Col = NewCol - 1
            End If
        End If
        vsBill.Row = NewRow
    End If
End Sub

Private Sub vsBill_GotFocus()
    With vsBill
        .FocusRect = flexFocusSolid
        .HighLight = flexHighlightWithFocus
        .BackColorSel = vbBlue
    End With
End Sub

Private Sub vsBill_KeyDown(KeyCode As Integer, Shift As Integer)
'功能：删除数据行
    Dim cur应收 As Currency, cur实收 As Currency
    Dim i As Long, j As Long, k As Long
    Dim lng药名ID As Long
    
    If KeyCode = vbKeyDelete Then
        With vsBill
            If .TextMatrix(.Row, (.Col \ MCOLS) * MCOLS) <> "" Then
                If MsgBox("要删除""" & .TextMatrix(.Row, (.Col \ MCOLS) * MCOLS) & """吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                lng药名ID = Val(.Cell(flexcpData, .Row, GetBillCol(2, .Col)))
                '删除明细内容
                Call DeleteDetails(lng药名ID)
                '移除规格
                mcll规格.Remove "_" & lng药名ID
                '重算价格
                Call ReCalc应收合计
                '清除当前味药信息
                For i = 0 To MCOLS - 1
                    .TextMatrix(.Row, (.Col \ MCOLS) * MCOLS + i) = ""
                    .Cell(flexcpData, .Row, (.Col \ MCOLS) * MCOLS + i) = Empty
                Next
                
                '后面的内容向前移
                For i = .Row To .Rows - 1
                    For j = 0 To .Cols - 1 Step MCOLS
                        If Not (i = .Row And j <= (.Col \ MCOLS) * MCOLS) Then
                            For k = 0 To MCOLS - 1
                                If j = 0 Then
                                    .TextMatrix(i - 1, .Cols - (MCOLS - k)) = .TextMatrix(i, j + k)
                                    .Cell(flexcpData, i - 1, .Cols - (MCOLS - k)) = .Cell(flexcpData, i, j + k)
                                Else
                                    .TextMatrix(i, j + k - MCOLS) = .TextMatrix(i, j + k)
                                    .Cell(flexcpData, i, j + k - MCOLS) = .Cell(flexcpData, i, j + k)
                                End If
                                .TextMatrix(i, j + k) = ""
                                .Cell(flexcpData, i, j + k) = Empty
                            Next
                        End If
                    Next
                Next
                '删除多余的空行
                If .Rows > MROWS Then
                    For i = .Rows - 1 To MROWS Step -1
                        If .TextMatrix(i, 0) = "" Then
                            .RemoveItem i
                        End If
                    Next
                End If
                Call .ShowCell(.Row, .Col)
                sta.Panels(3).Text = "共" & mcll规格.Count & "味药"
            End If
        End With
    End If
End Sub

Private Sub vsBill_KeyPress(KeyAscii As Integer)
'功能：非编辑状态时，自动移动单元格
    If KeyAscii = 13 Then
        KeyAscii = 0
        '定位到下一应输入单元格
        If vsBill.TextMatrix(vsBill.Row, (vsBill.Col \ MCOLS) * MCOLS) = "" Then
            If GetRow(vsBill.Row, vsBill.Col) > 1 Then
                Call zlCommFun.PressKey(vbKeyTab)
            End If
            Exit Sub
        Else
            Call EnterNextCell(vsBill.Row, vsBill.Col)
        End If
    ElseIf KeyAscii = Asc("*") Then
        KeyAscii = 0
        If CellCanEdit(vsBill.Row, vsBill.Col) Then
            If vsBill.Col <> (vsBill.Col \ MCOLS) * MCOLS Then
                Exit Sub
            End If
            If SelectChineDrug("") = False Then Exit Sub
        End If
    End If
End Sub

Private Sub vsBill_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'功能：非回车确认完后编辑的处理(这里Text:=EditText,但ValidateEdit事件中还没有)
    If Not mblnReturn Then '非回车确认失效
        vsBill.TextMatrix(Row, Col) = CStr(vsBill.Cell(flexcpData, Row, Col))
    End If
End Sub

Private Sub vsBill_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    If Col Mod MCOLS = col中药 And chkAuto.Value = 1 Then
        '自动完成快捷中药输入
        If Len(vsBill.EditText) >= Val(txtAuto.Text) Then
            Call vsBill_KeyPressEdit(Row, Col, 13)
        End If
    ElseIf Col Mod MCOLS = col剂量 Then
        '自动完成快捷数量输入
        If InStr(gstrABC, UCase(Chr(KeyCode))) > 0 And Between(KeyCode, vbKeyA, vbKeyZ) Then
            vsBill.EditCell
            vsBill.EditText = UCase(Chr(KeyCode))
            Call vsBill_KeyPressEdit(Row, Col, 13)
            vsBill.FinishEditing False  '控件bug未生效
        End If
    End If
End Sub

Private Sub vsBill_LostFocus()
  With vsBill
        .FocusRect = flexFocusLight
        .HighLight = flexHighlightAlways
        .BackColorSel = &HE7CFBA
    End With
End Sub

Private Sub vsBill_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsBill.EditSelStart = 0
    vsBill.EditSelLength = zlCommFun.ActualLen(vsBill.EditText)
End Sub

Private Sub vsBill_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'功能：限制某些列不允许编辑(该事件后于BeforeEdit,在EditText赋值之前)
    mblnReturn = False

    '必须依次输入
    If Not CellCanEdit(Row, Col) Then Cancel = True

    If Col Mod MCOLS = col剂量 Then
        vsBill.EditMaxLength = 8
    Else
        vsBill.EditMaxLength = 0
    End If
End Sub

Private Function CellCanEdit(ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
'功能：输入中药配方时,判断指定的单元格当前是否输入内容
'说明：在配方输入表格中,如果前一个未输入,则当前不允许输入
    '定位到上一个中药输入单元
    lngCol = (lngCol \ MCOLS) * MCOLS
    If lngCol - MCOLS >= vsBill.FixedCols Then
        lngCol = lngCol - MCOLS
    Else
        If lngRow - 1 >= vsBill.FixedRows Then
            lngRow = lngRow - 1
            lngCol = vsBill.Cols - MCOLS
        Else
            CellCanEdit = True
            Exit Function
        End If
    End If
    CellCanEdit = vsBill.TextMatrix(lngRow, lngCol) <> ""
End Function

Private Sub EnterNextCell(ByVal lngRow As Long, ByVal lngCol As Long)
'功能：进入下一个中药配方的输入单元格
    '当前位置未输入中药
    If vsBill.TextMatrix(lngRow, (lngCol \ MCOLS) * MCOLS) = "" Then Exit Sub

    '剂量未输入
    If lngCol Mod MCOLS = 1 And vsBill.TextMatrix(lngRow, lngCol) = "" Then Exit Sub

    If lngCol + 1 <= vsBill.Cols - 1 Then
        If (lngCol + 1) Mod MCOLS = col单位 And lngCol \ MCOLS + 1 = MIPTS Then
            If lngRow + 1 > vsBill.Rows - 1 Then
                vsBill.AddItem "", vsBill.Rows
                Call SetSplitLine
            End If
            lngCol = 0
            lngRow = lngRow + 1
        Else
            lngCol = lngCol + 1
        End If
    Else
        If lngRow + 1 > vsBill.Rows - 1 Then
            vsBill.AddItem "", vsBill.Rows
            Call SetSplitLine
        End If
        lngRow = lngRow + 1
        lngCol = vsBill.FixedCols
    End If

    vsBill.Row = lngRow: vsBill.Col = lngCol
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange And Not mblnOK Then
        If MsgBox("配方内容已被改变，确实要放弃这些改变退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1: Exit Sub
        End If
    End If
        
    Set mcolStock = Nothing
    Set mrsWarn = Nothing
    Set mrsPati = Nothing
    
    zlDatabase.SetPara "中药自动输入", chkAuto.Value, glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
    zlDatabase.SetPara "中药自动输入长度", Val(txtAuto.Text), glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
    
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Function SelectChineDrug(ByVal strInput As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：选择中药
    '入参：strInput-要查找的值
    '出参：
    '返回：成功,返回true, 否则返回False
    '编制：刘兴洪
    '日期：2010-07-27 13:56:04
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    'gblnStock:表示收费或记帐时如果指定了分配药房，是否只能输入该药房有库存的药品
    Dim vPoint As POINTAPI, strTmp As String
    Dim rsTemp As ADODB.Recordset, str特性 As String
    Dim lng药房ID As Long, strStock As String, strSQLAdd As String, str特准项目 As String, strSQL As String
    Dim strSQLInput As String, str撤档时间 As String, strWhere As String
    Dim int中药形态 As Integer, blnCancel As Boolean
    Dim str规格 As String, str诊疗 As String, lng药品ID As Long, lngTmp As Long
    Dim lng药名ID As Long, lng上次药名ID As Long
    
    int中药形态 = Get中药形态
    '门诊费用多这个条件mbytFun = 0 And,排开了划价
    If cbo药房.ListIndex < 0 Then cbo药房.ListIndex = cbo.FindIndex(cbo药房, lng药房ID)
    If cbo药房.ListIndex < 0 And cbo药房.ListCount > 0 Then cbo药房.ListIndex = 0
    If cbo药房.ListIndex < 0 Then
        MsgBox "药房未选择,请选择药房?", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
         If cbo药房.Enabled And cbo药房.Visible Then cbo药房.SetFocus
        Exit Function
    End If
    lng药房ID = cbo药房.ItemData(cbo药房.ListIndex)
  '特殊药品权限
    str特性 = ""
    If InStr(mstrPrivsOpt, ";麻醉药品记帐;") = 0 Then str特性 = str特性 & " And E.毒理分类<>'麻醉药'"
    If InStr(mstrPrivsOpt, ";毒性药品记帐;") = 0 Then str特性 = str特性 & " And E.毒理分类<>'毒性药'"
    If InStr(mstrPrivsOpt, ";贵重药品记帐;") = 0 Then str特性 = str特性 & " And E.价值分类 Not IN('贵重','昂贵')"
    '暂未有此限制
   ' If InStr(mstrPrivsOpt, "精神药品记帐") = 0 Then str特性 = str特性 & " And E.价值分类 Not IN('精神I类','精神II类')"
    If int中药形态 = 0 And lng药房ID <> 0 Then
        '只有散装才有库存
        strStock = _
            " Select 药品ID,Sum(Nvl(可用数量,0)) as 库存 From 药品库存" & _
            "  Where (Nvl(批次, 0) = 0 Or 效期 Is Null Or 效期 > Trunc(Sysdate))" & _
            "           And 性质 = 1 And 库房ID=[4]" & _
            "  Group by 药品ID  " & _
            "  Having Sum(Nvl(可用数量,0))<>0"
     Else
        strStock = "Select NULL as 药品ID,NULL as 库存 From Dual"
     End If
    
    If int中药形态 = 0 Then
        str规格 = _
        "   And Nvl(C.中药形态,0) = [6] And (D.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or D.撤档时间 IS NULL) And D.服务对象 IN([7],3)" & _
        "   And (D.站点='" & gstrNodeNo & "' Or D.站点 is Null)" & _
         IIf(gblnStock And lng药房ID <> 0, " And nvl(X.库存,0)<>0", "")
    Else
         str规格 = " And Exists(Select 1 From 药品规格 C Where C.药名ID=E.药名ID And Nvl(C.中药形态,0) = [6])"
    End If
    
    
    str撤档时间 = "" & _
        "   And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL) " & _
        "   And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
        "   And A.服务对象 IN([7],3)"
    
    str特准项目 = ""
    If int中药形态 = 0 Then
        If mint险类 <> 0 And mlng病人ID <> 0 Then
            '刘兴洪:24862
            If zl_Check特准项目(gclsInsure, mint险类, mlng病人ID, False) Then str特准项目 = Get保险特准项目(mlng病人ID, "D.ID")
        End If
    End If
        
    If strInput <> "" Then
            strWhere = " And (A.编码 Like [1] And B.码类=[3] Or B.名称 Like [2] And B.码类=[3] Or B.简码 Like upper([2]) And B.码类 IN([3],3))"
            If IsNumeric(strInput) Then                         '10,11.输入全是数字时只匹配编码'对于药品,则要匹配简码(码类为3的数字码)
                If Mid(gstrMatchMode, 1, 1) = "1" Then strWhere = " And (A.编码 Like [1] And B.码类=[3] Or B.简码 Like Upper([2]) And B.码类=3)"
            ElseIf zlCommFun.IsCharAlpha(strInput) Then         '01,11.输入全是字母时只匹配简码
                If Mid(gstrMatchMode, 2, 1) = "1" Then strWhere = " And B.简码 Like Upper([2]) And B.码类=[3]"
            ElseIf zlCommFun.IsCharChinese(strInput) Then
                strWhere = " And B.名称 Like [2] And B.码类=[3]"
            End If
             '非散装时按品种显示，且不显示库存
            strSQL = "" & _
            "   Select  distinct A.ID,A.编码,A.名称,A.计算单位" & _
            "   From 诊疗项目目录 A,诊疗项目别名 B" & _
            "   Where A.ID=B.诊疗项目ID  And A.类别='7' " & str撤档时间 & strWhere
            
            If int中药形态 = 0 Then
                '散装才显示到规格级,保持原来不变
                strSQL = _
                " Select distinct  A.ID as 药名ID,C.药品ID as ID,C.药品ID,D.编码,A.名称,D.规格,A.计算单位 as 剂量单位," & _
                        IIf(gbln住院单位, "C.住院单位", "D.计算单位") & " as 单位,D.产地,D.费用类型,d.执行科室 AS 执行科室_ID," & IIf(mint险类 <> 0, "N.名称 医保大类,", "") & _
                "       Decode(D.是否变价,1,'时价',LTrim(To_Char(Sum(F.现价)" & _
                        IIf(gbln住院单位, "*Nvl(C.住院包装,1)", "") & ",'999999" & gstrFeePrecisionFmt & "'))) as 单价," & _
                        IIf(InStr(1, mstrPrivsOpt, ";显示库存;") > 0, " LTrim(To_Char(X.库存" & IIf(gbln住院单位, "/Nvl(C.住院包装,1)", "") & ",'9999990.00000'))", "Decode(Sign(X.库存),1,'有','无')") & " as 库存" & _
                " From 药品特性 E,药品规格 C,收费项目目录 D,收费价目 F, " & vbNewLine & _
                            IIf(mint险类 <> 0, "保险支付项目 M,保险支付大类 N,", "") & vbNewLine & _
                "          (" & strSQL & ") A, " & vbNewLine & _
                "          (" & strStock & ") X" & vbNewLine & _
                " Where   A.ID=E.药名ID And A.ID=C.药名ID And C.药品ID=D.ID And C.药品ID=X.药品ID(+) " & vbNewLine & _
                "        And D.ID=F.收费细目ID " & vbNewLine & _
                         IIf(mint险类 <> 0, " And C.药品ID=M.收费细目ID(+) And M.险类(+)=[5] And M.大类ID=N.ID(+)" & vbNewLine, "") & _
                "        And exists(Select 1 From 收费执行科室 A1 Where A1.收费细目ID=C.药品ID And A1.执行科室ID=[4]   And (A1.病人来源 is NULL Or A1.病人来源=[7]) and (A1.开单科室ID is null or A1.开单科室ID=[8])  ) " & vbNewLine & _
                "        And Sysdate Between F.执行日期 and Nvl(F.终止日期,TO_DATE('3000-01-01','YYYY-MM-DD'))" & _
                         str规格 & str特性 & str特准项目 & _
                " Group by A.ID,C.药品ID,A.计算单位,D.编码,A.名称,D.规格,D.产地,D.费用类型,d.执行科室,D.是否变价," & IIf(mint险类 <> 0, "N.名称,", "") & "X.库存," & _
                        IIf(gbln住院单位, "C.住院单位,C.住院包装", "D.计算单位") & _
                " Order by D.编码"
            Else
                 '非散装时按品种显示，且不显示库存
                strSQL = strSQL & _
                "        And exists(Select 1 From 诊疗执行科室 A1 Where A1.诊疗项目ID=A.ID And A1.执行科室ID=[4]   And (A1.病人来源 is NULL Or A1.病人来源=[7]) and (A1.开单科室ID is null or A1.开单科室ID=[8])  ) " & vbNewLine
                strSQL = _
                    " Select Distinct A.ID,A.ID as 药名ID,A.编码,A.名称,A.计算单位 as 单位" & _
                    " From 药品特性 E,(" & strSQL & ") A" & _
                    " Where A.ID=E.药名ID  " & _
                    "         And Exists(Select 1 From 药品规格 C Where C.药名ID=E.药名ID And Nvl(C.中药形态,0) = [6])" & _
                    "         And Rownum<=100" & _
                    " Order by A.编码"
            End If
        
            vPoint = zlControl.GetCoordPos(vsBill.hWnd, vsBill.CellLeft, vsBill.CellTop)
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "中草药", False, "", "", False, False, True, vPoint.X, vPoint.Y, _
                        vsBill.CellHeight, blnCancel, True, True, strInput & "%", gstrLike & strInput & "%", gbytCode + 1, lng药房ID, mint险类, int中药形态, mint病人来源, mlng开单科室ID)
    Else
            If int中药形态 = 0 Then
                '散装才显示到规格级,保持原来不变
            strSQL = "" & _
                " Select 0 as 末级,ID,ID as 药名ID ,上级ID,编码,名称,Null as 规格,NULL as 剂量单位,NULL as 单位," & _
                "       NULL as 产地,NULL as 费用类型 , NULL as 执行科室_ID" & IIf(mint险类 = 0, "", ",Null as 医保大类") & ",NULL as 单价,NULL as 库存,NULL as 药品ID" & _
                " From 诊疗分类目录 " & _
                "  Where 类型=3 And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null)" & _
                " Start With 上级ID is NULL Connect by Prior ID=上级ID" & _
                " Union ALL "
                strSQL = strSQL & _
                "  Select 末级,-1*Rownum As Id,药名ID,上级ID,编码,名称,规格,剂量单位,单位,产地,费用类型,执行科室_ID,单价,库存,药品ID " & _
                "  From ( " & _
                " Select 1 as 末级,A.ID,A.ID as 药名ID,A.分类ID as 上级ID,D.编码,D.名称,D.规格,A.计算单位 as 剂量单位," & _
                            IIf(gbln住院单位, " C.住院单位", "D.计算单位") & " as 单位,D.产地,D.费用类型,d.执行科室 as 执行科室_ID" & IIf(mint险类 = 0, "", ",N.名称 医保大类") & "," & _
                "           Decode(D.是否变价,1,'时价',LTrim(To_Char(Sum(F.现价)" & _
                            IIf(gbln住院单位, "*Nvl(C.住院包装,1)", "") & ",'999999" & gstrFeePrecisionFmt & "'))) as 单价," & _
                            IIf(InStr(1, mstrPrivsOpt, ";显示库存;") > 0, " LTrim(To_Char(X.库存" & IIf(gbln住院单位, "/Nvl(C.住院包装,1)", "") & ",'999999" & gstrFeePrecisionFmt & "'))", "Decode(Sign(X.库存),1,'有','无')") & " as 库存,C.药品ID" & _
                " From  诊疗项目目录 A,药品特性 E,药品规格 C,收费项目目录 D,收费价目 F," & _
                            IIf(mint险类 = 0, "", "           保险支付项目 M,保险支付大类 N,") & _
                "           (" & strStock & ") X" & _
                " Where A.ID=E.药名ID And A.ID=C.药名ID And C.药品ID=D.ID And C.药品ID =F.收费细目ID And A.类别='7'  " & _
                        IIf(mint险类 = 0, "", "       And C.药品ID=M.收费细目ID(+) And   M.险类(+)=" & mint险类 & " And M.大类ID=N.ID(+)") & _
                "       And C.药品ID=X.药品ID(+) " & _
                "        And exists(Select 1 From 收费执行科室 A1 Where A1.收费细目ID=C.药品ID And A1.执行科室ID=[4]   And (A1.病人来源 is NULL Or A1.病人来源=[7]) and (A1.开单科室ID is null or A1.开单科室ID=[8])  ) " & vbNewLine & _
                "       And Sysdate Between F.执行日期 and Nvl(F.终止日期,TO_DATE('3000-01-01','YYYY-MM-DD'))" & _
                "       And D.服务对象 IN(" & mint病人来源 & ",3)" & str特准项目 & str规格 & str撤档时间 & _
                " Group by A.ID,A.计算单位 ,A.分类ID,D.编码,D.名称,D.规格,D.产地,D.费用类型,d.执行科室" & IIf(mint险类 = 0, "", ",N.名称") & ",D.是否变价,X.库存,C.药品ID," & _
                     IIf(gbln住院单位, "C.住院单位,C.住院包装", "D.计算单位") & _
                ")"
            Else

                strSQL = "" & _
                " Select 0 as 末级,ID,ID as 药名ID,上级ID,编码,名称,NULL as 单位,NULL as 处方职务ID" & _
                " From 诊疗分类目录 Where 类型=3 And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " Start With 上级ID is NULL Connect by Prior ID=上级ID"
                strSQL = strSQL & " UNION ALL " & _
                "Select Distinct 1 as 末级,A.ID,ID as 药名ID,A.分类ID as 上级ID,A.编码,A.名称,A.计算单位 as 单位,E.处方职务 as 处方职务ID" & _
                " From 诊疗项目目录 A,药品特性 E" & _
                " Where A.ID=E.药名ID" & str特性 & str撤档时间 & str规格 & _
                "        And exists(Select 1 From 诊疗执行科室 A1 Where A1.诊疗项目ID=A.ID And A1.执行科室ID=[4]   And (A1.病人来源 is NULL Or A1.病人来源=[7]) and (A1.开单科室ID is null or A1.开单科室ID=[8])  ) " & vbNewLine
            End If
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "中草药", False, "", "", False, True, False, 0, 0, 0, blnCancel, False, True, "", "", "", lng药房ID, "", int中药形态, mint病人来源, mlng开单科室ID)
    End If
    
    With vsBill
        If rsTemp Is Nothing Then
            If Not blnCancel Then
                MsgBox "未找到可用的中药项目，请先到诊疗项目管理中设置。！", vbInformation, gstrSysName
            End If
            With vsBill
              If strInput <> "" Then .TextMatrix(.Row, .Col) = CStr(.Cell(flexcpData, .Row, .Col))
              Exit Function
            End With
        End If
        
        lng药名ID = Val(NVL(rsTemp!药名ID)): lng药品ID = 0
        If int中药形态 = 0 Then lng药品ID = Val(NVL(rsTemp!药品ID))
        
        If ItemExist(lng药名ID, .Row, .Col) Then
           MsgBox "该味中药在配方中已经录入。", vbInformation, gstrSysName
            If strInput <> "" Then .TextMatrix(.Row, .Col) = CStr(.Cell(flexcpData, .Row, .Col))
            Exit Function
        End If
        
        lng上次药名ID = Val(.Cell(flexcpData, .Row, GetBillCol(2, .Col))) '上次的品种
        
        lng药品ID = -1
        If lng上次药名ID <> 0 Then
            If int中药形态 = 0 Then   '如果是第一味散装药，规格变了，可用药房跟着改变
                If .Row = .FixedRows And .Col = .FixedCols Then
                    If mcll规格("_" & lng上次药名ID) <> "" Then
                        lng药品ID = Val(Split(mcll规格("_" & lng上次药名ID), ",")(0))
                    Else
                        lng药品ID = 0
                    End If
                End If
            End If
            mcll规格.Remove "_" & lng上次药名ID
        End If
        
        '获取输入值
        If strInput <> "" Then .EditText = rsTemp!名称     '直接输入匹配时必要
         .TextMatrix(.Row, .Col) = rsTemp!名称
         If int中药形态 = 0 Then
            .TextMatrix(.Row, .Col + 2) = NVL(rsTemp!剂量单位)
         Else
            .TextMatrix(.Row, .Col + 2) = rsTemp!单位
         End If
         .Cell(flexcpData, .Row, .Col) = .TextMatrix(.Row, .Col)
         .Cell(flexcpData, .Row, .Col + 2) = lng药名ID    '记录中药ID
         If lng上次药名ID <> lng药名ID And lng上次药名ID <> 0 Then
            '删除上次药名ID
            Call DeleteDetails(lng上次药名ID)
            Err = 0: On Error Resume Next
            mcll规格.Remove "_" & lng上次药名ID
            Err = 0: On Error GoTo 0
         End If
         
        If mcll规格 Is Nothing Then Set mcll规格 = New Collection
        If int中药形态 = 0 Then
            Err = 0: On Error Resume Next
            mcll规格.Remove "_" & lng药名ID
            Err = 0: On Error GoTo 0
            mcll规格.Add NVL(rsTemp!药品ID) & ",0", "_" & lng药名ID
            
            '如果散装药品的规格变了，重设可用药房
            If lng药品ID <> Val(NVL(rsTemp!药品ID)) Then
                If cbo药房.ListIndex < 0 Then
                    If mlng病人ID <> 0 Then
                        lng药房ID = Get收费执行科室ID("7", Val(NVL(rsTemp!药品ID)), NVL(rsTemp!执行科室_ID, 0), mlng病人科室ID, mlng开单科室ID, mint病人来源, mlng中药房, mrsPati!病区ID)
                    Else
                        lng药房ID = Get收费执行科室ID("7", Val(NVL(rsTemp!药品ID)), NVL(rsTemp!执行科室_ID, 0), mlng病人科室ID, mlng开单科室ID, mint病人来源, mlng中药房)
                    End If
                Else
                    lng药房ID = cbo药房.ItemData(cbo药房.ListIndex)
                End If
                
                '设置中药房
                If Not gbln分离发药 Then
                    If cbo药房.ListIndex <> -1 Then
                        lngTmp = cbo药房.ItemData(cbo药房.ListIndex)
                    End If
                    If lngTmp <> lng药房ID And lng药房ID <> 0 Then
                        cbo药房.ListIndex = cbo.FindIndex(cbo药房, lng药房ID)
                        '改变了库房,需要重新刷新规格
                         Call 重新刷新所有中药规格
                    End If
                End If
            End If
        Else
            Err = 0: On Error Resume Next
            mcll规格.Remove "_" & lng药名ID
            Err = 0: On Error GoTo 0
            mcll规格.Add "", "_" & lng药名ID
            If cbo药房.ListIndex < 0 Then
                    If mlng中药房 <> 0 Then
                        cbo药房.ListIndex = cbo.FindIndex(cbo药房, mlng中药房)
                    Else
                       If cbo药房.ListCount <> 0 Then cbo药房.ListIndex = 0
                    End If
            End If
        End If
         If cbo药房.ListCount <> 0 And cbo药房.ListIndex < 0 Then cbo药房.ListIndex = 0
        '已输入数量时，修改药名
        Call 分解中药规格(lng药名ID, Val(.TextMatrix(.Row, .Col + 1)))
        If Val(.TextMatrix(.Row, .Col + 1)) <> 0 Then
            Call Show中药规格(lng药名ID, Val(.TextMatrix(.Row, .Col + 1)))
        End If
    End With
    Call ShowSpecs(lng药名ID)
    '问题:39319
    If Not mcll规格 Is Nothing Then
        sta.Panels(3).Text = "共" & mcll规格.Count & "味药"
    End If
    SelectChineDrug = True
End Function



Private Function ItemExist(ByVal lng中药ID As Long, ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
    '功能：判断中药配方输入表格中,指定的中药是否已经输入
    Dim i As Long, j As Long, lngTemp As Long
    Dim lngCurCol As Long
    
    lngTemp = GetBillCol(2, lngCol)
    For i = 1 To vsBill.Rows - 1
        For j = 0 To vsBill.Cols - 1 Step MCOLS
            lngCurCol = GetBillCol(2, j)
            If lngRow = i And lngTemp <> lngCurCol Or lngRow <> i Then
                If Val(vsBill.Cell(flexcpData, i, lngCurCol)) = lng中药ID Then
                    ItemExist = True
                    Exit Function
                End If
            End If
        Next
    Next
End Function
Private Function is分批药品(ByVal lng品名ID As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：是否分批药品
    '返回：是返回true,否则返回False
    '编制：刘兴洪
    '日期：2010-08-03 17:05:01
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim lng药品ID As Long, rsTemp As ADODB.Recordset
    Dim i As Long, strSQL As String
    
    If lng品名ID = 0 Then Exit Function
    Err = 0: On Error Resume Next
    lng药品ID = Val(Split(mcll规格("_" & lng品名ID) & ",", ",")(0))
    If Err <> 0 Then
        ShowMsgbox "未找到数据,请检查!"
        is分批药品 = False
        Exit Function
    End If
    If lng药品ID = 0 Then Exit Function
    For i = 1 To mobjDetails.Count
        If mobjDetails(i).Detail.ID = lng药品ID Then
           is分批药品 = mobjDetails(i).Detail.分批
           Exit Function
        End If
    Next
    
    On Error GoTo errHandle
    
    '未有数据,直接从库中读取
    strSQL = "Select 药房分批 From 药品规格 where 药品ID=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng药品ID)
    If Not rsTemp.EOF Then
        is分批药品 = NVL(rsTemp!药房分批, 0) <> 0
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetBillCol(ByVal int性质 As Integer, lngCol As Long) As Long
    '------------------------------------------------------------------------------------------------------------------------
    '功能：获取指定性质的列
    '入参:  int性质-0:草药名称列,1-剂量列;2-单位列
    '返回：指定性质的列
    '编制：刘兴洪
    '日期：2010-08-03 17:29:28
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    GetBillCol = (lngCol \ MCOLS) * MCOLS + int性质
End Function
Private Function GetBillDetailObject(ByVal lng药品ID As Long) As BillDetail
    '获取明细数据对象
    Dim i As Long
    For i = 1 To mobjDetails.Count
         If mobjDetails(i).Detail.ID = lng药品ID Then
            Set GetBillDetailObject = mobjDetails(i)
            Exit Function
         End If
    Next
End Function
Private Function CheckStock(ByVal lng药品ID As Long, dbl数量 As Double, Optional objDetail As Detail) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：检查药品库存
    '返回：存在库存,返回True,否则返回False
    '编制：刘兴洪
    '日期：2010-08-04 11:23:47
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim lng药房ID As Long
    If lng药品ID = 0 Then CheckStock = True: Exit Function
    
    If objDetail Is Nothing Then
        Set objDetail = GetBillDetailObject(lng药品ID).Detail
    End If
    
    If objDetail Is Nothing Then CheckStock = True: Exit Function
    If cbo药房.ListIndex < 0 Then
        If cbo药房.ListCount = 0 Then '33188
            CheckStock = True: Exit Function
        End If
        cbo药房.ListIndex = 0
    End If
    lng药房ID = cbo药房.ItemData(cbo药房.ListIndex)
    
    '药品库存检查
    With objDetail
        If Not gbln分离发药 Then
            If .分批 Or .变价 Then
                If Val(txt付数.Text) * dbl数量 > .库存 Then
                    MsgBox """" & .名称 & """为分批或时价药品，当前可用库存不足输入数量。", vbInformation, gstrSysName
                    Exit Function
                End If
            ElseIf mcolStock("_" & lng药房ID) <> 0 Then
                If Val(txt付数.Text) * dbl数量 > .库存 Then
                    If mcolStock("_" & lng药房ID) = 1 Then
                        If MsgBox("""" & .名称 & """的当前可用库存不足输入数量，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                             Exit Function
                        End If
                    ElseIf mcolStock("_" & lng药房ID) = 2 Then
                        MsgBox """" & .名称 & """的当前可用库存不足输入数量。", vbInformation, gstrSysName
                         Exit Function
                    End If
                End If
            End If
        ElseIf gstr中药房 <> "" And Val(txt付数.Text) * dbl数量 > .库存 Then
            If gblnStock Then
                MsgBox "[" & .名称 & "]的当前可用库存不足输入数量!", vbInformation, gstrSysName
                 Exit Function
            Else
                If MsgBox("[" & .名称 & "]的当前可用库存不足输入数量，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            End If
        End If
    End With
    CheckStock = True
End Function
Private Function IsCheckStockEnough(ByVal lng药品ID As Long, dbl数量 As Double, Optional objDetail As Detail = Nothing) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：判断库存是否充足
    '返回：充足返回true,否则返回False
    '编制：刘兴洪
    '日期：2010-08-04 17:11:52
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    If objDetail Is Nothing Then
        Set objDetail = GetBillDetailObject(lng药品ID).Detail
    End If
    If objDetail Is Nothing Then IsCheckStockEnough = True: Exit Function
    
    With objDetail
        If Val(txt付数.Text) * dbl数量 > .库存 Then Exit Function
    End With
    IsCheckStockEnough = True
End Function
Private Sub vsBill_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
'功能：输入数据确认
    Dim rsTmp As ADODB.Recordset
    Dim lng药名ID As Long, lng药品ID As Long
    Dim strSQL As String, blnCancel As Boolean
    Dim lng药房ID As Long, strStock As String, i As Long
    Dim vPoint As POINTAPI, strTmp As String
    Dim cur应收 As Currency, cur实收 As Currency
    Dim str特准项目 As String, blnOverFlow As Boolean
    Dim strInput As String, strSQLInput As String, strSQLAdd As String, strSQLItem As String
    Dim int中药形态 As Integer
    Dim dbl系数 As Double
    
    If KeyAscii = 13 Then
        mblnReturn = True '标记是按回车确认编辑
        KeyAscii = 0
        
        '截取回车后,如果用Msgbox使Edit焦点丢失,则会完成编辑,但不会激活AfterEdit事件
        If Col Mod MCOLS = col中药 Then
            '中药输入
            If vsBill.EditText = "" Then   'zyk
                Call zlCommFun.PressKey(vbKeyTab)
                Exit Sub
            End If
            strInput = vsBill.EditText
            If SelectChineDrug(strInput) = False Then
                vsBill.TextMatrix(Row, Col) = CStr(vsBill.Cell(flexcpData, Row, Col)): Exit Sub
            End If
        ElseIf Col Mod MCOLS = col剂量 Then
            '快捷输入转换
            vsBill.EditText = ConvertABCtoNUM(vsBill.EditText)
            '剂量输入合法性检查
            If Not IsNumeric(vsBill.EditText) Or Val(vsBill.EditText) > LONG_MAX Then
                MsgBox "药品剂量输入错误，不是数值类型或输入数值过大。", vbInformation, gstrSysName
                vsBill.TextMatrix(Row, Col) = CStr(vsBill.Cell(flexcpData, Row, Col)): Exit Sub
            End If
            
            
            '中药未输入,无效
            lng药名ID = Val(vsBill.Cell(flexcpData, Row, GetBillCol(2, Col)))
            If lng药名ID = 0 Then
                MsgBox "请先输入中草药。", vbInformation, gstrSysName
                vsBill.TextMatrix(Row, Col) = CStr(vsBill.Cell(flexcpData, Row, Col)): Exit Sub
            End If
            
            If Val(vsBill.EditText) < 0 Then
                If InStr(mstrPrivsOpt, ";草药负数记帐;") = 0 Then
                    MsgBox "你没有权限输入负数。", vbInformation, gstrSysName
                    vsBill.TextMatrix(Row, Col) = CStr(vsBill.Cell(flexcpData, Row, Col)): Exit Sub
                End If
                '除了散装以外,其他的只能输入正数
                If Get中药形态 <> 0 Then
                    MsgBox "药品形态只有散装的才能负数记帐。", vbInformation, gstrSysName
                    vsBill.TextMatrix(Row, Col) = CStr(vsBill.Cell(flexcpData, Row, Col)): Exit Sub
                Else
                    '散装的,要检查是否份批的,分批的也不允许负数记帐
                    If is分批药品(Val(vsBill.Cell(flexcpData, Row, Col + 1))) Then
                        MsgBox "分批药品不允许输入负数。", vbInformation, gstrSysName
                        vsBill.TextMatrix(Row, Col) = CStr(vsBill.Cell(flexcpData, Row, Col)): Exit Sub
                    End If
                End If
                If mint险类 > 0 Then
                    If Not gclsInsure.GetCapability(support负数记帐, mlng病人ID, mint险类) Then
                        MsgBox "本地医保不支持对医保病人进行负数记帐！", vbInformation, gstrSysName
                        vsBill.TextMatrix(Row, Col) = CStr(vsBill.Cell(flexcpData, Row, Col)): Exit Sub
                    End If
                End If
            End If
            If InStr(mstrPrivsOpt, ";药品输入小数;") = 0 Then
                If Val(vsBill.EditText) <> Int(vsBill.EditText) Then
                    MsgBox "你没有权限输入小数。", vbInformation, gstrSysName
                    vsBill.TextMatrix(Row, Col) = CStr(vsBill.Cell(flexcpData, Row, Col)): Exit Sub
                End If
            End If
            
            strTmp = vsBill.EditText
            If Val(vsBill.EditText) = 0 Then
                If MsgBox("剂量输入为零，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    vsBill.TextMatrix(Row, Col) = CStr(vsBill.Cell(flexcpData, Row, Col)): Exit Sub
                Else
                    vsBill.EditCell: vsBill.EditText = strTmp '焦点丢失后EditText也丢失
                End If
            End If
            
            int中药形态 = Get中药形态
            If int中药形态 = 0 Then
                '需要检查库存
                lng药品ID = Val(Split(mcll规格("_" & lng药名ID) & ",", ",")(0))
                dbl系数 = Get剂量系数(lng药品ID)
                dbl系数 = IIf(dbl系数 = 0, 1, dbl系数)
                If CheckStock(lng药品ID, FormatEx(Val(strTmp) / dbl系数, 5)) = False Then
                        vsBill.TextMatrix(Row, Col) = CStr(vsBill.Cell(flexcpData, Row, Col)): Exit Sub
                Else
                    vsBill.EditCell: vsBill.EditText = strTmp
                End If
            End If
            
            '检查处方限量:保存时检查
            vsBill.EditText = FormatEx(Val(vsBill.EditText), 5)
            strTmp = vsBill.EditText  '弹出Msgbox后vsBill.EditText会被清空,所以需要事先记录
            
            If 分解中药规格(lng药名ID, Val(strTmp)) = False Then
                
            End If
            Call Show中药规格(lng药名ID, Val(strTmp))
            '最大金额检查
            If gcurMaxMoney > 0 Then
                For i = 1 To mobjDetails.Count
                        If mobjDetails(i).Detail.药名ID = lng药名ID Then
                                If mobjDetails(i).InComes(1).应收金额 > gcurMaxMoney Then
                                    If MsgBox("药品为:" & mobjDetails(i).Detail.名称 & " 规格为" & mobjDetails(i).Detail.规格 & vbCrLf & _
                                                      "的当前金额超过了" & gcurMaxMoney & ",你确定要继续吗?", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                                        mobjDetails(i).数次 = Val(vsBill.Cell(flexcpData, Row, Col))
                                        Call 分解中药规格(lng药名ID, Val(vsBill.TextMatrix(Row, Col)))
                                        Call Show中药规格(lng药名ID, Val(vsBill.TextMatrix(Row, Col)))
                                        vsBill.TextMatrix(Row, Col) = CStr(vsBill.Cell(flexcpData, Row, Col))
                                        Exit Sub
                                    End If
                                End If
                        End If
                Next
            End If
            
            
            Call GetBillTotalIncomes(cur应收, cur实收, blnOverFlow)
            If blnOverFlow Then
                '溢出,恢复输入
                MsgBox "输入数量导致单据金额过大，请作适当调整。", vbInformation, gstrSysName
                Call 分解中药规格(lng药名ID, Val(vsBill.TextMatrix(Row, Col)))
                Call Show中药规格(lng药名ID, Val(vsBill.TextMatrix(Row, Col)))
                vsBill.TextMatrix(Row, Col) = CStr(vsBill.Cell(flexcpData, Row, Col))
                Exit Sub
            End If
                     
            '需要在实收求出之后
            '记帐费用报警加入行前报警,不输病人先输费用再输病人这种情况在最后保存时判断
            mrsWarn.Filter = ""
            If mrsWarn.RecordCount > 0 And Not mrsPati Is Nothing Then
                Call GetBillTotalIncomes(, cur实收)
                If cur实收 > 0 Then
                    gbytWarn = BillingWarn(mstrPrivsOpt, mrsPati!姓名 & IIf(NVL(mrsPati!住院号) = "", "", "(住院号:" & mrsPati!住院号 & " 床号:" & mrsPati!床号 & ")"), Val("" & mrsPati!病区ID), mrsPati!适用病人, mrsWarn, mrsPati!病人余额, _
                                Val("" & mrsPati!当日额) - mcurModiMoney, cur实收 + mcur非中药金额, Val("" & mrsPati!担保额), 7, "中草药", mstrWarn, , gblnPrice)
                                        
                    If gbytWarn = 2 Or gbytWarn = 3 Then
                        Call 分解中药规格(lng药名ID, Val(vsBill.TextMatrix(Row, Col)))
                        Call Show中药规格(lng药名ID, Val(vsBill.TextMatrix(Row, Col)))
                        vsBill.TextMatrix(Row, Col) = CStr(vsBill.Cell(flexcpData, Row, Col))
                        vsBill.TextMatrix(Row, Col) = CStr(vsBill.Cell(flexcpData, Row, Col))
                        sta.Panels(2).Text = "预交:" & Format(mrsPati!预交余额, "0.00") & "/费用:" & Format(mrsPati!预交余额 - mrsPati!病人余额, "0.00") & "/剩余:" & Format(mrsPati!病人余额, "0.00")
                        Exit Sub
                    End If
                End If
            End If
            '确认输入
            vsBill.TextMatrix(Row, Col) = strTmp
            vsBill.Cell(flexcpData, Row, Col) = vsBill.TextMatrix(Row, Col)
            '刷新金额显示
            txt应收.Text = Format(cur应收, gstrDec)
            txt实收.Text = Format(cur实收, gstrDec)
            mblnChange = True
        End If
        Call EnterNextCell(Row, Col)
    ElseIf Col Mod MCOLS = col剂量 Then
        lng药名ID = Val(vsBill.Cell(flexcpData, Row, GetBillCol(2, Col)))
        '药名未输入,不允输入剂量
        If lng药名ID = 0 Then
            KeyAscii = 0: Exit Sub
        End If
        strTmp = "0123456789" & gstrABC
        If InStr(mstrPrivsOpt, ";草药负数记帐;") > 0 Then
            If mint险类 > 0 Then
                If gclsInsure.GetCapability(support负数记帐, mlng病人ID, mint险类) Then strTmp = strTmp & "-"
            Else
                strTmp = strTmp & "-"
            End If
        End If
        '除了散装以外,其他的只能输入正数
        If InStr(1, strTmp, "-") > 0 Then
            If Get中药形态 <> 0 Then
                 strTmp = Replace(strTmp, "-", "")
            Else
                '散装的,要检查是否份批的,分批的也不允许负数记帐
                If is分批药品(lng药名ID) Then
                    strTmp = Replace(strTmp, "-", "")
                End If
            End If
        End If
        If InStr(mstrPrivsOpt, ";药品输入小数;") > 0 Then
            strTmp = strTmp & "."
        End If
        If InStr(strTmp & Chr(8) & Chr(27), UCase(Chr(KeyAscii))) = 0 Then
            KeyAscii = 0: Exit Sub
        End If
    End If
End Sub
Public Function CheckDrugDataValied(ByVal lng药品ID As Long, Optional strName As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：检查药品数据的合法性
    '返回：数据合法返回true,否则返回False
    '编制：刘兴洪
    '日期：2010-08-04 12:01:51
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    '毒麻贵重药品权限检查
    Set rsTemp = Read药品信息(lng药品ID)
    If Not rsTemp Is Nothing Then
        If IIf(IsNull(rsTemp!毒理分类), "", rsTemp!毒理分类) = "麻醉药" _
            And InStr(mstrPrivsOpt, ";麻醉药品记帐;") = 0 Then
            MsgBox """" & strName & """为麻醉药品，你没有权限对该类药品记帐！", vbInformation, gstrSysName
            Exit Function
        ElseIf IIf(IsNull(rsTemp!毒理分类), "", rsTemp!毒理分类) = "毒性药" _
            And InStr(mstrPrivsOpt, ";毒性药品记帐;") = 0 Then
            MsgBox """" & strName & """为毒性药品，你没有权限对该类药品记帐！", vbInformation, gstrSysName
            Exit Function
        ElseIf (IIf(IsNull(rsTemp!价值分类), "", rsTemp!价值分类) = "贵重" _
            Or IIf(IsNull(rsTemp!价值分类), "", rsTemp!价值分类) = "昂贵") _
            And InStr(mstrPrivsOpt, ";贵重药品记帐;") = 0 Then
            MsgBox """" & strName & """为贵重或昂贵药品，你没有权限对该类药品记帐！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    CheckDrugDataValied = True
    
End Function
Private Function zlGetDetail(ByVal lng药品ID As Long, Optional dbl数量 As Double, Optional ByRef objDetail As Detail) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：设置明细数据
    '返回：
    '编制：刘兴洪
    '日期：2010-08-04 16:28:00
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim lngTmp As Long, lng药房ID As Long
    
    Set objDetail = New Detail
    If cbo药房.ListIndex < 0 Then
        MsgBox "请选择中药房。", vbInformation, gstrSysName
        Exit Function
    End If
    
    On Error GoTo errH:
    

    If mint险类 > 0 Then
        strSQL = _
        " Select A.ID,A.类别,B.名称 as 类别名称,A.编码,A.名称," & _
        "       A.规格,A.计算单位,A.屏蔽费别,A.是否变价,A.加班加价," & _
        "       A.执行科室,A.费用类型,A.补充摘要,M.要求审批,C.药房分批,C.药名ID," & _
        "       C.住院单位,C.住院包装,J1.名称 as 诊疗名称,J1.计算单位 as 剂量单位,C.剂量系数,A.服务对象" & _
        " From 收费项目目录 A,收费项目类别 B,药品规格 C,保险支付项目 M,保险支付大类 N,诊疗项目目录 J1" & _
        " Where A.类别=B.编码 And A.ID=C.药品ID And A.ID=[1] and C.药名ID=J1.ID" & _
        " And A.ID=M.收费细目ID(+) And M.险类(+)=[2] And M.大类ID=N.ID(+)"
    Else
        strSQL = _
            "   Select A.ID,A.类别,B.名称 as 类别名称,A.编码,A.名称," & _
            "           A.规格,A.计算单位,A.屏蔽费别,A.是否变价,A.加班加价," & _
            "           A.执行科室,A.费用类型,A.补充摘要,0 as 要求审批,C.药房分批,C.药名ID," & _
            "           C.住院单位,C.住院包装,J1.名称 as 诊疗名称,J1.计算单位 as 诊疗单位,J1.计算单位 as 剂量单位,C.剂量系数,A.服务对象" & _
            " From 收费项目目录 A,收费项目类别 B,药品规格 C,诊疗项目目录 J1" & _
            " Where A.类别=B.编码 And A.ID=C.药品ID and C.药名ID=J1.ID And A.ID=[1]"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng药品ID, mint险类)
    '检查收费与发药分离时不允许输入时价及分批药品
    If gbln分离发药 Then
        If NVL(rsTmp!是否变价, 0) = 1 Or NVL(rsTmp!药房分批, 0) = 1 Then
            MsgBox "发药分离处理时不能输入时价或分批药品。", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    '检查对应保险支付项目,门诊费用多了这个条件:mbytFun=0
    If mint险类 <> 0 Then
        If Not CheckMediCareItem(lng药品ID, mint险类, "" & rsTmp!名称, Val(NVL(rsTmp!是否变价)) <> 1) Then
            Exit Function
        End If
    End If
    '毒麻贵重药品权限检查
    If CheckDrugDataValied(lng药品ID, NVL(rsTmp!名称)) = False Then Exit Function
    '相关库存检查
    '---------------------------------------------------------------------------------------
    objDetail.ID = rsTmp!ID
    objDetail.药名ID = rsTmp!药名ID
    objDetail.编码 = rsTmp!编码
    objDetail.名称 = rsTmp!名称
    objDetail.计算单位 = NVL(rsTmp!计算单位)
    objDetail.规格 = NVL(rsTmp!规格)
    objDetail.类别 = rsTmp!类别
    objDetail.类别名称 = rsTmp!类别名称
    objDetail.变价 = NVL(rsTmp!是否变价, 0) <> 0
    objDetail.分批 = NVL(rsTmp!药房分批, 0) <> 0
    objDetail.补充摘要 = NVL(rsTmp!补充摘要, 0) <> 0
    objDetail.处方职务 = Get处方职务(rsTmp!ID)
    objDetail.处方限量 = Get处方限量(rsTmp!ID)
    objDetail.加班加价 = NVL(rsTmp!加班加价, 0) <> 0
    objDetail.屏蔽费别 = NVL(rsTmp!屏蔽费别, 0) <> 0
    objDetail.住院包装 = NVL(rsTmp!住院包装, 1)
    objDetail.住院单位 = NVL(rsTmp!住院单位)
    objDetail.执行科室 = NVL(rsTmp!执行科室, 0)
    objDetail.类型 = NVL(rsTmp!费用类型)
    objDetail.要求审批 = NVL(rsTmp!要求审批, 0) = 1
    objDetail.中药形态 = Get中药形态
    objDetail.诊疗名称 = NVL(rsTmp!诊疗名称)
    objDetail.剂量单位 = NVL(rsTmp!剂量单位)
    objDetail.剂量系数 = NVL(rsTmp!剂量系数)
    objDetail.服务对象 = Val(NVL(rsTmp!服务对象))
    lng药房ID = cbo药房.ItemData(cbo药房.ListIndex)
    If Not gbln分离发药 Then
        objDetail.库存 = GetStock(rsTmp!ID, lng药房ID)
        If gbln住院单位 Then
            objDetail.库存 = objDetail.库存 / objDetail.住院包装
        End If
    ElseIf gstr中药房 <> "" Then
        objDetail.库存 = GetMultiStock(rsTmp!ID, gstr中药房)
        If objDetail.库存 = 0 And gblnStock Then
            MsgBox "[" & objDetail.名称 & "]的可用库存为零!", vbInformation, gstrSysName
            Exit Function
        End If
        If gbln住院单位 Then
            objDetail.库存 = objDetail.库存 / objDetail.住院包装
        End If
    End If
    zlGetDetail = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function SetBillDetail(ByVal lng药品ID As Long, Optional dbl数量 As Double, Optional lng序号 As Long = 0, _
    Optional objDetail As Detail = Nothing, Optional objBillDetail As BillDetail) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：新输入一味中草药时，计算价格，设置对象，以及显示输入
    '         lng序号-当前分配的序号
    '出参 : 返回明细行数据
    '编制：刘兴洪
    '日期：2010-08-02 17:55:57
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim objInComes As New BillInComes
    Dim rsTmp As ADODB.Recordset, strSQL As String, lngRow  As Integer
    Dim str摘要 As String, lng药房ID As Long
    Dim rs药品信息 As ADODB.Recordset
    
    lngRow = GetRow(vsBill.Row, vsBill.Col)
    If objDetail Is Nothing Then
        If zlGetDetail(lng药品ID, dbl数量, objDetail) = False Then Exit Function
    End If
    
    lng药房ID = cbo药房.ItemData(cbo药房.ListIndex)
    str摘要 = Get摘要(lng药品ID)
    
    If objDetail.补充摘要 Then
        If frmInputBox.InputBox(Me, "摘要", "请输入""" & objDetail.名称 & """的摘要信息:", 200, 3, True, False, str摘要) Then
            str摘要 = str摘要
        End If
    Else
'        If mint险类 <> 0 Then '90304
            str摘要 = gclsInsure.GetItemInfo(mint险类, mlng病人ID, lng药品ID, str摘要, 2)
'        End If
    End If
    Call Set摘要(lng药品ID, str摘要)
    Dim dblTemp As Double
    dblTemp = FormatEx(dbl数量 / objDetail.剂量系数, 5)
    If gbln住院单位 Then     '52722
        dblTemp = dblTemp / IIf(objDetail.住院包装 = 0, 1, objDetail.住院包装)
        Set objBillDetail = mobjDetails.Add(objDetail, lng药品ID, lngRow, 0, mlng病人ID, 0, 0, 0, "", "", "", 0, 0, _
                mstr费别, 0, objDetail.类别, objDetail.住院单位, "", Val(txt付数.Text), dblTemp, 0, lng药房ID, objInComes, "", lngRow & "_" & lng药品ID, , , , , , str摘要)
    Else
        Set objBillDetail = mobjDetails.Add(objDetail, lng药品ID, lngRow, 0, mlng病人ID, 0, 0, 0, "", "", "", 0, 0, _
                mstr费别, 0, objDetail.类别, objDetail.计算单位, "", Val(txt付数.Text), dblTemp, 0, lng药房ID, objInComes, "", lngRow & "_" & lng药品ID, , , , , , str摘要)
    End If


    
    mblnChange = True
    SetBillDetail = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function Get摘要(ByVal lng药品ID As Long) As String
    '------------------------------------------------------------------------------------------------------------------------
    '功能：获取曾经录入过的摘要
    '编制：刘兴洪
    '日期：2010-08-03 12:02:37
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim str摘要 As String
    Err = 0: On Error Resume Next
    str摘要 = mcllInput规则摘要("K" & lng药品ID)
    Get摘要 = str摘要
    Err = 0: On Error GoTo 0
End Function
Private Sub Set摘要(ByVal lng药品ID As Long, str摘要 As String)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：获取曾经录入过的摘要
    '编制：刘兴洪
    '日期：2010-08-03 12:02:37
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    mcllInput规则摘要.Remove "K" & lng药品ID
    mcllInput规则摘要.Add str摘要, "K" & lng药品ID
    Err = 0: On Error GoTo 0
End Sub
Private Sub 重新刷新所有中药规格()
    '------------------------------------------------------------------------------------------------------------------------
    '功能：重设所有中药的规格(按数量分配),并重新显示当前中药的数量按规格分配的列表
    '编制：刘兴洪
    '日期：2010-08-03 14:44:47
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, lng药名ID, dbl数量 As Double, lng药品ID As Long
    Dim int中药形态 As Long, rsTmp As ADODB.Recordset
    
    int中药形态 = Get中药形态
    With vsBill
        For i = .FixedRows To .Rows - 1
            For j = 0 To .Cols - 1 Step MCOLS
                lng药名ID = Val(.Cell(flexcpData, i, j + 2))
                dbl数量 = Val(.TextMatrix(i, j + 1))
                If lng药名ID <> 0 Then
                    If int中药形态 = 0 Then
                        '记录上次选择的药品ID,以便恢复选择
                        '问题:45410
                        lng药品ID = Val(Split(mcll规格("_" & lng药名ID) & ",", ",")(0))
                        mcll规格.Remove ("_" & lng药名ID)  '重选规格ID
                        Set rsTmp = Get散装规格(lng药名ID)   '取缺省规格
                        If rsTmp.RecordCount > 0 Then
                            If lng药品ID <> 0 Then rsTmp.Find "药品ID=" & lng药品ID, , adSearchForward, 1
                            If rsTmp.EOF Then rsTmp.MoveFirst
                            mcll规格.Add rsTmp!药品ID & ",0", "_" & lng药名ID
                        Else
                            mcll规格.Add IIf(lng药品ID = 0, "", lng药品ID & ",0"), "_" & lng药名ID
                        End If
                    End If
                    
                    Call 分解中药规格(lng药名ID, dbl数量, , False)
                    If mcll规格("_" & lng药名ID) = "" Or InStr(mcll规格("_" & lng药名ID), "|") > 0 Then
                        .Cell(flexcpForeColor, i, j + 1) = vbRed
                    Else
                        .Cell(flexcpForeColor, i, j + 1) = .ForeColor
                    End If
                End If
            Next
        Next
        lng药名ID = Val(.Cell(flexcpData, .Row, (.Col \ MCOLS) * MCOLS + 2))
        dbl数量 = Val(.TextMatrix(.Row, (.Col \ MCOLS) * MCOLS + 1))
    End With
    If lng药名ID <> 0 Then Call Show中药规格(lng药名ID, dbl数量)
    Call ReCalc应收合计
End Sub

Private Function Get散装规格(ByVal lng药名ID As Long) As ADODB.Recordset
'功能：获取当前药品的所有可用的散装规格
    Dim lng药房ID As Long, strSQL As String

    On Error GoTo errHandle
    
    If gblnStock Then
        If cbo药房.ListIndex <> -1 Then lng药房ID = cbo药房.ItemData(cbo药房.ListIndex)
        If lng药房ID <> 0 Then
            strSQL = " And Exists(Select 1 From 药品库存 B" & _
                " Where (Nvl(b.批次, 0) = 0 Or b.效期 Is Null Or b.效期>Trunc(Sysdate))" & _
                " And b.性质=1 And b.库房ID=[4] And a.药品ID=b.药品ID Group by b.药品ID" & _
                " Having Sum(b.可用数量)>0)"
        End If
    End If
    strSQL = "Select a.药名id, a.药品id, d.规格, d.产地, a.剂量系数,A.住院单位,A.住院包装,D.计算单位," & _
            "       d.编码, d.名称,A.中药形态,D.执行科室" & vbNewLine & _
            "From 药品规格 A, 收费项目目录 D" & vbNewLine & _
            "Where a.药名id = [1] And a.中药形态 = 0 And a.药品ID = d.ID" & strSQL & vbNewLine & _
            " And (D.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or D.撤档时间 IS NULL) And D.服务对象 IN([3],3)" & _
            " And (D.站点='" & gstrNodeNo & "' Or D.站点 is Null) Order By D.编码"
    Set Get散装规格 = zlDatabase.OpenSQLRecord(strSQL, "规格列表", lng药名ID, lng药房ID, mint病人来源, lng药房ID)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function 分解中药规格(ByVal lng药名ID As Long, ByVal dbl数量 As Double, _
    Optional objDetail As Detail, Optional ReCal应收 As Boolean = True, Optional int中药形态 As Integer = -1) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：分解中药规格数据
    '入参: dbl数量-计量单位数量
    '返回 :如果分析成功,返回true, 否则返回失败!
    '编制：刘兴洪
    '日期：2010-08-02 17:50:45
    '说明：31867
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, i As Long
    Dim str规格数量 As String, lng药房ID As Long
    Dim varData As Variant, varTemp As Variant
    Dim blnObj As Boolean
    
    If int中药形态 = -1 Then int中药形态 = Get中药形态
    blnObj = Not objDetail Is Nothing
    If int中药形态 = 0 Then
        '散装的在输入时已确定规格
        str规格数量 = mcll规格("_" & lng药名ID)
        If str规格数量 <> "" Then str规格数量 = Split(str规格数量, ",")(0) & "," & dbl数量
    Else
        '2.分配结果,药品id,数量;药品id,数量;...|剩余数量
        On Error GoTo errH
        If cbo药房.ListIndex <> -1 Then lng药房ID = cbo药房.ItemData(cbo药房.ListIndex)
        strSQL = "Select Zl_Dispensechspecs([1],[2],[3],[4],[5],[6]) as txt From dual"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "规格分配", lng药名ID, int中药形态, dbl数量, Val(txt付数.Text), lng药房ID, IIf(gbln分离发药, 1, 0))
        str规格数量 = "" & rsTemp!txt
    End If
    
    Call mcll规格.Remove("_" & lng药名ID)
    mcll规格.Add str规格数量, "_" & lng药名ID
    
    If str规格数量 = "" Then
        '删除所有药名ID信息
        Call DeleteDetails(lng药名ID)
        Exit Function
    End If
    
    '--返回:药品id,数量;药品id,数量;...(散装只选择一个规格)
    '--                             不能完全分配时返回:剂量为6和10的情况下,17克的分配=23755,6;23756,10|1
    '--                             不能分配时返回空,例如:剂量为6和10的情况下,3克的分配
    '删除所有药名ID信息
    Call DeleteDetails(lng药名ID)
    Dim objBillDetail As BillDetail
    '添加明细数据
    varData = Split(Split(str规格数量, "|")(0), ";")
    For i = 0 To UBound(varData)
         varTemp = Split(varData(i), ",")
         '设置明细
         If Not blnObj Then Set objDetail = Nothing
         
         If SetBillDetail(Val(varTemp(0)), Val(varTemp(1)), i, objDetail, objBillDetail) = False Then
            '分解失败
            Call DeleteDetails(lng药名ID)
            If ReCal应收 Then Call ReCalc应收合计
            Exit Function
         Else
            '设置收费项目数据
            Call zlCalcMoney(objBillDetail, True)
         End If
    Next
    If ReCal应收 Then Call ReCalc应收合计
    分解中药规格 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub ReCalc应收合计()
    '------------------------------------------------------------------------------------------------------------------------
    '功能：重算应收实收合计合计
    '编制：刘兴洪
    '日期：2010-08-04 11:40:25
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    '计算合计数
    Dim cur应收 As Currency, cur实收 As Currency
    Call GetBillTotalIncomes(cur应收, cur实收)
    txt应收.Text = Format(cur应收, gstrDec)
    txt实收.Text = Format(cur实收, gstrDec)
End Sub
Private Sub DeleteDetails(ByVal lng药名ID As Long)
    '删除药名ID的所有规格
    Dim blnNotFond As Boolean, i As Long
     Do While True
        '删除药名ID
        blnNotFond = True
        For i = 1 To mobjDetails.Count
            If mobjDetails(i).Detail.药名ID = lng药名ID Then
                 blnNotFond = False
                  mobjDetails.Remove i: Exit For
            End If
        Next
        If blnNotFond = True Then Exit Do
    Loop
End Sub
 
Private Sub GetBillTotalIncomes(Optional cur应收 As Currency, Optional cur实收 As Currency, Optional blnOvweFlow As Boolean)
'参数：blnOvweFlow=数据是否溢出
    Dim i As Long, j As Long
    
    cur应收 = 0: cur实收 = 0: blnOvweFlow = False
    For i = 1 To mobjDetails.Count
        For j = 1 To mobjDetails(i).InComes.Count
            '要用VAL转为Double进行运算
            If Abs(Val(cur应收) + Val(mobjDetails(i).InComes(j).应收金额)) > 922337203685477# Then
                blnOvweFlow = True: Exit Sub
            End If
            If Abs(Val(cur实收) + Val(mobjDetails(i).InComes(j).实收金额)) > 922337203685477# Then
                blnOvweFlow = True: Exit Sub
            End If
            cur应收 = cur应收 + mobjDetails(i).InComes(j).应收金额
            cur实收 = cur实收 + mobjDetails(i).InComes(j).实收金额
        Next
    Next
End Sub

 
Private Function Get中药形态() As Integer
    '------------------------------------------------------------------------------------------------------------------------
    '功能：获取当前的中药形态
    '返回：0-散装;1-饮片;2-免煎剂
    '编制：刘兴洪
    '日期：2010-07-27 14:58:51
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    For i = 0 To opt形态.UBound
        If opt形态(i).Value = True Then Exit For
    Next
    Get中药形态 = i
End Function
Private Sub zlCalcMoney(objBillDetail As BillDetail, Optional bln不调价 As Boolean)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：重新计算指定药品行的价格和金额
    '入参：objbillDetail-指定的明细数据
    '          bln不调价-不进行调价处理
    '出参：
    '返回：
    '编制：刘兴洪
    '日期：2010-08-03 14:18:34
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, objInCome As New BillInCome
    Dim dblAllTime As Double, dblMoney As Double, dblPrice As Double, dbl加班加价率 As Double
    Dim str费别 As String, cur金额 As Currency
    Dim strInfo As String, strSQL As String, i As Long, dblPriceSingle As Double

    On Error GoTo errH
     If Not bln不调价 Then Call AdjustCpt(objBillDetail.收费细目ID)

    strSQL = _
        " Select B.收入项目ID,C.名称,C.收据费目,B.现价,B.原价,B.加班加价率,B.附术收费率 " & _
        " From 收费项目目录 A,收费价目 B,收入项目 C " & _
        " Where B.收费细目ID = A.ID And C.ID = B.收入项目ID " & _
        " And ((Sysdate Between B.执行日期 and B.终止日期) Or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
        " And A.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, objBillDetail.收费细目ID)
    
    If Not rsTmp.EOF Then
        '先获取操作员以前输入的变价金额
        If objBillDetail.Detail.变价 Then
            If InStr(",5,6,7,", objBillDetail.收费类别) > 0 Then
                '计算药品时价(分批或不分批)
                '必然有记录(输入该项目时已判断)
                dblAllTime = objBillDetail.付数 * objBillDetail.数次
                If gbln住院单位 Then
                    '库存时价按售价数量进行计算
                    dblAllTime = dblAllTime * objBillDetail.Detail.住院包装
                End If
                If dblAllTime <> 0 Then
                    dblPrice = Get时价药品应收金额(objBillDetail.执行部门ID, objBillDetail.收费细目ID, dblAllTime, gstrDec, dblPriceSingle)
                    If dblAllTime <> 0 Then
                        '数量未分解完毕
                        MsgBox "第 " & Split(objBillDetail.Key, "_")(0) & " 行时价药品""" & objBillDetail.Detail.名称 & """库存不足,无法计算价格！", vbInformation, gstrSysName
                        dblMoney = 0
                    Else
                        '注意：货币型最多只能保留4位小数,且按Round处理,所以需要手工舍入;而用其它型在计算精度上又有问题
                        dblAllTime = objBillDetail.付数 * objBillDetail.数次
                        If gbln住院单位 Then
                            '按售价数量计算实价
                            dblAllTime = dblAllTime * objBillDetail.Detail.住院包装
                        End If
                        dblMoney = IIf(dblPriceSingle = 0, Format(dblPrice / dblAllTime, gstrFeePrecisionFmt), dblPriceSingle) '这里结果是按售价单位
                    End If
                Else
                    dblMoney = 0
                End If
            Else
                If objBillDetail.InComes.Count = 0 Then
                    '如果第一次计算金额,变价默认取原价
                    dblMoney = 0    'dblMoney = Nvl(rsTmp!原价, 0)
                Else
                    dblMoney = objBillDetail.InComes(1).标准单价
                    '如果用户输入的变价不满足变价范围，则取默认值
                    If Abs(dblMoney) > Abs(NVL(rsTmp!现价, 0)) Then
                        dblMoney = NVL(rsTmp!原价, 0)
                    End If
                End If
            End If
        End If

        '再清除原有记录
        Set objBillDetail.InComes = New BillInComes

        '填写现有费用记录
        For i = 1 To rsTmp.RecordCount
            Set objInCome = New BillInCome
            With objInCome
                .收入项目ID = rsTmp!收入项目ID
                .收入项目 = rsTmp!名称
                .收据费目 = NVL(rsTmp!收据费目)
                .原价 = NVL(rsTmp!原价, 0)
                .现价 = NVL(rsTmp!现价, 0)
                If objBillDetail.Detail.变价 Then
                    If InStr(",5,6,7,", objBillDetail.收费类别) > 0 And gbln住院单位 Then
                        .标准单价 = Format(dblMoney * objBillDetail.Detail.住院包装, gstrFeePrecisionFmt)
                    Else
                        .标准单价 = Format(dblMoney, gstrFeePrecisionFmt)
                    End If
                Else
                    If InStr(",5,6,7,", objBillDetail.收费类别) > 0 And gbln住院单位 Then
                        .标准单价 = Format(NVL(rsTmp!现价, 0) * objBillDetail.Detail.住院包装, gstrFeePrecisionFmt)
                    Else
                        .标准单价 = Format(NVL(rsTmp!现价, 0), gstrFeePrecisionFmt)
                    End If
                End If

                '应收金额=单价 * 付数 * 数次
                If InStr(",5,6,7,", objBillDetail.收费类别) > 0 _
                    And objBillDetail.Detail.变价 Then
                    .应收金额 = dblPrice '保证应收金额与零售金额没有误差
                Else
                    .应收金额 = .标准单价 * objBillDetail.付数 * objBillDetail.数次
                End If

                '加班费用率计算
                dbl加班加价率 = 0
                If mbln加班 And objBillDetail.Detail.加班加价 Then
                    dbl加班加价率 = NVL(rsTmp!加班加价率, 0) / 100
                    .应收金额 = .应收金额 + .应收金额 * dbl加班加价率
                End If

                .应收金额 = CCur(Format(.应收金额, gstrDec))

                dblAllTime = objBillDetail.付数 * objBillDetail.数次
                If gbln住院单位 Then dblAllTime = dblAllTime * objBillDetail.Detail.住院包装
                
                If objBillDetail.Detail.屏蔽费别 Then
                    .实收金额 = .应收金额
                Else
                    '药品按成本价加收,传入数量
                    
                    .实收金额 = CCur(Format(ActualMoney(mstr费别, .收入项目ID, .应收金额, _
                        objBillDetail.收费细目ID, objBillDetail.执行部门ID, dblAllTime, dbl加班加价率), gstrDec))
                End If
                objBillDetail.费别 = mstr费别

                '获取项目保险信息,门诊只有医保病人才算,门诊费用多了这个条件:And mbytFun = 0
                If mint险类 <> 0 Then
                    strInfo = gclsInsure.GetItemInsure(mlng病人ID, objBillDetail.收费细目ID, .实收金额, True, mint险类, _
                        objBillDetail.摘要 & "||" & dblAllTime)
                    If strInfo <> "" Then
                        objBillDetail.保险项目否 = Val(Split(strInfo, ";")(0)) <> 0
                        objBillDetail.保险大类ID = Val(Split(strInfo, ";")(1))
                        .统筹金额 = Format(Val(Split(strInfo, ";")(2)), gstrDec)
                        objBillDetail.保险编码 = CStr(Split(strInfo, ";")(3))
                                                
                        If UBound(Split(strInfo, ";")) >= 4 Then
                            If CStr(Split(strInfo, ";")(4)) <> "" Then objBillDetail.摘要 = CStr(Split(strInfo, ";")(4))
                            If UBound(Split(strInfo, ";")) >= 5 Then
                                If Split(strInfo, ";")(5) <> "" Then objBillDetail.Detail.类型 = Split(strInfo, ";")(5)
                            End If
                        End If
                    End If
                End If

                objBillDetail.InComes.Add .收入项目ID, .收入项目, .收据费目, .标准单价, .应收金额, .实收金额, .原价, .现价, "_" & .实收金额, .统筹金额
            End With
            rsTmp.MoveNext
        Next
    Else
        '如果没有收入项目,则清除对应的程序对象
        Set objBillDetail.InComes = New BillInComes
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
 
Private Sub Show中药规格(ByVal lng药名ID As Long, dbl数量 As Double, Optional int中药形态 As Long = -1)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：根据当前行和列，显示或隐藏中药规格列表
    '编制：刘兴洪
    '日期：2010-08-03 14:55:39
    '说明：如果是散装形态，则加载可选择的规格下拉列表
    '------------------------------------------------------------------------------------------------------------------------
    Dim str规格数量 As String, varData As Variant, arrValue As Variant
    Dim i As Long, str药品IDs As String, lngColBegin As Long
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strMsg As String, lng药品ID As Long
    Dim bln非散装换散装 As Boolean
    
    lngColBegin = (vsBill.Col \ MCOLS) * MCOLS
    vsBill.Cell(flexcpForeColor, vsBill.Row, lngColBegin + 1) = vsBill.ForeColor
    
    cmd形态.Visible = False
        
    With vs中药规格
        .Rows = .FixedRows
        .ColComboList(.ColIndex("规格")) = ""
        If dbl数量 = 0 Then Exit Sub
        If int中药形态 = -1 Then int中药形态 = Get中药形态
        
        str规格数量 = Trim(mcll规格("_" & lng药名ID))
        .Redraw = flexRDNone
        
        If str规格数量 = "" Then
            .Rows = .FixedRows + 1
            '不能分配时返回空,例如:剂量为6和10的情况下,3克的分配
            .MergeCells = flexMergeRestrictRows
            If int中药形态 = 0 Then
                strMsg = "该药品没有可用的散装形态，请选择其它药品或形态。"
            Else
                strMsg = "无法将所有数量按可用规格分配，请调整用量。"
            End If
            .TextMatrix(.Rows - 1, .ColIndex("规格")) = strMsg
            .TextMatrix(.Rows - 1, .ColIndex("产地")) = strMsg
            .MergeRow(.Rows - 1) = True
            .TextMatrix(.Rows - 1, .ColIndex("数量")) = dbl数量
            .Cell(flexcpData, .Rows - 1, .ColIndex("数量")) = .TextMatrix(.Rows - 1, .ColIndex("数量"))
            .Cell(flexcpForeColor, .Rows - 1, .ColIndex("数量")) = vbRed
            vsBill.Cell(flexcpForeColor, vsBill.Row, lngColBegin + 1) = vbRed
        Else
            varData = Split(Split(str规格数量, "|")(0), ";")
            If InStr(str规格数量, "|") > 0 Then
                 .Rows = .FixedRows + UBound(varData) + 2
            Else
                 .Rows = .FixedRows + UBound(varData) + 1
            End If
            
            For i = 0 To UBound(varData)
                arrValue = Split(varData(i), ",")
                str药品IDs = str药品IDs & "," & Val(arrValue(0))
                .Cell(flexcpData, .FixedRows + i, .ColIndex("规格")) = Val(arrValue(0)) '规格ID
                .TextMatrix(.FixedRows + i, .ColIndex("数量")) = arrValue(1)    '数量
                .Cell(flexcpData, .FixedRows + i, .ColIndex("数量")) = .TextMatrix(.FixedRows + i, .ColIndex("数量"))
            Next
            str药品IDs = Mid(str药品IDs, 2)
            
            On Error GoTo errH:
            If int中药形态 = 0 Then
                '读出所有可用(有库存)的散装规格，以便可以选择其它的规格
                Set rsTmp = Get散装规格(lng药名ID)
            Else
                strSQL = "" & _
                    "   Select /*+ Rule*/A.药品ID,D.规格,D.产地,A.剂量系数,A.中药形态 ," & _
                    "           A.住院单位,A.住院包装,D.计算单位 From 药品规格 A,收费项目目录 D,Table(f_Num2List([1])) B" & vbNewLine & _
                    "   Where A.药品ID = B.Column_value And A.药品ID = D.ID"
                '不加形态条件，因为该规格可能是换成的散装，而当前是选择的饮片
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "规格列表", str药品IDs)
                '如果只有一条记录,且是散装形态的,则认为是非散装的,换成散装的
                If rsTmp.RecordCount <> 0 Then
                    bln非散装换散装 = rsTmp.RecordCount = 1 And Val(NVL(rsTmp!中药形态)) = 0
                End If
            End If
            For i = .FixedRows To .Rows - 1
                If InStr(str规格数量, "|") > 0 And i = .Rows - 1 Then
                '最后一行显示未分配数量
                    .MergeCells = flexMergeRestrictRows
                    strMsg = "无法将所有数量按可用规格分配，请调整用量。"
                    .TextMatrix(i, .ColIndex("规格")) = strMsg
                    .TextMatrix(i, .ColIndex("产地")) = strMsg
                    .MergeRow(i) = True
                    .Cell(flexcpForeColor, i, .ColIndex("规格")) = vbRed
                    .TextMatrix(i, .ColIndex("数量")) = Split(str规格数量, "|")(1)
                    .Cell(flexcpData, i, .ColIndex("数量")) = .TextMatrix(i, .ColIndex("数量"))
                    vsBill.Cell(flexcpForeColor, vsBill.Row, lngColBegin + 1) = vbRed
                Else
                    lng药品ID = Val(CStr(.Cell(flexcpData, i, .ColIndex("规格"))))
                    rsTmp.Filter = "药品ID = " & lng药品ID
                    If rsTmp.RecordCount = 0 Then '散装，库存不足时（允许保存）
                        strMsg = "当前药房库存不足，或者没有散装规格。"
                        .TextMatrix(.Rows - 1, .ColIndex("规格")) = strMsg
                        .TextMatrix(.Rows - 1, .ColIndex("产地")) = strMsg
                        .MergeRow(.Rows - 1) = True
                        .Cell(flexcpForeColor, i, .ColIndex("数量")) = vbRed
                         vsBill.Cell(flexcpForeColor, vsBill.Row, lngColBegin + 1) = vbRed
                    Else
                        .TextMatrix(i, .ColIndex("规格")) = "" & rsTmp!规格
                        .Cell(flexcpData, i, .ColIndex("规格")) = "" & rsTmp!规格 '用于散装规格取消下拉选择时恢复
                        .TextMatrix(i, .ColIndex("产地")) = "" & rsTmp!产地
                        .Cell(flexcpData, i, .ColIndex("剂量系数")) = FormatEx(Val(.Cell(flexcpData, i, .ColIndex("数量"))) / IIf(Val(NVL(rsTmp!剂量系数)) = 0, 1, Val(NVL(rsTmp!剂量系数))), 5) '售价单位数量
                        .TextMatrix(i, .ColIndex("数量")) = FormatEx(Val(.Cell(flexcpData, i, .ColIndex("数量"))) / IIf(Val(NVL(rsTmp!剂量系数)) = 0, 1, Val(NVL(rsTmp!剂量系数))), 5)
                        .Cell(flexcpData, i, .ColIndex("数量")) = Val(.Cell(flexcpData, i, .ColIndex("数量"))) & ":" & IIf(Val(NVL(rsTmp!剂量系数)) = 0, 1, Val(NVL(rsTmp!剂量系数))) & ":" & IIf(Val(NVL(rsTmp!住院包装)) = 0, 1, Val(NVL(rsTmp!住院包装)))
                        '用包装单位显示
                        If gbln住院单位 Then
                             .TextMatrix(i, .ColIndex("数量")) = FormatEx(Val(.TextMatrix(i, .ColIndex("数量"))) / IIf(Val(NVL(rsTmp!住院包装)) = 0, 1, Val(NVL(rsTmp!住院包装))), 5) & NVL(rsTmp!住院单位)
                        Else
                             .TextMatrix(i, .ColIndex("数量")) = Val(.TextMatrix(i, .ColIndex("数量"))) & NVL(rsTmp!计算单位)
                        End If
                        .TextMatrix(i, .ColIndex("剂量系数")) = "" & rsTmp!剂量系数
                        If int中药形态 = 0 Then
                            '中药形态的,数量不足显示红色字体
                            If IsCheckStockEnough(lng药品ID, Val(.Cell(flexcpData, i, .ColIndex("剂量系数")))) Then
                                .Cell(flexcpForeColor, i, .ColIndex("数量")) = .ForeColor
                            Else
                                .Cell(flexcpForeColor, i, .ColIndex("数量")) = vbRed
                                vsBill.Cell(flexcpForeColor, vsBill.Row, lngColBegin + 1) = vbRed
                            End If
                        End If
                    End If
                End If
            Next
            
            '散装形态，允许选择规格
            If int中药形态 = 0 Or bln非散装换散装 Then
                If bln非散装换散装 Then
                    '需要更改规格
                    Set rsTmp = Get散装规格(lng药名ID)
                End If
                
                rsTmp.Filter = ""
                If rsTmp.RecordCount > 1 Then
                    str药品IDs = ""
                    For i = 1 To rsTmp.RecordCount
                        str药品IDs = str药品IDs & "|#" & rsTmp!药品ID & ";" & rsTmp!编码 & "-" & rsTmp!名称 & IIf(Not IsNull(rsTmp!规格), "(" & rsTmp!规格 & ")", "")
                        rsTmp.MoveNext
                    Next
                    .ColComboList(.ColIndex("规格")) = Mid(str药品IDs, 2)
                    rsTmp.MoveFirst
                    .RowData(.FixedRows) = rsTmp   '只有一行
                    .Cell(flexcpBackColor, .FixedRows, .ColIndex("规格")) = &HF0F4E4
                End If
            End If
        End If
        
        If int中药形态 <> 0 Then
            '非散装形态，未分配完时，允许换为散装
            If str规格数量 = "" Or InStr(str规格数量, "|") > 0 Then
                Set rsTmp = Get散装规格(lng药名ID)
                If rsTmp.RecordCount > 0 Then
                    strMsg = "无法将所有数量按可用规格分配，请调整用量或改用散装。"
                    .TextMatrix(.Rows - 1, .ColIndex("规格")) = strMsg
                    .TextMatrix(.Rows - 1, .ColIndex("产地")) = strMsg
                    .MergeRow(.Rows - 1) = True
                    
                    .Select .Rows - 1, .ColIndex("剂量系数")
                    cmd形态.Visible = True
                    cmd形态.Tag = rsTmp!药品ID  '缺省规格
                    cmd形态.Caption = "散装(&D)"
                    cmd形态.Top = vs中药规格.CellTop
                    cmd形态.Left = vs中药规格.CellLeft
                    cmd形态.Width = vs中药规格.CellWidth
                    cmd形态.Height = vs中药规格.CellHeight
                End If
            End If
        End If
        .Redraw = True
        .Visible = True
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub vs中药规格_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim str规格数量 As String, lng药品ID As Long
    Dim objDetail As Detail, dbl数量 As Double
    
    With vs中药规格
        Select Case .Col
        Case .ColIndex("规格")
            If .ComboData = "" Then
                '没有选择时移开焦点
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
            Else
                '散装中药，选择规格之后
                lng药品ID = CLng(.ComboData)
                If zlGetDetail(lng药品ID, Val(.Cell(flexcpData, Row, .ColIndex("剂量系数"))), objDetail) = False Then
                     .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Sub
                End If
                '计算单位数量:放在剂量系数中
                If CheckStock(lng药品ID, Val(.Cell(flexcpData, Row, .ColIndex("剂量系数"))), objDetail) = False Then
                     .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Sub
                End If
                Set rsTmp = .RowData(.FixedRows)
                rsTmp.Filter = "药品ID = " & lng药品ID
                str规格数量 = mcll规格("_" & rsTmp!药名ID)
                dbl数量 = Val(.Cell(flexcpData, Row, .ColIndex("数量")))
                mcll规格.Remove "_" & rsTmp!药名ID
                mcll规格.Add rsTmp!药品ID & "," & dbl数量, "_" & rsTmp!药名ID
                
                If 分解中药规格(Val(NVL(rsTmp!药名ID)), dbl数量, objDetail, True, 0) = False Then
                        mcll规格.Remove "_" & rsTmp!药名ID
                        mcll规格.Add str规格数量, "_" & rsTmp!药名ID
                        .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Sub
                End If
                
                If IsCheckStockEnough(lng药品ID, FormatEx(dbl数量 / IIf(Val(NVL(rsTmp!剂量系数)) = 0, 1, Val(NVL(rsTmp!剂量系数))), 5), objDetail) = False Then
                    vsBill.Cell(flexcpForeColor, vsBill.Row, GetBillCol(1, vsBill.Col)) = vbRed
                    .Cell(flexcpForeColor, Row, .ColIndex("数量")) = vbRed
                Else
                    vsBill.Cell(flexcpForeColor, vsBill.Row, GetBillCol(1, vsBill.Col)) = vsBill.ForeColor
                    .Cell(flexcpForeColor, Row, .ColIndex("数量")) = .ForeColor
                End If
                .TextMatrix(Row, .ColIndex("规格")) = Trim(NVL(rsTmp!规格))
                .Cell(flexcpData, Row, Col) = Trim(NVL(rsTmp!规格))   '用于恢复
                .TextMatrix(Row, .ColIndex("产地")) = Trim(NVL(rsTmp!产地))
                .TextMatrix(Row, .ColIndex("剂量系数")) = Trim(NVL(rsTmp!剂量系数))
                .Cell(flexcpData, Row, .ColIndex("剂量系数")) = FormatEx(dbl数量 / IIf(Val(NVL(rsTmp!剂量系数)) = 0, 1, Val(NVL(rsTmp!剂量系数))), 5)
                
                If vsBill.Enabled And vsBill.Visible Then vsBill.SetFocus
            End If
        Case Else
        End Select
    End With
End Sub

Private Sub vs中药规格_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vs中药规格
        Select Case NewCol
        Case .ColIndex("规格")
            If opt形态(0).Value And .ColComboList(NewCol) <> "" Then
                 .FocusRect = flexFocusSolid
            Else
                 .FocusRect = flexFocusLight
            End If
        End Select
    End With
End Sub

Private Sub vs中药规格_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vs中药规格
        Select Case Col
        Case .ColIndex("规格")
             If Not (opt形态(0).Value Or .ColComboList(Col) <> "") Then
                    Cancel = True
             End If
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub vs中药规格_ChangeEdit()
    Call vs中药规格_AfterEdit(vs中药规格.Row, vs中药规格.Col)
End Sub

Private Sub vs中药规格_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub vs中药规格_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If vs中药规格.ComboIndex <> -1 Then
            Call vs中药规格_KeyPress(13)
        End If
    End If
End Sub
Private Function Get中药规格(ByVal lng药名ID As Long, Optional ByVal lng形态 As Long = -1) As ADODB.Recordset
    '------------------------------------------------------------------------------------------------------------------------
    '功能：根据中药诊疗ID获取中药规格
    '返回：返回规格的记录集
    '编制：刘兴洪
    '日期：2010-08-05 11:19:04
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, lng药房ID As Long
    
    On Error GoTo errH
    If lng形态 = 0 Then
        Set Get中药规格 = Get散装规格(lng药名ID)
    Else
        If gblnStock Then
            lng药房ID = cbo药房.ItemData(cbo药房.ListIndex)
            If lng药房ID <> 0 Then
                strSQL = " And Exists(Select 1 From 药品库存 B" & _
                    " Where (Nvl(b.批次, 0) = 0 Or b.效期 Is Null Or b.效期>Trunc(Sysdate))" & _
                    " And b.性质=1 And b.库房ID=[4] And a.药品ID=b.药品ID Group by b.药品ID" & _
                    " Having Sum(b.可用数量)>0)"
            End If
        End If
    
        strSQL = "" & _
        "   Select A.药品ID,A.中药形态,D.编码,D.执行科室 " & _
        "   From 药品规格 A,收费项目目录 D Where A.药名ID = [1] And A.药品ID = D.ID" & _
                    IIf(lng形态 = -1, "", " And A.中药形态 = [3]") & strSQL & _
         "          And (D.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or D.撤档时间 IS NULL) And D.服务对象 IN([2],3)" & _
         "          And (D.站点='" & gstrNodeNo & "' Or D.站点 is Null) " & _
         "  Order by D.编码"
         
        Set Get中药规格 = zlDatabase.OpenSQLRecord(strSQL, "读取中药规格", lng药名ID, mint病人来源, lng形态, lng药房ID)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ShowSpecs(ByVal lng药名ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示中药的相关规格数
    '编制:刘兴洪
    '日期:2011-01-04 14:13:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, lng药房ID As Long, rsTemp As ADODB.Recordset
    Dim intCols As Integer, lngRow As Long, intCol As Integer
    Dim lngWidth As Long, i As Long
    
    On Error GoTo errH
    lng药房ID = 0
    If lng药名ID = 0 Then
        vsSpecShow.Clear
        vsSpecShow.Rows = 0
    End If
    If cbo药房.ListIndex >= 0 Then lng药房ID = cbo药房.ItemData(cbo药房.ListIndex)
     
    strSQL = "" & _
    "   Select  D.编码,D.规格,D.名称,E.名称 as 药房," & _
    "      " & IIf(gbln住院单位, "A.住院单位", "D.计算单位") & " as 住院单位 ," & _
                IIf(gbln住院单位, "nvl(A.住院包装,1)", "1") & "   as 住院包装,D.计算单位," & _
    "      Sum(nvl(M.可用数量,0))/" & IIf(gbln住院单位, "nvl(A.住院包装,1)", "1") & " as 可用数量" & _
    "   From 药品库存 M,药品规格 A,收费项目目录 D,部门表 E" & _
    "   Where M.药品ID = D.ID and M.药品ID=A.药品ID and M.库房ID=E.ID " & _
    "            And  (Nvl(M.批次, 0) = 0 Or M.效期 Is Null Or M.效期>Trunc(Sysdate)) " & _
    "            And A.药名ID = [1]   " & IIf(lng药房ID = 0, "", " And M.库房ID=[3] ") & _
     "           And (D.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or D.撤档时间 IS NULL) And D.服务对象 IN([2],3)" & _
     "           And (D.站点='" & gstrNodeNo & "' Or D.站点 is Null) " & _
     "  Group by E.名称,D.编码,D.规格,D.名称 ,D.计算单位" & IIf(gbln住院单位, ",A.住院单位", "") & "" & IIf(gbln住院单位, ",nvl(A.住院包装,1)", "") & _
     "  Having Sum(nvl(M.可用数量,0))>0 " & _
     "  Order by D.编码"
     
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "读取中药规格", lng药名ID, mint病人来源, lng药房ID)
    intCols = 3
    With vsSpecShow
        .Clear
        .Rows = 0: .Cols = intCols
        intCol = intCols: lngRow = -1
        lngWidth = (.Width / intCol) - 30
        For i = 0 To .Cols - 1
            .ColWidth(i) = lngWidth
        Next
        Do While Not rsTemp.EOF
            If intCol >= intCols Then
                lngRow = lngRow + 1
                .Rows = .Rows + 1
                intCol = 0
            End If
           .TextMatrix(lngRow, intCol) = IIf(lng药房ID = 0, "(" & NVL(rsTemp!药房) & ")", "") & NVL(rsTemp!规格, "无规格") & ":" & IIf(InStr(1, mstrPrivsOpt, ";显示库存;") = 0, "有库存", Val(NVL(rsTemp!可用数量)) & NVL(rsTemp!住院单位))
           intCol = intCol + 1
            rsTemp.MoveNext
        Loop
        If rsTemp.RecordCount = 0 Then
            .Rows = 1
            .TextMatrix(0, 0) = "无库存!"
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Public Function Get剂量系数(ByVal lng药品ID As Long) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取剂量系数
    '返回:
    '编制:刘兴洪
    '日期:2011-02-18 11:35:29
    '问题:35786
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl剂量系数 As Double, strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
     strSQL = "Select max(剂量系数) as 剂量系数 From 药品规格 where 药品ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng药品ID)
    dbl剂量系数 = NVL(rsTemp!剂量系数, 0)
    Get剂量系数 = dbl剂量系数
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


