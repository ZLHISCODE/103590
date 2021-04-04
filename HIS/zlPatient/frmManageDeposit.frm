VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.12#0"; "zlIDKind.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmManageDeposit 
   AutoRedraw      =   -1  'True
   Caption         =   "病人预交管理"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8910
   Icon            =   "frmManageDeposit.frx":0000
   KeyPreview      =   -1  'True
   ScaleHeight     =   5700
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picSearch 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   5940
      ScaleHeight     =   375
      ScaleWidth      =   2910
      TabIndex        =   5
      Top             =   195
      Width           =   2910
      Begin VB.TextBox txtValue 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   615
         TabIndex        =   6
         ToolTipText     =   "定位F3"
         Top             =   40
         Width           =   2235
      End
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   315
         Left            =   15
         TabIndex        =   7
         Top             =   48
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   556
         Appearance      =   2
         IDKindStr       =   $"frmManageDeposit.frx":08CA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   9
         FontName        =   "宋体"
         IDKind          =   -1
         ShowPropertySet =   -1  'True
         DefaultCardType =   "0"
         AllowAutoICCard =   -1  'True
         AllowAutoIDCard =   -1  'True
         BackColor       =   -2147483633
      End
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   5490
      Top             =   750
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDeposit.frx":0951
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDeposit.frx":0B6B
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDeposit.frx":0D85
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDeposit.frx":0F9F
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDeposit.frx":11B9
            Key             =   "View"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDeposit.frx":1933
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDeposit.frx":1B4D
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDeposit.frx":1D67
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDeposit.frx":1F81
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDeposit.frx":219B
            Key             =   "RollingCurtain"
            Object.Tag             =   "RollingCurtain"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   4605
      Top             =   765
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDeposit.frx":BB32
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDeposit.frx":BD4C
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDeposit.frx":BF66
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDeposit.frx":C180
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDeposit.frx":C39A
            Key             =   "View"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDeposit.frx":CB14
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDeposit.frx":CD2E
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDeposit.frx":CF48
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDeposit.frx":D162
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageDeposit.frx":D37C
            Key             =   "RollingCurtain"
            Object.Tag             =   "RollingCurtain"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.TabStrip tbPage 
      Height          =   300
      Left            =   75
      TabIndex        =   4
      Top             =   825
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   529
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid mshList 
      Height          =   3225
      Left            =   180
      TabIndex        =   3
      Top             =   1515
      Width           =   8670
      _cx             =   15293
      _cy             =   5689
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
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   3
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   5340
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmManageDeposit.frx":DA76
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6906
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3704
            MinWidth        =   3704
            Picture         =   "frmManageDeposit.frx":E30A
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
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   8910
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinWidth1       =   1995
      MinHeight1      =   720
      Width1          =   2010
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbr 
         Height          =   720
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   8790
         _ExtentX        =   15505
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   14
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "Preview"
               Description     =   "预览"
               Object.ToolTipText     =   "预览"
               Object.Tag             =   "预览"
               ImageKey        =   "Preview"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "Print"
               Description     =   "打印"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "收款"
               Key             =   "Deposit"
               Description     =   "收款"
               Object.ToolTipText     =   "进入收款窗口"
               Object.Tag             =   "收款"
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退款"
               Key             =   "Del"
               Description     =   "退款"
               Object.ToolTipText     =   "对当前选中单据退款"
               Object.Tag             =   "退款"
               ImageKey        =   "Del"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查阅"
               Key             =   "View"
               Description     =   "查阅"
               Object.ToolTipText     =   "查阅当前单据的内容"
               Object.Tag             =   "查阅"
               ImageKey        =   "View"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "过滤"
               Key             =   "Filter"
               Description     =   "过滤"
               Object.ToolTipText     =   "设置条件重新读取列表"
               Object.Tag             =   "过滤"
               ImageKey        =   "Filter"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "定位"
               Key             =   "Go"
               Description     =   "定位"
               Object.ToolTipText     =   "定位在当前列表中满足条件的记录上"
               Object.Tag             =   "定位"
               ImageKey        =   "Find"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "轧帐"
               Key             =   "轧帐"
               Object.ToolTipText     =   "收费轧帐"
               Object.Tag             =   "轧帐"
               ImageKey        =   "RollingCurtain"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "SplitRollingCurtain"
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助"
               Object.ToolTipText     =   "当前帮助主题"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Quit"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFile_PrintSet 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuFile_PreView 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFile_Print 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFile_Excel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMoneyEnum 
         Caption         =   "现金点钞(&E)"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileRollingCurtain 
         Caption         =   "收费轧帐(&M)"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuFileRollingCurtainSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileLocalSet 
         Caption         =   "参数设置(&R)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFileLocalSet_ 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuFile_quit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEdit_Deposit 
         Caption         =   "收款(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEdit_Del 
         Caption         =   "退款(&D)"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu mnuEidtBalanceDel 
         Caption         =   "余额退款(&Y)"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuEdit_View 
         Caption         =   "查阅(&V)"
      End
      Begin VB.Menu mnuEditSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditMzTozy 
         Caption         =   "门诊转住院(&M)"
      End
      Begin VB.Menu mnuEditZyToMz 
         Caption         =   "住院转门诊(&Z)"
      End
      Begin VB.Menu mnuEdit_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Print 
         Caption         =   "重打票据(&R)"
      End
      Begin VB.Menu mnuEdit_Print_Supplemental 
         Caption         =   "补打票据(&B)"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "报表(&R)"
      Visible         =   0   'False
      Begin VB.Menu mnuReportItem 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "查看(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "工具栏(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "标准按钮(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuView_Tlb_1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "文本标签(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "状态栏(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuView_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFilter 
         Caption         =   "过滤(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewGo 
         Caption         =   "定位(&G)"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewreFlash 
         Caption         =   "刷新(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpTitle 
         Caption         =   "帮助主题(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&WEB上的中联"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "中联主页(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "中联论坛(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "发送反馈(&K)..."
         End
      End
      Begin VB.Menu mnuHelp_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
End
Attribute VB_Name = "frmManageDeposit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mrsList As ADODB.Recordset  '单据列表
Private mstrFilter As String
Private mblnCancel As Boolean
Private mblnGo As Boolean, mlngGo As Long
Private mlngCurRow As Long, mlngTopRow As Long
Private mstrPrivs As String
Private mlngModul As Long
Private mblnNOMoved As Boolean '显明细时记录当前选择的单据是否在在线数据表中,以其它操作时无需再判断
Private mcllFilterA As Collection
Private mblnNotClick As Boolean
Private mblnUnLoad As Boolean
Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private mstrPrivs_RollingCurtain As String  '收费轧帐管理权限

Private Sub InitFilter()
    '-----------------------------------------------------------------------------------------------------------
    '功能:初始化过滤条件
    '入参:
    '出参:
    '返回:
    '编制:lesfeng
    '日期:2010-01-11 16:10:40
    '-----------------------------------------------------------------------------------------------------------
    Set mcllFilterA = New Collection
    mcllFilterA.Add Array("1901-01-01 00:00:00", "1901-01-01 00:00:00"), "收款时间"
    mcllFilterA.Add Array("", ""), "单据号"
    mcllFilterA.Add Array("", ""), "票据号"
    mcllFilterA.Add "", "门诊号"
    mcllFilterA.Add "", "住院号"
    mcllFilterA.Add "", "姓名"
    mcllFilterA.Add "", "结算号码"
    mcllFilterA.Add "", "收款人"
    mstrFilter = ""
End Sub
Private Sub InitPrepayType()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化预交类型
    '编制:刘兴洪
    '日期:2011-07-14 18:50:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    tbPage.Tabs.Clear
    mblnNotClick = True
    If InStr(1, mstrPrivs, ";门诊预交;") > 0 _
        And InStr(1, mstrPrivs, ";住院预交;") > 0 Then
        tbPage.Tabs.Add , "ALL", "所有预交"
        tbPage.Tabs("ALL").Selected = True
    End If
    If InStr(1, mstrPrivs, ";门诊预交;") > 0 Then
        tbPage.Tabs.Add , "K1", "门诊预交"
    End If
    If InStr(1, mstrPrivs, ";住院预交;") > 0 Then
        tbPage.Tabs.Add , "K2", "住院预交"
    End If
    If tbPage.SelectedItem Is Nothing And tbPage.Tabs.Count <> 0 Then
        tbPage.Tabs(0).Selected = True
    End If
    If tbPage.Tabs.Count = 0 Then
        MsgBox "你不具备门诊预交或住院预交权限,请与系统管理员联系!", vbOKOnly + vbInformation, gstrSysName
        mblnUnLoad = True
    End If
    mblnNotClick = False
End Sub

 
Private Sub cboType_Click()
    If mblnNotClick Then Exit Sub
    ShowBills mstrFilter
End Sub

Private Sub cbr_Resize()
     Call Form_Resize
End Sub

Private Sub Form_Activate()
    If mblnUnLoad Then Unload Me: Exit Sub
    
    Call InitLocPar(mlngModul)
End Sub

Public Sub ActiveIDKindKey()
    IDKind.ActiveFastKey
End Sub

Private Sub mnuEditMzTozy_Click()
    '门诊转住院
      If frmDeposit.zlShowEdit(Me, 0, 4, mstrPrivs, mlngModul, 1) = True Then
        If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then mnuViewReFlash_Click
    End If
End Sub

Private Sub mnuEditZyToMz_Click()
    '门诊转住院
      If frmDeposit.zlShowEdit(Me, 0, 4, mstrPrivs, mlngModul, 2) = True Then
        If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then mnuViewReFlash_Click
    End If
End Sub

Private Sub mnuEidtBalanceDel_Click()
    '余额退款
    If frmDeposit.zlShowEdit(Me, 0, 3, mstrPrivs, mlngModul) = True Then
        If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then mnuViewReFlash_Click
    End If
End Sub

Private Sub mnuFileLocalSet_Click()
    Call frmLocalSet.zlSetPara(Me, mstrPrivs, mlngModul)
    
'    If glng预交ID > 0 Then
'        If Not ExistBill(glng预交ID, 2) Then
'            zldatabase.SetPara "共用预交票据批次", 0, glngSys, mlngModul
'            glng预交ID = 0
'        End If
'    End If
End Sub

Private Sub mnuFileMoneyEnum_Click()
    Call frmMoneyEnum.ShowMe(Me)
End Sub

 

Private Sub mnuFileRollingCurtain_Click()
    Call zlExecuteChargeRollingCurtain(Me)
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Dim strNo As String, str病人ID As String, str住院号 As String
    With mshList
        strNo = mshList.TextMatrix(mshList.Row, .ColIndex("单据号"))
        str病人ID = Trim(.TextMatrix(.Row, .ColIndex("病人ID")))
        str住院号 = Trim(.TextMatrix(.Row, .ColIndex("住院号")))
    End With
    If strNo <> "" Then
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
                "NO=" & strNo, "病人ID=" & str病人ID, "住院号=" & str住院号)
    Else
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me)
    End If
End Sub

Private Sub mnuViewFilter_Click()
    frmDepositFilter.Show 1, Me
    If gblnOK Then
        mstrFilter = frmDepositFilter.mstrFilter
        Set mcllFilterA = frmDepositFilter.mcllFilter
        mblnCancel = (frmDepositFilter.chkCancel.Value = Checked And frmDepositFilter.chk收款.Value = 0)
        mnuViewReFlash_Click
    End If
End Sub
Private Sub mshList_DblClick()
    If mshList.MouseRow = 0 Then Exit Sub
    If mnuEdit_View.Enabled Then mnuEdit_View_Click
End Sub

Private Sub mshList_EnterCell()
    Dim strNo As String, lng记录状态 As Long
    Dim str票据号 As String
    With mshList
        If .Row = 0 Or .TextMatrix(.Row, .ColIndex("单据号")) = "" Then Exit Sub
        strNo = .TextMatrix(.Row, .ColIndex("单据号"))
        lng记录状态 = Val(.TextMatrix(.Row, .ColIndex("记录状态")))
        str票据号 = .TextMatrix(.Row, .ColIndex("票据号"))
        mlngGo = .Row: mlngCurRow = .Row: mlngTopRow = .TopRow
    End With
    If frmDepositFilter.mblnDateMoved Then
        mblnNOMoved = zlDatabase.NOMoved("病人预交记录", strNo, , "1", Me.Caption)
    Else
        mblnNOMoved = False
    End If
    
    'mshList.TextMatrix(mshList.Row, mshList.Cols - 1)
    Select Case lng记录状态
        Case 1
            SetMenu (True)
            mnuEdit_Print_Supplemental.Enabled = str票据号 = ""
        Case 2
            SetMenu (False)
            mnuEdit_View.Enabled = True
            tbr.Buttons.Item("View").Enabled = True
        Case 3
            SetMenu (False)
            mnuEdit_View.Enabled = True
            tbr.Buttons.Item("View").Enabled = True
    End Select
    
End Sub
Private Sub mshList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete And mnuEdit_Del.Enabled And mnuEdit_Del.Visible Then Call mnuEdit_Del_Click
End Sub

Private Sub mshList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuEdit, 2
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3
            '始终从当前行开始
            If mnuViewGo.Enabled Then Call SeekBill(False)
'        Case vbKeyReturn
'            If mnuEdit_View.Enabled Then mnuEdit_View_Click
        Case vbKeyEscape
            mblnGo = False
        Case Else
            IDKind.ActiveFastKey
    End Select
End Sub

Private Sub mnuEdit_Del_Click()
    Dim strNo As String, str操作员 As String
    Dim byt预交类型 As Byte
    With mshList
        strNo = .TextMatrix(.Row, .ColIndex("单据号"))
        str操作员 = .TextMatrix(.Row, .ColIndex("操作员"))
        byt预交类型 = Val(.TextMatrix(.Row, .ColIndex("预交类别ID")))
    End With
    If strNo = "" Then
        MsgBox "当前没有记录可以退款！", vbExclamation, gstrSysName
        Exit Sub
    End If
        
    '单据权限
    If Not BillOperCheck(6, str操作员, _
        CDate(mshList.TextMatrix(mshList.Row, mshList.ColIndex("操作时间"))), "退款") Then Exit Sub
    
    If Val(mshList.TextMatrix(mshList.Row, mshList.ColIndex("金额"))) < 0 Then
        MsgBox "该缴款记录金额为负,表示退款,不能执行该操作！", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    '是否已转入后备数据表中
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNo, 6, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '此时已转入在线数据表
    End If
    
    If is代收款(strNo) Then
         If InStr(mstrPrivs, "代收款退款") = 0 Then
            MsgBox "你没有权限进行代收款退款操作！", vbInformation, gstrSysName
            Exit Sub
        End If
    ElseIf InStr(mstrPrivs, "预交退款") = 0 Then
        MsgBox "你没有权限进行预交退款操作！", vbInformation, gstrSysName
        Exit Sub
    Else
        If HaveSpare(strNo) = 0 And InStr(mstrPrivs, "预交结清退款") = 0 Then
            MsgBox "该病人已没有预交余额,你没有权限作废这张单据！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If HaveBalance(strNo) <> 0 Then
            MsgBox "该笔预交已经被病人在结帐时使用,你不能作废这张单据！", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    On Error Resume Next
    Err.Clear
    If frmDeposit.zlShowEdit(Me, 0, 2, mstrPrivs, mlngModul, byt预交类型, strNo) = True Then
        If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then mnuViewReFlash_Click
    End If
End Sub

Private Sub mnuHelpTitle_Click()
ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub mnuEdit_Deposit_Click()
    Dim byt类型 As Byte
    If Not tbPage.SelectedItem Is Nothing Then
        byt类型 = Val(Mid(tbPage.SelectedItem.Key, 2))
    End If
    If frmDeposit.zlShowEdit(Me, 0, 0, mstrPrivs, mlngModul, byt类型) = True Then
        If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then mnuViewReFlash_Click
    End If
End Sub

Private Sub mnuEdit_View_Click()
    Dim blnViewCancel  As Boolean
    Dim strNo As String, str操作员 As String, byt预交类型 As Byte
    Dim int记录状态 As Integer, blnNOMoved As Boolean
    With mshList
        strNo = .TextMatrix(.Row, .ColIndex("单据号"))
        str操作员 = .TextMatrix(.Row, .ColIndex("操作员"))
        byt预交类型 = Val(.TextMatrix(.Row, .ColIndex("预交类别ID")))
        int记录状态 = Val(.TextMatrix(.Row, .ColIndex("记录状态")))
        blnViewCancel = int记录状态 = 2
    End With
    '是否已转入后备数据表中
    If mblnNOMoved Then
        blnNOMoved = zlDatabase.NOMoved("病人预交记录", strNo, , "1")
    End If
    
    If strNo = "" Then MsgBox "当前没有记录可以查阅！", vbExclamation, gstrSysName: Exit Sub
    '显示单据内容
    Call frmDeposit.zlShowEdit(Me, 0, 1, mstrPrivs, mlngModul, byt预交类型, strNo, blnViewCancel, blnNOMoved)

End Sub

Private Sub mnuFile_Quit_Click()
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuViewReFlash_Click()
    ShowBills mstrFilter
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = Not stbThis.Visible
    Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    cbr.Visible = Not cbr.Visible
    mnuViewToolText.Enabled = Not mnuViewToolText.Enabled
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim i As Integer
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For i = 1 To tbr.Buttons.Count
        tbr.Buttons(i).Caption = IIf(mnuViewToolText.Checked, tbr.Buttons(i).Tag, "")
    Next
    cbr.Bands(1).MinHeight = tbr.ButtonHeight
    Form_Resize
End Sub

Private Sub tbPage_Click()
        If mblnNotClick Then Exit Sub
        ShowBills mstrFilter
       If mshList.Enabled And mshList.Visible Then mshList.SetFocus
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Quit"
            mnuFile_Quit_Click
        Case "Go" '定位
            mnuViewGo_Click
        Case "Filter" '过滤
            mnuViewFilter_Click
        Case "View"
            mnuEdit_View_Click
        Case "Deposit"
            mnuEdit_Deposit_Click
        Case "Del"
            mnuEdit_Del_Click
        Case "Print"
            mnuFile_Print_Click
        Case "Preview"
            mnuFile_PreView_Click
        Case "轧帐"
            mnuFileRollingCurtain_Click
        Case "Help"
            mnuHelpTitle_Click
    End Select
End Sub

Private Sub tbr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool, 2
End Sub

Private Sub mnuFile_Excel_Click()
    Call OutputList(3)
End Sub

Private Sub mnuFile_PreView_Click()
    Call OutputList(2)
End Sub

Private Sub mnuFile_Print_Click()
    Call OutputList(1)
End Sub

Private Sub mnuFile_PrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub OutputList(bytStyle As Byte)
'功能：输入出列表
'参数：bytStyle=1-打印,2-预览,3-输出到Excel
    Dim objOut As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim bytR As Byte, intRow As Integer
    
    intRow = mshList.Row
    
    '表头
    objOut.Title.Text = "预交款收款清单"
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '表项
    With frmDepositFilter
        If IsNull(.dtpEnd.Value) Then
            objRow.Add "时间：" & Format(.dtpBegin.Value, "yyyy-MM-dd")
        Else
            objRow.Add "时间：" & Format(.dtpBegin.Value, "yyyy-MM-dd") & " 至 " & Format(.dtpEnd.Value, "yyyy-MM-dd")
        End If
        objRow.Add "性质：" & IIf(.chkCancel.Value = 1, "退款记录", "收款记录")
        objOut.UnderAppRows.Add objRow
    End With
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印日期：" & Format(zlDatabase.Currentdate(), "yyyy年MM月dd日")
    objOut.BelowAppRows.Add objRow
    
    '表体
    mshList.Redraw = False
    Set objOut.Body = mshList
    
    '输出
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    
    mshList.Row = intRow
    mshList.Col = 0: mshList.ColSel = mshList.Cols - 1
    mshList.Redraw = True
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage hWnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo hWnd
End Sub

Private Sub SetHeader()
    Dim strHead As String
    Dim i As Integer
    Dim lngRow As Long, lngCol As Long
    
 
    strHead = "单据号,4,850|票据号,4,1050|操作员,1,850|操作时间,4,1850|病人ID,1,750|门诊号,1,750|住院号,1,750|姓名,1,800|性别,4,500|年龄,4,500|科室,1,850|金额,7,850|结算方式,1,850|结算号码,1,1500|摘要,1,1500|记录状态,1,0|医疗付款方式,1,1500|预交类别,4,800|预交类别ID,1,0"
    With mshList
        .Redraw = False
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .FixedAlignment(i) = 4
            .ColKey(i) = UCase(Trim(.TextMatrix(0, i)))
        Next
        If Not Visible Then Call RestoreFlexState(mshList, App.ProductName & "\" & Me.Name)
        .ColHidden(.ColIndex("预交类别ID")) = True
        .RowHeight(0) = 320
        '恢复上次行
        If mlngCurRow = 0 Then mlngCurRow = 1
        If mlngTopRow = 0 Then mlngTopRow = 1
        If mlngCurRow <= .Rows - 1 Then
            .Row = mlngCurRow
        Else
            .Row = .Rows - 1
        End If
        If mlngTopRow <= .Rows - 1 Then
            .TopRow = mlngTopRow
        Else
            .TopRow = .Row
        End If
        .Col = 0: .ColSel = .Cols - 1
        Call mshList_EnterCell
        For lngRow = 1 To .Rows - 1
            .Row = lngRow
            For lngCol = 1 To .Cols - 1
                .Col = lngCol
                If .TextMatrix(lngRow, .ColIndex("记录状态")) = "2" Then
                    .CellForeColor = &HFF&
                ElseIf .TextMatrix(lngRow, .ColIndex("记录状态")) = "3" Then
                    .CellForeColor = &HFF0000
                End If
                
                If .TextMatrix(0, lngCol) = "金额" And IsNumeric(.TextMatrix(lngRow, lngCol)) Then
                   .TextMatrix(lngRow, lngCol) = Format(.TextMatrix(lngRow, lngCol), "0.00")
                End If
            Next lngCol
        Next lngRow
        .Redraw = True
    End With
End Sub

Private Sub ShowBills(Optional strIF As String, Optional blnSort As Boolean, Optional bytMode As Byte = 0, Optional objCard As Card)
'功能:按条件读取单据列表(过滤功能)
'参数:strIF=以"AND"开始的条件串
'     blnSort=不重新读取数据,仅重新显示已排序的内容
    Dim dbl金额  As Double, strKind As String, strFind As String
    Dim int预交类别 As Integer, strWhere As String, lng卡类别ID As Long
    Dim lng病人ID As Long, strPassWord As String, strErrMsg As String
    Dim strTable As String
    On Error GoTo errH
    
    If bytMode <> 0 Then
        If txtValue.Text = "" Then Exit Sub
        strKind = objCard.名称
        If (Left(txtValue.Text, 1) = "-" And IsNumeric(Mid(txtValue.Text, 2))) Then
            lng病人ID = Val(Mid(txtValue.Text, 2))
            strIF = " And B.病人ID=[1]"
            strFind = lng病人ID
        '89607: 李南春,2015/10/20,增加住院号查找,参数值取有效数字
        ElseIf (Left(txtValue.Text, 1) = "*" And IsNumeric(Mid(txtValue.Text, 2))) Then
            strFind = Val(Mid(txtValue.Text, 2))
            strIF = " And B.门诊号=[1]"
        ElseIf (Left(txtValue.Text, 1) = "+" And IsNumeric(Mid(txtValue.Text, 2))) Then
            strFind = Val(Mid(txtValue.Text, 2))
            strIF = " And B.病人ID=(Select Nvl(Max(病人ID),0) as 病人ID From 病案主页 Where 住院号=[1])"
        Else
            Select Case strKind
            Case "姓名"
                lng卡类别ID = IDKind.GetDefaultCardTypeID
                If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, txtValue.Text, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                If lng病人ID <= 0 Then
                    strFind = txtValue.Text & "%"
                    strIF = " And B.姓名 Like [1]"
                Else
                    strFind = lng病人ID
                    strIF = " And B.病人ID=[1]"
                End If
            Case "身份证号"
                strFind = txtValue.Text
                strIF = " And B.身份证号=[1]"
            Case "医保号"
                strFind = txtValue.Text
                strIF = " And B.医保号=[1]"
            Case "IC卡号"
                If gobjSquare.objSquareCard.zlGetPatiID("IC卡号", txtValue.Text, False, lng病人ID, _
                    strPassWord, strErrMsg) = False Then lng病人ID = 0
                strFind = lng病人ID
                strIF = " And B.病人ID=[1]"
            Case "门诊号"
                strFind = txtValue.Text
                strIF = " And B.门诊号=[1]"
            Case "住院号"
                strFind = txtValue.Text
                strIF = " And B.病人ID=(Select Nvl(Max(病人ID),0) as 病人ID From 病案主页 Where 住院号=[1])"
            Case Else
                '其他类别的,获取相关的病人ID
                '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|
                '是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密)
                '第7位后,就只能用索引,不然取不到数
                lng卡类别ID = objCard.接口序号
                If lng卡类别ID <> 0 Then
                    If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, txtValue.Text, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(strKind, txtValue.Text, False, lng病人ID, _
                        strPassWord, strErrMsg) = False Then lng病人ID = 0
                End If
                strFind = lng病人ID
                strIF = " And B.病人ID=[1]"
            End Select
        End If
    End If
    
    If Not blnSort Then
        Call zlCommFun.ShowFlash("正在读取单据列表,请稍候 ...", Me)
        DoEvents
        Me.Refresh
        strWhere = ""
        If tbPage.SelectedItem Is Nothing Then Exit Sub
        
        If gbln分站点显示 Then
             strWhere = strWhere & _
            " 　And Exists (Select 1 From 人员表 E, 部门人员 F, 部门表 G " & _
            " 　Where A.操作员姓名=e.姓名  And e.Id = f.人员id And f.部门id = g.Id And (g.站点 ='" & gstrNodeNo & "' Or g.站点 Is Null))"
        End If
        
        '115601:李南春,2017/10/23,指定查询列
        If frmDepositFilter.mblnDateMoved Then
            strTable = "(Select NO,实际票号,操作员姓名,收款时间,病人ID,金额,结算方式,结算号码,摘要,记录状态,主页ID,预交类别,科室ID,记录性质 From 病人预交记录  " & _
                        "UNION ALL " & _
                        "Select NO,实际票号,操作员姓名,收款时间,病人ID,金额,结算方式,结算号码,摘要,记录状态,主页ID,预交类别,科室ID,记录性质 From H病人预交记录)"
        Else
            strTable = "病人预交记录"
        End If
            
        If Left(tbPage.SelectedItem.Key, 1) = "K" Then
            int预交类别 = Val(Mid(tbPage.SelectedItem.Key, 2))
            strWhere = " And  Nvl(A.预交类别, 0) = " & IIf(bytMode = 0, "[12]", "[2]")
        End If
          
         gstrSQL = _
        "   Select A.NO as 单据号,A.实际票号 as 票据号,A.操作员姓名 as 操作员," & _
        "           To_Char(A.收款时间,'YYYY-MM-DD HH24:MI:SS') as 操作时间," & _
        "           A.病人ID,B.门诊号,B.住院号,B.姓名,B.性别,B.年龄,D.名称 as 科室," & _
        "           To_Char(Sum(A.金额),'9999999990.00') as 金额," & _
        "           A.结算方式,A.结算号码,A.摘要,A.记录状态," & _
        "           Decode(Nvl(A.主页ID,0),0,B.医疗付款方式,C.医疗付款方式) 医疗付款方式, " & _
        "           Decode(nvl(A.预交类别,2),1,'门诊预交', '住院预交') as 预交类别, nvl(A.预交类别,0) as 预交类别ID" & _
        " From " & strTable & " A,病人信息 B,病案主页 C,部门表 D " & _
        " Where A.病人ID=B.病人ID AND A.病人ID=C.病人ID(+) AND NVL(A.主页ID,0)=C.主页ID(+) And A.科室ID=D.ID(+) And A.记录性质=1 " & strIF & strWhere & _
        " Group by A.NO,A.记录状态,A.实际票号 ,Nvl(A.预交类别, 0),Decode(nvl(A.预交类别,2),1,'门诊预交', '住院预交'),A.操作员姓名," & _
        "           To_Char(A.收款时间,'YYYY-MM-DD HH24:MI:SS'),A.病人ID,B.门诊号,B.住院号,B.姓名,B.年龄," & _
        "           B.性别 , D.名称, A.结算方式,A.结算号码, A.摘要,Decode(Nvl(A.主页id, 0), 0, B.医疗付款方式, C.医疗付款方式)" & _
        " Order by 操作时间 Desc,单据号 Desc"
        
        Set mrsList = New ADODB.Recordset
            
        If bytMode = 0 Then
            Set mrsList = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CDate(mcllFilterA("收款时间")(0)), CDate(mcllFilterA("收款时间")(1)), _
            CStr(mcllFilterA("单据号")(0)), CStr(mcllFilterA("单据号")(1)), _
            CStr(mcllFilterA("票据号")(0)), CStr(mcllFilterA("票据号")(1)), CLng(Val(mcllFilterA("住院号"))), _
            CStr(mcllFilterA("姓名")), CStr(mcllFilterA("结算号码")), CStr(mcllFilterA("收款人")), CLng(Val(mcllFilterA("门诊号"))), int预交类别)
        Else
            Set mrsList = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strFind, int预交类别)
        End If
    End If
    
    mshList.Clear
    mshList.Rows = 2
    If mrsList.EOF Then
        Call SetHeader
        stbThis.Panels(2).Text = "当前设置没有过滤出任何单据"
        Call SetMenu(False)
    Else
        If (Left(txtValue.Text, 1) = "-" And IsNumeric(Mid(txtValue.Text, 2))) Then
            txtValue.Text = NVL(mrsList!姓名)
            IDKind.IDKind = 1
        ElseIf (Left(txtValue.Text, 1) = "*" And IsNumeric(Mid(txtValue.Text, 2))) Then
            txtValue.Text = NVL(mrsList!姓名)
            IDKind.IDKind = 1
        ElseIf (Left(txtValue.Text, 1) = "+" And IsNumeric(Mid(txtValue.Text, 2))) Then
            txtValue.Text = NVL(mrsList!姓名)
            IDKind.IDKind = 1
        End If
        
        Set mshList.DataSource = mrsList: Call SetHeader
        mrsList.MoveFirst: dbl金额 = 0
        Do While Not mrsList.EOF
            dbl金额 = dbl金额 + mrsList!金额
            mrsList.MoveNext
        Loop
        mrsList.MoveFirst
        stbThis.Panels(2) = "共 " & mrsList.RecordCount & " 张单据,合计:" & Format(dbl金额, "0.00")
        Call SetMenu(True)
    End If
    mnuEdit_Del.Enabled = Not mblnCancel And Not mrsList.EOF
    mnuEdit_Print.Enabled = Not mblnCancel And Not mrsList.EOF
    mnuEdit_Print_Supplemental.Enabled = Not mblnCancel And Not mrsList.EOF
    tbr.Buttons("Del").Enabled = Not mblnCancel And Not mrsList.EOF
    If Not blnSort Then Call zlCommFun.StopFlash
    Me.Refresh
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub SetMenu(blnUsed As Boolean)
'功能：根据有无记录设置菜单可用状态
    mnuFile_Print.Enabled = blnUsed
    mnuFile_PreView.Enabled = blnUsed
    mnuFile_Excel.Enabled = blnUsed
    tbr.Buttons("Print").Enabled = blnUsed
    tbr.Buttons("Preview").Enabled = blnUsed
    
    mnuEdit_Del.Enabled = blnUsed
    mnuEdit_View.Enabled = blnUsed
    mnuEdit_Print.Enabled = blnUsed
    mnuEdit_Print_Supplemental.Enabled = blnUsed
    tbr.Buttons("Del").Enabled = blnUsed
    tbr.Buttons("View").Enabled = blnUsed
    
    mnuViewGo.Enabled = blnUsed
    tbr.Buttons("Go").Enabled = blnUsed
End Sub

Private Sub Form_Load()
    Dim Curdate As Date
    Dim blnHavePrivs As Boolean
    
    mstrPrivs_RollingCurtain = ";" & GetPrivFunc(glngSys, 1506) & ";"
    mstrPrivs = gstrPrivs: mlngModul = glngModul
    mblnUnLoad = False
    Call InitFilter: Call InitPrepayType

    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs, "ZL" & glngSys \ 100 & "_INSIDE_1103_1")  '隐藏预交款缴款书
    Call RestoreWinState(Me, App.ProductName)
     
    
    '权限设置
    If InStr(mstrPrivs, "预交收款") = 0 And InStr(mstrPrivs, "代收款收取") = 0 Then
        mnuEdit_Deposit.Visible = False
        tbr.Buttons("Deposit").Visible = False
        mnuEdit_Print.Visible = False
        mnuEdit_2.Visible = False
    End If
    '52328
    mnuEdit_Print_Supplemental.Visible = _
        (InStr(mstrPrivs, ";代收款收取;") > 0 Or InStr(mstrPrivs, ";预交收款;") > 0) _
        And InStr(mstrPrivs, ";补打票据;") > 0
        
    If InStr(mstrPrivs, "预交退款") = 0 And InStr(mstrPrivs, "代收款退款") = 0 Then
        mnuEdit_Del.Visible = False
        tbr.Buttons("Del").Visible = False
    End If
    mnuEidtBalanceDel.Visible = InStr(1, mstrPrivs, ";预交退款;") > 0
    mnuEditMzTozy.Visible = InStr(1, mstrPrivs, ";门诊预交转住院;") > 0
    mnuEditZyToMz.Visible = InStr(1, mstrPrivs, ";住院预交转门诊;") > 0
    mnuEditSplit.Visible = InStr(1, mstrPrivs, ";门诊预交转住院;") > 0 Or InStr(1, mstrPrivs, ";住院预交转门诊;") > 0
    '收费轧帐管理
    blnHavePrivs = InStr(mstrPrivs_RollingCurtain, ";轧帐;") > 0
    mnuFileRollingCurtain.Visible = blnHavePrivs
    mnuFileRollingCurtainSplit.Visible = blnHavePrivs
    tbr.Buttons("轧帐").Visible = blnHavePrivs
    tbr.Buttons("SplitRollingCurtain").Visible = blnHavePrivs

    If InStr(";" & mstrPrivs & ";", ";重打票据;") = 0 Then
        mnuEdit_Print.Visible = False
    End If
    
    '缺省过滤条件
    Curdate = zlDatabase.Currentdate
    'by lesfeng 2010-03-06 性能优化
    mstrFilter = ""
    mstrFilter = mstrFilter & " And (收款时间  Between [1] And [2]) "
    mstrFilter = mstrFilter & " And 记录状态=1"
    mstrFilter = mstrFilter & " And 操作员姓名=[10]"
    
    mcllFilterA.Remove "收款时间"
    mcllFilterA.Add Array(Format(Curdate, "yyyy-mm-dd") & " 00:00:00", Format(Curdate, "yyyy-mm-dd") & " 23:59:59"), "收款时间"
    mcllFilterA.Remove "收款人"
    mcllFilterA.Add Trim(UserInfo.姓名), "收款人"
    mblnCancel = False
    
    Call SetHeader
    Call SetMenu(False)
    Set mobjIDCard = New clsIDCard
    Set mobjICCard = New clsICCard
    Call mobjIDCard.SetParent(Me.hWnd)
    Call mobjICCard.SetParent(Me.hWnd)
    Call InitIDKind
    stbThis.Panels(2).Text = "请刷新清单或重新设置过滤条件"
End Sub

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    If txtValue.Visible Then txtValue.SetFocus
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    txtValue.Text = objPatiInfor.卡号
    Call ShowBills("", , 1, objCard)
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strCardNo As String)
    If txtValue.Locked Or txtValue.Text <> "" Or Not Me.ActiveControl Is txtValue Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("IC卡", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtValue.Text = strCardNo
    Call ShowBills("", , 1, objCard)
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    If txtValue.Locked Or txtValue.Text <> "" Or Not Me.ActiveControl Is txtValue Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("身份证", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtValue.Text = strID
    Call ShowBills("", , 1, objCard)
End Sub

Private Sub txtValue_Change()
    If Me.ActiveControl Is txtValue Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtValue.Text = "")
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtValue.Text = "")
        IDKind.SetAutoReadCard txtValue.Text = ""
    End If
End Sub

Private Sub txtValue_GotFocus()
    Call zlControl.TxtSelAll(txtValue)
    Call zlCommFun.OpenIme(True)
    If txtValue.Text = "" And ActiveControl Is txtValue Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtValue.Text = "")
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtValue.Text = "")
        IDKind.SetAutoReadCard txtValue.Text = ""
    End If
End Sub

Private Sub txtValue_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        '0-门诊号;1-姓名;2-挂号单;3-就诊卡号;4-医保号
        Call ShowBills("", , 1, IDKind.GetCurCard)
        zlControl.TxtSelAll txtValue
    End If
End Sub

Private Sub txtValue_KeyPress(KeyAscii As Integer)
    '0-门诊号,1-姓名,2-挂号单,3-就诊卡号,4-医保号
    Dim blnCard As Boolean
    Dim strKind As String, intLen As Integer
    strKind = IDKind.GetCurCard.名称
    txtValue.PasswordChar = IIf(IDKind.GetCurCard.卡号密文规则 <> "" And IDKind.GetCurCard.卡号密文规则 <> "0", "*", "")
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtValue.IMEMode = 0
    
    '取缺省的刷卡方式
            '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|
            '是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密)
            '第7位后,就只能用索引,不然取不到数
    Select Case strKind
    Case "姓名"
        blnCard = zlCommFun.InputIsCard(txtValue, KeyAscii, gobjSquare.bln缺省卡号密文)
        intLen = gobjSquare.int缺省卡号长度
    Case "门诊号"
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    Case "挂号单"
    Case "医保号"
    Case Else
            If IDKind.GetCurCard.接口序号 <> 0 Then
                blnCard = zlCommFun.InputIsCard(txtValue, KeyAscii, IDKind.GetCurCard.卡号密文规则 <> "" And IDKind.GetCurCard.卡号密文规则 <> "0")
                intLen = IDKind.GetCurCard.卡号长度
            End If
    End Select
    
    '刷卡完毕或输入号码后回车
    If blnCard And Len(txtValue.Text) = intLen - 1 And KeyAscii <> 8 Then
        If KeyAscii <> 13 Then
            txtValue.Text = txtValue.Text & Chr(KeyAscii)
            txtValue.SelStart = Len(txtValue.Text)
        End If
        KeyAscii = 0
        Call ShowBills("", , 1, IDKind.GetCurCard)
        zlControl.TxtSelAll txtValue
   End If
End Sub

Private Sub txtValue_LostFocus()
    Call zlCommFun.OpenIme
    IDKind.SetAutoReadCard False
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
End Sub

Private Sub txtValue_Validate(Cancel As Boolean)
    txtValue.Text = Trim(txtValue.Text)
End Sub

'初始化IDKIND
Private Function InitIDKind() As Boolean
    Dim objCard As Card
    Dim lngCardID As Long
    Call IDKind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "姓|姓名|0|0|0|0|0|0;医|医保号|0|0|0|0|0|0;身|身份证号|0|0|0|0|0|0;IC|IC卡号|1|0|0|0|0|0;门|门诊号|0|0|0|0|0|0;住|住院号|0|0|0|0|0|0", txtValue)
    Set objCard = IDKind.GetfaultCard
    If IDKind.Cards.按缺省卡查找 And Not objCard Is Nothing Then
        gobjSquare.bln缺省卡号密文 = objCard.卡号密文规则 <> ""
        gobjSquare.int缺省卡号长度 = objCard.卡号长度
        Set gobjSquare.objDefaultCard = objCard
    Else
        gobjSquare.bln缺省卡号密文 = IDKind.Cards.加密显示
        gobjSquare.int缺省卡号长度 = 100
    End If
End Function

Private Sub Form_Resize()
    Dim cbrH As Long '工具条占用高度
    Dim staH As Long '状态栏占用高度
    
    On Error Resume Next
    
    If WindowState = 1 Then Exit Sub
    
    mshList.MousePointer = 0
    
    '靠齐控件宽度和高度
    cbrH = IIf(cbr.Visible, cbr.Height, 0)
    staH = IIf(stbThis.Visible, stbThis.Height, 0)
    With mshList
        .Left = Me.ScaleLeft
        tbPage.Top = Me.ScaleTop + cbrH + 20
        .Top = tbPage.Top + tbPage.Height + 10
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - .Top - staH
        tbPage.Width = Me.ScaleWidth
        tbPage.Left = ScaleLeft
    End With
    picSearch.Left = Me.Width - 3500
    If picSearch.Left < 5500 Then picSearch.Left = 5500
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrFilter = ""
    Unload frmDepositFilter
    Unload frmDepositFind
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub mnuViewGo_Click()
    If Not mblnCancel Then
        frmDepositFind.lbl操作员.Caption = "收款人"
    Else
        frmDepositFind.lbl操作员.Caption = "退款人"
    End If
    frmDepositFind.Show 1, Me
    If gblnOK Then Call SeekBill(frmDepositFind.optHead)
End Sub

Private Sub SeekBill(blnHead As Boolean)
    Dim i As Long
    Dim blnFill As Boolean
    
    Screen.MousePointer = 11
    mblnGo = True
    stbThis.Panels(2).Text = "正在定位满足条件的单据,按ESC终止 ..."
    Me.Refresh
    
    For i = IIf(blnHead, 1, mlngGo) To mshList.Rows - 1
        DoEvents
        
        '比较条件
        blnFill = True
        With frmDepositFind
            If .txtNO.Text <> "" Then
                blnFill = blnFill And mshList.TextMatrix(i, mshList.ColIndex("单据号")) = .txtNO.Text
            End If
            If .txtFact.Text <> "" Then
                blnFill = blnFill And mshList.TextMatrix(i, mshList.ColIndex("票据号")) = .txtFact.Text
            End If
            If .cbo操作员.ListIndex > 0 Then
                If Not mblnCancel Then
                    blnFill = blnFill And mshList.TextMatrix(i, mshList.ColIndex("收款人")) = zlCommFun.GetNeedName(.cbo操作员.Text)
                Else
                    blnFill = blnFill And mshList.TextMatrix(i, mshList.ColIndex("退款人")) = zlCommFun.GetNeedName(.cbo操作员.Text)
                End If
            End If
            If .txt姓名.Text <> "" Then
                blnFill = blnFill And UCase(mshList.TextMatrix(i, mshList.ColIndex("姓名"))) Like "*" & UCase(.txt姓名.Text) & "*"
            End If
            If IsNumeric(.txt住院号.Text) Then
                blnFill = blnFill And Val(mshList.TextMatrix(i, mshList.ColIndex("住院号"))) = Val(.txt住院号.Text)
            End If
        End With
        
        '满足则退出
        If blnFill Then
            mlngGo = i + 1
            mshList.Row = i: mshList.TopRow = i
            mshList.Col = 0: mshList.ColSel = mshList.Cols - 1
            stbThis.Panels(2).Text = "找到一条记录"
            Screen.MousePointer = 0: Exit Sub
        End If
        
        '按ESC取消
        If mblnGo = False Then
            stbThis.Panels(2).Text = "用户取消定位操作"
            Screen.MousePointer = 0: Exit Sub
        End If
    Next
    mlngGo = 1
    stbThis.Panels(2).Text = "已定位到清单尾部"
    Screen.MousePointer = 0
End Sub
 

Private Sub mnuEdit_Print_Click()
    Call PrintBill(0)
End Sub

Private Sub mnuEdit_Print_Supplemental_Click()
    Call PrintBill(1)
End Sub

Private Sub PrintBill(bytMode As Byte)
    '功能：当前收款记录重打或被打一张票据
    'bytMode=0-重打,1-补打
    Dim strSQL As String, strInvoice As String, strNo As String
    Dim lng领用ID As Long, blnValid As Boolean, blnInput As Boolean, lng病人ID As Long
    Dim byt预交类型 As Byte, str操作员 As String, str操作时间 As String, str发票号 As String
    Dim factProperty As Ty_FactProperty
    Dim strNos As String, intInvoiceFormat As Integer, blnTurnMZToZY As Boolean
                
    On Error GoTo errHandle
    With mshList
        strNo = .TextMatrix(.Row, .ColIndex("单据号"))
        str操作员 = .TextMatrix(.Row, .ColIndex("操作员"))
        byt预交类型 = Val(.TextMatrix(.Row, .ColIndex("预交类别ID")))
        str操作时间 = .TextMatrix(.Row, .ColIndex("操作时间"))
        str发票号 = Trim(.TextMatrix(.Row, .ColIndex("票据号")))
        lng病人ID = Val(.TextMatrix(.Row, .ColIndex("病人ID")))
    End With
    If strNo = "" Then
        MsgBox "当前没有记录可以重打票据！", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    '是否已转入后备数据表中
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNo, 6, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '此时已转入在线数据表
    End If
    factProperty = zl_GetInvoicePreperty(mlngModul, 2, CStr(byt预交类型))
    
    strNos = GetTurnMZToZYMultiNOs(bytMode = 0, strNo, blnTurnMZToZY, mblnNOMoved)
    If blnTurnMZToZY Then
        If strNos = "" Then
            MsgBox "当前没有记录可以重打票据！", vbExclamation, gstrSysName
            Exit Sub
        End If
        intInvoiceFormat = Val(zlDatabase.GetPara(284, glngSys, , "0"))
    Else
        strNos = strNo
        intInvoiceFormat = factProperty.intInvoiceFormat
    End If
    
    '单据权限
    If bytMode = 0 Then
        If Not BillOperCheck(6, str操作员, CDate(str操作时间), "重打") Then Exit Sub
    Else
        If str发票号 <> "" Then
            MsgBox "当前单据已打印过票据,不能进行补打！", vbExclamation, gstrSysName
            Exit Sub
        End If
    End If
    
    '如果严格控制票据使用
    If gblnBill预交 Then
        lng领用ID = CheckUsedBill(2, IIf(lng领用ID > 0, lng领用ID, factProperty.lngShareUseID), , CStr(byt预交类型))
        Select Case lng领用ID
            Case -1
                MsgBox "你没有自用和共用的预交票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
            Case -2
                MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
        End Select
        If lng领用ID <= 0 Then Exit Sub
    End If

    If ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me) Then
        If Not gblnBill预交 Then
            '有可能是第一次使用
            Do
                blnInput = False
                '非严格控制时直接从本地读取
                strInvoice = UCase(zlDatabase.GetPara("当前预交票据号", glngSys, mlngModul, ""))
                If strInvoice = "" Then
                    strInvoice = UCase(InputBox("没有找到已用的最大票据号码，无法确定将要使用的开始票据号。" & _
                                    vbCrLf & "请输入将要使用的开始票据号码：", gstrSysName, _
                                    "", Me.Left + 1500, Me.Top + 1500))
                    blnInput = True
                Else
                    strInvoice = zlCommFun.IncStr(strInvoice)
                    strInvoice = UCase(InputBox("请确认重打使用的开始票据号码：", gstrSysName, _
                                    strInvoice, Me.Left + 1500, Me.Top + 1500))
                    blnInput = True
                End If
                    
                '用户取消输入,允许打印
                If strInvoice = "" Then
                    If MsgBox("你确定不输入票据号继续打印吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                    blnValid = True
                Else
                    '检查输入有效性
                    If blnInput Then
                        If zlCommFun.ActualLen(strInvoice) <> gbyt预交 Then
                            MsgBox "输入的票据号码长度应该为 " & gbyt预交 & " 位！", vbInformation, gstrSysName
                        Else
                            blnValid = True
                        End If
                    Else
                        blnValid = True
                    End If
                End If
            Loop While Not blnValid
        Else
            Do
                '根据票据领用读取
                blnInput = False
                strInvoice = GetNextBill(lng领用ID)
                If strInvoice = "" Then
                    '如果中途换用靠后的号码,可能造成未用完,但下一号码已超出范围
                    strInvoice = UCase(InputBox("无法根据票据领用情况获取将要使用的开始票据号，" & _
                                    vbCrLf & "请你输入将要使用的开始票据号码：", gstrSysName, _
                                    "", Me.Left + 1500, Me.Top + 1500))
                    blnInput = True
                Else
                    strInvoice = UCase(InputBox("请确认重打使用的开始票据号码：", gstrSysName, _
                                    strInvoice, Me.Left + 1500, Me.Top + 1500))
                    blnInput = True
                End If
                
                '用户取消输入,不打印
                If strInvoice = "" Then Exit Sub
                
                '检查输入有效性
                If blnInput Then
                    If GetInvoiceGroupID(2, 1, lng领用ID, factProperty.lngShareUseID, strInvoice, CStr(byt预交类型)) = -3 Then
                        MsgBox "你输入的票据号码不在当前领用批次的有效领用范围内,请重新输入！", vbInformation, gstrSysName
                    Else
                        blnValid = True
                    End If
                Else
                    blnValid = True
                End If
            Loop While Not blnValid
        End If
        
        '执行数据处理
        'Zl_病人预交记录_Reprint
        strSQL = "Zl_病人预交记录_Reprint("
        '  单据号_In Varchar2,
        strSQL = strSQL & "'" & strNos & "',"
        '  票据号_In 票据使用明细.号码%Type,
        strSQL = strSQL & "'" & strInvoice & "',"
        '  领用id_In 票据使用明细.领用id%Type,
        strSQL = strSQL & "" & IIf(lng领用ID = 0, "NULL", lng领用ID) & ","
        '  使用人_In 票据使用明细.使用人%Type
        strSQL = strSQL & "'" & UserInfo.姓名 & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        '输出票据
        '78751:李南春,2014/10/20,增加预交票据打印格式
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me, _
            "NO=" & strNos, "收款时间=" & Format(str操作时间, "yyyy-mm-dd HH:MM:SS"), _
            "病人ID=" & lng病人ID, IIf(intInvoiceFormat = 0, "", "ReportFormat=" & intInvoiceFormat), 2)
        
        '更新本地票据
        If Not gblnBill预交 Then
            zlDatabase.SetPara "当前预交票据号", strInvoice, glngSys, mlngModul
        End If
                
        If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then mnuViewReFlash_Click
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function GetTurnMZToZYMultiNOs(ByVal blnRePrint As Boolean, _
    ByVal strNo As String, ByRef blnTurnMZToZY As Boolean, Optional ByVal blnNOMoved As Boolean) As String
    '功能：获取门诊转住院产生的预交单据，并返回同时转预交/一次打印的多张单据号
    '入参:strNo-需要重打NO
    '     blnRePrint-是否重打票据
    '     blnNOMoved-是否转入历史表空间
    '出参:
    '     blnTurnMZToZY-是否门诊转住院产生的
    '返回:一次打印的多张单据号，格式：A001,A002,A003,...
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strNos As String
    
    On Error GoTo errHandle
    blnTurnMZToZY = False
    strSQL = _
        "Select a.No, Max(a.记录状态) As 记录状态" & vbNewLine & _
        "From 病人预交记录 A, 病人预交记录 B" & vbNewLine & _
        "Where a.收款时间 = b.收款时间 And a.记录性质 = 1 And a.摘要 = '门诊转住院预交'" & vbNewLine & _
        "      And b.记录性质 = 1 And b.摘要 = '门诊转住院预交' And b.No = [1]" & vbNewLine & _
        "Group By a.NO" & vbNewLine & _
        "Order By NO"
    If blnNOMoved Then
        strSQL = Replace(strSQL, "病人预交记录", "H病人预交记录")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "", strNo)
    If rsTemp.EOF Then Exit Function
    
    blnTurnMZToZY = True
    If blnRePrint = False Then '补打
        If Val(zlDatabase.GetPara(283, glngSys, , "0")) = 1 Then
            With rsTemp
                Do While Not .EOF
                    If Val(NVL(!记录状态)) = 1 Then '不包含已作废单据
                        strNos = strNos & "," & NVL(rsTemp!NO)
                    End If
                    .MoveNext
                Loop
            End With
            If strNos <> "" Then strNos = Mid(strNos, 2)
        Else
            '如果不是多单据一次打印，则只补打当前单据
            strNos = strNo
        End If
        GetTurnMZToZYMultiNOs = strNos
        Exit Function
    End If
    
    '重打
    '应根据最后一次打印的情况来定
    strSQL = _
        "Select a.NO" & vbNewLine & _
        "From 票据打印内容 A" & vbNewLine & _
        "Where a.数据性质 = 2" & vbNewLine & _
        "      And a.ID In (Select ID" & vbNewLine & _
        "                From (Select b.Id" & vbNewLine & _
        "                      From 票据使用明细 A, 票据打印内容 B" & vbNewLine & _
        "                      Where a.打印id = b.Id And a.性质 = 1 And a.原因 In (1, 3) And b.数据性质 = 2 And b.No = [1]" & vbNewLine & _
        "                      Order By a.使用时间 Desc)" & vbNewLine & _
        "                Where Rownum < 2)" & vbNewLine & _
        "      And Not Exists(Select 1 From 病人预交记录 Where 记录性质 = 1 And 记录状态 = 2 And No = a.No)" & vbNewLine & _
        "Order By No"
    If blnNOMoved Then
        strSQL = Replace(strSQL, "票据打印内容", "H票据打印内容")
        strSQL = Replace(strSQL, "票据使用明细", "H票据使用明细")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "", strNo)
    If rsTemp.EOF Then Exit Function
    
    With rsTemp
        Do While Not .EOF
            strNos = strNos & "," & NVL(rsTemp!NO)
            .MoveNext
        Loop
    End With
    If strNos <> "" Then strNos = Mid(strNos, 2)
    GetTurnMZToZYMultiNOs = strNos
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub mshList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mshList.MouseRow = 0 Then
        mshList.MousePointer = 99
    Else
        mshList.MousePointer = 0
    End If
End Sub

Private Sub mshList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long
    
    lngCol = mshList.MouseCol
    
    If Button = 1 And mshList.MousePointer = 99 Then
        If mshList.TextMatrix(0, lngCol) = "" Then Exit Sub
        If mshList.TextMatrix(1, mshList.ColIndex("单据号")) = "" Then Exit Sub
        
        Set mshList.DataSource = Nothing

        mrsList.Sort = mshList.TextMatrix(0, lngCol) & IIf(mshList.ColData(lngCol) = 0, "", " DESC")
        mshList.ColData(lngCol) = (mshList.ColData(lngCol) + 1) Mod 2
        
        Call ShowBills(, True)
    End If
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

