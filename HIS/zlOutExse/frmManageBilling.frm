VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManageBilling 
   AutoRedraw      =   -1  'True
   Caption         =   "门诊记帐管理"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9675
   Icon            =   "frmManageBilling.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   Picture         =   "frmManageBilling.frx":08CA
   ScaleHeight     =   6210
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   5850
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmManageBilling.frx":0A58
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8229
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
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3722
            MinWidth        =   3722
            Picture         =   "frmManageBilling.frx":12EC
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
      TabIndex        =   5
      Top             =   0
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   9675
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinHeight1      =   720
      Width1          =   810
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbr 
         Height          =   720
         Left            =   30
         TabIndex        =   6
         Top             =   30
         Width           =   9555
         _ExtentX        =   16854
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
            NumButtons      =   17
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
               Caption         =   "记帐"
               Key             =   "Billing"
               Description     =   "记帐"
               Object.ToolTipText     =   "进入记帐窗口"
               Object.Tag             =   "记帐"
               ImageKey        =   "Billing"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "划价"
               Key             =   "Price"
               Description     =   "划价"
               Object.ToolTipText     =   "记帐划价"
               Object.Tag             =   "划价"
               ImageKey        =   "Price"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "审核"
               Key             =   "Auditing"
               Description     =   "审核"
               Object.ToolTipText     =   "记帐审核"
               Object.Tag             =   "审核"
               ImageKey        =   "Auditing"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Billing_"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修改"
               Key             =   "Modi"
               Description     =   "修改"
               Object.ToolTipText     =   "修改"
               Object.Tag             =   "修改"
               ImageKey        =   "Modi"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "销帐"
               Key             =   "Del"
               Description     =   "销帐"
               Object.ToolTipText     =   "对当前选中单据销帐"
               Object.Tag             =   "销帐"
               ImageKey        =   "Del"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Del_"
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查阅"
               Key             =   "View"
               Description     =   "查阅"
               Object.ToolTipText     =   "查阅当前单据的内容"
               Object.Tag             =   "查阅"
               ImageKey        =   "View"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "过滤"
               Key             =   "Filter"
               Description     =   "过滤"
               Object.ToolTipText     =   "按设置条件重新筛选记录"
               Object.Tag             =   "过滤"
               ImageKey        =   "Filter"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "定位"
               Key             =   "Go"
               Description     =   "定位"
               Object.ToolTipText     =   "定位到满足条件的记录上"
               Object.Tag             =   "定位"
               ImageKey        =   "Go"
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助"
               Object.ToolTipText     =   "当前帮助主题"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
      Height          =   2805
      Left            =   15
      TabIndex        =   0
      Top             =   1065
      Width           =   9660
      _ExtentX        =   17039
      _ExtentY        =   4948
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      MouseIcon       =   "frmManageBilling.frx":148A
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox picHsc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   45
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   9570
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3900
      Width           =   9570
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDetail 
      Height          =   1875
      Left            =   0
      TabIndex        =   1
      Top             =   3975
      Width           =   9660
      _ExtentX        =   17039
      _ExtentY        =   3307
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      MouseIcon       =   "frmManageBilling.frx":17A4
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   675
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":1ABE
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":1CD8
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":1EF2
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":210C
            Key             =   "View"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":2886
            Key             =   "Go"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":2AA0
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":2CBA
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":2ED4
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":30EE
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":3308
            Key             =   "Billing"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":3A02
            Key             =   "Price"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":40FC
            Key             =   "Auditing"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   90
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":47F6
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":4A10
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":4C2A
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":4E44
            Key             =   "View"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":55BE
            Key             =   "Go"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":57D8
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":59F2
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":5C0C
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":5E26
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":6040
            Key             =   "Billing"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":673A
            Key             =   "Price"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":6E34
            Key             =   "Auditing"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip tbs 
      Height          =   420
      Left            =   15
      TabIndex        =   2
      Top             =   735
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   741
      TabWidthStyle   =   2
      TabFixedWidth   =   2293
      TabFixedHeight  =   526
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "记帐单据(&1)"
            Key             =   "Auditing"
            Object.ToolTipText     =   "显示直接记帐或划价后审核了的记帐单据"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "划价单据(&2)"
            Key             =   "Price"
            Object.ToolTipText     =   "显示划价后未审核的记帐单据"
            ImageVarType    =   2
         EndProperty
      EndProperty
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
      Begin VB.Menu mnuEdit_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileLocalSet 
         Caption         =   "参数设置(&R)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFileLocalSet_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_quit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEdit_Billing 
         Caption         =   "门诊记帐(&B)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditCust 
         Caption         =   "自定义记帐单(&U)"
         Begin VB.Menu mnuEditCustBill 
            Caption         =   "(空)"
            Index           =   1
         End
      End
      Begin VB.Menu mnuEditPrice 
         Caption         =   "记帐划价(&P)"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuEditAuditing 
         Caption         =   "记帐审核(&A)"
         Shortcut        =   +{F2}
      End
      Begin VB.Menu mnuEditAuditingPati 
         Caption         =   "按病人审核(&N)"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuEditBilling_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Modi 
         Caption         =   "修改单据(&M)"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuEdit_Adjust 
         Caption         =   "调整时间(&J)"
         Shortcut        =   ^J
      End
      Begin VB.Menu mnuEdit_Adjust_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Del 
         Caption         =   "单据销帐(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEdit_Del_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_View 
         Caption         =   "查阅单据(&V)"
      End
      Begin VB.Menu mnuEdit_Print 
         Caption         =   "打印单据(&I)"
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
      Begin VB.Menu mnuView_5 
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
      Begin VB.Menu mnuView_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefeshOption 
         Caption         =   "刷新方式(&O)"
         Begin VB.Menu mnuViewRefeshOptionItem 
            Caption         =   "操作后不要刷新数据(&1)"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewRefeshOptionItem 
            Caption         =   "操作后提示是否刷新(&2)"
            Index           =   1
         End
         Begin VB.Menu mnuViewRefeshOptionItem 
            Caption         =   "操作后自动刷新数据(&3)"
            Index           =   2
         End
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
Attribute VB_Name = "frmManageBilling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mrsList As ADODB.Recordset  '单据列表
Private mrsTotal As ADODB.Recordset
Private mrsDetail As ADODB.Recordset
Private Type Type_SQLCondition
    Default As Boolean          '是否是缺省进入，此时没有条件值,缺省值在mstrFilter中
    DateB As Date
    DateE As Date
    NOB As String
    NOE As String
    Patient As String
    PatientIdentity As String
    DeptID As Long
    Operator As String
    PatientNo As String '问题号:38539
    PatientID As Long '问题号:38539
End Type
Private SQLCondition As Type_SQLCondition
Private mstrFilter As String
Private mbln记帐 As Boolean, mbln销帐 As Boolean
Private mstr操作员 As String
Private mstrPage As String, mblnMax As Boolean
Private mblnGo As Boolean, mlngGo As Long
Private mlngCurRow As Long, mlngTopRow As Long
Private mstrPrivs As String   '保存当前的授权功能
Private mlngModul As Long
Private mblnNOMoved As Boolean '显明细时记录当前选择的单据是否在在线数据表中,以其它操作时无需再判断
'消息相关对象变量
Private WithEvents mobjMsgModule As clsMipModule
Attribute mobjMsgModule.VB_VarHelpID = -1

Private Sub cbr_Resize()
    Form_Resize
End Sub

Private Sub Form_Activate()
    Call InitLocPar(mlngModul)
    Call mshList_GotFocus
End Sub
Private Sub mnuEdit_Adjust_Click()
    Dim strNo As String
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
    If strNo = "" Then
        MsgBox "当前没有单据可以调整！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '是否已转入后备数据表中
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNo, 2, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '此时已转入在线数据表
    End If
    
    '已经冲销过(部分)的单据不允许调整
    If BillExistDelete(strNo, 2) Then
        MsgBox "该单据包含已销帐内容,不允许调整！", vbInformation, gstrSysName
        Exit Sub
    End If

    '已结帐单据根据参数处理
    If HaveBilling(1, strNo) Then
        Select Case gbytBillOpt
            Case 0
            Case 1
                If MsgBox("该记帐单据已经结帐,要调整吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            Case 2
                MsgBox "该记帐单据已经结帐,不能调整！", vbExclamation, gstrSysName: Exit Sub
        End Select
    End If

    On Error Resume Next
    Err.Clear

    '显示单据内容
    Dim lng记帐ID As Long
    Dim varTemp As Variant
    
    lng记帐ID = mshList.TextMatrix(mshList.Row, GetColNum("记帐单ID"))
    
    If lng记帐ID = 0 Or gobjCustBill Is Nothing Then
        frmCharge.mlngModul = mlngModul
        frmCharge.mstrPrivs = mstrPrivs
        frmCharge.mbytInFun = 2
        frmCharge.mbytInState = 2
        frmCharge.mstrInNO = strNo
        Set frmCharge.mobjMsgModule = mobjMsgModule
        frmCharge.Show 1, Me
    Else
        '记帐ID、bytUseType、bytInState、strInNO、lngUnitID、lngDeptID、lng病人ID、mstrPrivs
        varTemp = Array(lng记帐ID, 3, 2, strNo, 0, 0, 0, mstrPrivs)
        gobjCustBill.CodeMan glngSys, -1, gcnOracle, Me, gstrDBUser, varTemp
        
        gblnOK = varTemp
    End If
End Sub

Private Sub mnuEdit_Modi_Click()
    Dim strNo As String
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
    
    If strNo = "" Then
        MsgBox "当前没有单据可以修改！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '是否已转入后备数据表中
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNo, 2, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '此时已转入在线数据表
    End If
    gstrModiNO = ""
    
    On Error Resume Next
    Err.Clear
    
    Dim lng记帐ID As Long
    Dim varTemp As Variant
    
    lng记帐ID = mshList.TextMatrix(mshList.Row, GetColNum("记帐单ID"))
    
    If lng记帐ID = 0 Or gobjCustBill Is Nothing Then
        frmCharge.mlngModul = mlngModul
        frmCharge.mstrPrivs = mstrPrivs
        frmCharge.mbytInFun = 2
        frmCharge.mstrInNO = strNo
        frmCharge.mbytInState = 0
        frmCharge.mbytBilling = IIf(tbs.SelectedItem.Key = "Auditing", 0, 1)
        Set frmCharge.mobjMsgModule = mobjMsgModule
        frmCharge.Show 1, Me
    Else
        '记帐ID、bytUseType、bytInState、strInNO、lngUnitID、lngDeptID、lng病人ID、mstrPrivs
        varTemp = Array(lng记帐ID, 3, 0, strNo, 0, 0, 0, mstrPrivs)
        gobjCustBill.CodeMan glngSys, -1, gcnOracle, Me, gstrDBUser, varTemp
        
        gblnOK = varTemp
    End If
    
    If gblnOK Then
        If gstrModiNO <> "" Then
            If mnuViewRefeshOptionItem(1).Checked Then
                If MsgBox("当前操作已更改单据清单内容,修改后的单据号为:[" & gstrModiNO & "],要刷新吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    mnuViewReFlash_Click
                End If
            ElseIf mnuViewRefeshOptionItem(2).Checked Then
                mnuViewReFlash_Click
            End If
        Else
            If mnuViewRefeshOptionItem(1).Checked Then
                If MsgBox("当前操作已更改单据清单内容,要刷新吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    mnuViewReFlash_Click
                End If
            ElseIf mnuViewRefeshOptionItem(2).Checked Then
                mnuViewReFlash_Click
            End If
        End If
    End If
End Sub

Private Sub mnuEdit_Print_Click()
    Dim strNo As String, strTime As String
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
    If strNo = "" Then
        MsgBox "当前没有单据可以打印！", vbInformation, gstrSysName
        Exit Sub
    End If
    If InStr(",0,1,", Val(mshList.TextMatrix(mshList.Row, GetColNum("符号")))) = 0 Then
        MsgBox "该单据为销帐单据或已被销帐，不能再打印！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '是否已转入后备数据表中
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNo, 2, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '此时已转入在线数据表
    End If
    
    If ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1122", Me) Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1122", Me, "NO=" & strNo, "药品单位=" & IIf(gbln药房单位, 1, 0), 2)
    End If
End Sub

Private Sub mnuEditAuditing_Click()
    On Error Resume Next
    frmCharge.mlngModul = mlngModul
    frmCharge.mstrPrivs = mstrPrivs
    frmCharge.mbytInFun = 2
    frmCharge.mbytInState = 0
    frmCharge.mbytBilling = 2
    Set frmCharge.mobjMsgModule = mobjMsgModule
    frmCharge.Show 1, Me
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewReFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewReFlash_Click
        End If
    End If
End Sub
Private Sub mnuEditAuditingPati_Click()
    On Error Resume Next
    If Not frmBillingAuditing.zlShowCard(Me, mlngModul, mstrPrivs) Then Exit Sub
    If mnuViewRefeshOptionItem(1).Checked Then
        If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            mnuViewReFlash_Click
        End If
    ElseIf mnuViewRefeshOptionItem(2).Checked Then
        mnuViewReFlash_Click
    End If
End Sub

Private Sub mnuEditPrice_Click()
    On Error Resume Next
    Err.Clear
    frmCharge.mlngModul = mlngModul
    frmCharge.mstrPrivs = mstrPrivs
    frmCharge.mbytInFun = 2
    frmCharge.mbytInState = 0
    frmCharge.mbytBilling = 1
    Set frmCharge.mobjMsgModule = mobjMsgModule
    frmCharge.Show 1, Me
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewReFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewReFlash_Click
        End If
    End If
End Sub

Private Sub mnuFileLocalSet_Click()
    Dim blnPre As Boolean
    
    blnPre = gbln药房单位
    
    With frmSetExpence
        .mlngModul = mlngModul
        .mstrPrivs = mstrPrivs
        .mbytInFun = 2
        .mblnSetDrugStore = False
        .Show 1, Me
    End With
        
    '更改了药品单位参数,重新刷新
    If gbln药房单位 <> blnPre Then
        ShowBills mstrFilter
    End If
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Dim strNo As String
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
    If strNo <> "" Then
        With mshList
            Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
                    "NO=" & .TextMatrix(.Row, GetColNum("单据号")), "病人ID=" & .TextMatrix(.Row, GetColNum("病人ID")), _
                    "开单人=" & .TextMatrix(.Row, GetColNum("开单人")))
        End With
    Else
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me)
    End If
End Sub

Private Sub mnuViewFilter_Click()
    With frmBillingFilter
        .mstrPrivs = mstrPrivs
        .chk记帐.Enabled = tbs.SelectedItem.Key = "Auditing"
        .chk销帐.Enabled = tbs.SelectedItem.Key = "Auditing"
        .Show 1, Me
        If gblnOK Then
            mstrFilter = .mstrFilter
            mbln记帐 = .chk记帐.Value = 1
            mbln销帐 = .chk销帐.Value = 1
            If .cbo操作员.Text <> "所有操作员" Then
                mstr操作员 = zlStr.NeedName(.cbo操作员.Text)
            Else
                mstr操作员 = ""
            End If
            SQLCondition.DateB = .dtpBegin.Value
            SQLCondition.DateE = .dtpEnd.Value
            SQLCondition.NOB = .txtNOBegin.Text
            SQLCondition.NOE = .txtNoEnd.Text
            SQLCondition.Patient = gstrLike & UCase(.txt姓名.Text) & "%"
            SQLCondition.PatientIdentity = .txt病人ID.Text
            SQLCondition.DeptID = .cbo科室.ItemData(.cbo科室.ListIndex)
            SQLCondition.Operator = mstr操作员
            '问题号:38539
            SQLCondition.PatientNo = .txtPatientNo
            SQLCondition.PatientID = .mlngPrePatient
            
            mnuViewReFlash_Click
        End If
    End With
End Sub

Private Sub mnuViewRefeshOptionItem_Click(Index As Integer)
    Dim i As Long
    For i = 0 To mnuViewRefeshOptionItem.UBound
        mnuViewRefeshOptionItem(i).Checked = i = Index
    Next
End Sub

Private Sub mshDetail_EnterCell()
    mshDetail.ForeColorSel = mshDetail.CellForeColor
End Sub

Private Sub mshDetail_GotFocus()
    Call SetActiveList(mshDetail)
End Sub

Private Sub mshDetail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mshDetail.MouseRow = 0 Then
        mshDetail.MousePointer = 99
    Else
        mshDetail.MousePointer = 0
    End If
End Sub

Private Sub mshDetail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long, strTime As String, blnDel As Boolean
    
    lngCol = mshDetail.MouseCol
    
    If Button = 1 And mshDetail.MousePointer = 99 Then
        If mshDetail.TextMatrix(0, lngCol) = "" Then Exit Sub
        If mshDetail.TextMatrix(1, 0) = "" Then Exit Sub
        If mrsDetail Is Nothing Then Exit Sub
        
        strTime = mshList.TextMatrix(mshList.Row, GetColNum("登记时间"))
        blnDel = Val(mshList.TextMatrix(mshList.Row, GetColNum("符号"))) = 2
        
        Set mshDetail.DataSource = Nothing

        mrsDetail.Sort = mshDetail.TextMatrix(0, lngCol) & IIf(mshDetail.ColData(lngCol) = 0, "", " DESC")
        mshDetail.ColData(lngCol) = (mshDetail.ColData(lngCol) + 1) Mod 2
        
        Call ShowDetail(, strTime, blnDel, True)
    End If
End Sub

Private Sub mshList_DblClick()
    If mshList.MouseRow = 0 Then Exit Sub
    If mnuEdit_View.Enabled Then mnuEdit_View_Click
End Sub

Private Sub mshList_EnterCell()
    Dim strNo As String, strTime As String, blnDel As Boolean
        
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
    
    If mshList.Row = 0 Or strNo = "" Then Exit Sub
    
    stbThis.Panels(2).Text = "共 " & NVL(mrsTotal!单据, 0) & " 张单据,合计:" & Format(NVL(mrsTotal!金额, 0), gstrDec)
    
    mlngGo = mshList.Row
    mlngCurRow = mshList.Row: mlngTopRow = mshList.TopRow
    
    strTime = mshList.TextMatrix(mshList.Row, GetColNum("登记时间"))
    blnDel = Val(mshList.TextMatrix(mshList.Row, GetColNum("符号"))) = 2
    
    mnuEdit_Adjust.Enabled = Not blnDel
    mnuEdit_Modi.Enabled = Not blnDel
    mnuEdit_Del.Enabled = Not blnDel
    tbr.Buttons("Modi").Enabled = Not blnDel
    tbr.Buttons("Del").Enabled = Not blnDel
        
    mshList.ForeColorSel = mshList.CellForeColor
    
    Call ShowDetail(strNo, strTime, blnDel)
End Sub

Private Sub mshList_GotFocus()
    Call SetActiveList(mshList)
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
        Case vbKeyReturn
            If mnuEdit_View.Enabled Then mnuEdit_View_Click
        Case vbKeyEscape
            mblnGo = False
    End Select
End Sub

Private Sub mnuEdit_Del_Click()
    Dim strNo As String, strTime As String
    Dim strSQL As String, i As Long, blnFlagPrint As Boolean
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
    
    If strNo = "" Then
        MsgBox "当前没有单据可以退费！", vbInformation, gstrSysName
        Exit Sub
    End If
    strTime = mshList.TextMatrix(mshList.Row, GetColNum("登记时间"))
    
    '可以在所有的检查完后,再决定是否导入到在线表,但所有检查时要把mblnNOMoved传入进行判断,为简化,暂定为先检查是否已转出
    '是否已转入后备数据表中
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNo, 2, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '此时已转入在线数据表
    End If
    
    '销帐权限检查
    If Not BillOperCheck(4, mshList.TextMatrix(mshList.Row, GetColNum("操作员")), CDate(strTime), "销帐", strNo, , 2) Then Exit Sub
    
    '是否已执行
    i = BillCanDelete(strNo, 2, , strTime, blnFlagPrint)
    If i <> 0 Then
        Select Case i
            Case 1 '该单据不存在
                MsgBox "指定的单据不存在！", vbInformation, gstrSysName
            Case 2 '已经全部完全执行
                MsgBox "该单据中的项目已经全部完全执行！", vbInformation, gstrSysName
            Case 3 '未完全执行部分剩余数量为0
                MsgBox "该单据中未完全执行部分项目剩余数量为零,没有可以销帐的费用！", vbInformation, gstrSysName
        End Select
        Exit Sub
    End If
    If blnFlagPrint Then
        If MsgBox("注意:检验医嘱的条码已打印，是否继续？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    
    '已结帐单据根据参数处理
    If HaveBilling(1, strNo, True, strTime) Then
        Select Case gbytBillOpt
            Case 1
                If MsgBox("该单据已经结帐,要销帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            Case 2
                MsgBox "该单据已经结帐,不能销帐！", vbExclamation, gstrSysName: Exit Sub
        End Select
    End If
    
    On Error Resume Next
    Err.Clear
    
    '显示单据内容
    Dim lng记帐ID As Long, varTemp As Variant
    
    lng记帐ID = mshList.TextMatrix(mshList.Row, GetColNum("记帐单ID"))
    
    If lng记帐ID = 0 Or gobjCustBill Is Nothing Then
        frmCharge.mlngModul = mlngModul
        frmCharge.mstrPrivs = mstrPrivs
        frmCharge.mbytInFun = 2
        frmCharge.mbytInState = 3
        frmCharge.mstrInNO = strNo
        frmCharge.mstrTime = strTime
        frmCharge.mbytBilling = IIf(tbs.SelectedItem.Key = "Auditing", 0, 1)
        Set frmCharge.mobjMsgModule = mobjMsgModule
        frmCharge.Show 1, Me
    Else
        '记帐ID、bytUseType、bytInState、strInNO、lngUnitID、lngDeptID、lng病人ID、mstrPrivs
        varTemp = Array(lng记帐ID, 3, 3, strNo, 0, 0, 0, mstrPrivs)
        gobjCustBill.CodeMan glngSys, -1, gcnOracle, Me, gstrDBUser, varTemp
        gblnOK = varTemp
    End If
    
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewReFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewReFlash_Click
        End If
    End If
End Sub

Private Sub mnuHelpTitle_Click()
ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub mnuEdit_Billing_Click()
    On Error Resume Next
    Err.Clear
    frmCharge.mlngModul = mlngModul
    frmCharge.mstrPrivs = mstrPrivs
    frmCharge.mbytInFun = 2
    frmCharge.mbytInState = 0
    frmCharge.mbytBilling = 0
    Set frmCharge.mobjMsgModule = mobjMsgModule
    frmCharge.Show 1, Me
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewReFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewReFlash_Click
        End If
    End If
End Sub

Private Sub mnuEditCustBill_Click(Index As Integer)
    '自定义记帐
    Dim varTemp As Variant
    '参数含义依次是：
    '记帐ID、bytUseType、bytInState、strInNO、lngUnitID、lngDeptID、lng病人ID、mstrPrivs、blnViewCancel
    
    varTemp = Array(mnuEditCustBill(Index).Tag, 3, 0, "", 0, 0, 0, mstrPrivs)
    gobjCustBill.CodeMan glngSys, -1, gcnOracle, Me, gstrDBUser, varTemp
    
    gblnOK = varTemp '返回值
    
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewReFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewReFlash_Click
        End If
    End If
End Sub

Private Sub mnuEdit_View_Click()
    Dim strNo As String, strTime As String, blnDel As Boolean
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
    If strNo = "" Then
        MsgBox "当前没有单据可以查阅！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    strTime = mshList.TextMatrix(mshList.Row, GetColNum("登记时间"))
    blnDel = Val(mshList.TextMatrix(mshList.Row, GetColNum("符号"))) = 2
    
    On Error Resume Next
    Err.Clear
    
    '显示单据内容
    Dim lng记帐ID As Long
    Dim varTemp As Variant
    lng记帐ID = mshList.TextMatrix(mshList.Row, GetColNum("记帐单ID"))
    
    If lng记帐ID = 0 Or gobjCustBill Is Nothing Then
        frmCharge.mlngModul = mlngModul
        frmCharge.mstrPrivs = mstrPrivs
        frmCharge.mbytInFun = 2
        frmCharge.mbytInState = 1
        frmCharge.mstrTime = strTime
        frmCharge.mblnDelete = blnDel
        frmCharge.mstrInNO = strNo
        frmCharge.mblnNOMoved = mblnNOMoved
        frmCharge.mbytBilling = IIf(tbs.SelectedItem.Key = "Auditing", 0, 1)
        Set frmCharge.mobjMsgModule = mobjMsgModule
        frmCharge.Show 1, Me
    Else
        '记帐ID、bytUseType、bytInState、strInNO、lngUnitID、lngDeptID、lng病人ID、mstrPrivs
        varTemp = Array(lng记帐ID, 3, 1, strNo, 0, 0, 0, mstrPrivs, blnDel)
        gobjCustBill.CodeMan glngSys, -1, gcnOracle, Me, gstrDBUser, varTemp
        
        gblnOK = varTemp
    End If
End Sub

Private Sub mnuFile_quit_Click()
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
    Dim i As Long
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For i = 1 To tbr.Buttons.Count
        tbr.Buttons(i).Caption = IIf(mnuViewToolText.Checked, tbr.Buttons(i).Tag, "")
    Next
    cbr.Bands(1).MinHeight = tbr.ButtonHeight
    Form_Resize
End Sub

Private Sub picHsc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If mshList.Height + Y < 1000 Or mshDetail.Height - Y < 1000 Then Exit Sub
        picHsc.Top = picHsc.Top + Y
        mshList.Height = mshList.Height + Y
        mshDetail.Top = mshDetail.Top + Y
        mshDetail.Height = mshDetail.Height - Y
        Me.Refresh
    End If
End Sub

Private Sub picHsc_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then mshList.SetFocus
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Quit"
            mnuFile_quit_Click
        Case "Go" '定位
            mnuViewGo_Click
        Case "Filter" '过滤
            mnuViewFilter_Click
        Case "View"
            mnuEdit_View_Click
        Case "Billing"
            mnuEdit_Billing_Click
        Case "Price"
            mnuEditPrice_Click
        Case "Auditing"
            mnuEditAuditing_Click
        Case "Modi"
            mnuEdit_Modi_Click
        Case "Del"
            mnuEdit_Del_Click
        Case "Print"
            mnuFile_Print_Click
        Case "Preview"
            mnuFile_PreView_Click
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
    objOut.Title.Text = "门诊记帐单据清单"
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '表项
    With frmBillingFilter
        objRow.Add "时间：" & Format(.dtpBegin.Value, .dtpBegin.CustomFormat) & " 至 " & Format(.dtpEnd.Value, .dtpEnd.CustomFormat)
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
    mshList.Col = 0: mshList.ColSel = mshList.COLS - 1
    mshList.Redraw = True
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage hWnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo hWnd
End Sub

Private Sub SetMenu(blnUsed As Boolean)
'功能：根据有无记录设置菜单可用状态
    mnuFile_Print.Enabled = blnUsed
    mnuFile_PreView.Enabled = blnUsed
    mnuFile_Excel.Enabled = blnUsed
    tbr.Buttons("Print").Enabled = blnUsed
    tbr.Buttons("Preview").Enabled = blnUsed
    
    mnuEdit_Adjust.Enabled = blnUsed
    mnuEdit_Modi.Enabled = blnUsed
    tbr.Buttons("Modi").Enabled = blnUsed
    
    mnuEdit_Del.Enabled = blnUsed
    mnuEdit_View.Enabled = blnUsed
    tbr.Buttons("Del").Enabled = blnUsed
    tbr.Buttons("View").Enabled = blnUsed
    
    mnuViewGo.Enabled = blnUsed
    tbr.Buttons("Go").Enabled = blnUsed
End Sub

Private Sub SetCustBill()
'设置与自定义记帐单相关的内容
    Dim rsTmp As New ADODB.Recordset
    Dim lngCount As Long, lngSum As Long
    
    On Error Resume Next
    
    mstrPrivs = mstrPrivs
    
    If gobjCustBill Is Nothing Then
        Set gobjCustBill = CreateObject("zl9CustAcc.clsCustAcc")
    End If
    If InStr(mstrPrivs, "专项记帐") = 0 Then
        mnuEditCust.Visible = False
        Exit Sub
    End If

    On Error GoTo errHandle
    

    
    '如果创建成功，再读出对应的菜单
    If Not gobjCustBill Is Nothing Then
        gstrSQL = "Select ID,名称 From 收费记帐单 Where substr(适用范围,1,1)='1' Order by 编号"
        Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
        lngSum = rsTmp.RecordCount
    End If
    
    If lngSum > 0 Then
        For lngCount = 1 To lngSum
            '增加到主菜单中
            If lngCount > 1 Then
                Load mnuEditCustBill(lngCount)
            End If
            mnuEditCustBill(lngCount).Caption = rsTmp("名称") & "(&" & lngCount & ")"
            mnuEditCustBill(lngCount).Tag = rsTmp("ID")
            
            rsTmp.MoveNext
        Next
    Else
        mnuEditCustBill(1).Enabled = False
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Load()
    Dim i As Long
    
    mstrPrivs = gstrPrivs
    mlngModul = glngModul
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
    
    Call SetCustBill
    Call RestoreWinState(Me, App.ProductName)
    Set stbThis.Panels(5).Picture = Me.Picture
    '刷新方式
    For i = 0 To mnuViewRefeshOptionItem.UBound
        If i = Val(zlDatabase.GetPara("刷新方式", glngSys, mlngModul, 2)) Then
            mnuViewRefeshOptionItem(i).Checked = True
        Else
            mnuViewRefeshOptionItem(i).Checked = False
        End If
    Next
    
    i = IIf(zlDatabase.GetPara("页面", glngSys, mlngModul, "1") = "1", 1, 2)
    tbs.Tabs(i).Selected = True
    
    mlngCurRow = 1: mlngTopRow = 1
    mstrPage = ""
    
    '权限设置
    If InStr(mstrPrivs, "门诊记帐") = 0 Then
        mnuEdit_Billing.Visible = False
        tbr.Buttons("Billing").Visible = False
        
        mnuEditCust.Visible = False
    End If
    If InStr(mstrPrivs, "记帐划价") = 0 Then
        mnuEditPrice.Visible = False
        tbr.Buttons("Price").Visible = False
    End If
    If InStr(mstrPrivs, "记帐审核") = 0 Then
        mnuEditAuditing.Visible = False
        mnuEditAuditingPati.Visible = False
        tbr.Buttons("Auditing").Visible = False
    End If
    If InStr(mstrPrivs, "门诊记帐") = 0 _
        And InStr(mstrPrivs, "记帐划价") = 0 _
        And InStr(mstrPrivs, "记帐审核") = 0 Then
        mnuEditBilling_.Visible = False
        tbr.Buttons("Billing_").Visible = False
    End If
    
    If InStr(mstrPrivs, "记录修改") = 0 Then
        mnuEdit_Modi.Visible = False
        tbr.Buttons("Modi").Visible = False
    End If
    If InStr(mstrPrivs, "记录调整") = 0 Then
        mnuEdit_Adjust.Visible = False
    End If
    If InStr(mstrPrivs, "记录修改") = 0 And InStr(mstrPrivs, "记录调整") = 0 Then
        mnuEdit_Adjust_.Visible = False
    End If
    
    If InStr(mstrPrivs, "门诊销帐") = 0 Then
        mnuEdit_Del.Visible = False
        mnuEdit_Del_.Visible = False
        tbr.Buttons("Del").Visible = False
        tbr.Buttons("Del_").Visible = False
    End If


    mbln记帐 = True
    mbln销帐 = False
    mstr操作员 = UserInfo.姓名
    
    Call SetHeader
    Call SetDetail
    Call SetMenu(False)
    
    stbThis.Panels(2).Text = "请刷新清单或重新设置过滤条件"
    
    '初始化消息处理对象模块
    Call zlMsgModuleInit
    
End Sub

Private Sub Form_Resize()
    Dim cbrH As Long, staH As Long, sngVsc As Single

    On Error Resume Next
    
    If WindowState = 1 Then Exit Sub
    
    mshList.MousePointer = 0
    
    '靠齐控件宽度和高度
    cbrH = IIf(cbr.Visible, cbr.Height, 0)
    staH = IIf(stbThis.Visible, stbThis.Height, 0)
    sngVsc = mshDetail.Height / (mshDetail.Height + mshList.Height)
    
    If mblnMax Then
        sngVsc = 0.3: mblnMax = False
    End If
    If Me.WindowState = 2 Then mblnMax = True
    
    tbs.Left = Me.ScaleLeft
    tbs.Top = Me.ScaleTop + cbrH + 15
    
    mshList.Left = 0
    mshList.Top = tbs.Top + tbs.TabFixedHeight + 30
    mshList.Width = Me.ScaleWidth
    mshList.Height = (Me.ScaleHeight - cbrH - staH - (tbs.TabFixedHeight + 45) - picHsc.Height) * (1 - sngVsc)
    
    picHsc.Top = mshList.Top + mshList.Height
    picHsc.Left = Me.ScaleLeft
    picHsc.Width = Me.ScaleWidth
    
    mshDetail.Left = Me.ScaleLeft
    mshDetail.Top = picHsc.Top + picHsc.Height
    mshDetail.Width = Me.ScaleWidth
    mshDetail.Height = Me.ScaleHeight - cbrH - staH - (tbs.TabFixedHeight + 45) - picHsc.Height - mshList.Height
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    
    mstrFilter = ""
    Unload frmBillingFilter
    Unload frmBillingGo
    
    Call SaveWinState(Me, App.ProductName)
    zlDatabase.SetPara "页面", tbs.SelectedItem.Index, glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
    '刷新方式
    For i = 0 To mnuViewRefeshOptionItem.UBound
        If mnuViewRefeshOptionItem(i).Checked Then
            zlDatabase.SetPara "刷新方式", i, glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
            Exit For
        End If
    Next
    '拆卸消息对象
    Call zlMsgModuleUnload
End Sub

Private Sub mnuViewGo_Click()
    frmBillingGo.Show 1, Me
    If gblnOK Then Call SeekBill(frmBillingGo.optHead)
End Sub

Private Sub SeekBill(blnHead As Boolean)
    Dim i As Long, bln As Boolean, intRows As Integer
    Dim blnFill As Boolean, j As Long
    
    Screen.MousePointer = 11
    mblnGo = True
    stbThis.Panels(2).Text = "正在定位满足条件的单据,按ESC终止 ..."
    Me.Refresh
    
    For i = IIf(blnHead, 1, mlngGo) To mshList.Rows - 1
        DoEvents
            
        '比较条件
        blnFill = True
        With frmBillingGo
            If .txtNO.Text <> "" Then
                blnFill = blnFill And mshList.TextMatrix(i, GetColNum("单据号")) = .txtNO.Text
            End If
            If .txt病人ID.Text <> "" Then
                blnFill = blnFill And mshList.TextMatrix(i, GetColNum("病人ID")) = .txt病人ID.Text
            End If
            If .txt姓名.Text <> "" Then
                blnFill = blnFill And UCase(mshList.TextMatrix(i, GetColNum("姓名"))) Like "*" & UCase(.txt姓名.Text) & "*"
            End If
        End With
    
        '满足则退出
        If blnFill Then
            mshList.Row = i: mshList.TopRow = i
            mshList.Col = 0: mshList.ColSel = mshList.COLS - 1
            
            Call mshList_EnterCell
            mlngGo = i + 1
            
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

Private Function GetColNum(strHead As String) As Integer
    Dim i As Long
    For i = 0 To mshList.COLS - 1
        If mshList.TextMatrix(0, i) = strHead Then GetColNum = i: Exit Function
    Next
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
        If mshList.TextMatrix(1, GetColNum("单据号")) = "" Then Exit Sub
        If mrsList Is Nothing Then Exit Sub
        
        Set mshList.DataSource = Nothing

        mrsList.Sort = mshList.TextMatrix(0, lngCol) & IIf(mshList.ColData(lngCol) = 0, "", " DESC")
        mshList.ColData(lngCol) = (mshList.ColData(lngCol) + 1) Mod 2
        
        Call ShowBills(, True)
    End If
End Sub

Private Sub SetHeader()
    Dim strHead As String
    Dim i As Long
    
    strHead = "单据号,1,850|开单科室,1,850|开单人,1,800|病人ID,1,750|门诊号,1,900|姓名,1,700|费别,1,900|应收金额,7,850|实收金额,7,850|操作员,1,800|登记时间,1,1850|说明,1,850|符号,1,0|记帐单ID,1,0"
    With mshList
        .COLS = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(mshList, App.ProductName & "\" & Me.Name)
                
         '医生列
        i = GetColNum("开单人")
        If InStr(mstrPrivs, "医生查询") = 0 Then
            .ColWidth(i) = 0
        ElseIf mshList.ColWidth(i) = 0 Then
            .ColWidth(i) = 800
        End If
                
        i = GetColNum("符号"): mshList.ColWidth(i) = 0
        i = GetColNum("记帐单ID"): mshList.ColWidth(i) = 0
        
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
        
        .Col = 0: .ColSel = .COLS - 1
                
        Call mshList_EnterCell
    End With
End Sub

Private Sub SetDetail()
    Dim strHead As String
    Dim i As Long
    
    strHead = "类别,1,750|名称,1,1800" & IIf(gTy_System_Para.byt药品名称显示 = 2, "|商品名,1,2000", "") & "|规格,1,1000|单位,4,500|数量,7,850|单价,7,850|应收金额,7,850|实收金额,7,850|执行科室,1,850|类型,1,850|说明,1,1000|记录状态,1,0"
    
    With mshDetail
        .Redraw = False
        
        .COLS = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(mshDetail, App.ProductName & "\" & Me.Name)
        For i = 0 To .COLS - 1
            If .TextMatrix(0, i) = "商品名" Then
                If gTy_System_Para.byt药品名称显示 = 2 Then
                    If .ColWidth(i) <= 0 Then .ColWidth(i) = 2000
                Else
                    .ColWidth(i) = 0
                End If
            End If
        Next
        
        .RowHeight(0) = 320
        .ColWidth(.COLS - 1) = 0
        
        .Row = 1: .Col = 0: .ColSel = .COLS - 1
        
        Call mshDetail_EnterCell
        
        .Redraw = True
    End With
End Sub

Private Sub ShowBills(Optional ByVal strFilter As String, Optional blnSort As Boolean)
'功能:按条件读取单据列表(过滤功能)
'参数:strFilter=以"AND"开始的条件串
'     blnSort=不重新读取数据,仅重新显示已排序的内容
    Dim i As Long, j As Long, k As Long
    Dim strSQL As String
    
    On Error GoTo errH
    
    If Not blnSort Then
        Call zlCommFun.ShowFlash("正在读取单据列表,请稍候 ...", Me)
        DoEvents
        Me.Refresh
        
        SQLCondition.Default = (strFilter = "")
        If strFilter = "" Then
            '缺省过滤条件(一天内)
            strFilter = " And 登记时间 Between Trunc(Sysdate) And Trunc(Sysdate+1)-1/24/60/60"
        End If
            
        '操作员单独控制
        If mstr操作员 <> "" Then
            If tbs.SelectedItem.Key = "Auditing" Then
                strFilter = strFilter & " And 操作员姓名||''=[10]"
            Else
                strFilter = strFilter & " And 划价人||''=[10]"
            End If
        End If
        
        strFilter = " Where 记录性质=2 And 门诊标志 in(1,4) " & strFilter
        
        If frmBillingFilter.mblnDateMoved And tbs.SelectedItem.Key = "Auditing" Then  '筛选时的时间在最后一次转出之前,且当前列表不是划价单
            strFilter = zlGetFullFieldsTable("门诊费用记录", 2, strFilter, False)
        Else
            strFilter = zlGetFullFieldsTable("门诊费用记录", 0, strFilter, False)
        End If
        
        '单据号,开单科室,开单人,病人ID,门诊号,姓名,费别,应收金额,实收金额,操作员,登记时间,说明,符号,记帐单ID
        If tbs.SelectedItem.Key = "Auditing" Then
            '记帐单状态
            If mbln记帐 And mbln销帐 Then
                strFilter = strFilter & " And 记录状态 IN([11],[12],[13])"
            ElseIf mbln记帐 Then
                strFilter = strFilter & " And 记录状态 IN([11],[13])"
            ElseIf mbln销帐 Then
                strFilter = strFilter & " And 记录状态=[12]"
            End If
            
            strFilter = strFilter & " And 记录状态<>0 And 操作员姓名 IS NOT NULL"
            
            strSQL = _
                "Select A.NO as 单据号,B.名称 as 开单科室,A.开单人,A.病人ID,A.标识号 as 门诊号,A.姓名,A.费别," & _
                " To_Char(Sum(Decode(A.记录状态,2,-1,1)*A.应收金额),'9999999" & gstrDec & "') as 应收金额," & _
                " To_Char(Sum(Decode(A.记录状态,2,-1,1)*A.实收金额),'9999999" & gstrDec & "') as 实收金额," & _
                " A.操作员姓名 as 操作员,To_Char(A.登记时间,'YYYY-MM-DD HH24:MI:SS') as 登记时间," & _
                " Decode(Max(A.记录状态),2,'销帐记录','记帐记录') as 说明,Max(A.记录状态) as 符号,A.记帐单ID" & _
                " From (" & strFilter & ") A,部门表 B" & _
                " Where A.开单部门ID = B.ID" & _
                " And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & vbNewLine & _
                " Group by A.NO,B.名称,A.开单人,A.病人ID,A.标识号,A.姓名,A.费别,A.操作员姓名,A.登记时间,A.记帐单ID" & _
                " Order by A.登记时间 Desc,A.NO Desc"
        Else
            '记帐划价单状态,即未审核状态
            strFilter = strFilter & " And 记录状态=[14] And 操作员姓名 IS NULL And 划价人 is Not NULL"
            
            strSQL = _
                "Select A.NO as 单据号,B.名称 as 开单科室,A.开单人,A.病人ID,A.标识号 as 门诊号,A.姓名,A.费别," & _
                " To_Char(Sum(Decode(A.记录状态,2,-1,1)*A.应收金额),'9999999" & gstrDec & "') as 应收金额," & _
                " To_Char(Sum(Decode(A.记录状态,2,-1,1)*A.实收金额),'9999999" & gstrDec & "') as 实收金额," & _
                " A.划价人 as 操作员,To_Char(A.登记时间,'YYYY-MM-DD HH24:MI:SS') as 登记时间," & _
                " Decode(Max(A.记录状态),2,'销帐记录','记帐记录') as 说明,Max(A.记录状态) as 符号,A.记帐单ID" & _
                " From (" & strFilter & ") A,部门表 B" & _
                " Where A.开单部门ID = B.ID" & _
                " And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & vbNewLine & _
                " Group by A.NO,B.名称,A.开单人,A.病人ID,A.标识号,A.姓名,A.费别,A.划价人,A.登记时间,A.记帐单ID" & _
                " Order by A.登记时间 Desc,A.NO Desc"
        End If
        With SQLCondition
            If .Default Then
                Set mrsList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "", "", "", "", "", "", "", "", "", mstr操作员, 1, 2, 3, 0)
            Else
                Set mrsList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .DateB, .DateE, .NOB, .NOE, .Patient, .PatientIdentity, .DeptID, .PatientNo, .PatientID, mstr操作员, 1, 2, 3, 0)
            End If
        End With
    End If
    
    mshList.Clear
    mshList.Rows = 2
    
    mshDetail.Clear
    mshDetail.Rows = 2
    
    If mrsList.EOF Then
        stbThis.Panels(2).Text = "当前设置没有过滤出任何单据"
        Call SetMenu(False)
    Else
        '求实收合计金额
        If Not blnSort Then
            strSQL = "Select Sum(实收金额) as 金额,Count(Distinct NO) as 单据 From (" & _
                Replace(strFilter, "记录状态 IN([11],[13]", "记录状态 IN([11],[12],[13]") & ") A,部门表 B Where A.开单部门ID = B.ID" & _
                " And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)"

            With SQLCondition
                If .Default Then
                    Set mrsTotal = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "", "", "", "", "", "", "", "", "", mstr操作员, 1, 2, 3, 0)
                Else
                    Set mrsTotal = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .DateB, .DateE, .NOB, .NOE, .Patient, Val(.PatientIdentity), .DeptID, .PatientNo, .PatientID, mstr操作员, 1, 2, 3, 0)
                End If
            End With
        End If

        Set mshList.DataSource = mrsList
        stbThis.Panels(2).Text = "共 " & NVL(mrsTotal!单据, 0) & " 张单据,合计:" & Format(NVL(mrsTotal!金额, 0), gstrDec)
        Call SetMenu(True)
    End If

    mshList.Redraw = False
    '设置颜色
    If mbln销帐 And Not mbln记帐 Then
        mshList.ForeColor = &HC0
    Else
        mshList.ForeColor = ForeColor
        k = GetColNum("符号")
        For i = 1 To mshList.Rows - 1
            If Val(mshList.TextMatrix(i, k)) = 2 Then
                '销帐记录用红色
                mshList.Row = i
                For j = 0 To mshList.COLS - 1
                    mshList.Col = j
                    mshList.CellForeColor = &HC0
                Next
            ElseIf Val(mshList.TextMatrix(i, k)) = 3 Then
                '包含销帐的用蓝色
                mshList.Row = i
                For j = 0 To mshList.COLS - 1
                    mshList.Col = j
                    mshList.CellForeColor = &HC00000
                Next
            End If
        Next
    End If
    
    Call SetHeader
    Call SetDetail
        
    mshList.Redraw = True
    
    If Not blnSort Then Call zlCommFun.StopFlash
    
    Me.Refresh
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub tbs_Click()
    If Not Visible Then Exit Sub
    If tbs.SelectedItem.Key = mstrPage Then Exit Sub
    mstrPage = tbs.SelectedItem.Key
     
    ShowBills mstrFilter
    On Error Resume Next
    mshList.SetFocus
End Sub

Private Sub SetActiveList(obj As Object)
    If obj Is mshList Then
        mshList.BackColorSel = &HC0C0C0
        mshDetail.BackColorSel = &HE0E0E0
    ElseIf obj Is mshDetail Then
        mshList.BackColorSel = &HE0E0E0
        mshDetail.BackColorSel = &HC0C0C0
    End If
End Sub

Private Sub ShowDetail(Optional ByVal strNo As String, Optional ByVal strTime As String, _
    Optional ByVal blnDel As Boolean, Optional ByVal blnSort As Boolean)

    Dim strSQL As String, i As Long, j As Long
    
    On Error GoTo errH
    
    If Not blnSort Then
        If frmBillingFilter.mblnDateMoved And tbs.SelectedItem.Key = "Auditing" Then
            '记帐划价单不检查是否在后备表中,因为不会转出到后备表
            mblnNOMoved = zlDatabase.NOMoved("门诊费用记录", strNo, , "2")
        Else
            mblnNOMoved = False   '必须要有这一句
        End If
        strSQL = _
        " Select C.名称 as 类别,Nvl(E.名称,B.名称) as 名称," & IIf(gTy_System_Para.byt药品名称显示 = 2, "E1.名称 as 商品名,", "") & "B.规格," & _
                IIf(gbln药房单位, "Decode(X.药品ID,NULL,A.计算单位,X." & gstr药房单位 & ")", "A.计算单位") & " as 单位," & _
        "       To_Char(Avg(Nvl(A.付数,1)*" & IIf(blnDel, "-1*", "") & "A.数次)" & _
                IIf(gbln药房单位, "/Nvl(X." & gstr药房包装 & ",1)", "") & ",'9999990.00000') as 数量, " & _
        "       To_Char(Sum(A.标准单价)" & IIf(gbln药房单位, "*Nvl(X." & gstr药房包装 & ",1)", "") & ",'999999" & gstrFeePrecisionFmt & "') as 单价, " & _
        "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.应收金额),'9999999" & gstrDec & "') as 应收金额, " & _
        "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.实收金额),'9999999" & gstrDec & "') as 实收金额, " & _
        "       D.名称 as 执行科室,Nvl(A.费用类型,B.费用类型) as 类型," & _
        "       Decode(Nvl(A.执行状态,0),0,'未执行',1,'完全执行',2,'部分执行','第'||ABS(A.执行状态)||'次退费') as 说明," & _
        "       A.记录状态" & _
        " From " & IIf(mblnNOMoved, zlGetFullFieldsTable("门诊费用记录"), "门诊费用记录 A") & "," & _
        "       收费项目目录 B,收费项目类别 C,部门表 D,收费项目别名 E," & IIf(gTy_System_Para.byt药品名称显示 = 2, "收费项目别名 E1,", "") & "药品规格 X" & _
        " Where A.收费细目ID=B.ID and A.收费类别=C.编码 And A.执行部门ID=D.ID(+) And A.收费细目ID=X.药品ID(+)" & _
        "       And A.记录性质=2 And A.NO=[1] And A.门诊标志 in(1,4) And A.记录状态" & IIf(blnDel, "=2", " IN(0,1,3)") & _
                IIf(strTime <> "", " And A.登记时间=[2]", "") & _
        "       And A.收费细目ID=E.收费细目ID(+) And E.码类(+)=1 And E.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
                IIf(gTy_System_Para.byt药品名称显示 = 2, "       And A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3", "") & _
        " Group by Nvl(A.价格父号,A.序号),C.名称,Nvl(E.名称,B.名称)," & IIf(gTy_System_Para.byt药品名称显示 = 2, "E1.名称 ,", "") & " B.规格," & _
        " A.计算单位,D.名称,Nvl(A.费用类型,B.费用类型),A.执行状态,A.记录状态,X.药品ID,X." & gstr药房单位 & ",Nvl(X." & gstr药房包装 & ",1)" & _
        " Order by Nvl(A.价格父号,A.序号)"
        If strTime <> "" Then
            Set mrsDetail = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo, CDate(strTime))
        Else
            Set mrsDetail = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo)
        End If
    End If
        
    mshDetail.Redraw = False
    
    mshDetail.Clear
    mshDetail.Rows = 2
    
    mshDetail.ForeColor = IIf(blnDel, &HC0, ForeColor)

    If Not mrsDetail.EOF Then Set mshDetail.DataSource = mrsDetail
    
    '设置颜色
    If blnDel Then
        '退费直接为红色
        mshDetail.ForeColor = &HC0
    Else
        '原始单据退过的为蓝色
        mshDetail.ForeColor = ForeColor
        For i = 1 To mshDetail.Rows - 1
            If Val(mshDetail.TextMatrix(i, mshDetail.COLS - 1)) = 3 Then
                mshDetail.Row = i
                For j = 0 To mshDetail.COLS - 1
                    mshDetail.Col = j
                    mshDetail.CellForeColor = &HC00000
                Next
            End If
        Next
    End If

    Call SetDetail
    
    mshDetail.Redraw = True
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub


Private Function zlMsgModuleInit() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化消息模块
    '入参:lngModule -模块号
    '     strPivs-权限串
    '出参:objMsgModule-返回消息对象
    '返回:初始化成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-03-11 11:46:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    Set mobjMsgModule = New clsMipModule
    Call mobjMsgModule.InitMessage(glngSys, mlngModul, mstrPrivs)
    Call AddMipModule(mobjMsgModule)
    zlMsgModuleInit = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Public Function zlMsgModuleUnload() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:拆卸消息模块
    '入参:objMsgModule-消息对象
    '编制:刘兴洪
    '日期:2014-03-11 11:46:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    
    If mobjMsgModule Is Nothing Then Exit Function
    Call mobjMsgModule.CloseMessage
    Call DelMipModule(mobjMsgModule)
    Set mobjMsgModule = Nothing
    zlMsgModuleUnload = False
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function


