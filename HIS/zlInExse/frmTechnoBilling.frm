VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTechnoBilling 
   AutoRedraw      =   -1  'True
   Caption         =   "医技科室记帐"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9885
   Icon            =   "frmTechnoBilling.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   Picture         =   "frmTechnoBilling.frx":08CA
   ScaleHeight     =   6195
   ScaleWidth      =   9885
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picHsc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   45
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   9780
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3885
      Width           =   9780
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   5835
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmTechnoBilling.frx":0A58
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8599
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
            Object.Width           =   3722
            MinWidth        =   3722
            Picture         =   "frmTechnoBilling.frx":0DCC
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
      TabIndex        =   4
      Top             =   0
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   9885
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinWidth1       =   6300
      MinHeight1      =   720
      Width1          =   4500
      NewRow1         =   0   'False
      Caption2        =   "医技科室"
      Child2          =   "cboUnit"
      MinWidth2       =   2100
      MinHeight2      =   300
      Width2          =   1800
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar tbr 
         Height          =   720
         Left            =   165
         TabIndex        =   5
         Top             =   30
         Width           =   6525
         _ExtentX        =   11509
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
               Caption         =   "记帐"
               Key             =   "Billing"
               Description     =   "记帐"
               Object.ToolTipText     =   "记帐"
               Object.Tag             =   "记帐"
               ImageKey        =   "Billing"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修改"
               Key             =   "Modi"
               Description     =   "修改"
               Object.ToolTipText     =   "修改"
               Object.Tag             =   "修改"
               ImageKey        =   "Modi"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "销帐"
               Key             =   "Del"
               Description     =   "销帐"
               Object.ToolTipText     =   "对当前选中单据销帐"
               Object.Tag             =   "销帐"
               ImageKey        =   "Del"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Del_"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查阅"
               Key             =   "View"
               Description     =   "查阅"
               Object.ToolTipText     =   "查阅"
               Object.Tag             =   "查阅"
               ImageKey        =   "View"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "过滤"
               Key             =   "Filter"
               Description     =   "过滤"
               Object.ToolTipText     =   "按设置条件重新筛选记录"
               Object.Tag             =   "过滤"
               ImageKey        =   "Filter"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "定位"
               Key             =   "Go"
               Description     =   "定位"
               Object.ToolTipText     =   "定位到满足条件的记录上"
               Object.Tag             =   "定位"
               ImageKey        =   "Go"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
      Begin VB.ComboBox cboUnit 
         Height          =   300
         Left            =   7695
         TabIndex        =   2
         Text            =   "cboUnit"
         Top             =   240
         Width           =   2100
      End
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   5205
      Top             =   90
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
            Picture         =   "frmTechnoBilling.frx":0F6A
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":1184
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":139E
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":15B8
            Key             =   "View"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":1D32
            Key             =   "Go"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":1F4C
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":2166
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":2380
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":259A
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":27B4
            Key             =   "Billing"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":2EAE
            Key             =   "Exe"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":35A8
            Key             =   "Cancel"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   4620
      Top             =   90
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
            Picture         =   "frmTechnoBilling.frx":3CA2
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":3EBC
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":40D6
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":42F0
            Key             =   "View"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":4A6A
            Key             =   "Go"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":4C84
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":4E9E
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":50B8
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":52D2
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":54EC
            Key             =   "Billing"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":5BE6
            Key             =   "Exe"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTechnoBilling.frx":62E0
            Key             =   "Cancel"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img32 
      Left            =   3105
      Top             =   90
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
            Picture         =   "frmTechnoBilling.frx":69DA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   3690
      Top             =   90
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
            Picture         =   "frmTechnoBilling.frx":72B4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
      Height          =   3135
      Left            =   15
      TabIndex        =   0
      Top             =   735
      Width           =   9870
      _ExtentX        =   17410
      _ExtentY        =   5530
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
      MouseIcon       =   "frmTechnoBilling.frx":784E
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDetail 
      Height          =   1875
      Left            =   0
      TabIndex        =   1
      Top             =   3960
      Width           =   9870
      _ExtentX        =   17410
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
      MouseIcon       =   "frmTechnoBilling.frx":7B68
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
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
      Begin VB.Menu mnuEditBilling 
         Caption         =   "记帐单(&B)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditSimple 
         Caption         =   "简单记帐(&S)"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuEditCust 
         Caption         =   "自定义记帐(&U)"
         Begin VB.Menu mnuEditCustBill 
            Caption         =   "(空)"
            Index           =   1
         End
      End
      Begin VB.Menu mnuEditBilling_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditModi 
         Caption         =   "修改单据(&M)"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuEditAdjust 
         Caption         =   "调整时间(&J)"
         Shortcut        =   ^J
      End
      Begin VB.Menu mnuEditAdjust_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDel 
         Caption         =   "单据销帐(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditDelApply 
         Caption         =   "销帐申请(&Q)"
      End
      Begin VB.Menu mnuEditDelAudit 
         Caption         =   "销帐审核(&H)"
      End
      Begin VB.Menu mnuEditDel_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditView 
         Caption         =   "查阅单据(&V)"
      End
      Begin VB.Menu mnuEditPrint 
         Caption         =   "打印单据(&P)"
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
         Begin VB.Menu mnuViewToolUnit 
            Caption         =   "医技科室(&U)"
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
         Caption         =   "过滤(&T)"
         Shortcut        =   ^T
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
Attribute VB_Name = "frmTechnoBilling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mrsList As ADODB.Recordset  '单据列表
Private mrsTotal As ADODB.Recordset
Private mrsDetail As ADODB.Recordset  '单据列表
Private mstrFilter As String
Private mbln记帐 As Boolean, mbln销帐 As Boolean

Private Type Type_SQLCondition
    Default As Boolean          '是否是缺省进入，此时没有条件值,缺省值在mstrFilter中
    DateB As Date
    DateE As Date
    NOB As String
    NOE As String
    InPatientID As Double
    Patient As String
    Operator As String
End Type
Private SQLCondition As Type_SQLCondition

Private mblnGo As Boolean, mlngGo As Long
Private mlngCurRow As Long, mlngTopRow As Long

Private mblnMax As Boolean
Private mlngDeptID As Long, mlngUnitID As Long

Private mstrPrivs As String     '保存当前模块的授权功能
Private mstrPrivsOpt As String '记帐操作1150模块的授权功能
Private mlngModul As Long
Private mblnNOMoved As Boolean '记录当前选择的单据是否是在后备数据表中
Private mrsDept As ADODB.Recordset

Private Sub cboUnit_Click()
    If cboUnit.ItemData(cboUnit.ListIndex) = mlngDeptID Then Exit Sub
    
    mlngDeptID = cboUnit.ItemData(cboUnit.ListIndex)
    mlngUnitID = Get病区ID(mlngDeptID)
    
    If Visible Then Call ShowBills(mstrFilter)
End Sub

Private Sub cboUnit_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long, lng医生ID As Long
    If KeyAscii <> 13 Then Exit Sub
    
    If cboUnit.ListIndex <> -1 Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    
    If mrsDept Is Nothing Then Call InitUnits
    
    
    
    If zlSelectDept(Me, mlngModul, cboUnit, mrsDept, cboUnit.Text, True, "") = False Then
        KeyAscii = 0: Exit Sub
    End If
    Exit Sub


End Sub

Private Sub cboUnit_Validate(Cancel As Boolean)
    If cboUnit.ListIndex >= 0 Then Exit Sub
    zlcontrol.CboLocate cboUnit, mlngDeptID, True
    If cboUnit.ListIndex < 0 And cboUnit.ListCount <> 0 Then cboUnit.ListIndex = 0

End Sub

Private Sub cbr_Resize()
    Form_Resize
End Sub

Private Sub Form_Activate()
    Call InitLocPar(mlngModul)
    Call mshList_GotFocus
End Sub

Private Sub mnuEditAdjust_Click()
    Dim strNO As String
    
    strNO = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
    If strNO = "" Then
        MsgBox "当前没有单据可以调整！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '是否已转入后备数据表中
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNO, 2, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '此时已转入在线数据表
    End If
    
    '未全部审核或多次审核的不允许修改
    If Not BillIdentical(strNO) Then
        MsgBox "单据中包含部份未审核或分多次审核的内容，不允许修改。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '已经冲销过(部分)的单据不允许调整
    If BillExistDelete(strNO, 2) Then
        MsgBox "该单据包含已销帐内容,不允许调整！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '是否已经结帐
    If HaveBilling(2, strNO) <> 0 Then
        Select Case gbytBillOpt
            Case 0
            Case 1
                If MsgBox("该记帐单据包含已经结帐的内容,要调整吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            Case 2
                MsgBox "该记帐单据包含已经结帐的内容,不能调整！", vbExclamation, gstrSysName: Exit Sub
        End Select
    End If
    
    On Error Resume Next
    Err.Clear
    
    If BillisSimple(strNO) Then '简单记帐
        frmSimpleBilling.mstrPrivs = mstrPrivs
        frmSimpleBilling.mbytInState = 2
        frmSimpleBilling.mstrInNO = strNO
        frmSimpleBilling.mbytUseType = 2
        frmSimpleBilling.mlngDeptID = mlngDeptID
        frmSimpleBilling.mlngUnitID = mlngUnitID
        frmSimpleBilling.mlngModule = mlngModul
        frmSimpleBilling.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    Else '记帐单
        Dim lng记帐ID As Long
        Dim varTemp As Variant
        
        lng记帐ID = mshList.TextMatrix(mshList.Row, GetColNum("记帐单ID"))
        
        If lng记帐ID = 0 Or gobjCustBill Is Nothing Then
            frmCharge.mstrPrivs = mstrPrivs
            frmCharge.mbytInState = 2
            frmCharge.mstrInNO = strNO
            frmCharge.mbytUseType = 2
            frmCharge.mlngDeptID = mlngDeptID
            frmCharge.mlngUnitID = mlngUnitID
            frmCharge.mlngModule = mlngModul
            frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
        Else
            '记帐ID、bytUseType、bytInState、strInNO、lngUnitID、lngDeptID、lng病人ID、mstrPrivs
            varTemp = Array(lng记帐ID, 2, 2, strNO, mlngUnitID, mlngDeptID, 0, mstrPrivs)
            gobjCustBill.CodeMan glngSys, -1, gcnOracle, Me, gstrDBUser, varTemp
            
            gblnOK = varTemp
        End If
    End If
End Sub

Private Sub mnuEditBilling_Click()
    Err.Clear
    On Error Resume Next
    
    gbytBilling = 0
    frmCharge.mstrPrivs = mstrPrivs
    frmCharge.mbytInState = 0
    frmCharge.mbytUseType = 2
    frmCharge.mlngDeptID = mlngDeptID
    frmCharge.mlngUnitID = mlngUnitID
    frmCharge.mlngModule = mlngModul
    frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ShowBills(mstrFilter)
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            Call ShowBills(mstrFilter)
        End If
    End If
End Sub

Private Sub mnuEditCustBill_Click(Index As Integer)
    '自定义记帐
    Dim varTemp As Variant
            
    '参数含义依次是：
    '记帐ID、bytUseType、bytInState、strInNO、lngUnitID、lngDeptID、lng病人ID、mstrPrivs、blnViewCancel
    varTemp = Array(mnuEditCustBill(Index).Tag, 2, 0, "", mlngUnitID, mlngDeptID, 0, mstrPrivs)
    gobjCustBill.CodeMan glngSys, -1, gcnOracle, Me, gstrDBUser, varTemp
    
    gblnOK = varTemp '返回值
    
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ShowBills(mstrFilter)
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            Call ShowBills(mstrFilter)
        End If
    End If
End Sub

Private Sub mnuEditDelApply_Click()
    If mlngDeptID = 0 Then
        MsgBox "请先选择当前科室!", vbInformation, gstrSysName
        cboUnit.SetFocus
        Exit Sub
    End If
    With frmReCharge
        .mlngDeptID = mlngDeptID
        .mbytUseType = 1
        .mbytFun = 0
        .mstrPrivs = mstrPrivs
        .Show IIf(gfrmMain Is Nothing, 0, 1), Me
    End With
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

Private Sub mnuEditDelAudit_Click()
    If mlngDeptID = 0 Then
        MsgBox "请先选择当前科室!", vbInformation, gstrSysName
        cboUnit.SetFocus
        Exit Sub
    End If
    With frmReCharge
        .mlngDeptID = mlngDeptID
        .mbytUseType = 1
        .mbytFun = 1
        .mstrPrivs = mstrPrivs
        .Show IIf(gfrmMain Is Nothing, 0, 1), Me
    End With
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

Private Sub mnuEditPrint_Click()
    Dim strNO As String, strTime As String
    
    strNO = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
    
    If strNO = "" Then
        MsgBox "当前没有单据可以打印！", vbInformation, gstrSysName
        Exit Sub
    End If

    If mshList.TextMatrix(mshList.Row, GetColNum("符号")) <> 1 Then
        MsgBox "该单据为销帐单据或已被销帐，不能再打印！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '是否已转入后备数据表中
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNO, 2, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '此时已转入在线数据表
    End If
    
    strTime = mshList.TextMatrix(mshList.Row, GetColNum("登记时间"))
    
    If ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1135", Me) Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1135", Me, "NO=" & strNO, "登记时间=" & strTime, "药品单位=" & IIf(gbln住院单位, 1, 0), "PrintEmpty=0", "重打=1", 2)
    End If
End Sub

Private Sub mnuEditSimple_Click()
    Err.Clear
    On Error Resume Next
    
    gbytBilling = 0
    frmSimpleBilling.mstrPrivs = mstrPrivs
    frmSimpleBilling.mbytInState = 0
    frmSimpleBilling.mbytUseType = 2
    frmSimpleBilling.mlngDeptID = mlngDeptID
    frmSimpleBilling.mlngUnitID = mlngUnitID
    frmSimpleBilling.mlngModule = mlngModul
    frmSimpleBilling.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ShowBills(mstrFilter)
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            Call ShowBills(mstrFilter)
        End If
    End If
End Sub

Private Sub mnuEditModi_Click()
    Dim strNO As String, intInsure As Integer
    Dim strInfo As String, strUnitIDs As String
    
    strNO = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
    If strNO = "" Then
        MsgBox "当前没有单据可以修改！", vbInformation, gstrSysName
        Exit Sub
    End If

    '是否已转入后备数据表中
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNO, 2, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '此时已转入在线数据表
    End If
    
    '未全部审核或多次审核的不允许修改
    If Not BillIdentical(strNO) Then
        MsgBox "单据中包含部份未审核或分多次审核的内容，不允许修改。", vbInformation, gstrSysName
        Exit Sub
    End If

    '权限判断
    If Not BillOperCheck(5, mshList.TextMatrix(mshList.Row, GetColNum("操作员")), _
        CDate(mshList.TextMatrix(mshList.Row, GetColNum("登记时间"))), "修改", strNO) Then Exit Sub
    
    '全院销帐
    If InStr(mstrPrivsOpt, ";全院销帐;") = 0 Then
        If strUnitIDs = "" Then strUnitIDs = GetUserUnits(True)
        
        If InStr("," & strUnitIDs & ",", "," & Val(mshList.TextMatrix(mshList.Row, GetColNum("开单部门ID"))) & ",") = 0 Then
            MsgBox "你没有权限对其它科室的单据销帐,不允许修改该单据！", vbInformation, gstrSysName
            Exit Sub
        End If
    End If

    '留观病人权限
    strInfo = Check留观病人(strNO, mstrPrivsOpt)
    If strInfo <> "" Then
        MsgBox "单据中包含" & strInfo & ",你没有权限对该单据进行操作！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '是否有出院修改(记帐)权限
    If Not BillCanBeOperate(strNO, mstrPrivsOpt, "修改") Then Exit Sub
    
    '去掉了医保连接匹配检查
    
    '包含分批或时价的药品不允许修改
    If Not BillCanModi(strNO, 2) Then
        MsgBox "该张单据中包含分批或时价药品,不允许修改！", vbInformation, gstrSysName
        Exit Sub
    End If

    '已经冲销过(部分)的单据不允许修改
    If BillExistDelete(strNO, 2) Then
        MsgBox "该单据包含已销帐费用,不允许修改！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '如果包含部分执行或全部执行的项目,则不一定可以全部冲销,不允许修改
    If HaveExecute(2, strNO, 2) Then
        MsgBox "该单据中包含完全执行或部分执行的项目,不允许修改！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '是否已经结帐
    If HaveBilling(2, strNO) <> 0 Then
        intInsure = BillExistInsure(strNO)
        If intInsure <> 0 Then
            If Not gclsInsure.GetCapability(support允许冲销已结帐的记帐单据, , intInsure) Then
                '医保病人的单据固定为已结帐就禁止修改
                MsgBox "该医保记帐单据包含已经结帐的内容,不能修改！", vbExclamation, gstrSysName: Exit Sub
            End If
        Else
            Select Case gbytBillOpt
                Case 0
                Case 1
                    If MsgBox("该记帐单据包含已经结帐的内容,要修改吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                Case 2
                    MsgBox "该记帐单据包含已经结帐的内容,不能修改！", vbExclamation, gstrSysName: Exit Sub
            End Select
        End If
    End If
    
    gstrModiNO = ""
    
    On Error Resume Next
    Err.Clear
    
    gbytBilling = 0 '记帐修改
    If BillisSimple(strNO) Then '简单记帐
        frmSimpleBilling.mstrPrivs = mstrPrivs
        frmSimpleBilling.mbytInState = 0
        frmSimpleBilling.mstrInNO = strNO
        frmSimpleBilling.mbytUseType = 2
        frmSimpleBilling.mlngDeptID = mlngDeptID
        frmSimpleBilling.mlngUnitID = mlngUnitID
        frmSimpleBilling.mlngModule = mlngModul
        frmSimpleBilling.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    Else '记帐单
        Dim lng记帐ID As Long
        Dim varTemp As Variant
        
        lng记帐ID = mshList.TextMatrix(mshList.Row, GetColNum("记帐单ID"))
        
        If lng记帐ID = 0 Or gobjCustBill Is Nothing Then
            frmCharge.mstrPrivs = mstrPrivs
            frmCharge.mbytInState = 0
            frmCharge.mstrInNO = strNO
            frmCharge.mbytUseType = 2
            frmCharge.mlngDeptID = mlngDeptID
            frmCharge.mlngUnitID = mlngUnitID
            frmCharge.mlngModule = mlngModul
            frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
        Else
            '记帐ID、bytUseType、bytInState、strInNO、lngUnitID、lngDeptID、lng病人ID、mstrPrivs
            varTemp = Array(lng记帐ID, 2, 0, strNO, mlngUnitID, mlngDeptID, 0, mstrPrivs)
            gobjCustBill.CodeMan glngSys, -1, gcnOracle, Me, gstrDBUser, varTemp
            
            gblnOK = varTemp
        End If
    End If

    If gblnOK Then
        If gstrModiNO <> "" Then
            If mnuViewRefeshOptionItem(1).Checked Then
                If MsgBox("当前操作已更改单据清单内容,修改后的单据号为:[" & gstrModiNO & "],要刷新吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    Call ShowBills(mstrFilter)
                End If
            ElseIf mnuViewRefeshOptionItem(2).Checked Then
                Call ShowBills(mstrFilter)
            End If
        Else
            If mnuViewRefeshOptionItem(1).Checked Then
                If MsgBox("当前操作已更改单据清单内容,要刷新吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    Call ShowBills(mstrFilter)
                End If
            ElseIf mnuViewRefeshOptionItem(2).Checked Then
                Call ShowBills(mstrFilter)
            End If
        End If
    End If
End Sub

Private Sub mnuFileLocalSet_Click()
    Dim bln住院单位 As Boolean
    
    bln住院单位 = gbln住院单位
    
    frmSetExpence.mlngModul = mlngModul
    frmSetExpence.mstrPrivs = mstrPrivs
    frmSetExpence.mbytInFun = 0
    frmSetExpence.mbytUseType = 2
    frmSetExpence.Show 1, Me
    If gblnOK Then
        
        If bln住院单位 <> gbln住院单位 Then
            If Not (mshList.Rows = 2 And mshList.TextMatrix(1, GetColNum("单据号")) = "") Then
                Call mnuViewReFlash_Click
            End If
        End If
    End If
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Dim strNO As String
    
    strNO = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
    If strNO = "" Then
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
            "病区=" & mlngUnitID, "病人科室=" & mlngDeptID)
    Else
        With mshList
            Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
                "病区=" & mlngUnitID, "病人科室=" & mlngDeptID, "NO=" & strNO, _
                "住院号=" & .TextMatrix(.Row, GetColNum("住院号")), _
                "开单人=" & .TextMatrix(.Row, GetColNum("开单人")))
        End With
    End If
End Sub

Private Sub mnuViewFilter_Click()
    
    If frmTechnoFilter.mlngDept <> mlngDeptID Then
        frmTechnoFilter.mlngDept = mlngDeptID
        frmTechnoFilter.LoadOper
    End If
    
    frmTechnoFilter.Show 1, Me
    If gblnOK Then
        With frmTechnoFilter
            mstrFilter = .mstrFilter
            mbln记帐 = .chk记帐.Value = 1
            mbln销帐 = .chk销帐.Value = 1
            
            SQLCondition.Default = False
            SQLCondition.DateB = .dtpBegin.Value
            SQLCondition.DateE = .dtpEnd.Value
            SQLCondition.NOB = .txtNOBegin.Text
            SQLCondition.NOE = .txtNoEnd.Text
            SQLCondition.InPatientID = Val(.txt住院号.Text)
            SQLCondition.Patient = gstrLike & UCase(.txt姓名.Text) & "%"
            SQLCondition.Operator = zlStr.NeedName(.cbo操作员.Text)
        End With
        
        mnuViewReFlash_Click
    End If
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
    If mnuEditView.Enabled Then mnuEditView_Click
End Sub

Private Sub mshList_EnterCell()
    Dim strNO As String, strTime As String, blnDel As Boolean
        
    strNO = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
    
    If mshList.Row = 0 Or strNO = "" Then Exit Sub
    
    stbThis.Panels(2).Text = "共 " & Nvl(mrsTotal!单据, 0) & " 张单据,合计:" & Format(Nvl(mrsTotal!金额, 0), gstrDec)
    
    mlngGo = mshList.Row
    mlngCurRow = mshList.Row: mlngTopRow = mshList.TopRow
    
    strTime = mshList.TextMatrix(mshList.Row, GetColNum("登记时间"))
    blnDel = Val(mshList.TextMatrix(mshList.Row, GetColNum("符号"))) = 2
    
    mnuEditAdjust.Enabled = Not blnDel
    mnuEditModi.Enabled = Not blnDel And Val(mshList.TextMatrix(mshList.Row, GetColNum("医嘱序号"))) = 0 _
                        And Val(mshList.TextMatrix(mshList.Row, GetColNum("记录性质"))) <> 3
    mnuEditDel.Enabled = Not blnDel
    tbr.Buttons("Modi").Enabled = mnuEditModi.Enabled
    tbr.Buttons("Del").Enabled = mnuEditDel.Enabled
        
    mshList.ForeColorSel = mshList.CellForeColor
    
    Call ShowDetail(strNO, strTime, blnDel)
End Sub

Private Sub mshList_GotFocus()
    Call SetActiveList(mshList)
End Sub

Private Sub mshList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete And mnuEditDel.Enabled And mnuEditDel.Visible Then Call mnuEditDel_Click
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
            If Me.ActiveControl Is cboUnit Then
            Else
                If mnuEditView.Enabled Then mnuEditView_Click
            End If
        Case vbKeyEscape
            mblnGo = False
    End Select
End Sub

Private Sub mnuEditDel_Click()
    Dim intInsure As Integer, strInfo As String
    Dim strNO As String, strTime As String
    Dim intTmp As Integer, i As Long, blnFlagPrint As Boolean
    
    strNO = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
    If strNO = "" Then
        MsgBox "当前没有单据可以销帐！", vbInformation, gstrSysName
        Exit Sub
    End If
    strTime = mshList.TextMatrix(mshList.Row, GetColNum("登记时间"))
    
    '权限判断
    If Not BillOperCheck(5, mshList.TextMatrix(mshList.Row, GetColNum("操作员")), CDate(strTime), "销帐", strNO) Then Exit Sub

    '未全部审核或多次审核的不允许修改
    If Not BillIdentical(strNO) Then
        MsgBox "单据中包含部份未审核或分多次审核的内容，不允许修改。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '是否已转入后备数据表中
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNO, 2, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '此时已转入在线数据表
    End If
    
    '留观病人权限
    strInfo = Check留观病人(strNO, mstrPrivsOpt, strTime)
    If strInfo <> "" Then
        MsgBox "单据中包含" & strInfo & ",你没有权限对该单据进行操作！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '是否已执行
    i = BillCanDelete(strNO, 2, , strTime, mstrPrivsOpt, blnFlagPrint)
    If i <> 0 Then
        Select Case i
            Case 1 '该单据不存在
                MsgBox "指定单据中的内容不存在！", vbInformation, gstrSysName
            Case 2 '已经全部完全执行
                MsgBox "指定单据中的内容已经全部完全执行！", vbInformation, gstrSysName
            Case 3 '未完全执行部分剩余数量为0
                MsgBox "指定单据中的内容未执行部分项目剩余数量为零,没有可以销帐的费用！", vbInformation, gstrSysName
        End Select
        Exit Sub
    End If
    If blnFlagPrint Then
        If MsgBox("注意:检验医嘱的条码已打印，是否继续？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    
    '是否有出院修改(记帐)权限
    If Not BillCanBeOperate(strNO, mstrPrivsOpt, "销帐", strTime) Then Exit Sub
    
    '是否已经结帐
    intInsure = BillExistInsure(strNO)
    intTmp = HaveBilling(2, strNO, False, strTime)
    If intTmp <> 0 Then
        If intInsure <> 0 Then
            If Not gclsInsure.GetCapability(support允许冲销已结帐的记帐单据, , intInsure) Then
                '医保病人的单据,固定为已结帐的禁止销帐
                If intTmp = 1 Then
                    MsgBox "该医保记帐单据未销帐部分已经结帐,不能销帐！", vbExclamation, gstrSysName
                    Exit Sub
                Else
                    MsgBox "该医保记帐单据包含已经结帐的内容,只能对未结帐部分进行销帐！", vbExclamation, gstrSysName
                End If
            End If
        Else
            Select Case gbytBillOpt
                Case 0
                Case 1
                    If MsgBox("该记帐单据包含已经结帐的内容,要销帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                Case 2
                    If intTmp = 1 Then
                        MsgBox "该记帐单据未销帐部分已经结帐,不能销帐！", vbExclamation, gstrSysName
                        Exit Sub
                    Else
                        MsgBox "该记帐单据包含已经结帐的内容,只能对未结帐部分进行销帐！", vbExclamation, gstrSysName: Exit Sub
                    End If
            End Select
        End If
    End If
    
    '医保销帐不允许对负数记录进行销帐
    If intInsure <> 0 Then
        If CheckNONegative(strNO) Then
            MsgBox "该单据存在负数记帐记录,不允许进行医保销帐操作！", vbInformation, gstrSysName
             Exit Sub
        End If
    End If
        
    '是否存在重算冲减记录
    If CheckRecalcRecord(strNO) Then
        MsgBox "发现该记帐单据存在按费别重算的打折冲减记录!" & vbCrLf & _
            "结帐前请按费别重算费用，否则病人将享受已销帐单据的打折优惠金额！", vbInformation, Me.Caption
    End If
    
    On Error Resume Next
    Err.Clear
    
    If BillisSimple(strNO) Then '简单记帐
        frmSimpleBilling.mbytUseType = 2
        frmSimpleBilling.mstrPrivs = mstrPrivs
        frmSimpleBilling.mbytInState = 3
        frmSimpleBilling.mstrInNO = strNO
        frmSimpleBilling.mstrTime = strTime
        frmSimpleBilling.mlngDeptID = mlngDeptID
        frmSimpleBilling.mlngUnitID = mlngUnitID
        frmSimpleBilling.mlngModule = mlngModul
        frmSimpleBilling.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    Else '记帐单
        Dim lng记帐ID As Long, varTemp As Variant
        
        lng记帐ID = mshList.TextMatrix(mshList.Row, GetColNum("记帐单ID"))
        
        If lng记帐ID = 0 Or gobjCustBill Is Nothing Then
            frmCharge.mbytUseType = 2
            frmCharge.mstrPrivs = mstrPrivs
            frmCharge.mbytInState = 3
            frmCharge.mstrInNO = strNO
            frmCharge.mstrTime = strTime
            frmCharge.mlngDeptID = mlngDeptID
            frmCharge.mlngUnitID = mlngUnitID
            frmCharge.mlngModule = mlngModul
            frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
        Else
            '记帐ID、bytUseType、bytInState、strInNO、lngUnitID、lngDeptID、lng病人ID、mstrPrivs
            varTemp = Array(lng记帐ID, 2, 3, strNO, 0, 0, 0, mstrPrivs)
            gobjCustBill.CodeMan glngSys, -1, gcnOracle, Me, gstrDBUser, varTemp
            
            gblnOK = varTemp
        End If
    End If

    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前操作已更改单据清单内容,要刷新吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ShowBills(mstrFilter)
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            Call ShowBills(mstrFilter)
        End If
    End If
End Sub

Private Sub mnuHelpTitle_Click()
ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub mnuEditView_Click()
    Dim strNO As String, strTime As String, blnDel As Boolean
    
    strNO = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
    If strNO = "" Then
        MsgBox "当前没有单据可以查阅！", vbInformation, gstrSysName
        Exit Sub
    End If

    strTime = mshList.TextMatrix(mshList.Row, GetColNum("登记时间"))
    blnDel = Val(mshList.TextMatrix(mshList.Row, GetColNum("符号"))) = 2

    On Error Resume Next
    Err.Clear
    
    If BillisSimple(strNO) Then '简单记帐
        frmSimpleBilling.mbytUseType = 2
        frmSimpleBilling.mlngDeptID = mlngDeptID
        frmSimpleBilling.mstrPrivs = mstrPrivs
        frmSimpleBilling.mbytInState = 1
        frmSimpleBilling.mstrInNO = strNO
        frmSimpleBilling.mblnNOMoved = mblnNOMoved
        frmSimpleBilling.mstrTime = strTime
        frmSimpleBilling.mblnDelete = blnDel
        frmSimpleBilling.mlngModule = mlngModul
        frmSimpleBilling.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    Else '记帐单
        Dim lng记帐ID As Long
        Dim varTemp As Variant
        
        lng记帐ID = mshList.TextMatrix(mshList.Row, GetColNum("记帐单ID"))
        
        If lng记帐ID = 0 Or gobjCustBill Is Nothing Then
            frmCharge.mlngDeptID = mlngDeptID
            frmCharge.mbytUseType = 2
            frmCharge.mstrPrivs = mstrPrivs
            frmCharge.mbytInState = 1
            frmCharge.mstrInNO = strNO
            frmCharge.mblnNOMoved = mblnNOMoved
            frmCharge.mstrTime = strTime
            frmCharge.mblnDelete = blnDel
            frmCharge.mlngModule = mlngModul
            frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
        Else
            '记帐ID、bytUseType、bytInState、strInNO、lngUnitID、lngDeptID、lng病人ID、mstrPrivs
            varTemp = Array(lng记帐ID, 2, 1, strNO, 0, mlngDeptID, 0, mstrPrivs, blnDel)
            gobjCustBill.CodeMan glngSys, -1, gcnOracle, Me, gstrDBUser, varTemp
            
            gblnOK = varTemp
        End If
    End If
End Sub

Private Sub mnuFile_quit_Click()
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuViewReFlash_Click()
    Call ShowBills(mstrFilter)
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = Not stbThis.Visible
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim i As Long
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For i = 1 To tbr.Buttons.Count
        tbr.Buttons(i).Caption = IIf(mnuViewToolText.Checked, tbr.Buttons(i).Tag, "")
    Next
    cbr.Bands(1).minHeight = tbr.ButtonHeight
    Form_Resize
End Sub

Private Sub mnuViewToolUnit_Click()
    mnuViewToolUnit.Checked = Not mnuViewToolUnit.Checked
    If mnuViewToolButton.Checked Then cbr.Bands(1).Visible = False
    cbr.Bands(2).Visible = Not cbr.Bands(2).Visible
    If mnuViewToolButton.Checked Then cbr.Bands(1).Visible = True
    cbr.Visible = cbr.Bands(2).Visible Or cbr.Bands(1).Visible
    Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    cbr.Bands(1).Visible = Not cbr.Bands(1).Visible
    mnuViewToolText.Enabled = Not mnuViewToolText.Enabled
    cbr.Visible = cbr.Bands(2).Visible Or cbr.Bands(1).Visible
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
            mnuEditView_Click
        Case "Billing"
            mnuEditBilling_Click
        Case "Modi"
            mnuEditModi_Click
        Case "Del"
            mnuEditDel_Click
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
    objOut.Title.Text = "住院记帐单据清单"
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '表项
    With frmTechnoFilter
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
    mshList.Col = 0: mshList.ColSel = mshList.Cols - 1
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
    
    mnuEditAdjust.Enabled = blnUsed
    mnuEditModi.Enabled = blnUsed
    tbr.Buttons("Modi").Enabled = blnUsed
    
    mnuEditDel.Enabled = blnUsed
    mnuEditView.Enabled = blnUsed
    mnuEditPrint.Enabled = blnUsed
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
    
    If gobjCustBill Is Nothing Then
        Set gobjCustBill = CreateObject("zl9CustAcc.clsCustAcc")
    End If
    If InStr(mstrPrivsOpt, ";专项记帐;") = 0 Then
        mnuEditCust.Visible = False
        Exit Sub
    End If
    
    On Error GoTo errHandle
    
    '如果创建成功，再读出对应的菜单
    If Not gobjCustBill Is Nothing Then
        gstrSQL = "Select ID,名称 From 收费记帐单 Where substr(适用范围,4,1)='1' Order by 编号"
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
    mstrPrivsOpt = GetInsidePrivs(Enum_Inside_Program.p记帐操作)
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
    
    mlngCurRow = 1: mlngTopRow = 1
    
    '权限设置
    If InStr(mstrPrivsOpt, ";住院记帐;") = 0 Then
        mnuEditBilling.Visible = False
        mnuEditSimple.Visible = False
        mnuEditBilling_.Visible = False
        tbr.Buttons("Billing").Visible = False
        
        mnuEditCust.Visible = False
    End If
    '55380
    If InStr(mstrPrivsOpt, ";药品销帐;") = 0 _
        And InStr(mstrPrivsOpt, ";卫材销帐;") = 0 _
        And InStr(mstrPrivsOpt, ";诊疗销帐;") = 0 Then
        mnuEditDel.Visible = False
        If InStr(mstrPrivsOpt, ";药品销帐申请;") = 0 _
            And InStr(mstrPrivsOpt, ";卫材销帐申请;") = 0 _
            And InStr(mstrPrivsOpt, ";诊疗销帐申请;") = 0 _
            And InStr(mstrPrivsOpt, ";销帐审核;") = 0 Then
            mnuEditDel_.Visible = False
        End If
        tbr.Buttons("Del").Visible = False
    End If
    '55380
    If InStr(mstrPrivsOpt, ";药品销帐申请;") = 0 _
        Or InStr(mstrPrivsOpt, ";诊疗销帐申请;") = 0 _
        Or InStr(mstrPrivsOpt, ";卫材销帐申请;") = 0 _
        Or InStr(1, mstrPrivsOpt, ";部分销帐;") = 0 Then
        mnuEditDelApply.Visible = False
    End If
    
    If InStr(mstrPrivsOpt, ";销帐审核;") = 0 Then
        mnuEditDelAudit.Visible = False
    End If
    
    If InStr(mstrPrivsOpt, ";记录修改;") = 0 Then
        mnuEditModi.Visible = False
        tbr.Buttons("Modi").Visible = False
    End If
    If InStr(mstrPrivsOpt, ";记录调整;") = 0 Then
        mnuEditAdjust.Visible = False
    End If
    If InStr(mstrPrivsOpt, ";记录修改;") = 0 _
        And InStr(mstrPrivsOpt, ";记录调整;") = 0 Then
        mnuEditAdjust_.Visible = False
    End If
    '55380
    If InStr(mstrPrivsOpt, ";住院记帐;") = 0 _
        And InStr(mstrPrivsOpt, ";记录修改;") = 0 _
        And (InStr(mstrPrivsOpt, ";药品销帐;") = 0 _
        And InStr(mstrPrivsOpt, ";卫材销帐;") = 0 _
        And InStr(mstrPrivsOpt, ";诊疗销帐;") = 0) Then
        tbr.Buttons("Del_").Visible = False
    End If
        
    '科室
    If Not InitUnits Then Unload Me: Exit Sub
    If cboUnit.ListIndex = -1 Then
        MsgBox "没有发现你所属科室,且你不具有所有科室权限,不能使用医技科室记帐！", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    
    mbln记帐 = True
    mbln销帐 = False
    
    Call SetHeader
    Call SetDetail
    Call SetMenu(False)
    
    stbThis.Panels(2).Text = "请刷新清单或重新设置过滤条件"
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
    
    mshList.Left = 0
    mshList.Top = cbrH
    mshList.Width = Me.ScaleWidth
    mshList.Height = (Me.ScaleHeight - cbrH - staH - picHsc.Height) * (1 - sngVsc)
    
    picHsc.Left = Me.ScaleLeft
    picHsc.Top = mshList.Top + mshList.Height
    picHsc.Width = Me.ScaleWidth
    
    mshDetail.Left = Me.ScaleLeft
    mshDetail.Top = picHsc.Top + picHsc.Height
    mshDetail.Width = Me.ScaleWidth
    mshDetail.Height = Me.ScaleHeight - staH - cbrH - picHsc.Height - mshList.Height
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    
    mstrFilter = ""
    mlngDeptID = 0
    mlngUnitID = 0
    
    Unload frmTechnoFilter
    Unload frmTechnoGo
    Call SaveWinState(Me, App.ProductName)
    
    '刷新方式
    For i = 0 To mnuViewRefeshOptionItem.UBound
        If mnuViewRefeshOptionItem(i).Checked Then
            zlDatabase.SetPara "刷新方式", i, glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
            Exit For
        End If
    Next
    
End Sub

Private Sub mnuViewGo_Click()
    frmTechnoGo.Show 1, Me
    If gblnOK Then Call SeekBill(frmTechnoGo.optHead)
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
        With frmTechnoGo
            If .txtNO.Text <> "" Then
                blnFill = blnFill And mshList.TextMatrix(i, GetColNum("单据号")) = .txtNO.Text
            End If
            If .txt住院号.Text <> "" Then
                blnFill = blnFill And mshList.TextMatrix(i, GetColNum("住院号")) = .txt住院号.Text
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
            mshList.Col = 0: mshList.ColSel = mshList.Cols - 1
            
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
    For i = 0 To mshList.Cols - 1
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
    
    strHead = "单据号,1,850|开单人,1,800|病人科室,1,850|住院号,1,750|床号,1,500|姓名,1,700|费别,1,900|应收金额,7,850|实收金额,7,850|" & _
            "操作员,1,800|登记时间,1,1850|说明,1,850|符号,1,0|记录性质,1,0|多病人单,1,0|记帐单ID,1,0|病人ID,1,0|主页ID,1,0|医嘱序号,1,0|开单部门ID,1,0"
    With mshList
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(mshList, App.ProductName & "\" & Me.Name)
        .RowHeight(0) = 320
        
        i = GetColNum("符号"): mshList.ColWidth(i) = 0
        i = GetColNum("记录性质"): mshList.ColWidth(i) = 0
        i = GetColNum("多病人单"): mshList.ColWidth(i) = 0
        i = GetColNum("记帐单ID"): mshList.ColWidth(i) = 0
        
        '查看医生的权限
        i = GetColNum("开单人")
        If InStr(mstrPrivsOpt, ";医生查询;") = 0 Then
            mshList.ColWidth(i) = 0
        ElseIf mshList.ColWidth(i) = 0 Then
            mshList.ColWidth(i) = 800
        End If
        
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
    End With
End Sub

Private Sub ShowBills(Optional ByVal strIF As String, Optional blnSort As Boolean)
'功能:按条件读取单据列表(过滤功能)
'参数:strIF=以"AND"开始的条件串
'     blnSort=不重新读取数据,仅重新显示已排序的内容
    Dim i As Long, j As Long, k As Long
    Dim Curdate As Date, strSql As String
    
    On Error GoTo errH
        
    If Not blnSort Then
        Call zlCommFun.ShowFlash("正在读取单据列表,请稍候 ...", Me)
        DoEvents
        Me.Refresh
        
        '取缺省条件(当日记帐)
        SQLCondition.Default = (strIF = "")
        If strIF = "" Then
            strIF = " And 登记时间 Between trunc(sysdate) And trunc(sysdate+1)-1/24/60/60 And 记录状态 IN(1,3)"
            strIF = strIF & " And 操作员姓名||''=[7]"
        End If
        strIF = strIF & " And 开单部门ID+0=[8]"
        strIF = " Where 记录性质=2 And 门诊标志=2 And 记录状态<>0 And 操作员姓名 is Not NULL And Nvl(多病人单,0)=0  " & strIF
        
        '筛选时的时间在最后一次转出之前
        If frmTechnoFilter.mblnDateMoved Then
            strIF = zlGetFullFieldsTable("住院费用记录", 2, strIF, False)
        Else
            strIF = zlGetFullFieldsTable("住院费用记录", 0, strIF, False)
        End If
                
        '单据号,开单人,病人科室,住院号,床号,姓名,费别,应收金额,实收金额,操作员,登记时间,说明,符号,记录性质,多病人单,记帐单ID
        strSql = _
        " Select A.NO as 单据号,A.开单人," & _
        "        B.名称 as 病人科室,A.标识号 as 住院号,C.出院病床 as 床号,A.姓名,A.费别," & _
        "        To_Char(Sum(Decode(A.记录状态,2,-1,1)*A.应收金额),'9999999" & gstrDec & "') as 应收金额," & _
        "        To_Char(Sum(Decode(A.记录状态,2,-1,1)*A.实收金额),'9999999" & gstrDec & "') as 实收金额," & _
        "        A.操作员姓名 as 操作员,To_Char(A.登记时间,'YYYY-MM-DD HH24:MI:SS') as 登记时间," & _
        "        Decode(A.记录性质,3,Decode(Max(A.记录状态),2,'自动销帐','自动记帐'),Decode(Max(A.记录状态),2,'销帐记录','记帐记录')) as 说明," & _
        "        Max(A.记录状态) as 符号,A.记录性质,A.多病人单,A.记帐单ID,A.病人ID,A.主页ID,Nvl(A.医嘱序号,0) 医嘱序号,A.开单部门ID" & _
        " From (" & strIF & ") A,部门表 B,病案主页 C" & _
        " Where A.病人科室ID=B.ID(+) And A.病人ID=C.病人ID(+) And A.主页ID=C.主页ID (+)" & _
        " Group by A.NO,A.开单人,B.名称,A.标识号,C.出院病床,A.姓名,A.费别,A.操作员姓名," & _
        "          A.登记时间,A.记录性质,A.多病人单,A.记帐单ID,A.病人ID,A.主页ID,Nvl(A.医嘱序号,0),A.开单部门ID" & _
        " Order by A.登记时间 Desc,A.NO Desc"
        
        With SQLCondition
            If .Default Then .Operator = UserInfo.姓名
            Set mrsList = zlDatabase.OpenSQLRecord(strSql, Me.Caption, .DateB, .DateE, .NOB, .NOE, .InPatientID, .Patient, .Operator, mlngDeptID)
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
            strSql = "Select Sum(实收金额) as 金额,Count(Distinct NO) as 单据 From (" & Replace(strIF, "记录状态 IN(1,3)", "记录状态 IN(1,2,3)") & ")"
            With SQLCondition
                Set mrsTotal = zlDatabase.OpenSQLRecord(strSql, Me.Caption, .DateB, .DateE, .NOB, .NOE, .InPatientID, .Patient, .Operator, mlngDeptID)
            End With
        End If
    
        Set mshList.DataSource = mrsList
        stbThis.Panels(2).Text = "共 " & Nvl(mrsTotal!单据, 0) & " 张单据,合计:" & Format(Nvl(mrsTotal!金额, 0), gstrDec)
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
                For j = 0 To mshList.Cols - 1
                    mshList.Col = j
                    mshList.CellForeColor = &HC0
                Next
            ElseIf Val(mshList.TextMatrix(i, k)) = 3 Then
                '包含销帐的用蓝色
                mshList.Row = i
                For j = 0 To mshList.Cols - 1
                    mshList.Col = j
                    mshList.CellForeColor = &HC00000
                Next
            End If
        Next
    End If
    
    Call SetHeader
    If mshList.Row = 0 Or mshList.TextMatrix(mshList.Row, GetColNum("单据号")) = "" Then Call SetDetail
        
    mshList.Redraw = True
    
    If Not blnSort Then Call zlCommFun.StopFlash
    
    Me.Refresh
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function InitUnits() As Boolean
'功能：初始化医技科室
    Dim i As Long, strSql As String
    
    On Error GoTo errH
        
    '包含门诊/住院医技科室
    If InStr(mstrPrivs, ";所有科室;") > 0 Then
        strSql = _
            " Select Distinct A.ID,A.编码,A.名称,A.简码" & _
            " From 部门表 A,部门性质说明 B " & _
            " Where B.部门ID = A.ID " & _
            " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
            " And B.服务对象 IN(1,2,3) And B.工作性质 IN('检查','检验','手术','治疗','营养')" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
            " Order by A.编码"
    Else
        strSql = _
            " Select Distinct A.ID,A.编码,A.名称,A.简码,Nvl(C.缺省,0) as 缺省" & _
            " From 部门表 A,部门性质说明 B,部门人员 C " & _
            " Where B.部门ID = A.ID And A.ID=C.部门ID And C.人员ID=[1]" & _
            " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
            " And B.服务对象 IN(1,2,3) And B.工作性质 IN('检查','检验','手术','治疗','营养')" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
            " Order by A.编码"
    End If
    Set mrsDept = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID)
    
    If Not mrsDept.EOF Then
        For i = 1 To mrsDept.RecordCount
            cboUnit.AddItem mrsDept!编码 & "-" & mrsDept!名称
            cboUnit.ItemData(cboUnit.NewIndex) = mrsDept!ID
            If cboUnit.ListIndex = -1 Then
                If InStr(mstrPrivs, ";所有科室;") > 0 Then
                    If UserInfo.部门ID = mrsDept!ID Then cboUnit.ListIndex = cboUnit.NewIndex
                Else
                    If mrsDept!缺省 = 1 Then cboUnit.ListIndex = cboUnit.NewIndex
                End If
            End If
            mrsDept.MoveNext
        Next
        If cboUnit.ListIndex = -1 And cboUnit.ListCount > 0 Then cboUnit.ListIndex = 0
    ElseIf InStr(mstrPrivs, ";所有科室;") > 0 Then
        MsgBox "没有可用的医技科室,请先到部门管理中设置！", vbInformation, gstrSysName
        Exit Function
    End If
    InitUnits = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetActiveList(obj As Object)
    If obj Is mshList Then
        mshList.BackColorSel = &HC0C0C0
        mshDetail.BackColorSel = &HE0E0E0
    ElseIf obj Is mshDetail Then
        mshList.BackColorSel = &HE0E0E0
        mshDetail.BackColorSel = &HC0C0C0
    End If
End Sub

Private Sub SetDetail()
    Dim strHead As String
    Dim i As Long
    
    strHead = "类别,1,650|名称,1,1600" & IIf(gTy_System_Para.byt药品名称显示 = 2, "|商品名,1,1600", "") & "|规格,1,1000|单位,4,500|数量,7,850|单价,7,850|应收金额,7,850|实收金额,7,850|统筹金额,7,850|执行科室,1,850|类型,1,850|说明,1,1000|记录状态,1,0"
    
    With mshDetail
        .Redraw = False
        
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(mshDetail, App.ProductName & "\" & Me.Name)
        '刘兴洪:27990 2010-02-22 17:34:32
        For i = 0 To .Cols - 1
            If .TextMatrix(0, i) = "商品名" Then
                If gTy_System_Para.byt药品名称显示 = 2 Then
                    If .ColWidth(i) <= 0 Then .ColWidth(i) = 1600
                Else
                    .ColWidth(i) = 0
                End If
            End If
        Next
                
        .RowHeight(0) = 320
        .ColWidth(.Cols - 1) = 0
        
        .Row = 1: .Col = 0: .ColSel = .Cols - 1
        
        Call mshDetail_EnterCell
        
        .Redraw = True
    End With
End Sub

Private Sub ShowDetail(Optional ByVal strNO As String, Optional ByVal strTime As String, _
    Optional ByVal blnDel As Boolean, Optional ByVal blnSort As Boolean)
    
    Dim strSql As String, i As Long, j As Long
    
    On Error GoTo errH
        
    If Not blnSort Then
        
        If frmTechnoFilter.mblnDateMoved Then
            mblnNOMoved = zlDatabase.NOMoved("住院费用记录", strNO, , 2, Me.Caption)
        Else
            mblnNOMoved = False   '必须要有这一句
        End If
        
        strSql = _
        " Select C.名称 as 类别,Nvl(E.名称,B.名称) as 名称," & IIf(gTy_System_Para.byt药品名称显示 = 2, "E1.名称 as 商品名,", "") & "B.规格," & _
                IIf(gbln住院单位, "Decode(X.药品ID,NULL,A.计算单位,X.住院单位)", "A.计算单位") & " as 单位," & _
        "       To_Char(Avg(Nvl(A.付数,1)*" & IIf(blnDel, "-1*", "") & "A.数次)" & _
                IIf(gbln住院单位, "/Nvl(X.住院包装,1)", "") & ",'9999990.00000') as 数量, " & _
        "       To_Char(Sum(A.标准单价)" & IIf(gbln住院单位, "*Nvl(X.住院包装,1)", "") & ",'99999" & gstrFeePrecisionFmt & "') as 单价, " & _
        "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.应收金额),'9999999" & gstrDec & "') as 应收金额, " & _
        "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.实收金额),'9999999" & gstrDec & "') as 实收金额, " & _
        "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.统筹金额),'9999999" & gstrDec & "') as 统筹金额, " & _
        "       D.名称 as 执行科室,Nvl(A.费用类型,B.费用类型) as 类型," & _
        "       Decode(Nvl(A.执行状态,0),0,'未执行',1,'已执行',2,'部分执行','第'||ABS(A.执行状态)||'次退费') as 说明 , A.记录状态" & _
        " From " & IIf(mblnNOMoved, zlGetFullFieldsTable("住院费用记录"), "住院费用记录 A") & " ," & _
        "       收费项目目录 B,收费项目类别 C,部门表 D,收费项目别名 E,药品规格 X" & _
                  IIf(gTy_System_Para.byt药品名称显示 = 2, ",收费项目别名 E1", "") & _
        " Where A.收费细目ID=B.ID And A.收费类别=C.编码 And A.执行部门ID=D.ID(+)" & _
        "       And A.NO=[1] And A.记录性质=2 And A.门诊标志=2 And Nvl(A.多病人单,0)=0" & _
        "       And A.收费细目ID=X.药品ID(+) And A.记录状态" & IIf(blnDel, "=2", " IN(1,3)") & IIf(strTime <> "", " And A.登记时间=[2]", "") & _
        "       And A.收费细目ID=E.收费细目ID(+) And E.码类(+)=1 And E.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
                IIf(gTy_System_Para.byt药品名称显示 = 2, "       And A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3", "") & _
        " Group by Nvl(A.价格父号,A.序号),C.名称,Nvl(E.名称,B.名称)," & IIf(gTy_System_Para.byt药品名称显示 = 2, "E1.名称 ,", "") & " B.规格,A.计算单位," & _
        "       D.名称,Nvl(A.费用类型,B.费用类型),A.执行状态,A.记录状态,X.药品ID,X.住院单位,Nvl(X.住院包装,1)" & _
        " Order by Nvl(A.价格父号,A.序号)"
        If strTime <> "" Then
            Set mrsDetail = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNO, CDate(strTime))
        Else
            Set mrsDetail = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNO)
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
            If Val(mshDetail.TextMatrix(i, mshDetail.Cols - 1)) = 3 Then
                mshDetail.Row = i
                For j = 0 To mshDetail.Cols - 1
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

Private Sub mnuViewRefeshOptionItem_Click(Index As Integer)
    Dim i As Long
    For i = 0 To mnuViewRefeshOptionItem.UBound
        mnuViewRefeshOptionItem(i).Checked = i = Index
    Next
End Sub


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

