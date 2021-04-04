VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmStuffPlanList 
   Caption         =   "卫材计划管理"
   ClientHeight    =   5895
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11970
   Icon            =   "frmStuffPlanList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5895
   ScaleWidth      =   11970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin TabDlg.SSTab TabShow 
      Height          =   345
      Left            =   1680
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   609
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "申购部门(&0)"
      TabPicture(0)   =   "frmStuffPlanList.frx":014A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "被申购部门(&1)"
      TabPicture(1)   =   "frmStuffPlanList.frx":0166
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
   End
   Begin VB.PictureBox picSeparate_s 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   30
      MousePointer    =   7  'Size N S
      ScaleHeight     =   300
      ScaleWidth      =   4815
      TabIndex        =   6
      Top             =   2790
      Width           =   4815
   End
   Begin VB.CommandButton Cmd查阅 
      Caption         =   "查阅(&V)"
      Height          =   350
      Left            =   5160
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1100
   End
   Begin ComCtl3.CoolBar cbrTool 
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   11775
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tlbTool"
      MinHeight1      =   720
      Width1          =   6210
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Caption2        =   "库房"
      Child2          =   "cboStock"
      MinHeight2      =   300
      Width2          =   4095
      NewRow2         =   0   'False
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   6915
         TabIndex        =   4
         Text            =   "cboStock"
         Top             =   240
         Width           =   4770
      End
      Begin MSComctlLib.Toolbar tlbTool 
         Height          =   720
         Left            =   165
         TabIndex        =   1
         Top             =   30
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsCold"
         HotImageList    =   "ilsHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   16
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "PrintView"
               Description     =   "预览"
               Object.ToolTipText     =   "预览"
               Object.Tag             =   "预览"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "Print"
               Description     =   "打印"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "PrintSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "增加"
               Key             =   "Add"
               Description     =   "增加"
               Object.ToolTipText     =   "增加"
               Object.Tag             =   "增加"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修改"
               Key             =   "Modify"
               Description     =   "修改"
               Object.ToolTipText     =   "修改"
               Object.Tag             =   "修改"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "Delete"
               Description     =   "删除"
               Object.ToolTipText     =   "删除"
               Object.Tag             =   "删除"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "EditSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "审核"
               Key             =   "Verify"
               Description     =   "审核"
               Object.ToolTipText     =   "审核"
               Object.Tag             =   "审核"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "取消"
               Key             =   "Cancel"
               Object.ToolTipText     =   "取消审核"
               Object.Tag             =   "取消"
               ImageKey        =   "Cancel"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "清除"
               Key             =   "Clear"
               Description     =   "清除"
               Object.ToolTipText     =   "清除"
               Object.Tag             =   "清除"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "VerifySeparate"
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "过滤"
               Key             =   "Search"
               Description     =   "过滤"
               Object.ToolTipText     =   "过滤"
               Object.Tag             =   "过滤"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "刷新"
               Key             =   "Refresh"
               Description     =   "刷新"
               Object.ToolTipText     =   "刷新"
               Object.Tag             =   "刷新"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FindSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助主题"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Exit"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   11
            EndProperty
         EndProperty
         MouseIcon       =   "frmStuffPlanList.frx":0182
         Begin VB.Timer LimitTime 
            Enabled         =   0   'False
            Interval        =   8000
            Left            =   6660
            Top             =   180
         End
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   5532
      Width           =   11964
      _ExtentX        =   21114
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmStuffPlanList.frx":049C
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16034
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
   Begin MSComctlLib.ImageList ilsCold 
      Left            =   0
      Top             =   600
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
            Picture         =   "frmStuffPlanList.frx":0D30
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":0F50
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":1170
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":138C
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":15AC
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":17CC
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":19E8
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":1C04
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":1E1E
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":1F78
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":2194
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":23B4
            Key             =   "Cancel"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHot 
      Left            =   600
      Top             =   600
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
            Picture         =   "frmStuffPlanList.frx":250E
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":272E
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":294E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":2B6A
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":2D8A
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":2FAA
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":31C6
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":33E2
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":35FC
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":3756
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":3976
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanList.frx":3B96
            Key             =   "Cancle"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
      Height          =   1815
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   3201
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      RowSizingMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDetail 
      Height          =   1965
      Left            =   0
      TabIndex        =   7
      Top             =   3360
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   3466
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   0
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuFilePreView 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileBillPrint 
         Caption         =   "单据打印(&B)"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuFileBillPreview 
         Caption         =   "单据预览(&L)"
      End
      Begin VB.Menu mnuFileLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnuFileLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileParameter 
         Caption         =   "参数设置(&R)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFileLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditAdd 
         Caption         =   "新增(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "修改(&M)"
      End
      Begin VB.Menu mnuEditDel 
         Caption         =   "删除(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditVerify 
         Caption         =   "审核(&C)"
      End
      Begin VB.Menu mnuEditCancel 
         Caption         =   "取消(&Q)"
      End
      Begin VB.Menu mnuEditClear 
         Caption         =   "清除(&S)"
      End
      Begin VB.Menu mnuEditLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDisplay 
         Caption         =   "查看单据(&W)"
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
         Begin VB.Menu mnuViewToolText 
            Caption         =   "文本标签(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "状态栏(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewSearch 
         Caption         =   "过滤(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewLine4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
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
            Caption         =   "发送反馈(&M)..."
         End
      End
      Begin VB.Menu mnuHelpLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
End
Attribute VB_Name = "frmStuffPlanList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngMode As Long
Private mstrFind As String
Private mintPreCol As Integer           '前一次单据头的排序列
Private mintsort As Integer             '前一次单据头的排序
Private mblnBootUp As Boolean
Private mlastRow As Long                '上次电击的行

Private mdtStartDate As Date
Private mdtEndDate As Date
Private mdtVerifyStart As Date
Private mdtVerifyEnd As Date
Private mstrPrivs As String
Private mintUnit  As Integer                '显示单位:0-散装单位,1-包装单位
Private mintOldY  As Integer
Private mstrOthers() As String  '0-记录状态(计划类型),1-开始单号,2-结束单号,3-材料id,4-对方部门id(或入出类别id或编制方法(计划用)),5-填制人,6-审核人,7-供应商ID,8-生产商,9-开始生产日期,10-结束生产日期,11-开始发票号,12-结束发票号
Private mblnCostView As Boolean             '查看成本价相关信息 true-允许查看 false-不允许查看
Private mblnProvider As Boolean             '查看上次供应商相关信息 true-允许查看 false-不允许查看
Private mstrCaption As String           '窗体标题
Private mintFindDay As Integer          '查询天数范围
 
'---------------------------------------------------------------------------------------------------------
'设置相关的过滤条件:2008-08-22 16:35:52
'刘兴宏:
Private mblnNoClick As Boolean
Private mstr工作性质 As String
Private mbln操作员限制 As Boolean

Private mstrTitle As String '标题

'----------------------------------------------------------------------------------------------------------
'刘兴宏:增加小数位数的格式串
'修改:2007/03/06
Private mOraFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Public Sub ShowList(ByVal lngMode As Long, ByVal strTitle As String, ByVal frmMain As Variant)
    '--------------------------------------------------------------------------------------------------------------------------
    '功能:显示指定模块的入口
    '参数:lngMode-模块号
    '     strTitle-标题
    '     frmMain-父窗口
    '返回:
    '编制:刘兴宏
    '日期:2007/12/26
    '问题:11282
    '--------------------------------------------------------------------------------------------------------------------------
    Dim strOthers(0 To 12) As String
    Dim i As Integer
    Dim intCol As Integer
    
    mstrCaption = strTitle
    mblnBootUp = False
    mlngMode = lngMode
    mstrTitle = strTitle
    mstrPrivs = gstrPrivs
    
    mintFindDay = Val(zlDatabase.GetPara("查询天数", glngSys, mlngMode, 1))
    mdtStartDate = Format(DateAdd("d", -mintFindDay, sys.Currentdate), "yyyy-MM-dd")
    mdtEndDate = Format(sys.Currentdate, "yyyy-MM-dd")
    
    mdtVerifyStart = "1901-01-01"
    mdtVerifyEnd = "1901-01-01"
    
    mstrFind = " AND A.审核日期 is Null And A.编制日期 Between [2] And [3]"
    
    If Not CheckDepend Then Exit Sub            '数据依赖性测试
    mblnCostView = zlStr.IsHavePrivs(mstrPrivs, "查看成本价")
    mblnProvider = zlStr.IsHavePrivs(mstrPrivs, "查看供应商")
    
    Me.Caption = strTitle
    SetVisable  '根据权限设置不同的显示项目
        
    For i = 0 To 12
        strOthers(i) = ""
    Next
    '设置生产日期
    strOthers(9) = "1901-01-01"
    strOthers(10) = "1901-01-01"
    mintUnit = Val(zlDatabase.GetPara("卫材单位", glngSys, mlngMode, "0"))
  
    '刘兴宏:增加小数格式化串
    With mOraFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价, True)
        .FM_金额 = GetFmtString(mintUnit, g_金额, True)
        .FM_零售价 = GetFmtString(mintUnit, g_售价, True)
        .FM_数量 = GetFmtString(mintUnit, g_数量, True)
    End With
    mstrOthers = strOthers
    mstrPrivs = gstrPrivs
    mlastRow = 0
    
    If mlngMode <> 1725 Then
        TabShow.Visible = False
    Else
        TabShow.Visible = True
    End If
    
    GetList (mstrFind)   '列出单据头
    RestoreWinState Me, App.ProductName, mstrTitle
    '恢复个性化参数设置后，还需要对权限控制的列进一步设置
    With mshDetail
        For intCol = 1 To .Cols - 1
            If mlngMode = 1725 Or mlngMode = 1724 Then
                If InStr(1, .TextMatrix(0, intCol), "成本价") > 0 Or InStr(1, .TextMatrix(0, intCol), "成本金额") > 0 Then
                    .ColWidth(intCol) = IIf(mblnCostView = True, 900, 0)
                End If
            End If
            If mlngMode = 1725 Then
                If InStr(1, .TextMatrix(0, intCol), "上次供应商") > 0 Then
                    .ColWidth(intCol) = IIf(mblnProvider = True, 1000, 0)
                End If
            End If
        Next
    End With
    
    '2006-04-25:刘兴宏,统一增加报表发布到模块的功能
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
    If mlngMode = 1725 Then
        cbrTool.Bands(2).Caption = "部门"
    End If
    mblnBootUp = True
    
    If IsObject(frmMain) Then
        Me.Show , frmMain
    Else
        OS.ShowChildWindow Me.hwnd, frmMain
    End If
    Me.ZOrder 0
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub


'检查数据依赖性
Private Function CheckDepend() As Boolean
    
    Dim rsTemp As New Recordset
    Dim strStock As String
    
    On Error GoTo ErrHandle
    CheckDepend = False
    strStock = " And b.编码 In('V','K','12','W')"
    
    If mlngMode = 1725 Then
        gstrSQL = "" & _
            "   SELECT DISTINCT a.id, a.编码||'-'||a.名称 as 名称" & _
            "   FROM 部门表 a  " & _
            "   where (a.撤档时间 is null or TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01') " & _
            "       And (a.站点=[2] or a.站点 is null) " & _
            IIf(InStr(1, mstrPrivs, ";所有部门;") > 0, "", " and  id in (Select 部门id from 部门人员 where 人员id =[1])") & _
            "   Order by 名称"
            mstr工作性质 = ""
            mbln操作员限制 = Not zlStr.IsHavePrivs(mstrPrivs, "所有部门")
    Else
        gstrSQL = "" & _
            "   SELECT DISTINCT a.id , a.编码||'-'||a.名称 as 名称 " & _
            "   FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " & _
            "   Where c.工作性质 = b.名称 And (a.站点=[2] or a.站点 is null) " & _
            "             " & strStock & _
            "           AND a.id = c.部门id " & _
            "           AND (a.撤档时间 is null or TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01')" & _
            IIf(InStr(1, mstrPrivs, ";所有库房;") > 0, "", " and a.id in (Select 部门id from 部门人员 where 人员id =[1])") & _
            "   Order by 名称"
            mstr工作性质 = "V,K,12,W"
            mbln操作员限制 = Not zlStr.IsHavePrivs(mstrPrivs, "所有库房")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrTitle, UserInfo.Id, gstrNodeNo)
    
    If rsTemp.EOF Then
        If mlngMode = 1725 Then
            MsgBox "没有划分部门体系或你不具备相关的权限,请查看部门管理或找系统管理员授权！", vbInformation, gstrSysName
        Else
            MsgBox "没有划分卫材库性质的部门或不具备相关的权限,请查看部门管理或找系统管事员授权！", vbInformation, gstrSysName
        End If
        rsTemp.Close
        Exit Function
    End If
            
    With cboStock
        .Clear
        If mlngMode <> 1725 Then
            If InStr(1, mstrPrivs, ";所有库房;") > 0 Then
                .AddItem "全院"
                .ItemData(.NewIndex) = 0
            End If
        End If
        
        If InStr(1, mstrPrivs, ";所有部门;") > 0 Then
            .AddItem "所有部门"
            .ItemData(.NewIndex) = 0
        End If
        Do While Not rsTemp.EOF
            .AddItem rsTemp!名称
            .ItemData(.NewIndex) = rsTemp!Id
            If rsTemp!Id = UserInfo.部门ID Then
                .ListIndex = .NewIndex
            End If
            rsTemp.MoveNext
        Loop
        rsTemp.Close
        If .ListIndex < 0 Then .ListIndex = 0
    End With
    
    CheckDepend = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Initialize()
    Call InitCommonControls
End Sub

Private Sub cboStock_Click()
    If mblnNoClick Then Exit Sub
    If cboStock.ListIndex >= 0 Then cboStock.Tag = cboStock.ItemData(cboStock.ListIndex)
    If mblnBootUp Then mnuViewRefresh_Click
End Sub
Private Sub cboStock_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cboStock.ListCount = 0 Then Call zlControl.ControlSetFocus(mshList): Exit Sub
    
    If cboStock.ListIndex >= 0 Then
        If Val(cboStock.Tag) = cboStock.ItemData(cboStock.ListIndex) Then
            Call zlControl.ControlSetFocus(mshList, True)
            Exit Sub
        End If
    End If
    
    If Select部门选择器(Me, cboStock, Trim(cboStock.Text), mstr工作性质, mbln操作员限制) = False Then
        Exit Sub
    End If
    If cboStock.ListIndex >= 0 Then
        cboStock.Tag = cboStock.ItemData(cboStock.ListIndex)
    End If
End Sub

Private Sub cboStock_LostFocus()
    Dim i As Long
    If cboStock.ListCount = 0 Then Exit Sub
    If cboStock.ListIndex < 0 Then
        For i = 0 To cboStock.ListCount - 1
            If Val(cboStock.Tag) = cboStock.ItemData(i) Then
                mblnNoClick = True
                cboStock.ListIndex = i: Exit For
            End If
        Next
    End If
    mblnNoClick = False
End Sub

Private Sub cbrTool_Resize()
    Form_Resize
End Sub

Private Sub GetList(ByVal strFind As String)
    Dim rsTemp As New Recordset
    
    On Error GoTo ErrHandle
    Call FS.ShowFlash("正在搜索材料计划记录,请稍候 ...", Me)
    DoEvents
    Screen.MousePointer = vbHourglass
    mshList.Redraw = False
    If mlngMode = 1725 Then
        If TabShow.Tab = 0 Then
        If cboStock.ItemData(cboStock.ListIndex) <> 0 Then '选中了所有库房时就不需要库房ID了
            strFind = strFind & " and nvl(a.部门id,0) =[1] "
        End If
        
        gstrSQL = "" & _
            "   SELECT a.no,a.id,b.名称 as 部门, decode(a.计划类型,1,'月度计划',2,'季度计划',3,'年度计划',4,'周度计划') as 计划类型," & _
            "           a.期间,a.编制人,to_char(a.编制日期,'yyyy-mm-dd HH24:MI:SS') as 编制日期, a.审核人," & _
            "           to_char(a.审核日期,'yyyy-mm-dd HH24:MI:SS') as 审核日期,a.编制说明 " & _
            "   From 材料采购计划 a,部门表 b  " & _
            "  Where a.单据=1 and a.部门id=b.id " & strFind & _
            " ORDER BY a.no desc "
        Else
            If cboStock.ItemData(cboStock.ListIndex) <> 0 Then '选中了所有库房时就不需要库房ID了
                strFind = strFind & " and nvl(a.库房id,0) =[1] "
            End If
            
            gstrSQL = "" & _
            "   SELECT a.no,a.id,b.名称 as 部门, decode(a.计划类型,1,'月度计划',2,'季度计划',3,'年度计划',4,'周度计划') as 计划类型," & _
            "           a.期间,a.编制人,to_char(a.编制日期,'yyyy-mm-dd HH24:MI:SS') as 编制日期, a.审核人," & _
            "           to_char(a.审核日期,'yyyy-mm-dd HH24:MI:SS') as 审核日期,a.编制说明 " & _
            "   From 材料采购计划 a,部门表 b  " & _
            "  Where a.单据=1 and a.库房id=b.id " & strFind & _
            " ORDER BY a.no desc "
        End If
    Else
        gstrSQL = "" & _
            "   SELECT no,id, decode(计划类型,1,'月度计划',2,'季度计划',3,'年度计划',4,'周度计划') as 计划类型 ," & _
            "           期间,decode(编制方法,1,'往年同期线形参照法',2,'临近期间平均参照法',3,'材料储备定额参照法',4, '卫材日销售量参照法', '部门申购参照法') as 编制方法 ," & _
            "           编制人,to_char(编制日期,'yyyy-mm-dd HH24:MI:SS') as 编制日期, 审核人," & _
            "           to_char(审核日期,'yyyy-mm-dd HH24:MI:SS') as 审核日期,编制说明 " & _
            "   From 材料采购计划 a " & _
            "  Where nvl(库房id,0) =[1] and 单据=0 " & strFind & _
            " ORDER BY a.no desc "
    End If
    
    'mstrOthers(0 To 12) As String ' 0-记录状态(计划类型),1-开始单号,2-结束单号,3-材料id,4-对方部门id(或入出类别id或编制方法(计划用)),5-填制人,6-审核人,7-供应商ID,8-生产商,9-开始生产日期,10-结束生产日期,11-开始发票号,12-结束发票号
    '参数范围:[1]-库房id,[2]:开始填制日期,[3]结束填制日期,[4]开始审核日期,[5] 结束审核日期,[6]-记录状态,[7]开始单据号,[8]结束单据号,[9]材料id,[10]对方部门id,[11]填制人,[12]审核人[13]-供应商ID,[14]-生产商,[15]-开始生产日期,[16]-结束生产日期,[17]-开始发票号,[18]-结束发票号
    
    '初始生产日期
    mstrOthers(9) = IIf(Trim(mstrOthers(9)) = "", "1901-01-01", mstrOthers(9))
    mstrOthers(10) = IIf(Trim(mstrOthers(10)) = "", "1901-01-01", mstrOthers(10))
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, cboStock.ItemData(cboStock.ListIndex), _
        CDate(Format(mdtStartDate, "yyyy-mm-dd") & " 00:00:00"), CDate(Format(mdtEndDate, "yyyy-mm-dd") & " 23:59:59"), _
        CDate(Format(mdtVerifyStart, "yyyy-mm-dd") & " 00:00:00"), CDate(Format(mdtVerifyEnd, "yyyy-mm-dd") & " 23:59:59"), _
        Val(mstrOthers(0)), mstrOthers(1), mstrOthers(2), Val(mstrOthers(3)), _
        Val(mstrOthers(4)), mstrOthers(5), mstrOthers(6), _
        Val(mstrOthers(7)), mstrOthers(8), CDate(mstrOthers(9) & " 00:00:00"), CDate(mstrOthers(10) & " 23:59:59"), _
         mstrOthers(11), mstrOthers(12))
          
          
    Set mshList.Recordset = rsTemp
    With mshList
        If .Rows = 1 Then
            .Rows = .Rows + 100
            .Row = 1
            .Redraw = True
            .TopRow = 1
            .Rows = .Rows - 99
        End If
        .Row = 1
        .Col = 0
        .ColSel = .Cols - 1
    End With
    SetListColWidth
    mshList.Redraw = True
    Call FS.StopFlash
    Screen.MousePointer = vbDefault
    SetEnable
    stbThis.Panels(2).Text = "当前共有" & rsTemp.RecordCount & "张单据"
    rsTemp.Close
    Call mshlist_EnterCell
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'表头列宽初始
Private Sub SetListColWidth()
    Dim intCol As Integer
    
    With mshList
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        
        If mblnBootUp = False Then
            For intCol = 0 To .Cols - 1
                .ColWidth(intCol) = 1500
                .ColAlignmentFixed(intCol) = 4
            Next
        Else
            For intCol = 0 To .Cols - 1
                .ColAlignmentFixed(intCol) = 4
            Next
        End If
        
        .ColWidth(1) = 0
    End With
End Sub

'根据权限设置不同的显示项目
Private Sub SetVisable()
    '外购入库所有权限：参数设置、基本、所有库房、登记、修改、删除、验收、清除

    If InStr(1, mstrPrivs, ";增加;") = 0 Then
        mnuEditAdd.Visible = False
        tlbTool.Buttons("Add").Visible = False
    End If
    
    If InStr(1, mstrPrivs, ";修改;") = 0 Then
        mnuEditModify.Visible = False
        tlbTool.Buttons("Modify").Visible = False
    End If
    
    
    If InStr(1, mstrPrivs, ";删除;") = 0 Then
        mnuEditDel.Visible = False
        tlbTool.Buttons("Delete").Visible = False
         '对没有所有编辑权限时，把菜单和工具栏上的相应的分割线屏蔽。
        If mnuEditAdd.Visible = False And mnuEditModify.Visible = False Then
            mnuEditLine1.Visible = False
            tlbTool.Buttons("EditSeparate").Visible = False
        End If
    End If
    
    If InStr(1, mstrPrivs, ";审核;") = 0 Then
        mnuEditVerify.Visible = False
        tlbTool.Buttons("Verify").Visible = False
    End If
    If InStr(1, mstrPrivs, ";取消;") = 0 Then
        mnuEditCancel.Visible = False
        tlbTool.Buttons("Cancel").Visible = False
    End If
    
    If InStr(1, mstrPrivs, ";单据打印;") = 0 Then
        mnuFileBillPrint.Visible = False
        mnuFileBillPreview.Visible = False
    End If
    
    If InStr(1, mstrPrivs, ";清除;") = 0 Then
        mnuEditClear.Visible = False
        tlbTool.Buttons("Clear").Visible = False
        If mnuEditVerify.Visible = False And mnuEditCancel.Visible = False Then
            mnuEditLine2.Visible = False
            tlbTool.Buttons("VerifySeparate").Visible = False
        End If
    End If
    
End Sub
Private Sub Cmd查阅_Click()
    Call mnuEditDisplay_Click
End Sub

Private Sub Form_Load()
    PrintRange "查询范围:" & Format(mdtStartDate, "yyyy年MM月dd日") & "至" & Format(mdtEndDate, "yyyy年MM月dd日")
End Sub

Private Sub Form_Resize()
    '窗体位置设置
    
    On Error Resume Next
    
    If Me.WindowState = 1 Then Exit Sub
    
    cbrTool.Bands(1).MinHeight = tlbTool.Height
    With cbrTool
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth - .Left
    End With
    
    With picSeparate_s
        .Height = 300
        .Left = 0
        .Width = cbrTool.Width
    End With
    
    If mlngMode = 1725 Then
        With TabShow
            .Top = IIf(cbrTool.Visible, cbrTool.Height, 0)
            .Left = 0
        End With
        
        With mshList
            .Top = TabShow.Top + TabShow.Height
            .Left = 0
            .Width = cbrTool.Width
            .Height = picSeparate_s.Top - .Top
        End With
    Else
        With mshList
            .Top = IIf(cbrTool.Visible, cbrTool.Height, 0)
            .Left = 0
            .Width = cbrTool.Width
            .Height = picSeparate_s.Top - .Top
        End With
    End If
    
    With Cmd查阅
        .Left = Me.ScaleWidth - .Width - 100
        .Top = mshList.Top + mshList.Height + 30
        .ZOrder
    End With
    
    With mshDetail
        .Top = picSeparate_s.Top + picSeparate_s.Height + 100
        .Left = 0
        .Height = ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
        .Width = cbrTool.Width
    End With
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, mstrTitle
End Sub


 
Private Sub mnuEditAdd_Click()
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    strNo = ""
    '新增
    Select Case mlngMode
    Case 1725
        If cboStock.ItemData(cboStock.ListIndex) <> 0 Then
            frmStuffRequestPlanCard.ShowCard Me, strNo, 1, mstrPrivs, blnSuccess
        End If
    Case 1724
        frmStuffPlanCard.ShowCard Me, strNo, 1, blnSuccess
    End Select
    If blnSuccess = True Then
        mnuViewRefresh_Click
    End If
End Sub


Private Sub mnuEditCancel_Click()
    '验收
    
    Dim strNo As String
    Dim blnSuccess As Boolean
    Dim lngBillId  As Long
    
    With mshList
        strNo = Trim(.TextMatrix(.Row, 0))
        lngBillId = .TextMatrix(.Row, 1)
        If strNo = "" Then Exit Sub
        If MsgBox("你确实要取消单据号为“" & strNo & "”的" & IIf(mlngMode = 1725, "请购单", "采购计划单") & "的审批信息吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) <> vbYes Then Exit Sub
    End With
    err = 0: On Error GoTo ErrHand:
    'zl_材料计划管理_Cancel(ID)
    gstrSQL = "zl_材料计划管理_Cancel(" & lngBillId & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    mnuViewRefresh_Click
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub mnuEditClear_Click()
    '清除
    Dim lngBillId As Long
    Dim intRow As Integer
    Dim strTitle As String
    Dim intReturn As Integer
    Dim intRecord As Integer
     
    With mshList
        
        On Error GoTo ErrHandle
        intRow = .Row
        lngBillId = .TextMatrix(intRow, 1)
        intReturn = MsgBox("你确实要清除单据号为“" & .TextMatrix(.Row, 0) & "”的" & IIf(mlngMode = 1725, "请购单", "采购计划单") & "吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
        intRecord = .Rows - 1
        If intReturn = vbYes Then
            gstrSQL = "zl_材料计划管理_DELETE('" & lngBillId & "')"
            If gstrSQL = "" Then Exit Sub
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            intRecord = intRecord - 1
            If .Rows > 2 Then
                .RemoveItem intRow
            ElseIf .Rows = 2 Then
                .Rows = 3
                .RemoveItem intRow
                SetEnable
            End If
            If intRow < .Rows - 1 Then
                .Row = intRow
            Else
                If .Rows = 2 Then
                    .Row = 1
                Else
                    .Row = intRow - 1
                End If
            End If
            .Col = 0
            .ColSel = .Cols - 1
        End If
    End With
    
    mlastRow = 0
    Call mshlist_EnterCell
    stbThis.Panels(2).Text = "当前共有" & intRecord & "张单据"
    Exit Sub

ErrHandle:
    Exit Sub
End Sub

Private Sub mnuEditVerify_Click()
    '验收
    
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    With mshList
        strNo = .TextMatrix(.Row, 0)
        Select Case mlngMode
        Case 1725
            frmStuffRequestPlanCard.ShowCard Me, strNo, 3, mstrPrivs, blnSuccess
        Case 1724
            frmStuffPlanCard.ShowCard Me, strNo, 3, blnSuccess
        End Select

    
    End With
    If blnSuccess = True Then
        mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditDel_Click()
    '删除
    Dim lngBillId As Long
    Dim intRow As Integer
    Dim strTitle As String
    Dim intReturn As Integer
    Dim intRecord As Integer
     
    With mshList
        
        On Error GoTo ErrHandle
        intRow = .Row
        lngBillId = .TextMatrix(intRow, 1)
        
        intReturn = MsgBox("你确实要删除单据号为“" & .TextMatrix(.Row, 0) & "”的" & IIf(mlngMode = 1725, "请购单", "采购计划单") & "吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
        intRecord = .Rows - 1
        If intReturn = vbYes Then
            gstrSQL = "zl_材料计划管理_DELETE('" & lngBillId & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            intRecord = intRecord - 1
            If .Rows > 2 Then
                .RemoveItem intRow
            ElseIf .Rows = 2 Then
                .Rows = 3
                .RemoveItem intRow
                
                SetEnable
            End If
            If intRow < .Rows - 1 Then
                .Row = intRow
            Else
                If .Rows = 2 Then
                    .Row = 1
                Else
                    .Row = intRow - 1
                End If
            End If
            .Col = 0
            .ColSel = .Cols - 1
        End If
    End With
    mlastRow = 0
    Call mshlist_EnterCell
    stbThis.Panels(2).Text = "当前共有" & intRecord & "张单据"
    Exit Sub
ErrHandle:
    Exit Sub
End Sub

Private Sub mnuEditDisplay_Click()
    '查看单据
    
    Dim strNo As String
    With mshList
        strNo = .TextMatrix(.Row, 0)
        Select Case mlngMode
        Case 1725
            frmStuffRequestPlanCard.ShowCard Me, strNo, 4, mstrPrivs
        Case 1724
            frmStuffPlanCard.ShowCard Me, strNo, 4
        End Select
    End With
End Sub

Private Sub mnuEditModify_Click()
    '修改
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    blnSuccess = False
    With mshList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        strNo = .TextMatrix(.Row, 0)
        Select Case mlngMode
        Case 1725
            frmStuffRequestPlanCard.ShowCard Me, strNo, 2, mstrPrivs, blnSuccess
        Case 1724
            frmStuffPlanCard.ShowCard Me, strNo, 2, blnSuccess
        End Select
        If blnSuccess = True Then
            mnuViewRefresh_Click
        End If
    End With
End Sub

Private Sub mnuFileBillPreview_Click()
    With mshList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        If mstrCaption = "卫材计划管理" Then
            ReportOpen gcnOracle, glngSys, "zl1_bill_1724", Me, "单据编号=" & .TextMatrix(.Row, 0), "单位=" & mintUnit, 1
        Else
            ReportOpen gcnOracle, glngSys, "zl1_bill_1725", Me, "单据编号=" & .TextMatrix(.Row, 0), "单位=" & mintUnit, 1
        End If
    End With
End Sub

Private Sub mnuFileBillPrint_Click()
    With mshList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        If mstrCaption = "卫材计划管理" Then
            ReportOpen gcnOracle, glngSys, "zl1_bill_1724", Me, "单据编号=" & .TextMatrix(.Row, 0), "单位=" & mintUnit, 2
        Else
            ReportOpen gcnOracle, glngSys, "zl1_bill_1725", Me, "单据编号=" & .TextMatrix(.Row, 0), "单位=" & mintUnit, 2
        End If
    End With
End Sub

Private Sub mnuFileExcel_Click()
    '输出到Excel
    If Me.ActiveControl Is mshList Then
        mshList.Redraw = flexRDNone
        subPrint 3
        mshList.Redraw = flexRDDirect
        mshList.Col = 0
        mshList.ColSel = mshList.Cols - 1
    ElseIf Me.ActiveControl Is mshDetail Then
        mshDetail.Redraw = flexRDNone
        subPrint 3
        mshDetail.Redraw = flexRDDirect
        mshDetail.Col = 0
        mshDetail.ColSel = mshDetail.Cols - 1
    End If
End Sub

Private Sub mnufileexit_Click()
    '退出
    Unload Me
End Sub

Private Sub mnuFileParameter_Click()
'参数设置
    Dim strReg As String
    frmParaset.设置参数 mlngMode, mstrPrivs, Me, mstrCaption
    strReg = Val(zlDatabase.GetPara("卫材单位", glngSys, mlngMode, "0"))
    mintUnit = Val(strReg)
    mlastRow = 0
  
    '刘兴宏:增加小数格式化串
    With mOraFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价, True)
        .FM_金额 = GetFmtString(mintUnit, g_金额, True)
        .FM_零售价 = GetFmtString(mintUnit, g_售价, True)
        .FM_数量 = GetFmtString(mintUnit, g_数量, True)
    End With
    mintFindDay = Val(zlDatabase.GetPara("查询天数", glngSys, mlngMode, 1))
    mdtStartDate = Format(DateAdd("d", -mintFindDay, sys.Currentdate), "yyyy-MM-dd")
    mdtEndDate = Format(sys.Currentdate, "yyyy-MM-dd")
    
    GetList (mstrFind)  '列出单据头
'    Call mshlist_EnterCell
End Sub

Private Sub mnuFilePreView_Click()
    '打印预览
    mshList.Redraw = False
    subPrint 2
    mshList.Redraw = True
    mshList.Col = 0
    mshList.ColSel = mshList.Cols - 1
    
End Sub

Private Sub mnuFilePrint_Click()
    '打印
    mshList.Redraw = False
    subPrint 1
    mshList.Redraw = True
    mshList.Col = 0
    mshList.ColSel = mshList.Cols - 1
End Sub

Private Sub mnuFilePrintSet_Click()
    '打印设置
    zlPrintSet
End Sub

Private Sub mnuHelpAbout_Click()
    '关于
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
    '帮助主题
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name)
End Sub

Private Sub mnuHelpWebHome_Click()
    '中联主页
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    '发送反馈
    Call zlMailTo(Me.hwnd)
End Sub

Private Sub mnuViewRefresh_Click()
    '刷新
    mlastRow = 0
    GetList mstrFind
End Sub

Private Sub mnuViewSearch_Click()
    '查找
    Dim strFind As String
    Dim strOthers() As String
    strFind = FrmStuffPlanSearch.GetSearch(Me, mdtStartDate, mdtEndDate, mdtVerifyStart, mdtVerifyEnd, strOthers)
    
    If strFind <> "" Then
        mstrFind = strFind
        mstrOthers = strOthers
        mlastRow = 0
        GetList mstrFind
        If Format(mdtStartDate, "yyyy-mm-dd") = "1901-01-01" And Format(mdtVerifyStart, "yyyy-mm-dd") = "1901-01-01" Then
        ElseIf Format(mdtStartDate, "yyyy-mm-dd") <> "1901-01-01" And Format(mdtVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
            PrintRange "查询范围:填制日期 " & Format(mdtStartDate, "yyyy年MM月dd日") & "至" & Format(mdtEndDate, "yyyy年MM月dd日") & "  审核日期 " & Format(mdtVerifyStart, "yyyy年MM月dd日") & "至" & Format(mdtVerifyEnd, "yyyy年MM月dd日")
        ElseIf Format(mdtStartDate, "yyyy-mm-dd") <> "1901-01-01" Then
            PrintRange "查询范围:填制日期 " & Format(mdtStartDate, "yyyy年MM月dd日") & "至" & Format(mdtEndDate, "yyyy年MM月dd日")
        ElseIf Format(mdtVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
            PrintRange "查询范围:审核日期 " & Format(mdtVerifyStart, "yyyy年MM月dd日") & "至" & Format(mdtVerifyEnd, "yyyy年MM月dd日")
        End If
     End If
End Sub

Private Sub mnuViewStatus_Click()
    With mnuViewStatus
        .Checked = Not .Checked
        stbThis.Visible = .Checked
    End With
    
    Form_Resize
End Sub
Private Sub mnuReportItem_Click(Index As Integer)
    Dim strNo As String
    Dim lng库房ID As Long
    
    With mshList
        strNo = Trim(.TextMatrix(.Row, 0))
    End With
    
    If cboStock.ListIndex < 0 Then
        lng库房ID = 0
    Else
        lng库房ID = cboStock.ItemData(cboStock.ListIndex)
    End If
    
    If mlngMode = 1725 Then
        '2006-04-25:刘兴宏:增加自定义报表发布到模块的功能
        If Format(mdtStartDate, "yyyy-mm-dd") = "1990-01-01" Then
            Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, "NO=" & strNo, "部门id=" & lng库房ID)
        Else
            Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, "NO=" & strNo, "部门id=" & lng库房ID, "开始时间=" & Format(mdtStartDate, "yyyy-mm-dd"), "结束时间=" & Format(mdtEndDate, "yyyy-mm-dd"))
        End If
    Else
        '2006-04-25:刘兴宏:增加自定义报表发布到模块的功能
        If Format(mdtStartDate, "yyyy-mm-dd") = "1990-01-01" Then
            Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, "NO=" & strNo, "库房=" & lng库房ID)
        Else
            Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, "NO=" & strNo, "库房=" & lng库房ID, "开始时间=" & Format(mdtStartDate, "yyyy-mm-dd"), "结束时间=" & Format(mdtEndDate, "yyyy-mm-dd"))
        End If
    End If
End Sub
Private Sub mnuViewToolButton_Click()
    With mnuViewToolButton
        .Checked = Not .Checked
        cbrTool.Bands(1).Visible = .Checked
        mnuViewToolText.Enabled = .Checked
    End With
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim intCount As Integer      '工具条索引
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    With tlbTool.Buttons
        If mnuViewToolText.Checked = False Then
            '取消所有的文本标签显示
            For intCount = 1 To .count
                .Item(intCount).Caption = ""
            Next
        Else
            '让所有的文本标签显示。说明：Tag中放的文本标签
            For intCount = 1 To .count
                .Item(intCount).Caption = .Item(intCount).Tag
            Next
        End If
    End With
    
    cbrTool.Bands(1).MinHeight = tlbTool.Height
    
    Form_Resize
End Sub


Private Sub mshList_Click()
    With mshList
         If .Row < 1 Then Exit Sub
         If .MouseRow = 0 Then
            ListSort
            Exit Sub
         End If
    End With
End Sub

Private Sub mshlist_DblClick()
    If mnuEditModify.Visible = False Then Exit Sub
    If mnuEditModify.Enabled = False Then Exit Sub
    If mshList.MouseRow = 0 Then Exit Sub
    mnuEditModify_Click
End Sub

Private Sub mshlist_EnterCell()
    Dim rsTemp As New Recordset
    Dim IntBill As Integer                      '单据类型  如：1、外购入库；2、
    Dim strUnit As String                       '单位名称:如门诊单位，住院单位等
    Dim str包装系数 As String
    
    mlastRow = mshList.Row
    
    On Error GoTo ErrHandle
    If mshList.Row >= 1 And LTrim(mshList.TextMatrix(mshList.Row, 0)) <> "" Then
        mshList.Col = 0
        mshList.ColSel = mshList.Cols - 1

        mshDetail.Redraw = False
        Select Case mintUnit
            Case 0
                str包装系数 = "1"
            Case Else
                str包装系数 = "D.换算系数"
        End Select
        If mlngMode = 1725 Then
            gstrSQL = "" & _
                "   SELECT M.编码,M.名称 as 通用名称, M.规格," & IIf(mintUnit = 0, "M.计算单位", "D.包装单位") & " as  单位," & _
                "           trim(to_char(b.请购数量 /" & str包装系数 & "," & mOraFMT.FM_数量 & ")) 请购数量," & _
                "           trim(to_char(b.上期销量 /" & str包装系数 & "," & mOraFMT.FM_数量 & ")) 上期销量," & _
                "           trim(to_char(b.本期销量 /" & str包装系数 & "," & mOraFMT.FM_数量 & ")) 本期销量," & _
                "           trim(to_char(b.计划数量 /" & str包装系数 & "," & mOraFMT.FM_数量 & ")) 审批数量," & _
                "           trim(to_char(b.单价 *" & str包装系数 & "," & mOraFMT.FM_成本价 & ")) 成本价," & _
                "           trim(to_char(b.金额," & mOraFMT.FM_金额 & ")) 成本金额, b.上次供应商,b.上次生产商 " & _
                "   FROM 材料采购计划 a, 材料计划内容 b,部门表 c,材料特性 D,收费项目目录 M" & _
                "   Where a.id = b.计划id " & _
                "           and nvl(a.库房id,0)=c.id(+) " & _
                "           and b.材料id=d.材料id and b.材料id=M.id " & _
                "           AND b.计划ID =[1]" & _
                "   Order by 序号"
        Else
            gstrSQL = "" & _
                "   SELECT M.编码,M.名称 as 通用名称, M.规格," & IIf(mintUnit = 0, "M.计算单位", "D.包装单位") & " as  单位," & _
                "           trim(to_char(b.前期数量 /" & str包装系数 & "," & mOraFMT.FM_数量 & ")) 前期数量," & _
                "           trim(to_char(b.上期数量 /" & str包装系数 & "," & mOraFMT.FM_数量 & ")) 上期数量," & _
                "           trim(to_char(b.库存数量 /" & str包装系数 & "," & mOraFMT.FM_数量 & ")) 库存数量," & _
                "           trim(to_char(b.上期销量 /" & str包装系数 & "," & mOraFMT.FM_数量 & ")) 上期销量," & _
                "           trim(to_char(b.本期销量 /" & str包装系数 & "," & mOraFMT.FM_数量 & ")) 本期销量," & _
                "           trim(to_char(b.计划数量 /" & str包装系数 & "," & mOraFMT.FM_数量 & ")) 计划数量," & _
                "           trim(to_char(b.单价 *" & str包装系数 & "," & mOraFMT.FM_成本价 & ")) 成本价," & _
                "           trim(to_char(b.金额," & mOraFMT.FM_金额 & ")) 成本金额, b.上次供应商 供应商,b.上次生产商 " & _
                "   FROM 材料采购计划 a, 材料计划内容 b,部门表 c,材料特性 D,收费项目目录 M" & _
                "   Where a.id = b.计划id " & _
                "           and nvl(a.库房id,0)=c.id(+) " & _
                "           and b.材料id=d.材料id and b.材料id=M.id " & _
                "           AND b.计划ID =[1] " & _
                "   Order by 序号"
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mshList.TextMatrix(mshList.Row, 1)))
        Set mshDetail.Recordset = rsTemp
        
        With mshDetail
            If .Rows = 1 Then
                .Rows = .Rows + 100
                .Row = 1
                .Redraw = True

                .TopRow = 1
                .Rows = .Rows - 99
            End If
            .Row = 1
            .Col = 0
            .ColSel = .Cols - 1
        End With

        mshDetail.Redraw = True
    Else
        If mlngMode = 1725 Then
            With mshDetail
                .Cols = 12
                .Rows = 2
                .Clear
                .TextMatrix(0, 0) = "编码"
                .TextMatrix(0, 1) = "部门"
                .TextMatrix(0, 2) = "规格"
                .TextMatrix(0, 3) = "单位"
                .TextMatrix(0, 4) = "上期销量"
                .TextMatrix(0, 5) = "本期销量"
                .TextMatrix(0, 6) = "请购数量"
                .TextMatrix(0, 7) = "审批数量"
                .TextMatrix(0, 8) = "成本价"
                .TextMatrix(0, 9) = "成本金额"
                .TextMatrix(0, 10) = "上次供应商"
                .TextMatrix(0, 11) = "上次生产商"
                .Row = 1
                .Col = 0
                .ColSel = .Cols - 1
            End With
        Else
            With mshDetail
                .Cols = 14
                .Rows = 2
                .Clear
                .TextMatrix(0, 0) = "编码"
                .TextMatrix(0, 1) = "名称"
                .TextMatrix(0, 2) = "规格"
                .TextMatrix(0, 3) = "单位"
                .TextMatrix(0, 4) = "前期数量"
                .TextMatrix(0, 5) = "上期数量"
                .TextMatrix(0, 6) = "库存数量"
                .TextMatrix(0, 7) = "上期销量"
                .TextMatrix(0, 8) = "本期销量"
                .TextMatrix(0, 9) = "计划数量"
                .TextMatrix(0, 10) = "成本价"
                .TextMatrix(0, 11) = "成本金额"
                .TextMatrix(0, 12) = "供应商"
                .TextMatrix(0, 13) = "上次生产商"
    
                .Row = 1
                .Col = 0
                .ColSel = .Cols - 1
            End With
        End If
    End If
    SetDetailColWidth
    SetEnable
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetDetailColWidth()
    Dim intCol As Integer
    
    With mshDetail
        If mblnBootUp = False Then
            .ColWidth(0) = 0
            .ColWidth(1) = 2500
            For intCol = 2 To .Cols - 1
                .ColWidth(intCol) = 1000
            Next
        End If
        
        Select Case mlngMode
            Case 1725
                For intCol = 0 To .Cols - 1
                    .ColAlignment(intCol) = 1
                    If InStr(1, .TextMatrix(0, intCol), "数量") > 0 Or InStr(1, .TextMatrix(0, intCol), "销量") > 0 Then
                        .ColAlignment(intCol) = flexAlignRightCenter     '数量和销量
                    End If
                    If InStr(1, .TextMatrix(0, intCol), "成本价") > 0 Then
                        .ColAlignment(intCol) = flexAlignRightCenter     '单价
                        .ColWidth(intCol) = IIf(mblnCostView = False, 0, 1000)
                    End If
                    If InStr(1, .TextMatrix(0, intCol), "成本金额") > 0 Then
                        .ColAlignment(intCol) = flexAlignRightCenter     '金额
                        .ColWidth(intCol) = IIf(mblnCostView = False, 0, 1000)
                    End If
                    If InStr(1, .TextMatrix(0, intCol), "上次供应商") > 0 Then
                        .ColAlignment(intCol) = flexAlignLeftCenter     '上次供应商
                        .ColWidth(intCol) = IIf(mblnProvider = False, 0, 1000)
                    End If
                    .ColAlignmentFixed(intCol) = 4
                Next
            Case 1724
                For intCol = 0 To .Cols - 1
                    .ColAlignment(intCol) = 1
                    If InStr(1, .TextMatrix(0, intCol), "数量") > 0 Or InStr(1, .TextMatrix(0, intCol), "销量") > 0 Then
                        .ColAlignment(intCol) = flexAlignRightCenter     '数量和销量
                    End If
                    If InStr(1, .TextMatrix(0, intCol), "成本价") > 0 Then
                        .ColAlignment(intCol) = flexAlignRightCenter     '单价
                        .ColWidth(intCol) = IIf(mblnCostView = False, 0, 1000)
                    End If
                    If InStr(1, .TextMatrix(0, intCol), "成本金额") > 0 Then
                        .ColAlignment(intCol) = flexAlignRightCenter     '金额
                        .ColWidth(intCol) = IIf(mblnCostView = False, 0, 1000)
                    End If
                    .ColAlignmentFixed(intCol) = 4
                Next
        End Select
    End With
End Sub

Private Sub mshlist_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If mnuEditModify.Visible = False Then Exit Sub
        If mnuEditModify.Enabled = False Then Exit Sub
        mnuEditModify_Click
    End If
        
End Sub

Private Sub mshlist_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    If mnuEdit.Visible = False Then Exit Sub
    PopupMenu mnuEdit, 2
End Sub

Private Sub picSeparate_s_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
        mintOldY = Y
End Sub

Private Sub picSeparate_s_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '分割条设置
    
    If Button <> 1 Then Exit Sub
    
    With picSeparate_s
        If .Top + Y < 2000 Then Exit Sub
        If .Top + Y > ScaleHeight - 2000 Then Exit Sub
        .Move .Left, .Top + Y - mintOldY
    End With
    
    With mshList
        .Height = picSeparate_s.Top - .Top
    End With
    
    With Cmd查阅
        .Top = mshList.Top + mshList.Height + 30
    End With
    
    With mshDetail
        .Top = picSeparate_s.Top + picSeparate_s.Height + 100
        .Height = ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
    End With
End Sub

Private Sub picSeparate_s_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
        mintOldY = 0
End Sub

Private Sub tabShow_Click(PreviousTab As Integer)
    Call GetList(mstrFind)    '列出单据头
End Sub

Private Sub tlbTool_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "PrintView"
            mnuFilePreView_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Add"
            mnuEditAdd_Click
        Case "Modify"
            mnuEditModify_Click
        Case "Delete"
            mnuEditDel_Click
        Case "Verify"
            mnuEditVerify_Click
        Case "Cancel"
            mnuEditCancel_Click
        Case "Clear"
            mnuEditClear_Click
        Case "Search"
            mnuViewSearch_Click
        Case "Refresh"
            mnuViewRefresh_Click
        Case "Help"
            mnuHelpTitle_Click
        Case "Exit"
            mnufileexit_Click
    End Select
End Sub

'设置菜单和工具按钮的可用属性
Private Sub SetEnable()
    With mshList
        .ToolTipText = ""
        If .TextMatrix(.Row, 0) = "" Or .Row = 0 Then         '没有单
            mnuFilePreView.Enabled = False
            mnuFilePrint.Enabled = False
            mnuFileBillPreview.Enabled = False
            mnuFileBillPrint.Enabled = False
            mnuFileExcel.Enabled = False
            tlbTool.Buttons("Print").Enabled = False
            tlbTool.Buttons("PrintView").Enabled = False
            
            If mnuEditModify.Visible = True Then
                mnuEditModify.Enabled = False
                tlbTool.Buttons("Modify").Enabled = False
            End If
            If mnuEditDel.Visible = True Then
                mnuEditDel.Enabled = False
                tlbTool.Buttons("Delete").Enabled = False
            End If
            If mnuEditVerify.Visible = True Then
                mnuEditVerify.Enabled = False
                tlbTool.Buttons("Verify").Enabled = False
            End If
            
            If mnuEditCancel.Visible = True Then
                mnuEditCancel.Enabled = False
                tlbTool.Buttons("Cancel").Enabled = False
            End If
            
            If mnuEditClear.Visible = True Then
                mnuEditClear.Enabled = False
                tlbTool.Buttons("Clear").Enabled = False
            End If
            
            If mnuEditDisplay.Visible = True Then
                mnuEditDisplay.Enabled = False
            End If
            Cmd查阅.Enabled = False
        Else
            Cmd查阅.Enabled = True
            mnuFilePreView.Enabled = True
            mnuFilePrint.Enabled = True
            mnuFileBillPreview.Enabled = True
            mnuFileBillPrint.Enabled = True
            mnuFileExcel.Enabled = True
            tlbTool.Buttons("Print").Enabled = True
            tlbTool.Buttons("PrintView").Enabled = True
            
            If .TextMatrix(.Row, .Cols - 3) = "" Then    '未审核单
                If mnuEditModify.Visible = True Then
                    mnuEditModify.Enabled = True
                    tlbTool.Buttons("Modify").Enabled = True
                End If
                If mnuEditDel.Visible = True Then
                    mnuEditDel.Enabled = True
                    tlbTool.Buttons("Delete").Enabled = True
                End If
                If mnuEditVerify.Visible = True Then
                    mnuEditVerify.Enabled = True
                    tlbTool.Buttons("Verify").Enabled = True
                End If
                    
                If mnuEditCancel.Visible = True Then
                    mnuEditCancel.Enabled = False
                    tlbTool.Buttons("Cancel").Enabled = False
                End If
                
                If mnuEditClear.Visible = True Then
                    mnuEditClear.Enabled = False
                    tlbTool.Buttons("Clear").Enabled = False
                End If
                
                If mnuEditDisplay.Visible = True Then
                    mnuEditDisplay.Enabled = True
                End If
            Else    '审核单
                If mnuEditModify.Visible = True Then
                    mnuEditModify.Enabled = False
                    tlbTool.Buttons("Modify").Enabled = False
                End If
                If mnuEditDel.Visible = True Then
                    mnuEditDel.Enabled = False
                    tlbTool.Buttons("Delete").Enabled = False
                End If
                If mnuEditVerify.Visible = True Then
                    mnuEditVerify.Enabled = False
                    tlbTool.Buttons("Verify").Enabled = False
                End If
                If mnuEditCancel.Visible = True Then
                    mnuEditCancel.Enabled = True
                    tlbTool.Buttons("Cancel").Enabled = True
                End If
                
                If mnuEditClear.Visible = True Then
                    mnuEditClear.Enabled = True
                    tlbTool.Buttons("Clear").Enabled = True
                End If
                
                If mnuEditDisplay.Visible = True Then
                    mnuEditDisplay.Enabled = True
                End If
            End If
        End If
    End With
End Sub

Private Sub subPrint(bytMode As Byte)
    '功能:进行打印,预览和输出到EXCEL
    '参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    Dim objPrint As Object
    Dim objRow As New zlTabAppRow
    Dim strRange As String
    
    
    If Format(mdtStartDate, "yyyy-mm-dd") = "1901-01-01" And Format(mdtVerifyStart, "yyyy-mm-dd") = "1901-01-01" Then
        strRange = "审核日期 " & Format(mdtVerifyStart, "yyyy年MM月dd日") & "至" & Format(mdtVerifyEnd, "yyyy年MM月dd日")
    ElseIf Format(mdtVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
        strRange = "填制日期 " & Format(mdtStartDate, "yyyy年MM月dd日") & "至" & Format(mdtEndDate, "yyyy年MM月dd日") & "  审核日期 " & Format(mdtVerifyStart, "yyyy年MM月dd日") & "至" & Format(mdtVerifyEnd, "yyyy年MM月dd日")
    Else
        strRange = "填制日期 " & Format(mdtStartDate, "yyyy年MM月dd日") & "至" & Format(mdtEndDate, "yyyy年MM月dd日")
    End If
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
    objPrint.Title.Text = Me.Caption
        
    objRow.Add "时间：" & strRange
    objRow.Add "部门：" & cboStock.Text
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow

    objRow.Add "打印人:" & UserInfo.用户名
    objRow.Add "打印日期:" & Format(sys.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    
    If Me.ActiveControl Is mshList Then
        Set objPrint.Body = mshList
    Else
        Set objPrint.Body = mshDetail
    End If
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub tlbTool_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
        Case "Payment"
        Case "Imprest"
    End Select
End Sub

Private Sub tlbTool_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool
    End If
End Sub

'对单据头列排序
Private Sub ListSort()
    Dim intCol As Integer
    Dim intRow As Integer
    Dim intTemp As String

    With mshList
        If .Rows > 1 Then
'            .Redraw = False
'            intCol = .MouseCol
'            .Col = intCol
'            .ColSel = intCol
'            intTemp = .TextMatrix(.Row, 0)
'            If intCol = mintPreCol And mintsort = flexSortStringNoCaseDescending Then
'               .Sort = flexSortStringNoCaseAscending
'               mintsort = flexSortStringNoCaseAscending
'            Else
'               .Sort = flexSortStringNoCaseDescending
'               mintsort = flexSortStringNoCaseDescending
'            End If
'            mintPreCol = intCol
'            .Row = Grid.MshGrdFindRow(mshList, intTemp, 0)
'            If .RowPos(.Row) + .RowHeight(.Row) > .Height Then
'                .TopRow = .Row
'            Else
'                .TopRow = 1
'            End If
            .Col = 0
            .ColSel = .Cols - 1
            .Redraw = True
            .SetFocus
        Else
            .ColSel = 0
        End If
    End With
End Sub

Private Sub PrintRange(ByVal strRange As String)
    '功能:打印时间范围
    picSeparate_s.Cls
    picSeparate_s.CurrentX = 50
    picSeparate_s.CurrentY = 100
    picSeparate_s.Print strRange
End Sub
Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

