VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmDrugQualityList 
   Caption         =   "药品质量管理"
   ClientHeight    =   4980
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9750
   Icon            =   "frmDrugQualityList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin ComCtl3.CoolBar cbrTool 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9750
      _ExtentX        =   17198
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   9750
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tlbTool"
      MinWidth1       =   6000
      MinHeight1      =   720
      Width1          =   6210
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Caption2        =   "库房"
      Child2          =   "cboStock"
      MinHeight2      =   300
      Width2          =   3615
      NewRow2         =   0   'False
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   6825
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   2835
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
            NumButtons      =   15
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
               Caption         =   "处理"
               Key             =   "Verify"
               Description     =   "处理"
               Object.ToolTipText     =   "处理"
               Object.Tag             =   "处理"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "清除"
               Key             =   "Clear"
               Description     =   "清除"
               Object.ToolTipText     =   "清除"
               Object.Tag             =   "清除"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "VerifySeparate"
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "过滤"
               Key             =   "Search"
               Description     =   "过滤"
               Object.ToolTipText     =   "过滤"
               Object.Tag             =   "过滤"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "刷新"
               Key             =   "Refresh"
               Description     =   "刷新"
               Object.ToolTipText     =   "刷新"
               Object.Tag             =   "刷新"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FindSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助主题"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Exit"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   11
            EndProperty
         EndProperty
         MouseIcon       =   "frmDrugQualityList.frx":014A
      End
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   4620
      Width           =   9750
      _ExtentX        =   17198
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDrugQualityList.frx":0464
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12118
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
      Left            =   2280
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQualityList.frx":0CF8
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQualityList.frx":0F14
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQualityList.frx":1130
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQualityList.frx":134A
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQualityList.frx":1564
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQualityList.frx":177E
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQualityList.frx":1998
            Key             =   "Clear"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQualityList.frx":2092
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQualityList.frx":22AC
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQualityList.frx":2406
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQualityList.frx":2622
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHot 
      Left            =   1545
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQualityList.frx":283E
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQualityList.frx":2A5A
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQualityList.frx":2C76
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQualityList.frx":2E90
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQualityList.frx":30AC
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQualityList.frx":32C6
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQualityList.frx":34E0
            Key             =   "Clear"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQualityList.frx":3BDA
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQualityList.frx":3DF4
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQualityList.frx":3F4E
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQualityList.frx":416A
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   1965
      Left            =   360
      TabIndex        =   5
      Top             =   1320
      Width           =   6255
      _cx             =   11033
      _cy             =   3466
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
      BackColorSel    =   16769992
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDrugQualityList.frx":4386
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
      ExplorerBar     =   1
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
      VirtualData     =   0   'False
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
   Begin VB.Label lblRange 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "查询范围:1999年8月12日至1999年9月12日"
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   4200
      Width           =   3330
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
         Caption         =   "处理(&V)"
      End
      Begin VB.Menu mnuEditClear 
         Caption         =   "清除(&C)"
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
Attribute VB_Name = "frmDrugQualityList"
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
Private mstrPrivs As String
Private mblnViewCost As Boolean         '查看成本价 true-允许查看 flase-不允许查看
Private Const MStrCaption As String = "药品质量管理"

Private mbln中药库房 As Boolean

Private strStart As Date
Private strEnd As Date
Private strVerifyStart As Date
Private strVerifyEnd As Date

Private Type Type_SQLCondition
    date填制时间开始 As Date
    date填制时间结束 As Date
    date审核时间开始 As Date
    date审核时间结束 As Date
    lng药品 As Long
    lng供应商 As Long
    str填制人 As String
    str审核人 As String
End Type

Private SQLCondition As Type_SQLCondition

Private mlng库房id As Long
Private mintUnit As Integer                 '单位系数：1-售价;2-门诊;3-住院;4-药库

'从参数表中取药品价格、数量、金额小数位数（显示精度）
Private mintShowCostDigit As Integer            '成本价小数位数
Private mintShowPriceDigit As Integer           '售价小数位数
Private mintShowNumberDigit As Integer          '数量小数位数
Private mintShowMoneyDigit As Integer           '金额小数位数

Private mstrCostFormat As String
Private mstrPriceFormat As String
Private mstrNumberFormat As String
Private mstrMoneyFormat As String

Private Const mconint售价单位 As Integer = 1
Private Const mconint门诊单位 As Integer = 2
Private Const mconint住院单位 As Integer = 3
Private Const mconint药库单位 As Integer = 4
'检查数据依赖性
Private Function CheckDepend() As Boolean
    Dim rsDepend As New Recordset
    
    CheckDepend = False
    On Error GoTo errHandle
    If InStr(mstrPrivs, "所有库房") = 0 Then
        gstrSQL = "SELECT DISTINCT a.id, a.名称 " _
                & "FROM 部门性质说明 c, 部门性质分类 b, 部门表 a, 部门人员 d " _
                & "Where (a.站点 = [1] Or a.站点 is Null) And c.工作性质 = b.名称 " _
                & "  AND Instr('HIJKLMN',b.编码,1) > 0 " _
                & "  AND a.id = c.部门id AND a.id=d.部门id and d.人员id=[2] " _
                & "  AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01'"
    Else
        gstrSQL = "SELECT DISTINCT a.id, a.名称 " _
                & "FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " _
                & "Where (a.站点 = [1] Or a.站点 is Null) And c.工作性质 = b.名称 " _
                & "  AND Instr('HIJKLMN',b.编码,1) > 0 " _
                & "  AND a.id = c.部门id " _
                & "  AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01'"
    End If
    Set rsDepend = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption, gstrNodeNo, UserInfo.用户ID)
    
    If rsDepend.EOF Then
        MsgBox "至少应该设置一个具有药库性质，药房性质，或者制剂室性质的部门,请查看部门管理！", vbInformation, gstrSysName
        rsDepend.Close
        Exit Function
    End If
            
    With cboStock
        .Clear
        Do While Not rsDepend.EOF
            .AddItem rsDepend!名称
            .ItemData(.NewIndex) = rsDepend!id
            If rsDepend!id = UserInfo.部门ID Then
                .ListIndex = .NewIndex
            End If
            rsDepend.MoveNext
        Loop
        rsDepend.Close
        If .ListCount > 0 Then
            .ListIndex = 0
        End If
    End With
    CheckDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function Get出库单据号(ByVal lngBillId As Long)
    Dim rstemp As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select 出库单No From 药品质量记录 Where Id = [1] "
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "取出库单据号", lngBillId)
    
    If Not rstemp.EOF Then
        Get出库单据号 = IIf(IsNull(rstemp!出库单no), "", rstemp!出库单no)
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub cboStock_Click()
    If mlng库房id <> Me.cboStock.ItemData(Me.cboStock.ListIndex) Then
        mlng库房id = Me.cboStock.ItemData(Me.cboStock.ListIndex)
        Call GetDrugDigit(mlng库房id, MStrCaption, mintUnit, mintShowCostDigit, mintShowPriceDigit, mintShowNumberDigit, mintShowMoneyDigit)
        
        '组织格式化串
        mstrCostFormat = "'999999999990." & String(mintShowCostDigit, "0") & "'"
        mstrPriceFormat = "'999999999990." & String(mintShowPriceDigit, "0") & "'"
        mstrNumberFormat = "'999999999990." & String(mintShowNumberDigit, "0") & "'"
        mstrMoneyFormat = "'999999999990." & String(mintShowMoneyDigit, "0") & "'"
    
        If mblnBootUp Then mnuViewRefresh_Click
    End If
End Sub

Private Sub cbrTool_Resize()
    Form_Resize
End Sub

Private Sub GetList(ByVal strFind As String)
    Dim rsList As ADODB.Recordset
    Dim strUnit As String
    Dim str包装系数 As String
    Dim str比例系数 As String
    Dim strSql药名 As String
    Dim n As Integer
    Dim str库房性质 As String
    
    On Error GoTo errHandle
    Call FS.ShowFlash("正在搜索药品质量管理记录,请稍候 ...", Me)
    DoEvents
    Screen.MousePointer = vbHourglass
    
    vsfList.Redraw = False
    
    mbln中药库房 = False
    str库房性质 = ""
    gstrSQL = "Select a.工作性质 From 部门性质说明 A Where a.部门id =[1]"
    Set rsList = zldatabase.OpenSQLRecord(gstrSQL, "判断是库房性质", cboStock.ItemData(cboStock.ListIndex))
    Do While Not rsList.EOF
        str库房性质 = str库房性质 & "," & rsList!工作性质
        rsList.MoveNext
    Loop
    If str库房性质 Like "*中药*" Or str库房性质 Like "*制剂室*" Then mbln中药库房 = True
    
    Select Case mintUnit
        Case mconint售价单位
            strUnit = "F.计算单位"
            str包装系数 = "1 as 比例系数 "
            str比例系数 = "1 "
        Case mconint门诊单位
            strUnit = "B.门诊单位"
            str包装系数 = "B.门诊包装 as 比例系数 "
            str比例系数 = "B.门诊包装 "
        Case mconint住院单位
            strUnit = "B.住院单位"
            str包装系数 = "B.住院包装 as 比例系数 "
            str比例系数 = "B.住院包装 "
        Case mconint药库单位
            strUnit = "B.药库单位"
            str包装系数 = "B.药库包装 as 比例系数 "
            str比例系数 = "B.药库包装 "
    End Select
        
    If gint药品名称显示 = 0 Then
        strSql药名 = ",('['||F.编码||']'||F.名称) AS 药品信息"
    ElseIf gint药品名称显示 = 1 Then
        strSql药名 = ",('['||F.编码||']'||NVL(D.名称,F.名称)) AS 药品信息"
    Else
        strSql药名 = ",('['||F.编码||']'||F.名称) AS 药品信息,D.名称 As 商品名"
    End If
        
    gstrSQL = "SELECT DISTINCT A.ID" & strSql药名 & _
        ",A.药品ID,A.批次,A.供药单位ID,F.规格,A.产地 as 生产商," & IIf(mbln中药库房, "B.原产地,", "") & "A.批号," & _
        " ltrim(to_char(a.成本单价*" & str比例系数 & "," & mstrCostFormat & ")) as 成本价," & _
        " ltrim(to_char(a.成本金额," & mstrMoneyFormat & ")) as 成本金额," & _
        " ltrim(to_char(a.销售单价*" & str比例系数 & "," & mstrPriceFormat & ")) as 零售价," & _
        " ltrim(to_char(a.销售金额," & mstrMoneyFormat & ")) as 销售金额," & _
        strUnit & " AS 单位,LTRIM(TO_CHAR(A.毁损数量/(" & str比例系数 & ")," & mstrNumberFormat & ")) AS 毁损数量," & str包装系数 & _
        " ,A.毁损原因,C.名称 AS 供应商,A.登记人,TO_CHAR(A.登记时间,'YYYY-MM-DD') AS 登记时间,A.解决办法,A.处理人,TO_CHAR(A.处理时间,'YYYY-MM-DD') AS 处理时间 " & _
        " FROM 药品质量记录 A, 药品规格 B, 收费项目目录 F, 收费项目别名 D, 供应商 C " & _
        " WHERE A.药品ID = B.药品ID And B.药品ID=F.ID " & _
        " AND F.id = D.收费细目ID(+) AND A.供药单位ID = C.ID(+) And D.性质(+)=3 And D.码类(+)=1" & _
        " AND SUBSTR(类型(+),1,1)='1' " & _
        " AND A.库房ID = [9] " & _
        strFind & _
        " ORDER BY 登记时间 DESC "

    Set rsList = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption, _
        SQLCondition.date填制时间开始, _
        SQLCondition.date填制时间结束, _
        SQLCondition.date审核时间开始, _
        SQLCondition.date审核时间结束, _
        SQLCondition.lng药品, _
        SQLCondition.lng供应商, _
        SQLCondition.str填制人, _
        SQLCondition.str审核人, _
        cboStock.ItemData(cboStock.ListIndex))
        
    Set vsfList.DataSource = rsList
    With vsfList
        If .rows = 1 Then
            .rows = .rows + 100
            .Row = 1
            .Redraw = True
            .TopRow = 1
            .rows = .rows - 99
        End If
        .Row = 1
        .Col = 0
        .ColSel = .Cols - 1
        
        For n = 0 To .Cols - 1
            .FixedAlignment(n) = flexAlignCenterCenter
        Next
    End With
    SetListColWidth
    
    vsfList.Redraw = True
    Call FS.StopFlash
    Screen.MousePointer = vbDefault
    SetEnable
    staThis.Panels(2).Text = "当前共有" & rsList.RecordCount & "张单据"
    rsList.Close
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'表头列宽初始
Private Sub SetListColWidth()
    Dim intCol As Integer
    
    With vsfList
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(8 + IIf(gint药品名称显示 = 2, 1, 0) + IIf(mbln中药库房, 1, 0)) = flexAlignRightCenter
        .ColAlignment(9 + IIf(gint药品名称显示 = 2, 1, 0) + IIf(mbln中药库房, 1, 0)) = flexAlignRightCenter
        .ColAlignment(10 + IIf(gint药品名称显示 = 2, 1, 0) + IIf(mbln中药库房, 1, 0)) = flexAlignRightCenter
        .ColAlignment(11 + IIf(gint药品名称显示 = 2, 1, 0) + IIf(mbln中药库房, 1, 0)) = flexAlignRightCenter
        .ColAlignment(13 + IIf(gint药品名称显示 = 2, 1, 0) + IIf(mbln中药库房, 1, 0)) = flexAlignRightCenter
        
        If mblnBootUp = False Then
'            For intCol = 0 To .Cols - 1
'                .ColWidth(intCol) = 1500
'            Next
            .ColWidth(1) = 2000
            If gint药品名称显示 = 2 Then .ColWidth(2) = 2000
            .ColWidth(5 + IIf(gint药品名称显示 = 2, 1, 0)) = 2000
        End If
        If mblnViewCost = False Then
            .ColWidth(9 + IIf(mbln中药库房, 1, 0)) = 0 '成本价
            .ColWidth(10 + IIf(mbln中药库房, 1, 0)) = 0 '成本金额
        End If
        
        .ColWidth(0) = 0
        .ColWidth(2 + IIf(gint药品名称显示 = 2, 1, 0)) = 0
        .ColWidth(3 + IIf(gint药品名称显示 = 2, 1, 0)) = 0
        .ColWidth(4 + IIf(gint药品名称显示 = 2, 1, 0)) = 0
        .ColWidth(14 + IIf(gint药品名称显示 = 2, 1, 0) + IIf(mbln中药库房, 1, 0)) = 0
        
    End With
End Sub

'根据权限设置不同的显示项目
Private Sub SetVisable()
    '外购入库所有权限：参数设置、基本、所有库房、登记、修改、删除、验收、冲销
    
    If Not zlStr.IsHavePrivs(mstrPrivs, "质量登记") Then
        mnuEditAdd.Visible = False
        mnuEditModify.Visible = False
        mnuEditDel.Visible = False
        
        mnuEditLine1.Visible = False
        
        tlbTool.Buttons("Add").Visible = False
        tlbTool.Buttons("Modify").Visible = False
        tlbTool.Buttons("Delete").Visible = False
        
        tlbTool.Buttons("EditSeparate").Visible = False
        
    End If
    
    If Not zlStr.IsHavePrivs(mstrPrivs, "解决处理") Then
        mnuEditVerify.Visible = False
        tlbTool.Buttons("Verify").Visible = False
    End If
    
    If Not zlStr.IsHavePrivs(mstrPrivs, "清除记录") Then
        mnuEditClear.Visible = False
        tlbTool.Buttons("Clear").Visible = False
         '对没有所有编辑权限时，把菜单和工具栏上的相应的分割线屏蔽。
        If mnuEditVerify.Visible = False Then
            mnuEditLine2.Visible = False
            tlbTool.Buttons("VerifySeparate").Visible = False
        End If
    End If
    
        
End Sub


Private Sub Form_Load()
    '恢复设置
    Dim strStart As String
    Dim strEnd As String
    Dim strFind As String
    Dim dateCurrentDate As Date
    Dim strTemp As String
    Dim int查询天数 As Integer
    
    mlngMode = glngModul
    mstrPrivs = gstrprivs
    mblnViewCost = zlStr.IsHavePrivs(mstrPrivs, "查看成本价")
    
    mblnBootUp = False
    If Not CheckDepend Then
        Unload Me
        Exit Sub
    End If
    
    mlng库房id = Me.cboStock.ItemData(Me.cboStock.ListIndex)
    Call GetDrugDigit(mlng库房id, MStrCaption, mintUnit, mintShowCostDigit, mintShowPriceDigit, mintShowNumberDigit, mintShowMoneyDigit)
    
    '组织格式化串
    mstrCostFormat = "'999999999990." & String(mintShowCostDigit, "0") & "'"
    mstrPriceFormat = "'999999999990." & String(mintShowPriceDigit, "0") & "'"
    mstrNumberFormat = "'999999999990." & String(mintShowNumberDigit, "0") & "'"
    mstrMoneyFormat = "'999999999990." & String(mintShowMoneyDigit, "0") & "'"
    
    SetVisable  '根据权限设置不同的显示项目
        
    dateCurrentDate = Sys.Currentdate
    int查询天数 = Val(zldatabase.GetPara("查询天数", glngSys, mlngMode, 7)) - 1
    strStart = Format(DateAdd("d", -int查询天数, dateCurrentDate), "yyyy-MM-dd")
    strEnd = Format(dateCurrentDate, "yyyy-MM-dd")
    
    strFind = " AND A.处理时间 is Null And A.登记时间 Between [1] And [2] "
    SQLCondition.date填制时间开始 = CDate(Format(strStart, "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date填制时间结束 = CDate(Format(strEnd, "yyyy-mm-dd") & " 23:59:59")
    
    strVerifyStart = "1901-01-01"
    strVerifyEnd = "1901-01-01"
    
    mstrFind = strFind
    
    lblRange.Caption = "查询范围:" & Format(dateCurrentDate, "yyyy年MM月dd日") & "至" & Format(dateCurrentDate, "yyyy年MM月dd日")
    GetList (mstrFind)  '列出单据头
    '恢复个性化设置
    RestoreWinState Me, App.ProductName, MStrCaption
    '恢复个性化设置后，权限控制的列还需要进一步控制
    vsfList.ColWidth(9) = IIf(mblnViewCost = True, 1000, 0) '成本价
    vsfList.ColWidth(10) = IIf(mblnViewCost = True, 1400, 0) '成本金额
    
    Call zldatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs)
    
    mblnBootUp = True
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
    
    With lblRange
        .Top = IIf(staThis.Visible = True, Me.ScaleHeight - staThis.Height - .Height - 100, Me.ScaleHeight - .Height - 100)
        .Left = 0
        .Width = cbrTool.Width
    End With
    
    With vsfList
        .Top = IIf(cbrTool.Visible, cbrTool.Height, 0)
        .Left = 0
        .Width = cbrTool.Width
        .Height = lblRange.Top - .Top - 50
    End With
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, MStrCaption
   
End Sub


Private Sub mnuEditAdd_Click()
    Dim strNo As String
    Dim BlnSuccess As Boolean
    
    strNo = ""
    '新增
    BlnSuccess = frmDrugQualityCard.ShowCard(Me, 1, 0, mstrPrivs)
    
    
    If BlnSuccess = True Then
        mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditClear_Click()
    '删除
    Dim lngBillId As Long
    Dim intRow As Integer
    Dim intReturn As Integer
    Dim intRecord As Integer
    Dim strNo As String
    Dim StrDate As String
    
    Dim Dbl数量 As Double
    Dim strsql As String
    Dim rscord As Recordset
    
    On Error GoTo errHandle
    strsql = "select 毁损数量 from 药品质量记录 where id=[1]"
    Set rscord = zldatabase.OpenSQLRecord(strsql, "mnuEditClear_Click", vsfList.TextMatrix(vsfList.Row, 0))
    If rscord.EOF Then
       MsgBox "该药品质量记录已被其他人删除，请检查！", vbOKOnly, gstrSysName
       Exit Sub
    Else
        Dbl数量 = rscord!毁损数量
    End If
    rscord.Close
    
    With vsfList
        intRow = .Row
        lngBillId = .TextMatrix(intRow, 0)
        intReturn = MsgBox("你确实要清除药品信息为“" & .TextMatrix(.Row, 1) & "”的药品质量管理单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
        intRecord = .rows - 1
        If intReturn = vbYes Then
            strNo = Get出库单据号(lngBillId)
            
            If Trim(strNo) <> "" Then gcnOracle.BeginTrans
            
            gstrSQL = "zl_药品质量管理_delete(" & lngBillId & ")"
            Call zldatabase.ExecuteProcedure(gstrSQL, MStrCaption)
            
            If Trim(strNo) <> "" Then
                StrDate = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                
                gstrSQL = "ZL_药品其他出库_STRIKE("
                '行次
                gstrSQL = gstrSQL & "1"
                '原记录状态
                gstrSQL = gstrSQL & ",1"
                'NO
                gstrSQL = gstrSQL & ",'" & strNo & "'"
                '序号
                gstrSQL = gstrSQL & ",1"
                '药品ID
                gstrSQL = gstrSQL & "," & Val(.TextMatrix(intRow, 2 + IIf(gint药品名称显示 = 2, 1, 0)))
                '冲销数量
                'gstrSQL = gstrSQL & "," & Val(.TextMatrix(intRow, 9 + IIf(gint药品名称显示 = 2, 1, 0))) * Val(.TextMatrix(intRow, 10 + IIf(gint药品名称显示 = 2, 1, 0)))
                'gstrSQL = gstrSQL & "," & Val(.TextMatrix(intRow, 13 + IIf(gint药品名称显示 = 2, 1, 0)))
                gstrSQL = gstrSQL & "," & Dbl数量
                
                '填制人
                gstrSQL = gstrSQL & ",'" & UserInfo.用户姓名 & "'"
                '填制日期
                gstrSQL = gstrSQL & ",to_date('" & Format(StrDate, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS')"
                gstrSQL = gstrSQL & ")"
                
                Call zldatabase.ExecuteProcedure(gstrSQL, MStrCaption)
                
                gcnOracle.CommitTrans
            End If
            
            intRecord = intRecord - 1
            If .rows > 2 Then
                .RemoveItem intRow
            ElseIf .rows = 2 Then
                .rows = 3
                .RemoveItem intRow
                SetEnable
            End If
            If intRow < .rows - 1 Then
                .Row = intRow
            Else
                If .rows = 2 Then
                    .Row = 1
                Else
                    .Row = intRow - 1
                End If
            End If
            .Col = 0
            .ColSel = .Cols - 1
        End If
    End With
    staThis.Panels(2).Text = "当前共有" & intRecord & "张单据"
    Exit Sub

errHandle:
    If Trim(strNo) <> "" Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
        
End Sub

Private Sub mnuEditVerify_Click()
    '验收
    
    Dim lngRecordID As Long
    Dim BlnSuccess As Boolean
    
    With vsfList
        lngRecordID = .TextMatrix(.Row, 0)
        BlnSuccess = frmDrugQualityCard.ShowCard(Me, 3, lngRecordID, mstrPrivs)
    
    End With
    If BlnSuccess = True Then
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
    Dim rsParrel As New Recordset
    
    With vsfList
        
        On Error GoTo errHandle
        intRow = .Row
        lngBillId = .TextMatrix(intRow, 0)
        intReturn = MsgBox("你确实要删除药品信息为“" & .TextMatrix(.Row, 1) & "”的药品质量管理单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
        intRecord = .rows - 1
        If intReturn = vbYes Then
            gstrSQL = "select nvl(处理人,'0') from 药品质量记录 where id=[1]"
            Set rsParrel = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption, lngBillId)
            
            If rsParrel.EOF Then
                MsgBox "该药品质量记录已被其他人删除，请检查！", vbOKOnly, gstrSysName
                Exit Sub
            ElseIf rsParrel.Fields(0) <> "0" Then
                MsgBox "该药品质量记录已被其他人处理，请检查！", vbOKOnly, gstrSysName
                Exit Sub
            End If
            rsParrel.Close
            
            gstrSQL = "zl_药品质量管理_delete(" & lngBillId & ")"
        
            If gstrSQL = "" Then Exit Sub
            Call zldatabase.ExecuteProcedure(gstrSQL, MStrCaption)
            
            intRecord = intRecord - 1
            If .rows > 2 Then
                .RemoveItem intRow
            ElseIf .rows = 2 Then
                .rows = 3
                .RemoveItem intRow
                SetEnable
            End If
            If intRow < .rows - 1 Then
                .Row = intRow
            Else
                If .rows = 2 Then
                    .Row = 1
                Else
                    .Row = intRow - 1
                End If
            End If
            .Col = 0
            .ColSel = .Cols - 1
        End If
    End With
    staThis.Panels(2).Text = "当前共有" & intRecord & "张单据"
    Exit Sub

errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Sub

Private Sub mnuEditDisplay_Click()
    '查看单据
    
    Dim lngRecordID As Long
    
    With vsfList
        lngRecordID = .TextMatrix(.Row, 0)
        frmDrugQualityCard.ShowCard Me, 4, lngRecordID, mstrPrivs
    End With
End Sub

Private Sub mnuEditModify_Click()
    '修改
    Dim lngRecordID As Long
    Dim BlnSuccess As Boolean
    
    BlnSuccess = False
    With vsfList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        lngRecordID = .TextMatrix(.Row, 0)
        BlnSuccess = frmDrugQualityCard.ShowCard(Me, 2, lngRecordID, mstrPrivs)
        If BlnSuccess = True Then
            mnuViewRefresh_Click
        End If
    End With
End Sub

Private Sub mnuFileBillPreview_Click()
    With vsfList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        ReportOpen gcnOracle, glngSys, "zl1_bill_1331", Me
    End With
End Sub

Private Sub mnuFileBillPrint_Click()
    Call mnuFileBillPreview_Click
End Sub

Private Sub mnuFileExcel_Click()
    '输出到Excel
    vsfList.Redraw = False
    subPrint 3
    vsfList.Redraw = True
    vsfList.Col = 0
    vsfList.ColSel = vsfList.Cols - 1
End Sub

Private Sub mnufileexit_Click()
    '退出
    Unload Me
    
End Sub

Private Sub mnuFileParameter_Click()
    '参数设置
    Dim dateCurrentDate As Date
    Dim int查询天数 As Date
    
    frm参数设置.设置参数 Me, mstrPrivs, MStrCaption
    
    dateCurrentDate = Sys.Currentdate
    int查询天数 = Val(zldatabase.GetPara("查询天数", glngSys, mlngMode, 7)) - 1
    strStart = Format(DateAdd("d", -int查询天数, dateCurrentDate), "yyyy-MM-dd")
    strEnd = Format(dateCurrentDate, "yyyy-MM-dd")

    SQLCondition.date填制时间开始 = CDate(Format(strStart, "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date填制时间结束 = CDate(Format(strEnd, "yyyy-mm-dd") & " 23:59:59")
    
    Call GetList(mstrFind)
End Sub

Private Sub mnuFilePreView_Click()
    '打印预览
    vsfList.Redraw = False
    subPrint 2
    vsfList.Redraw = True
    vsfList.Col = 0
    vsfList.ColSel = vsfList.Cols - 1
    
End Sub

Private Sub mnuFilePrint_Click()
    '打印
    vsfList.Redraw = False
    subPrint 1
    vsfList.Redraw = True
    vsfList.Col = 0
    vsfList.ColSel = vsfList.Cols - 1
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
'    ReportMan gcnOracle, Me
   Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub mnuHelpWebHome_Click()
    '中联主页
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    '发送反馈
    Call zlMailTo(Me.hWnd)
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    '默认参数：药品=药品id，库房=库房id，供应商=供应商id，开始时间=填制开始时间，结束时间=填制结束时间
    Dim str开始时间 As String
    Dim str结束时间 As String
    
    str开始时间 = IIf(Format(SQLCondition.date填制时间开始, "yyyy-mm-dd") = "1899-12-30", "", Format(SQLCondition.date填制时间开始, "yyyy-mm-dd"))
    str结束时间 = IIf(Format(SQLCondition.date填制时间结束, "yyyy-mm-dd") = "1899-12-30", "", Format(SQLCondition.date填制时间结束, "yyyy-mm-dd"))
    
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
        "药品=" & IIf(SQLCondition.lng药品 = 0, "", SQLCondition.lng药品), _
        "库房=" & IIf(Val(cboStock.ItemData(cboStock.ListIndex)) = 0, "", Val(cboStock.ItemData(cboStock.ListIndex))), _
        "供应商=" & IIf(SQLCondition.lng供应商 = 0, "", SQLCondition.lng供应商), _
        "开始时间=" & str开始时间, _
        "结束时间=" & str结束时间)
End Sub

Private Sub mnuViewRefresh_Click()
    '刷新
    GetList mstrFind
End Sub

Private Sub mnuViewSearch_Click()
    '查找
    Dim strFind As String
    
    strFind = FrmDrugQualitySearch.GetSearch(Me, strStart, strEnd, strVerifyStart, strVerifyEnd, _
                SQLCondition.date填制时间开始, _
                SQLCondition.date填制时间结束, _
                SQLCondition.date审核时间开始, _
                SQLCondition.date审核时间结束, _
                SQLCondition.lng药品, _
                SQLCondition.lng供应商, _
                SQLCondition.str填制人, _
                SQLCondition.str审核人, _
                cboStock.ItemData(cboStock.ListIndex))
    
    If strFind <> "" Then
        mstrFind = strFind
        GetList mstrFind
        If Format(strStart, "yyyy-mm-dd") = "1901-01-01" And Format(strVerifyStart, "yyyy-mm-dd") = "1901-01-01" Then
            lblRange.Visible = False
        ElseIf Format(strStart, "yyyy-mm-dd") <> "1901-01-01" And Format(strVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
            lblRange.Visible = True
            lblRange = "查询范围:登记时间 " & Format(strStart, "yyyy年MM月dd日") & "至" & Format(strEnd, "yyyy年MM月dd日") & "  处理时间 " & Format(strVerifyStart, "yyyy年MM月dd日") & "至" & Format(strVerifyEnd, "yyyy年MM月dd日")
        ElseIf Format(strStart, "yyyy-mm-dd") <> "1901-01-01" Then
            lblRange.Visible = True
            lblRange = "查询范围:登记时间 " & Format(strStart, "yyyy年MM月dd日") & "至" & Format(strEnd, "yyyy年MM月dd日")
        ElseIf Format(strVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
            lblRange.Visible = True
            lblRange = "查询范围:登记时间 " & Format(strVerifyStart, "yyyy年MM月dd日") & "至" & Format(strVerifyEnd, "yyyy年MM月dd日")
        End If
             
    End If
    
End Sub

Private Sub mnuViewStatus_Click()
    With mnuViewStatus
        .Checked = Not .Checked  ' Xor True
        staThis.Visible = .Checked
    End With
    
    Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    With mnuViewToolButton
        .Checked = Not .Checked   ' Xor True
        cbrTool.Bands(1).Visible = .Checked
        mnuViewToolText.Enabled = .Checked
    End With
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim intCount As Integer      '工具条索引
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked   ' Xor True
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


Private Sub vsfList_Click()
    With vsfList
         If .Row < 1 Then Exit Sub
         If .MouseRow = 0 Then
            ListSort
            Exit Sub
         End If
    End With
End Sub

Private Sub vsfList_DblClick()
    If mnuEditModify.Visible = False Then Exit Sub
    If mnuEditModify.Enabled = False Then Exit Sub
    If vsfList.MouseRow = 0 Then Exit Sub
    mnuEditDisplay_Click
End Sub

Private Sub vsfList_EnterCell()
    SetEnable
End Sub

Private Sub vsfList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If mnuEditModify.Visible = False Then Exit Sub
        If mnuEditModify.Enabled = False Then Exit Sub
        mnuEditDisplay_Click
    End If
        
End Sub

Private Sub vsfList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    If mnuEdit.Visible = False Then Exit Sub
    
    PopupMenu mnuEdit, 2
    
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
    With vsfList
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
            If mnuEditClear.Visible = True Then
                mnuEditClear.Enabled = False
                tlbTool.Buttons("Clear").Enabled = False
            End If
            
            If mnuEditVerify.Visible = True Then
                mnuEditVerify.Enabled = False
                tlbTool.Buttons("Verify").Enabled = False
            End If
            
            If mnuEditDisplay.Visible = True Then
                mnuEditDisplay.Enabled = False
            End If
        Else
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
'    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As Object
    Dim objRow As New zlTabAppRow
    Dim strRange As String
    
    If Format(strStart, "yyyy-mm-dd") = "1901-01-01" And Format(strVerifyStart, "yyyy-mm-dd") = "1901-01-01" Then
        strRange = "审核日期 " & Format(strVerifyStart, "yyyy年MM月dd日") & "至" & Format(strVerifyEnd, "yyyy年MM月dd日")
    ElseIf Format(strVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
        strRange = "填制日期 " & Format(strStart, "yyyy年MM月dd日") & "至" & Format(strEnd, "yyyy年MM月dd日") & "  审核日期 " & Format(strVerifyStart, "yyyy年MM月dd日") & "至" & Format(strVerifyEnd, "yyyy年MM月dd日")
    Else
        strRange = "填制日期 " & Format(strStart, "yyyy年MM月dd日") & "至" & Format(strEnd, "yyyy年MM月dd日")
    End If
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
    objPrint.Title.Text = MStrCaption
        
    objRow.Add "时间：" & strRange
    objRow.Add "部门：" & cboStock.Text
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
        
    objRow.Add "打印人:" & UserInfo.用户姓名
    objRow.Add "打印日期:" & Format(Sys.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    
    Set objPrint.Body = vsfList
    
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
'            mnuEditAddPayment_Click
        Case "Imprest"
'            mnuEditAddImprest_Click
    End Select
End Sub

Private Sub tlbTool_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool
    End If
End Sub

'对单据头列排序
Private Sub ListSort()
    Dim intCol As Integer
    Dim intRow As Integer
    Dim intTemp As String
    
    With vsfList
        If .rows > 1 Then
            .Redraw = False
            intCol = .MouseCol
            .Col = intCol
            .ColSel = intCol
            intTemp = .TextMatrix(.Row, 0)
            If intCol = mintPreCol And mintsort = flexSortStringNoCaseDescending Then
               .Sort = flexSortStringNoCaseAscending
               mintsort = flexSortStringNoCaseAscending
            Else
               .Sort = flexSortStringNoCaseDescending
               mintsort = flexSortStringNoCaseDescending
            End If
            
            mintPreCol = intCol
            .Row = FindRow(vsfList, intTemp, 0)
            If .RowPos(.Row) + .RowHeight(.Row) > .Height Then
                .TopRow = .Row
            Else
                .TopRow = 1
            End If
            .Col = 0
            .ColSel = .Cols - 1
            .Redraw = True
            .SetFocus
        Else
            .ColSel = 0
        End If
    End With
End Sub

'寻找与某一列相等的行
Public Function FindRow(ByVal FlexTemp As MSHFlexGrid, ByVal intTemp As Variant, ByVal intCol As Integer) As Integer
    Dim i As Integer
    
    With FlexTemp
        For i = 1 To .rows - 1
            If IsDate(intTemp) Then
               If Format(.TextMatrix(i, intCol), "yyyy-mm-dd") = Format(intTemp, "yyyy-mm-dd") Then
                  FindRow = i
                  Exit Function
               End If
            Else
                If .TextMatrix(i, intCol) = intTemp Then
                  FindRow = i
                  Exit Function
                End If
            End If
        Next
    End With
    FindRow = 1
End Function


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

