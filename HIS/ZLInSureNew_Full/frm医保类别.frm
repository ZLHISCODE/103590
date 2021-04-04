VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm医保类别 
   Caption         =   "保险类别设置"
   ClientHeight    =   6105
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9765
   Icon            =   "frm医保类别.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   9765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cmb中心 
      Height          =   300
      Left            =   6960
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   900
      Width           =   1815
   End
   Begin VB.PictureBox picSplitV 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5490
      Left            =   5370
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5490
      ScaleWidth      =   45
      TabIndex        =   1
      Top             =   690
      Width           =   45
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   3660
      Top             =   5340
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保类别.frx":0E42
            Key             =   "Fix"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保类别.frx":115C
            Key             =   "FixD"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保类别.frx":1476
            Key             =   "Common"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保类别.frx":1790
            Key             =   "CommonD"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   2850
      Top             =   5310
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保类别.frx":1AAA
            Key             =   "Fix"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保类别.frx":1DC4
            Key             =   "FixD"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保类别.frx":20DE
            Key             =   "Common"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保类别.frx":23F8
            Key             =   "CommonD"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   7770
      Top             =   480
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
            Picture         =   "frm医保类别.frx":2712
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保类别.frx":292C
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保类别.frx":2B46
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保类别.frx":2D60
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保类别.frx":2F7A
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保类别.frx":3194
            Key             =   "Select"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保类别.frx":388E
            Key             =   "Parameter"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保类别.frx":3F88
            Key             =   "View"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保类别.frx":41A2
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保类别.frx":43BC
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   7110
      Top             =   510
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
            Picture         =   "frm医保类别.frx":45D6
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保类别.frx":47F0
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保类别.frx":4A0A
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保类别.frx":4C24
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保类别.frx":4E3E
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保类别.frx":5058
            Key             =   "Select"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保类别.frx":5752
            Key             =   "Parameter"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保类别.frx":5E4C
            Key             =   "View"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保类别.frx":6066
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保类别.frx":6280
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSplitH 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   5550
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   3000
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2880
      Width           =   3000
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9765
      _ExtentX        =   17224
      _ExtentY        =   1244
      BandCount       =   1
      _CBWidth        =   9765
      _CBHeight       =   705
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinHeight1      =   645
      Width1          =   615
      FixedBackground1=   0   'False
      Key1            =   "Comm"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   645
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   9645
         _ExtentX        =   17013
         _ExtentY        =   1138
         ButtonWidth     =   820
         ButtonHeight    =   1138
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   15
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "Preview"
               Object.ToolTipText     =   "打印预览"
               Object.Tag             =   "预览"
               ImageKey        =   "Preview"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "Print"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "增加"
               Key             =   "New"
               Object.ToolTipText     =   "增加保险类别"
               Object.Tag             =   "增加"
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修改"
               Key             =   "Modify"
               Object.ToolTipText     =   "修改保险类别"
               Object.Tag             =   "修改"
               ImageKey        =   "Modify"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "Delete"
               Object.ToolTipText     =   "删除保险类别"
               Object.Tag             =   "删除"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split1"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "选择"
               Key             =   "Select"
               Object.ToolTipText     =   "设为当前使用医保"
               Object.Tag             =   "选择"
               ImageKey        =   "Select"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split4"
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "参数"
               Key             =   "Parameter"
               Description     =   "参数"
               Object.ToolTipText     =   "参数设置"
               Object.Tag             =   "参数"
               ImageKey        =   "Parameter"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split2"
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查看"
               Key             =   "View"
               Description     =   "View"
               Object.ToolTipText     =   "查看方式"
               Object.Tag             =   "查看"
               ImageKey        =   "View"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "大图标"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "小图标"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "列表"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "详细资料"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split3"
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Object.ToolTipText     =   "帮助主题"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Quit"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ListView lvwKind_S 
      Height          =   4755
      Left            =   30
      TabIndex        =   0
      Top             =   810
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   8387
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils32"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   5745
      Width           =   9765
      _ExtentX        =   17224
      _ExtentY        =   635
      SimpleText      =   $"frm医保类别.frx":649A
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm医保类别.frx":64E1
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12144
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh分段 
      Height          =   1800
      Left            =   5220
      TabIndex        =   7
      Top             =   3960
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   3175
      _Version        =   393216
      Rows            =   5
      Cols            =   4
      FixedCols       =   0
      BackColorFixed  =   -2147483638
      BackColorBkg    =   -2147483643
      GridColor       =   4210752
      GridColorFixed  =   4210752
      GridLinesFixed  =   1
      ScrollBars      =   2
      MergeCells      =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh参数 
      Height          =   1560
      Left            =   5160
      TabIndex        =   6
      Top             =   2040
      Width           =   3270
      _ExtentX        =   5768
      _ExtentY        =   2752
      _Version        =   393216
      BackColor       =   16777215
      Rows            =   6
      FixedCols       =   0
      BackColorFixed  =   13684944
      BackColorBkg    =   -2147483643
      GridColor       =   4210752
      GridColorFixed  =   4210752
      GridLinesFixed  =   1
      ScrollBars      =   2
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lbl中心 
      AutoSize        =   -1  'True
      Caption         =   "医保中心(&N)"
      Height          =   180
      Left            =   5580
      TabIndex        =   9
      Top             =   960
      Width           =   990
   End
   Begin VB.Label lbl参数 
      Alignment       =   2  'Center
      BackColor       =   &H00E6F5FD&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "相关运行参数"
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   5160
      TabIndex        =   5
      Top             =   1755
      Width           =   3360
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuFilePreview 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnuFileSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "类别(&E)"
      Begin VB.Menu mnuEditAdd 
         Caption         =   "增加(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "修改(&M)"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "删除(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditSplit0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelect 
         Caption         =   "设为当前使用医保(&S)"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuEditDeselect 
         Caption         =   "取消选择(&E)"
      End
   End
   Begin VB.Menu mnuCenter 
      Caption         =   "中心(&C)"
      Begin VB.Menu mnuCenterAdd 
         Caption         =   "增加(&A)"
      End
      Begin VB.Menu mnuCenterModify 
         Caption         =   "修改(&M)"
      End
      Begin VB.Menu mnuCenterDelete 
         Caption         =   "删除(&D)"
      End
      Begin VB.Menu mnuCenterSplitPara 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCenterParameter 
         Caption         =   "运行参数设置(&P)"
      End
      Begin VB.Menu mnuCenterSplitYear 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCenterYear 
         Caption         =   "年龄段(&I)"
         Index           =   0
      End
      Begin VB.Menu mnuCenterSplitSect 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCenterSect 
         Caption         =   "支付费用档(&E)"
      End
      Begin VB.Menu mnuCenterSplitSpec 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCenterHome 
         Caption         =   "申请家庭病床(&H)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCenterSpec 
         Caption         =   "申请门诊特殊病(&S)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCenterEspecial 
         Caption         =   "申请特治特检(&T)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCenterOut 
         Caption         =   "申请转院(&O)"
         Visible         =   0   'False
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
         Begin VB.Menu mnuViewToolspilt1 
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
      Begin VB.Menu mnuViewSplit0 
         Caption         =   "-"
      End
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
         Index           =   3
      End
      Begin VB.Menu mnuViewSplit1 
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
         Caption         =   "Web上的中联"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "中联主页(&H)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "发送反馈(&K)..."
         End
      End
      Begin VB.Menu mnuHelpSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
End
Attribute VB_Name = "frm医保类别"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const mstrLvw As String = "名称,2000,0,1;序号,800,0,2;医院编码,1440,0,0;说明,2000,0,0"

Dim msngStartX As Single, msngStartY As Single    '移动前鼠标的位置
Dim mblnLoad As Boolean  '窗口还未打开时为真
Dim mintColumn As Integer '
Dim mstrKey As String       '当前选择的ListItem的Key值
Dim mbln年龄段 As Boolean   '是否具有编辑年龄段的权限

Private Sub Form_Activate()
    If mblnLoad = True Then
        Call Form_Resize '为了使CoolBar自适应高度
        FillList
    End If
    mblnLoad = False
End Sub

Private Sub Form_Load()
    mblnLoad = True
    
    Call 权限控制
    '允许进行列删除的ListView须做标记
    lvwKind_S.Tag = "可变化的"
    '-----------
    RestoreWinState Me, App.ProductName
    '如果ListView的还未被设置，比如第一次使用，那就调用缺省的初始化
    If lvwKind_S.ColumnHeaders.Count = 0 Then
        zlControl.LvwSelectColumns lvwKind_S, mstrLvw, True
    End If
    '根据lvwKind_S显示设置对应菜单
    mnuViewIcon_Click lvwKind_S.View
    
    lvwKind_S.Sorted = True
    lvwKind_S.SortKey = 1
    
    zlControl.CboSetHeight cmb中心, 3600
    Call InitTable
End Sub

Private Sub InitTable()
'功能：初始化表格
    With msh参数
        .Rows = 2: .Cols = 2
        .TextMatrix(0, 0) = "参数名"
        .TextMatrix(0, 1) = "参数值"
        .ColWidth(0) = 1900
        .ColWidth(1) = 3200
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        
        .COL = 0
        .Row = 0
        .ColSel = .Cols - 1
        .RowSel = 0
        .FillStyle = flexFillRepeat
        .CellAlignment = 4
        .FillStyle = flexFillSingle
        .AllowBigSelection = False
        .Row = 1
    End With
    
    With msh分段
        .Cols = 4
        .ColWidth(0) = 1500
        .ColWidth(1) = 1000
        .ColWidth(2) = 1000
        .ColWidth(3) = 2000
        .ColAlignment(0) = 1
        .ColAlignment(1) = 7
        .ColAlignment(2) = 7
        .ColAlignment(3) = 1
    End With
End Sub

Private Sub Form_Resize()
    Dim sngTop As Single, sngBottom As Single
    On Error Resume Next
    
    sngTop = IIf(cbrThis.Visible, cbrThis.Top + cbrThis.Height, 0)
    sngBottom = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0)
    
    lvwKind_S.Top = sngTop
    lvwKind_S.Height = IIf(sngBottom - lvwKind_S.Top > 0, sngBottom - lvwKind_S.Top, 0)
    lvwKind_S.Left = ScaleLeft
    
    picSplitV.Top = sngTop
    picSplitV.Height = IIf(sngBottom - picSplitV.Top > 0, sngBottom - picSplitV.Top, 0)
    picSplitV.Left = lvwKind_S.Left + lvwKind_S.Width
    
    With cmb中心
        '设置控件的左边距与宽度
        lbl中心.Left = picSplitV.Left + picSplitV.Width
        .Left = lbl中心.Left + lbl中心.Width + 30
        .Width = IIf(ScaleWidth - cmb中心.Left > 0, ScaleWidth - cmb中心.Left, 0)
    
        lbl参数.Left = lbl中心.Left
        lbl参数.Width = IIf(ScaleWidth - lbl参数.Left > 0, ScaleWidth - lbl参数.Left, 0)
    End With
    With lbl参数
        msh参数.Left = .Left
        msh参数.Width = .Width
        picSplitH.Left = .Left
        picSplitH.Width = .Width
        msh分段.Left = .Left
        msh分段.Width = .Width
    End With
    
    If cmb中心.Visible = True Then
        cmb中心.Top = sngTop
        lbl中心.Top = sngTop + 60
        lbl参数.Top = cmb中心.Top + cmb中心.Height + 120
    Else
        lbl参数.Top = sngTop + 90
    End If
    
    msh参数.Top = lbl参数.Top + lbl参数.Height
    picSplitH.Top = msh参数.Top + msh参数.Height
    msh分段.Top = picSplitH.Top + picSplitH.Height
    msh分段.Height = IIf(sngBottom - msh分段.Top > 0, sngBottom - msh分段.Top, 0)
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub lvwKind_S_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintColumn = ColumnHeader.Index - 1 Then '仍是刚才那列
        lvwKind_S.SortOrder = IIf(lvwKind_S.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvwKind_S.SortKey = mintColumn
        lvwKind_S.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwKind_S_DblClick()
    If mnuEditModify.Visible = True And mnuEditModify.Enabled = True Then
        Call mnuEditModify_Click
    End If
End Sub

Private Sub lvwKind_S_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If mstrKey = Item.Key Then Exit Sub
    mstrKey = Item.Key
    
    Dim rsTemp As New ADODB.Recordset
    Dim lngCount As Long, lngIndex As Long
    
    On Error GoTo errHandle
    If mbln年龄段 = True Then
        '首先隐藏所有年龄段
        mnuCenterYear(0).Visible = False
        mnuCenterSplitYear.Visible = False
        For lngCount = 1 To mnuCenterYear.UBound
            Unload mnuCenterYear(lngCount)
        Next
        
        '然后再按新的人群分类打开
        gstrSQL = "select * from 保险人群 where 险类=[1] order by 序号"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CInt(Mid(Item.Key, 2)))
        
        Do Until rsTemp.EOF
            lngIndex = rsTemp("序号") - 1
            If lngIndex = 0 Then
                mnuCenterYear(0).Visible = True
                mnuCenterSplitYear.Visible = True
            Else
                Load mnuCenterYear(lngIndex)
            End If
            
            mnuCenterYear(lngIndex).Caption = rsTemp("名称") & "(&" & rsTemp("序号") & ")"
            rsTemp.MoveNext
        Loop
        rsTemp.Close
    End If
        
    cmb中心.Clear
    cmb中心.Visible = (Item.Tag = "1")
    lbl中心.Visible = cmb中心.Visible
    Call Form_Resize
    
    If cmb中心.Visible = False Then
        '该医保只能有一个中心
        cmb中心.AddItem "1." & Item.Text
        cmb中心.ListIndex = 0
        Exit Sub
    End If
    
    On Error GoTo errHandle
    
    gstrSQL = "select 序号,编码,名称 from 保险中心目录 where 险类=[1] order by 序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CInt(Mid(Item.Key, 2)))
    Do Until rsTemp.EOF
        cmb中心.AddItem rsTemp("编码") & "." & rsTemp("名称")
        cmb中心.ItemData(cmb中心.NewIndex) = rsTemp("序号")
        rsTemp.MoveNext
    Loop
    
    If cmb中心.ListCount > 0 Then
        cmb中心.ListIndex = 0
    Else
        Call FillItem
    End If
    
    '如果是沈阳医保，允许进行特殊业务的申请
    If Mid(Item.Key, 2) = TYPE_沈阳市 Then
        mnuCenterSplitSpec.Visible = True
        mnuCenterSpec.Visible = True
        mnuCenterEspecial.Visible = True
        mnuCenterHome.Visible = True
        mnuCenterOut.Visible = True
    Else
        mnuCenterSplitSpec.Visible = False
        mnuCenterSpec.Visible = False
        mnuCenterEspecial.Visible = False
        mnuCenterHome.Visible = False
        mnuCenterOut.Visible = False
    End If
    Exit Sub

errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmb中心_Click()
    '由于刷新参数与分档的时间并不长，且为了保证一更新ListIndex就能刷新
    '所以没有再保存上一次的ListIndex值
    Call FillItem
End Sub


Private Sub mnuCenterEspecial_Click()
    '
End Sub

Private Sub mnuCenterHome_Click()
    '
End Sub

Private Sub mnuCenterOut_Click()
    frm特殊业务申请.Show 1, Me
End Sub

Private Sub mnuCenterSpec_Click()
    '
End Sub

Private Sub mnuEditAdd_Click()
    If frm医保类别编辑.编辑医保类别("") = True Then
        lvwKind_S_ItemClick lvwKind_S.SelectedItem
    End If
End Sub

Private Sub mnuEditModify_Click()
    If lvwKind_S.SelectedItem Is Nothing Then Exit Sub
    
    If frm医保类别编辑.编辑医保类别(Mid(lvwKind_S.SelectedItem.Key, 2)) = True Then
        lvwKind_S_ItemClick lvwKind_S.SelectedItem
    End If
End Sub

Private Sub mnuEditDelete_Click()
    Dim intIndex As Integer
    
    If lvwKind_S.SelectedItem Is Nothing Then Exit Sub
    If MsgBox("你确认要删除“" & lvwKind_S.SelectedItem.Text & "”医保类别吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
        On Error GoTo errHandle
        
        gstrSQL = "zl_保险类别_delete(" & Mid(lvwKind_S.SelectedItem.Key, 2) & ")"
        
        MousePointer = vbHourglass
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        With lvwKind_S
            mstrKey = ""
            intIndex = .SelectedItem.Index
            .ListItems.Remove .SelectedItem.Key
            If .ListItems.Count > 0 Then
                intIndex = IIf(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
                .ListItems(intIndex).Selected = True
                .ListItems(intIndex).EnsureVisible
                lvwKind_S_ItemClick .SelectedItem
            Else
                cmb中心.Clear
                cmb中心.Visible = False
                Call Form_Resize
                Call FillItem
            End If
        End With
        MousePointer = vbDefault
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    MousePointer = vbDefault
End Sub

Private Sub mnuEditDeselect_Click()
    Dim lst As ListItem
    Dim strIcon As String
    
    SaveSetting "ZLSOFT", "公共全局", "是否支持医保", "No"
    gintInsure = 0
    SaveSetting "ZLSOFT", "公共全局", "医保类别", 0
    For Each lst In lvwKind_S.ListItems
        strIcon = IIf(Left(lst.Icon, 3) = "Fix", "Fix", "Common")
        
        lst.Icon = strIcon
        lst.SmallIcon = strIcon
    Next
    Call SetMenu
End Sub

Private Sub mnuEditSelect_Click()
    Dim lst As ListItem
    Dim strIcon As String
    
    SaveSetting "ZLSOFT", "公共全局", "是否支持医保", "Yes"
    For Each lst In lvwKind_S.ListItems
        If lst Is lvwKind_S.SelectedItem Then
            '设为当前医保
            gintInsure = Mid(lst.Key, 2)
            SaveSetting "ZLSOFT", "公共全局", "医保类别", gintInsure
            strIcon = IIf(Left(lst.Icon, 3) = "Fix", "FixD", "CommonD")
        Else
            strIcon = IIf(Left(lst.Icon, 3) = "Fix", "Fix", "Common")
        End If
        
        lst.Icon = strIcon
        lst.SmallIcon = strIcon
    Next
    Call SetMenu
End Sub

Private Sub mnuCenterAdd_Click()
    If lvwKind_S.SelectedItem Is Nothing Then Exit Sub
    
    Call frm医保类别中心.编辑保险中心(Mid(lvwKind_S.SelectedItem.Key, 2), "")
End Sub

Private Sub mnuCenterModify_Click()
    If lvwKind_S.SelectedItem Is Nothing Then Exit Sub
    
    Call frm医保类别中心.编辑保险中心(Mid(lvwKind_S.SelectedItem.Key, 2), cmb中心.ItemData(cmb中心.ListIndex))
End Sub

Private Sub mnuCenterDelete_Click()
    If cmb中心.ListIndex < 0 Then Exit Sub
    If MsgBox("你确认要删除“" & cmb中心.Text & "”医保中心吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
        On Error GoTo errHandle
        
        gstrSQL = "zl_保险中心目录_delete(" & Mid(lvwKind_S.SelectedItem.Key, 2) & "," & cmb中心.ItemData(cmb中心.ListIndex) & ")"
        
        MousePointer = vbHourglass
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        With cmb中心
            .RemoveItem .ListIndex
            If .ListCount > 0 Then
                .ListIndex = 0
            Else
                Call FillItem
            End If
        End With
        MousePointer = vbDefault
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    MousePointer = vbDefault
End Sub

Private Sub mnuCenterParameter_Click()
'功能：修改医保类别的运行参数
'注意：不同医保其参数设置是由不同的程序实现的
    Dim blnReturn As Boolean
    Dim lng险类 As Long
    
    If lvwKind_S.SelectedItem Is Nothing Then Exit Sub
    
    On Error GoTo errHandle
    lng险类 = Val(Mid(lvwKind_S.SelectedItem.Key, 2))
    Select Case lng险类
        Case TYPE_南京市
            blnReturn = frmSet南京市.参数设置(lng险类)
        Case TYPE_昭通
            blnReturn = frmSet昭通.参数设置
        Case TYPE_徐州市
            blnReturn = frmSet徐州市.参数设置()
        Case TYPE_徐州农保
            blnReturn = frmset徐州农保.参数设置()
        Case TYPE_徐州
            blnReturn = frmset徐州.参数设置()
        Case TYPE_成都市农医
            blnReturn = frmSet成都市农医.参数设置
        Case TYPE_余姚
            blnReturn = 医保设置_余姚()
        Case TYPE_北京尚洋
            blnReturn = 医保设置_北京尚洋()
        Case TYPE_浙江
            blnReturn = frmSet浙江.参数设置()
        Case TYPE_新都
            blnReturn = 医保设置_新都()
        Case TYPE_重庆市
            '设置功能完全由窗体完成
            blnReturn = frmSet重庆.参数设置()
        Case TYPE_华东
            blnReturn = 医保设置_华东()
        Case TYPE_广元
            blnReturn = 医保设置_广元()
        Case TYPE_涪陵
            blnReturn = 医保设置_涪陵()
        Case TYPE_重庆壁山
            blnReturn = frmset壁山.参数设置()
        Case TYPE_重庆松藻
            blnReturn = frmSet松藻.参数设置(Mid(lvwKind_S.SelectedItem.Key, 2), cmb中心.ItemData(cmb中心.ListIndex))
        Case TYPE_重庆中梁山
            blnReturn = frmSet中梁山.参数设置()
        Case TYPE_云南省, TYPE_昆明市, TYPE_云南建水
            Dim msgReturn As VbMsgBoxResult
            
            msgReturn = MsgBox("请问本医保是否支持患慢性病、特种病的医保病人？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName)
            
            gstrSQL = "zl_保险参数_Delete(" & lng险类 & ",0)"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            
            '新增参数数据
            gstrSQL = "zl_保险参数_Insert(" & lng险类 & ",0,'支持慢性病、特种病','" & IIf(msgReturn = vbYes, "1", "0") & "',1)"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            
            blnReturn = True
        Case TYPE_贵阳市
            blnReturn = frmSet贵阳.参数设置(Mid(lvwKind_S.SelectedItem.Key, 2))
        Case TYPE_自贡市
            '设置功能完全由窗体完成
            blnReturn = frmSet中软.参数设置(Mid(lvwKind_S.SelectedItem.Key, 2), cmb中心.ItemData(cmb中心.ListIndex))
        Case TYPE_四川自贡
            blnReturn = frmSet自贡.参数设置(Mid(lvwKind_S.SelectedItem.Key, 2), cmb中心.ItemData(cmb中心.ListIndex))
        Case TYPE_泸州市
            '设置功能完全由窗体完成
            blnReturn = frmSet泸州.参数设置(Mid(lvwKind_S.SelectedItem.Key, 2), cmb中心.ItemData(cmb中心.ListIndex))
        Case TYPE_铜仁
            '设置功能完全由窗体完成
            blnReturn = frmSet铜仁.参数设置(Mid(lvwKind_S.SelectedItem.Key, 2), cmb中心.ItemData(cmb中心.ListIndex))
        Case TYPE_咸阳市
            '设置功能完全由窗体完成
            blnReturn = frmSet咸阳.参数设置()
        Case TYPE_成都市
            blnReturn = 医保设置_成都
        Case TYPE_成都莲合
            blnReturn = 医保设置_莲合
        Case TYPE_开县
            blnReturn = 医保设置_开县
        Case type_成都郊县
            blnReturn = 医保设置_成都郊县
        Case TYPE_成都南充
            blnReturn = 医保设置_成都南充
        Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
            blnReturn = 医保设置_福建巨龙(lng险类)
        Case type_米易
            blnReturn = 医保设置_米易
        Case TYPE_四川眉山
            blnReturn = 医保设置_眉山
        Case TYPE_沈阳市
            blnReturn = 医保设置_沈阳
        Case TYPE_乐山
            blnReturn = 医保设置_乐山
        Case TYPE_大连市, TYPE_大连开发区
            '200311
            If cmb中心.ListIndex < 0 Then Exit Sub
            blnReturn = 医保设置_大连(Val(Mid(lvwKind_S.SelectedItem.Key, 2)), cmb中心.ItemData(cmb中心.ListIndex))
        Case TYPE_重大校园卡
            '刘兴宏(200403)
            blnReturn = 医保设置_重大校园卡(Val(Mid(lvwKind_S.SelectedItem.Key, 2)), 0)
        Case TYPE_重庆银海版
            blnReturn = 医保设置_重庆银海版()
        Case TYPE_重庆渝北
            '20040715
            blnReturn = 医保设置_重庆渝北()
        Case TYPE_黔南
            '200410
            blnReturn = 医保设置_黔南()
        Case TYPE_成都德阳
            '200411
            blnReturn = 医保设置_成都德阳()
        Case TYPE_成都内江
            '200411
            blnReturn = 医保设置_成都内江()
        Case TYPE_兴安
            '20050125
            blnReturn = 医保设置_兴安()
        Case TYPE_吉林
            blnReturn = 医保设置_吉林()
        Case TYPE_临沧奉庆
            blnReturn = 医保设置_奉庆()
        Case TYPE_北京
            blnReturn = 医保设置_北京
        Case TYPE_毕节
            blnReturn = 医保设置_毕节
        Case TYPE_宁海
            blnReturn = 医保设置_宁海
        Case TYPE_慈溪农医
            blnReturn = 医保设置_慈溪农医
        Case TYPE_广元旺苍
            blnReturn = 医保设置_广元旺苍(Mid(lvwKind_S.SelectedItem.Key, 2), cmb中心.ItemData(cmb中心.ListIndex))
        Case TYPE_南充阆中
            blnReturn = 医保设置_南充阆中(Mid(lvwKind_S.SelectedItem.Key, 2), cmb中心.ItemData(cmb中心.ListIndex))
        Case TYPE_渝北农医
            blnReturn = 医保设置_渝北农医
        Case TYPE_兴成核工业
            blnReturn = 医保设置_兴成(Mid(lvwKind_S.SelectedItem.Key, 2), cmb中心.ItemData(cmb中心.ListIndex))
        Case TYPE_陕西大兴
            blnReturn = 医保设置_神木大兴(Mid(lvwKind_S.SelectedItem.Key, 2), cmb中心.ItemData(cmb中心.ListIndex))
        Case TYPE_山西
            '陈东：20050304
            blnReturn = frmSet山西.参数设置()
        Case TYPE_铜山县
            blnReturn = frmSet铜山县.参数设置()
        Case Is > 900, TYPE_徐州六院            '中联医保
            blnReturn = frmSet中联.参数设置(Mid(lvwKind_S.SelectedItem.Key, 2), cmb中心.ItemData(cmb中心.ListIndex))
    End Select
    
    If blnReturn = True Then
        '设置成功，刷新显示
        Call Fill参数
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub mnuCenterSect_Click()
    Dim blnReturn As Boolean
    
    If lvwKind_S.SelectedItem Is Nothing Then Exit Sub
    
    blnReturn = frm医保类别档次.档次设置(Mid(lvwKind_S.SelectedItem.Key, 2), cmb中心.ItemData(cmb中心.ListIndex))
    If blnReturn = True Then
        '设置成功，刷新显示
        Call Fill分段
    End If
End Sub

Private Sub mnuCenterYear_Click(Index As Integer)
    Dim blnReturn As Boolean
    Dim STRNAME As String
    
    If lvwKind_S.SelectedItem Is Nothing Then Exit Sub
    
    STRNAME = Left(mnuCenterYear(Index).Caption, InStr(mnuCenterYear(Index).Caption, "(") - 1)
    blnReturn = frm医保年龄段.档次设置(Mid(lvwKind_S.SelectedItem.Key, 2), cmb中心.ItemData(cmb中心.ListIndex), Index + 1, STRNAME)
    If blnReturn = True Then
        '设置成功，刷新显示
        Call Fill分段
    End If
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFilePreview_Click()
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub mnuFilePrintSet_Click()
    zlPrintSet
End Sub

Private Sub subPrint(ByVal bytMode As Byte)
    '功能:进行打印,预览和输出到EXCEL
    '参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    
    Dim objPrint As New zlPrintGrds
    Dim objRow As New zlTabAppRow
    
    Set objPrint.Grds = New Collection
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
    objPrint.Title.Text = "保险类别设置"
        
    objRow.Add "保险类别：" & lvwKind_S.SelectedItem.Text
    If cmb中心.Visible = True Then
        objRow.Add "医保中心：" & cmb中心.Text
    End If
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add " "
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人:" & UserInfo.姓名    '& "   打印日期:" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印日期:" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    
    objPrint.Grds.Add msh参数
    objPrint.Grds.Add msh分段
    
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrViewGrds objPrint, 1
          Case 2
              zlPrintOrViewGrds objPrint, 2
          Case 3
              zlPrintOrViewGrds objPrint, 3
      End Select
    Else
        zlPrintOrViewGrds objPrint, bytMode
    End If
End Sub


Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(hwnd)
End Sub


Private Sub mnuViewRefresh_Click()
    Call FillList
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    cbrThis.Visible = mnuViewToolButton.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim intCOUNT As Integer
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For intCOUNT = 1 To tbrThis.Buttons.Count
        tbrThis.Buttons(intCOUNT).Caption = IIf(mnuViewToolText.Checked, tbrThis.Buttons(intCOUNT).Tag, "")
    Next
    
    cbrThis.Bands(1).MinHeight = tbrThis.Height
    Call Form_Resize
End Sub

Private Sub mnuViewIcon_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
    Next
    mnuViewIcon(Index).Checked = True
    lvwKind_S.View = Index
End Sub

Private Sub msh参数_DblClick()
    If mnuCenterParameter.Visible = True And mnuCenterParameter.Enabled = True Then
        Call mnuCenterParameter_Click
    End If
End Sub

Private Sub msh分段_DblClick()
    If msh分段.RowData(msh分段.Row) <> 0 Then
        '调用年龄段程序
        If mbln年龄段 = True And mnuCenterYear(0).Enabled = True Then
            Call mnuCenterYear_Click(msh分段.RowData(msh分段.Row) - 1)
        End If
    Else
        '调用费用档次程序
        If mnuCenterSect.Visible And mnuCenterSect.Enabled Then
            Call mnuCenterSect_Click
        End If
    End If
End Sub

Private Sub picSplitV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        msngStartX = x
    End If
End Sub

Private Sub picSplitV_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sngTemp As Single
    
    If Button = 1 Then
        sngTemp = picSplitV.Left + x - msngStartX
        If sngTemp > 2000 And Me.ScaleWidth - (sngTemp + picSplitV.Width) > 1000 Then
            picSplitV.Left = sngTemp
            lvwKind_S.Width = picSplitV.Left - lvwKind_S.Left
            
            Call Form_Resize
        End If
        lvwKind_S.SetFocus
    End If
End Sub

Private Sub picSplitH_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        msngStartY = y
    End If
End Sub

Private Sub picSplitH_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sngTemp As Single
    
    If Button = 1 Then
        sngTemp = picSplitH.Top + y - msngStartY
        If sngTemp - msh参数.Top > 500 And (msh分段.Top + msh分段.Height) - (sngTemp + picSplitV.Width) > 1000 Then
            picSplitH.Top = sngTemp
            msh参数.Height = picSplitH.Top - msh参数.Top
            
            Call Form_Resize
        End If
        msh参数.SetFocus
    End If
End Sub

Private Sub tbrThis_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim i As Integer
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
    Next
    mnuViewIcon(ButtonMenu.Index - 1).Checked = True
    lvwKind_S.View = ButtonMenu.Index - 1
End Sub

Private Sub tbrThis_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuViewTool, 2
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Preview"
            mnuFilePreview_Click
        Case "Print"
            mnuFilePrint_Click
        Case "New"
            mnuEditAdd_Click
        Case "Modify"
            mnuEditModify_Click
        Case "Delete"
            mnuEditDelete_Click
        Case "Select"
            mnuEditSelect_Click
        Case "Parameter"
            mnuCenterParameter_Click
        Case "Help"
            mnuHelpTitle_Click
        Case "Quit"
            mnuFileExit_Click
        Case "View"
            mnuViewIcon(lvwKind_S.View).Checked = False
            If lvwKind_S.View = 3 Then
                mnuViewIcon(0).Checked = True
                lvwKind_S.View = 0
            Else
                mnuViewIcon(lvwKind_S.View + 1).Checked = True
                lvwKind_S.View = lvwKind_S.View + 1
            End If
    End Select
End Sub

Private Sub FillList()
'功能：显示所有医保类别列表
    Dim rsTemp As New ADODB.Recordset
    Dim strIcon As String
    Dim lst As ListItem, strKey As String
    Dim lngCol  As Long, varValue As Variant

    If Not lvwKind_S.SelectedItem Is Nothing Then
        strKey = lvwKind_S.SelectedItem.Key
    End If
    
    lvwKind_S.ListItems.Clear
    mstrKey = ""
    
    gstrSQL = "select 序号,名称,说明,医院编码,是否固定,具有中心,是否禁止 from 保险类别 Where 医保部件 Is NULL"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    Do Until rsTemp.EOF
        strIcon = IIf(rsTemp("是否固定") = 1, "Fix", "Common")
        If rsTemp("序号") = gintInsure Then strIcon = strIcon & "D"
        
        Set lst = lvwKind_S.ListItems.Add(, "K" & rsTemp("序号"), rsTemp("名称"), strIcon, strIcon)
        If rsTemp("序号") = gintInsure Then
            lst.Selected = True
        End If
        '根据ListView的列名从数据库取数
        For lngCol = 2 To lvwKind_S.ColumnHeaders.Count
            varValue = rsTemp(lvwKind_S.ColumnHeaders(lngCol).Text).Value
            lst.SubItems(lngCol - 1) = IIf(IsNull(varValue), "", varValue)
        Next
        lst.Tag = IIf(rsTemp("具有中心") = 1, 1, 0)
        If rsTemp("是否禁止") = 1 Then
            lst.Ghosted = True
        End If
        rsTemp.MoveNext
    Loop
    
    If lvwKind_S.ListItems.Count > 0 Then
        On Error Resume Next
        Set lst = lvwKind_S.ListItems(strKey)
        If Err <> 0 Then
            Err.Clear
            If lvwKind_S.SelectedItem Is Nothing Then
                Set lst = lvwKind_S.ListItems(1)
                lst.Selected = True
            Else
                Set lst = lvwKind_S.SelectedItem
            End If
        Else
            lst.Selected = True
        End If
        lst.EnsureVisible
        lvwKind_S_ItemClick lst
    Else
        cmb中心.Clear
        cmb中心.Visible = False
        Call Form_Resize
        Call FillItem
    End If
End Sub

Private Sub FillItem()
'功能：根据医保中心的序号显示参数及分段设置
    
    Call Fill参数
    Call Fill分段
    Call SetMenu
End Sub

Private Sub Fill参数()
'功能：根据医保中心的序号显示参数
    Dim rsTemp As New ADODB.Recordset
    Dim lst As ListItem, lngRow As Long
    Dim strTemp As String
    
    With msh参数
        .Rows = 2
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 1) = ""
    End With
    '如果没有中心，肯定不会有显示的
    If lvwKind_S.SelectedItem Is Nothing Or cmb中心.ListIndex < 0 Then Exit Sub
    
    With msh参数
        Select Case Val(Mid(lvwKind_S.SelectedItem.Key, 2))
            Case TYPE_成都市
                .Rows = 3
                .TextMatrix(1, 0) = "连接串"
                .TextMatrix(1, 1) = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("ConnectionStrINg"), "dsn=cnnSyb;uID=face;pwd=facepass")
                .TextMatrix(2, 0) = "卡号长度"
                .TextMatrix(2, 1) = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("CardNOLength"), 20)
            Case TYPE_成都莲合
                .Rows = 3
                .TextMatrix(1, 0) = "连接串"
                .TextMatrix(1, 1) = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("LHConnectionStrINg"), "dsn=lhyb;uid=sa;pwd=;")
                .TextMatrix(2, 0) = "医保内码"
                .TextMatrix(2, 1) = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("intercode"), 713)
            Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
                gstrSQL = "select 参数名,参数值 from 保险参数 where 险类=[1] and (中心=[2] or 中心 is null) order by 序号"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CInt(Mid(lvwKind_S.SelectedItem.Key, 2)), CInt(cmb中心.ItemData(cmb中心.ListIndex)))
                
                If rsTemp.RecordCount = 0 Then Exit Sub
                .Rows = rsTemp.RecordCount + 1
                lngRow = 1
                Do Until rsTemp.EOF
                    .TextMatrix(lngRow, 0) = rsTemp("参数名")
                    .TextMatrix(lngRow, 1) = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
                    
                    lngRow = lngRow + 1
                    rsTemp.MoveNext
                Loop
            Case Else
                '保险自贡医保，中联医保
                '如果中心为Null，表示该参数对所有中心有效
                
                '固定参数不能修改与查看，只能由
                gstrSQL = "select 参数名,参数值 from 保险参数 where 险类=[1] and (中心=[2] or 中心 is null) and 参数名 not like '%密码%' and 参数名 not in('卡验证码') and (是否固定<>1 Or 是否固定 Is null ) order by 序号"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CInt(Mid(lvwKind_S.SelectedItem.Key, 2)), CInt(cmb中心.ItemData(cmb中心.ListIndex)))
                
                If rsTemp.RecordCount = 0 Then Exit Sub
                
                .Rows = rsTemp.RecordCount + 1
                lngRow = 1
                Do Until rsTemp.EOF
                    .TextMatrix(lngRow, 0) = rsTemp("参数名")
                    Select Case rsTemp("参数名")
                        Case "卡种类"  '中联医保
                            .TextMatrix(lngRow, 1) = IIf(rsTemp("参数值") = "1", "输入型", "读取型")
                        Case "密码验证"
                            .TextMatrix(lngRow, 1) = IIf(rsTemp("参数值") = "1", "需要", "不需要")
                        Case "收费使用医保基金"
                            .TextMatrix(lngRow, 1) = IIf(rsTemp("参数值") = "1", "可以", "不可以")
                        Case "支持慢性病、特种病", "定点医疗机构", "传输数据", "门诊连续收费", "支持特殊门诊", "入院时选择参保前在院"
                            .TextMatrix(lngRow, 1) = IIf(rsTemp("参数值") = "1", "是", "否")
                        Case "收费个人帐户使用范围", "结算个人帐户使用范围"
                            strTemp = IIf(IsNull(rsTemp("参数值")), "00", rsTemp("参数值"))
                            
                            .TextMatrix(lngRow, 1) = IIf(Left(strTemp, 1) = "1", "全自费部分、", "") & _
                                                     IIf(Mid(strTemp, 2, 1) = "1", "首先自付部分、", "") & _
                                                     IIf(Mid(strTemp, 3, 1) = "1", "超限部分、", "")
                            If .TextMatrix(lngRow, 1) <> "" Then
                                .TextMatrix(lngRow, 1) = Mid(.TextMatrix(lngRow, 1), 1, Len(.TextMatrix(lngRow, 1)) - 1)
                            End If
                        Case "先扣起付线"
                            strTemp = IIf(IsNull(rsTemp("参数值")), "0", rsTemp("参数值"))
                            .TextMatrix(lngRow, 1) = IIf(strTemp = "1", "是", "否")
                        Case Else
                            If rsTemp!参数名 Like "*口令*" Then
                                .TextMatrix(lngRow, 1) = "********"
                            Else
                                .TextMatrix(lngRow, 1) = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
                            End If
                    End Select
                    
                    lngRow = lngRow + 1
                    rsTemp.MoveNext
                Loop
                
                Select Case Val(Mid(lvwKind_S.SelectedItem.Key, 2))
                    Case TYPE_重庆壁山
                        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, 0) = "当前使用的串口"
                        If IsNumeric(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "当前使用的串口", "")) = True Then
                            .TextMatrix(.Rows - 1, 1) = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "当前使用的串口", "") + 1
                        End If
                End Select
        End Select
    End With
End Sub

Private Sub Fill分段()
'功能：根据医保中心的序号显示年龄分段与费用档次
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long, lng险类 As Long, lng中心 As Long
    Dim str备注 As String
    
    '如果没有中心，肯定不会有显示的
    rsTemp.CursorLocation = adUseClient
    With msh分段
        .Clear
        .BackColor = &HFFFFFF
        If lvwKind_S.SelectedItem Is Nothing Or cmb中心.ListIndex < 0 Then
            '没有数据，显示一张空表
            .Rows = 9
            Set表头 0, "保险人群年龄段", True, 1
            Set表头 1, "名称,下限,上限,备注", False, 1
            Set表头 3, "支付费用档", True, 0
            Set表头 4, "名称,下限,上限,备注", False, 0
            Exit Sub
        End If
        
        lng险类 = Mid(lvwKind_S.SelectedItem.Key, 2)
        lng中心 = cmb中心.ItemData(cmb中心.ListIndex)
        
        gstrSQL = "select A.在职,A.名称,A.下限,A.上限,B.名称 as 人群名称,B.序号 " & _
                "   ,nvl(全额统筹,0) as 全额统筹,nvl(无起付线,0) as 无起付线,nvl(无封顶线,0) as 无封顶线" & _
                " from 保险年龄段 A ,保险人群 B" & _
                " where A.险类(+)=B.险类 and A.在职(+)=B.序号 and B.险类=[1] and A.中心(+)=[2]" & _
                " Order by A.在职,A.年龄段"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng险类, lng中心)
        
        If rsTemp.RecordCount = 0 Then
            .Rows = 3 '两个标题行
        Else
            .Rows = rsTemp.RecordCount + 2
        End If
        Set表头 0, "保险人群年龄段", True, 1
        Set表头 1, "名称,下限,上限,备注", False, 1
        lngRow = 2
        Do Until rsTemp.EOF
            .MergeRow(lngRow) = False
            .TextMatrix(lngRow, 0) = IIf(IsNull(rsTemp("名称")), rsTemp("人群名称"), rsTemp("名称"))
            .TextMatrix(lngRow, 1) = Format(rsTemp("下限"), "###;-###; ; ")
            .TextMatrix(lngRow, 2) = Format(rsTemp("上限"), "###;-###; ; ")
            
            If IsNull(rsTemp("名称")) = True Then
                str备注 = "尚未定义"
            Else
                str备注 = ""
                 If rsTemp("全额统筹") = 1 Then str备注 = ",全额统筹"
                 If rsTemp("无起付线") = 1 Then str备注 = str备注 & ",无起付线"
                 If rsTemp("无封顶线") = 1 Then str备注 = str备注 & ",无封顶线"
                 
                 str备注 = Mid(str备注, 2)
            End If
            .TextMatrix(lngRow, 3) = str备注
            
            .RowData(lngRow) = rsTemp("序号")
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        
        gstrSQL = "select 名称,下限,上限 from 保险费用档 " & _
            "where 险类=[1] and 中心=[2]"
        If lng险类 = TYPE_四川眉山 Then gstrSQL = gstrSQL & " And 档次<>0 "
        gstrSQL = gstrSQL & " Order by 档次"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng险类, lng中心)
        lngRow = .Rows
        If rsTemp.RecordCount = 0 Then
            .Rows = .Rows + 3 '两个标题行
        Else
            .Rows = .Rows + rsTemp.RecordCount + 2
        End If
        Set表头 lngRow, "保险支付费用档", True, 0
        Set表头 lngRow + 1, "名称,下限,上限,备注", False, 0
        lngRow = lngRow + 2
        Do Until rsTemp.EOF
            .MergeRow(lngRow) = False
            .TextMatrix(lngRow, 0) = rsTemp("名称")
            .TextMatrix(lngRow, 1) = Format(rsTemp("下限"), "########0.00;-########0.00; ; ")
            .TextMatrix(lngRow, 2) = Format(rsTemp("上限"), "########0.00;-########0.00; ; ")
            
            .RowData(lngRow) = 0
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
End Sub

Private Sub SetMenu()
'功能：根据当前的显示内容设置菜单可用性
    Dim blnItem As Boolean
    Dim lngIndex As Long
    Dim blnEnable As Boolean
    
    If lvwKind_S.SelectedItem Is Nothing Then
        '当前没有可设置的
        stbThis.Panels(2).Text = "共有医保类别" & lvwKind_S.ListItems.Count & "个"
        mnuEditModify.Enabled = False
        mnuEditDelete.Enabled = False
        mnuEditSelect.Enabled = False
        mnuEditDeselect.Enabled = False
        
        mnuCenterAdd.Enabled = False
        mnuCenterModify.Enabled = False
        mnuCenterDelete.Enabled = False
        mnuCenterParameter.Enabled = False
        mnuCenterSect.Enabled = False
    Else
        blnEnable = Val(Mid(lvwKind_S.SelectedItem.Key, 2)) <> TYPE_自贡市 And Val(Mid(lvwKind_S.SelectedItem.Key, 2)) <> TYPE_铜仁 And Val(Mid(lvwKind_S.SelectedItem.Key, 2)) <> TYPE_重庆中梁山
        
        stbThis.Panels(2).Text = "共有医保类别" & lvwKind_S.ListItems.Count & "个，所选的为" & lvwKind_S.SelectedItem.Text
        mnuEditModify.Enabled = True
        mnuEditDelete.Enabled = (Left(lvwKind_S.SelectedItem.Icon, 6) = "Common")
        mnuEditDeselect.Enabled = (Right(lvwKind_S.SelectedItem.Icon, 1) = "D")
        mnuEditSelect.Enabled = Not mnuEditDeselect.Enabled
        
        mnuCenterAdd.Enabled = cmb中心.Visible
        mnuCenterModify.Enabled = cmb中心.Visible And cmb中心.ListIndex > -1
        mnuCenterDelete.Enabled = mnuCenterModify.Enabled
        
        mnuCenterParameter.Enabled = cmb中心.ListIndex > -1
        mnuCenterSect.Enabled = mnuCenterParameter.Enabled And blnEnable
    End If
    
    For lngIndex = mnuCenterYear.LBound To mnuCenterYear.UBound
        mnuCenterYear(lngIndex).Enabled = mnuCenterSect.Enabled
    Next
    
    tbrThis.Buttons("Modify").Enabled = mnuEditModify.Enabled
    tbrThis.Buttons("Delete").Enabled = mnuEditDelete.Enabled
    tbrThis.Buttons("Select").Enabled = mnuEditSelect.Enabled
    tbrThis.Buttons("Parameter").Enabled = mnuCenterParameter.Enabled
End Sub

Private Sub 权限控制()
    If InStr(gstrPrivs, "增删改") = 0 Then
        tbrThis.Buttons("New").Visible = False
        tbrThis.Buttons("Modify").Visible = False
        tbrThis.Buttons("Delete").Visible = False
        tbrThis.Buttons("Split1").Visible = False
        
        mnuEditAdd.Visible = False
        mnuEditModify.Visible = False
        mnuEditDelete.Visible = False
        mnuEditSplit0.Visible = False
        
        mnuCenterAdd.Visible = False
        mnuCenterModify.Visible = False
        mnuCenterDelete.Visible = False
        mnuCenterSplitPara.Visible = False
    End If
    
    If InStr(gstrPrivs, "年龄段") = 0 Then
        mbln年龄段 = False
        mnuCenterYear(0).Visible = False
        mnuCenterSplitYear.Visible = False
    Else
        mbln年龄段 = True
    End If
    
    If InStr(gstrPrivs, "保险费用档") = 0 Then
        mnuCenterSect.Visible = False
        mnuCenterSplitSect.Visible = False
    End If
    
    If InStr(gstrPrivs, "运行参数设置") = 0 Then
        tbrThis.Buttons("Parameter").Visible = False
        tbrThis.Buttons("Split2").Visible = False
        
        If gstrPrivs = "基本" Then
            '完全屏蔽
            mnuCenter.Visible = False
            mnuCenterSplitPara.Visible = True
            mnuCenterParameter.Visible = False
        Else
            '只屏蔽相关项目
            mnuCenterParameter.Visible = False
            mnuCenterSplitPara.Visible = False
        End If
    End If
End Sub


Private Sub Set表头(ByVal lngRow As Long, strCaptions As String, ByVal blnMerge As Boolean, ByVal lngIndex As Long)
'功能：为费用档的几个子表设置表头
    Dim lngCol As Long
    Dim varCaptions As Variant
    
    With msh分段
        .MergeRow(lngRow) = blnMerge
        .RowData(lngRow) = lngIndex
        
        varCaptions = Split(strCaptions, ",")
        For lngCol = 0 To .Cols - 1
            If blnMerge = False Then
                .TextMatrix(lngRow, lngCol) = varCaptions(lngCol)
            Else
                .TextMatrix(lngRow, lngCol) = strCaptions
            End If
        Next
        
        .AllowBigSelection = True
        .Row = lngRow
        .COL = 0
        .RowSel = lngRow
        .ColSel = .Cols - 1
        .FillStyle = flexFillRepeat
        .CellAlignment = 4
        .CellBackColor = IIf(blnMerge, &HE6F5FD, &HD0D0D0)
        .CellForeColor = IIf(blnMerge, &H800000, 0)
        .FillStyle = flexFillSingle
        .AllowBigSelection = False
    End With
End Sub


