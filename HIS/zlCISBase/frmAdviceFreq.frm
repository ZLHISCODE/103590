VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAdviceFreq 
   Caption         =   "医嘱频率设置"
   ClientHeight    =   6690
   ClientLeft      =   3270
   ClientTop       =   2880
   ClientWidth     =   9540
   Icon            =   "frmAdviceFreq.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   9540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Tag             =   "12"
   Begin MSComctlLib.ImageList ils32 
      Left            =   2265
      Top             =   1965
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
            Picture         =   "frmAdviceFreq.frx":0442
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceFreq.frx":089A
            Key             =   "ItemNo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceFreq.frx":0CEE
            Key             =   "No"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceFreq.frx":1146
            Key             =   "Root"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   2280
      Top             =   1365
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
            Picture         =   "frmAdviceFreq.frx":159E
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceFreq.frx":19F6
            Key             =   "ItemNo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceFreq.frx":1E4A
            Key             =   "No"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceFreq.frx":22A2
            Key             =   "Root"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picNote 
      AutoRedraw      =   -1  'True
      Height          =   1005
      Left            =   3270
      ScaleHeight     =   945
      ScaleWidth      =   6120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5265
      Width           =   6180
      Begin VB.TextBox txtNote 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C00000&
         Height          =   780
         Left            =   465
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   90
         Width           =   5250
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshTime 
      Height          =   3975
      Left            =   3270
      TabIndex        =   1
      Top             =   1215
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   7011
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   280
      BackColorBkg    =   -2147483628
      GridColor       =   -2147483632
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   3255
      ScaleHeight     =   300
      ScaleWidth      =   6240
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   810
      Width           =   6240
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   30
         TabIndex        =   7
         Top             =   60
         Width           =   90
      End
   End
   Begin MSComctlLib.ListView lvwItem 
      Height          =   5025
      Left            =   60
      TabIndex        =   0
      Top             =   1215
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   8864
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils32"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "编码"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "名称"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "简码"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "英文名称"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "频率次数"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "频率间隔"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "间隔单位"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   6330
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   635
      SimpleText      =   "CoolBar1"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmAdviceFreq.frx":26FA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11748
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
   Begin ComCtl3.CoolBar cbrMain 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   9540
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbrMain"
      MinHeight1      =   720
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrMain 
         Height          =   720
         Left            =   30
         TabIndex        =   5
         Top             =   30
         Width           =   9420
         _ExtentX        =   16616
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ilsMenu"
         HotImageList    =   "ilsMenuHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   12
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "预览"
               Object.ToolTipText     =   "预览"
               Object.Tag             =   "预览"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "打印"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "增加"
               Key             =   "增加"
               Object.ToolTipText     =   "增加"
               Object.Tag             =   "增加"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修改"
               Key             =   "修改"
               Object.ToolTipText     =   "修改"
               Object.Tag             =   "修改"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "删除"
               Object.ToolTipText     =   "删除"
               Object.Tag             =   "删除"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "时间"
               Key             =   "时间"
               Object.ToolTipText     =   "调整频率缺省时间"
               Object.Tag             =   "时间"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "编辑_"
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查看"
               Key             =   "查看"
               Object.ToolTipText     =   "查看方式"
               Object.Tag             =   "查看"
               ImageIndex      =   7
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
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "帮助"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   990
      Top             =   270
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceFreq.frx":2F8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceFreq.frx":31A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceFreq.frx":33C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceFreq.frx":35DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceFreq.frx":37F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceFreq.frx":3A10
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceFreq.frx":3C2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceFreq.frx":3E44
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceFreq.frx":405E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsMenuHot 
      Left            =   345
      Top             =   270
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceFreq.frx":4278
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceFreq.frx":4492
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceFreq.frx":46AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceFreq.frx":48C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceFreq.frx":4AE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceFreq.frx":4CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceFreq.frx":4F14
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceFreq.frx":512E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceFreq.frx":5348
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip tabClass 
      Height          =   465
      Left            =   90
      TabIndex        =   2
      Top             =   870
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   820
      TabWidthStyle   =   2
      TabFixedWidth   =   1764
      TabFixedHeight  =   616
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "西医(&1)"
            Key             =   "XY"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "中医(&2)"
            Key             =   "ZY"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "固定项(&3)"
            Key             =   "GDX"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Image imgLR_s 
      Height          =   5370
      Left            =   3135
      MousePointer    =   9  'Size W E
      Top             =   885
      Width           =   45
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
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditTime 
         Caption         =   "时间设置(&T)"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuEdit_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditAdd 
         Caption         =   "增加(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditModi 
         Caption         =   "修改(&M)"
      End
      Begin VB.Menu mnuEditDel 
         Caption         =   "删除(&D)"
         Shortcut        =   {DEL}
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
      Begin VB.Menu mnuView_1 
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
         Checked         =   -1  'True
         Index           =   3
      End
      Begin VB.Menu mnuView_2 
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
Attribute VB_Name = "frmAdviceFreq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnItem As Boolean
Private mstrPrivs As String
Private mstrPreKey As String
Private mlngMode As Long
Private Sub Form_Load()
    Dim blnDo As Boolean
    
    mlngMode = glngModul
    mstrPrivs = gstrPrivs
    
    Call RestoreWinState(Me, App.ProductName)
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs)
    
    mnuViewIcon_Click lvwItem.View
    Call zlControl.PicShowFlat(picNote, -1)
    
    '权限设置
    If InStr(mstrPrivs, "增删改") = 0 And InStr(mstrPrivs, "时间安排") = 0 Then
        mnuEdit.Visible = False
        tbrMain.Buttons("增加").Visible = False
        tbrMain.Buttons("修改").Visible = False
        tbrMain.Buttons("删除").Visible = False
        tbrMain.Buttons("时间").Visible = False
        tbrMain.Buttons("编辑_").Visible = False
    ElseIf InStr(mstrPrivs, "增删改") = 0 Then
        mnuEditAdd.Visible = False
        mnuEditModi.Visible = False
        mnuEditDel.Visible = False
        mnuEdit_1.Visible = False
        tbrMain.Buttons("增加").Visible = False
        tbrMain.Buttons("修改").Visible = False
        tbrMain.Buttons("删除").Visible = False
    ElseIf InStr(mstrPrivs, "时间安排") = 0 Then
        mnuEditTime.Visible = False
        mnuEdit_1.Visible = False
        tbrMain.Buttons("时间").Visible = False
    End If
        
    mstrPreKey = ""
    Call LoadItems
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    imgLR_s.Top = IIf(cbrMain.Visible, cbrMain.Height, 0)
    imgLR_s.Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - IIf(cbrMain.Visible, cbrMain.Height, 0)
    
    tabClass.Top = IIf(cbrMain.Visible, cbrMain.Height, 0) + 30
    tabClass.Left = 30
    
    lvwItem.Left = 0
    lvwItem.Top = tabClass.Top + 345
    lvwItem.Width = imgLR_s.Left
    lvwItem.Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - lvwItem.Top
    
    picInfo.Left = imgLR_s.Left + imgLR_s.Width
    picInfo.Top = imgLR_s.Top
    picInfo.Width = Me.ScaleWidth - picInfo.Left
    
    With mshTime
        .Left = imgLR_s.Left + imgLR_s.Width
        .Top = picInfo.Top + picInfo.Height + 15
        .Width = Me.ScaleWidth - .Left
        .Height = imgLR_s.Height - picInfo.Height - picNote.Height - 45
    End With
    
    With picNote
        .Left = mshTime.Left
        .Top = mshTime.Top + mshTime.Height + 15
        .Width = mshTime.Width
    End With

    Call zlControl.PicShowFlat(picNote, -1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub lvwItem_Click()
    SetMenuEnable
End Sub

Private Sub lvwItem_DblClick()
    If mnuEdit.Visible And mnuEditModi.Enabled And mnuEditModi.Visible Then mnuEditModi_Click
End Sub

Private Sub lvwItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then lvwItem_DblClick
End Sub

Private Sub lvwItem_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And mnuEdit.Visible Then Me.PopupMenu mnuEdit
End Sub

Private Sub mnuEditAdd_Click()
    frmAdviceFreqEdit.mbytType = tabClass.SelectedItem.Index
    frmAdviceFreqEdit.Show 1
End Sub

Private Sub mnuEditDel_Click()
    Dim strSql As String, intIdx As Integer
    
    If lvwItem.SelectedItem Is Nothing Then Exit Sub
    If MsgBox("确实要删除该频率项目吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    strSql = "ZL_诊疗频率项目_DELETE('" & lvwItem.SelectedItem.Text & "')"
    
    On Error GoTo errH
    gcnOracle.BeginTrans
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    gcnOracle.CommitTrans
    On Error GoTo 0
    
    intIdx = lvwItem.SelectedItem.Index
    lvwItem.ListItems.Remove intIdx
    If lvwItem.ListItems.Count > 0 Then
        If intIdx <= lvwItem.ListItems.Count Then
            lvwItem.ListItems(intIdx).Selected = True
        Else
            lvwItem.ListItems(lvwItem.ListItems.Count).Selected = True
        End If
        lvwItem.SelectedItem.EnsureVisible
        Call lvwItem_ItemClick(lvwItem.SelectedItem)
    Else
        lblInfo.Caption = ""
        txtNote.Text = ""
        mshTime.Clear
        mshTime.ClearStructure
        mshTime.Rows = 2: mshTime.Cols = 2
        Call SetMenuEnable
    End If
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditModi_Click()
    If lvwItem.SelectedItem Is Nothing Then Exit Sub
    frmAdviceFreqEdit.mbytType = tabClass.SelectedItem.Index
    frmAdviceFreqEdit.mstrCode = lvwItem.SelectedItem.Text
    frmAdviceFreqEdit.Show 1
    If gblnOK Then mnuViewRefresh_Click
End Sub

Private Sub mnuFileExcel_Click()
    GrdPrint 3
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Public Function LoadItems(Optional strSeekKey As String) As Boolean
'功能：读取诊疗频率项目清单
'参数：strSeekKey;要强行定位的项目
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    Dim objItem As ListItem
    Dim strSaveKey As String
    Dim lngSelect As Long      '适用范围   1-西医，2-中医，3其他固定项
    
    On Error GoTo errH
    
    If Not lvwItem.SelectedItem Is Nothing Then strSaveKey = lvwItem.SelectedItem.Key
    lvwItem.ListItems.Clear
    lngSelect = tabClass.SelectedItem.Index
    If lngSelect <> 1 And lngSelect <> 2 Then
        strSql = " Nvl(适用范围, 0) <> 1 And Nvl(适用范围, 0) <> 2 "
    Else
        strSql = " Nvl(适用范围, 0) = [1] "
    End If
    
    strSql = "Select * From 诊疗频率项目 Where " & strSql & " Order by 编码"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngSelect)
    
    If Not rsTmp.EOF Then
        Do While Not rsTmp.EOF
            Set objItem = lvwItem.ListItems.Add(, "_" & rsTmp!编码, rsTmp!编码, "Root", "Root")
            objItem.Tag = IIf(IsNull(rsTmp!频率次数), 0, rsTmp!频率次数)
            For i = 2 To lvwItem.ColumnHeaders.Count
                objItem.SubItems(i - 1) = Nvl(Nvl(rsTmp(lvwItem.ColumnHeaders(i).Text).Value))
            Next
            '定位
            If objItem.Key = strSaveKey And strSeekKey = "" Then objItem.Selected = True
            If objItem.Key = strSeekKey Then objItem.Selected = True
            rsTmp.MoveNext
        Loop
    Else
        lblInfo.Caption = ""
        txtNote.Text = ""
        mshTime.Clear
        mshTime.ClearStructure
        mshTime.FixedRows = 0
        mshTime.FixedCols = 0
        mshTime.Rows = 0
        mshTime.Cols = 0
    End If
    LoadItems = True
    
    If Not (lvwItem.SelectedItem Is Nothing) Then
        lvwItem.SelectedItem.EnsureVisible
        Call lvwItem_ItemClick(lvwItem.SelectedItem)
    Else
        Call SetMenuEnable
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub mnuFilePreview_Click()
    GrdPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    GrdPrint 1
End Sub

Private Sub mnuFilePrintSet_Click()
    zlPrintSet
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage hWnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo hWnd
End Sub

Private Sub mnuhelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuEditTime_Click()
    If lvwItem.SelectedItem Is Nothing Then Exit Sub
    
    On Error Resume Next
    frmAdviceFreqTime.mstrCode = lvwItem.SelectedItem.Text
    frmAdviceFreqTime.Show 1
    If gblnOK Then
        mstrPreKey = ""
        Call lvwItem_ItemClick(lvwItem.SelectedItem)
    End If
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    '默认参数：编码=项目编码，范围=适用范围(1-西医 2-中医 3-固定项)
    Dim str编码 As String
    Dim int范围 As Integer
    
    If Not Me.lvwItem.SelectedItem Is Nothing Then
        str编码 = Mid(Me.lvwItem.SelectedItem.Key, 2)
    End If
    
    int范围 = tabClass.SelectedItem.Index
    
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
        "编码=" & IIf(str编码 = "", "", str编码), _
        "范围=" & int范围)
End Sub

Private Sub mnuViewToolButton_Click()
    tbrMain.Visible = Not tbrMain.Visible
    cbrMain.Visible = Not cbrMain.Visible
    mnuViewToolButton.Checked = cbrMain.Visible
    If mnuViewToolButton.Checked = False Then
        mnuViewToolText.Enabled = False
    Else
        mnuViewToolText.Enabled = True
    End If
    Form_Resize
End Sub

Private Sub mnuViewIcon_Click(Index As Integer)
    mnuViewIcon(0).Checked = False
    mnuViewIcon(1).Checked = False
    mnuViewIcon(2).Checked = False
    mnuViewIcon(3).Checked = False
    lvwItem.View = Index
    mnuViewIcon(lvwItem.View).Checked = True
End Sub

Private Sub mnuViewRefresh_Click()
    Call tabClass_Click
End Sub

Private Sub mnuViewStatus_Click()
    stbThis.Visible = Not stbThis.Visible
    mnuViewStatus.Checked = stbThis.Visible
    
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim i As Long
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For i = 1 To tbrMain.Buttons.Count
        tbrMain.Buttons.Item(i).Caption = IIf(mnuViewToolText.Checked, tbrMain.Buttons.Item(i).Key, "")
    Next
    cbrMain.Bands(1).MinHeight = tbrMain.Height
    
    Form_Resize
End Sub

Private Sub mshTime_DblClick()
    If mnuEdit.Visible And mnuEditTime.Visible And mnuEditTime.Enabled Then mnuEditTime_Click
End Sub

Private Sub mshTime_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then mshTime_DblClick
End Sub

Private Sub picNote_Resize()
    On Error Resume Next
    With txtNote
        .Width = picNote.Width - .Left * 2
        .Height = picNote.Height - .Top * 2
    End With
    Call zlControl.PicShowFlat(picNote, -1)
End Sub

Private Sub imgLR_s_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    
    imgLR_s.Left = imgLR_s.Left + x
    If imgLR_s.Left < tabClass.Width + 30 Then imgLR_s.Left = tabClass.Width + 30
    If Me.Width - imgLR_s.Left - imgLR_s.Width < 1000 Then imgLR_s.Left = Me.Width - imgLR_s.Width - 1000
    Form_Resize
End Sub

Private Sub tabClass_Click()
    mstrPreKey = ""
    Call LoadItems
    If Visible Then lvwItem.SetFocus
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "增加"
            mnuEditAdd_Click
        Case "修改"
            mnuEditModi_Click
        Case "删除"
            mnuEditDel_Click
        Case "帮助"
            
        Case "预览"
            mnuFilePreview_Click
        Case "打印"
            mnuFilePrint_Click
        Case "时间"
            mnuEditTime_Click
        Case "查看"
            lvwItem.View = IIf(lvwItem.View = 3, 0, lvwItem.View + 1)
            mnuViewIcon_Click lvwItem.View
        Case "退出"
            mnuFileExit_Click
    End Select
End Sub

Private Sub tbrMain_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Text
        Case "大图标"
            mnuViewIcon_Click 0
        Case "小图标"
            mnuViewIcon_Click 1
        Case "列表"
            mnuViewIcon_Click 2
        Case "详细资料"
            mnuViewIcon_Click 3
    End Select
End Sub

Private Sub lvwItem_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call zlControl.LvwSortColumn(lvwItem, ColumnHeader.Index)
End Sub

Private Sub lvwItem_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Item.Key <> mstrPreKey Then
        mstrPreKey = Item.Key
        
        mshTime.Clear
        mshTime.ClearStructure
        mshTime.FixedCols = 0: mshTime.FixedRows = 0
        mshTime.Rows = 0: mshTime.Cols = 0
        
        If Mid(Item.Key, 2) > 0 Then
            Call ShowTimeScheme(Mid(Item.Key, 2))
        ElseIf Mid(mstrPreKey, 2) = -1 Then
            txtNote.Text = "说明：" & vbCrLf & "    该类项目只能用于临嘱，固定为一次性执行。"
        ElseIf Mid(mstrPreKey, 2) = -2 Then
            txtNote.Text = "说明：" & vbCrLf & "    该类项目只能用于长嘱，如护理等级。护理等级固定为持续性，按天计时执行。"
        ElseIf Mid(mstrPreKey, 2) = -3 Then
            txtNote.Text = "说明：" & vbCrLf & "    该类项目只能用于长嘱，适用于长期备用医嘱，必要时才发送执行。"
        ElseIf Mid(mstrPreKey, 2) = -4 Then
            txtNote.Text = "说明：" & vbCrLf & "    该类项目只能用于长嘱，适用于持续性长嘱。"
        ElseIf Mid(mstrPreKey, 2) = -5 Then
            txtNote.Text = "说明：" & vbCrLf & "    该类项目只能用于临嘱，适用于临时备用医嘱，在12小时内需要时发送执行，超过12小时未执行则自动停止。"
        End If
        lblInfo.Caption = "【" & Item.SubItems(1) & "】的缺省执行时间"
        Call SetMenuEnable
    End If
End Sub

Private Sub GrdPrint(intMode As Byte)
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    
    If gstrUserName = "" Then Call GetUserInfo
        
    objPrint.Title = "【" & lvwItem.SelectedItem.SubItems(1) & "】的缺省执行时间"
    
    Set objRow = New zlTabAppRow
    objRow.Add ""
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add ""
    objPrint.UnderAppRows.Add objRow
            
    Set objRow = New zlTabAppRow
    objRow.Add "打印人:" & gstrUserName
    objRow.Add "打印时间:" & Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
    objPrint.BelowAppRows.Add objRow
    
    Set objPrint.Body = mshTime
    If intMode = 1 Then
        Select Case zlPrintAsk(objPrint)
        Case 1
            zlPrintOrView1Grd objPrint, 1
        Case 2
            zlPrintOrView1Grd objPrint, 2
        Case 3
            zlPrintOrView1Grd objPrint, 3
        End Select
    Else
        zlPrintOrView1Grd objPrint, intMode
    End If
End Sub

Private Sub tbrMain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuViewTool, 2
End Sub

Private Sub SetMenuEnable()
    If tabClass.SelectedItem.Index = 3 Then
        mnuEditTime.Enabled = False
        mnuEditModi.Enabled = False
        mnuEditDel.Enabled = False
        mnuFilePreview.Enabled = False
        mnuFilePrint.Enabled = False
        mnuFileExcel.Enabled = False
        mnuEditAdd.Enabled = False
    Else
        mnuEditTime.Enabled = True
        mnuEditModi.Enabled = True
        mnuEditDel.Enabled = True
        mnuFilePreview.Enabled = True
        mnuFilePrint.Enabled = True
        mnuFileExcel.Enabled = True
        mnuEditAdd.Enabled = True
    End If
                
    If lvwItem.SelectedItem Is Nothing Then
        mnuEditTime.Enabled = False
        mnuEditModi.Enabled = False
        mnuEditDel.Enabled = False
    ElseIf lvwItem.SelectedItem.SubItems(6) = "分钟" Then
        '分钟间隔不需要设置时间
        mnuEditTime.Enabled = False
    End If
    
                
    If lvwItem.ListItems.Count = 0 Then
        mnuFilePreview.Enabled = False
        mnuFilePrint.Enabled = False
        mnuFileExcel.Enabled = False
    End If
        
    tbrMain.Buttons("增加").Enabled = mnuEditAdd.Enabled
    tbrMain.Buttons("删除").Enabled = mnuEditDel.Enabled
    tbrMain.Buttons("修改").Enabled = mnuEditModi.Enabled
    tbrMain.Buttons("时间").Enabled = mnuEditTime.Enabled
    tbrMain.Buttons("预览").Enabled = mnuFilePreview.Enabled
    tbrMain.Buttons("打印").Enabled = mnuFilePrint.Enabled
    
    stbThis.Panels(2).Text = "共有 " & lvwItem.ListItems.Count & " 个频率项目"
End Sub

Private Function ShowTimeScheme(ByVal str编码 As String) As Boolean
'功能：根据当前频率项目显示它的时间方案表
'参数：str编码=频率项目编码
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long, j As Long
    Dim lng频率次数 As Long, lng频率间隔 As Long, str间隔单位 As String
    Dim arrTime As Variant
    
    On Error GoTo errH
    
    With mshTime
        
        '频率项目信息
        strSql = "Select * From 诊疗频率项目 Where 编码=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, str编码)
        If Not rsTmp.EOF Then
            str间隔单位 = IIf(IsNull(rsTmp!间隔单位), "", rsTmp!间隔单位)
            lng频率次数 = IIf(IsNull(rsTmp!频率次数), 0, rsTmp!频率次数)
            lng频率间隔 = IIf(IsNull(rsTmp!频率间隔), 0, rsTmp!频率间隔)
        End If
        If str间隔单位 = "分钟" Then
            txtNote.Text = "说明：" & vbCrLf & "    间隔单位为""分钟""，不需要指定执行时间。相对于医嘱开始执行时间为准计算执行时间。"
        ElseIf str间隔单位 = "小时" Then
            txtNote.Text = "说明：" & vbCrLf & "    间隔单位为""小时""，时间格式为""小时数[:分钟数]""。小时数相对于医嘱开始时间，分钟数为0时可以不写。如：1次/3小时 = 2:30，表示医嘱开始后第2小时内的30分钟第一次执行。"
        ElseIf str间隔单位 = "天" Then
            If lng频率间隔 = 1 Then
                txtNote.Text = "说明：" & vbCrLf & "    间隔单位为""天""，频率间隔为""1""，时间格式为""绝对时间""。如：3次/天 = 8:00-12:00-16:00 或 8:12:16，表示每天的8点,12点,16点执行。"
            Else
                txtNote.Text = "说明：" & vbCrLf & "    间隔单位为""天""，频率间隔大于""1""，时间格式为""相对天数/绝对时间""。如：1次/3天 = 2/8:00 或2/8，表示医嘱开始的第2天8:00点第一次执行。"
            End If
        ElseIf str间隔单位 = "周" Then
            txtNote.Text = "说明：" & vbCrLf & "    间隔单位为""周""，时间格式为""星期数/绝对时间""，星期数用1-7表示星期一到星期日。如：3次/周 = 1/8:00-3/8:00-5/8:00，表示医嘱开始后在每周星期一、三、五的8:00执行。"
        End If
        
        '频率时间方案
        strSql = _
            "Select A.方案序号,A.时间方案,B.名称,B.编码" & _
            " From 诊疗频率时间 A,诊疗项目目录 B" & _
            " Where A.给药途径ID=B.ID(+) And A.执行频率=[1]" & _
            " Order by 方案序号"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, str编码)
        
        '时间方案表头
        If str间隔单位 = "分钟" Then
            .FixedRows = 0: .Cols = 2: .Rows = 0
        ElseIf str间隔单位 = "周" Or str间隔单位 = "天" And lng频率间隔 > 1 Then
            .Cols = 2 + lng频率次数 * 2
            .Rows = IIf(rsTmp.EOF, 1, rsTmp.RecordCount) + 2
            .FixedRows = 2
            .FixedCols = 1
            
            .TextMatrix(0, 0) = "序号": .TextMatrix(1, 0) = .TextMatrix(0, 0)
            .TextMatrix(0, 1) = IIf(tabClass.SelectedItem.Index = 1, "给药途径", "中药用法"): .TextMatrix(1, 1) = .TextMatrix(0, 1)
            For i = 2 To .Cols - 1 Step 2
                .TextMatrix(0, i) = "第" & ((i - 2) \ 2) + 1 & "次"
                .TextMatrix(0, i + 1) = .TextMatrix(0, i)
                If str间隔单位 = "周" Then
                    .TextMatrix(1, i) = "星期"
                    .TextMatrix(1, i + 1) = "时间"
                    .ColWidth(i) = 450
                    .ColWidth(i + 1) = 1000
                Else
                    .TextMatrix(1, i) = "天"
                    .TextMatrix(1, i + 1) = "时间"
                    .ColWidth(i) = 300
                    .ColWidth(i + 1) = 1000
                End If
                .ColAlignment(i) = 4
                .ColAlignment(i + 1) = 1
            Next
        ElseIf str间隔单位 = "小时" Or str间隔单位 = "天" And lng频率间隔 = 1 Then
            .Cols = 2 + lng频率次数
            .Rows = IIf(rsTmp.EOF, 1, rsTmp.RecordCount) + 1
            .FixedRows = 1
            .FixedCols = 1
            
            .TextMatrix(0, 0) = "序号"
            .TextMatrix(0, 1) = IIf(tabClass.SelectedItem.Index = 1, "给药途径", "中药用法")
            For i = 2 To .Cols - 1
                .TextMatrix(0, i) = "第" & i - 1 & "次"
                .ColWidth(i) = 1000
                .ColAlignment(i) = 1
            Next
        End If
        .ColWidth(0) = 450
        .ColWidth(1) = 1500
        .ColAlignment(0) = 4
        .ColAlignment(1) = 1
        For i = 0 To .Cols - 1
            .ColAlignmentFixed(i) = 4
        Next
        .MergeCells = flexMergeRestrictAll
        .MergeCol(0) = True: .MergeCol(1) = True
        If .Rows <> 0 Then
            .MergeRow(0) = True: .MergeRow(1) = True
        End If
        
        '时间数据
        For i = 1 To rsTmp.RecordCount
            .TextMatrix(i + .FixedRows - 1, 0) = rsTmp!方案序号
            .TextMatrix(i + .FixedRows - 1, 1) = IIf(IsNull(rsTmp!名称), "<不确定>", rsTmp!名称)
            
            arrTime = Split(rsTmp!时间方案, "-")
            If str间隔单位 = "周" Or str间隔单位 = "天" And lng频率间隔 > 1 Then
                For j = 0 To lng频率次数 - 1
                    .TextMatrix(i + .FixedRows - 1, j * 2 + 2) = Split(arrTime(j), "/")(0)
                    .TextMatrix(i + .FixedRows - 1, j * 2 + 3) = Split(arrTime(j), "/")(1)
                Next
            ElseIf str间隔单位 = "小时" Or str间隔单位 = "天" And lng频率间隔 = 1 Then
                For j = 0 To lng频率次数 - 1
                    .TextMatrix(i + .FixedRows - 1, j + 2) = arrTime(j)
                Next
            End If
            rsTmp.MoveNext
        Next
    End With
    
    ShowTimeScheme = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub
