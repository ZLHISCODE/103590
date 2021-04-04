VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmIllManage 
   BackColor       =   &H8000000A&
   Caption         =   "疾病编码管理"
   ClientHeight    =   6750
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   8730
   Icon            =   "frmIllManage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6750
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cdgFile 
      Left            =   6000
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2625
      Left            =   4980
      MousePointer    =   9  'Size W E
      ScaleHeight     =   2625
      ScaleWidth      =   45
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2070
      Width           =   45
   End
   Begin MSComctlLib.ListView lvwMain 
      Height          =   2475
      Left            =   4050
      TabIndex        =   2
      Top             =   2070
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   4366
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
      BackColor       =   16777215
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView tvwMain_S 
      Height          =   4575
      Left            =   150
      TabIndex        =   1
      Top             =   1170
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   8070
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   4380
      Top             =   5490
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":030A
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":075E
            Key             =   "Stop"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   3810
      Top             =   5430
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":0BB0
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":1004
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":1458
            Key             =   "Stop"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar coolbar1 
      Align           =   1  'Align Top
      Height          =   1125
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   1984
      BandCount       =   2
      _CBWidth        =   8730
      _CBHeight       =   1125
      _Version        =   "6.7.9782"
      Child1          =   "Toolbar1"
      MinHeight1      =   720
      Width1          =   8535
      Key1            =   "only"
      NewRow1         =   0   'False
      Caption2        =   "编码类别"
      Child2          =   "cmbType"
      MinWidth2       =   3495
      MinHeight2      =   300
      Width2          =   1590
      FixedBackground2=   0   'False
      NewRow2         =   -1  'True
      Begin VB.ComboBox cmbType 
         Height          =   300
         ItemData        =   "frmIllManage.frx":18AA
         Left            =   945
         List            =   "frmIllManage.frx":18C3
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   780
         Width           =   7695
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   720
         Left            =   165
         TabIndex        =   5
         Top             =   30
         Width           =   8475
         _ExtentX        =   14949
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "Ilsrw"
         HotImageList    =   "Ilscolor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   17
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "Preview"
               Object.ToolTipText     =   "预览"
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
               Key             =   "Split0"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "分类"
               Key             =   "Class"
               Object.ToolTipText     =   "增加分类"
               Object.Tag             =   "分类"
               ImageKey        =   "Class"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "疾病"
               Key             =   "Disease"
               Object.ToolTipText     =   "增加疾病"
               Object.Tag             =   "疾病"
               ImageKey        =   "Disease"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修改"
               Key             =   "Modify"
               Object.ToolTipText     =   "修改"
               Object.Tag             =   "修改"
               ImageKey        =   "Modify"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "Delete"
               Object.ToolTipText     =   "删除"
               Object.Tag             =   "删除"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split1"
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "启用"
               Key             =   "启用"
               Object.ToolTipText     =   "启用疾病编码"
               Object.Tag             =   "启用"
               ImageKey        =   "Start"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "停用"
               Key             =   "停用"
               Object.ToolTipText     =   "停用疾病编码"
               Object.Tag             =   "停用"
               ImageKey        =   "Stop"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Splits"
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查找"
               Key             =   "Find"
               Description     =   "查找"
               Object.ToolTipText     =   "查找"
               Object.Tag             =   "查找"
               ImageKey        =   "Find"
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split3"
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查看"
               Key             =   "View"
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
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split4"
               Style           =   3
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Quit"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList Ilsrw 
      Left            =   6900
      Top             =   1020
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
            Picture         =   "frmIllManage.frx":1925
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":1B3F
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":1D59
            Key             =   "Class"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":1F75
            Key             =   "Disease"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":2191
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":23B1
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":25D1
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":27EB
            Key             =   "View"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":2A0B
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":2C2B
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":2E4B
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":3065
            Key             =   "Stop"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Ilscolor 
      Left            =   5880
      Top             =   1020
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
            Picture         =   "frmIllManage.frx":327F
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":349F
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":36BF
            Key             =   "Class"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":38DB
            Key             =   "Disease"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":3AF7
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":3D17
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":3F37
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":4151
            Key             =   "View"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":4371
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":4591
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":47B1
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllManage.frx":49CB
            Key             =   "Stop"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   6390
      Width           =   8730
      _ExtentX        =   15399
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
            Picture         =   "frmIllManage.frx":4BE5
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10319
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
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuFilePrintView 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileEXCEL 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnuFileImport 
         Caption         =   "导入疾病编码"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuFileExport 
         Caption         =   "导出疾病编码"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuFileLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditClass 
         Caption         =   "增加分类(&C)"
      End
      Begin VB.Menu mnuEditDisease 
         Caption         =   "增加疾病(&D)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "修改(&M)"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "删除(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditStart 
         Caption         =   "启用(&S)"
      End
      Begin VB.Menu mnuEditStop 
         Caption         =   "停用(&T)"
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
            Caption         =   "标准文本(&S)"
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
      Begin VB.Menu mnuViewLine1 
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
      Begin VB.Menu mnuViewLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewOrder 
         Caption         =   "检查序号排列(&O)"
      End
      Begin VB.Menu mnuViewSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewStop 
         Caption         =   "显示停用项目(&P)"
      End
      Begin VB.Menu mnuViewLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewColumn 
         Caption         =   "选择列(&C)"
      End
      Begin VB.Menu mnuViewAll 
         Caption         =   "全部显示(&A)"
      End
      Begin VB.Menu mnuViewLine4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "查找(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "刷新(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "帮助主题(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "Web上的中联"
         Begin VB.Menu mnuHelpWebHomePage 
            Caption         =   "中联主页(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "中联论坛(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "发送反馈(&K)..."
         End
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
   Begin VB.Menu mnuShort1 
      Caption         =   "快捷菜单"
      Visible         =   0   'False
      Begin VB.Menu mnushortMenu1 
         Caption         =   "增加分类(&C)"
         Index           =   1
      End
      Begin VB.Menu mnushortMenu1 
         Caption         =   "修改分类(&M)"
         Index           =   2
      End
      Begin VB.Menu mnushortMenu1 
         Caption         =   "删除分类(&D)"
         Index           =   3
      End
   End
   Begin VB.Menu mnuShort2 
      Caption         =   "快捷菜单"
      Visible         =   0   'False
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "增加疾病(&D)"
         Index           =   1
      End
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "修改疾病(&M)"
         Index           =   2
      End
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "删除疾病(&D)"
         Index           =   3
      End
      Begin VB.Menu mnuShortLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "大图标(&G)"
         Index           =   0
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "小图标(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "列表(&L)"
         Index           =   2
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "详细资料(&D)"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmIllManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mintColumn As Integer
Dim mstr编码类别 As String
Dim mblnLoad As Boolean

Dim msngStartX As Single    '移动前鼠标的位置
Dim mlng序号 As Long
Dim mstrNodeKey As String, mstrTypeText As String
'每节中四段的含义依次是名称、宽度、对齐、可选性
Private Const mstr疾病 As String = "编码,1200,0,1;附码,1200,0,0;名称,2440,0,2;拼音码,1000,0,0;五笔码,1000,0,0;性别限制,800,0,0;手术类型,800,0,0;统计码,800,0,0;提醒疗效,800,0,0;分娩信息,800,0,0;说明,3000,0,0;建档时间,1400,0,0;撤档时间,1400,0,0"

Private mlngMode As Long
Private mstrPrivs As String '权限串
Private mconnExcel As New ADODB.Connection
Private mintProgress As Integer

Private Sub Form_Load()
    Dim i As Long
    
    mlngMode = glngModul
    mstrPrivs = gstrPrivs
    
    Call 权限控制
    '允许进行列删除的ListView须做标记
    lvwMain.Tag = "可变化的"
    RestoreWinState Me, App.ProductName
    Call zldatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs)
    
    For i = 0 To 3
        Me.mnuViewIcon(i).Checked = False
    Next
    Me.mnuViewIcon(Me.lvwMain.View).Checked = True
    
    mnuViewAll.Checked = (Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "显示所有", 0)) = 1)
    
    mblnLoad = True
    '如果ListView的还未被设置，比如第一次使用，那就调用缺省的初始化
'    If lvwMain.ColumnHeaders.Count = 0 Then
        zlControl.LvwSelectColumns lvwMain, mstr疾病, True
'    End If
End Sub

Private Sub Form_Activate()
    On Error GoTo errHandle
    
    If mblnLoad = True Then
        mblnLoad = False '马上把它改过来
                
        '20031112byZT：通过权限判断是否使用中医
        gbln使用中医 = InStr(mstrPrivs, "中医") > 0
        
        '初始化编码类别列表框的内容
        Call Fill类别
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Unload Me
End Sub

Private Sub Form_Resize()
    If WindowState = 1 Then Exit Sub
    
    Dim sngTop As Single, sngBottom As Single
    On Error Resume Next
    sngTop = IIF(CoolBar1.Visible, CoolBar1.Top + CoolBar1.Height, 0)
    sngBottom = Me.ScaleHeight - IIF(stbThis.Visible, stbThis.Height, 0)
    
    tvwMain_S.Left = ScaleLeft
    tvwMain_S.Top = sngTop
    tvwMain_S.Height = sngBottom - sngTop
    
    With picSplit
        .Left = tvwMain_S.Left + tvwMain_S.Width
        .Top = tvwMain_S.Top
        .Height = tvwMain_S.Height
    End With
    
    lvwMain.Top = tvwMain_S.Top
    lvwMain.Height = tvwMain_S.Height
    
    If tvwMain_S.Visible = True Then
        lvwMain.Left = picSplit.Left + picSplit.Width
    Else
        lvwMain.Left = ScaleLeft
    End If
    lvwMain.Width = ScaleWidth - lvwMain.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mconnExcel.State = 1 Then mconnExcel.Close
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "显示所有", IIF(mnuViewAll.Checked, 1, 0)
    SaveWinState Me, App.ProductName
End Sub

Private Sub cmbType_Click()
    If cmbType.Text = mstrTypeText Then Exit Sub
    
    If cmbType.ItemData(cmbType.ListIndex) = 1 Then
        tvwMain_S.Visible = True
        picSplit.Visible = True
    Else
        tvwMain_S.Visible = False
        picSplit.Visible = False
    End If
    Call Form_Resize
    
    Call FillTree
End Sub

Private Sub coolbar1_HeightChanged(ByVal NewHeight As Single)
    Form_Resize
End Sub

Private Sub lvwMain_DblClick()
    
    If mnuEditModify.Visible And mnuEditModify.Enabled Then
        '对当前项目进行编辑
        Call mnuEditModify_Click
    End If

End Sub

Private Sub lvwMain_GotFocus()
    Call SetMenu
End Sub


Private Sub lvwMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
     Call SetMenu
End Sub

Private Sub lvwMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    
    If Button = 2 Then
        mnuShortMenu2(1).Enabled = mnuEditDisease.Enabled
        mnuShortMenu2(2).Enabled = mnuEditModify.Enabled
        mnuShortMenu2(3).Enabled = mnuEditDelete.Enabled
        For i = 0 To 3
            mnuShortIcon(i).Checked = mnuViewIcon(i).Checked
        Next
        PopupMenu mnuShort2, vbPopupMenuRightButton
    End If
End Sub

Private Sub lvwMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintColumn = ColumnHeader.Index - 1 Then '仍是刚才那列
        lvwMain.SortOrder = IIF(lvwMain.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvwMain.SortKey = mintColumn
        lvwMain.SortOrder = lvwAscending
    End If
    If Not lvwMain.SelectedItem Is Nothing Then
        lvwMain.SelectedItem.EnsureVisible
    End If
End Sub


Private Sub mnuEditStart_Click()
    Call StopAndResume(False)
End Sub

Private Sub mnuEditDelete_Click()
'删除
    Dim strKey As String
    Dim intIndex As Long
    
    If AllowContinue = False Then Exit Sub
    
    On Error GoTo errHandle
    If ActiveControl Is tvwMain_S Then
        If MsgBox("你确认要删除名称为“" & tvwMain_S.SelectedItem.Text & "”的分类吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
            Me.MousePointer = 11
            
            gstrSQL = "ZL_疾病编码分类_delete(" & Mid(tvwMain_S.SelectedItem.Key, 2) & ")"
            Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            
            Me.MousePointer = 0
            
            strKey = tvwMain_S.SelectedItem.Key
            If Not tvwMain_S.SelectedItem.Next Is Nothing Then
                tvwMain_S.SelectedItem.Next.Selected = True
            Else
                If Not tvwMain_S.SelectedItem.Parent Is Nothing Then
                    tvwMain_S.SelectedItem.Parent.Selected = True
                End If
            End If
            Call FillList
            tvwMain_S.Nodes.Remove strKey
        End If
    Else
        If MsgBox("你确认要删除名称为“" & lvwMain.SelectedItem.Text & "”的疾病吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
            Me.MousePointer = 11
            
            On Error Resume Next
            gstrSQL = "ZL_疾病编码目录_delete(" & Mid(lvwMain.SelectedItem.Key, 2) & ")"
            Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            If Err.Number <> 0 Then
                If InStr(Err.Description, "ORA-20005") > 0 Then
                    MsgBox "项目已经使用不能删除，只能停用", vbInformation, gstrSysName
   
                Else
                    MsgBox Err.Description, vbInformation, gstrSysName
                End If
                
                Me.MousePointer = 0
                Exit Sub
            End If
            
            Me.MousePointer = 0
            
            With lvwMain
                intIndex = .SelectedItem.Index
                .ListItems.Remove .SelectedItem.Key
                If .ListItems.Count > 0 Then
                    intIndex = IIF(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
                    .ListItems(intIndex).Selected = True
                    .ListItems(intIndex).EnsureVisible
                End If
                Call SetMenu
            End With
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Me.MousePointer = 0
End Sub



Private Sub mnuEditModify_Click()
    Dim nodTemp As Node
    Dim str名称 As String
    
    If AllowContinue = False Then Exit Sub
    
    Set nodTemp = tvwMain_S.SelectedItem
    If ActiveControl Is tvwMain_S And tvwMain_S.Visible = True Then
        '修改分类
        If nodTemp Is Nothing Then
            Exit Sub
        Else
            
            With tvwMain_S.SelectedItem
                frmIllSortEdit.疾病编辑 "", "", mstr编码类别, Mid(nodTemp.Key, 2)
            End With
        End If
    Else
        '修改疾病
        If lvwMain.SelectedItem Is Nothing Then Exit Sub
        If nodTemp Is Nothing Then
            If tvwMain_S.Visible = True Then Exit Sub '这种情况是不允许
            
            frmIllItemEdit.疾病编辑 tvwMain_S.Visible, "无", "", mstr编码类别, Mid(lvwMain.SelectedItem.Key, 2)
        Else
            
            frmIllItemEdit.疾病编辑 tvwMain_S.Visible, nodTemp.Text, Mid(nodTemp.Key, 2), mstr编码类别, Mid(lvwMain.SelectedItem.Key, 2)
        End If
    End If
End Sub

Private Sub mnuEditClass_Click()
    Dim nodTemp As Node
    Dim str名称 As String
    
    If AllowContinue = False Then Exit Sub
    
    Set nodTemp = tvwMain_S.SelectedItem
    '增加分类
    If nodTemp Is Nothing Then
        frmIllSortEdit.疾病编辑 "无", "", mstr编码类别
    Else
        frmIllSortEdit.疾病编辑 nodTemp.Text, Mid(nodTemp.Key, 2), mstr编码类别
    End If
End Sub

Private Sub mnuEditDisease_Click()
    Dim nodTemp As Node
    Dim str名称 As String
    
    If AllowContinue = False Then Exit Sub
    
    Set nodTemp = tvwMain_S.SelectedItem
    '增加疾病
    If nodTemp Is Nothing Then
        If tvwMain_S.Visible = True Then Exit Sub '这种情况是不允许
        
        frmIllItemEdit.疾病编辑 tvwMain_S.Visible, "无", "", mstr编码类别
    Else
        If Mid(nodTemp.Key, 2) = 1 Then
            If MsgBox("在该项目下增加疾病会引起系统自带报表的计算错误，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        
        frmIllItemEdit.疾病编辑 tvwMain_S.Visible, nodTemp.Text, Mid(nodTemp.Key, 2), mstr编码类别
    End If
End Sub

Private Function AllowContinue() As Boolean
'检查是否允许继续编辑编码
    If MsgBox("国际疾病分类需要统一的标准，这是一件非常严肃的事，" & vbCrLf & _
        "你最好能在当地卫生统计的权威机构指导下完成本操作。" & vbCrLf & vbCrLf & _
        "是否继续？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    
    AllowContinue = True
End Function

Private Sub mnuEditStop_Click()
    Call StopAndResume(True)
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnuFileExport_Click()
    Dim strPath As String
    strPath = zlcommfun.OpenDir(Me.hwnd, "导出目录", App.Path)
    If strPath <> "" Then
        If Not Right(strPath, 1) = "\" Then
            strPath = strPath & "\"
        End If
        Call FuncCreateSQL(strPath)
    End If
End Sub

Private Sub mnuFileImport_Click()
    Dim objfrm As Form
    Set objfrm = New frmIllImport
    Call objfrm.ShowMe(Me)
    Call FillTree
End Sub

Private Sub mnuFilePrintView_Click()
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub mnuFilePrintset_Click()
    zlPrintSet
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub mnuHelpWebHomePage_Click()
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hwnd)
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    '默认参数：类别=类别编码，分类=分类id，项目=目录id
    Dim str类别编码 As String
    Dim lng分类id As Long
    Dim lng项目id As Long
    
    If cmbType.ListIndex <> -1 Then
        str类别编码 = Mid(cmbType.List(cmbType.ListIndex), 1, 1)
    End If
    
    If Not tvwMain_S.SelectedItem Is Nothing Then
        lng分类id = Mid(tvwMain_S.SelectedItem.Key, 2)
    End If
    
    If Not lvwMain.SelectedItem Is Nothing Then
        lng项目id = Mid(lvwMain.SelectedItem.Key, 2)
    End If
    
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
        "类别=" & str类别编码, _
        "分类=" & IIF(lng分类id = 0, "", lng分类id), _
        "项目=" & IIF(lng项目id = 0, "", lng项目id))
End Sub

Private Sub mnuShortIcon_Click(Index As Integer)
    mnuViewIcon_Click Index
End Sub

Private Sub mnuShortMenu1_Click(Index As Integer)
    Select Case Index
        Case 1
            Call mnuEditClass_Click
        Case 2
            Call mnuEditModify_Click
        Case 3
            Call mnuEditDelete_Click
    End Select
End Sub

Private Sub mnuShortMenu2_Click(Index As Integer)
    Select Case Index
        Case 1
            Call mnuEditDisease_Click
        Case 2
            Call mnuEditModify_Click
        Case 3
            Call mnuEditDelete_Click
    End Select
End Sub

Private Sub mnuViewAll_Click()
    mnuViewAll.Checked = Not mnuViewAll.Checked
    Call FillList
End Sub

Private Sub mnuViewColumn_Click()
    If zlControl.LvwSelectColumns(lvwMain, mstr疾病) = True Then
        '列有变化就要重新刷新
        Call FillList
    End If
End Sub

Private Sub mnuViewFind_Click()
    frmIllFind.ShowFind mstr编码类别, mnuViewStop.Checked
End Sub

Private Sub mnuViewOrder_Click()
'检查序号是否按上下级序号有序的排列的，主要是防止跳号或顺序有问题
    Dim nodTemp As Node
    
    mlng序号 = 1
    
    If tvwMain_S.SelectedItem Is Nothing Then Exit Sub
    
    MousePointer = vbHourglass
    Set nodTemp = CheckOrder(tvwMain_S.SelectedItem.Root)
    If nodTemp Is Nothing Then
        MsgBox "检查完毕，序号正确排列。", vbInformation, gstrSysName
    Else
        nodTemp.Selected = True
        nodTemp.EnsureVisible
        Call FillList
        MsgBox "该分类的正确序号应该是" & mlng序号 & "，请修改。", vbExclamation, gstrSysName
    End If
    MousePointer = 0
    
End Sub

Private Function CheckOrder(ByVal nod As Node) As Node
    Dim lngTemp As Long
    Dim nodTemp As Node
    
    '检查节点本身
    lngTemp = Mid(nod.Text, 2, InStr(nod.Text, "】") - 2)
    If lngTemp <> mlng序号 Then
        Set CheckOrder = nod
        Exit Function
    End If
    
    mlng序号 = mlng序号 + 1
    '递归检查其子节点
    Set nod = nod.Child
    Do Until nod Is Nothing
        Set nodTemp = CheckOrder(nod)
        
        '如果有返回值，那表示已经出错了
        If Not nodTemp Is Nothing Then
            Set CheckOrder = nodTemp
            Exit Function
        End If
        Set nod = nod.Next
    Loop
    
End Function

Private Sub mnuViewRefresh_Click()
    Call FillTree
End Sub

Private Sub mnuViewIcon_Click(Index As Integer)
    Dim i As Integer
    
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
    Next
    mnuViewIcon(Index).Checked = True
    lvwMain.View = Index
End Sub

Private Sub mnuViewStop_Click()
    mnuViewStop.Checked = Not mnuViewStop.Checked
    Call mnuViewRefresh_Click
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    CoolBar1.Visible = mnuViewToolButton.Checked
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim buttTemp As Button
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For Each buttTemp In Toolbar1.Buttons
        If mnuViewToolText.Checked Then
            buttTemp.Caption = buttTemp.Tag
        Else
            buttTemp.Caption = ""
        End If
    Next
    CoolBar1.Bands("only").MinHeight = Toolbar1.Height
    Form_Resize
End Sub

Private Sub picsplit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        msngStartX = X
        picSplit.Tag = "可移动"
    End If
End Sub

Private Sub picsplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sngTemp As Single
    
    If picSplit.Tag = "可移动" Then
        sngTemp = picSplit.Left + X - msngStartX
        
        If sngTemp > 1500 And ScaleWidth - sngTemp > 1500 Then
            tvwMain_S.Width = sngTemp - ScaleLeft
            picSplit.Left = sngTemp
            lvwMain.Left = sngTemp + picSplit.Width
            lvwMain.Width = ScaleWidth - lvwMain.Left
        End If
    End If
End Sub

Private Sub picSplit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picSplit.Tag = "" '改变标志
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Preview"
            mnuFilePrintView_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Class"
            mnuEditClass_Click
        Case "Disease"
            mnuEditDisease_Click
        Case "Modify"
            mnuEditModify_Click
        Case "Delete"
            mnuEditDelete_Click
        Case "停用"
            mnuEditStop_Click
        Case "启用"
            mnuEditStart_Click
        Case "View"
            mnuViewIcon(lvwMain.View).Checked = False
            If lvwMain.View = 3 Then
                mnuViewIcon(0).Checked = True
                lvwMain.View = 0
            Else
                mnuViewIcon(lvwMain.View + 1).Checked = True
                lvwMain.View = lvwMain.View + 1
            End If
        Case "Find"
            Call mnuViewFind_Click
        Case "Help"
            Call mnuHelpHelp_Click
        Case "Quit"
            Call mnuFileExit_Click
    End Select
End Sub

Private Sub Toolbar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool
    End If
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Form_Resize
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim i As Integer
    
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
    Next
    mnuViewIcon(ButtonMenu.Index - 1).Checked = True
    lvwMain.View = ButtonMenu.Index - 1
End Sub

Private Sub tvwMain_S_GotFocus()
    Call SetMenu
End Sub

Private Sub tvwMain_S_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If mnuEdit.Visible = False Then Exit Sub
        mnuShortMenu1(1).Enabled = mnuEditClass.Enabled
        mnuShortMenu1(2).Enabled = mnuEditModify.Enabled
        mnuShortMenu1(3).Enabled = mnuEditDelete.Enabled
        PopupMenu mnuShort1
    End If
End Sub

Private Sub tvwMain_S_NodeClick(ByVal Node As MSComctlLib.Node)
    If mstrNodeKey = Node.Key Then Exit Sub
    
    FillList
End Sub

Private Sub subPrint(bytMode As Byte)
'功能:进行打印,预览和输出到EXCEL
'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As Object
            
    Set objPrint = New zlPrintLvw
    objPrint.Title.Text = "编码表"
    Set objPrint.Body.objData = lvwMain
    objPrint.BelowAppItems.Add "打印人：" & gstrUserName
    objPrint.BelowAppItems.Add "打印时间：" & Format(zldatabase.Currentdate, "yyyy年MM月dd日")
    If bytMode = 1 Then
        Select Case zlPrintAsk(objPrint)
        Case 1
            zlPrintOrViewLvw objPrint, 1
        Case 2
            zlPrintOrViewLvw objPrint, 2
        Case 3
            zlPrintOrViewLvw objPrint, 3
        Case Else
        End Select
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If

End Sub

Private Sub Fill类别()
'功能：装入疾病编码类别
    Dim rsTemp As New ADODB.Recordset
    
    mstrTypeText = ""
    
    On Error GoTo errHandle
    gstrSQL = ""
    If gbln购买中医 = False Or gbln使用中医 = False Then
        gstrSQL = " where 编码<>'B' and 编码<>'Z'"
    End If
    gstrSQL = "select 编码,类别,是否分类 from 疾病编码类别 " & gstrSQL & " order by 优先级"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption)

    cmbType.Clear
    Do Until rsTemp.EOF
        cmbType.AddItem rsTemp("编码") & ". " & rsTemp("类别")
        cmbType.ItemData(cmbType.NewIndex) = rsTemp("是否分类")
        rsTemp.MoveNext
    Loop
    
    cmbType.ListIndex = 0
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub FillTree()
'如果当前编码有类别，那么把分类装入到树中
    Dim rsTemp As New ADODB.Recordset
    Dim nodTemp As Node
    Dim strKey As String
    Dim strTemp As String
    
    If Not tvwMain_S.SelectedItem Is Nothing Then
        strKey = tvwMain_S.SelectedItem.Key
    End If
    
    mstrTypeText = cmbType.Text
    mstr编码类别 = Left(mstrTypeText, 1)
    
    rsTemp.CursorLocation = adUseClient

    On Error GoTo errHandle
    tvwMain_S.Nodes.Clear
    If tvwMain_S.Visible = True Then
        '只处理有类别的编码
        gstrSQL = "select ID,上级ID,序号,名称, 撤档时间 from 疾病编码分类 where 类别=[1] " & vbNewLine & _
            IIF(mnuViewStop.Checked, "", " And (撤档时间 is null or 撤档时间>=to_date('3000-01-01','yyyy-mm-dd'))") & _
            " Start With 上级ID is null connect by prior id=上级ID order by level,序号"

        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstr编码类别)
                
        Do Until rsTemp.EOF
            strTemp = IIF(Format(rsTemp!撤档时间 & "", "YYYY-MM-DD") = "3000-01-01", "", Nvl(rsTemp!撤档时间))
            
            If IsNull(rsTemp("上级ID")) Then
                Set nodTemp = tvwMain_S.Nodes.Add(, , "K" & rsTemp("ID"), "【" & rsTemp("序号") & "】" & Trim(rsTemp("名称")), "Root", "Root")
            Else
                Set nodTemp = tvwMain_S.Nodes.Add("K" & rsTemp("上级ID"), tvwChild, "K" & rsTemp("ID"), "【" & rsTemp("序号") & "】" & Trim(rsTemp("名称")), "Root", "Root")
            End If
            If strTemp <> "" Then
                nodTemp.ForeColor = vbRed
                nodTemp.Tag = strTemp
            End If
            rsTemp.MoveNext
        Loop
    End If
    On Error Resume Next
    Set nodTemp = tvwMain_S.Nodes(strKey)
    If Err <> 0 Then
        Set nodTemp = tvwMain_S.Nodes(1)
        nodTemp.Selected = True
        nodTemp.Expanded = True
    Else
        Err.Clear
        nodTemp.Selected = True
        nodTemp.Expanded = True
        nodTemp.EnsureVisible
    End If
    Call FillList
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub FillList()
'功能:更新ListView中的数据
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim i As Integer
    Dim lst As ListItem
    Dim strKey As String
    
    On Error GoTo errHandleList
    
    If tvwMain_S.SelectedItem Is Nothing And tvwMain_S.Visible = True Then
        lvwMain.ListItems.Clear
        Call SetMenu
        Exit Sub
    End If
    If Not lvwMain.SelectedItem Is Nothing Then
        '保留原有键值
        strKey = lvwMain.SelectedItem.Key
    End If
    
    rsTemp.CursorLocation = adUseClient
    
    If tvwMain_S.Visible = False Then
        '没有上下级的关系
        gstrSQL = " A.类别=[1] "
    Else
        mstrNodeKey = tvwMain_S.SelectedItem.Key '记录当前访问的节点
        If mnuViewAll.Checked = True Then
            gstrSQL = " A.分类ID in  " & _
                "(select ID from 疾病编码分类 start with id=[1] " & _
                " connect by prior id=上级ID)"
        Else
            gstrSQL = " A.分类ID=[1] "
        End If
    End If
        
    gstrSQL = "" & _
    "   Select A.ID,A.编码,附码,A.名称,A.简码 as 拼音码,A.五笔码,A.说明,A.性别限制,A.疗效限制 as 提醒疗效,A.手术类型," & _
    "          A.统计码,decode(A.分娩,1,'录入') 分娩信息,to_char(A.建档时间,'yyyy-mm-dd') as  建档时间, " & _
    "          to_char(A.撤档时间,'yyyy-mm-dd') as 撤档时间" & _
    "   From 疾病编码目录 A  " & _
    "   Where " & gstrSQL & IIF(mnuViewStop.Checked, "", " and (a.撤档时间 is null or a.撤档时间>=to_date('3000-01-01','yyyy-mm-dd'))")
    
    If tvwMain_S.Visible = False Then
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstr编码类别)
    Else
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Mid(tvwMain_S.SelectedItem.Key, 2)))
    End If
    Dim strIco As String
    
    With lvwMain.ListItems
        .Clear
        Do Until rsTemp.EOF
            '得出正确的图标
            '添加节点
            strTemp = IIF(Nvl(rsTemp!撤档时间) = "3000-01-01", "", Nvl(rsTemp!撤档时间))
            strIco = IIF(strTemp <> "", "Stop", "Item")
            
       
            Set lst = .Add(, "K" & rsTemp("id"), rsTemp("编码"), strIco, strIco)
            If strTemp <> "" Then
                lst.ForeColor = vbRed
            Else
                lst.ForeColor = lvwMain.ForeColor
            End If
            Dim varValue As Variant
            '根据ListView的列名从数据库取数
            For i = 2 To lvwMain.ColumnHeaders.Count
                varValue = rsTemp(lvwMain.ColumnHeaders(i).Text).value
                If lvwMain.ColumnHeaders(i).Text = "撤档时间" Then
                    If Nvl(varValue) = "3000-01-01" Then
                        lst.SubItems(i - 1) = ""
                    Else
                        lst.SubItems(i - 1) = IIF(IsNull(varValue), "", varValue)
                        
                    End If
                Else
                    lst.SubItems(i - 1) = IIF(IsNull(varValue), "", varValue)
                End If
                If strTemp <> "" Then lst.ListSubItems(i - 1).ForeColor = vbRed
            Next
            rsTemp.MoveNext
        Loop
    End With
    
    If lvwMain.ListItems.Count > 0 Then
        On Error Resume Next
        Set lst = lvwMain.ListItems(strKey)
        If Err <> 0 Then
            Set lst = lvwMain.ListItems(1)
            lst.Selected = True
            lst.EnsureVisible
        Else
            Err.Clear
            lst.Selected = True
            lst.EnsureVisible
        End If
    End If
    Call SetMenu
    
    Exit Sub
errHandleList:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Sub

Private Sub 权限控制()
'功能:由于有的用户权限不够,故使一些菜单项或按钮不可见
    gbln购买中医 = InStr(mstrPrivs, "中医") > 0
    If InStr(mstrPrivs, "增删改") = 0 Then
        mnuEdit.Visible = False
        mnuEditClass.Visible = False
        mnuEditDisease.Visible = False
        mnuEditModify.Visible = False
        mnuFileExport.Visible = False
        mnuFileImport.Visible = False
        mnuShortMenu2(1).Visible = False
        mnuShortMenu2(2).Visible = False
        mnuShortMenu2(3).Visible = False
        mnuShortLine.Visible = False
        
        Toolbar1.Buttons("Split0").Visible = False
        Toolbar1.Buttons("Class").Visible = False
        Toolbar1.Buttons("Disease").Visible = False
        Toolbar1.Buttons("Modify").Visible = False
        Toolbar1.Buttons("Delete").Visible = False
        Toolbar1.Buttons("Splits").Visible = False
        Toolbar1.Buttons("停用").Visible = False
        Toolbar1.Buttons("启用").Visible = False
        
    End If
      
End Sub

Public Sub SetMenu()
'功能:设置修改和删除按钮的有效值
'参数:blnEnabled 有效值
'    Dim blnNew As Boolean
    Dim blnModify As Boolean
    Dim blnStop As Boolean
    
    blnStop = False
    blnModify = True
    If ActiveControl Is tvwMain_S Then
        If tvwMain_S.SelectedItem Is Nothing Then
            blnModify = False
            stbThis.Panels(2).Text = "请先设置分类。"
        Else
            stbThis.Panels(2).Text = "当前分类共有" & tvwMain_S.SelectedItem.Children & "个子类，" & lvwMain.ListItems.Count & "条疾病编码。"
        End If
        If Not tvwMain_S.SelectedItem Is Nothing Then
            blnStop = tvwMain_S.SelectedItem.ForeColor = vbRed
        End If
    Else
        If lvwMain.SelectedItem Is Nothing Or lvwMain.ListItems.Count = 0 Then
            blnModify = False
        End If
        stbThis.Panels(2).Text = "当前分类共有" & lvwMain.ListItems.Count & "条疾病编码。"
        If Not lvwMain.SelectedItem Is Nothing Then
            blnStop = lvwMain.SelectedItem.SubItems(lvwMain.ColumnHeaders("_撤档时间").Index - 1) <> ""
        End If
    End If
    
    Toolbar1.Buttons("停用").Enabled = Not blnStop
    Toolbar1.Buttons("启用").Enabled = blnStop
    
    mnuEditStart.Enabled = blnStop
    mnuEditStop.Enabled = Not blnStop
    
    '只有树形列表可见时，才可增加分类
    Toolbar1.Buttons("Class").Enabled = tvwMain_S.Visible
    mnuEditClass.Enabled = tvwMain_S.Visible
    mnuViewAll.Enabled = tvwMain_S.Visible
    mnuViewOrder.Enabled = tvwMain_S.Visible
    
    mnuEditDisease.Enabled = (Not tvwMain_S.Visible) Or (Not tvwMain_S.SelectedItem Is Nothing)
    Toolbar1.Buttons("Disease").Enabled = mnuEditDisease.Enabled
    
    Toolbar1.Buttons("Modify").Enabled = blnModify
    Toolbar1.Buttons("Delete").Enabled = blnModify
    mnuEditDelete.Enabled = blnModify
    mnuEditModify.Enabled = blnModify

    EnablePrint lvwMain.ListItems.Count > 0
End Sub

Private Sub EnablePrint(ByVal blnEnabled As Boolean)
'功能:设置打印和预鉴按钮的有效值
'参数:blnEnabled 有效值
    Toolbar1.Buttons("Print").Enabled = blnEnabled
    Toolbar1.Buttons("Preview").Enabled = blnEnabled
    mnuFilePrintView.Enabled = blnEnabled
    mnuFilePrint.Enabled = blnEnabled
    mnuFileExcel.Enabled = blnEnabled
End Sub
Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

Private Sub StopAndResume(ByVal blnStop As Boolean)
    '--------------------------------------------------------------------------------------
    '功能:停用或启用物资
    '参数:blnStop-是否停用
    '返回:
    '编制:刘兴宏
    '问题:11689
    '修改:2007/12/28
    '--------------------------------------------------------------------------------------
    
    Dim lng疾病ID As Long, lng分类id As Long
    Dim strSQL As String, intIndex As Integer
    Dim i As Integer
    Dim ReMoveRow As Long
    Dim nodTemp As Node
    Dim str名称 As String
    Dim strDate As String
    
    
    If ActiveControl Is tvwMain_S And tvwMain_S.Visible = True Then
        '修改分类
        Set nodTemp = tvwMain_S.SelectedItem
        If nodTemp Is Nothing Then
            Exit Sub
        Else
            With tvwMain_S.SelectedItem
                If AllowContinue = False Then Exit Sub
                If MsgBox("你是否真的要" & IIF(blnStop, "停用", "启用") & "“" & tvwMain_S.SelectedItem.Text & "”本分类下所有分类项目以及分类项目下所有的疾病编码吗？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
                lng分类id = Val(Mid(tvwMain_S.SelectedItem.Key, 2))
                If lng分类id <= 0 Then Exit Sub
                Err = 0: On Error GoTo ErrHand:
                If blnStop Then
                    strSQL = "Zl_疾病编码分类_STOP(" & lng分类id & ")"
                Else
                    If nodTemp.Tag <> "" Then
                        strDate = "To_Date('" & nodTemp.Tag & "','YYYY-MM-DD HH24:MI:SS')"
                        strSQL = "Zl_疾病编码分类_REUSE(" & lng分类id & "," & strDate & ")"
                    Else
                        strSQL = "Zl_疾病编码分类_REUSE(" & lng分类id & ")"
                    End If
                    
                End If
                zldatabase.ExecuteProcedure strSQL, Me.Caption
            End With
            '刷新
            Call FillTree
        End If
    Else
        '修改疾病
        If lvwMain.SelectedItem Is Nothing Then Exit Sub
        If MsgBox("你是否真的要" & IIF(blnStop, "停用", "启用") & "“" & lvwMain.SelectedItem.Text & "”的疾病吗？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
        
        lng疾病ID = Val(Mid(lvwMain.SelectedItem.Key, 2))
        If lng疾病ID <= 0 Then Exit Sub
        
        Err = 0: On Error GoTo ErrHand:
        If blnStop Then
            strSQL = "Zl_疾病编码目录_STOP(" & lng疾病ID & ")"
        Else
            strSQL = "Zl_疾病编码目录_REUSE(" & lng疾病ID & ")"
        End If
        zldatabase.ExecuteProcedure strSQL, Me.Caption
        
        With lvwMain.SelectedItem
            .Icon = IIF(blnStop, "Stop", "Item")
            .SmallIcon = IIF(blnStop, "Stop", "Item")
            .ForeColor = IIF(blnStop, vbRed, &H80000008)
        End With
        For i = 2 To lvwMain.ColumnHeaders.Count
             lvwMain.SelectedItem.ListSubItems(i - 1).ForeColor = IIF(blnStop, vbRed, &H80000008)
        Next
        If mnuViewStop.Checked Then
            If Not blnStop Then
                lvwMain.SelectedItem.SubItems(lvwMain.ColumnHeaders("_撤档时间").Index - 1) = ""
            Else
                lvwMain.SelectedItem.SubItems(lvwMain.ColumnHeaders("_撤档时间").Index - 1) = Format(zldatabase.Currentdate, "yyyy-mm-dd")
            End If
            Call SetMenu
            Exit Sub
        End If
        
        If blnStop = False Then
            lvwMain.SelectedItem.SubItems(lvwMain.ColumnHeaders("_撤档时间").Index - 1) = ""
            Call SetMenu
            Exit Sub
        End If
        Me.MousePointer = 0
        With lvwMain
            intIndex = .SelectedItem.Index
            .ListItems.Remove .SelectedItem.Key
            If .ListItems.Count > 0 Then
                intIndex = IIF(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
                .ListItems(intIndex).Selected = True
                .ListItems(intIndex).EnsureVisible
            End If
            Call SetMenu
        End With
    End If
    

    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncCreateSQL(ByVal strFilePath As String)
'功能:生成安装脚步
'参数:strFilePath-保存文件的位置
    Dim rsType As New ADODB.Recordset
    Dim rsContent As New ADODB.Recordset
    Dim strTitle As String, strTemp As String, strValue As String
    Dim strName As String, strType As String
    
    Dim lngId As Long, i As Long
    Dim colTemp As Collection
    Dim colSQL As Collection
    Dim blnOver As Boolean
    Dim objFile As New FileSystemObject
    Dim strFileName As String
    
    On Error GoTo errH
    '只导出疾病编码类别为 D-ICD-10;Y-损伤中毒;M-肿瘤形态学;S-手术编码
    gstrSQL = "Select a.Id, a.上级id, a.序号, a.名称,a.简码, a.类别, a.编码范围, a.是否病人,NULL AS ID_TEMP, NULL AS 上级ID_TEMP " & vbNewLine & _
                "From 疾病编码分类 A, 疾病编码类别 B" & vbNewLine & _
                "Where a.类别 = b.编码 And (a.撤档时间 Is Null Or Trunc(a.撤档时间) = To_Date('3000-01-01', 'YYYY-MM-DD')) And A.类别 In ('D','Y','M','S')" & vbNewLine & _
                "Order By b.优先级,a.序号"
    Call zldatabase.OpenRecordset(rsType, gstrSQL, Me.Caption, adOpenStatic, adLockOptimistic)
    
    gstrSQL = "Select a.Id, NULL AS ID_TEMP, a.编码, a.序号, a.附码, a.统计码, a.名称, a.简码, a.五笔码, a.说明, a.性别限制, a.疗效限制, a.手术类型, a.分娩, a.分类id, NULL As 分类ID_TEMP, a.适用范围, a.类别" & vbNewLine & _
        "From 疾病编码目录 A, 疾病编码类别 B" & vbNewLine & _
        "Where a.类别 = b.编码 And (a.撤档时间 Is Null Or Trunc(a.撤档时间) = To_Date('3000-01-01', 'YYYY-MM-DD')) And A.类别 In ('D','Y','M','S') " & vbNewLine & _
        "Order By b.优先级, a.编码, a.序号 "
    Call zldatabase.OpenRecordset(rsContent, gstrSQL, Me.Caption, adOpenStatic, adLockOptimistic)
    
    'ID更换
    lngId = 1: Set colTemp = New Collection
    For i = 1 To rsType.RecordCount
        colTemp.Add lngId, "_" & rsType!ID
        lngId = lngId + 1
        rsType.MoveNext
    Next
    
    rsType.Filter = ""
    For i = 1 To rsType.RecordCount
        rsType!ID_TEMP = colTemp("_" & rsType!ID)
        If Nvl(rsType!上级id, 0) <> 0 Then rsType!上级ID_TEMP = colTemp("_" & rsType!上级id)
        rsType.MoveNext
    Next
    
    rsContent.Filter = ""
    lngId = 1
    For i = 1 To rsContent.RecordCount
        rsContent!ID_TEMP = lngId
        rsContent!分类ID_TEMP = colTemp("_" & rsContent!分类id)
        lngId = lngId + 1
        rsContent.MoveNext
    Next
    
    Set colSQL = New Collection
    Set colTemp = New Collection
    
    With rsType
        rsType.Filter = ""
        strType = ""
        strTitle = "Insert Into 疾病编码分类(ID, 上级id, 类别, 序号, 名称, 简码, 编码范围, 是否病人) " & vbCrLf
        For i = 1 To .RecordCount
            strName = FuncGetStr(!名称)
            strTemp = "Select " & !ID_TEMP & "," & IIF(Val(!上级id & "") = 0, "Null", !上级ID_TEMP) & ",'" & Trim(!类别) & "'," & !序号 & ",'" & strName & "','" & !简码 & "','" & FuncGetStr(!编码范围 & "") & "'," & Val(!是否病人 & "") & " From Dual UNION ALL" & vbCrLf
            If Len(strTitle & strValue & strTemp) > 100000 Or (!类别 & "" <> strType And strType <> "") Then
                colSQL.Add "--类别=" & strType  '添加一行空行
                strValue = Mid(strValue, 1, InStrRev(strValue, "UNION ALL") - 1) & ";"
                colSQL.Add strTitle & strValue
                strValue = strTemp
                blnOver = True
            Else
                blnOver = False
                strValue = strValue & strTemp
            End If
            strType = !类别 & ""
            .MoveNext
            If .EOF Then
                If Not blnOver Then
                    colSQL.Add "--类别=" & strType  '添加一行空行
                    strValue = Mid(strValue, 1, InStrRev(strValue, "UNION ALL") - 1) & ";"
                    colSQL.Add strTitle & strValue
                    Exit For
                End If
            End If
        Next
    End With
    strFileName = strFilePath & "疾病编码.SQL"
    If objFile.FileExists(strFileName) Then objFile.DeleteFile strFileName, True
    SaveLog strFileName, "--疾病编码分类", "-1"
    For i = 1 To colSQL.Count
        SaveLog strFileName, colSQL(i), "-1"
    Next
    
    Set colSQL = New Collection
    With rsContent
        .Filter = "": strValue = "": strType = ""
        strTitle = "Insert Into 疾病编码目录 (ID, 分类id, 类别, 编码, 序号, 附码, 名称, 简码, 五笔码, 说明, 性别限制, 疗效限制, 手术类型, 分娩, 适用范围)" & vbCrLf
        For i = 1 To .RecordCount
            strName = FuncGetStr(!名称)
            strTemp = "Select " & !ID_TEMP & "," & !分类ID_TEMP & ",'" & !类别 & "','" & !编码 & "'," & !序号 & ",'" & !附码 & "','" & strName & "','" & !简码 & "','" & !五笔码 & "','" & !说明 & "','" & !性别限制 & "','" & _
                    !疗效限制 & "','" & !手术类型 & "','" & !分娩 & "','" & !适用范围 & "' From Dual UNION ALL" & vbCrLf
      
            If Len(strTitle & strValue & strTemp) > 100000 Or (!类别 & "" <> strType And strType <> "") Then
                colSQL.Add "--类别=" & strType  '添加一行空行
                strValue = Mid(strValue, 1, InStrRev(strValue, "UNION ALL") - 1) & ";"
                colSQL.Add strTitle & strValue
                strValue = strTemp
                blnOver = True
            Else
                blnOver = False
                strValue = strValue & strTemp
            End If
            strType = !类别 & ""
            .MoveNext
            If .EOF Then
                If Not blnOver Then
                    colSQL.Add "--类别=" & strType  '添加一行空行
                    strValue = Mid(strValue, 1, InStrRev(strValue, "UNION ALL") - 1) & ";"
                    colSQL.Add strTitle & strValue
                    Exit For
                End If
            End If
        Next
    End With
    SaveLog strFileName, "--疾病编码目录", "-1"
    For i = 1 To colSQL.Count
        SaveLog strFileName, colSQL(i), "-1"
    Next
    MsgBox "导出成功,文件位置:" & vbCrLf & strFileName, vbInformation + vbOKOnly, Me.Caption
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



