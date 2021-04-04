VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frm保险病种 
   Caption         =   "医保病种管理"
   ClientHeight    =   6150
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8025
   Icon            =   "frm保险病种.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6150
   ScaleWidth      =   8025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   5790
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   635
      SimpleText      =   $"frm保险病种.frx":0E42
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm保险病种.frx":0E89
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9075
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
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   1270
      BandCount       =   1
      _CBWidth        =   8025
      _CBHeight       =   720
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinHeight1      =   660
      Width1          =   615
      FixedBackground1=   0   'False
      Key1            =   "Comm"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   660
         Left            =   30
         TabIndex        =   5
         Top             =   30
         Width           =   7905
         _ExtentX        =   13944
         _ExtentY        =   1164
         ButtonWidth     =   820
         ButtonHeight    =   1164
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   13
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "Preview"
               Object.ToolTipText     =   "打印预览"
               Object.Tag             =   "预览"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "Print"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageIndex      =   2
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
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修改"
               Key             =   "Modify"
               Object.ToolTipText     =   "修改保险类别"
               Object.Tag             =   "修改"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "Delete"
               Object.ToolTipText     =   "删除保险类别"
               Object.Tag             =   "删除"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split1"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查看"
               Key             =   "View"
               Object.ToolTipText     =   "查看方式"
               Object.Tag             =   "查看"
               ImageIndex      =   6
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
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "split2"
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查找"
               Key             =   "Find"
               Object.ToolTipText     =   "查找"
               Object.Tag             =   "查找"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "split3"
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Object.ToolTipText     =   "帮助主题"
               Object.Tag             =   "帮助"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Exit"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ListView lvwLimit 
      Height          =   960
      Left            =   2160
      TabIndex        =   8
      Top             =   4710
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   1693
      View            =   3
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "编码"
         Text            =   "编码"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "名称"
         Text            =   "名称"
         Object.Width           =   4586
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "规格"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "大类"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "性质"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.PictureBox picSplitH 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   1590
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   6075
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4530
      Width           =   6075
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   720
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险病种.frx":171D
            Key             =   "Fix"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险病种.frx":1A37
            Key             =   "FixD"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险病种.frx":1D51
            Key             =   "Common"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险病种.frx":206B
            Key             =   "CommonD"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险病种.frx":2385
            Key             =   "Disease"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   120
      Top             =   5160
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
            Picture         =   "frm保险病种.frx":2C5F
            Key             =   "Fix"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险病种.frx":2F79
            Key             =   "FixD"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险病种.frx":3293
            Key             =   "Common"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险病种.frx":35AD
            Key             =   "CommonD"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险病种.frx":38C7
            Key             =   "Disease"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险病种.frx":3E61
            Key             =   "Limit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   4485
      Top             =   390
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
            Picture         =   "frm保险病种.frx":42B3
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险病种.frx":44CD
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险病种.frx":46E7
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险病种.frx":4901
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险病种.frx":4B1B
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险病种.frx":4D35
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险病种.frx":4F4F
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险病种.frx":5169
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险病种.frx":5383
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   5205
      Top             =   360
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
            Picture         =   "frm保险病种.frx":559D
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险病种.frx":57B7
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险病种.frx":59D1
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险病种.frx":5BEB
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险病种.frx":5E05
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险病种.frx":601F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险病种.frx":6239
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险病种.frx":6453
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险病种.frx":666D
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwKind_S 
      Height          =   4755
      Left            =   60
      TabIndex        =   0
      Top             =   810
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   8387
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ils32"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.PictureBox picSplitV 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5490
      Left            =   1590
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5490
      ScaleWidth      =   45
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   690
      Width           =   45
   End
   Begin MSComctlLib.ListView lvwItem 
      Height          =   3450
      Left            =   1950
      TabIndex        =   1
      Top             =   840
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   6085
      View            =   3
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "编码"
         Text            =   "编码"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "名称"
         Text            =   "名称"
         Object.Width           =   4586
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "简码"
         Text            =   "简码"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "疾病类别"
         Text            =   "疾病类别"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "特殊封顶线"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "封顶线金额"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label lbl限制 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "病种特准使用项目"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1245
      Left            =   1650
      TabIndex        =   6
      Top             =   4470
      Width           =   285
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
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditAdd 
         Caption         =   "增加(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "修改(&M)"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "删除(&D)"
         Shortcut        =   {DEL}
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
      Begin VB.Menu mnuViewSplit 
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
      Begin VB.Menu mnuViewSplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "查找(&F)"
      End
      Begin VB.Menu mnuViewSplit3 
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
   Begin VB.Menu mnuShort 
      Caption         =   "快捷菜单"
      Visible         =   0   'False
      Begin VB.Menu mnuShortMenu 
         Caption         =   "增加(&A)"
         Index           =   0
      End
      Begin VB.Menu mnuShortMenu 
         Caption         =   "修改(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuShortMenu 
         Caption         =   "删除(&D)"
         Index           =   2
      End
      Begin VB.Menu mnuShortSplit 
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
Attribute VB_Name = "frm保险病种"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim msngStartX As Single, msngStartY As Single    '移动前鼠标的位置
Dim mintColumn As Integer
Dim mstrKey As String
Dim mblnLoad As Boolean


Private Sub Form_Activate()
    If mblnLoad = True Then
        '显示当前项
        lvwKind_S.SelectedItem.EnsureVisible
        lvwKind_S_ItemClick lvwKind_S.SelectedItem
    End If
    mblnLoad = False
End Sub

Private Sub Form_Load()
    RestoreWinState Me, App.ProductName
    Call 权限控制
    
    mblnLoad = True
End Sub

Private Sub Form_Resize()
    Dim sngTop As Single, sngBottom As Single
    On Error Resume Next
    
    sngTop = IIf(cbrThis.Visible, cbrThis.Top + cbrThis.Height, 0)
    sngBottom = ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0)
    
    lvwKind_S.Top = sngTop
    lvwKind_S.Height = IIf(sngBottom - lvwKind_S.Top > 0, sngBottom - lvwKind_S.Top, 0)
    lvwKind_S.Left = ScaleLeft
    
    picSplitV.Top = sngTop
    picSplitV.Height = IIf(sngBottom - picSplitV.Top > 0, sngBottom - picSplitV.Top, 0)
    picSplitV.Left = lvwKind_S.Left + lvwKind_S.Width
    
    lbl限制.Left = picSplitV.Left + picSplitV.Width
    With lvwLimit
        .Left = lbl限制.Left + lbl限制.Width
        .Width = IIf(ScaleWidth - .Left > 0, ScaleWidth - .Left, 0)
        .Top = sngBottom - .Height
        
    End With
    lbl限制.Top = lvwLimit.Top
    lbl限制.Height = lvwLimit.Height
    
    picSplitH.Left = lbl限制.Left
    picSplitH.Width = IIf(ScaleWidth - picSplitH.Left > 0, ScaleWidth - picSplitH.Left, 0)
    picSplitH.Top = lvwLimit.Top - picSplitH.Height
    
    lvwItem.Top = sngTop
    lvwItem.Left = picSplitH.Left
    lvwItem.Width = picSplitH.Width
    lvwItem.Height = IIf(picSplitH.Top - lvwItem.Top > 0, picSplitH.Top - lvwItem.Top, 0)
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub lvwItem_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintColumn = ColumnHeader.Index - 1 Then '仍是刚才那列
        lvwItem.SortOrder = IIf(lvwItem.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvwItem.SortKey = mintColumn
        lvwItem.SortOrder = lvwAscending
    End If
    If Not lvwItem.SelectedItem Is Nothing Then
        lvwItem.SelectedItem.EnsureVisible
    End If
End Sub

Private Sub lvwItem_DblClick()
    If mnuEditModify.Visible = True And mnuEditModify.Enabled = True Then
        Call mnuEditModify_Click
    End If
End Sub

Private Sub lvwItem_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call FillItem
End Sub

Private Sub lvwItem_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    
    If Button = 2 Then
        mnuShortMenu(0).Enabled = mnuEditAdd.Enabled
        mnuShortMenu(1).Enabled = mnuEditModify.Enabled
        mnuShortMenu(2).Enabled = mnuEditDelete.Enabled
        For i = 0 To 3
            mnuShortIcon(i).Checked = mnuViewIcon(i).Checked
        Next
        PopupMenu mnuShort, vbPopupMenuRightButton
    End If
End Sub

Private Sub lvwKind_S_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If mstrKey = Item.Key Then Exit Sub
    Call FillList
End Sub


Private Sub lvwLimit_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call zlControl.LvwSortColumn(lvwLimit, ColumnHeader.Index)
End Sub

Private Sub mnuEditAdd_Click()
    Dim lng险类 As Long
    
    lng险类 = Mid(mstrKey, 2)
    If frm保险病种编辑.编辑病种(lng险类, "") = True Then
        '主记录的内容已经更新了
        Call SetMenu
    End If
End Sub

Private Sub mnuEditModify_Click()
    Dim lng险类 As Long
    
    If lvwItem.SelectedItem Is Nothing Then Exit Sub
    
    lng险类 = Mid(mstrKey, 2)
    If frm保险病种编辑.编辑病种(lng险类, Mid(lvwItem.SelectedItem.Key, 2)) = True Then
        '主记录的内容已经更新了
        Call SetMenu
    End If
End Sub

Private Sub mnuEditDelete_Click()
    Dim intIndex As Integer
    
    On Error GoTo errHandle
    
    If lvwItem.SelectedItem Is Nothing Then Exit Sub
    If MsgBox("你确认要删除编码为“" & lvwItem.SelectedItem.Text & "”的医保病种吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    MousePointer = vbHourglass
    
    gstrSQL = "zl_保险病种_DELETE(" & Mid(lvwItem.SelectedItem.Key, 2) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    With lvwItem
        intIndex = .SelectedItem.Index
        .ListItems.Remove .SelectedItem.Key
        If .ListItems.Count > 0 Then
            intIndex = IIf(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
            .ListItems(intIndex).Selected = True
            .ListItems(intIndex).EnsureVisible
        End If
    End With
    Call SetMenu
    MousePointer = vbDefault
    Exit Sub
    
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    MousePointer = vbDefault
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

Private Sub subPrint(bytMode As Byte)
'功能:进行打印,预览和输出到EXCEL
'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As New zlPrintLvw
    
    
    objPrint.Title.Text = "医保病种"
    Set objPrint.Body.objData = lvwItem
    objPrint.UnderAppItems.Add "医保类别：" & lvwKind_S.SelectedItem.Text
    objPrint.BelowAppItems.Add "打印人：" & gstrUserName
    objPrint.BelowAppItems.Add "打印时间：" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrViewLvw objPrint, 1
          Case 2
              zlPrintOrViewLvw objPrint, 2
          Case 3
              zlPrintOrViewLvw objPrint, 3
      End Select
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
   ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hwnd)
End Sub

Private Sub mnuShortIcon_Click(Index As Integer)
    mnuViewIcon_Click Index
End Sub

Private Sub mnuShortMenu_Click(Index As Integer)
    Select Case Index
        Case 0
            mnuEditAdd_Click
        Case 1
            mnuEditModify_Click
        Case 2
            mnuEditDelete_Click
    End Select
End Sub

Private Sub mnuViewFind_Click()
    frm病种查找.lng险类 = Mid(Me.lvwKind_S.SelectedItem.Key, 2)
    frm病种查找.Show , Me
End Sub

Private Sub mnuViewIcon_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
    Next
    mnuViewIcon(Index).Checked = True
    lvwItem.View = Index
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
    Dim lngCount As Integer
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For lngCount = 1 To tbrThis.Buttons.Count
        tbrThis.Buttons(lngCount).Caption = IIf(mnuViewToolText.Checked = True, tbrThis.Buttons(lngCount).Tag, "")
    Next
    cbrThis.Bands(1).MinHeight = tbrThis.Height
    cbrThis.Refresh
    Call Form_Resize
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
        If sngTemp - lvwItem.Top > 1200 And ScaleHeight - (sngTemp + picSplitH.Height) > 1500 Then
            picSplitH.Top = sngTemp
            lvwLimit.Height = lvwLimit.Top + lvwLimit.Height - (picSplitH.Top + picSplitH.Height)
            lvwLimit.Top = picSplitH.Top + picSplitH.Height
            
            Call Form_Resize
        End If
        lvwKind_S.SetFocus
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
        If sngTemp > 1200 And ScaleWidth - (sngTemp + picSplitV.Width) > 1000 Then
            picSplitV.Left = sngTemp
            lvwKind_S.Width = picSplitV.Left - lvwKind_S.Left
            
            Call Form_Resize
        End If
        lvwKind_S.SetFocus
    End If
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Preview"
            mnuFilePreview_Click
        Case "Print"
            mnuFilePrint_Click
        Case "New"
            mnuEditAdd_Click
        Case "Delete"
            mnuEditDelete_Click
        Case "Modify"
            mnuEditModify_Click
        Case "View"
            mnuViewIcon(lvwItem.View).Checked = False
            If lvwItem.View = 3 Then
                mnuViewIcon(0).Checked = True
                lvwItem.View = 0
            Else
                mnuViewIcon(lvwItem.View + 1).Checked = True
                lvwItem.View = lvwItem.View + 1
            End If
        Case "Find"
            mnuViewFind_Click
        Case "Help"
            mnuHelpTitle_Click
        Case "Exit"
            mnuFileExit_Click
    End Select
End Sub

Private Sub tbrThis_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim i As Integer
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
    Next
    mnuViewIcon(ButtonMenu.Index - 1).Checked = True
    lvwItem.View = ButtonMenu.Index - 1
End Sub

Private Sub tbrThis_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuViewTool, 2
    
End Sub

Private Sub FillList()
'功能：显示当前类别下的医保大类
    Dim rsTemp As New ADODB.Recordset
    Dim lst As ListItem
    Dim strItemKey As String
    
    Me.MousePointer = vbHourglass
    mstrKey = lvwKind_S.SelectedItem.Key
    If Not lvwItem.SelectedItem Is Nothing Then
        '保存以前的选择项
        strItemKey = lvwItem.SelectedItem.Key
    End If
    lvwItem.ListItems.Clear
    
    'Modified by 朱玉宝 20031218 地区：福州
    If Mid(mstrKey, 2) = TYPE_福建巨龙 Or Mid(mstrKey, 2) = TYPE_福建省 Or Mid(mstrKey, 2) = TYPE_福州市 Or Mid(mstrKey, 2) = TYPE_南平市 Then
        gstrSQL = "select ID,substr(名称,1,instr(名称,'@@')-1) 编码,substr(名称,instr(名称,'@@')+2) 名称,简码,decode(类别,1,'慢性病',2,'特种病','普通病') as 类别,'' 特殊封顶线,'' 封顶线金额 from 保险病种 where 险类=" & Mid(mstrKey, 2) & " Order by 编码"
    ElseIf Mid(mstrKey, 2) = TYPE_重庆银海版 Then
        '调试重庆医保银海版 204-04-07
        gstrSQL = "select ID,编码,名称,简码,decode(类别,1,'特殊病',2,'急诊病',3,'恶性肿瘤',4,'精神病','普通病') as 类别 " & _
                  " ,decode(特殊封顶线,1,'有','') as 特殊封顶线,decode(特殊封顶线,1,decode(封顶线金额,null,'无封顶线',to_char(封顶线金额,'FM900090009990')) ,'') as 封顶线金额" & _
                  " from 保险病种 where 险类=" & Mid(mstrKey, 2) & " Order by 编码"
    Else
        gstrSQL = "select ID,编码,名称,简码,decode(类别,1,'慢性病',2,'特种病','普通病') as 类别 " & _
                  " ,decode(特殊封顶线,1,'有','') as 特殊封顶线,decode(特殊封顶线,1,decode(封顶线金额,null,'无封顶线',to_char(封顶线金额,'FM900090009990')) ,'') as 封顶线金额" & _
                  " from 保险病种 where 险类=" & Mid(mstrKey, 2) & " Order by 编码"
    End If
    Call OpenRecordset(rsTemp, Me.Caption)
    
    Do Until rsTemp.EOF
        Set lst = lvwItem.ListItems.Add(, "K" & rsTemp("ID"), rsTemp("编码"), "Disease", "Disease")
        lst.SubItems(1) = IIf(IsNull(rsTemp("名称")), "", rsTemp("名称"))
        lst.SubItems(2) = IIf(IsNull(rsTemp("简码")), "", rsTemp("简码"))
        lst.SubItems(3) = rsTemp("类别")
        lst.SubItems(4) = Nvl(rsTemp("特殊封顶线"), "")
        lst.SubItems(5) = Nvl(rsTemp("封顶线金额"), "")
        rsTemp.MoveNext
    Loop
    
    If lvwItem.ListItems.Count > 0 Then
        On Error Resume Next
        Set lst = lvwItem.ListItems(strItemKey)
        If Err <> 0 Then
            Set lst = lvwItem.ListItems(1)
            lst.Selected = True
            lst.EnsureVisible
        Else
            Err.Clear
            lst.Selected = True
            lst.EnsureVisible
        End If
    End If
    Call SetMenu
    Me.MousePointer = vbDefault
End Sub

Private Sub FillItem()
    Dim rsTemp As New ADODB.Recordset
    Dim lst As ListItem
    
    lvwLimit.ListItems.Clear
    If lvwItem.SelectedItem Is Nothing Then Exit Sub
    
    gstrSQL = " Select A.ID,A.编码,A.名称,'' 规格,decode(B.性质,1,'1-允许',2,'2-排斥','0-不限') 性质,0 排序 from 保险支付大类 A,保险特准项目 B where A.ID=B.收费细目ID and B.大类=1 and B.病种ID=" & Mid(lvwItem.SelectedItem.Key, 2) & _
              " Union " & _
              " Select A.ID,A.编码,A.名称,A.规格,decode(B.性质,1,'1-允许',2,'2-排斥','0-不限') 性质,1 排序 from 收费细目 A,保险特准项目 B where A.ID=B.收费细目ID and B.大类=0 and B.病种ID=" & Mid(lvwItem.SelectedItem.Key, 2) & _
              " Order by 排序,编码"
    Call OpenRecordset(rsTemp, Me.Caption)
    
    Do Until rsTemp.EOF
        Set lst = lvwLimit.ListItems.Add(, "K" & rsTemp!排序 & "_" & rsTemp("ID"), rsTemp("编码"), "Fix", "Fix")
        lst.SubItems(1) = rsTemp("名称")
        lst.SubItems(2) = Nvl(rsTemp("规格"))
        lst.SubItems(3) = IIf(rsTemp!排序 = "0", "是", "否")
        lst.SubItems(4) = Nvl(rsTemp("性质"))
        rsTemp.MoveNext
    Loop
End Sub

Private Sub SetMenu()
'功能：根据当前内容设置菜单的可用性
    Dim bln非自贡 As Boolean
    Dim bln非泸州 As Boolean
    
    Call FillItem
    stbThis.Panels(2).Text = lvwKind_S.SelectedItem.Text & "共有" & lvwItem.ListItems.Count & "条病种记录"
    
    bln非自贡 = Val(Mid(lvwKind_S.SelectedItem.Key, 2)) <> TYPE_自贡市
    bln非泸州 = Val(Mid(lvwKind_S.SelectedItem.Key, 2)) <> TYPE_泸州市
    
    mnuEditAdd.Enabled = bln非泸州
    mnuEditModify.Enabled = Not (lvwItem.SelectedItem Is Nothing) And bln非泸州
    mnuEditDelete.Enabled = Not (lvwItem.SelectedItem Is Nothing) And bln非自贡 And bln非泸州
    tbrThis.Buttons("New").Enabled = mnuEditAdd.Enabled
    tbrThis.Buttons("Modify").Enabled = mnuEditModify.Enabled
    tbrThis.Buttons("Delete").Enabled = mnuEditDelete.Enabled
End Sub

Private Sub 权限控制()
    If InStr(gstrPrivs, "增删改") = 0 Then
        tbrThis.Buttons("New").Visible = False
        tbrThis.Buttons("Modify").Visible = False
        tbrThis.Buttons("Delete").Visible = False
        tbrThis.Buttons("Split1").Visible = False
        
        mnuEdit.Visible = False
        mnuEditAdd.Visible = False
        mnuEditModify.Visible = False
        
        mnuShortMenu(0).Visible = False
        mnuShortMenu(1).Visible = False
        mnuShortMenu(2).Visible = False
        mnuShortSplit.Visible = False
    End If
End Sub

Public Sub ShowForm(frmParent As Form)
'功能：装入医保类别
'说明：使用本功能的主要原因是在出错退出时窗体不会闪
    Dim rsTemp As New ADODB.Recordset
    Dim strIcon As String, lst As ListItem
    
    
    gstrSQL = "select 序号,名称,是否固定 from 保险类别 where nvl(是否禁止,0)<>1 And 医保部件 Is NULL order by 序号"
    Call OpenRecordset(rsTemp, Me.Caption)
    
    If rsTemp.RecordCount = 0 Then
        '如果是在窗体初始化时调用，就不用处理其它内容了
        MsgBox "没有可用保险类别，不能使用本功能。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '现在才能开始使用控件
    If frm保险病种.Visible = True Then
        frm保险病种.Show
        Exit Sub
    End If
    
    mstrKey = ""
    lvwKind_S.ListItems.Clear
    Do Until rsTemp.EOF
        strIcon = IIf(rsTemp("是否固定") = 1, "Fix", "Common")
        If rsTemp("序号") = gintInsure Then strIcon = strIcon & "D"
        
        Set lst = lvwKind_S.ListItems.Add(, "K" & rsTemp("序号"), rsTemp("名称"), strIcon, strIcon)
        If rsTemp("序号") = gintInsure Then
            lst.Selected = True
        End If
        
        rsTemp.MoveNext
    Loop
    If lvwKind_S.SelectedItem Is Nothing Then
        lvwKind_S.ListItems(1).Selected = True
    End If
    frm保险病种.Show , frmParent
End Sub


Public Function CheckForm() As Boolean
'功能：装入医保类别
'说明：使用本功能的主要原因是在出错退出时窗体不会闪
    Dim rsTemp As New ADODB.Recordset
    Dim strIcon As String, lst As ListItem
    
    
    gstrSQL = "select 序号,名称,是否固定 from 保险类别 where nvl(是否禁止,0)<>1 And 医保部件 Is NULL order by 序号"
    Call OpenRecordset(rsTemp, Me.Caption)
    
    If rsTemp.RecordCount = 0 Then
        '如果是在窗体初始化时调用，就不用处理其它内容了
        MsgBox "没有可用保险类别，不能使用本功能。", vbInformation, gstrSysName
        Exit Function
    End If
    
    '现在才能开始使用控件
    If frm保险病种.Visible = True Then
        CheckForm = True
        Exit Function
    End If
    
    mstrKey = ""
    lvwKind_S.ListItems.Clear
    Do Until rsTemp.EOF
        strIcon = IIf(rsTemp("是否固定") = 1, "Fix", "Common")
        If rsTemp("序号") = gintInsure Then strIcon = strIcon & "D"
        
        Set lst = lvwKind_S.ListItems.Add(, "K" & rsTemp("序号"), rsTemp("名称"), strIcon, strIcon)
        If rsTemp("序号") = gintInsure Then
            lst.Selected = True
        End If
        
        rsTemp.MoveNext
    Loop
    If lvwKind_S.SelectedItem Is Nothing Then
        lvwKind_S.ListItems(1).Selected = True
    End If
    CheckForm = True
End Function
