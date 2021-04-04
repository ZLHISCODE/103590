VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frmUnit 
   Caption         =   "合约单位管理"
   ClientHeight    =   4980
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   6645
   Icon            =   "frmUnit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.ImageList ils32 
      Left            =   2580
      Top             =   1020
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":0442
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":089A
            Key             =   "ItemNo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":0CEE
            Key             =   "Write"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   2700
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":1146
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":159E
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":19F6
            Key             =   "ItemNo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":1E4A
            Key             =   "No"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":22A2
            Key             =   "Write"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3225
      Left            =   3120
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3225
      ScaleWidth      =   30
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1590
      Width           =   30
   End
   Begin MSComctlLib.ListView lvwMain 
      Height          =   2235
      Left            =   3120
      TabIndex        =   1
      Top             =   1380
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   3942
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
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "名称"
         Text            =   "名称"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "编码"
         Text            =   "编码"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "简码"
         Text            =   "简码"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "地址"
         Text            =   "地址"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "电话"
         Text            =   "电话"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "开户银行"
         Text            =   "开户银行"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "帐号"
         Text            =   "帐号"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Key             =   "联系人"
         Text            =   "联系人"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Key             =   "站点"
         Text            =   "院区"
         Object.Width           =   1411
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwMain_S 
      Height          =   3345
      Left            =   240
      TabIndex        =   0
      Top             =   990
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   5900
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   18
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList Ilsrw 
      Left            =   4470
      Top             =   60
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
            Picture         =   "frmUnit.frx":26FA
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":291A
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":2B3A
            Key             =   "Parent"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":2D56
            Key             =   "Child"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":2F72
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":3192
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":33B2
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":35D2
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":37F2
            Key             =   "View"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":3A12
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":3C32
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Ilscolor 
      Left            =   3270
      Top             =   210
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
            Picture         =   "frmUnit.frx":3E52
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":4072
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":4292
            Key             =   "Parent"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":44AE
            Key             =   "Child"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":46CA
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":48EA
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":4B0A
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":4D2A
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":4F4A
            Key             =   "View"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":516A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":538A
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   6645
      _CBHeight       =   780
      _Version        =   "6.7.8988"
      Child1          =   "Toolbar1"
      MinHeight1      =   720
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   720
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   6525
         _ExtentX        =   11509
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "Ilsrw"
         HotImageList    =   "Ilscolor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   15
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
               Key             =   "Split"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "分类"
               Key             =   "Parent"
               Object.ToolTipText     =   "增加分类"
               Object.Tag             =   "分类"
               ImageKey        =   "Parent"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "项目"
               Key             =   "Child"
               Object.ToolTipText     =   "增加项目"
               Object.Tag             =   "项目"
               ImageKey        =   "Child"
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
               Key             =   "Start"
               Object.ToolTipText     =   "启用"
               Object.Tag             =   "启用"
               ImageKey        =   "Start"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "停用"
               Key             =   "Stop"
               Object.ToolTipText     =   "停用"
               Object.Tag             =   "停用"
               ImageKey        =   "Stop"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "sdf"
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Object.ToolTipText     =   "帮助"
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   4620
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   635
      SimpleText      =   $"frmUnit.frx":55AA
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmUnit.frx":55F1
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6641
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
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFileset 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuFilepre 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnusplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnufileexit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditAddParent 
         Caption         =   "增加分类(&P)"
      End
      Begin VB.Menu mnuEditAddChild 
         Caption         =   "增加项目(&C)"
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "修改(&M)"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "删除(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditStart 
         Caption         =   "启用(&S)"
      End
      Begin VB.Menu mnuEditStop 
         Caption         =   "停用(&T)"
      End
      Begin VB.Menu mnuEditSplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditExpand 
         Caption         =   "加长下级编码(&X)"
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
         Begin VB.Menu mnuviewspilt1 
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
      Begin VB.Menu mnuviewsplit1 
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
      Begin VB.Menu mnuViewSplit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewSelect 
         Caption         =   "选择列(&C)"
      End
      Begin VB.Menu mnuViewSplit4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewShowAll 
         Caption         =   "显示所有下级(&H)"
      End
      Begin VB.Menu mnuViewShowStop 
         Caption         =   "显示停用单位(&P)"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "刷新(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpTopic 
         Caption         =   "帮助主题(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "Web上的中联"
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
      Begin VB.Menu mnuHelpSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
   Begin VB.Menu mnuShort1 
      Caption         =   "快捷菜单"
      Visible         =   0   'False
      Begin VB.Menu mnuShortMenu1 
         Caption         =   "增加分类(&P)"
         Index           =   1
      End
      Begin VB.Menu mnuShortMenu1 
         Caption         =   "修改(&M)"
         Index           =   2
      End
      Begin VB.Menu mnuShortMenu1 
         Caption         =   "删除(&D)"
         Index           =   3
      End
   End
   Begin VB.Menu mnuShort2 
      Caption         =   "快捷菜单"
      Visible         =   0   'False
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "增加项目(&C)"
         Index           =   1
      End
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "修改(&M)"
         Index           =   2
      End
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "删除(&D)"
         Index           =   3
      End
      Begin VB.Menu mnuShortsplit1 
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
         Checked         =   -1  'True
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Dim msngStartX As Single         '移动前鼠标的位置
Dim mblnItem As Boolean         '为真表示单击到ListView某一项上
Dim mintColumn As Integer
Dim mblnLoad As Boolean
Dim mstrKey As String
Private mstrPrivs As String
Private mlngModul As Long
Private Const mstrLvw As String = "名称,1300,0,1;编码,800,0,2;简码,900,0,0;地址,1440,0,0;电话,1440,0,0;开户银行,1440,0,0;帐号,1440,0,0;联系人,1440,0,0;建档时间,1100,0,0;撤档时间,1100,0,0;所属分类,2000,0,0;院区,800,0,0"

Private Sub Form_Activate()
    If mblnLoad = True Then
        Call 权限控制
        Call Form_Resize '为了使CoolBar自适应高度
        FillTree
    End If
    mblnLoad = False
End Sub

Private Sub Form_Load()
    mlngModul = glngModul
    mstrPrivs = gstrPrivs
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
    mblnLoad = True
    '允许进行列删除的ListView须做标记
    lvwMain.Tag = "可变化的"
    RestoreWinState Me, App.ProductName
    If lvwMain.ColumnHeaders(9).Text = "站点" Then
        lvwMain.ColumnHeaders(9).Text = "院区"
    End If
    mnuViewShowAll.Checked = zlDatabase.GetPara("显示所有下级", glngSys, mlngModul, 0) <> "0"
    mnuViewShowStop.Checked = zlDatabase.GetPara("显示停用单位", glngSys, mlngModul, 0) <> "0"
    '根据LvwMain显示设置对应菜单
     mnuViewIcon_Click lvwMain.View
End Sub

Private Sub Form_Resize()
    Dim sngTop As Single, sngBottom As Single
    sngTop = IIf(CoolBar1.Visible, CoolBar1.Top + CoolBar1.Height, 0)
    sngBottom = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0)
    
    tvwMain_S.Top = sngTop
    tvwMain_S.Height = IIf(sngBottom - tvwMain_S.Top > 0, sngBottom - tvwMain_S.Top, 0)
    tvwMain_S.Left = 0
    
    picSplit.Top = sngTop
    picSplit.Height = IIf(sngBottom - picSplit.Top > 0, sngBottom - picSplit.Top, 0)
    picSplit.Left = tvwMain_S.Left + tvwMain_S.Width
    
    lvwMain.Left = picSplit.Left + picSplit.Width
    lvwMain.Top = sngTop
    lvwMain.Height = IIf(sngBottom - lvwMain.Top > 0, sngBottom - lvwMain.Top, 0)
    If Me.ScaleWidth - lvwMain.Left > 0 Then lvwMain.Width = Me.ScaleWidth - lvwMain.Left
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrKey = ""
    zlDatabase.SetPara "显示所有下级", IIf(mnuViewShowAll.Checked, 1, 0), glngSys, mlngModul
    zlDatabase.SetPara "显示停用单位", IIf(mnuViewShowStop.Checked, 1, 0), glngSys, mlngModul
    SaveWinState Me, App.ProductName
End Sub

Private Sub lvwMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintColumn = ColumnHeader.Index - 1 Then '仍是刚才那列
        lvwMain.SortOrder = IIf(lvwMain.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvwMain.SortKey = mintColumn
        lvwMain.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwMain_DblClick()
    If mblnItem = True And mnuEditModify.Visible Then mnuEditModify_Click
End Sub

Private Sub lvwMain_GotFocus()
    With lvwMain
        stbThis.Panels(2).Text = "单位列表中共显示有" & .ListItems.Count & "个合约单位。"
    End With
    Call SetMenu
End Sub

Private Sub lvwMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mblnItem = True
    SetMenu
    stbThis.Panels(2).Text = "单位列表中共显示有" & lvwMain.ListItems.Count & "个合约单位。"
End Sub

Private Sub lvwMain_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If mnuEditModify.Enabled And mnuEditModify.Visible Then mnuEditModify_Click
    End If
End Sub

Private Sub lvwMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnItem = False
End Sub

Private Sub lvwMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    If Button = 2 Then
        mnuShortMenu2(2).Enabled = mnuEditModify.Enabled
        mnuShortMenu2(3).Enabled = mnuEditDelete.Enabled
        For i = 0 To 3
            mnuShortIcon(i).Checked = mnuViewIcon(i).Checked
        Next
        PopupMenu mnuShort2, vbPopupMenuRightButton
    End If
End Sub

Private Sub mnuEditAddChild_Click()
    Dim str编码 As String
    Dim str名称 As String
    Dim i As Integer
    Dim blnReturn As Boolean
    
    With tvwMain_S.SelectedItem
        If .Key = "Root" Then
           blnReturn = frmUnitEdit.编辑单位("无", "", "", , True)
        Else
            i = InStr(.Text, "】")
            str编码 = Mid(.Text, 2, i - 2)
            str名称 = Mid(.Text, i + 1)
            blnReturn = frmUnitEdit.编辑单位(str名称, Mid(.Key, 2), str编码, , True)
        End If
        If blnReturn = True Then tvwMain_S_NodeClick tvwMain_S.SelectedItem
    End With
End Sub

Private Sub mnuEditAddParent_Click()
    Dim str编码 As String
    Dim str名称 As String
    Dim i As Integer
    Dim strKey As String
    Dim blnReturn As Boolean
    
    With tvwMain_S.SelectedItem
        strKey = .Key
        If .Key = "Root" Then
           blnReturn = frmUnitEdit.编辑单位("无", "", "", , False)
        Else
            i = InStr(.Text, "】")
            str编码 = Mid(.Text, 2, i - 2)
            str名称 = Mid(.Text, i + 1)
           blnReturn = frmUnitEdit.编辑单位(str名称, Mid(.Key, 2), str编码, , False)
        End If
    End With
    If blnReturn = True Then
        FillTree
    End If
End Sub

Private Sub mnuEditExpand_Click()
    Dim strTemp As String
    Dim str父编码 As String
    Dim str编码 As String
    Dim intNew As Integer '目前最长的
    Dim intChild As Integer
    
    On Error GoTo errHandle
    With tvwMain_S.SelectedItem
        If .Key = "Root" Then
            str父编码 = ""
            intNew = GetDownCodeLength("", "合约单位")
            intChild = GetLocalCodeLength("", "合约单位")
        Else
            str父编码 = Mid(.Text, 2, InStr(.Text, "】") - 2)
            intNew = GetDownCodeLength(Mid(.Key, 2), "合约单位")
            intChild = GetLocalCodeLength(Mid(.Key, 2), "合约单位")
        End If
        If intNew = 0 Or intChild = 0 Then Exit Sub
        If intNew = 10 Then
            MsgBox "不能再加长编码，某一个下级已经用足了长度。", vbExclamation, gstrSysName
            Exit Sub
        End If
        
        intNew = frmCount.GetLength(intChild, 10 - (intNew - intChild))
        If intNew = 0 Then Exit Sub
        strTemp = str父编码 & String(intNew - intChild, "0")
        
        If .Key = "Root" Then
            gstrSQL = "zl_合约单位_EXPAND('" & strTemp & "'," & Len(str父编码) + 1 & ",0)"
        Else
            gstrSQL = "zl_合约单位_EXPAND('" & strTemp & "'," & Len(str父编码) + 1 & "," & Mid(.Key, 2) & ")"
        End If
'        Call SQLTest(App.ProductName, Me.Caption, gstrSQL)
'        gcnOracle.Execute gstrSQL, , adCmdStoredProc
'        Call SQLTest
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
        FillTree
    End With
    Exit Sub
errHandle:
    If errCenter() = 1 Then Resume
    Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditModify_Click()
    Dim str编码 As String
    Dim str名称 As String
    Dim i As Integer
    Dim strKey As String
    Dim blnReturn As Boolean
    
    With tvwMain_S.SelectedItem
        strKey = .Key
        If ActiveControl Is tvwMain_S Then
            If .Key = "Root" Then Exit Sub
            If .Parent.Key = "Root" Then
               blnReturn = frmUnitEdit.编辑单位("无", "", "", Mid(.Key, 2))
            Else
                i = InStr(.Parent.Text, "】")
                str编码 = Mid(.Parent.Text, 2, i - 2)
                str名称 = Mid(.Parent.Text, i + 1)
                blnReturn = frmUnitEdit.编辑单位(str名称, Mid(.Parent.Key, 2), str编码, Mid(.Key, 2))
            End If
        Else
            If .Key = "Root" Then
                blnReturn = frmUnitEdit.编辑单位("无", "", "", Mid(lvwMain.SelectedItem.Key, 2))
            Else
                i = InStr(.Text, "】")
                str编码 = Mid(.Text, 2, i - 2)
                str名称 = Mid(.Text, i + 1)
                blnReturn = frmUnitEdit.编辑单位(str名称, Mid(.Key, 2), str编码, Mid(lvwMain.SelectedItem.Key, 2))
            End If
        End If
    End With
    If blnReturn = True Then
        FillTree
    End If
End Sub

Private Sub mnuEditDelete_Click()
    On Error GoTo errHandle
    
    If ActiveControl Is tvwMain_S Then
        If MsgBox("删除分类同时也将删除该类别的项目，" & vbCrLf & "你确认要删除名称为“" & tvwMain_S.SelectedItem.Text & "”的分类项目吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        gstrSQL = "zl_合约单位_delete(" & Mid(tvwMain_S.SelectedItem.Key, 2) & ")"
    Else
        If MsgBox("你确认要删除名称为“" & lvwMain.SelectedItem.Text & "”的合约单位吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        gstrSQL = "zl_合约单位_delete(" & Mid(lvwMain.SelectedItem.Key, 2) & ")"
    End If
    Me.MousePointer = 11
    
'    Call SQLTest(App.ProductName, Me.Caption, gstrSQL)
'    gcnOracle.Execute gstrSQL, , adCmdStoredProc
'    Call SQLTest
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    
    If ActiveControl Is tvwMain_S Then
        FillTree
        Call tvwMain_S_GotFocus
    Else
        FillList tvwMain_S.SelectedItem.Key
        Call lvwMain_GotFocus
    End If
    Me.MousePointer = 0
    Exit Sub
errHandle:
    If errCenter() = 1 Then Resume
    Call SaveErrLog
    Me.MousePointer = 0
End Sub

Private Sub mnuEditStart_Click()
    On Error GoTo errHandle
    Dim strKey As String

    strKey = lvwMain.SelectedItem.Key
    gstrSQL = "zl_合约单位_reuse(" & Mid(lvwMain.SelectedItem.Key, 2) & ")"
    '执行启用过程
'    Call SQLTest(App.ProductName, Me.Caption, gstrSQL)
'    gcnOracle.Execute gstrSQL, , adCmdStoredProc
'    Call SQLTest
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    
    '改变图标和颜色
    With lvwMain.SelectedItem
        .Icon = "Item"
        .SmallIcon = "Item"
        .ForeColor = RGB(0, 0, 0)
        
        Dim i As Integer
        For i = 1 To lvwMain.ColumnHeaders.Count - 1
            .ListSubItems(i).ForeColor = RGB(0, 0, 0)
        Next
    End With
    '改变状态栏和菜单
    SetMenu
    Exit Sub
errHandle:
    If errCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditStop_Click()
    On Error GoTo errHandle
    Dim strKey As String

    strKey = lvwMain.SelectedItem.Key
    gstrSQL = "zl_合约单位_stop(" & Mid(lvwMain.SelectedItem.Key, 2) & ")"
    '执行启用过程
'    Call SQLTest(App.ProductName, Me.Caption, gstrSQL)
'    gcnOracle.Execute gstrSQL, , adCmdStoredProc
'    Call SQLTest
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    
    '改变图标和颜色
    If mnuViewShowStop.Checked = True Then '要显示停用部门
        With lvwMain.SelectedItem
            .Icon = "ItemNo"
            .SmallIcon = "ItemNo"
            .ForeColor = RGB(255, 0, 0)
            
            Dim i As Integer
            For i = 1 To lvwMain.ColumnHeaders.Count - 1
                .ListSubItems(i).ForeColor = RGB(255, 0, 0)
            Next
        End With
        SetMenu
    Else '不显示停用部门
        With lvwMain
            .ListItems.Remove .SelectedItem.Key
            If .ListItems.Count > 0 Then
                .ListItems(1).Selected = True
                .ListItems(1).EnsureVisible
                lvwMain_ItemClick .SelectedItem
            Else
                Call lvwMain_GotFocus
            End If
        End With
    End If
    Exit Sub
errHandle:
    If errCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me)
End Sub

Private Sub mnuShortMenu1_Click(Index As Integer)
    Select Case Index
        Case 1
            mnuEditAddParent_Click
        Case 2
            mnuEditModify_Click
        Case 3
            mnuEditDelete_Click
    End Select
End Sub

Private Sub mnuShortMenu2_Click(Index As Integer)
    Select Case Index
        Case 1
            mnuEditAddChild_Click
        Case 2
            mnuEditModify_Click
        Case 3
            mnuEditDelete_Click
    End Select
End Sub

Private Sub mnuShortIcon_Click(Index As Integer)
    mnuViewIcon_Click Index
End Sub

Private Sub mnuViewIcon_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
    Next
    mnuViewIcon(Index).Checked = True
    lvwMain.View = Index
End Sub

Private Sub mnuViewRefresh_Click()
    Call FillTree
End Sub

Private Sub mnuViewSelect_Click()
    If zlControl.LvwSelectColumns(lvwMain, mstrLvw) = True Then
        '列有变化就要重新刷新
        FillList tvwMain_S.SelectedItem.Key
    End If
End Sub

Private Sub mnuViewShowAll_Click()
    mnuViewShowAll.Checked = Not mnuViewShowAll.Checked
    FillList tvwMain_S.SelectedItem.Key
End Sub

Private Sub mnuViewShowStop_Click()
    mnuViewShowStop.Checked = Not mnuViewShowStop.Checked
    FillList tvwMain_S.SelectedItem.Key
End Sub

Private Sub picsplit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        msngStartX = X
    End If
End Sub

Private Sub picsplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sngTemp As Single
    If Button = 1 Then
        sngTemp = picSplit.Left + X - msngStartX
        If sngTemp > 0 And Me.ScaleWidth - (sngTemp + picSplit.Width) > 0 Then
            picSplit.Left = sngTemp
            tvwMain_S.Width = picSplit.Left - tvwMain_S.Left
            lvwMain.Left = picSplit.Left + picSplit.Width
            lvwMain.Width = Me.ScaleWidth - lvwMain.Left
        End If
    End If
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnufilepre_Click()
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub mnufileset_Click()
    zlPrintSet
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Parent"
            mnuEditAddParent_Click
        Case "Child"
            mnuEditAddChild_Click
        Case "Modify"
            mnuEditModify_Click
        Case "Delete"
            mnuEditDelete_Click
        Case "Stop"
            mnuEditStop_Click
        Case "Start"
            mnuEditStart_Click
        Case "Quit"
            mnuFileExit_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Preview"
            mnufilepre_Click
        Case "Help"
            mnuhelptopic_Click
        Case "View"
            mnuViewIcon(lvwMain.View).Checked = False
            If lvwMain.View = 3 Then
                mnuViewIcon(0).Checked = True
                lvwMain.View = 0
            Else
                mnuViewIcon(lvwMain.View + 1).Checked = True
                lvwMain.View = lvwMain.View + 1
            End If
    End Select
    
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    CoolBar1.Visible = mnuViewToolButton.Checked
    CoolBar1.Bands("only").MinHeight = Toolbar1.Height
    Form_Resize
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
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

Private Sub mnuhelptopic_Click()
ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hwnd)
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim i As Integer
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
    Next
    mnuViewIcon(ButtonMenu.Index - 1).Checked = True
    lvwMain.View = ButtonMenu.Index - 1
End Sub

Private Sub Toolbar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool
    End If
End Sub

Private Sub tvwMain_S_GotFocus()
    stbThis.Panels(2).Text = "本单位分类有" & lvwMain.ListItems.Count & "个下级项目"
    SetMenu
End Sub

Private Sub tvwMain_S_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If mnuShortMenu1(1).Visible = False Then Exit Sub
        mnuShortMenu1(2).Enabled = mnuEditModify.Enabled
        mnuShortMenu1(3).Enabled = mnuEditDelete.Enabled
        PopupMenu mnuShort1, vbPopupMenuRightButton
    End If
End Sub

Private Sub tvwMain_S_NodeClick(ByVal Node As MSComctlLib.Node)
    If mstrKey = Node.Key Then Exit Sub
    mstrKey = Node.Key
    
    FillList Node.Key
    mnuEditExpand.Enabled = lvwMain.ListItems.Count <> 0 Or Node.Children <> 0
    tvwMain_S_GotFocus
End Sub

Private Sub subPrint(bytMode As Byte)
'功能:进行打印,预览和输出到EXCEL
'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As New zlPrintLvw
    objPrint.Title.Text = "合约单位"
    Set objPrint.Body.objData = lvwMain
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

Private Sub FillTree()
'功能:装入合约单位的所有分类到tvwMain_S
    Dim strTemp As String
    Dim strKey As String
    Dim rs合约单位 As New ADODB.Recordset
    
    mstrKey = ""
    
    rs合约单位.CursorLocation = adUseClient
    rs合约单位.CursorType = adOpenKeyset
    rs合约单位.LockType = adLockReadOnly
    If Not tvwMain_S.SelectedItem Is Nothing Then
        strKey = tvwMain_S.SelectedItem.Key
    End If
    
    On Error GoTo errHandle
    gstrSQL = "select ID,上级ID,编码,名称 from 合约单位  " & _
        "where 末级 <> 1 start with 上级ID is null connect by prior ID =上级ID"
    Set rs合约单位 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    tvwMain_S.Nodes.Clear
    tvwMain_S.Nodes.Add , , "Root", "所有合约单位", "Root", "Root"
    tvwMain_S.Nodes("Root").Sorted = True
    Do Until rs合约单位.EOF
        
        If IsNull(rs合约单位("上级id")) Then
            tvwMain_S.Nodes.Add "Root", tvwChild, "C" & rs合约单位("id"), "【" & rs合约单位("编码") & "】" & rs合约单位("名称"), "Write", "Write"
        Else
            tvwMain_S.Nodes.Add "C" & rs合约单位("上级id"), tvwChild, "C" & rs合约单位("id"), "【" & rs合约单位("编码") & "】" & rs合约单位("名称"), "Write", "Write"
        End If
        tvwMain_S.Nodes("C" & rs合约单位("ID")).Sorted = True
        rs合约单位.MoveNext
    Loop
    
    Dim nod As Node
    On Error Resume Next
    Set nod = tvwMain_S.Nodes(strKey)
    If Err <> 0 Then
        Set nod = tvwMain_S.Nodes("Root")
        nod.Selected = True
        nod.Expanded = True
        tvwMain_S_NodeClick nod
    Else
        Err.Clear
        nod.Selected = True
        nod.Expanded = True
        nod.EnsureVisible
        tvwMain_S_NodeClick nod
    End If
    Exit Sub
errHandle:
    If errCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub FillList(ByVal str合约单位ID As String)
'功能:装入对应分类的子分类和项目到lvwMain
'参数:str合约单位ID 分类的标识
    Dim rs合约单位 As New ADODB.Recordset
    Dim fld As Field
    Dim lst As ListItem
    Dim strKey As String
    Dim str停用 As String
    
    If Not lvwMain.SelectedItem Is Nothing Then
        '保留原有键值
        strKey = lvwMain.SelectedItem.Key
    End If
    
    rs合约单位.CursorLocation = adUseClient
    rs合约单位.CursorType = adOpenKeyset
    rs合约单位.LockType = adLockReadOnly
    
    If mnuViewShowStop.Checked = False Then
        str停用 = " (撤档时间 is null or 撤档时间 = to_date('3000-01-01','YYYY-MM-DD')) and "
    End If
    
    On Error GoTo errHandle
    'by lesfeng 2010-03-08 性能优化
    If mnuViewShowAll.Checked = True Then
        gstrSQL = "select A.ID,A.上级ID,A.名称,A.编码,A.简码,A.地址,A.电话,A.开户银行,A.帐号,A.联系人,A.建档时间,A.撤档时间,A.站点 院区,B.名称 as 所属分类 from " & _
            "(select ID,上级ID,名称,编码,简码,地址,电话,开户银行,帐号,联系人,to_char(建档时间,'YYYY-MM-DD') as 建档时间,to_char(撤档时间,'YYYY-MM-DD') as 撤档时间,站点 院区 from 合约单位 where " & _
            IIf(str停用 = "", "", str停用) & " 末级=1 connect by prior id=上级id start with  " & _
            IIf(str合约单位ID = "Root", "上级ID is null ", "上级ID = [1] ") & ") A,合约单位 B where A.上级ID=B.ID(+)"
    Else
        gstrSQL = "select ID,上级ID,名称,编码,简码,地址,电话,开户银行,帐号,联系人,to_char(建档时间,'YYYY-MM-DD') as 建档时间,to_char(撤档时间,'YYYY-MM-DD') as 撤档时间,'' as 所属分类,站点 院区 from 合约单位 where " & _
            IIf(str停用 = "", "", str停用) & " 末级=1 and " & IIf(str合约单位ID = "Root", "上级ID is null ", "上级ID = [1]")
    End If
    
    Set rs合约单位 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Mid(str合约单位ID, 2))
    
    LockWindowUpdate lvwMain.hwnd
    lvwMain.ListItems.Clear
    Do Until rs合约单位.EOF
        If CDate(IIf(IsNull(rs合约单位("撤档时间")), "3000-01-01", rs合约单位("撤档时间"))) = CDate("3000-01-01") Then
            Set lst = lvwMain.ListItems.Add(, "C" & rs合约单位("ID"), rs合约单位("名称"), "Item", "Item")
        Else
            Set lst = lvwMain.ListItems.Add(, "C" & rs合约单位("ID"), rs合约单位("名称"), "ItemNo", "ItemNo")
            lst.ForeColor = RGB(255, 0, 0)
        End If
        
        Dim lngCol  As Long
        Dim varValue As Variant
        '根据ListView的列名从数据库取数
        For lngCol = 2 To lvwMain.ColumnHeaders.Count
            varValue = rs合约单位(lvwMain.ColumnHeaders(lngCol).Text).Value
            lst.SubItems(lngCol - 1) = IIf(IsNull(varValue), "", varValue)
            If lst.Icon = "ItemNo" Then lst.ListSubItems(lngCol - 1).ForeColor = RGB(255, 0, 0)
        Next
        rs合约单位.MoveNext
    Loop
    LockWindowUpdate 0
    
    If lvwMain.ListItems.Count > 0 Then
        Dim Item As ListItem
        On Error Resume Next
        Set Item = lvwMain.ListItems(strKey)
        If Err <> 0 Then
            Set Item = lvwMain.ListItems(1)
            Item.Selected = True
            Item.EnsureVisible
            lvwMain_ItemClick Item
        Else
            Err.Clear
            Item.Selected = True
            Item.EnsureVisible
            lvwMain_ItemClick Item
        End If
    Else
        Call SetMenu
    End If
    Exit Sub
errHandle:
    If errCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetMenu()
'功能:设置修改和删除按钮的有效值
'参数:blnEnabled 有效值
'    Dim blnNew As Boolean
    Dim blnModify As Boolean
    Dim blnStart As Boolean
    Dim blnStop As Boolean
    
    If ActiveControl Is tvwMain_S Then
        blnStart = False
        blnStop = False
        blnModify = tvwMain_S.SelectedItem.Key <> "Root"
    Else
        If lvwMain.SelectedItem Is Nothing Or lvwMain.ListItems.Count = 0 Then
            blnStart = False
            blnStop = False
            blnModify = False
        Else
            blnStart = (lvwMain.SelectedItem.Icon = "ItemNo")
            blnStop = (lvwMain.SelectedItem.Icon = "Item")
            blnModify = (lvwMain.SelectedItem.Icon = "Item")
        End If
    End If
    '整体赋值
'    Toolbar1.Buttons("Parent").Enabled = blnNew
'    Toolbar1.Buttons("Child").Enabled = blnNew
'    mnuEditAddParent.Enabled = blnNew
'    mnuEditAddChild.Enabled = blnNew
    
    Toolbar1.Buttons("Modify").Enabled = blnModify
    Toolbar1.Buttons("Delete").Enabled = blnModify
    mnuEditDelete.Enabled = blnModify
    mnuEditModify.Enabled = blnModify
    
    Toolbar1.Buttons("Start").Enabled = blnStart
    Toolbar1.Buttons("Stop").Enabled = blnStop
    mnuEditStart.Enabled = blnStart
    mnuEditStop.Enabled = blnStop

    EnablePrint (lvwMain.ListItems.Count > 0)
End Sub

Private Sub 权限控制()
'功能:由于有的用户权限不够,故使一些菜单项或按钮不可见
    If InStr(mstrPrivs, "增删改") = 0 Then
        mnuEdit.Visible = False
        mnuEditModify.Visible = False
        mnuShortMenu1(1).Visible = False
        mnuShortMenu2(1).Visible = False
        mnuShortMenu2(2).Visible = False
        mnuShortMenu2(3).Visible = False
        mnuShortsplit1.Visible = -False
        Toolbar1.Buttons("Split").Visible = False
        Toolbar1.Buttons("Parent").Visible = False
        Toolbar1.Buttons("Child").Visible = False
        Toolbar1.Buttons("Modify").Visible = False
        Toolbar1.Buttons("Delete").Visible = False
        Toolbar1.Buttons("Split1").Visible = False
        Toolbar1.Buttons("Start").Visible = False
        Toolbar1.Buttons("Stop").Visible = False
    End If
End Sub

Private Sub EnablePrint(ByVal blnEnabled As Boolean)
'功能:设置打印和预鉴按钮的有效值
'参数:blnEnabled 有效值
    Toolbar1.Buttons("Print").Enabled = blnEnabled
    Toolbar1.Buttons("Preview").Enabled = blnEnabled
    mnuFilepre.Enabled = blnEnabled
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

