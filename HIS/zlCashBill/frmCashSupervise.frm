VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCashSupervise 
   Caption         =   "收费财务监控"
   ClientHeight    =   6795
   ClientLeft      =   -135
   ClientTop       =   240
   ClientWidth     =   10800
   Icon            =   "frmCashSupervise.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   10800
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picGroup 
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   30
      ScaleHeight     =   420
      ScaleWidth      =   3015
      TabIndex        =   12
      Top             =   780
      Width           =   3015
      Begin VB.ComboBox cbo人员组 
         Height          =   300
         Left            =   810
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   105
         Width           =   2100
      End
      Begin VB.Label lblGroup 
         AutoSize        =   -1  'True
         Caption         =   "人员分组"
         Height          =   180
         Left            =   60
         TabIndex        =   14
         Top             =   165
         Width           =   720
      End
   End
   Begin MSComctlLib.ImageList ilssmall 
      Left            =   4350
      Top             =   4140
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
            Picture         =   "frmCashSupervise.frx":0442
            Key             =   "man"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   3150
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5175
      ScaleMode       =   0  'User
      ScaleWidth      =   38.572
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   840
      Width           =   45
   End
   Begin MSComctlLib.ImageList Ilsrw 
      Left            =   5760
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashSupervise.frx":0766
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashSupervise.frx":0980
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashSupervise.frx":0BA0
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashSupervise.frx":0DC0
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashSupervise.frx":0FE0
            Key             =   "View"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashSupervise.frx":1200
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashSupervise.frx":1420
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashSupervise.frx":1640
            Key             =   "Filter"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Ilscolor 
      Left            =   5160
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashSupervise.frx":185A
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashSupervise.frx":1A7A
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashSupervise.frx":1C9A
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashSupervise.frx":1EBA
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashSupervise.frx":20DA
            Key             =   "View"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashSupervise.frx":22FA
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashSupervise.frx":251A
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashSupervise.frx":273A
            Key             =   "Filter"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   10800
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "Toolbar1"
      MinHeight1      =   720
      Width1          =   8070
      Key1            =   "only"
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Caption2        =   "人员性质"
      Child2          =   "cboKind"
      MinWidth2       =   1110
      MinHeight2      =   300
      Width2          =   2010
      NewRow2         =   0   'False
      BandStyle2      =   1
      AllowVertical2  =   0   'False
      Begin VB.ComboBox cboKind 
         Height          =   300
         Left            =   9600
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   240
         Width           =   1170
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   720
         Left            =   165
         TabIndex        =   1
         Top             =   30
         Width           =   8505
         _ExtentX        =   15002
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "Ilsrw"
         HotImageList    =   "Ilscolor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   11
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
               Key             =   "Split1"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "增加"
               Key             =   "New"
               Object.ToolTipText     =   "增加"
               Object.Tag             =   "增加"
               ImageKey        =   "New"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "PayFree"
                     Text            =   "手工缴款(&A)"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "PayAll"
                     Text            =   "全额缴款(&B)"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "PayDay"
                     Text            =   "按日缴款(&C)"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "Delete"
               Object.ToolTipText     =   "删除"
               Object.Tag             =   "删除"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split2"
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "过滤"
               Key             =   "Filter"
               Object.ToolTipText     =   "过滤条件设置"
               Object.Tag             =   "过滤"
               ImageKey        =   "Filter"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Quit"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ilsbig 
      Left            =   4560
      Top             =   2790
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
            Picture         =   "frmCashSupervise.frx":2954
            Key             =   "man"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picContainer 
      BackColor       =   &H00FFFFFF&
      Height          =   5550
      Left            =   3210
      ScaleHeight     =   5490
      ScaleWidth      =   6195
      TabIndex        =   3
      Top             =   810
      Width           =   6255
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshRecord 
         Height          =   2370
         Left            =   120
         TabIndex        =   6
         Top             =   2730
         Width           =   5940
         _ExtentX        =   10478
         _ExtentY        =   4180
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         BackColorFixed  =   -2147483648
         BackColorBkg    =   16777215
         GridColor       =   8421504
         GridColorFixed  =   8421504
         GridLinesFixed  =   1
         GridLinesUnpopulated=   1
         SelectionMode   =   1
         MergeCells      =   2
         AllowUserResizing=   1
         Appearance      =   0
         MouseIcon       =   "frmCashSupervise.frx":2C78
         _NumberOfBands  =   1
         _Band(0).Cols   =   8
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshTotal 
         Height          =   1455
         Left            =   180
         TabIndex        =   7
         Top             =   690
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   2566
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorFixed  =   -2147483648
         BackColorBkg    =   -2147483643
         BackColorUnpopulated=   -2147483644
         GridColor       =   8421504
         GridColorFixed  =   8421504
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         GridLinesFixed  =   1
         AllowUserResizing=   1
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
      End
      Begin VB.Label lblSplit 
         BackStyle       =   0  'Transparent
         Height          =   60
         Left            =   600
         MousePointer    =   7  'Size N S
         TabIndex        =   9
         Top             =   2370
         Width           =   1065
      End
      Begin VB.Label lblCaption1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "暂存金额情况"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   1920
         TabIndex        =   8
         Top             =   150
         Width           =   2175
      End
      Begin VB.Label lblCaption2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "缴款记录"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   2310
         TabIndex        =   4
         Top             =   2250
         Width           =   1455
      End
   End
   Begin MSComctlLib.ListView lvwMain_S 
      Height          =   5040
      Left            =   45
      TabIndex        =   2
      Top             =   1275
      Width           =   3030
      _ExtentX        =   5345
      _ExtentY        =   8890
      Arrange         =   2
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDropMode     =   1
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ilsbig"
      SmallIcons      =   "ilssmall"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDropMode     =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Name"
         Text            =   "收款员"
         Object.Width           =   3175
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   10
      Top             =   6435
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   635
      SimpleText      =   $"frmCashSupervise.frx":2F92
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmCashSupervise.frx":2FD9
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13970
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
      Begin VB.Menu mnuFileSet 
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
      Begin VB.Menu mnusplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuPay 
      Caption         =   "缴款(&J)"
      Begin VB.Menu mnuPayNewFree 
         Caption         =   "新增手工缴款(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuPayNewAll 
         Caption         =   "新增全额缴款(&B)"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuPayNewDay 
         Caption         =   "新增按日缴款(C)"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuDelSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPayDelete 
         Caption         =   "删除缴款记录(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuPayPrint 
         Caption         =   "重打缴款单(&P)"
      End
      Begin VB.Menu mnuEditSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditPersonGroup 
         Caption         =   "人员分组(&F)"
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
      Begin VB.Menu mnuviewsplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewAll 
         Caption         =   "所有收款员(&A)"
      End
      Begin VB.Menu mnuViewHave 
         Caption         =   "有暂存金额的(&H)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuviewspilt2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "大图标(&G)"
         Checked         =   -1  'True
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
      Begin VB.Menu mnuViewSplit4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFilter 
         Caption         =   "过滤(&I)"
      End
      Begin VB.Menu mnuViewFlash 
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
   Begin VB.Menu mnuAdd 
      Caption         =   "快捷菜单"
      Visible         =   0   'False
      Begin VB.Menu mnuAddAll 
         Caption         =   "显示所有收款员(&A)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuAddHave 
         Caption         =   "显示有暂存金额的收款员(&H)"
      End
      Begin VB.Menu mnuaddsplit 
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
Attribute VB_Name = "frmCashSupervise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private msngStart As Single    '移动前鼠标的位置
Private mdatBegin As Date, mdatEnd As Date
Private mblnLoad As Boolean  '窗口还未打开时为真
Private mstrKey As String
Private mstrOperator As String, mstrPrivs As String, mlngModul As Long
Private mrsHandin As Recordset '缴款记录
Private mblnDateMoved As Boolean '当前时间范围是否在转出之前
Private mblnGroups As Boolean '是否存在分组
Private mblnNotClick As Boolean
Private Sub cboKind_Click()
    If cboKind.Text <> cboKind.Tag And Me.Visible Then
        Call FillTree
        cboKind.Tag = cboKind.Text
    End If
End Sub
Private Function LoadGroups() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载组信息
    '编制:刘兴洪
    '返回:成功,返回true,否则返回False
    '日期:2010-11-29 10:32:02
    '问题:33633
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim lngPreID As Long, strWhere As String
    
    On Error GoTo errHandle
    gstrSQL = "" & _
    "   Select A.Id, A.组名称,A.简码, A.说明, A.负责人id, A.删除日期,B.姓名 as 负责人  " & _
    "   From 财务缴款分组 A,人员表 B " & _
    "   Where A.负责人ID=B.Id(+) And A.删除日期>Sysdate " & _
    "   Order by ID"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If cbo人员组.ListIndex >= 0 Then lngPreID = cbo人员组.ItemData(cbo人员组.ListIndex)
    mblnNotClick = True
    With rsTemp
        cbo人员组.Clear
        mblnGroups = .RecordCount <> 0
        If Not zlStr.IsHavePrivs(mstrPrivs, "所有人员组") Then
            rsTemp.Filter = "  负责人ID=" & UserInfo.ID
        End If
        Do While Not .EOF
            cbo人员组.AddItem Nvl(rsTemp!组名称)
            cbo人员组.ItemData(cbo人员组.NewIndex) = Val(Nvl(rsTemp!ID))
            If Val(Nvl(rsTemp!ID)) = lngPreID Then cbo人员组.ListIndex = cbo人员组.NewIndex
            rsTemp.MoveNext
        Loop
        If cbo人员组.ListCount > 0 And cbo人员组.ListIndex < 0 Then cbo人员组.ListIndex = 0
        If mblnGroups = True And cbo人员组.ListCount = 0 Then
            ShowMsgbox "你没有任何组的操作权限,请与系统管理员联系并授权(所有人员组或分配成组负责人)!"
            picGroup.Visible = False
            Call Form_Resize
            Call picGroup_Resize
            Exit Function
        End If
    End With
    picGroup.Visible = mblnGroups
    Call Form_Resize
    Call picGroup_Resize
    Call FillTree
    mblnNotClick = False
    
    LoadGroups = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
 
 
Private Sub cbo人员组_Click()
    If mblnNotClick = True Then Exit Sub
    '加载人员组信息
    Call FillTree
End Sub
Private Sub Form_Activate()
    If mblnLoad = True Then
        Call Form_Resize '为了使CoolBar自适应高度
        If LoadGroups = False Then mblnLoad = False:   Unload Me: Exit Sub
        'If FillTree() = False Then mblnLoad = False:   Unload Me: Exit Sub
    End If
    mblnLoad = False
End Sub

Private Sub Form_Load()
    mblnLoad = True
    mstrPrivs = gstrPrivs
    mlngModul = glngModul
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
    
    Call 权限控制
    Call InitFace
    '-----------
    RestoreWinState Me, App.ProductName
    
    mnuViewAll.Checked = zlDatabase.GetPara("显示所有收款员", glngSys, mlngModul, "0") = "1"
    mnuViewHave.Checked = Not mnuViewAll.Checked
    'Call FillTree
    '根据LvwMain显示设置对应菜单
     mnuViewIcon_Click lvwMain_S.View
End Sub

Private Sub InitFace()
    '初始化表格
    Dim arrTemp1 As Variant, arrTemp2 As Variant, arrTemp3 As Variant
    Dim intColumn As Integer, i As Long
    
    '初始化数据
    mdatEnd = TruncateDate(zlDatabase.Currentdate)
    mdatBegin = TruncateDate(DateAdd("m", -1, mdatEnd))
    mblnDateMoved = zlDatabase.DateMoved(Format(mdatBegin, "yyyy-MM-dd hh:mm:ss"), , , Me.Caption)
    
    arrTemp1 = Array("日期", "结算方式", "结算金额", "结算号", "截止时间", "登记人", "摘要", "缴款部门")
    arrTemp2 = Array(" 1999年10月31日 ", "  结算方式 ", "-########0.00", Space(15), " yyyy-MM-dd HH:mm:ss ", " 东方不败 ", Space(30), Space(15))
    arrTemp3 = Array(1, 1, 7, 1, 1, 1, 1, 1)
    mshRecord.Row = 0
    For intColumn = 0 To mshRecord.Cols - 1
        mshRecord.Col = intColumn
        mshRecord.Text = arrTemp1(intColumn)
        mshRecord.ColWidth(intColumn) = TextWidth(arrTemp2(intColumn))
        mshRecord.ColAlignment(intColumn) = arrTemp3(intColumn)
        mshRecord.CellAlignment = 4
    Next                              '初始化缴款记录表
    mshRecord.ColAlignment(2) = 7
    mshRecord.MergeCol(0) = True
    
    mshTotal.ColAlignment(0) = 2
    arrTemp1 = Array("结算方式", "期初暂存", "缴款合计", "期末暂存")
    arrTemp2 = Array("1234567890", "123456789.123", "123456789.123", "123456789.123")
    arrTemp3 = Array(1, 7, 7, 7)
    mshTotal.Row = 0
    For intColumn = 0 To mshTotal.Cols - 1
        mshTotal.Col = intColumn
        mshTotal.Text = arrTemp1(intColumn)
        mshTotal.ColWidth(intColumn) = TextWidth(arrTemp2(intColumn))
        mshTotal.ColAlignment(intColumn) = arrTemp3(intColumn)
        mshTotal.CellAlignment = 4
    Next '初始化缴款记录表
    
    arrTemp1 = Array("全部", "门诊挂号员", "门诊收费员", "预交收款员", "住院结帐员", "入院登记员", "发卡登记人")
    cboKind.Clear
    For i = 0 To UBound(arrTemp1)
        cboKind.AddItem arrTemp1(i)
    Next
    cboKind.ListIndex = 0   '调用click事件
    
End Sub
 
Private Sub Form_Resize()
    Dim sngTop As Single, sngBottom As Single
    On Error Resume Next
    sngTop = IIf(CoolBar1.Visible, CoolBar1.Top + CoolBar1.Height, 0)
    sngBottom = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0)
    
    picGroup.Top = IIf(picGroup.Visible, sngTop, 0)
    lvwMain_S.Top = IIf(picGroup.Visible, picGroup.Height + 50, 0) + sngTop
    lvwMain_S.Height = IIf(sngBottom - lvwMain_S.Top > 0, sngBottom - lvwMain_S.Top, 0)
    lvwMain_S.Left = 0
    picGroup.Width = lvwMain_S.Width
    picGroup.Left = 0
    
    picSplit.Top = sngTop
    picSplit.Height = IIf(sngBottom - picSplit.Top > 0, sngBottom - picSplit.Top, 0)
    picSplit.Left = lvwMain_S.Left + lvwMain_S.Width
    
    picContainer.Left = picSplit.Left + picSplit.Width
    picContainer.Top = sngTop
    picContainer.Width = ScaleWidth - picContainer.Left
    picContainer.Height = sngBottom - picContainer.Top
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrKey = ""
    Set mrsHandin = Nothing
    zlDatabase.SetPara "显示所有收款员", IIf(mnuViewAll.Checked, 1, 0), glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
    
    SaveWinState Me, App.ProductName
End Sub

Private Sub lvwMain_S_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwMain_S.SortOrder = IIf(lvwMain_S.SortOrder = lvwAscending, lvwDescending, lvwAscending)
End Sub

Private Sub lvwMain_S_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If mstrKey = Item.Key Then Exit Sub
    mstrKey = Item.Key
    
    FillList Item.Text
End Sub

Private Sub lvwMain_S_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    If Button = 2 Then
        mnuAddAll.Checked = mnuViewAll.Checked
        mnuAddHave.Checked = mnuViewHave.Checked
        For i = 0 To 3
            mnuShortIcon(i).Checked = mnuViewIcon(i).Checked
        Next
        PopupMenu mnuAdd, 2
    End If
End Sub

Private Sub mnuEditPersonGroup_Click()
    If frmGroupAndPesons.ShowGroups(Me, mlngModul, mstrPrivs) = False Then
        Exit Sub
    End If
    '重新加载数据
    Call LoadGroups
End Sub

Private Sub mnuHelpTopic_Click()
   Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hWnd)
End Sub

Private Sub mnuPayPrint_Click()
   Dim lng单据ID As Long
   
   lng单据ID = mshRecord.RowData(mshRecord.Row)
   If lng单据ID = 0 Then Exit Sub
   
   If MsgBox("你确定要重打" & mshRecord.TextMatrix(mshRecord.Row, 0) & "的缴款单吗?", vbYesNo + vbQuestion, gstrSysName) = vbYes Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_BILL_1500", Me, "单据ID=" & lng单据ID, 2)
   End If
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    If lvwMain_S.SelectedItem Is Nothing Then Exit Sub
    
    If Val(mshRecord.RowData(1)) = 0 Then
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
            "收费员=" & lvwMain_S.SelectedItem.Text)
    Else
        With mshRecord
            Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
                "收费员=" & lvwMain_S.SelectedItem.Text, "单据ID=" & .RowData(.Row), "截止时间=" & .TextMatrix(.Row, MshGetColNum(mshRecord, "截止时间")), _
                "登记人=" & .TextMatrix(.Row, MshGetColNum(mshRecord, "登记人")))
        End With
    End If
End Sub

Private Sub mnuViewFlash_Click()
    '刷新,先加组
     Call LoadGroups
    
End Sub

Private Sub mnuViewIcon_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
    Next
    mnuViewIcon(Index).Checked = True
    lvwMain_S.View = Index
End Sub

Private Sub mnuShortIcon_Click(Index As Integer)
    mnuViewIcon_Click Index
End Sub

Private Sub mnuAddAll_Click()
    mnuViewAll_Click
End Sub

Private Sub mnuAddHave_Click()
    mnuviewHave_Click
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub mnuFilePreview_Click()
    subPrint 2
End Sub

Private Sub mnuFileSet_Click()
        zlPrintSet
End Sub

Private Sub mnuPayDelete_Click()
    On Error GoTo errH
    Dim rsTmp As New Recordset
    Dim datSys As Date, i As Long, strTmp As String
    On Error GoTo errH:
    
    With mshRecord
        If .RowData(.Row) = 0 Then Exit Sub
        
        If MsgBox("你确实要删除日期为" & Trim(.TextMatrix(.Row, 0)) & "的缴款登记卡吗？", _
            vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
            
            gstrSQL = zlGetFullFieldsTable("人员缴款记录", 1, "Where id=[1]", False, "")
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(.RowData(.Row)))
            
            If Not rsTmp Is Nothing Then
                If Not rsTmp.EOF Then
                     MsgBox "当前选择的缴款记录在后备数据表中!" & vbCrLf _
                         & "请与系统管理员联系,转入到在线数据表再操作!", vbInformation, gstrSysName
                     Exit Sub
                End If
            End If
            
            datSys = zlDatabase.Currentdate
            i = datSys - CDate(.TextMatrix(.Row, 0))
            
            If i > 1 Then strTmp = "警告:即将删除的缴款记录是昨天以前的!" & vbCrLf & vbCrLf
            strTmp = strTmp & "为避免误删除,须再次确认才能继续." & vbCrLf & vbCrLf & "请输入OK"
            
            If UCase(InputBox(strTmp, "操作确认")) <> "OK" Then
                MsgBox "输入的确认关键字不是OK!将取消本次操作!", vbInformation + vbOKOnly, Me.Caption
                Exit Sub
            End If
            
            gstrSQL = "zl_人员缴款记录_delete(" & .RowData(.Row) & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            If mnuViewHave.Checked = True Then
                FillTree
            Else
                FillList lvwMain_S.SelectedItem.Text
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuViewFilter_Click()
    Dim List1 As ListItem, strPersonelKind As String
    
    If cboKind.ListIndex > 0 Then strPersonelKind = cboKind.Text
    If Not lvwMain_S.SelectedItem Is Nothing Then mstrOperator = lvwMain_S.SelectedItem.Text
    If frmTimeSet.ShowMe(Me, 0, 0, mlngModul, mstrPrivs, mdatBegin, mdatEnd, mstrOperator, mblnDateMoved, strPersonelKind, mnuViewHave.Checked) = True Then
        
        If mstrOperator <> "" Then
            For Each List1 In lvwMain_S.ListItems
                If List1.Text = mstrOperator Then
                    List1.Selected = True
                    Call List1.EnsureVisible
                    Exit For
                End If
            Next
        End If
        
        If Not lvwMain_S.SelectedItem Is Nothing Then
            FillList lvwMain_S.SelectedItem.Text
        End If
    End If
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    CoolBar1.Visible = mnuViewToolButton.Checked
    CoolBar1.Bands("only").MinHeight = Toolbar1.Height
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

Private Sub mshRecord_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngCol As Long, i As Long, lngID As Long
    Dim bln空格 As Boolean, strColName As String
    
    lngCol = mshRecord.MouseCol
    
    If Button = 1 And mshRecord.MousePointer = 99 Then
        strColName = mshRecord.TextMatrix(0, lngCol)
        If strColName = "" Then Exit Sub
        If mrsHandin Is Nothing Then Exit Sub
        
        mshRecord.ColData(lngCol) = (mshRecord.ColData(lngCol) + 1) Mod 2
        strColName = Switch(strColName = "日期", "登记时间", strColName = "结算金额", "金额", strColName <> "", strColName)
        mrsHandin.Sort = strColName & IIf(mshRecord.ColData(lngCol) = 0, "", " DESC")
                
        i = 1
        Do Until mrsHandin.EOF
            If mrsHandin("单据ID") <> lngID Then
                lngID = mrsHandin("单据ID")
                bln空格 = Not bln空格
            End If
        
            mshRecord.TextMatrix(i, 0) = Format(mrsHandin("登记时间"), "yyyy年MM月dd日") & IIf(bln空格, " ", "")
            mshRecord.TextMatrix(i, 1) = mrsHandin("结算方式")
            mshRecord.TextMatrix(i, 2) = Format(mrsHandin("金额"), "##########0.00;-##########0.00;;")
            mshRecord.TextMatrix(i, 3) = IIf(IsNull(mrsHandin("结算号")), "", mrsHandin("结算号"))
            mshRecord.TextMatrix(i, 4) = Format(Nvl(mrsHandin!截止时间), "yyyy-MM-dd HH:mm:ss")
            mshRecord.TextMatrix(i, 5) = IIf(IsNull(mrsHandin("登记人")), " ", mrsHandin("登记人"))
            mshRecord.TextMatrix(i, 6) = IIf(IsNull(mrsHandin("摘要")), " ", mrsHandin("摘要"))
            mshRecord.TextMatrix(i, 7) = IIf(IsNull(mrsHandin("缴款部门")), " ", mrsHandin("缴款部门"))
            mshRecord.RowData(i) = mrsHandin("单据ID")
            i = i + 1
            mrsHandin.MoveNext
        Loop
        mshRecord.Row = 1
    End If
End Sub

Private Sub mshRecord_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mshRecord.MouseRow = 0 Then
        mshRecord.MousePointer = 99
    Else
        mshRecord.MousePointer = 0
    End If
End Sub

Private Sub mshRecord_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And mnuPay.Visible Then PopupMenu mnuPay, 2
End Sub

Private Sub mnuPayNewFree_Click()
'功能：新增手工缴款记录
    If frmCashPay.编辑缴款记录(lvwMain_S.SelectedItem.Text, Mid(lvwMain_S.SelectedItem.Key, 2)) = True Then
        If mnuViewHave.Checked = True Then
            FillTree
        Else
            FillList lvwMain_S.SelectedItem.Text
        End If
    End If
End Sub
Private Sub mnuPayNewDay_Click()
    If frmCashPayAll.ShowMe(lvwMain_S.SelectedItem.Text, Mid(lvwMain_S.SelectedItem.Key, 2), Me, PM_按日缴款) Then
        If mnuViewHave.Checked Then
            Call FillTree
        Else
            Call FillList(lvwMain_S.SelectedItem.Text)
        End If
    End If
End Sub
Private Sub mnuPayNewAll_Click()
'功能：新增全额缴款记录
    If frmCashPayAll.ShowMe(lvwMain_S.SelectedItem.Text, Mid(lvwMain_S.SelectedItem.Key, 2), Me, PM_全额缴款) Then
        If mnuViewHave.Checked Then
            Call FillTree
        Else
            Call FillList(lvwMain_S.SelectedItem.Text)
        End If
    End If
End Sub

Private Sub mnuViewAll_Click()
    mnuViewAll.Checked = Not mnuViewAll.Checked
    mnuViewHave.Checked = Not mnuViewAll.Checked
    If Me.Visible Then FillTree
End Sub

Private Sub mnuviewHave_Click()
    mnuViewHave.Checked = Not mnuViewHave.Checked
    mnuViewAll.Checked = Not mnuViewHave.Checked
    
    If Me.Visible Then FillTree
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Form_Resize
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub picSplit_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then msngStart = x
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sngTemp As Single
    If Button = 1 Then
        sngTemp = picSplit.Left + x - msngStart
        If sngTemp - lvwMain_S.Left > 1500 And Me.ScaleWidth - sngTemp > 3000 Then
            picSplit.Left = sngTemp
            lvwMain_S.Width = picSplit.Left - lvwMain_S.Left
            picGroup.Width = lvwMain_S.Width
            
            picContainer.Left = picSplit.Left + picSplit.Width
            picContainer.Width = Me.ScaleWidth - picContainer.Left
            
        End If
    End If
End Sub

Private Sub picContainer_Resize()
    On Error Resume Next
    lblCaption1.Left = (picContainer.ScaleWidth - lblCaption1.Width) / 2
    mshTotal.Top = lblCaption1.Top + lblCaption1.Height + 300
    mshTotal.Left = -15
    mshTotal.Width = picContainer.ScaleWidth + 30
    mshTotal.Height = lblSplit.Top - mshTotal.Top
    
    lblSplit.Width = picContainer.ScaleWidth
    
    lblCaption2.Top = lblSplit.Top + lblSplit.Height + 100
    lblCaption2.Left = (picContainer.ScaleWidth - lblCaption2.Width) / 2
    mshRecord.Top = lblCaption2.Top + lblCaption2.Height + 300
    mshRecord.Left = -15
    mshRecord.Width = picContainer.ScaleWidth + 30
    mshRecord.Height = picContainer.ScaleHeight - mshRecord.Top
    
End Sub

Private Sub lblSplit_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then msngStart = y
End Sub

Private Sub lblSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sngTemp As Single
    If Button = 1 Then
        sngTemp = lblSplit.Top + y - msngStart
        If sngTemp - mshTotal.Top > 1000 And picContainer.ScaleHeight - sngTemp > 1500 Then
            lblSplit.Top = sngTemp
            mshTotal.Height = sngTemp - mshTotal.Top
            lblCaption2.Top = lblSplit.Top + lblSplit.Height + 100
            mshRecord.Top = lblCaption2.Top + lblCaption2.Height + 300
            mshRecord.Height = picContainer.ScaleHeight - mshRecord.Top
        End If
    End If
End Sub

 

Private Sub picGroup_Resize()
    Err = 0: On Error Resume Next
    With picGroup
        '33633
        cbo人员组.Width = .ScaleWidth - cbo人员组.Left - 50
    End With
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "New"
            If Button.ButtonMenus("PayFree").Visible Then
                mnuPayNewFree_Click
            ElseIf Button.ButtonMenus("PayAll").Visible Then
                mnuPayNewAll_Click
            ElseIf Button.ButtonMenus("PayDay").Visible Then
                mnuPayNewDay_Click
            End If
        Case "Delete"
            mnuPayDelete_Click
        Case "Filter"
            mnuViewFilter_Click
        Case "Quit"
            mnuFileExit_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Preview"
            mnuFilePreview_Click
        Case "Help"
            mnuHelpTopic_Click
        Case "View"
            mnuViewIcon(lvwMain_S.View).Checked = False
            If lvwMain_S.View = 3 Then
                mnuViewIcon(0).Checked = True
                lvwMain_S.View = 0
            Else
                mnuViewIcon(lvwMain_S.View + 1).Checked = True
                lvwMain_S.View = lvwMain_S.View + 1
            End If
    End Select
End Sub

Private Sub ClearTable()
    Dim i As Integer

    mshTotal.Rows = 2
    mshRecord.Rows = 2
    For i = 0 To mshRecord.Cols - 1
        mshRecord.TextMatrix(1, i) = ""
    Next
    For i = 0 To mshTotal.Cols - 1
        mshTotal.TextMatrix(1, i) = ""
    Next
    mshRecord.RowData(1) = 0
    Call SetMenu
End Sub

Private Function FillTree() As Boolean
'功能:装入所有收费员到lvwMain_S
    Dim strKey As String, strKind As String
    Dim rs收费员 As New ADODB.Recordset
    Dim lng组ID As Long
    On Error GoTo errH
    
    '得到收款员名单
    mstrKey = ""
    gstrSQL = ""
    If cboKind.ListIndex > 0 Then
        strKind = cboKind.Text
        gstrSQL = " And C.人员性质=[1]"
    ElseIf mnuViewHave.Checked = False Then
        gstrSQL = " And C.人员性质 in ('门诊挂号员','门诊收费员','预交收款员','住院结帐员','入院登记员','发卡登记人')"
    End If
    
    If mnuViewHave.Checked = True Then
        '在指点定期间内有暂存金的操作员
        gstrSQL = "" & _
        "   Select Distinct A.收款员,B.ID " & _
        "    From 人员缴款余额 A,人员表 B,人员性质说明 C" & IIf(mblnGroups, ",缴款成员组成 M", "") & vbNewLine & _
        "   Where A.收款员=B.姓名 And 余额<>0 And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null) " & _
        "           And B.id=C.人员ID" & gstrSQL & vbNewLine & _
                IIf(mblnGroups, " And B.ID=M.成员ID And M.组ID=[2] ", "") & _
        "   Order by 收款员"
    Else
        '所有期间内操作员
        gstrSQL = "" & _
        "   Select Distinct A.姓名 as 收款员,A.ID  " & _
        "   From 人员表 A,人员性质说明 C" & IIf(mblnGroups, ",缴款成员组成 M", "") & vbNewLine & _
        "   Where A.ID=C.人员ID And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
        "           And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & gstrSQL & vbNewLine & _
                IIf(mblnGroups, " And A.ID=M.成员ID And M.组ID=[2] ", "") & _
        "   Order by 收款员"
    End If
    If cbo人员组.ListIndex < 0 Then
        lng组ID = 0
    Else
        lng组ID = cbo人员组.ItemData(cbo人员组.ListIndex)
    End If
    DoEvents
    Me.Refresh
    Set rs收费员 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strKind, lng组ID)
    'If rs收费员.EOF Then MsgBox "当前没有收费员，不能运行本程序。", vbExclamation, gstrSysName: Exit Function
    If Not lvwMain_S.SelectedItem Is Nothing Then
        strKey = lvwMain_S.SelectedItem.Key
    End If
    
    With lvwMain_S.ListItems
        .Clear
        Do Until rs收费员.EOF
            If Not IsNull(rs收费员("收款员")) Then
                .Add , "C" & rs收费员("ID"), rs收费员("收款员"), "man", "man"
            End If
            rs收费员.MoveNext
        Loop
        If .Count > 0 Then
            Dim Item As ListItem
            On Error Resume Next
            Set Item = lvwMain_S.ListItems(strKey)
            If Err <> 0 Then
                Set Item = lvwMain_S.ListItems(1)
                Item.Selected = True
                Item.EnsureVisible
            Else
                Err.Clear
                Item.Selected = True
                Item.EnsureVisible
            End If
            FillList lvwMain_S.SelectedItem.Text
        Else
            FillList ""
        End If
    End With
    FillTree = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub FillList(ByVal str收费员 As String)
'功能:显示指定收费员的收款汇总表和缴款记录
'参数:str收费员 收费员的名字
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    Dim lngID As Long
    Dim bln空格 As Boolean, strDate As String
    
    On Error GoTo errH
    
    If str收费员 = "" Then
        Call ClearTable
        Exit Sub
    End If
    
    '显示统计表
    strDate = Format(mdatEnd, "yyyyMMdd")
    gstrSQL = _
        "Select 结算方式,sum(余额+缴款合计) as 期初,sum(缴款合计-期末缴款) as 期内,sum(余额+期末缴款) as 期末 from( " & _
        "Select 结算方式,金额 as 缴款合计, " & _
        "Decode(Sign(To_Char(登记时间,'YYYYMMDD')-[3]),1,金额,0) as 期末缴款,0 as 余额 " & _
        "From " & IIf(mblnDateMoved, zlGetFullFieldsTable("人员缴款记录"), "人员缴款记录 ") & _
        "Where 收款员 = [1] and 登记时间>=[2] " & _
        "Union All " & _
        "Select 结算方式,0 as 缴款合计,0 as 期末缴款,余额 " & _
        "From 人员缴款余额 " & _
        "Where 性质=1 and 余额<>0 and 收款员 =[1]) " & _
        " group by 结算方式 " & _
        " having sum(余额+缴款合计)<>0 or sum(缴款合计-期末缴款)<>0 or sum(余额+期末缴款)<>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str收费员, mdatBegin, strDate)
    
    If rsTmp.EOF Then
        mshTotal.Rows = 2
        For i = 0 To mshTotal.Cols - 1
            mshTotal.TextMatrix(1, i) = ""
        Next
    Else
        mshTotal.Rows = rsTmp.RecordCount + 1
        i = 1
        Do Until rsTmp.EOF
            mshTotal.TextMatrix(i, 0) = rsTmp("结算方式")
            mshTotal.TextMatrix(i, 1) = Format(rsTmp("期初"), "##########0.00;-##########0.00; ;")
            mshTotal.TextMatrix(i, 2) = Format(rsTmp("期内"), "##########0.00;-##########0.00; ;")
            mshTotal.TextMatrix(i, 3) = Format(rsTmp("期末"), "##########0.00;-##########0.00; ;")
            i = i + 1
            rsTmp.MoveNext
        Loop
    End If
    rsTmp.Close
    
    
    '显示缴款记录
    gstrSQL = _
        "Select 单据ID,登记时间,结算方式,金额,结算号,截止时间,登记人,摘要,B.名称 缴款部门" & _
        " From " & IIf(mblnDateMoved, zlGetFullFieldsTable("人员缴款记录"), "人员缴款记录") & " A,部门表 B Where A.收款部门ID=B.ID(+) And 收款员=[1]" & _
        " And 登记时间 Between [2] And [3] order by 登记时间"
    Set mrsHandin = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str收费员, mdatBegin, DateAdd("s", -1, DateAdd("d", 1, mdatEnd)))
    
    If mrsHandin.EOF Then
        mshRecord.Rows = 2
        For i = 0 To mshRecord.Cols - 1
            mshRecord.TextMatrix(1, i) = ""
        Next
        mshRecord.RowData(1) = 0
        Call SetMenu
    Else
        mshRecord.Rows = mrsHandin.RecordCount + 1
        i = 1
        Do Until mrsHandin.EOF
            If mrsHandin("单据ID") <> lngID Then
                lngID = mrsHandin("单据ID")
                bln空格 = Not bln空格
            End If
        
            mshRecord.TextMatrix(i, 0) = Format(mrsHandin("登记时间"), "yyyy年MM月dd日") & IIf(bln空格, " ", "")
            mshRecord.TextMatrix(i, 1) = mrsHandin("结算方式")
            mshRecord.TextMatrix(i, 2) = Format(mrsHandin("金额"), "##########0.00;-##########0.00;;")
            mshRecord.TextMatrix(i, 3) = IIf(IsNull(mrsHandin("结算号")), "", mrsHandin("结算号"))
            mshRecord.TextMatrix(i, 4) = Format(Nvl(mrsHandin!截止时间), "yyyy-MM-dd HH:mm:ss")
            mshRecord.TextMatrix(i, 5) = IIf(IsNull(mrsHandin("登记人")), " ", mrsHandin("登记人"))
            mshRecord.TextMatrix(i, 6) = IIf(IsNull(mrsHandin("摘要")), " ", mrsHandin("摘要"))
            mshRecord.TextMatrix(i, 7) = IIf(IsNull(mrsHandin("缴款部门")), " ", mrsHandin("缴款部门"))
            mshRecord.RowData(i) = mrsHandin("单据ID")
            i = i + 1
            mrsHandin.MoveNext
        Loop
        mshRecord.Row = 1
        Call SetMenu
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub subPrint(bytMode As Byte)
'功能:进行打印,预览和输出到EXCEL
'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    Dim objPrint1Grd As New zlPrint1Grd
    Dim objAppRow As New zlTabAppRow
    
    If lvwMain_S.ListItems.Count = 0 Then Exit Sub
    objPrint1Grd.Title.Text = "缴款记录表"
    objPrint1Grd.Title.Color = RGB(255, 0, 0)
    objPrint1Grd.Title.Font.Name = lblCaption2.Font.Name
    objPrint1Grd.Title.Font.Size = lblCaption2.Font.Size
    
    objAppRow.Add "收款员：" & lvwMain_S.SelectedItem.Text
    objAppRow.Add "时间范围：" & Format(mdatBegin, "YYYY年MM月DD日") & "至" & Format(mdatEnd, "YYYY年MM月DD日")
    objPrint1Grd.UnderAppRows.Add objAppRow
    Set objPrint1Grd.Body = mshRecord
    
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint1Grd)
          Case 1
               zlPrintOrView1Grd objPrint1Grd, 1
          Case 2
              zlPrintOrView1Grd objPrint1Grd, 2
          Case 3
              zlPrintOrView1Grd objPrint1Grd, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint1Grd, bytMode
    End If
    
    Set objPrint1Grd = Nothing
    Set objAppRow = Nothing
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim i As Integer
    
    Select Case ButtonMenu.Key
        Case "PayFree"
            Call mnuPayNewFree_Click
        Case "PayAll"
            Call mnuPayNewAll_Click
        Case "PayDay"
            Call mnuPayNewDay_Click
        Case Else
            For i = 0 To 3
                mnuViewIcon(i).Checked = False
            Next
            mnuViewIcon(ButtonMenu.Index - 1).Checked = True
            lvwMain_S.View = ButtonMenu.Index - 1
    End Select
End Sub

Private Sub Toolbar1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuViewTool, 2
End Sub

Private Sub 权限控制()
'功能:由于有的用户权限不够,故使一些菜单项或按钮不可见
    If InStr(mstrPrivs, "删除缴款") = 0 And InStr(mstrPrivs, "手工缴款") = 0 _
        And InStr(mstrPrivs, "全额缴款") = 0 _
        And InStr(mstrPrivs, "按日缴款") = 0 And zlStr.IsHavePrivs(mstrPrivs, "成员分组") = False Then
        mnuPay.Visible = False
        Toolbar1.Buttons("New").Visible = False
        Toolbar1.Buttons("Delete").Visible = False
        Toolbar1.Buttons("Split2").Visible = False
    Else
        If InStr(mstrPrivs, "手工缴款") = 0 And InStr(mstrPrivs, "全额缴款") = 0 And InStr(mstrPrivs, "按日缴款") = 0 Then
            Toolbar1.Buttons("New").Visible = False
        End If
        If InStr(mstrPrivs, "手工缴款") = 0 Then
            mnuPayNewFree.Visible = False
            Toolbar1.Buttons("New").ButtonMenus("PayFree").Visible = False
        End If
        If InStr(mstrPrivs, "全额缴款") = 0 Then
            mnuPayNewAll.Visible = False
            Toolbar1.Buttons("New").ButtonMenus("PayAll").Visible = False
        End If
        If InStr(mstrPrivs, "按日缴款") = 0 Then
            mnuPayNewDay.Visible = False
            Toolbar1.Buttons("New").ButtonMenus("PayDay").Visible = False
        End If
        If InStr(mstrPrivs, "删除缴款") = 0 Then
            mnuPayDelete.Visible = False
            Toolbar1.Buttons("Delete").Visible = False
        End If
        If InStr(mstrPrivs, "重打缴款单") = 0 Then
            mnuPayPrint.Visible = False
        End If
        mnuEditPersonGroup.Visible = zlStr.IsHavePrivs(mstrPrivs, "成员分组")
        mnuEditSplit.Visible = mnuEditPersonGroup.Visible
    End If
End Sub

Private Sub SetMenu()
    Dim blnNew As Boolean
    Dim blnDelete As Boolean
    Dim lngCount As Long, lngID As Long
    Dim i As Integer
    
    blnNew = Not (lvwMain_S.SelectedItem Is Nothing)
    blnDelete = mshRecord.RowData(mshRecord.Row) <> 0
    
    mnuPayNewFree.Enabled = blnNew
    mnuPayNewAll.Enabled = blnNew
    mnuPayNewDay.Enabled = blnNew
    Toolbar1.Buttons("New").Enabled = blnNew
    
    mnuPayDelete.Enabled = blnDelete
    Toolbar1.Buttons("Delete").Enabled = blnDelete
    mnuPayPrint.Enabled = blnDelete
    
    blnDelete = mshRecord.RowData(1) <> 0
    mnuFilePreview.Enabled = blnDelete
    mnuFilePrint.Enabled = blnDelete
    mnuFileExcel.Enabled = blnDelete
    Toolbar1.Buttons("Preview").Enabled = blnDelete
    Toolbar1.Buttons("Print").Enabled = blnDelete
    
    
    For i = 1 To mshRecord.Rows - 1
        If lngID <> mshRecord.RowData(i) Then
            lngID = mshRecord.RowData(i)
            lngCount = lngCount + 1
        End If
    Next
    If lvwMain_S.SelectedItem Is Nothing Then
        stbThis.Panels(2).Text = ""
    Else
        stbThis.Panels(2).Text = lvwMain_S.SelectedItem.Text & "在" & _
            Format(mdatBegin, "yyyy年MM月dd日") & "——" & _
            Format(mdatEnd, "yyyy年MM月dd日") & "之间共有" & lngCount & "条缴款记录。"
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

