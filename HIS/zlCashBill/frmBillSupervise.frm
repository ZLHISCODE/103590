VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmBillSupervise 
   Caption         =   "票据使用监控"
   ClientHeight    =   6510
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   9195
   Icon            =   "frmBillSupervise.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   9195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   6150
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   635
      SimpleText      =   $"frmBillSupervise.frx":0442
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmBillSupervise.frx":0489
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11139
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
   Begin VB.PictureBox picH 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   6480
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleMode       =   0  'User
      ScaleWidth      =   1530.013
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2820
      Width           =   1785
   End
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2715
      Left            =   3180
      ScaleHeight     =   2715
      ScaleMode       =   0  'User
      ScaleWidth      =   38.572
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3120
      Width           =   45
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh汇总 
      Height          =   2145
      Left            =   4170
      TabIndex        =   5
      Top             =   3480
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   3784
      _Version        =   393216
      Rows            =   3
      Cols            =   4
      FixedRows       =   2
      FixedCols       =   0
      BackColorFixed  =   -2147483648
      BackColorBkg    =   -2147483643
      BackColorUnpopulated=   -2147483644
      GridColor       =   8421504
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   0
      GridLinesFixed  =   1
      MergeCells      =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin MSComctlLib.ImageList Ilscolor 
      Left            =   7965
      Top             =   930
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
            Picture         =   "frmBillSupervise.frx":0D1D
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":0F3D
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":115D
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":137D
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":159D
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":17BD
            Key             =   "Cancel"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":19DD
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":1BFD
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":1E17
            Key             =   "View"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":2033
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":2253
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":2473
            Key             =   "Check"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Ilsrw 
      Left            =   8670
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
            Picture         =   "frmBillSupervise.frx":268D
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":28A7
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":2AC7
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":2CE7
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":2F07
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":3127
            Key             =   "Cancel"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":3347
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":3567
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":3781
            Key             =   "View"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":399D
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":3BBD
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":3DDD
            Key             =   "Check"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   9195
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "Toolbar1"
      MinHeight1      =   720
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Caption2        =   "使用类别"
      Child2          =   "cbo类别"
      MinWidth2       =   1995
      MinHeight2      =   300
      Width2          =   1695
      NewRow2         =   0   'False
      Begin VB.ComboBox cbo类别 
         Height          =   300
         Left            =   7110
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   1995
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   720
         Left            =   165
         TabIndex        =   7
         Top             =   30
         Width           =   5940
         _ExtentX        =   10478
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
               Caption         =   "领用"
               Key             =   "New"
               Object.ToolTipText     =   "领用票据"
               Object.Tag             =   "领用"
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修改"
               Key             =   "Modify"
               Object.ToolTipText     =   "修改记录"
               Object.Tag             =   "修改"
               ImageKey        =   "Modify"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "Delete"
               Object.ToolTipText     =   "删除记录"
               Object.Tag             =   "删除"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split1"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "报损"
               Key             =   "Cancel"
               Object.ToolTipText     =   "票据报损"
               Object.Tag             =   "报损"
               ImageKey        =   "Cancel"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "核对"
               Key             =   "Check"
               Object.ToolTipText     =   "核对票据明细"
               Object.Tag             =   "核对"
               ImageKey        =   "Check"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split2"
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "过滤"
               Key             =   "Filter"
               Description     =   "过滤条件"
               Object.ToolTipText     =   "过滤条件"
               Object.Tag             =   "过滤"
               ImageKey        =   "Filter"
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
   Begin MSComctlLib.ImageList ils32 
      Left            =   2880
      Top             =   1350
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":3FF7
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":4449
            Key             =   "C2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":4763
            Key             =   "C3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":4A7D
            Key             =   "C5"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":4D97
            Key             =   "C1"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":50B1
            Key             =   "C4"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":53CB
            Key             =   "C7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillSupervise.frx":56E5
            Key             =   "C6"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   2970
      Top             =   2220
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
            Picture         =   "frmBillSupervise.frx":59FF
            Key             =   "Item"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvw领用_S 
      Height          =   1185
      Left            =   3960
      TabIndex        =   8
      Top             =   900
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   2090
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
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvwMain 
      Height          =   3345
      Left            =   330
      TabIndex        =   9
      Top             =   1890
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   5900
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ils32"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "票据种类"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblDown 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "使用情况"
      Height          =   240
      Left            =   7170
      TabIndex        =   4
      Top             =   3030
      Width           =   1095
   End
   Begin VB.Label lblUp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "领用记录"
      Height          =   240
      Left            =   7020
      TabIndex        =   3
      Top             =   1110
      Width           =   1095
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
      Begin VB.Menu mnuFileSpit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuBill 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuBillGet 
         Caption         =   "领用票据(&N)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuBillModify 
         Caption         =   "修改记录(&M)"
      End
      Begin VB.Menu mnuBillDelete 
         Caption         =   "删除记录(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuBillSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBillCancel 
         Caption         =   "票据报损(&R)"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuBillCheck 
         Caption         =   "核对领用单(&B)"
         Index           =   0
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuBillCheck 
         Caption         =   "核对票据明细(&C)"
         Index           =   1
         Shortcut        =   ^H
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
         Caption         =   "显示所有领用记录(&A)"
      End
      Begin VB.Menu mnuViewHave 
         Caption         =   "仅显示未用完(&P)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewCheck 
         Caption         =   "显示核对信息(&H)"
      End
      Begin VB.Menu mnuviewsplit2 
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
      Begin VB.Menu mnuViewDetail 
         Caption         =   "明细清单(&E)"
      End
      Begin VB.Menu mnuViewSelect 
         Caption         =   "选择列(&C)"
      End
      Begin VB.Menu mnuViewFilter 
         Caption         =   "过滤(&I)"
      End
      Begin VB.Menu mnuViewSplit45 
         Caption         =   "-"
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
         Caption         =   "显示所有领用记录(&A)"
      End
      Begin VB.Menu mnuAddHave 
         Caption         =   "仅显示未用完(&P)"
      End
      Begin VB.Menu mnuAddSplit 
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
   Begin VB.Menu mnuAdd2 
      Caption         =   "快捷菜单"
      Visible         =   0   'False
      Begin VB.Menu mnuAddDetail 
         Caption         =   "明细清单(&D)"
      End
   End
End
Attribute VB_Name = "frmBillSupervise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private msngStart As Single    '移动前鼠标的位置
Private mdatBegin As Date, mdatEnd As Date
Private mstrOperator As String  '过滤的票据领用人
Private mlngModul As Long

Private mblnUnload As Boolean
Private mblnLoad As Boolean  '窗口还未打开时为真
Private mblnItem As Boolean  '为真表示单击到ListView某一项上
Private mintColumn As Integer '
Private mstr票据 As String   '上一次的票据种类
Private mstr票据长度  As String  '票据长度
Private mstrKey As String    '上一次的记录
Private mstrPrivs As String
Private Const mstrLvw As String = "开始号码,1000,0,1;终止号码,1000,0,2;使用类别,1000,0,1;" & _
    "领用人,800,0,2;当前号码,1000,0,0;剩余数量,600,0,0;" & _
    "批次,1000,0,0;使用方式,600,0,0;登记时间,1200,0,0;登记人,800,0,2;核对人,800,0,2;" & _
    "前缀文本,0,0,2;签字人,800,0,2;签字时间,1200,0,2"
Private mblnNotClick As Boolean
Private mblnNOMoved As Boolean '当前票据是不否在后备数据表中
Private mblnDateMoved As Boolean '当前时间范围是否在转出之前
Private mbln药店  As Boolean

Private Sub cbo类别_Click()
    Call SetDefaultUserType
    If mblnNotClick = True Then Exit Sub
    mstrKey = ""
    Call Fill记录
End Sub

Private Sub Form_Activate()
    If mblnUnload Then Unload Me: Exit Sub
    Call LoadCombox
    
    If mblnLoad = True Then
        Call Form_Resize '为了使CoolBar自适应高度
        'Call Fill记录
         If lvwMain.Enabled And lvwMain.View Then lvwMain.SetFocus
        Call lvwMain_ItemClick(lvwMain.SelectedItem)
    End If
End Sub

Private Sub Form_Load()
    mblnLoad = True
    mblnUnload = False
    mstrPrivs = gstrPrivs
    mlngModul = glngModul
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
    
    Call PrivilegeCTRL
    If Not InitFace Then
        mblnUnload = True
        Exit Sub
    End If
    lvw领用_S.Tag = "可变化的"
    '-----------
    RestoreWinState Me, App.ProductName
    
    mnuViewAll.Checked = zlDatabase.GetPara("显示所有领用记录", glngSys, mlngModul, "0") = "1"
    mnuViewHave.Checked = Not mnuViewAll.Checked
    mnuViewCheck.Checked = zlDatabase.GetPara("查看核对信息", glngSys, mlngModul, "0") = "1"
    '如果ListView的还未被设置，比如第一次使用，那就调用缺省的初始化
    If lvw领用_S.ColumnHeaders.Count <> UBound(Split(mstrLvw, ";")) + 1 Then
        zlControl.LvwSelectColumns lvw领用_S, mstrLvw, True
    End If
    '根据lvw领用_S显示设置对应菜单
     mnuViewIcon_Click lvw领用_S.View
     
     
    '创建第三方票据打印部件
    On Error Resume Next
    gblnBillPrint = False
    Set gobjBillPrint = CreateObject("zlBillPrint.clsBillPrint")
    If Not gobjBillPrint Is Nothing Then
        gblnBillPrint = gobjBillPrint.zlInitialize(gcnOracle, glngSys, glngModul, UserInfo.编号, UserInfo.姓名)
    End If
    Err.Clear: On Error GoTo 0
End Sub

Private Function InitFace() As Boolean
    Dim arrTemp1 As Variant, arrTemp2 As Variant, arrTemp3 As Variant
    Dim i As Integer, strTmp As String
    Dim objListItem As ListItem, strKeyValue As String
    
    '初始化数据
    mdatEnd = TruncateDate(zlDatabase.Currentdate)
    mdatBegin = TruncateDate(DateAdd("m", -1, mdatEnd))
    mblnDateMoved = zlDatabase.DateMoved(Format(mdatBegin, "yyyy-MM-dd hh:mm:ss"), , , Me.Caption)
    
    If InStr(mstrPrivs, ";所有操作员;") = 0 Then
        mstrOperator = UserInfo.姓名
    Else
        mstrOperator = ""
    End If
    
    mstr票据长度 = zlDatabase.GetPara(20, glngSys, , "7|7|7|7|7")
    If mbln药店 = False Then
        If zlStr.IsHavePrivs(mstrPrivs, "收费收据") Then
            strTmp = strTmp & "|" & "收费收据,1"
        End If
        If zlStr.IsHavePrivs(mstrPrivs, "预交收据") And _
            (zlStr.IsHavePrivs(mstrPrivs, "预交门诊票据") _
                Or zlStr.IsHavePrivs(mstrPrivs, "预交住院票据")) Then
            strTmp = strTmp & "|" & "预交收据,2"
        End If
        If zlStr.IsHavePrivs(mstrPrivs, "结帐收据") Then
            strTmp = strTmp & "|" & "结帐收据,3"
        End If
        If zlStr.IsHavePrivs(mstrPrivs, "挂号收据") Then
            strTmp = strTmp & "|" & "挂号收据,4"
        End If
        If zlStr.IsHavePrivs(mstrPrivs, "医疗卡") Then
            strTmp = strTmp & "|" & "医疗卡,5"
        End If
        If zlStr.IsHavePrivs(mstrPrivs, "消费卡") Then
            strTmp = strTmp & "|" & "消费卡,6"
        End If
        If strTmp = "" Then
            MsgBox "你没有操作任何票据的权限!", vbInformation, App.ProductName
            Exit Function
        Else
            strTmp = Mid(strTmp, 2)
        End If
        
        arrTemp1 = Split(strTmp, "|")
        For i = 0 To UBound(arrTemp1)
            arrTemp2 = Split(arrTemp1(i), ",")
            Set objListItem = lvwMain.ListItems.Add(, "C" & arrTemp2(1), arrTemp2(0), "C" & arrTemp2(1))
            
            GetRegInFor g私有模块, Me.Name, "C" & arrTemp2(1), strKeyValue
            objListItem.Tag = strKeyValue
        Next
    Else
        lvwMain.ListItems.Add , "C1", "收费收据", "C1"
        lvwMain.ListItems.Add , "C5", "会员卡", "C7"
    End If
    lvwMain.ListItems(1).Selected = True
    
    '初始化表格
    arrTemp1 = Array("使用人", "使用", "使用", "使用", "使用", "收回", "收回", "收回")
    arrTemp2 = Array("使用人", "正常", "重打", "报损", "总数", "作废", "重打", "总数")
    arrTemp3 = Array(1000, 800, 800, 800, 800, 800, 800, 800)
    With msh汇总
        .Cols = 8
        .MergeCol(0) = True
        .MergeRow(0) = True
        .ColAlignment(0) = 1
        For i = 0 To .Cols - 1
            .TextMatrix(0, i) = arrTemp1(i)
            .TextMatrix(1, i) = arrTemp2(i)
            .ColWidth(i) = arrTemp3(i)
        Next                              '初始化缴款记录表
        .AllowBigSelection = True
        .FillStyle = flexFillRepeat
        .Row = 0: .Col = 0: .RowSel = 1: .ColSel = .Cols - 1
        .CellAlignment = 4
        .FillStyle = flexFillSingle
        .AllowBigSelection = False
    End With
    
    InitFace = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Resize()
    Dim sngTop As Single, sngBottom As Single
    
    On Error Resume Next
    sngTop = IIf(cbrThis.Visible, cbrThis.Top + cbrThis.Height, 0)
    sngBottom = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0)
    
    lvwMain.Top = sngTop
    lvwMain.Height = IIf(sngBottom - lvwMain.Top > 0, sngBottom - lvwMain.Top, 0)
    lvwMain.Left = 0
    
    picSplit.Top = sngTop
    picSplit.Height = IIf(sngBottom - picSplit.Top > 0, sngBottom - picSplit.Top, 0)
    picSplit.Left = lvwMain.Left + lvwMain.Width
    
    lblUp.Top = sngTop
    lblUp.Left = picSplit.Left + picSplit.Width
    If Me.ScaleWidth - lblUp.Left > 0 Then lblUp.Width = Me.ScaleWidth - lblUp.Left
    
    lvw领用_S.Left = lblUp.Left
    lvw领用_S.Top = lblUp.Top + lblUp.Height
    lvw领用_S.Width = lblUp.Width
    
    picH.Left = lblUp.Left
    picH.Top = lvw领用_S.Top + lvw领用_S.Height
    picH.Width = lblUp.Width
    
    lblDown.Left = lblUp.Left
    lblDown.Top = picH.Top + picH.Height
    lblDown.Width = lblUp.Width
    
    msh汇总.Left = lblUp.Left
    msh汇总.Top = lblDown.Top + lblDown.Height
    msh汇总.Width = lblUp.Width
    msh汇总.Height = sngBottom - msh汇总.Top
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    
    mstrKey = ""
    mstr票据 = ""
    mblnItem = False
    zlDatabase.SetPara "显示所有领用记录", IIf(mnuViewAll.Checked, 1, 0), glngSys, mlngModul, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
    zlDatabase.SetPara "查看核对信息", IIf(mnuViewCheck.Checked, 1, 0), glngSys, mlngModul, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
    
    SaveWinState Me, App.ProductName
    For i = 1 To lvwMain.ListItems.Count
        SaveRegInFor g私有模块, Me.Name, lvwMain.ListItems(i).Key, lvwMain.ListItems(i).Tag
    Next
    
    If Not gobjBillPrint Is Nothing Then
        Call gobjBillPrint.zlTerminate
        Set gobjBillPrint = Nothing
    End If
End Sub

Private Sub lvwMain_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    lvwMain.Drag 0
End Sub

Private Sub lvwMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If mstr票据 = Item.Key Then Exit Sub

    Call LoadCombox
    mstr票据 = Item.Key
    
    '调整列标题显示名称
    If CurrentIsBill(Val(Mid(Item.Key, 2))) Then
        lvw领用_S.ColumnHeaders(1).Text = "开始号码"
        lvw领用_S.ColumnHeaders(2).Text = "终止号码"
        lvw领用_S.ColumnHeaders(5).Text = "当前号码"
    Else
        lvw领用_S.ColumnHeaders(1).Text = "开始卡号"
        lvw领用_S.ColumnHeaders(2).Text = "终止卡号"
        lvw领用_S.ColumnHeaders(5).Text = "当前卡号"
    End If
    
    Call Fill记录
End Sub

Private Sub lvwMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If lvwMain.HitTest(x, y) Is Nothing Then Exit Sub
        
        lvwMain.Drag 1
    End If
End Sub

Private Sub mnuAddDetail_Click()
    Call mnuViewDetail_Click
End Sub

Private Sub mnuBillCheck_Click(Index As Integer)
    Dim lng领用ID As Long, str前缀 As String, blnChecked As Boolean
    Dim lng票种 As gBillType, strSQL As String
    
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    
    lng票种 = Val(Mid(lvwMain.SelectedItem.Key, 2))
    If lvw领用_S.SelectedItem Is Nothing And Index = 1 Then
        Call frmBillUses.ShowMe(Me, mstrPrivs, 1, True, mblnNOMoved, lng票种, 0, "")
        Exit Sub
    End If
    
    If lvw领用_S.SelectedItem Is Nothing Then Exit Sub
    lng领用ID = Val(Mid(lvw领用_S.SelectedItem.Key, 2))
    If Index = 0 Then
        blnChecked = (lvw领用_S.SelectedItem.SubItems(GetItemCOL("核对人")) <> "")
        If blnChecked Then
            If MsgBox("你确认要取消该领用单的核对记录吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Sub
            
            On Error GoTo errHandle
            If lng票种 = gBillType.消费卡 Then
                ' Zl_消费卡领用记录_Check
                strSQL = " Zl_消费卡领用记录_Check("
                '  Id_In       消费卡领用记录.Id%Type,
                strSQL = strSQL & "" & lng领用ID & ","
                '  核对结果_In 消费卡领用记录.核对结果%Type,
                strSQL = strSQL & "" & "NULL" & ","
                '  核对人_In   消费卡领用记录.核对人%Type,
                strSQL = strSQL & "" & "NULL" & ","
                '  备注_In     消费卡领用记录.备注%Type,
                strSQL = strSQL & "" & "NULL" & ","
                '  核对模式_In 消费卡领用记录.核对模式%Type
                strSQL = strSQL & "" & "NULL" & ")"
            Else
                'Zl_票据领用记录_Check
                strSQL = "Zl_票据领用记录_Check("
                '  Id_In       In 票据领用记录.Id%Type,
                strSQL = strSQL & "" & lng领用ID & ","
                '  核对结果_In In 票据领用记录.核对结果%Type,
                strSQL = strSQL & "" & "NULL" & ","
                '  核对人_In   In 票据领用记录.核对人%Type,
                strSQL = strSQL & "" & "NULL" & ","
                '  备注_In     In 票据领用记录.备注%Type,
                strSQL = strSQL & "" & "NULL" & ","
                '  核对模式_In In 票据领用记录.核对模式%Type
                strSQL = strSQL & "" & "NULL" & ")"
            End If
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        Else
            If frmBillEdit.ShowMe(Me, 1, mlngModul, mstrPrivs, lng领用ID, , lng票种) = False Then Exit Sub
        End If
        Call Fill记录
        Call SetMenu
    Else
        str前缀 = lvw领用_S.SelectedItem.SubItems(GetItemCOL("前缀文本"))
        Call frmBillUses.ShowMe(Me, mstrPrivs, 1, True, mblnNOMoved, lng票种, lng领用ID, str前缀)
            
        Call Fill记录
        If mnuViewCheck.Checked Then Fill汇总
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub

Private Sub mnuBillDelete_Click()
    On Error GoTo errHandle
    Dim intIndex As Long
    Dim lng票种 As gBillType
    
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    
    lng票种 = Val(Mid(lvwMain.SelectedItem.Key, 2))
    If MsgBox("你确认要删除开始" & _
        IIf(lng票种 = gBillType.就诊卡 Or lng票种 = gBillType.消费卡, "卡号", "号码") & _
        "为“" & lvw领用_S.SelectedItem.Text & "”的" & lvwMain.SelectedItem.Text & "领用记录吗？", _
        vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    If mblnNOMoved Then
        MsgBox "当前选择的领用记录在后备数据表中!" & vbCrLf _
            & "请与系统管理员联系,转入到在线数据表再操作!", vbInformation, gstrSysName
        Exit Sub
    End If
    If zlIsModify(Val(Mid(lvw领用_S.SelectedItem.Key, 2))) = False Then Exit Sub
    
    Me.MousePointer = 11
    If lng票种 = gBillType.消费卡 Then
        gstrSQL = "Zl_消费卡领用记录_Delete(" & Mid(lvw领用_S.SelectedItem.Key, 2) & ")"
    Else
        gstrSQL = "zl_票据领用记录_delete(" & Mid(lvw领用_S.SelectedItem.Key, 2) & ")"
    End If
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)

    Me.MousePointer = 0
    
    With lvw领用_S
        intIndex = .SelectedItem.Index
        .ListItems.Remove .SelectedItem.Key
        If .ListItems.Count > 0 Then
            intIndex = IIf(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
            .ListItems(intIndex).Selected = True
            .ListItems(intIndex).EnsureVisible
        End If
    End With
    Call Fill汇总
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Me.MousePointer = 0
End Sub

Private Sub mnuBillGet_Click()
    Dim int票种 As gBillType, str类别 As String
    
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    
    int票种 = Val(Mid(lvwMain.SelectedItem.Key, 2))
    Select Case int票种
    Case gBillType.收费收据, gBillType.结帐收据
        str类别 = Trim(cbo类别.Text)
    Case gBillType.预交收据, gBillType.就诊卡, gBillType.消费卡
        If cbo类别.ListIndex < 0 Then Exit Sub
        str类别 = cbo类别.ItemData(cbo类别.ListIndex)
    End Select
    If frmBillEdit.ShowMe(Me, 0, mlngModul, mstrPrivs, 0, str类别, int票种) = False Then Exit Sub
    
    Call Fill记录
    Call Fill汇总
End Sub

Private Sub mnuBillModify_Click()
    Dim lngLen As Long
    Dim int票种 As gBillType

    If lvw领用_S.SelectedItem Is Nothing Or mnuBill.Visible = False Then Exit Sub
    
    int票种 = Val(Mid(lvwMain.SelectedItem.Key, 2))
    '102181:李南春,2016/11/10,医疗卡票据长度
    If CurrentIsBill(int票种) = True Then
        lngLen = Val(Split(mstr票据长度, "|")(Mid(lvwMain.SelectedItem.Key, 2) - 1))
        If msh汇总.Rows > 3 And Len(lvw领用_S.SelectedItem.Text) <> lngLen Then
            MsgBox lvwMain.SelectedItem.Text & "的号码规定长度应该是" & lngLen & "位。" & _
                vbCrLf & "当条记录的号码长度不符规定，而它又已使用，故不能修改。", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    If mblnNOMoved Then
        MsgBox "当前选择的领用记录在后备数据表中!" & vbCrLf _
            & "请与系统管理员联系,转入到在线数据表再操作!", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If zlIsModify(Val(Mid(lvw领用_S.SelectedItem.Key, 2))) = False Then Exit Sub
    
    
    If frmBillEdit.ShowMe(Me, 0, mlngModul, mstrPrivs, Mid(lvw领用_S.SelectedItem.Key, 2), , int票种) = False Then Exit Sub
    
    Call Fill记录
    Call Fill汇总
End Sub

Public Function zlIsModify(ByVal lngID As Long, Optional blnMsg As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是否允许修改他人的票据
    '入参:lngID-领用ID
    '     blnMsg-是否提示信息
    '出参:
    '返回:允许修改,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-02-01 10:49:54
    '问题:27372
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    '检查当前是否允许修改他们的单据
    If zlStr.IsHavePrivs(mstrPrivs, "允许操作他人登记票据") Then
       zlIsModify = True: Exit Function
    End If
    '因为在读取时，已经有检查,现不用再判断.这段代码以后可能存在改动。所以保留
    zlIsModify = True: Exit Function
    '检查是否为本身单据
    gstrSQL = "Select ID From 票据领用记录 where id=[1] and 登记人=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngID, UserInfo.姓名)
    If rsTemp.EOF And blnMsg Then
        ShowMsgbox "注意:" & vbCrLf & "    你不能操作其他人登记的票据!"
    End If
    zlIsModify = rsTemp.EOF = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Sub mnuBillCancel_Click()
    If lvw领用_S.SelectedItem Is Nothing Then Exit Sub
    If mblnNOMoved Then
        MsgBox "当前选择的领用记录在后备数据表中!" & vbCrLf _
            & "请与系统管理员联系,转入到在线数据表再操作!", vbInformation, gstrSysName
        Exit Sub
    End If
    If zlIsModify(Val(Mid(lvw领用_S.SelectedItem.Key, 2))) = False Then Exit Sub
    If frmBillDiscard.编辑票据报损(Me, mstrPrivs, _
        Val(Mid(lvwMain.SelectedItem.Key, 2)), Val(Mid(lvw领用_S.SelectedItem.Key, 2))) Then
        Call Fill记录
    End If
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

Private Sub mnuReportItem_Click(Index As Integer)
    Dim str领用人 As String, str领用ID As String
    
    If Not lvw领用_S.SelectedItem Is Nothing Then
        str领用ID = Mid(lvw领用_S.SelectedItem.Key, 2)
        str领用人 = lvw领用_S.SelectedItem.SubItems(GetItemCOL("领用人"))
    End If
    Call ReportOpen(gcnOracle, _
        Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
        "票种=" & Val(Mid(lvwMain.SelectedItem.Key, 2)), "领用人=" & str领用人, "领用ID=" & str领用ID)
End Sub

Private Sub mnuViewCheck_Click()
    mnuViewCheck.Checked = Not mnuViewCheck.Checked
    Call Fill汇总
End Sub

Private Sub mnuViewDetail_Click()
    Dim lng领用ID As Long, lng原因 As Long, lng性质 As Long
    Dim strCondition As String, str提示 As String, str使用人 As String, str前缀 As String
    Dim blnOne As Boolean, lng票种 As gBillType
    
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    
    With msh汇总
        If .Rows = 3 Or .TextMatrix(.Row, .Col) = " " Then
            Exit Sub
        End If
        
        blnOne = .Row < 2 Or .Row = .Rows - 1
        Select Case .Col
            Case 0
                str提示 = "全部明细清单"
                strCondition = ""
            Case 1
                str提示 = "正常使用明细清单"
                strCondition = " and 原因=[2]": lng原因 = 1
            Case 2
                str提示 = "重打使用明细清单"
                strCondition = " and 原因=[2]": lng原因 = 3
            Case 3
                str提示 = "报损使用明细清单"
                strCondition = " and 原因=[2]": lng原因 = 5
            Case 4
                str提示 = "全部使用明细清单"
                strCondition = " and 性质=[3]": lng性质 = 1
            Case 5
                str提示 = "作废收回明细清单"
                strCondition = " and 原因=[2]": lng原因 = 2
            Case 6
                str提示 = "重打收回明细清单"
                strCondition = " and 原因=[2]": lng原因 = 4
            Case 7
                str提示 = "全部收回明细清单"
                strCondition = " and 性质=[3]": lng性质 = 2
        End Select
        If blnOne = False Then
            str使用人 = .TextMatrix(.Row, 0)
            str提示 = str使用人 & "的" & str提示
            strCondition = strCondition & " and 使用人||''=[4]"
        Else
            str提示 = "所有人的" & str提示
        End If
    End With
    lng票种 = Val(Mid(lvwMain.SelectedItem.Key, 2))
    lng领用ID = Val(Mid(lvw领用_S.SelectedItem.Key, 2))
    str前缀 = lvw领用_S.SelectedItem.SubItems(GetItemCOL("前缀文本"))
    
    Call frmBillUses.ShowMe(Me, mstrPrivs, 0, mnuViewCheck.Checked, mblnNOMoved, _
        lng票种, lng领用ID, str前缀, strCondition, lng原因, lng性质, str使用人, str提示)
End Sub

Private Function GetItemCOL(strColName As String)
'功能:根据名称返回listsubitems列表的列号
    Dim lngCol As Long
    
    For lngCol = 2 To lvw领用_S.ColumnHeaders.Count
        'ColumnHeaders的第一列是listitem的开始号码,listsubitems的第一列是从ColumnHeaders的第2列开始的
        If lvw领用_S.ColumnHeaders(lngCol).Text = strColName Then
            GetItemCOL = lngCol - 1
            Exit For
        End If
    Next
End Function

Private Sub mnuViewFlash_Click()
    Call Fill记录
End Sub

Private Sub mnuAddAll_Click()
    mnuViewAll_Click
End Sub

Private Sub mnuAddHave_Click()
    mnuviewHave_Click
End Sub

Private Sub lvw领用_S_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintColumn = ColumnHeader.Index - 1 Then '仍是刚才那列
        lvw领用_S.SortOrder = IIf(lvw领用_S.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvw领用_S.SortKey = mintColumn
        lvw领用_S.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvw领用_S_DblClick()
    If mblnItem = True And mnuBillModify.Enabled And mnuBillModify.Visible Then
        Call mnuBillModify_Click
    End If
End Sub

Public Sub lvw领用_S_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim rsTmp As Recordset
    Dim int票种 As gBillType
    
    mblnItem = True
    
    On Error GoTo errHandle
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    If mstrKey = Item.Key Then Exit Sub
    mstrKey = Item.Key
    
    int票种 = Val(Mid(lvwMain.SelectedItem.Key, 2))
    '当前领用记录是否在后备表中
    mblnNOMoved = False
    If mblnDateMoved Then
        If int票种 = gBillType.消费卡 Then
            gstrSQL = "Select id From H消费卡领用记录 Where id=[1]"
        Else
            gstrSQL = "Select id From H票据领用记录 Where id=[1]"
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Mid(lvw领用_S.SelectedItem.Key, 2))
        If rsTmp.RecordCount > 0 Then mblnNOMoved = True
    End If
    
    Call Fill汇总
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub lvw领用_S_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If mnuBillModify.Enabled And mnuBillModify.Visible Then
            Call mnuBillModify_Click
        End If
    End If
End Sub
 
 Sub lvw领用_S_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    If Button = 2 Then
        mnuAddAll.Checked = mnuViewAll.Checked
        mnuAddHave.Checked = mnuViewHave.Checked
        For i = 0 To 3
            mnuShortIcon(i).Checked = mnuViewIcon(i).Checked
        Next
        PopupMenu mnuAdd, vbPopupMenuRightButton
    End If
End Sub

Private Sub mnuViewFilter_Click()
    Dim lngKind As Long
    
    lngKind = Val(Mid(lvwMain.SelectedItem.Key, 2))
    If frmTimeSet.ShowMe(Me, 1, lngKind, mlngModul, mstrPrivs, _
        mdatBegin, mdatEnd, mstrOperator, mblnDateMoved) Then
        Call Fill记录
    End If
End Sub

Private Sub mnuViewSelect_Click()
    If zlControl.LvwSelectColumns(lvw领用_S, mstrLvw) = True Then
        '列有变化就要重新刷新
        Fill记录
    End If
End Sub

Private Sub msh汇总_DblClick()
    If mnuViewDetail.Enabled = True And mnuViewDetail.Visible = True Then
        Call mnuViewDetail_Click
    End If
End Sub

Private Sub msh汇总_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If msh汇总.Rows = 3 Then Exit Sub
    msh汇总.SetFocus
    If Button = 2 Then PopupMenu mnuAdd2, vbPopupMenuRightButton
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim i As Integer
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
    Next
    mnuViewIcon(ButtonMenu.Index - 1).Checked = True
    lvw领用_S.View = ButtonMenu.Index - 1
End Sub

Private Sub mnuViewIcon_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
    Next
    mnuViewIcon(Index).Checked = True
    lvw领用_S.View = Index
End Sub

Private Sub mnuShortIcon_Click(Index As Integer)
    mnuViewIcon_Click Index
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

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    cbrThis.Visible = mnuViewToolButton.Checked
    cbrThis.Bands("only").minHeight = Toolbar1.Height
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
    cbrThis.Bands("only").minHeight = Toolbar1.Height
    Form_Resize
End Sub

Private Sub mnuViewAll_Click()
    mnuViewAll.Checked = Not mnuViewAll.Checked
    mnuViewHave.Checked = Not mnuViewAll.Checked
    Fill记录
End Sub

Private Sub mnuviewHave_Click()
    mnuViewHave.Checked = Not mnuViewHave.Checked
    mnuViewAll.Checked = Not mnuViewHave.Checked
    Fill记录
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

Private Sub picH_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then msngStart = y
End Sub

Private Sub picH_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sngTemp As Single
    Dim sngBottom As Single
    
    If Button = 1 Then
        sngTemp = picH.Top + y - msngStart
        sngBottom = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0)
        If sngTemp - lvw领用_S.Top > 1500 And sngBottom - sngTemp > 2000 Then
            picH.Top = sngTemp
            lvw领用_S.Height = sngTemp - lvw领用_S.Top
            
            lblDown.Top = sngTemp + 45
            msh汇总.Top = lblDown.Top + lblDown.Height
            msh汇总.Height = sngBottom - msh汇总.Top
        End If
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "New"
            mnuBillGet_Click
        Case "Modify"
            mnuBillModify_Click
        Case "Delete"
            mnuBillDelete_Click
        Case "Cancel"
            mnuBillCancel_Click
        Case "Check"
            Call mnuBillCheck_Click(1)
        Case "Quit"
            mnuFileExit_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Preview"
            mnuFilePreview_Click
        Case "Filter"
            mnuViewFilter_Click
        Case "Help"
            mnuHelpTopic_Click
        Case "View"
            mnuViewIcon(lvw领用_S.View).Checked = False
            If lvw领用_S.View = 3 Then
                mnuViewIcon(0).Checked = True
                lvw领用_S.View = 0
            Else
                mnuViewIcon(lvw领用_S.View + 1).Checked = True
                lvw领用_S.View = lvw领用_S.View + 1
            End If
    End Select
End Sub

Private Sub Toolbar1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool, vbPopupMenuRightButton
    End If
End Sub

Private Sub Fill记录()
'功能:装入所有收费员到lvw领用_S_S
    Dim rsTmp As ADODB.Recordset
    Dim lst As ListItem, int票种 As gBillType, strWhere As String
    Dim strKey As String, str类别 As String, str使用类别 As String
    Dim lngCol  As Long, strColName As String
    Dim varValue As Variant
        
    If Not lvw领用_S.SelectedItem Is Nothing Then
        strKey = lvw领用_S.SelectedItem.Key '保留原有键值
    End If
    'mstrOperator:没有所有操作员权限时,只则显示本人领用票据或共享票据
    '问题:35834
    int票种 = Val(Mid(lvwMain.SelectedItem.Key, 2))
    strWhere = ""
    Select Case int票种
    Case gBillType.收费收据, gBillType.结帐收据
        str类别 = cbo类别.Text
        If str类别 = " " Then
            strWhere = strWhere & " And nvl(使用类别,'LXH')=[6]"
            str类别 = "LXH"
        ElseIf str类别 <> "所有类别" Then
            strWhere = strWhere & " And nvl(使用类别,'LXH')=[6]"
        End If
    Case gBillType.预交收据
        If cbo类别.ListIndex < 0 Then Exit Sub
        str类别 = cbo类别.ItemData(cbo类别.ListIndex)
        '58071
        If Val(str类别) <> -1 Then
            str类别 = Val(str类别)
            strWhere = strWhere & " And nvl(使用类别,'0')=[6]"
        End If
    Case gBillType.就诊卡
        If cbo类别.ListIndex < 0 Then Exit Sub
        str类别 = cbo类别.ItemData(cbo类别.ListIndex)
        If Val(str类别) <> 0 Then
            str类别 = Val(str类别)
            strWhere = strWhere & " And nvl(使用类别,'0')=[6]"
        End If
    Case Else
    End Select

    str使用类别 = "A.使用类别,"
    If int票种 = gBillType.预交收据 Then
        '58071
        str使用类别 = "decode(nvl(A.使用类别,'0'),'0','','1','门诊','住院') as 使用类别,"
    ElseIf int票种 = gBillType.就诊卡 Then
        str使用类别 = "nvl(M.名称,'就诊卡') As 使用类别,"
    End If
    
    If int票种 = gBillType.消费卡 Then
        If cbo类别.ListIndex < 0 Then Exit Sub
        str类别 = cbo类别.ItemData(cbo类别.ListIndex)
        
        gstrSQL = _
            "Select A.ID, nvl(M.名称,'消费卡') As 使用类别,A.领用人,A.前缀文本," & vbNewLine & _
            "       A.开始卡号 As 开始号码,A.终止卡号 As 终止号码," & vbNewLine & _
            "       Decode(A.使用方式,1,'自用','共用') as 使用方式," & vbNewLine & _
            "       to_Char(A.登记时间,'YYYY-MM-DD') as 登记时间," & vbNewLine & _
            "       A.登记人,A.当前卡号 As 当前号码,A.剩余数量,A.批次,A.核对人," & vbNewLine & _
            "       A.签字人,to_Char(A.签字时间,'YYYY-MM-DD HH24:mi:ss') as 签字时间" & vbNewLine & _
            "From " & IIf(mblnDateMoved, zlGetFullFieldsTable("消费卡领用记录"), "消费卡领用记录 A") & _
                " ,人员表 B,消费卡类别目录 M" & vbNewLine & _
            "Where a.接口编号 = m.编号(+) And a.接口编号=[6]" & vbNewLine & _
                    IIf(mnuViewHave.Checked, " And A.剩余数量<>0", "") & vbNewLine & _
            "      And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & vbNewLine & _
            "      And A.领用人=B.姓名 And A.登记时间 Between [2] And [3]" & vbNewLine & _
                    IIf(mstrOperator = "", "", " And (A.领用人=[4] Or nvl(A.使用方式,0)=2)")
    Else
        gstrSQL = _
            "Select A.ID," & str使用类别 & _
            "       A.领用人,A.前缀文本,A.开始号码,A.终止号码," & vbNewLine & _
            "       Decode(A.使用方式,1,'自用','共用') as 使用方式," & vbNewLine & _
            "       to_Char(A.登记时间,'YYYY-MM-DD') as 登记时间," & vbNewLine & _
            "       A.登记人,A.当前号码,A.剩余数量,A.批次,A.核对人," & vbNewLine & _
            "       A.签字人,to_Char(A.签字时间,'YYYY-MM-DD HH24:mi:ss') as 签字时间" & vbNewLine & _
            "From " & IIf(mblnDateMoved, zlGetFullFieldsTable("票据领用记录"), "票据领用记录 A") & _
                " ,人员表 B" & IIf(int票种 = gBillType.就诊卡, ",医疗卡类别 M", "") & vbNewLine & _
            "Where A.票种=[1] " & IIf(mnuViewHave.Checked, " And A.剩余数量<>0", "") & vbNewLine & _
            "      And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & strWhere & vbNewLine & _
            "      And A.领用人=B.姓名 And A.登记时间 Between [2] And [3]" & vbNewLine & _
                    IIf(mstrOperator = "", "", " And (A.领用人=[4] Or nvl(A.使用方式,0)=2)") & vbNewLine & _
                    IIf(int票种 = gBillType.就诊卡, " And to_number(nvl(A.使用类别,'0'))=M.ID(+)", "")
    End If

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
        int票种, mdatBegin, DateAdd("s", -1, DateAdd("d", 1, mdatEnd)), _
        mstrOperator, UserInfo.姓名, str类别)
    
    LockWindowUpdate lvw领用_S.hWnd
    With lvw领用_S
        .ListItems.Clear
        Do Until rsTmp.EOF
            Set lst = .ListItems.Add(, "C" & rsTmp("ID"), rsTmp("开始号码"), "Item", "Item")
            
            '根据ListView的列名从数据库取数
            For lngCol = 2 To lvw领用_S.ColumnHeaders.Count
                strColName = lvw领用_S.ColumnHeaders(lngCol).Text
                If strColName = "开始卡号" Then strColName = "开始号码"
                If strColName = "终止卡号" Then strColName = "终止号码"
                If strColName = "当前卡号" Then strColName = "当前号码"
                varValue = rsTmp(strColName).Value
                lst.SubItems(lngCol - 1) = IIf(IsNull(varValue), "", varValue)
            Next
            rsTmp.MoveNext
        Loop
        If .ListItems.Count > 0 Then
            Dim Item As ListItem
            On Error Resume Next
            Set Item = .ListItems(strKey)
            If Err <> 0 Then
                Set Item = .ListItems(1)
                Item.Selected = True
                Item.EnsureVisible
                lvw领用_S_ItemClick Item
            Else
                Err.Clear
                Item.Selected = True
                Item.EnsureVisible
                mstrKey = "" '清空状态变量,刷新汇总列表
                lvw领用_S_ItemClick Item
            End If
        Else
            Call Fill汇总
        End If
    End With
    LockWindowUpdate 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    LockWindowUpdate 0
End Sub

Public Sub Fill汇总()
'功能:对票据使用情况汇总
    Dim rsTmp As ADODB.Recordset
    Dim lngCol As Long
    Dim lngSum(1 To 5) As Long
    Dim lngCSum(1 To 5) As Long, int票种 As Integer
    On Error GoTo errH
    
    If lvw领用_S.SelectedItem Is Nothing Then
        msh汇总.Rows = 3
        For lngCol = 0 To msh汇总.Cols - 1
            msh汇总.TextMatrix(2, lngCol) = ""
        Next
        msh汇总.Row = 2
        Call SetMenu
        Exit Sub
    End If
    
    int票种 = Val(Mid(lvwMain.SelectedItem.Key, 2))
    If int票种 = gBillType.消费卡 Then
        If mnuViewCheck.Checked Then
            gstrSQL = _
                "Select 使用人, Sum(Decode(原因, 1, 1, 6, 1, 0)) As 正常," & vbNewLine & _
                "       Sum(Decode(原因, 2, 1, 0)) As 作废," & vbNewLine & _
                "       Sum(Decode(原因, 3, 1, 0)) As 重打," & vbNewLine & _
                "       Sum(Decode(原因, 4, 1, 0)) As 重打收回," & vbNewLine & _
                "       Sum(Decode(原因, 5, 1, 0)) As 报损," & vbNewLine & _
                "       Sum(Decode(核对结果, 1, 1, 0)) As C正常," & vbNewLine & _
                "       Sum(Decode(核对结果, 2, 1, 0)) As C作废," & vbNewLine & _
                "       Sum(Decode(核对结果, 3, 1, 0)) As C重打," & vbNewLine & _
                "       Sum(Decode(核对结果, 4, 1, 0)) As C重打收回," & vbNewLine & _
                "       Sum(Decode(核对结果, 5, 1, 0)) As C报损" & vbNewLine & _
                "From " & IIf(mblnDateMoved, zlGetFullFieldsTable("消费卡使用记录"), "消费卡使用记录") & vbNewLine & _
                "Where 领用id = [1]" & vbNewLine & _
                "Group By 使用人"
        Else
            gstrSQL = _
                "Select 使用人,Sum(Decode(原因, 1, 1, 6, 1, 0)) As 正常," & vbNewLine & _
                "       Sum(Decode(原因, 2, 1, 0)) As 作废, " & vbNewLine & _
                "       Sum(decode(原因,3,1,0)) As 重打," & vbNewLine & _
                "       Sum(decode(原因,4,1,0)) As 重打收回," & vbNewLine & _
                "       Sum(decode(原因,5,1,0)) As 报损 " & vbNewLine & _
                "From " & IIf(mblnDateMoved, zlGetFullFieldsTable("消费卡使用记录"), "消费卡使用记录") & vbNewLine & _
                "Where 领用ID = [1]" & vbNewLine & _
                "Group By 使用人"
        End If
    Else
        If mnuViewCheck.Checked Then
            gstrSQL = _
                "Select 使用人, Sum(Decode(原因, 1, 1, 6, 1, 0)) As 正常," & vbNewLine & _
                "       Sum(Decode(原因, 2, 1, 0)) As 作废," & vbNewLine & _
                "       Sum(Decode(原因, 3, 1, 0)) As 重打," & vbNewLine & _
                "       Sum(Decode(原因, 4, 1, 0)) As 重打收回," & vbNewLine & _
                "       Sum(Decode(原因, 5, 1, 0)) As 报损," & vbNewLine & _
                "       Sum(Decode(核对结果, 1, 1, 0)) As C正常," & vbNewLine & _
                "       Sum(Decode(核对结果, 2, 1, 0)) As C作废," & vbNewLine & _
                "       Sum(Decode(核对结果, 3, 1, 0)) As C重打," & vbNewLine & _
                "       Sum(Decode(核对结果, 4, 1, 0)) As C重打收回," & vbNewLine & _
                "       Sum(Decode(核对结果, 5, 1, 0)) As C报损" & vbNewLine & _
                "From " & IIf(mblnDateMoved, zlGetFullFieldsTable("票据使用明细"), "票据使用明细") & vbNewLine & _
                "Where 领用id = [1]" & vbNewLine & _
                "Group By 使用人"
        Else
            gstrSQL = _
                "Select 使用人,Sum(Decode(原因, 1, 1, 6, 1, 0)) As 正常," & vbNewLine & _
                "       Sum(Decode(原因, 2, 1, 0)) As 作废, " & vbNewLine & _
                "       Sum(decode(原因,3,1,0)) As 重打," & vbNewLine & _
                "       Sum(decode(原因,4,1,0)) As 重打收回," & vbNewLine & _
                "       Sum(decode(原因,5,1,0)) As 报损 " & vbNewLine & _
                "From " & IIf(mblnDateMoved, zlGetFullFieldsTable("票据使用明细"), "票据使用明细") & vbNewLine & _
                "Where 领用ID = [1]" & vbNewLine & _
                "Group By 使用人"
        End If
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Mid(lvw领用_S.SelectedItem.Key, 2))
    With msh汇总
        .Redraw = False
        If rsTmp.EOF Then
            .Rows = 3
            For lngCol = 0 To .Cols - 1
                .TextMatrix(2, lngCol) = ""
            Next
        Else
            .Rows = rsTmp.RecordCount + 3 '两行是表头，另一行是合计
            lngCol = 2
            Do Until rsTmp.EOF
                .TextMatrix(lngCol, 0) = rsTmp!使用人
                If mnuViewCheck.Checked Then
                    .TextMatrix(lngCol, 1) = Format(rsTmp!正常, "#########;-#########;" & IIf(rsTmp!正常 = 0 And rsTmp!C正常 <> 0, "0", "") & "; ") & Format(rsTmp!C正常, "\/#########;-#########; ; "): lngSum(1) = lngSum(1) + rsTmp!正常: lngCSum(1) = lngCSum(1) + rsTmp!C正常
                    .TextMatrix(lngCol, 2) = Format(rsTmp!重打, "#########;-#########;" & IIf(rsTmp!重打 = 0 And rsTmp!C重打 <> 0, "0", "") & "; ") & Format(rsTmp!C重打, "\/#########;-#########; ; "): lngSum(2) = lngSum(2) + rsTmp!重打: lngCSum(2) = lngCSum(2) + rsTmp!C重打
                    .TextMatrix(lngCol, 3) = Format(rsTmp!报损, "#########;-#########;" & IIf(rsTmp!报损 = 0 And rsTmp!C报损 <> 0, "0", "") & "; ") & Format(rsTmp!C报损, "\/#########;-#########; ; "): lngSum(3) = lngSum(3) + rsTmp!报损: lngCSum(3) = lngCSum(3) + rsTmp!C报损
                    .TextMatrix(lngCol, 4) = Format(rsTmp!正常 + rsTmp!重打 + rsTmp!报损, "#########;-#########;" & IIf((rsTmp!正常 + rsTmp!重打 + rsTmp!报损) = 0 And (rsTmp!C正常 + rsTmp!C重打 + rsTmp!C报损) <> 0, "0", "") & "; ") & Format(rsTmp!C正常 + rsTmp!C重打 + rsTmp!C报损, "\/#########;-#########; ; ")
                    .TextMatrix(lngCol, 5) = Format(rsTmp!作废, "#########;-#########;" & IIf(rsTmp!作废 = 0 And rsTmp!C作废 <> 0, "0", "") & "; ") & Format(rsTmp!C作废, "\/#########;-#########; ; "): lngSum(4) = lngSum(4) + rsTmp!作废: lngCSum(4) = lngCSum(4) + rsTmp!C作废
                    .TextMatrix(lngCol, 6) = Format(rsTmp!重打收回, "#########;-#########;" & IIf(rsTmp!重打收回 = 0 And rsTmp!C重打收回 <> 0, "0", "") & "; ") & Format(rsTmp!C重打收回, "\/#########;-#########; ; "): lngSum(5) = lngSum(5) + rsTmp!重打收回: lngCSum(5) = lngCSum(5) + rsTmp!C重打收回
                    .TextMatrix(lngCol, 7) = Format(rsTmp!作废 + rsTmp!重打收回, "#########;-#########;" & IIf((rsTmp!作废 + rsTmp!重打收回) = 0 And (rsTmp!C作废 + rsTmp!C重打收回) <> 0, "0", "") & "; ") & Format(rsTmp!C作废 + rsTmp!C重打收回, "\/#########;-#########; ; ")

                Else
                    .TextMatrix(lngCol, 1) = Format(rsTmp!正常, "#########;-#########; ; "): lngSum(1) = lngSum(1) + rsTmp!正常
                    .TextMatrix(lngCol, 2) = Format(rsTmp!重打, "#########;-#########; ; "): lngSum(2) = lngSum(2) + rsTmp!重打
                    .TextMatrix(lngCol, 3) = Format(rsTmp!报损, "#########;-#########; ; "): lngSum(3) = lngSum(3) + rsTmp!报损
                    .TextMatrix(lngCol, 4) = Format(rsTmp!正常 + rsTmp!重打 + rsTmp!报损, "#########;-#########; ; ")
                    .TextMatrix(lngCol, 5) = Format(rsTmp!作废, "#########;-#########; ; "): lngSum(4) = lngSum(4) + rsTmp!作废
                    .TextMatrix(lngCol, 6) = Format(rsTmp!重打收回, "#########;-#########; ; "): lngSum(5) = lngSum(5) + rsTmp!重打收回
                    .TextMatrix(lngCol, 7) = Format(rsTmp!作废 + rsTmp!重打收回, "#########;-#########; ; ")
                End If
                
                lngCol = lngCol + 1
                rsTmp.MoveNext
            Loop
            lngCol = .Rows - 1
            .TextMatrix(lngCol, 0) = "   合计"
            If mnuViewCheck.Checked Then
                .TextMatrix(lngCol, 1) = Format(lngSum(1), "#########;-#########;" & IIf(lngSum(1) = 0 And lngCSum(1) <> 0, "0", "") & "; ") & Format(lngCSum(1), "\/#########;-#########; ; ")
                .TextMatrix(lngCol, 2) = Format(lngSum(2), "#########;-#########;" & IIf(lngSum(2) = 0 And lngCSum(2) <> 0, "0", "") & "; ") & Format(lngCSum(2), "\/#########;-#########; ; ")
                .TextMatrix(lngCol, 3) = Format(lngSum(3), "#########;-#########;" & IIf(lngSum(3) = 0 And lngCSum(3) <> 0, "0", "") & "; ") & Format(lngCSum(3), "\/#########;-#########; ; ")
                .TextMatrix(lngCol, 4) = Format(lngSum(1) + lngSum(2) + lngSum(3), "#########;-#########;" & IIf((lngSum(1) + lngSum(2) + lngSum(3)) = 0 And (lngCSum(1) + lngCSum(2) + lngCSum(3)) <> 0, "0", "") & "; ") & Format(lngCSum(1) + lngCSum(2) + lngCSum(3), "\/#########;-#########; ; ")
                .TextMatrix(lngCol, 5) = Format(lngSum(4), "#########;-#########;" & IIf(lngSum(4) = 0 And lngCSum(4) <> 0, "0", "") & "; ") & Format(lngCSum(4), "\/#########;-#########; ; ")
                .TextMatrix(lngCol, 6) = Format(lngSum(5), "#########;-#########;" & IIf(lngSum(5) = 0 And lngCSum(5) <> 0, "0", "") & "; ") & Format(lngCSum(5), "\/#########;-#########; ; ")
                .TextMatrix(lngCol, 7) = Format(lngSum(5) + lngSum(4), "#########;-#########;" & IIf((lngSum(5) + lngSum(4)) = 0 And (lngCSum(5) + lngCSum(4)) <> 0, "0", "") & "; ") & Format(lngCSum(5) + lngCSum(4), "\/#########;-#########; ; ")
            
            Else
                .TextMatrix(lngCol, 1) = Format(lngSum(1), "#########;-#########; ; ")
                .TextMatrix(lngCol, 2) = Format(lngSum(2), "#########;-#########; ; ")
                .TextMatrix(lngCol, 3) = Format(lngSum(3), "#########;-#########; ; ")
                .TextMatrix(lngCol, 4) = Format(lngSum(1) + lngSum(2) + lngSum(3), "#########;-#########; ; ")
                .TextMatrix(lngCol, 5) = Format(lngSum(4), "#########;-#########; ; ")
                .TextMatrix(lngCol, 6) = Format(lngSum(5), "#########;-#########; ; ")
                .TextMatrix(lngCol, 7) = Format(lngSum(5) + lngSum(4), "#########;-#########; ; ")
            End If
        End If
        .Redraw = True
        .Row = 2
    End With
    Call SetMenu
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
    Dim objPrint As Object
    Dim objRow As New zlTabAppRow
    
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    
    If ActiveControl Is msh汇总 Then
        If msh汇总.Rows = 3 Then Exit Sub
        
        Set objPrint = New zlPrint1Grd
        objPrint.Title.Text = lvwMain.SelectedItem.Text & "使用情况"
        Set objPrint.Body = msh汇总
        objRow.Add "领用人：" & lvw领用_S.SelectedItem.SubItems(GetItemCOL("领用人"))
        If CurrentIsBill(Val(Mid(lvwMain.SelectedItem.Key, 2))) Then
            objRow.Add "号码：" & lvw领用_S.SelectedItem.Text & _
                "——" & lvw领用_S.SelectedItem.SubItems(GetItemCOL("终止号码"))
        Else
            objRow.Add "卡号：" & lvw领用_S.SelectedItem.Text & _
                "——" & lvw领用_S.SelectedItem.SubItems(GetItemCOL("终止卡号"))
        End If
        objPrint.UnderAppRows.Add objRow
        
        Set objRow = New zlTabAppRow
        objRow.Add "打印人：" & UserInfo.姓名
        objRow.Add "打印时间：" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
        objPrint.BelowAppRows.Add objRow
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
    
    Else
        Set objPrint = New zlPrintLvw
        objPrint.Title.Text = lvwMain.SelectedItem.Text & "领用记录"
        Set objPrint.Body.objData = lvw领用_S
        objPrint.UnderAppItems.Add "领用时间：" & Format(mdatBegin, "yyyy年MM月dd日") & _
            "——" & Format(mdatEnd, "yyyy年MM月dd日")
        objPrint.BelowAppItems.Add "打印人：" & UserInfo.姓名
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
    End If
End Sub

Private Sub PrivilegeCTRL()
'功能:由于有的用户权限不够,故使一些菜单项或按钮不可见
    If InStr(mstrPrivs, "增删改") = 0 _
        And InStr(mstrPrivs, "票据报损") = 0 _
        And InStr(mstrPrivs, "票据核对") = 0 Then
        mnuBill.Visible = False
        Toolbar1.Buttons("New").Visible = False
        Toolbar1.Buttons("Modify").Visible = False
        Toolbar1.Buttons("Delete").Visible = False
        Toolbar1.Buttons("Split1").Visible = False
        Toolbar1.Buttons("Cancel").Visible = False
        Toolbar1.Buttons("Check").Visible = False
        Toolbar1.Buttons("Split2").Visible = False
    ElseIf InStr(mstrPrivs, "增删改") = 0 Then
        mnuBillGet.Visible = False
        mnuBillModify.Visible = False
        mnuBillDelete.Visible = False
        mnuBillSplit.Visible = False
        Toolbar1.Buttons("New").Visible = False
        Toolbar1.Buttons("Modify").Visible = False
        Toolbar1.Buttons("Delete").Visible = False
        Toolbar1.Buttons("Split1").Visible = False
    ElseIf InStr(mstrPrivs, "票据报损") = 0 _
        And InStr(mstrPrivs, "票据核对") = 0 Then
        mnuBillSplit.Visible = False
        mnuBillCancel.Visible = False
        mnuBillCheck(0).Visible = False
        mnuBillCheck(1).Visible = False
        Toolbar1.Buttons("Split2").Visible = False
        Toolbar1.Buttons("Cancel").Visible = False
        Toolbar1.Buttons("Check").Visible = False
    ElseIf InStr(mstrPrivs, "票据报损") = 0 Then
        mnuBillCancel.Visible = False
        Toolbar1.Buttons("Cancel").Visible = False
    ElseIf InStr(mstrPrivs, "票据核对") = 0 Then
        mnuBillCheck(0).Visible = False
        mnuBillCheck(1).Visible = False
        Toolbar1.Buttons("Check").Visible = False
    End If
    mbln药店 = (glngSys \ 100 = 8)
End Sub

Private Sub SetMenu()
    Dim blnDetail As Boolean, blnModify As Boolean, blnChecked As Boolean
    Dim blnHavePrivs As Boolean  '是否有操作权限
    
    blnModify = Not (lvw领用_S.SelectedItem Is Nothing)
    
    blnHavePrivs = zlStr.IsHavePrivs(mstrPrivs, "允许操作他人登记票据")
    If blnHavePrivs = False And blnModify Then
        '需检判断登记人:
       blnHavePrivs = lvw领用_S.SelectedItem.SubItems(GetItemCOL("登记人")) = UserInfo.姓名
    End If
    
    blnDetail = (msh汇总.Rows > 3)
    
    If Not (lvw领用_S.SelectedItem Is Nothing) Then
        blnChecked = (lvw领用_S.SelectedItem.SubItems(GetItemCOL("核对人")) <> "")
    End If

    mnuBillModify.Enabled = blnModify And blnHavePrivs
    mnuBillCancel.Enabled = blnModify
    
    mnuBillCheck(0).Enabled = blnModify And Not blnDetail    '核对领用单
    If blnChecked Then
        mnuBillCheck(0).Caption = "取消核对领用单(&B)"
    Else
        mnuBillCheck(0).Caption = "核对领用单(&B)"
    End If
    Toolbar1.Buttons("Modify").Enabled = blnModify And blnHavePrivs
    Toolbar1.Buttons("Cancel").Enabled = blnModify
    Toolbar1.Buttons("Check").Enabled = blnModify And blnDetail
    
    
    mnuBillDelete.Enabled = blnModify And Not blnDetail And blnHavePrivs
    Toolbar1.Buttons("Delete").Enabled = blnModify And Not blnDetail And blnHavePrivs
    mnuViewDetail.Enabled = blnDetail
    

    mnuFilePreview.Enabled = blnModify
    mnuFilePrint.Enabled = blnModify
    mnuFileExcel.Enabled = blnModify
    Toolbar1.Buttons("Preview").Enabled = blnModify
    Toolbar1.Buttons("Print").Enabled = blnModify
    
    stbThis.Panels(2).Text = "《" & lvwMain.SelectedItem.Text & "》在" & _
        Format(mdatBegin, "yyyy年MM月dd日") & "——" & _
        Format(mdatEnd, "yyyy年MM月dd日") & "之间共有" & lvw领用_S.ListItems.Count & "条领用记录。"
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

Private Function LoadCombox() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载Combox数据
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-04-27 10:22:29
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim int票种 As gBillType, str类别 As String
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    If lvwMain.SelectedItem Is Nothing Then Exit Function
    
    int票种 = Val(Mid(lvwMain.SelectedItem.Key, 2))
    str类别 = lvwMain.SelectedItem.Tag
    
    Select Case int票种
    Case gBillType.收费收据, gBillType.结帐收据
        strSQL = "Select 编码,名称,简码,缺省标志 From 票据使用类别 "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        mblnNotClick = True
        With cbo类别
            .Clear
            
            .AddItem "所有类别"
            If str类别 = "所有类别" Then .ListIndex = .NewIndex
            
            Do While Not rsTemp.EOF
                .AddItem Nvl(rsTemp!名称)
                .ItemData(.NewIndex) = 1
                If Val(Nvl(rsTemp!缺省标志)) = 1 And .ListIndex < 0 Then .ListIndex = .NewIndex
                If str类别 = Nvl(rsTemp!名称) Then .ListIndex = .NewIndex
                rsTemp.MoveNext
            Loop
            
            .AddItem " "
            .ItemData(.NewIndex) = -1
            If str类别 = " " Then .ListIndex = .NewIndex
            
            If .ListIndex < 0 And .ListCount > 0 Then .ListIndex = 0
        End With
        cbrThis.Bands(2).Visible = True
        cbrThis.Bands(2).Caption = "使用类别"
        mblnNotClick = False
    Case gBillType.预交收据
        mblnNotClick = True
        With cbo类别
            .Clear
            If zlStr.IsHavePrivs(mstrPrivs, "预交门诊票据") _
                And zlStr.IsHavePrivs(mstrPrivs, "预交住院票据") Then
                .AddItem "所有预交"
                .ItemData(.NewIndex) = -1
                If Val(str类别) = -1 Then .ListIndex = .NewIndex
            End If
            If zlStr.IsHavePrivs(mstrPrivs, "预交门诊票据") Then
                .AddItem "门诊预交"
                .ItemData(.NewIndex) = 1
                If Val(str类别) = 1 Then .ListIndex = .NewIndex
            End If
            If zlStr.IsHavePrivs(mstrPrivs, "预交住院票据") Then
                .AddItem "住院预交"
                .ItemData(.NewIndex) = 2
                If Val(str类别) = 2 Then .ListIndex = .NewIndex
            End If
            '58071
            If zlStr.IsHavePrivs(mstrPrivs, "预交门诊票据") _
                And zlStr.IsHavePrivs(mstrPrivs, "预交住院票据") Then
                .AddItem ""
                .ItemData(.NewIndex) = 0
                If Val(str类别) = 0 Then .ListIndex = .NewIndex
            End If
            If .ListIndex < 0 And .ListCount > 0 Then .ListIndex = 0
        End With
        cbrThis.Bands(2).Visible = True
        cbrThis.Bands(2).Caption = "使用类别"
        mblnNotClick = False
    Case gBillType.就诊卡
        strSQL = _
            "Select ID, 编码, 名称, 缺省标志" & vbNewLine & _
            "From 医疗卡类别" & vbNewLine & _
            "Where Nvl(是否启用, 0) >= 1" & vbNewLine & _
            "Order By 编码"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        mblnNotClick = True
        With cbo类别
            .Clear
            Do While Not rsTemp.EOF
                .AddItem Nvl(rsTemp!编码) & "-" & Nvl(rsTemp!名称)
                .ItemData(.NewIndex) = Val(Nvl(rsTemp!ID))
                If Val(Nvl(rsTemp!缺省标志)) = 1 And .ListIndex < 0 Then .ListIndex = .NewIndex
                If Val(str类别) = Val(Nvl(rsTemp!ID)) Then .ListIndex = .NewIndex
                rsTemp.MoveNext
            Loop
            If .ListIndex < 0 And .ListCount > 0 Then .ListIndex = 0
        End With
        cbrThis.Bands(2).Visible = True
        cbrThis.Bands(2).Caption = "卡类别"
        mblnNotClick = False
    Case gBillType.消费卡
        strSQL = "Select 编号, 名称 From 消费卡类别目录 Where Nvl(启用, 0) >= 1 Order By 编号"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        mblnNotClick = True
        With cbo类别
            .Clear
            Do While Not rsTemp.EOF
                .AddItem Nvl(rsTemp!编号) & "-" & Nvl(rsTemp!名称)
                .ItemData(.NewIndex) = Val(Nvl(rsTemp!编号))
                If Val(str类别) = Val(Nvl(rsTemp!编号)) Then .ListIndex = .NewIndex
                rsTemp.MoveNext
            Loop
            If .ListIndex < 0 And .ListCount > 0 Then .ListIndex = 0
        End With
        cbrThis.Bands(2).Visible = True
        cbrThis.Bands(2).Caption = "卡类别"
        mblnNotClick = False
    Case Else
        cbrThis.Bands(2).Visible = False
    End Select
    LoadCombox = True
     
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetDefaultUserType()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置缺省的使用类别
    '编制:刘兴洪
    '日期:2011-04-27 14:23:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim int票种 As gBillType
    
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    int票种 = Val(Mid(lvwMain.SelectedItem.Key, 2))
    Select Case int票种
    Case gBillType.收费收据, gBillType.结帐收据
        lvwMain.SelectedItem.Tag = cbo类别.Text
    Case gBillType.预交收据, gBillType.就诊卡, gBillType.消费卡
        If cbo类别.ListIndex >= 0 Then
            lvwMain.SelectedItem.Tag = cbo类别.ItemData(cbo类别.ListIndex)
        Else
            lvwMain.SelectedItem.Tag = ""
        End If
    Case Else
    End Select
End Sub
