VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm应付款管理 
   Caption         =   "应付款管理"
   ClientHeight    =   6480
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9570
   FillColor       =   &H00404080&
   Icon            =   "frm应付款管理.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6480
   ScaleWidth      =   9570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox PicColor 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   7230
      ScaleHeight     =   540
      ScaleWidth      =   2265
      TabIndex        =   13
      Top             =   75
      Width           =   2295
      Begin VB.Label lblColor 
         BackColor       =   &H000000FF&
         Height          =   195
         Index           =   3
         Left            =   1215
         TabIndex        =   21
         Top             =   300
         Width           =   270
      End
      Begin VB.Label lblColor 
         BackColor       =   &H80000001&
         Height          =   195
         Index           =   2
         Left            =   45
         TabIndex        =   20
         Top             =   300
         Width           =   270
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00FF8080&
         Height          =   195
         Index           =   1
         Left            =   1215
         TabIndex        =   19
         Top             =   45
         Width           =   270
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00404080&
         Height          =   195
         Index           =   0
         Left            =   45
         TabIndex        =   18
         Top             =   45
         Width           =   270
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "冲销"
         Height          =   180
         Index           =   3
         Left            =   1605
         TabIndex        =   17
         Top             =   307
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "被冲销"
         Height          =   180
         Index           =   2
         Left            =   375
         TabIndex        =   16
         Top             =   307
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "已付款"
         Height          =   180
         Index           =   1
         Left            =   1605
         TabIndex        =   15
         Top             =   52
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "计划付款"
         Height          =   180
         Index           =   0
         Left            =   375
         TabIndex        =   14
         Top             =   52
         Width           =   720
      End
   End
   Begin VB.PictureBox PicRang 
      BackColor       =   &H8000000C&
      Height          =   315
      Left            =   2835
      ScaleHeight     =   255
      ScaleWidth      =   6630
      TabIndex        =   11
      Top             =   765
      Width           =   6690
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "查询范围:1999年8月12日至1999年9月12日"
         ForeColor       =   &H80000018&
         Height          =   180
         Left            =   75
         TabIndex        =   12
         Top             =   45
         Width           =   3330
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
      Height          =   2445
      Left            =   2805
      TabIndex        =   10
      Top             =   1095
      Width           =   6750
      _ExtentX        =   11906
      _ExtentY        =   4313
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   11
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   1050
      Top             =   2970
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
            Picture         =   "frm应付款管理.frx":08CA
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款管理.frx":0D22
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款管理.frx":117A
            Key             =   "ItemNo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款管理.frx":15CE
            Key             =   "No"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款管理.frx":1A26
            Key             =   "Write"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwList 
      Height          =   5355
      Left            =   0
      TabIndex        =   3
      Top             =   750
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   9446
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6120
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   2356
            Picture         =   "frm应付款管理.frx":1E7E
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11792
            MinWidth        =   600
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
      Left            =   6150
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款管理.frx":2712
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款管理.frx":2932
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款管理.frx":2B52
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款管理.frx":2D6E
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款管理.frx":2F8E
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款管理.frx":31AE
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款管理.frx":33CA
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款管理.frx":35E6
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款管理.frx":3800
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款管理.frx":395A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款管理.frx":3B76
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款管理.frx":3D96
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款管理.frx":3FB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款管理.frx":41CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款管理.frx":43E4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHot 
      Left            =   6750
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款管理.frx":45FE
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款管理.frx":481E
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款管理.frx":4A3E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款管理.frx":4C5A
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款管理.frx":4E7A
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款管理.frx":509A
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款管理.frx":52B6
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款管理.frx":54D2
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款管理.frx":56EC
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款管理.frx":5846
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款管理.frx":5A66
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款管理.frx":5C86
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款管理.frx":5EA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款管理.frx":60BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm应付款管理.frx":62D4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Height          =   780
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   1376
      BandCount       =   2
      BandBorders     =   0   'False
      _CBWidth        =   9570
      _CBHeight       =   780
      _Version        =   "6.7.8988"
      Child1          =   "tlbThis"
      MinHeight1      =   720
      Width1          =   11040
      NewRow1         =   0   'False
      MinHeight2      =   0
      NewRow2         =   0   'False
      BandStyle2      =   1
      Begin MSComctlLib.Toolbar tlbThis 
         Height          =   720
         Left            =   165
         TabIndex        =   2
         Top             =   30
         Width           =   9315
         _ExtentX        =   16431
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
            NumButtons      =   18
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
               Key             =   "Split"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "计划"
               Key             =   "SplitDue"
               Description     =   "计划"
               Object.ToolTipText     =   "制定付款计划"
               Object.Tag             =   "计划"
               ImageIndex      =   15
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split1"
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "审核"
               Key             =   "Verify"
               Description     =   "审核"
               Object.ToolTipText     =   "审核"
               Object.Tag             =   "审核"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "冲销"
               Key             =   "Strike"
               Description     =   "冲销"
               Object.ToolTipText     =   "冲销"
               Object.Tag             =   "冲销"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split2"
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "过滤"
               Key             =   "Search"
               Description     =   "过滤"
               Object.ToolTipText     =   "设置过滤条件"
               Object.Tag             =   "过滤"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "定位"
               Key             =   "Find"
               Description     =   "定位"
               Object.ToolTipText     =   "单据定位"
               Object.Tag             =   "定位"
               ImageIndex      =   14
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "刷新"
               Key             =   "Refresh"
               Description     =   "刷新"
               Object.ToolTipText     =   "刷新"
               Object.Tag             =   "刷新"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FindSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助主题"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Exit"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   11
            EndProperty
         EndProperty
         MouseIcon       =   "frm应付款管理.frx":64EE
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSplit 
      Height          =   1800
      Left            =   2805
      TabIndex        =   6
      Top             =   3975
      Width           =   6750
      _ExtentX        =   11906
      _ExtentY        =   3175
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      RowSizingMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin VB.Label lblTemp 
      Caption         =   "应付款总额："
      Height          =   165
      Index           =   2
      Left            =   2880
      TabIndex        =   9
      Top             =   5865
      Width           =   4440
   End
   Begin VB.Label lblTemp 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   1
      Left            =   2805
      TabIndex        =   8
      Top             =   5790
      Width           =   6750
   End
   Begin VB.Label lblHsc_s 
      Height          =   5355
      Left            =   2745
      MousePointer    =   9  'Size W E
      TabIndex        =   4
      Top             =   750
      Width           =   60
   End
   Begin VB.Label lblTemp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "应付款计划"
      ForeColor       =   &H80000018&
      Height          =   180
      Index           =   0
      Left            =   2985
      TabIndex        =   7
      Top             =   3705
      Width           =   900
   End
   Begin VB.Label lblVsc_s 
      BackColor       =   &H80000010&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2850
      MousePointer    =   7  'Size N S
      TabIndex        =   5
      Top             =   3645
      Width           =   6750
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
      Begin VB.Menu mnuFileExcel 
         Caption         =   "输出到&Excel"
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
      Begin VB.Menu mnuEditAdd 
         Caption         =   "增加(&A)"
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "修改(&M)"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "删除(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSplit 
         Caption         =   "计划(&S)"
      End
      Begin VB.Menu mnuEditLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditVerify 
         Caption         =   "审核(&V)"
      End
      Begin VB.Menu mnuEditStrike 
         Caption         =   "冲销(&C)"
      End
      Begin VB.Menu mnuEditLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditNO 
         Caption         =   "单据(&D)"
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
         Begin VB.Menu mnuViewLine1 
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
      Begin VB.Menu mnuViewSp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewSavePrint 
         Caption         =   "存盘打印(&C)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewVerifyPrint 
         Caption         =   "审核打印(&V)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewOpen 
         Caption         =   "过滤(&J)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "单据定位(&F)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewLine5 
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
         Caption         =   "&Web上的中联"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "中联主页(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "中联论坛(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "发送反馈(&K)"
         End
      End
      Begin VB.Menu mnuHelpLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)"
      End
   End
End
Attribute VB_Name = "frm应付款管理"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String
Private mblnFirst  As Boolean

Private Enum HeadCol
    ID
    单位ID
    收发ID
    付款序号
    计划序号
    记录状态
    付款标志
    单据号
    供应商
    品名
    规格
    单位
    批号
    产地
    随货单号
    发票号
    发票日期
    数量
    发票金额
    采购价
    采购金额
    填制人
    填制日期
    审核人
    审核日期
    当前库房
    当前库房库存
    全院库存
    药库单位
End Enum

Private mdtStartDate As Date    '填制日期
Private mdtEndDate As Date
Private mdtVerifyStartDate As Date  '审核日期
Private mdtVerifyEndDate As Date
Private mstrFind As String
Private mstr类型 As String
Private msngDownX As Single, msngDownY As Single, mSelKey As String
Private mlngModule As Long
Private mcllFilter As Collection
Private mbln付款标志 As Boolean

'by lesfeng 2010-1-7 性能优化
Private Sub InitFilter()
    '-----------------------------------------------------------------------------------------------------------
    '功能:初始化过滤条件
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2008-08-18 16:10:40
    '-----------------------------------------------------------------------------------------------------------
    Set mcllFilter = New Collection
    mcllFilter.Add Array("1901-01-01 00:00:00", "1901-01-01 00:00:00"), "填制日期"
    mcllFilter.Add Array("1901-01-01 00:00:00", "1901-01-01 00:00:00"), "审核日期"
    mcllFilter.Add Array("", ""), "单据号"
    mcllFilter.Add "", "随货单号"
    mcllFilter.Add "", "供应商id"
    mcllFilter.Add "", "库房ID"
    mcllFilter.Add "", "填制人"
    mcllFilter.Add "", "审核人"
    mstrFind = ""
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    Call 权限控制
End Sub

Private Sub Form_Load()
    Dim strReg As String
    mstrPrivs = gstrPrivs
    mblnFirst = True
    mlngModule = glngModul
    mstr类型 = "0000"
    '恢复参数
    'by lesfeng 2010-1-7 性能优化
    Call InitFilter
    
    mnuViewSavePrint.Checked = IIf(Val(zldatabase.GetPara("存盘打印", glngSys, mlngModule)) = 1, 1, 0) = 1
    mnuViewVerifyPrint.Checked = IIf(Val(zldatabase.GetPara("审核打印", glngSys, mlngModule)) = 1, 1, 0) = 1
    mbln付款标志 = Val(zldatabase.GetPara("外购入库需要经过标记付款后才能进行付款管理", glngSys, 0)) = 1
    
    mdtStartDate = Format(DateAdd("d", -7, zldatabase.Currentdate), "yyyy-MM-dd")
    mdtEndDate = Format(zldatabase.Currentdate, "yyyy-MM-dd")
    mdtVerifyStartDate = "1901-01-01"
    mdtVerifyEndDate = "1901-01-01"
        
    'by lesfeng 2010-1-7 性能优化
'    mstrFind = " AND A.记录状态 = 1 And A.审核日期 is Null And A.填制日期 Between [1] And [2]"
    mstrFind = " And (A.填制日期 Between [1] And [2]) and 审核日期 is null"
    lblRange = "查询范围:" & Format(DateAdd("d", -7, zldatabase.Currentdate), "yyyy年MM月dd日") & "至" & Format(zldatabase.Currentdate, "yyyy年MM月dd日")
    
    mcllFilter.Remove "填制日期"
    mcllFilter.Add Array(Format(mdtStartDate, "yyyy-mm-dd") & " 00:00:00", Format(mdtEndDate, "yyyy-mm-dd") & " 23:59:59"), "填制日期"
       
    mSelKey = ""
    List供应商
    RestoreWinState Me, App.ProductName
   
    '2006-04-25:刘兴宏,统一增加报表发布到模块的功能
    Call zldatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
    
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Dim lng分类id As Long
    Dim lng应付记录ID As Long
    Dim lng收发ID As Long
    Dim str发票号 As String
    Dim strNO As String
    Dim lng单位ID As Long
    Dim lng记录状态 As Long
    Dim lng付款序号 As Long
    If Not tvwList.SelectedItem Is Nothing Then
        lng分类id = Val(Mid(tvwList.SelectedItem.Key, 2))
    End If
    
    lng应付记录ID = Val(mshList.TextMatrix(mshList.Row, HeadCol.ID))
    lng收发ID = Val(mshList.TextMatrix(mshList.Row, HeadCol.收发ID))
    lng单位ID = Val(mshList.TextMatrix(mshList.Row, HeadCol.单位ID))
    str发票号 = Trim(mshList.TextMatrix(mshList.Row, HeadCol.发票号))
    strNO = Trim(mshList.TextMatrix(mshList.Row, HeadCol.单据号))
    lng记录状态 = Val(mshList.TextMatrix(mshList.Row, HeadCol.记录状态))
    lng付款序号 = Val(mshList.TextMatrix(mshList.Row, HeadCol.付款序号))
    
    '2006-04-25:刘兴宏:增加自定义报表发布到模块的功能
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, "分类=" & lng分类id, "应付记录=" & lng应付记录ID, "入库单据=" & lng收发ID, "发票号=" & str发票号, "NO=" & strNO, "供应商=" & lng单位ID, "记录状态=" & lng记录状态, "付款序号=" & lng付款序号)
    
End Sub

Private Sub List供应商()
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:加载供应商数型数据
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim rstTemp As New ADODB.Recordset
    Dim nodTemp As Node
    Dim strKey As String
    Dim strPrentKey As String
    
    If tvwList.SelectedItem Is Nothing Then
        strKey = "Root"
        strPrentKey = "Root"
    Else
        If tvwList.SelectedItem.Parent Is Nothing Then
            strKey = "Root"
            strPrentKey = "Root"
        Else
            strKey = tvwList.SelectedItem.Key
            strPrentKey = tvwList.SelectedItem.Parent.Key
        End If
    End If
    tvwList.Nodes.Clear
    tvwList.Nodes.Add , , "Root", "所有供应商", 1
    tvwList.Nodes("Root").Sorted = True
    Dim i As Long
    Dim str类型 As String
    
    
    str类型 = ""
    For i = 1 To Len(mstr类型)
        If Mid(mstr类型, i, 1) = 1 And Check相关权限(mstrPrivs, i) Then
            str类型 = str类型 & " or substr(类型," & i & ",1)=1"
        End If
    Next
    If str类型 <> "" Then
        str类型 = " And (" & Mid(str类型, 4) & ") "
    End If
    
    Dim str权限 As String
    str权限 = " and " & Get分类权限(mstrPrivs, "")
    
    'by lesfeng 2010-1-7 性能优化
    gstrSQL = "" & _
        "   Select id,上级id,编码,名称,类型,末级" & _
        "   From 供应商" & _
        "       where (撤档时间 is null or To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01') " & _
        "           and ( 末级<>1 or (末级=1 " & zl_获取站点限制() & "  " & _
                    str类型 & str权限 & "))" & _
        "   start with 上级id is null connect by prior id=上级id "
        
    Err = 0
    On Error GoTo ErrHand:
    zldatabase.OpenRecordset rstTemp, gstrSQL, Me.Caption
    With rstTemp
        While Not .EOF
            If IsNull(!上级ID) Then
                Set nodTemp = tvwList.Nodes.Add("Root", tvwChild, "K" & rstTemp("ID"), "【" & Nvl(!编码) & "】" & Nvl(!名称), IIf(Nvl(!末级, 0) = 0, 5, 2))
            Else
                Set nodTemp = tvwList.Nodes.Add("K" & !上级ID, tvwChild, "K" & rstTemp("ID"), "【" & !编码 & "】" & !名称, IIf(!末级 = 0, 5, 2))
            End If
            If strKey = "K" & Nvl(!ID) Then
                nodTemp.Selected = True
                nodTemp.Expanded = True
            End If
            nodTemp.Tag = Nvl(!类型)
            nodTemp.Sorted = True
            rstTemp.MoveNext
        Wend
    End With
    If tvwList.SelectedItem Is Nothing Then
        Err = 0
        On Error Resume Next
        If strPrentKey <> "" Then
            tvwList.Nodes(strPrentKey).Selected = True
            tvwList.Nodes(strPrentKey).Expanded = True
        End If
        If Err <> 0 Then
            tvwList.Nodes("Root").Selected = True
            tvwList.Nodes("Root").Expanded = True
        End If
    End If
    Err.Clear
    tvwList_NodeClick tvwList.SelectedItem
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Full应付记录()
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim i As Long
    Dim str类型 As String
    Dim lng上级id As Long
    
    str类型 = ""
    For i = 1 To Len(mstr类型)
        If Mid(mstr类型, i, 1) = 1 Then
            str类型 = str类型 & " or substr(b.类型," & i & ",1)=1"
        End If
    Next
    If str类型 <> "" Then
        str类型 = " And (" & Mid(str类型, 4) & ") "
    End If
    mshList.Redraw = False
    Dim str权限 As String
    str权限 = " and " & Get分类权限(mstrPrivs, "a.", False)
    
    If tvwList.SelectedItem.Key = "Root" Then
        'IIf(mstr类型 = "0000", "And (b.类型 is null or b.类型='" & mstr类型 & "')", "       and b.类型='" & mstr类型 & "'")
        strSQL = "" & _
            "   Select  a.id,a.单位ID as  单位id,a.收发ID,nvl(付款序号,0) as 付款序号,decode(a.计划序号,null,-1,0,-1,a.计划序号) as 计划序号,nvl(a.记录状态,1) as 记录状态," & _
            IIf(mbln付款标志, "decode(a.付款标志,1,'付款','') 付款标志,", "'' 付款标志,") & _
            "           a.no as 单据号,'['||b.编码||']'||b.名称 as 供应商,a.品名,a.规格,a.计量单位,a.批号,a.产地,a.随货单号, a.发票号,to_char(a.发票日期,'yyyy-mm-dd') as 发票日期," & _
            "           a.数量,a.发票金额,a.采购价,a.采购金额," & _
            "           a.填制人,to_char(a.填制日期,'yyyy-mm-dd hh24:mi:ss') as 填制日期," & _
            "           a.审核人,to_char(a.审核日期,'yyyy-mm-dd hh24:mi:ss') as 审核日期" & _
            IIf(mbln付款标志, ", e.名称 当前库房, c.全院库存, d.当前库房库存, c.药库单位 ", "") & _
            "   From 应付记录 a,供应商 b " & _
            IIf(mbln付款标志, ",(Select a.药品id, Round(a.全院库存 / b.药库包装, 5) 全院库存, b.药库单位 From (Select 药品id, Sum(实际数量) 全院库存 From 药品库存 Group By 药品id) A, 药品规格 B Where a.药品id = b.药品id) C, " & _
                              " (Select a.库房id, a.药品id, Round(a.当前库房库存 / b.药库包装, 5) 当前库房库存 From (Select 库房id, 药品id, Sum(实际数量) 当前库房库存 From 药品库存 Group By 库房id, 药品id) A, 药品规格 B Where a.药品id = b.药品id) D, " & _
                              " 部门表 E ", "") & _
            "   Where a.单位ID=b.id " & _
            IIf(mbln付款标志, " and a.项目id=c.药品id(+) and a.项目id=d.药品id(+) and a.库房id=d.库房id(+) and a.库房id=e.id(+) ", "") & _
            "     and not a.记录性质 in (-1,2) " & zl_获取站点限制(True, "b") & "   " & _
            "" & str类型 & str权限 & mstrFind & _
            "   Order By a.填制日期 desc,a.NO"
            lng上级id = 0
    Else
        strSQL = "" & _
            "   Select  a.id,a.单位ID as  单位id,a.收发ID,nvl(付款序号,0) as 付款序号,decode(a.计划序号,null,-1,0,-1,a.计划序号) as 计划序号,nvl(a.记录状态,1) as 记录状态," & _
            IIf(mbln付款标志, "decode(a.付款标志,1,'付款','') 付款标志,", "'' 付款标志,") & _
            "           a.no as 单据号,'['||b.编码||']'||b.名称 as 供应商,a.品名,a.规格,a.计量单位,a.批号,a.产地,a.随货单号, a.发票号,to_char(a.发票日期,'yyyy-mm-dd') as 发票日期," & _
            "           a.数量,a.发票金额,a.采购价,a.采购金额," & _
            "           a.填制人,to_char(a.填制日期,'yyyy-mm-dd hh24:mi:ss') as 填制日期," & _
            "           a.审核人,to_char(a.审核日期,'yyyy-mm-dd hh24:mi:ss') as 审核日期" & _
            IIf(mbln付款标志, ", e.名称 当前库房, c.全院库存, d.当前库房库存, c.药库单位 ", "") & _
            "   From 应付记录 a,供应商 b " & _
            IIf(mbln付款标志, ",(Select a.药品id, Round(a.全院库存 / b.药库包装, 5) 全院库存, b.药库单位 From (Select 药品id, Sum(实际数量) 全院库存 From 药品库存 Group By 药品id) A, 药品规格 B Where a.药品id = b.药品id) C, " & _
                              " (Select a.库房id, a.药品id, Round(a.当前库房库存 / b.药库包装, 5) 当前库房库存, b.药库单位 From (Select 库房id, 药品id, Sum(实际数量) 当前库房库存 From 药品库存 Group By 库房id, 药品id) A, 药品规格 B Where a.药品id = b.药品id) D, " & _
                              " 部门表 E", "") & _
            "   Where  a.单位ID=b.id " & _
            IIf(mbln付款标志, " and a.项目id=c.药品id(+) and a.项目id=d.药品id(+) and a.库房id=d.库房id(+) and a.库房id=e.id(+) ", "") & _
            "     and not a.记录性质 in (-1,2)  And a.单位ID in (select ID From 供应商 where  " & zl_获取站点限制(False) & "  start with id= [12] connect by prior id=上级id )" & _
            "" & str类型 & mstrFind & str权限 & _
            "   Order By a.填制日期 desc,a.NO"
            lng上级id = Val(Mid(tvwList.SelectedItem.Key, 2))
    End If
    
    'by lesfeng 2010-1-7 性能优化
    '填制日期: [1] [2]
    '审核日期: [3] [4]
    '单据号:   [5] [6]
    '随货单号: [7]
    '供应商id: [8]
    '填制人: [9]
    '审核人: [10]
    On Error GoTo errHandle
    Set rsTemp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(mcllFilter("填制日期")(0)), CDate(mcllFilter("填制日期")(1)), _
        CDate(mcllFilter("审核日期")(0)), CDate(mcllFilter("审核日期")(1)), CStr(mcllFilter("单据号")(0)), CStr(mcllFilter("单据号")(1)), _
        CStr(mcllFilter("随货单号")), CLng(Val(mcllFilter("供应商id"))), CStr(mcllFilter("填制人")), CStr(mcllFilter("审核人")), _
        Val(mcllFilter("库房ID")), lng上级id)

    mshList.Redraw = False
    If rsTemp.RecordCount > 0 Then
        Set mshList.Recordset = rsTemp
        mshList.Row = 1
        mshList.Col = HeadCol.单据号
        mshList.ColSel = mshList.Cols - 1
    Else
        mshList.Clear
        mshList.Rows = 2
    End If
    stbThis.Panels(2).Text = "当前共有" & rsTemp.RecordCount & "张单据"
    
    setGridColor
    formatMSH
    合计
    mshList.Redraw = True
    Full付款计划
    
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub 合计()
    Dim dblSum As Double
    Dim lngRow As Long
    With mshList
        For lngRow = 1 To .Rows - 1
            dblSum = dblSum + Val(.TextMatrix(lngRow, HeadCol.发票金额))
        Next
        '获取当前未付余额
        
    End With
    lblTemp(2) = "应付款总额：" & Format(dblSum, "###0.00;-###0.00;0;0") & "元"
    
End Sub

Private Sub SetRowColor(ByVal lngRow As Long, ByVal lngColor As Long, Optional blnList As Boolean = True)
    Dim intCol As Integer
    Dim objTmp As Object
    Set objTmp = IIf(blnList, mshList, mshSplit)
    With objTmp
        .Row = lngRow
        For intCol = 0 To .Cols - 1
            .Col = intCol
            .CellForeColor = lngColor
        Next
    End With
End Sub

Private Sub setGridColor()
    Dim lngStatus As Long
    Dim intRow As Integer
    Dim intCol As Integer
    
    With mshList
        'If mrsDue.RecordCount = 0 Then Exit Sub
        'mrsDue.MoveFirst
        For intRow = 1 To .Rows - 1
            lngStatus = Val(.TextMatrix(intRow, HeadCol.计划序号))
            If lngStatus <> -1 Then    '计划序号不为0说明已经有计划进行了付款
                SetRowColor intRow, &H404080
            End If
            lngStatus = Val(.TextMatrix(intRow, HeadCol.付款序号))
            If lngStatus <> 0 Then '此记录已付款
                SetRowColor intRow, &HFF8080
            Else
                lngStatus = Val(.TextMatrix(intRow, HeadCol.记录状态))
            
                If lngStatus Mod 3 = 0 Then             '被冲销记录
                    SetRowColor intRow, &H80000001
                ElseIf lngStatus Mod 3 = 2 Then         '冲销记录
                    SetRowColor intRow, &HFF
                End If
            End If
        Next
        .Col = 0
        .Row = 1
    End With
    
End Sub

Private Sub Full付款计划()
    Dim strSQL As String
    Dim lngID As Long
    Dim rsTemp As New ADODB.Recordset
    
    lngID = Val(mshList.TextMatrix(mshList.Row, HeadCol.ID))
    
    On Error GoTo errHandle
    If lngID = 0 Then
        Set mshSplit.Recordset = Nothing
        mshSplit.Clear
        mshSplit.Rows = 2
    Else
        'by lesfeng 2010-1-7 性能优化
        strSQL = "Select 付款序号,计划序号,计划金额,to_char(计划日期,'yyyy-MM-dd'),计划人,to_char(制定日期,'yyyy-MM-dd') From 应付记录 Where ID= [1] And 记录性质=-1 Order By 计划序号"
        Set rsTemp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, lngID)
        If rsTemp.RecordCount > 0 Then
            Set mshSplit.Recordset = rsTemp
        Else
            Set mshSplit.Recordset = Nothing
            mshSplit.Clear
            mshSplit.Rows = 2
        End If
    End If
    
    With mshSplit
        .Redraw = False
        .FormatString = "^付款序号|^计划序号|^计划金额|^计划付款日期|^计划人|^制定计划日期"
        .ColAlignment(0) = 7
        .ColWidth(0) = 0
        .ColWidth(1) = 1000: .ColAlignment(1) = 4
        .ColWidth(2) = 1100: .ColAlignment(2) = 7
        .ColWidth(3) = 1300: .ColAlignment(3) = 4
        .ColWidth(4) = 1100: .ColAlignment(4) = 1
        .ColWidth(5) = 1300: .ColAlignment(5) = 4
        setPlanGrdColor
        .Col = 1
        .ColSel = .Cols - 1
        .Redraw = True
    End With
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Form_Resize()
    Err = 0
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    
    If Me.WindowState <> vbMaximized Then
        If Me.Height < 5000 Then
            Me.Height = 5000
        End If
        If Me.Width < 4500 Then
            Me.Width = 4500
        End If
    End If
    
    If cbrThis.Bands(1).MinHeight <> tlbThis.Height Then cbrThis.Bands(1).MinHeight = tlbThis.Height
    cbrThis.Move 0, 0, Me.ScaleWidth
    
    If lblHsc_s.Left > Me.ScaleWidth - 2000 Then lblHsc_s.Left = Me.ScaleWidth - 2000
    
    lblHsc_s.Top = IIf(cbrThis.Visible, cbrThis.Height + 30, 0)
    lblHsc_s.Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - lblHsc_s.Top - 15
    tvwList.Move 0, lblHsc_s.Top, lblHsc_s.Left, lblHsc_s.Height
    If lblVsc_s.Top > Me.ScaleHeight - 2000 Then lblVsc_s.Top = Me.ScaleHeight - 2000
    
    lblVsc_s.Left = lblHsc_s.Left + lblHsc_s.Width
    lblVsc_s.Width = Me.ScaleWidth - lblVsc_s.Left
    With PicRang
        .Top = lblHsc_s.Top
        .Width = lblVsc_s.Width
        .Left = lblVsc_s.Left
    End With
    With mshList
        .Left = lblVsc_s.Left
        .Top = lblHsc_s.Top + PicRang.Height + 50
        .Width = lblVsc_s.Width
        .Height = lblVsc_s.Top - .Top
    End With
    
    lblTemp(0).Move (lblVsc_s.Width - lblTemp(0).Width) / 2 + lblVsc_s.Left, (lblVsc_s.Height - lblTemp(0).Height) / 2 + lblVsc_s.Top + 1
    With mshSplit
        .Left = lblVsc_s.Left
        .Top = lblVsc_s.Top + lblVsc_s.Height
        .Width = lblVsc_s.Width
        .Height = tvwList.Top + tvwList.Height - .Top - lblTemp(1).Height
        
    End With
    lblTemp(1).Move lblVsc_s.Left, mshSplit.Top + mshSplit.Height, lblVsc_s.Width
    lblTemp(2).Move lblTemp(1).Left + 60, lblTemp(1).Top + (lblTemp(1).Height - lblTemp(2).Height) / 2, lblTemp(1).Width - 120
    With PicColor
        .Left = ScaleWidth - .Width - 100
        .Top = 80
    End With
    mnuViewToolButton.Checked = cbrThis.Visible
    mnuViewStatus.Checked = stbThis.Visible
    mnuViewToolText.Checked = tlbThis.Buttons(1).Caption <> ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub mnuEditAdd_Click()
    Dim blnReturn As Boolean
    
    '增加
    If tvwList.SelectedItem.Image = 2 Then
        Call frm应付款编辑.ShowCard(Me, 0, g新增, mstrPrivs, Val(Mid(tvwList.SelectedItem.Key, 2)), , blnReturn)
    Else
        Call frm应付款编辑.ShowCard(Me, 0, g新增, mstrPrivs, 0, , blnReturn)
    End If
    If blnReturn = False Then Exit Sub
    '填充
    Full应付记录
    
End Sub

Private Sub mnuEditDelete_Click()
'删除
    Dim strSQL As String
    Dim intRow As Integer
    Dim lngID As Long
    
    lngID = Val(mshList.TextMatrix(mshList.Row, HeadCol.ID))
    If lngID = 0 Then Exit Sub
    
    If MsgBox("你确定要删除该应付记录吗？", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then Exit Sub
    strSQL = "ZL_应付记录_DELETE(" & lngID & ")"
    
    zldatabase.ExecuteProcedure strSQL, Me.Caption
    
    With mshList
        intRow = .Row
        If .Rows > 2 Then
            .RemoveItem intRow
        ElseIf .Rows = 2 Then
            .Rows = 3
            .RemoveItem intRow
            SetEnabled
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
    End With
    Full付款计划
End Sub

Private Sub mnuEditModify_Click()
    Dim blnReturn As Boolean
    '修改
    Dim lngID As Long
    Dim int记录状态 As Integer
    
    lngID = Val(mshList.TextMatrix(mshList.Row, HeadCol.ID))
    If lngID = 0 Then Exit Sub
    
    Call frm应付款编辑.ShowCard(Me, lngID, g修改, mstrPrivs, , , blnReturn)
    If blnReturn = False Then Exit Sub
    
    Full应付记录
End Sub

Private Sub mnuEditNO_Click()
    Dim blnReturn  As Boolean
    Dim lngID As Long
    Dim int记录状态 As Integer
    Dim bytRec As RecBillStatus
    
    lngID = Val(mshList.TextMatrix(mshList.Row, HeadCol.ID))
    int记录状态 = Val(mshList.TextMatrix(mshList.Row, HeadCol.记录状态))
    If lngID = 0 Then Exit Sub
'    Select Case int记录状态
'    Case 1
'        bytRec = 正常记录
'    Case 2
'        bytRec = 冲销记录
'    Case Else
'        bytRec = 被冲销记录
'    End Select
    '审核
    Call frm应付款编辑.ShowCard(Me, lngID, g查看, mstrPrivs, 0, int记录状态, blnReturn)
    If blnReturn = False Then Exit Sub
    mnuViewRefresh_Click
End Sub

Private Sub mnuEditSplit_Click()
    '计划
    Dim lngID As Long
    If mnuEditSplit.Enabled = False Then Exit Sub
    lngID = Val(mshList.TextMatrix(mshList.Row, HeadCol.ID))
    If lngID = 0 Then Exit Sub
    
    frm付款计划.计划 Me, lngID
    Full付款计划
End Sub

Private Sub mnuEditStrike_Click()
    '冲销
    Dim blnReturn As Boolean
    Dim lngID As Long
    lngID = Val(mshList.TextMatrix(mshList.Row, HeadCol.ID))
    If lngID = 0 Then Exit Sub
    
    Call frm应付款编辑.ShowCard(Me, lngID, g取消, mstrPrivs, 0, 正常记录, blnReturn)
    If blnReturn = False Then Exit Sub
    mnuViewRefresh_Click
    
End Sub

Private Sub mnuEditVerify_Click()
    Dim blnReturn  As Boolean
    Dim lngID As Long
    lngID = Val(mshList.TextMatrix(mshList.Row, HeadCol.ID))
    If lngID = 0 Then Exit Sub
    '审核
    Call frm应付款编辑.ShowCard(Me, lngID, g审核, mstrPrivs, 0, 正常记录, blnReturn)
    If blnReturn = False Then Exit Sub
    mnuViewRefresh_Click
End Sub

Private Sub mnuFileExcel_Click()
'输出到Excel
    mshList.Redraw = False
    subPrint 3
    mshList.Redraw = True
    mshList.Col = 0
    mshList.ColSel = mshList.Cols - 1
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
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
    zlPrintSet
End Sub

Private Sub lblHsc_s_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    msngDownX = X
End Sub

Private Sub lblHsc_s_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        With lblHsc_s
            If .Left + X - msngDownX < 2000 Then Exit Sub
            If .Left + X - msngDownX > ScaleWidth - 2000 Then Exit Sub
            .Left = .Left + X - msngDownX
        End With
        Call Form_Resize
    End If
End Sub

Private Sub lblVsc_s_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    msngDownY = Y
End Sub

Private Sub lblVsc_s_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        With lblVsc_s
            If .Top + Y - msngDownY < 2000 Then Exit Sub
            If .Top + Y - msngDownY > ScaleHeight - 2000 Then Exit Sub
            .Top = .Top + Y - msngDownY
        End With
        Call Form_Resize
    End If
End Sub

Private Sub mnuViewFind_Click()
'定位
End Sub

Private Sub mnuViewOpen_Click()
    Dim strCon As String
    Dim strFind As String
    Dim str类型 As String
    Dim cllFilter As Collection
    
    str类型 = mstr类型
    'by lesfeng 2010-1-7 性能优化
    strFind = frm应付款过滤.GetSearch(Me, mstrPrivs, mdtStartDate, mdtEndDate, mdtVerifyStartDate, mdtVerifyEndDate, mstr类型, cllFilter)
    If strFind = "" Then Exit Sub
    mstrFind = strFind
    Set mcllFilter = cllFilter
    '加载数据
    '
    If Format(mdtStartDate, "yyyy-mm-dd") = "1901-01-01" And Format(mdtVerifyStartDate, "yyyy-mm-dd") = "1901-01-01" Then
    ElseIf Format(mdtStartDate, "yyyy-mm-dd") <> "1901-01-01" And Format(mdtVerifyStartDate, "yyyy-mm-dd") <> "1901-01-01" Then
        strCon = "查询范围:填制日期 " & Format(mdtStartDate, "yyyy年MM月dd日") & "至" & Format(mdtEndDate, "yyyy年MM月dd日") & "  审核日期 " & Format(mdtVerifyStartDate, "yyyy年MM月dd日") & "至" & Format(mdtVerifyEndDate, "yyyy年MM月dd日")
    ElseIf Format(mdtStartDate, "yyyy-mm-dd") <> "1901-01-01" Then
        strCon = "查询范围:填制日期 " & Format(mdtStartDate, "yyyy年MM月dd日") & "至" & Format(mdtEndDate, "yyyy年MM月dd日")
    ElseIf Format(mdtVerifyStartDate, "yyyy-mm-dd") <> "1901-01-01" Then
        strCon = "查询范围:审核日期 " & Format(mdtVerifyStartDate, "yyyy年MM月dd日") & "至" & Format(mdtVerifyEndDate, "yyyy年MM月dd日")
    End If
    lblRange = strCon
    If str类型 <> mstr类型 Then
        '条件重置
        mSelKey = ""
        List供应商
    Else
        Full应付记录
    End If
End Sub

Private Sub mnuViewRefresh_Click()
    '刷新
    Err = 0
    On Error Resume Next
    mSelKey = ""
    tvwList_NodeClick tvwList.SelectedItem
End Sub

Private Sub mnuViewSavePrint_Click()
    mnuViewSavePrint.Checked = Not mnuViewSavePrint.Checked
    Call zldatabase.SetPara("存盘打印", IIf(mnuViewSavePrint.Checked, "1", "0"), glngSys, mlngModule)
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    PicColor.Visible = mnuViewToolButton.Checked And mnuViewToolButton.Checked
    cbrThis.Visible = mnuViewToolButton.Checked
    cbrThis.Bands(1).MinHeight = tlbThis.Height
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim buttTemp As Button
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    PicColor.Visible = mnuViewToolButton.Checked And mnuViewToolButton.Checked
    For Each buttTemp In tlbThis.Buttons
        If mnuViewToolText.Checked Then
            buttTemp.Caption = buttTemp.Tag
        Else
            buttTemp.Caption = ""
        End If
    Next
    cbrThis.Bands(1).MinHeight = tlbThis.Height
    Form_Resize
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Form_Resize
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
       ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(hwnd)
End Sub

Private Sub mnuViewVerifyPrint_Click()
    mnuViewVerifyPrint.Checked = Not mnuViewVerifyPrint.Checked
    Call zldatabase.SetPara("审核打印", IIf(mnuViewVerifyPrint.Checked, "1", "0"), glngSys, mlngModule)
    
End Sub

Private Sub mshList_Click()
    Dim strSQL As String
    Dim lngID As Long
    lngID = Val(mshList.TextMatrix(mshList.Row, HeadCol.ID))
    If lngID <> Val(mshSplit.Tag) Then
        Full付款计划
        lngID = lngID
    End If
End Sub

Private Sub mshList_DblClick()
    If mnuEditModify.Enabled And mnuEditModify.Visible Then
        mnuEditModify_Click
    Else
        mnuEditNO_Click
    End If
End Sub

Private Sub mshList_EnterCell()
    SetEnabled
End Sub

Private Sub mshList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    PopupMenu mnuEdit
End Sub

Private Sub mshList_RowColChange()
    mshList_Click
End Sub

Private Sub mshSplit_DblClick()
    mnuEditSplit_Click
End Sub

Private Sub tlbthis_ButtonClick(ByVal Button As MSComctlLib.Button)
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
            mnuEditDelete_Click
        Case "SplitDue"
            mnuEditSplit_Click
        Case "Find"
            mnuViewFind_Click
        Case "Search"
            Call mnuViewOpen_Click
        Case "Refresh"
            Call mnuViewRefresh_Click
        Case "Help"
            Call mnuHelpTitle_Click
        Case "Exit"
            Call mnuFileExit_Click
        Case "Verify"
            Call mnuEditVerify_Click
        Case "Strike"
            mnuEditStrike_Click
    End Select
End Sub

Private Sub tlbthis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then Me.PopupMenu mnuViewTool
End Sub

Private Sub tvwList_NodeClick(ByVal Node As MSComctlLib.Node)
    If mSelKey = Node.Key Then Exit Sub
    mSelKey = Node.Key
    Full应付记录
    SetEnabled
End Sub

Private Sub SetEnabled()
    Dim blnData As Boolean '有无数据
    Dim blnVerfiy As Boolean '是否审核
    Dim blnCancel As Boolean  '冲销单据
    Dim blnPayMoney As Boolean '已经付款
    Dim blnSys As Boolean       '系统传入数据
    Dim blnPlan As Boolean
    Dim blnSign As Boolean
    
    If mshList.Rows <= 1 Then
        blnData = False
        blnVerfiy = False
        blnPayMoney = False
        blnPlan = False
        blnSign = False
    Else
        With mshList
            blnData = Val(mshList.TextMatrix(1, HeadCol.ID)) <> 0
            blnVerfiy = Trim(.TextMatrix(.Row, HeadCol.审核日期)) <> ""
            blnCancel = Val(.TextMatrix(.Row, HeadCol.记录状态)) <> 1
            blnPlan = Val(.TextMatrix(.Row, HeadCol.记录状态)) = 1 Or Val(.TextMatrix(.Row, HeadCol.记录状态)) = 3
            blnPayMoney = Val(.TextMatrix(.Row, HeadCol.付款序号)) <> 0
            blnSys = Val(mshList.TextMatrix(.Row, HeadCol.收发ID)) <> 0
            blnSign = mshList.TextMatrix(.Row, HeadCol.付款标志) = "付款"
        End With
    End If
    
    mnuEditModify.Enabled = blnData And Not blnVerfiy And Not blnSys
    mnuEditDelete.Enabled = blnData And Not blnVerfiy And Not blnSys
    mnuEditStrike.Enabled = blnData And blnVerfiy And Not blnCancel And (Not blnPayMoney) And Not blnSys
    If mbln付款标志 Then
        mnuEditVerify.Enabled = blnData And Not blnVerfiy And (Not blnSys Or blnSign)
    Else
        mnuEditVerify.Enabled = blnData And Not blnVerfiy And Not blnSys
    End If
    mnuEditSplit.Enabled = blnData And blnVerfiy And blnPlan And (Not blnPayMoney)
    mnuEditNO.Enabled = blnData
    
    tlbThis.Buttons("Modify").Enabled = mnuEditModify.Enabled
    tlbThis.Buttons("Delete").Enabled = mnuEditDelete.Enabled
    tlbThis.Buttons("Verify").Enabled = mnuEditVerify.Enabled
    tlbThis.Buttons("Strike").Enabled = mnuEditStrike.Enabled
    tlbThis.Buttons("SplitDue").Enabled = mnuEditSplit.Enabled
    
    mnuFilePreView.Enabled = blnData
    mnuFilePrint.Enabled = blnData
    mnuFileExcel.Enabled = blnData
    tlbThis.Buttons("Print").Enabled = blnData
    tlbThis.Buttons("PrintView").Enabled = blnData
    
End Sub


Private Sub formatMSH()
    Dim intCol As Integer
    With mshList
        .Cols = 29
        For intCol = 0 To .Cols - 1
            .ColAlignmentFixed(intCol) = 4
        Next
        .TextMatrix(0, HeadCol.ID) = "ID"
        .TextMatrix(0, HeadCol.收发ID) = "收发ID"
        .TextMatrix(0, HeadCol.付款序号) = "付款序号"
        .TextMatrix(0, HeadCol.计划序号) = "计划序号"
        .TextMatrix(0, HeadCol.记录状态) = "记录状态"
        
        .TextMatrix(0, HeadCol.付款标志) = "付款标志"
        
        .TextMatrix(0, HeadCol.单据号) = "单据号"
        .TextMatrix(0, HeadCol.随货单号) = "随货单号"
        .TextMatrix(0, HeadCol.发票号) = "发票号"
        .TextMatrix(0, HeadCol.发票日期) = "发票日期"
        .TextMatrix(0, HeadCol.发票金额) = "发票金额"
        .TextMatrix(0, HeadCol.供应商) = "供应商"
        .TextMatrix(0, HeadCol.品名) = "品名"
        .TextMatrix(0, HeadCol.规格) = "规格"
        .TextMatrix(0, HeadCol.单位) = "单位"
        .TextMatrix(0, HeadCol.批号) = "批号"
        .TextMatrix(0, HeadCol.产地) = "产地"
        .TextMatrix(0, HeadCol.数量) = "数量"
        .TextMatrix(0, HeadCol.采购价) = "采购价"
        .TextMatrix(0, HeadCol.采购金额) = "采购金额"
        
        .TextMatrix(0, HeadCol.填制人) = "填制人"
        .TextMatrix(0, HeadCol.填制日期) = "填制日期"
        .TextMatrix(0, HeadCol.审核人) = "审核人"
        .TextMatrix(0, HeadCol.审核日期) = "审核日期"
        
        .ColWidth(HeadCol.ID) = 0
        .ColWidth(HeadCol.收发ID) = 0
        .ColWidth(HeadCol.单位ID) = 0
        .ColWidth(HeadCol.付款序号) = 0
        .ColWidth(HeadCol.计划序号) = 0
        .ColWidth(HeadCol.记录状态) = 0
    
        If mblnFirst = False Then
            SetEnabled
            Exit Sub
        End If
               
        .ColWidth(HeadCol.供应商) = 2000
        .ColWidth(HeadCol.单据号) = 1400
        .ColWidth(HeadCol.随货单号) = 1400
        .ColWidth(HeadCol.发票号) = 1400
        .ColWidth(HeadCol.发票日期) = 1400
        .ColWidth(HeadCol.发票金额) = 1400
        .ColWidth(HeadCol.品名) = 2400
        .ColWidth(HeadCol.规格) = 2000
        .ColWidth(HeadCol.单位) = 800
        .ColWidth(HeadCol.批号) = 1400
        .ColWidth(HeadCol.产地) = 2000
        .ColWidth(HeadCol.数量) = 1400
        .ColWidth(HeadCol.采购价) = 1400
        .ColWidth(HeadCol.采购金额) = 1400
        .ColWidth(HeadCol.填制人) = 1000
        .ColWidth(HeadCol.填制日期) = 1600
        .ColWidth(HeadCol.审核人) = 1000
        .ColWidth(HeadCol.审核日期) = 1600
        
        .ColAlignment(HeadCol.ID) = 1
        .ColAlignment(HeadCol.收发ID) = 1
        .ColAlignment(HeadCol.付款序号) = 1
        .ColAlignment(HeadCol.计划序号) = 1
        .ColAlignment(HeadCol.记录状态) = 1
        
        .ColAlignment(HeadCol.供应商) = 1
        .ColAlignment(HeadCol.填制人) = 4
        .ColAlignment(HeadCol.填制日期) = 4
        .ColAlignment(HeadCol.审核人) = 4
        .ColAlignment(HeadCol.审核日期) = 4
        
        .ColAlignment(HeadCol.单据号) = 4
        .ColAlignment(HeadCol.随货单号) = 1
        .ColAlignment(HeadCol.发票号) = 1
        .ColAlignment(HeadCol.发票日期) = 4
        .ColAlignment(HeadCol.发票金额) = 7
        .ColAlignment(HeadCol.品名) = 1
        .ColAlignment(HeadCol.规格) = 1
        .ColAlignment(HeadCol.单位) = 4
        .ColAlignment(HeadCol.批号) = 1
        .ColAlignment(HeadCol.产地) = 1
        .ColAlignment(HeadCol.数量) = 7
        .ColAlignment(HeadCol.采购价) = 7
        .ColAlignment(HeadCol.采购金额) = 7
        .ColAlignment(HeadCol.填制人) = 1
        
        If mbln付款标志 Then
            .TextMatrix(0, HeadCol.当前库房) = "当前库房"
            .TextMatrix(0, HeadCol.当前库房库存) = "当前库房库存"
            .TextMatrix(0, HeadCol.全院库存) = "全院库存"
            .TextMatrix(0, HeadCol.药库单位) = "药库单位"
            .ColWidth(HeadCol.付款标志) = 800
            .ColWidth(HeadCol.当前库房) = 1400
            .ColWidth(HeadCol.当前库房库存) = 1400
            .ColWidth(HeadCol.全院库存) = 1400
            .ColWidth(HeadCol.药库单位) = 800
            .ColAlignment(HeadCol.当前库房库存) = 7
            .ColAlignment(HeadCol.全院库存) = 7
        Else
            .ColWidth(HeadCol.付款标志) = 0
            .ColWidth(HeadCol.当前库房) = 0
            .ColWidth(HeadCol.当前库房库存) = 0
            .ColWidth(HeadCol.全院库存) = 0
            .ColWidth(HeadCol.药库单位) = 0
        End If
    End With
    SetEnabled
End Sub

Private Sub subPrint(bytMode As Byte)
'功能:进行打印,预览和输出到EXCEL
'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
'    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As Object
    Dim objRow As New zlTabAppRow
    
    Set objPrint = New zlPrint1Grd
    
    If Me.ActiveControl Is mshSplit Then
        
        objRow.Add "应付单位:" & mshList.TextMatrix(mshList.Row, HeadCol.供应商)
        objPrint.UnderAppRows.Add objRow
        
        Set objRow = New zlTabAppRow
        objRow.Add "应付单据号:" & mshList.TextMatrix(mshList.Row, HeadCol.单据号)
        objRow.Add "品名:" & mshList.TextMatrix(mshList.Row, HeadCol.品名)
        objRow.Add "规格:" & mshList.TextMatrix(mshList.Row, HeadCol.规格)
        objPrint.UnderAppRows.Add objRow
        
        objPrint.Title.Text = "付款计划清册表"
        Set objPrint.Body = mshSplit
    Else
        Dim strRange As String
        If Format(mdtStartDate, "yyyy-mm-dd") = "1901-01-01" And Format(mdtVerifyStartDate, "yyyy-mm-dd") = "1901-01-01" Then
            strRange = "审核日期：" & Format(mdtVerifyStartDate, "yyyy年MM月dd日") & "至" & Format(mdtVerifyEndDate, "yyyy年MM月dd日")
        ElseIf Format(mdtVerifyStartDate, "yyyy-mm-dd") <> "1901-01-01" Then
            strRange = "填制日期：" & Format(mdtStartDate, "yyyy年MM月dd日") & "至" & Format(mdtEndDate, "yyyy年MM月dd日") & "  审核日期：" & Format(mdtVerifyStartDate, "yyyy年MM月dd日") & "至" & Format(mdtVerifyEndDate, "yyyy年MM月dd日")
        Else
            strRange = "填制日期：" & Format(mdtStartDate, "yyyy年MM月dd日") & "至" & Format(mdtEndDate, "yyyy年MM月dd日")
        End If
        objRow.Add strRange
        objPrint.Title.Text = "应付款清册表"
        Set objPrint.Body = mshList
    End If
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印时间：" & Format(zldatabase.Currentdate, "yyyy年MM月dd日")
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
End Sub

Private Sub 权限控制()
    Dim blnAdd As Boolean
    Dim blnModify As Boolean
    Dim blnDelete As Boolean
    Dim blnVerify As Boolean
    Dim blnCancel As Boolean
    Dim blnPlan As Boolean
    
    blnAdd = InStr(1, mstrPrivs, "登记") <> 0
    blnModify = InStr(1, mstrPrivs, "修改") <> 0
    blnDelete = InStr(1, mstrPrivs, "删除") <> 0
    blnVerify = InStr(1, mstrPrivs, "审核") <> 0
    blnCancel = InStr(1, mstrPrivs, "冲销") <> 0
    blnPlan = InStr(1, mstrPrivs, "付款计划")
        
    mnuEditAdd.Visible = blnAdd
    mnuEditModify.Visible = blnModify
    mnuEditDelete.Visible = blnDelete
    mnuEditVerify.Visible = blnVerify
    mnuEditStrike.Visible = blnCancel
    mnuEditSplit.Visible = blnPlan
    
    tlbThis.Buttons("Add").Visible = blnAdd
    tlbThis.Buttons("Modify").Visible = blnModify
    tlbThis.Buttons("Delete").Visible = blnDelete
    
    tlbThis.Buttons("Verify").Visible = blnVerify
    tlbThis.Buttons("Strike").Visible = blnCancel
    
    tlbThis.Buttons("SplitDue").Visible = blnPlan
    
    
    If (Not blnAdd And Not blnModify And Not blnDelete) Or Not blnPlan Then
        tlbThis.Buttons("Split").Visible = False
        mnuEditLine1.Visible = False
    End If
    
    If (Not blnVerify And Not blnCancel) Or Not blnPlan Then
        mnuEditLine2.Visible = False
        tlbThis.Buttons("Split1").Visible = False
    End If
    
    If Not (blnAdd Or blnModify Or blnDelete Or blnPlan Or blnVerify Or blnVerify Or blnCancel) Then
        tlbThis.Buttons("Split2").Visible = False
        mnuEditLine3.Visible = False
    End If
    
End Sub

Private Sub setPlanGrdColor()
    Dim lngRow As Long
    
    With mshSplit
            For lngRow = 1 To .Rows - 1
                If Val(.TextMatrix(lngRow, 0)) <> 0 Then
                    '已付
                    SetRowColor lngRow, &HFF8080, False
                Else
                    '未付
                    
                End If
            Next
            
    End With
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub
