VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmStuffQuery 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   Caption         =   "卫材库存查询"
   ClientHeight    =   7110
   ClientLeft      =   165
   ClientTop       =   825
   ClientWidth     =   9495
   Icon            =   "frmStuffQuery.frx":0000
   LockControls    =   -1  'True
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   7110
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picColor 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   5160
      ScaleHeight     =   255
      ScaleWidth      =   2055
      TabIndex        =   12
      Top             =   6120
      Width           =   2055
      Begin VB.PictureBox picColor1 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   14
         Top             =   0
         Width           =   260
      End
      Begin VB.PictureBox picColor2 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   960
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   13
         Top             =   0
         Width           =   260
      End
      Begin VB.Label lblColor2 
         AutoSize        =   -1  'True
         Caption         =   "近效期"
         Height          =   180
         Left            =   1320
         TabIndex        =   16
         Top             =   30
         Width           =   540
      End
      Begin VB.Label lblColor1 
         AutoSize        =   -1  'True
         Caption         =   "停用"
         Height          =   180
         Left            =   360
         TabIndex        =   15
         Top             =   30
         Width           =   360
      End
   End
   Begin VB.PictureBox picFind 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1470
      ScaleHeight     =   255
      ScaleWidth      =   2655
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   6810
      Width           =   2655
      Begin VB.TextBox txt材料信息 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   780
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   0
         Width           =   1245
      End
      Begin VB.Label lbl材料信息 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "材料信息"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   30
         TabIndex        =   10
         Top             =   37
         Width           =   720
      End
   End
   Begin MSComctlLib.ImageList imglvw 
      Left            =   2985
      Top             =   2205
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":0982
            Key             =   "root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":268C
            Key             =   "child"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picVLine_S 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5460
      Left            =   2940
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5460
      ScaleWidth      =   45
      TabIndex        =   1
      Top             =   1305
      Width           =   45
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   1125
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   1984
      BandCount       =   2
      _CBWidth        =   9495
      _CBHeight       =   1125
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinHeight1      =   720
      Width1          =   2730
      NewRow1         =   0   'False
      Caption2        =   "库房"
      Child2          =   "cob库房"
      MinHeight2      =   300
      Width2          =   6780
      NewRow2         =   -1  'True
      Begin VB.ComboBox cob库房 
         Height          =   300
         Left            =   585
         TabIndex        =   5
         Text            =   "cob库房"
         Top             =   780
         Width           =   8820
      End
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   720
         Left            =   165
         TabIndex        =   3
         Top             =   30
         Width           =   9240
         _ExtentX        =   16298
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgTbrStard"
         HotImageList    =   "imgTbrHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   15
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
               Description     =   "打印"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "重置"
               Key             =   "重置"
               Object.ToolTipText     =   "重置条件"
               Object.Tag             =   "重置"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               ImageIndex      =   3
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查找"
               Key             =   "查找"
               Object.ToolTipText     =   "查找"
               Object.Tag             =   "查找"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "明细"
               Key             =   "明细"
               Object.ToolTipText     =   "材料明细帐"
               Object.Tag             =   "明细"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "总帐"
               Key             =   "总帐"
               Object.ToolTipText     =   "材料总帐"
               Object.Tag             =   "总帐"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "条码"
               Key             =   "条码"
               Object.ToolTipText     =   "卫材条码"
               Object.Tag             =   "条码"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "刷新"
               Key             =   "刷新"
               Object.ToolTipText     =   "刷新"
               Object.Tag             =   "刷新"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "字体"
               Key             =   "字体"
               Object.ToolTipText     =   "字体"
               Object.Tag             =   "字体"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "帮助"
               Description     =   "帮助"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "退出"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imgTbrHot 
      Left            =   1425
      Top             =   780
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
            Picture         =   "frmStuffQuery.frx":4396
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":45B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":47CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":49E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":4C04
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":4E1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":5038
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":5254
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":5470
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":568C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":58A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":6180
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgTbrStard 
      Left            =   690
      Top             =   810
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
            Picture         =   "frmStuffQuery.frx":649A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":66B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":68D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":6AEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":6D08
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":6F22
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":713C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":7358
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":7574
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":7790
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":79AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQuery.frx":8284
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf材料信息_S 
      Height          =   2985
      Left            =   3000
      TabIndex        =   6
      Top             =   1260
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   5265
      _Version        =   393216
      BackColor       =   16777215
      FixedCols       =   0
      GridColor       =   -2147483631
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf分批库存_S 
      Height          =   870
      Left            =   3240
      TabIndex        =   7
      Top             =   5100
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   1535
      _Version        =   393216
      BackColor       =   16777215
      FixedCols       =   0
      GridColor       =   -2147483631
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   11
      Top             =   6744
      Width           =   9492
      _ExtentX        =   16748
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmStuffQuery.frx":83E6
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11668
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
   Begin MSComctlLib.TreeView tvwSection_S 
      Height          =   4350
      Left            =   60
      TabIndex        =   0
      Top             =   1275
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   7673
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imglvw"
      Appearance      =   1
   End
   Begin VB.Label lbl分批_S 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "分批库存"
      ForeColor       =   &H8000000E&
      Height          =   180
      Left            =   3270
      MousePointer    =   7  'Size N S
      TabIndex        =   4
      Top             =   4920
      Width           =   6585
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
      Begin VB.Menu mnuExcel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnuFileBatch 
         Caption         =   "批量打印明细帐(&B)"
      End
      Begin VB.Menu mnuViewLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "参数设置(&R)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
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
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "工具栏(&T)"
         Begin VB.Menu mnuViewToolbarStand 
            Caption         =   "标准按钮(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolbarText 
            Caption         =   "文本标签(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "状态栏(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFont 
         Caption         =   "字体(&F)"
         Begin VB.Menu mnuViewFontSize 
            Caption         =   "小字体"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewFontSize 
            Caption         =   "中字体"
            Index           =   1
         End
         Begin VB.Menu mnuViewFontSize 
            Caption         =   "大字体"
            Index           =   2
         End
      End
      Begin VB.Menu mnuViewForeColor 
         Caption         =   "前景色(&C)"
      End
      Begin VB.Menu mnuViewBackColor 
         Caption         =   "背景色(&B)"
      End
      Begin VB.Menu mnuviewLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "刷新(&R)"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "查找(&F)"
         Shortcut        =   ^F
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
            Caption         =   "发送反馈(&K)..."
         End
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
   Begin VB.Menu mnuPopu 
      Caption         =   "弹出"
      Visible         =   0   'False
      Begin VB.Menu mnuPopuOpen 
         Caption         =   "打开(&O)"
      End
      Begin VB.Menu mnuPopuBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopuFontSize 
         Caption         =   "小字体"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuPopuFontSize 
         Caption         =   "中字体"
         Index           =   1
      End
      Begin VB.Menu mnuPopuFontSize 
         Caption         =   "大字体"
         Index           =   2
      End
   End
   Begin VB.Menu mnuReportBill 
      Caption         =   "报表菜单"
      Visible         =   0   'False
      Begin VB.Menu mnuBill 
         Caption         =   "单据(&D)"
      End
   End
End
Attribute VB_Name = "frmStuffQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--------------------------------------------------------
Public mblnDo As Boolean
Public mbytUint As Byte '卫材单位

Dim mintFont As Integer
Dim WithEvents mrsData  As ADODB.Recordset
Attribute mrsData.VB_VarHelpID = -1
Dim mrsTreeData As ADODB.Recordset

Dim mstrStartDate As String
Dim mstrEndDate As String

Dim mblnFirst As Boolean              '确定是否第一次使用本系统
Dim mbln库存数 As Boolean
Dim mbln包含停用 As Boolean
Dim mintMonths As Integer
Dim mstrPrivs As String
Dim mblnColor As Boolean

Private mlngCardRow As Long
Private mlngRow As Long
Private mstrCardSort As String                 '排序列

Private mblnNoClick As Boolean

Private Const MLNG白色 As Long = &H80000005
Private Const MLNG黑色 As Long = &H80000008
Private Const MLNG蓝色 As Long = &H8000000D
Private Const MLNGSEL As Long = &HA87B82
Private Const MLNG本色 As Long = &H8000000F
Private Const MLNG灰色 As Long = &HC0C0C0
Private Const MLNG红色 As Long = &HC0           '停用
Private mblnCostView As Boolean             '查看成本价相关信息 true-允许查看 false-不允许查看
Private Const mstrCaption As String = "卫材库存查询"


Private mstrOthers() As String  '  0-编码,1-名称,2-简码,3-规格,4-产地,5-指定产地
'----------------------
'三张报表的变量设置
Public WithEvents mobjReport As zl9Report.clsReport
Attribute mobjReport.VB_VarHelpID = -1
Private mlngCurReport As Long
Private mobjCurSheet As Object
Dim mstrNoS As String
'-----------------------
Private mlngModule As Long
'----------------------------------------------------------------------------------------------------------
'刘兴宏:增加小数位数的格式串
'修改:2007/03/06
Private mFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------
Private Sub cbrThis_Resize()
    Form_Resize
End Sub

Private Sub cob库房_Click()
    If mblnNoClick Then Exit Sub
    If Me.tvwSection_S.Nodes.Count = 0 Then Exit Sub
    Me.tvwSection_S.Tag = ""
    ReFreshStuffData Me.cob库房.ItemData(Me.cob库房.ListIndex), mstrStartDate, mstrEndDate, IIf(Left(Me.tvwSection_S.SelectedItem.Key, 1) = "R", 0, Mid(Me.tvwSection_S.SelectedItem.Key, 2))
End Sub

Private Sub cob库房_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cob库房.ListCount = 0 Then Call zlControl.ControlSetFocus(Msf材料信息_S): Exit Sub
    
    If cob库房.ListIndex >= 0 Then
        If Val(cob库房.Tag) = cob库房.ItemData(cob库房.ListIndex) Then
            Call zlControl.ControlSetFocus(Msf材料信息_S, True)
            Exit Sub
        End If
    End If
    
    If Select部门选择器(Me, cob库房, Trim(cob库房.Text), "V,K,12,W", IIf(InStr(1, mstrPrivs, "所有库房") = 0, True, False)) = False Then
        Exit Sub
    End If
    If cob库房.ListIndex >= 0 Then
        cob库房.Tag = cob库房.ItemData(cob库房.ListIndex)
    End If
End Sub


Private Sub cob库房_LostFocus()
    Dim i As Long
    
    If cob库房.ListCount = 0 Then Exit Sub
    If cob库房.ListIndex < 0 Then
        For i = 0 To cob库房.ListCount - 1
            If Val(cob库房.Tag) = cob库房.ItemData(i) Then
                mblnNoClick = True
                cob库房.ListIndex = i: Exit For
            End If
        Next
    End If
    mblnNoClick = False
End Sub


Private Sub mrsData_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If mblnColor Then Exit Sub
    If mrsData.RecordCount = 0 Then
        RefreshBatch Me.cob库房.ItemData(Me.cob库房.ListIndex), 0
        Exit Sub
    End If
    If mrsData.EOF Then mrsData.MoveFirst
    If mrsData.EOF Then Exit Sub
    RefreshBatch Me.cob库房.ItemData(Me.cob库房.ListIndex), mrsData.Fields("材料id").Value
    If Me.tvwSection_S.Tag <> "T" Then Exit Sub
    err = 0
    On Error Resume Next
    Me.tvwSection_S.Nodes("_" & mrsData.Fields("分类id").Value).Selected = True
    Me.tvwSection_S.Nodes("_" & mrsData.Fields("分类id").Value).Expanded = True
End Sub

Private Sub Form_Activate()
    If Not mblnFirst Then Exit Sub
    
    mblnFirst = False
    
    tbrThis.Buttons("条码").Visible = gblnCode
    
    If Not ReFreshTreeView() Then Unload Me: Exit Sub
    ReFreshStuffData Me.cob库房.ItemData(Me.cob库房.ListIndex), mstrStartDate, mstrEndDate, IIf(Left(Me.tvwSection_S.SelectedItem.Key, 1) = "R", 0, Mid(Me.tvwSection_S.SelectedItem.Key, 2))
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim strOthers(0 To 6) As String
    mlngModule = glngModul
    mstrPrivs = gstrPrivs
    mblnCostView = zlStr.IsHavePrivs(mstrPrivs, "查看成本价")
    
    For i = 0 To 6
        strOthers(i) = ""
    Next
    mstrOthers = strOthers
    mblnFirst = True
    Call GetParaSet
    mnuViewForeColor.Visible = False
    mnuViewBackColor.Visible = False
    
    Call mnuViewFontSize_Click(mintFont)
    
    With Msf分批库存_S
        .Clear
        .Cols = IIf(gblnCode = True, 15, 13)
        .Rows = 2
        .FixedRows = 1
        .FixedCols = 0
        
        .ColAlignmentFixed(0) = 4
        .ColAlignmentFixed(1) = 4
        .ColAlignmentFixed(2) = 4
        .ColAlignmentFixed(3) = 4
        .ColAlignmentFixed(4) = 4
        .ColAlignmentFixed(5) = 4
        .ColAlignmentFixed(6) = 4
        .ColAlignmentFixed(7) = 4
        .ColAlignmentFixed(8) = 4
        .ColAlignmentFixed(9) = 4
        .ColAlignmentFixed(10) = 4
        
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
        .ColAlignment(4) = 7
        .ColAlignment(5) = 7
        .ColAlignment(6) = 7
        .ColAlignment(7) = 7
        .ColAlignment(8) = 7
        .ColAlignment(9) = 7
        .ColAlignment(10) = 7

        .TextMatrix(0, 0) = "库房"
        .TextMatrix(0, 1) = "批号"
        .TextMatrix(0, 2) = "失效期"
        .TextMatrix(0, 3) = "产地"
        .TextMatrix(0, 4) = "可用库存"
        .TextMatrix(0, 5) = "库存数量"
        .TextMatrix(0, 6) = "库存金额"
        .TextMatrix(0, 7) = "成本价"
        .TextMatrix(0, 8) = "成本金额"
        .TextMatrix(0, 9) = "库存差价"
        .TextMatrix(0, 10) = "最后进价"
        
        .ColWidth(0) = 1000
        .ColWidth(1) = 1000
        .ColWidth(2) = 1200
        .ColWidth(3) = 1200
        .ColWidth(4) = 1500
        .ColWidth(5) = 1500
        .ColWidth(6) = 1500
        .ColWidth(10) = 1500
        
        If gblnCode = True Then
            .TextMatrix(0, 11) = "商品条码"
            .TextMatrix(0, 12) = "内部条码"
            .TextMatrix(0, 13) = "售价"
            .TextMatrix(0, 14) = "供应商"
            
            .ColAlignment(11) = 7
            .ColAlignment(12) = 7
            .ColAlignment(13) = 7
            .ColAlignment(14) = 1
            
            .ColAlignmentFixed(11) = 4
            .ColAlignmentFixed(12) = 4
            .ColAlignmentFixed(13) = 4
            .ColAlignmentFixed(14) = 4
            
            .ColWidth(11) = 2000
            .ColWidth(12) = 2000
            .ColWidth(13) = 1500
            .ColWidth(14) = 1500
        Else
            .TextMatrix(0, 11) = "售价"
            .TextMatrix(0, 12) = "供应商"
            .ColAlignment(11) = 7
            .ColAlignment(12) = 1
            .ColAlignmentFixed(11) = 4
            .ColAlignmentFixed(12) = 4
            .ColWidth(11) = 1500
            .ColWidth(12) = 1500
        End If
    End With
    Call SetFormat(True)
    RestoreWinState Me, App.ProductName, mstrCaption
    Msf分批库存_S.ColWidth(7) = IIf(mblnCostView = False, 0, 1500)
    Msf分批库存_S.ColWidth(8) = IIf(mblnCostView = False, 0, 1500)
    Msf分批库存_S.ColWidth(9) = IIf(mblnCostView = False, 0, 1500)
    
    '加载报表
    Set mobjReport = New zl9Report.clsReport

    '2006-04-25:刘兴宏,统一增加报表发布到模块的功能
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
    
    Call 设置权限
    
    Call SetParent(picFind.hwnd, stbThis.hwnd)
    picFind.Top = 80
    picFind.Left = stbThis.Panels(1).Width + 80
    
    stbThis.Panels(2).Picture = picColor
End Sub

Private Sub Form_Resize()
    Dim intTop As Integer, intButton As Integer
    If Me.WindowState = 1 Then Exit Sub
    intTop = IIf(Me.cbrThis.Visible, Me.cbrThis.Height, 0)
    intButton = IIf(Me.stbThis.Visible, Me.stbThis.Height, 0)
    
    On Error Resume Next
    Me.picVLine_S.Top = intTop + Me.ScaleTop
    Me.picVLine_S.Height = Me.ScaleHeight - Me.tvwSection_S.Top - intButton
    If Me.picVLine_S.Left < 500 Then Me.picVLine_S.Left = 500
    If Me.picVLine_S.Left > Me.ScaleWidth - 500 Then Me.picVLine_S.Left = Me.ScaleWidth - 500
    
    Me.tvwSection_S.Left = Me.ScaleLeft
    Me.tvwSection_S.Width = Me.picVLine_S.Left - Me.tvwSection_S.Left
    Me.tvwSection_S.Top = Me.ScaleTop + intTop
    
    If Me.ScaleWidth - Me.picVLine_S.Left - Me.picVLine_S.Width < 500 Then
        Me.Width = Me.picVLine_S.Left + Me.picVLine_S.Width + 500
    End If
    If Me.ScaleHeight - Me.lbl分批_S.Top - Me.lbl分批_S.Height < 500 Then
        Me.Height = Me.lbl分批_S.Top + Me.lbl分批_S.Height + 2000
    End If
    If Me.ScaleHeight < 500 Then
        Me.Height = 2000
    End If
    Me.tvwSection_S.Height = Me.ScaleHeight - tvwSection_S.Top - intButton
    
    Me.lbl分批_S.Left = Me.picVLine_S.Left + Me.picVLine_S.Width
    Me.lbl分批_S.Width = Me.ScaleWidth - Me.lbl分批_S.Left
    With Me.Msf分批库存_S
        .Left = Me.lbl分批_S.Left
        .Width = Me.lbl分批_S.Width
    End With
    
    Me.Msf材料信息_S.Left = Me.lbl分批_S.Left
    Me.Msf材料信息_S.Width = Me.lbl分批_S.Width
        
    If Me.Msf分批库存_S.Visible Then
        With Me.Msf分批库存_S
            .Top = Me.lbl分批_S.Top + Me.lbl分批_S.Height
            .Height = Me.ScaleHeight - .Top - intButton
        End With
        Me.Msf材料信息_S.Top = intTop + 50
        Me.Msf材料信息_S.Height = Me.lbl分批_S.Top - Me.Msf材料信息_S.Top
    Else
        Me.Msf材料信息_S.Top = intTop + 50
        Me.Msf材料信息_S.Height = Me.ScaleHeight - Me.Msf材料信息_S.Top - intButton
    End If
    
    With picColor
        .Top = Me.ScaleHeight - .Height - 30
        .Left = Me.ScaleWidth - stbThis.Panels(3).Width - stbThis.Panels(4).Width - .Width - 300
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    SaveWinState Me, App.ProductName, mstrCaption
End Sub

Private Sub lbl分批_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Me.lbl分批_S.Top = Me.lbl分批_S.Top + y
        If Me.lbl分批_S.Top < 5000 Then Me.lbl分批_S.Top = 5000
        If Me.Height - Me.lbl分批_S.Top < 2000 Then Me.lbl分批_S.Top = Me.Height - 2000
        Form_Resize
    End If
End Sub

Private Sub mnuEXCEL_Click()
    grdPrint 1
End Sub

Private Sub mnuFileBatch_Click()
    With FrmPrintList
        .mstrPrivs = mstrPrivs
        .Show 1, Me
    End With
End Sub

Private Sub mnufileexit_Click()
    Unload Me
End Sub
Private Sub GetParaSet()
    '功能:获取参数设置
    mbytUint = Val(zlDatabase.GetPara("卫材单位", glngSys, mlngModule))
 
    '刘兴宏:增加小数格式化串
    With mFMT
        .FM_成本价 = GetFmtString(mbytUint, g_成本价)
        .FM_金额 = GetFmtString(mbytUint, g_金额)
        .FM_零售价 = GetFmtString(mbytUint, g_售价)
        .FM_数量 = GetFmtString(mbytUint, g_数量)
    End With
    
    
    mbln库存数 = IIf(Val(zlDatabase.GetPara("只显示有库存卫材", glngSys, mlngModule)) = 1, 1, 0) = 1
    mintMonths = Val(zlDatabase.GetPara("报警月数", glngSys, mlngModule, 3))  '
    mbln包含停用 = IIf(Val(zlDatabase.GetPara("包含停用卫材", glngSys, mlngModule)) = 1, 1, 0) = 1
    mintFont = Val(zlDatabase.GetPara("字体字号", glngSys, mlngModule, 9))
   
End Sub
Private Sub mnuFileOpen_Click()
    mblnDo = False
    Call frmStuffQueryParaSet.参数设置(Me, mlngModule, mstrPrivs)
    If Not mblnDo Then Exit Sub
    If Me.tvwSection_S.Nodes.Count = 0 Then Exit Sub
    
    Call GetParaSet
    Me.tvwSection_S.Tag = ""
    ReFreshStuffData Me.cob库房.ItemData(Me.cob库房.ListIndex), mstrStartDate, mstrEndDate, IIf(Left(Me.tvwSection_S.SelectedItem.Key, 1) = "R", 0, Mid(Me.tvwSection_S.SelectedItem.Key, 2))
End Sub

Private Sub mnuFilePrint_Click()
    grdPrint 3
End Sub

Private Sub mnuFilePrintSet_Click()
     zlPrintSet
End Sub

Private Sub mnuFilePrintView_Click()
  grdPrint 0
End Sub
Private Sub grdPrint(blnIsPreview As Byte)
    '---------------------------------------------------
    '功能：    根据屏幕组织表上附加项目，打印预览
    '参数：
    '     blnIsPreview: 0表示预览 1表示输出到EXCEL 其它表示打印
    '返回：
    '---------------------------------------------------
    Dim objPrint As New zlPrint1Grd
    Dim objRow As zlTabAppRow
    
    objPrint.Title.Text = "卫材库存查询"
    Set objRow = New zlTabAppRow
    objRow.Add "库房：" & Me.cob库房.Text
    objRow.Add "卫材用途：" & Me.tvwSection_S.SelectedItem.Text
    objRow.Add "截止日期：" & Format(Sys.Currentdate, "yyyy年MM月DD日")
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add "打印人:" & UserInfo.用户名
    objRow.Add "打印时间:" & Format(Sys.Currentdate, "yyyy年MM月DD日 HH:MM")
    objPrint.BelowAppRows.Add objRow
    Set objPrint.Body = Msf材料信息_S
    
    Call Msf材料信息_S_LostFocus
    If blnIsPreview = 0 Then
         zlPrintOrView1Grd objPrint, 2
    Else
      If blnIsPreview = 1 Then
            zlPrintOrView1Grd objPrint, 3
      Else
        Select Case zlPrintAsk(objPrint)
            Case 1
                 zlPrintOrView1Grd objPrint, 1
            Case 2
                zlPrintOrView1Grd objPrint, 2
            Case 3
                zlPrintOrView1Grd objPrint, 3
        End Select
      End If
    End If
    Set objPrint = Nothing
End Sub

Private Sub mnuHelpAbout_Click()
   ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Sub mnuViewFind_Click()
    Dim rsTmp As New ADODB.Recordset
    Dim strFind As String
    Me.tvwSection_S.Tag = ""
    FrmStuffQueryFind.Show 1, Me
    strFind = FrmStuffQueryFind.mstrTemp
    mstrOthers = FrmStuffQueryFind.mstrOthers
    
    Unload FrmStuffQueryFind
    If strFind = "" Then Exit Sub
    If Not ReFreshFilterData(cob库房.ItemData(cob库房.ListIndex), strFind) Then Exit Sub
    Me.tvwSection_S.Tag = "T"
End Sub

Private Sub mnuViewFontSize_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 2
        Me.mnuViewFontSize(i).Checked = False
    Next
    Me.mnuViewFontSize(Index).Checked = True

    Select Case Index
    Case 0
        Me.Msf材料信息_S.Font.Size = 9
        Me.tvwSection_S.Font.Size = 9
        Msf分批库存_S.Font.Size = 9
     Case 1
        Me.Msf材料信息_S.Font.Size = 11
        Me.tvwSection_S.Font.Size = 11
        Msf分批库存_S.Font.Size = 11
    Case 2
        Me.Msf材料信息_S.Font.Size = 15
        Me.tvwSection_S.Font.Size = 15
        Msf分批库存_S.Font.Size = 15
    End Select
    mintFont = Index
    Call zlDatabase.SetPara("字体字号", mintFont, glngSys, mlngModule)
    Form_Resize
    Me.Refresh
End Sub

Private Sub mnuViewForeColor_Click()
    Dim lngForeColor As Long
    lngForeColor = zlGetColor(Me.Msf材料信息_S.ForeColor)
    Me.Msf材料信息_S.Redraw = False
    Me.Msf材料信息_S.ForeColor = lngForeColor
    Me.Msf材料信息_S.Redraw = True
    
End Sub
Private Sub mnuViewBackColor_Click()
    Dim lngBackColor As Long
    lngBackColor = zlGetColor(Me.Msf材料信息_S.BackColor)
    Me.Msf材料信息_S.BackColor = lngBackColor
    
End Sub

Private Sub showReportMXZ()
    If mrsData Is Nothing Then Exit Sub
    If Not (mrsData.State = 1) Then Exit Sub
    If mrsData.RecordCount = 0 Then Exit Sub
    If ISCheckReport("ZL1_INSIDE_1721_2") = False Then Exit Sub
    
    If cob库房.ItemData(cob库房.ListIndex) = 0 Then
        Call mobjReport.ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1721_2", Me, "材料=" & mrsData.Fields("名称").Value & "|" & mrsData.Fields("材料id").Value, "库房=所有库房|is not null", "单位=" & Choose(mbytUint, "散装单位", "包装单位") & "|" & mbytUint, "开始日期=" & Format(DateAdd("m", -1, Sys.Currentdate), "yyyy-MM-DD"), "结束日期=" & Format(Sys.Currentdate, "yyyy-MM-DD"))
    Else
        Call mobjReport.ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1721_2", Me, "材料=" & mrsData.Fields("名称").Value & "|" & mrsData.Fields("材料id").Value, "库房=" & cob库房.Text & "|=  " & cob库房.ItemData(cob库房.ListIndex), "单位=" & Choose(mbytUint, "散装单位", "包装单位") & "|" & mbytUint, "开始日期=" & Format(DateAdd("m", -1, Sys.Currentdate), "yyyy-MM-DD"), "结束日期=" & Format(Sys.Currentdate, "yyyy-MM-DD"), "单位=" & mbytUint)      '"包含未审核单据=0| And A.审核人 Is Not NULL"
    End If
    
End Sub


Private Sub showReportCode()
    If mrsData Is Nothing Then Exit Sub
    If Not (mrsData.State = 1) Then Exit Sub
    If mrsData.RecordCount = 0 Then Exit Sub
    
    If cob库房.ItemData(cob库房.ListIndex) = 0 Then Exit Sub
    If Msf分批库存_S.Row = 0 Then Exit Sub
    If Msf分批库存_S.TextMatrix(Msf分批库存_S.Row, 11) = "" And Msf分批库存_S.TextMatrix(Msf分批库存_S.Row, 12) = "" Then Exit Sub
    
    Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1721_4", Me, "库房=" & cob库房.Text & "|=  " & cob库房.ItemData(cob库房.ListIndex), "商品条码=" & Msf分批库存_S.TextMatrix(Msf分批库存_S.Row, 11), "内部条码=" & Msf分批库存_S.TextMatrix(Msf分批库存_S.Row, 12))
        
End Sub

Private Sub mnuViewRefresh_Click()
    ReFreshStuffData Me.cob库房.ItemData(Me.cob库房.ListIndex), mstrStartDate, mstrEndDate, IIf(Left(Me.tvwSection_S.SelectedItem.Key, 1) = "R", 0, Mid(Me.tvwSection_S.SelectedItem.Key, 2))
End Sub

Private Sub ShowReportMXB()
    On Error Resume Next
    Call mobjReport.ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1721_3", Me, "库房=" & Me.cob库房.Text & "|" & IIf(Me.cob库房.ItemData(Me.cob库房.ListIndex) = 0, " is not null ", "=" & Me.cob库房.ItemData(Me.cob库房.ListIndex)), "单位=" & Choose(mbytUint, "散装单位", "包装单位") & "|" & mbytUint)
End Sub

Private Sub mnuViewStatus_Click()
    Me.mnuViewStatus.Checked = Not Me.mnuViewStatus.Checked
    Me.stbThis.Visible = Me.mnuViewStatus.Checked
    Form_Resize
End Sub

Private Sub ShowReportSumAccount()
    err = 0
    On Error Resume Next
    Call mobjReport.ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1721_1", Me, "库房=" & Me.cob库房.Text & "|" & IIf(Me.cob库房.ItemData(Me.cob库房.ListIndex) = 0, " is not null ", "=" & Me.cob库房.ItemData(Me.cob库房.ListIndex)))
End Sub
Private Sub SetReportCtrlIndexEnabled()
    '设置指定报表的Enable属性
    Dim i As Long
    For i = 0 To mnuReportItem.UBound
        If Split(mnuReportItem(i).Tag & ",", ",")(1) = "ZL1_INSIDE_1721_2" Then
            tbrThis.Buttons("明细").Enabled = mrsData.RecordCount <> 0
            mnuReportItem.Item(i).Enabled = mrsData.RecordCount <> 0
        End If
    Next
End Sub
Private Sub mnuReportItem_Click(Index As Integer)


    If Split(mnuReportItem(Index).Tag & ",", ",")(1) = "ZL1_INSIDE_1721_2" Then
        '明细帐
        Call showReportMXZ
        Exit Sub
    End If
    
    If Split(mnuReportItem(Index).Tag & ",", ",")(1) = "ZL1_INSIDE_1721_3" Then
        '明细表
        Call ShowReportMXB
        Exit Sub
    End If
    
    
    If Split(mnuReportItem(Index).Tag & ",", ",")(1) = "ZL1_INSIDE_1721_1" Then
        '总帐
        Call ShowReportSumAccount
        Exit Sub
    End If
    
    
    Dim lng库房ID As Long, lng分类id As Long, lng材料ID As Long
    If cob库房.ListIndex < 0 Then
        lng库房ID = 0
    Else
        lng库房ID = cob库房.ItemData(cob库房.ListIndex)
    End If
    
    If Not tvwSection_S.SelectedItem Is Nothing Then
        lng分类id = Val(Mid(tvwSection_S.SelectedItem.Key, 2))
    End If
    
    lng材料ID = 0
    If Not mrsData Is Nothing Then
        If mrsData.State = 1 Then
            If Not mrsData.EOF Then
               lng材料ID = mrsData.Fields("材料id").Value
            End If
        End If
    End If
    
    '2006-04-25:刘兴宏:增加自定义报表发布到模块的功能
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, "库房=" & lng库房ID, "分类=" & lng分类id, "材料=" & lng材料ID)
    
End Sub

Private Sub mnuViewToolbarStAnd_Click()
    Dim intCount As Integer
    Me.mnuViewToolbarStand.Checked = Not Me.mnuViewToolbarStand.Checked
    Me.mnuViewToolbarText.Enabled = Me.mnuViewToolbarStand.Checked
    Me.cbrThis.Visible = Me.mnuViewToolbarStand.Checked
    
    If Me.mnuViewToolbarText.Checked Then
        For intCount = 1 To Me.tbrThis.Buttons.Count
            Me.tbrThis.Buttons(intCount).Caption = Me.tbrThis.Buttons(intCount).Tag
        Next
    Else
        For intCount = 1 To Me.tbrThis.Buttons.Count
            Me.tbrThis.Buttons(intCount).Caption = ""
        Next
    End If
    Me.cbrThis.Bands(1).MinHeight = Me.tbrThis.Height
    Me.cbrThis.Refresh
    Form_Resize

End Sub
Private Sub mnuViewToolbarText_Click()
    Dim intCount As Integer
    Me.mnuViewToolbarText.Checked = Not Me.mnuViewToolbarText.Checked
    If Me.mnuViewToolbarText.Checked Then
        For intCount = 1 To Me.tbrThis.Buttons.Count
            Me.tbrThis.Buttons(intCount).Caption = Me.tbrThis.Buttons(intCount).Tag
        Next
    Else
        For intCount = 1 To Me.tbrThis.Buttons.Count
            Me.tbrThis.Buttons(intCount).Caption = ""
        Next
    End If
    Me.cbrThis.Bands(1).MinHeight = Me.tbrThis.Height
    Me.cbrThis.Refresh
    Form_Resize

End Sub

Private Sub Msf分批库存_S_EnterCell()
    On Error Resume Next
    Dim intCol As Integer
    Dim lngColor As Long
    Dim LngSelectRow As Long
    
    With Msf分批库存_S
        .Redraw = False
        
        LngSelectRow = .Row     '保存当前选中行
        If mlngRow <> 0 Then
            .Row = mlngRow       '清除上次选中行
            For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellBackColor = MLNG白色
                .CellForeColor = IIf(.RowData(.Row) = 0, MLNG黑色, glng报警)
            Next
            .Col = 0
        End If
        
        mlngRow = LngSelectRow
        .Row = mlngRow     '设置当前选中行
        If Not Me.ActiveControl Is Nothing Then
            For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellBackColor = IIf(Me.ActiveControl.Name = "Msf分批库存_S", MLNGSEL, MLNG灰色)
                If Me.ActiveControl.Name = "Msf分批库存_S" Then
                    lngColor = IIf(.RowData(.Row) = 0, MLNG白色, glng报警)
                Else
                    lngColor = IIf(.RowData(.Row) = 0, MLNG黑色, glng报警)
                End If
                .CellForeColor = lngColor
            Next
        End If
        .Col = 0
        
        .Redraw = True
    End With
End Sub

Private Sub Msf分批库存_S_GotFocus()
    Dim intCol As Integer
    With Msf分批库存_S
        .GridColorFixed = MLNG黑色
        .GridColor = MLNG黑色
        For intCol = 0 To .Cols - 1
            .Redraw = False
            .Col = intCol
            .CellBackColor = MLNGSEL
            .CellForeColor = IIf(.RowData(.Row) = 0, MLNG白色, glng报警)
            .Redraw = True
        Next
        .Col = 0
    End With
    With Msf材料信息_S
        For intCol = 0 To .Cols - 1
            .Redraw = False
            .Col = intCol
            .CellBackColor = MLNG灰色
            .CellForeColor = IIf(Trim(.TextMatrix(.Row, 16)) = "", MLNG黑色, MLNG红色)
            .Redraw = True
        Next
        .Col = 0
    End With
End Sub

Private Sub Msf分批库存_S_LostFocus()
    Dim intCol As Integer
    With Msf分批库存_S
        .GridColorFixed = &H80000011
        .GridColor = &H80000011
        For intCol = 0 To .Cols - 1
            .Redraw = False
            .Col = intCol
            .CellBackColor = MLNG灰色
            .CellForeColor = IIf(.RowData(.Row) = 0, MLNG黑色, glng报警)
            .Redraw = True
        Next
        .Col = 0
    End With
End Sub

Private Sub Msf材料信息_S_DblClick()
    showReportMXZ
End Sub

Private Sub Msf材料信息_S_EnterCell()
    On Error Resume Next
    Dim intCol As Integer
    Dim lngColor As Long
    Dim LngSelectRow As Long
    
    With Msf材料信息_S
        .Redraw = False
        
        LngSelectRow = .Row     '保存当前选中行
        If mlngCardRow <> 0 Then
            .Row = mlngCardRow       '清除上次选中行
            For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellBackColor = MLNG白色
                .CellForeColor = IIf(Trim(.TextMatrix(.Row, 16)) = "", MLNG黑色, MLNG红色)
            Next
            .Col = 0
        End If
        
        mlngCardRow = LngSelectRow
        .Row = mlngCardRow       '设置当前选中行
        If Not ActiveControl Is Nothing Then
            For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellBackColor = IIf(Me.ActiveControl.Name = "Msf材料信息_S", MLNGSEL, MLNG灰色)
                If Me.ActiveControl.Name = "Msf材料信息_S" Then
                    lngColor = IIf(Trim(.TextMatrix(.Row, 16)) = "", MLNG白色, MLNG红色)
                Else
                    lngColor = IIf(Trim(.TextMatrix(.Row, 16)) = "", MLNG黑色, MLNG红色)
                End If
                .CellForeColor = lngColor
            Next
        End If
        .Col = 0
        
        .Redraw = True
        
        '读取其批次信息
        If Val(.TextMatrix(.Row, 0)) <> 0 Then
            mrsData.MoveFirst
            mrsData.Find "材料ID=" & Val(.TextMatrix(.Row, 0))
        End If
    End With
End Sub

Private Sub Msf材料信息_S_GotFocus()
    Dim intCol As Integer
    With Msf材料信息_S
        .GridColorFixed = MLNG黑色
        .GridColor = MLNG黑色
        For intCol = 0 To .Cols - 1
            .Redraw = False
            .Col = intCol
            .CellBackColor = MLNGSEL
            .CellForeColor = IIf(Trim(.TextMatrix(.Row, 16)) = "", MLNG白色, MLNG红色)
            .Redraw = True
        Next
        .Col = 0
    End With
    With Msf分批库存_S
        For intCol = 0 To .Cols - 1
            .Redraw = False
            .Col = intCol
            .CellBackColor = MLNG灰色
            .CellForeColor = IIf(.RowData(.Row) = 0, MLNG黑色, glng报警)
            .Redraw = True
        Next
        .Col = 0
    End With
End Sub

Private Sub Msf材料信息_S_LostFocus()
    Dim intCol As Integer
    With Msf材料信息_S
        .GridColorFixed = &H80000011
        .GridColor = &H80000011
        For intCol = 0 To .Cols - 1
            .Redraw = False
            .Col = intCol
            .CellBackColor = MLNG灰色
            .CellForeColor = IIf(Trim(.TextMatrix(.Row, 16)) = "", MLNG黑色, MLNG红色)
            .Redraw = True
        Next
        .Col = 0
    End With
End Sub

Private Sub Msf材料信息_S_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim StrHeader As String
    Dim intCol As Integer, intMouseCol As Integer

    '实现列排序
    If Button = 1 Then
        With Msf材料信息_S
            If .MouseRow <> 0 Then Exit Sub
            If mrsData Is Nothing Then Exit Sub
            If mrsData.State = 0 Then Exit Sub
            If mrsData.EOF Then Exit Sub
            
            intMouseCol = .MouseCol
            StrHeader = .TextMatrix(0, intMouseCol)
            If StrHeader = "库存数量" Then
                StrHeader = "实际数量"
            ElseIf StrHeader = "库存金额" Then
                StrHeader = "实际金额"
            ElseIf StrHeader = "库存差价" Then
                StrHeader = "实际差价"
            End If
            
            If Mid(mstrCardSort, 2) = StrHeader Then
                mstrCardSort = IIf(Mid(mstrCardSort, 1, 1) = "A", "D", "A") & StrHeader
                mrsData.Sort = StrHeader & IIf(Mid(mstrCardSort, 1, 1) = "D", " Desc", " Asc")
            Else
                mstrCardSort = "A" & StrHeader
                mrsData.Sort = StrHeader & " Asc"
            End If
            
            FS.ShowFlash ("正在排序中，请稍候...")
            Call SetFormat(False)
            Call Msf材料信息_S_EnterCell
            FS.StopFlash
        End With
    Else
        PopupMenu mnuView
    End If
End Sub

Private Sub picVLine_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Me.picVLine_S.Left = Me.picVLine_S.Left + x
        Form_Resize
    End If
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    With Button
        Select Case .Key
        Case "预览"
            mnuFilePrintView_Click
        Case "打印"
            grdPrint 3
        Case "总帐"
            ShowReportSumAccount
        Case "明细"
            showReportMXZ
        Case "条码"
            showReportCode
        Case "查找"
            mnuViewFind_Click
        Case "刷新"
            mnuViewRefresh_Click
        Case "总帐"
            ShowReportSumAccount
        Case "字体"
             PopupMenu mnuViewFont
        Case "前景色"
            mnuViewForeColor_Click
        Case "背景色" '
            mnuViewBackColor_Click
        Case "帮助"
            mnuHelpTitle_Click
        Case "退出"
           mnufileexit_Click
        End Select
    End With
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnuViewToolbar
    End If
End Sub

Private Sub tvwSection_S_GotFocus()
    If Me.tvwSection_S.Tag = "T" Then Me.tvwSection_S.Tag = "F"
End Sub

Private Sub tvwSection_S_NodeClick(ByVal Node As MSComctlLib.Node)
    If Me.tvwSection_S.Tag = "T" Then Exit Sub
    ReFreshStuffData Me.cob库房.ItemData(Me.cob库房.ListIndex), mstrStartDate, mstrEndDate, IIf(Left(Me.tvwSection_S.SelectedItem.Key, 1) = "R", 0, Mid(Me.tvwSection_S.SelectedItem.Key, 2))
End Sub

Private Function ReFreshTreeView() As Boolean
    '-------------------------------------------------------------------------
    '--功能:重新获取的树型结构数据
    '--参数:
    '--返回:如果数据库打开成功,则返True,否则返回False
    '-------------------------------------------------------------------------
    Dim objNode As Node
    Dim RecDept As New ADODB.Recordset
    Dim RecStuff As New ADODB.Recordset
    Dim str材质 As String
    
    ReFreshTreeView = False
    
    On Error GoTo ErrHand
    
    gstrSQL = "" & _
        "   Select distinct a.ID,a.编码 || '-' || a.名称 As 名称 " & _
        "   From 部门表 a,部门性质说明 b,部门性质分类 C " & _
        "   Where a.id=b.部门id And b.工作性质=c.名称 And C.名称 In ('制剂室', '卫材库', '发料部门', '虚拟库房') " & _
        "       And (a.站点=[2] or a.站点 is null) " & _
                IIf(InStr(1, mstrPrivs, "所有库房") <> 0, "", " And A.id In (Select 部门ID From 部门人员 Where 人员ID=[1])") & _
        " and (to_char(a.撤档时间,'yyyy-mm-dd')='3000-01-01' or a.撤档时间 is null) " & _
        " Order by a.编码 || '-' || a.名称 "
        
    Set RecDept = zlDatabase.OpenSQLRecord(gstrSQL, "所有库房", UserInfo.Id, gstrNodeNo)
    
    With RecDept
          
        If .RecordCount = 0 Then
            MsgBox "库房或发料部门体系未建立或权限不足，不能执行本程序!", vbInformation, gstrSysName
            Exit Function
        End If
        
        If InStr(1, mstrPrivs, "所有库房") <> 0 Then
            Me.cob库房.Clear
            Me.cob库房.AddItem "所有库房"
            Me.cob库房.ItemData(Me.cob库房.NewIndex) = 0
            Me.cob库房.ListIndex = Me.cob库房.NewIndex
        End If
        Do While Not .EOF
            Me.cob库房.AddItem .Fields("名称").Value
            Me.cob库房.ItemData(Me.cob库房.NewIndex) = .Fields("ID").Value
            .MoveNext
        Loop
        Me.cob库房.ListIndex = 0
    End With
    
    Set mrsTreeData = New ADODB.Recordset
    gstrSQL = "" & _
        "   Select id,上级id,编码,名称" & _
        "   From 诊疗分类目录 " & _
        "   where 类型=7" & _
        "   start with 上级id is null " & _
        "   connect by prior id=上级id " & _
        "   Order by level,id"
    
    Set mrsTreeData = zlDatabase.OpenSQLRecord(gstrSQL, "卫材分类")
    
    With mrsTreeData
        If .RecordCount = 0 Then
            MsgBox "卫材分类体系未建立，不能执行本程序!", vbInformation, gstrSysName
            Exit Function
        End If
        Me.tvwSection_S.Nodes.Clear
        tvwSection_S.Nodes.Add , , "Root", "所有分类", "root", "root"
        Do While Not .EOF
            
            If IsNull(.Fields("上级id").Value) Then
                Set objNode = Me.tvwSection_S.Nodes.Add("Root", 4, "_" & .Fields("id").Value, zlStr.NVL(!编码) & " -" & .Fields("名称").Value, "child")
            Else
                Set objNode = Me.tvwSection_S.Nodes.Add("_" & .Fields("上级id").Value, 4, "_" & .Fields("id").Value, zlStr.NVL(!编码) & " -" & .Fields("名称").Value, "child")
            End If
            .MoveNext
         Loop
         
         If tvwSection_S.Nodes(1).Children <> 0 Then
            tvwSection_S.Nodes(1).Child.Selected = True
         Else
            tvwSection_S.Nodes(1).Selected = True
         End If
    End With
    ReFreshTreeView = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Unload Me
End Function

Private Sub ReFreshStuffData(ByVal lngDeptId As Long, strStartDate As String, strEndDate As String, lngUseId As Long)
    '-------------------------------------------------------------------------
    '--功能:重新获取的材料库存数
    '--参数:
    '       lngDeptId:材料房id
    '       strStartDate:开始日期
    '       strEndDate:结束日期
    '       lngUseId:用途id值
    '--返回:
    '-------------------------------------------------------------------------
    Dim gstrSQL1 As String, strSQL As String
    Dim intRow As Long
    Dim bln批次 As Long
    Dim intCol As Long
    Dim ite As ListItem
    gstrSQL1 = ""
    
    Call FS.ShowFlash("正在查找数据,请稍候 ...", Me)
    DoEvents
    
      
   If lngDeptId = 0 Then
        Select Case mbytUint
        Case 0
            gstrSQL = ",Q.计算单位 as 单位,'' as 上次采购价,decode(Q.是否变价,1,decode(m.上次售价,Null, m.指导零售价,m.上次售价),nvl(P.现价,0)) as 最后售价," & _
            " 1 as 系数,Sum(B.可用数量) As 可用数量,Sum(B.实际数量) As 实际数量" & _
                      ",Sum(B.实际金额) As 实际金额,Sum(B.实际差价) As 实际差价" & _
                      ",Decode(To_Char(Q.撤档时间,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(   Q.撤档时间,'yyyy-MM-dd')) 撤档时间,1 as 除数, g.名称 as 供应商 "
            
            gstrSQL1 = " Group by M.材料ID,L.分类id,Q.编码,Q.名称,Q.规格,Q.是否变价,Q.产地,M.库房分批,Q.计算单位,p.现价,nvl(P.现价,0)" & _
                       ",nvl(M.换算系数,0),m.上次售价,m.指导零售价,Decode(To_Char(Q.撤档时间,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(Q.撤档时间,'yyyy-MM-dd')), g.名称 order by q.编码 "
            
        Case Else
            gstrSQL = ",M.包装单位 as 单位,'' as 上次采购价,decode(Q.是否变价,1,decode(m.上次售价,Null, m.指导零售价,m.上次售价),nvl(P.现价,0))*nvl(M.换算系数,0) as 最后售价, " & _
            " nvl(M.换算系数,0) as 系数" & _
                      ",Sum(B.可用数量/Decode(M.换算系数,0,1,null,1,M.换算系数)) as 可用数量" & _
                      ",Sum(B.实际数量/Decode(M.换算系数,0,1,null,1,M.换算系数)) as 实际数量,Sum(B.实际金额) As 实际金额" & _
                      ",Sum(B.实际差价) As 实际差价,Decode(To_Char(q.撤档时间,'yyyy-MM-dd'),'3000-01-01',' '" & _
                      ",To_Char(q.撤档时间,'yyyy-MM-dd')) 撤档时间,Decode(M.换算系数,0,1,null,1,M.换算系数) as 除数, g.名称 as 供应商 "
            
            gstrSQL1 = " Group by M.材料ID,l.分类id,Q.编码,Q.名称,Q.规格,Q.是否变价,Q.产地,M.库房分批,M.包装单位,p.现价,nvl(P.现价,0)*nvl(M.换算系数,0)" & _
                       ",nvl(M.换算系数,0),m.上次售价,m.指导零售价,Decode(M.换算系数,0,1,null,1,M.换算系数)" & _
                       ",Decode(To_Char(Q.撤档时间,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(Q.撤档时间,'yyyy-MM-dd')), g.名称 order by q.编码 "
        End Select
    Else
        Select Case mbytUint
        Case 0
            gstrSQL = ",Q.计算单位 as 单位,Decode(M.库房分批,0,Avg(S.上次采购价),Null) as 上次采购价,decode(Q.是否变价,1" & _
                      ",decode(m.上次售价,Null, m.指导零售价,m.上次售价) ,nvl(P.现价,0)) as 最后售价,1 as 系数" & _
                      ",Sum(S.可用数量) as 可用数量, Sum(S.实际数量) as 实际数量,Sum(S.实际金额) as 实际金额,Sum(S.实际差价) as 实际差价" & _
                      ",Decode(To_Char(q.撤档时间,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(q.撤档时间,'yyyy-MM-dd')) 撤档时间,1 as 除数, g.名称 as 供应商 "
            
            gstrSQL1 = " Group by M.材料ID,Q.编码,L.分类id,Q.名称,Q.规格,Q.是否变价,Q.产地,nvl(M.最大效期,0),M.库房分批,Q.计算单位,p.现价" & _
                       ",nvl(P.现价,0),nvl(M.换算系数,0),m.上次售价,m.指导零售价,Decode(To_Char(Q.撤档时间,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(Q.撤档时间,'yyyy-MM-dd')), g.名称 order by q.编码 "
        Case Else
            gstrSQL = ",M.包装单位 as 单位,Decode(M.库房分批,0,Avg(S.上次采购价*nvl(M.换算系数,0)),Null) as 上次采购价,decode(Q.是否变价,1," & _
                      "decode(m.上次售价,Null, m.指导零售价,m.上次售价) ,nvl(P.现价,0))*nvl(M.换算系数,0) as 最后售价," & _
                      "nvl(M.换算系数,0) as 系数,Sum(S.可用数量 /Decode(M.换算系数,0,1,null,1,M.换算系数)) as 可用数量" & _
                      ",Sum(S.实际数量 /Decode(M.换算系数,0,1,null,1,M.换算系数)) as 实际数量,Sum(S.实际金额) as 实际金额" & _
                      ",Sum(S.实际差价) as 实际差价,Decode(To_Char(q.撤档时间,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(q.撤档时间,'yyyy-MM-dd')) 撤档时间" & _
                      ",Decode(M.换算系数,0,1,null,1,M.换算系数) as 除数, g.名称 as 供应商 "
            
            gstrSQL1 = " Group by M.材料ID,Q.编码,L.分类id,Q.名称,Q.规格,Q.是否变价,Q.产地,nvl(M.最大效期,0),M.库房分批,M.包装单位,P.现价" & _
                       ",M.换算系数,m.上次售价,m.指导零售价,nvl(P.现价,0)*nvl(M.换算系数,0),Decode(To_Char(Q.撤档时间,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(Q.撤档时间,'yyyy-MM-dd')), g.名称 order by q.编码 "
        End Select
    End If
    
    Set mrsData = New ADODB.Recordset
    On Error GoTo ErrHand:
    
    If lngDeptId = 0 Then
        strSQL = "Select  M.材料ID,L.分类id,Q.编码,Q.名称 as 名称,Q.规格,Q.产地,Null as 效期,Decode(M.库房分批,1,'是','否') as 库房分批 " & gstrSQL & _
                " From 材料特性 M,收费项目目录 Q,诊疗项目目录 L ,收费价目 P ,供应商 G, " & _
                "     (SELECT a.库房id, a.药品id, a.批次, a.效期, a.性质, a.可用数量, a.实际数量, a.实际金额, a.实际差价, a.上次供应商id, a.上次采购价, a.上次批号, a.上次生产日期" & _
                "           ,a.上次产地, a.灭菌效期, a.批准文号, a.零售价, a.上次扣率 " & _
                "       FROM 药品库存 A, 材料特性 B, 诊疗项目目录 C WHERE a.药品id = b.材料id And b.诊疗id = c.Id And a.性质=1 " & _
                IIf(tvwSection_S.SelectedItem.Key = "Root", "", "    And C.分类id in ( Select id From 诊疗分类目录 Q start with Q.id= [1] connect by prior id=上级id)") & " ) B " & _
                " Where m.上次供应商id = g.Id(+) And P.收费细目Id=M.材料id and M.诊疗ID=L.id And (Q.站点=[3] or Q.站点 is null) " & _
                "       And P.收费细目id=Q.id " & IIf(mbln包含停用, "", " and (TO_CHAR(Q.撤档时间, 'yyyy-mm-dd') = '3000-01-01' OR Q.撤档时间 IS NULL) ") & _
                "       And sysdate between P.执行日期 And nvl(P.终止日期,To_Date('3000-01-01','yyyy-MM-DD')) and  M.材料id=B.药品id(+)  " & _
                GetPriceClassString("P") & _
                IIf(mbln库存数, " And  B.实际数量<>0 ", "") & IIf(tvwSection_S.SelectedItem.Key = "Root", "", "    And L.分类id in ( Select id From 诊疗分类目录 Q start with Q.id= [1] connect by prior id=上级id)")
        strSQL = strSQL + gstrSQL1
    Else
        strSQL = "Select M.材料ID,L.分类id,Q.编码,Q.名称 as 名称,Q.规格,Q.产地,nvl(M.最大效期,0) as 效期,Decode(M.库房分批,1,'是','否') as 库房分批 " & gstrSQL & _
                " From 材料特性 M,收费项目目录 Q,诊疗项目目录 L,收费价目 P, 供应商 G, " & _
                "      (Select a.药品ID,a.上次采购价,sum(a.可用数量) as 可用数量,sum(a.实际数量) as 实际数量,sum(a.实际金额) as 实际金额,sum(a.实际差价) as 实际差价 " & _
                "       From 药品库存 A, 材料特性 B, 诊疗项目目录 C Where a.药品id = b.材料id And b.诊疗id = c.Id And a.库房id=[2] And a.性质=1 " & _
                IIf(tvwSection_S.SelectedItem.Key = "Root", "", "    And C.分类id in ( Select id From 诊疗分类目录 Q start with Q.id= [1] connect by prior id=上级id)") & _
                " Group by a.药品ID,a.上次采购价) S,(Select Distinct 收费细目id, 执行科室id From 收费执行科室 Where 执行科室id = [2]) K " & _
                " Where m.上次供应商id = g.Id(+) And M.材料id=P.收费细目Id and M.诊疗ID=L.id  and P.收费细目id=Q.id And (Q.站点=[3] or Q.站点 is null) And m.材料id=k.收费细目id " & _
                "       And M.材料id=S.药品id(+) " & IIf(mbln包含停用, "", " and (TO_CHAR(Q.撤档时间, 'yyyy-mm-dd') = '3000-01-01' OR Q.撤档时间 IS NULL)  ") & _
                "       And sysdate between P.执行日期 And nvl(P.终止日期,To_Date('3000-01-01','yyyy-MM-DD')) " & _
                GetPriceClassString("P") & _
                    IIf(mbln库存数, " And S.实际数量<>0 ", "") & IIf(tvwSection_S.SelectedItem.Key = "Root", "", " And L.分类id in ( Select id From 诊疗分类目录 Q start with Q.id= [1] connect by prior id=上级id)")
        strSQL = strSQL + gstrSQL1
    End If
    gstrSQL = strSQL
        
    Set mrsData = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lngUseId, lngDeptId, gstrNodeNo)
    
    With mrsData
        If .RecordCount = 0 Then
            mnuExcel.Enabled = False
            mnuFilePrint.Enabled = False
            mnuFilePrintView.Enabled = False
            mnuViewFind.Enabled = False
            tbrThis.Buttons.Item(1).Enabled = False
            tbrThis.Buttons.Item(2).Enabled = False
            tbrThis.Buttons.Item(6).Enabled = False
            'tbrThis.Buttons.Item(7).Enabled = False
        Else
            mnuExcel.Enabled = True
            mnuFilePrint.Enabled = True
            mnuFilePrintView.Enabled = True
            mnuViewFind.Enabled = True
            tbrThis.Buttons.Item(1).Enabled = True
            tbrThis.Buttons.Item(2).Enabled = True
            tbrThis.Buttons.Item(6).Enabled = True
           ' tbrThis.Buttons.Item(7).Enabled = True
        End If
        Call SetReportCtrlIndexEnabled
        
        Call FS.StopFlash
        Call SetFormat(False)
    End With
    
    With Msf材料信息_S
        .Row = 1
        Call Msf材料信息_S_EnterCell
    End With
    Exit Sub
ErrHand:
    Call FS.StopFlash
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Sub
End Sub

Private Sub RefreshBatch(lng库房ID As Long, lng材料ID As Long)
    '-------------------------------------------------------------------------
    '--功能:重新获取的卫材分批库存数
    '--参数:
    '       lng库房Id:材料房id
    '       lng材料ID:用途id值
    '--返回:
    '-------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim intRow As Long
    Dim intCol As Long
    Dim lngColor As Long
    
    Dim int分批 As Integer
    Dim int在用 As Integer
     
    On Error GoTo ErrHand
    Me.Msf分批库存_S.Redraw = False
    Me.Msf分批库存_S.Rows = 1
    gstrSQL = "Select 1 From 部门性质说明 Where 部门id=[1] And 工作性质  like '发料部门'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "发料部门判断", lng库房ID)
            
    If rsTemp.EOF Then
        int在用 = 0
    Else
        int在用 = 1
    End If

    gstrSQL = "" & _
        "   Select Decode(nvl(库房分批,0),1,Decode(Nvl(在用分批,0),1,2,1),0) As 分批 " & _
        "   From 材料特性 " & _
        "   Where 材料id=[1]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取分批性质", lng材料ID)
        
        
    '如果库房分批且在用分批（int分批=2）；仅库房分批（int分批=1）；不分批（int分批=0）
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        int分批 = Val(zlStr.NVL(rsTemp!分批))
        If lng库房ID = 0 Or (int在用 = 1 And int分批 = 2) Or (int在用 = 0 And int分批 <> 0) Then
            '是所有库房 或者 是库房且库存分批，则显示分库房分批库存
            If lng库房ID = 0 Then
                gstrSQL = "" & _
                    "   Select D.名称 as 库房,Null As 批号,Null As 失效期,0 报警,Null As 产地,Null As 最后进价," & _
                    "        Sum(S.可用数量)/" & mrsData("除数").Value & " as 可用数量," & _
                    "        Sum(S.实际数量)/" & mrsData("除数").Value & " as 实际数量," & _
                    "        Sum(S.实际金额) As 实际金额," & _
                    "        Sum(S.实际差价) As 实际差价," & _
                    "       avg(s.平均成本价)* " & mrsData("除数").Value & " as 平均成本价," & _
                    "        Null As 填制日期,Null As No,Null as 供货单位, " & _
                    "        decode(sum(s.实际数量),0,0,sum(s.实际金额)/sum(s.实际数量)) as 售价 " & _
                    "   From 药品库存 S,部门表 D  " & _
                    "   Where S.库房id=D.id And S.性质=1 And S.药品id=[2]" & _
                    "       And (S.实际数量<>0 or S.实际金额<>0 or S.实际差价<>0)" & _
                    " Group By D.名称 "
            Else
               gstrSQL = "Select D.名称 as 库房,s.上次批号 As 批号, s.效期 as 失效期, s.上次产地 As 产地,Decode(sign(Add_Months(Sysdate," & mintMonths & ")-s.效期),-1,0,1) 报警," & _
                        "        S.可用数量/" & mrsData("除数").Value & " as 可用数量,S.实际数量/" & mrsData("除数").Value & " as 实际数量,S.实际金额,S.实际差价," & _
                        "        S.上次采购价*" & mrsData("除数").Value & " as 最后进价,S.商品条码,S.内部条码,s.平均成本价*" & mrsData("除数").Value & "平均成本价," & _
                        "decode(c.是否变价,1,decode(s.批次,null,decode(s.实际数量,0,0,s.实际金额/s.实际数量),s.零售价 ),b.现价) * " & mrsData("除数").Value & " as 售价 , g.名称 As 供应商 " & _
                        " From 药品库存 S,部门表 D,材料特性 A,收费价目 B,收费项目目录 C, 供应商 G " & _
                        " Where s.上次供应商id = g.Id(+) And S.库房id=D.id  And S.药品id=A.材料id" & _
                        "       And S.药品id=[2] And S.性质=1 And S.库房id=[1] and b.收费细目id=a.材料id and B.收费细目ID=c.id  and sysdate between b.执行日期 and b.终止日期 " & _
                        GetPriceClassString("B") & " And (S.实际数量<>0 or S.实际金额<>0 or S.实际差价<>0)" & _
                        " order by D.编码"
            End If
            
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lng库房ID, lng材料ID)
            
            With rsTemp
                Me.Msf分批库存_S.Rows = .RecordCount + 1
                Do While Not .EOF
                    Me.Msf分批库存_S.TextMatrix(.AbsolutePosition, 0) = !库房
                    Me.Msf分批库存_S.TextMatrix(.AbsolutePosition, 1) = IIf(IsNull(!批号), "", !批号)
                    Me.Msf分批库存_S.TextMatrix(.AbsolutePosition, 2) = Format(!失效期, "yyyy年MM月dd日")
                    Me.Msf分批库存_S.TextMatrix(.AbsolutePosition, 3) = IIf(IsNull(!产地), "", !产地)
                    Me.Msf分批库存_S.TextMatrix(.AbsolutePosition, 4) = Format(!可用数量, mFMT.FM_数量)
                    Me.Msf分批库存_S.TextMatrix(.AbsolutePosition, 5) = Format(!实际数量, mFMT.FM_数量)
                    Me.Msf分批库存_S.TextMatrix(.AbsolutePosition, 6) = Format(!实际金额, mFMT.FM_金额)
'                    Me.Msf分批库存_S.TextMatrix(.AbsolutePosition, 7) = Format(((NVL(!实际金额, 0) - NVL(!实际差价, 0)) / IIf(NVL(!实际数量, 0) = 0, 1, NVL(!实际数量, 1))), mFMT.FM_成本价)
'                    Me.Msf分批库存_S.TextMatrix(.AbsolutePosition, 8) = Format((NVL(!实际金额, 0) - NVL(!实际差价, 0)), mFMT.FM_金额)
                    Me.Msf分批库存_S.TextMatrix(.AbsolutePosition, 7) = Format(!平均成本价, mFMT.FM_成本价) '成本价
                    Me.Msf分批库存_S.TextMatrix(.AbsolutePosition, 8) = Format(!平均成本价 * !实际数量, mFMT.FM_金额) '成本金额
                    Me.Msf分批库存_S.TextMatrix(.AbsolutePosition, 9) = Format(!实际差价, mFMT.FM_金额)
                    Me.Msf分批库存_S.TextMatrix(.AbsolutePosition, 10) = Format(!最后进价, mFMT.FM_成本价)
                    
                    If gblnCode = True And lng库房ID > 0 Then
                        Me.Msf分批库存_S.TextMatrix(.AbsolutePosition, 11) = IIf(IsNull(!商品条码), "", !商品条码)
                        Me.Msf分批库存_S.TextMatrix(.AbsolutePosition, 12) = IIf(IsNull(!内部条码), "", !内部条码)
                    End If
                    If gblnCode = True Then
                        Me.Msf分批库存_S.TextMatrix(.AbsolutePosition, 13) = Format(!售价, mFMT.FM_零售价)
                        If lng库房ID <> 0 Then
                            Me.Msf分批库存_S.ColWidth(14) = 1500
                            Me.Msf分批库存_S.TextMatrix(.AbsolutePosition, 14) = IIf(IsNull(!供应商), "", !供应商)
                        Else
                            Me.Msf分批库存_S.ColWidth(14) = 0
                        End If
                    Else
                        Me.Msf分批库存_S.TextMatrix(.AbsolutePosition, 11) = Format(!售价, mFMT.FM_零售价)
                        If lng库房ID <> 0 Then
                            Me.Msf分批库存_S.ColWidth(12) = 1500
                            Me.Msf分批库存_S.TextMatrix(.AbsolutePosition, 12) = IIf(IsNull(!供应商), "", !供应商)
                         Else
                            Me.Msf分批库存_S.ColWidth(12) = 0
                        End If
                    End If
                    
                    Me.Msf分批库存_S.RowData(.AbsolutePosition) = !报警
                    '根据记录状态的不同，进行着色
                    lngColor = IIf(!报警 = 0, glng正常, glng报警)
                    For intCol = 0 To Msf分批库存_S.Cols - 1
                        Msf分批库存_S.Col = intCol
                        Msf分批库存_S.Row = .AbsolutePosition
                        Msf分批库存_S.CellForeColor = lngColor
                    Next
                    .MoveNext
                Loop
            End With
        End If
    End If
    If lng库房ID = 0 Then
        Me.Msf分批库存_S.ColWidth(0) = 1000
        Me.Msf分批库存_S.ColWidth(1) = 0
        Me.Msf分批库存_S.ColWidth(2) = 0
        Me.Msf分批库存_S.ColWidth(3) = 0
        Me.Msf分批库存_S.ColWidth(10) = 0
        Me.Msf分批库存_S.ColWidth(11) = 0
        Me.Msf分批库存_S.ColWidth(12) = 0
    Else
        Me.Msf分批库存_S.ColWidth(0) = 0
        Me.Msf分批库存_S.ColWidth(1) = 1500
        Me.Msf分批库存_S.ColWidth(2) = 1500
        Me.Msf分批库存_S.ColWidth(3) = 1500
        Me.Msf分批库存_S.ColWidth(10) = 0
        Me.Msf分批库存_S.ColWidth(11) = 1800
        Me.Msf分批库存_S.ColWidth(12) = 1800
    End If
    If mblnCostView = False Then
        Me.Msf分批库存_S.ColWidth(7) = 0
        Me.Msf分批库存_S.ColWidth(8) = 0
        Me.Msf分批库存_S.ColWidth(9) = 0
        Me.Msf分批库存_S.ColWidth(10) = 0
    End If
    If Me.Msf分批库存_S.Rows = 1 Then
        Me.Msf分批库存_S.Visible = False
        Me.lbl分批_S.Visible = False
        Me.Msf分批库存_S.Rows = 2
    Else
        Me.Msf分批库存_S.Visible = True
        Me.lbl分批_S.Visible = True
    End If
    Me.Msf分批库存_S.FixedRows = 1
    Me.Msf分批库存_S.Redraw = True
    Call Form_Resize
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Exit Sub
End Sub

Private Function ReFreshFilterData(ByVal lngDeptId As Long, strFind As String) As Boolean
    '-------------------------------------------------------------------------
    '--功能:重新所指定的卫材
    '--参数:
    '       lngDeptId:材料房id
    '       strFind:查打条件
    '--返回:
    '-------------------------------------------------------------------------
    Dim gstrSQL1 As String
    Dim intRow As Long
    Dim bln批次 As Long
    Dim intCol As Long
    Dim rsTemp As New ADODB.Recordset
    Dim str剂型 As String
    Dim ite As ListItem
    gstrSQL1 = ""
    On Error GoTo ErrHand:
    
    Call FS.ShowFlash("正在查找数据,请稍候 ...", Me)
    DoEvents
    
    ReFreshFilterData = False
    If lngDeptId = 0 Then
        Select Case mbytUint
            Case 0
                gstrSQL = ",Q.计算单位 as 单位,0 as 上次采购价,Decode(q.是否变价, 1,decode(m.上次售价,Null, m.指导零售价,m.上次售价), Nvl(p.现价, 0)) as 最后售价,nvl(M.换算系数,0) as 系数,(B.可用数量) as 可用数量" & _
                          ",(B.实际数量) as 实际数量,(B.实际金额) as 实际金额,(B.实际差价) as 实际差价" & _
                          ",Decode(To_Char(Q.撤档时间,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(Q.撤档时间,'yyyy-MM-dd')) 撤档时间,1 as 除数, g.名称 as 供应商 "
            Case Else
                gstrSQL = ",M.包装单位 as 单位,0 as 上次采购价,Decode(q.是否变价, 1,decode(m.上次售价,Null, m.指导零售价,m.上次售价), Nvl(p.现价, 0))*nvl(M.换算系数,0) as 最后售价,nvl(M.换算系数,0) as 系数" & _
                          ",(B.可用数量/Decode(M.换算系数,0,1,null,1,M.换算系数)) as 可用数量" & _
                          ",(B.实际数量/Decode(M.换算系数,0,1,null,1,M.换算系数)) as 实际数量" & _
                          ",(B.实际金额) As 实际金额,(B.实际差价) As 实际差价" & _
                          ",Decode(To_Char(Q.撤档时间,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(Q.撤档时间,'yyyy-MM-dd')) 撤档时间" & _
                          ",Decode(M.换算系数,0,1,null,1,M.换算系数) as 除数, g.名称 as 供应商 "
        End Select
    Else
        Select Case mbytUint
        Case 0
            gstrSQL = ",Q.计算单位 as 单位,S.上次采购价,Decode(q.是否变价, 1,decode(m.上次售价,Null, m.指导零售价,m.上次售价), Nvl(p.现价, 0)) as 最后售价,nvl(M.换算系数,0) as 系数,(S.可用数量) as 可用数量" & _
                      ",(S.实际数量) as 实际数量,(S.实际金额) as 实际金额,(S.实际差价) as 实际差价" & _
                      ",Decode(To_Char(Q.撤档时间,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(Q.撤档时间,'yyyy-MM-dd')) 撤档时间,1 as 除数, g.名称 as 供应商  "
        Case Else
            gstrSQL = ",M.包装单位 as 单位,S.上次采购价*nvl(M.换算系数,0) as 上次采购价,Decode(q.是否变价, 1,decode(m.上次售价,Null, m.指导零售价,m.上次售价), Nvl(p.现价, 0))*nvl(M.换算系数,0) as 最后售价" & _
                      ",nvl(M.换算系数,0) as 系数,S.可用数量 /Decode(M.换算系数,0,1,null,1,M.换算系数) as 可用数量" & _
                      ",S.实际数量 /Decode(M.换算系数,0,1,null,1,M.换算系数) as 实际数量,S.实际金额 as 实际金额,S.实际差价 as 实际差价" & _
                      ",Decode(To_Char(Q.撤档时间,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(Q.撤档时间,'yyyy-MM-dd')) 撤档时间" & _
                      ",Decode(M.换算系数,0,1,null,1,M.换算系数) as 除数, g.名称 as 供应商 "
        End Select
    End If
    
    If lngDeptId = 0 Then
       gstrSQL = "" & _
            "   Select distinct B.库房ID,M.材料ID,B.批次,L.分类id,Q.编码,Q.名称 as 名称,Q.规格,Q.产地,nvl(M.最大效期,0) as 效期" & _
            "       ,Decode(M.库房分批,1,'是','否') as 库房分批 " & gstrSQL & _
            "   From 材料特性 M,收费项目目录 Q,诊疗项目目录 L ,收费价目 P ,供应商 G, " & _
            "       (select 库房id, 药品id, 批次, 效期, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 上次供应商id, 上次采购价" & _
            "             ,上次批号, 上次生产日期, 上次产地, 灭菌效期, 批准文号, 零售价, 上次扣率 from 药品库存 where 性质=1) B " & _
            "   Where m.上次供应商id = g.Id(+) And P.收费细目Id=M.材料id  and M.诊疗id=L.id and P.收费细目id=Q.id And (Q.站点=[6] or Q.站点 is null) And M.材料id=B.药品id(+) " & _
                        IIf(mbln包含停用, "", " and (TO_CHAR(Q.撤档时间, 'yyyy-mm-dd') = '3000-01-01' OR Q.撤档时间 IS NULL) ") & _
            "           And sysdate between P.执行日期 And nvl(P.终止日期,To_Date('3000-01-01','yyyy-MM-DD'))  " & _
            GetPriceClassString("P") & _
                        IIf(mbln库存数, "  And B.实际数量<>0 ", "") & IIf(strFind = "", "", " And " & strFind)
        
        gstrSQL = "" & _
            "   SELECT 材料ID,分类id,编码,名称,规格,产地,效期,库房分批,单位,max(上次采购价) 上次采购价,最后售价," & _
            "           系数,SUM(可用数量) AS 可用数量 , SUM(实际数量) AS 实际数量 ,SUM(实际金额) AS 实际金额," & _
            "           SUM(实际差价) AS 实际差价,撤档时间,除数,供应商 " & _
            "   From (" & gstrSQL & ") " & _
            " GROUP BY 材料ID,分类id,编码,名称,规格,产地,效期,库房分批,最后售价,系数 ,撤档时间, 除数,单位,供应商 Order By 编码 "
    Else
       gstrSQL = "" & _
            "   Select Distinct M.材料ID,S.批次,Q.编码,L.分类id,Q.名称 as 名称,Q.规格,Q.产地,nvl(M.最大效期,0) as 效期" & _
            "       ,Decode(M.库房分批,1,'是','否') as 库房分批 " & gstrSQL & _
            "   From 材料特性 M,收费项目目录 Q,诊疗项目目录 L ,收费价目 P ,供应商 G, " & _
            "       (Select 库房ID,药品id 材料ID,批次,上次采购价,sum(可用数量) as 可用数量, sum(实际数量) as 实际数量,sum(实际金额) as 实际金额" & _
            "           ,sum(实际差价) as 实际差价 From 药品库存 Where 库房id+0=[1] And 性质=1 " & _
            "         Group by 库房ID,药品ID,批次,上次采购价) S " & _
            "   Where m.上次供应商id = g.Id(+) And P.收费细目Id=M.材料id and M.诊疗id=L.id and P.收费细目id=Q.id And (Q.站点=[6] or Q.站点 is null) " & _
            "       And M.材料id=S.材料id(+)" & IIf(mbln包含停用, "", " and (TO_CHAR(Q.撤档时间, 'yyyy-mm-dd') = '3000-01-01' OR Q.撤档时间 IS NULL) ") & _
            "       And sysdate between P.执行日期 And nvl(P.终止日期,To_Date('3000-01-01','yyyy-MM-DD')) " & _
            GetPriceClassString("P") & _
                    IIf(mbln库存数, " And S.实际数量<>0 And S.实际数量 is not null ", "") & " And " & strFind
        
       gstrSQL = "" & _
            "   SELECT 材料ID,分类id,编码,名称,规格,产地,效期,库房分批,单位,Decode(库房分批,'否',Avg(上次采购价),Null) 上次采购价,最后售价," & _
             "      系数,SUM(可用数量) AS 可用数量 , SUM(实际数量) AS 实际数量 ,SUM(实际金额) AS 实际金额," & _
            "       SUM(实际差价) AS 实际差价,撤档时间,除数,供应商 " & _
            "   from (" & gstrSQL & ") " & _
            "   GROUP BY 材料ID,编码,分类id,名称,规格,产地,效期,库房分批,最后售价,系数 ,撤档时间, 除数,单位 ,供应商 Order By 编码 "

    End If
        
   '0-编码,1-名称,2-简码,3-规格,4-产地,5-指定产地,6-站点
    '参数:[1]库房,[2]-编码 ,[3]-名称,[4]-简码,[5]-规格,[6]-站点,[7]-指定产地
        
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lngDeptId, mstrOthers(0), mstrOthers(1), mstrOthers(2), mstrOthers(3), gstrNodeNo, mstrOthers(5))
    
    Call FS.StopFlash
    With rsTemp
        If .RecordCount = 0 Then
            MsgBox "在指定条件中的数据不存在!", vbInformation, gstrSysName
            Exit Function
        End If
        ReFreshFilterData = True
        Set mrsData = rsTemp
      
        Call SetFormat(False)
    End With
    
    With Msf材料信息_S
        .Row = 1
        Call Msf材料信息_S_EnterCell
    End With
    Exit Function
ErrHand:
    Call FS.StopFlash
    If ErrCenter() = 1 Then Resume
    ReFreshFilterData = False
    Exit Function
End Function


'-------------------------------
'
'报表
'
'''''''''''''''''''''''''''''''''

Private Sub mnuBill_Click()
    Dim strNo As String
    Dim byt单据 As Integer
    Dim byt记录状态 As Integer
          
    Select Case Mid(mstrNoS, 4)
        Case "_INSIDE_1721_1"  '总帐
            strNo = Mid(Trim(mobjCurSheet.TextMatrix(mobjCurSheet.Row, 3)), 3)
            byt单据 = Val(mobjCurSheet.TextMatrix(mobjCurSheet.Row, 1))
            byt记录状态 = Val(mobjCurSheet.TextMatrix(mobjCurSheet.Row, 4))
        Case "_INSIDE_1721_2"  '明细帐
            strNo = Trim(mobjCurSheet.TextMatrix(mobjCurSheet.Row, 3))
            byt单据 = Val(mobjCurSheet.TextMatrix(mobjCurSheet.Row, 2))
            byt记录状态 = Val(mobjCurSheet.TextMatrix(mobjCurSheet.Row, 1))
        Case "_INSIDE_1721_3"  '明细表
        
    End Select
    
    If strNo = "" Or byt单据 = 0 Or byt记录状态 = 99 Then Exit Sub
    If byt单据 = 0 Then Exit Sub
    ShowBill Me, strNo, byt记录状态, byt单据
End Sub

Private Sub mobjReport_ReportActive(ByVal strNo As String, Form As Object)
    mlngCurReport = Form.hwnd
    mstrNoS = strNo
End Sub


Private Sub mobjReport_SheetDblClick(ByVal strNo As String, Sheet As Object, frmParent As Object)
    mlngCurReport = frmParent.hwnd
    mstrNoS = strNo
    Set mobjCurSheet = Sheet
    If Mid(UCase(strNo), 4) = "_INSIDE_1723_3" Then Exit Sub
    mnuBill_Click
End Sub

Private Sub mobjReport_SheetMouseDown(ByVal strNo As String, Button As Integer, Shift As Integer, x As Single, y As Single, Sheet As Object, frmParent As Object)
    mlngCurReport = frmParent.hwnd
    mstrNoS = strNo
    Set mobjCurSheet = Sheet
    If Mid(UCase(strNo), 4) <> "_INSIDE_1723_3" Then
        If Button = 2 Then PopupMenu mnuReportBill, 2
    End If
End Sub

Private Sub SetMenu(ByVal intState As Integer)
    If intState = 0 Then mnuReportBill.Visible = False: Exit Sub
End Sub

Private Sub ShowBill(frmObject As Object, strNo As String, int记录状态 As Integer, int单据 As Integer, Optional bln在用 As Boolean = False)
    '--------------------------------------------------------------------------------------
    '功能:显示指定单据
    '参数:
    '       frmObject:窗体
    '           strNo:单据号
    '     int记录状态:单据状态(mod(记录状态,3)=1-正常记录;mod(记录状态,3)=2-冲销记录;mod(记录状态,3)=0-已经冲销的记录)
    '         int单据:单据类别( 库房:1-外购入库单;2-其它入库;3-移库单;4-领用;5-其它出库;6-盘存;7-更换单;
    '                           在用:1-领用;2-销售;3-报废单;4-权属变更)
    '                           15-材料外购入库,16-材料自制入库,17-材料其他入库,18-材料差价调整,19-材料移库,20-部门材料领用,21-材料其他出库,22-材料盘点，23-材料盘点记录单；24-收费处方发料；25-记帐单处方发料；26-记帐表处方发料；
    '--------------------------------------------------------------------------------------
    Dim strPrivsTemp As String
    
    On Error GoTo ErrHandle
    Select Case int单据
        Case 15
            strPrivsTemp = GetPrivFunc(glngSys, 1712)
            frmPurchaseCard.ShowCard frmObject, strNo, 4, int记录状态, strPrivsTemp
        Case 16
            strPrivsTemp = GetPrivFunc(glngSys, 1713)
            frmSelfMakeCard.ShowCard frmObject, strNo, 4, int记录状态, strPrivsTemp
        Case 17
            strPrivsTemp = GetPrivFunc(glngSys, 1714)
            frmOtherInputCard.ShowCard frmObject, strNo, 4, int记录状态, strPrivsTemp
        Case 18
            strPrivsTemp = GetPrivFunc(glngSys, 1715)
            frmDiffPriceAdjustCard.ShowCard frmObject, strNo, 4, int记录状态, strPrivsTemp
        Case 19
            strPrivsTemp = GetPrivFunc(glngSys, 1716)
            frmTransferCard.ShowCard frmObject, strNo, 4, int记录状态, strPrivsTemp
        Case 20
            strPrivsTemp = GetPrivFunc(glngSys, 1717)
            frmDrawCard.ShowCard frmObject, strNo, 4, int记录状态, strPrivsTemp
        Case 21
            strPrivsTemp = GetPrivFunc(glngSys, 1718)
            frmOtherOutputCard.ShowCard frmObject, strNo, 4, int记录状态, strPrivsTemp
        Case 22
            strPrivsTemp = GetPrivFunc(glngSys, 1719)
            frmCheckCard.ShowCard frmObject, strNo, 4, int记录状态, strPrivsTemp
        Case 13
            Dim rsTemp As New ADODB.Recordset
            gstrSQL = "Select id,单据,NO,nvl(价格id,0) as 价格id" & _
                " From 药品收发记录" & _
                " Where No=[1] And 单据=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取价格记录ID", strNo, int单据)
            
            With rsTemp
                If .EOF Or .BOF Then Exit Sub
            End With
            gstrUserName = UserInfo.用户名
            Call frmStuffPrice.ShowBill(frmObject, B_查阅, Val(zlStr.NVL(rsTemp!价格id)), 0)
'            With frmStuffPrice
'                .mlngBillId = rsTemp!价格id
'                .mlngStuffId = 1
'                .Show 1, frmObject
'            End With
        Case Else
            With Frm单据See
                .int记录状态 = int记录状态
                .byt单据 = int单据
                .strNo = strNo
                .Show 1, frmObject
            End With
    End Select
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub 设置权限()

    Dim lngCount As Long
    Dim i As Long
    
    lngCount = 0        '统计是否存在相关菜单数据
    tbrThis.Buttons("明细").Visible = False
    tbrThis.Buttons("总帐").Visible = False
    For i = 0 To mnuReportItem.UBound
        If Split(mnuReportItem(i).Tag & ",", ",")(1) = "ZL1_INSIDE_1721_2" Then
            tbrThis.Buttons("明细").Visible = True
            lngCount = lngCount + 1
        End If
        If Split(mnuReportItem(i).Tag & ",", ",")(1) = "ZL1_INSIDE_1721_1" Then
            tbrThis.Buttons("总帐").Visible = True
            lngCount = lngCount + 1
        End If
    Next
End Sub
Private Function ISCheckReport(ByVal strReportCode As String) As Boolean
    '功能:检查指定报表是否有权限
    '参数:strReportCode-报表编号
    Dim i As Long
    
    For i = 0 To mnuReportItem.UBound
        If Split(mnuReportItem(i).Tag & ",", ",")(1) = strReportCode Then
            ISCheckReport = mnuReportItem(i).Enabled And mnuReport.Visible
            Exit Function
        End If
    Next
    ISCheckReport = False
End Function

Private Sub SetFormat(ByVal BlnSetHeader As Boolean)
    On Error Resume Next
    
    Dim intCol As Integer
    With Msf材料信息_S
        .Clear
        .Rows = 2
        .Cols = 19
        .TextMatrix(0, 0) = "材料ID"
        .TextMatrix(0, 1) = "分类ID"
        .TextMatrix(0, 2) = "编码"
        .TextMatrix(0, 3) = "名称"
        .TextMatrix(0, 4) = "规格"
        .TextMatrix(0, 5) = "产地"
        .TextMatrix(0, 6) = "效期"
        .TextMatrix(0, 7) = "库房分批"
        .TextMatrix(0, 8) = "单位"
        .TextMatrix(0, 9) = "上次采购价"
        .TextMatrix(0, 10) = "最后售价"
        .TextMatrix(0, 11) = "系数"
        .TextMatrix(0, 12) = "可用数量"
        .TextMatrix(0, 13) = "库存数量"
        .TextMatrix(0, 14) = "库存金额"
        .TextMatrix(0, 15) = "库存差价"
        .TextMatrix(0, 16) = "撤档时间"
        .TextMatrix(0, 17) = "除数"
        .TextMatrix(0, 18) = "上次供应商"
        If Not BlnSetHeader Then
            If mrsData.RecordCount = 0 Then Exit Sub
            Call DataBound
        End If
        
        For intCol = 0 To .Cols - 1
            .ColAlignmentFixed(intCol) = 4
        Next
        
        If BlnSetHeader Then
            If mblnFirst Then
                .ColWidth(0) = 0
                .ColWidth(1) = 0
                .ColWidth(2) = 1000
                .ColWidth(3) = 2000
                .ColWidth(4) = 900
                .ColWidth(5) = 1400
                .ColWidth(6) = 0
                .ColWidth(7) = 800
                .ColWidth(8) = 800
                If Me.cob库房.ItemData(Me.cob库房.ListIndex) = -1 Or Me.cob库房.ItemData(Me.cob库房.ListIndex) = 0 Then
                    .ColWidth(9) = 0
                Else
                    .ColWidth(9) = IIf(mblnCostView = False, 0, 1000)
                End If
                .ColWidth(10) = 1000
                .ColWidth(11) = 0
                .ColWidth(12) = 1000
                .ColWidth(13) = 1000
                .ColWidth(14) = 1000
                .ColWidth(15) = IIf(mblnCostView = False, 0, 1000)
                .ColWidth(16) = 0
                .ColWidth(17) = 0
                .ColWidth(18) = 1500
            End If
        Else
            .ColWidth(0) = 0
            .ColWidth(1) = 0
            .ColWidth(6) = 0
            If Me.cob库房.ItemData(Me.cob库房.ListIndex) = -1 Or Me.cob库房.ItemData(Me.cob库房.ListIndex) = 0 Then
                .ColWidth(9) = 0
            Else
                .ColWidth(9) = IIf(mblnCostView = False, 0, 1000)
            End If
            .ColWidth(11) = 0
            .ColWidth(15) = IIf(mblnCostView = False, 0, 1000)
            .ColWidth(16) = 0
            .ColWidth(17) = 0
            
            .ColAlignment(2) = 1
            .ColAlignment(3) = 1
            .ColAlignment(4) = 1
            .ColAlignment(9) = 7
            .ColAlignment(10) = 7
            .ColAlignment(11) = 7
            .ColAlignment(12) = 7
            .ColAlignment(13) = 7
            .ColAlignment(14) = 7
            .ColAlignment(15) = 7
        End If
        .Row = 1
    End With
End Sub

Private Sub DataBound()
    Dim lngColor As Long
    Dim lngRow As Long, lngCol As Long
    
    If mrsData.RecordCount <> 0 Then mrsData.MoveFirst
    With Msf材料信息_S
        .Redraw = False
        mblnColor = True
        Do While Not mrsData.EOF
            If mrsData.AbsolutePosition > .Rows - 1 Then .Rows = .Rows + 1
            .Row = mrsData.AbsolutePosition
            '填充数据
            .TextMatrix(.Row, 0) = mrsData!材料ID
            .TextMatrix(.Row, 1) = mrsData!分类id
            .TextMatrix(.Row, 2) = mrsData!编码
            .TextMatrix(.Row, 3) = mrsData!名称
            .TextMatrix(.Row, 4) = zlStr.NVL(mrsData!规格, "")
            .TextMatrix(.Row, 5) = zlStr.NVL(mrsData!产地, "")
            .TextMatrix(.Row, 6) = zlStr.NVL(mrsData!效期, "")
            .TextMatrix(.Row, 7) = zlStr.NVL(mrsData!库房分批, "否")
            .TextMatrix(.Row, 8) = zlStr.NVL(mrsData!单位, "")
            .TextMatrix(.Row, 9) = Format(mrsData!上次采购价, mFMT.FM_成本价)
            .TextMatrix(.Row, 10) = Format(mrsData!最后售价, mFMT.FM_零售价)
            .TextMatrix(.Row, 11) = Format(mrsData!系数, GFM_VBXS)
            .TextMatrix(.Row, 12) = Format(mrsData!可用数量, mFMT.FM_数量)
            .TextMatrix(.Row, 13) = Format(mrsData!实际数量, mFMT.FM_数量)
            .TextMatrix(.Row, 14) = Format(mrsData!实际金额, mFMT.FM_金额)
            .TextMatrix(.Row, 15) = Format(mrsData!实际差价, mFMT.FM_金额)
            .TextMatrix(.Row, 16) = zlStr.NVL(mrsData!撤档时间, "")
            .TextMatrix(.Row, 17) = zlStr.NVL(mrsData!除数, 1)
            .TextMatrix(.Row, 18) = zlStr.NVL(mrsData!供应商, "")
            '上色
            If mbln包含停用 Then
                lngColor = IIf(Trim(.TextMatrix(.Row, 16)) = "", MLNG黑色, MLNG红色)
                For lngCol = 0 To .Cols - 1
                    .Col = lngCol
                    .CellForeColor = lngColor
                Next
            End If
            mrsData.MoveNext
        Loop
        .Redraw = True
        mblnColor = False
    End With
End Sub

Private Sub txt材料信息_GotFocus()
    Call zlControl.TxtSelAll(txt材料信息)
End Sub

Private Sub txt材料信息_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strFind As String
    Dim strTemp As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txt材料信息.Text) = "" Then Exit Sub
    
    '0-编码,1-名称,2-简码,3-规格,4-产地,5-指定产地
    '参数:[1]库房,[2]-编码 ,[3]-名称,[4]-简码,[5]-规格,[6]-产地,[7]-指定产地

    txt材料信息.Text = Replace(txt材料信息.Text, "'", "")
    strTemp = GetMatchingSting(txt材料信息.Text)
    mstrOthers(0) = strTemp
    mstrOthers(1) = strTemp
    mstrOthers(2) = strTemp
    mstrOthers(3) = strTemp
'    strFind = "(Q.名称 like [3] "
'    strFind = strFind & " Or Q.编码 like [2] "
'    strFind = strFind & " Or M.材料id in (Select 收费细目ID from 收费项目别名  where 简码 like [4] ))"
    
    strFind = " M.材料id in (Select Distinct a.Id " & _
        " From 收费项目目录 A, 收费项目别名 B " & _
        " Where a.Id = b.收费细目id And (a.名称 Like [3] Or a.编码 Like [2] Or 简码 Like [4]) "
    
    If gblnCode = True Then
        strFind = strFind & " Union All " & _
        " Select 药品id From 药品库存 " & _
        " Where 性质 = 1 And 库房id + 0 = [1] And (商品条码 Like [2] Or 内部条码 Like [2])) "
    Else
        strFind = strFind & ") "
    End If
    
    If Not ReFreshFilterData(cob库房.ItemData(cob库房.ListIndex), strFind) Then Exit Sub
    
    Me.tvwSection_S.Tag = "T"
End Sub



Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub



