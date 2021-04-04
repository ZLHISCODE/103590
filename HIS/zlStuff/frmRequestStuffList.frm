VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRequestStuffList 
   Caption         =   "卫材申领管理"
   ClientHeight    =   4980
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9615
   Icon            =   "frmRequestStuffList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4980
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picSeparate_s 
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   180
      MousePointer    =   7  'Size N S
      ScaleHeight     =   360
      ScaleWidth      =   4815
      TabIndex        =   0
      Top             =   2700
      Width           =   4815
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "查询范围:1999年8月12日至1999年9月12日"
         Height          =   180
         Left            =   0
         TabIndex        =   4
         Top             =   200
         Width           =   3690
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         Caption         =   "成本金额："
         Height          =   180
         Left            =   0
         TabIndex        =   3
         Top             =   20
         Width           =   900
      End
      Begin VB.Label lbl2 
         AutoSize        =   -1  'True
         Caption         =   "售价金额："
         Height          =   180
         Left            =   1890
         TabIndex        =   2
         Top             =   20
         Width           =   900
      End
      Begin VB.Label lbl3 
         AutoSize        =   -1  'True
         Caption         =   "差价金额："
         Height          =   180
         Left            =   3690
         TabIndex        =   1
         Top             =   20
         Width           =   900
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDetail 
      Height          =   975
      Left            =   360
      TabIndex        =   5
      Top             =   3120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1720
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   0
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin ComCtl3.CoolBar cbrTool 
      Height          =   780
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   11775
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tlbTool"
      MinWidth1       =   6000
      MinHeight1      =   720
      Width1          =   6210
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Caption2        =   "申领部门"
      Child2          =   "cboStock"
      MinWidth2       =   3000
      MinHeight2      =   300
      Width2          =   3345
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar tlbTool 
         Height          =   720
         Left            =   165
         TabIndex        =   8
         Top             =   30
         Width           =   7515
         _ExtentX        =   13256
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
            NumButtons      =   16
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
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Hank"
                     Text            =   "手工填写"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Auto"
                     Text            =   "自动生成"
                  EndProperty
               EndProperty
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
               Key             =   "Edit1"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "核查"
               Key             =   "Check"
               Object.ToolTipText     =   "核查"
               Object.Tag             =   "核查"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "接收"
               Key             =   "Receive"
               Object.ToolTipText     =   "接收"
               Object.Tag             =   "接收"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "冲销"
               Key             =   "DisReceive"
               Object.ToolTipText     =   "冲销"
               Object.Tag             =   "冲销"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "EditSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "过滤"
               Key             =   "Search"
               Description     =   "过滤"
               Object.ToolTipText     =   "过滤"
               Object.Tag             =   "过滤"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "刷新"
               Key             =   "Refresh"
               Description     =   "刷新"
               Object.ToolTipText     =   "刷新"
               Object.Tag             =   "刷新"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FindSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助主题"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Exit"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   11
            EndProperty
         EndProperty
         MouseIcon       =   "frmRequestStuffList.frx":014A
         Begin VB.Timer LimitTime 
            Enabled         =   0   'False
            Interval        =   8000
            Left            =   6660
            Top             =   180
         End
      End
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   8685
         TabIndex        =   7
         Text            =   "cboStock"
         Top             =   240
         Width           =   3000
      End
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   9
      Top             =   4620
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmRequestStuffList.frx":0464
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11880
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
      Left            =   0
      Top             =   600
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
            Picture         =   "frmRequestStuffList.frx":0CF8
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":0F18
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":1138
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":1354
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":1574
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":1794
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":19B0
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":1BCC
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":1DE6
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":1F40
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":215C
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":237C
            Key             =   "check"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHot 
      Left            =   600
      Top             =   600
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
            Picture         =   "frmRequestStuffList.frx":2596
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":27B6
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":29D6
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":2BF2
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":2E12
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":3032
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":324E
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":346A
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":3684
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":37DE
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":39FE
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffList.frx":3C1E
            Key             =   "check"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid mshList 
      Height          =   885
      Left            =   240
      TabIndex        =   10
      Top             =   1320
      Width           =   4935
      _cx             =   8705
      _cy             =   1561
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
      ForeColorSel    =   0
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
      FormatString    =   ""
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
         Shortcut        =   {F12}
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
         Begin VB.Menu mnuEditAddHank 
            Caption         =   "手工填写(&H)"
         End
         Begin VB.Menu mnuEditAddAuto 
            Caption         =   "自动生成(&A)"
         End
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
      Begin VB.Menu mnuEditCheck 
         Caption         =   "核查(&V)"
      End
      Begin VB.Menu mnuEditReceive 
         Caption         =   "接受(&R)"
      End
      Begin VB.Menu mnuEditDisReceive 
         Caption         =   "冲销(&D)"
      End
      Begin VB.Menu mnuEditLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditImport 
         Caption         =   "导入申购单(&I)"
         Visible         =   0   'False
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
Attribute VB_Name = "frmRequestStuffList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFind As String
Private mblnBootUp As Boolean
Private mlastRow As Long                '上次电击的行
Private mintPreCol As Integer           '前一次单据头的排序列
Private mintsort As Integer             '前一次单据头的排序
Private mintPreDetailCol As Integer     '前一次单据体的排序列
Private mintDetailsort As Integer       '前一次单据体的排序

Private mdtStartDate As Date
Private mdtEndDate As Date
Private mdtVerifyStart As Date
Private mdtVerifyEnd As Date
Private mstrPrivs As String
Private mintUnit As String
Private mstrOthers() As String ' 0-记录状态(计划类型),1-开始单号,2-结束单号,3-材料id,4-对方部门id(或入出类别id或编制方法(计划用)),5-填制人,6-审核人,7-供应商ID,8-生产商,9-开始生产日期,10-结束生产日期,11-开始发票号,12-结束发票号,13-条码信息
Private mlngModule As Long
Private mblnCostView As Boolean             '查看成本价相关信息 true-允许查看 false-不允许查看
Private Const mstrCaption As String = "卫材申领管理"
Private mbln申领核查 As Boolean     '单据是否需要核查 true-需要 false-不需要
Private mintFindDay As Integer      '查询天数范围
Private mint明确批次 As Integer             '表示在填写申领单时，是否明确卫材的批次
Private mint冲销申请 As Integer                          '0-不需要申请;1-需要申请


'----------------------------------------------------------------------------------------------------------
'刘兴宏:增加小数位数的格式串
'修改:2007/03/06
Private mOraFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------


Private Sub cboStock_Click()
    If mblnBootUp Then mnuViewRefresh_Click
End Sub

Private Sub cboStock_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cboStock.ListCount = 0 Then Call zlControl.ControlSetFocus(mshList): Exit Sub
    
    If cboStock.ListIndex >= 0 Then
        If Val(cboStock.Tag) = cboStock.ItemData(cboStock.ListIndex) Then
            Call zlControl.ControlSetFocus(mshList, True)
            Exit Sub
        End If
    End If
    
    If Select部门选择器(Me, cboStock, Trim(cboStock.Text), "W,V,K", True) = False Then
        Exit Sub
    End If
    If cboStock.ListIndex >= 0 Then
        cboStock.Tag = cboStock.ItemData(cboStock.ListIndex)
    End If
End Sub

Private Sub cbrTool_Resize()
    If mblnBootUp = False Then Exit Sub
    Form_Resize
End Sub

Public Sub ShowList(ByVal frmMain As Variant)
    Dim strFind As String
    
    mblnBootUp = False
    If Not CheckDepend Then Exit Sub            '数据依赖性测试
                
    SetVisable  '根据权限设置不同的显示项目
    mintFindDay = Val(zlDatabase.GetPara("查询天数", glngSys, mlngModule, 1))
    mdtStartDate = Format(DateAdd("d", -mintFindDay, sys.Currentdate), "yyyy-MM-dd")
    mdtEndDate = Format(sys.Currentdate, "yyyy-MM-dd")
    mdtVerifyStart = "1901-01-01"
    mdtVerifyEnd = "1901-01-01"
    
    strFind = " AND A.记录状态 = 1 And A.审核日期 is Null And A.填制日期 Between To_Date('" & Format(mdtStartDate, "yyyy-mm-dd") & " 00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(mdtEndDate, "yyyy-mm-dd") & " 23:59:59','YYYY-MM-DD HH24:MI:SS')"
    mstrFind = strFind
    
    GetList (mstrFind)  '列出单据头
    RestoreWinState Me, App.ProductName, mstrCaption
    '恢复个性化参数设置后，还需要对权限控制的列进一步设置
    With mshList
        .ColWidth(.ColIndex("成本金额")) = IIf(mblnCostView = True, 900, 0) '发料部门人员是否可以看成本价
        .ColWidth(.ColIndex("差价金额")) = IIf(mblnCostView = True, 900, 0)
    End With
    
    With mshDetail
        .ColWidth(12) = IIf(mblnCostView = True, 900, 0)
        .ColWidth(13) = IIf(mblnCostView = True, 900, 0)
        .ColWidth(16) = IIf(mblnCostView = True, 900, 0)
    End With
    
    mblnBootUp = True
    
    If IsObject(frmMain) Then
        Me.Show , frmMain
    Else
        zlCommFun.ShowChildWindow Me.hwnd, frmMain
    End If
    
    Me.ZOrder 0

End Sub

'检查数据依赖性
Private Function CheckDepend() As Boolean
    
    Dim rsDepend As New Recordset
    Dim strStock As String
    
    CheckDepend = False

    On Error GoTo ErrHandle
    strStock = " And B.名称 In('制剂室','卫材库','发料部门')"
    
    gstrSQL = "" & _
        "   SELECT DISTINCT a.id, a.名称 " & _
        "   FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " & _
        "   Where c.工作性质 = b.名称 And (A.站点=[1] or A.站点 is null) " & strStock & _
        "           AND a.id = c.部门id " & _
        "           AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01'"
    Set rsDepend = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, gstrNodeNo)
    If rsDepend.EOF Then
        MsgBox "部门性质信息不全,请查看部门管理！", vbInformation, gstrSysName
        rsDepend.Close
        Exit Function
    End If
    rsDepend.Close
    
    gstrSQL = "" & _
        "   SELECT DISTINCT a.id, a.名称 " & _
        "   FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " & _
        "   Where c.工作性质 = b.名称 And (A.站点=[2] or A.站点 is null) " & strStock & _
        "           AND a.id = c.部门id " & _
        "           and a.id in (select 部门id from 部门人员 where 人员id= [1]) " & _
        "           AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01'"
    Set rsDepend = zlDatabase.OpenSQLRecord(gstrSQL, "卫材申领管理", UserInfo.Id, gstrNodeNo)
    If rsDepend.EOF Then
        MsgBox "你不是卫材库、发料部门、或制剂室的工作人员，不能进入！", vbInformation, gstrSysName
        rsDepend.Close
        Exit Function
    End If
    
    With cboStock
        .Clear
        Do While Not rsDepend.EOF
            .AddItem rsDepend!名称
            .ItemData(.NewIndex) = rsDepend!Id
            If rsDepend!Id = glngDeptId Then
                .ListIndex = .NewIndex
            End If
            rsDepend.MoveNext
        Loop
        rsDepend.Close
        
        If .ListIndex = -1 Then
            .ListIndex = 0
        End If
    End With
    
    CheckDepend = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetList(ByVal strFind As String)
    Dim rsList As New Recordset
    Dim strUserPart As String
    Dim intCol As Integer
    
    '用于统计合计金额
    Dim dbl1 As Double
    Dim dbl2 As Double
    Dim dbl3 As Double
    Dim n As Long
    Dim strFormat As String
    
    On Error GoTo ErrHandle
    strFormat = "0.00##"
    
    mlastRow = 0
    
    mshList.Redraw = False
    strUserPart = " And A.库房ID+0=[1]"
    
    gstrSQL = "SELECT A.NO, C.名称 AS 发料库房,LTRIM(TO_CHAR (SUM (A.成本金额), " & mOraFMT.FM_金额 & ")) AS 成本金额, " & _
        " LTRIM(TO_CHAR ( (SUM (A.零售金额)), " & mOraFMT.FM_金额 & ")) AS 售价金额,LTRIM(TO_CHAR (SUM (A.零售金额 - A.成本金额)," & mOraFMT.FM_金额 & ")) AS 差价金额, A.填制人, " & _
        " TO_CHAR (MIN(A.填制日期), 'YYYY-MM-DD HH24:MI:SS') AS 填制日期, " & _
        " a.核查人, To_Char(Min(a.核查日期), 'YYYY-MM-DD HH24:MI:SS') As 核查日期," & _
        " A.审核人, TO_CHAR (MIN(A.审核日期), 'YYYY-MM-DD HH24:MI:SS') AS 审核日期, A.记录状态, A.配药人 发送人,A.摘要,nvl(a.发药方式,0) as 发药方式 " & _
        " FROM 药品收发记录 A, 部门表 B,部门表 C " & _
        " WHERE A.库房ID = B.ID AND A.对方部门ID=C.ID AND A.单据 = 19 AND  A.入出系数=1 " & _
        " And (A.配药人 Is NULL Or A.配药日期 Is Not NULL)" & _
        strUserPart & strFind & _
        " GROUP BY A.NO,C.名称,A.填制人,A.核查人,A.审核人,A.记录状态,A.配药人,A.摘要,nvl(a.发药方式,0) " & _
        " ORDER BY NO DESC, 填制日期 ASC "
        
     'mstrOthers(0 To 13) As String ' 0-记录状态(计划类型),1-开始单号,2-结束单号,3-材料id,4-对方部门id(或入出类别id或编制方法(计划用)),5-填制人,6-审核人,7-供应商ID,8-生产商,9-开始生产日期,10-结束生产日期,11-开始发票号,12-结束发票号,13-条码信息
    '参数范围:[1]-库房id,[2]:开始填制日期,[3]结束填制日期,[4]开始审核日期,[5] 结束审核日期,[6]-记录状态,[7]开始单据号,[8]结束单据号,[9]材料id,[10]对方部门id,[11]填制人,[12]审核人[13]-供应商ID,[14]-生产商,[15]-开始生产日期,[16]-结束生产日期,[17]-开始发票号,[18]-结束发票号,[19]-条码信息
    
    '初始生产日期
    mstrOthers(9) = IIf(Trim(mstrOthers(9)) = "", "1901-01-01", mstrOthers(9))
    mstrOthers(10) = IIf(Trim(mstrOthers(10)) = "", "1901-01-01", mstrOthers(10))
    
    Set rsList = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex), _
        CDate(Format(mdtStartDate, "yyyy-mm-dd") & " 00:00:00"), CDate(Format(mdtEndDate, "yyyy-mm-dd") & " 23:59:59"), _
        CDate(Format(mdtVerifyStart, "yyyy-mm-dd") & " 00:00:00"), CDate(Format(mdtVerifyEnd, "yyyy-mm-dd") & " 23:59:59"), _
        Val(mstrOthers(0)), mstrOthers(1), mstrOthers(2), Val(mstrOthers(3)), _
        Val(mstrOthers(4)), mstrOthers(5), mstrOthers(6), _
        Val(mstrOthers(7)), mstrOthers(8), CDate(mstrOthers(9) & " 00:00:00"), CDate(mstrOthers(10) & " 23:59:59"), _
         mstrOthers(11), mstrOthers(12), mstrOthers(13) & "%")
      
    Set mshList.DataSource = rsList
    With mshList
        If .Rows = 1 Then
            .Rows = .Rows + 100
            .Row = 1
'            .Redraw = True
            
            .TopRow = 1
            .Rows = .Rows - 99
        End If
        .Row = 1
        .Col = 0
'        .ColSel = .Cols - 1
        
        For intCol = 0 To .Cols - 1
            .ColKey(intCol) = .TextMatrix(0, intCol)
        Next
    End With
    
    SetListColWidth
    
    '统计合计金额
    If (Not rsList.EOF) And (Not rsList.BOF) Then
        rsList.MoveFirst
        Do While Not rsList.EOF
            dbl1 = dbl1 + IIf(IsNull(rsList!成本金额), 0, rsList!成本金额)
            dbl2 = dbl2 + IIf(IsNull(rsList!售价金额), 0, rsList!售价金额)
            dbl3 = dbl3 + IIf(IsNull(rsList!差价金额), 0, rsList!差价金额)
            rsList.MoveNext
        Loop
        rsList.MoveFirst
        
        lbl1.Caption = "成本金额合计：" & Format(dbl1, strFormat)
        lbl2.Caption = "售价金额合计：" & Format(dbl2, strFormat)
        lbl3.Caption = "差价金额合计：" & Format(dbl3, strFormat)
    Else
        lbl1.Caption = "成本金额合计：" & Format(0, strFormat)
        lbl2.Caption = "售价金额合计：" & Format(0, strFormat)
        lbl3.Caption = "差价金额合计：" & Format(0, strFormat)
    End If
    
    mshlist_EnterCell    '列出单据体
    
    SetStrikeColor
    mshList.Redraw = True
    Call SetEnable
    
    staThis.Panels(2).Text = "当前共有" & rsList.RecordCount & "张单据"
    rsList.Close
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub SetStrikeColor()
    Dim intStatus As Integer
    Dim intRow As Integer
    Dim intCol As Integer
    
    With mshList
        If .Rows <= 2 Then Exit Sub
        For intRow = 1 To .Rows - 1
            intStatus = Val(.TextMatrix(intRow, .ColIndex("记录状态")))
            If intStatus Mod 3 = 0 Then
                .Row = intRow
                For intCol = 0 To .Cols - 1
                    .Col = intCol
                    .CellForeColor = &H80000001
                Next
            End If
            If intStatus Mod 3 = 2 Then
                .Row = intRow
                For intCol = 0 To .Cols - 1
                    .Col = intCol
                    If .TextMatrix(intRow, .ColIndex("审核人")) = "" Then
                        .CellForeColor = &HC0C0FF
                    Else
                        .CellForeColor = &HFF
                    End If
                Next
            End If
        Next
    End With
End Sub

'表头列宽初始
Private Sub SetListColWidth()
    Dim intCol As Integer
    
    With mshList
        .ColAlignment(.ColIndex("成本金额")) = flexAlignRightCenter
        If mblnBootUp = False Then
            For intCol = 1 To .Cols - 1
                If intCol = 1 Then
                   .ColWidth(intCol) = 2000
                ElseIf intCol = .ColIndex("记录状态") Or intCol = .ColIndex("发药方式") Then
                    .ColWidth(intCol) = 0
                Else
                    .ColWidth(intCol) = 1000
                End If
            Next
        End If
        
        .ColWidth(.ColIndex("成本金额")) = IIf(mblnCostView = False, 0, 1000) '发料部门人员是否可以看成本价
        If mblnCostView = False Then
            .ColWidth(.ColIndex("差价金额")) = 0 '发料部门人员是否可以看成本价
        End If
    End With
End Sub


Private Sub SetDetailColWidth()
    Dim intCol As Integer
    Dim i As Integer
    On Error Resume Next
    
    With mshDetail
        .ColAlignment(8) = flexAlignRightCenter     '灭菌效期
        .ColAlignment(9) = flexAlignRightCenter     '填写数量
        .ColAlignment(10) = flexAlignCenterCenter    '实际数量
        .ColAlignment(11) = flexAlignRightCenter     '单位
        .ColAlignment(12) = flexAlignRightCenter     '成本价
        .ColAlignment(13) = flexAlignRightCenter    '成本金额
        .ColAlignment(14) = flexAlignRightCenter    '售价
        .ColAlignment(15) = flexAlignRightCenter    '售价金额
        .ColAlignment(16) = flexAlignRightCenter    '差价
                
        If mblnBootUp = False Then
            .ColWidth(0) = 0
            .ColWidth(1) = 2500
            For intCol = 2 To .Cols - 1
                .ColWidth(intCol) = 1000
            Next
            .ColWidth(16) = 0
        End If
        
        .ColWidth(12) = IIf(mblnCostView = False, 0, 1000)
        .ColWidth(13) = IIf(mblnCostView = False, 0, 1000)
        .ColWidth(16) = IIf(mblnCostView = False, 0, 1000)
    End With
End Sub


'根据权限设置不同的显示项目
Private Sub SetVisable()
    '基本，申领
    If mbln申领核查 = False Then
        mnuEditCheck.Visible = False
        tlbTool.Buttons("Check").Visible = False
    Else
        mnuEditCheck.Visible = True
        tlbTool.Buttons("Check").Visible = True
    End If
    
    If Not zlStr.IsHavePrivs(gstrPrivs, "申领") Then
        mnuEditAdd.Visible = False
        mnuEditModify.Visible = False
        mnuEditDel.Visible = False
        
        tlbTool.Buttons("Add").Visible = False
        tlbTool.Buttons("Modify").Visible = False
        tlbTool.Buttons("Delete").Visible = False
        tlbTool.Buttons("Edit1").Visible = False
        mnuEditLine1.Visible = False
    End If
    If Not zlStr.IsHavePrivs(gstrPrivs, "审核") Then
        mnuEditReceive.Visible = False
        tlbTool.Buttons("Receive").Visible = False
    End If
    If Not zlStr.IsHavePrivs(gstrPrivs, "冲销") Then
        mnuEditDisReceive.Visible = False
        If mnuEditReceive.Visible = False Then mnuEditLine2.Visible = False
        tlbTool.Buttons("DisReceive").Visible = False
        tlbTool.Buttons("EditSeparate").Visible = mnuEditLine2.Visible
    End If
    If Not zlStr.IsHavePrivs(gstrPrivs, "单据打印") Then
        mnuFileBillPrint.Visible = False
        mnuFileBillPreview.Visible = False
    End If
    
End Sub

Private Sub Form_Activate()
    If mint冲销申请 = 1 Then
        mnuEditDisReceive.Caption = "申请冲销(&R)"
        tlbTool.Buttons("DisReceive").Caption = "申请冲销"
    Else
        mnuEditDisReceive.Caption = "冲销(&D)"
        tlbTool.Buttons("DisReceive").Caption = "冲销"
    End If
End Sub

Private Sub Form_Load()
    Dim strReg As String
    Dim strOthers(0 To 13) As String
    Dim i As Integer
    mlngModule = glngModul
    mbln申领核查 = IIf((zlDatabase.GetPara("申领需要核查后才能移库", glngSys, mlngModule, "0")) = 0, False, True)
    
    '取移库的冲销申请参数
    mint冲销申请 = Val(zlDatabase.GetPara("冲销申请", glngSys, 1716))

    '如果申领按批次申领，则不能使用"导入采购单"功能。
    mint明确批次 = IIf(IS批次申领, 1, 0)
'    If mint明确批次 = 1 Then
'        mnuEditImport.Visible = False
'    End If
    
    For i = 0 To 13
        strOthers(i) = ""
    Next
    '设置生产日期
    strOthers(9) = "1901-01-01"
    strOthers(10) = "1901-01-01"
    
    '0-记录状态(计划类型),1-开始单号,2-结束单号,3-材料id,4-对方部门id(或入出类别id或编制方法(计划用)),5-填制人,6-审核人,7-供应商ID,8-生产商,9-开始生产日期,10-结束生产日期,11-开始发票号,12-结束发票号,13-条码信息
    mstrOthers = strOthers

    lblRange.Caption = "查询范围:" & Format(sys.Currentdate, "yyyy年MM月dd日") & "至" & Format(sys.Currentdate, "yyyy年MM月dd日")
    
    mstrPrivs = gstrPrivs
    mblnCostView = zlStr.IsHavePrivs(mstrPrivs, "查看成本价")
    If Not zlStr.IsHavePrivs(mstrPrivs, "参数设置") Then
        mnuFileParameter.Visible = False
    Else
        mnuFileParameter.Visible = True
    End If
    
    strReg = Val(zlDatabase.GetPara("卫材单位", glngSys, mlngModule, "0"))
    mintUnit = Val(strReg)
  
    '刘兴宏:增加小数格式化串
    With mOraFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价, True)
        .FM_金额 = GetFmtString(mintUnit, g_金额, True)
        .FM_零售价 = GetFmtString(mintUnit, g_售价, True)
        .FM_数量 = GetFmtString(mintUnit, g_数量, True)
    End With
    
    
    lbl1.Caption = ""
    lbl2.Caption = ""
    lbl3.Caption = ""
    lbl2.Left = lbl1.Left + lbl1.Width + 3000
    lbl3.Left = lbl2.Left + lbl2.Width + 3000
    If mblnCostView = False Then
        lbl1.Visible = False
        lbl3.Visible = False
    End If
   
   '2006-04-25:刘兴宏,统一增加报表发布到模块的功能
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
     
End Sub

Private Sub Form_Resize()
    '窗体位置设置
    
    On Error Resume Next
    If Me.WindowState <> vbMaximized Then
        If Me.Height < 8145 Then
            Me.Height = 8145
        End If
    End If
    
    cbrTool.Bands(1).MinHeight = tlbTool.Height
    With cbrTool
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth - .Left
    End With
    
    With picSeparate_s
        .Height = 360
        .Left = 0
        .Width = cbrTool.Width
        
    End With
    
    With mshList
        .Top = IIf(cbrTool.Visible, cbrTool.Height, 0)
        .Left = 0
        .Width = cbrTool.Width
        .Height = picSeparate_s.Top - .Top
    End With
    
    With mshDetail
        .Top = picSeparate_s.Top + picSeparate_s.Height + 100
        .Left = 0
        .Height = ScaleHeight - .Top - IIf(staThis.Visible, staThis.Height, 0)
        .Width = cbrTool.Width
    End With
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, mstrCaption
End Sub

Private Sub mnuEditAddAuto_Click()
    '自动生成
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    strNo = ""
    frmRequestStuffCard.ShowCard Me, strNo, 5, , mstrPrivs, blnSuccess, cboStock.ItemData(cboStock.ListIndex)
    If blnSuccess = True Then
        mnuViewRefresh_Click
    End If
    
End Sub

Private Sub mnuEditAddHank_Click()
    '手工填写
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    strNo = ""
    '新增
    frmRequestStuffCard.ShowCard Me, strNo, 1, , mstrPrivs, blnSuccess, cboStock.ItemData(cboStock.ListIndex)
    If blnSuccess = True Then
        mnuViewRefresh_Click
    End If
    
End Sub

Private Sub mnuEditCheck_Click()
    '核查，验收
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    blnSuccess = False
    With mshList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        strNo = .TextMatrix(.Row, 0)
        
        frmRequestStuffCard.ShowCard Me, strNo, 3, mshList.TextMatrix(mshList.Row, mshList.ColIndex("记录状态")), mstrPrivs, blnSuccess, cboStock.ItemData(cboStock.ListIndex)
        If blnSuccess = True Then
            mnuViewRefresh_Click
        End If
    End With
End Sub

Private Sub mnuEditDel_Click()
    '删除
    Dim StrBillNo As String
    Dim intRow As Integer
    Dim strTitle As String
    Dim intReturn As Integer
    Dim intRecord As Integer
    
    With mshList
        If .TextMatrix(intRow, .Cols - 4) = "" And Val(.TextMatrix(.Row, .ColIndex("记录状态"))) = 1 Then
            If Not Check申领(StrBillNo) Then
                MsgBox "你没有权限删除移库单！", vbInformation, gstrSysName
                Exit Sub
            End If
        
            strTitle = "药品申领单"
        ElseIf Val(.TextMatrix(.Row, .Cols - 3)) Mod 3 = 2 And mint冲销申请 = 1 Then
            strTitle = "冲销申请单"
        End If
        
        On Error GoTo ErrHandle
        intRow = .Row
        StrBillNo = .TextMatrix(intRow, 0)
        
        intReturn = MsgBox("你确实要删除单据号为“" & StrBillNo & "”的" & strTitle & "吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
        intRecord = .Rows - 1
        If intReturn = vbYes Then
            gstrSQL = "zl_材料移库_Delete('" & StrBillNo & "'," & Val(.TextMatrix(.Row, .ColIndex("记录状态"))) & " )"
            If gstrSQL = "" Then Exit Sub
            Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption & "-删除申领单")
            intRecord = intRecord - 1
            mlastRow = 0
            If .Rows > 2 Then
                .RemoveItem intRow
            ElseIf .Rows = 2 Then
                .Rows = 3
                .RemoveItem intRow
                With mshDetail
                    .Rows = 1
                    .Rows = 2
                    .FixedRows = 1
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                End With
                SetEnable
                
            End If
                
            '.RowHeight(intRow) = 0
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
'            .ColSel = .Cols - 1
            mshlist_EnterCell
        End If
    End With
    staThis.Panels(2).Text = "当前共有" & intRecord & "张单据"
    Exit Sub

ErrHandle:
    If ErrCenter() = 1 Then Resume 'Resume这种情况不用调用
    Call SaveErrLog
    
End Sub

Private Sub mnuEditDisplay_Click()
    '查看单据
    
    Dim strNo As String
    With mshList
        strNo = .TextMatrix(.Row, 0)
        frmRequestStuffCard.ShowCard Me, strNo, 4, .TextMatrix(.Row, .ColIndex("记录状态")), mstrPrivs, , cboStock.ItemData(cboStock.ListIndex)
    End With
End Sub


'Modified By 朱玉宝 2003-12-10 地区：泸州
Private Sub mnuEditDisReceive_Click()
    Dim strNo As String, blnSuccess As Boolean
    Dim int处理方式 As Integer
    
    If mnuEditDisReceive.Caption = "申请冲销(&R)" Then
        int处理方式 = 1
    Else
        int处理方式 = 0
    End If
    
    With mshList
        strNo = .TextMatrix(.Row, 0)
        frmRequestStuffCard.ShowCard Me, strNo, 7, .TextMatrix(.Row, .ColIndex("记录状态")), mstrPrivs, blnSuccess, cboStock.ItemData(cboStock.ListIndex), int处理方式
        If Not blnSuccess Then Exit Sub
    End With
    Call mnuViewRefresh_Click
End Sub

Private Sub mnuEditImport_Click()
    Dim blnSuccess As Boolean
    
    frmPurchaseImportFromPlane.ShowCard Me, cboStock.Text, cboStock.ItemData(cboStock.ListIndex), mintUnit, InStr(mstrPrivs, "所有库房") <> 0, blnSuccess, 1, 1722, mint明确批次
    
    If blnSuccess = True Then
        mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditModify_Click()
    '修改
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    blnSuccess = False
    With mshList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        strNo = .TextMatrix(.Row, 0)
        If Not Check申领(strNo) Then
            MsgBox "你没有权限修改移库单！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        frmRequestStuffCard.ShowCard Me, strNo, 2, mshList.TextMatrix(.Row, mshList.ColIndex("记录状态")), mstrPrivs, blnSuccess, cboStock.ItemData(cboStock.ListIndex)
        If blnSuccess = True Then
            mnuViewRefresh_Click
        End If
    End With
End Sub

Private Sub mnuEditReceive_Click()
    Dim strNo As String, blnSuccess As Boolean
    With mshList
        strNo = .TextMatrix(.Row, 0)
        frmRequestStuffCard.ShowCard Me, strNo, 6, .TextMatrix(.Row, .ColIndex("记录状态")), mstrPrivs, blnSuccess, cboStock.ItemData(cboStock.ListIndex)
    End With
    If blnSuccess = True Then
        mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuFileBillPreview_Click()
    With mshList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        ReportOpen gcnOracle, glngSys, "zl1_bill_1722", Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & .TextMatrix(.Row, .ColIndex("记录状态")), "单位系数=" & mintUnit, 1
    End With
End Sub

Private Sub mnuFileBillPrint_Click()
    
    With mshList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        ReportOpen gcnOracle, glngSys, "zl1_bill_1722", Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & .TextMatrix(.Row, .ColIndex("记录状态")), "单位系数=" & mintUnit, 2
    End With
End Sub

Private Sub mnuFileExcel_Click()
    '输出到Excel
    If Me.ActiveControl Is mshList Then
        mshList.Redraw = False
        subPrint 3
        mshList.Redraw = True
        mshList.Col = 0
'        mshList.ColSel = mshList.Cols - 1
    ElseIf Me.ActiveControl Is mshDetail Then
        mshDetail.Redraw = False
        subExcel 3
        mshDetail.Redraw = True
        mshDetail.Col = 0
        mshDetail.ColSel = mshDetail.Cols - 1
    End If
End Sub

Private Sub mnufileexit_Click()
    '退出
    Unload Me
    
End Sub

Private Sub mnuFileParameter_Click()
    Dim strReg As String
    '参数设置
'    frmRequestPara.设置参数 mlngModule, Me, mstrCaption, mstrPrivs
    frmParaset.设置参数 mlngModule, mstrPrivs, Me, mstrCaption
    
    strReg = Val(zlDatabase.GetPara("卫材单位", glngSys, mlngModule, "0"))
    mintUnit = Val(strReg)
    mbln申领核查 = IIf((zlDatabase.GetPara("申领需要核查后才能移库", glngSys, mlngModule, "0")) = 0, False, True)
    If mbln申领核查 = False Then
        mnuEditCheck.Visible = False
        tlbTool.Buttons("Check").Visible = False
    Else
        mnuEditCheck.Visible = True
        tlbTool.Buttons("Check").Visible = True
    End If
    mintFindDay = Val(zlDatabase.GetPara("查询天数", glngSys, mlngModule, 1))
    mdtStartDate = Format(DateAdd("d", -mintFindDay, sys.Currentdate), "yyyy-MM-dd")
    mdtEndDate = Format(sys.Currentdate, "yyyy-MM-dd")
    
    Call GetList(mstrFind)
End Sub

Private Sub mnuFilePreView_Click()
    '打印预览
    mshList.Redraw = False
    subPrint 2
    mshList.Redraw = True
    mshList.Col = 0
'    mshList.ColSel = mshList.Cols - 1
    
End Sub

Private Sub mnuFilePrint_Click()
    '打印
    mshList.Redraw = False
    subPrint 1
    mshList.Redraw = True
    mshList.Col = 0
'    mshList.ColSel = mshList.Cols - 1
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
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name)
End Sub

Private Sub mnuHelpWebHome_Click()
    '中联主页
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    '发送反馈
    Call zlMailTo(Me.hwnd)
End Sub

Private Sub mnuViewRefresh_Click()
    '刷新
    GetList mstrFind
End Sub

Private Sub mnuViewSearch_Click()
    
   '查找
    
    Dim strCon As String
    Dim strFind As String
    Dim strOthers() As String
    
    strFind = FrmTransferSearch.GetSearch(Me, 1716, mdtStartDate, mdtEndDate, mdtVerifyStart, mdtVerifyEnd, mstrPrivs, strOthers)
    
    If strFind <> "" Then
        mstrFind = strFind
        mstrOthers = strOthers
        GetList mstrFind
        If Format(mdtStartDate, "yyyy-mm-dd") = "1901-01-01" And Format(mdtVerifyStart, "yyyy-mm-dd") = "1901-01-01" Then
            lblRange.Visible = False
        ElseIf Format(mdtStartDate, "yyyy-mm-dd") <> "1901-01-01" And Format(mdtVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
            lblRange.Visible = True
            lblRange = "查询范围:填制日期 " & Format(mdtStartDate, "yyyy年MM月dd日") & "至" & Format(mdtEndDate, "yyyy年MM月dd日") & "  审核日期 " & Format(mdtVerifyStart, "yyyy年MM月dd日") & "至" & Format(mdtVerifyEnd, "yyyy年MM月dd日")
        ElseIf Format(mdtStartDate, "yyyy-mm-dd") <> "1901-01-01" Then
            lblRange.Visible = True
            lblRange = "查询范围:填制日期 " & Format(mdtStartDate, "yyyy年MM月dd日") & "至" & Format(mdtEndDate, "yyyy年MM月dd日")
        ElseIf Format(mdtVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
            lblRange.Visible = True
            lblRange = "查询范围:审核日期 " & Format(mdtVerifyStart, "yyyy年MM月dd日") & "至" & Format(mdtVerifyEnd, "yyyy年MM月dd日")
        End If
    End If
    
End Sub

Private Sub mnuViewStatus_Click()
    With mnuViewStatus
        .Checked = Not .Checked
        staThis.Visible = .Checked
    End With
    
    Form_Resize
End Sub
Private Sub mnuReportItem_Click(Index As Integer)
    Dim strNo As String
    Dim intRecodeSta As Integer
    Dim lng库房ID As Long
    Dim lngCol As Long
    
    With mshList
        strNo = Trim(.TextMatrix(.Row, 0))
        lngCol = GetCol(mshList, "记录状态")
        If lngCol < 0 Then
            intRecodeSta = 1
        Else
            intRecodeSta = Val(.TextMatrix(.Row, lngCol))
        End If
    End With
    
    If cboStock.ListIndex < 0 Then
        lng库房ID = 0
    Else
        lng库房ID = cboStock.ItemData(cboStock.ListIndex)
    End If
    
    '2006-04-25:刘兴宏:增加自定义报表发布到模块的功能
    If Format(mdtStartDate, "yyyy-mm-dd") = "1990-01-01" Then
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, "NO=" & strNo, "记录状态=" & intRecodeSta, "申领部门=" & lng库房ID)
    Else
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, "NO=" & strNo, "记录状态=" & intRecodeSta, "申领部门=" & lng库房ID, "开始时间=" & Format(mdtStartDate, "yyyy-mm-dd"), "结束时间=" & Format(mdtEndDate, "yyyy-mm-dd"))
    End If
End Sub
Private Sub mnuViewToolButton_Click()
    With mnuViewToolButton
        .Checked = Not .Checked
        cbrTool.Visible = .Checked
        mnuViewToolText.Enabled = .Checked
    End With
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim intCount As Integer      '工具条索引
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    With tlbTool.Buttons
        If mnuViewToolText.Checked = False Then
            '取消所有的文本标签显示
            For intCount = 1 To .Count
                .Item(intCount).Caption = ""
            Next
        Else
            '让所有的文本标签显示。说明：Tag中放的文本标签
            For intCount = 1 To .Count
                .Item(intCount).Caption = .Item(intCount).Tag
            Next
        End If
    End With
    
    cbrTool.Bands(1).MinHeight = tlbTool.Height
    
    Form_Resize
End Sub

Private Sub mshDetail_Click()
    With mshDetail
        If .Row < 1 Or .TextMatrix(.Row, 0) = "" Then Exit Sub
        If .MouseRow = 0 Then
            DetailSort          '列排序
            Exit Sub
        End If
    End With
End Sub

Private Sub mshList_Click()
    With mshList
         If .Row < 1 Then Exit Sub
         If .MouseRow = 0 Then
            ListSort
            Exit Sub
         End If
    End With
End Sub

Private Sub mshlist_DblClick()
    If mnuEditModify.Visible = False Then Exit Sub
    If mnuEditModify.Enabled = False Then Exit Sub
    If mshList.MouseRow = 0 Then Exit Sub
    mnuEditModify_Click
End Sub

Private Sub mshlist_EnterCell()
    Dim rsDetail As New Recordset
    Dim strUnitQuantity As String               '单位和数量格式化串
    Dim IntBill As Integer                      '单据类型  如：1、外购入库；2、
    Dim strUnit As String                       '单位名称:如门诊单位，住院单位等
    Dim str包装系数 As String
    Dim strOrder As String
    Dim strCompare As String
    
    On Error GoTo ErrHandle

    If mlastRow = mshList.Row Or LTrim(mshList.TextMatrix(mshList.Row, 0)) = "" Then
        If LTrim(mshList.TextMatrix(mshList.Row, 0)) = "" Then
            With mshDetail
                .Cols = IIf(gblnCode = True, 21, 19)
                .Rows = 2
                
                .Clear
                .TextMatrix(0, 0) = "序号"
                .TextMatrix(0, 1) = "卫材信息"
                .TextMatrix(0, 2) = "材料来源"
                .TextMatrix(0, 3) = "规格"
                .TextMatrix(0, 4) = "产地"
                .TextMatrix(0, 5) = "批准文号"
                .TextMatrix(0, 6) = "批号"
                .TextMatrix(0, 7) = "效期"
                .TextMatrix(0, 8) = "灭菌效期"
                .TextMatrix(0, 9) = "填写数量"
                .TextMatrix(0, 10) = "实际数量"
                .TextMatrix(0, 11) = "单位"
                .TextMatrix(0, 12) = "成本价"
                .TextMatrix(0, 13) = "成本金额"
                .TextMatrix(0, 14) = "售价"
                .TextMatrix(0, 15) = "售价金额"
                .TextMatrix(0, 16) = "差价"
                .TextMatrix(0, 17) = "库房货位"
                .TextMatrix(0, 18) = "名称"
                
                If gblnCode = True Then
                    .TextMatrix(0, 19) = "商品条码"
                    .TextMatrix(0, 20) = "内部条码"
                End If
                
                .Row = 1
                .Col = 0
                .ColSel = .Cols - 1
            End With
        End If
        Exit Sub
    End If
    mlastRow = mshList.Row
    SetEnable
    
    strOrder = zlDatabase.GetPara("单据排序", glngSys, mlngModule, "00")
    strCompare = Mid(strOrder, 1, 1)
    
    If mshList.Row >= 1 And LTrim(mshList.TextMatrix(mshList.Row, 0)) <> "" Then
        mshList.Col = 0
'        mshList.ColSel = mshList.Cols - 1
        
        mshDetail.Redraw = False
        
        Select Case mintUnit
            Case 0
                str包装系数 = "1"
            Case Else
                str包装系数 = "B.换算系数"
        End Select
            
        gstrSQL = "" & _
            "   SELECT * " & _
            "   FROM (  " & _
            "           SELECT DISTINCT 序号,('['||D.编码||']'||D.名称) AS 卫材信息,B.材料来源," & _
            "                       D.规格,A.产地,A.批准文号, A.批号, A.效期,to_char(a.灭菌效期,'yyyy-mm-dd') as 灭菌效期," & _
            "                       (TO_CHAR(A.填写数量 /" & str包装系数 & "," & mOraFMT.FM_数量 & ")) AS 填写数量," & _
            "                       (TO_CHAR(A.实际数量 /" & str包装系数 & "," & mOraFMT.FM_数量 & ")) AS 实际数量," & _
                                    IIf(mintUnit = 0, "D.计算单位", "b.包装单位") & "  AS 单位," & _
            "                       TO_CHAR (A.成本价*" & str包装系数 & "," & mOraFMT.FM_成本价 & ") AS 成本价," & _
            "                       TO_CHAR (A.成本金额, " & mOraFMT.FM_金额 & ") AS 成本金额," & _
            "                       TO_CHAR (A.零售价*" & str包装系数 & "," & mOraFMT.FM_零售价 & ") AS 售价," & _
            "                       TO_CHAR (A.零售金额, " & mOraFMT.FM_金额 & ") AS 售价金额," & _
            "                       TO_CHAR (A.差价," & mOraFMT.FM_金额 & ") AS 差价 ,C.库房货位 ,NVL(E.名称,D.名称) as 名称 "
            
        If gblnCode = True Then
            gstrSQL = gstrSQL & " ,A.商品条码,A.内部条码 "
        End If
        
        gstrSQL = gstrSQL & _
            "           FROM 药品收发记录 A, 材料特性 B, 收费项目别名 E, 收费项目目录 D, 材料储备限额 C " & _
            "           WHERE A.药品ID = B.材料ID AND B.材料ID=D.ID " & _
            "                   AND B.材料ID = E.收费细目ID(+) AND E.性质(+)=3 " & _
            "                   AND A.记录状态 = [2]" & _
            "                   AND A.单据 = 19 AND 入出系数=1 " & _
            "                   AND A.NO =[1]  AND A.药品ID=C.材料ID(+) AND A.库房ID=C.库房ID(+)" & _
            "   )" & _
            " ORDER BY " & IIf(strCompare = "0", "序号", IIf(strCompare = "1", "卫材信息", IIf(strCompare = "2", "名称", "库房货位"))) & IIf(Right(strOrder, 1) = "0", " ASC", " DESC")
        
        Set rsDetail = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mshList.TextMatrix(mshList.Row, 0), Val(mshList.TextMatrix(mshList.Row, mshList.ColIndex("记录状态"))))
                
        Set mshDetail.Recordset = rsDetail
    
        
        With rsDetail
            .Close
        End With
        With mshDetail
            If .Rows = 1 Then
                .Rows = .Rows + 100
                .Row = 1
                .Redraw = True
                .TopRow = 1
                .Rows = .Rows - 99
            End If
            .Row = 1
            .Col = 0
            .ColSel = .Cols - 1
        End With
        mshDetail.Redraw = True
    End If
    SetDetailColWidth
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mshlist_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If mnuEditModify.Visible = False Then Exit Sub
        If mnuEditModify.Enabled = False Then Exit Sub
        mnuEditModify_Click
    End If
        
End Sub

Private Sub mshlist_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    If mnuEdit.Visible = False Then Exit Sub
    
    PopupMenu mnuEdit, 2
    
End Sub

Private Sub picSeparate_s_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    '分割条设置
    
    If Button <> 1 Then Exit Sub
    
    With picSeparate_s
        If .Top + Y < 2000 Then Exit Sub
        If .Top + Y > ScaleHeight - 2000 Then Exit Sub
        .Move .Left, .Top + Y
    End With
    
    With mshList
        .Top = IIf(cbrTool.Visible, cbrTool.Height, 0)
        .Height = picSeparate_s.Top - .Top
    End With
    
    With mshDetail
        .Top = picSeparate_s.Top + picSeparate_s.Height + 100
        .Height = ScaleHeight - .Top - IIf(staThis.Visible, staThis.Height, 0)
    End With
    
End Sub

Private Sub tlbTool_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "PrintView"
            mnuFilePreView_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Add"
            mnuEditAddHank_Click
        Case "Modify"
            mnuEditModify_Click
        Case "Check"
            mnuEditCheck_Click
        Case "Delete"
            mnuEditDel_Click
        Case "Receive"
            mnuEditReceive_Click
        Case "DisReceive"
            mnuEditDisReceive_Click
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
    Dim bln已发送 As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    With mshList
        .ToolTipText = ""
        If .TextMatrix(.Row, 0) = "" Or .Row = 0 Then         '没有单
            mnuFilePreView.Enabled = False
            mnuFilePrint.Enabled = False
            mnuFileBillPreview.Enabled = False
            mnuFileBillPrint.Enabled = False
            mnuFileExcel.Enabled = False
            tlbTool.Buttons("Print").Enabled = False
            tlbTool.Buttons("PrintView").Enabled = False
            
            If mnuEditCheck.Visible = True Then
                mnuEditCheck.Enabled = False
                tlbTool.Buttons("Check").Enabled = False
            End If
            If mnuEditModify.Visible = True Then
                mnuEditModify.Enabled = False
                tlbTool.Buttons("Modify").Enabled = False
            End If
            If mnuEditDel.Visible = True Then
                mnuEditDel.Enabled = False
                tlbTool.Buttons("Delete").Enabled = False
            End If
            If mnuEditDisplay.Visible = True Then
                mnuEditDisplay.Enabled = False
            End If
            If mnuEditReceive.Visible Then
                mnuEditReceive.Enabled = False
                tlbTool.Buttons("Receive").Enabled = False
            End If
            If mnuEditDisReceive.Visible Then
                mnuEditDisReceive.Enabled = False
                tlbTool.Buttons("DisReceive").Enabled = False
            End If
         Else
            mnuFilePreView.Enabled = True
            mnuFilePrint.Enabled = True
            mnuFileBillPreview.Enabled = True
            mnuFileBillPrint.Enabled = True
            mnuFileExcel.Enabled = True
            tlbTool.Buttons("Print").Enabled = True
            tlbTool.Buttons("PrintView").Enabled = True
                        
            If .TextMatrix(.Row, .ColIndex("审核日期")) = "" Then    '未审核单
                bln已发送 = (mshList.TextMatrix(mshList.Row, .ColIndex("发送人")) <> "")
                If mnuEditModify.Visible = True Then
                    mnuEditModify.Enabled = True
                    tlbTool.Buttons("Modify").Enabled = True
                End If
                If mnuEditDel.Visible = True Then
                    mnuEditDel.Enabled = True
                    tlbTool.Buttons("Delete").Enabled = True
                End If
                 
                If mnuEditDisplay.Visible = True Then
                    mnuEditDisplay.Enabled = True
                End If
                
                If mnuEditModify.Visible = True Then
                    mnuEditModify.Enabled = True
                    tlbTool.Buttons("Modify").Enabled = True
                End If
                
                '如果要进行核查
                If mnuEditCheck.Visible = True Then
                    mnuEditCheck.Enabled = Not bln已发送
                    tlbTool.Buttons("Check").Enabled = Not bln已发送
                    If .TextMatrix(.Row, .ColIndex("核查日期")) = "" Then    '核查日期
                        '未核查
                        If mnuEditReceive.Visible Then
                            mnuEditReceive.Enabled = bln已发送  '未核查已发送的可以接受（移库模块填单）
                            tlbTool.Buttons("Receive").Enabled = bln已发送
                        End If
                    Else
                        '已核查
                        If mnuEditReceive.Visible Then
                            mnuEditReceive.Enabled = bln已发送
                            tlbTool.Buttons("Receive").Enabled = bln已发送
                        End If
                    End If
                Else
                '不进行核查
                    If mnuEditReceive.Visible Then
                        mnuEditReceive.Enabled = bln已发送
                        tlbTool.Buttons("Receive").Enabled = bln已发送
                    End If
                End If

                If mnuEditDisReceive.Visible Then
                    If bln已发送 Then
                        mnuEditDisReceive.Enabled = Not bln已发送
                        tlbTool.Buttons("DisReceive").Enabled = Not bln已发送
                    Else
                        mnuEditDisReceive.Enabled = False
                        tlbTool.Buttons("DisReceive").Enabled = False
                    End If
                End If
                
                '如果冲销单还未审核，则允许删除
                If mint冲销申请 = 1 Then
                    If Val(.TextMatrix(.Row, .ColIndex("记录状态"))) Mod 3 = 2 Then
                        mnuEditModify.Enabled = False
                        tlbTool.Buttons("Modify").Enabled = False
                        mnuEditReceive.Enabled = False
                        tlbTool.Buttons("Receive").Enabled = False
                        mnuEditDisReceive.Enabled = False
                        tlbTool.Buttons("DisReceive").Enabled = False
                        
                        mnuEditDel.Enabled = True
                        tlbTool.Buttons("Delete").Enabled = True
                    End If
                Else
                    If mnuEditDisReceive.Visible Then
                        If bln已发送 Then
                            mnuEditDisReceive.Enabled = Not bln已发送
                            tlbTool.Buttons("DisReceive").Enabled = Not bln已发送
                        Else
                            mnuEditDisReceive.Enabled = False
                            tlbTool.Buttons("DisReceive").Enabled = False
                        End If
                    End If
                End If
            ElseIf .TextMatrix(.Row, .ColIndex("记录状态")) = 1 Then    '审核单
                '判断是否接受（不支持已冲销单据的接受功能，必需全退或输负数的方式解决，因为要实现这个功能，需要汇总统计剩余数量）
                If mnuEditCheck.Visible = True Then
                    mnuEditCheck.Enabled = False
                    tlbTool.Buttons("Check").Enabled = False
                End If
                    
                If mnuEditModify.Visible = True Then
                    mnuEditModify.Enabled = False
                    tlbTool.Buttons("Modify").Enabled = False
                End If
                
                If mnuEditDel.Visible = True Then
                    mnuEditDel.Enabled = False
                    tlbTool.Buttons("Delete").Enabled = False
                End If
                 
                If mnuEditDisplay.Visible = True Then
                    mnuEditDisplay.Enabled = True
                End If
                If mnuEditReceive.Visible Then
                    mnuEditReceive.Enabled = False
                    tlbTool.Buttons("Receive").Enabled = False
                End If
                If mnuEditDisReceive.Visible Then
                    mnuEditDisReceive.Enabled = True
                    tlbTool.Buttons("DisReceive").Enabled = True
                End If
            Else   '2,3 冲销单
                If .TextMatrix(.Row, .ColIndex("记录状态")) Mod 3 = 0 Then
                    .ToolTipText = "冲销单据的原单据"
                    If mnuEditDisReceive.Visible = True Then
                        mnuEditDisReceive.Enabled = True
                        tlbTool.Buttons("DisReceive").Enabled = True
                    End If
                ElseIf .TextMatrix(.Row, .ColIndex("记录状态")) Mod 3 = 2 Then
                    .ToolTipText = "冲销单据"
                    If mnuEditDisReceive.Visible = True Then
                        mnuEditDisReceive.Enabled = False
                        tlbTool.Buttons("DisReceive").Enabled = False
                    End If
                End If
                If mnuEditCheck.Visible = True Then
                    mnuEditCheck.Enabled = False
                    tlbTool.Buttons("Check").Enabled = False
                End If
                If mnuEditModify.Visible = True Then
                    mnuEditModify.Enabled = False
                    tlbTool.Buttons("Modify").Enabled = False
                End If
                If mnuEditDel.Visible = True Then
                    mnuEditDel.Enabled = False
                    tlbTool.Buttons("Delete").Enabled = False
                End If
                 
                If mnuEditDisplay.Visible = True Then
                    mnuEditDisplay.Enabled = True
                End If
                If mnuEditReceive.Visible Then
                    mnuEditReceive.Enabled = False
                    tlbTool.Buttons("Receive").Enabled = False
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
    
    If Format(mdtStartDate, "yyyy-mm-dd") = "1901-01-01" And Format(mdtVerifyStart, "yyyy-mm-dd") = "1901-01-01" Then
        strRange = "审核日期 " & Format(mdtVerifyStart, "yyyy年MM月dd日") & "至" & Format(mdtVerifyEnd, "yyyy年MM月dd日")
    ElseIf Format(mdtVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
        strRange = "填制日期 " & Format(mdtStartDate, "yyyy年MM月dd日") & "至" & Format(mdtEndDate, "yyyy年MM月dd日") & "  审核日期 " & Format(mdtVerifyStart, "yyyy年MM月dd日") & "至" & Format(mdtVerifyEnd, "yyyy年MM月dd日")
    Else
        strRange = "填制日期 " & Format(mdtStartDate, "yyyy年MM月dd日") & "至" & Format(mdtEndDate, "yyyy年MM月dd日")
    End If
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
    objPrint.Title.Text = mstrCaption
        
    objRow.Add "时间：" & strRange
    
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
        
    objRow.Add "打印人:" & gstrUserName
    objRow.Add "打印日期:" & Format(sys.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    
    Set objPrint.Body = mshList
    
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
    If ButtonMenu.Index = 1 Then
        mnuEditAddHank_Click
    Else
        mnuEditAddAuto_Click
    End If
End Sub

Private Sub tlbTool_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool
    End If
End Sub


'对单据头列排序
Private Sub ListSort()
    Dim intCol As Integer
    Dim intRow As Integer
    Dim intTemp As String
    
    With mshList
        If .Rows > 1 Then
            .Redraw = False
            intCol = .MouseCol
            .Col = intCol
            .ColSel = intCol
            intTemp = .TextMatrix(.Row, 0)
            
            Select Case intCol
                Case 2
                    If intCol = mintPreCol And mintsort = flexSortNumericDescending Then
                       .Sort = flexSortNumericAscending
                       mintsort = flexSortNumericAscending
                    Else
                       .Sort = flexSortNumericDescending
                       mintsort = flexSortNumericDescending
                    End If
                Case Else
                    If intCol = mintPreCol And mintsort = flexSortStringNoCaseDescending Then
                       .Sort = flexSortStringNoCaseAscending
                       mintsort = flexSortStringNoCaseAscending
                    Else
                       .Sort = flexSortStringNoCaseDescending
                       mintsort = flexSortStringNoCaseDescending
                    End If
            End Select
            mintPreCol = intCol
            .Row = grid.MshGrdFindRow(mshList, intTemp, 0)
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

'对单据头列排序
Private Sub DetailSort()
    Dim intCol As Integer
    Dim intRow As Integer
    Dim intTemp As Integer
    
    With mshDetail
        If .Rows > 1 Then
            .Redraw = False
            intCol = .MouseCol
            .Col = intCol
            .ColSel = intCol
            intTemp = .TextMatrix(.Row, 0)
                
            Select Case intCol
                Case 6, 7, 9, 10, 11, 12, 13
                    If intCol = mintPreDetailCol And mintDetailsort = flexSortNumericDescending Then
                       .Sort = flexSortNumericAscending
                       mintDetailsort = flexSortNumericAscending
                    Else
                       .Sort = flexSortNumericDescending
                       mintDetailsort = flexSortNumericDescending
                    End If
                    
                Case Else
                    If intCol = mintPreDetailCol And mintDetailsort = flexSortStringNoCaseDescending Then
                       .Sort = flexSortStringNoCaseAscending
                       mintDetailsort = flexSortStringNoCaseAscending
                    Else
                       .Sort = flexSortStringNoCaseDescending
                       mintDetailsort = flexSortStringNoCaseDescending
                    End If
            End Select
                
            mintPreDetailCol = intCol
            .Row = grid.MshGrdFindRow(mshDetail, intTemp, 0)
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

Private Sub subExcel(bytMode As Byte)
'功能:进行打印,预览和输出到EXCEL
'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
'    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As Object
    Dim objRow As zlTabAppRow
    Dim strRange As String
    
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
    objPrint.Title.Text = mstrCaption
        
    Set objRow = New zlTabAppRow
    objRow.Add ""
    objRow.Add "NO." & Trim(mshList.TextMatrix(mshList.Row, mshList.ColIndex("no")))
    objPrint.UnderAppRows.Add objRow

    Set objRow = New zlTabAppRow
    objRow.Add "移出库房：" & mshList.TextMatrix(mshList.Row, mshList.ColIndex("发料库房"))
    objRow.Add "移入库房：" & gstrDeptName
    objPrint.UnderAppRows.Add objRow
        
    Set objRow = New zlTabAppRow
    objRow.Add "摘要:" & mshList.TextMatrix(mshList.Row, mshList.ColIndex("摘要"))
    objPrint.BelowAppRows.Add objRow
        
    Set objRow = New zlTabAppRow
    objRow.Add "填制人:" & mshList.TextMatrix(mshList.Row, mshList.ColIndex("填制人")) & "  填制日期:" & mshList.TextMatrix(mshList.Row, mshList.ColIndex("填制日期"))
    
    objRow.Add "审核人:" & mshList.TextMatrix(mshList.Row, mshList.ColIndex("审核人")) & "  审核日期:" & mshList.TextMatrix(mshList.Row, mshList.ColIndex("审核日期"))
    objPrint.BelowAppRows.Add objRow
    
    Set objPrint.Body = mshDetail
    zlPrintOrView1Grd objPrint, bytMode
End Sub

Private Function Check申领(ByVal StrBillNo As String) As Boolean
    Dim rsCheck As New ADODB.Recordset
    On Error GoTo ErrHandle
    
    '先检查是不是申领单
    gstrSQL = " Select Nvl(发药方式,0) 申领 From 药品收发记录 " & _
              " Where 单据=19 And NO=[1] And 序号=1"
    
    Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "检查是不是申领单", StrBillNo)
              
    Check申领 = Not (rsCheck!申领 = 0)
    Exit Function
ErrHandle:
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
    Call zlWebForum(Me.hwnd)
End Sub

