VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frmMainList 
   Caption         =   "药品协定入库管理"
   ClientHeight    =   4980
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9480
   Icon            =   "frmMainList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picColor 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4920
      ScaleHeight     =   255
      ScaleWidth      =   3615
      TabIndex        =   8
      Top             =   4320
      Width           =   3615
      Begin VB.PictureBox picColor3 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2280
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   11
         Top             =   0
         Width           =   260
      End
      Begin VB.PictureBox picColor2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1320
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   10
         Top             =   0
         Width           =   260
      End
      Begin VB.PictureBox picColor1 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   9
         Top             =   0
         Width           =   260
      End
      Begin VB.Label lblColor3 
         AutoSize        =   -1  'True
         Caption         =   "财务冲销"
         Height          =   180
         Left            =   2640
         TabIndex        =   14
         Top             =   30
         Width           =   720
      End
      Begin VB.Label lblColor1 
         AutoSize        =   -1  'True
         Caption         =   "正常冲销"
         Height          =   180
         Left            =   360
         TabIndex        =   13
         Top             =   37
         Width           =   720
      End
      Begin VB.Label lblColor2 
         AutoSize        =   -1  'True
         Caption         =   "正常"
         Height          =   180
         Left            =   1680
         TabIndex        =   12
         Top             =   37
         Width           =   360
      End
   End
   Begin VB.PictureBox picSeparate_s 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   120
      MousePointer    =   7  'Size N S
      ScaleHeight     =   300
      ScaleWidth      =   4815
      TabIndex        =   4
      Top             =   2760
      Width           =   4815
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "查询范围:1999年8月12日至1999年9月12日"
         Height          =   180
         Left            =   0
         TabIndex        =   5
         Top             =   120
         Width           =   3690
      End
   End
   Begin ComCtl3.CoolBar cbrTool 
      Height          =   780
      Left            =   0
      TabIndex        =   0
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
      Caption2        =   "库房"
      Child2          =   "cboStock"
      MinWidth2       =   3000
      MinHeight2      =   300
      Width2          =   3345
      NewRow2         =   0   'False
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   8685
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   3000
      End
      Begin MSComctlLib.Toolbar tlbTool 
         Height          =   720
         Left            =   165
         TabIndex        =   1
         Top             =   30
         Width           =   7785
         _ExtentX        =   13732
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
            NumButtons      =   15
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
               Key             =   "EditSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "审核"
               Key             =   "Verify"
               Description     =   "审核"
               Object.ToolTipText     =   "审核"
               Object.Tag             =   "审核"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "冲销"
               Key             =   "Strike"
               Description     =   "冲销"
               Object.ToolTipText     =   "冲销"
               Object.Tag             =   "冲销"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "VerifySeparate"
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "过滤"
               Key             =   "Search"
               Description     =   "过滤"
               Object.ToolTipText     =   "过滤"
               Object.Tag             =   "过滤"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "刷新"
               Key             =   "Refresh"
               Description     =   "刷新"
               Object.ToolTipText     =   "刷新"
               Object.Tag             =   "刷新"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FindSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助主题"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Exit"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   11
            EndProperty
         EndProperty
         MouseIcon       =   "frmMainList.frx":014A
      End
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   4620
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMainList.frx":0464
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11642
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
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":0CF8
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":0F18
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":1138
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":1354
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":1574
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":1794
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":19B0
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":1BCC
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":1DE6
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":1F40
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":215C
            Key             =   "Quit"
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
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":237C
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":259C
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":27BC
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":29D8
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":2BF8
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":2E18
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":3034
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":3250
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":346A
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":35C4
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":37E4
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   1005
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   6255
      _cx             =   11033
      _cy             =   1773
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
      ForeColorSel    =   -2147483640
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
      FormatString    =   $"frmMainList.frx":3A04
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
   Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   5655
      _cx             =   9975
      _cy             =   1720
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
      BackColorSel    =   16053482
      ForeColorSel    =   -2147483640
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
      FormatString    =   $"frmMainList.frx":3A79
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
         Shortcut        =   ^A
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
      Begin VB.Menu mnuEditVerify 
         Caption         =   "审核(&C)"
      End
      Begin VB.Menu mnuEditStrike 
         Caption         =   "冲销(&K)"
      End
      Begin VB.Menu mnuEditLine2 
         Caption         =   "-"
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
Attribute VB_Name = "frmMainList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngMode As Long
Private mstrFind As String
Private mblnBootUp As Boolean
Private mlastRow As Long                '上次电击的行
Private mstrTitle As String             '窗体的标题
Private mintPreCol As Integer           '前一次单据头的排序列
Private mintsort As Integer             '前一次单据头的排序
Private mintPreDetailCol As Integer     '前一次单据体的排序列
Private mintDetailsort As Integer       '前一次单据体的排序
Private mstrPrivs As String             '当前用户具有的当前模块的功能
Private mblnViewCost As Boolean         '查看成本价权限 true-能查看成本价 flase-不能查看成本价

Private strStart As Date
Private strEnd As Date
Private strVerifyStart As Date
Private strVerifyEnd As Date

Private Type Type_SQLCondition
    strNO开始 As String
    strNO结束 As String
    date填制时间开始 As Date
    date填制时间结束 As Date
    date审核时间开始 As Date
    date审核时间结束 As Date
    lng药品 As Long
    lng移出库房 As Long
    str填制人 As String
    str审核人 As String
End Type

Private SQLCondition As Type_SQLCondition

Private mlng库房id As Long
Private mintUnit As Integer                 '单位系数：1-售价;2-门诊;3-住院;4-药库
Private mint库存检查 As Integer                       '表示药品出库时是否进行库存检查：0-不检查;1-检查，不足提醒；2-检查，不足禁止

'从参数表中取药品价格、数量、金额小数位数（显示精度）
Private mintShowCostDigit As Integer            '成本价小数位数
Private mintShowPriceDigit As Integer           '售价小数位数
Private mintShowNumberDigit As Integer          '数量小数位数
Private mintShowMoneyDigit As Integer           '金额小数位数

Private mstrNumberFormat As String
Private mstrCostFormat As String
Private mstrPriceFormat As String
Private mstrMoneyFormat As String

Private Const mconint售价单位 As Integer = 1
Private Const mconint门诊单位 As Integer = 2
Private Const mconint住院单位 As Integer = 3
Private Const mconint药库单位 As Integer = 4
Private Sub cboStock_Click()
    If mlng库房id <> cboStock.ItemData(cboStock.ListIndex) Then
        mlng库房id = cboStock.ItemData(cboStock.ListIndex)
        Call GetDrugDigit(mlng库房id, Me.Tag, mintUnit, mintShowCostDigit, mintShowPriceDigit, mintShowNumberDigit, mintShowMoneyDigit)
        
        '组织格式化串
        mstrCostFormat = "'999999999990." & String(mintShowCostDigit, "0") & "'"
        mstrPriceFormat = "'999999999990." & String(mintShowPriceDigit, "0") & "'"
        mstrNumberFormat = "'999999999990." & String(mintShowNumberDigit, "0") & "'"
        mstrMoneyFormat = "'999999999990." & String(mintShowMoneyDigit, "0") & "'"
        
        If mblnBootUp Then mnuViewRefresh_Click
    End If
End Sub

Private Sub cbrTool_Resize()
    If mblnBootUp = False Then Exit Sub
    Form_Resize
End Sub

Public Sub ShowList(ByVal lngMode As Long, ByVal strTitle As String, ByVal FrmMain As Variant)
    Dim strFind As String
    Dim dateCurDate As Date
    Dim strTemp As String
    Dim dateCurrentDate As Date
    
    mblnBootUp = False
    mlngMode = lngMode
    mstrTitle = strTitle
    mstrPrivs = gstrprivs
    Me.Tag = strTitle
    
    If Not CheckDepend Then     '数据依赖性测试
        Unload Me
        Exit Sub
    End If
    
    mlng库房id = Me.cboStock.ItemData(Me.cboStock.ListIndex)
    Call GetDrugDigit(mlng库房id, Me.Tag, mintUnit, mintShowCostDigit, mintShowPriceDigit, mintShowNumberDigit, mintShowMoneyDigit)
    
    '组织格式化串
    mstrCostFormat = "'999999999990." & String(mintShowCostDigit, "0") & "'"
    mstrPriceFormat = "'999999999990." & String(mintShowPriceDigit, "0") & "'"
    mstrNumberFormat = "'999999999990." & String(mintShowNumberDigit, "0") & "'"
    mstrMoneyFormat = "'999999999990." & String(mintShowMoneyDigit, "0") & "'"
    
    SetVisable  '根据权限设置不同的显示项目
    
    dateCurDate = Sys.Currentdate()
    strStart = Format(DateAdd("d", -7, dateCurDate), "yyyy-MM-dd")
    strEnd = Format(dateCurDate, "yyyy-MM-dd")
    
    strVerifyStart = "1901-01-01"
    strVerifyEnd = "1901-01-01"
    
    strFind = " AND A.记录状态 = 1 And A.审核日期 is Null And A.填制日期 Between [3] And [4] "
    SQLCondition.date填制时间开始 = CDate(Format(strStart, "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date填制时间结束 = CDate(Format(strEnd, "yyyy-mm-dd") & " 23:59:59")
        
    mstrFind = strFind
    
    GetList (mstrFind)  '列出单据头
    RestoreWinState Me, App.ProductName, mstrTitle
    '恢复个性化设置后，权限控制显示的列需要进一步控制
    With vsfDetail
        If mblnViewCost = False Then
            .ColHidden(.ColIndex("购价")) = True
            .ColHidden(.ColIndex("购价金额")) = True
            .ColHidden(.ColIndex("差价")) = True
        End If
    End With
    
    Call zldatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs)
    
    mblnBootUp = True
    
    If IsObject(FrmMain) Then
        Me.Show , FrmMain
    Else
        'ZLBH融合显示
        OS.ShowChildWindow Me.hWnd, FrmMain
    End If
    Me.ZOrder 0
End Sub

'检查数据依赖性
Private Function CheckDepend() As Boolean
    
    Dim rsDepend As New Recordset
    Dim strStock As String
    
    CheckDepend = False
    On Error GoTo errHandle
    Select Case mlngMode
        Case 1344
            strStock = "LMN"
        
        Case 1343
            strStock = "HIJKLMN"
        Case Else
            
            
    End Select

    gstrSQL = "SELECT DISTINCT a.id, a.名称 " _
         & "FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " _
        & "Where (a.站点 = '" & gstrNodeNo & "' Or a.站点 is Null) And c.工作性质 = b.名称 " _
          & "AND Instr([1],b.编码,1) > 0 " _
         & " AND a.id = c.部门id " _
          & "AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01'"
    Set rsDepend = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strStock)
    
    If rsDepend.EOF Then
        MsgBox "部门性质信息不全,请查看部门管理！", vbInformation, gstrSysName
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
            If Not IsHavePrivs(gstrprivs, "所有药房") Then
                MsgBox "你不是药房工作人员且不具有所有库房的权限，不能进入！", vbInformation, gstrSysName
                Unload Me
                Exit Function
            End If
            .ListIndex = 0
        ElseIf Not IsHavePrivs(gstrprivs, "所有药房") Then
            .Enabled = False
        End If
    End With
    CheckDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetList(ByVal strFind As String)
    Dim rsList As New Recordset
    Dim strUserPart As String
    Dim n As Integer
    
    On Error GoTo errHandle
    mlastRow = 0
'    strUserPart = " And A.库房ID+0=" & cboStock.ItemData(cboStock.ListIndex)
    strUserPart = " And A.库房ID+0=[11] "
    vsfList.Redraw = flexRDNone
    Select Case mlngMode
        Case 1344           '协定药品入库管理
            gstrSQL = "SELECT /*+ Rule*/ a.no, a.填制人, " _
                  & "TO_CHAR (a.填制日期, 'yyyy-mm-dd HH24:Mi:SS') AS 填制日期, a.审核人, " _
                  & "TO_CHAR (a.审核日期, 'yyyy-mm-dd HH24:Mi:SS') AS 审核日期, a.记录状态, a.摘要 " _
             & "FROM 药品收发记录 a, 部门表 b " _
            & "Where a.库房id = b.ID AND a.单据 = 3 and 入出系数=1 and  a.序号=1  " _
                        & strUserPart & strFind _
           & "ORDER BY no DESC, 填制日期 ASC "
       
    End Select
    
'    Call SQLTest(App.Title, Me.Caption, gstrSQL)
'    rsList.Open gstrSQL, gcnOracle
'    Call SQLTest
    
    Set rsList = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
        SQLCondition.strNO开始, _
        SQLCondition.strNO结束, _
        SQLCondition.date填制时间开始, _
        SQLCondition.date填制时间结束, _
        SQLCondition.date审核时间开始, _
        SQLCondition.date审核时间结束, _
        SQLCondition.lng药品, _
        SQLCondition.lng移出库房, _
        SQLCondition.str填制人, _
        SQLCondition.str审核人, _
        cboStock.ItemData(cboStock.ListIndex))
   
    Set vsfList.DataSource = rsList
    With vsfList
        If .rows = 1 Then
            .rows = .rows + 100
            .Row = 1
            .Redraw = flexRDDirect
            
            .TopRow = 1
            .rows = .rows - 99
            
        End If
        .Row = 1
        .Col = 0
        .ColSel = .Cols - 1
        
        For n = 0 To .Cols - 1
            .FixedAlignment(n) = flexAlignCenterCenter
        Next
    End With
    SetListColWidth
    vsfList.ColWidth(vsfList.Cols - 2) = 0
    vsfList_EnterCell    '列出单据体
    
    SetStrikeColor
    vsfList.Redraw = flexRDDirect
    
    staThis.Panels(2).Text = "当前共有" & rsList.RecordCount & "张单据"
    rsList.Close
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetStrikeColor()
    Dim intStatus As Integer
    Dim intRow As Integer
    Dim intCol As Integer
    
    With vsfList
        If .rows <= 2 Then Exit Sub
        For intRow = 1 To .rows - 1
            intStatus = .TextMatrix(intRow, .Cols - 2)
            If intStatus = 3 Then
                .Cell(flexcpForeColor, intRow, 0, intRow, .Cols - 1) = &H80000001 ' &H80000018
            End If
            If intStatus = 2 Then
                .Cell(flexcpForeColor, intRow, 0, intRow, .Cols - 1) = &HFF '&H80000001
            End If
        Next
    End With
                
End Sub

'表头列宽初始
Private Sub SetListColWidth()
    Dim intCol As Integer
    
    With vsfList
        If mblnBootUp = False Then
            For intCol = 1 To .Cols - 1
                If intCol = 1 Then
                    .ColWidth(intCol) = 2000
                ElseIf intCol = .Cols - 2 Then
                    .ColWidth(intCol) = 0
                Else
                    .ColWidth(intCol) = 1000
                End If
            Next
        End If
            
    End With
End Sub


Private Sub SetDetailColWidth()
    Dim intCol As Integer
    Dim rsDetail As New Recordset
    Dim bln中药库房 As Boolean
    Dim str库房性质 As String
    
    On Error GoTo errHandle
    str库房性质 = ""
    gstrSQL = "Select a.工作性质 From 部门性质说明 A Where a.部门id =[1]"
    Set rsDetail = zldatabase.OpenSQLRecord(gstrSQL, "判断是库房性质", cboStock.ItemData(cboStock.ListIndex))
    Do While Not rsDetail.EOF
        str库房性质 = str库房性质 & "," & rsDetail!工作性质
        rsDetail.MoveNext
    Loop
    If str库房性质 Like "*中药*" Or str库房性质 Like "*制剂室*" Then bln中药库房 = True
        
    With vsfDetail
        Select Case mlngMode
            Case 1344
                .ColAlignment(.ColIndex("数量")) = flexAlignRightCenter     '数量
                .ColAlignment(.ColIndex("单位")) = flexAlignCenterCenter    '单位
                .ColAlignment(.ColIndex("购价")) = flexAlignRightCenter     '购价
                .ColAlignment(.ColIndex("购价金额")) = flexAlignRightCenter '购价金额
                .ColAlignment(.ColIndex("售价")) = flexAlignRightCenter     '售价
                .ColAlignment(.ColIndex("售价金额")) = flexAlignRightCenter '售价金额
                .ColAlignment(.ColIndex("差价")) = flexAlignRightCenter     '差价
        End Select
        
        If mblnBootUp = False Then
            .ColWidth(0) = 0
            .ColWidth(.ColIndex("药品信息")) = 2500
            For intCol = 2 To .Cols - 1
                .ColWidth(intCol) = 1000
            Next
        End If
        If mblnViewCost = False Then
            .ColHidden(.ColIndex("购价")) = True
            .ColHidden(.ColIndex("购价金额")) = True
            .ColHidden(.ColIndex("差价")) = True
        End If
        
        If bln中药库房 Then
            .ColHidden(.ColIndex("原产地")) = False
        Else
            .ColHidden(.ColIndex("原产地")) = True
        End If
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'根据权限设置不同的显示项目
Private Sub SetVisable()
    '外购入库所有权限：参数设置、基本、所有库房、登记、修改、删除、验收、冲销

    Select Case mlngMode
        Case 1344
            If Not IsHavePrivs(gstrprivs, "登记") Then
                mnuEditAdd.Visible = False
                tlbTool.Buttons("Add").Visible = False
            End If
            
            If Not IsHavePrivs(gstrprivs, "修改") Then
                mnuEditModify.Visible = False
                tlbTool.Buttons("Modify").Visible = False
            End If
            
            If Not IsHavePrivs(gstrprivs, "删除") Then
                mnuEditDel.Visible = False
                tlbTool.Buttons("Delete").Visible = False
                 '对没有所有编辑权限时，把菜单和工具栏上的相应的分割线屏蔽。
                If mnuEditAdd.Visible = False And mnuEditModify.Visible = False Then
                    mnuEditLine1.Visible = False
                    tlbTool.Buttons("EditSeparate").Visible = False
                End If
            End If
            
            If Not IsHavePrivs(gstrprivs, "审核") Then
                mnuEditVerify.Visible = False
                tlbTool.Buttons("Verify").Visible = False
            End If
            
            If Not IsHavePrivs(gstrprivs, "冲销") Then
                mnuEditStrike.Visible = False
                tlbTool.Buttons("Strike").Visible = False
                
                If mnuEditVerify.Visible = False Then
                    mnuEditLine2.Visible = False
                    tlbTool.Buttons("VerifySeparate").Visible = False
                End If
            End If
            If Not IsHavePrivs(gstrprivs, "单据打印") Then
               mnuFileBillPrint.Visible = False
               mnuFileBillPreview.Visible = False
            End If
        Case 1301
        
        Case Else
        
    End Select
    
End Sub

Private Sub Form_Load()
    Dim dateCurDate As Date
    
    mblnViewCost = IsHavePrivs(mstrPrivs, "查看成本价")
    '恢复设置
    dateCurDate = Sys.Currentdate()
    lblRange.Caption = "查询范围:" & Format(dateCurDate, "yyyy年MM月dd日") & "至" & Format(dateCurDate, "yyyy年MM月dd日")
    
    staThis.Panels(2).Picture = picColor
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
        .Height = 300
        .Left = 0
        .Width = cbrTool.Width
        
    End With
    
    With vsfList
        .Top = IIf(cbrTool.Visible, cbrTool.Height, 0)
        .Left = 0
        .Width = cbrTool.Width
        .Height = picSeparate_s.Top - .Top
    End With
    
    With vsfDetail
        .Top = picSeparate_s.Top + picSeparate_s.Height + 100
        .Left = 0
        .Height = ScaleHeight - .Top - IIf(staThis.Visible, staThis.Height, 0)
        .Width = cbrTool.Width
    End With
        
    With picColor
        .Top = Me.ScaleHeight - .Height - 30
        .Left = Me.ScaleWidth - staThis.Panels(3).Width - staThis.Panels(4).Width - .Width - 300
    End With
    
    If mlngMode <> 1343 Then
        picColor3.Visible = False
        lblColor3.Visible = False
        picColor.Width = lblColor2.Left + lblColor2.Width + 20
    Else
        lblColor3.Caption = "未审核冲销"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, mstrTitle
End Sub

Private Sub mnuEditAdd_Click()
    Dim strNo As String
    Dim BlnSuccess As Boolean
    
    '检查本期是否已经审核结存，如果未审核结存则不能进行新增业务操作
'    If CheckIsAccount(Val(cboStock.ItemData(cboStock.ListIndex))) = False Then
'        Exit Sub
'    End If
    
    strNo = ""
    '新增
    Select Case mlngMode
        Case 1344
            frmAccordDrugCard.ShowCard Me, strNo, 1, , BlnSuccess
        Case 1306
            '‘frmtputCard.ShowCard Me, strNO, 1, , blnSuccess
    End Select
    
    If BlnSuccess = True Then
        mnuViewRefresh_Click
    End If
End Sub

Private Sub MnuEditVerify_Click()
    '验收
    Dim strNo As String
    Dim BlnSuccess As Boolean
    
    With vsfList
        strNo = .TextMatrix(.Row, 0)
        Select Case mlngMode
            Case 1344
                frmAccordDrugCard.ShowCard Me, strNo, 3, .TextMatrix(.Row, .Cols - 2), BlnSuccess
            Case 1343
                'frmSelfMakeCard.ShowCard Me, StrNo, 3, .TextMatrix(.Row, .Cols - 2), BlnSuccess
            
        End Select
        
    End With
    If BlnSuccess = True Then
        mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditDel_Click()
    '删除
    Dim StrBillNo As String
    Dim intRow As Integer
    Dim strTitle As String
    Dim intReturn As Integer
    Dim intRecord As Integer
    
    With vsfList
        Select Case mlngMode
            Case 1344
                strTitle = "协定药品入库单"
            Case Else
                
        End Select
        
        On Error GoTo errHandle
        intRow = .Row
        StrBillNo = .TextMatrix(intRow, 0)
        intReturn = MsgBox("你确实要删除单据号为“" & StrBillNo & "”的" & strTitle & "吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
        intRecord = .rows - 1
        If intReturn = vbYes Then
            Select Case mlngMode
                Case 1344
                    gstrSQL = "zl_协定入库_Delete('" & StrBillNo & "')"
            End Select
            If gstrSQL = "" Then Exit Sub
            Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            
            intRecord = intRecord - 1
            
            mlastRow = 0
            If .rows > 2 Then
                .RemoveItem intRow
            ElseIf .rows = 2 Then
                .rows = 3
                .RemoveItem intRow
                With vsfDetail
                    .rows = 1
                    .rows = 2
                    .FixedRows = 1
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                End With
                SetEnable
                
            End If
                
            '.RowHeight(intRow) = 0
            If intRow < .rows - 1 Then
                .Row = intRow
            Else
                If .rows = 2 Then
                    .Row = 1
                Else
                    .Row = intRow - 1
                End If
            End If
            .Col = 0
            .ColSel = .Cols - 1
            vsfList_EnterCell
        End If
    End With
    staThis.Panels(2).Text = "当前共有" & intRecord & "张单据"
    Exit Sub
    
errHandle:
    If ErrCenter() = 1 Then Resume 'Resume这种情况不用调用
    Call SaveErrLog
    
End Sub

Private Sub mnuEditDisplay_Click()
    '查看单据
    
    Dim strNo As String
    With vsfList
        strNo = .TextMatrix(.Row, 0)
        Select Case mlngMode
            Case 1344
                frmAccordDrugCard.ShowCard Me, strNo, 4, .TextMatrix(.Row, .Cols - 2)
            Case 1343
                'frmOtherInputCard.ShowCard Me, StrNo, 4, .TextMatrix(.Row, .Cols - 2)
        
        End Select
        
    End With
    
End Sub

Private Sub mnuEditStrike_Click()
    '冲销
    
    With vsfList
        If MsgBox("你确实要全部冲销单据号为“" & .TextMatrix(.Row, 0) & "”的单据吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            If StrikeSave = True Then
                mnuViewRefresh_Click
            End If
        End If
    End With
End Sub

Private Function StrikeSave() As Boolean
    Dim n As Integer
    
    StrikeSave = False
    With vsfList
        Select Case mlngMode
            Case 1344
                mint库存检查 = MediWork_GetCheckStockRule(cboStock.ItemData(cboStock.ListIndex))
                
                '检查可用数量是否足够，参数设置为不检查库存时不进行（传入单据，药名，库存检查，单据号，序号）
                If mint库存检查 <> 0 And .TextMatrix(.Row, 0) <> "" Then
                    For n = 1 To vsfDetail.rows - 1
                        If vsfDetail.TextMatrix(n, 0) <> "" Then
                            If CheckStrickUsable(3, 0, 0, vsfDetail.TextMatrix(n, 1), _
                                0, 0, mint库存检查, Trim(.TextMatrix(.Row, 0)), Val(vsfDetail.TextMatrix(n, 0))) = False Then
                                Exit Function
                            End If
                        End If
                    Next
                End If
                
                gstrSQL = "zl_协定入库_Strike('" & .TextMatrix(.Row, 0) & "','" & .TextMatrix(.Row, .Cols - 4) & "')"
            
        End Select
        
        On Error GoTo errHandle
        
        Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        '提示停用药品
        Call CheckStopMedi(单据号.协药入库 & "|" & .TextMatrix(.Row, 0))
    End With
    StrikeSave = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

    'MsgBox "存盘失败！", vbInformation, gstrSysName
End Function

Private Sub mnuEditModify_Click()
    '修改
    Dim strNo As String
    Dim BlnSuccess As Boolean
    
    BlnSuccess = False
    With vsfList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        strNo = .TextMatrix(.Row, 0)
        
        Select Case mlngMode
            Case 1344
                frmAccordDrugCard.ShowCard Me, strNo, 2, vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 2), BlnSuccess
            Case 1302
                
        End Select
        If BlnSuccess = True Then
            mnuViewRefresh_Click
        End If
    End With
End Sub

Private Sub mnuFileBillPreview_Click()
    Dim int单位系数 As Integer
    
    With vsfList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub

        Select Case mintUnit
            Case mconint售价单位
                int单位系数 = 4
            Case mconint门诊单位
                int单位系数 = 2
            Case mconint住院单位
                int单位系数 = 1
            Case mconint药库单位
                int单位系数 = 3
        End Select
    
        ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1344", "zl8_bill_1344"), Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & .TextMatrix(.Row, .Cols - 2), "单位系数=" & int单位系数, 1
    End With
End Sub

Private Sub MnuFileBillprint_Click()
    Dim int单位系数 As Integer
    
    With vsfList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        
        Select Case mintUnit
            Case mconint售价单位
                int单位系数 = 4
            Case mconint门诊单位
                int单位系数 = 2
            Case mconint住院单位
                int单位系数 = 1
            Case mconint药库单位
                int单位系数 = 3
        End Select
        ReportOpen gcnOracle, glngSys, "zl1_bill_1344", Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & .TextMatrix(.Row, .Cols - 2), "单位系数=" & int单位系数, 2
    End With
End Sub

Private Sub mnuFileExcel_Click()
    '输出到Excel
    If Me.ActiveControl Is vsfList Then
        vsfList.Redraw = flexRDNone
        subPrint 3
        vsfList.Redraw = flexRDDirect
        vsfList.Col = 0
        vsfList.ColSel = vsfList.Cols - 1
    ElseIf Me.ActiveControl Is vsfDetail Then
        vsfDetail.Redraw = flexRDNone
        subExcel 3
        vsfDetail.Redraw = flexRDDirect
        vsfDetail.Col = 0
        vsfDetail.ColSel = vsfDetail.Cols - 1
    End If
End Sub

Private Sub mnufileexit_Click()
    '退出
    Unload Me
    
End Sub

Private Sub mnuFileParameter_Click()
    '参数设置
    frm参数设置.设置参数 Me, mstrPrivs, 1344, Me.Tag
    Call GetList(mstrFind)
End Sub

Private Sub mnuFilePreView_Click()
    '打印预览
    vsfList.Redraw = flexRDNone
    subPrint 2
    vsfList.Redraw = flexRDDirect
    vsfList.Col = 0
    vsfList.ColSel = vsfList.Cols - 1
End Sub

Private Sub mnuFilePrint_Click()
    '打印

    vsfList.Redraw = flexRDNone
    subPrint 1
    vsfList.Redraw = flexRDDirect
    vsfList.Col = 0
    vsfList.ColSel = vsfList.Cols - 1
    
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
'    ReportMan gcnOracle, Me
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub mnuHelpWebHome_Click()
    '中联主页
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    '发送反馈
    Call zlMailTo(Me.hWnd)
End Sub

Private Sub mnuReportItem_Click(index As Integer)
    '默认参数：药品=药品id，库房=库房id，开始时间=填制开始时间，结束时间=填制结束时间，NO=入库单NO
    Dim str开始时间 As String
    Dim str结束时间 As String
    Dim strNo As String
    
    If vsfList.Row >= 1 And LTrim(vsfList.TextMatrix(vsfList.Row, 0)) <> "" Then
        strNo = vsfList.TextMatrix(vsfList.Row, 0)
    End If
    
    str开始时间 = IIf(Format(SQLCondition.date填制时间开始, "yyyy-mm-dd") = "1899-12-30", "", Format(SQLCondition.date填制时间开始, "yyyy-mm-dd"))
    str结束时间 = IIf(Format(SQLCondition.date填制时间结束, "yyyy-mm-dd") = "1899-12-30", "", Format(SQLCondition.date填制时间结束, "yyyy-mm-dd"))
    
    Call ReportOpen(gcnOracle, Split(mnuReportItem(index).Tag, ",")(0), Split(mnuReportItem(index).Tag, ",")(1), Me, _
        "药品=" & IIf(SQLCondition.lng药品 = 0, "", SQLCondition.lng药品), _
        "库房=" & Val(cboStock.ItemData(cboStock.ListIndex)), _
        "开始时间=" & str开始时间, _
        "结束时间=" & str结束时间, _
        "NO=" & strNo)
End Sub
Private Sub mnuViewRefresh_Click()
    '刷新
    GetList mstrFind
End Sub

Private Sub mnuViewSearch_Click()
    '查找
    Dim strFind As String
    
    Select Case mlngMode
        Case 1344
            strFind = FrmListSearch.GetSearch(Me, mlngMode, cboStock.ItemData(cboStock.ListIndex), strStart, strEnd, strVerifyStart, strVerifyEnd, _
                SQLCondition.strNO开始, _
                SQLCondition.strNO结束, _
                SQLCondition.date填制时间开始, _
                SQLCondition.date填制时间结束, _
                SQLCondition.date审核时间开始, _
                SQLCondition.date审核时间结束, _
                SQLCondition.lng药品, _
                SQLCondition.lng移出库房, _
                SQLCondition.str填制人, _
                SQLCondition.str审核人)
        Case Else
    End Select
    '将所有模块的填制时间保存到注册表中
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & mlngMode, "默认查询时间", SQLCondition.date填制时间开始 & "|" & SQLCondition.date填制时间结束)
    
    If strFind <> "" Then
        mstrFind = strFind
        GetList mstrFind
        If Format(strStart, "yyyy-mm-dd") = "1901-01-01" And Format(strVerifyStart, "yyyy-mm-dd") = "1901-01-01" Then
            lblRange.Visible = False
        ElseIf Format(strStart, "yyyy-mm-dd") <> "1901-01-01" And Format(strVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
            lblRange.Visible = True
            lblRange = "查询范围:填制日期 " & Format(strStart, "yyyy年MM月dd日") & "至" & Format(strEnd, "yyyy年MM月dd日") & "  审核日期 " & Format(strVerifyStart, "yyyy年MM月dd日") & "至" & Format(strVerifyEnd, "yyyy年MM月dd日")
        ElseIf Format(strStart, "yyyy-mm-dd") <> "1901-01-01" Then
            lblRange.Visible = True
            lblRange = "查询范围:填制日期 " & Format(strStart, "yyyy年MM月dd日") & "至" & Format(strEnd, "yyyy年MM月dd日")
        ElseIf Format(strVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
            lblRange.Visible = True
            lblRange = "查询范围:审核日期 " & Format(strVerifyStart, "yyyy年MM月dd日") & "至" & Format(strVerifyEnd, "yyyy年MM月dd日")
        End If
             
    End If
    
End Sub

Private Sub mnuViewStatus_Click()
    With mnuViewStatus
        .Checked = Not .Checked  ' Xor True
        staThis.Visible = .Checked
    End With
    
    Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    With mnuViewToolButton
        .Checked = Not .Checked   ' Xor True
        cbrTool.Bands(1).Visible = .Checked
        mnuViewToolText.Enabled = .Checked
    End With
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim intCount As Integer      '工具条索引
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked   ' Xor True
    With tlbTool.Buttons
        If mnuViewToolText.Checked = False Then
            '取消所有的文本标签显示
            For intCount = 1 To .count
                .Item(intCount).Caption = ""
            Next
        Else
            '让所有的文本标签显示。说明：Tag中放的文本标签
            For intCount = 1 To .count
                .Item(intCount).Caption = .Item(intCount).Tag
            Next
        End If
    End With
    
    cbrTool.Bands(1).MinHeight = tlbTool.Height
    
    Form_Resize
End Sub

Private Sub vsfDetail_GotFocus()
    Call SetGridFocus(vsfDetail, True)
End Sub

Private Sub vsfDetail_LostFocus()
    Call SetGridFocus(vsfDetail, False)
End Sub

Private Sub vsfList_DblClick()
    If mnuEditModify.Visible = False Then Exit Sub
    If mnuEditModify.Enabled = False Then Exit Sub
    If vsfList.MouseRow = 0 Then Exit Sub
    mnuEditModify_Click
End Sub

Private Sub vsfList_EnterCell()
    Dim rsDetail As New Recordset
    Dim strUnitQuantity As String               '单位和数量格式化串
    Dim IntBill As Integer                      '单据类型  如：1、外购入库；2、
    Dim str包装系数 As String
    Dim strSql药名 As String
    Dim intCol As Integer
    Dim n As Integer
    
    If mlastRow = vsfList.Row Then Exit Sub
    mlastRow = vsfList.Row
    
    On Error GoTo errHandle
    With vsfList
        .Redraw = flexRDNone
        .ForeColorSel = .Cell(flexcpForeColor, mlastRow, 1)
        .Redraw = flexRDDirect
    End With
    
    SetEnable
    
    If vsfList.Row >= 1 And LTrim(vsfList.TextMatrix(vsfList.Row, 0)) <> "" Then
        vsfList.Col = 0
        vsfList.ColSel = vsfList.Cols - 1
        
        vsfDetail.Redraw = flexRDNone
        
        Select Case mintUnit
            Case mconint售价单位
                strUnitQuantity = "LTRIM(to_char(A.实际数量," & mstrNumberFormat & ")) AS 数量," _
                    & "F.计算单位 AS 单位,"
                str包装系数 = "1"
            Case mconint门诊单位
                strUnitQuantity = "LTRIM(to_char(A.实际数量 / B.门诊包装," & mstrNumberFormat & ")) AS 数量," _
                    & "B.门诊单位 AS 单位,"
                str包装系数 = "B.门诊包装"
            Case mconint住院单位
                strUnitQuantity = "LTRIM(to_char(A.实际数量 / B.住院包装," & mstrNumberFormat & ")) AS 数量," _
                    & "B.住院单位 AS 单位,"
                str包装系数 = "B.住院包装"
            Case mconint药库单位
                strUnitQuantity = "LTRIM(to_char(A.实际数量 / B.药库包装," & mstrNumberFormat & ")) AS 数量," _
                    & "B.药库单位 AS 单位,"
                str包装系数 = "B.药库包装"
        End Select
        
        If gint药品名称显示 = 0 Then
            strSql药名 = "('['||F.编码||']'||F.名称) AS 药品信息"
        ElseIf gint药品名称显示 = 1 Then
            strSql药名 = "('['||F.编码||']'||NVL(E.名称,F.名称)) AS 药品信息"
        Else
            strSql药名 = "('['||F.编码||']'||F.名称) AS 药品信息,E.名称 As 商品名"
        End If
        
        Select Case mlngMode
            Case 1344
                IntBill = 3
                gstrSQL = " SELECT * FROM " & _
                    "    (SELECT DISTINCT 序号," & strSql药名 & ",B.药品来源,B.基本药物,F.规格,A.产地,A.原产地," & _
                    strUnitQuantity & _
                    "    LTRIM(TO_CHAR (A.成本价*" & str包装系数 & ", " & mstrCostFormat & ")) AS 购价," & _
                    "    LTRIM(TO_CHAR (A.成本金额, " & mstrMoneyFormat & ")) AS 购价金额," & _
                    "    LTRIM(TO_CHAR (A.零售价*" & str包装系数 & "," & mstrPriceFormat & ")) AS 售价," & _
                    "    LTRIM(TO_CHAR (A.零售金额, " & mstrMoneyFormat & ")) AS 售价金额," & _
                    "    LTRIM(TO_CHAR (A.差价, " & mstrMoneyFormat & ")) AS 差价 " & _
                    "    FROM 药品收发记录 A, 药品规格 B,收费项目别名 E,收费项目目录 F " & _
                    "    WHERE A.药品ID = B.药品ID AND B.药品ID=F.ID" & _
                    "    AND B.药品ID = E.收费细目ID(+) AND E.性质(+)=3 " & _
                    "    AND 记录状态 = [2] " & _
                    "    AND A.单据 = 3 AND 入出系数=1 " & _
                    "    AND A.NO = [1])" & _
                    " ORDER BY 序号 "
        End Select
        
        Set rsDetail = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, vsfList.TextMatrix(vsfList.Row, 0), Val(vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 2)))
        
        Set vsfDetail.DataSource = rsDetail
        rsDetail.Close
        
        With vsfDetail
            .Row = 1
            .Col = 0
            .ColSel = .Cols - 1
            
            For n = 0 To .Cols - 1
                .ColKey(n) = .TextMatrix(0, n)
                .FixedAlignment(n) = flexAlignCenterCenter
            Next
        End With
        
        vsfDetail.Redraw = flexRDDirect
    ElseIf LTrim(vsfList.TextMatrix(vsfList.Row, 0)) = "" Then
        With vsfDetail
            .Redraw = flexRDNone
            Select Case mlngMode
                Case 1344
                    .Cols = IIf(gint药品名称显示 = 2, 15, 14)
                    .rows = 2
                    .Clear
                    
                    intCol = 0
                    
                    .TextMatrix(0, intCol) = "序号": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "药品信息": intCol = intCol + 1
                    
                    If gint药品名称显示 = 2 Then
                        .TextMatrix(0, intCol) = "商品名": intCol = intCol + 1
                    End If
                    
                    .TextMatrix(0, intCol) = "药品来源": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "基本药物": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "规格": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "生产商": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "原产地": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "数量": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "单位": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "购价": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "购价金额": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "售价": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "售价金额": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "差价": intCol = intCol + 1
                
            End Select
            .Row = 1
            .Col = 0
            .ColSel = .Cols - 1
            
            For n = 0 To .Cols - 1
                .ColKey(n) = .TextMatrix(0, n)
                .FixedAlignment(n) = flexAlignCenterCenter
            Next
            .Redraw = flexRDDirect
        End With
    End If
    SetDetailColWidth
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsfList_GotFocus()
    Call SetGridFocus(vsfList, True)
End Sub

Private Sub vsfList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If mnuEditModify.Visible = False Then Exit Sub
        If mnuEditModify.Enabled = False Then Exit Sub
        mnuEditModify_Click
    End If
        
End Sub

Private Sub vsfList_LostFocus()
    Call SetGridFocus(vsfList, False)
End Sub

Private Sub vsfList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    If mnuEdit.Visible = False Then Exit Sub
    
    PopupMenu mnuEdit, 2
    
End Sub

Private Sub picSeparate_s_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    '分割条设置
    
    If Button <> 1 Then Exit Sub
    
    With picSeparate_s
        If .Top + y < 2000 Then Exit Sub
        If .Top + y > ScaleHeight - 2000 Then Exit Sub
        .Move .Left, .Top + y
    End With
    
    With vsfList
        .Top = IIf(cbrTool.Visible, cbrTool.Height, 0)
        .Height = picSeparate_s.Top - .Top
    End With
    
    With vsfDetail
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
            mnuEditAdd_Click
        Case "Modify"
            mnuEditModify_Click
        Case "Delete"
            mnuEditDel_Click
        Case "Verify"
            MnuEditVerify_Click
        Case "Strike"
            mnuEditStrike_Click
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
    With vsfList
        .ToolTipText = ""
        If .TextMatrix(.Row, 0) = "" Or .Row = 0 Then         '没有单
            mnuFilePreView.Enabled = False
            mnuFilePrint.Enabled = False
            mnuFileBillPreview.Enabled = False
            mnuFileBillPrint.Enabled = False
            mnuFileExcel.Enabled = False
            tlbTool.Buttons("Print").Enabled = False
            tlbTool.Buttons("PrintView").Enabled = False
            
            If mnuEditModify.Visible = True Then
                mnuEditModify.Enabled = False
                tlbTool.Buttons("Modify").Enabled = False
            End If
            If mnuEditDel.Visible = True Then
                mnuEditDel.Enabled = False
                tlbTool.Buttons("Delete").Enabled = False
            End If
            If mnuEditVerify.Visible = True Then
                mnuEditVerify.Enabled = False
                tlbTool.Buttons("Verify").Enabled = False
            End If
            
            If mnuEditStrike.Visible = True Then
                mnuEditStrike.Enabled = False
                tlbTool.Buttons("Strike").Enabled = False
            End If
             
            If mnuEditDisplay.Visible = True Then
                mnuEditDisplay.Enabled = False
            End If
         Else
            mnuFilePreView.Enabled = True
            mnuFilePrint.Enabled = True
            mnuFileBillPreview.Enabled = True
            mnuFileBillPrint.Enabled = True
            mnuFileExcel.Enabled = True
            tlbTool.Buttons("Print").Enabled = True
            tlbTool.Buttons("PrintView").Enabled = True
            
            
            If .TextMatrix(.Row, .Cols - 4) = "" Then    '未审核单
                If mnuEditModify.Visible = True Then
                    mnuEditModify.Enabled = True
                    tlbTool.Buttons("Modify").Enabled = True
                End If
                If mnuEditDel.Visible = True Then
                    mnuEditDel.Enabled = True
                    tlbTool.Buttons("Delete").Enabled = True
                End If
                If mnuEditVerify.Visible = True Then
                    mnuEditVerify.Enabled = True
                    tlbTool.Buttons("Verify").Enabled = True
                End If
                
                If mnuEditStrike.Visible = True Then
                    mnuEditStrike.Enabled = False
                    tlbTool.Buttons("Strike").Enabled = False
                End If
                 
                If mnuEditDisplay.Visible = True Then
                    mnuEditDisplay.Enabled = True
                End If
            ElseIf .TextMatrix(.Row, .Cols - 2) = 1 Then    '审核单
                If mnuEditModify.Visible = True Then
                    mnuEditModify.Enabled = False
                    tlbTool.Buttons("Modify").Enabled = False
                End If
                If mnuEditDel.Visible = True Then
                    mnuEditDel.Enabled = False
                    tlbTool.Buttons("Delete").Enabled = False
                End If
                If mnuEditVerify.Visible = True Then
                    mnuEditVerify.Enabled = False
                    tlbTool.Buttons("Verify").Enabled = False
                End If
                
                If mnuEditStrike.Visible = True Then
                    mnuEditStrike.Enabled = True
                    tlbTool.Buttons("Strike").Enabled = True
                End If
                 
                If mnuEditDisplay.Visible = True Then
                    mnuEditDisplay.Enabled = True
                End If
            Else   '2,3 冲销单
                If .TextMatrix(.Row, .Cols - 2) = 3 Then
                    .ToolTipText = "冲销单据的原单据"
                ElseIf .TextMatrix(.Row, .Cols - 2) = 2 Then
                    .ToolTipText = "冲销单据"
                End If
                If mnuEditModify.Visible = True Then
                    mnuEditModify.Enabled = False
                    tlbTool.Buttons("Modify").Enabled = False
                End If
                If mnuEditDel.Visible = True Then
                    mnuEditDel.Enabled = False
                    tlbTool.Buttons("Delete").Enabled = False
                End If
                If mnuEditVerify.Visible = True Then
                    mnuEditVerify.Enabled = False
                    tlbTool.Buttons("Verify").Enabled = False
                End If
                
                If mnuEditStrike.Visible = True Then
                    mnuEditStrike.Enabled = False
                    tlbTool.Buttons("Strike").Enabled = False
                End If
                 
                If mnuEditDisplay.Visible = True Then
                    mnuEditDisplay.Enabled = True
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
    
    If Format(strStart, "yyyy-mm-dd") = "1901-01-01" And Format(strVerifyStart, "yyyy-mm-dd") = "1901-01-01" Then
        strRange = "审核日期 " & Format(strVerifyStart, "yyyy年MM月dd日") & "至" & Format(strVerifyEnd, "yyyy年MM月dd日")
    ElseIf Format(strVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
        strRange = "填制日期 " & Format(strStart, "yyyy年MM月dd日") & "至" & Format(strEnd, "yyyy年MM月dd日") & "  审核日期 " & Format(strVerifyStart, "yyyy年MM月dd日") & "至" & Format(strVerifyEnd, "yyyy年MM月dd日")
    Else
        strRange = "填制日期 " & Format(strStart, "yyyy年MM月dd日") & "至" & Format(strEnd, "yyyy年MM月dd日")
    End If
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
    objPrint.Title.Text = mstrTitle
        
    objRow.Add "时间：" & strRange
    objRow.Add "部门：" & cboStock.Text
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
        
    objRow.Add "打印人:" & gstrUserName
    objRow.Add "打印日期:" & Format(Sys.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    
    Set objPrint.Body = vsfList
    
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

Private Sub tlbTool_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool
    End If
End Sub

'寻找与某一列相等的行
Public Function FindRow(ByVal FlexTemp As MSHFlexGrid, ByVal intTemp As Variant, ByVal intCol As Integer) As Integer
    Dim i As Integer
    
    With FlexTemp
        For i = 1 To .rows - 1
            If IsDate(intTemp) Then
               If Format(.TextMatrix(i, intCol), "yyyy-mm-dd") = Format(intTemp, "yyyy-mm-dd") Then
                  FindRow = i
                  Exit Function
               End If
            Else
                If .TextMatrix(i, intCol) = intTemp Then
                  FindRow = i
                  Exit Function
                End If
            End If
        Next
    End With
    FindRow = 1
End Function

Private Sub subExcel(bytMode As Byte)
'功能:输出单据到EXCEL
'参数:bytMode=3 输出到EXCEL
'
    Dim objPrint As Object
    Dim objRow As zlTabAppRow
    Dim strRange As String
    
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
    objPrint.Title.Text = mstrTitle
        
    Set objRow = New zlTabAppRow
    objRow.Add ""
    objRow.Add "NO." & Trim(vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "NO")))
    objPrint.UnderAppRows.Add objRow

    Set objRow = New zlTabAppRow
    objRow.Add "库房：" & Trim(cboStock.Text)
    objPrint.UnderAppRows.Add objRow
        
    Set objRow = New zlTabAppRow
    objRow.Add "摘要:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "摘要"))
    objPrint.BelowAppRows.Add objRow
        
    Set objRow = New zlTabAppRow
    objRow.Add "填制人:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "填制人")) & "  填制日期:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "填制日期"))
    
    objRow.Add "审核人:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "审核人")) & "  审核日期:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "审核日期"))
    objPrint.BelowAppRows.Add objRow
    
    Set objPrint.Body = vsfDetail
    zlPrintOrView1Grd objPrint, bytMode
End Sub




Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

