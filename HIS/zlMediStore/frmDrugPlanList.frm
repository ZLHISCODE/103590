VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frmDrugPlanList 
   Caption         =   "药品计划管理"
   ClientHeight    =   5895
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9120
   Icon            =   "frmDrugPlanList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5895
   ScaleWidth      =   9120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picSeparate_s 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   30
      MousePointer    =   7  'Size N S
      ScaleHeight     =   300
      ScaleWidth      =   4815
      TabIndex        =   5
      Top             =   2790
      Width           =   4815
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "查询范围:1999年8月12日至1999年9月12日"
         Height          =   180
         Left            =   0
         TabIndex        =   6
         Top             =   120
         Width           =   3690
      End
   End
   Begin VB.CommandButton Cmd查阅 
      Caption         =   "查阅(&V)"
      Height          =   350
      Left            =   5160
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1100
   End
   Begin ComCtl3.CoolBar cbrTool 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   9120
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tlbTool"
      MinWidth1       =   4995
      MinHeight1      =   720
      Width1          =   4995
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Caption2        =   "库房"
      Child2          =   "cboStock"
      MinHeight2      =   300
      Width2          =   1995
      NewRow2         =   0   'False
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   7740
         TabIndex        =   3
         Text            =   "cboStock"
         Top             =   240
         Width           =   1290
      End
      Begin MSComctlLib.Toolbar tlbTool 
         Height          =   720
         Left            =   165
         TabIndex        =   1
         Top             =   30
         Width           =   6840
         _ExtentX        =   12065
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
            NumButtons      =   19
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
               Caption         =   "复核"
               Key             =   "Check"
               Description     =   "复核"
               Object.ToolTipText     =   "复核"
               Object.Tag             =   "复核"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "清除"
               Key             =   "Clear"
               Description     =   "清除"
               Object.ToolTipText     =   "清除"
               Object.Tag             =   "清除"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "取消"
               Key             =   "Cancel"
               Description     =   "取消审核"
               Object.ToolTipText     =   "取消"
               Object.Tag             =   "取消"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "VerifySeparate"
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "过滤"
               Key             =   "Search"
               Description     =   "过滤"
               Object.ToolTipText     =   "过滤"
               Object.Tag             =   "过滤"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "刷新"
               Key             =   "Refresh"
               Description     =   "刷新"
               Object.ToolTipText     =   "刷新"
               Object.Tag             =   "刷新"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "PlugInSeparator"
               Style           =   3
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "功能"
               Key             =   "PlugItem"
               ImageIndex      =   14
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FindSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助主题"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Exit"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   11
            EndProperty
         EndProperty
         MouseIcon       =   "frmDrugPlanList.frx":014A
      End
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   5535
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDrugPlanList.frx":0464
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11007
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
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":0CF8
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":0F18
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":1138
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":1354
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":1574
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":1794
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":19B0
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":1BCC
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":1DE6
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":1F40
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":215C
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":237C
            Key             =   "Cancel"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":2596
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":2C90
            Key             =   "PlugIn"
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
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":3B6A
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":3D8A
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":3FAA
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":41C6
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":43E6
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":4606
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":4822
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":4A3E
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":4C58
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":4DB2
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":4FD2
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":51F2
            Key             =   "Cancel"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":540C
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":5B06
            Key             =   "PlugIn"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   1005
      Left            =   0
      TabIndex        =   7
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
      FormatString    =   $"frmDrugPlanList.frx":82B8
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
      TabIndex        =   8
      Top             =   3240
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
      FormatString    =   $"frmDrugPlanList.frx":832D
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
      Begin VB.Menu mnuPlugIn 
         Caption         =   "扩展(&E)"
         Visible         =   0   'False
         Begin VB.Menu mnuPlugItem 
            Caption         =   "功能"
            Index           =   0
         End
      End
      Begin VB.Menu mnuEditLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditVerify 
         Caption         =   "审核(&V)"
      End
      Begin VB.Menu mnuEditCheck 
         Caption         =   "复核(&R)"
      End
      Begin VB.Menu mnuEditClear 
         Caption         =   "清除(&S)"
      End
      Begin VB.Menu mnuEditCancel 
         Caption         =   "取消(&C)"
      End
      Begin VB.Menu mnuEditLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditExeAmount 
         Caption         =   "修改执行数量(&E)"
      End
      Begin VB.Menu mnuEditExport 
         Caption         =   "采购计划导出(&X)"
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
      Begin VB.Menu mnuViewColDefine 
         Caption         =   "列设置(&C)"
      End
      Begin VB.Menu mnuViewLine3 
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
Attribute VB_Name = "frmDrugPlanList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngMode As Long
Private mstrFind As String
Private mintPreCol As Integer           '前一次单据头的排序列
Private mintsort As Integer             '前一次单据头的排序
Private mblnBootUp As Boolean
Private mlastRow As Long                '上次点击的行
Private mstrPrivs As String
Private mint价格显示 As Integer         '0:显示成本价;  1:显示售价;  2:显示成本和售价
Private mblnViewCost As Boolean             '查看成本价 true-可以查看成本价 false-不可以查看成本价
Private Const MStrCaption As String = "药品计划管理"
Private mintPlanPoint As Integer        '全院计划不管站点 0-要管站点，1-不管站点

Private strStart As Date
Private strEnd As Date
Private strVerifyStart As Date
Private strVerifyEnd As Date
Private mlng库房ID As Long  '库房id
Private mintUnit As Integer             '单位系数：1-售价;2-门诊;3-住院;4-药库
Private mobjPlugIn As Object             '外挂接口

'从参数表中取药品价格、数量、金额小数位数
Private mintShowCostDigit As Integer            '成本价小数位数
Private mintShowPriceDigit As Integer           '售价小数位数
Private mintShowNumberDigit As Integer          '数量小数位数
Private mintShowMoneyDigit As Integer           '金额小数位数

Private mstrCostFormat As String
Private mstrPriceFormat As String
Private mstrNumberFormat As String
Private mstrMoneyFormat As String

Private Type Type_SQLCondition
    strNO开始 As String
    strNO结束 As String
    date填制时间开始 As Date
    date填制时间结束 As Date
    date审核时间开始 As Date
    date审核时间结束 As Date
    date复核时间开始 As Date
    date复核时间结束 As Date
    str填制人 As String
    str审核人 As String
    str复核人 As String
    lng计划类型 As Long
    lng编制方法 As Long
    lng药品 As Long
End Type

Private SQLCondition As Type_SQLCondition

Private Sub PlugInFun(ByVal strFunName As String)
    '执行外挂功能
    Dim strParam As String
    Dim lng库房ID As Long
    Dim int单据 As Integer
    Dim strNo As String
    
    With vsfList
        If .TextMatrix(.Row, 0) <> "" Then
            lng库房ID = Val(cboStock.ItemData(cboStock.ListIndex))
            int单据 = 0
            strNo = .TextMatrix(.Row, .ColIndex("NO"))
            
            strParam = lng库房ID & "," & int单据 & "," & strNo
        End If
    End With
    
    Call zlPlugIn_Fun(glngSys, mlngMode, mobjPlugIn, Me, strFunName, strParam)
End Sub
'检查数据依赖性
Private Function CheckDepend() As Boolean
    Dim rsDepend As New Recordset
    
    On Error GoTo errHandle
    CheckDepend = False
    
    gstrSQL = "SELECT DISTINCT a.id, a.名称 " _
            & "FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " _
            & "Where (a.站点 = [1] Or a.站点 is Null) And c.工作性质 = b.名称 " _
            & "  AND Instr('HIJKLMN',b.编码,1) > 0 " _
            & "  AND a.id = c.部门id " _
            & "  AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01' " _
            & IIf(zlStr.IsHavePrivs(mstrPrivs, "所有库房"), "", " And a.ID IN (Select 部门ID From 部门人员 Where 人员ID=[2])")
    Set rsDepend = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, gstrNodeNo, UserInfo.用户ID)
    
    If rsDepend.EOF Then
        MsgBox "没有设置药库性质的部门,请查看部门管理！", vbInformation, gstrSysName
        rsDepend.Close
        Exit Function
    End If
            
    With cboStock
        .Clear
        
        '无站点时才能设置全院
        '0-要管站点，1-不管站点
        If (gstrNodeNo = "-" Or gstrNodeNo = "0") Or mintPlanPoint = 1 Then
            .AddItem "全院"
            .ItemData(.NewIndex) = 0
        End If
        
        Do While Not rsDepend.EOF
            
            .AddItem rsDepend!名称
            .ItemData(.NewIndex) = rsDepend!id
            If rsDepend!id = UserInfo.部门ID Then
                .ListIndex = .NewIndex
            End If
            rsDepend.MoveNext
        Loop
        rsDepend.Close
        
        If .ListIndex = -1 Then
            If Not zlStr.IsHavePrivs(mstrPrivs, "所有库房") Then
                MsgBox "你不是药房工作人员且不具有所有库房的权限，不能进入！", vbInformation, gstrSysName
                Unload Me
                Exit Function
            End If
            .ListIndex = 0
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

Private Sub cboStock_Click()
    If mblnBootUp Then
        mlng库房ID = Me.cboStock.ItemData(Me.cboStock.ListIndex)
    
        Call GetDrugDigit(mlng库房ID, MStrCaption, mintUnit, mintShowCostDigit, mintShowPriceDigit, mintShowNumberDigit, mintShowMoneyDigit)
        
        '组织格式化串
        mstrCostFormat = "'999999999990." & String(mintShowCostDigit, "0") & "'"
        mstrPriceFormat = "'999999999990." & String(mintShowPriceDigit, "0") & "'"
        mstrNumberFormat = "'999999999990." & String(mintShowNumberDigit, "0") & "'"
        mstrMoneyFormat = "'999999999990." & String(mintShowMoneyDigit, "0") & "'"
        mnuViewRefresh_Click
    End If
End Sub

Private Sub cboStock_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim str工作性质 As String
    
    str工作性质 = "H,I,J,K,L,M,N"

    If KeyCode <> vbKeyReturn Then Exit Sub
    If cboStock.ListCount = 0 Then Call zlControl.ControlSetFocus(vsfList): Exit Sub
    
    If cboStock.ListIndex >= 0 Then
        If Val(cboStock.Tag) = cboStock.ItemData(cboStock.ListIndex) Then
            Call zlControl.ControlSetFocus(vsfList, True)
            Exit Sub
        End If
    End If
    
    If Select部门选择器(Me, cboStock, Trim(cboStock.Text), str工作性质, IIf(zlStr.IsHavePrivs(mstrPrivs, "所有库房"), False, True)) = False Then
        Exit Sub
    End If
    If cboStock.ListIndex >= 0 Then
        cboStock.Tag = cboStock.ItemData(cboStock.ListIndex)
    End If
End Sub

Private Sub cboStock_KeyPress(KeyAscii As Integer)
    '屏蔽输入单引号
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub


Private Sub cboStock_Validate(Cancel As Boolean)
    If cboStock.ListCount > 0 Then
        If cboStock.ListIndex = -1 Then
            MsgBox "请选择一个药库或者药房！", vbInformation, gstrSysName
            Cancel = True
        End If
    End If
End Sub


Private Sub cbrTool_Resize()
    Form_Resize
End Sub

Private Sub GetList(ByVal strFind As String)
    Dim rsList As New Recordset
    Dim n As Integer
    Dim strsql As String
     
    On Error GoTo errHandle
    Call FS.ShowFlash("正在搜索药品付款记录,请稍候 ...", Me)
    DoEvents
    Screen.MousePointer = vbHourglass
    
    vsfList.Redraw = flexRDNone
    strsql = ""
    If SQLCondition.lng药品 <> 0 Then
        strsql = ", 药品计划内容 C "
        strFind = " And a.Id = c.计划id and C.药品id=[15] " & strFind
    End If
    
    gstrSQL = " SELECT a.NO, a.ID, DECODE(a.计划类型,0,'临时',1,'月度计划',2,'季度计划',3,'年度计划','周计划') AS 计划类型 ," & _
        "a.期间,DECODE(A.编制方法, 0, '根据申领产生', 1, '往年同期线形参照法', 2, '临近期间平均参照法', 3, '药品储备定额参照法', 4, '药品日销售量参照法', '自定义区间参照法') AS 编制方法 ," & _
        "a.编制人,TO_CHAR(a.编制日期,'YYYY-MM-DD HH24:MI:SS') AS 编制日期, a.审核人, " & _
        "TO_CHAR(a.审核日期,'YYYY-MM-DD HH24:MI:SS') AS 审核日期,a.复核人,TO_CHAR(a.复核日期,'YYYY-MM-DD HH24:MI:SS') AS 复核日期, b.名称 申领药房, a.编制说明 " & _
        " FROM 药品采购计划 A, 部门表 B " & strsql & _
        " WHERE a.药房ID = b.ID(+) And NVL(a.库房ID,0)+0= [11] " & strFind & _
        " ORDER BY A.NO DESC "

    Set rsList = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, _
        SQLCondition.strNO开始, _
        SQLCondition.strNO结束, _
        SQLCondition.date填制时间开始, _
        SQLCondition.date填制时间结束, _
        SQLCondition.date审核时间开始, _
        SQLCondition.date审核时间结束, _
        SQLCondition.str填制人, _
        SQLCondition.str审核人, _
        SQLCondition.lng计划类型, _
        SQLCondition.lng编制方法, _
        mlng库房ID, _
        SQLCondition.date复核时间开始, _
        SQLCondition.date复核时间结束, _
        SQLCondition.str复核人, _
        SQLCondition.lng药品)
    Set vsfList.DataSource = rsList
    With vsfList
        If .rows = 1 Then
            .rows = .rows + 100
            .Row = 1
            .TopRow = 1
            .rows = .rows - 99
        End If
        .Row = 1
        .Col = 0
        .ColSel = .Cols - 1
        
        For n = 0 To .Cols - 1
            .ColKey(n) = .TextMatrix(0, n)
            .FixedAlignment(n) = flexAlignCenterCenter
        Next
    End With
    SetListColWidth
    
    '是否显示申领药房栏
    Call View申领药房(rsList)
    
    vsfList.Redraw = flexRDDirect
    Call FS.StopFlash
    Screen.MousePointer = vbDefault
    SetEnable
    staThis.Panels(2).Text = "当前共有" & rsList.RecordCount & "张单据"
    rsList.Close
    Call vsfList_EnterCell
    vsfList.Row = 1
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub View申领药房(ByVal rsTmp As ADODB.Recordset)
'功能：检查药房ID有无信息，来确定是否显示药房信息栏。
    vsfList.colHidden(vsfList.ColIndex("申领药房")) = True
    If rsTmp.RecordCount <= 0 Then Exit Sub
    With rsTmp
        .MoveFirst
        Do While Not .EOF
            If nvl(!申领药房) <> "" Then
                vsfList.colHidden(vsfList.ColIndex("申领药房")) = False
                Exit Sub
            End If
            .MoveNext
        Loop
    End With
End Sub

'表头列宽初始
Private Sub SetListColWidth()
    Dim intCol As Integer
    
    With vsfList
        .ColAlignment(.ColIndex("NO")) = flexAlignLeftCenter
        .ColAlignment(.ColIndex("计划类型")) = flexAlignLeftCenter
        .ColAlignment(.ColIndex("期间")) = flexAlignLeftCenter
        
        If mblnBootUp = False Then
            For intCol = 0 To .Cols - 1
                .ColWidth(intCol) = 1500
            Next
        End If
        .ColWidth(1) = 0
        
    End With
End Sub

'根据权限设置不同的显示项目
Private Sub SetVisable()
    '药品计划管理所有权限：参数设置、基本、所有库房、登记、修改、删除、验收、清除、取消审核、复核、取消复核、修改执行数量

    If Not zlStr.IsHavePrivs(mstrPrivs, "增加") Then
        mnuEditAdd.Visible = False
        tlbTool.Buttons("Add").Visible = False
    End If
    
    If Not zlStr.IsHavePrivs(mstrPrivs, "修改") Then
        mnuEditModify.Visible = False
        tlbTool.Buttons("Modify").Visible = False
    End If
    
    If Not zlStr.IsHavePrivs(mstrPrivs, "删除") Then
        mnuEditDel.Visible = False
        tlbTool.Buttons("Delete").Visible = False
         '对没有所有编辑权限时，把菜单和工具栏上的相应的分割线屏蔽。
        If mnuEditAdd.Visible = False And mnuEditModify.Visible = False Then
            mnuEditLine1.Visible = False
            tlbTool.Buttons("EditSeparate").Visible = False
        End If
    End If
    
    If Not zlStr.IsHavePrivs(mstrPrivs, "审核") Then
        mnuEditVerify.Visible = False
        tlbTool.Buttons("Verify").Visible = False
    End If
    
    If Not zlStr.IsHavePrivs(mstrPrivs, "清除") Then
        mnuEditClear.Visible = False
        tlbTool.Buttons("Clear").Visible = False
    End If
    
    If Not zlStr.IsHavePrivs(mstrPrivs, "取消审核") And Not zlStr.IsHavePrivs(mstrPrivs, "取消复核") Then
        mnuEditCancel.Visible = False
        tlbTool.Buttons("Cancel").Visible = False
        If mnuEditVerify.Visible = False And mnuEditClear.Visible = False Then
            mnuEditLine2.Visible = False
            tlbTool.Buttons("VerifySeparate").Visible = False
        End If
    End If
    
    If Not zlStr.IsHavePrivs(mstrPrivs, "采购计划打印") Then
        mnuFileBillPrint.Visible = False
        mnuFileBillPreview.Visible = False
    End If

    If Not zlStr.IsHavePrivs(mstrPrivs, "复核") Then
        mnuEditCheck.Visible = False
        tlbTool.Buttons("Check").Visible = False
    End If
    
    If Not zlStr.IsHavePrivs(mstrPrivs, "修改执行数量") Then
        mnuEditExeAmount.Visible = False
    End If
End Sub


Private Sub Cmd查阅_Click()
    Call mnuEditDisplay_Click
End Sub

Private Sub Form_Activate()
    If vsfList.Visible Then
        vsfList.SetFocus
        vsfList.Row = 1
        vsfDetail.Row = IIf(vsfDetail.rows > 1, 1, 0)
    End If
End Sub

Private Sub Form_Load()
    '恢复设置
    Dim strStart As String
    Dim strEnd As String
    Dim strFind As String
    Dim dateCurrentDate As Date
    Dim strTemp As String
    Dim int查询天数 As Integer
    
    mblnBootUp = False
    mlngMode = glngModul
    mstrPrivs = gstrprivs
    
    mblnViewCost = zlStr.IsHavePrivs(mstrPrivs, "查看成本价")
    mint价格显示 = Val(zlDataBase.GetPara("价格显示方式", glngSys, 模块号.药品计划))
    mintPlanPoint = Val(zlDataBase.GetPara("全院计划不管站点", glngSys, mlngMode, 0))
    
    If Not CheckDepend Then
        Unload Me
        Exit Sub
    End If
    
    On Error Resume Next
    '实例化采购平台接口
    If gobjDrugPurchase Is Nothing Then
        Set gobjDrugPurchase = CreateObject("zlDrugPurchase.clsDrugPurchase")
    End If
    Err.Clear
    On Error GoTo 0
    If Not gobjDrugPurchase Is Nothing Then
        mnuEditExport.Visible = True
    End If
    
    mlng库房ID = Me.cboStock.ItemData(Me.cboStock.ListIndex)
    
    Call GetDrugDigit(mlng库房ID, MStrCaption, mintUnit, mintShowCostDigit, mintShowPriceDigit, mintShowNumberDigit, mintShowMoneyDigit)
    
    '组织格式化串
    mstrCostFormat = "'999999999990." & String(mintShowCostDigit, "0") & "'"
    mstrPriceFormat = "'999999999990." & String(mintShowPriceDigit, "0") & "'"
    mstrNumberFormat = "'999999999990." & String(mintShowNumberDigit, "0") & "'"
    mstrMoneyFormat = "'999999999990." & String(mintShowMoneyDigit, "0") & "'"
    
    mlastRow = 0
    SetVisable  '根据权限设置不同的显示项目

    strVerifyStart = "1901-01-01"
    strVerifyEnd = "1901-01-01"
    
    dateCurrentDate = Sys.Currentdate
    
    int查询天数 = Val(zlDataBase.GetPara("查询天数", glngSys, mlngMode, 7)) - 1
    strStart = Format(DateAdd("d", -int查询天数, dateCurrentDate), "yyyy-MM-dd")
    strEnd = Format(dateCurrentDate, "yyyy-MM-dd")
    
    strFind = " AND A.审核日期 is Null And A.编制日期 Between [3] And [4] "
    SQLCondition.date填制时间开始 = CDate(Format(strStart, "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date填制时间结束 = CDate(Format(strEnd, "yyyy-mm-dd") & " 23:59:59")
    
    mstrFind = strFind
    
    lblRange.Caption = "查询范围:" & Format(dateCurrentDate, "yyyy年MM月dd日") & "至" & Format(dateCurrentDate, "yyyy年MM月dd日")
    GetList (mstrFind)  '列出单据头
   
    RestoreWinState Me, App.ProductName, MStrCaption
    Call zlDataBase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs)
    
    '计划业务外挂部件
    Call zlPlugIn_Ini(glngSys, glngModul, mobjPlugIn)
    
    '外挂部件有扩展功能
    Call zlPlugIn_SetVBMenu(glngSys, glngModul, mobjPlugIn, Me)
    
    '外挂部件有扩展功能
    Call zlPlugIn_SetVBToolbar(glngSys, glngModul, mobjPlugIn, Me, tlbTool, "PlugItem", "PlugInSeparator")
        
    mblnBootUp = True
End Sub

Private Sub Form_Resize()
    '窗体位置设置
    
    On Error Resume Next
    
    If Me.WindowState = 1 Then Exit Sub
    
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
    
    With Cmd查阅
        .Left = Me.ScaleWidth - .Width - 100
        .Top = vsfList.Top + vsfList.Height + 30
        .ZOrder
    End With
    
    With vsfDetail
        .Top = picSeparate_s.Top + picSeparate_s.Height + 100
        .Left = 0
        .Height = ScaleHeight - .Top - IIf(staThis.Visible, staThis.Height, 0)
        .Width = cbrTool.Width
    End With
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, MStrCaption
    
    Call zlPlugIn_Unload(mobjPlugIn)
End Sub


Private Sub mnuEditAdd_Click()
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    strNo = ""
    '新增
    frmDrugPlanCard.ShowCard Me, strNo, 1, blnSuccess, cboStock.ItemData(cboStock.ListIndex)
    
    If blnSuccess = True Then
        mnuViewRefresh_Click
    End If
End Sub


Private Sub mnuEditCancel_Click()
    '取消审核
    Dim lngBillId As Long
    Dim intRow As Integer
    Dim intReturn As Integer
    Dim intType As Integer
    
    With vsfList
        
        On Error GoTo errHandle
        intRow = .Row
        lngBillId = .TextMatrix(intRow, 1)
        intType = IIf(.TextMatrix(intRow, .ColIndex("复核人")) = "", 0, 1)
        intReturn = MsgBox("你确实要取消审核单据号为“" & .TextMatrix(.Row, 0) & "”的采购计划单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)

        If intReturn = vbYes Then
            gstrSQL = "zl_药品计划管理_Cancel(" & lngBillId & "," & intType & ")"
            
            If gstrSQL = "" Then Exit Sub
            Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
            
        End If
    End With
    
    Call mnuViewRefresh_Click
    Exit Sub

errHandle:
    Exit Sub
End Sub

Private Sub mnuEditCheck_Click()
    '复核
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    With vsfList
        strNo = .TextMatrix(.Row, 0)
        frmDrugPlanCard.ShowCard Me, strNo, 5, blnSuccess, cboStock.ItemData(cboStock.ListIndex)
    End With
    
    If blnSuccess = True Then
        mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditClear_Click()
    '清除
    Dim lngBillId As Long
    Dim intRow As Integer
    Dim strTitle As String
    Dim intReturn As Integer
    Dim intRecord As Integer
     
    With vsfList
        
        On Error GoTo errHandle
        intRow = .Row
        lngBillId = .TextMatrix(intRow, 1)
        intReturn = MsgBox("你确实要清除单据号为“" & .TextMatrix(.Row, 0) & "”的采购计划单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
        intRecord = .rows - 1
        If intReturn = vbYes Then
            gstrSQL = "zl_药品计划管理_DELETE('" & lngBillId & "')"
            
            If gstrSQL = "" Then Exit Sub
            Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
            
            intRecord = intRecord - 1
            If .rows > 2 Then
                .RemoveItem intRow
            ElseIf .rows = 2 Then
                .rows = 3
                .RemoveItem intRow
                SetEnable
            End If
                
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
        End If
    End With
    
    mlastRow = 0
    Call vsfList_EnterCell
    staThis.Panels(2).Text = "当前共有" & intRecord & "张单据"
    Exit Sub

errHandle:
    Exit Sub
End Sub

Private Sub mnuEditExeAmount_Click()
    '修改执行数量
    
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    With vsfList
        strNo = .TextMatrix(.Row, 0)
        frmDrugPlanCard.ShowCard Me, strNo, 6, blnSuccess, cboStock.ItemData(cboStock.ListIndex)
    
    End With
    If blnSuccess = True Then
        mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditExport_Click()
    gobjDrugPurchase.PurchasePlan gcnOracle ', UserInfo.用户ID
End Sub

Private Sub mnuEditVerify_Click()
    '验收
    
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    With vsfList
        strNo = .TextMatrix(.Row, 0)
        frmDrugPlanCard.ShowCard Me, strNo, 3, blnSuccess, cboStock.ItemData(cboStock.ListIndex)
    
    End With
    If blnSuccess = True Then
        mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditDel_Click()
    '删除
    Dim lngBillId As Long
    Dim intRow As Integer
    Dim strTitle As String
    Dim intReturn As Integer
    Dim intRecord As Integer
     
    With vsfList
        
        On Error GoTo errHandle
        intRow = .Row
        lngBillId = .TextMatrix(intRow, 1)
        intReturn = MsgBox("你确实要删除单据号为“" & .TextMatrix(.Row, 0) & "”的采购计划单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
        intRecord = .rows - 1
        If intReturn = vbYes Then
            gstrSQL = "zl_药品计划管理_DELETE('" & lngBillId & "')"
            
            If gstrSQL = "" Then Exit Sub
            Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
            
            intRecord = intRecord - 1
            If .rows > 2 Then
                .RemoveItem intRow
            ElseIf .rows = 2 Then
                .rows = 3
                .RemoveItem intRow
                
                SetEnable
                
            End If
                
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
        End If
    End With
    
    mlastRow = 0
    Call vsfList_EnterCell
    staThis.Panels(2).Text = "当前共有" & intRecord & "张单据"
    Exit Sub

errHandle:
    Exit Sub
End Sub

Private Sub mnuEditDisplay_Click()
    '查看单据
    
    Dim strNo As String
    With vsfList
        strNo = .TextMatrix(.Row, 0)
        frmDrugPlanCard.ShowCard Me, strNo, 4, , cboStock.ItemData(cboStock.ListIndex)
        
    End With
    
End Sub

Private Sub mnuEditModify_Click()
    '修改
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    blnSuccess = False
    With vsfList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        strNo = .TextMatrix(.Row, 0)
        frmDrugPlanCard.ShowCard Me, strNo, 2, blnSuccess, cboStock.ItemData(cboStock.ListIndex)
        If blnSuccess = True Then
            mnuViewRefresh_Click
        End If
    End With
End Sub

Private Sub mnuFileBillPreview_Click()
    With vsfList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        'If Val(vsfDetail.TextMatrix(1, 9)) = 0 Then
        If mint价格显示 = 1 Then
            '按售价和售价金额显示
            ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1330", "zl8_bill_1330"), Me, "单据编号=" & .TextMatrix(.Row, 0), 1, "ReportFormat=2"
        ElseIf mint价格显示 = 0 Then
            '按成本价和成本金额显示
            ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1330", "zl8_bill_1330"), Me, "单据编号=" & .TextMatrix(.Row, 0), 1, "ReportFormat=1"
        Else
            ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1330", "zl8_bill_1330"), Me, "单据编号=" & .TextMatrix(.Row, 0), 1, "ReportFormat=3"
        End If
    End With
End Sub

Private Sub mnuFileBillPrint_Click()
    With vsfList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        'If Val(vsfDetail.TextMatrix(1, 9)) = 0 Then
        If mint价格显示 = 1 Then
            '按售价和售价金额显示
            ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1330", "zl8_bill_1330"), Me, "单据编号=" & .TextMatrix(.Row, 0), 2, "ReportFormat=2"
        ElseIf mint价格显示 = 0 Then
            '按成本价和成本金额显示
            ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1330", "zl8_bill_1330"), Me, "单据编号=" & .TextMatrix(.Row, 0), 2, "ReportFormat=1"
        Else
            ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1330", "zl8_bill_1330"), Me, "单据编号=" & .TextMatrix(.Row, 0), 2, "ReportFormat=3"
        End If
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
        subPrint 3
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
    Dim dateCurrentDate As Date
    Dim int查询天数 As Integer
    Dim strDept As String
    Dim strTemp As String
    Dim i As Integer
    
    frm参数设置.设置参数 Me, mstrPrivs, MStrCaption
    mint价格显示 = Val(zlDataBase.GetPara("价格显示方式", glngSys, 模块号.药品计划))
    mlastRow = 0
    mintPlanPoint = Val(zlDataBase.GetPara("全院计划不管站点", glngSys, mlngMode, 0))
    With cboStock
        If mintPlanPoint = 1 Or (gstrNodeNo = "-" Or gstrNodeNo = "0") Then
            strDept = ""
            For i = 0 To .ListCount - 1
                If .List(i) <> "全院" Then
                    strDept = strDept & .ItemData(i) & "," & .List(i) & "|"
                End If
            Next
            
            If strDept <> "" Then
                .Clear
                
                .AddItem "全院"
                .ItemData(.NewIndex) = 0
                
                For i = 0 To UBound(Split(strDept, "|")) - 1
                    .AddItem Mid(Split(strDept, "|")(i), InStr(1, Split(strDept, "|")(i), ",") + 1)
                    
                    .ItemData(.NewIndex) = Mid(Split(strDept, "|")(i), 1, InStr(1, Split(strDept, "|")(i), ",") - 1)
                    If Mid(Split(strDept, "|")(i), 1, InStr(1, Split(strDept, "|")(i), ",") - 1) = UserInfo.部门ID Then
                        .ListIndex = .NewIndex
                        mlng库房ID = .NewIndex
                    End If
                Next
            End If
            
            For i = 0 To .ListCount - 1
                If .ItemData(i) = mlng库房ID Then
                    .ListIndex = i
                End If
            Next
        Else
            For i = 0 To .ListCount - 1
                If .List(i) = "全院" Then
                    .RemoveItem i
                    .ListIndex = 0
                    Exit For
                End If
            Next
        End If
    End With
    
    Call GetDrugDigit(mlng库房ID, MStrCaption, mintUnit, mintShowCostDigit, mintShowPriceDigit, mintShowNumberDigit, mintShowMoneyDigit)
        
    '组织格式化串
    mstrCostFormat = "'999999999990." & String(mintShowCostDigit, "0") & "'"
    mstrPriceFormat = "'999999999990." & String(mintShowPriceDigit, "0") & "'"
    mstrNumberFormat = "'999999999990." & String(mintShowNumberDigit, "0") & "'"
    mstrMoneyFormat = "'999999999990." & String(mintShowMoneyDigit, "0") & "'"
    
    dateCurrentDate = Sys.Currentdate
    
    int查询天数 = Val(zlDataBase.GetPara("查询天数", glngSys, mlngMode, 7)) - 1
    strStart = Format(DateAdd("d", -int查询天数, dateCurrentDate), "yyyy-MM-dd")
    strEnd = Format(dateCurrentDate, "yyyy-MM-dd")
    
    SQLCondition.date填制时间开始 = CDate(Format(strStart, "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date填制时间结束 = CDate(Format(strEnd, "yyyy-mm-dd") & " 23:59:59")
    
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

Private Sub mnuPlugItem_Click(index As Integer)
    Call PlugInFun(mnuPlugItem(index).Tag)
End Sub

Private Sub mnuReportItem_Click(index As Integer)
    '默认参数：库房=库房id，开始时间=填制开始时间，结束时间=填制结束时间，NO=计划单NO
    Dim str开始时间 As String
    Dim str结束时间 As String
    Dim strNo As String
    
    If vsfList.Row >= 1 And LTrim(vsfList.TextMatrix(vsfList.Row, 0)) <> "" Then
        strNo = vsfList.TextMatrix(vsfList.Row, 0)
    End If
    
    str开始时间 = IIf(Format(SQLCondition.date填制时间开始, "yyyy-mm-dd") = "1899-12-30", "", Format(SQLCondition.date填制时间开始, "yyyy-mm-dd"))
    str结束时间 = IIf(Format(SQLCondition.date填制时间结束, "yyyy-mm-dd") = "1899-12-30", "", Format(SQLCondition.date填制时间结束, "yyyy-mm-dd"))
    
    Call ReportOpen(gcnOracle, Split(mnuReportItem(index).Tag, ",")(0), Split(mnuReportItem(index).Tag, ",")(1), Me, _
        "库房=" & IIf(Val(cboStock.ItemData(cboStock.ListIndex)) = 0, "", Val(cboStock.ItemData(cboStock.ListIndex))), _
        "开始时间=" & str开始时间, _
        "结束时间=" & str结束时间, _
        "NO=" & strNo)
End Sub

Private Sub mnuViewColDefine_Click()
    Dim strColumn_All As String, strColumn_Select As String, strColumn_UnSelect As String
    Dim str选择列 As String
    Dim str屏蔽列 As String
    Dim strAllCol As String
    Dim arr总列, arr设置列
    
    On Error Resume Next
    
    Select Case mlngMode
    Case 模块号.药品计划           '药品计划管理
        strColumn_All = "药名,0|商品名,0|药品来源,1|规格,1|生产商,0|原产地,1|单位,1|医保类型,1|前期数量,1|上期数量,1|库存上限,1|库存下限,1|" & _
                        "库存数量,0|上期销量,1|本期销量,1|计划数量,0|执行数量,0|送货单位,1|送货数量,1|成本价,0|成本金额,0|售价,0|售价金额,0|上次供应商,1|说明,1|基本药物,1|批准文号,1"
        str选择列 = "药名|商品名|药品来源|规格|生产商|原产地|单位|医保类型|前期数量|上期数量|库存上限|库存下限|库存数量|上期销量|本期销量|计划数量|执行数量|送货单位|送货数量|成本价|成本金额|售价|售价金额|上次供应商|说明|基本药物|批准文号"
        str屏蔽列 = ""
    End Select
    
    '取已选择列的信息
    strColumn_Select = zlDataBase.GetPara("选择列", glngSys, mlngMode, "")
    strColumn_UnSelect = zlDataBase.GetPara("屏蔽列", glngSys, mlngMode, "")
    
    If strColumn_Select <> "" Then
        If strColumn_UnSelect <> "" Then
            strAllCol = strColumn_Select & "|" & strColumn_UnSelect
        Else
            strAllCol = strColumn_Select
        End If
        arr总列 = Split(strColumn_All, "|")
        arr设置列 = Split(strAllCol, "|")
        If UBound(arr总列) <> UBound(arr设置列) Then
            strColumn_Select = "药名|商品名|药品来源|规格|生产商|原产地|单位|医保类型|前期数量|上期数量|库存上限|库存下限|库存数量|上期销量|本期销量|计划数量|执行数量|送货单位|送货数量|成本价|成本金额|售价|售价金额|上次供应商|说明|基本药物|批准文号"
            strColumn_UnSelect = ""
            zlDataBase.SetPara "选择列", strColumn_Select, glngSys, 模块号.药品计划
            zlDataBase.SetPara "屏蔽列", strColumn_UnSelect, glngSys, 模块号.药品计划
        End If
    Else
        strColumn_Select = "药名|商品名|药品来源|规格|生产商|原产地|单位|医保类型|前期数量|上期数量|库存上限|库存下限|库存数量|上期销量|本期销量|计划数量|执行数量|送货单位|送货数量|成本价|成本金额|售价|售价金额|上次供应商|说明|基本药物|批准文号"
        strColumn_UnSelect = ""
        zlDataBase.SetPara "选择列", strColumn_Select, glngSys, 模块号.药品计划
        zlDataBase.SetPara "屏蔽列", strColumn_UnSelect, glngSys, 模块号.药品计划
    End If
    
    If Not frm列设置.ShowME(Me, strColumn_All, strColumn_Select) Then Exit Sub
    
    zlDataBase.SetPara "选择列", Split(strColumn_Select, "||")(0), glngSys, mlngMode
    zlDataBase.SetPara "屏蔽列", Split(strColumn_Select, "||")(1), glngSys, mlngMode
End Sub

Private Sub mnuViewRefresh_Click()
    '刷新
    mlastRow = 0
    GetList mstrFind
End Sub

Private Sub mnuViewSearch_Click()
    '查找
    Dim strFind As String
    
    strFind = FrmDrugPlanSearch.GetSearch(Me, strStart, strEnd, strVerifyStart, strVerifyEnd, _
                SQLCondition.strNO开始, _
                SQLCondition.strNO结束, _
                SQLCondition.date填制时间开始, _
                SQLCondition.date填制时间结束, _
                SQLCondition.date审核时间开始, _
                SQLCondition.date审核时间结束, _
                SQLCondition.date复核时间开始, _
                SQLCondition.date复核时间结束, _
                SQLCondition.str填制人, _
                SQLCondition.str审核人, _
                SQLCondition.str复核人, _
                SQLCondition.lng计划类型, _
                SQLCondition.lng编制方法, _
                SQLCondition.lng药品)
    
    If strFind <> "" Then
        mstrFind = strFind
        mlastRow = 0
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
    Dim strSql药名 As String
    Dim n As Integer
    Dim intCol As Integer
    Dim strUnit As String

    If mlastRow = vsfList.Row Then Exit Sub
    mlastRow = vsfList.Row
    
    On Error GoTo errHandle
    If vsfList.Row >= 1 And LTrim(vsfList.TextMatrix(vsfList.Row, 0)) <> "" Then
        vsfList.Col = 0
        vsfList.ColSel = vsfList.Cols - 1

        If gint药品名称显示 = 0 Then
            strSql药名 = ",('['||D.编码||']'||D.通用名) AS 药品信息"
        ElseIf gint药品名称显示 = 1 Then
            strSql药名 = ",('['||D.编码||']'||NVL(D.商品名,D.通用名)) AS 药品信息"
        Else
            strSql药名 = ",('['||D.编码||']'||D.通用名) AS 药品信息,D.商品名"
        End If
        Select Case mintUnit '单位系数：1-售价;2-门诊;3-住院;4-药库
        Case 1
            gstrSQL = "decode(d.送货单位,null,'',d.送货单位|| '(1'||d.送货单位||'='||d.送货包装/1|| d.计算单位 ||')') as 送货单位,to_char(b.送货数量,'999999999990.0') as 送货数量,"
            strUnit = "1,"
        Case 2
            gstrSQL = "decode(d.送货单位,null,'',d.送货单位||'（1'||d.送货单位||'='||d.送货包装/d.门诊包装|| d.门诊单位 ||')') as 送货单位,to_char(b.送货数量,'999999999990.0') as 送货数量,"
            strUnit = "d.门诊包装,"
        Case 3
            gstrSQL = "decode(d.送货单位,null,'',d.送货单位||'（1'||d.送货单位||'='||d.送货包装/d.住院包装|| d.住院单位 ||')') as 送货单位,to_char(b.送货数量,'999999999990.0') as 送货数量,"
            strUnit = "d.住院包装,"
        Case Else
            gstrSQL = "decode(d.送货单位,null,'',d.送货单位||'（1'||d.送货单位||'='||d.送货包装/d.药库包装|| d.药库单位 ||')') as 送货单位,to_char(b.送货数量,'999999999990.0') as 送货数量,"
            strUnit = "d.药库包装,"
        End Select
        
        gstrSQL = "SELECT B.序号" & strSql药名 & ",D.药品来源,D.规格, Decode(" & mintUnit & ", 1, d.计算单位, 2, d.门诊单位, 3, d.住院单位, d.药库单位) As 单位,d.医保类型," & _
                " TRIM(TO_CHAR(B.前期数量 / " & strUnit & mstrNumberFormat & ")) 前期数量," & _
                " TRIM(TO_CHAR(B.上期数量 / " & strUnit & mstrNumberFormat & ")) 上期数量," & _
                " TRIM(TO_CHAR(B.库存数量 / " & strUnit & mstrNumberFormat & ")) 库存数量," & _
                " TRIM(TO_CHAR(B.上期销量 / " & strUnit & mstrNumberFormat & ")) 上期销量," & _
                " TRIM(TO_CHAR(B.本期销量 / " & strUnit & mstrNumberFormat & ")) 本期销量," & _
                " TRIM(TO_CHAR(B.计划数量 / " & strUnit & mstrNumberFormat & ")) 计划数量," & _
                gstrSQL & _
                " Trim(To_Char(B.单价 * " & strUnit & mstrCostFormat & ")) 成本价," & _
                " Trim(To_Char(B.金额, " & mstrMoneyFormat & ")) 成本金额, " & _
                " Trim(To_Char(B.售价 * " & strUnit & mstrPriceFormat & ")) 售价, " & _
                " Trim(To_Char(B.售价金额, " & mstrMoneyFormat & ")) 售价金额, " & _
                " B.上次供应商,B.上次生产商,D.原产地,NVL(B.说明,'') 说明, " & _
                " TRIM(TO_CHAR(B.执行数量 / " & strUnit & mstrNumberFormat & ")) 执行数量,b.批准文号 " & _
                " FROM 药品采购计划 A, 药品计划内容 B,部门表 C," & _
                "     (SELECT DISTINCT A.药品ID, F.编码,F.名称 As 通用名,B.名称 As 商品名,f.费用类型 As 医保类型,A.药品来源,f.计算单位,A.住院包装,A.门诊包装,A.药库包装," & _
                "      F.规格, a.门诊单位,a.住院单位,A.药库单位,a.送货单位,a.送货包装,a.原产地 " & _
                "     FROM 药品规格 A, 收费项目别名 B, 收费项目目录 F " & _
                "     WHERE A.药品ID = B.收费细目ID(+) AND B.性质(+)=3 " & _
                "     AND A.药品ID = F.ID) D " & _
                " WHERE A.ID = B.计划ID AND NVL(A.库房ID,0)=C.ID(+) " & _
                " AND B.药品ID=D.药品ID AND B.计划ID = [1] " & IIf(SQLCondition.lng药品 > 0, " And B.药品ID=[2] ", "") & _
                " ORDER BY 序号"
        Set rsDetail = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, Val(vsfList.TextMatrix(vsfList.Row, 1)), SQLCondition.lng药品)
        
        vsfDetail.Redraw = flexRDNone
        
        Set vsfDetail.DataSource = rsDetail
        rsDetail.Close
        
        With vsfDetail
            .Row = IIf(.rows > 1, 1, 0)
            .Col = 0
            .ColSel = .Cols - 1
            
            For n = 0 To .Cols - 1
                .ColKey(n) = .TextMatrix(0, n)
                .FixedAlignment(n) = flexAlignCenterCenter
            Next
            
            If Trim(vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("编制方法"))) = "自定义区间参照法" Then
                .TextMatrix(0, .ColIndex("前期数量")) = "本期数量"
                .TextMatrix(0, .ColIndex("上期数量")) = "本期销量"
                .TextMatrix(0, .ColIndex("上期销量")) = "上月销量"
                .TextMatrix(0, .ColIndex("本期销量")) = "本月销量"
            End If
            
            .Redraw = flexRDDirect
        End With
    ElseIf LTrim(vsfList.TextMatrix(vsfList.Row, 0)) = "" Then
        With vsfDetail
            .Redraw = flexRDNone
            .Cols = IIf(gint药品名称显示 = 2, 23, 22)
            .rows = 2
            .Clear
            
            intCol = 0
            
            .TextMatrix(0, intCol) = "序号": intCol = intCol + 1
            .TextMatrix(0, intCol) = "药品信息": intCol = intCol + 1
            
            If gint药品名称显示 = 2 Then
                .TextMatrix(0, intCol) = "商品名": intCol = intCol + 1
            End If
            
            .TextMatrix(0, intCol) = "药品来源": intCol = intCol + 1
            .TextMatrix(0, intCol) = "规格": intCol = intCol + 1
            .TextMatrix(0, intCol) = "单位": intCol = intCol + 1
            .TextMatrix(0, intCol) = "医保类型": intCol = intCol + 1
            .TextMatrix(0, intCol) = "前期数量": intCol = intCol + 1
            .TextMatrix(0, intCol) = "上期数量": intCol = intCol + 1
            .TextMatrix(0, intCol) = "库存数量": intCol = intCol + 1
            .TextMatrix(0, intCol) = "上期销量": intCol = intCol + 1
            .TextMatrix(0, intCol) = "本期销量": intCol = intCol + 1
            .TextMatrix(0, intCol) = "计划数量": intCol = intCol + 1
            .TextMatrix(0, intCol) = "成本价": intCol = intCol + 1
            .TextMatrix(0, intCol) = "成本金额": intCol = intCol + 1
            .TextMatrix(0, intCol) = "售价": intCol = intCol + 1
            .TextMatrix(0, intCol) = "售价金额": intCol = intCol + 1
            .TextMatrix(0, intCol) = "上次供应商": intCol = intCol + 1
            .TextMatrix(0, intCol) = "上次生产商": intCol = intCol + 1
            .TextMatrix(0, intCol) = "原产地": intCol = intCol + 1
            .TextMatrix(0, intCol) = "说明": intCol = intCol + 1
            .TextMatrix(0, intCol) = "执行数量": intCol = intCol + 1
            .TextMatrix(0, intCol) = "批准文号": intCol = intCol + 1
            
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
    SetEnable
    
    With vsfDetail
        If .rows <= 1 Then Exit Sub
        If .TextMatrix(1, 0) <> "" Then
            If mint价格显示 = 0 Then
                vsfDetail.ColWidth(.ColIndex("售价")) = 0
                vsfDetail.ColWidth(.ColIndex("售价金额")) = 0
            ElseIf mint价格显示 = 1 Then
                vsfDetail.ColWidth(.ColIndex("成本价")) = 0
                vsfDetail.ColWidth(.ColIndex("成本金额")) = 0
            End If
        End If
        If mblnViewCost = False Then
            .ColWidth(.ColIndex("成本价")) = 0
            .ColWidth(.ColIndex("成本金额")) = 0
        End If
    End With
    vsfDetail.Row = 1
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetDetailColWidth()
    Dim rsDetail As New Recordset
    Dim bln中药库房 As Boolean
    Dim str库房性质 As String
    
    On Error GoTo errHandle
    
    With vsfDetail
        .ColWidth(.ColIndex("序号")) = 500
        .ColWidth(.ColIndex("药品信息")) = 1500
        .ColWidth(.ColIndex("药品来源")) = 1000
        .ColWidth(.ColIndex("规格")) = 800
        .ColWidth(.ColIndex("单位")) = 800
        .ColWidth(.ColIndex("前期数量")) = 1200
        .ColWidth(.ColIndex("上期数量")) = 1200
        .ColWidth(.ColIndex("库存数量")) = 1200
        .ColWidth(.ColIndex("上期销量")) = 1200
        .ColWidth(.ColIndex("本期销量")) = 1200
        .ColWidth(.ColIndex("计划数量")) = 1200
        .ColWidth(.ColIndex("成本价")) = 1200
        .ColWidth(.ColIndex("成本金额")) = 1200
        .ColWidth(.ColIndex("售价")) = 1200
        .ColWidth(.ColIndex("售价金额")) = 1200
        .ColWidth(.ColIndex("上次供应商")) = 1200
        .ColWidth(.ColIndex("上次生产商")) = 1200
        .ColWidth(.ColIndex("原产地")) = 1200
        .ColWidth(.ColIndex("说明")) = 1200
        .ColWidth(.ColIndex("执行数量")) = 1200
        .ColWidth(.ColIndex("批准文号")) = 1200
        .ColAlignment(.ColIndex("前期数量")) = flexAlignRightCenter
        .ColAlignment(.ColIndex("上期数量")) = flexAlignRightCenter
        .ColAlignment(.ColIndex("库存数量")) = flexAlignRightCenter
        .ColAlignment(.ColIndex("上期销量")) = flexAlignRightCenter
        .ColAlignment(.ColIndex("本期销量")) = flexAlignRightCenter
        .ColAlignment(.ColIndex("计划数量")) = flexAlignRightCenter
        .ColAlignment(.ColIndex("成本价")) = flexAlignRightCenter
        .ColAlignment(.ColIndex("成本金额")) = flexAlignRightCenter
        .ColAlignment(.ColIndex("售价")) = flexAlignRightCenter
        .ColAlignment(.ColIndex("售价金额")) = flexAlignRightCenter
        .ColAlignment(.ColIndex("执行数量")) = flexAlignRightCenter
        If .TextMatrix(1, 0) <> "" Then
            '0:显示成本价;  1:显示售价;  2:显示成本和售价
            If mint价格显示 = 0 Then
                .ColWidth(.ColIndex("售价")) = 0
                .ColWidth(.ColIndex("售价金额")) = 0
            ElseIf mint价格显示 = 1 Then
                .ColWidth(.ColIndex("成本价")) = 0
                .ColWidth(.ColIndex("成本金额")) = 0
            End If
        End If
        If mblnViewCost = False Then
            .ColWidth(.ColIndex("成本价")) = 0
            .ColWidth(.ColIndex("成本金额")) = 0
        End If
        
        str库房性质 = ""
        gstrSQL = "Select a.工作性质 From 部门性质说明 A Where a.部门id =[1]"
        Set rsDetail = zlDataBase.OpenSQLRecord(gstrSQL, "判断是库房性质", cboStock.ItemData(cboStock.ListIndex))
        Do While Not rsDetail.EOF
            str库房性质 = str库房性质 & "," & rsDetail!工作性质
            rsDetail.MoveNext
        Loop
        If str库房性质 Like "*中药*" Or str库房性质 Like "*制剂室*" Then bln中药库房 = True
        If bln中药库房 Then
            .colHidden(.ColIndex("原产地")) = False
        Else
            .colHidden(.ColIndex("原产地")) = True
        End If
    End With
    
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
        .Height = picSeparate_s.Top - .Top
    End With
    
    With Cmd查阅
        .Top = vsfList.Top + vsfList.Height + 30
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
            mnuEditVerify_Click
        Case "Clear"
            mnuEditClear_Click
        Case "Cancel"
            mnuEditCancel_Click
        Case "Check"
            mnuEditCheck_Click
        Case "Search"
            mnuViewSearch_Click
        Case "Refresh"
            mnuViewRefresh_Click
        Case "Help"
            mnuHelpTitle_Click
        Case "Exit"
            mnufileexit_Click
        Case Else
            'zlPlugIn外挂功能
            If Button.Key Like "PlugItem*" Then
                Call PlugInFun(Button.Caption)
            End If
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
            
            If mnuEditClear.Visible = True Then
                mnuEditClear.Enabled = False
                tlbTool.Buttons("Clear").Enabled = False
            End If
            
            If mnuEditCancel.Visible = True Then
                mnuEditCancel.Enabled = False
                tlbTool.Buttons("Cancel").Enabled = False
            End If
            
            If mnuEditDisplay.Visible = True Then
                mnuEditDisplay.Enabled = False
            End If
            
            If mnuEditCheck.Visible = True Then
                mnuEditCheck.Enabled = False
                tlbTool.Buttons("Check").Enabled = False
            End If
            
            If mnuEditExeAmount.Visible = True Then
                mnuEditExeAmount.Enabled = False
            End If
        Else
            mnuFilePreView.Enabled = True
            mnuFilePrint.Enabled = True
            mnuFileBillPreview.Enabled = True
            mnuFileBillPrint.Enabled = True
            mnuFileExcel.Enabled = True
            tlbTool.Buttons("Print").Enabled = True
            tlbTool.Buttons("PrintView").Enabled = True
            
            If .TextMatrix(.Row, .ColIndex("审核人")) = "" Then    '未审核单
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
                
                If mnuEditClear.Visible = True Then
                    mnuEditClear.Enabled = False
                    tlbTool.Buttons("Clear").Enabled = False
                End If
                
                If mnuEditCancel.Visible = True Then
                    mnuEditCancel.Enabled = False
                    tlbTool.Buttons("Cancel").Enabled = False
                End If
            
                If mnuEditDisplay.Visible = True Then
                    mnuEditDisplay.Enabled = True
                End If
                
                If mnuEditCheck.Visible = True Then
                    mnuEditCheck.Enabled = False
                    tlbTool.Buttons("Check").Enabled = False
                End If
                
                If mnuEditExeAmount.Visible = True Then
                    mnuEditExeAmount.Enabled = False
                End If
            Else    '审核单
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
                
                If mnuEditDisplay.Visible = True Then
                    mnuEditDisplay.Enabled = True
                End If
                
                If mnuEditClear.Visible = True Then
                    mnuEditClear.Enabled = True
                    tlbTool.Buttons("Clear").Enabled = True
                End If
                
                If mnuEditExeAmount.Visible = True Then
                    mnuEditExeAmount.Enabled = True
                End If
                
                If .TextMatrix(.Row, .ColIndex("复核人")) = "" Then    '未复核单
                    If mnuEditCancel.Visible = True And zlStr.IsHavePrivs(mstrPrivs, "取消审核") Then
                        mnuEditCancel.Enabled = True
                        tlbTool.Buttons("Cancel").Enabled = True
                    End If
                    
                    If mnuEditCheck.Visible = True Then
                        mnuEditCheck.Enabled = True
                        tlbTool.Buttons("Check").Enabled = True
                    End If
                Else
                    '已复核单
                    If mnuEditCancel.Visible = True And zlStr.IsHavePrivs(mstrPrivs, "取消复核") Then
                        mnuEditCancel.Enabled = True
                        tlbTool.Buttons("Cancel").Enabled = True
                    End If
                    
                    If mnuEditCheck.Visible = True Then
                        mnuEditCheck.Enabled = False
                        tlbTool.Buttons("Check").Enabled = False
                    End If
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
    
    objPrint.Title.Text = MStrCaption
        
    objRow.Add "时间：" & strRange
    objRow.Add "部门：" & cboStock.Text
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow

    objRow.Add "打印人:" & UserInfo.用户姓名
    objRow.Add "打印日期:" & Format(Sys.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    
    If Me.ActiveControl Is vsfDetail Then
        Set objPrint.Body = vsfDetail
    Else
        Set objPrint.Body = vsfList
    End If
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
    Select Case ButtonMenu.Key
        Case "Payment"
'            mnuEditAddPayment_Click
        Case "Imprest"
'            mnuEditAddImprest_Click
    End Select
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


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

