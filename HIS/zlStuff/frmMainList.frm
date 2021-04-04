VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMainList 
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
   Begin VSFlex8Ctl.VSFlexGrid vsfCostlyInfo 
      Height          =   615
      Left            =   6360
      TabIndex        =   8
      Top             =   3480
      Width           =   2295
      _cx             =   4048
      _cy             =   1085
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
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
      ExplorerBar     =   0
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
      VirtualData     =   -1  'True
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
   Begin TabDlg.SSTab TabShow 
      Height          =   345
      Left            =   255
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   609
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "移出库房(&0)"
      TabPicture(0)   =   "frmMainList.frx":014A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "移入库房(&1)"
      TabPicture(1)   =   "frmMainList.frx":0166
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
   End
   Begin MSComctlLib.ImageList ilsHot 
      Left            =   7290
      Top             =   1830
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
            Picture         =   "frmMainList.frx":0182
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":03A2
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":05C2
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":07DE
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":09FE
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":0C1E
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":0E3A
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":1056
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":1270
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":13CA
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":15EA
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":180A
            Key             =   "Prepare"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":1F04
            Key             =   "Send"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":211E
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":2338
            Key             =   "Cancel"
            Object.Tag             =   "Cancel"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Cmd查阅 
      Caption         =   "查阅(&V)"
      Height          =   350
      Left            =   5250
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2700
      Width           =   1100
   End
   Begin VB.PictureBox picSeparate_s 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   600
      MousePointer    =   7  'Size N S
      ScaleHeight     =   630
      ScaleWidth      =   4815
      TabIndex        =   4
      Top             =   2685
      Width           =   4815
   End
   Begin ComCtl3.CoolBar cbrTool 
      Height          =   1125
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   1984
      BandCount       =   2
      _CBWidth        =   9480
      _CBHeight       =   1125
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
         Left            =   585
         TabIndex        =   2
         Text            =   "cboStock"
         Top             =   780
         Width           =   8805
      End
      Begin MSComctlLib.Toolbar tlbTool 
         Height          =   720
         Left            =   165
         TabIndex        =   1
         Top             =   30
         Width           =   9225
         _ExtentX        =   16272
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
            NumButtons      =   21
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
               Object.Visible         =   0   'False
               Key             =   "PrepareSplit"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "核查"
               Key             =   "Check"
               Object.ToolTipText     =   "核查"
               Object.Tag             =   "核查"
               ImageKey        =   "Prepare"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "取消"
               Key             =   "CancelCheck"
               Object.ToolTipText     =   "取消核查"
               Object.Tag             =   "取消"
               ImageKey        =   "Cancel"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "备料"
               Key             =   "Prepare"
               Object.ToolTipText     =   "备料"
               Object.Tag             =   "备料"
               ImageKey        =   "Prepare"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "发料"
               Key             =   "Send"
               Object.ToolTipText     =   "发料"
               Object.Tag             =   "发料"
               ImageKey        =   "Send"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "回退"
               Key             =   "Back"
               Object.ToolTipText     =   "回退"
               Object.Tag             =   "回退"
               ImageKey        =   "Back"
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "EditSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "审核"
               Key             =   "Verify"
               Description     =   "审核"
               Object.ToolTipText     =   "审核"
               Object.Tag             =   "审核"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "冲销"
               Key             =   "Strike"
               Description     =   "冲销"
               Object.ToolTipText     =   "冲销"
               Object.Tag             =   "冲销"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "VerifySeparate"
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "过滤"
               Key             =   "Search"
               Description     =   "过滤"
               Object.ToolTipText     =   "过滤"
               Object.Tag             =   "过滤"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "刷新"
               Key             =   "Refresh"
               Description     =   "刷新"
               Object.ToolTipText     =   "刷新"
               Object.Tag             =   "刷新"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FindSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助主题"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Exit"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   11
            EndProperty
         EndProperty
         MouseIcon       =   "frmMainList.frx":2552
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
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
            Picture         =   "frmMainList.frx":286C
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
      Left            =   6675
      Top             =   1845
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
            Picture         =   "frmMainList.frx":3100
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":3320
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":3540
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":375C
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":397C
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":3B9C
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":3DB8
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":3FD4
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":41EE
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":4348
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":4564
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":4784
            Key             =   "Prepare"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":4E7E
            Key             =   "Send"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":5098
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":52B2
            Key             =   "Cancel"
            Object.Tag             =   "Cancel"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
      Height          =   1455
      Left            =   480
      TabIndex        =   9
      Top             =   1320
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   2566
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   -2147483628
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      RowSizingMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VSFlex8Ctl.VSFlexGrid mshDetail 
      Height          =   945
      Left            =   360
      TabIndex        =   10
      Top             =   3480
      Width           =   5985
      _cx             =   10557
      _cy             =   1667
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
      BackColor       =   -2147483628
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483628
      GridColor       =   12632256
      GridColorFixed  =   0
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483628
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmMainList.frx":54CC
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
      ExplorerBar     =   7
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
      Begin VB.Image imgLeft 
         Height          =   240
         Left            =   30
         Picture         =   "frmMainList.frx":55A1
         Top             =   45
         Width           =   240
      End
   End
   Begin VB.Label lblCostly 
      BackColor       =   &H8000000A&
      Caption         =   "高值材料信息"
      Height          =   195
      Left            =   6360
      TabIndex        =   7
      Top             =   3300
      Width           =   1455
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
      Begin VB.Menu mnuEditCheckBatch 
         Caption         =   "备货卫材批量核查(&H)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditCheck 
         Caption         =   "核查(&H)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditCancelCheck 
         Caption         =   "取消核查(&Q)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditCheckLine 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditVerifyBatch 
         Caption         =   "备货卫材批量审核(&T)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditVerify 
         Caption         =   "审核(&C)"
      End
      Begin VB.Menu mnuEditStrike 
         Caption         =   "冲销(&K)"
      End
      Begin VB.Menu mnuEditLine0 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditRestore 
         Caption         =   "卫材退货(&R)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditBill 
         Caption         =   "修改发票信息(&B)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditReg 
         Caption         =   "修改注册证号(&G)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditAcc 
         Caption         =   "财务审核(&V)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditImport 
         Caption         =   "导入计划单(&I)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditImportFile 
         Caption         =   "导入外部文件(&F)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditPrepare 
         Caption         =   "备料(&P)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditSend 
         Caption         =   "发料(&S)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditBack 
         Caption         =   "回退(&B)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditPrePareSp 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditVerifySelect 
         Caption         =   "财务审核单查询(&Y)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditDisplay 
         Caption         =   "查看单据(&W)"
      End
      Begin VB.Menu mnuEditLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditTMPrint 
         Caption         =   "卫材条码打印管理"
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
      Begin VB.Menu mnuViewLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewColDefine 
         Caption         =   "列选择(&C)"
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
Private mstrPrivs As String                     '权限
Private mblnFirst As Boolean
Private mblnPopupmenuCall As Boolean
Private mstrOrder As String             '记录排序方式
Private mStr库房 As String              '记录当前操作员所能操作的所有库房
Private mbln申领核查 As Boolean     '单据是否需要核查 true-需要 false-不需要
Private mintFindDay As Integer      '查询天数范围

'---------------------------------------------------------------------------------------------------------
'设置相关的过滤条件:2008-08-22 16:35:52
'刘兴宏:
Private mblnNoClick As Boolean
Private mstr工作性质 As String
Private mbln操作员限制 As Boolean

'----------------------------------------------------------------------------------------------------------
'刘兴宏:增加小数位数的格式串
'修改:2007/03/06
Private mOraFMT As g_FmtString
Private mFMT As g_FmtString

'----------------------------------------------------------------------------------------------------------
'日期设置
Private mdtStartDate As Date
Private mdtEndDate As Date
Private mdtVerifyStart As Date
Private mdtVerifyEnd As Date
Private mintOldY As Integer
Private mintUnit As Integer                 '0：散装单位；1：包装单位
Private mstrPrintRange As String      '打印范围文本
Private mstrMoneySum As String        '金额合计
Private mint有发票 As Integer
Private mint无发票 As Integer
Private mstrOthers() As String ' 0-记录状态(计划类型),1-开始单号,2-结束单号,3-材料id,4-对方部门id(或入出类别id或编制方法(计划用)),5-填制人,6-审核人,7-供应商ID,8-生产商,9-开始生产日期,10-结束生产日期,11-开始发票号,12-结束发票号,13-条码信息
'------------------------------------------------------------------------------------------------------------------------
'--刘兴宏:20060803,问题:8740
Private mbln向发料部门领用 As Boolean       '领用有效
Private mbln只具备普通科室       As Boolean       '当前人员所在科室只具备普通科室,领用有效
'------------------------------------------------------------------------------------------------------------------------

Private mint移库处理流程 As Integer                    '1-需要备料、发送、接收这一过程  0-不需要这一过程
Private mbln需要核查    As Boolean              '只针对外购入库
Private mint领用审核方式 As Integer             '领用审核：0－普通审核；1－需要先财务审核
Private mstr高值耗材 As String              '记录过滤条件中是否选择了高值耗材
Private mint冲销申请 As Integer             '0-不需要申请冲销 1-需要申请冲销
Private mint冲销方式 As Integer             '0－正常冲销方式；1－产生冲销申请单据；2－审核已产生的冲销申请单据
Private mblnCostView As Boolean             '查看成本价相关信息 true-允许查看 false-不允许查看
Private mbln移库明确批次 As Boolean         '是否明确批次，仅对移库单有效
Public Sub SetMenu()
    '隐藏备料、发送、审核与冲销
    If mlngMode <> 1716 Then Exit Sub
    
    mnuEditPrepare.Visible = False
    mnuEditSend.Visible = False
    mnuEditBack.Visible = False
    
    tlbTool.Buttons("Prepare").Visible = False
    tlbTool.Buttons("Send").Visible = False
    tlbTool.Buttons("Back").Visible = False
    
    mnuEditVerify.Visible = False
    mnuEditStrike.Visible = False
    tlbTool.Buttons("Verify").Visible = False
    tlbTool.Buttons("Strike").Visible = False
 
    mnuEditLine1.Visible = False
    mnuEditLine0.Visible = False
    mnuEditLine2.Visible = False
    mnuEditPrePareSp.Visible = False
    tlbTool.Buttons("EditSeparate").Visible = False
    tlbTool.Buttons("VerifySeparate").Visible = False
    
    '根据当前页面开启
    If TabShow.Tab = 0 Then
        If mlngMode = 1716 Then
            mint移库处理流程 = IIf(Val(zlDatabase.GetPara("移库流程", glngSys, mlngMode, "0", , , , cboStock.ItemData(cboStock.ListIndex))) = 1, 1, 0)
            
            If mint移库处理流程 = 0 Then
                mnuEditPrepare.Visible = False
                mnuEditVerify.Visible = zlStr.IsHavePrivs(mstrPrivs, "审核")
                mnuEditStrike.Visible = zlStr.IsHavePrivs(mstrPrivs, "冲销")
                mnuEditLine0.Visible = mnuEditVerify.Visible Or mnuEditAdd.Visible Or mnuEditModify.Visible Or mnuEditDel.Visible
                mnuEditLine1.Visible = mnuEditVerify.Visible And (mnuEditAdd.Visible Or mnuEditModify.Visible Or mnuEditDel.Visible)
                tlbTool.Buttons("Verify").Visible = mnuEditVerify.Visible
                tlbTool.Buttons("Strike").Visible = mnuEditStrike.Visible
                
                tlbTool.Buttons("VerifySeparate").Visible = mnuEditLine0.Visible
                 tlbTool.Buttons("PrintSeparate").Visible = mnuEditLine0.Visible
                mnuEditVerify.Caption = "审核(&C)"
                tlbTool.Buttons("Verify").Caption = "审核"
                tlbTool.Buttons("Verify").Tag = "审核"
                tlbTool.Buttons("Verify").ToolTipText = "审核"
            Else
                mnuEditVerify.Caption = "接收(&C)"
                tlbTool.Buttons("Verify").Caption = "接收"
                tlbTool.Buttons("Verify").Tag = "接收"
                tlbTool.Buttons("Verify").ToolTipText = "接收"
                mnuEditPrepare.Visible = zlStr.IsHavePrivs(mstrPrivs, "发送")
                mnuEditLine1.Visible = mnuEditPrepare.Visible
                mnuEditPrePareSp.Visible = mnuEditPrepare.Visible
            End If
            
            mint冲销申请 = IIf(Val(zlDatabase.GetPara("冲销申请", glngSys, mlngMode, "0")) = 1, 1, 0)
            If mint冲销申请 = 0 Then
                tlbTool.Buttons("Strike").Visible = False
                tlbTool.Buttons("VerifySeparate").Visible = False
                mnuEditStrike.Caption = "冲销(&K)"
            Else
                tlbTool.Buttons("Strike").ToolTipText = "审核冲销"
                tlbTool.Buttons("Strike").Caption = "审核冲销"
                tlbTool.Buttons("Strike").Tag = "审核冲销"
                tlbTool.Buttons("Strike").Visible = zlStr.IsHavePrivs(mstrPrivs, "冲销")
                tlbTool.Buttons("VerifySeparate").Visible = tlbTool.Buttons("Strike").Visible
                mnuEditStrike.Caption = "审核冲销(&K)"
            End If
            
            mnuEditSend.Visible = mnuEditPrepare.Visible
            mnuEditBack.Visible = mnuEditPrepare.Visible
            tlbTool.Buttons("Prepare").Visible = mnuEditPrepare.Visible
            tlbTool.Buttons("Send").Visible = mnuEditPrepare.Visible
            tlbTool.Buttons("Back").Visible = mnuEditPrepare.Visible
            tlbTool.Buttons("EditSeparate").Visible = mnuEditPrepare.Visible
        End If
    Else
        If mlngMode = 1716 Then '移库
            mint冲销申请 = IIf(Val(zlDatabase.GetPara("冲销申请", glngSys, mlngMode, "0")) = 1, 1, 0)
            If mint冲销申请 = 0 Then
                tlbTool.Buttons("Strike").ToolTipText = "冲销"
                tlbTool.Buttons("Strike").Caption = "冲销"
                tlbTool.Buttons("Strike").Tag = "冲销"
                mnuEditStrike.Caption = "冲销(&K)"
            Else
                tlbTool.Buttons("Strike").ToolTipText = "申请冲销"
                tlbTool.Buttons("Strike").Caption = "申请冲销"
                tlbTool.Buttons("Strike").Tag = "申请冲销"
                mnuEditStrike.Caption = "申请冲销(&K)"
            End If
        End If
        mnuEditVerify.Visible = zlStr.IsHavePrivs(mstrPrivs, "审核")
        mnuEditStrike.Visible = zlStr.IsHavePrivs(mstrPrivs, "冲销")
        
        mnuEditLine1.Visible = mnuEditAdd.Visible Or mnuEditDel.Visible Or mnuEditModify.Visible
        mnuEditPrePareSp.Visible = mnuEditVerify.Visible Or mnuEditStrike.Visible
        
        tlbTool.Buttons("Verify").Visible = mnuEditVerify.Visible
        tlbTool.Buttons("Strike").Visible = mnuEditStrike.Visible
        tlbTool.Buttons("VerifySeparate").Visible = True
    End If

End Sub


Private Sub cboStock_Click()
    On Error Resume Next
    Dim lng库房ID As Long
    Dim rsCheck As New ADODB.Recordset
    Dim str性质 As String
    
    If mblnNoClick Then Exit Sub
    If cboStock.ListIndex >= 0 Then cboStock.Tag = cboStock.ItemData(cboStock.ListIndex)
    
    On Error GoTo ErrHandle
    '检查该库房是否为卫材库，只有卫材库才允许退货
    lng库房ID = cboStock.ItemData(cboStock.ListIndex)
    If mlngMode = 1716 Or mlngMode = 1717 Or mlngMode = 1722 Then
        str性质 = " in ('卫材库')"
    Else
        str性质 = " in ('卫材库','虚拟库房')"
    End If
    
    gstrSQL = " SELECT DISTINCT 0 " & _
              " FROM 部门性质说明 " & _
              " WHERE 工作性质 " & str性质 & _
              "         AND 部门ID =[1]"
              
    Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "检查当前库房是否为卫材库", lng库房ID)
    
    mnuEditRestore.Enabled = (rsCheck.RecordCount <> 0)
'    mnuEditLine0.Enabled = (rsCheck.RecordCount <> 0)
    
    '切换库房刷新菜单项
    SetMenu
    
    If mblnBootUp Then mnuViewRefresh_Click
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
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
    
    If Select部门选择器(Me, cboStock, Trim(cboStock.Text), mstr工作性质, mbln操作员限制) = False Then
        Exit Sub
    End If
    If cboStock.ListIndex >= 0 Then
        cboStock.Tag = cboStock.ItemData(cboStock.ListIndex)
    End If
End Sub

Private Sub cboStock_LostFocus()
    Dim i As Long
        If cboStock.ListCount = 0 Then Exit Sub
    If cboStock.ListIndex < 0 Then
        For i = 0 To cboStock.ListCount - 1
            If Val(cboStock.Tag) = cboStock.ItemData(i) Then
                mblnNoClick = True
                cboStock.ListIndex = i: Exit For
            End If
        Next
    End If
    mblnNoClick = False
End Sub

Private Sub cbrTool_Resize()
    If mblnBootUp = False Then Exit Sub
    Form_Resize
End Sub

Public Sub ShowList(ByVal lngMode As Long, ByVal strTitle As String, ByVal frmMain As Variant)
    Dim strFind As String
    
    mblnBootUp = False
    mlngMode = lngMode
    mstrTitle = strTitle
    mstrPrivs = ";" & gstrPrivs & ";"
    
    
    mbln向发料部门领用 = False
    If mlngMode = 1717 Then
        '刘兴宏:增加可以向发料部门领用
        '问题:8468
        mbln向发料部门领用 = Val(zlDatabase.GetPara(132, glngSys, 0)) = 1
    End If
        
        
    If Not CheckDepend Then Exit Sub            '数据依赖性测试
    Me.Caption = strTitle
    Me.Tag = strTitle
                
    SetPopedom  '根据权限设置不同的显示项目
    mintFindDay = Val(zlDatabase.GetPara("查询天数", glngSys, mlngMode, 1))
    mdtStartDate = Format(DateAdd("d", -mintFindDay, sys.Currentdate), "yyyy-MM-dd")
    mdtEndDate = Format(sys.Currentdate, "yyyy-MM-dd")
    mdtVerifyStart = "1901-01-01"
    mdtVerifyEnd = "1901-01-01"
    
    
    strFind = " AND A.记录状态 = 1 And A.审核日期 is Null And A.填制日期 Between To_Date('" & Format(mdtStartDate, "yyyy-mm-dd") & " 00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(mdtEndDate, "yyyy-mm-dd") & " 23:59:59','YYYY-MM-DD HH24:MI:SS')"
    
    mstrFind = strFind
    
   
    Call tabShow_Click(0)
    If mlngMode <> 1716 Then GetList (mstrFind) '列出单据头
    
    
    
    RestoreWinState Me, App.ProductName, mstrTitle
    Call SetColCostPriceWidth
    mblnBootUp = True
    
    Call SetTlbAndMenuCaption
    Call SetMenu
    
    '2006-04-25:刘兴宏,统一增加报表发布到模块的功能
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
     
    If IsObject(frmMain) Then
        Me.Show , frmMain
    Else
        OS.ShowChildWindow Me.hwnd, frmMain
    End If
    
    Me.ZOrder 0
End Sub

Private Sub SetColCostPriceWidth()
    Dim intCol As Integer
    Dim blnCol成本 As Boolean
    
    With mshList
        For intCol = 1 To .Cols - 1
            If .TextMatrix(0, intCol) = "库存差价" Or .TextMatrix(0, intCol) = "差价金额" Or .TextMatrix(0, intCol) = "成本金额" Or .TextMatrix(0, intCol) = "结算金额" Then
                .ColWidth(intCol) = IIf(mblnCostView = True, 1000, 0)
            End If
        Next
    End With
    With mshDetail
        Select Case mlngMode
            Case 1712 '卫材外购
                .ColWidth(.ColIndex("结算价")) = IIf(mblnCostView = True, 1000, 0)
                .ColWidth(.ColIndex("结算金额")) = IIf(mblnCostView = True, 1500, 0)
                .ColWidth(.ColIndex("差价")) = IIf(mblnCostView = True, 1000, 0)
            Case 1713 '自制入库
                .ColWidth(.ColIndex("采购价")) = IIf(mblnCostView = True, 1000, 0)
                .ColWidth(.ColIndex("采购金额")) = IIf(mblnCostView = True, 1500, 0)
                .ColWidth(.ColIndex("差价")) = IIf(mblnCostView = True, 1000, 0)
            Case 1714 '其他入库
                .ColWidth(.ColIndex("成本价")) = IIf(mblnCostView = True, 1000, 0)
                .ColWidth(.ColIndex("成本金额")) = IIf(mblnCostView = True, 1500, 0)
                .ColWidth(.ColIndex("差价")) = IIf(mblnCostView = True, 1000, 0)
            Case 1716 '卫材移库
                .ColWidth(.ColIndex("成本价")) = IIf(mblnCostView = True, 1000, 0)
                .ColWidth(.ColIndex("成本金额")) = IIf(mblnCostView = True, 1500, 0)
                .ColWidth(.ColIndex("差价")) = IIf(mblnCostView = True, 1000, 0)
            Case 1717 '卫材领用
                .ColWidth(.ColIndex("成本价")) = IIf(mblnCostView = True, 1000, 0)
                .ColWidth(.ColIndex("成本金额")) = IIf(mblnCostView = True, 1500, 0)
                .ColWidth(.ColIndex("差价")) = IIf(mblnCostView = True, 1000, 0)
            Case 1718 '其他出库
                .ColWidth(.ColIndex("成本价")) = IIf(mblnCostView = True, 1000, 0)
                .ColWidth(.ColIndex("成本金额")) = IIf(mblnCostView = True, 1500, 0)
                .ColWidth(.ColIndex("差价")) = IIf(mblnCostView = True, 1000, 0)
        End Select
    End With
End Sub
Private Sub SetTlbAndMenuCaption()
    '设置菜单和工具栏的相关属性
    
    If mlngMode = 1716 Then
        mnuEditVerify.Caption = "接收(&C)"
        tlbTool.Buttons("Verify").Caption = "接收"
        tlbTool.Buttons("Verify").Tag = "接收"
        tlbTool.Buttons("Verify").ToolTipText = "接收"
        TabShow.Visible = True
    Else
        TabShow.Visible = False
    End If
    

End Sub


'检查数据依赖性
Private Function CheckDepend() As Boolean
    Dim rsTemp As New Recordset
    
    On Error GoTo ErrHandle
    CheckDepend = False
    mbln只具备普通科室 = False
 
    '获取可操作的库房
    Select Case mlngMode
        Case 1712                       '卫材外购入库管理
            mstr工作性质 = "V,K,12"
        Case 1713                       '卫材自制入库管理
            mstr工作性质 = "V,K,12"
        Case 1714                       '卫材其他入库管理
            mstr工作性质 = "V,K,W,12"
        Case 1715                       '卫材库存差价调整
            mstr工作性质 = "V,K,12"
        Case 1716                       '卫材移库管理
            '对移库单,可以跟发料部门移库.只用移到发料部门的材料才能被销售.
            mstr工作性质 = "V,K,W"
        Case 1717                       '卫材领用管理
            '问题:8468:20060803,主要是修改可以向发料部门领用
            mstr工作性质 = "V,K" & IIf(mbln向发料部门领用 = False, "", ",W")
            
            gstrSQL = "" & _
                "   SELECT /*+ Rule*/ DISTINCT a.id, a.名称 " & _
                "   FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " & _
                "     , Table(Cast(f_Str2List([3]) as zlTools.t_StrList)) D " & _
                "   Where c.工作性质 = b.名称 and (a.站点=[2] or a.站点 is null) and b.编码 = D.Column_Value " & _
                "           AND a.id = c.部门id " & _
                "           AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01'" & _
                "           And a.ID IN (Select 部门ID From 部门人员 Where 人员ID=[1])"
            '"         AND instr(',V,K,W',b.编码)>0"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取人员库房性质", UserInfo.Id, gstrNodeNo, mstr工作性质)
            If rsTemp.EOF Then
                mbln只具备普通科室 = True
                '由于是普通科室，因此具备向所有库房申领,但是不具备审核;冲销等权限
                If InStr(mstrPrivs, ";所有库房;") = 0 Then
                    mstrPrivs = mstrPrivs & ";所有库房;"
                End If
                mstrPrivs = Replace(mstrPrivs, ";审核;", ";")
                mstrPrivs = Replace(mstrPrivs, ";冲销;", ";")
            Else
                mbln只具备普通科室 = False
            End If
        Case 1718                       '卫材其他出库管理
            mstr工作性质 = "W,V,K,12"
        Case 1719                       '卫材盘点管理
            mstr工作性质 = "V,K,12"
        Case Else
    End Select
    
    gstrSQL = "" & _
        "   SELECT /*+ Rule*/ DISTINCT a.id,a.编码, a.名称,a.简码" & _
        "   FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " & _
        "     , Table(Cast(f_Str2List([3]) as zlTools.t_StrList)) D " & _
        "   Where c.工作性质 = b.名称 and (a.站点=[2] or a.站点 is null) " & _
        "           AND b.编码 = D.Column_Value " & _
        "           AND a.id = c.部门id " & _
        "           AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01'" & _
        IIf(InStr(mstrPrivs, "所有库房") <> 0, "", " And a.ID IN (Select 部门ID From 部门人员 Where 人员ID=[1])")
    
    mbln操作员限制 = Not zlStr.IsHavePrivs(mstrPrivs, "所有库房")
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取相应的库房", UserInfo.Id, gstrNodeNo, mstr工作性质)
    
    If rsTemp.EOF Then
        If mlngMode = 1717 And mbln向发料部门领用 = False Then
            ShowMsgBox "至少应该设置一个具有卫材库性质" & vbCrLf & "或者制剂室性质的部门,请查看部门管理！"
        Else
            ShowMsgBox "至少应该设置一个具有卫材库性质，发料部门性质" & vbCrLf & "或者制剂室性质的部门,请查看部门管理！"
        End If
        rsTemp.Close
        Exit Function
    End If
    
    
    '装入库房数据
    With cboStock
        .Clear
        Do While Not rsTemp.EOF
            .AddItem rsTemp!名称
            .ItemData(.NewIndex) = rsTemp!Id
            mStr库房 = mStr库房 & rsTemp!Id & "," & rsTemp!名称 & "|"
            If rsTemp!Id = UserInfo.部门ID Then
                .ListIndex = .NewIndex
            End If
            rsTemp.MoveNext
        Loop
        rsTemp.Close
        If .ListIndex = -1 Then
            .ListIndex = 0
        End If
    End With
    mbln需要核查 = False
    If mlngMode = 1712 Then
        '外购入库，需要确定是否需要核查功能
        mbln需要核查 = Val(zlDatabase.GetPara("卫材外购需要核查", glngSys, "0")) = 1
    End If
    
    mint领用审核方式 = 0
    If mlngMode = 1717 Then
        mint领用审核方式 = Val(zlDatabase.GetPara("审核流程", glngSys, mlngMode, "0"))
    End If
    
    CheckDepend = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetList(ByVal strFind As String)
    Dim rsTemp As New Recordset
    Dim strUserPart As String
    Dim dbl1 As Double, dbl2 As Double, dbl3 As Double, dbl4 As Double
    Dim strTemp As String
    mlastRow = 0
    
    On Error GoTo ErrHandle
    Call FS.ShowFlash("正在搜索卫生材料记录,请稍候 ...", Me)
    DoEvents
    Screen.MousePointer = vbHourglass
    strUserPart = " And A.库房ID+0=[1]"
    mshList.Redraw = False
    
    Select Case mlngMode
        Case 1712           '卫材外购入库管理
            gstrSQL = "" & _
                "   SELECT  A.No, Decode(Nvl(A.发药方式, 0), 0, '入库单', '退库单') as 单据说明,C.名称 AS 供应商,ltrim(to_char(SUM(A.成本金额)," & mOraFMT.FM_金额 & ")) AS 结算金额," & _
                "           ltrim(to_char((SUM(A.零售金额))," & mOraFMT.FM_金额 & ")) AS 售价金额," & _
                "           LTRIM(TO_CHAR((SUM(A.零售金额 - A.成本金额))," & mOraFMT.FM_金额 & " )) AS 差价金额," & _
                "           Decode(Sign(Nvl(Max(a.费用id), 0) - 1), 1, 0, Nvl(Max(a.费用id), 0)) as 财务标志, A.填制人,TO_CHAR(min(A.填制日期), 'yyyy-mm-dd HH24:Mi:SS') AS 填制日期, " & _
                            IIf(mbln需要核查 = False, "", "           A.配药人 as 核查人,TO_CHAR(min(A.配药日期), 'yyyy-mm-dd HH24:Mi:SS') AS 核查日期, ") & _
                "           A.审核人,TO_CHAR(min(A.审核日期), 'yyyy-mm-dd HH24:Mi:SS') AS 审核日期," & _
                "           A.记录状态, " & _
                "           A.摘要, Max(费用id) As 费用id " & _
                "   FROM 药品收发记录 A, 部门表 B, 供应商 C,材料特性 D, 应付记录 E " & _
                "   Where A.库房id = B.ID AND A.供药单位id = C.Id and (c.站点=[21] or c.站点 is null) AND A.单据 = 15 and a.药品id=d.材料id(+) and e.系统标识(+) = 5 And e.记录性质(+) = 0 And a.Id = e.收发id(+) " & mstr高值耗材 & _
                        strUserPart & strFind & _
                "   GROUP BY A.No,C.名称,A.填制人,A.审核人,A.配药人 ,A.记录状态,A.摘要,A.发药方式 " & _
                "   ORDER BY   No Desc,填制日期 asc"
                
        Case 1713           '卫材自制入库管理
             
            gstrSQL = "" & _
                "   SELECT  a.no, c.名称 AS 制剂室,ltrim(TO_CHAR (SUM (nvl(a.成本金额,0))," & mOraFMT.FM_金额 & ")) AS 成本金额," & _
                "           ltrim(TO_CHAR ( (SUM (nvl(a.零售金额,0))), " & mOraFMT.FM_金额 & ")) AS 售价金额," & _
                "           LTRIM(TO_CHAR((SUM(A.零售金额 - nvl(A.成本金额,0)))," & mOraFMT.FM_金额 & " )) AS 差价金额," & _
                "           a.填制人, " & _
                "           TO_CHAR (min(a.填制日期), 'yyyy-mm-dd HH24:Mi:SS') AS 填制日期, a.审核人, " & _
                "           TO_CHAR (min(a.审核日期), 'yyyy-mm-dd HH24:Mi:SS') AS 审核日期, a.记录状态, a.摘要 " & _
                "   FROM 药品收发记录 a, 部门表 b ,部门表 c " & _
                "   Where   a.库房id = b.ID AND a.对方部门id=c.id AND a.单据 = 16 and a.入出系数=1 " & _
                            strUserPart & strFind & _
                "   GROUP BY a.no,c.名称,a.填制人,a.审核人,a.记录状态,a.摘要 " & _
                "   ORDER BY no DESC, 填制日期 ASC "
    
        Case 1714           '卫材其他入库管理
            gstrSQL = "" & _
                "   SELECT  a.no, c.名称 AS 入出类别,ltrim(TO_CHAR (SUM (a.成本金额)," & mOraFMT.FM_金额 & " )) AS 成本金额," & _
                "           ltrim(TO_CHAR ((SUM (a.零售金额)), " & mOraFMT.FM_金额 & " )) AS 售价金额," & _
                "           LTRIM(TO_CHAR((SUM(A.零售金额 - A.成本金额))," & mOraFMT.FM_金额 & " )) AS 差价金额," & _
                "           a.填制人,TO_CHAR (min(a.填制日期), 'yyyy-mm-dd HH24:Mi:SS') AS 填制日期, a.审核人," & _
                "           TO_CHAR (min(a.审核日期), 'yyyy-mm-dd HH24:Mi:SS') AS 审核日期, a.记录状态, a.摘要 " & _
                "   FROM 药品收发记录 a, 部门表 b,药品入出类别 c " & _
                "   Where a.库房id = b.ID AND a.入出类别id = c.id AND a.单据 = 17 " & _
                            strUserPart & strFind & _
                "   GROUP BY a.no,c.名称,a.填制人,a.审核人,a.记录状态,a.摘要 " & _
                "   ORDER BY no DESC,填制日期 ASC "
                
        Case 1715           '库存差价调整管理
            gstrSQL = "" & _
                "   SELECT  a.no, ltrim(TO_CHAR (SUM (a.零售价), " & mOraFMT.FM_金额 & " )) AS 库存金额," & _
                "           ltrim(TO_CHAR (SUM (a.成本价), " & mOraFMT.FM_金额 & " )) AS 库存差价," & _
                "           ltrim(TO_CHAR ( (SUM (a.差价))," & mOraFMT.FM_金额 & " )) AS 调整额, " & _
                "           a.填制人,TO_CHAR (min(a.填制日期), 'yyyy-mm-dd HH24:Mi:SS') AS 填制日期, a.审核人," & _
                "           TO_CHAR(min(a.审核日期), 'yyyy-mm-dd HH24:Mi:SS') AS 审核日期, a.记录状态, a.摘要,max(nvl(a.发药方式,0)) as 发药方式  " & _
                "   FROM 药品收发记录 a, 部门表 b " & _
                "   Where   a.库房id = b.ID  AND a.单据 = 18 " & _
                            strUserPart & strFind & _
                "   GROUP BY a.no,a.填制人,a.审核人,a.记录状态,a.摘要 " & _
                "   ORDER BY no DESC,填制日期 ASC "
            
        Case 1716           '卫材移库管理
            If mbln申领核查 = True Then
                strTemp = " and (Nvl(a.发药方式, 0) = 1 And a.核查人 Is Not Null Or Nvl(a.发药方式, 0) = 0)"
            Else
                strTemp = ""
            End If
                If TabShow.Tab = 0 Then
                    strUserPart = " And A.库房ID+0=[1]"
                    
                    gstrSQL = "" & _
                        "   SELECT  a.no, c.名称 AS 移入库房," & _
                        "           LTRIM(TO_CHAR (SUM (A.成本金额), " & mOraFMT.FM_金额 & ")) AS 成本金额, " & _
                        "           ltrim(TO_CHAR ((SUM (a.零售金额))," & mOraFMT.FM_金额 & ")) AS 售价金额," & _
                        "           LTRIM(TO_CHAR((SUM(A.零售金额 - A.成本金额))," & mOraFMT.FM_金额 & " )) AS 差价金额," & _
                        "           a.填制人, " & _
                        "           TO_CHAR (min(a.填制日期), 'yyyy-mm-dd HH24:Mi:SS') AS 填制日期,a.审核人 as 接收人,  " & _
                        "           To_Char(Min(a.审核日期), 'yyyy-mm-dd HH24:Mi:SS') As 接收日期," & _
                        "           A.配药人 AS 备料人,TO_CHAR(MIN(A.配药日期),'YYYY-MM-DD HH24:MI:SS') AS 发送日期," & _
                        "            a.记录状态, a.摘要 " & _
                        "   FROM 药品收发记录 a, 部门表 b ,部门表 c " & _
                        "   Where   a.库房id = b.ID AND a.对方部门id=c.id AND a.单据 = 19 AND  a.入出系数=-1 " & _
                                    strUserPart & strFind & strTemp & _
                        "   GROUP BY a.no,c.名称,a.填制人,a.审核人,a.配药人,a.记录状态,a.摘要 " & _
                        "   ORDER BY a.no DESC,填制日期 ASC,a.配药人 asc "
                Else
                    strUserPart = " And A.对方部门ID+0=[1]"
                    gstrSQL = "" & _
                        "   SELECT  a.no, B.名称 AS 移出库房," & _
                        "           LTRIM(TO_CHAR (SUM (A.成本金额), " & mOraFMT.FM_金额 & ")) AS 成本金额, " & _
                        "           ltrim(TO_CHAR ((SUM (a.零售金额))," & mOraFMT.FM_金额 & ")) AS 售价金额," & _
                        "           LTRIM(TO_CHAR((SUM(A.零售金额 - A.成本金额))," & mOraFMT.FM_金额 & " )) AS 差价金额," & _
                        "           a.填制人, " & _
                        "           TO_CHAR (min(a.填制日期), 'yyyy-mm-dd HH24:Mi:SS') AS 填制日期," & _
                        "           A.审核人 AS 接收人,TO_CHAR(MIN(A.审核日期),'YYYY-MM-DD HH24:MI:SS') AS 审核日期," & _
                        "           a.记录状态, a.摘要 " & _
                        "   FROM 药品收发记录 a, 部门表 b ,部门表 c " & _
                        "   Where   a.库房id = b.ID AND a.对方部门id=c.id AND a.单据 = 19 AND  a.入出系数=-1 " & _
                                    strUserPart & strFind & strTemp & _
                        "   GROUP BY a.no,b.名称,a.填制人,a.审核人,a.记录状态,a.摘要 " & _
                        "   ORDER BY a.no DESC,填制日期 ASC,a.审核人 asc "
                End If
        Case 1717           '卫材领用管理
            gstrSQL = "" & _
                "   SELECT  a.no, c.名称 AS 领用部门," & _
                "           LTRIM(TO_CHAR (SUM (A.成本金额), " & mOraFMT.FM_金额 & ")) AS 成本金额, " & _
                "           ltrim(TO_CHAR ((SUM (a.零售金额))," & mOraFMT.FM_金额 & ")) AS 售价金额," & _
                "           LTRIM(TO_CHAR((SUM(A.零售金额 - A.成本金额))," & mOraFMT.FM_金额 & " )) AS 差价金额," & _
                "           a.领用人,a.填制人, " & _
                "           TO_CHAR (min(a.填制日期), 'yyyy-mm-dd HH24:Mi:SS') AS 填制日期 "
                
                If mint领用审核方式 = 1 Then
                    gstrSQL = gstrSQL & ", a.配药人 As 核查人, TO_CHAR (min(a.配药日期), 'yyyy-mm-dd HH24:Mi:SS') AS 核查日期 "
                End If
                
            gstrSQL = gstrSQL & ", a.审核人, TO_CHAR (min(a.审核日期), 'yyyy-mm-dd HH24:Mi:SS') AS 审核日期, a.记录状态, a.摘要 " & _
                "   FROM 药品收发记录 a, 部门表 b ,部门表 c " & _
                "   Where   a.库房id = b.ID AND a.对方部门id=c.id AND a.单据 = 20 " & IIf(mbln只具备普通科室, " and a.对方部门id in (Select 部门ID From 部门人员 Where 人员ID=[20])", "") & _
                            strUserPart & strFind & _
                "   GROUP BY a.no,c.名称,a.领用人,a.填制人,a.配药人,a.审核人,a.记录状态,a.摘要 " & _
                "   ORDER BY no DESC, 填制日期 ASC "
                
        Case 1718          '卫材其他出库管理
            gstrSQL = "" & _
                "   SELECT /*+rule*/ a.no, c.名称 AS 入出类别,d.名称 AS 对方单位," & _
                "           LTRIM(TO_CHAR (SUM (A.成本金额), " & mOraFMT.FM_金额 & ")) AS 成本金额, " & _
                "           ltrim(TO_CHAR ((SUM (a.零售金额))," & mOraFMT.FM_金额 & ")) AS 售价金额," & _
                "           LTRIM(TO_CHAR((SUM(A.零售金额 - A.成本金额))," & mOraFMT.FM_金额 & " )) AS 差价金额," & _
                "           LTrim(To_Char((Sum(A.单量 * A.实际数量)), '9999999999990.99')) As 外销金额," & _
                "           a.填制人,TO_CHAR (min(a.填制日期), 'yyyy-mm-dd HH24:Mi:SS') AS 填制日期, a.审核人," & _
                "           TO_CHAR (min(a.审核日期), 'yyyy-mm-dd HH24:Mi:SS') AS 审核日期, a.记录状态, a.摘要, Max(费用id) As 费用id " & _
                "   FROM 药品收发记录 a, 部门表 b,药品入出类别 c,材料外销单位 d " & _
                "   Where a.库房id = b.ID AND a.入出类别id = c.id AND A.发药窗口=D.编码 And a.单据 = 21 " & strUserPart & strFind & _
                "   GROUP BY a.no,c.名称,d.名称,a.填制人,a.审核人,a.记录状态,a.摘要 "

            gstrSQL = gstrSQL & _
                " Union All " & _
                "   SELECT  a.no, c.名称 AS 入出类别,'' AS 对方单位," & _
                "           LTRIM(TO_CHAR (SUM (A.成本金额), " & mOraFMT.FM_金额 & ")) AS 成本金额, " & _
                "           ltrim(TO_CHAR ((SUM (a.零售金额))," & mOraFMT.FM_金额 & ")) AS 售价金额," & _
                "           LTRIM(TO_CHAR((SUM(A.零售金额 - A.成本金额))," & mOraFMT.FM_金额 & " )) AS 差价金额," & _
                "           LTrim(To_Char((Sum(A.单量 * A.实际数量)), '9999999999990.99')) As 外销金额," & _
                "           a.填制人,TO_CHAR (min(a.填制日期), 'yyyy-mm-dd HH24:Mi:SS') AS 填制日期, a.审核人," & _
                "           TO_CHAR (min(a.审核日期), 'yyyy-mm-dd HH24:Mi:SS') AS 审核日期, a.记录状态, a.摘要, Max(费用id) As 费用id " & _
                "   FROM 药品收发记录 a, 部门表 b,药品入出类别 c " & _
                "   Where a.库房id = b.ID AND a.入出类别id = c.id AND A.发药窗口 Is Not Null And A.发药窗口 Not In (Select 编码 From 材料外销单位) And a.单据 = 21 " & strUserPart & strFind & _
                "   GROUP BY a.no,c.名称,a.填制人,a.审核人,a.记录状态,a.摘要 "
            
            gstrSQL = gstrSQL & _
                " Union All " & _
                "   SELECT  a.no, c.名称 AS 入出类别,'' AS 对方单位," & _
                "           LTRIM(TO_CHAR (SUM (A.成本金额), " & mOraFMT.FM_金额 & ")) AS 成本金额, " & _
                "           ltrim(TO_CHAR ((SUM (a.零售金额))," & mOraFMT.FM_金额 & ")) AS 售价金额," & _
                "           LTRIM(TO_CHAR((SUM(A.零售金额 - A.成本金额))," & mOraFMT.FM_金额 & " )) AS 差价金额," & _
                "           LTrim(To_Char((Sum(A.单量 * A.实际数量)), '9999999999990.99')) As 外销金额," & _
                "           a.填制人,TO_CHAR (min(a.填制日期), 'yyyy-mm-dd HH24:Mi:SS') AS 填制日期, a.审核人," & _
                "           TO_CHAR (min(a.审核日期), 'yyyy-mm-dd HH24:Mi:SS') AS 审核日期, a.记录状态, a.摘要, Max(费用id) As 费用id " & _
                "   FROM 药品收发记录 a, 部门表 b,药品入出类别 c " & _
                "   Where a.库房id = b.ID AND a.入出类别id = c.id And A.发药窗口 Is Null And a.单据 = 21 " & strUserPart & strFind & _
                "   GROUP BY a.no,c.名称,a.填制人,a.审核人,a.记录状态,a.摘要 " & _
                "   ORDER BY no DESC,填制日期 ASC "
        Case 1719         '卫材盘点
            '频次字段保存的 盘店时间
            gstrSQL = "" & _
                "   SELECT distinct a.no, 频次 AS 盘点时间," & _
                "           a.填制人,TO_CHAR (min(a.填制日期), 'yyyy-mm-dd HH24:Mi:SS') AS 填制日期, a.审核人," & _
                "           TO_CHAR (min(a.审核日期), 'yyyy-mm-dd HH24:Mi:SS') AS 审核日期, " & _
                "           ltrim(to_char((Sum(Nvl(扣率,0)*零售价))," & mOraFMT.FM_金额 & ")) 盘点金额," & _
                "           ltrim(to_char((Sum(零售金额))," & mOraFMT.FM_金额 & ")) 金额差,a.记录状态, a.摘要 " & _
                "   FROM 药品收发记录 a, 部门表 b " & _
                "   Where a.库房id = b.ID AND a.单据 =22  " & strUserPart & strFind & _
                "   Group by a.no,频次,a.填制人,a.审核人,a.记录状态, a.摘要 " & _
                "   ORDER BY no DESC,填制日期 ASC "
    End Select
    
     'mstrOthers(0 To 13) As String ' 0-记录状态(计划类型),1-开始单号,2-结束单号,3-材料id,4-对方部门id(或入出类别id或编制方法(计划用)),5-填制人,6-审核人,7-供应商ID,8-生产商,9-开始生产日期,10-结束生产日期,11-开始发票号,12-结束发票号,13-条码信息
    '参数范围:[1]-库房id,[2]:开始填制日期,[3]结束填制日期,[4]开始审核日期,[5] 结束审核日期,[6]-记录状态,[7]开始单据号,[8]结束单据号,[9]材料id,[10]对方部门id,[11]填制人,[12]审核人[13]-供应商ID,[14]-生产商,[15]-开始生产日期,[16]-结束生产日期,[17]-开始发票号,[18]-结束发票号,[19]-条码信息
    
    '初始生产日期
    mstrOthers(9) = IIf(Trim(mstrOthers(9)) = "", "1901-01-01", mstrOthers(9))
    mstrOthers(10) = IIf(Trim(mstrOthers(10)) = "", "1901-01-01", mstrOthers(10))
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
        cboStock.ItemData(cboStock.ListIndex), _
        CDate(Format(mdtStartDate, "yyyy-mm-dd") & " 00:00:00"), _
        CDate(Format(mdtEndDate, "yyyy-mm-dd") & " 23:59:59"), _
        CDate(Format(mdtVerifyStart, "yyyy-mm-dd") & " 00:00:00"), _
        CDate(Format(mdtVerifyEnd, "yyyy-mm-dd") & " 23:59:59"), _
        Val(mstrOthers(0)), _
        mstrOthers(1), _
        mstrOthers(2), _
        Val(mstrOthers(3)), _
        Val(mstrOthers(4)), _
        mstrOthers(5), _
        mstrOthers(6), _
        Val(mstrOthers(7)), _
        mstrOthers(8), _
        CDate(mstrOthers(9) & " 00:00:00"), _
        CDate(mstrOthers(10) & " 23:59:59"), _
        mstrOthers(11), _
        mstrOthers(12), _
        mstrOthers(13) & "%", _
        UserInfo.Id, _
        gstrNodeNo)
        
    Set mshList.DataSource = rsTemp
    
    With mshList
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
    
    Call SetListColWidth
    
    '统计合计金额
    If (Not rsTemp.EOF) And (Not rsTemp.BOF) Then
        rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            Select Case mlngMode
                Case 1712
                    dbl1 = dbl1 + IIf(IsNull(rsTemp!结算金额), 0, rsTemp!结算金额)
                    dbl2 = dbl2 + IIf(IsNull(rsTemp!售价金额), 0, rsTemp!售价金额)
                    dbl3 = dbl3 + IIf(IsNull(rsTemp!差价金额), 0, rsTemp!差价金额)
                Case 1713, 1714, 1716, 1717
                    dbl1 = dbl1 + IIf(IsNull(rsTemp!成本金额), 0, rsTemp!成本金额)
                    dbl2 = dbl2 + IIf(IsNull(rsTemp!售价金额), 0, rsTemp!售价金额)
                    dbl3 = dbl3 + IIf(IsNull(rsTemp!差价金额), 0, rsTemp!差价金额)
                Case 1715
                    dbl1 = dbl1 + IIf(IsNull(rsTemp!库存金额), 0, rsTemp!库存金额)
                    dbl2 = dbl2 + IIf(IsNull(rsTemp!库存差价), 0, rsTemp!库存差价)
                    dbl3 = dbl3 + IIf(IsNull(rsTemp!调整额), 0, rsTemp!调整额)
                Case 1718
                    dbl1 = dbl1 + IIf(IsNull(rsTemp!成本金额), 0, rsTemp!成本金额)
                    dbl2 = dbl2 + IIf(IsNull(rsTemp!售价金额), 0, rsTemp!售价金额)
                    dbl3 = dbl3 + IIf(IsNull(rsTemp!差价金额), 0, rsTemp!差价金额)
                    dbl4 = dbl4 + IIf(IsNull(rsTemp!外销金额), 0, rsTemp!外销金额)
                Case 1719
                    dbl2 = dbl2 + IIf(IsNull(rsTemp!盘点金额), 0, rsTemp!盘点金额)
                    dbl3 = dbl3 + IIf(IsNull(rsTemp!金额差), 0, rsTemp!金额差)
            End Select
            rsTemp.MoveNext
        Loop
        rsTemp.MoveFirst
    End If
    Dim strText As String
    Select Case mlngMode
        Case 1712
            If mblnCostView = False Then
                strText = "售价金额合计：" & Format(dbl2, mFMT.FM_金额)
            Else
                strText = "结算金额合计：" & Format(dbl1, mFMT.FM_金额)
                strText = strText & Space(10) & " 售价金额合计：" & Format(dbl2, mFMT.FM_金额)
                strText = strText & Space(10) & "差价金额合计：" & Format(dbl3, mFMT.FM_金额)
            End If
        Case 1713, 1714, 1716, 1717
            If mblnCostView = False Then
                strText = "售价金额合计：" & Format(dbl2, mFMT.FM_金额)
            Else
                strText = "成本金额合计：" & Format(dbl1, mFMT.FM_金额)
                strText = strText & Space(10) & "售价金额合计：" & Format(dbl2, mFMT.FM_金额)
                strText = strText & Space(10) & "差价金额合计：" & Format(dbl3, mFMT.FM_金额)
            End If
        Case 1715
            strText = "库存金额合计：" & Format(dbl1, mFMT.FM_金额)
            strText = strText & Space(10) & "库存差价合计：" & Format(dbl2, mFMT.FM_金额)
            strText = strText & Space(10) & "调整额合计：" & Format(dbl3, mFMT.FM_金额)
        Case 1718
            If mblnCostView = False Then
                strText = "售价金额合计：" & Format(dbl2, mFMT.FM_金额)
                strText = strText & Space(10) & "外销金额合计：" & Format(dbl4, mFMT.FM_金额)
            Else
                strText = "成本金额合计：" & Format(dbl1, mFMT.FM_金额)
                strText = strText & Space(10) & "售价金额合计：" & Format(dbl2, mFMT.FM_金额)
                strText = strText & Space(10) & "差价金额合计：" & Format(dbl3, mFMT.FM_金额)
                strText = strText & Space(10) & "外销金额合计：" & Format(dbl4, mFMT.FM_金额)
            End If
        Case 1719
            strText = "盘点金额合计：" & Format(dbl2, mFMT.FM_金额)
            strText = strText & Space(10) & "金额差合计：" & Format(dbl3, mFMT.FM_金额)
    End Select
    mstrMoneySum = strText
    PrintRange strText & Space(10) & vbCrLf & mstrPrintRange
    
    
    Call mshlist_EnterCell    '列出单据体
    
    Call SetStrikeColor
    
    With mshList
        .Row = 1
        .Col = 0
        .ColSel = .Cols - 1
    End With
    mshList.Redraw = True
    
    Call FS.StopFlash
    
    Screen.MousePointer = vbDefault
    stbThis.Panels(2).Text = "当前共有" & rsTemp.RecordCount & "张单据"
    
    rsTemp.Close
    If mshList.Visible = True Then
        mshList.SetFocus
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetStrikeColor()
    Dim int记录状态 As Integer, int财务标志 As Integer
    Dim intRow As Integer, intCol As Integer
    Dim intCol记录状态 As Integer, intCol财务标志 As Integer
    Dim int自动审核 As Integer
    Dim intCol审核人 As Integer
        
    With mshList
        If .Rows <= 2 Then Exit Sub
        intCol记录状态 = GetCol(mshList, "记录状态")
        If intCol记录状态 < 0 Then Exit Sub
        intCol财务标志 = GetCol(mshList, "财务标志")
        int自动审核 = GetCol(mshList, "费用ID")
        If mlngMode = 1716 Then '卫材移库
            intCol审核人 = GetCol(mshList, "接收人")
        Else
            intCol审核人 = GetCol(mshList, "审核人")
        End If
        
        For intRow = 1 To .Rows - 1
            int记录状态 = Val(.TextMatrix(intRow, intCol记录状态))
            If intCol财务标志 >= 0 Then int财务标志 = Val(.TextMatrix(intRow, intCol财务标志))
            
            If int记录状态 Mod 3 = 0 Then
                .Row = intRow
                For intCol = 0 To .Cols - 1
                    .Col = intCol
                    .CellForeColor = &H80000001
                Next
            ElseIf int记录状态 Mod 3 = 2 Then
                .Row = intRow
                For intCol = 0 To .Cols - 1
                    .Col = intCol
                    If .TextMatrix(intRow, intCol审核人) = "" Then
                        .CellForeColor = &HC0C0FF
                    Else
                        .CellForeColor = IIf(int财务标志 = 1, &HC0C0FF, &HFF)   '财务标志显示不一样
                    End If
                Next
            End If
            
            If int自动审核 > 1 Then
                If Val(.TextMatrix(intRow, int自动审核)) > 1 Then
                    .Row = intRow
                    For intCol = 0 To .Cols - 1
                        .Col = intCol
                        .CellForeColor = IIf(Val(.TextMatrix(intRow, int自动审核)) > 1, &H808080, &H80000008)
                    Next
                End If
            End If
        Next
    End With
End Sub

'表头列宽初始
Private Sub SetListColWidth()
    Dim intCol As Integer
    
    With mshList
        Select Case mlngMode
            Case 1712
                .ColAlignment(3) = flexAlignRightCenter
                .ColAlignment(4) = flexAlignRightCenter
                .ColAlignment(5) = flexAlignRightCenter
            Case 1713
                .ColAlignment(2) = flexAlignRightCenter
                .ColAlignment(3) = flexAlignRightCenter
                    
            Case 1714
                .ColAlignment(2) = flexAlignRightCenter
                .ColAlignment(3) = flexAlignRightCenter
            Case 1715
                .ColAlignment(1) = flexAlignRightCenter
                .ColAlignment(2) = flexAlignRightCenter
                .ColAlignment(3) = flexAlignRightCenter
                
            Case 1716
                .ColAlignment(2) = flexAlignRightCenter
            Case 1718
                .ColAlignment(2) = flexAlignLeftCenter
                .ColAlignment(3) = flexAlignRightCenter
                .ColAlignment(4) = flexAlignRightCenter
                .ColAlignment(5) = flexAlignRightCenter
                .ColAlignment(6) = flexAlignRightCenter
            Case 1719
                .ColAlignment(2) = flexAlignRightCenter         '售价金额
            Case 1720
                
            Case Else
            
        End Select
        intCol = GetCol(mshList, "记录状态")
        If intCol >= 0 Then mshList.ColWidth(intCol) = 0
        intCol = GetCol(mshList, "财务标志")
        If intCol >= 0 Then .ColWidth(intCol) = 0
        intCol = GetCol(mshList, "费用ID")
        If intCol >= 0 Then .ColWidth(intCol) = 0
        
        If mblnBootUp = False Then
            For intCol = 1 To .Cols - 1
                If intCol = 1 Then
                    If mlngMode = 1715 Then
                        .ColWidth(intCol) = 1000
                    ElseIf intCol = GetCol(mshList, "单据类型") Then
                        .ColWidth(intCol) = 900
                    Else
                        .ColWidth(intCol) = 2000
                    End If
                    
                ElseIf intCol = GetCol(mshList, "记录状态") Then
                    .ColWidth(intCol) = 0
                ElseIf intCol = GetCol(mshList, "财务标志") Then
                    .ColWidth(intCol) = 0
                ElseIf intCol = GetCol(mshList, "供应商") Then
                     .ColWidth(intCol) = 2000
                Else
                    .ColWidth(intCol) = 1000
                End If
                If mlngMode = 1715 Then
                    If intCol = GetCol(mshList, "发药方式") Then
                        .ColWidth(intCol) = 0
                    End If
                End If
                
                If .TextMatrix(0, intCol) = "库存差价" Or .TextMatrix(0, intCol) = "差价金额" Or .TextMatrix(0, intCol) = "结算金额" Or .TextMatrix(0, intCol) = "成本金额" Then
                    .ColWidth(intCol) = IIf(mblnCostView = True, 1000, 0)
                End If
            Next
        End If
    End With
End Sub


Private Sub SetDetailColWidth()
    Dim intCol As Integer
    Dim blnCol成本 As Boolean
    
    With mshDetail
        For intCol = 0 To .Cols - 1
            .FixedAlignment(intCol) = flexAlignCenterCenter
            '设置主键
            .ColKey(intCol) = .TextMatrix(0, intCol)
        Next
        
        zl_vsGrid_Para_Restore mlngMode, mshDetail, mstrTitle, "mshDetail", False, True
        
        For intCol = 0 To .Cols - 1
'            .FixedAlignment(intCol) = flexAlignCenterCenter
'
'            '设置主键
'            .ColKey(intCol) = .TextMatrix(0, intCol)
            If mblnFirst Then
                If .ColWidth(intCol) = 0 Then .ColWidth(intCol) = 1000
            End If
            Select Case .ColKey(intCol)
            Case "单位", "发票号", "付款序号", "发标日期", "序号", "零售单位", "发票代码"
                .ColAlignment(intCol) = flexAlignCenterCenter
                If .ColKey(intCol) = "序号" Then
                    .ColWidth(intCol) = 0: .ColHidden(intCol) = True
                End If
            Case Else
                If .ColKey(intCol) = "卫材信息" And mblnFirst = True Then
                    If .ColWidth(intCol) = 0 Then .ColWidth(intCol) = 2500
                End If
                .ColAlignment(intCol) = flexAlignLeftCenter
            End Select
            '.coldata(i):1-固定,-1-不能选,0-可选
            If .ColKey(intCol) = "卫材信息" Then .ColData(intCol) = 1
            If .ColKey(intCol) = "序号" Then .ColData(intCol) = -1
            If .ColKey(intCol) Like "*数量*" Or _
                .ColKey(intCol) Like "*价*" Or _
                .ColKey(intCol) Like "*金*" Or _
                .ColKey(intCol) Like "*额*" Or _
                .ColKey(intCol) Like "*率*" Then
                .ColAlignment(intCol) = flexAlignRightCenter
            End If
            '对于零售价\零售金额\零售差价需要默认为零
            Select Case .ColKey(intCol)
            Case "零售价", "零售单位", "零售金额", "零售差价"
                .ColHidden(intCol) = True
            Case Else
            End Select
            
            '条码要根据权限来判断是否显示
            Select Case .ColKey(intCol)
            Case "商品条码", "内部条码"
                If gblnCode = False Then
                    .ColWidth(intCol) = 0
                    .ColHidden(intCol) = True
                End If
            End Select
        Next
        
        For intCol = 1 To mshList.Cols - 1
            If mshList.ColHeaderCaption(0, intCol) = "成本金额" Then
                blnCol成本 = True
                Exit For
            End If
        Next
        
        Select Case mlngMode
            Case 1712 '卫材外购
                .ColWidth(.ColIndex("结算价")) = IIf(mblnCostView = False, 0, 1000)
                If mblnCostView = False Then .ColData(.ColIndex("结算价")) = -1
                .ColWidth(.ColIndex("结算金额")) = IIf(mblnCostView = False, 0, 1500)
                If mblnCostView = False Then .ColData(.ColIndex("结算金额")) = -1
                .ColWidth(.ColIndex("差价")) = IIf(mblnCostView = False, 0, 1500)
                If mblnCostView = False Then .ColData(.ColIndex("差价")) = -1
            Case 1713 '自制入库
                .ColWidth(.ColIndex("采购价")) = IIf(mblnCostView = False, 0, 1000)
                If mblnCostView = False Then .ColData(.ColIndex("采购价")) = -1
                .ColWidth(.ColIndex("采购金额")) = IIf(mblnCostView = False, 0, 1500)
                If mblnCostView = False Then .ColData(.ColIndex("采购金额")) = -1
                .ColWidth(.ColIndex("差价")) = IIf(mblnCostView = False, 0, 1500)
                If mblnCostView = False Then .ColData(.ColIndex("差价")) = -1
            Case 1714 '其他入库
                .ColWidth(.ColIndex("成本价")) = IIf(mblnCostView = False, 0, 1000)
                If mblnCostView = False Then .ColData(.ColIndex("成本价")) = -1
                .ColWidth(.ColIndex("成本金额")) = IIf(mblnCostView = False, 0, 1500)
                If mblnCostView = False Then .ColData(.ColIndex("成本金额")) = -1
                .ColWidth(.ColIndex("差价")) = IIf(mblnCostView = False, 0, 1500)
                If mblnCostView = False Then .ColData(.ColIndex("差价")) = -1
            Case 1716 '卫材移库
                .ColWidth(.ColIndex("成本价")) = IIf(mblnCostView = False, 0, 1000)
                If mblnCostView = False Then .ColData(.ColIndex("成本价")) = -1
                .ColWidth(.ColIndex("成本金额")) = IIf(mblnCostView = False, 0, 1500)
                If mblnCostView = False Then .ColData(.ColIndex("成本金额")) = -1
                .ColWidth(.ColIndex("差价")) = IIf(mblnCostView = False, 0, 1500)
                If mblnCostView = False Then .ColData(.ColIndex("差价")) = -1
                If blnCol成本 = True Then
                    mshList.ColWidth(intCol) = IIf(mblnCostView = False, 0, 1500)
                End If
            Case 1717 '卫材领用
                .ColWidth(.ColIndex("成本价")) = IIf(mblnCostView = False, 0, 1000)
                If mblnCostView = False Then .ColData(.ColIndex("成本价")) = -1
                .ColWidth(.ColIndex("成本金额")) = IIf(mblnCostView = False, 0, 1500)
                If mblnCostView = False Then .ColData(.ColIndex("成本金额")) = -1
                .ColWidth(.ColIndex("差价")) = IIf(mblnCostView = False, 0, 1500)
                If mblnCostView = False Then .ColData(.ColIndex("差价")) = -1
                If blnCol成本 = True Then
                    mshList.ColWidth(intCol) = IIf(mblnCostView = False, 0, 1500)
                End If
            Case 1718 '其他出库
                .ColWidth(.ColIndex("成本价")) = IIf(mblnCostView = False, 0, 1000)
                If mblnCostView = False Then .ColData(.ColIndex("成本价")) = -1
                .ColWidth(.ColIndex("成本金额")) = IIf(mblnCostView = False, 0, 1500)
                If mblnCostView = False Then .ColData(.ColIndex("成本金额")) = -1
                .ColWidth(.ColIndex("差价")) = IIf(mblnCostView = False, 0, 1500)
                If mblnCostView = False Then .ColData(.ColIndex("差价")) = -1
                If blnCol成本 = True Then
                    mshList.ColWidth(intCol) = IIf(mblnCostView = False, 0, 1500)
                End If
        End Select
    End With
End Sub


'根据权限设置不同的显示项目
Private Sub SetPopedom()

    '外购入库所有权限：参数设置、基本、所有库房、登记、修改、删除、验收、冲销、单据打印
    Select Case mlngMode
        Case 1712, 1713, 1714, 1715, 1716, 1717, 1718, 1719
            If mlngMode = 1712 Then
                '刘兴宏:增加核查功能,2007/05/14
                mnuEditCheck.Visible = mbln需要核查 And InStr(1, mstrPrivs, ";核查;") <> 0
                mnuEditCancelCheck.Visible = mnuEditCheck.Visible
                tlbTool.Buttons("Check").Visible = mnuEditCheck.Visible
                tlbTool.Buttons("CancelCheck").Visible = mnuEditCheck.Visible
                tlbTool.Buttons("PrepareSplit").Visible = mnuEditCheck.Visible
            End If
            
            If mlngMode = 1717 Then
                mnuEditCheck.Visible = (mint领用审核方式 = 1) And InStr(1, mstrPrivs, ";财务审核;") <> 0
                mnuEditCancelCheck.Visible = mnuEditCheck.Visible
                mnuEditCheckLine.Visible = mnuEditCheck.Visible
                tlbTool.Buttons("Check").Visible = mnuEditCheck.Visible
                tlbTool.Buttons("CancelCheck").Visible = mnuEditCheck.Visible
                tlbTool.Buttons("PrepareSplit").Visible = mnuEditCheck.Visible
            End If
             
            If InStr(1, mstrPrivs, ";登记;") = 0 Then
                mnuEditAdd.Visible = False
                mnuEditRestore.Visible = False
                tlbTool.Buttons("Add").Visible = False
            Else
                mnuEditRestore.Visible = True
            End If
            
            If InStr(1, mstrPrivs, ";修改;") = 0 Then
                mnuEditModify.Visible = False
                tlbTool.Buttons("Modify").Visible = False
            End If
            
            If InStr(1, mstrPrivs, ";删除;") = 0 Then
                mnuEditDel.Visible = False
                tlbTool.Buttons("Delete").Visible = False
                 '对没有所有编辑权限时，把菜单和工具栏上的相应的分割线屏蔽。
                If mnuEditAdd.Visible = False And mnuEditModify.Visible = False Then
                    mnuEditLine1.Visible = False
                    tlbTool.Buttons("EditSeparate").Visible = False
                End If
            End If
            
            If InStr(1, ";" & mstrPrivs & ";", ";审核;") = 0 Then
                mnuEditVerify.Visible = False
                mnuEditBill.Visible = False
                mnuEditReg.Visible = False
                tlbTool.Buttons("Verify").Visible = False
            End If
            
            If InStr(1, mstrPrivs, ";冲销;") = 0 Then
                mnuEditStrike.Visible = False
                tlbTool.Buttons("Strike").Visible = False
                
                If mnuEditVerify.Visible = False Then
                    mnuEditLine2.Visible = False
                    tlbTool.Buttons("VerifySeparate").Visible = False
                End If
            End If
            If InStr(1, mstrPrivs, ";单据打印;") = 0 Then
                mnuFileBillPrint.Visible = False
                mnuFileBillPreview.Visible = False
            End If
        Case Else
        
    End Select
                        
    If mlngMode = 1712 Then
        mnuEditLine0.Visible = True
        mnuEditCheckBatch.Visible = mbln需要核查 And InStr(1, mstrPrivs, ";核查;") <> 0
        mnuEditVerifyBatch.Visible = InStr(1, mstrPrivs, ";审核;") <> 0
        mnuEditCheckLine.Visible = mnuEditCheck.Visible And (mnuEditVerify.Visible Or mnuEditStrike.Visible)
        If InStr(1, ";" & mstrPrivs & ";", ";审核;") <> 0 Then
            mnuEditBill.Visible = True
            mnuEditReg.Visible = True
        Else
            mnuEditLine0.Visible = False
        End If
        If InStr(1, mstrPrivs, ";财务审核;") <> 0 Then
            mnuEditLine0.Visible = True
            mnuEditAcc.Visible = True
        Else
            If (mnuEditBill.Visible = False And mnuEditReg.Visible = False) Then mnuEditLine0.Visible = False
            mnuEditAcc.Visible = False
        End If
        If InStr(1, mstrPrivs, ";导入计划单;") <> 0 Then
            mnuEditLine0.Visible = True
            mnuEditImport.Visible = True
        Else
            mnuEditImport.Visible = False
        End If
    ElseIf mlngMode = 1716 Then
        '移库单
        mnuEditPrepare.Visible = InStr(1, ";" & mstrPrivs & ";", ";发送;") <> 0
        mnuEditSend.Visible = mnuEditPrepare.Visible
        mnuEditBack.Visible = mnuEditPrepare.Visible
        mnuEditPrePareSp.Visible = mnuEditPrepare.Visible
            
        tlbTool.Buttons("PrepareSplit").Visible = mnuEditPrepare.Visible
        tlbTool.Buttons("Send").Visible = mnuEditPrepare.Visible
        tlbTool.Buttons("Back").Visible = mnuEditPrepare.Visible
        tlbTool.Buttons("Prepare").Visible = mnuEditPrepare.Visible
        
        If InStr(1, ";" & mstrPrivs & ";", ";审核;") = 0 And _
           InStr(1, ";" & mstrPrivs & ";", ";冲销;") = 0 Then
            TabShow.TabVisible(1) = False
        End If
            
    Else
        mnuEditBill.Visible = False
        mnuEditBill.Visible = False
        mnuEditAcc.Visible = False
        mnuEditImport.Visible = False
        mnuEditLine0.Visible = False
    End If
    mnuEditRestore.Visible = mnuEditRestore.Visible And mlngMode = 1712
End Sub




Private Sub Cmd查阅_Click()
    Call mnuEditDisplay_Click
End Sub

Private Sub Form_Activate()
    PrintRange mstrMoneySum & Space(10) & vbCrLf & mstrPrintRange
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
End Sub

Private Sub Form_Load()
    Dim strOthers(0 To 13) As String
    Dim i As Integer
    mblnFirst = True
    
    mblnCostView = zlStr.IsHavePrivs(mstrPrivs, "查看成本价")
    mbln申领核查 = IIf((zlDatabase.GetPara("申领需要核查后才能移库", glngSys, 1722, "0")) = 0, False, True)
    
    mbln移库明确批次 = IS批次移库
'    If mlngMode = 1716 Then
'        mnuEditImport.Caption = "导入申购单(&I)"
'        mnuEditImport.Visible = True
'    End If
    
    For i = 0 To 13
        strOthers(i) = ""
    Next
    '设置生产日期
    strOthers(9) = "1901-01-01"
    strOthers(10) = "1901-01-01"
    
    '0-记录状态(计划类型),1-开始单号,2-结束单号,3-材料id,4-对方部门id(或入出类别id或编制方法(计划用)),5-填制人,6-审核人,7-供应商ID,8-生产商,9-开始生产日期,10-结束生产日期,11-开始发票号,12-结束发票号,13-条码信息
    mstrOthers = strOthers
    
    '恢复设置
    Me.Caption = mstrTitle
    mstrPrintRange = "查询范围:" & Format(sys.Currentdate, "yyyy年MM月dd日") & "至" & Format(sys.Currentdate, "yyyy年MM月dd日")
    
    PrintRange mstrMoneySum & Space(10) & vbCrLf & mstrPrintRange
    
    mintUnit = Val(IIf(Val(zlDatabase.GetPara("卫材单位", glngSys, mlngMode, "0")) = 1, 1, 0))
    mstrOrder = zlDatabase.GetPara("单据排序", glngSys, mlngMode, "00")
  
    '刘兴宏:增加小数格式化串
    With mOraFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价, True)
        .FM_金额 = GetFmtString(mintUnit, g_金额, True)
        .FM_零售价 = GetFmtString(mintUnit, g_售价, True)
        .FM_数量 = GetFmtString(mintUnit, g_数量, True)
        .FM_散装零售价 = GetFmtString(0, g_售价, True)
    End With
    With mFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价)
        .FM_金额 = GetFmtString(mintUnit, g_金额)
        .FM_零售价 = GetFmtString(mintUnit, g_售价)
        .FM_数量 = GetFmtString(mintUnit, g_数量)
        .FM_散装零售价 = GetFmtString(0, g_售价)
    End With
        
    mnuViewLine3.Visible = mlngMode = 1712
    mnuViewColDefine.Visible = mlngMode = 1712
    mnuEditLine0.Visible = mlngMode = 1712
    mnuEditVerifySelect.Visible = mlngMode = 1712
    TabShow.Visible = (mlngMode = 1716)
    
    mnuEditTMPrint.Visible = mlngMode = 1712
    mnuEditLine3.Visible = mnuEditTMPrint.Visible
    
    '导入外部文件
    If InStr(mstrPrivs, ";登记;") > 0 Then
        mnuEditImportFile.Visible = (mlngMode = 1712 Or mlngMode = 1714)
    Else
        mnuEditImportFile.Visible = False
    End If
    
    '高值材料
    With vsfCostlyInfo
        '.Cols = 4
        '.Rows = 2
        .RowHeight(0) = 300
        .AutoSizeMode = flexAutoSizeColWidth
        .Visible = False
        lblCostly.Visible = .Visible
    End With
    
    If mlngMode = 1712 Then
        If gobjPlugIn Is Nothing Then
            On Error Resume Next
            Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
            If Not gobjPlugIn Is Nothing Then
                Call gobjPlugIn.Initialize(gcnOracle, glngSys, glngModul)
                If InStr(",438,0,", "," & err.Number & ",") = 0 Then
                    MsgBox "zlPlugIn 外挂部件执行 Initialize 时出错：" & vbCrLf & err.Number & vbCrLf & err.Description, vbInformation, gstrSysName
                End If
            End If
            err.Clear: On Error GoTo 0
        End If
         
        Call LoadPlugInMnu(Not gobjPlugIn Is Nothing)
    End If
    
End Sub

Private Sub Form_Resize()
    '窗体位置设置
    
    On Error Resume Next
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
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
        .Height = 400
        .Left = 0
        .Width = cbrTool.Width
        
    End With
   
    With TabShow
        .Left = 0
        .Top = cbrTool.Height
    End With
    
    With mshList
        .Top = IIf(cbrTool.Visible, cbrTool.Height, 0) + IIf(TabShow.Visible, TabShow.Height, 0)
        .Left = 0
        .Width = cbrTool.Width
        .Height = picSeparate_s.Top - .Top
    End With
    
    With Cmd查阅
        .Left = Me.ScaleWidth - .Width - 100
        .Top = mshList.Top + mshList.Height + 30
    End With
    
    If mlngMode = 1712 And vsfCostlyInfo.Visible Then
        '卫材外购入库需要显示高值卫材信息
        With mshDetail
            .Top = picSeparate_s.Top + picSeparate_s.Height + 100
            .Left = 0
            .Height = ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0) - lblCostly.Height - vsfCostlyInfo.Height
            .Width = cbrTool.Width
        End With
        With lblCostly
            .Top = mshDetail.Top + mshDetail.Height + 40
            .Left = 0
            .Width = cbrTool.Width
        End With
        With vsfCostlyInfo
            .Top = lblCostly.Top + lblCostly.Height
            .Left = 0
            .Width = cbrTool.Width
        End With
    Else
        With mshDetail
            .Top = picSeparate_s.Top + picSeparate_s.Height + 100
            .Left = 0
            .Height = ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
            .Width = cbrTool.Width
        End With
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, mstrTitle
    
    Set gobjPlugIn = Nothing
End Sub


Private Sub imgLeft_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(mshDetail.hwnd)
    lngLeft = vRect.Left + imgLeft.Left
    lngTop = vRect.Top + imgLeft.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, mshDetail, lngLeft, lngTop, imgLeft.Height)

    zl_vsGrid_Para_Save mlngMode, mshDetail, mstrTitle, "mshDetail", False, True

End Sub

Private Sub mnuEditAcc_Click()
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    With mshList
        strNo = .TextMatrix(.Row, 0)
        frmPurchaseCard.ShowCard Me, strNo, 7, .TextMatrix(.Row, GetCol(mshList, "记录状态")), mstrPrivs, blnSuccess
        If blnSuccess = True Then
            mnuViewRefresh_Click
        End If
    End With
End Sub

Private Sub mnuEditCheckBatch_Click()
    Dim frmPVB As New frmPurchaseVerifyBatch
    
    If Val(cboStock.Tag) > 0 Then
        frmPVB.ShowMe Me, mstrPrivs, 1, Val(cboStock.Tag)
        Call mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditImport_Click()
    Dim blnSuccess As Boolean
    
    If mlngMode = 1712 Then
        frmPurchaseImportFromPlane.ShowCard Me, cboStock.Text, cboStock.ItemData(cboStock.ListIndex), mintUnit, InStr(mstrPrivs, "所有库房") <> 0, blnSuccess
    ElseIf mlngMode = 1716 Then
        frmPurchaseImportFromPlane.ShowCard Me, cboStock.Text, cboStock.ItemData(cboStock.ListIndex), mintUnit, InStr(mstrPrivs, "所有库房") <> 0, blnSuccess, 1, 1716, IIf(mbln移库明确批次, 1, 0)
    End If
    
    If blnSuccess = True Then
        mnuViewRefresh_Click
    End If
End Sub
Private Sub mnuEditAdd_Click()
    Dim strNo As String
    Dim blnSuccess As Boolean

    strNo = ""
    '新增
    Select Case mlngMode
        '卫材外购入库
        Case 1712
            '解决Popupmenu模态窗体，不能继续Popupmenu
            If mblnPopupmenuCall Then
                mnuEditAdd.Tag = "1"
            Else
                frmPurchaseCard.ShowCard Me, strNo, 1, , mstrPrivs, blnSuccess
                mnuEditAdd.Tag = ""
            End If
        '卫材自制入库
        Case 1713
            frmSelfMakeCard.ShowCard Me, strNo, 1, , mstrPrivs, blnSuccess
        '卫材其他入库
        Case 1714
            frmOtherInputCard.ShowCard Me, strNo, 1, , mstrPrivs, blnSuccess
        '库存差价调整
        Case 1715
            frmDiffPriceAdjustCard.ShowCard Me, strNo, 1, , mstrPrivs, blnSuccess
        '卫材移库
        Case 1716
            frmTransferCard.ShowCard Me, strNo, 1, , mstrPrivs, blnSuccess
        '卫材领用
        Case 1717
            frmDrawCard.ShowCard Me, strNo, 1, , mstrPrivs, blnSuccess
        '卫材其他出库
        Case 1718
            frmOtherOutputCard.ShowCard Me, strNo, 1, , mstrPrivs, blnSuccess
    End Select
    
    If blnSuccess = True Then
        mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditBack_Click()
    Dim strNo As String
    err = 0: On Error GoTo ErrHand
    '回退上一次状态
    '如果未备料直接退出（只能从发送回退到备料，由备料回退到非备料）
    strNo = mshList.TextMatrix(mshList.Row, 0)
    If strNo = "" Then Exit Sub
    
    gstrSQL = "ZL_材料移库_BACK('" & strNo & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "回退")
    Call mnuViewRefresh_Click
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub mnuEditBill_Click()
    Dim strNo As String
    Dim blnSuccess As Boolean

    With mshList
        strNo = .TextMatrix(.Row, 0)
        frmPurchaseCard.ShowCard Me, strNo, 5, .TextMatrix(.Row, GetCol(mshList, "记录状态")), mstrPrivs, blnSuccess

        If blnSuccess = True Then
            mnuViewRefresh_Click
        End If
    End With
    
End Sub
Private Function CancelCheck() As Boolean
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:取消核查功能
    '参数:
    '返回:取消成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/05/13
    '-----------------------------------------------------------------------------------------------------------------------
    Dim blnSuccess As Boolean
    
    
    CancelCheck = False
    With mshList
        If mlngMode = 1712 Then
            ' Zl_材料外购_Cancelcheck
            '  No_In In 药品收发记录.NO%Type
            gstrSQL = "ZL_材料外购_CANCELCHECK('" & .TextMatrix(.Row, 0) & "')"
        ElseIf mlngMode = 1717 Then
            '领用
            gstrSQL = "Zl_材料领用_CancelVerify('" & .TextMatrix(.Row, 0) & "')"
        Else
            Exit Function
        End If
    End With
    
    err = 0: On Error GoTo ErrHandle
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "--取消核查 ")
        
    CancelCheck = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub mnuEditCancelCheck_Click()
    '----------------------------------------------------------------------------------------------------------------------------
    '功能:取消核查(只有外购入库才具备取消核查功能)
    '编制:刘兴宏
    '日期:2007/05/15
    '----------------------------------------------------------------------------------------------------------------------------
    Dim blnRefresh As Boolean
 
    If mlngMode = 1712 Or mlngMode = 1717 Then
        With mshList
            blnRefresh = (MsgBox("你确实要取消核查单据号为“" & .TextMatrix(.Row, 0) & "”的单据吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes)
            If blnRefresh Then
                blnRefresh = CancelCheck()
                If blnRefresh Then mnuViewRefresh_Click
            End If
        End With
    End If
End Sub



Private Sub mnuEditCheck_Click()
    '----------------------------------------------------------------------------------------------------------------------------
    '功能:核查指定的单据(只有外购入库才具备核查功能)
    '编制:刘兴宏
    '日期:2007/05/15
    '----------------------------------------------------------------------------------------------------------------------------
    Dim strNo  As String
    Dim blnSuccess As Boolean
    With mshList
        strNo = mshList.TextMatrix(mshList.Row, 0)
        Select Case mlngMode
            Case 1712
                frmPurchaseCard.ShowCard Me, strNo, 9, .TextMatrix(.Row, GetCol(mshList, "记录状态")), mstrPrivs, blnSuccess
            Case 1717
                frmDrawCard.ShowCard Me, strNo, 5, .TextMatrix(.Row, GetCol(mshList, "记录状态")), mstrPrivs, blnSuccess
        End Select
        If blnSuccess = True Then
            mnuViewRefresh_Click
        End If
    End With
End Sub

Private Sub mnuEditImportFile_Click()
    Dim rsTmp As ADODB.Recordset
    Dim blnVirtualStock As Boolean
    
    If cboStock.ListCount < 1 Then Exit Sub
    
    On Error GoTo ErrHandle

    With frmPurchaseImportFile
        .EntryPort mlngMode, cboStock.ItemData(cboStock.ListIndex) & ";" & cboStock.Text
        .Show vbModal, Me
        If .Result Then
            Call mnuViewRefresh_Click
        End If
    End With
    Exit Sub
    
ErrHandle:
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Sub mnuEditPrepare_Click()
    Dim strNo As String
    On Error GoTo ErrHand
    strNo = mshList.TextMatrix(mshList.Row, 0)
    
    If Trim(strNo) = "" Then Exit Sub
    
    gstrSQL = "zl_材料移库_PREPARE('" & strNo & "','" & UserInfo.用户名 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "备料")
    Call mnuViewRefresh_Click
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub mnuEditReg_Click()
    Dim strNo As String
    Dim blnSuccess As Boolean

    With mshList
        strNo = .TextMatrix(.Row, 0)
        frmPurchaseCard.ShowCard Me, strNo, 10, .TextMatrix(.Row, GetCol(mshList, "记录状态")), mstrPrivs, blnSuccess

        If blnSuccess = True Then
            mnuViewRefresh_Click
        End If
    End With
End Sub
Private Sub mnuEditRestore_Click()
    If mblnPopupmenuCall Then
        mnuEditRestore.Tag = "1"
    Else
        Dim strNo As String
        Dim blnSuccess As Boolean
        Call frmPurchaseCard.ShowCard(Me, strNo, 8, , mstrPrivs, blnSuccess)
        mnuEditRestore.Tag = ""
        If blnSuccess Then mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditSend_Click()
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    err = 0: On Error GoTo ErrHand
    strNo = mshList.TextMatrix(mshList.Row, 0)
    If Trim(strNo) = "" Then Exit Sub
    
    Call frmTransferCard.ShowCard(Me, strNo, 10, 1, mstrPrivs, blnSuccess)
    If blnSuccess Then mnuViewRefresh_Click
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub mnuEditTMPrint_Click()
    frmBarCodePrint.ShowMe Me, mOraFMT.FM_数量, cboStock
End Sub


Private Sub mnuEditVerify_Click()
    '验收
    
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    With mshList
        strNo = .TextMatrix(.Row, 0)
        Select Case mlngMode
            Case 1712
                frmPurchaseCard.ShowCard Me, strNo, 3, .TextMatrix(.Row, GetCol(mshList, "记录状态")), mstrPrivs, blnSuccess
            Case 1713
                frmSelfMakeCard.ShowCard Me, strNo, 3, .TextMatrix(.Row, GetCol(mshList, "记录状态")), mstrPrivs, blnSuccess
            Case 1714
                frmOtherInputCard.ShowCard Me, strNo, 3, .TextMatrix(.Row, GetCol(mshList, "记录状态")), mstrPrivs, blnSuccess
            Case 1715
                frmDiffPriceAdjustCard.ShowCard Me, strNo, 3, .TextMatrix(.Row, GetCol(mshList, "记录状态")), mstrPrivs, blnSuccess
            Case 1716
                frmTransferCard.ShowCard Me, strNo, 3, .TextMatrix(.Row, GetCol(mshList, "记录状态")), mstrPrivs, blnSuccess
            Case 1717
                frmDrawCard.ShowCard Me, strNo, 3, .TextMatrix(.Row, GetCol(mshList, "记录状态")), mstrPrivs, blnSuccess
            Case 1718
                If Val(.TextMatrix(.Row, GetCol(mshList, "费用ID"))) > 1 Then
                    MsgBox "备货卫材发料自动出库的单据不允许手工审核！", vbInformation, gstrSysName
                    Exit Sub
                End If
                frmOtherOutputCard.ShowCard Me, strNo, 3, .TextMatrix(.Row, GetCol(mshList, "记录状态")), mstrPrivs, blnSuccess
        End Select
    End With
    If blnSuccess = True Then
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
    Dim rsCheck As New ADODB.Recordset
    Dim intCol记录状态 As Integer
     
    With mshList
        Select Case mlngMode
            Case 1712
                If Val(.TextMatrix(.Row, GetCol(mshList, "费用ID"))) > 1 Then
                    MsgBox "备货卫材发料自动产生的入库单据不允许删除！", vbInformation, gstrSysName
                    Exit Sub
                End If
                
                strTitle = "外购入库单"
            Case 1713
                strTitle = "自制入库单"
            Case 1714
                strTitle = "其他入库单"
            Case 1715
                strTitle = "库存差价调整单"
            Case 1716
                strTitle = "卫材移库单"
            Case 1717
                strTitle = "卫材领用单"
            Case 1718
                If Val(.TextMatrix(.Row, GetCol(mshList, "费用ID"))) > 1 Then
                    MsgBox "备货卫材发料自动出库的单据不允许删除！", vbInformation, gstrSysName
                    Exit Sub
                End If
                
                strTitle = "卫材其他出库单"
            Case 1719
                strTitle = "卫材盘点单"
        End Select
        
        On Error GoTo ErrHandle
        intRow = .Row
        StrBillNo = .TextMatrix(intRow, 0)
        intReturn = MsgBox("你确实要删除单据号为“" & StrBillNo & "”的" & strTitle & "吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
        
        intRecord = .Rows - 1
        
        If intReturn = vbYes Then
            Select Case mlngMode
                Case 1712
                    gstrSQL = "zl_材料外购_Delete('" & StrBillNo & "')"
                Case 1713
                    gstrSQL = "zl_自制材料入库_Delete('" & StrBillNo & "')"
                Case 1714
                    gstrSQL = "zl_材料其他入库_Delete('" & StrBillNo & "')"
                Case 1715
                    gstrSQL = "zl_材料库存差价调整_Delete('" & StrBillNo & "')"
                Case 1716
                    intCol记录状态 = GetCol(mshList, "记录状态")
                    If .TextMatrix(.Row, intCol记录状态) = 1 Then
                    '已备料（填写了配料人）或已发送的单据，不允许入库方修改此类单据
                        If TabShow.Tab = 1 Then
                            If TestPrepare(StrBillNo) Then
                                MsgBox "已发送的单据不允许删除！", vbInformation, gstrSysName
                                Exit Sub
                            End If
                        End If
                    End If
                    gstrSQL = "zl_材料移库_Delete('" & StrBillNo & "'," & .TextMatrix(.Row, intCol记录状态) & ")"
                    
'
'
'                    '先检查是不是申领单
'                    gstrSQL = " Select Nvl(发药方式,0) 申领 From 药品收发记录 " & _
'                              " Where 单据=19 And NO='" & strBillNo & "' And 序号=1"
'                    Call OpenRecordset(rsCheck, "检查是不是申领单")
'                    If rsCheck!申领 = 0 Then
'                        gstrSQL = "zl_材料移库_Delete('" & strBillNo & "')"
'                    Else
'                        gstrSQL = "zl_材料申领_Delete('" & strBillNo & "')"
'                    End If
'
                Case 1717
                    gstrSQL = "zl_材料领用_Delete('" & StrBillNo & "')"
                Case 1718
                    gstrSQL = "zl_材料其他出库_Delete('" & StrBillNo & "')"
                Case 1719
                    gstrSQL = "zl_材料盘点_Delete('" & StrBillNo & "')"
                Case Else
                
            End Select
            If gstrSQL = "" Then Exit Sub
            zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
            
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
            mshlist_EnterCell
        End If
    End With
    stbThis.Panels(2).Text = "当前共有" & intRecord & "张单据"
    Exit Sub

ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Sub

Private Sub mnuEditDisplay_Click()
    '查看单据
    
    Dim strNo As String
    With mshList
        strNo = .TextMatrix(.Row, 0)
        If strNo = "" Then Exit Sub
        Select Case mlngMode
            Case 1712
                frmPurchaseCard.ShowCard Me, strNo, 4, Val(.TextMatrix(.Row, GetCol(mshList, "记录状态"))), mstrPrivs
            Case 1713
                frmSelfMakeCard.ShowCard Me, strNo, 4, Val(.TextMatrix(.Row, GetCol(mshList, "记录状态"))), mstrPrivs
            Case 1714
                frmOtherInputCard.ShowCard Me, strNo, 4, Val(.TextMatrix(.Row, GetCol(mshList, "记录状态"))), mstrPrivs
            Case 1715
                frmDiffPriceAdjustCard.ShowCard Me, strNo, 4, Val(.TextMatrix(.Row, GetCol(mshList, "记录状态"))), mstrPrivs
            Case 1716
                frmTransferCard.ShowCard Me, strNo, 4, Val(.TextMatrix(.Row, GetCol(mshList, "记录状态"))), mstrPrivs
            Case 1717
                frmDrawCard.ShowCard Me, strNo, 4, Val(.TextMatrix(.Row, GetCol(mshList, "记录状态"))), mstrPrivs
            Case 1718
                frmOtherOutputCard.ShowCard Me, strNo, 4, Val(.TextMatrix(.Row, GetCol(mshList, "记录状态"))), mstrPrivs
        End Select
        
    End With
    
End Sub

Private Sub mnuEditStrike_Click()
    Dim blnPurchase As Boolean, blnRefresh As Boolean
    
    
    '如果是外购(blnPurchase为真)，则直接进入冲销
    '询问是否冲销(blnPurchase为提示框返回值)，是则进入冲销
    blnPurchase = (InStr(1, "1712,1714,1716,1717,1718", mlngMode) <> 0)
    With mshList
        If Not blnPurchase Then
            blnPurchase = (MsgBox("你确实要冲销单据号为“" & .TextMatrix(.Row, 0) & "”的单据吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes)
        End If
        If blnPurchase Then
            blnRefresh = StrikeSave
            If blnRefresh Then mnuViewRefresh_Click
        End If
    End With
End Sub
Private Function CheckSelfMakeStock(ByVal str单据号 As String) As Boolean
    '------------------------------------------------------------------------------
    '功能:在冲销时，检查自制入库的可用数量是否充足
    '返回:允足,返回true,否则返回False
    '编制:刘兴宏
    '日期:2008/02/15
    '------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, int库存检查 As Integer
    err = 0: On Error GoTo ErrHand:
    gstrSQL = "" & _
        "   Select 药品id,nvl(批次,0) as 批次,库房id,sum(实际数量) as 实际数量 " & _
        "   From 药品收发记录 A " & _
        "   where 单据 = 16 And A.NO = [1] And A.记录状态 = 1 And A.入出系数=1" & _
        "   Group by 药品ID,nvl(批次,0),库房ID"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str单据号)
    With rsTemp
        If .EOF Then Exit Function
        int库存检查 = Get出库检查(cboStock.ItemData(cboStock.ListIndex))
        Do While Not .EOF
            If Check可用数量(Val(zlStr.Nvl(rsTemp!库房ID)), Val(zlStr.Nvl(rsTemp!药品ID)), _
                Val(zlStr.Nvl(rsTemp!批次, 0)), Val(zlStr.Nvl(rsTemp!实际数量)), int库存检查) = False Then Exit Function
            .MoveNext
        Loop
    End With
    CheckSelfMakeStock = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function
Private Function StrikeSave() As Boolean
    Dim blnSuccess As Boolean
    
    StrikeSave = False
    With mshList
        Select Case mlngMode
            Case 1712
                frmPurchaseCard.ShowCard Me, .TextMatrix(.Row, 0), 6, mshList.TextMatrix(mshList.Row, GetCol(mshList, "记录状态")), mstrPrivs, blnSuccess
                StrikeSave = blnSuccess
                Exit Function
            Case 1713
                If CheckSelfMakeStock(.TextMatrix(.Row, 0)) = False Then Exit Function
                gstrSQL = "zl_自制材料入库_Strike('" & .TextMatrix(.Row, 0) & "','" & UserInfo.用户名 & "')"
            Case 1714
                frmOtherInputCard.ShowCard Me, .TextMatrix(.Row, 0), 6, mshList.TextMatrix(mshList.Row, GetCol(mshList, "记录状态")), mstrPrivs, blnSuccess
                StrikeSave = blnSuccess
                Exit Function
            Case 1715
                gstrSQL = "zl_材料库存差价调整_Strike('" & .TextMatrix(.Row, 0) & "','" & UserInfo.用户名 & "')"
            Case 1716
                If mnuEditStrike.Caption = "冲销(&K)" Then
                    mint冲销方式 = 0
                ElseIf mnuEditStrike.Caption = "申请冲销(&K)" Then
                    mint冲销方式 = 1
                ElseIf mnuEditStrike.Caption = "审核冲销(&K)" Then
                    mint冲销方式 = 2
                End If
                
                frmTransferCard.ShowCard Me, .TextMatrix(.Row, 0), 6, mshList.TextMatrix(mshList.Row, GetCol(mshList, "记录状态")), mstrPrivs, blnSuccess, mint冲销方式
                StrikeSave = blnSuccess
                Exit Function
            Case 1717
                frmDrawCard.ShowCard Me, .TextMatrix(.Row, 0), 6, mshList.TextMatrix(mshList.Row, GetCol(mshList, "记录状态")), mstrPrivs, blnSuccess
                StrikeSave = blnSuccess
                Exit Function
            Case 1718
                frmOtherOutputCard.ShowCard Me, .TextMatrix(.Row, 0), 6, mshList.TextMatrix(mshList.Row, GetCol(mshList, "记录状态")), mstrPrivs, blnSuccess
                StrikeSave = blnSuccess
                Exit Function
            Case 1719
                gstrSQL = "zl_卫材盘点_Strike('" & .TextMatrix(.Row, 0) & "','" & UserInfo.用户名 & "')"
            Case Else
            
        End Select
        
        On Error GoTo ErrHandle
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "--冲销 ")
    End With
    StrikeSave = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

Private Sub mnuEditModify_Click()
    '修改
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    blnSuccess = False
    With mshList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        strNo = .TextMatrix(.Row, 0)
        
        Select Case mlngMode
            Case 1712
                If Val(.TextMatrix(.Row, GetCol(mshList, "费用ID"))) > 1 Then
                    MsgBox "备货卫材发料自动产生的入库单据不允许修改！", vbInformation, gstrSysName
                    Exit Sub
                End If
                frmPurchaseCard.ShowCard Me, strNo, 2, Val(mshList.TextMatrix(mshList.Row, GetCol(mshList, "记录状态"))), mstrPrivs, blnSuccess
            Case 1713
                frmSelfMakeCard.ShowCard Me, strNo, 2, Val(mshList.TextMatrix(mshList.Row, GetCol(mshList, "记录状态"))), mstrPrivs, blnSuccess
            Case 1714
                frmOtherInputCard.ShowCard Me, strNo, 2, Val(mshList.TextMatrix(mshList.Row, GetCol(mshList, "记录状态"))), mstrPrivs, blnSuccess
            Case 1715
                frmDiffPriceAdjustCard.ShowCard Me, strNo, 2, Val(mshList.TextMatrix(mshList.Row, GetCol(mshList, "记录状态"))), mstrPrivs, blnSuccess
            Case 1716
            
                '已备料（填写了配料人）或已发送的单据，不允许入库方修改此类单据
                If TabShow.Tab = 1 Then
                    If TestPrepare(strNo) Then
                        MsgBox "已发送的单据不允许修改！", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
                
                frmTransferCard.ShowCard Me, strNo, 2, Val(mshList.TextMatrix(mshList.Row, GetCol(mshList, "记录状态"))), mstrPrivs, blnSuccess, mint冲销方式
            Case 1717
                frmDrawCard.ShowCard Me, strNo, 2, Val(mshList.TextMatrix(mshList.Row, GetCol(mshList, "记录状态"))), mstrPrivs, blnSuccess
            Case 1718
                If Val(.TextMatrix(.Row, GetCol(mshList, "费用ID"))) > 1 Then
                    MsgBox "备货卫材发料自动出库的单据不允许修改！", vbInformation, gstrSysName
                    Exit Sub
                End If
                frmOtherOutputCard.ShowCard Me, strNo, 2, Val(mshList.TextMatrix(mshList.Row, GetCol(mshList, "记录状态"))), mstrPrivs, blnSuccess
        End Select
        If blnSuccess = True Then
            mnuViewRefresh_Click
        End If
    End With
End Sub

Private Sub mnuEditVerifyBatch_Click()
    Dim frmPVB As New frmPurchaseVerifyBatch
    
''    If mbln需要核查 Then
''        If InStr(1, mstrPrivs, ";核查;") <= 0 And InStr(1, mstrPrivs, ";审核;") <= 0 Then
''            MsgBox "你启用了卫材参数设置的“卫材外购需要核查”参数，但没有“核查”权限，也没有“审核”权限！", vbInformation, gstrSysName
''            Exit Sub
''        End If
''    Else
''        If InStr(1, mstrPrivs, ";审核;") <= 0 Then
''            MsgBox "你没有“审核”权限，请联系管理员！", vbInformation, gstrSysName
''            Exit Sub
''        End If
''    End If
    If Val(cboStock.Tag) > 0 Then
        frmPVB.ShowMe Me, mstrPrivs, IIf(mbln需要核查, 2, 0), Val(cboStock.Tag)
        Call mnuViewRefresh_Click
    End If

End Sub

Private Sub mnuEditVerifySelect_Click()
    frmPurchaseVerifySelect.ShowMe Me, mStr库房, cboStock.ListIndex
End Sub

Private Sub mnuFileBillPreview_Click()
    Dim int单位系数 As Integer
    
    On Error GoTo ErrHandle
    With mshList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        
        Select Case mintUnit
            Case 0
                int单位系数 = 0
            Case 1
                int单位系数 = 1
        End Select
        
        Select Case mlngMode
            Case 1712
                Dim rsTemp As New ADODB.Recordset
                Dim bln退库单 As Boolean
                
                gstrSQL = "Select Nvl(发药方式,0) 标志 From 药品收发记录 Where NO=[1] And 记录状态=[2] and 单据=15 And Rownum<2"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[判断是否是退库单]", .TextMatrix(.Row, 0), Val(.TextMatrix(.Row, GetCol(mshList, "记录状态"))))
                
                bln退库单 = (rsTemp!标志 = 1)
            
                ReportOpen gcnOracle, glngSys, "zl1_bill_1712", Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & Val(.TextMatrix(.Row, GetCol(mshList, "记录状态"))), "单位系数=" & int单位系数, IIf(bln退库单, "卫材退货单", "卫材外购入库单"), 1
            Case 1713
                ReportOpen gcnOracle, glngSys, "zl1_bill_1713", Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & Val(.TextMatrix(.Row, GetCol(mshList, "记录状态"))), "单位系数=" & int单位系数, 1
            Case 1714
                ReportOpen gcnOracle, glngSys, "zl1_bill_1714", Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & Val(.TextMatrix(.Row, GetCol(mshList, "记录状态"))), "单位系数=" & int单位系数, 1
            Case 1715
                ReportOpen gcnOracle, glngSys, "zl1_bill_1715", Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & Val(.TextMatrix(.Row, GetCol(mshList, "记录状态"))), "单位系数=" & int单位系数, 1
            Case 1716
                ReportOpen gcnOracle, glngSys, "zl1_bill_1716", Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & Val(.TextMatrix(.Row, GetCol(mshList, "记录状态"))), "单位系数=" & int单位系数, 1
            Case 1717
                ReportOpen gcnOracle, glngSys, "zl1_bill_1717", Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & Val(.TextMatrix(.Row, GetCol(mshList, "记录状态"))), "单位系数=" & int单位系数, 1
            Case 1718
                ReportOpen gcnOracle, glngSys, "zl1_bill_1718", Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & Val(.TextMatrix(.Row, GetCol(mshList, "记录状态"))), "单位系数=" & int单位系数, 1
            Case 1719
                ReportOpen gcnOracle, glngSys, "zl1_bill_1719", Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & Val(.TextMatrix(.Row, GetCol(mshList, "记录状态"))), "单位系数=" & int单位系数, 1
            Case Else
            
        End Select
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuFileBillPrint_Click()
    Dim strUnit As String
    Dim int单位系数 As Integer
    
    On Error GoTo ErrHandle
    With mshList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        Select Case mintUnit
            Case 0
                int单位系数 = 0
            Case 1
                int单位系数 = 1
        End Select
        
        Select Case mlngMode
            Case 1712
                Dim rsTemp As New ADODB.Recordset
                Dim bln退库单 As Boolean
                gstrSQL = "Select Nvl(发药方式,0) 标志 From 药品收发记录 Where NO=[1] And 记录状态=[2] and 单据=15 And Rownum<2"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[判断是否是退库单]", .TextMatrix(.Row, 0), Val(.TextMatrix(.Row, GetCol(mshList, "记录状态"))))
                bln退库单 = (rsTemp!标志 = 1)
                ReportOpen gcnOracle, glngSys, "zl1_bill_1712", Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & .TextMatrix(.Row, GetCol(mshList, "记录状态")), "单位系数=" & int单位系数, IIf(bln退库单, "卫材退货单", "卫材外购入库单"), 2
                
            Case 1713
                ReportOpen gcnOracle, glngSys, "zl1_bill_1713", Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & .TextMatrix(.Row, GetCol(mshList, "记录状态")), "单位系数=" & int单位系数, 2
            Case 1714
                ReportOpen gcnOracle, glngSys, "zl1_bill_1714", Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & .TextMatrix(.Row, GetCol(mshList, "记录状态")), "单位系数=" & int单位系数, 2
            Case 1715
                ReportOpen gcnOracle, glngSys, "zl1_bill_1715", Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & .TextMatrix(.Row, GetCol(mshList, "记录状态")), "单位系数=" & int单位系数, 2
            Case 1716
                ReportOpen gcnOracle, glngSys, "zl1_bill_1716", Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & .TextMatrix(.Row, GetCol(mshList, "记录状态")), "单位系数=" & int单位系数, 2
            Case 1717
                ReportOpen gcnOracle, glngSys, "zl1_bill_1717", Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & .TextMatrix(.Row, GetCol(mshList, "记录状态")), "单位系数=" & int单位系数, 2
            Case 1718
                ReportOpen gcnOracle, glngSys, "zl1_bill_1718", Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & .TextMatrix(.Row, GetCol(mshList, "记录状态")), "单位系数=" & int单位系数, 2
            Case 1719
                ReportOpen gcnOracle, glngSys, "zl1_bill_1719", Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & .TextMatrix(.Row, GetCol(mshList, "记录状态")), "单位系数=" & int单位系数, 2
            Case Else
            
        End Select
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuFileExcel_Click()
    '输出到Excel
    
    If Me.ActiveControl Is mshList Then
        mshList.Redraw = False
        subPrint 3
        mshList.Redraw = True
        mshList.Col = 0
        mshList.ColSel = mshList.Cols - 1
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
    '参数设置
    frmParaset.设置参数 mlngMode, mstrPrivs, Me, Me.Tag
    mintUnit = IIf(Val(zlDatabase.GetPara("卫材单位", glngSys, mlngMode, "0")) = 1, 1, 0)
    
    mint领用审核方式 = 0
    If mlngMode = 1717 Then
        mint领用审核方式 = Val(zlDatabase.GetPara("审核流程", glngSys, mlngMode, "0"))
    End If
    
    SetPopedom
    Call SetMenu
    mstrOrder = zlDatabase.GetPara("单据排序", glngSys, mlngMode, "00")
    mintUnit = Val(IIf(Val(zlDatabase.GetPara("卫材单位", glngSys, mlngMode, "0")) = 1, 1, 0))
    
    mintFindDay = Val(zlDatabase.GetPara("查询天数", glngSys, mlngMode, 1))
    mdtStartDate = Format(DateAdd("d", -mintFindDay, sys.Currentdate), "yyyy-MM-dd")
    mdtEndDate = Format(sys.Currentdate, "yyyy-MM-dd")
  
    '刘兴宏:增加小数格式化串
    With mOraFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价, True)
        .FM_金额 = GetFmtString(mintUnit, g_金额, True)
        .FM_零售价 = GetFmtString(mintUnit, g_售价, True)
        .FM_数量 = GetFmtString(mintUnit, g_数量, True)
        .FM_散装零售价 = GetFmtString(0, g_售价, True)
    End With
    With mFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价)
        .FM_金额 = GetFmtString(mintUnit, g_金额)
        .FM_零售价 = GetFmtString(mintUnit, g_售价)
        .FM_数量 = GetFmtString(mintUnit, g_数量)
        .FM_散装零售价 = GetFmtString(0, g_售价)
    End With
    
    Call GetList(mstrFind)
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
    '打印设置
    zlPrintSet
End Sub

Private Sub mnuHelpAbout_Click()
    '关于
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
    '帮助主题
    Dim StrWinName As String
    With mshList
        Select Case mlngMode
            Case 1712
                StrWinName = "frmMainList1"
            Case 1713
                StrWinName = "frmMainList2"
            Case 1714
                StrWinName = "frmMainList3"
            Case 1715
                StrWinName = "frmMainList4"
            Case 1716
                StrWinName = "frmMainList5"
            Case 1717
                StrWinName = "frmMainList6"
            Case 1718
                StrWinName = "frmMainList7"
            Case 1719
                StrWinName = "frmMainList8"
        End Select
    End With
    Call ShowHelp(App.ProductName, Me.hwnd, StrWinName, Int(glngSys / 100))
End Sub

Private Sub mnuHelpWebHome_Click()
    '中联主页
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    '发送反馈
    Call zlMailTo(Me.hwnd)
End Sub

Private Sub mnuPlugItem_Click(Index As Integer)
    Call ExcPlugInFun(mnuPlugItem(Index).Tag)
End Sub

Private Sub mnuViewColDefine_Click()
    Dim strColumn_All As String, strColumn_Select As String
    
    Select Case mlngMode
    Case 1712           '卫材外购入库管理
        strColumn_All = "卫材,0|商品名,1|规格,1|产地,1|批准文号,1|批号,0|生产日期,1|商品条码,1|灭菌日期,1|灭菌失效期,1|效期,0|注册证号,1|单位,1|数量,0|指导批发价,1|采购价,0|扣率,1|" & _
                        "加成率,1|结算价,0|结算金额,0|售价,0|售价金额,0|差价,0|零售单位,1|零售价,1|零售金额,1|零售差价,1|随货单号,1|验收结论,1|发票号,0|发票代码,0|发票日期,0|发票金额,0"
    Case 1713
        strColumn_All = "卫材,0|规格,1|产地,1|批号,0|生产日期,1|灭菌日期,1|灭菌失效期,1|效期,0|注册证号,1|单位,1|数量,0|指导批发价,1|采购价,1|扣率,1|" & _
                        "加成率,1|结算价,0|结算金额,0|售价,0|售价金额,0|差价,0|零售单位,1|零售价,1|零售金额,1|零售差价,1|随货单号,1|发票号,0|发票代码,0|发票日期,0|发票金额,0"
    
    Case Else
        Exit Sub
    End Select

    '取已选择列的信息'Me.Caption
    strColumn_Select = zlDatabase.GetPara("选择列", glngSys, mlngMode)
    If Not frmColSet.ShowMe(Me, strColumn_All, strColumn_Select) Then Exit Sub
    Call zlDatabase.SetPara("选择列", Split(strColumn_Select, "||")(0), glngSys, mlngMode)
    Call zlDatabase.SetPara("屏蔽列", Split(strColumn_Select, "||")(1), glngSys, mlngMode)
End Sub

Private Sub mnuViewRefresh_Click()
    '刷新
    GetList mstrFind
End Sub

Private Sub mnuViewSearch_Click()
    '查找
    Dim strFind As String
    Dim strOthers() As String

    Select Case mlngMode
        Case 1715, 1716, 1717, 1718, 1719
            strFind = FrmTransferSearch.GetSearch(Me, mlngMode, mdtStartDate, mdtEndDate, mdtVerifyStart, mdtVerifyEnd, mstrPrivs, strOthers)
        Case 1712
            strFind = FrmPurchaseSearch.GetSearch(Me, mdtStartDate, mdtEndDate, mdtVerifyStart, mdtVerifyEnd, strOthers, mstr高值耗材, mint无发票, mint有发票)
        Case 1713
            strFind = FrmSelfMakeSearch.GetSearch(Me, mdtStartDate, mdtEndDate, mdtVerifyStart, mdtVerifyEnd, strOthers)
        Case 1714
            strFind = FrmOtherInputSearch.GetSearch(Me, mdtStartDate, mdtEndDate, mdtVerifyStart, mdtVerifyEnd, strOthers)
    End Select
    
    If strFind <> "" Then
        mstrFind = strFind
        mstrOthers = strOthers
        
        GetList mstrFind
        If Format(mdtStartDate, "yyyy-mm-dd") = "1901-01-01" And Format(mdtVerifyStart, "yyyy-mm-dd") = "1901-01-01" Then
            mstrPrintRange = ""
        ElseIf Format(mdtStartDate, "yyyy-mm-dd") <> "1901-01-01" And Format(mdtVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
            mstrPrintRange = "查询范围:填制日期 " & Format(mdtStartDate, "yyyy年MM月dd日") & "至" & Format(mdtEndDate, "yyyy年MM月dd日") & "  审核日期 " & Format(mdtVerifyStart, "yyyy年MM月dd日") & "至" & Format(mdtVerifyEnd, "yyyy年MM月dd日")
        ElseIf Format(mdtStartDate, "yyyy-mm-dd") <> "1901-01-01" Then
            mstrPrintRange = "查询范围:填制日期 " & Format(mdtStartDate, "yyyy年MM月dd日") & "至" & Format(mdtEndDate, "yyyy年MM月dd日")
        ElseIf Format(mdtVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
            mstrPrintRange = "查询范围:审核日期 " & Format(mdtVerifyStart, "yyyy年MM月dd日") & "至" & Format(mdtVerifyEnd, "yyyy年MM月dd日")
        End If
        PrintRange mstrMoneySum & Space(10) & vbCrLf & mstrPrintRange
     End If
    
End Sub

Private Sub mnuViewStatus_Click()
    With mnuViewStatus
        .Checked = Not .Checked
        stbThis.Visible = .Checked
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
    If Format(mdtStartDate, "yyyy-mm-dd") = "1901-01-01" Then
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, "NO=" & strNo, "记录状态=" & intRecodeSta, "库房=" & lng库房ID)
    Else
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, "NO=" & strNo, "记录状态=" & intRecodeSta, "库房=" & lng库房ID, "开始时间=" & Format(mdtStartDate, "yyyy-mm-dd"), "结束时间=" & Format(mdtEndDate, "yyyy-mm-dd"))
    End If
End Sub
Private Sub mnuViewToolButton_Click()
    With mnuViewToolButton
        .Checked = Not .Checked
        cbrTool.Bands(1).Visible = .Checked
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

Private Sub mshDetail_AfterMoveColumn(ByVal Col As Long, Position As Long)
      zl_vsGrid_Para_Save mlngMode, mshDetail, mstrTitle, "mshDetail", False, True
End Sub

Private Sub mshDetail_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
      zl_vsGrid_Para_Save mlngMode, mshDetail, mstrTitle, "mshDetail", False, True
End Sub

Private Sub mshDetail_Click()
'    With mshDetail
'         If .Row < 1 Or .TextMatrix(.Row, 0) = "" Then Exit Sub
'         If .MouseRow = 0 Then
'            DetailSort          '列排序
'            Exit Sub
'         End If
'    End With
End Sub

Private Sub mshDetail_EnterCell()
    '卫材外购入库
    On Error GoTo ErrHandle
    If mlngMode = 1712 Then
        Dim rsTmp As ADODB.Recordset
        Dim strTmp As String
        
        vsfCostlyInfo.Visible = False
        lblCostly.Visible = False
        
        If mshDetail.Rows <= 1 Or mshDetail.Row <= 0 Then
            Call Form_Resize
            Exit Sub
        End If
        
        strTmp = "Select A.科室, A.病人姓名, A.住院号, A.床号, nvl(C.高值材料,0) 高值材料 " _
               & "From 收发记录补充信息 A, 药品收发记录 B, 材料特性 C " _
               & "Where A.收发id = B.ID And B.药品id = C.材料id and A.收发id=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strTmp, Me.Caption, mshDetail.TextMatrix(mshDetail.Row, mshDetail.ColIndex("收发ID")))
        Set vsfCostlyInfo.DataSource = rsTmp
        If rsTmp.RecordCount > 0 Then
            If rsTmp!高值材料 = 1 Then
                vsfCostlyInfo.Visible = True
                lblCostly.Visible = True
                vsfCostlyInfo.ColHidden(vsfCostlyInfo.ColIndex("高值材料")) = True
                vsfCostlyInfo.ColHidden(0) = True
            End If
        End If
        rsTmp.Close
        With vsfCostlyInfo
            .ColWidth(.ColIndex("科室")) = 2000
            .ColWidth(.ColIndex("病人姓名")) = 2000
            .ColWidth(.ColIndex("住院号")) = 2000
            .ColWidth(.ColIndex("床号")) = 1000
            .ColAlignment(.ColIndex("住院号")) = flexAlignLeftCenter
        End With
        Call Form_Resize
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
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
    Dim rsTemp As New Recordset
    Dim strUnitQuantity As String               '单位和数量格式化串
    Dim IntBill As Integer                      '单据类型  如：1、外购入库；2、
    Dim strUnit As String                       '单位名称:如门诊单位，住院单位等
    Dim str包装系数 As String
    Dim str零售价 As String
    Dim intCol As Integer
    Dim str排序 As String
    Dim str列名 As String
    Dim strTemp As String
    
'    If mlastRow = mshList.Row Then Exit Sub
    mlastRow = mshList.Row
        
    On Error GoTo ErrHandle
'    If mshList.Row >= 1 And LTrim(mshList.TextMatrix(mshList.Row, 0)) <> "" Then
    
        mshList.Col = 0
        mshList.ColSel = mshList.Cols - 1
        If mshList.RowIsVisible(mshList.Row) = False Then
           mshList.TopRow = mshList.Row
        End If
        
        If Mid(mstrOrder, 1, 1) = "0" Then
            str排序 = " 序号"
        ElseIf Mid(mstrOrder, 1, 1) = "1" Then
            str排序 = " 卫材信息"
        ElseIf Mid(mstrOrder, 1, 1) = "2" Then
            str排序 = " 名称"
        End If
        
        If Mid(mstrOrder, 2, 1) = "0" Then
            str排序 = str排序 & " asc"
        ElseIf Mid(mstrOrder, 2, 1) = "1" Then
            str排序 = str排序 & " desc"
        End If
        Select Case mintUnit
            Case 0
                strUnitQuantity = "ltrim(rtrim((to_char(A.实际数量 ," & mOraFMT.FM_数量 & ")))) AS 数量," _
                    & "D.计算单位 AS 单位,"
                str包装系数 = "1"
            Case 1
                strUnitQuantity = "ltrim(rtrim((to_char(A.实际数量 / b.换算系数," & mOraFMT.FM_数量 & ")))) AS 数量," _
                    & "B.包装单位 AS 单位,"
                str包装系数 = "B.换算系数"
        End Select
        
        Dim int单据 As Integer
        Select Case mlngMode
            Case 1712       '卫材外购入库
                IntBill = 1
                strTemp = ""
                
                If mint无发票 = 1 And mint有发票 = 0 Then
                    strTemp = " and c.发票号 is null "
                End If
                If mint有发票 = 1 And mint无发票 = 0 Then
                    strTemp = " and c.发票号 is not null "
                End If
                
                If mintUnit <> 0 Then
                    str列名 = "序号,卫材信息,商品名,规格,产地,批准文号,批号,生产日期,失效期,注册证号,数量,单位,结算价,结算金额,扣率,售价,售价金额,差价,零售价,零售单位,零售金额,零售差价,随货单号,发票号,发票代码,发票日期,付款序号,发票金额,收发id,商品条码,内部条码"
                    str零售价 = "" & _
                        "                   ltrim(rtrim(to_char(A.零售价* " & str包装系数 & "," & mOraFMT.FM_零售价 & "))) as 售价 , " & _
                        "                   ltrim(rtrim(to_char(A.零售金额," & mOraFMT.FM_金额 & ")))  as 售价金额, " & _
                        "                   ltrim(rtrim(to_char(A.差价, " & mOraFMT.FM_金额 & "))) as 差价," & _
                        "                   ltrim(rtrim(to_char(A.零售价," & mOraFMT.FM_散装零售价 & "))) as 零售价 , " & _
                        "                   D.计算单位 as 零售单位 , " & _
                        "                   ltrim(rtrim(to_char(A.零售金额," & mOraFMT.FM_金额 & ")))  as 零售金额," & _
                        "                   ltrim(rtrim(to_char(A.差价," & mOraFMT.FM_金额 & "))) as 零售差价,"
                Else
                    str列名 = "序号,卫材信息,商品名,规格,产地,批准文号,批号,生产日期,失效期,注册证号,数量,单位,结算价,结算金额,扣率,售价,售价金额,差价,随货单号,发票号,发票代码,付款序号,发票金额,收发id,商品条码,内部条码"
                    str零售价 = "" & _
                    "                   ltrim(rtrim(to_char(A.零售价*" & str包装系数 & "," & mOraFMT.FM_零售价 & "))) as 售价 , " & _
                    "                   ltrim(rtrim(to_char(A.零售金额," & mOraFMT.FM_金额 & ")))  as 售价金额, " & _
                    "                   ltrim(rtrim(to_char(A.差价," & mOraFMT.FM_金额 & "))) as 差价,"

                End If
                gstrSQL = "" & _
                    "   SELECT " & str列名 & _
                    "   From (  SELECT distinct a.序号, '[' || D.编码 || ']' || D.名称 AS 卫材信息,E.名称 As 商品名,D.规格,d.编码,zlSpellCode(d.名称) 名称, A.产地,A.批准文号, A.批号, to_char(A.生产日期,'yyyy-mm-dd') as 生产日期, to_char(A.效期,'yyyy-mm-dd') as 失效期 ,a.注册证号," & _
                                        strUnitQuantity & _
                    "                   ltrim(rtrim(to_char((A.成本价*" & str包装系数 & ")," & mOraFMT.FM_成本价 & "))) AS 结算价, ltrim(rtrim(to_char(A.成本金额," & mOraFMT.FM_金额 & "))) AS 结算金额," & _
                    "                   DECODE(A.扣率, NULL, 0, A.扣率) AS 扣率, " & str零售价 & _
                    "                   C.随货单号,C.发票号,c.发票代码 ,to_char(C.发票日期,'yyyy-mm-dd') as 发票日期,rtrim(ltrim(to_char(nvl(c.付款序号,0),'9999999999999999'))) as 付款序号, ltrim(rtrim(to_char(C.发票金额," & mOraFMT.FM_金额 & "))) as 发票金额, A.ID 收发ID, a.商品条码, a.内部条码 " & _
                    "           FROM  药品收发记录 A, 材料特性 b,收费项目目录 D,收费项目别名 E, " & _
                    "                 (Select 收发id,随货单号,发票号,发票代码,发票日期,付款序号,发票金额 From 应付记录 Where 系统标识=5 And 记录性质=0) C " & _
                    "           Where  A.药品id = B.材料id and A.药品id=D.id AND A.Id = C.收发id (+) And D.ID = E.收费细目id(+) And E.性质(+) = 3 " & strTemp & _
                    "                   AND A.记录状态 =[3] " & _
                    "                   AND A.单据 = [1] " & _
                    "                   AND A.No =[2]" & _
                    "       ) " & _
                    "   ORDER BY " & str排序
                    int单据 = 15
            Case 1713 '卫材自制入库管理
                IntBill = 2
                 gstrSQL = "" & _
                    "   select 序号,卫材信息,规格,批号,失效期,数量,单位,采购价,采购金额,售价,售价金额,差价 " & _
                    "   FROM (  SELECT DISTINCT 序号,('[' || d.编码 || ']' || d.名称) AS 卫材信息,d.规格,d.编码,zlSpellCode(d.名称) 名称,a.批号, to_char(A.效期,'yyyy-mm-dd') as 失效期," & _
                                        strUnitQuantity & _
                    "                   To_Char((a.成本价*" & str包装系数 & ")," & mOraFMT.FM_成本价 & ") AS 采购价," & _
                    "                   TO_CHAR (a.成本金额, " & mOraFMT.FM_金额 & ") AS 采购金额," & _
                    "                   TO_CHAR (a.零售价*" & str包装系数 & ", " & mOraFMT.FM_零售价 & ") AS 售价," & _
                    "                   TO_CHAR (a.零售金额," & mOraFMT.FM_金额 & ") AS 售价金额," & _
                    "                   TO_CHAR (a.差价, " & mOraFMT.FM_金额 & ") AS 差价 " & _
                    "           FROM 药品收发记录 a , 材料特性 b,收费项目目录 D " & _
                    "           Where a.药品id = b.材料id and a.药品id=d.id " & _
                    "                   AND a.记录状态 = [3] " & _
                    "                   AND a.单据 = [1] AND 入出系数=1 " & _
                    "                   AND a.no = [2] " & _
                    "         )" & _
                    "   ORDER BY " & str排序
                    
                    int单据 = 16
            Case 1714       '其他入库
                IntBill = 4
                If mintUnit <> 0 Then
                    str列名 = "序号,卫材信息,规格,产地,批准文号,批号,生产日期,失效期,数量,单位,成本价,成本金额,售价,售价金额,差价,零售价,零售金额,零售差价,商品条码,内部条码"
                    str零售价 = "" & _
                        "                   ltrim(rtrim(to_char(((A.零售金额-to_number(nvl(to_char(A.用法," & gOraFmt_Max.FM_金额 & "),'0')," & gOraFmt_Max.FM_金额 & " ))/a.实际数量)* " & str包装系数 & "," & mOraFMT.FM_零售价 & "))) as 售价 , " & _
                        "                   ltrim(rtrim(to_char(A.零售金额-to_number(nvl(to_char(A.用法," & gOraFmt_Max.FM_金额 & "),'0')," & gOraFmt_Max.FM_金额 & " )," & mOraFMT.FM_金额 & ")))  as 售价金额, " & _
                        "                   ltrim(rtrim(to_char(A.差价-to_number(nvl(to_char(A.用法," & gOraFmt_Max.FM_金额 & "),'0')," & gOraFmt_Max.FM_金额 & " ), " & mOraFMT.FM_金额 & "))) as 差价," & _
                        "                   ltrim(rtrim(to_char(A.零售价," & mOraFMT.FM_散装零售价 & "))) as 零售价 , " & _
                        "                   D.计算单位 as 零售单位 , " & _
                        "                   ltrim(rtrim(to_char(A.零售金额," & mOraFMT.FM_金额 & ")))  as 零售金额," & _
                        "                   ltrim(rtrim(to_char(A.差价," & mOraFMT.FM_金额 & "))) as 零售差价 "
                Else
                    str列名 = "序号,卫材信息,规格,产地,批准文号,批号,生产日期,失效期,数量,单位,成本价,成本金额,售价,售价金额,差价,商品条码,内部条码"
                    str零售价 = "" & _
                    "                   ltrim(rtrim(to_char(A.零售价*" & str包装系数 & "," & mOraFMT.FM_零售价 & "))) as 售价 , " & _
                    "                   ltrim(rtrim(to_char(A.零售金额," & mOraFMT.FM_金额 & ")))  as 售价金额, " & _
                    "                   ltrim(rtrim(to_char(A.差价," & mOraFMT.FM_金额 & "))) as 差价"

                End If
                                
                                
                gstrSQL = "" & _
                    "   Select * " & _
                    "   From (  SELECT distinct 序号, ('[' || D.编码 || ']' || D.名称) AS 卫材信息," & _
                    "                   D.规格,d.编码,zlSpellCode(d.名称) 名称, A.产地,A.批准文号, A.批号, to_char(A.生产日期,'yyyy-mm-dd') as 生产日期, to_char(A.效期,'yyyy-mm-dd') as 失效期," & strUnitQuantity & _
                    "                   to_char(A.成本价*" & str包装系数 & "," & mOraFMT.FM_成本价 & ") AS 成本价, to_char(A.成本金额," & mOraFMT.FM_金额 & ") AS 成本金额," & str零售价 & _
                    "           , a.商品条码, a.内部条码 " & _
                    "           FROM 药品收发记录 A, 材料特性 b,收费项目目录 D  " & _
                    "           Where  A.药品id = B.材料id and a.药品id=d.id  " & _
                    "                   AND A.记录状态 =  [3] " & _
                    "                   AND A.单据 = [1] " & _
                    "                   AND A.No =[2] " & _
                    "       ) " & _
                    "   ORDER BY " & str排序
                int单据 = 17
                
            Case 1715 '卫材库存差价调整
                IntBill = 5
                
                gstrSQL = "" & _
                    "   Select 序号,卫材信息,规格,产地,批号,失效期,单位,库存金额,库存差价,调整额 " & _
                    "   From (  SELECT distinct 序号, ('[' || D.编码 || ']' || D.名称) AS 卫材信息," & _
                    "                   D.规格,d.编码,zlSpellCode(d.名称) 名称, A.产地, A.批号, to_char(A.效期,'yyyy-mm-dd') as 失效期," & IIf(mintUnit = 0, "D.计算单位", "B.包装单位") & " as 单位," & _
                    "                   to_char(A.零售价," & mOraFMT.FM_金额 & ") AS 库存金额,to_char(A.成本价," & mOraFMT.FM_金额 & ") AS 库存差价," & _
                    "                   to_char(A.差价," & mOraFMT.FM_金额 & ")  as 调整额 " & _
                    "           FROM 药品收发记录 A, 材料特性 b,收费项目目录 D" & _
                    "           Where  A.药品id = B.材料id and A.药品id=d.id " & _
                    "                   AND A.记录状态 =  [3] " & _
                    "                   AND A.单据 = [1] " & _
                    "                   AND A.No =[2] " & _
                    "       ) " & _
                    "   ORDER BY " & str排序
                int单据 = 18
                    
            Case 1716       '卫材移库管理
                IntBill = 6
                
                gstrSQL = "" & _
                    "   SELECT 序号,卫材信息,规格,产地,批准文号,批号,失效期,填写数量,实际数量,单位,成本价,成本金额,售价,售价金额,差价,商品条码,内部条码 " & _
                    "   FROM (  SELECT DISTINCT 序号,('[' || D.编码 || ']' || d.名称) AS 卫材信息,d.规格,d.编码,zlSpellCode(d.名称) 名称,a.产地,a.批准文号, a.批号, to_char(A.效期,'yyyy-mm-dd') as 失效期," & _
                    "                   (to_char(A.填写数量 /" & str包装系数 & "," & mOraFMT.FM_数量 & ")) AS 填写数量,(to_char(A.实际数量 /" & str包装系数 & "," & mOraFMT.FM_数量 & ")) AS 实际数量," & IIf(mintUnit = 0, "D.计算单位", "B.包装单位") & " as 单位," & _
                    "                   TO_CHAR (a.成本价*" & str包装系数 & "," & mOraFMT.FM_成本价 & ") AS 成本价," & _
                    "                   TO_CHAR (a.成本金额, " & mOraFMT.FM_金额 & ") AS 成本金额," & _
                    "                   TO_CHAR (a.零售价*" & str包装系数 & ", " & mOraFMT.FM_零售价 & ") AS 售价," & _
                    "                   TO_CHAR (a.零售金额, " & mOraFMT.FM_金额 & ") AS 售价金额," & _
                    "                   TO_CHAR (a.差价, " & mOraFMT.FM_金额 & ") AS 差价, a.商品条码, a.内部条码 " & _
                    "           FROM 药品收发记录 a, 材料特性 b,收费项目目录 D " & _
                    "           Where a.药品id = b.材料id and a.药品id=d.id " & _
                    "                   AND a.记录状态 = [3] " & _
                    "                   AND a.单据 = [1] AND 入出系数=-1 " & _
                    "                   AND a.no = [2] " & _
                    "           )" & _
                    "   ORDER BY " & str排序
                    int单据 = 19
                
            Case 1717       '领用
                IntBill = 7
                
                gstrSQL = "" & _
                    "   SELECT 序号,卫材信息,规格,产地,批准文号,批号,失效期,填写数量,实际数量,单位,成本价,成本金额,售价,售价金额,差价,商品条码,内部条码 " & _
                    "   FROM (  SELECT DISTINCT 序号,('[' || D.编码 || ']' || D.名称) AS 卫材信息,D.规格,d.编码,zlSpellCode(d.名称) 名称,a.产地,a.批准文号, a.批号, to_char(A.效期,'yyyy-mm-dd') as 失效期," & _
                    "                   (to_char(A.填写数量 /" & str包装系数 & "," & mOraFMT.FM_数量 & ")) AS 填写数量,(to_char(A.实际数量 /" & str包装系数 & "," & mOraFMT.FM_数量 & ")) AS 实际数量," & IIf(mintUnit = 0, "D.计算单位", "b.包装单位") & " as 单位," & _
                    "                   TO_CHAR (A.成本价*" & str包装系数 & ", " & mOraFMT.FM_成本价 & ") AS 成本价," & _
                    "                   TO_CHAR (a.成本金额, " & mOraFMT.FM_金额 & ") AS 成本金额," & _
                    "                   TO_CHAR (a.零售价*" & str包装系数 & "," & mOraFMT.FM_零售价 & ") AS 售价," & _
                    "                   TO_CHAR (a.零售金额, " & mOraFMT.FM_金额 & ") AS 售价金额," & _
                    "                   TO_CHAR (a.差价, " & mOraFMT.FM_金额 & ") AS 差价, a.商品条码, a.内部条码 " & _
                    "           FROM 药品收发记录 a , 材料特性 b,收费项目目录 D" & _
                    "           Where a.药品id = b.材料id and a.药品id=d.id " & _
                    "                   AND A.记录状态 = [3] " & _
                    "                   AND a.单据 =[1] " & _
                    "                   AND a.no = [2] )" & _
                    "   ORDER BY " & str排序
                int单据 = 20
                    
            Case 1718   '其他出库
                IntBill = 11
                If mshList.TextMatrix(mshList.Row, 1) = "材料外销" Then
                    str列名 = "序号,卫材信息,规格,产地,批准文号,批号,失效期,数量,单位,成本价,成本金额,售价,售价金额,差价,外销价,外销金额,增值税率,税金,商品条码,内部条码"
                Else
                    str列名 = "序号,卫材信息,规格,产地,批准文号,批号,失效期,数量,单位,成本价,成本金额,售价,售价金额,差价,商品条码,内部条码"
                End If
                
                gstrSQL = "" & _
                    "   Select " & str列名 & _
                    "   From (  SELECT distinct 序号, ('[' || d.编码 || ']' ||d.名称) AS 卫材信息," & _
                    "                   d.规格,d.编码,zlSpellCode(d.名称) 名称, A.产地,A.批准文号, A.批号, to_char(A.效期,'yyyy-mm-dd') as 失效期," & strUnitQuantity & _
                    "                   to_char(a.成本价*" & str包装系数 & "," & mOraFMT.FM_成本价 & ") AS 成本价, to_char(A.成本金额," & mOraFMT.FM_金额 & ") AS 成本金额," & _
                    "                   to_char(A.零售价*" & str包装系数 & "," & mOraFMT.FM_零售价 & ") as 售价 , to_char(A.零售金额," & mOraFMT.FM_金额 & ")  as 售价金额, to_char(A.差价," & mOraFMT.FM_金额 & ") as 差价 "
                    
                If mshList.TextMatrix(mshList.Row, 1) = "材料外销" Then
                    gstrSQL = gstrSQL & " ,LTRIM(TO_CHAR(A.单量*" & str包装系数 & "," & mOraFMT.FM_零售价 & ")) AS 外销价,LTRIM(TO_CHAR(A.单量*A.实际数量," & mOraFMT.FM_金额 & ")) AS 外销金额,LTRIM(TO_CHAR(Nvl(A.频次,0)/100," & mOraFMT.FM_金额 & ")) As 增值税率,LTRIM(TO_CHAR(A.单量*A.实际数量*(Nvl(A.频次,0)/100/(1+Nvl(A.频次,0)/100))," & mOraFMT.FM_金额 & ")) As 税金 "
                End If
                    
                gstrSQL = gstrSQL & ", a.商品条码, a.内部条码  FROM 药品收发记录 A , 材料特性 b,收费项目目录 D" & _
                    "           Where  A.药品id = B.材料id and a.药品id=d.id " & _
                    "                   AND A.记录状态 =  [3] " & _
                    "                   AND A.单据 = [1] " & _
                    "                   AND A.No =[2] " & _
                    "       ) " & _
                    "   ORDER BY " & str排序
                int单据 = 21
            Case 1719 '卫材盘点管理
                IntBill = 12
                
                gstrSQL = "" & _
                    "   SELECT * " & _
                    "   FROM (  SELECT DISTINCT 序号,('[' || d.编码 || ']' || d.名称) AS 卫材信息," & _
                    "                   d.规格,d.编码,zlSpellCode(d.名称) 名称,a.产地," & IIf(strUnit = "包装单位", "d.包装单位", "b." & strUnit) & " as 单位,a.批号, to_char(A.效期,'yyyy-mm-dd') as 失效期," & _
                    "                   (to_char(A.填写数量 /" & str包装系数 & "," & mOraFMT.FM_数量 & ")) AS 帐面数," & _
                    "                   (to_char(A.扣率 /" & str包装系数 & "," & mOraFMT.FM_数量 & ")) AS 实盘数," & _
                    "                   Decode(Sign(A.扣率-A.填写数量),-1,'亏',1,'盈','平') as 标志," & _
                    "                   (to_char(A.实际数量 /" & str包装系数 & "," & mOraFMT.FM_数量 & ")) AS 数量差," & _
                    "                   TO_CHAR (a.零售价*" & str包装系数 & ", " & mOraFMT.FM_零售价 & ") AS 售价," & _
                    "                   TO_CHAR (a.零售金额, " & mOraFMT.FM_金额 & ") AS 金额差," & _
                    "                   TO_CHAR (a.差价, " & mOraFMT.FM_金额 & ") AS 差价差, " & _
                    "                   TO_CHAR ((A.扣率 / " & str包装系数 & ")*(a.零售价*" & str包装系数 & "), " & mOraFMT.FM_金额 & ") as 盘点金额 " & _
                    "           FROM 药品收发记录 a, 材料特性 b,收费项目目录 D" & _
                    "           Where a.药品id = b.材料id and a.药品id=d.id  " & _
                    "                   AND 记录状态 = [3] " & _
                    "                   AND a.单据 = [1] " & _
                    "                   AND a.no = [2] " & _
                    "       )" & _
                    "   ORDER BY " & str排序
                int单据 = 22
            End Select
            
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, int单据, mshList.TextMatrix(mshList.Row, 0), Val(mshList.TextMatrix(mshList.Row, GetCol(mshList, "记录状态"))))
        
        Set mshDetail.DataSource = rsTemp
        With mshDetail
            If rsTemp.RecordCount = 0 Then
                .Rows = 2
                .Clear 1
            End If
            rsTemp.Close
        End With
        If mlngMode = 1712 Then
            mshDetail.ColHidden(mshDetail.ColIndex("收发ID")) = True
            Call mshDetail_EnterCell
        End If

    SetDetailColWidth
    SetEnable
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

Private Sub mshlist_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    If mnuEdit.Visible = False Then Exit Sub
    '解决Popupmenu模态窗体，不能继续Popupmenu
    mblnPopupmenuCall = True
    PopupMenu mnuEdit, 2
    mblnPopupmenuCall = False
    If mnuEditAdd.Tag = "1" Then
        Call mnuEditAdd_Click
    ElseIf mnuEditRestore.Tag = "1" Then
        Call mnuEditRestore_Click
    End If
    
End Sub

Private Sub picSeparate_s_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
        If Button <> 1 Then Exit Sub
        mintOldY = y
End Sub

Private Sub picSeparate_s_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    '分割条设置
    
    If Button <> 1 Then Exit Sub
    
    With picSeparate_s
        If .Top + y < 2000 Then Exit Sub
        If .Top + y > ScaleHeight - 2000 Then Exit Sub
        .Move .Left, .Top + y - mintOldY
    End With
    
    With mshList
        .Top = IIf(cbrTool.Visible, cbrTool.Height, 0) + IIf(TabShow.Visible, TabShow.Height, 0)
        .Height = picSeparate_s.Top - .Top
    End With
    
    With Cmd查阅
        .Top = mshList.Top + mshList.Height + 30
    End With
    
    With mshDetail
        .Top = picSeparate_s.Top + picSeparate_s.Height + 100
        If mlngMode = 1712 Then
            .Height = ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0) _
                    - IIf(vsfCostlyInfo.Visible, vsfCostlyInfo.Height, 0) _
                    - IIf(lblCostly.Visible, lblCostly.Height, 0)
        Else
            .Height = ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
        End If
    End With
    
End Sub

Private Sub picSeparate_s_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    mintOldY = 0
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
        Case "Prepare"
            mnuEditPrepare_Click
        Case "Send"
            mnuEditSend_Click
        Case "Back"
            mnuEditBack_Click
        Case "Verify"
            mnuEditVerify_Click
        Case "Check"
            mnuEditCheck_Click
        Case "CancelCheck"
            mnuEditCancelCheck_Click
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
    Dim strVerify As String
    Dim bln核查 As Boolean
    Dim intCol As Integer
    Dim intTemp As Integer
    
    If mlngMode = 1712 Then
        If mbln需要核查 Then
            mnuEditCheckBatch.Enabled = InStr(mstrPrivs, ";核查;") > 0
        End If
        mnuEditVerifyBatch.Enabled = InStr(mstrPrivs, ";审核;") > 0
    End If
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
            tlbTool.Buttons("Strike").Enabled = False
        
            mnuEditCheck.Enabled = False
            mnuEditCancelCheck.Enabled = False
            tlbTool.Buttons("Check").Enabled = False
            tlbTool.Buttons("CancelCheck").Enabled = False
        
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
            
            If mnuEditBill.Visible = True Then
                mnuEditBill.Enabled = False
            End If
            
            If mnuEditReg.Visible = True Then
                mnuEditReg.Enabled = False
            End If
            
            If mnuEditAcc.Visible Then
                mnuEditAcc.Enabled = False
            End If
            
            If mnuEditImport.Visible Then
                mnuEditImport.Enabled = True
            End If
            
            If mnuEditPrepare.Visible Then
                mnuEditPrepare.Enabled = False
                mnuEditSend.Enabled = False
                mnuEditBack.Enabled = False
                tlbTool.Buttons("Prepare").Enabled = False
                tlbTool.Buttons("Send").Enabled = False
                tlbTool.Buttons("Back").Enabled = False
            End If
        Else
            mnuFilePreView.Enabled = True
            mnuFilePrint.Enabled = True
            mnuFileBillPreview.Enabled = True
            mnuFileBillPrint.Enabled = True
            mnuFileExcel.Enabled = True
            tlbTool.Buttons("Print").Enabled = True
            tlbTool.Buttons("PrintView").Enabled = True
            
            '只有外购入库单才有
            If mnuEditBill.Visible = True Then
                mnuEditBill.Enabled = False
            End If
            
            If mnuEditReg.Visible = True Then
                mnuEditReg.Enabled = False
            End If
            
            If mnuEditAcc.Visible Then
                mnuEditAcc.Enabled = False
            End If
            
            If mnuEditImport.Visible Then
                mnuEditImport.Enabled = True
            End If
            
            If mlngMode = 1719 Then
                strVerify = .TextMatrix(.Row, .Cols - 6)
            Else
                If mlngMode = 1716 Then '卫材移库
                    If mint移库处理流程 = 1 Then
                        If TabShow.Tab = 0 Then
                            If Val(.TextMatrix(.Row, GetCol(mshList, "记录状态"))) Mod 3 = 2 Then
                                strVerify = .TextMatrix(.Row, GetCol(mshList, "接收人")) '冲销单据是否审核
                            Else
                                strVerify = .TextMatrix(.Row, GetCol(mshList, "备料人")) '是否备料
                            End If
                        Else
                            strVerify = .TextMatrix(.Row, GetCol(mshList, "接收人")) '是否接收
                        End If
                    Else
                        strVerify = .TextMatrix(.Row, GetCol(mshList, "接收人")) '是否审核
                    End If
                Else
                    strVerify = .TextMatrix(.Row, GetCol(mshList, "审核日期"))    '审核日期
                End If
            End If
            
            If strVerify = "" Then    '未审核单
                If mlngMode = 1712 Then
                
                    '刘兴宏:加入核查流程2007/05/13
                    intCol = GetCol(mshList, "核查日期")
                    If intCol >= 0 Then
                        bln核查 = Trim(.TextMatrix(.Row, intCol)) <> ""
                    Else
                        bln核查 = False
                    End If
                    
                    If mnuEditModify.Visible = True Then
                        mnuEditModify.Enabled = Not bln核查
                        tlbTool.Buttons("Modify").Enabled = Not bln核查
                    End If
                    If mnuEditDel.Visible = True Then
                        mnuEditDel.Enabled = Not bln核查
                        tlbTool.Buttons("Delete").Enabled = Not bln核查
                    End If
                    
                    mnuEditCheck.Enabled = Not bln核查
                    mnuEditCancelCheck.Enabled = bln核查
                    tlbTool.Buttons("Check").Enabled = Not bln核查
                    tlbTool.Buttons("CancelCheck").Enabled = bln核查
                    
                    If mnuEditVerify.Visible = True Then
                        mnuEditVerify.Enabled = IIf(mbln需要核查, bln核查, True)
                        tlbTool.Buttons("Verify").Enabled = IIf(mbln需要核查, bln核查, True)
                    End If
                    
                ElseIf mlngMode = 1717 Then
                    intCol = GetCol(mshList, "核查日期")
                    If intCol >= 0 Then
                        bln核查 = Trim(.TextMatrix(.Row, intCol)) <> ""
                    Else
                        bln核查 = False
                    End If
                    
                    If mnuEditModify.Visible = True Then
                        mnuEditModify.Enabled = Not bln核查
                        tlbTool.Buttons("Modify").Enabled = Not bln核查
                    End If
                    If mnuEditDel.Visible = True Then
                        mnuEditDel.Enabled = Not bln核查
                        tlbTool.Buttons("Delete").Enabled = Not bln核查
                    End If
                    
                    mnuEditCheck.Enabled = Not bln核查
                    mnuEditCancelCheck.Enabled = bln核查
                    tlbTool.Buttons("Check").Enabled = Not bln核查
                    tlbTool.Buttons("CancelCheck").Enabled = bln核查
                    
                    If mnuEditVerify.Visible = True Then
                        mnuEditVerify.Enabled = IIf(mint领用审核方式 = 1, bln核查, True)
                        tlbTool.Buttons("Verify").Enabled = mnuEditVerify.Enabled
                    End If
                Else
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
                End If
                If mnuEditStrike.Visible = True Then
                    mnuEditStrike.Enabled = False
                    tlbTool.Buttons("Strike").Enabled = False
                End If
                 
                If mnuEditDisplay.Visible = True Then
                    mnuEditDisplay.Enabled = True
                End If
                
                '移库单，根据当前选择的页面，当前单据设置按钮状态
                If mlngMode = 1716 Then
                    If TabShow.Tab = 0 Then
                        If Val(.TextMatrix(.Row, GetCol(mshList, "记录状态"))) Mod 3 = 2 Then
                            tlbTool.Buttons("Modify").Enabled = False
                            tlbTool.Buttons("Delete").Enabled = False
                            tlbTool.Buttons("Prepare").Enabled = False
                            tlbTool.Buttons("Send").Enabled = False
                            tlbTool.Buttons("Back").Enabled = False
                            tlbTool.Buttons("Verify").Enabled = False
                            tlbTool.Buttons("Strike").Enabled = True
                            mnuEditModify.Enabled = False
                            mnuEditDel.Enabled = False
                            mnuEditPrepare.Enabled = False
                            mnuEditVerify.Enabled = False
                            mnuEditStrike.Enabled = True
                        Else
                            mnuEditPrepare.Enabled = (.TextMatrix(.Row, 0) <> "")
                            mnuEditSend.Enabled = False
                            mnuEditBack.Enabled = False
                            tlbTool.Buttons("Prepare").Enabled = mnuEditPrepare.Enabled
                            tlbTool.Buttons("Send").Enabled = False
                            tlbTool.Buttons("Back").Enabled = False
                            
                            '如果该单据已审核，不允许备药与发送
                            If TestVerify(mshList.TextMatrix(mshList.Row, 0)) Then
                                mnuEditPrepare.Enabled = False
                                mnuEditSend.Enabled = False
                                mnuEditBack.Enabled = False
                                tlbTool.Buttons("Prepare").Enabled = False
                                tlbTool.Buttons("Send").Enabled = False
                                tlbTool.Buttons("Back").Enabled = False
                                tlbTool.Buttons("Strike").Enabled = False
                            Else
                                tlbTool.Buttons("Strike").Enabled = False
                            End If
                        End If
                    Else
                        If Val(.TextMatrix(.Row, GetCol(mshList, "记录状态"))) Mod 3 = 2 Then
                            tlbTool.Buttons("Modify").Enabled = False
                            tlbTool.Buttons("Strike").Enabled = False
                            tlbTool.Buttons("Verify").Enabled = False
                            tlbTool.Buttons("Delete").Enabled = True
                            mnuEditModify.Enabled = False
                            mnuEditStrike.Enabled = False
                            mnuEditVerify.Enabled = False
                            mnuEditDel.Enabled = True
                        Else
                            mnuEditVerify.Enabled = TestPrepare(.TextMatrix(.Row, 0))
                            tlbTool.Buttons("Verify").Enabled = mnuEditVerify.Enabled
                        End If
                    End If
                End If
                
            ElseIf .TextMatrix(.Row, GetCol(mshList, "记录状态")) = 1 Then  '审核单
                If mlngMode = 1712 Or mlngMode = 1717 Then
                    '刘兴宏:加入核查功能2007/05/13
                    mnuEditCheck.Enabled = False
                    tlbTool.Buttons("Check").Enabled = False
                    mnuEditCheck.Enabled = False
                    mnuEditCancelCheck.Enabled = False
                    tlbTool.Buttons("Check").Enabled = False
                    tlbTool.Buttons("CancelCheck").Enabled = False
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
                    If mlngMode = 1715 And .TextMatrix(.Row, .Cols - 1) = "1" Then
                        mnuEditStrike.Enabled = False
                        tlbTool.Buttons("Strike").Enabled = False
                    Else
                        mnuEditStrike.Enabled = True
                        tlbTool.Buttons("Strike").Enabled = True
                    End If
                End If
                 
                If mnuEditDisplay.Visible = True Then
                    mnuEditDisplay.Enabled = True
                End If
                
                '只有外购入库单才有
                If mnuEditBill.Visible = True Then
'                    If Val(mshDetail.TextMatrix(mshDetail.Row, mshDetail.Cols - 2)) = 0 Then
                        mnuEditBill.Enabled = True
'                    End If
                End If
                
                If mnuEditReg.Visible = True Then
                    If Val(mshDetail.TextMatrix(mshDetail.Row, mshDetail.Cols - 2)) = 0 Then
                        mnuEditReg.Enabled = True
                    End If
                End If
                
                If mlngMode = 1716 And TabShow.Tab = 0 Then
                    mnuEditAcc.Enabled = (Trim(mshList.TextMatrix(mshList.Row, GetCol(mshList, "发送日期"))) <> "")
                Else
                    mnuEditAcc.Enabled = (Trim(mshList.TextMatrix(mshList.Row, GetCol(mshList, "审核日期"))) <> "")
                End If
                
                If mlngMode = 1716 Then
                    If TabShow.Tab = 0 Then
                        mnuEditPrepare.Enabled = False
                        mnuEditSend.Enabled = (mshList.TextMatrix(mshList.Row, mshList.Cols - 3) = "")
                        mnuEditBack.Enabled = True
                        tlbTool.Buttons("Prepare").Enabled = False
                        tlbTool.Buttons("Send").Enabled = mnuEditSend.Enabled
                        tlbTool.Buttons("Back").Enabled = True
                        '如果该单据已审核，不允许备料与发送
                        If TestVerify(mshList.TextMatrix(mshList.Row, 0)) Then
                            mnuEditPrepare.Enabled = False
                            mnuEditSend.Enabled = False
                            mnuEditBack.Enabled = False
                            tlbTool.Buttons("Prepare").Enabled = False
                            tlbTool.Buttons("Send").Enabled = False
                            tlbTool.Buttons("Back").Enabled = False
                        End If
                        
                        If mnuEditStrike.Visible = True Then
                            If mint冲销申请 = 1 Then
                                mnuEditStrike.Enabled = False
                                tlbTool.Buttons("Strike").Enabled = False
                            Else
                                mnuEditStrike.Enabled = True
                                tlbTool.Buttons("Strike").Enabled = True
                            End If
                        End If
                    Else
                        If mnuEditStrike.Visible = True Then
                            mnuEditStrike.Enabled = True
                            tlbTool.Buttons("Strike").Enabled = True
                        End If
                    End If
                End If
                
            Else   '2,3 冲销单（已付款的单据不允许财务审核，同样，财务审核后的单据不允许冲销）
                If mnuEditBill.Visible = True Then
                    mnuEditBill.Enabled = True
                End If
                
                If mnuEditReg.Visible = True Then
                    mnuEditReg.Enabled = True
                End If
                
                If mlngMode = 1712 Or mlngMode = 1717 Then
                    '刘兴宏:加入核查功能2007/05/13
                    mnuEditCheck.Enabled = False
                    mnuEditCancelCheck.Enabled = False
                    tlbTool.Buttons("Check").Enabled = False
                    tlbTool.Buttons("CancelCheck").Enabled = False
                    
                End If
                
                If Val(.TextMatrix(.Row, GetCol(mshList, "记录状态"))) Mod 3 = 0 Then
                    '财务审核
                    intTemp = GetCol(mshList, "财务标志")
                    If intTemp >= 0 Then intTemp = Val(.TextMatrix(.Row, intTemp))
                    .ToolTipText = IIf(intTemp = 1, "财务审核的原单据", "冲销单据的原单据")
                    If mnuEditStrike.Visible = True Then
                        mnuEditStrike.Enabled = True
                        tlbTool.Buttons("Strike").Enabled = True
                    Else
                        tlbTool.Buttons("Strike").Enabled = False
                    End If
                    If mlngMode = 1716 And TabShow.Tab = 0 Then
                        mnuEditAcc.Enabled = (Trim(mshList.TextMatrix(mshList.Row, GetCol(mshList, "发送日期"))) = "")
                    Else
                        mnuEditAcc.Enabled = (Trim(mshList.TextMatrix(mshList.Row, GetCol(mshList, "审核日期"))) = "")
                    End If
                    
                    If mlngMode = 1716 Then
                        If TabShow.Tab = 0 Then
                            mnuEditPrepare.Enabled = False
                            mnuEditSend.Enabled = (mshList.TextMatrix(mshList.Row, mshList.Cols - 3) = "")
                            mnuEditBack.Enabled = True
                            tlbTool.Buttons("Prepare").Enabled = False
                            tlbTool.Buttons("Send").Enabled = mnuEditSend.Enabled
                            tlbTool.Buttons("Back").Enabled = True
                            '如果该单据已审核，不允许备料与发送
                            If TestVerify(mshList.TextMatrix(mshList.Row, 0)) Then
                                mnuEditPrepare.Enabled = False
                                mnuEditSend.Enabled = False
                                mnuEditBack.Enabled = False
                                tlbTool.Buttons("Prepare").Enabled = False
                                tlbTool.Buttons("Send").Enabled = False
                                tlbTool.Buttons("Back").Enabled = False
                            End If
                            
                            If mnuEditStrike.Visible = True Then
                                If mint冲销申请 = 1 Then
                                    mnuEditStrike.Enabled = False
                                    tlbTool.Buttons("Strike").Enabled = False
                                Else
                                    mnuEditStrike.Enabled = True
                                    tlbTool.Buttons("Strike").Enabled = True
                                End If
                            End If
                        Else
                            If mnuEditStrike.Visible = True Then
                                mnuEditStrike.Enabled = True
                                tlbTool.Buttons("Strike").Enabled = True
                            End If
                        End If
                    End If
                ElseIf .TextMatrix(.Row, GetCol(mshList, "记录状态")) Mod 3 = 2 Then
                    .ToolTipText = "冲销单据"
                    '财务审核
                    intTemp = GetCol(mshList, "财务标志")
                    If intTemp >= 0 Then intTemp = Val(.TextMatrix(.Row, intTemp))
                    .ToolTipText = IIf(intTemp = 1, "财务审核的冲销单据", "冲销单据")
                    
                    If mnuEditStrike.Visible = True Then
                        mnuEditStrike.Enabled = False
                        tlbTool.Buttons("Strike").Enabled = False
                    End If
                    
                    If mlngMode = 1716 Then '移库
                        If TabShow.Tab = 0 Then
                            If mnuEditVerify.Visible = True Then
                                mnuEditVerify.Enabled = False
                                tlbTool.Buttons("Verify").Enabled = False
                            End If
                            
                            If mint冲销申请 = 1 Then mnuEditStrike.Visible = True
                            If strVerify = "" Then
                                mnuEditStrike.Enabled = True
                                tlbTool.Buttons("Strike").Enabled = True
                            Else
                                mnuEditStrike.Enabled = False
                                tlbTool.Buttons("Strike").Enabled = False
                            End If
                        Else

                        End If
                    End If
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
                 
                If mnuEditDisplay.Visible = True Then
                    mnuEditDisplay.Enabled = True
                End If
                
                If mlngMode = 1304 Or mlngMode = 1716 Then
                    If TabShow.Tab = 0 Then
                        mnuEditPrepare.Enabled = False
                        mnuEditSend.Enabled = False
                        mnuEditBack.Enabled = False
                        tlbTool.Buttons("Prepare").Enabled = False
                        tlbTool.Buttons("Send").Enabled = mnuEditSend.Enabled
                        tlbTool.Buttons("Back").Enabled = False
                        '如果该单据已审核，不允许备药与发送
                        If TestVerify(mshList.TextMatrix(mshList.Row, 0)) Then
                            mnuEditPrepare.Enabled = False
                            mnuEditSend.Enabled = False
                            mnuEditBack.Enabled = False
                            tlbTool.Buttons("Prepare").Enabled = False
                            tlbTool.Buttons("Send").Enabled = False
                            tlbTool.Buttons("Back").Enabled = False
                        End If
                    End If
                End If
            End If
        End If
    End With
    Cmd查阅.Enabled = mnuEditDisplay.Enabled
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
    
    objPrint.Title.Text = mstrTitle
        
    objRow.Add "时间：" & strRange
    objRow.Add "部门：" & cboStock.Text
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
        
    objRow.Add "打印人:" & UserInfo.用户名
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

Private Sub subExcel(bytMode As Byte)
'功能:进行输出到EXCEL
'参数:bytMode3 输出到EXCEL

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
    objRow.Add "NO." & Trim(mshList.TextMatrix(mshList.Row, GetCol(mshList, "NO")))
    objPrint.UnderAppRows.Add objRow
    
    Select Case mlngMode
        Case 1712       '卫材外购入库管理
            Set objRow = New zlTabAppRow
            objRow.Add "库房：" & Trim(cboStock.Text)
            objRow.Add "供应商：" & Trim(mshList.TextMatrix(mshList.Row, GetCol(mshList, "供应商")))
            objPrint.UnderAppRows.Add objRow
                
        Case 1713       '卫材自制入库管理
            Set objRow = New zlTabAppRow
            objRow.Add "库房：" & Trim(cboStock.Text)
            objRow.Add "制剂室：" & Trim(mshList.TextMatrix(mshList.Row, GetCol(mshList, "制剂室")))
            objPrint.UnderAppRows.Add objRow
            
        Case 1714      '卫材其他入库管理
            Set objRow = New zlTabAppRow
            objRow.Add "库房：" & Trim(cboStock.Text)
            objRow.Add "入出类别：" & Trim(mshList.TextMatrix(mshList.Row, GetCol(mshList, "入出类别")))
            objPrint.UnderAppRows.Add objRow
        Case 1715       '库存差价调整管理
            Set objRow = New zlTabAppRow
            objRow.Add "库房：" & Trim(cboStock.Text)
            objPrint.UnderAppRows.Add objRow
            
        Case 1716       '卫材移库管理
            Set objRow = New zlTabAppRow
            If TabShow.Tab = 0 Then
                objRow.Add "移出库房：" & Trim(cboStock.Text)
                objRow.Add "移入库房：" & Trim(mshList.TextMatrix(mshList.Row, GetCol(mshList, "移入库房")))
            Else
                objRow.Add "移入库房：" & Trim(cboStock.Text)
                objRow.Add "移出库房：" & Trim(mshList.TextMatrix(mshList.Row, GetCol(mshList, "移出库房")))
            End If
            objPrint.UnderAppRows.Add objRow
        Case 1717       '卫材领用管理
            Set objRow = New zlTabAppRow
            objRow.Add "发卫材库房：" & Trim(cboStock.Text)
            objRow.Add "领用部门：" & Trim(mshList.TextMatrix(mshList.Row, GetCol(mshList, "领用部门")))
            objPrint.UnderAppRows.Add objRow
            
        Case 1718       '卫材其他出库管理
            Set objRow = New zlTabAppRow
            objRow.Add "库房：" & Trim(cboStock.Text)
            objRow.Add "入出类别：" & Trim(mshList.TextMatrix(mshList.Row, GetCol(mshList, "入出类别")))
            objPrint.UnderAppRows.Add objRow
        Case 1719       '卫材盘点管理
            Set objRow = New zlTabAppRow
            objRow.Add "盘点库房：" & Trim(cboStock.Text)
            objRow.Add "盘点时间：" & Trim(mshList.TextMatrix(mshList.Row, GetCol(mshList, "盘点时间")))
            objPrint.UnderAppRows.Add objRow
    End Select
        
    Set objRow = New zlTabAppRow
    objRow.Add "摘要:" & mshList.TextMatrix(mshList.Row, GetCol(mshList, "摘要"))
    objPrint.BelowAppRows.Add objRow
        
    Set objRow = New zlTabAppRow
    objRow.Add "填制人:" & mshList.TextMatrix(mshList.Row, GetCol(mshList, "填制人")) & "  填制日期:" & mshList.TextMatrix(mshList.Row, GetCol(mshList, "填制日期"))
    
    '单独处理移库模块
    If mlngMode = 1716 Then
        If TabShow.Tab = 0 Then
            objRow.Add "审核人:  审核日期:"
        Else
            objRow.Add "审核人:" & mshList.TextMatrix(mshList.Row, GetCol(mshList, "接收人")) & "  审核日期:" & mshList.TextMatrix(mshList.Row, GetCol(mshList, "审核日期"))
        End If
    Else
        objRow.Add "审核人:" & mshList.TextMatrix(mshList.Row, GetCol(mshList, "审核人")) & "  审核日期:" & mshList.TextMatrix(mshList.Row, GetCol(mshList, "审核日期"))
    End If
    objPrint.BelowAppRows.Add objRow
    
    Set objPrint.Body = mshDetail
    zlPrintOrView1Grd objPrint, bytMode
End Sub

Private Sub tlbTool_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool
    End If
End Sub

'对单据头列排序
Private Sub ListSort()
    Dim intCol As Integer
    Dim intRow As Integer
    Dim intTemp As String
    
    err = 0: On Error Resume Next
    With mshList
        If .Rows > 1 Then
            .Redraw = False
            intCol = .MouseCol
            .Col = intCol
            .ColSel = intCol
            intTemp = .TextMatrix(.Row, intCol)

            Select Case mlngMode
                Case 1712, 1718
                    If InStr(1, "345", intCol) <> 0 Then '345为数字,其他为字符
                        If intCol = mintPreCol And mintsort = flexSortNumericDescending Then
                           .Sort = flexSortNumericAscending
                           mintsort = flexSortNumericAscending
                        Else
                           .Sort = flexSortNumericDescending
                           mintsort = flexSortNumericDescending
                        End If
                    Else
                        If intCol = mintPreCol And mintsort = flexSortStringNoCaseDescending Then
                           .Sort = flexSortStringNoCaseAscending
                           mintsort = flexSortStringNoCaseAscending
                        Else
                           .Sort = flexSortStringNoCaseDescending
                           mintsort = flexSortStringNoCaseDescending
                        End If
                    End If
                Case 1713, 1714   '2,34列为数字，其他为字符
                    If InStr(1, "234", intCol) <> 0 Then
                        If intCol = mintPreCol And mintsort = flexSortNumericDescending Then
                           .Sort = flexSortNumericAscending
                           mintsort = flexSortNumericAscending
                        Else
                           .Sort = flexSortNumericDescending
                           mintsort = flexSortNumericDescending
                        End If
                    Else
                        If intCol = mintPreCol And mintsort = flexSortStringNoCaseDescending Then
                           .Sort = flexSortStringNoCaseAscending
                           mintsort = flexSortStringNoCaseAscending
                        Else
                           .Sort = flexSortStringNoCaseDescending
                           mintsort = flexSortStringNoCaseDescending
                        End If
                    End If
                Case 1715               '1,2,3列为数字，其他为字符
                    If InStr(1, "123", intCol) <> 0 Then
                        If intCol = mintPreCol And mintsort = flexSortNumericDescending Then
                           .Sort = flexSortNumericAscending
                           mintsort = flexSortNumericAscending
                        Else
                           .Sort = flexSortNumericDescending
                           mintsort = flexSortNumericDescending
                        End If
                    Else
                        If intCol = mintPreCol And mintsort = flexSortStringNoCaseDescending Then
                           .Sort = flexSortStringNoCaseAscending
                           mintsort = flexSortStringNoCaseAscending
                        Else
                           .Sort = flexSortStringNoCaseDescending
                           mintsort = flexSortStringNoCaseDescending
                        End If
                    End If
                Case 1716, 1717 '2,3,4列为数字，其他为字符
                    If InStr(1, "234", intCol) <> 0 Then
                        If intCol = mintPreCol And mintsort = flexSortNumericDescending Then
                           .Sort = flexSortNumericAscending
                           mintsort = flexSortNumericAscending
                        Else
                           .Sort = flexSortNumericDescending
                           mintsort = flexSortNumericDescending
                        End If
                    Else
                        If intCol = mintPreCol And mintsort = flexSortStringNoCaseDescending Then
                           .Sort = flexSortStringNoCaseAscending
                           mintsort = flexSortStringNoCaseAscending
                        Else
                           .Sort = flexSortStringNoCaseDescending
                           mintsort = flexSortStringNoCaseDescending
                        End If
                    End If
                Case 1719               '全为字符
                    If intCol = mintPreCol And mintsort = flexSortStringNoCaseDescending Then
                        .Sort = flexSortStringNoCaseAscending
                        mintsort = flexSortStringNoCaseAscending
                    Else
                       .Sort = flexSortStringNoCaseDescending
                       mintsort = flexSortStringNoCaseDescending
                    End If
                Case Else

            End Select
            mintPreCol = intCol
            .Row = grid.MshGrdFindRow(mshList, intTemp, intCol)
            If .RowPos(.Row) + .RowHeight(.Row) > .Height Then
                .TopRow = .Row
            Else
                .TopRow = 1
            End If
            .Col = 0
            .ColSel = .Cols - 1
            If .Row = 0 Then
                .Row = 1
            End If
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
            
            Select Case mlngMode
                Case 1712                   '6,8,9,10,11,12,13,16为数字，其他为字符
                    Select Case intCol
                        Case 6, 8, 9, 10, 11, 12, 13, 16
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
                        
                
                Case 1713, 1714, 1718       '6,8,9,10,11,12为数字，其他为字符
                    Select Case intCol
                        Case 6, 8, 9, 10, 11, 12
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
                Case 1715                   '7,8为数字，其他为字符
                    Select Case intCol
                        Case 7, 8
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
                Case 1716, 1717             '6,7,9,10,11,12,13为数字，其他为字符
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
                Case 1719                   '7,8,10,11,12,13为数字，其他为字符
                    Select Case intCol
                        Case 7, 8, 10, 11, 12, 13
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

Private Sub PrintRange(ByVal strRange As String)
    '功能:打印时间范围
    picSeparate_s.Cls
    picSeparate_s.CurrentX = 50
    picSeparate_s.CurrentY = 50
    picSeparate_s.Print strRange
End Sub
Private Function TestVerify(ByVal strNo As String) As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    '检查该单据是否通过审核，仅针对移库单
    gstrSQL = "" & _
        "   Select 审核人 From 药品收发记录 " & _
        "   Where 单据=19 And NO=[1] And Rownum<2"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查是否通过审核", strNo)
        
    If Not IsNull(rsTemp!审核人) Then
        TestVerify = True
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function TestPrepare(ByVal strNo As String) As Boolean
    Dim IntBill As Integer
    Dim rsTemp As New ADODB.Recordset
    '检查配药人是否已经填写
    On Error GoTo ErrHandle
    Select Case mlngMode
    Case 1712
        IntBill = 15
    Case 1716
        IntBill = 19
    Case Else
        Exit Function
    End Select
    
    gstrSQL = "Select 配药人 From 药品收发记录 Where 单据=[1] And NO=[2]  And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查是否通过核查", IntBill, strNo)
    If Not IsNull(rsTemp!配药人) Then
        TestPrepare = True
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub tabShow_Click(PreviousTab As Integer)
    If mlngMode <> 1716 Then Exit Sub
    Call SetMenu

    Call GetList(mstrFind)
End Sub


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

Private Sub LoadPlugInMnu(ByVal blnHave As Boolean)
'参数：blnHave true 表示插件对象存在
    Dim strTmp As String
    Dim arrTmp As Variant
    Dim i As Integer
    
    mnuPlugIn.Visible = blnHave
 
    If blnHave Then
        'blnHave 为true 时可以确保 gobjPlugIn 对象不为 Nothing
        On Error Resume Next
        strTmp = gobjPlugIn.GetFuncNames(glngSys, glngModul)
        If InStr(",438,0,", "," & err.Number & ",") = 0 Then
            MsgBox "zlPlugIn 外挂部件执行 GetFuncNames 时出错：" & vbCrLf & err.Number & vbCrLf & err.Description, vbInformation, gstrSysName
        End If
        err.Clear: On Error GoTo 0
        
        If strTmp = "" Then Exit Sub
        
        strTmp = Replace(strTmp, "Auto:", "")
        arrTmp = Split(strTmp, ",")
        For i = 0 To UBound(arrTmp)
            If i <> 0 Then
                Load mnuPlugItem(i)
            End If
            
            mnuPlugItem(i).Caption = CStr(arrTmp(i))
            mnuPlugItem(i).Tag = CStr(arrTmp(i))
            
            If i <= 9 Then
                mnuPlugItem(i).Caption = CStr(arrTmp(i)) & "(&" & IIf(i = 9, 0, i + 1) & ")"
            End If
        Next
    End If
End Sub

Private Sub ExcPlugInFun(ByVal strFunName As String)
    Dim lng库房ID As Long
    Dim int单据 As Integer
    Dim strNo As String
    
    With mshList
        lng库房ID = Val(cboStock.ItemData(cboStock.ListIndex))
        If mlngMode = 1712 Then int单据 = 15
        strNo = .TextMatrix(.Row, 0)
    End With
    
    On Error Resume Next
    
    If gobjPlugIn Is Nothing Then
        On Error Resume Next
        Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        If Not gobjPlugIn Is Nothing Then
            Call gobjPlugIn.Initialize(gcnOracle, glngSys, glngModul)
            If InStr(",438,0,", "," & err.Number & ",") = 0 Then
                MsgBox "zlPlugIn 外挂部件执行 Initialize 时出错：" & vbCrLf & err.Number & vbCrLf & err.Description, vbInformation, gstrSysName
            End If
        End If
        err.Clear: On Error GoTo 0
    End If
    
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        Call gobjPlugIn.DrugStuffWorkNoramal(mlngMode, strFunName, lng库房ID, strNo, int单据)
        If InStr(",438,0,", "," & err.Number & ",") = 0 Then
            MsgBox "zlPlugIn 外挂部件执行 ExecuteFunc 时出错：" & vbCrLf & err.Number & vbCrLf & err.Description, vbInformation, gstrSysName
        End If
        err.Clear: On Error GoTo 0
    End If
End Sub
