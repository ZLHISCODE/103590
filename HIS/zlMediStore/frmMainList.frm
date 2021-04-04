VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmMainList 
   ClientHeight    =   4980
   ClientLeft      =   2400
   ClientTop       =   4365
   ClientWidth     =   9480
   Icon            =   "frmMainList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4980
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picColor 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4680
      ScaleHeight     =   255
      ScaleWidth      =   3615
      TabIndex        =   14
      Top             =   4320
      Width           =   3615
      Begin VB.PictureBox picColor1 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   17
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
         TabIndex        =   16
         Top             =   0
         Width           =   260
      End
      Begin VB.PictureBox picColor3 
         BackColor       =   &H00FF00FF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2280
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   15
         Top             =   0
         Width           =   260
      End
      Begin VB.Label lblColor2 
         AutoSize        =   -1  'True
         Caption         =   "正常"
         Height          =   180
         Left            =   1680
         TabIndex        =   20
         Top             =   37
         Width           =   360
      End
      Begin VB.Label lblColor1 
         AutoSize        =   -1  'True
         Caption         =   "正常冲销"
         Height          =   180
         Left            =   360
         TabIndex        =   19
         Top             =   37
         Width           =   720
      End
      Begin VB.Label lblColor3 
         AutoSize        =   -1  'True
         Caption         =   "财务冲销"
         Height          =   180
         Left            =   2640
         TabIndex        =   18
         Top             =   30
         Width           =   720
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   1005
      Left            =   60
      TabIndex        =   11
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
      FormatString    =   $"frmMainList.frx":014A
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
   Begin TabDlg.SSTab TabShow 
      Height          =   330
      Left            =   60
      TabIndex        =   7
      Top             =   870
      Visible         =   0   'False
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   582
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "移出库房(&0)"
      TabPicture(0)   =   "frmMainList.frx":01BF
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "移入库房(&1)"
      TabPicture(1)   =   "frmMainList.frx":01DB
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
   End
   Begin VB.CommandButton Cmd查阅 
      Caption         =   "查阅(&V)"
      Height          =   350
      Left            =   5250
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2700
      Width           =   1100
   End
   Begin VB.PictureBox picSeparate_s 
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   120
      MousePointer    =   7  'Size N S
      ScaleHeight     =   360
      ScaleWidth      =   4815
      TabIndex        =   4
      Top             =   2550
      Width           =   4815
      Begin VB.Label lbl4 
         AutoSize        =   -1  'True
         Caption         =   "外调外销金额："
         Height          =   180
         Left            =   4560
         TabIndex        =   12
         Top             =   0
         Width           =   1260
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         Caption         =   "成本金额："
         Height          =   180
         Left            =   0
         TabIndex        =   10
         Top             =   20
         Width           =   900
      End
      Begin VB.Label lbl2 
         AutoSize        =   -1  'True
         Caption         =   "售价金额："
         Height          =   180
         Left            =   1890
         TabIndex        =   9
         Top             =   20
         Width           =   900
      End
      Begin VB.Label lbl3 
         AutoSize        =   -1  'True
         Caption         =   "差价金额："
         Height          =   180
         Left            =   3690
         TabIndex        =   8
         Top             =   20
         Width           =   900
      End
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "查询范围:1999年8月12日至1999年9月12日"
         Height          =   180
         Left            =   0
         TabIndex        =   5
         Top             =   200
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
         TabIndex        =   2
         Text            =   "cboStock"
         Top             =   240
         Width           =   3000
      End
      Begin MSComctlLib.Toolbar tlbTool 
         Height          =   720
         Left            =   165
         TabIndex        =   1
         Top             =   30
         Width           =   7875
         _ExtentX        =   13891
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
            NumButtons      =   24
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
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Object.Visible         =   0   'False
                     Key             =   "FromStore"
                     Text            =   "向库房领药"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Object.Visible         =   0   'False
                     Key             =   "FromLeave"
                     Text            =   "向留存领药"
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
               Key             =   "EditSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "核查"
               Key             =   "Prepare"
               Object.ToolTipText     =   "核查"
               Object.Tag             =   "核查"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "备药"
               Key             =   "PreparePhysic"
               Object.ToolTipText     =   "备药"
               Object.Tag             =   "备药"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "发送"
               Key             =   "SendPhysic"
               Object.ToolTipText     =   "发送"
               Object.Tag             =   "发送"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "回退"
               Key             =   "Back"
               Object.ToolTipText     =   "回退到上次状态"
               Object.Tag             =   "回退"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "PrepareSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "审核"
               Key             =   "Verify"
               Description     =   "审核"
               Object.ToolTipText     =   "审核"
               Object.Tag             =   "审核"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "冲销"
               Key             =   "Strike"
               Description     =   "冲销"
               Object.ToolTipText     =   "冲销"
               Object.Tag             =   "冲销"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "申请冲销"
               Key             =   "ApplyStrike"
               Description     =   "申请冲销"
               Object.ToolTipText     =   "申请冲销"
               Object.Tag             =   "申请冲销"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "审核冲销"
               Key             =   "VerifyStrike"
               Description     =   "审核冲销"
               Object.ToolTipText     =   "审核冲销"
               Object.Tag             =   "审核冲销"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "VerifySeparate"
               Style           =   3
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "过滤"
               Key             =   "Search"
               Description     =   "过滤"
               Object.ToolTipText     =   "过滤"
               Object.Tag             =   "过滤"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "刷新"
               Key             =   "Refresh"
               Description     =   "刷新"
               Object.ToolTipText     =   "刷新"
               Object.Tag             =   "刷新"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "PlugInSeparator"
               Style           =   3
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "功能1"
               Key             =   "PlugItem"
               ImageIndex      =   17
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FindSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助主题"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageIndex      =   14
            EndProperty
            BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Exit"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   15
            EndProperty
         EndProperty
         MouseIcon       =   "frmMainList.frx":01F7
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
            Picture         =   "frmMainList.frx":0511
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
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":0DA5
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":0FC5
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":11E5
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":1401
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":1621
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":1841
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":1F3B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":292D
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":331F
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":3539
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":3755
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":3971
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":3B8B
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":3CE5
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":3F01
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":4121
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":A983
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
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":B85D
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":BA7D
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":BC9D
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":BEB9
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":C0D9
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":C2F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":C9F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":D3E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":DDD7
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":DFF1
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":E20D
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":E429
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":E643
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":E79D
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":E9BD
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":EBDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainList.frx":1543F
            Key             =   "PlugIn"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
      Height          =   975
      Left            =   0
      TabIndex        =   13
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
      FormatString    =   $"frmMainList.frx":17BF1
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
      Begin VB.Menu mnuFileCodePrint 
         Caption         =   "条码打印(&C)"
         Begin VB.Menu mnuFileAllCodePrint 
            Caption         =   "单据中药品条码打印(&A)"
         End
         Begin VB.Menu mnuFileSelCodePrint 
            Caption         =   "选中行药品条码打印(&S)"
         End
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
      Begin VB.Menu mnuEditPrepare 
         Caption         =   "核查(&H)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditPreparePhysic 
         Caption         =   "备药(&P)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditSendPhysic 
         Caption         =   "发送(&S)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditBack 
         Caption         =   "回退(&O)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditLine3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditVerify 
         Caption         =   "审核(&C)"
      End
      Begin VB.Menu mnuEditMark 
         Caption         =   "发票核对(&M)"
      End
      Begin VB.Menu mnuEditStrike 
         Caption         =   "冲销(&K)"
      End
      Begin VB.Menu mnuEditApplyStrike 
         Caption         =   "申请冲销"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditVerifyStrike 
         Caption         =   "审核冲销"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditWriteOff 
         Caption         =   "批量冲销"
      End
      Begin VB.Menu mnuEditLine0 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditRestore 
         Caption         =   "药库退货(&R)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditBill 
         Caption         =   "修改发票信息(&B)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditAcc 
         Caption         =   "财务审核(&V)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditPric 
         Caption         =   "成本价调价(&Z)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditHandBack 
         Caption         =   "药品退药计划(&P)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditMediPlanImport 
         Caption         =   "药品计划单批量导入(&I)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditDeliveryInvoice 
         Caption         =   "送货发票导入(&E)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditVerifySelect 
         Caption         =   "财务审核单据查询(&Y)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditDisplay 
         Caption         =   "查看单据(&W)"
      End
      Begin VB.Menu mnuEditCodePrintLine 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditAllCodePrint 
         Caption         =   "单据中药品条码打印(&A)"
         Visible         =   0   'False
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
Private mblnStock As Boolean            '当前操作员是否是药库人员，仅对领用单据有效
Private mblnBootUp As Boolean
Private mlastRow As Long                '上次点击的行
Private mstrTitle As String             '窗体的标题
Private mbln核查 As Boolean
Private mstrPrivs As String                     '权限
Private mintListRow As Integer
Private mStr库房 As String
Private mblnDo As Boolean
Private mbln操作员限制  As Boolean
Private mblnViewCost As Boolean      '查看成本价 true-可以查看 false-不可以查看

Private mstrNumberFormat As String
Private mstrCostFormat As String
Private mstrPriceFormat As String
Private mstrMoneyFormat As String

Private mlng库房ID As Long
Private mintUnit As Integer                 '单位系数：1-售价;2-门诊;3-住院;4-药库
Private mblnBandEvent As Boolean            '记录是否对vsf控件调用了datasource方法，true-是，false-否

Private mobjPlugIn As Object             '外挂接口

'从参数表中取药品价格、数量、金额小数位数（显示精度）
Private mintShowCostDigit As Integer            '成本价小数位数
Private mintShowPriceDigit As Integer           '售价小数位数
Private mintShowNumberDigit As Integer          '数量小数位数
Private mintShowMoneyDigit As Integer           '金额小数位数

Private Const mconint售价单位 As Integer = 1
Private Const mconint门诊单位 As Integer = 2
Private Const mconint住院单位 As Integer = 3
Private Const mconint药库单位 As Integer = 4

'日期设置
Private strStart As Date
Private strEnd As Date
Private strVerifyStart As Date
Private strVerifyEnd As Date

Private int移库处理流程 As Integer                    '1-需要备药、发送、接收这一过程  0-不需要这一过程
Private mint冲销申请 As Integer                       '0-不需要申请;1-需要申请
Private mint领用冲销申请 As Integer                   '药品领用模块冲销方式：0-不需要申请;1-需要申请
Private mint库存检查 As Integer                       '表示药品出库时是否进行库存检查：0-不检查;1-检查，不足提醒；2-检查，不足禁止

Private Type Type_SQLCondition
    strNO开始 As String
    strNO结束 As String
    date填制时间开始 As Date
    date填制时间结束 As Date
    date审核时间开始 As Date
    date审核时间结束 As Date
    lng药品 As Long
    lng库房 As Long
    str填制人 As String
    str审核人 As String
    lng生产商 As Long
    str产地 As String
    str发票号开始 As String
    str发票号结束 As String
    lng入出类别 As Long
    int填制审核一并查询 As Integer
    int无标记 As Integer
    int有标记 As Integer
    int无发票 As Integer
    int有发票 As Integer
    lng药品分类 As Long
    str剂型 As String
    date发票审核日期开始 As Date
    date发票审核日期结束 As Date
End Type

Private SQLCondition As Type_SQLCondition

Private mstr屏蔽列 As String

Private Enum 外购主表
    NO = 0
    供应商 = 1
    成本金额 = 2
    售价金额 = 3
    差价金额 = 4
    零售金额 = 5
    零售差价 = 6
    冲销类型 = 7
    填制人 = 8
    填制日期 = 9
    修改人 = 10
    修改日期 = 11
    核查人 = 12
    核查日期 = 13
    审核人 = 14
    审核日期 = 15
    记录状态 = 16
    单据说明 = 17
    摘要 = 18
    
    列数 = 19
End Enum

Private Function Is申领(ByVal strBillNo As String) As Boolean
    Dim rsCheck As New ADODB.Recordset
    '先检查是不是申领单
    On Error GoTo errHandle
    gstrSQL = " Select Nvl(发药方式,0) 申领 From 药品收发记录 " & _
              " Where 单据=6 And NO=[1] And 入出系数 = -1 and rownum = 1"
    Set rsCheck = zlDataBase.OpenSQLRecord(gstrSQL, "检查是不是申领单", strBillNo)
    
    Is申领 = Not (rsCheck!申领 = 0)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub PlugInFun(ByVal strFunName As String)
    '执行外挂功能
    Dim strParam As String
    Dim lng库房ID As Long
    Dim int单据 As Integer
    Dim strNo As String
    
    If mlngMode <> 模块号.外购入库 Then Exit Sub
    
    With vsfList
        If .TextMatrix(.Row, 0) <> "" Then
            lng库房ID = Val(cboStock.ItemData(cboStock.ListIndex))
            If mlngMode = 模块号.外购入库 Then int单据 = 1
            strNo = .TextMatrix(.Row, 0)
            
            strParam = lng库房ID & "," & int单据 & "," & strNo
        End If
    End With
    
    Call zlPlugIn_Fun(glngSys, mlngMode, mobjPlugIn, Me, strFunName, strParam)
End Sub

Private Sub SetDetailFocus()
    vsfDetail.ForeColorFixed = glngFixedForeColorByFocus
    vsfDetail.BackColorSel = glngRowByFocus
'    If vsfDetail.Row > 0 Then
'        vsfDetail.ForeColorSel = vsfDetail.Cell(flexcpForeColor, vsfDetail.Row)
'    End If
    
    vsfList.ForeColorFixed = glngFixedForeColorNotFocus
    vsfList.BackColorSel = glngRowByNotFocus
End Sub

Public Sub SetMenu()
    '隐藏备药、发送、审核与冲销
    mnuEditPreparePhysic.Visible = False
    mnuEditSendPhysic.Visible = False
    mnuEditBack.Visible = False
    tlbTool.Buttons("PreparePhysic").Visible = False
    tlbTool.Buttons("SendPhysic").Visible = False
    tlbTool.Buttons("Back").Visible = False
    mnuEditVerify.Visible = False
    mnuEditStrike.Visible = False
    mnuEditWriteOff.Visible = False
    tlbTool.Buttons("Verify").Visible = False
    tlbTool.Buttons("Strike").Visible = False
    mnuEditLine3.Visible = False
    mnuEditLine0.Visible = False
    tlbTool.Buttons("PrepareSeparate").Visible = False
    tlbTool.Buttons("VerifySeparate").Visible = False
    
    '根据当前页面开启
    If TabShow.Tab = 0 Then
        If mlngMode = 模块号.药品移库 Then
            If int移库处理流程 = 0 Then
                mnuEditPreparePhysic.Visible = False
                mnuEditVerify.Visible = zlStr.IsHavePrivs(mstrPrivs, "审核")
'                mnuEditStrike.Visible = zlStr.IsHavePrivs(mstrPrivs, "冲销")
                mnuEditWriteOff.Visible = False
                mnuEditStrike.Visible = False
                mnuEditLine0.Visible = True
                tlbTool.Buttons("Verify").Visible = mnuEditVerify.Visible
                tlbTool.Buttons("Strike").Visible = mnuEditStrike.Visible
                tlbTool.Buttons("VerifySeparate").Visible = True
                mnuEditVerify.Caption = "审核(&C)"
                tlbTool.Buttons("Verify").Caption = "审核"
                tlbTool.Buttons("Verify").Tag = "审核"
                tlbTool.Buttons("Verify").ToolTipText = "审核"
            Else
                mnuEditVerify.Caption = "接收(&C)"
                tlbTool.Buttons("Verify").Caption = "接收"
                tlbTool.Buttons("Verify").Tag = "接收"
                tlbTool.Buttons("Verify").ToolTipText = "接收"
                mnuEditPreparePhysic.Visible = zlStr.IsHavePrivs(mstrPrivs, "发送")
            End If
            If mint冲销申请 = 1 Then
                mnuEditStrike.Visible = True
                mnuEditWriteOff.Visible = False
                mnuEditStrike.Caption = "审核冲销(&K)"
                tlbTool.Buttons("Strike").Visible = True
                tlbTool.Buttons("Strike").Caption = IIf(mnuViewToolText.Checked = False, "", "审核冲销")
                tlbTool.Buttons("Strike").Tag = "审核冲销"
                tlbTool.Buttons("Strike").ToolTipText = "审核冲销"
                mnuEditLine0.Visible = True
                tlbTool.Buttons("VerifySeparate").Visible = True
            End If
                        
            mnuEditSendPhysic.Visible = mnuEditPreparePhysic.Visible
            mnuEditBack.Visible = mnuEditPreparePhysic.Visible
            mnuEditLine3.Visible = mnuEditPreparePhysic.Visible
            tlbTool.Buttons("PreparePhysic").Visible = mnuEditPreparePhysic.Visible
            tlbTool.Buttons("SendPhysic").Visible = mnuEditPreparePhysic.Visible
            tlbTool.Buttons("Back").Visible = mnuEditPreparePhysic.Visible
            tlbTool.Buttons("PrepareSeparate").Visible = mnuEditPreparePhysic.Visible
        Else
            If mlngMode = 模块号.药品领用 Then
                tlbTool.Buttons("Add").Style = tbrDropdown
                tlbTool.Buttons("Add").ButtonMenus(1).Visible = True
                tlbTool.Buttons("Add").ButtonMenus(2).Visible = True
    
                If mblnStock Then
                    mnuEditVerify.Visible = zlStr.IsHavePrivs(mstrPrivs, "审核")
                    mnuEditStrike.Visible = zlStr.IsHavePrivs(mstrPrivs, "冲销")
                    mnuEditWriteOff.Visible = zlStr.IsHavePrivs(mstrPrivs, "冲销")
                    If mint领用冲销申请 = 1 Then
                        mnuEditStrike.Visible = False
                        mnuEditApplyStrike.Visible = zlStr.IsHavePrivs(mstrPrivs, "冲销申请")
                        mnuEditVerifyStrike.Visible = zlStr.IsHavePrivs(mstrPrivs, "冲销审核")
                    End If
                End If
            Else
                mnuEditVerify.Visible = zlStr.IsHavePrivs(mstrPrivs, "审核")
                mnuEditStrike.Visible = zlStr.IsHavePrivs(mstrPrivs, "冲销")
                mnuEditWriteOff.Visible = zlStr.IsHavePrivs(mstrPrivs, "冲销")
            End If
            
            If mlngMode = 模块号.外购入库 Then
                mnuEditLine0.Visible = True
                If zlStr.IsHavePrivs(mstrPrivs, "核查成本价") And mbln核查 Then
                    mnuEditPrepare.Visible = True
                    mnuEditBack.Visible = True
                    tlbTool.Buttons("Prepare").Visible = True
                    tlbTool.Buttons("Back").Visible = True
                    tlbTool.Buttons("PrepareSeparate").Visible = True
                End If
                
                If zlStr.IsHavePrivs(mstrPrivs, "付款标记") And gtype_UserSysParms.P173_经过标记付款后才能进行付款管理 = 1 Then
                    mnuEditMark.Visible = True
                Else
                    mnuEditMark.Visible = False
                End If
            End If
            tlbTool.Buttons("Verify").Visible = mnuEditVerify.Visible
            tlbTool.Buttons("Strike").Visible = mnuEditStrike.Visible
            tlbTool.Buttons("ApplyStrike").Visible = mnuEditApplyStrike.Visible
            tlbTool.Buttons("VerifyStrike").Visible = mnuEditVerifyStrike.Visible
            tlbTool.Buttons("VerifySeparate").Visible = (mnuEditVerify.Visible Or mnuEditStrike.Visible)
        End If
    Else
        mnuEditVerify.Visible = zlStr.IsHavePrivs(mstrPrivs, "审核")
        mnuEditStrike.Visible = zlStr.IsHavePrivs(mstrPrivs, "冲销")
        mnuEditWriteOff.Visible = zlStr.IsHavePrivs(mstrPrivs, "冲销")
        
        If mlngMode = 模块号.药品移库 Then
            If mint冲销申请 = 1 Then
                mnuEditStrike.Caption = "申请冲销(&R)"
                tlbTool.Buttons("Strike").Caption = IIf(mnuViewToolText.Checked = False, "", "申请冲销")
                tlbTool.Buttons("Strike").Tag = "申请冲销"
                tlbTool.Buttons("Strike").ToolTipText = "申请冲销"
            Else
                mnuEditStrike.Caption = "冲销(&K)"
                tlbTool.Buttons("Strike").Caption = IIf(mnuViewToolText.Checked = False, "", "冲销")
                tlbTool.Buttons("Strike").ToolTipText = "冲销"
            End If
        End If
        
        If mlngMode = 模块号.外购入库 Then
            mnuEditLine0.Visible = True
        End If
        tlbTool.Buttons("Verify").Visible = mnuEditVerify.Visible
        tlbTool.Buttons("Strike").Visible = mnuEditStrike.Visible
        tlbTool.Buttons("VerifySeparate").Visible = True
    End If
    
    If mlngMode = 模块号.自制入库 Or mlngMode = 模块号.差价调整 Then
        mnuEditWriteOff.Visible = False
    End If
End Sub

Private Sub cboStock_Click()

    Dim lng库房ID As Long
    Dim rsCheck As New ADODB.Recordset
    
    On Error GoTo errHandle
    lng库房ID = cboStock.ItemData(cboStock.ListIndex)

    If mlng库房ID <> lng库房ID Then
        mlng库房ID = Me.cboStock.ItemData(Me.cboStock.ListIndex)
        Call GetDrugDigit(mlng库房ID, Me.Tag, mintUnit, mintShowCostDigit, mintShowPriceDigit, mintShowNumberDigit, mintShowMoneyDigit)
        
        '重新组织格式化串
        mstrCostFormat = "'999999999990." & String(mintShowCostDigit, "0") & "'"
        mstrPriceFormat = "'999999999990." & String(mintShowPriceDigit, "0") & "'"
        mstrNumberFormat = "'999999999990." & String(mintShowNumberDigit, "0") & "'"
        mstrMoneyFormat = "'999999999990." & String(mintShowMoneyDigit, "0") & "'"
    
        '检查该库房是否为药库，只有药库才允许退货
        gstrSQL = " SELECT DISTINCT 0 " & _
                  " FROM 部门性质说明 " & _
                  " WHERE 工作性质 LIKE '%药库' " & _
                  " AND 部门ID = [1]"
        Set rsCheck = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[检查当前库房是否为药库]", lng库房ID)
                  
        mnuEditRestore.Enabled = (rsCheck.RecordCount > 0)
'        mnuEditLine0.Enabled = (rsCheck.RecordCount = 0)
        
        If mblnBootUp Then mnuViewRefresh_Click
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cboStock_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim str工作性质 As String
    '获取可操作的库房
    Select Case mlngMode
        Case 模块号.外购入库
            If InStr(1, mstrPrivs, "允许药房外购入库") = 0 Then
                str工作性质 = "H,I,J"
            Else
                str工作性质 = "H,I,J,K,L,M,N"
            End If
        Case 模块号.自制入库
            str工作性质 = "H,I,J,K,L,M,N"
        Case 模块号.其他入库
            str工作性质 = "H,I,J,K,L,M,N"
        Case 模块号.差价调整
            str工作性质 = "H,I,J,K,L,M,N"
        Case 模块号.药品移库
            str工作性质 = "H,I,J,K,L,M,N"
        Case 模块号.药品领用
            str工作性质 = "H,I,J,K,L,M,N"
        Case 模块号.其他出库
            str工作性质 = "H,I,J,K,L,M,N"
        Case Else
    End Select
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cboStock.ListCount = 0 Then Call zlControl.ControlSetFocus(vsfList): Exit Sub
    
    If cboStock.ListIndex >= 0 Then
        If Val(cboStock.Tag) = cboStock.ItemData(cboStock.ListIndex) Then
            Call zlControl.ControlSetFocus(vsfList, True)
            Exit Sub
        End If
    End If
    
    If Select部门选择器(Me, cboStock, Trim(cboStock.Text), str工作性质, mbln操作员限制) = False Then
        Exit Sub
    End If
    If cboStock.ListIndex >= 0 Then
        cboStock.Tag = cboStock.ItemData(cboStock.ListIndex)
    End If
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
    If mblnBootUp = False Then Exit Sub
    Form_Resize
End Sub

Public Sub ShowList(ByVal lngMode As Long, ByVal strTitle As String, ByVal FrmMain As Variant)
    Dim strFind As String
    Dim dateCurrentDate As Date
    Dim strTemp As String
    Dim int查询天数 As Integer
    
    mblnBootUp = False
    mlngMode = lngMode
    mstrTitle = strTitle
    mstrPrivs = gstrprivs
    Me.Tag = strTitle
    
    int移库处理流程 = Val(zlDataBase.GetPara("移库流程", glngSys, 模块号.药品移库))
    mint冲销申请 = Val(zlDataBase.GetPara("冲销申请", glngSys, 模块号.药品移库))
    mint领用冲销申请 = Val(zlDataBase.GetPara("冲销申请", glngSys, 模块号.药品领用))
    
    If mlngMode = 模块号.外购入库 Or mlngMode = 模块号.其他入库 Then
        If zlDataBase.GetPara("选择列", glngSys, mlngMode) = "" Then
            mstr屏蔽列 = "零售价|零售单位|零售金额|零售差价"
        Else
            mstr屏蔽列 = zlDataBase.GetPara("屏蔽列", glngSys, mlngMode)
        End If
    End If
    
    '数据依赖性测试
    If Not CheckDepend Then
        Unload Me
        Exit Sub
    End If
    
    '实例化采购平台接口
    If mlngMode = 模块号.外购入库 Then
        On Error Resume Next
        If gobjDrugPurchase Is Nothing Then
            Set gobjDrugPurchase = CreateObject("zlDrugPurchase.clsDrugPurchase")
        End If
        Err.Clear
        On Error GoTo 0
        If Not gobjDrugPurchase Is Nothing Then
            mnuEditDeliveryInvoice.Visible = True
        End If
    End If
    
    mlng库房ID = Me.cboStock.ItemData(Me.cboStock.ListIndex)
    Call GetDrugDigit(mlng库房ID, Me.Tag, mintUnit, mintShowCostDigit, mintShowPriceDigit, mintShowNumberDigit, mintShowMoneyDigit)
    
    '组织格式化串
    mstrCostFormat = "'999999999990." & String(mintShowCostDigit, "0") & "'"
    mstrPriceFormat = "'999999999990." & String(mintShowPriceDigit, "0") & "'"
    mstrNumberFormat = "'999999999990." & String(mintShowNumberDigit, "0") & "'"
    mstrMoneyFormat = "'999999999990." & String(mintShowMoneyDigit, "0") & "'"
    
    SetVisable  '根据权限设置不同的显示项目
        
    dateCurrentDate = Sys.Currentdate

    int查询天数 = Val(zlDataBase.GetPara("查询天数", glngSys, mlngMode, 7)) - 1
    strStart = Format(DateAdd("d", -int查询天数, dateCurrentDate), "yyyy-MM-dd")
    strEnd = Format(dateCurrentDate, "yyyy-MM-dd")
    
    strVerifyStart = "1901-01-01"
    strVerifyEnd = "1901-01-01"
    
    strFind = " AND A.记录状态 = 1 And A.审核日期 is Null And A.填制日期 Between [3] And [4] "
    SQLCondition.date填制时间开始 = CDate(Format(strStart, "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date填制时间结束 = CDate(Format(strEnd, "yyyy-mm-dd") & " 23:59:59")
    
    mstrFind = strFind
    
    Call TabShow_Click(0)
    If mlngMode <> 模块号.药品移库 Then GetList (mstrFind) '列出单据头
    
    '恢复用户个性化设置
    RestoreWinState Me, App.ProductName, mstrTitle
    
    If mlngMode = 模块号.差价调整 Then
        vsfList.ColWidth(vsfList.Cols - 1) = 0
        vsfList.ColWidth(vsfList.Cols - 3) = 0
    End If
    If mlngMode = 模块号.外购入库 Then
        vsfList.ColWidth(外购主表.记录状态) = 0
        vsfList.ColWidth(外购主表.单据说明) = 1000
    End If
    If mlngMode = 模块号.药品领用 Then
        vsfList.ColWidth(vsfList.Cols - 4) = 0
        vsfList.ColWidth(vsfList.Cols - 3) = 1000
    End If
    '用户个性化设置后，重新设置权限控制的列是否显示
    If mblnViewCost = False Then
        With vsfList
            Select Case mlngMode
                Case 模块号.外购入库
                    .colHidden(外购主表.成本金额) = True
                    .colHidden(外购主表.差价金额) = True
                Case 模块号.自制入库
                    .colHidden(.ColIndex("成本金额")) = True
                    .colHidden(.ColIndex("差价金额")) = True
                Case 模块号.其他入库
                    .colHidden(.ColIndex("成本金额")) = True
                    .colHidden(.ColIndex("差价金额")) = True
                Case 模块号.药品移库
                    .colHidden(.ColIndex("成本金额")) = True
                    .colHidden(.ColIndex("差价金额")) = True
                Case 模块号.药品领用
                    .colHidden(.ColIndex("成本金额")) = True
                    .colHidden(.ColIndex("差价金额")) = True
                Case 模块号.其他出库
                    .colHidden(.ColIndex("成本金额")) = True
                    .colHidden(.ColIndex("差价金额")) = True
            End Select
        End With
        
        With vsfDetail
            Select Case mlngMode
                Case 模块号.外购入库
                    .colHidden(.ColIndex("成本价")) = True
                    .colHidden(.ColIndex("成本金额")) = True
                    .colHidden(.ColIndex("差价")) = True
                Case 模块号.其他入库
                    .colHidden(.ColIndex("成本价")) = True
                    .colHidden(.ColIndex("成本金额")) = True
                    .colHidden(.ColIndex("差价")) = True
                Case 模块号.自制入库
                    .colHidden(.ColIndex("采购价")) = True
                    .colHidden(.ColIndex("采购金额")) = True
                    .colHidden(.ColIndex("差价")) = True
                Case 模块号.药品移库
                    .colHidden(.ColIndex("成本价")) = True
                    .colHidden(.ColIndex("成本金额")) = True
                    .colHidden(.ColIndex("差价")) = True
                Case 模块号.药品领用
                    .colHidden(.ColIndex("成本价")) = True
                    .colHidden(.ColIndex("成本金额")) = True
                    .colHidden(.ColIndex("差价")) = True
                Case 模块号.其他出库
                    .colHidden(.ColIndex("成本价")) = True
                    .colHidden(.ColIndex("成本金额")) = True
                    .colHidden(.ColIndex("差价")) = True
            End Select
        End With
    End If
            
    Call zlDataBase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs)
    
    mblnBootUp = True
    
    If IsObject(FrmMain) Then
        Me.Show , FrmMain
    Else
        OS.ShowChildWindow Me.hWnd, FrmMain
    End If
    
    Me.ZOrder 0
End Sub

'检查数据依赖性
Private Function CheckDepend() As Boolean
    Dim rsDepend As New ADODB.Recordset
    Dim strStock As String, strCaption As String
    
    CheckDepend = False
    On Error GoTo errHandle
    
    '获取可操作的库房
    Select Case mlngMode
        Case 模块号.外购入库
            If InStr(1, mstrPrivs, "允许药房外购入库") = 0 Then
                strStock = "HIJ"
            Else
                strStock = "HIJKLMN"
            End If
        Case 模块号.自制入库
            strStock = "HIJKLMN"
        Case 模块号.其他入库
            strStock = "HIJKLMN"
        Case 模块号.差价调整
            strStock = "HIJKLMN"
        Case 模块号.药品移库
            strStock = "HIJKLMN"
        Case 模块号.药品领用
            strStock = "HIJKLMN"
        Case 模块号.其他出库
            strStock = "HIJKLMN"
        Case Else
    End Select
    
    '如果是药品领用，则检查当前科室是否是领用部门，且允许向库房领药
    If mlngMode <> 模块号.药品领用 Then
        gstrSQL = "SELECT DISTINCT a.id, a.名称 " _
                & "FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " _
                & "Where (a.站点 = [3] Or a.站点 is Null) And c.工作性质 = b.名称 " _
                & "  AND Instr([2],b.编码,1) > 0 " _
                & "  AND a.id = c.部门id " _
                & "  AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01'" _
                & IIf(zlStr.IsHavePrivs(gstrprivs, "所有库房"), "", " And a.ID IN (Select 部门ID From 部门人员 Where 人员ID=[1])")
        Set rsDepend = zlDataBase.OpenSQLRecord(gstrSQL, mstrTitle, UserInfo.用户ID, strStock, gstrNodeNo)
        
    Else
        '先判断是不是药库人员使用本模块
        gstrSQL = "SELECT DISTINCT a.id, a.名称 " _
                & "FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " _
                & "Where (a.站点 = [3] Or a.站点 is Null) And c.工作性质 = b.名称 " _
                & "  AND Instr([2],b.编码,1) > 0 " _
                & "  AND a.id = c.部门id " _
                & "  AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01'" _
                & "  And a.ID IN (Select 部门ID From 部门人员 Where 人员ID=[1])"
        Set rsDepend = zlDataBase.OpenSQLRecord(gstrSQL, mstrTitle, UserInfo.用户ID, strStock, gstrNodeNo)
                  
        mblnStock = (rsDepend.RecordCount <> 0)
        
        If mblnStock Then
            '根据权限所有库房提取库房数据
            gstrSQL = "SELECT DISTINCT a.id, a.名称 " _
                    & "FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " _
                    & "Where (a.站点 = [3] Or a.站点 is Null) And c.工作性质 = b.名称 " _
                    & "  AND Instr([2],b.编码,1) > 0 " _
                    & "  AND a.id = c.部门id " _
                    & "  AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01'" _
                    & IIf(zlStr.IsHavePrivs(gstrprivs, "所有库房"), "", " And a.ID IN (Select 部门ID From 部门人员 Where 人员ID=[1])")
            Set rsDepend = zlDataBase.OpenSQLRecord(gstrSQL, mstrTitle, UserInfo.用户ID, strStock, gstrNodeNo)
        Else
            '提取该人员所属的领用部门
            gstrSQL = " Select C.ID " & _
                      " From 部门性质说明 A,部门性质分类 B,部门表 C " & _
                      " Where (c.站点 = [3] Or c.站点 is Null) And A.工作性质=B.名称 And A.部门ID=C.ID " & _
                      "   AND TO_CHAR(C.撤档时间, 'yyyy-MM-dd')='3000-01-01' And B.编码='O'" & _
                      "   And C.ID IN (Select 部门ID From 部门人员 Where 人员ID=[1])"
            Set rsDepend = zlDataBase.OpenSQLRecord(gstrSQL, mstrTitle, UserInfo.用户ID, "O", gstrNodeNo)
                      
            If rsDepend.RecordCount = 0 Then
                MsgBox "你不是领药部门的操作人员，不能使用本模块！[部门管理]", vbInformation, gstrSysName
                Exit Function
            End If
            
            '再根据药品领药控制，提取这些领药部门允许领用库房的数据
            gstrSQL = "SELECT DISTINCT a.id, a.名称 " _
                    & "FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " _
                    & "Where (a.站点 = [3] Or a.站点 is Null) And c.工作性质 = b.名称 " _
                    & "  AND Instr([2],b.编码,1) > 0 " _
                    & "  AND a.id = c.部门id " _
                    & "  AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01'" _
                    & "  And a.ID IN (Select 对方库房ID From 药品领用控制 Where 领用部门ID IN (" & gstrSQL & "))"
            Set rsDepend = zlDataBase.OpenSQLRecord(gstrSQL, mstrTitle, UserInfo.用户ID, strStock, gstrNodeNo)
        End If
    End If
        
    If rsDepend.EOF Then
        If mlngMode <> 模块号.药品领用 Or mblnStock Then
            If mlngMode = 模块号.外购入库 And InStr(1, mstrPrivs, "允许药房外购入库") = 0 Then
                MsgBox "该人员无“允许药房外购入库”权限，请与管理员联系！", vbInformation, gstrSysName
            Else
                MsgBox "至少应该设置一个具有药库性质，药房性质，或者制剂室性质的部门,请查看部门管理！", vbInformation, gstrSysName
            End If
        Else
            MsgBox "你没有权限向任何库房领用药品，请设置领用流向！[基础参数设置]", vbInformation, gstrSysName
        End If
        If rsDepend.State = 1 Then rsDepend.Close
        Exit Function
    End If
    
    '装入库房数据
    With cboStock
        .Clear
        Do While Not rsDepend.EOF
            .AddItem rsDepend!名称
            .ItemData(.NewIndex) = rsDepend!id
            mStr库房 = mStr库房 & rsDepend!id & "," & rsDepend!名称 & "|"
            If mlngMode <> 模块号.药品领用 Or mblnStock Then
                If rsDepend!id = UserInfo.部门ID Then
                    .ListIndex = .NewIndex
                End If
            End If
            rsDepend.MoveNext
        Loop
        If .ListIndex = -1 Then .ListIndex = 0
        rsDepend.Close
    End With
    
    '检查是否需要核查环节（仅针对外购入库）
    strCaption = "核查"
    If mlngMode = 模块号.外购入库 Then
        mbln核查 = (gtype_UserSysParms.P75_外购入库需要核查 = 1)
                
        mnuEditPrepare.Caption = strCaption & "(&H)"
        tlbTool.Buttons("Prepare").Caption = strCaption
        tlbTool.Buttons("Prepare").Tag = strCaption
        tlbTool.Buttons("Prepare").ToolTipText = strCaption
    ElseIf mlngMode = 模块号.药品移库 Then
        If int移库处理流程 = 0 Then
            mnuEditVerify.Caption = "审核(&C)"
            tlbTool.Buttons("Verify").Caption = "审核"
            tlbTool.Buttons("Verify").Tag = "审核"
            tlbTool.Buttons("Verify").ToolTipText = "审核"
        Else
            mnuEditVerify.Caption = "接收(&C)"
            tlbTool.Buttons("Verify").Caption = "接收"
            tlbTool.Buttons("Verify").Tag = "接收"
            tlbTool.Buttons("Verify").ToolTipText = "接收"
        End If
        TabShow.Visible = True
    End If
    
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
    Dim strTransfer As String, strTransfer_Order As String
    Dim strsql As String
    Dim strSqlForm As String
    
    '用于统计合计金额
    Dim dbl1 As Double
    Dim dbl2 As Double
    Dim dbl3 As Double
    Dim dbl4 As Double
    Dim n As Long
    Dim StrFormat As String
    
    StrFormat = "0.00##"
    On Error GoTo errHandle
    mlastRow = 0
    Call FS.ShowFlash("正在搜索药品记录,请稍候 ...", Me)
    DoEvents
    Screen.MousePointer = vbHourglass
    
    If strFind = "" Then
        strFind = " AND A.记录状态 = 1 And A.审核日期 is Null And A.填制日期 Between [3] And [4] "
        SQLCondition.date填制时间开始 = CDate(Format(strStart, "yyyy-mm-dd") & " 00:00:00")
        SQLCondition.date填制时间结束 = CDate(Format(strEnd, "yyyy-mm-dd") & " 23:59:59")
    End If
    
    Select Case mlngMode
    Case 模块号.药品移库
        If TabShow.Tab = 0 Then
            strUserPart = " And A.库房ID+0=[16] "
            strTransfer = " A.审核人 AS 接收人,TO_CHAR(MIN(A.审核日期),'YYYY-MM-DD HH24:MI:SS') AS 接收日期,A.配药人 AS 备药人,TO_CHAR(MIN(A.配药日期),'YYYY-MM-DD HH24:MI:SS') AS 发送日期"
            strTransfer_Order = ",A.审核人,A.配药人"
        Else
            strUserPart = " And A.对方部门ID+0=[16] "
            strTransfer = " A.配药人 AS 备药人,TO_CHAR(MIN(A.配药日期),'YYYY-MM-DD HH24:MI:SS') AS 发送日期,A.审核人 AS 接收人,TO_CHAR(MIN(A.审核日期),'YYYY-MM-DD HH24:MI:SS') AS 接收日期"
            strTransfer_Order = ",A.配药人,A.审核人"
        End If
    Case 模块号.药品领用
        If mblnStock Then
            strUserPart = " And A.库房ID+0=[16] "
        Else
            strUserPart = " Select C.ID " & _
                      " From 部门性质说明 A,部门性质分类 B,部门表 C " & _
                      " Where (c.站点 = [18] Or c.站点 is Null) And A.工作性质=B.名称 And A.部门ID=C.ID " & _
                      " AND TO_CHAR(C.撤档时间, 'yyyy-MM-dd')='3000-01-01' And B.编码='O'" & _
                      " And C.ID IN (Select 部门ID From 部门人员 Where 人员ID=[17])"

            strUserPart = " And A.库房ID+0=[16] And A.对方部门ID+0 IN (" & strUserPart & ")"
        End If
    Case Else
        strUserPart = " And A.库房ID+0=[16] "
    End Select
    
    If mlngMode = 模块号.差价调整 Or mlngMode = 模块号.药品移库 Or mlngMode = 模块号.药品领用 Or mlngMode = 模块号.其他出库 Or mlngMode = 模块号.外购入库 Then
        If SQLCondition.str剂型 <> "" And SQLCondition.lng药品分类 = 0 Then
            strSqlForm = " , 药品规格 F, 诊疗项目目录 G, 药品特性 H"
            strFind = strFind & " and a.药品id = f.药品id And f.药名id = g.Id And g.Id = h.药名id(+) and h.药品剂型 in (select * from Table(Cast(f_Str2list([21]) As zlTools.t_Strlist))) and (g.类别='5' or g.类别='6' or g.类别='7')"
        ElseIf SQLCondition.str剂型 = "" And SQLCondition.lng药品分类 <> 0 Then
            strSqlForm = " , 药品规格 F, 诊疗项目目录 G"
            strFind = strFind & " and a.药品id = f.药品id And f.药名id = g.Id And g.分类id + 0=[22] and (g.类别='5' or g.类别='6' or g.类别='7')"
        ElseIf SQLCondition.str剂型 <> "" And SQLCondition.lng药品分类 <> 0 Then
            strSqlForm = " , 药品规格 F, 诊疗项目目录 G, 药品特性 H"
            strFind = strFind & " and a.药品id = f.药品id And f.药名id = g.Id And g.Id = h.药名id(+) and h.药品剂型 in (select * from Table(Cast(f_Str2list([21]) As zlTools.t_Strlist))) and (g.类别='5' or g.类别='6' or g.类别='7') and g.分类id + 0=[22]"
        End If
    End If

    vsfList.Redraw = flexRDNone
    Select Case mlngMode
        Case 模块号.外购入库           '药品外购入库管理
            
            If SQLCondition.int填制审核一并查询 = 0 Then
                gstrSQL = "SELECT A.NO, C.名称 AS 供应商,LTRIM(TO_CHAR(SUM(A.成本金额)," & mstrMoneyFormat & ")) AS 成本金额," & _
                    " LTRIM(TO_CHAR(SUM(A.零售金额)," & mstrMoneyFormat & " )) AS 售价金额, LTRIM(TO_CHAR(SUM(A.差价)," & mstrMoneyFormat & " )) AS 差价金额, " & _
                    " LTRIM(TO_CHAR(SUM(A.零售金额)," & mstrMoneyFormat & " )) AS 零售金额, LTRIM(TO_CHAR(SUM(A.零售金额 - A.成本金额)," & mstrMoneyFormat & " )) AS 零售差价, Nvl(A.费用id, 0) 冲销类型,A.填制人," & _
                    " TO_CHAR(MIN(A.填制日期), 'YYYY-MM-DD HH24:MI:SS') AS 填制日期,A.修改人,TO_CHAR(MIN(A.修改日期), 'YYYY-MM-DD HH24:MI:SS') AS 修改日期,A.配药人 As 核查人,TO_CHAR(MIN(A.配药日期), 'YYYY-MM-DD HH24:MI:SS') As 核查日期,A.审核人," & _
                    " TO_CHAR(MIN(A.审核日期), 'YYYY-MM-DD HH24:MI:SS') AS 审核日期,A.记录状态,Decode(Nvl(A.发药方式, 0), 0, '入库单', '退库单') 单据说明,A.摘要 " & _
                    " FROM 药品收发记录 A, 部门表 B,供应商 C,应付记录 D " & strSqlForm & _
                    " WHERE A.库房ID + 0 = B.ID AND A.供药单位ID+0 = C.ID AND SUBSTR(C.类型,1,1)=1 AND D.系统标识(+)=1 AND D.记录性质(+)=0 " & _
                    " AND A.ID=D.收发ID(+) AND A.单据 = 1 " & strUserPart & strFind & _
                    " GROUP BY A.NO,C.名称,Nvl(A.费用id, 0),A.填制人,A.修改人,A.配药人,A.配药日期,A.审核人,A.记录状态, A.发药方式,A.摘要 " & _
                    " ORDER BY NO DESC,填制日期 ASC"
            Else
                gstrSQL = "SELECT A.NO, C.名称 AS 供应商,LTRIM(TO_CHAR(SUM(A.成本金额)," & mstrMoneyFormat & ")) AS 成本金额," & _
                    " LTRIM(TO_CHAR(SUM(A.零售金额)," & mstrMoneyFormat & " )) AS 售价金额, LTRIM(TO_CHAR(SUM(A.差价)," & mstrMoneyFormat & " )) AS 差价金额, " & _
                    " LTRIM(TO_CHAR(SUM(A.零售金额)," & mstrMoneyFormat & " )) AS 零售金额, LTRIM(TO_CHAR(SUM(A.零售金额 - A.成本金额)," & mstrMoneyFormat & " )) AS 零售差价, Nvl(A.费用id, 0) 冲销类型,A.填制人," & _
                    " TO_CHAR(MIN(A.填制日期), 'YYYY-MM-DD HH24:MI:SS') AS 填制日期,A.修改人,TO_CHAR(MIN(A.修改日期), 'YYYY-MM-DD HH24:MI:SS') AS 修改日期,A.配药人 As 核查人,TO_CHAR(MIN(A.配药日期), 'YYYY-MM-DD HH24:MI:SS') As 核查日期,A.审核人," & _
                    " TO_CHAR(MIN(A.审核日期), 'YYYY-MM-DD HH24:MI:SS') AS 审核日期,A.记录状态,Decode(Nvl(A.发药方式, 0), 0, '入库单', '退库单') 单据说明,A.摘要 " & _
                    " FROM 药品收发记录 A, 部门表 B,供应商 C,应付记录 D " & strSqlForm & _
                    " WHERE A.库房ID + 0 = B.ID AND A.供药单位ID+0 = C.ID AND SUBSTR(C.类型,1,1)=1 AND D.系统标识(+)=1 AND D.记录性质(+)=0 " & _
                    " AND A.ID=D.收发ID(+) AND A.单据 = 1 " & strUserPart & strFind & " And (A.填制日期 Between [3] And [4]) And A.审核日期 Is Null " & _
                    " GROUP BY A.NO,C.名称,Nvl(A.费用id, 0),A.填制人,A.修改人,A.配药人,A.配药日期,A.审核人,A.记录状态,A.发药方式,A.摘要 " & _
                    " Union All " & _
                    " SELECT A.NO, C.名称 AS 供应商,LTRIM(TO_CHAR(SUM(A.成本金额)," & mstrMoneyFormat & ")) AS 成本金额," & _
                    " LTRIM(TO_CHAR(SUM(A.零售金额)," & mstrMoneyFormat & " )) AS 售价金额, LTRIM(TO_CHAR(SUM(A.差价)," & mstrMoneyFormat & " )) AS 差价金额, " & _
                    " LTRIM(TO_CHAR(SUM(A.零售金额)," & mstrMoneyFormat & " )) AS 零售金额, LTRIM(TO_CHAR(SUM(A.零售金额 - A.成本金额)," & mstrMoneyFormat & " )) AS 零售差价, Nvl(A.费用id, 0) 冲销类型,A.填制人," & _
                    " TO_CHAR(MIN(A.填制日期), 'YYYY-MM-DD HH24:MI:SS') AS 填制日期,A.修改人,TO_CHAR(MIN(A.修改日期), 'YYYY-MM-DD HH24:MI:SS') AS 修改日期,A.配药人 As 核查人,TO_CHAR(MIN(A.配药日期), 'YYYY-MM-DD HH24:MI:SS') As 核查日期,A.审核人," & _
                    " TO_CHAR(MIN(A.审核日期), 'YYYY-MM-DD HH24:MI:SS') AS 审核日期,A.记录状态,Decode(Nvl(A.发药方式, 0), 0, '入库单', '退库单') 单据说明,A.摘要 " & _
                    " FROM 药品收发记录 A, 部门表 B,供应商 C,应付记录 D " & strSqlForm & _
                    " WHERE A.库房ID + 0 = B.ID AND A.供药单位ID+0 = C.ID AND SUBSTR(C.类型,1,1)=1 AND D.系统标识(+)=1 AND D.记录性质(+)=0 " & _
                    " AND A.ID=D.收发ID(+) AND A.单据 = 1 " & strUserPart & strFind & " And (A.审核日期 Between [5] And [6]) " & _
                    " GROUP BY A.NO,C.名称,Nvl(A.费用id, 0),A.填制人,A.修改人,A.配药人,A.配药日期,A.审核人,A.记录状态,A.发药方式,A.摘要 " & _
                    " ORDER BY NO DESC,填制日期 ASC"
            End If
        Case 模块号.自制入库           '药品自制入库管理
            gstrSQL = "SELECT A.NO, C.名称 AS 制剂室,LTRIM(TO_CHAR (SUM (A.成本金额), " & mstrMoneyFormat & ")) AS 成本金额," & _
                " LTRIM(TO_CHAR ( (SUM (A.零售金额)), " & mstrMoneyFormat & ")) AS 售价金额,  LTRIM(TO_CHAR((SUM(A.零售金额 - A.成本金额))," & mstrMoneyFormat & " )) AS 差价金额, A.填制人, " & _
                " TO_CHAR (MIN(A.填制日期), 'YYYY-MM-DD HH24:MI:SS') AS 填制日期, A.审核人, " & _
                " TO_CHAR (MIN(A.审核日期), 'YYYY-MM-DD HH24:MI:SS') AS 审核日期, A.记录状态, A.摘要 " & _
                " FROM 药品收发记录 A, 部门表 B ,部门表 C " & _
                " WHERE A.库房ID = B.ID AND A.对方部门ID=C.ID AND A.单据 = 2 AND A.入出系数=1 " & _
                strUserPart & strFind & _
                " GROUP BY A.NO,C.名称,A.填制人,A.审核人,A.记录状态,A.摘要 " & _
                " ORDER BY NO DESC, 填制日期 ASC "
    
        Case 模块号.其他入库           '药品其他入库管理
'            gstrSQL = "SELECT /*+ Rule*/ A.NO, C.名称 AS 入出类别,LTRIM(TO_CHAR (SUM (A.成本金额), " & mstrMoneyFormat & ")) AS 成本金额," & _
'                " LTRIM(TO_CHAR (SUM (A.零售金额)-Sum(To_Number(Nvl(A.用法, 0))), " & mstrMoneyFormat & ")) AS 售价金额,LTRIM(TO_CHAR(SUM(A.零售金额 - A.成本金额- To_Number(Nvl(A.用法, 0)))," & mstrMoneyFormat & " )) AS 差价金额, " & _
'                " LTRIM(TO_CHAR (SUM (A.零售金额), " & mstrMoneyFormat & ")) AS 零售金额,LTRIM(TO_CHAR(SUM(A.零售金额 - A.成本金额)," & mstrMoneyFormat & " )) AS 零售差价, " & _
'                " A.填制人,TO_CHAR (MIN(A.填制日期), 'YYYY-MM-DD HH24:MI:SS') AS 填制日期, A.审核人," & _
'                " TO_CHAR (MIN(A.审核日期), 'YYYY-MM-DD HH24:MI:SS') AS 审核日期, A.记录状态, A.摘要 " & _
'                " FROM 药品收发记录 A, 部门表 B,药品入出类别 C " & _
'                " WHERE A.库房ID = B.ID AND A.入出类别ID = C.ID AND A.单据 = 4 " & _
'                strUserPart & StrFind & _
'                " GROUP BY A.NO,C.名称,A.填制人,A.审核人,A.记录状态,A.摘要 " & _
'                " ORDER BY NO DESC,填制日期 ASC "
            gstrSQL = "SELECT  A.NO, C.名称 AS 入出类别,LTRIM(TO_CHAR (SUM (A.成本金额), " & mstrMoneyFormat & ")) AS 成本金额," & _
                " LTRIM(TO_CHAR (SUM (A.零售金额), " & mstrMoneyFormat & ")) AS 售价金额,LTRIM(TO_CHAR(SUM(A.差价)," & mstrMoneyFormat & " )) AS 差价金额, " & _
                " LTRIM(TO_CHAR (SUM (A.零售金额), " & mstrMoneyFormat & ")) AS 零售金额,LTRIM(TO_CHAR(SUM(A.零售金额 - A.成本金额)," & mstrMoneyFormat & " )) AS 零售差价, " & _
                " A.填制人,TO_CHAR (MIN(A.填制日期), 'YYYY-MM-DD HH24:MI:SS') AS 填制日期,A.修改人,TO_CHAR (MIN(A.修改日期), 'YYYY-MM-DD HH24:MI:SS') AS 修改日期, A.审核人," & _
                " TO_CHAR (MIN(A.审核日期), 'YYYY-MM-DD HH24:MI:SS') AS 审核日期, A.记录状态, A.摘要 " & _
                " FROM 药品收发记录 A, 部门表 B,药品入出类别 C " & _
                " WHERE A.库房ID = B.ID AND A.入出类别ID = C.ID AND A.单据 = 4 " & _
                strUserPart & strFind & _
                " GROUP BY A.NO,C.名称,A.填制人,A.修改人,A.审核人,A.记录状态,A.摘要 " & _
                " ORDER BY NO DESC,填制日期 ASC "
        Case 模块号.差价调整           '库存差价调整管理
            gstrSQL = "SELECT  A.NO, LTRIM(TO_CHAR (SUM (A.零售价), " & mstrMoneyFormat & ")) AS 库存金额,LTRIM(TO_CHAR (SUM (A.成本价),'9999999999999990.00000')) AS 库存差价," & _
                " LTRIM(TO_CHAR ( (SUM (A.差价)), '9999999999999990.00000')) AS 调整额, A.填制人,TO_CHAR (MIN(A.填制日期), 'YYYY-MM-DD HH24:MI:SS') AS 填制日期, A.审核人," & _
                " TO_CHAR (MIN(A.审核日期), 'YYYY-MM-DD HH24:MI:SS') AS 审核日期, A.记录状态, A.摘要,0 发药方式 " & _
                " FROM 药品收发记录 A, 部门表 B  " & strSqlForm & _
                " WHERE A.库房ID = B.ID  AND A.单据 = 5 And nvl(发药方式,0)=0 " & _
                strUserPart & strFind & _
                " GROUP BY A.NO,A.填制人,A.审核人,A.记录状态,A.摘要 " & _
                "UNION ALL " & _
                "SELECT A.NO, LTRIM(TO_CHAR (SUM (A.零售价), " & mstrMoneyFormat & ")) AS 库存金额,LTRIM(TO_CHAR (SUM (A.成本价)," & mstrMoneyFormat & ")) AS 库存差价," & _
                " LTRIM(TO_CHAR ( (SUM (A.差价)), " & mstrMoneyFormat & ")) AS 调整额, A.填制人,TO_CHAR (MIN(A.填制日期), 'YYYY-MM-DD HH24:MI:SS') AS 填制日期, A.审核人," & _
                " TO_CHAR (MIN(A.审核日期), 'YYYY-MM-DD HH24:MI:SS') AS 审核日期, A.记录状态, A.摘要,1 发药方式 " & _
                " FROM 药品收发记录 A, 部门表 B  " & strSqlForm & _
                " WHERE A.库房ID = B.ID  AND A.单据 = 5 And  nvl(发药方式,0)=1  and a.库房id=" & cboStock.ItemData(cboStock.ListIndex) & _
                strFind & _
                " GROUP BY A.NO,A.填制人,A.审核人,A.记录状态,A.摘要 " & _
                " ORDER BY NO DESC,填制日期 ASC "
        Case 模块号.药品移库           '药品移库管理（查我的出库，条件不限；查我的入库，则只能看到未备药或已发送的单据）
            gstrSQL = "SELECT A.NO," & IIf(TabShow.Tab = 0, "C.名称 As 移入库房,", "B.名称 AS 移出库房,") & " LTRIM(TO_CHAR (SUM (A.成本金额), " & mstrMoneyFormat & ")) AS 成本金额, " & _
                " LTRIM(TO_CHAR ( (SUM (A.零售金额)), " & mstrMoneyFormat & ")) AS 售价金额, LTRIM(TO_CHAR((SUM(A.零售金额 - A.成本金额))," & mstrMoneyFormat & " )) AS 差价金额,A.填制人, " & _
                " TO_CHAR (MIN(A.填制日期), 'YYYY-MM-DD HH24:MI:SS') AS 填制日期,A.修改人,TO_CHAR (MIN(A.修改日期), 'YYYY-MM-DD HH24:MI:SS') AS 修改日期," & strTransfer & " ,A.记录状态, A.摘要 " & _
                " FROM 药品收发记录 A, 部门表 B ,部门表 C  " & strSqlForm & _
                " WHERE A.库房ID = B.ID AND A.对方部门ID=C.ID AND A.单据 = 6 AND  A.入出系数=-1" & _
                IIf(TabShow.Tab = 0, " ", " And (A.配药人 Is NULL Or A.配药日期 Is Not NULL)") & _
                strUserPart & strFind & _
                " GROUP BY A.NO," & IIf(TabShow.Tab = 0, "C.名称", "B.名称") & ",A.填制人,A.修改人" & strTransfer_Order & ",A.记录状态,A.摘要 " & _
                " ORDER BY NO DESC, 填制日期 ASC "
        Case 模块号.药品领用           '药品领用管理
            gstrSQL = "SELECT A.NO, C.名称 AS 领用部门,LTRIM(TO_CHAR (SUM (A.成本金额), " & mstrMoneyFormat & ")) AS 成本金额, " & _
                " LTRIM(TO_CHAR ( (SUM (A.零售金额)), " & mstrMoneyFormat & ")) AS 售价金额, LTRIM(TO_CHAR((SUM(A.零售金额 - A.成本金额))," & mstrMoneyFormat & " )) AS 差价金额,A.填制人, " & _
                " TO_CHAR (MIN(A.填制日期), 'YYYY-MM-DD HH24:MI:SS') AS 填制日期, A.修改人,TO_CHAR (MIN(A.修改日期), 'YYYY-MM-DD HH24:MI:SS') AS 修改日期, A.审核人, " & _
                " TO_CHAR (MIN(A.审核日期), 'YYYY-MM-DD HH24:MI:SS') AS 审核日期, A.记录状态, A.摘要,Decode(Nvl(A.发药方式,0),0,'向库房领药','向留存领药') 领药方式,A.冲销原因  " & _
                " FROM 药品收发记录 A, 部门表 B ,部门表 C  " & strSqlForm & _
                " WHERE A.库房ID = B.ID AND A.对方部门ID=C.ID AND A.单据 = 7  " & _
                strUserPart & strFind & _
                " GROUP BY A.NO,C.名称,A.填制人,A.修改人,A.审核人,A.记录状态,A.摘要,Nvl(A.发药方式,0),A.冲销原因 " & _
                " ORDER BY NO DESC, 填制日期 ASC "
        Case 模块号.其他出库          '药品其他出库管理
            If SQLCondition.int填制审核一并查询 = 0 Then
                gstrSQL = "SELECT A.NO, C.名称 AS 入出类别,Decode(C.名称, '药品外调', D.名称, Decode(C.名称, '药品外销', E.名称, '')) AS 对方单位,LTRIM(TO_CHAR (SUM (A.成本金额), " & mstrMoneyFormat & ")) AS 成本金额, " & _
                    " LTRIM(TO_CHAR ( (SUM (A.零售金额)), " & mstrMoneyFormat & ")) AS 售价金额, LTRIM(TO_CHAR((SUM(A.零售金额 - A.成本金额))," & mstrMoneyFormat & " )) AS 差价金额,LTRIM(TO_CHAR((SUM(A.单量 * A.实际数量))," & mstrMoneyFormat & " )) AS 外调外销金额,A.填制人," & _
                    " TO_CHAR (MIN(A.填制日期), 'YYYY-MM-DD HH24:MI:SS') AS 填制日期, A.修改人,TO_CHAR (MIN(A.修改日期), 'YYYY-MM-DD HH24:MI:SS') AS 修改日期, A.审核人," & _
                    " TO_CHAR (MIN(A.审核日期), 'YYYY-MM-DD HH24:MI:SS') AS 审核日期, A.记录状态, A.摘要 " & _
                    " FROM 药品收发记录 A, 部门表 B,药品入出类别 C,药品外调单位 D, 药品外销单位 E  " & strSqlForm & _
                    " WHERE A.库房ID = B.ID AND A.入出类别ID = C.ID AND A.发药窗口=D.编码(+) And A.发药窗口 = E.编码(+) AND A.单据 = 11 " & _
                    strUserPart & strFind & _
                    " GROUP BY A.NO,C.名称,D.名称,E.名称,A.填制人, A.修改人,A.审核人,A.记录状态,A.摘要 " & _
                    " ORDER BY NO DESC,填制日期 ASC "
            Else
                gstrSQL = "SELECT A.NO, C.名称 AS 入出类别,Decode(C.名称, '药品外调', D.名称, Decode(C.名称, '药品外销', E.名称, '')) AS 对方单位,LTRIM(TO_CHAR (SUM (A.成本金额), " & mstrMoneyFormat & ")) AS 成本金额, " & _
                    " LTRIM(TO_CHAR ( (SUM (A.零售金额)), " & mstrMoneyFormat & ")) AS 售价金额, LTRIM(TO_CHAR((SUM(A.零售金额 - A.成本金额))," & mstrMoneyFormat & " )) AS 差价金额,LTRIM(TO_CHAR((SUM(A.单量 * A.实际数量))," & mstrMoneyFormat & " )) AS 外调外销金额,A.填制人," & _
                    " TO_CHAR (MIN(A.填制日期), 'YYYY-MM-DD HH24:MI:SS') AS 填制日期, A.修改人,TO_CHAR (MIN(A.修改日期), 'YYYY-MM-DD HH24:MI:SS') AS 修改日期, A.审核人," & _
                    " TO_CHAR (MIN(A.审核日期), 'YYYY-MM-DD HH24:MI:SS') AS 审核日期, A.记录状态, A.摘要 " & _
                    " FROM 药品收发记录 A, 部门表 B,药品入出类别 C,药品外调单位 D, 药品外销单位 E  " & strSqlForm & _
                    " WHERE A.库房ID = B.ID AND A.入出类别ID = C.ID AND A.发药窗口=D.编码(+) And A.发药窗口 = E.编码(+) AND A.单据 = 11 " & _
                    strUserPart & strFind & " And (A.填制日期 Between [3] And [4]) And A.审核日期 Is Null " & _
                    " GROUP BY A.NO,C.名称,D.名称,E.名称,A.填制人, A.修改人,A.审核人,A.记录状态,A.摘要 " & _
                    " Union " & _
                    " SELECT A.NO, C.名称 AS 入出类别,Decode(C.名称, '药品外调', D.名称, Decode(C.名称, '药品外销', E.名称, '')) AS 对方单位,LTRIM(TO_CHAR (SUM (A.成本金额), " & mstrMoneyFormat & ")) AS 成本金额, " & _
                    " LTRIM(TO_CHAR ( (SUM (A.零售金额)), " & mstrMoneyFormat & ")) AS 售价金额, LTRIM(TO_CHAR((SUM(A.零售金额 - A.成本金额))," & mstrMoneyFormat & " )) AS 差价金额,LTRIM(TO_CHAR((SUM(A.单量 * A.实际数量))," & mstrMoneyFormat & " )) AS 外调外销金额,A.填制人," & _
                    " TO_CHAR (MIN(A.填制日期), 'YYYY-MM-DD HH24:MI:SS') AS 填制日期, A.修改人,TO_CHAR (MIN(A.修改日期), 'YYYY-MM-DD HH24:MI:SS') AS 修改日期, A.审核人," & _
                    " TO_CHAR (MIN(A.审核日期), 'YYYY-MM-DD HH24:MI:SS') AS 审核日期, A.记录状态, A.摘要 " & _
                    " FROM 药品收发记录 A, 部门表 B,药品入出类别 C,药品外调单位 D, 药品外销单位 E " & strSqlForm & _
                    " WHERE A.库房ID = B.ID AND A.入出类别ID = C.ID AND A.发药窗口=D.编码(+) And A.发药窗口 = E.编码(+) AND A.单据 = 11 " & _
                    strUserPart & strFind & " And (A.审核日期 Between [5] And [6]) " & _
                    " GROUP BY A.NO,C.名称,D.名称,E.名称,A.填制人, A.修改人,A.审核人,A.记录状态,A.摘要 " & _
                    " ORDER BY NO DESC,填制日期 ASC "
            End If
    End Select

    Set rsList = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, _
        SQLCondition.strNO开始, _
        SQLCondition.strNO结束, _
        SQLCondition.date填制时间开始, _
        SQLCondition.date填制时间结束, _
        SQLCondition.date审核时间开始, _
        SQLCondition.date审核时间结束, _
        SQLCondition.lng药品, _
        SQLCondition.lng库房, _
        SQLCondition.str填制人, _
        SQLCondition.str审核人, _
        SQLCondition.lng生产商, _
        SQLCondition.str产地, _
        SQLCondition.str发票号开始, _
        SQLCondition.str发票号结束, _
        SQLCondition.lng入出类别, _
        cboStock.ItemData(cboStock.ListIndex), _
        UserInfo.用户ID, _
        gstrNodeNo, _
        SQLCondition.date发票审核日期开始, _
        SQLCondition.date发票审核日期结束, _
        SQLCondition.str剂型, _
        SQLCondition.lng药品分类)
    
    mblnBandEvent = True
    Set vsfList.DataSource = rsList
    mblnBandEvent = False
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
        
        For n = 0 To .Cols - 1
            .ColKey(n) = .TextMatrix(0, n)
            .FixedAlignment(n) = flexAlignCenterCenter
        Next
    End With
    SetListColWidth
    DoEvents
    
    '统计合计金额
    lbl1.Caption = ""
    lbl2.Caption = ""
    lbl3.Caption = ""
    If (Not rsList.EOF) And (Not rsList.BOF) Then
        rsList.MoveFirst
        Do While Not rsList.EOF
            Select Case mlngMode
                Case 模块号.外购入库
                    dbl1 = dbl1 + IIf(IsNull(rsList!成本金额), 0, rsList!成本金额)
                    dbl2 = dbl2 + IIf(IsNull(rsList!售价金额), 0, rsList!售价金额)
                    dbl3 = dbl3 + IIf(IsNull(rsList!差价金额), 0, rsList!差价金额)
                Case 模块号.自制入库, 模块号.其他入库, 模块号.药品移库, 模块号.药品领用
                    dbl1 = dbl1 + IIf(IsNull(rsList!成本金额), 0, rsList!成本金额)
                    dbl2 = dbl2 + IIf(IsNull(rsList!售价金额), 0, rsList!售价金额)
                    dbl3 = dbl3 + IIf(IsNull(rsList!差价金额), 0, rsList!差价金额)
                Case 模块号.其他出库
                    dbl1 = dbl1 + IIf(IsNull(rsList!成本金额), 0, rsList!成本金额)
                    dbl2 = dbl2 + IIf(IsNull(rsList!售价金额), 0, rsList!售价金额)
                    dbl3 = dbl3 + IIf(IsNull(rsList!差价金额), 0, rsList!差价金额)
                    dbl4 = dbl4 + IIf(IsNull(rsList!外调外销金额), 0, rsList!外调外销金额)
                Case 模块号.差价调整
                    dbl1 = dbl1 + IIf(IsNull(rsList!库存金额), 0, rsList!库存金额)
                    dbl2 = dbl2 + IIf(IsNull(rsList!库存差价), 0, rsList!库存差价)
                    dbl3 = dbl3 + IIf(IsNull(rsList!调整额), 0, rsList!调整额)
            End Select
            rsList.MoveNext
        Loop
        
        rsList.MoveFirst
        
        Select Case mlngMode
            Case 模块号.外购入库
                lbl1.Caption = "成本金额合计：" & Format(dbl1, StrFormat)
                lbl2.Caption = "售价金额合计：" & Format(dbl2, StrFormat)
                lbl3.Caption = "差价金额合计：" & Format(dbl3, StrFormat)
            Case 模块号.自制入库, 模块号.其他入库, 模块号.药品移库, 模块号.药品领用
                lbl1.Caption = "成本金额合计：" & Format(dbl1, StrFormat)
                lbl2.Caption = "售价金额合计：" & Format(dbl2, StrFormat)
                lbl3.Caption = "差价金额合计：" & Format(dbl3, StrFormat)
            Case 模块号.其他出库
                lbl1.Caption = "成本金额合计：" & Format(dbl1, StrFormat)
                lbl2.Caption = "售价金额合计：" & Format(dbl2, StrFormat)
                lbl3.Caption = "差价金额合计：" & Format(dbl3, StrFormat)
                lbl4.Caption = "外调(销)金额合计：" & Format(dbl4, StrFormat)
            Case 模块号.差价调整
                lbl1.Caption = "库存金额合计：" & Format(dbl1, StrFormat)
                lbl2.Caption = "库存差价合计：" & Format(dbl2, StrFormat)
                lbl3.Caption = "调整额合计：" & Format(dbl3, StrFormat)
        End Select
    
    End If

    With vsfList
        If mintListRow >= .rows Then
            mintListRow = .rows - 1
        End If
        .Row = IIf(.rows > 1, IIf(mintListRow > 1, mintListRow, 1), 1)
        .Col = 0
'        .ColSel = .Cols - 1
    End With
    
    vsfList_EnterCell    '列出单据体
    
    SetListColor
    
    SetEnable
    vsfList.Redraw = flexRDDirect
    Call FS.StopFlash

    Screen.MousePointer = vbDefault
    staThis.Panels(2).Text = "当前共有" & rsList.RecordCount & "张单据"
    rsList.Close
    
    If vsfList.Visible = True Then
        vsfList.SetFocus
        vsfList.Row = IIf(vsfList.rows > 1, IIf(mintListRow > 1, mintListRow, 1), 1)
        vsfList.TopRow = vsfList.Row
    End If
    If mblnDo Then
        RestoreFlexState vsfList, App.ProductName & "\" & Me.Name & mstrTitle
        RestoreFlexState vsfDetail, App.ProductName & "\" & Me.Name & mstrTitle
    End If
    If mblnViewCost = False Then
        lbl1.Visible = False
        lbl3.Visible = False
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetListColor()
    Dim intStatus As Integer
    Dim intRow As Integer
    Dim intCol As Integer
    Dim int冲销状态 As Integer      '0-已申请的冲销记录;1-已审核的冲销记录
    
    With vsfList
        If .rows <= 2 Then Exit Sub
        For intRow = 1 To .rows - 1
'            intStatus = .TextMatrix(intRow, IIf(mlngMode = 模块号.差价调整 Or mlngMode = 模块号.外购入库 Or mlngMode = 模块号.药品领用, .Cols - 3, .Cols - 2))
            intStatus = .TextMatrix(intRow, .ColIndex("记录状态"))
            If intStatus Mod 3 = 0 Then
                .Cell(flexcpForeColor, intRow, 0, intRow, .Cols - 1) = &H80000001
            End If
            If intStatus Mod 3 = 2 Then
                '移库冲销单根据审核日期判断状态
                If mlngMode = 模块号.药品移库 Then
                    If Trim(.TextMatrix(intRow, .ColIndex("接收日期"))) <> "" Then
                        int冲销状态 = 1
                    Else
                        int冲销状态 = 0
                    End If
                End If
                
                '领用冲销单根据审核日期判断状态
                If mlngMode = 模块号.药品领用 Then
                    If Trim(.TextMatrix(intRow, .ColIndex("审核日期"))) <> "" Then
                        int冲销状态 = 1
                    Else
                        int冲销状态 = 0
                    End If
                End If
                
                If mlngMode = 模块号.外购入库 Then
                    '外购入库中财务审核时冲销的单据为浅红色，其他普通冲销单据为红色
                    .Cell(flexcpForeColor, intRow, 0, intRow, .Cols - 1) = IIf(Val(.TextMatrix(intRow, 外购主表.冲销类型)) = 1, &HFF00FF, &HFF)
                ElseIf mlngMode = 模块号.药品移库 Or mlngMode = 模块号.药品领用 Then
                    '移库、领用中申请冲销单据为浅红色，已冲销单据为红色
                    .Cell(flexcpForeColor, intRow, 0, intRow, .Cols - 1) = IIf(int冲销状态 = 0, &HFF00FF, &HFF)
                Else
                    '其他单据已冲销单据为红色
                    .Cell(flexcpForeColor, intRow, 0, intRow, .Cols - 1) = &HFF
                End If
                
                '领用存在冲销单据，冲销原因列显示
                If mlngMode = 模块号.药品领用 Then If vsfList.colHidden(.ColIndex("冲销原因")) Then vsfList.colHidden(.ColIndex("冲销原因")) = False
                
            End If
        Next
    End With
End Sub

'表头列宽初始
Private Sub SetListColWidth()
    Dim intCol As Integer
    
    With vsfList
        Select Case mlngMode
            Case 模块号.外购入库
                .ColAlignment(外购主表.成本金额) = flexAlignRightCenter
                .ColAlignment(外购主表.售价金额) = flexAlignRightCenter
                .ColAlignment(外购主表.差价金额) = flexAlignRightCenter
                .ColAlignment(外购主表.零售金额) = flexAlignRightCenter
                .ColAlignment(外购主表.零售差价) = flexAlignRightCenter
            Case 模块号.自制入库
                .ColAlignment(2) = flexAlignRightCenter
                .ColAlignment(3) = flexAlignRightCenter
                .ColAlignment(4) = flexAlignRightCenter
            Case 模块号.其他入库
                .ColAlignment(2) = flexAlignRightCenter
                .ColAlignment(3) = flexAlignRightCenter
                .ColAlignment(4) = flexAlignRightCenter
                .ColAlignment(5) = flexAlignRightCenter
                .ColAlignment(6) = flexAlignRightCenter
            Case 模块号.差价调整
                .ColAlignment(1) = flexAlignRightCenter
                .ColAlignment(2) = flexAlignRightCenter
                .ColAlignment(3) = flexAlignRightCenter
            Case 模块号.药品移库
                .ColAlignment(2) = flexAlignRightCenter
                .ColAlignment(3) = flexAlignRightCenter
                .ColAlignment(4) = flexAlignRightCenter
            Case 模块号.药品领用
                .ColAlignment(2) = flexAlignRightCenter
                .ColAlignment(3) = flexAlignRightCenter
                .ColAlignment(4) = flexAlignRightCenter
            Case 模块号.其他出库
                .ColAlignment(3) = flexAlignRightCenter         '售价金额
                .ColAlignment(4) = flexAlignRightCenter
                .ColAlignment(5) = flexAlignRightCenter
                .ColAlignment(6) = flexAlignRightCenter
        End Select
        
        For intCol = 1 To .Cols - 1
            If intCol = 1 Then
                If mlngMode = 模块号.差价调整 Then
                    .ColWidth(intCol) = 1000
                Else
                    .ColWidth(intCol) = 2000
                End If
            ElseIf intCol = .Cols - 2 Then
                  .ColWidth(intCol) = 0
            Else
                .ColWidth(intCol) = 1000
            End If
        Next
        If mlngMode = 模块号.差价调整 Then
            .ColWidth(.Cols - 1) = 0
            .ColWidth(.Cols - 3) = 0
        End If
        If mlngMode = 模块号.外购入库 Then
'            .ColWidth(外购主表.零售金额) = IIf(InStr("|" & mstr屏蔽列 & "|", "|零售金额|") = 0, 1000, 0)
'            .ColWidth(外购主表.零售差价) = IIf(InStr("|" & mstr屏蔽列 & "|", "|零售差价|") = 0, 1000, 0)
            .ColWidth(外购主表.零售金额) = 0
            .ColWidth(外购主表.零售差价) = 0
            .ColWidth(外购主表.冲销类型) = 0
            .ColWidth(外购主表.记录状态) = 0
            .ColWidth(外购主表.单据说明) = 1000
            If mbln核查 = True Then
                If .ColWidth(外购主表.核查人) = 0 Then .ColWidth(外购主表.核查人) = 1000
                If .ColWidth(外购主表.核查日期) = 0 Then .ColWidth(外购主表.核查日期) = 1000
            Else
                .ColWidth(外购主表.核查人) = 0
                .ColWidth(外购主表.核查日期) = 0
            End If
        End If
        If mlngMode = 模块号.其他入库 Then
'            .ColWidth(5) = IIf(InStr("|" & mstr屏蔽列 & "|", "|零售金额|") = 0, 1000, 0)
'            .ColWidth(6) = IIf(InStr("|" & mstr屏蔽列 & "|", "|零售差价|") = 0, 1000, 0)
            .ColWidth(5) = 0
            .ColWidth(6) = 0
        End If
        If mlngMode = 模块号.药品领用 Then
            .ColWidth(.Cols - 4) = 0
            .ColWidth(.Cols - 3) = 1000
            .ColWidth(.Cols - 2) = 1000
            
            .colHidden(.ColIndex("冲销原因")) = True '冲销原因默认不显示
        End If
        If mblnViewCost = False Then
            Select Case mlngMode
                Case 模块号.外购入库
                    .colHidden(外购主表.成本金额) = True
                    .colHidden(外购主表.差价金额) = True
                Case 模块号.自制入库
                    .colHidden(.ColIndex("成本金额")) = True
                    .colHidden(.ColIndex("差价金额")) = True
                Case 模块号.其他入库
                    .colHidden(.ColIndex("成本金额")) = True
                    .colHidden(.ColIndex("差价金额")) = True
                Case 模块号.药品移库
                    .colHidden(.ColIndex("成本金额")) = True
                    .colHidden(.ColIndex("差价金额")) = True
                Case 模块号.药品领用
                    .colHidden(.ColIndex("成本金额")) = True
                    .colHidden(.ColIndex("差价金额")) = True
                Case 模块号.其他出库
                    .colHidden(.ColIndex("成本金额")) = True
                    .colHidden(.ColIndex("差价金额")) = True
            End Select
        End If
    End With
End Sub

Private Sub SetDetailColWidth()
    Dim intCol As Integer
    Dim rsDetail As New Recordset
    Dim bln中药库房 As Boolean
    Dim str库房性质 As String
    
    On Error GoTo errHandle

    With vsfDetail
        Select Case mlngMode
            Case 模块号.外购入库
                .ColAlignment(.ColIndex("数量")) = flexAlignRightCenter     '数量
                .ColAlignment(.ColIndex("单位")) = flexAlignCenterCenter    '单位
                .ColAlignment(.ColIndex("成本价")) = flexAlignRightCenter     '成本价
                .ColAlignment(.ColIndex("成本金额")) = flexAlignRightCenter     '成本金额
                .ColAlignment(.ColIndex("扣率")) = flexAlignRightCenter     '扣率
                .ColAlignment(.ColIndex("售价")) = flexAlignRightCenter    '售价
                .ColAlignment(.ColIndex("售价金额")) = flexAlignRightCenter    '售价金额
                .ColAlignment(.ColIndex("差价")) = flexAlignRightCenter    '差价
                .ColAlignment(.ColIndex("批准文号")) = flexAlignLeftCenter     '批准文号
                .ColAlignment(.ColIndex("发票金额")) = flexAlignRightCenter    '发票金额
                .ColAlignment(.ColIndex("差价让利比")) = flexAlignRightCenter    '差价让利比
                .ColAlignment(.ColIndex("随货单号")) = flexAlignLeftCenter     '随货单号
                .ColAlignment(.ColIndex("零售价")) = flexAlignRightCenter    '零售价
                .ColAlignment(.ColIndex("零售金额")) = flexAlignRightCenter    '零售金额
                .ColAlignment(.ColIndex("零售差价")) = flexAlignRightCenter    '零售差价
            Case 模块号.自制入库
                .ColAlignment(.ColIndex("数量")) = flexAlignRightCenter     '数量
                .ColAlignment(.ColIndex("单位")) = flexAlignCenterCenter    '单位
                .ColAlignment(.ColIndex("采购价")) = flexAlignRightCenter     '成本价
                .ColAlignment(.ColIndex("采购金额")) = flexAlignRightCenter     '成本金额
                .ColAlignment(.ColIndex("售价")) = flexAlignRightCenter    '售价
                .ColAlignment(.ColIndex("售价金额")) = flexAlignRightCenter    '售价金额
                .ColAlignment(.ColIndex("差价")) = flexAlignRightCenter    '差价
            Case 模块号.其他入库
                .ColAlignment(.ColIndex("数量")) = flexAlignRightCenter     '数量
                .ColAlignment(.ColIndex("单位")) = flexAlignCenterCenter    '单位
                .ColAlignment(.ColIndex("成本价")) = flexAlignRightCenter     '成本价
                .ColAlignment(.ColIndex("成本金额")) = flexAlignRightCenter     '成本金额
                .ColAlignment(.ColIndex("售价")) = flexAlignRightCenter    '售价
                .ColAlignment(.ColIndex("售价金额")) = flexAlignRightCenter    '售价金额
                .ColAlignment(.ColIndex("差价")) = flexAlignRightCenter    '差价
                .ColAlignment(.ColIndex("零售价")) = flexAlignRightCenter    '零售价
                .ColAlignment(.ColIndex("零售金额")) = flexAlignRightCenter    '零售金额
                .ColAlignment(.ColIndex("零售差价")) = flexAlignRightCenter    '零售差价
            Case 模块号.差价调整
                .ColAlignment(.ColIndex("单位")) = flexAlignCenterCenter    '单位
                .ColAlignment(.ColIndex("库存金额")) = flexAlignRightCenter     '库存金额
                .ColAlignment(.ColIndex("库存差价")) = flexAlignRightCenter     '库存差价
                .ColAlignment(.ColIndex("调整额")) = flexAlignRightCenter     '调整额
                .ColAlignment(.ColIndex("新成本价")) = flexAlignRightCenter     '新成本价
            Case 模块号.药品移库
                .ColAlignment(.ColIndex("填写数量")) = flexAlignRightCenter     '填写数量
                .ColAlignment(.ColIndex("实际数量")) = flexAlignRightCenter     '实际数量
                .ColAlignment(.ColIndex("单位")) = flexAlignCenterCenter    '单位
                .ColAlignment(.ColIndex("成本价")) = flexAlignRightCenter     '成本价
                .ColAlignment(.ColIndex("成本金额")) = flexAlignRightCenter     '成本金额
                .ColAlignment(.ColIndex("售价")) = flexAlignRightCenter    '售价
                .ColAlignment(.ColIndex("售价金额")) = flexAlignRightCenter    '售价金额
                .ColAlignment(.ColIndex("差价")) = flexAlignRightCenter    '差价
            Case 模块号.药品领用
                .ColAlignment(.ColIndex("填写数量")) = flexAlignRightCenter     '填写数量
                .ColAlignment(.ColIndex("实际数量")) = flexAlignRightCenter     '实际数量
                .ColAlignment(.ColIndex("单位")) = flexAlignCenterCenter    '单位
                .ColAlignment(.ColIndex("成本价")) = flexAlignRightCenter     '成本价
                .ColAlignment(.ColIndex("成本金额")) = flexAlignRightCenter     '成本金额
                .ColAlignment(.ColIndex("售价")) = flexAlignRightCenter    '售价
                .ColAlignment(.ColIndex("售价金额")) = flexAlignRightCenter    '售价金额
                .ColAlignment(.ColIndex("差价")) = flexAlignRightCenter    '差价
            Case 模块号.其他出库
                .ColAlignment(.ColIndex("数量")) = flexAlignRightCenter     '数量
                .ColAlignment(.ColIndex("单位")) = flexAlignCenterCenter    '单位
                .ColAlignment(.ColIndex("成本价")) = flexAlignRightCenter     '成本价
                .ColAlignment(.ColIndex("成本金额")) = flexAlignRightCenter     '成本金额
                .ColAlignment(.ColIndex("售价")) = flexAlignRightCenter    '售价
                .ColAlignment(.ColIndex("售价金额")) = flexAlignRightCenter    '售价金额
                .ColAlignment(.ColIndex("差价")) = flexAlignRightCenter    '差价
                If vsfList.TextMatrix(vsfList.Row, 1) = "药品外销" Then
                    .ColAlignment(.ColIndex("外销价")) = flexAlignRightCenter    '外调价/外销价
                    .ColAlignment(.ColIndex("外销金额")) = flexAlignRightCenter    '外调金额/外销金额
                Else
                    .ColAlignment(.ColIndex("外调价")) = flexAlignRightCenter    '外调价/外销价
                    .ColAlignment(.ColIndex("外调金额")) = flexAlignRightCenter    '外调金额/外销金额
                End If
                .ColAlignment(.ColIndex("增值税率")) = flexAlignRightCenter    '增值税率
                .ColAlignment(.ColIndex("税金")) = flexAlignRightCenter    '税金
        End Select
        
        If mblnBootUp = False Then
            .ColWidth(0) = 500
            If mlngMode = 模块号.外购入库 Then
                .ColWidth(1) = 1000
            Else
                .ColWidth(1) = 2500
            End If
            For intCol = 2 To .Cols - 1
                .ColWidth(intCol) = 1000
            Next
        End If
        
        str库房性质 = ""
        gstrSQL = "Select a.工作性质 From 部门性质说明 A Where a.部门id =[1]"
        Set rsDetail = zlDataBase.OpenSQLRecord(gstrSQL, "判断是库房性质", cboStock.ItemData(cboStock.ListIndex))
        Do While Not rsDetail.EOF
            str库房性质 = str库房性质 & "," & rsDetail!工作性质
            rsDetail.MoveNext
        Loop
        If str库房性质 Like "*中药*" Or str库房性质 Like "*制剂室*" Then bln中药库房 = True
        
        Select Case mlngMode
            Case 模块号.外购入库
                .ColWidth(.ColIndex("付款序号")) = 0
                .ColWidth(.ColIndex("招标药品")) = 0
                .ColWidth(.ColIndex("差价让利比")) = 0
'                .ColWidth(.ColIndex("零售价")) = IIf(InStr("|" & mstr屏蔽列 & "|", "|零售价|") = 0, 1000, 0)
'                .ColWidth(.ColIndex("零售金额")) = IIf(InStr("|" & mstr屏蔽列 & "|", "|零售金额|") = 0, 1000, 0)
'                .ColWidth(.ColIndex("零售差价")) = IIf(InStr("|" & mstr屏蔽列 & "|", "|零售差价|") = 0, 1000, 0)
                .ColWidth(.ColIndex("零售价")) = 0
                .ColWidth(.ColIndex("零售金额")) = 0
                .ColWidth(.ColIndex("零售差价")) = 0
                If mblnViewCost = False Then
                    .colHidden(.ColIndex("成本价")) = True
                    .colHidden(.ColIndex("成本金额")) = True
                    .colHidden(.ColIndex("差价")) = True
                End If
                If bln中药库房 Then
                    .colHidden(.ColIndex("原产地")) = False
                Else
                    .colHidden(.ColIndex("原产地")) = True
                End If
            Case 模块号.其他入库
'                .ColWidth(.ColIndex("零售价")) = IIf(InStr("|" & mstr屏蔽列 & "|", "|零售价|") = 0, 1000, 0)
'                .ColWidth(.ColIndex("零售金额")) = IIf(InStr("|" & mstr屏蔽列 & "|", "|零售金额|") = 0, 1000, 0)
'                .ColWidth(.ColIndex("零售差价")) = IIf(InStr("|" & mstr屏蔽列 & "|", "|零售差价|") = 0, 1000, 0)
                .ColWidth(.ColIndex("零售价")) = 0
                .ColWidth(.ColIndex("零售金额")) = 0
                .ColWidth(.ColIndex("零售差价")) = 0
                If mblnViewCost = False Then
                    .colHidden(.ColIndex("成本价")) = True
                    .colHidden(.ColIndex("成本金额")) = True
                    .colHidden(.ColIndex("差价")) = True
                End If
                If bln中药库房 Then
                    .colHidden(.ColIndex("原产地")) = False
                Else
                    .colHidden(.ColIndex("原产地")) = True
                End If
            Case 模块号.自制入库
                If mblnViewCost = False Then
                    .colHidden(.ColIndex("采购价")) = True
                    .colHidden(.ColIndex("采购金额")) = True
                    .colHidden(.ColIndex("差价")) = True
                End If
            Case 模块号.差价调整
'                If mblnViewCost = False Then
'                    .ColHidden(.ColIndex("库存差价")) = True
'                    .ColHidden(.ColIndex("新成本价")) = True
'                End If
                If bln中药库房 Then
                    .colHidden(.ColIndex("原产地")) = False
                Else
                    .colHidden(.ColIndex("原产地")) = True
                End If
            Case 模块号.药品移库
                If mblnViewCost = False Then
                    .colHidden(.ColIndex("成本价")) = True
                    .colHidden(.ColIndex("成本金额")) = True
                    .colHidden(.ColIndex("差价")) = True
                End If
                If bln中药库房 Then
                    .colHidden(.ColIndex("原产地")) = False
                Else
                    .colHidden(.ColIndex("原产地")) = True
                End If
            Case 模块号.药品领用
                If mblnViewCost = False Then
                    .colHidden(.ColIndex("成本价")) = True
                    .colHidden(.ColIndex("成本金额")) = True
                    .colHidden(.ColIndex("差价")) = True
                End If
                .ColWidth(.ColIndex("名称")) = 0
                If bln中药库房 Then
                    .colHidden(.ColIndex("原产地")) = False
                Else
                    .colHidden(.ColIndex("原产地")) = True
                End If
            Case 模块号.其他出库
                If mblnViewCost = False Then
                    .colHidden(.ColIndex("成本价")) = True
                    .colHidden(.ColIndex("成本金额")) = True
                    .colHidden(.ColIndex("差价")) = True
                End If
                .ColWidth(.ColIndex("名称")) = 0
                
                If vsfList.TextMatrix(vsfList.Row, 1) = "药品外调" Then
                   .ColWidth(.ColIndex("增值税率")) = 0
                   .ColWidth(.ColIndex("税金")) = 0
                ElseIf vsfList.TextMatrix(vsfList.Row, 1) = "药品外销" Then
                Else
                    .ColWidth(.ColIndex("外调价")) = 0
                    .ColWidth(.ColIndex("外调金额")) = 0
                    .ColWidth(.ColIndex("增值税率")) = 0
                    .ColWidth(.ColIndex("税金")) = 0
                End If
                If bln中药库房 Then
                    .colHidden(.ColIndex("原产地")) = False
                Else
                    .colHidden(.ColIndex("原产地")) = True
                End If
        End Select
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
    '外购入库所有权限：参数设置、基本、所有库房、登记、修改、删除、验收、冲销、单据打印
    
    If zlStr.IsHavePrivs(mstrPrivs, "付款标记") And gtype_UserSysParms.P173_经过标记付款后才能进行付款管理 = 1 And mlngMode = 模块号.外购入库 Then
        mnuEditMark.Visible = True
    Else
        mnuEditMark.Visible = False
    End If
    Select Case mlngMode
        Case 模块号.外购入库, 模块号.自制入库, 模块号.其他入库, 模块号.差价调整, 模块号.药品移库, 模块号.药品领用, 模块号.其他出库
            If Not zlStr.IsHavePrivs(mstrPrivs, "登记") Then
                mnuEditAdd.Visible = False
                mnuEditRestore.Visible = False
                tlbTool.Buttons("Add").Visible = False
            Else
                mnuEditRestore.Visible = True
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
            
            If mlngMode = 模块号.药品领用 Then
                If Not mblnStock Or Not zlStr.IsHavePrivs(mstrPrivs, "审核") Then
                    mnuEditVerify.Visible = False
                    mnuEditBill.Visible = False
                    tlbTool.Buttons("Verify").Visible = False
                End If
            Else
                If Not zlStr.IsHavePrivs(mstrPrivs, "审核") Then
                    mnuEditVerify.Visible = False
                    mnuEditBill.Visible = False
                    tlbTool.Buttons("Verify").Visible = False
                End If
            End If
            
            If Not zlStr.IsHavePrivs(mstrPrivs, "冲销") Then
                mnuEditStrike.Visible = False
                mnuEditWriteOff.Visible = False
                tlbTool.Buttons("Strike").Visible = False
                
                If mnuEditVerify.Visible = False Then
                    mnuEditLine2.Visible = False
                    tlbTool.Buttons("VerifySeparate").Visible = False
                End If
            End If
            If Not zlStr.IsHavePrivs(mstrPrivs, "单据打印") Then
                mnuFileBillPrint.Visible = False
                mnuFileBillPreview.Visible = False
            End If
        Case Else
    End Select
    
    If mlngMode = 模块号.外购入库 Then
        mnuEditVerifySelect.Visible = True
        If zlStr.IsHavePrivs(mstrPrivs, "审核") Then
            mnuEditLine0.Visible = True
            mnuEditBill.Visible = True
        Else
            mnuEditLine0.Visible = False
        End If
        mnuEditRestore.Visible = zlStr.IsHavePrivs(mstrPrivs, "退货")
        If zlStr.IsHavePrivs(mstrPrivs, "财务审核") Then
            mnuEditLine0.Visible = True
            mnuEditAcc.Visible = True
        Else
            If mnuEditBill.Visible = False Then mnuEditLine0.Visible = False
            mnuEditAcc.Visible = False
        End If
        
        If zlStr.IsHavePrivs(mstrPrivs, "药品退药计划") Then
            mnuEditHandBack.Visible = True
        Else
            mnuEditHandBack.Visible = False
        End If
        
        If zlStr.IsHavePrivs(mstrPrivs, "药品计划单批量导入") Then
            mnuEditMediPlanImport.Visible = True
        Else
            mnuEditMediPlanImport.Visible = False
        End If
                
        If zlStr.IsHavePrivs(mstrPrivs, "核查成本价") And mbln核查 Then
            mnuEditPrepare.Visible = True
            mnuEditBack.Visible = True
            mnuEditLine3.Visible = True
            Me.tlbTool.Buttons("Prepare").Visible = True
            Me.tlbTool.Buttons("Back").Visible = True
            Me.tlbTool.Buttons("PrepareSeparate").Visible = True
        End If
        mnuEditPrepare.Enabled = mbln核查
        mnuEditBack.Enabled = mbln核查
        Me.tlbTool.Buttons("Prepare").Enabled = mbln核查
        Me.tlbTool.Buttons("Back").Enabled = mbln核查
        
    ElseIf mlngMode = 模块号.药品移库 Then
        mnuEditRestore.Visible = False
        mnuEditLine2.Visible = False
        If zlStr.IsHavePrivs(mstrPrivs, "发送") Then
            mnuEditPreparePhysic.Visible = True
            mnuEditSendPhysic.Visible = True
            mnuEditBack.Visible = True
            mnuEditLine3.Visible = True
            tlbTool.Buttons("PreparePhysic").Visible = True
            tlbTool.Buttons("SendPhysic").Visible = True
            tlbTool.Buttons("Back").Visible = True
            tlbTool.Buttons("PrepareSeparate").Visible = True
        End If
        If Not zlStr.IsHavePrivs(mstrPrivs, "审核") And _
           Not zlStr.IsHavePrivs(mstrPrivs, "冲销") Then
            TabShow.TabVisible(1) = False
        End If
    Else
        mnuEditBill.Visible = False
        mnuEditAcc.Visible = False
        mnuEditRestore.Visible = False
        mnuEditLine0.Visible = False
        mnuEditHandBack.Visible = False
    End If
End Sub

Private Sub Cmd查阅_Click()
    Call mnuEditDisplay_Click
End Sub

Private Sub Form_Activate()
    If vsfList.Visible = True Then
        vsfList.SetFocus
'        If vsfList.rows > 1 Then
'            vsfList.Row = 1
'        End If
        If vsfDetail.rows > 1 Then
            vsfDetail.Row = 1
        End If
    End If
End Sub

Private Sub Form_Load()
    '恢复设置
    Dim dateCurrentDate As Date
    mbln操作员限制 = Not zlStr.IsHavePrivs(mstrPrivs, "所有库房")
    mblnViewCost = zlStr.IsHavePrivs(mstrPrivs, "查看成本价")
    
    Me.Caption = mstrTitle
    dateCurrentDate = Sys.Currentdate
    lblRange.Caption = "查询范围:" & Format(dateCurrentDate, "yyyy年MM月dd日") & "至" & Format(dateCurrentDate, "yyyy年MM月dd日")
    
    mnuViewLine3.Visible = (mlngMode = 模块号.外购入库 Or mlngMode = 模块号.其他入库)
    mnuViewColDefine.Visible = (mlngMode = 模块号.外购入库 Or mlngMode = 模块号.其他入库)
    
    mnuEditLine0.Visible = (mlngMode = 模块号.外购入库)
    
    mnuFileCodePrint.Visible = (mlngMode = 模块号.外购入库 Or mlngMode = 模块号.其他入库 Or mlngMode = 模块号.药品领用 Or mlngMode = 模块号.药品移库 Or mlngMode = 模块号.其他出库)
    mnuEditCodePrintLine.Visible = mnuEditAllCodePrint.Visible
    
    TabShow.Visible = (mlngMode = 模块号.药品移库)
    mblnDo = Val(zlDataBase.GetPara("使用个性化风格")) <> 0
    
    lbl1.Caption = ""
    lbl2.Caption = ""
    lbl3.Caption = ""
    lbl4.Caption = ""
    lbl2.Left = lbl1.Left + lbl1.Width + 2500
    lbl3.Left = lbl2.Left + lbl2.Width + 2500
    lbl4.Left = lbl3.Left + lbl3.Width + 2500
    If mblnViewCost = False Then
        lbl1.Visible = False
        lbl3.Visible = False
        lbl2.Left = lbl1.Left
        lbl4.Left = lbl2.Left + lbl2.Width + 2500
    End If
    
    Me.Top = (Screen.Height - Me.Height) / 2
    If mlngMode = 模块号.差价调整 Or mlngMode = 模块号.自制入库 Then
        mnuEditWriteOff.Visible = False
    End If
    
    If mlngMode = 模块号.外购入库 Then
        '外购业务外挂部件
        Call zlPlugIn_Ini(glngSys, glngModul, mobjPlugIn)
        
        '外挂部件有扩展功能
        Call zlPlugIn_SetVBMenu(glngSys, glngModul, mobjPlugIn, Me)
        
        '外挂部件有扩展功能
        Call zlPlugIn_SetVBToolbar(glngSys, glngModul, mobjPlugIn, Me, tlbTool, "PlugItem", "PlugInSeparator")
    End If
    
    staThis.Panels(2).Picture = picColor
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
        .Height = 360
        .Left = 0
        .Width = cbrTool.Width
    End With
    
    With TabShow
        .Left = 0
        .Top = cbrTool.Height
    End With
    
    With vsfList
        .Top = IIf(cbrTool.Visible, cbrTool.Height, 0) + IIf(TabShow.Visible, TabShow.Height, 0)
        .Left = 0
        .Width = cbrTool.Width
        .Height = picSeparate_s.Top - .Top
    End With
    
    With Cmd查阅
        .Left = Me.ScaleWidth - .Width - 100
        .Top = vsfList.Top + vsfList.Height + 30
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
    
    If mlngMode <> 模块号.外购入库 And mlngMode <> 模块号.药品移库 And mlngMode <> 模块号.药品领用 Then
        picColor3.Visible = False
        lblColor3.Visible = False
        picColor.Width = lblColor2.Left + lblColor2.Width + 20
    Else
        If mlngMode = 模块号.药品移库 Or mlngMode = 模块号.药品领用 Then
            lblColor3.Caption = "未审核冲销"
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, mstrTitle
    
    Call zlPlugIn_Unload(mobjPlugIn)
    
    mblnDo = False
End Sub

Private Sub mnuEditAcc_Click()
    Dim strNo As String
    Dim blnSuccess As Boolean
    Dim int记录状态 As Integer
    
    If cboStock.ListIndex = -1 Then
        MsgBox "请选择库房！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    With vsfList
        strNo = .TextMatrix(.Row, 0)
        int记录状态 = .TextMatrix(.Row, .Cols - 3)
        frmPurchaseCard.ShowCard Me, strNo, 编辑.财务审核, int记录状态, blnSuccess
        
        If blnSuccess = True Then
            mintListRow = vsfList.Row
            mnuViewRefresh_Click
        End If
    End With
End Sub

Private Sub mnuEditAdd_Click()
    Dim strNo As String
    Dim blnSuccess As Boolean
    Dim rsStock As ADODB.Recordset
    
    '检查本期是否已经审核结存，如果未审核结存则不能进行新增业务操作
'    If CheckIsAccount(Val(cboStock.ItemData(cboStock.ListIndex))) = False Then
'        Exit Sub
'    End If
    If cboStock.ListIndex = -1 Then
        MsgBox "请选择库房！", vbInformation, gstrSysName
        Exit Sub
    End If
    strNo = ""
    '新增
    Select Case mlngMode
        Case 模块号.外购入库
            frmPurchaseCard.ShowCard Me, strNo, 编辑.新增, , blnSuccess
        Case 模块号.自制入库
            frmSelfMakeCard.ShowCard Me, strNo, 编辑.新增, , blnSuccess
        Case 模块号.其他入库
            frmOtherInputCard.ShowCard Me, strNo, 编辑.新增, , blnSuccess
        Case 模块号.差价调整
            frmDiffPriceAdjustCard.ShowCard Me, strNo, 编辑.新增, , blnSuccess
        Case 模块号.药品移库
            Set rsStock = ReturnSQL(Val(cboStock.ItemData(cboStock.ListIndex)), "mnuEditAdd_Click", True, 模块号.药品移库)
            If rsStock.EOF Then
                '请设置流向控制
                MsgBox "该库房未设置药品流向控制，不能进行新增单据！", vbOKOnly, gstrSysName
                Exit Sub
            End If
            
            frmTransferCard.ShowCard Me, strNo, 编辑.新增, , blnSuccess
        Case 模块号.药品领用
            frmDrawCard.ShowCard Me, strNo, 编辑.新增, mblnStock, , 0, blnSuccess
        Case 模块号.其他出库
            frmOtherOutputCard.ShowCard Me, strNo, 编辑.新增, , blnSuccess
    End Select
    
    If blnSuccess = True Then
        mintListRow = vsfList.Row + 1
        mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditApplyStrike_Click()
    mnuEditApplyStrike.Tag = "1"
    Call mnuEditStrike_Click
End Sub

Private Sub mnuEditBack_Click()
    Dim strNo As String
    On Error GoTo ErrHand
    '1、移库：回退上一次状态，如果未备药直接退出（只能从发送回退到备药，由备药回退到非备药）
    '2、外购：撤销核查
    strNo = vsfList.TextMatrix(vsfList.Row, 0)
    
    If TestDelete(strNo) Then
        MsgBox "该单据已被删除！", vbInformation, gstrSysName
        mintListRow = vsfList.Row
        mnuViewRefresh_Click
        Exit Sub
    End If
    
    If TestVerify(strNo) Then
        MsgBox "该单据已被审核！", vbInformation, gstrSysName
        mintListRow = vsfList.Row + 1
        mnuViewRefresh_Click
        Exit Sub
    End If
    
    Select Case mlngMode
    Case 模块号.外购入库   '外购入库撤销核查
        gstrSQL = "Zl_药品外购_CancelCheck('" & strNo & "')"
    Case 模块号.药品移库   '移库退回
        gstrSQL = "ZL_药品移库_BACK('" & strNo & "')"
    End Select
        
    Call zlDataBase.ExecuteProcedure(gstrSQL, "回退")
    mintListRow = vsfList.Row
    Call mnuViewRefresh_Click
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub mnuEditBill_Click()
    Dim strNo As String
    Dim blnSuccess As Boolean
    Dim int记录状态 As Integer
    
    With vsfList
        strNo = .TextMatrix(.Row, 0)
        int记录状态 = .TextMatrix(.Row, .Cols - 3)
        frmPurchaseCard.ShowCard Me, strNo, 编辑.修改发票, int记录状态, blnSuccess
        
        If blnSuccess = True Then
            mintListRow = vsfList.Row
            mnuViewRefresh_Click
        End If
    End With
    
End Sub

Private Sub mnuEditDeliveryInvoice_Click()
    gobjDrugPurchase.DeliveryInvoice gcnOracle
End Sub

Private Sub mnuEditHandBack_Click()
    frmHandBackPlan.ShowForm Me, mlng库房ID, mintUnit
End Sub

Private Sub mnuEditMark_Click()
    Call frmPurchaseMark.ShowME(mStr库房, cboStock.ListIndex, Me, mstrPrivs)
End Sub

Private Sub mnuEditMediPlanImport_Click()
    frmMediPlanImport.ShowCard Me, Val(cboStock.ItemData(cboStock.ListIndex))
    mnuViewRefresh_Click
End Sub

Private Sub mnuEditPrepare_Click()
    Dim strNo As String
    Dim blnSuccess As Boolean
    Dim rsTemp As New ADODB.Recordset
    '针对外购入库是核查（仅允许修改成本价）
    '针对移库单是配送
    strNo = vsfList.TextMatrix(vsfList.Row, 0)
    If Trim(strNo) = "" Then Exit Sub
    
    Select Case mlngMode
    Case 模块号.外购入库
        frmPurchaseCard.ShowCard Me, strNo, 编辑.核查, vsfList.TextMatrix(vsfList.Row, 外购主表.记录状态), blnSuccess
    Case 模块号.药品移库
        If TestPrepare(strNo) Then
            MsgBox "此移库单[" & strNo & "]的所有药品已经配送！", vbInformation, gstrSysName
            Exit Sub
        End If
        
    End Select
    
    If blnSuccess = True Then
        mintListRow = vsfList.Row
        mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditPreparePhysic_Click()
    Dim strNo As String
    Dim strCheckString As String
    
    On Error GoTo ErrHand
    strNo = vsfList.TextMatrix(vsfList.Row, 0)
    If Trim(strNo) = "" Then Exit Sub
    
    strCheckString = CheckBill(Trim(strNo))
    If strCheckString <> "" Then
        MsgBox strCheckString, vbInformation, gstrSysName
        mintListRow = vsfList.Row + 1
        Call mnuViewRefresh_Click
        Exit Sub
    End If
    
    gstrSQL = "zl_药品移库_PREPARE('" & strNo & "','" & UserInfo.用户姓名 & "')"
    Call zlDataBase.ExecuteProcedure(gstrSQL, "备药")
    
    mintListRow = vsfList.Row
    
    Call mnuViewRefresh_Click
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Public Function CheckBill(ByVal strNo As String) As String
    Dim rs As New ADODB.Recordset
    
    CheckBill = ""
    On Error GoTo errHandle
    gstrSQL = " Select 审核日期,配药日期,配药人 From 药品收发记录 " & _
              " Where 单据=6 And NO=[1] And 记录状态=1 And RowNum=1 "
    Set rs = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[检查单据]", strNo)
            
    With rs
        '返回空，表示已经删除
        If .EOF Then
            CheckBill = "该单据已经被其他操作员删除！"
        ElseIf Not IsNull(!审核日期) Then
            CheckBill = "该单据已经被其他操作员审核！"
        ElseIf Not IsNull(!配药日期) Then
            CheckBill = "该单据已经被其他操作员发送！"
        ElseIf Not IsNull(!配药人) Then
            CheckBill = "该单据已经被其他操作员备药！"
        End If
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Check已付款记录(ByVal strNo As String) As Boolean
    Dim strsql As String
    Dim rsCheck As ADODB.Recordset
    
    On Error GoTo errHandle
    strsql = "Select count(Id) 已付款 From 应付记录 Where 收发id In(Select Id From 药品收发记录 Where 单据=5 And No=[1]) And nvl(付款序号,0)>0 "
    Set rsCheck = zlDataBase.OpenSQLRecord(strsql, Me.Caption & "[检查应付记录]", strNo)
    
    Check已付款记录 = (rsCheck!已付款 > 0)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub mnuEditPric_Click()
    frmDiffPriceAdjustCard.ShowCard Me, "", 编辑.新增, , False, 2
End Sub

Private Sub mnuEditRestore_Click()
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    If cboStock.ListIndex = -1 Then
        MsgBox "请选择库房！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Call frmPurchaseCard.ShowCard(Me, strNo, 编辑.药库退货, , blnSuccess)
    If blnSuccess Then
        mintListRow = vsfList.Row + 1
        mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditSendPhysic_Click()
    Dim strNo As String
    Dim blnSuccess As Boolean
    On Error GoTo ErrHand
    strNo = vsfList.TextMatrix(vsfList.Row, 0)
    If Trim(strNo) = "" Then Exit Sub
    
    Call frmTransferCard.ShowCard(Me, strNo, 编辑.发送, 1, blnSuccess)
    If blnSuccess Then
        mintListRow = vsfList.Row
        mnuViewRefresh_Click
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub mnuEditVerify_Click()
    '验收
    Dim strNo As String
    Dim blnSuccess As Boolean
    Dim int审核方式 As Integer  '1-审核已申请的冲销单；
    
    With vsfList
        strNo = .TextMatrix(.Row, 0)
        Select Case mlngMode
            Case 模块号.外购入库
                If mbln核查 Then
                    If Not TestPrepare(strNo) Then
                        MsgBox "该单据还未通过核查，不允许审核！", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
                frmPurchaseCard.ShowCard Me, strNo, 编辑.审核, .TextMatrix(.Row, 外购主表.记录状态), blnSuccess
            Case 模块号.自制入库
                frmSelfMakeCard.ShowCard Me, strNo, 编辑.审核, .TextMatrix(.Row, .Cols - 2), blnSuccess
            Case 模块号.其他入库
                frmOtherInputCard.ShowCard Me, strNo, 编辑.审核, .TextMatrix(.Row, .Cols - 2), blnSuccess
            Case 模块号.差价调整
                frmDiffPriceAdjustCard.ShowCard Me, strNo, 编辑.审核, .TextMatrix(.Row, .Cols - 3), blnSuccess, IIf(Val(vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 1)) = 0, 1, 2)
            Case 模块号.药品移库
                frmTransferCard.ShowCard Me, strNo, 编辑.审核, .TextMatrix(.Row, .Cols - 2), blnSuccess
            Case 模块号.药品领用
                frmDrawCard.ShowCard Me, strNo, 编辑.审核, mblnStock, .TextMatrix(.Row, .Cols - 4), 0, blnSuccess
            Case 模块号.其他出库
                frmOtherOutputCard.ShowCard Me, strNo, 编辑.审核, .TextMatrix(.Row, .Cols - 2), blnSuccess
        End Select
        
    End With
    If blnSuccess = True Then
        mintListRow = vsfList.Row
        mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditDel_Click()
    '删除
    Dim strBillNo As String
    Dim intRow As Integer
    Dim strTitle As String
    Dim intReturn As Integer
    Dim intRecord As Integer
    Dim rsCheck As New ADODB.Recordset
     
    With vsfList
        Select Case mlngMode
            Case 模块号.外购入库
                strTitle = "外购入库单"
            Case 模块号.自制入库
                strTitle = "自制入库单"
            Case 模块号.其他入库
                strTitle = "其他入库单"
            Case 模块号.差价调整
                strTitle = "库存差价调整单"
            Case 模块号.药品移库
                strTitle = "药品移库单"
            Case 模块号.药品领用
                strTitle = "药品领用单"
            Case 模块号.其他出库
                strTitle = "药品其他出库单"
        End Select
        
        On Error GoTo errHandle
        intRow = .Row
        strBillNo = .TextMatrix(intRow, 0)
        intReturn = MsgBox("你确实要删除单据号为“" & strBillNo & "”的" & strTitle & "吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
        intRecord = .rows - 1
        If intReturn = vbYes Then
            Select Case mlngMode
                Case 模块号.外购入库
                    gstrSQL = "zl_药品外购_Delete('" & strBillNo & "')"
                Case 模块号.自制入库
                    gstrSQL = "zl_自制入库_Delete('" & strBillNo & "')"
                Case 模块号.其他入库
                    gstrSQL = "zl_药品其他入库_Delete('" & strBillNo & "')"
                Case 模块号.差价调整
                    gstrSQL = "zl_药品库存差价调整_Delete('" & strBillNo & "')"
                Case 模块号.药品移库
                    If Val(.TextMatrix(.Row, .Cols - 2)) = 1 Then
                        '已备药（填写了配药人）或已发送的单据，不允许入库方修改此类单据
                        If TestDelete(strBillNo) Then
                            MsgBox "该单据已被删除！", vbInformation, gstrSysName
                            mintListRow = vsfList.Row
                            mnuViewRefresh_Click
                            Exit Sub
                        End If
                        If TestPrepare(strBillNo) Then
                            MsgBox "已备药和发送的单据不允许删除！", vbInformation, gstrSysName
                            mintListRow = vsfList.Row + 1
                            mnuViewRefresh_Click
                            Exit Sub
                        End If
                    End If
                    
'                    If Is申领(StrBillNo) Then
'                        If Not zlStr.IsHavePrivs(mstrPrivs, "允许修改申领单") Then
'                            MsgBox "你没有权限修改申领单！", vbInformation, gstrSysName
'                            Exit Sub
'                        End If
'                    End If
                    
                    '传入记录状态，是为了可能是删除未审核的冲销申请单据
                    gstrSQL = "zl_药品移库_Delete('" & strBillNo & "'," & Val(.TextMatrix(.Row, .Cols - 2)) & " )"

                Case 模块号.药品领用
                    gstrSQL = "zl_药品领用_Delete('" & strBillNo & "'," & Val(.TextMatrix(.Row, .Cols - 4)) & " )"
                Case 模块号.其他出库
                    gstrSQL = "zl_药品其他出库_Delete('" & strBillNo & "')"
                Case Else
                
            End Select
            If gstrSQL = "" Then Exit Sub
            
            Call zlDataBase.ExecuteProcedure(gstrSQL, Me.Caption)
            
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
    If intRecord = 0 Then
        lbl1.Caption = ""
        lbl2.Caption = ""
        lbl3.Caption = ""
    End If
    mintListRow = vsfList.Row
    mnuViewRefresh_Click
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
            Case 模块号.外购入库
                frmPurchaseCard.ShowCard Me, strNo, 编辑.查阅, .TextMatrix(.Row, 外购主表.记录状态)
                
            Case 模块号.自制入库
                frmSelfMakeCard.ShowCard Me, strNo, 编辑.查阅, .TextMatrix(.Row, .Cols - 2)
            Case 模块号.其他入库
                frmOtherInputCard.ShowCard Me, strNo, 编辑.查阅, .TextMatrix(.Row, .Cols - 2)
            Case 模块号.差价调整
                frmDiffPriceAdjustCard.ShowCard Me, strNo, 编辑.查阅, .TextMatrix(.Row, .Cols - 3), False, IIf(Val(vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 1)) = 0, 1, 2)
            Case 模块号.药品移库
                frmTransferCard.ShowCard Me, strNo, 编辑.查阅, .TextMatrix(.Row, .Cols - 2)
            Case 模块号.药品领用
                frmDrawCard.ShowCard Me, strNo, 编辑.查阅, mblnStock, .TextMatrix(.Row, .Cols - 4), 0
            Case 模块号.其他出库
                frmOtherOutputCard.ShowCard Me, strNo, 编辑.查阅, .TextMatrix(.Row, .Cols - 2)
            Case Else
        
        End Select
        
    End With
    
End Sub

Private Sub mnuEditStrike_Click()
    Dim blnPurchase As Boolean, blnRefresh As Boolean
    Dim str模块号 As String
    
    '如果是外购(blnPurchase为真)，则直接进入冲销
    '询问是否冲销(blnPurchase为提示框返回值)，是则进入冲销
    
    If cboStock.ListIndex = -1 Then
        MsgBox "请选择库房！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    str模块号 = 模块号.外购入库 & "," & 模块号.其他入库 & "," & 模块号.药品移库 & "," & 模块号.药品领用 & "," & 模块号.其他出库
    blnPurchase = (InStr(1, str模块号, mlngMode) <> 0)
    With vsfList
        If Not blnPurchase Then
            blnPurchase = (MsgBox("你确实要全部冲销单据号为“" & .TextMatrix(.Row, 0) & "”的单据吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes)
        End If
        If blnPurchase Then
            blnRefresh = StrikeSave
            If blnRefresh Then
                mintListRow = vsfList.Row
                mnuViewRefresh_Click
            End If
        End If
    End With
End Sub

Private Function StrikeSave() As Boolean
    Dim blnSuccess As Boolean
    Dim int处理方式 As Integer
    Dim n As Integer
    Dim int单据 As Integer
    
    StrikeSave = False
    With vsfList
        Select Case mlngMode
            Case 模块号.外购入库
                frmPurchaseCard.ShowCard Me, .TextMatrix(.Row, 0), 编辑.冲销, vsfList.TextMatrix(vsfList.Row, 外购主表.记录状态), blnSuccess
                StrikeSave = blnSuccess
                Exit Function
            Case 模块号.自制入库
                mint库存检查 = MediWork_GetCheckStockRule(cboStock.ItemData(cboStock.ListIndex))
                
                '检查可用数量是否足够，参数设置为不检查库存时不进行（传入单据，药名，库存检查，单据号，序号）
                If mint库存检查 <> 0 And .TextMatrix(.Row, 0) <> "" Then
                    For n = 1 To vsfDetail.rows - 1
                        If vsfDetail.TextMatrix(n, 0) <> "" Then
                            If CheckStrickUsable(单据号.自制入库, 0, 0, vsfDetail.TextMatrix(n, vsfDetail.ColIndex("药品信息")), _
                                0, vsfDetail.TextMatrix(n, vsfDetail.ColIndex("数量")), mint库存检查, Trim(.TextMatrix(.Row, 0)), vsfDetail.TextMatrix(n, vsfDetail.ColIndex("序号"))) = False Then
                                Exit Function
                            End If
                        End If
                    Next
                End If
                gstrSQL = "zl_自制入库_Strike('" & .TextMatrix(.Row, 0) & "','" & UserInfo.用户姓名 & "')"
            Case 模块号.其他入库
                frmOtherInputCard.ShowCard Me, .TextMatrix(.Row, 0), 编辑.冲销, vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 2), blnSuccess
                StrikeSave = blnSuccess
                Exit Function
            Case 模块号.差价调整
                If Val(.TextMatrix(.Row, .Cols - 1)) = 1 Then
                    If Check已付款记录(.TextMatrix(.Row, 0)) Then
                        MsgBox "药品已付款，不能冲销！", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
                gstrSQL = "zl_药品库存差价调整_Strike('" & .TextMatrix(.Row, 0) & "','" & UserInfo.用户姓名 & "')"
            Case 模块号.药品移库
                If mnuEditStrike.Caption = "申请冲销(&R)" Then
                    int处理方式 = 1
                ElseIf mnuEditStrike.Caption = "审核冲销(&K)" Then
                    int处理方式 = 2
                Else
                    int处理方式 = 0
                End If
               
                frmTransferCard.ShowCard Me, .TextMatrix(.Row, 0), 编辑.冲销, vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 2), blnSuccess, int处理方式
                StrikeSave = blnSuccess
                Exit Function
            Case 模块号.药品领用
                If mnuEditApplyStrike.Tag = "1" Then
                    int处理方式 = 1
                    mnuEditApplyStrike.Tag = "0"
                ElseIf mnuEditVerifyStrike.Tag = "1" Then
                    int处理方式 = 2
                    mnuEditVerifyStrike.Tag = "0"
                Else
                    int处理方式 = 0
                End If
            
                frmDrawCard.ShowCard Me, .TextMatrix(.Row, 0), 编辑.冲销, mblnStock, vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 4), 0, blnSuccess, int处理方式
                StrikeSave = blnSuccess
                Exit Function
            Case 模块号.其他出库
                frmOtherOutputCard.ShowCard Me, .TextMatrix(.Row, 0), 编辑.冲销, vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 2), blnSuccess
                StrikeSave = blnSuccess
                Exit Function
            Case Else
            
        End Select
        
        On Error GoTo errHandle
        
        Call zlDataBase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        If mlngMode = 模块号.自制入库 Or mlngMode = 模块号.差价调整 Then
            If mlngMode = 模块号.自制入库 Then
                int单据 = 单据号.自制入库
            ElseIf mlngMode = 模块号.差价调整 Then
                int单据 = 单据号.差价调整
            End If
            '提示停用药品
            Call CheckStopMedi(int单据 & "|" & .TextMatrix(.Row, 0))
        End If
    End With
    
    StrikeSave = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    
    'MsgBox "存盘失败！", vbInformation, gstrSysName
    Call SaveErrLog

End Function

Private Sub mnuEditModify_Click()
    '修改
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    blnSuccess = False
    With vsfList
        If cboStock.ListIndex = -1 Then
            MsgBox "请选择库房！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        strNo = .TextMatrix(.Row, 0)
        
        Select Case mlngMode
            Case 模块号.外购入库
                frmPurchaseCard.ShowCard Me, strNo, 编辑.修改, vsfList.TextMatrix(vsfList.Row, 外购主表.记录状态), blnSuccess
            Case 模块号.自制入库
                frmSelfMakeCard.ShowCard Me, strNo, 编辑.修改, vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 2), blnSuccess
            Case 模块号.其他入库
                frmOtherInputCard.ShowCard Me, strNo, 编辑.修改, vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 2), blnSuccess
            Case 模块号.差价调整
                frmDiffPriceAdjustCard.ShowCard Me, strNo, 编辑.修改, vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 3), blnSuccess, IIf(Val(vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 1)) = 0, 1, 2)
            Case 模块号.药品移库
                '已备药（填写了配药人）或已发送的单据，不允许入库方修改此类单据
                If TabShow.Tab = 1 Then
                    If TestPrepare(strNo) Then
                        MsgBox "已发送的单据不允许修改！", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
                frmTransferCard.ShowCard Me, strNo, 编辑.修改, vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 2), blnSuccess
            Case 模块号.药品领用
                frmDrawCard.ShowCard Me, strNo, 编辑.修改, mblnStock, vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 4), 0, blnSuccess
            Case 模块号.其他出库
                frmOtherOutputCard.ShowCard Me, strNo, 编辑.修改, vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 2), blnSuccess
        End Select
        If blnSuccess = True Then
            mintListRow = vsfList.Row
            mnuViewRefresh_Click
        End If
    End With
End Sub

Private Sub mnuEditVerifySelect_Click()
    frmPurchaseVerifySelect.ShowME Me, mStr库房, cboStock.ListIndex
End Sub

Private Sub mnuEditVerifyStrike_Click()
    mnuEditVerifyStrike.Tag = "1"
    Call mnuEditStrike_Click
End Sub

Private Sub mnuEditWriteOff_Click()
    Dim strStock As String
    Dim i As Integer
    
    
    If cboStock.ListIndex = -1 Then
        MsgBox "请选择库房！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    With Me.cboStock
        For i = 0 To .ListCount - 1
            strStock = strStock & .List(i) & "," & .ItemData(i) & "|"
        Next
    End With
    
    Call frm批量冲销.ShowME(mlngMode, Me, strStock, Me.cboStock.ListIndex)
End Sub


Private Sub mnuFileAllCodePrint_Click()
    If Trim(vsfList.TextMatrix(vsfList.Row, 0)) = "" Or vsfList.rows <= 1 Then Exit Sub
    CodePrint vsfList.TextMatrix(vsfList.Row, 0)
End Sub

Private Sub mnuEditAllCodePrint_Click()
    If Trim(vsfList.TextMatrix(vsfList.Row, 0)) = "" Or vsfList.rows <= 1 Then Exit Sub
    CodePrint vsfList.TextMatrix(vsfList.Row, 0)
End Sub

Private Sub mnuFileSelCodePrint_Click()
    If Trim(vsfList.TextMatrix(vsfList.Row, 0)) = "" Or vsfList.rows <= 1 Then Exit Sub
    CodePrint Val(vsfDetail.TextMatrix(vsfDetail.Row, vsfDetail.ColIndex("药品ID")))
End Sub

Private Sub CodePrint(ByVal varPar As Variant)
'功能：打印要品条码
'传参：varPar是long型则打印对应药品条码；是String型则根据单据号打印单据中的药品条码
    Dim rsTemp As New ADODB.Recordset
    Dim int单据 As Integer
    Dim strReport As String

    On Error GoTo errHandle
    
    If Not zlStr.IsHavePrivs(mstrPrivs, "药品条码打印") Then
        MsgBox "对不起，你没有该权限！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Select Case mlngMode
        Case 模块号.外购入库
            int单据 = 1
            strReport = "ZL1_INSIDE_1300_1"
        Case 模块号.其他入库
            int单据 = 4
            strReport = "ZL1_INSIDE_1302_1"
        Case 模块号.其他出库
            int单据 = 11
            strReport = "ZL1_INSIDE_1306_1"
        Case 模块号.药品移库
            int单据 = 6
            strReport = "ZL1_INSIDE_1304_1"
        Case 模块号.药品领用
            int单据 = 7
            strReport = "ZL1_INSIDE_1305_2"
    End Select

    
    
    If TypeName(varPar) = "String" Then '打印整张单据条码
        gstrSQL = "select distinct 药品ID from 药品收发记录 where 单据 = [2] and  NO = [1] order by 药品ID"
        
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "药品条码打印", varPar, int单据)
        
        Do While Not rsTemp.EOF
            ReportOpen gcnOracle, glngSys, strReport, Me, "药品=" & rsTemp!药品id, 2
            rsTemp.MoveNext
        Loop
        
    Else '打印对应药品条码
        If varPar = 0 Then Exit Sub
        ReportOpen gcnOracle, glngSys, strReport, Me, "药品=" & varPar, 2
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuFileBillPreview_Click()
    Dim int单位系数 As Integer
    Dim bln退库单 As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
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
        
        Select Case mlngMode
            Case 模块号.外购入库
                '判断是否是退库单
                gstrSQL = "Select Nvl(发药方式,0) 标志 From 药品收发记录 Where NO=[1] And 记录状态=[2] And Rownum<2"
                Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[判断是否是退库单]", .TextMatrix(.Row, 0), Val(.TextMatrix(.Row, 外购主表.记录状态)))
                
                bln退库单 = (rsTemp!标志 = 1)
                ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_" & 模块号.外购入库, "zl8_bill_" & 模块号.外购入库), Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & .TextMatrix(.Row, 外购主表.记录状态), "单位系数=" & int单位系数, IIf(bln退库单, "药品退货单", "药品外购入库单"), 1
            Case 模块号.自制入库
                ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_" & 模块号.自制入库, "zl8_bill_" & 模块号.自制入库), Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & .TextMatrix(.Row, .Cols - 2), "单位系数=" & int单位系数, 1
            Case 模块号.其他入库
                ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_" & 模块号.其他入库, "zl8_bill_" & 模块号.其他入库), Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & .TextMatrix(.Row, .Cols - 2), "单位系数=" & int单位系数, 1
            Case 模块号.差价调整
                ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_" & 模块号.差价调整, "zl8_bill_" & 模块号.差价调整), Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & .TextMatrix(.Row, .Cols - 3), "单位系数=" & int单位系数, 1
            Case 模块号.药品移库
                ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_" & 模块号.药品移库, "zl8_bill_" & 模块号.药品移库), Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & .TextMatrix(.Row, .Cols - 2), "单位系数=" & int单位系数, 1
            Case 模块号.药品领用
                ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_" & 模块号.药品领用, "zl8_bill_" & 模块号.药品领用), Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & .TextMatrix(.Row, .Cols - 4), "单位系数=" & int单位系数, 1
            Case 模块号.其他出库
                ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_" & 模块号.其他出库, "zl8_bill_" & 模块号.其他出库), Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & .TextMatrix(.Row, .Cols - 2), "单位系数=" & int单位系数, 1
            Case Else
            
        End Select
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuFileBillPrint_Click()
    Dim int单位系数 As Integer
    Dim bln退库单 As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
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
        
        Select Case mlngMode
            Case 模块号.外购入库
                '判断是否是退库单
                gstrSQL = "Select Nvl(发药方式,0) 标志 From 药品收发记录 Where NO=[1] And 记录状态=[2] And Rownum<2"
                Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[判断是否是退库单]", .TextMatrix(.Row, 0), Val(.TextMatrix(.Row, .Cols - 3)))

                bln退库单 = (rsTemp!标志 = 1)
                ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_" & 模块号.外购入库, "zl8_bill_" & 模块号.外购入库), Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & .TextMatrix(.Row, 外购主表.记录状态), "单位系数=" & int单位系数, IIf(bln退库单, "药品退货单", "药品外购入库单"), 2
            Case 模块号.自制入库
                ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_" & 模块号.自制入库, "zl8_bill_" & 模块号.自制入库), Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & .TextMatrix(.Row, .Cols - 2), "单位系数=" & int单位系数, 2
            Case 模块号.其他入库
                ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_" & 模块号.其他入库, "zl8_bill_" & 模块号.其他入库), Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & .TextMatrix(.Row, .Cols - 2), "单位系数=" & int单位系数, 2
            Case 模块号.差价调整
                ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_" & 模块号.差价调整, "zl8_bill_" & 模块号.差价调整), Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & .TextMatrix(.Row, .Cols - 3), "单位系数=" & int单位系数, 2
            Case 模块号.药品移库
                ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_" & 模块号.药品移库, "zl8_bill_" & 模块号.药品移库), Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & .TextMatrix(.Row, .Cols - 2), "单位系数=" & int单位系数, 2
            Case 模块号.药品领用
                ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_" & 模块号.药品领用, "zl8_bill_" & 模块号.药品领用), Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & .TextMatrix(.Row, .Cols - 4), "单位系数=" & int单位系数, 2
            Case 模块号.其他出库
                ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_" & 模块号.其他出库, "zl8_bill_" & 模块号.其他出库), Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & .TextMatrix(.Row, .Cols - 2), "单位系数=" & int单位系数, 2
            Case Else
            
        End Select
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
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
    Dim int查询天数 As Integer
    Dim dateCurrentDate As Date
    
    frm参数设置.设置参数 Me, mstrPrivs, Me.Tag
    
    Call GetDrugDigit(mlng库房ID, Me.Tag, mintUnit, mintShowCostDigit, mintShowPriceDigit, mintShowNumberDigit, mintShowMoneyDigit)
    
    '重新组织格式化串
    mstrCostFormat = "'999999999990." & String(mintShowCostDigit, "0") & "'"
    mstrPriceFormat = "'999999999990." & String(mintShowPriceDigit, "0") & "'"
    mstrNumberFormat = "'999999999990." & String(mintShowNumberDigit, "0") & "'"
    mstrMoneyFormat = "'999999999990." & String(mintShowMoneyDigit, "0") & "'"
    
    int移库处理流程 = Val(zlDataBase.GetPara("移库流程", glngSys, 模块号.药品移库))
    mint冲销申请 = Val(zlDataBase.GetPara("冲销申请", glngSys, 模块号.药品移库))
    
    dateCurrentDate = Sys.Currentdate
    int查询天数 = Val(zlDataBase.GetPara("查询天数", glngSys, mlngMode, 1)) - 1
    strStart = Format(DateAdd("d", -int查询天数, dateCurrentDate), "yyyy-MM-dd")
    strEnd = Format(dateCurrentDate, "yyyy-MM-dd")
    
    SQLCondition.date填制时间开始 = CDate(Format(strStart, "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date填制时间结束 = CDate(Format(strEnd, "yyyy-mm-dd") & " 23:59:59")
    
    Call SetMenu
    Call GetList(mstrFind)
End Sub

Private Sub mnuFilePreView_Click()
    '打印预览
    Dim lngCurRow As Long
    
    lngCurRow = vsfList.Row
    vsfList.Redraw = flexRDNone
    subPrint 2
    vsfList.Row = lngCurRow
    vsfList.Redraw = flexRDDirect
    vsfList.Col = 0
    vsfList.ColSel = vsfList.Cols - 1
    
End Sub

Private Sub mnuFilePrint_Click()
    '打印
    Dim lngCurRow As Long
    
    lngCurRow = vsfList.Row
    vsfList.Redraw = flexRDNone
    subPrint 1
    vsfList.Row = lngCurRow
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
    Dim StrWinName As String
    With vsfList
        Select Case mlngMode
            Case 模块号.外购入库
                StrWinName = "frmMainList1"
            Case 模块号.自制入库
                StrWinName = "frmMainList2"
            Case 模块号.其他入库
                StrWinName = "frmMainList3"
            Case 模块号.差价调整
                StrWinName = "frmMainList4"
            Case 模块号.药品移库
                StrWinName = "frmMainList5"
            Case 模块号.药品领用
                StrWinName = "frmMainList6"
            Case 模块号.其他出库
                StrWinName = "frmMainList7"
        End Select
    End With
    Call ShowHelp(App.ProductName, Me.hWnd, StrWinName)
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
    '默认参数：模块号.外购入库(外购入库)- 药品=药品id，库房=库房id，供应商=供应商id，产地=产地名称，NO=入库单NO
    '          模块号.自制入库(自制入库)- 药品=药品id，库房=库房id，开始时间=填制开始时间，结束时间=填制结束时间，NO=入库单NO
    '          模块号.其他入库(其他入库)- 药品=药品id，库房=库房id，产地=产地名称，开始时间=填制开始时间，结束时间=填制结束时间，NO=入库单NO
    '          模块号.差价调整(库存差价调整)- 药品=药品id，库房=库房id，开始时间=填制开始时间，结束时间=填制结束时间，NO=调整单据NO
    '          模块号.药品移库(药品移库)- 药品=药品id，库房=移出库房id，移入库房=移入库房id，开始时间=填制开始时间，结束时间=填制结束时间，NO=移库单NO
    '          模块号.药品领用(药品领用)- 药品=药品id，库房=库房id，开始时间=填制开始时间，结束时间=填制结束时间，NO=领用单NO
    '          模块号.其他出库(其他出库)- 药品=药品id，库房=库房id，开始时间=填制开始时间，结束时间=填制结束时间，NO=出库单NO
    Dim str开始时间 As String
    Dim str结束时间 As String
    Dim strNo As String
    Dim strReportName As String
    
    strReportName = Split(mnuReportItem(index).Tag, ",")(1)
    
    If strReportName = "ZL1_INSIDE_模块号.药品领用_1" Then
        ReportOpen gcnOracle, glngSys, "ZL1_INSIDE_模块号.药品领用_1", Me, "期间=" & Format(Sys.Currentdate, "YYYY"), "库房=" & cboStock.Text & "|" & cboStock.ItemData(cboStock.ListIndex), "单位=住院单位" & "|4"
    Else
        If vsfList.Row >= 1 And LTrim(vsfList.TextMatrix(vsfList.Row, 0)) <> "" Then
            strNo = vsfList.TextMatrix(vsfList.Row, 0)
        End If
        
        str开始时间 = IIf(Format(SQLCondition.date填制时间开始, "yyyy-mm-dd") = "1899-12-30", "", Format(SQLCondition.date填制时间开始, "yyyy-mm-dd"))
        str结束时间 = IIf(Format(SQLCondition.date填制时间结束, "yyyy-mm-dd") = "1899-12-30", "", Format(SQLCondition.date填制时间结束, "yyyy-mm-dd"))
            
        Call ReportOpen(gcnOracle, Split(mnuReportItem(index).Tag, ",")(0), Split(mnuReportItem(index).Tag, ",")(1), Me, _
            "药品=" & IIf(SQLCondition.lng药品 = 0, "", SQLCondition.lng药品), _
            "库房=" & IIf(Val(cboStock.ItemData(cboStock.ListIndex)) = 0, "", Val(cboStock.ItemData(cboStock.ListIndex))), _
            "移入库房=" & IIf(SQLCondition.lng库房 = 0, "", SQLCondition.lng库房), _
            "供应商=" & IIf(SQLCondition.lng生产商 = 0, "", SQLCondition.lng生产商), _
            "产地=" & SQLCondition.str产地, _
            "开始时间=" & str开始时间, _
            "结束时间=" & str结束时间, _
            "NO=" & strNo)
    End If
End Sub
Private Sub mnuViewColDefine_Click()
    Dim strColumn_All As String, strColumn_Select As String, strColumn_UnSelect As String
    Dim str可选列 As String
    Dim str屏蔽列 As String '默认屏蔽列
    Dim strAllCol As String
    Dim arr总列, arr设置列
    Dim strChange As String
    Dim strOldColName As String, strNewColName As String
    Dim intCol As Integer
    
    On Error Resume Next
    
    Select Case mlngMode
    Case 模块号.外购入库           '药品外购入库管理
        strColumn_All = "药名,0|药品来源,1|基本药物,1|药价级别,1|规格,1|生产商,0|原产地,1|批号,0|生产日期,1|效期,0|单位,1|数量,0|指导批发价,1|采购价,1|扣率,1|" & _
                        "成本价,0|成本金额,0|加成率,1|售价,0|售价金额,0|差价,0|零售价,1|零售单位,1|零售金额,1|零售差价,1|批准文号,1|外观,1|" & _
                        "产品合格证,1|随货单号,1|随货日期,1|验收结论,1|发票号,0|发票代码,0|发票日期,0|发票金额,0"
        str可选列 = "药名|药品来源|基本药物|药价级别|规格|生产商|原产地|批号|生产日期|效期|单位|数量|指导批发价|采购价|扣率|" & _
                        "成本价|成本金额|加成率|售价|售价金额|差价|零售价|零售单位|零售金额|零售差价|批准文号|外观|产品合格证|随货单号|随货日期|验收结论|发票号|发票代码|发票日期|发票金额"
        str屏蔽列 = "零售价|零售单位|零售金额|零售差价"
    Case 模块号.自制入库           '药品自制入库管理
    Case 模块号.其他入库           '药品其他入库管理
        strColumn_All = "药名,0|药品来源,1|基本药物,1|规格,1|生产商,0|原产地,1|批号,0|生产日期,1|效期,0|单位,1|数量,0|冲销数量,0|成本价,1|成本金额,1|" & _
                        "售价,0|售价金额,0|差价,0|零售价,1|零售单位,1|零售金额,1|零售差价,1|批准文号,1|外观,1"
        str可选列 = "药名|药品来源|基本药物|规格|生产商|原产地|批号|生产日期|效期|单位|数量|冲销数量|成本价|成本金额|" & _
                        "售价|售价金额|差价|零售价|零售单位|零售金额|零售差价|批准文号|外观"
        str屏蔽列 = "零售价|零售单位|零售金额|零售差价"
    Case 模块号.差价调整           '库存差价调整管理
    Case 模块号.药品移库           '药品移库管理
    Case 模块号.药品领用           '药品领用管理
    Case 模块号.其他出库           '药品其他出库管理
    End Select
    
    '取已选择列的信息'Me.Caption
    strColumn_Select = zlDataBase.GetPara("选择列", glngSys, mlngMode, "")
    strColumn_UnSelect = zlDataBase.GetPara("屏蔽列", glngSys, mlngMode, "")

    '外购入库、其他入库默认无"零售价|零售单位|零售金额|零售差价"这几列
    If mlngMode = 模块号.外购入库 Or mlngMode = 模块号.其他入库 Then
        If strColumn_Select <> "" Then
            '兼容老版本处理，列名称变化，格式：老列名,新列名|老列名,新列名...
            strChange = "产地,生产商|结算价,成本价|结算金额,成本金额"
        
            For intCol = 0 To UBound(Split(strChange, "|"))
                strOldColName = Split(Split(strChange, "|")(intCol), ",")(0)
                strNewColName = Split(Split(strChange, "|")(intCol), ",")(1)
                        
                If InStr(1, "|" & strColumn_Select & "|", "|" & strOldColName & "|") <> 0 Then
                    strColumn_Select = Replace("|" & strColumn_Select & "|", "|" & strOldColName & "|", "|" & strNewColName & "|")
                    strColumn_Select = Left(strColumn_Select, Len(strColumn_Select) - 1)
                    strColumn_Select = Mid(strColumn_Select, 2)
                End If
                
                If InStr(1, "|" & strColumn_UnSelect & "|", "|" & strOldColName & "|") <> 0 Then
                    strColumn_UnSelect = Replace("|" & strColumn_UnSelect & "|", "|" & strOldColName & "|", "|" & strNewColName & "|")
                    strColumn_UnSelect = Left(strColumn_UnSelect, Len(strColumn_UnSelect) - 1)
                    strColumn_UnSelect = Mid(strColumn_UnSelect, 2)
                End If
            Next
            
            If strColumn_UnSelect <> "" Then
                strAllCol = strColumn_Select & "|" & strColumn_UnSelect
            Else
                strAllCol = strColumn_Select
            End If
            arr总列 = Split(str可选列, "|")
            arr设置列 = Split(strAllCol, "|")
            
            If UBound(arr总列) <> UBound(arr设置列) Or InStr(1, "|" & strColumn_Select & "|", "|生产商|") = 0 Or InStr(1, "|" & strColumn_UnSelect & "|", "|生产商|") <> 0 Or (mlngMode = 模块号.其他入库 And (InStr(1, "|" & strColumn_Select & "|", "|采购价|") <> 0 Or InStr(1, "|" & strColumn_UnSelect & "|", "|采购价|") <> 0)) Then
                Select Case mlngMode
                Case 模块号.外购入库
                    strColumn_Select = "药名|药品来源|基本药物|药价级别|规格|生产商|原产地|批号|生产日期|效期|单位|数量|指导批发价|采购价|扣率|成本价|成本金额|加成率|售价|售价金额|差价|批准文号|外观|产品合格证|随货单号|随货日期|发票号|发票代码|发票日期|发票金额"
                    strColumn_UnSelect = "零售价|零售单位|零售金额|零售差价"
                    zlDataBase.SetPara "选择列", strColumn_Select, glngSys, mlngMode
                    zlDataBase.SetPara "屏蔽列", strColumn_UnSelect, glngSys, mlngMode
                Case 模块号.其他入库
                    strColumn_Select = "药名|药品来源|基本药物|规格|生产商|原产地|批号|生产日期|效期|单位|数量|冲销数量|成本价|成本金额|售价|售价金额|差价|批准文号|外观"
                    strColumn_UnSelect = "零售价|零售单位|零售金额|零售差价"
                    zlDataBase.SetPara "选择列", strColumn_Select, glngSys, mlngMode
                    zlDataBase.SetPara "屏蔽列", strColumn_UnSelect, glngSys, mlngMode
                End Select
            End If
        Else
            Select Case mlngMode
            Case 模块号.外购入库
                strColumn_Select = "药名|药品来源|基本药物|药价级别|规格|生产商|原产地|批号|生产日期|效期|单位|数量|指导批发价|采购价|扣率|成本价|成本金额|加成率|售价|售价金额|差价|批准文号|外观|产品合格证|随货单号|随货日期|发票号|发票代码|发票日期|发票金额"
                strColumn_UnSelect = "零售价|零售单位|零售金额|零售差价"
                zlDataBase.SetPara "选择列", strColumn_Select, glngSys, mlngMode
                zlDataBase.SetPara "屏蔽列", strColumn_UnSelect, glngSys, mlngMode
            Case 模块号.其他入库
                strColumn_Select = "药名|药品来源|基本药物|规格|生产商|原产地|批号|生产日期|效期|单位|数量|冲销数量|成本价|成本金额|售价|售价金额|差价|批准文号|外观"
                strColumn_UnSelect = "零售价|零售单位|零售金额|零售差价"
                zlDataBase.SetPara "选择列", strColumn_Select, glngSys, mlngMode
                zlDataBase.SetPara "屏蔽列", strColumn_UnSelect, glngSys, mlngMode
            End Select
        End If
    End If
    
    If Not frm列设置.ShowME(Me, strColumn_All, strColumn_Select) Then Exit Sub
    
    zlDataBase.SetPara "选择列", Split(strColumn_Select, "||")(0), glngSys, mlngMode
    zlDataBase.SetPara "屏蔽列", Split(strColumn_Select, "||")(1), glngSys, mlngMode
End Sub

Private Sub mnuViewRefresh_Click()
    '刷新
    If cboStock.ListIndex = -1 Then
        MsgBox "请选择库房！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    GetList mstrFind
End Sub

Private Sub mnuViewSearch_Click()
    '查找
    Dim strFind As String
    
    If cboStock.ListIndex = -1 Then
        MsgBox "请选择库房！", vbInformation, gstrSysName
        Exit Sub
    End If
    Select Case mlngMode
        Case 模块号.差价调整, 模块号.药品移库, 模块号.药品领用, 模块号.其他出库
            FrmTransferSearch.In_入出类型 = IIf(TabShow.Tab = 0, -1, 1)
            strFind = FrmTransferSearch.GetSearch(Me, mlngMode, mlng库房ID, strStart, strEnd, strVerifyStart, strVerifyEnd, _
                SQLCondition.strNO开始, _
                SQLCondition.strNO结束, _
                SQLCondition.date填制时间开始, _
                SQLCondition.date填制时间结束, _
                SQLCondition.date审核时间开始, _
                SQLCondition.date审核时间结束, _
                SQLCondition.lng药品, _
                SQLCondition.lng库房, _
                SQLCondition.str填制人, _
                SQLCondition.str审核人, _
                SQLCondition.lng药品分类, _
                SQLCondition.str剂型, _
                SQLCondition.int填制审核一并查询)
        Case 模块号.外购入库
            strFind = FrmPurchaseSearch.GetSearch(Me, strStart, strEnd, strVerifyStart, strVerifyEnd, _
                SQLCondition.strNO开始, _
                SQLCondition.strNO结束, _
                SQLCondition.date填制时间开始, _
                SQLCondition.date填制时间结束, _
                SQLCondition.date审核时间开始, _
                SQLCondition.date审核时间结束, _
                SQLCondition.lng药品, _
                SQLCondition.str填制人, _
                SQLCondition.str审核人, _
                SQLCondition.lng生产商, _
                SQLCondition.str产地, _
                SQLCondition.str发票号开始, _
                SQLCondition.str发票号结束, _
                SQLCondition.lng药品分类, _
                SQLCondition.str剂型, _
                SQLCondition.date发票审核日期开始, _
                SQLCondition.date发票审核日期结束, _
                SQLCondition.int无标记, _
                SQLCondition.int有标记, _
                SQLCondition.int无发票, _
                SQLCondition.int有发票, _
                SQLCondition.int填制审核一并查询)
                
'                Call FrmPurchaseSearch.GetInfo(SQLCondition.int无标记, SQLCondition.int有标记, SQLCondition.int无发票, SQLCondition.int有发票)
        Case 模块号.自制入库
            strFind = FrmSelfMakeSearch.GetSearch(Me, strStart, strEnd, strVerifyStart, strVerifyEnd, _
                SQLCondition.strNO开始, _
                SQLCondition.strNO结束, _
                SQLCondition.date填制时间开始, _
                SQLCondition.date填制时间结束, _
                SQLCondition.date审核时间开始, _
                SQLCondition.date审核时间结束, _
                SQLCondition.lng药品, _
                SQLCondition.str填制人, _
                SQLCondition.str审核人)
        Case 模块号.其他入库
            strFind = FrmOtherInputSearch.GetSearch(Me, strStart, strEnd, strVerifyStart, strVerifyEnd, _
                SQLCondition.strNO开始, _
                SQLCondition.strNO结束, _
                SQLCondition.date填制时间开始, _
                SQLCondition.date填制时间结束, _
                SQLCondition.date审核时间开始, _
                SQLCondition.date审核时间结束, _
                SQLCondition.lng药品, _
                SQLCondition.str填制人, _
                SQLCondition.str审核人, _
                SQLCondition.str产地, _
                SQLCondition.lng入出类别)
    End Select
    
    If strFind <> "" Or SQLCondition.int填制审核一并查询 = 1 Then
        mstrFind = strFind
        vsfList.rows = 1
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
    Call SetMenu
        
    cbrTool.Bands(1).MinHeight = tlbTool.Height
    
    Form_Resize
End Sub





Private Sub vsfDetail_GotFocus()
    Call SetGridFocus(vsfDetail, True)
End Sub

Private Sub vsfDetail_LostFocus()
    Call SetGridFocus(vsfDetail, False)
End Sub

Private Sub vsfDetail_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    If mnuFileCodePrint.Visible = False Then Exit Sub
    
    PopupMenu mnuFileCodePrint, 2
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
    Dim intBill As Integer                      '单据类型  如：1、外购入库；2、
    Dim strUnit As String                       '单位名称:如门诊单位，住院单位等
    Dim str包装系数 As String
    Dim strOrder As String
    Dim strCompare As String
    Dim strSql效期 As String
    Dim n As Long
    Dim strSql药名 As String
    Dim intCol As Integer
    Dim intRow As Integer
    Dim strSqlOrder As String
    Dim strTemp As String
    
    On Error GoTo errHandle
    
    If mblnBandEvent = True Then Exit Sub
    If mlastRow = vsfList.Row Then
        SetEnable
        Exit Sub
    End If
    mlastRow = vsfList.Row
    
    With vsfList
        .Redraw = flexRDNone
        .ForeColorSel = .Cell(flexcpForeColor, mlastRow, 1)
        .Redraw = flexRDDirect
    End With
    
    strOrder = zlDataBase.GetPara("排序", glngSys, mlngMode)
    strCompare = Mid(strOrder, 1, 1)
    
    strSqlOrder = "序号"
    
    If strCompare = "0" Then
        '按序号排序
        strSqlOrder = "序号"
    ElseIf strCompare = "1" Then
        '按编码排序
        strSqlOrder = "药品信息"
    ElseIf strCompare = "2" Then
        '按名称排序
        strSqlOrder = "Substr(药品信息, Instr(药品信息, ']') + 1)"
    ElseIf strCompare = "3" Then
        ''按库房货位排序
        strSqlOrder = "库房货位"
    End If
    
    strSqlOrder = strSqlOrder & IIf(Right(strOrder, 1) = "0", " ASC", " DESC") & ",药品信息,序号"

    If vsfList.Row >= 1 And LTrim(vsfList.TextMatrix(vsfList.Row, 0)) <> "" Then
        vsfList.Col = 0
        vsfList.ColSel = vsfList.Cols - 1
        
        vsfDetail.Redraw = flexRDNone
        
        Select Case mintUnit
            Case mconint售价单位
                strUnit = "F.计算单位"
                strUnitQuantity = "LTRIM(to_char(A.实际数量," & mstrNumberFormat & ")) AS 数量," _
                    & "F.计算单位 AS 单位,"
                str包装系数 = "1"
            Case mconint门诊单位
                strUnit = "B.门诊单位"
                strUnitQuantity = "LTRIM(to_char(A.实际数量 / B.门诊包装," & mstrNumberFormat & ")) AS 数量," _
                    & "B.门诊单位 AS 单位,"
                str包装系数 = "B.门诊包装"
            Case mconint住院单位
                strUnit = "B.住院单位"
                strUnitQuantity = "LTRIM(to_char(A.实际数量 / B.住院包装," & mstrNumberFormat & ")) AS 数量," _
                    & "B.住院单位 AS 单位,"
                str包装系数 = "B.住院包装"
            Case mconint药库单位
                strUnit = "B.药库单位"
                strUnitQuantity = "LTRIM(to_char(A.实际数量 / B.药库包装," & mstrNumberFormat & ")) AS 数量," _
                    & "B.药库单位 AS 单位,"
                str包装系数 = "B.药库包装"
        End Select
        
        strSql效期 = IIf(gtype_UserSysParms.P149_效期显示方式 = 1, "TO_CHAR(A.效期-1,'YYYY-MM-DD') AS 有效期至", "TO_CHAR(A.效期,'YYYY-MM-DD') AS 失效期")
        
        If gint药品名称显示 = 0 Then
            strSql药名 = ",('['||F.编码||']'||F.名称) AS 药品信息"
        ElseIf gint药品名称显示 = 1 Then
            strSql药名 = ",('['||F.编码||']'||NVL(E.名称,F.名称)) AS 药品信息"
        Else
            strSql药名 = ",('['||F.编码||']'||F.名称) AS 药品信息,E.名称 As 商品名"
        End If
        
        Select Case mlngMode
            Case 模块号.外购入库
                intBill = 1
                strTemp = ""
                
                If SQLCondition.int有标记 = 1 And SQLCondition.int无标记 = 0 Then
                    strTemp = strTemp & " and c.付款标志=1"
                End If
                If SQLCondition.int无标记 = 1 And SQLCondition.int有标记 = 0 Then
                    strTemp = strTemp & " and (c.付款标志=0 or c.付款标志 is null)"
                End If
                If SQLCondition.int无发票 = 1 And SQLCondition.int有发票 = 0 Then
                    strTemp = strTemp & " and c.发票号 is null"
                End If
                If SQLCondition.int有发票 = 1 And SQLCondition.int无发票 = 0 Then
                    strTemp = strTemp & " and c.发票号 is not null"
                End If
'                gstrSQL = "SELECT * FROM (SELECT DISTINCT A.序号,decode(c.付款标志,Null,'未标记',0,'未标记','已标记') 付款标志 " & strSql药名 & ",B.药品来源,B.基本药物,F.规格, A.产地, A.批号, " & strSql效期 & " ," & _
'                    strUnitQuantity & _
'                    " LTRIM(TO_CHAR(A.成本价*" & str包装系数 & "," & mstrCostFormat & ")) AS 成本价, LTRIM(TO_CHAR(A.成本金额," & mstrMoneyFormat & ")) AS 成本金额," & _
'                    " DECODE(A.扣率, NULL, 0, A.扣率) AS 扣率, " & _
'                    " LTRIM(TO_CHAR(Decode(To_Number(Nvl(A.用法, 0)), 0, A.零售价, (A.零售金额 - To_Number(Nvl(A.用法, 0))) / A.实际数量)*" & str包装系数 & "," & mstrPriceFormat & ")) AS 售价 , " & _
'                    " LTRIM(TO_CHAR(A.零售金额- To_Number(Nvl(A.用法, 0))," & mstrMoneyFormat & "))  AS 售价金额, LTRIM(TO_CHAR(A.差价- To_Number(Nvl(A.用法, 0))," & mstrMoneyFormat & ")) AS 差价," & _
'                    " A.批准文号, C.发票号 ,TO_CHAR(C.发票日期,'YYYY-MM-DD') AS 发票日期,NVL(C.付款序号,'0') AS 付款序号, " & _
'                    " LTRIM(TO_CHAR(Decode(C.发票号,Null,0,C.发票金额)," & mstrMoneyFormat & ")) AS 发票金额,B.招标药品,B.差价让利比,C.随货单号, " & _
'                    " Decode(Nvl(F.是否变价, 0) * Nvl(A.批次, 0), 0, '',LTRIM(TO_CHAR(A.零售价," & mstrPriceFormat & "))) As 零售价," & _
'                    " Decode(Nvl(F.是否变价, 0) * Nvl(A.批次, 0), 0, '',LTRIM(TO_CHAR(A.零售金额," & mstrMoneyFormat & ")))  AS 零售金额, " & _
'                    " Decode(Nvl(F.是否变价, 0) * Nvl(A.批次, 0), 0, '',LTRIM(TO_CHAR(A.差价," & mstrMoneyFormat & "))) AS 零售差价 " & _
'                    " FROM 药品收发记录 A, 药品规格 B,收费项目目录 F,收费项目别名 E ,应付记录 C " & _
'                    " WHERE  A.药品ID = B.药品ID AND B.药品ID=F.ID " & _
'                    " AND A.ID = C.收发ID (+) AND C.系统标识(+)=1 AND C.记录性质(+)<>-1 " & _
'                    " AND B.药品ID=E.收费细目ID(+) AND E.性质(+)=3 " & _
'                    " AND A.单据 = [1] AND A.记录状态 = [3] " & _
'                    " AND A.NO =[2] " & strTemp & _
'                    " ) ORDER BY " & strSqlOrder
                gstrSQL = "SELECT * FROM (SELECT DISTINCT A.序号,decode(c.付款标志,Null,'未标记',0,'未标记','已标记') 付款标志 " & strSql药名 & ",B.药品来源,B.基本药物,F.规格, A.产地 as 生产商,A.原产地, A.批号, " & strSql效期 & " ," & _
                    strUnitQuantity & _
                    " LTRIM(TO_CHAR(A.成本价*" & str包装系数 & "," & mstrCostFormat & ")) AS 成本价, LTRIM(TO_CHAR(A.成本金额," & mstrMoneyFormat & ")) AS 成本金额," & _
                    " DECODE(A.扣率, NULL, 0, A.扣率) AS 扣率, " & _
                    " LTRIM(TO_CHAR(A.零售价*" & str包装系数 & "," & mstrPriceFormat & ")) AS 售价 , " & _
                    " LTRIM(TO_CHAR(A.零售金额," & mstrMoneyFormat & "))  AS 售价金额, LTRIM(TO_CHAR(A.差价," & mstrMoneyFormat & ")) AS 差价," & _
                    " A.批准文号, C.发票号 ,c.发票代码,TO_CHAR(C.发票日期,'YYYY-MM-DD') AS 发票日期,NVL(C.付款序号,'0') AS 付款序号, " & _
                    " LTRIM(TO_CHAR(Decode(C.发票号,Null,decode(c.发票代码,Null,0,c.发票金额),C.发票金额)," & mstrMoneyFormat & ")) AS 发票金额,B.招标药品,B.差价让利比,C.随货单号,TO_CHAR(C.随货日期,'YYYY-MM-DD') AS 随货日期, " & _
                    " Decode(Nvl(F.是否变价, 0) * Nvl(A.批次, 0), 0, '',LTRIM(TO_CHAR(A.零售价," & mstrPriceFormat & "))) As 零售价," & _
                    " Decode(Nvl(F.是否变价, 0) * Nvl(A.批次, 0), 0, '',LTRIM(TO_CHAR(A.零售金额," & mstrMoneyFormat & ")))  AS 零售金额, " & _
                    " Decode(Nvl(F.是否变价, 0) * Nvl(A.批次, 0), 0, '',LTRIM(TO_CHAR(A.差价," & mstrMoneyFormat & "))) AS 零售差价,F.ID 药品ID " & _
                    " FROM 药品收发记录 A, 药品规格 B,收费项目目录 F,收费项目别名 E ,应付记录 C " & _
                    " WHERE  A.药品ID = B.药品ID AND B.药品ID=F.ID " & _
                    " AND A.ID = C.收发ID (+) AND C.系统标识(+)=1 AND C.记录性质(+)=0 " & _
                    " AND B.药品ID=E.收费细目ID(+) AND E.性质(+)=3 " & _
                    " AND A.单据 = [1] AND A.记录状态 = [3] " & _
                    " AND A.NO =[2] " & strTemp & _
                    " ) ORDER BY " & strSqlOrder
            Case 模块号.自制入库
                intBill = 2
                gstrSQL = " SELECT * FROM (SELECT DISTINCT A.序号" & strSql药名 & ",B.药品来源,B.基本药物,F.规格,A.批号, " & strSql效期 & "," & _
                    strUnitQuantity & _
                    " LTRIM(TO_CHAR(A.成本价*" & str包装系数 & "," & mstrCostFormat & ")) AS 采购价," & _
                    " LTRIM(TO_CHAR (A.成本金额, " & mstrMoneyFormat & ")) AS 采购金额," & _
                    " LTRIM(TO_CHAR (A.零售价*" & str包装系数 & ", " & mstrPriceFormat & ")) AS 售价," & _
                    " LTRIM(TO_CHAR (A.零售金额, " & mstrMoneyFormat & ")) AS 售价金额," & _
                    " LTRIM(TO_CHAR (A.差价, " & mstrMoneyFormat & ")) AS 差价 " & _
                    " FROM 药品收发记录 A, 药品规格 B,收费项目目录 F, 收费项目别名 E " & _
                    " WHERE A.药品ID = B.药品ID AND B.药品ID=F.ID " & _
                    " AND B.药品ID = E.收费细目ID (+) AND E.性质(+)=3 " & _
                    " AND 记录状态 = [3] " & _
                    " AND A.单据 = [1] AND 入出系数=1 " & _
                    " AND A.NO = [2] " & _
                    " ) ORDER BY " & strSqlOrder
            Case 模块号.其他入库
                intBill = 4
'                gstrSQL = "SELECT * FROM (SELECT DISTINCT A.序号" & strSql药名 & ",B.药品来源,B.基本药物," & _
'                    "F.规格, A.产地, A.批号, " & strSql效期 & "," & strUnitQuantity & _
'                    " LTRIM(TO_CHAR(A.成本价*" & str包装系数 & "," & mstrCostFormat & ")) AS 成本价, LTRIM(TO_CHAR(A.成本金额," & mstrMoneyFormat & ")) AS 成本金额," & _
'                    " LTRIM(TO_CHAR(Decode(To_Number(Nvl(A.用法, 0)), 0, A.零售价, (A.零售金额 - To_Number(Nvl(A.用法, 0))) / A.实际数量)*" & str包装系数 & "," & mstrPriceFormat & ")) AS 售价," & _
'                    " LTRIM(TO_CHAR(A.零售金额- To_Number(Nvl(A.用法, 0))," & mstrMoneyFormat & "))  AS 售价金额, LTRIM(TO_CHAR(A.差价- To_Number(Nvl(A.用法, 0))," & mstrMoneyFormat & ")) AS 差价, " & _
'                    " Decode(Nvl(F.是否变价, 0) * Nvl(A.批次, 0), 0, '',LTRIM(TO_CHAR(A.零售价," & mstrPriceFormat & "))) As 零售价," & _
'                    " Decode(Nvl(F.是否变价, 0) * Nvl(A.批次, 0), 0, '',LTRIM(TO_CHAR(A.零售金额," & mstrMoneyFormat & ")))  AS 零售金额, " & _
'                    " Decode(Nvl(F.是否变价, 0) * Nvl(A.批次, 0), 0, '',LTRIM(TO_CHAR(A.差价," & mstrMoneyFormat & "))) AS 零售差价 " & _
'                    " FROM 药品收发记录 A, 药品规格 B,收费项目目录 F,收费项目别名 E  " & _
'                    " WHERE  A.药品ID = B.药品ID AND B.药品ID=F.ID" & _
'                    " AND B.药品ID=E.收费细目ID(+) AND E.性质(+)=3 " & _
'                    " AND 记录状态 = [3] " & _
'                    " AND A.单据 = [1] " & _
'                    " AND A.NO =[2] " & _
'                    " ) ORDER BY " & strSqlOrder
                gstrSQL = "SELECT * FROM (SELECT DISTINCT A.序号" & strSql药名 & ",B.药品来源,B.基本药物," & _
                    "F.规格, A.产地 as 生产商,A.原产地, A.批号, " & strSql效期 & "," & strUnitQuantity & _
                    " LTRIM(TO_CHAR(A.成本价*" & str包装系数 & "," & mstrCostFormat & ")) AS 成本价, LTRIM(TO_CHAR(A.成本金额," & mstrMoneyFormat & ")) AS 成本金额," & _
                    " LTRIM(TO_CHAR(A.零售价*" & str包装系数 & "," & mstrPriceFormat & ")) AS 售价," & _
                    " LTRIM(TO_CHAR(A.零售金额," & mstrMoneyFormat & "))  AS 售价金额, LTRIM(TO_CHAR(A.差价," & mstrMoneyFormat & ")) AS 差价, " & _
                    " Decode(Nvl(F.是否变价, 0) * Nvl(A.批次, 0), 0, '',LTRIM(TO_CHAR(A.零售价," & mstrPriceFormat & "))) As 零售价," & _
                    " Decode(Nvl(F.是否变价, 0) * Nvl(A.批次, 0), 0, '',LTRIM(TO_CHAR(A.零售金额," & mstrMoneyFormat & ")))  AS 零售金额, " & _
                    " Decode(Nvl(F.是否变价, 0) * Nvl(A.批次, 0), 0, '',LTRIM(TO_CHAR(A.差价," & mstrMoneyFormat & "))) AS 零售差价,F.ID 药品ID " & _
                    " FROM 药品收发记录 A, 药品规格 B,收费项目目录 F,收费项目别名 E  " & _
                    " WHERE  A.药品ID = B.药品ID AND B.药品ID=F.ID" & _
                    " AND B.药品ID=E.收费细目ID(+) AND E.性质(+)=3 " & _
                    " AND 记录状态 = [3] " & _
                    " AND A.单据 = [1] " & _
                    " AND A.NO =[2] " & _
                    " ) ORDER BY " & strSqlOrder
            Case 模块号.差价调整
                intBill = 5
                gstrSQL = "SELECT * FROM (SELECT DISTINCT A.序号" & strSql药名 & ",B.药品来源,B.基本药物," & _
                    "F.规格, A.产地 as 生产商,A.原产地, A.批号, " & strSql效期 & "," & strUnit & _
                    " AS 单位,LTRIM(TO_CHAR(A.零售价," & mstrMoneyFormat & ")) AS 库存金额,LTRIM(TO_CHAR(A.成本价," & mstrMoneyFormat & ")) AS 库存差价," & _
                    " LTRIM(TO_CHAR(A.差价," & mstrMoneyFormat & "))  AS 调整额, " & _
                    " LTRIM(TO_CHAR(A.单量*" & str包装系数 & "," & mstrCostFormat & ")) AS 新成本价 " & _
                    " FROM 药品收发记录 A, 药品规格 B,收费项目目录 F,收费项目别名 E " & _
                    " WHERE A.药品ID = B.药品ID AND B.药品ID=F.ID " & _
                    " AND B.药品ID=E.收费细目ID(+) AND E.性质(+)=3 " & _
                    " AND 记录状态 = [3] " & _
                    " AND A.单据 =[1] " & _
                    " AND A.NO =[2] " & _
                    " ) ORDER BY " & strSqlOrder
                    
            Case 模块号.药品移库       '药品移库管理
                intBill = 6
                gstrSQL = " SELECT * FROM (SELECT DISTINCT A.序号" & strSql药名 & ",B.药品来源,B.基本药物,F.规格,A.产地 as 生产商,A.原产地, " & _
                    " A.批号, " & strSql效期 & ",LTRIM(TO_CHAR(A.填写数量 /" & str包装系数 & "," & mstrNumberFormat & ")) AS 填写数量," & _
                    " LTRIM(TO_CHAR(A.实际数量 /" & str包装系数 & "," & mstrNumberFormat & ")) AS 实际数量," & strUnit & " AS 单位," & _
                    " LTRIM(TO_CHAR (A.成本价*" & str包装系数 & ", " & mstrCostFormat & ")) AS 成本价," & _
                    " LTRIM(TO_CHAR (A.成本金额, " & mstrMoneyFormat & ")) AS 成本金额," & _
                    " LTRIM(TO_CHAR (A.零售价*" & str包装系数 & ", " & mstrPriceFormat & ")) AS 售价," & _
                    " LTRIM(TO_CHAR (A.零售金额, " & mstrMoneyFormat & ")) AS 售价金额," & _
                    " LTRIM(TO_CHAR (A.差价, " & mstrMoneyFormat & ")) AS 差价 ,C.库房货位,F.ID 药品ID " & _
                    " FROM 药品收发记录 A, 药品规格 B,收费项目目录 F,收费项目别名 E,药品储备限额 C " & _
                    " WHERE A.药品ID = B.药品ID AND B.药品ID=F.ID" & _
                    " AND B.药品ID = E.收费细目ID (+) AND E.性质(+)=3 " & _
                    " AND 记录状态 = [3] " & _
                    " AND A.单据 = [1] AND 入出系数=-1 " & _
                    " AND A.NO = [2] AND A.药品ID=C.药品ID(+) AND A.库房ID=C.库房ID(+)) " & _
                    " ORDER BY " & strSqlOrder
            Case 模块号.药品领用
                intBill = 7
                gstrSQL = " SELECT * FROM (SELECT DISTINCT A.序号" & strSql药名 & ",B.药品来源,B.基本药物,F.规格,A.产地 as 生产商,A.原产地, " & _
                    " A.批号, " & strSql效期 & ",LTRIM(TO_CHAR(A.填写数量 /" & str包装系数 & "," & mstrNumberFormat & ")) AS 填写数量," & _
                    " LTRIM(TO_CHAR(A.实际数量 /" & str包装系数 & "," & mstrNumberFormat & ")) AS 实际数量," & strUnit & " AS 单位," & _
                    " LTRIM(TO_CHAR (A.成本价*" & str包装系数 & ", " & mstrCostFormat & ")) AS 成本价," & _
                    " LTRIM(TO_CHAR (A.成本金额, " & mstrMoneyFormat & ")) AS 成本金额," & _
                    " LTRIM(TO_CHAR (A.零售价*" & str包装系数 & ", " & mstrPriceFormat & ")) AS 售价," & _
                    " LTRIM(TO_CHAR (A.零售金额, " & mstrMoneyFormat & ")) AS 售价金额," & _
                    " LTRIM(TO_CHAR (A.差价, " & mstrMoneyFormat & ")) AS 差价 ,C.库房货位 ,NVL(E.名称,F.名称) as 名称,F.ID 药品ID " & _
                    " FROM 药品收发记录 A, 药品规格 B, 收费项目目录 F,收费项目别名 E ,药品储备限额 C " & _
                    " WHERE A.药品ID = B.药品ID AND B.药品ID=F.ID " & _
                    " AND B.药品ID = E.收费细目ID (+) AND E.性质(+)=3 " & _
                    " AND 记录状态 = [3] " & _
                    " AND A.单据 = [1] " & _
                    " AND A.NO = [2] AND A.药品ID=C.药品ID(+) AND A.库房ID=C.库房ID(+))" & _
                    " ORDER BY " & strSqlOrder
            Case 模块号.其他出库
                intBill = 11
                gstrSQL = "SELECT * FROM (SELECT DISTINCT A.序号" & strSql药名 & ",B.药品来源,B.基本药物," & _
                        " F.规格, A.产地 as 生产商,A.原产地, A.批号, " & strSql效期 & "," & strUnitQuantity & _
                        " LTRIM(TO_CHAR(A.成本价*" & str包装系数 & "," & mstrCostFormat & ")) AS 成本价, LTRIM(TO_CHAR(A.成本金额," & mstrMoneyFormat & ")) AS 成本金额," & _
                        " LTRIM(TO_CHAR(A.零售价*" & str包装系数 & "," & mstrPriceFormat & ")) AS 售价 , LTRIM(TO_CHAR(A.零售金额," & mstrMoneyFormat & "))  AS 售价金额, LTRIM(TO_CHAR(A.差价," & mstrMoneyFormat & ")) AS 差价, " & _
                        " C.库房货位 ,NVL(E.名称,F.名称) as 名称 ,F.ID 药品ID, "
                    
                If vsfList.TextMatrix(vsfList.Row, 1) = "药品外调" Then
                    gstrSQL = gstrSQL & " LTRIM(TO_CHAR(A.单量*" & str包装系数 & "," & mstrPriceFormat & ")) AS 外调价,LTRIM(TO_CHAR(A.单量*A.实际数量," & mstrMoneyFormat & ")) AS 外调金额,'' As 增值税率,'' As 税金 "
                ElseIf vsfList.TextMatrix(vsfList.Row, 1) = "药品外销" Then
                    gstrSQL = gstrSQL & " LTRIM(TO_CHAR(A.单量*" & str包装系数 & "," & mstrPriceFormat & ")) AS 外销价,LTRIM(TO_CHAR(A.单量*A.实际数量," & mstrMoneyFormat & ")) AS 外销金额,LTRIM(TO_CHAR(Nvl(A.频次,0)/100," & mstrMoneyFormat & ")) As 增值税率,LTRIM(TO_CHAR(A.单量*A.实际数量*(Nvl(A.频次,0)/100/(1+Nvl(A.频次,0)/100))," & mstrMoneyFormat & ")) As 税金 "
                Else
                    gstrSQL = gstrSQL & " '' As 外调价,'' As 外调金额,'' As 增值税率,'' As 税金 "
                End If
                
                gstrSQL = gstrSQL & " FROM 药品收发记录 A, 药品规格 B,收费项目目录 F,收费项目别名 E ,药品储备限额 C " & _
                    " WHERE  A.药品ID = B.药品ID AND B.药品ID=F.ID" & _
                    " AND B.药品ID=E.收费细目ID(+) AND E.性质(+)=3 " & _
                    " AND 记录状态 = [3] " & _
                    " AND A.单据 =[1] " & _
                    " AND A.NO =[2] AND A.药品ID=C.药品ID(+) AND A.库房ID=C.库房ID(+))" & _
                    " ORDER BY " & strSqlOrder
                
        End Select
        
        If mlngMode = 模块号.药品领用 Then
            Set rsDetail = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, intBill, vsfList.TextMatrix(vsfList.Row, 0), Val(vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 4)))
        Else
            Set rsDetail = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, intBill, vsfList.TextMatrix(vsfList.Row, 0), Val(vsfList.TextMatrix(vsfList.Row, vsfList.Cols - IIf(mlngMode = 模块号.差价调整 Or mlngMode = 模块号.外购入库, 3, 2))))
        End If
        
        Set vsfDetail.DataSource = rsDetail
        With vsfDetail
            If rsDetail.RecordCount > 0 Then
                .Row = 1
            End If
            .Col = 0
            .ColSel = .Cols - 1
            
            For n = 0 To .Cols - 1
                .ColKey(n) = .TextMatrix(0, n)
                .FixedAlignment(n) = flexAlignCenterCenter
            Next
            
            .colHidden(.ColIndex("药品ID")) = True '药品ID列不显示
            
            '重新更新界面序号，因为药品移库是2条数据，但是只提取一条，所以会出现1，3，5，7或者2，4，6，8，因此需要将其改成连续的
            If mlngMode = 模块号.药品移库 Then
                For intRow = 0 To .rows - 1
                    If intRow <> 0 Then
                        .TextMatrix(intRow, vsfDetail.ColIndex("序号")) = intRow
                    End If
                Next
            End If
        End With
        
        '对各种单据的数量格式化
        If rsDetail.RecordCount > 0 Then
            With vsfDetail
                Select Case mlngMode
                Case 模块号.外购入库
                    For n = 1 To .rows - 1
                        .TextMatrix(n, .ColIndex("数量")) = zlStr.FormatEx(.TextMatrix(n, .ColIndex("数量")), mintShowNumberDigit, , True)
                    Next
                Case 模块号.自制入库
                    For n = 1 To .rows - 1
                        .TextMatrix(n, .ColIndex("数量")) = zlStr.FormatEx(.TextMatrix(n, .ColIndex("数量")), mintShowNumberDigit, , True)
                    Next
                Case 模块号.其他入库
                    For n = 1 To .rows - 1
                        .TextMatrix(n, .ColIndex("数量")) = zlStr.FormatEx(.TextMatrix(n, .ColIndex("数量")), mintShowNumberDigit, , True)
                    Next
                Case 模块号.药品移库
                    For n = 1 To .rows - 1
                        .TextMatrix(n, .ColIndex("填写数量")) = zlStr.FormatEx(.TextMatrix(n, .ColIndex("填写数量")), mintShowNumberDigit, , True)
                        .TextMatrix(n, .ColIndex("实际数量")) = zlStr.FormatEx(.TextMatrix(n, .ColIndex("实际数量")), mintShowNumberDigit, , True)
                    Next
                Case 模块号.药品领用
                    For n = 1 To .rows - 1
                        .TextMatrix(n, .ColIndex("填写数量")) = zlStr.FormatEx(.TextMatrix(n, .ColIndex("填写数量")), mintShowNumberDigit, , True)
                        .TextMatrix(n, .ColIndex("实际数量")) = zlStr.FormatEx(.TextMatrix(n, .ColIndex("实际数量")), mintShowNumberDigit, , True)
                    Next
                Case 模块号.其他出库
                    For n = 1 To .rows - 1
                        .TextMatrix(n, .ColIndex("数量")) = zlStr.FormatEx(.TextMatrix(n, .ColIndex("数量")), mintShowNumberDigit, , True)
                    Next
                End Select
            End With
        End If
        
        vsfDetail.Redraw = flexRDDirect
    ElseIf LTrim(vsfList.TextMatrix(vsfList.Row, 0)) = "" Then
        With vsfDetail
            .Redraw = flexRDNone
            Select Case mlngMode
                Case 模块号.外购入库
                    .Cols = IIf(gint药品名称显示 = 2, 32, 31)
                    .rows = 2
                    .Clear
                    
                    intCol = 0
                    
                    .TextMatrix(0, intCol) = "序号": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "付款标志": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "药品信息": intCol = intCol + 1
                    
                    If gint药品名称显示 = 2 Then
                        .TextMatrix(0, intCol) = "商品名": intCol = intCol + 1
                    End If
                    
                    .TextMatrix(0, intCol) = "药品来源": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "基本药物": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "规格": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "生产商": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "原产地": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "批号": intCol = intCol + 1
                    .TextMatrix(0, intCol) = IIf(gtype_UserSysParms.P149_效期显示方式 = 1, "有效期至", "失效期"): intCol = intCol + 1
                    .TextMatrix(0, intCol) = "数量": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "单位": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "成本价": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "成本金额": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "扣率": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "售价": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "售价金额": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "差价": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "批准文号": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "发票号": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "发票代码": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "发票日期": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "付款序号": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "发票金额": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "招标药品": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "差价让利比": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "随货单号": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "随货日期": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "零售价": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "零售金额": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "零售差价": intCol = intCol + 1
                Case 模块号.自制入库
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
                    
                    .TextMatrix(0, intCol) = "批号": intCol = intCol + 1
                    .TextMatrix(0, intCol) = IIf(gtype_UserSysParms.P149_效期显示方式 = 1, "有效期至", "失效期"): intCol = intCol + 1
                    .TextMatrix(0, intCol) = "数量": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "单位": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "采购价": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "采购金额": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "售价": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "售价金额": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "差价": intCol = intCol + 1
                Case 模块号.其他入库
                    .Cols = IIf(gint药品名称显示 = 2, 20, 19)
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
                    .TextMatrix(0, intCol) = "批号": intCol = intCol + 1
                    .TextMatrix(0, intCol) = IIf(gtype_UserSysParms.P149_效期显示方式 = 1, "有效期至", "失效期"): intCol = intCol + 1
                    .TextMatrix(0, intCol) = "数量": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "单位": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "成本价": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "成本金额": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "售价": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "售价金额": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "差价": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "零售价": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "零售金额": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "零售差价": intCol = intCol + 1
                Case 模块号.差价调整
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
                    .TextMatrix(0, intCol) = "批号": intCol = intCol + 1
                    .TextMatrix(0, intCol) = IIf(gtype_UserSysParms.P149_效期显示方式 = 1, "有效期至", "失效期"): intCol = intCol + 1
                    .TextMatrix(0, intCol) = "单位": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "库存金额": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "库存差价": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "调整额": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "新成本价": intCol = intCol + 1
                
                Case 模块号.药品移库
                    .Cols = IIf(gint药品名称显示 = 2, 19, 18)
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
                    .TextMatrix(0, intCol) = "批号": intCol = intCol + 1
                    .TextMatrix(0, intCol) = IIf(gtype_UserSysParms.P149_效期显示方式 = 1, "有效期至", "失效期"): intCol = intCol + 1
                    .TextMatrix(0, intCol) = "填写数量": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "实际数量": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "单位": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "成本价": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "成本金额": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "售价": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "售价金额": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "差价": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "库房货位": intCol = intCol + 1
                Case 模块号.药品领用
                    .Cols = IIf(gint药品名称显示 = 2, 20, 19)
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
                    .TextMatrix(0, intCol) = "批号": intCol = intCol + 1
                    .TextMatrix(0, intCol) = IIf(gtype_UserSysParms.P149_效期显示方式 = 1, "有效期至", "失效期"): intCol = intCol + 1
                    .TextMatrix(0, intCol) = "填写数量": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "实际数量": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "单位": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "成本价": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "成本金额": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "售价": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "售价金额": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "差价": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "库房货位": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "名称": intCol = intCol + 1
                Case 模块号.其他出库
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
                    .TextMatrix(0, intCol) = "基本药物": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "规格": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "生产商": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "原产地": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "批号": intCol = intCol + 1
                    .TextMatrix(0, intCol) = IIf(gtype_UserSysParms.P149_效期显示方式 = 1, "有效期至", "失效期"): intCol = intCol + 1
                    .TextMatrix(0, intCol) = "数量": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "单位": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "成本价": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "成本金额": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "售价": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "售价金额": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "差价": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "库房货位": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "名称": intCol = intCol + 1
                    
                    If vsfList.TextMatrix(vsfList.Row, 1) = "药品外销" Then
                        .TextMatrix(0, intCol) = "外销价": intCol = intCol + 1
                        .TextMatrix(0, intCol) = "外销金额": intCol = intCol + 1
                    Else
                        .TextMatrix(0, intCol) = "外调价": intCol = intCol + 1
                        .TextMatrix(0, intCol) = "外调金额": intCol = intCol + 1
                    End If
                    
                    .TextMatrix(0, intCol) = "增值税率": intCol = intCol + 1
                    .TextMatrix(0, intCol) = "税金": intCol = intCol + 1
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
    SetEnable
    Call ShowColor(rsDetail)
    If mlngMode = 模块号.药品移库 Then Call CheckNumber
    
    If mblnDo Then
        RestoreFlexState vsfDetail, App.ProductName & "\" & Me.Name & mstrTitle
    End If
    
    If vsfDetail.rows > 1 Then
        vsfDetail.Row = 1
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetListFocuse()
    Dim intStatus As Integer
    Dim lngForeColor As Long
    
    With vsfList
        .ForeColorFixed = glngFixedForeColorByFocus
        .BackColorSel = glngRowByFocus
    
'        If .Row > 0 Then
'            .ForeColorSel = .Cell(flexcpForeColor, .Row)
'        End If
    End With
    
    vsfDetail.ForeColorFixed = glngFixedForeColorNotFocus
    vsfDetail.BackColorSel = glngRowByNotFocus
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
    
    mnuEditAllCodePrint.Visible = True
    mnuEditAllCodePrint.Visible = mlngMode = 1300 Or mlngMode = 1302 Or mlngMode = 1304 Or mlngMode = 1305 Or mlngMode = 1306
    mnuEditCodePrintLine.Visible = mnuEditAllCodePrint.Visible
    PopupMenu mnuEdit, 2
    mnuEditAllCodePrint.Visible = False
    mnuEditCodePrintLine.Visible = mnuEditAllCodePrint.Visible
    
End Sub

Private Sub picSeparate_s_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    '分割条设置
    
    If Button <> 1 Then Exit Sub
    With picSeparate_s
        If .Top + y < 2000 Then Exit Sub
        If .Top + y > ScaleHeight - 2000 Then Exit Sub
        .Move .Left, .Top + y
    End With
'    Call Form_Resize
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
    Me.Refresh
End Sub

Private Sub TabShow_Click(PreviousTab As Integer)
    If mlngMode <> 模块号.药品移库 And mlngMode <> 模块号.药品领用 Then Exit Sub
    
    lbl1.Caption = ""
    lbl2.Caption = ""
    lbl3.Caption = ""
    
    mintListRow = 1
    
    Call SetMenu
    Call GetList(mstrFind)
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
        Case "PreparePhysic"
            mnuEditPreparePhysic_Click
        Case "SendPhysic"
            mnuEditSendPhysic_Click
        Case "Back"
            mnuEditBack_Click
        Case "Verify"
            mnuEditVerify_Click
        Case "Strike"
            mnuEditStrike_Click
        Case "ApplyStrike"
            mnuEditApplyStrike_Click
        Case "VerifyStrike"
            mnuEditVerifyStrike_Click
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
'        Case "Mark"
'            Call frmPurchaseMark.showMe(cboStock.ItemData(cboStock.ListIndex), Me)
    End Select
    
End Sub

Private Sub tlbTool_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim blnSuccess As Boolean
    Select Case ButtonMenu.Key
    Case "FromStore"
        frmDrawCard.ShowCard Me, "", 编辑.新增, mblnStock, , 0, blnSuccess
    Case "FromLeave"
        frmDrawCard.ShowCard Me, "", 编辑.新增, mblnStock, , 1, blnSuccess
    End Select
    
    If blnSuccess = True Then
        mintListRow = vsfList.Row + 1
        mnuViewRefresh_Click
    End If
End Sub
'设置菜单和工具按钮的可用属性
Private Sub SetEnable()
    Dim strVerify As String
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
            
            If mnuEditApplyStrike.Visible = True Then
                mnuEditApplyStrike.Enabled = False
                tlbTool.Buttons("ApplyStrike").Enabled = False
            End If
            
            If mnuEditVerifyStrike.Visible = True Then
                mnuEditVerifyStrike.Enabled = False
                tlbTool.Buttons("VerifyStrike").Enabled = False
            End If

            If mnuEditDisplay.Visible = True Then
                mnuEditDisplay.Enabled = False
            End If
            
            If mnuEditPrepare.Visible Then
                mnuEditPrepare.Enabled = False
                mnuEditBack.Enabled = False
                tlbTool.Buttons("Prepare").Enabled = False
                tlbTool.Buttons("Back").Enabled = False
            End If
            
            If mnuEditPreparePhysic.Visible Then
                mnuEditPreparePhysic.Enabled = False
                mnuEditSendPhysic.Enabled = False
                mnuEditBack.Enabled = False
                tlbTool.Buttons("PreparePhysic").Enabled = False
                tlbTool.Buttons("SendPhysic").Enabled = False
                tlbTool.Buttons("Back").Enabled = False
            End If
            
            If mnuEditBill.Visible = True Then
                mnuEditBill.Enabled = False
            End If
            If mnuEditAcc.Visible Then
                mnuEditAcc.Enabled = False
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
            If mnuEditAcc.Visible Then
                mnuEditAcc.Enabled = False
            End If
            
            If mlngMode = 模块号.药品移库 Then
                If TabShow.Tab = 0 Then
                    strVerify = .TextMatrix(.Row, .Cols - 6)
                Else
                    strVerify = .TextMatrix(.Row, .Cols - 4)
                End If
            Else
                strVerify = .TextMatrix(.Row, .Cols - 4)
            End If
            
            If mlngMode = 模块号.药品领用 Then strVerify = .TextMatrix(.Row, .Cols - 5)
            
            
            If strVerify = "" Then    '未审核单
                If mnuEditModify.Visible = True Then
                    mnuEditModify.Enabled = IIf(mlngMode = 模块号.药品移库, IIf(.TextMatrix(.Row, .Cols - 4) = "", True, False), True)
                    tlbTool.Buttons("Modify").Enabled = mnuEditModify.Enabled
                End If
                If mnuEditDel.Visible = True Then
                    mnuEditDel.Enabled = IIf(mlngMode = 模块号.药品移库, IIf(.TextMatrix(.Row, .Cols - 4) = "", True, False), True)
                    tlbTool.Buttons("Delete").Enabled = mnuEditDel.Enabled
                End If
                If mnuEditVerify.Visible = True Then
                    mnuEditVerify.Enabled = True
                    tlbTool.Buttons("Verify").Enabled = True
                End If
                If mnuEditStrike.Visible = True Then
                    mnuEditStrike.Enabled = False
                    tlbTool.Buttons("Strike").Enabled = False
                End If
                If mnuEditApplyStrike.Visible = True Then
                    mnuEditApplyStrike.Enabled = False
                    tlbTool.Buttons("ApplyStrike").Enabled = False
                End If
                If mnuEditVerifyStrike.Visible = True Then
                    mnuEditVerifyStrike.Enabled = False
                    tlbTool.Buttons("VerifyStrike").Enabled = False
                End If
                If mnuEditDisplay.Visible = True Then
                    mnuEditDisplay.Enabled = True
                End If
                
                '未审核的单据，允许重复核查
                mnuEditPrepare.Enabled = True
                tlbTool.Buttons("Prepare").Enabled = True
                
                If mnuEditBack.Visible = True Then
                    If .TextMatrix(.Row, .Cols - 6) <> "" Then
                        mnuEditBack.Enabled = True
                        tlbTool.Buttons("Back").Enabled = True
                    Else
                        mnuEditBack.Enabled = False
                        tlbTool.Buttons("Back").Enabled = False
                    End If
                End If
                
                '如果是外购入库单，检查是否通过核查，未通过核查的单据，不允许审核
                If mlngMode = 模块号.外购入库 And mbln核查 Then
                    mnuEditVerify.Enabled = TestPrepare(.TextMatrix(.Row, 0))
                    tlbTool.Buttons("Verify").Enabled = TestPrepare(.TextMatrix(.Row, 0))
                End If
                '移库单，根据当前选择的页面，当前单据设置按钮状态
                If mlngMode = 模块号.药品移库 Then
                    If TabShow.Tab = 0 Then
                        mnuEditPreparePhysic.Enabled = (.TextMatrix(.Row, .Cols - 4) = "")
                        mnuEditSendPhysic.Enabled = (.TextMatrix(.Row, .Cols - 4) <> "") And (.TextMatrix(.Row, .Cols - 3) = "")
                        mnuEditBack.Enabled = Not mnuEditPreparePhysic.Enabled
                        tlbTool.Buttons("PreparePhysic").Enabled = mnuEditPreparePhysic.Enabled
                        tlbTool.Buttons("SendPhysic").Enabled = mnuEditSendPhysic.Enabled
                        tlbTool.Buttons("Back").Enabled = mnuEditBack.Enabled
                        '如果该单据已审核，不允许备药与发送
                        If TestVerify(vsfList.TextMatrix(vsfList.Row, 0)) Then
                            mnuEditPreparePhysic.Enabled = False
                            mnuEditSendPhysic.Enabled = False
                            mnuEditBack.Enabled = False
                            tlbTool.Buttons("PreparePhysic").Enabled = False
                            tlbTool.Buttons("SendPhysic").Enabled = False
                            tlbTool.Buttons("Back").Enabled = False
                        End If
                        
                        '如果冲销单还未审核，则允许审核冲销单
                        If mint冲销申请 = 1 And Val(.TextMatrix(.Row, .Cols - 2)) Mod 3 = 2 Then
                            mnuEditPreparePhysic.Enabled = False
                            mnuEditSendPhysic.Enabled = False
                            mnuEditBack.Enabled = False
                            mnuEditDel.Enabled = False
                            tlbTool.Buttons("PreparePhysic").Enabled = False
                            tlbTool.Buttons("SendPhysic").Enabled = False
                            tlbTool.Buttons("Back").Enabled = False
                            tlbTool.Buttons("Delete").Enabled = False
                            
                            mnuEditStrike.Enabled = True
                            tlbTool.Buttons("Strike").Enabled = True
                            
                            mnuEditVerify.Enabled = False
                            tlbTool.Buttons("Verify").Enabled = mnuEditVerify.Enabled
                        End If
                    Else
                        If int移库处理流程 = 1 Then
                            mnuEditVerify.Enabled = TestPrepare(.TextMatrix(.Row, 0))
                        Else
                            mnuEditVerify.Enabled = True
                        End If
                        tlbTool.Buttons("Verify").Enabled = mnuEditVerify.Enabled
                        
                        '如果冲销单还未审核，则允许删除
                        If mint冲销申请 = 1 And Val(.TextMatrix(.Row, .Cols - 2)) Mod 3 = 2 Then
                            mnuEditModify.Enabled = False
                            tlbTool.Buttons("Modify").Enabled = False
                            mnuEditVerify.Enabled = False
                            tlbTool.Buttons("Verify").Enabled = False
                            mnuEditStrike.Enabled = False
                            tlbTool.Buttons("Strike").Enabled = False
                            
                            mnuEditDel.Enabled = True
                            tlbTool.Buttons("Delete").Enabled = True
                        End If
                    End If
                End If
                If mlngMode = 模块号.药品领用 Then
                     '如果冲销单还未审核，则允许审核冲销单
                    If mint领用冲销申请 = 1 And Val(.TextMatrix(.Row, .Cols - 4)) Mod 3 = 2 Then
                        mnuEditDel.Enabled = True
                        tlbTool.Buttons("Delete").Enabled = True
                        mnuEditModify.Enabled = False
                        tlbTool.Buttons("Modify").Enabled = False
                        mnuEditVerify.Enabled = False
                        tlbTool.Buttons("Verify").Enabled = False
                        mnuEditApplyStrike.Enabled = False
                        tlbTool.Buttons("ApplyStrike").Enabled = False
                        mnuEditVerifyStrike.Enabled = True
                        tlbTool.Buttons("VerifyStrike").Enabled = True
                    End If
                End If
            ElseIf .TextMatrix(.Row, IIf(mlngMode = 模块号.差价调整 Or mlngMode = 模块号.外购入库 Or mlngMode = 模块号.药品领用, IIf(mlngMode = 模块号.药品领用, .Cols - 4, .Cols - 3), .Cols - 2)) = 1 Then '审核单
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
                
                '库存差价调整中，如果是成本价调整单则不能冲销
                If mlngMode = 模块号.差价调整 Then
                    If Val(vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 1)) = 1 Then
                        mnuEditStrike.Enabled = False
                        tlbTool.Buttons("Strike").Enabled = False
                    End If
                End If
                
                If mlngMode = 模块号.药品领用 Then
                    If mint领用冲销申请 = 1 Then
                        mnuEditApplyStrike.Enabled = True
                        tlbTool.Buttons("ApplyStrike").Enabled = True
                        mnuEditVerifyStrike.Enabled = False
                        tlbTool.Buttons("VerifyStrike").Enabled = False
                    End If
                End If
                
                '只有外购入库单才有
                mnuEditBill.Enabled = True
                mnuEditAcc.Enabled = True
                If mnuEditPrepare.Visible Then
                    mnuEditPrepare.Enabled = False
                    mnuEditBack.Enabled = False
                    tlbTool.Buttons("Prepare").Enabled = False
                    tlbTool.Buttons("Back").Enabled = False
                End If
                If mlngMode = 模块号.药品移库 Then
                    If TabShow.Tab = 0 Then
                        '如果该单据已审核，不允许备药与发送
                        If TestVerify(vsfList.TextMatrix(vsfList.Row, 0)) Then
                            mnuEditPreparePhysic.Enabled = False
                            mnuEditSendPhysic.Enabled = False
                            mnuEditBack.Enabled = False
                            tlbTool.Buttons("PreparePhysic").Enabled = False
                            tlbTool.Buttons("SendPhysic").Enabled = False
                            tlbTool.Buttons("Back").Enabled = False
                        End If
                        If mint冲销申请 = 1 Then
                            mnuEditStrike.Enabled = False
                            tlbTool.Buttons("Strike").Enabled = False
                        End If
                    End If
                End If
            Else   '2,3 冲销单（已付款的单据不允许财务审核，同样，财务审核后的单据不允许冲销）
                If .TextMatrix(.Row, IIf(mlngMode = 模块号.差价调整 Or mlngMode = 模块号.外购入库 Or mlngMode = 模块号.药品领用, IIf(mlngMode = 模块号.药品领用, .Cols - 4, .Cols - 3), .Cols - 2)) Mod 3 = 0 Then
                    .ToolTipText = "冲销单据的原单据"
                    If mnuEditStrike.Visible = True Then
                        mnuEditStrike.Enabled = True
                        tlbTool.Buttons("Strike").Enabled = True
                    End If
                    If mnuEditApplyStrike.Visible = True Then
                        mnuEditApplyStrike.Enabled = True
                        tlbTool.Buttons("ApplyStrike").Enabled = True
                    End If
                    '允许部分冲销的单据财务审核
                    mnuEditAcc.Enabled = True
                    
                    '允许部分冲销的原单据修改发票信息
                    If mnuEditBill.Visible = True Then
                        mnuEditBill.Enabled = True
                    End If
                ElseIf .TextMatrix(.Row, IIf(mlngMode = 模块号.差价调整 Or mlngMode = 模块号.外购入库 Or mlngMode = 模块号.药品领用, IIf(mlngMode = 模块号.药品领用, .Cols - 4, .Cols - 3), .Cols - 2)) Mod 3 = 2 Then
                    If mlngMode = 模块号.外购入库 Then
                        If Val(.TextMatrix(.Row, 外购主表.冲销类型)) = 1 Then
                            .ToolTipText = "财务审核冲销单据"
                        Else
                            .ToolTipText = "冲销单据"
                        End If
                    Else
                        .ToolTipText = "冲销单据"
                    End If
                    If mnuEditStrike.Visible = True Then
                        mnuEditStrike.Enabled = False
                        tlbTool.Buttons("Strike").Enabled = False
                    End If
                    If mnuEditApplyStrike.Visible = True Then
                        mnuEditApplyStrike.Enabled = False
                        tlbTool.Buttons("ApplyStrike").Enabled = False
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
                If mnuEditVerifyStrike.Visible = True Then
                    mnuEditVerifyStrike.Enabled = False
                    tlbTool.Buttons("VerifyStrike").Enabled = False
                End If
                If mnuEditPrepare.Visible Then
                    mnuEditPrepare.Enabled = False
                    mnuEditBack.Enabled = False
                    tlbTool.Buttons("Prepare").Enabled = False
                    tlbTool.Buttons("Back").Enabled = False
                End If
                If mnuEditDisplay.Visible = True Then
                    mnuEditDisplay.Enabled = True
                End If
                If mlngMode = 模块号.药品移库 Then
                    If TabShow.Tab = 0 Then
                        mnuEditStrike.Enabled = False
                        tlbTool.Buttons("Strike").Enabled = False
                        mnuEditPreparePhysic.Enabled = False
                        mnuEditSendPhysic.Enabled = False
                        mnuEditBack.Enabled = False
                        tlbTool.Buttons("PreparePhysic").Enabled = False
                        tlbTool.Buttons("SendPhysic").Enabled = mnuEditSendPhysic.Enabled
                        tlbTool.Buttons("Back").Enabled = False
                        '如果该单据已审核，不允许备药与发送
                        If TestVerify(vsfList.TextMatrix(vsfList.Row, 0)) Then
                            mnuEditPreparePhysic.Enabled = False
                            mnuEditSendPhysic.Enabled = False
                            mnuEditBack.Enabled = False
                            tlbTool.Buttons("PreparePhysic").Enabled = False
                            tlbTool.Buttons("SendPhysic").Enabled = False
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
        
    objRow.Add "打印人:" & UserInfo.用户姓名
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
    objRow.Add "NO." & Trim(vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "NO")))
    objPrint.UnderAppRows.Add objRow
    
    Select Case mlngMode
        Case 模块号.外购入库       '药品外购入库管理
            Set objRow = New zlTabAppRow
            objRow.Add "库房：" & Trim(cboStock.Text)
            objRow.Add "供应商：" & Trim(vsfList.TextMatrix(vsfList.Row, 外购主表.供应商))
            objPrint.UnderAppRows.Add objRow
                
        Case 模块号.自制入库       '药品自制入库管理
            Set objRow = New zlTabAppRow
            objRow.Add "库房：" & Trim(cboStock.Text)
            objRow.Add "制剂室：" & Trim(vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "制剂室")))
            objPrint.UnderAppRows.Add objRow
            
        Case 模块号.其他入库       '药品其他入库管理
            Set objRow = New zlTabAppRow
            objRow.Add "库房：" & Trim(cboStock.Text)
            objRow.Add "入出类别：" & Trim(vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "入出类别")))
            objPrint.UnderAppRows.Add objRow
        Case 模块号.差价调整       '库存差价调整管理
            Set objRow = New zlTabAppRow
            objRow.Add "库房：" & Trim(cboStock.Text)
            objPrint.UnderAppRows.Add objRow
            
        Case 模块号.药品移库       '药品移库管理
            Set objRow = New zlTabAppRow
            If TabShow.Tab = 0 Then
                objRow.Add "移出库房：" & Trim(cboStock.Text)
                objRow.Add "移入库房：" & Trim(vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "移入库房")))
            Else
                objRow.Add "移入库房：" & Trim(cboStock.Text)
                objRow.Add "移出库房：" & Trim(vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "移出库房")))
            End If
            objPrint.UnderAppRows.Add objRow
            
        Case 模块号.药品领用       '药品领用管理
            Set objRow = New zlTabAppRow
            objRow.Add "发药库房：" & Trim(cboStock.Text)
            objRow.Add "领用部门：" & Trim(vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "领用部门")))
            objPrint.UnderAppRows.Add objRow
            
        Case 模块号.其他出库       '药品其他出库管理
            Set objRow = New zlTabAppRow
            objRow.Add "库房：" & Trim(cboStock.Text)
            objRow.Add "入出类别：" & Trim(vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "入出类别")))
            objPrint.UnderAppRows.Add objRow
    End Select
        
    Set objRow = New zlTabAppRow
    objRow.Add "摘要:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "摘要"))
    objPrint.BelowAppRows.Add objRow
        
    Set objRow = New zlTabAppRow
    objRow.Add "填制人:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "填制人")) & "  填制日期:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "填制日期"))
    
    If mlngMode = 模块号.药品移库 Then
        objRow.Add "审核人:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "接收人")) & "  审核日期:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "接收日期"))
    Else
        objRow.Add "审核人:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "审核人")) & "  审核日期:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "审核日期"))
    End If
    
    objPrint.BelowAppRows.Add objRow
    
    Set objPrint.Body = vsfDetail
    zlPrintOrView1Grd objPrint, bytMode
End Sub




Private Sub tlbTool_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool
    End If
End Sub
'寻找与某一列相等的行
Private Function FindRow(ByVal FlexTemp As MSHFlexGrid, ByVal intTemp As Variant, ByVal intCol As Integer) As Integer
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

Private Function TestPrepare(ByVal strNo As String) As Boolean
    Dim intBill As Integer
    Dim rsTemp As New ADODB.Recordset
    '检查配药人是否已经填写
    
    On Error GoTo errHandle
    Select Case mlngMode
    Case 模块号.外购入库
        intBill = 1
    Case 模块号.药品移库
        intBill = 6
    Case Else
        Exit Function
    End Select
    
    gstrSQL = "Select 配药人 From 药品收发记录 Where 单据=[1] And NO=[2] And Rownum<2"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "检查是否通过核查", intBill, strNo)

    If Not IsNull(rsTemp!配药人) Then
        TestPrepare = True
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function TestDelete(ByVal strNo As String) As Boolean
    Dim intBill As Integer
    Dim rsTemp As New ADODB.Recordset
    '检查单据是否删除
    On Error GoTo errHandle

    Select Case mlngMode
    Case 模块号.外购入库
        intBill = 1
    Case 模块号.药品移库
        intBill = 6
    Case Else
        Exit Function
    End Select
    
    gstrSQL = "Select id From 药品收发记录 Where 单据=[1] And NO=[2] And Rownum<2"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "检查是否通过核查", intBill, strNo)
    
    TestDelete = (rsTemp.RecordCount = 0)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function TestVerify(ByVal strNo As String) As Boolean
    Dim int单据 As Integer
    Dim rsTemp As New ADODB.Recordset
    '检查该单据是否通过审核
    On Error GoTo errHandle

    Select Case mlngMode
        Case 模块号.外购入库
            int单据 = 1
        Case 模块号.药品移库
            int单据 = 6
    End Select
    
    gstrSQL = "Select 审核人 From 药品收发记录 " & _
        " Where 单据=[1] And NO=[2] And Rownum<2"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "检查是否通过审核", int单据, strNo)
    
    If Not IsNull(rsTemp!审核人) Then
        TestVerify = True
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ShowColor(ByVal rsDetail As ADODB.Recordset)
    Dim lngCol As Long, lngCols As Long
    Dim lngRow As Long, lngRows As Long
    Dim bln招标药品 As Boolean
    Dim dbl差价让利比 As Double
    
    '为外购入库单上色
    If mlngMode <> 模块号.外购入库 Then Exit Sub
    If rsDetail.State = 0 Then Exit Sub
    If rsDetail.RecordCount = 0 Then Exit Sub
    vsfDetail.Redraw = flexRDNone
    lngRows = vsfDetail.rows - 1
    lngCols = vsfDetail.Cols - 1
    rsDetail.MoveFirst
    
    For lngRow = 1 To lngRows
        '招标药品需要上色
        vsfDetail.Row = lngRow
        bln招标药品 = (nvl(rsDetail!招标药品, 0) = 1)
        dbl差价让利比 = nvl(rsDetail!差价让利比, 0)
        
        If bln招标药品 Then
            vsfDetail.Cell(flexcpForeColor, lngRow, 0, lngRow, lngCols) = IIf(dbl差价让利比 = 0, &H800000, &H800080)
        Else
            vsfDetail.Cell(flexcpForeColor, lngRow, 0, lngRow, lngCols) = IIf(dbl差价让利比 = 0, &H0, &H40&)
        End If

        rsDetail.MoveNext
    Next
    
    vsfDetail.Row = 1
    vsfDetail.Col = 0: vsfDetail.ColSel = lngCols
    vsfDetail.Redraw = flexRDDirect
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

Private Sub CheckNumber()
    '如果填写数量和实际数量不一致，则用红色字体标注实际数量用以提醒
    Dim intRow As Integer
    Dim blnColor As Boolean

    With vsfDetail
        If .TextMatrix(1, 1) = "" Then Exit Sub
        For intRow = 1 To .rows - 1
            blnColor = False
            If .TextMatrix(intRow, .ColIndex("药品ID")) = "" Then Exit Sub
            If Val(.TextMatrix(intRow, .ColIndex("填写数量"))) <> Val(.TextMatrix(intRow, .ColIndex("实际数量"))) Then blnColor = True
            .Cell(flexcpForeColor, intRow, .ColIndex("实际数量"), intRow, .ColIndex("实际数量")) = IIf(blnColor, vbRed, vbBlack)
        Next
    End With
                
End Sub
