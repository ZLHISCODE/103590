VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmCheckMain 
   ClientHeight    =   4980
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9480
   Icon            =   "frmCheckMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picColor 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   5280
      ScaleHeight     =   255
      ScaleWidth      =   3615
      TabIndex        =   15
      Top             =   4320
      Width           =   3615
      Begin VB.PictureBox picColor3 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2280
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
         Top             =   0
         Width           =   260
      End
      Begin VB.Label lblColor3 
         AutoSize        =   -1  'True
         Caption         =   "财务冲销"
         Height          =   180
         Left            =   2640
         TabIndex        =   21
         Top             =   30
         Width           =   720
      End
      Begin VB.Label lblColor1 
         AutoSize        =   -1  'True
         Caption         =   "正常冲销"
         Height          =   180
         Left            =   360
         TabIndex        =   20
         Top             =   37
         Width           =   720
      End
      Begin VB.Label lblColor2 
         AutoSize        =   -1  'True
         Caption         =   "正常"
         Height          =   180
         Left            =   1680
         TabIndex        =   19
         Top             =   37
         Width           =   360
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
      Height          =   1155
      Left            =   0
      TabIndex        =   12
      Top             =   3000
      Width           =   6255
      _cx             =   11033
      _cy             =   2037
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
      BackColorAlternate=   15724527
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
      FormatString    =   $"frmCheckMain.frx":030A
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
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   1455
      Left            =   0
      TabIndex        =   11
      Top             =   1040
      Width           =   6255
      _cx             =   11033
      _cy             =   2566
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
      FormatString    =   $"frmCheckMain.frx":037F
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
      Height          =   360
      Left            =   30
      TabIndex        =   7
      Top             =   720
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   635
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "盘点记录单清单(&1)"
      TabPicture(0)   =   "frmCheckMain.frx":03F4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "盘点表清单(&2)"
      TabPicture(1)   =   "frmCheckMain.frx":0410
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
   End
   Begin MSComctlLib.ImageList ilsHot 
      Left            =   7110
      Top             =   720
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
            Picture         =   "frmCheckMain.frx":042C
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":064C
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":086C
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":0A88
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":0CA8
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":0EC8
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":10E4
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":1300
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":151A
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":1734
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":188E
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":1AAE
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsCold 
      Left            =   6510
      Top             =   720
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
            Picture         =   "frmCheckMain.frx":1CCE
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":1EEE
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":210E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":232A
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":254A
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":276A
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":2986
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":2BA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":2DBC
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":2FD6
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":3130
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":334C
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Cmd查阅 
      Caption         =   "查阅(&V)"
      Height          =   350
      Left            =   8040
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2555
      Width           =   1100
   End
   Begin VB.PictureBox picSeparate_s 
      BorderStyle     =   0  'None
      Height          =   370
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   375
      ScaleWidth      =   7935
      TabIndex        =   4
      Top             =   2520
      Width           =   7935
      Begin VB.Label lbl成本金额差 
         AutoSize        =   -1  'True
         Caption         =   "成本金额差合计："
         Height          =   180
         Left            =   6480
         TabIndex        =   14
         Top             =   0
         Width           =   1440
      End
      Begin VB.Label lblSum成本金额 
         AutoSize        =   -1  'True
         Caption         =   "盘点成本金额合计："
         Height          =   180
         Left            =   4680
         TabIndex        =   13
         Top             =   0
         Width           =   1620
      End
      Begin VB.Label lbl3 
         AutoSize        =   -1  'True
         Caption         =   "账面金额差合计："
         Height          =   180
         Left            =   3000
         TabIndex        =   10
         Top             =   0
         Width           =   1440
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         Caption         =   "盘点金额合计："
         Height          =   180
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   1260
      End
      Begin VB.Label lbl2 
         AutoSize        =   -1  'True
         Caption         =   "金额差合计："
         Height          =   180
         Left            =   1680
         TabIndex        =   8
         Top             =   0
         Width           =   1080
      End
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "查询范围：1999年8月12日至1999年9月12日"
         Height          =   180
         Left            =   0
         TabIndex        =   5
         Top             =   200
         Width           =   3420
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
         Width           =   7875
         _ExtentX        =   13891
         _ExtentY        =   1270
         ButtonWidth     =   1138
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
               Caption         =   "记录单"
               Key             =   "Bill"
               Description     =   "增加"
               Object.ToolTipText     =   "记录单"
               Object.Tag             =   "记录单"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "盘点表"
               Key             =   "Table"
               Object.ToolTipText     =   "盘点表"
               Object.Tag             =   "盘点表"
               ImageIndex      =   3
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Auto"
                     Text            =   "自动产生盘点表"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Total"
                     Text            =   "汇总记录单产生盘点表"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Zero"
                     Text            =   "全部盘为零"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修改"
               Key             =   "Modify"
               Description     =   "修改"
               Object.ToolTipText     =   "修改"
               Object.Tag             =   "修改"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "Delete"
               Description     =   "删除"
               Object.ToolTipText     =   "删除"
               Object.Tag             =   "删除"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "EditSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "审核"
               Key             =   "Verify"
               Description     =   "审核"
               Object.ToolTipText     =   "审核"
               Object.Tag             =   "审核"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "冲销"
               Key             =   "Strike"
               Description     =   "冲销"
               Object.ToolTipText     =   "冲销"
               Object.Tag             =   "冲销"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "VerifySeparate"
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "确认"
               Key             =   "Affirmant"
               Object.ToolTipText     =   "月度确认"
               Object.Tag             =   "确认"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "AffirmantSplit"
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "过滤"
               Key             =   "Search"
               Description     =   "过滤"
               Object.ToolTipText     =   "过滤"
               Object.Tag             =   "过滤"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "刷新"
               Key             =   "Refresh"
               Description     =   "刷新"
               Object.ToolTipText     =   "刷新"
               Object.Tag             =   "刷新"
               ImageIndex      =   10
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
               ImageIndex      =   11
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Exit"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   12
            EndProperty
         EndProperty
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
      Begin VB.Menu mnuEditAddBill 
         Caption         =   "增加记录单(&B)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditAddTable 
         Caption         =   "增加盘点表(&T)"
         Begin VB.Menu mnuEditAddTableAuto 
            Caption         =   "自动产生盘点表(&A)"
         End
         Begin VB.Menu mnuEditAddTableTotal 
            Caption         =   "汇总记录单产生盘点表(&T)"
         End
         Begin VB.Menu mnuEditAddTableZero 
            Caption         =   "全部盘为零(&Z)"
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
      Begin VB.Menu mnuEditVerify 
         Caption         =   "审核(&C)"
      End
      Begin VB.Menu mnuEditStrike 
         Caption         =   "冲销(&K)"
      End
      Begin VB.Menu mnuEditLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditAffirmant 
         Caption         =   "月度确认(&O)"
      End
      Begin VB.Menu mnuEditAffirmantSplit 
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
Attribute VB_Name = "frmCheckMain"
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
Private mblnViewCost As Boolean         '查看成本价
'Private Const mstrTitle As String = "药品盘点管理"

Public mstrPrivs As String              '权限

'日期设置
Private strStart As Date
Private strEnd As Date
Private strVerifyStart As Date
Private strVerifyEnd As Date

Private mlng库房ID As Long
Private mintUnit As Integer                 '单位系数：1-售价;2-门诊;3-住院;4-药库

Private Const mcstComment As String = "黑-盘平;红-盘盈;蓝-盘亏;粗体-停用药品"

'从参数表中取药品价格、数量、金额小数位数（显示精度）
Private mintShowCostDigit As Integer            '成本价小数位数
Private mintShowPriceDigit As Integer           '售价小数位数
Private mintShowNumberDigit As Integer          '数量小数位数
Private mintShowMoneyDigit As Integer           '金额小数位数

Private mstrCostFormat As String
Private mstrPriceFormat As String
Private mstrNumberFormat As String
Private mstrMoneyFormat As String

Private Const mconint售价单位 As Integer = 1
Private Const mconint门诊单位 As Integer = 2
Private Const mconint住院单位 As Integer = 3
Private Const mconint药库单位 As Integer = 4

Private Type Type_SQLCondition
    strNO开始 As String
    strNO结束 As String
    date填制时间开始 As Date
    date填制时间结束 As Date
    date审核时间开始 As Date
    date审核时间结束 As Date
    lng药品 As Long
    lng移入库房 As Long
    str填制人 As String
    str审核人 As String
    lng药品分类 As Long
    str剂型 As String
End Type

Private SQLCondition As Type_SQLCondition
Private Sub cboStock_Click()
    If mlng库房ID <> Me.cboStock.ItemData(Me.cboStock.ListIndex) Then
        mlng库房ID = Me.cboStock.ItemData(Me.cboStock.ListIndex)
        Call GetDrugDigit(mlng库房ID, mstrTitle, mintUnit, mintShowCostDigit, mintShowPriceDigit, mintShowNumberDigit, mintShowMoneyDigit)
        
        '重新组织格式化串
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
    Dim dateCurrentDate As Date
    Dim strTemp As String
    Dim int查询天数 As Integer
    
    mblnBootUp = False
    mlngMode = lngMode
    mstrTitle = strTitle
    mstrPrivs = gstrprivs
    Me.Caption = strTitle
    
    If Not CheckDepend Then Exit Sub            '数据依赖性测试
    
    mlng库房ID = Me.cboStock.ItemData(Me.cboStock.ListIndex)
    Call GetDrugDigit(mlng库房ID, mstrTitle, mintUnit, mintShowCostDigit, mintShowPriceDigit, mintShowNumberDigit, mintShowMoneyDigit)
    
    '组织格式化串
    mstrCostFormat = "'999999999990." & String(mintShowCostDigit, "0") & "'"
    mstrPriceFormat = "'999999999990." & String(mintShowPriceDigit, "0") & "'"
    mstrNumberFormat = "'999999999990." & String(mintShowNumberDigit, "0") & "'"
    mstrMoneyFormat = "'999999999990." & String(mintShowMoneyDigit, "0") & "'"
        
    strVerifyStart = "1901-01-01"
    strVerifyEnd = "1901-01-01"
    
    dateCurrentDate = Sys.Currentdate
    int查询天数 = Val(zlDatabase.GetPara("查询天数", glngSys, mlngMode, 7)) - 1
    strStart = Format(DateAdd("d", -int查询天数, dateCurrentDate), "yyyy-MM-dd")
    strEnd = Format(dateCurrentDate, "yyyy-MM-dd")
    
    strFind = " AND A.记录状态 = 1 And A.审核日期 is Null And A.填制日期 Between [3] And [4] "
    SQLCondition.date填制时间开始 = CDate(Format(strStart, "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date填制时间结束 = CDate(Format(strEnd, "yyyy-mm-dd") & " 23:59:59")
    
    mstrFind = strFind
    
    SetVisable  '根据权限设置不同的显示项目
    TabShow.Tab = 0
    GetList (mstrFind)  '列出单据头
    
    RestoreWinState Me, App.ProductName, mstrTitle
        
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
    
    Dim rsDepend As New Recordset
    Dim strStock As String
    
    On Error GoTo errHandle
    CheckDepend = False
    
    strStock = "HIJKLMN"
    gstrSQL = "SELECT DISTINCT a.id, a.名称 " _
             & "FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " _
            & "Where (a.站点 = '" & gstrNodeNo & "' Or a.站点 is Null) And c.工作性质 = b.名称 " _
              & "AND Instr([1],b.编码,1) > 0 " _
             & " AND a.id = c.部门id " _
              & "AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01'" _
              & IIf(zlStr.IsHavePrivs(mstrPrivs, "所有库房"), "", " And a.ID IN (Select 部门ID From 部门人员 Where 人员ID=[2])")

    Set rsDepend = zlDatabase.OpenSQLRecord(gstrSQL, mstrTitle, strStock, UserInfo.用户ID)
    
    If rsDepend.EOF Then
        MsgBox "至少应该设置一个具有药库性质，药房性质，或者制剂室性质的部门,请查看部门管理！", vbInformation, gstrSysName
        rsDepend.Close
        Exit Function
    End If
            
    With cboStock
        .Clear
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

Private Sub GetList(ByVal strFind As String)
    Dim rsList As New Recordset
    Dim strUserPart As String
    Dim str包装系数 As String
    Dim strSqlForm As String
    Dim n As Integer
    
    '用于统计合计金额
    Dim dbl1 As Double
    Dim dbl2 As Double
    Dim dbl3 As Double
    Dim dbl盘点成本金额 As Double
    Dim dbl盘点金额差 As Double

    mlastRow = 0
    On Error GoTo errHandle

    Call FS.ShowFlash("正在搜索药品记录,请稍候 ...", Me)
    DoEvents
    Screen.MousePointer = vbHourglass
    strUserPart = " And A.库房ID+0=[11] "
    
    Select Case mintUnit
        Case mconint售价单位
            str包装系数 = "1"
        Case mconint门诊单位
            str包装系数 = "B.门诊包装"
        Case mconint住院单位
            str包装系数 = "B.住院包装"
        Case mconint药库单位
            str包装系数 = "B.药库包装"
    End Select
    
    vsfList.Redraw = flexRDNone
    '频次字段保存的 盘点时间
    If TabShow.Tab = 1 Then
        If SQLCondition.str剂型 <> "" And SQLCondition.lng药品分类 = 0 Then
            strSqlForm = " , 诊疗项目目录 G, 药品特性 H"
            strFind = strFind & " And b.药名id = g.Id And g.Id = h.药名id(+) and h.药品剂型 in (select * from Table(Cast(f_Str2list([13]) As zlTools.t_Strlist))) and (g.类别='5' or g.类别='6' or g.类别='7')"
        ElseIf SQLCondition.str剂型 = "" And SQLCondition.lng药品分类 <> 0 Then
            strSqlForm = " , 诊疗项目目录 G"
            strFind = strFind & " And b.药名id = g.Id And g.分类id + 0=[12] and (g.类别='5' or g.类别='6' or g.类别='7')"
        ElseIf SQLCondition.str剂型 <> "" And SQLCondition.lng药品分类 <> 0 Then
            strSqlForm = " , 诊疗项目目录 G, 药品特性 H"
            strFind = strFind & " And b.药名id = g.Id And g.Id = h.药名id(+) and h.药品剂型 in (select * from Table(Cast(f_Str2list([13]) As zlTools.t_Strlist))) and (g.类别='5' or g.类别='6' or g.类别='7') and g.分类id + 0=[12]"
        End If
        
        gstrSQL = "Select NO, 盘点时间, 填制人, 填制日期, 审核人, 审核日期, " & _
                "   to_char(Sum(盘点金额), " & mstrMoneyFormat & ") 盘点金额, to_char(Sum(金额差), " & mstrMoneyFormat & ") 金额差,to_char(Sum(账面金额差), " & mstrMoneyFormat & ") 账面金额差,to_char(Sum(盘点成本金额)," & mstrMoneyFormat & ") 盘点成本金额, to_char(Sum(成本金额差)," & mstrMoneyFormat & ") 成本金额差, 记录状态, 摘要" & _
                " from ( SELECT a.no,a.序号, 频次 AS 盘点时间," _
                & "a.填制人,TO_CHAR (min(a.填制日期), 'yyyy-mm-dd HH24:Mi:SS') AS 填制日期, a.审核人," _
                & "TO_CHAR (min(a.审核日期), 'yyyy-mm-dd HH24:Mi:SS') AS 审核日期, " _
                & "ltrim(to_char(A.成本价+A.入出系数*A.零售金额*Decode(记录状态, 1, 1, Decode(Mod(记录状态, 3), 0, 1, -1))," & mstrMoneyFormat & ")) 盘点金额," _
                & "ltrim(to_char(零售金额*a.入出系数," & mstrMoneyFormat & ")) 金额差," _
                & "ltrim(to_char((A.扣率-A.填写数量) * a.零售价* Decode(记录状态, 1, 1, Decode(Mod(记录状态, 3), 0, 1, -1))," & mstrMoneyFormat & ")) AS 账面金额差," _
                & "ltrim(to_char((a.成本价+to_char(a.零售金额*a.入出系数*Decode(a.记录状态, 1, 1, Decode(Mod(a.记录状态, 3), 0, 1, -1))," & mstrMoneyFormat & "))-(a.成本金额+to_char(a.差价*a.入出系数*Decode(a.记录状态, 1, 1, Decode(Mod(a.记录状态, 3), 0, 1, -1))," & mstrMoneyFormat & "))," & mstrMoneyFormat & ")) as 盘点成本金额," _
                & "ltrim(to_char(a.零售金额*a.入出系数-a.差价*a.入出系数," & mstrMoneyFormat & ")) as 成本金额差," _
                & " a.记录状态, a.摘要 " _
                & " FROM 药品收发记录 a,药品规格 B " & strSqlForm _
                & " Where A.药品ID=B.药品ID And A.单据 = 12  " & strUserPart & strFind _
                & " Group By a.No,a.序号, 频次, a.填制人, a.审核人, a.成本价, a.入出系数, a.成本价,a.成本金额," & str包装系数 & ", a.零售金额, a.记录状态, a.扣率, a.填写数量, a.零售价, a.单量, a.差价, a.摘要) " _
                & " Group By NO, 盘点时间, 填制人, 填制日期, 审核人, 审核日期, 记录状态, 摘要 ORDER BY no DESC,填制日期 ASC"
    Else
        If SQLCondition.str剂型 <> "" And SQLCondition.lng药品分类 = 0 Then
            strSqlForm = " , 药品规格 F, 诊疗项目目录 G, 药品特性 H"
            strFind = strFind & " and a.药品id = f.药品id And f.药名id = g.Id And g.Id = h.药名id(+) and h.药品剂型 in (select * from Table(Cast(f_Str2list([13]) As zlTools.t_Strlist))) and (g.类别='5' or g.类别='6' or g.类别='7')"
        ElseIf SQLCondition.str剂型 = "" And SQLCondition.lng药品分类 <> 0 Then
            strSqlForm = " , 药品规格 F, 诊疗项目目录 G"
            strFind = strFind & " and a.药品id = f.药品id And f.药名id = g.Id And g.分类id + 0=[12] and (g.类别='5' or g.类别='6' or g.类别='7')"
        ElseIf SQLCondition.str剂型 <> "" And SQLCondition.lng药品分类 <> 0 Then
            strSqlForm = " , 药品规格 F, 诊疗项目目录 G, 药品特性 H"
            strFind = strFind & " and a.药品id = f.药品id And f.药名id = g.Id And g.Id = h.药名id(+) and h.药品剂型 in(select * from Table(Cast(f_Str2list([13]) As zlTools.t_Strlist))) and (g.类别='5' or g.类别='6' or g.类别='7') and g.分类id + 0=[12]"
        End If
        gstrSQL = " SELECT a.no, 频次 AS 盘点时间," _
                    & "a.填制人,TO_CHAR (min(a.填制日期), 'yyyy-mm-dd HH24:Mi:SS') AS 填制日期,a.摘要 " _
                    & " FROM 药品收发记录 a " & strSqlForm _
                    & " Where a.单据 = 14  " & strUserPart & strFind _
                    & " Group by a.no,频次,a.填制人,a.摘要 " _
                    & " ORDER BY no DESC,填制日期 ASC "
    End If
    
    Set rsList = zlDatabase.OpenSQLRecord(gstrSQL, mstrTitle, _
        SQLCondition.strNO开始, _
        SQLCondition.strNO结束, _
        SQLCondition.date填制时间开始, _
        SQLCondition.date填制时间结束, _
        SQLCondition.date审核时间开始, _
        SQLCondition.date审核时间结束, _
        SQLCondition.lng药品, _
        SQLCondition.lng移入库房, _
        SQLCondition.str填制人, _
        SQLCondition.str审核人, _
        cboStock.ItemData(cboStock.ListIndex), _
        SQLCondition.lng药品分类, _
        SQLCondition.str剂型)
        
    Set vsfList.DataSource = rsList
    With vsfList
        If .rows = 1 Then
            .rows = .rows + 100
            .Row = 1
            .Redraw = flexRDDirect
            
            .TopRow = 1
            .rows = .rows - 99
        End If
        .ColAlignment(.ColIndex("盘点成本金额")) = flexAlignRightCenter
        .ColAlignment(.ColIndex("盘点金额")) = flexAlignRightCenter
        .ColAlignment(.ColIndex("金额差")) = flexAlignRightCenter
        .ColAlignment(.ColIndex("账面金额差")) = flexAlignRightCenter
        .ColAlignment(.ColIndex("成本金额差")) = flexAlignRightCenter
        .Row = 1
        .Col = 0
        .ColSel = .Cols - 1
        If TabShow.Tab = 1 Then
            .ColWidth(.Cols - 2) = 0         '始终隐藏"记录状态"这一列
        End If
        
        For n = 0 To .Cols - 1
            .FixedAlignment(n) = flexAlignCenterCenter
        Next
    End With
    SetListColWidth
    
    '统计合计金额
    lbl1.Caption = "盘点金额合计："
    lbl2.Caption = "金额差合计："
    lbl3.Caption = "账面金额差合计："
    
    If TabShow.Tab = 1 Then
        If mblnViewCost = False Then
            lblSum成本金额.Visible = False
            lbl成本金额差.Visible = False
        Else
            lblSum成本金额.Visible = True
            lbl成本金额差.Visible = True
        End If
        If (Not rsList.EOF) And (Not rsList.BOF) Then
            rsList.MoveFirst
            Do While Not rsList.EOF
                dbl1 = dbl1 + IIf(IsNull(rsList!盘点金额), 0, rsList!盘点金额)
                dbl2 = dbl2 + IIf(IsNull(rsList!金额差), 0, rsList!金额差)
                dbl3 = dbl3 + IIf(IsNull(rsList!账面金额差), 0, rsList!账面金额差)
                dbl盘点成本金额 = dbl盘点成本金额 + IIf(IsNull(rsList!盘点成本金额), 0, rsList!盘点成本金额)
                dbl盘点金额差 = dbl盘点金额差 + IIf(IsNull(rsList!成本金额差), 0, rsList!成本金额差)
                rsList.MoveNext
            Loop
            rsList.MoveFirst
            
            lbl1.Caption = "盘点金额合计：" & Format(dbl1, "0." & String(mintShowMoneyDigit, "0"))
            lbl2.Caption = "金额差合计：" & Format(dbl2, "0." & String(mintShowMoneyDigit, "0"))
            lbl3.Caption = "账面金额差合计：" & Format(dbl3, "0." & String(mintShowMoneyDigit, "0"))
            lblSum成本金额.Caption = "盘点成本金额合计：" & Format(dbl盘点成本金额, "0." & String(mintShowMoneyDigit, "0"))
            lbl成本金额差.Caption = "成本金额差：" & Format(dbl盘点金额差, "0." & String(mintShowMoneyDigit, "0"))
        End If
    Else
        lblSum成本金额.Visible = False
        lbl成本金额差.Visible = False
    End If
    
    lbl2.Left = lbl1.Width + lbl1.Left + 200
    lbl3.Left = lbl2.Width + lbl2.Left + 200
    lblSum成本金额.Left = lbl3.Width + lbl3.Left + 200
    lbl成本金额差.Left = lblSum成本金额.Width + lblSum成本金额.Left + 200
    
    vsfList_EnterCell    '列出单据体
    
    SetStrikeColor
    With vsfList
        .Row = 1
        .Col = 0
        .ColSel = .Cols - 1
    End With
    vsfList.Redraw = flexRDDirect
    Call FS.StopFlash
    Screen.MousePointer = vbDefault
    staThis.Panels(2).Text = "当前共有" & rsList.RecordCount & "张单据"
    rsList.Close
    If vsfList.Visible = True Then
        vsfList.SetFocus
        vsfList.Row = 1
    End If
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
            intStatus = IIf(TabShow.Tab = 0, 1, Val(.TextMatrix(intRow, .Cols - 2)))
            If intStatus Mod 3 = 0 Then
                .Cell(flexcpForeColor, intRow, 0, intRow, .Cols - 1) = &H80000001
            End If
            If intStatus Mod 3 = 2 Then
                .Cell(flexcpForeColor, intRow, 0, intRow, .Cols - 1) = &HFF
            End If
        Next
    End With
End Sub

'表头列宽初始
Private Sub SetListColWidth()
    Dim intCol As Integer
    
    With vsfList
        If TabShow.Tab = 1 Then
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
        Else
            If mblnBootUp = False Then
                .ColWidth(1) = 2000
                .ColWidth(4) = 3000
            End If
        End If
        .ColWidth(.ColIndex("盘点成本金额")) = 1500
    End With
    
    Call RestoreFlexState(vsfList, TabShow.TabCaption(TabShow.Tab))
    If TabShow.Tab = 1 And mblnViewCost = False Then
        vsfList.ColHidden(vsfList.ColIndex("盘点成本金额")) = True
        vsfList.ColHidden(vsfList.ColIndex("成本金额差")) = True
    End If
End Sub

Private Sub SetDetailColWidth()
    Dim intCol As Integer
    
    With vsfDetail
        .ColAlignment(.ColIndex("单位")) = flexAlignCenterCenter    '单位
        .ColAlignment(.ColIndex("实盘数")) = flexAlignRightCenter '实盘数
        If TabShow.Tab = 1 Then
            .ColAlignment(.ColIndex("帐面数")) = flexAlignRightCenter     '帐面数
            .ColAlignment(.ColIndex("标志")) = flexAlignCenterCenter    '标志
            .ColAlignment(.ColIndex("数量差")) = flexAlignRightCenter     '数量差
            .ColAlignment(.ColIndex("成本价")) = flexAlignRightCenter    '成本价
            .ColAlignment(.ColIndex("售价")) = flexAlignRightCenter    '售价
            .ColAlignment(.ColIndex("金额差")) = flexAlignRightCenter    '金额差
            .ColAlignment(.ColIndex("差价差")) = flexAlignRightCenter    '差价差
            .ColAlignment(.ColIndex("盘点金额")) = flexAlignRightCenter    '盘点金额
            .ColAlignment(.ColIndex("账面金额差")) = flexAlignRightCenter    '账面金额差
            .ColAlignment(.ColIndex("盘点成本金额")) = flexAlignRightCenter    '盘点成本金额
            .ColAlignment(.ColIndex("成本金额差")) = flexAlignRightCenter    '成本金额差
            
        End If
        
        If TabShow.Tab = 1 Then
            If mblnBootUp = False Then
                .ColWidth(0) = 500
                .ColWidth(.ColIndex("药品信息")) = 2500
                For intCol = 2 To .Cols - 1
                    .ColWidth(intCol) = 1000
                Next
                .ColWidth(.ColIndex("撤档时间")) = 0
                .ColWidth(.ColIndex("盘点成本金额")) = 1500
            End If
        Else
            .ColWidth(0) = 500
            .ColWidth(.ColIndex("药品信息")) = 2500
            For intCol = 2 To .Cols - 1
                .ColWidth(intCol) = 1000
            Next
        End If
        
        Call RestoreFlexState(vsfDetail, TabShow.TabCaption(TabShow.Tab))
        If TabShow.Tab = 1 And mblnViewCost = False Then
            .ColHidden(.ColIndex("成本价")) = True
            .ColHidden(.ColIndex("差价差")) = True
            .ColHidden(.ColIndex("盘点成本金额")) = True
            .ColHidden(.ColIndex("成本金额差")) = True
        End If
    End With
End Sub


'根据权限设置不同的显示项目
Private Sub SetVisable()
    '外购入库所有权限：参数设置、基本、所有库房、登记、修改、删除、验收、冲销、单据打印
'    If Not zlStr.IsHavePrivs(mstrPrivs, "参数设置") Then
'         mnuFileParameter.Visible = False
'         mnuFileLine3.Visible = False                '相应的分割线
'    End If
     
    If Not zlStr.IsHavePrivs(mstrPrivs, "登记") Then
        mnuEditAddBill.Visible = False
        mnuEditAddTable.Visible = False
        tlbTool.Buttons("Bill").Visible = False
        tlbTool.Buttons("Table").Visible = False
    End If
    
    If Not zlStr.IsHavePrivs(mstrPrivs, "修改") Then
        mnuEditModify.Visible = False
        tlbTool.Buttons("Modify").Visible = False
    End If
    
    If Not zlStr.IsHavePrivs(mstrPrivs, "删除") Then
        mnuEditDel.Visible = False
        tlbTool.Buttons("Delete").Visible = False
         '对没有所有编辑权限时，把菜单和工具栏上的相应的分割线屏蔽。
        If mnuEditAddBill.Visible = False And mnuEditModify.Visible = False Then
            mnuEditLine1.Visible = False
            tlbTool.Buttons("EditSeparate").Visible = False
        End If
    End If
    
    If Not zlStr.IsHavePrivs(mstrPrivs, "审核") Then
        mnuEditVerify.Visible = False
        tlbTool.Buttons("Verify").Visible = False
    End If
    
    If Not zlStr.IsHavePrivs(mstrPrivs, "月度确认") Then
        mnuEditAffirmant.Visible = False
        mnuEditAffirmantSplit.Visible = False
        tlbTool.Buttons("Affirmant").Visible = False
        tlbTool.Buttons("AffirmantSplit").Visible = False
    End If
    
    If Not zlStr.IsHavePrivs(mstrPrivs, "冲销") Then
        mnuEditStrike.Visible = False
        tlbTool.Buttons("Strike").Visible = False
        
        If mnuEditVerify.Visible = False Then
            mnuEditLine2.Visible = False
            tlbTool.Buttons("VerifySeparate").Visible = False
        End If
    End If
    If Not zlStr.IsHavePrivs(mstrPrivs, "全部盘为零") Then
        mnuEditAddTableZero.Visible = False
        tlbTool.Buttons("Table").ButtonMenus("Zero").Visible = False
    End If
    If Not zlStr.IsHavePrivs(mstrPrivs, "单据打印") Then
        mnuFileBillPrint.Visible = False
        mnuFileBillPreview.Visible = False
    End If
    If Not zlStr.IsHavePrivs(mstrPrivs, "药品盘点表") Then
        mnuEditAddTable.Visible = False
        tlbTool.Buttons("Table").Visible = False
        TabShow.TabVisible(1) = False
    End If
End Sub

Private Sub Cmd查阅_Click()
    Call mnuEditDisplay_Click
End Sub

Private Sub Form_Activate()
    If vsfList.Visible = True Then
        vsfList.SetFocus
        vsfList.Row = 1
        vsfDetail.Row = 1
    End If
End Sub

Private Sub Form_Load()
    '恢复设置
    Dim dateCurrentDate As Date
    
    Me.Caption = mstrTitle
    mblnViewCost = zlStr.IsHavePrivs(mstrPrivs, "查看成本价")
    dateCurrentDate = Sys.Currentdate
    lblRange.Caption = "查询范围:" & Format(dateCurrentDate, "yyyy年MM月dd日") & "至" & Format(dateCurrentDate, "yyyy年MM月dd日")
    
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs)
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
        Me.Top = (Screen.Height - Me.Height) / 2
    End If
   
    cbrTool.Bands(1).MinHeight = tlbTool.Height
    
    With cbrTool
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth - .Left
    End With
    
    With picSeparate_s
        .Height = 370
        .Left = 0
        .Width = cbrTool.Width
        
    End With
    
    With TabShow
        .Top = IIf(cbrTool.Visible, cbrTool.Height, 0)
        .Left = 0
        .Width = cbrTool.Width
    End With
    
    With vsfList
        .Top = TabShow.Top + TabShow.Height
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
    picColor3.Visible = False
    lblColor3.Visible = False
    picColor.Width = lblColor2.Left + lblColor2.Width + 20
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, mstrTitle
    Call SaveFlexState(vsfList, TabShow.TabCaption(TabShow.Tab))
    Call SaveFlexState(vsfDetail, TabShow.TabCaption(TabShow.Tab))
End Sub
Private Sub mnuEditaddBill_Click()
    Dim strNo As String
    Dim BlnSuccess As Boolean
    
    frmCheckCourseCard.ShowCard Me, strNo, 1, , BlnSuccess
    
    If BlnSuccess Then Call mnuViewRefresh_Click
End Sub

Private Sub mnuEditAddTableAuto_Click()
    Dim strNo As String
    Dim BlnSuccess As Boolean
    
    '检查本期是否已经审核结存，如果未审核结存则不能进行新增业务操作
'    If CheckIsAccount(Val(cboStock.ItemData(cboStock.ListIndex))) = False Then
'        Exit Sub
'    End If
    
    frmCheckCard.ShowCard Me, strNo, 1, , BlnSuccess
    
    If BlnSuccess Then
        Call mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditAddTableTotal_Click()
    Dim strNo As String
    Dim BlnSuccess As Boolean
    
    '检查本期是否已经审核结存，如果未审核结存则不能进行新增业务操作
'    If CheckIsAccount(Val(cboStock.ItemData(cboStock.ListIndex))) = False Then
'        Exit Sub
'    End If
    
    frmCheckCard.ShowCard Me, strNo, 5, , BlnSuccess
    
    If BlnSuccess Then
        Call mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditAddTableZero_Click()
    Dim strNo As String
    Dim BlnSuccess As Boolean
    
    '检查本期是否已经审核结存，如果未审核结存则不能进行新增业务操作
'    If CheckIsAccount(Val(cboStock.ItemData(cboStock.ListIndex))) = False Then
'        Exit Sub
'    End If
   
    frmCheckCard.ShowCard Me, strNo, 6, , BlnSuccess
    
    If BlnSuccess Then
        Call mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditAffirmant_Click()
    Dim str审核日期 As String       '缺省做为确认记录的结束日期
    '填写月度确认记录
    If TabShow.Tab = 1 Then
        str审核日期 = vsfList.TextMatrix(vsfList.Row, 5)
    End If
    With frm月度确认
        Call .ShowEditor(Me.cboStock.ItemData(Me.cboStock.ListIndex), str审核日期)
    End With
End Sub

Private Sub mnuEditVerify_Click()
    '验收
    
    Dim strNo As String
    Dim BlnSuccess As Boolean
    
    With vsfList
        strNo = .TextMatrix(.Row, 0)
        frmCheckCard.ShowCard Me, strNo, 3, .TextMatrix(.Row, .Cols - 2), BlnSuccess
    End With
    
    If BlnSuccess Then Call mnuViewRefresh_Click
End Sub

Private Sub mnuEditDel_Click()
    '删除
    Dim strBillNo As String
    Dim intRow As Integer
    Dim strTitle As String
    Dim intReturn As Integer
    Dim intRecord As Integer
     
    With vsfList
        strTitle = IIf(TabShow.Tab = 0, "盘点记录单", "盘点表")
        
        On Error GoTo errHandle
        intRow = .Row
        strBillNo = .TextMatrix(intRow, 0)
        intReturn = MsgBox("你确实要删除单据号为“" & strBillNo & "”的" & strTitle & "吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
        intRecord = .rows - 1
        If intReturn = vbYes Then
            If TabShow.Tab = 1 Then
                gstrSQL = "zl_药品盘点_Delete('" & strBillNo & "')"
            Else
                gstrSQL = "zl_药品盘点记录单_Delete('" & strBillNo & "')"
            End If
            Call zlDatabase.ExecuteProcedure(gstrSQL, mstrTitle)
            
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
        If TabShow.Tab = 0 Then
            frmCheckCourseCard.ShowCard Me, strNo, 4
        Else
            frmCheckCard.ShowCard Me, strNo, 4, .TextMatrix(.Row, .Cols - 2)
        End If
    End With
End Sub

Private Sub mnuEditStrike_Click()
    Dim blnPurchase As Boolean, blnRefresh As Boolean
    
    '如果是外购(blnPurchase为真)，则直接进入冲销
    '询问是否冲销(blnPurchase为提示框返回值)，是则进入冲销
    blnPurchase = (InStr(1, "1300,1302,1304,1305,1306", mlngMode) <> 0)
    With vsfList
        If Not blnPurchase Then
            blnPurchase = (MsgBox("你确实要全部冲销单据号为“" & .TextMatrix(.Row, 0) & "”的单据吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes)
        End If
        If blnPurchase Then
            blnRefresh = StrikeSave
            If blnRefresh Then mnuViewRefresh_Click
        End If
    End With
End Sub

Private Function StrikeSave() As Boolean
    Dim BlnSuccess As Boolean
    Dim rsTemp As ADODB.Recordset
    Dim int库存检查 As Integer
    Dim strMsg As String
    Dim n As Integer
    
    StrikeSave = False
    
    int库存检查 = MediWork_GetCheckStockRule(mlng库房ID)
    
    On Error GoTo errHandle
    If int库存检查 <> 0 Then
        gstrSQL = "Select A.药品信息 " & _
            " From (Select Distinct '(' || I.编码 || ')' || Nvl(N.名称, I.名称) As 药品信息, A.实际数量, Nvl(K.实际数量, 0) As 库存数量 " & _
            " From 药品收发记录 A, (Select 药品id, 库房id, 实际数量, Nvl(批次, 0) 批次 From 药品库存 Where 性质 = 1) K, 药品规格 B, 收费项目目录 I, 收费项目别名 N " & _
            " Where A.药品id = K.药品id(+) And A.库房id = K.库房id(+) And Nvl(A.批次, 0) = K.批次(+) And A.药品id = B.药品id And " & _
            " A.药品id = I.ID And A.药品id = N.收费细目id(+) And N.性质(+) = 3 And A.单据 = 12 And A.入出系数 = 1 And A.NO = [1]) A " & _
            " Where A.实际数量 > A.库存数量 "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查库存", vsfList.TextMatrix(vsfList.Row, 0))
        
        With rsTemp
            If .RecordCount > 0 Then
                For n = 1 To .RecordCount
                    If n > 5 Then
                        strMsg = strMsg & vbCrLf & "还有其他" & .RecordCount - 5 & "个药品......"
                        Exit For
                    End If
                    strMsg = IIf(strMsg = "", "", strMsg & "," & vbCrLf) & !药品信息
                    .MoveNext
                Next
                
                If int库存检查 = 1 Then
                    If MsgBox("注意，以下药品库存不足：" & vbCrLf & strMsg & vbCrLf & Space(4) & "是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                ElseIf int库存检查 = 2 Then
                    MsgBox "对不起，以下药品库存不足，不能冲销！" & vbCrLf & strMsg, vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End With
    End If
    
    With vsfList
        gstrSQL = "zl_药品盘点_Strike('" & .TextMatrix(.Row, 0) & "','" & UserInfo.用户姓名 & "')"
        
        Call zlDatabase.ExecuteProcedure(gstrSQL, mstrTitle)
        
        '提示停用药品
        Call CheckStopMedi(单据号.盘点表 & "|" & .TextMatrix(.Row, 0))
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
    Dim BlnSuccess As Boolean
    
    BlnSuccess = False
    With vsfList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        strNo = .TextMatrix(.Row, 0)
        If TabShow.Tab = 0 Then
            frmCheckCourseCard.ShowCard Me, strNo, 2, 1, BlnSuccess
        Else
            frmCheckCard.ShowCard Me, strNo, 2, vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 2), BlnSuccess
        End If
        
        If BlnSuccess Then Call mnuViewRefresh_Click
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
        ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1307", "zl8_bill_1307"), Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & .TextMatrix(.Row, .Cols - 2), "单位系数=" & int单位系数, 1
    End With
End Sub
Private Sub mnuFileBillPrint_Click()
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
        ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1307", "zl8_bill_1307"), Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & .TextMatrix(.Row, .Cols - 2), "单位系数=" & int单位系数, 2
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
    Dim dateCurrentDate As Date
    Dim int查询天数 As Integer
    
    frm参数设置.设置参数 Me, mstrPrivs, mstrTitle
    
    dateCurrentDate = Sys.Currentdate
    int查询天数 = Val(zlDatabase.GetPara("查询天数", glngSys, mlngMode, 7)) - 1
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
    Dim StrWinName As String
    With vsfList
        StrWinName = "frmMainList8"
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

Private Sub mnuReportItem_Click(Index As Integer)
    '默认参数：药品=药品id，库房=库房id，开始时间=填制开始时间，结束时间=填制结束时间，盘点单=盘点单NO，盘点表=盘点表NO
    Dim str开始时间 As String
    Dim str结束时间 As String
    Dim strNo As String
    Dim strReportName As String
    
    strReportName = Split(mnuReportItem(Index).Tag, ",")(1)
    
    Select Case strReportName
        Case "ZL1_INSIDE_1307"
            Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1307", Me, "库房=" & Me.cboStock.Text & "|" & IIf(Me.cboStock.ItemData(Me.cboStock.ListIndex) = 0, " is not null ", "=" & Me.cboStock.ItemData(Me.cboStock.ListIndex)))
        Case "ZL1_INSIDE_1307_1"
            Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1307_1", Me, "库房=" & Me.cboStock.Text & "|" & IIf(Me.cboStock.ItemData(Me.cboStock.ListIndex) = 0, " is not null ", "=" & Me.cboStock.ItemData(Me.cboStock.ListIndex)), "单位=" & Choose(mintUnit, "售价单位", "门诊单位", "住院单位", "药库单位") & "|" & Choose(mintUnit, 1, 3, 4, 2))
        Case Else
            If vsfList.Row >= 1 And LTrim(vsfList.TextMatrix(vsfList.Row, 0)) <> "" Then
                strNo = vsfList.TextMatrix(vsfList.Row, 0)
            End If
            
            str开始时间 = IIf(Format(SQLCondition.date填制时间开始, "yyyy-mm-dd") = "1899-12-30", "", Format(SQLCondition.date填制时间开始, "yyyy-mm-dd"))
            str结束时间 = IIf(Format(SQLCondition.date填制时间结束, "yyyy-mm-dd") = "1899-12-30", "", Format(SQLCondition.date填制时间结束, "yyyy-mm-dd"))
            
            Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
                "药品=" & IIf(SQLCondition.lng药品 = 0, "", SQLCondition.lng药品), _
                "库房=" & IIf(Val(cboStock.ItemData(cboStock.ListIndex)) = 0, "", Val(cboStock.ItemData(cboStock.ListIndex))), _
                "开始时间=" & str开始时间, _
                "结束时间=" & str结束时间, _
                "盘点单=" & strNo, _
                "盘点表=" & strNo)
    End Select
End Sub

Private Sub mnuViewRefresh_Click()
    '刷新
    GetList mstrFind
End Sub

Private Sub mnuViewSearch_Click()
    '查找
    
    Dim strFind As String
    
    strFind = FrmTransferSearch.GetSearch(Me, mlngMode, mlng库房ID, strStart, strEnd, strVerifyStart, strVerifyEnd, _
                SQLCondition.strNO开始, _
                SQLCondition.strNO结束, _
                SQLCondition.date填制时间开始, _
                SQLCondition.date填制时间结束, _
                SQLCondition.date审核时间开始, _
                SQLCondition.date审核时间结束, _
                SQLCondition.lng药品, _
                SQLCondition.lng移入库房, _
                SQLCondition.str填制人, _
                SQLCondition.str审核人, _
                SQLCondition.lng药品分类, _
                SQLCondition.str剂型)
    
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




Private Sub vsfDetail_EnterCell()
    With vsfDetail
        If .Row = 0 Then Exit Sub
        
        .Redraw = flexRDNone
        .ForeColorSel = .Cell(flexcpForeColor, .Row, 1)
        .Redraw = flexRDDirect
    End With
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
    Dim intBill As Integer                      '单据类型  如：1、外购入库；2、
    Dim str包装系数 As String
    Dim str单位字段 As String
    Dim strOrder As String
    Dim strCompare As String
    Dim strSql效期 As String
    Dim lngColor As Long
    Dim n As Long
    Dim i As Integer
    Dim intCol As Integer
    Dim strSql药名 As String
    Dim strSqlOrder As String
    
    If mlastRow = vsfList.Row Then Exit Sub
    mlastRow = vsfList.Row
    
    On Error GoTo errHandle
    With vsfList
        .Redraw = flexRDNone
        .ForeColorSel = .Cell(flexcpForeColor, mlastRow, 1)
        .Redraw = flexRDDirect
    End With
    
    strOrder = zlDatabase.GetPara("排序", glngSys, 模块号.药品盘点)
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
    End If
    
    strSqlOrder = strSqlOrder & IIf(Right(strOrder, 1) = "0", " ASC", " DESC") & ",药品信息,序号"
    
    If vsfList.Row >= 1 And LTrim(vsfList.TextMatrix(vsfList.Row, 0)) <> "" Then
        vsfList.Col = 0
        vsfList.ColSel = vsfList.Cols - 1
        
        vsfDetail.Redraw = flexRDNone
        Select Case mintUnit
            Case mconint售价单位
                str包装系数 = "1"
                str单位字段 = "I.计算单位"
            Case mconint门诊单位
                str包装系数 = "B.门诊包装"
                str单位字段 = "B.门诊单位"
            Case mconint住院单位
                str包装系数 = "B.住院包装"
                str单位字段 = "B.住院单位"
            Case mconint药库单位
                str包装系数 = "B.药库包装"
                str单位字段 = "B.药库单位"
        End Select
        
        strSql效期 = IIf(gtype_UserSysParms.P149_效期显示方式 = 1, "TO_CHAR(A.效期-1,'YYYY-MM-DD') AS 有效期至", "TO_CHAR(A.效期,'YYYY-MM-DD') AS 失效期")
        
        If gint药品名称显示 = 0 Then
            strSql药名 = ",('['||I.编码||']'||I.名称) AS 药品信息"
        ElseIf gint药品名称显示 = 1 Then
            strSql药名 = ",('['||I.编码||']'||NVL(N.名称,I.名称)) AS 药品信息"
        Else
            strSql药名 = ",('['||I.编码||']'||I.名称) AS 药品信息,N.名称 As 商品名"
        End If
        
        intBill = IIf(TabShow.Tab = 1, 12, 14)
        If TabShow.Tab = 1 Then
            gstrSQL = "Select DISTINCT a.序号" & strSql药名 & "," _
                    & "     B.药品来源,B.基本药物,I.规格,a.产地," & str单位字段 & " as 单位,a.批号," & strSql效期 & ",a.批准文号," _
                    & "     LTRIM(to_char(A.填写数量 /" & str包装系数 & ",decode(a.扣率,0,'999999999990.00000'," & mstrNumberFormat & "))) AS 帐面数," _
                    & "     LTRIM(to_char(A.扣率 /" & str包装系数 & "," & mstrNumberFormat & ")) AS 实盘数," _
                    & "     Decode(Sign(A.扣率-A.填写数量),-1,'亏',1,'盈','平') as 标志," _
                    & "     LTRIM(to_char(A.实际数量 /" & str包装系数 & ",decode(a.扣率,0,'999999999990.00000'," & mstrNumberFormat & "))) AS 数量差," _
                    & "     LTRIM(TO_CHAR (a.单量*" & str包装系数 & ", " & mstrCostFormat & ")) AS 成本价," _
                    & "     LTRIM(TO_CHAR (a.零售价*" & str包装系数 & ", " & mstrPriceFormat & ")) AS 售价," _
                    & "     LTRIM(TO_CHAR (a.零售金额*a.入出系数,decode(a.扣率,0,'999999999990.00000', " & mstrMoneyFormat & "))) AS 金额差," _
                    & "     LTRIM(TO_CHAR ((A.扣率-A.填写数量) * a.零售价* Decode(记录状态, 1, 1, Decode(Mod(记录状态, 3), 0, 1, -1))," & mstrMoneyFormat & ")) AS 账面金额差," _
                    & "     LTRIM(TO_CHAR (a.差价*a.入出系数, decode(a.扣率,0,'999999999990.00000'," & mstrMoneyFormat & "))) AS 差价差, " _
                    & "     LTrim(To_Char((a.扣率 / b.门诊包装)*(a.零售价 * b.门诊包装), " & mstrMoneyFormat & ")) As 盘点金额," _
                    & "     LTrim(To_Char(((a.成本价+to_char(a.零售金额*a.入出系数*Decode(a.记录状态, 1, 1, Decode(Mod(a.记录状态, 3), 0, 1, -1))," & mstrMoneyFormat & "))-(a.成本金额+to_char(a.差价*a.入出系数*Decode(a.记录状态, 1, 1, Decode(Mod(a.记录状态, 3), 0, 1, -1))," & mstrMoneyFormat & ")))," & mstrMoneyFormat & ")) as 盘点成本金额, " _
                    & "     ltrim(To_Char((a.零售金额*a.入出系数 - a.差价*a.入出系数 ), " & mstrMoneyFormat & ")) As 成本金额差," _
                    & " Nvl(I.撤档时间, To_Date('3000-01-01', 'YYYY-MM-DD')) As 撤档时间,e.库房货位 " _
                    & " From (Select a.入出系数,a.记录状态,a.序号,a.药品id,a.产地,a.批号,a.效期,A.填写数量,A.扣率,A.实际数量,a.成本价,a.成本金额,a.零售价,a.零售金额,a.差价,a.单量,a.批准文号,a.库房id" _
                    & "         From 药品收发记录 a" _
                    & "        Where a.记录状态= [2] " _
                    & "             And a.单据= 12 And a.No=[1]) a," _
                    & "        药品规格 b,收费项目目录 I ,收费项目别名 n,药品储备限额 e" _
                    & " Where a.药品id = b.药品id And a.药品id = i.Id" _
                    & "        And a.药品id=n.收费细目id(+) And n.性质(+)=3 " _
                    & "        And a.药品id = e.药品id and a.库房id = e.库房id " _
                    & " ORDER BY " & strSqlOrder
        Else
            gstrSQL = "Select DISTINCT a.序号" & strSql药名 & "," _
                    & "     B.药品来源,B.基本药物,I.规格,a.产地," & str单位字段 & " as 单位,a.批号," & strSql效期 & ",a.批准文号," _
                    & "     to_char(A.扣率 /" & str包装系数 & "," & mstrNumberFormat & ") AS 实盘数" _
                    & " From (Select a.序号,a.药品id,a.产地,a.批号,a.效期,A.填写数量,A.扣率,A.实际数量,a.零售价,a.零售金额,a.差价,a.批准文号" _
                    & "         From 药品收发记录 a" _
                    & "        Where a.记录状态= 1 And a.单据= 14 And a.No=[1]) a," _
                    & "        药品规格 b,收费项目目录 I ,收费项目别名 n" _
                    & " Where a.药品id = b.药品id And a.药品id = i.Id" _
                    & "        And a.药品id=n.收费细目id(+) And n.性质(+)=3 " _
                    & " ORDER BY " & strSqlOrder
        End If
        Set rsDetail = zlDatabase.OpenSQLRecord(gstrSQL, mstrTitle, vsfList.TextMatrix(vsfList.Row, 0), vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 2))
        
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
                
        '格式化数量
'        With vsfDetail
'            Select Case TabShow.Tab
'            Case 0
'                For n = 1 To .rows - 1
'                    .TextMatrix(n, .ColIndex("实盘数")) = zlStr.FormatEx(.TextMatrix(n, .ColIndex("实盘数")), mintShowNumberDigit)
'                Next
'            Case 1
'                For n = 1 To .rows - 1
'                    .TextMatrix(n, .ColIndex("帐面数")) = zlStr.FormatEx(.TextMatrix(n, .ColIndex("帐面数")), mintShowNumberDigit)
'                    .TextMatrix(n, .ColIndex("实盘数")) = zlStr.FormatEx(.TextMatrix(n, .ColIndex("实盘数")), mintShowNumberDigit)
'                    .TextMatrix(n, .ColIndex("数量差")) = zlStr.FormatEx(.TextMatrix(n, .ColIndex("数量差")), mintShowNumberDigit)
'                Next
'            End Select
'        End With
        
        vsfDetail.Redraw = flexRDDirect
    ElseIf LTrim(vsfList.TextMatrix(vsfList.Row, 0)) = "" Then
        With vsfDetail
            .Cols = IIf(TabShow.Tab = 1, 24, 11)
            If gint药品名称显示 = 2 Then .Cols = .Cols + 1
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
            .TextMatrix(0, intCol) = "产地": intCol = intCol + 1
            .TextMatrix(0, intCol) = "单位": intCol = intCol + 1
            .TextMatrix(0, intCol) = "批号": intCol = intCol + 1
            .TextMatrix(0, intCol) = IIf(gtype_UserSysParms.P149_效期显示方式 = 1, "有效期至", "失效期"): intCol = intCol + 1
            .TextMatrix(0, intCol) = "批准文号": intCol = intCol + 1
            If TabShow.Tab = 0 Then
                .TextMatrix(0, intCol) = "实盘数": intCol = intCol + 1
            Else
                .TextMatrix(0, intCol) = "帐面数": intCol = intCol + 1
                .TextMatrix(0, intCol) = "实盘数": intCol = intCol + 1
                .TextMatrix(0, intCol) = "标志": intCol = intCol + 1
                .TextMatrix(0, intCol) = "数量差": intCol = intCol + 1
                .TextMatrix(0, intCol) = "成本价": intCol = intCol + 1
                .TextMatrix(0, intCol) = "售价": intCol = intCol + 1
                .TextMatrix(0, intCol) = "金额差": intCol = intCol + 1
                .TextMatrix(0, intCol) = "差价差": intCol = intCol + 1
                .TextMatrix(0, intCol) = "盘点成本金额": intCol = intCol + 1
                .TextMatrix(0, intCol) = "盘点金额": intCol = intCol + 1
                .TextMatrix(0, intCol) = "成本金额差": intCol = intCol + 1
                .TextMatrix(0, intCol) = "账面金额差": intCol = intCol + 1
                .TextMatrix(0, intCol) = "撤档时间": intCol = intCol + 1
                .TextMatrix(0, intCol) = "库房货位": intCol = intCol + 1
            End If
            
            .Row = 1
            .Col = 0
            .ColSel = .Cols - 1
            
            For n = 0 To .Cols - 1
                .ColKey(n) = .TextMatrix(0, n)
                .FixedAlignment(n) = flexAlignCenterCenter
            Next
        End With
    End If
    SetDetailColWidth
    SetEnable
    
    '上色
    If TabShow.Tab = 1 Then
        With vsfDetail
            .Redraw = flexRDNone
            For n = 1 To .rows - 1
                If .TextMatrix(n, 0) <> "" Then
                    If .TextMatrix(n, .ColIndex("标志")) = "盈" Then
                        lngColor = vbRed
                    ElseIf .TextMatrix(n, .ColIndex("标志")) = "亏" Then
                        lngColor = vbBlue
                    Else
                        lngColor = vbBlack
                    End If
                    
                    '盘亏盘盈行用颜色区分；
                    If lngColor <> vbBlack Then
                        .Cell(flexcpForeColor, n, 0, n, .Cols - 1) = lngColor
                    End If
                    
                    '如果是停用药品，该行粗体显示
                    If Format(.TextMatrix(n, .ColIndex("撤档时间")), "YYYY-MM-DD") <> "3000-01-01" Then
                        .Cell(flexcpFontBold, n, 0, n, .Cols - 1) = True
                    End If
                End If
            Next
            .Redraw = flexRDDirect
        End With
    End If
    
    vsfDetail.Row = 1
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

Private Sub TabShow_Click(PreviousTab As Integer)
    Call SaveFlexState(vsfList, TabShow.TabCaption(PreviousTab))
    Call SaveFlexState(vsfDetail, TabShow.TabCaption(PreviousTab))
    mblnBootUp = False
    If TabShow.Tab = 1 Then
        vsfDetail.ToolTipText = mcstComment
    Else
        vsfDetail.ToolTipText = ""
    End If
    GetList (mstrFind)  '列出单据头
    mblnBootUp = True
End Sub
Private Sub tlbTool_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "PrintView"
            mnuFilePreView_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Bill"
            mnuEditaddBill_Click
        Case "Table"
            mnuEditAddTableAuto_Click
        Case "Modify"
            mnuEditModify_Click
        Case "Delete"
            mnuEditDel_Click
        Case "Verify"
            mnuEditVerify_Click
        Case "Strike"
            mnuEditStrike_Click
        Case "Affirmant"
            mnuEditAffirmant_Click
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
    Dim strVerify As String, blnVisible As Boolean
    
    blnVisible = (TabShow.Tab = 1)
    mnuEditVerify.Visible = blnVisible And zlStr.IsHavePrivs(mstrPrivs, "审核")
    mnuEditStrike.Visible = blnVisible And zlStr.IsHavePrivs(mstrPrivs, "冲销")
    mnuEditLine2.Visible = blnVisible
    tlbTool.Buttons("Verify").Visible = blnVisible And zlStr.IsHavePrivs(mstrPrivs, "审核")
    tlbTool.Buttons("Strike").Visible = blnVisible And zlStr.IsHavePrivs(mstrPrivs, "冲销")
    tlbTool.Buttons("VerifySeparate").Visible = blnVisible
    
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
            mnuFileExcel.Enabled = True
            tlbTool.Buttons("Print").Enabled = True
            tlbTool.Buttons("PrintView").Enabled = True
            
            mnuFileBillPreview.Enabled = TabShow.Tab = 1
            mnuFileBillPrint.Enabled = TabShow.Tab = 1
            
            If TabShow.Tab = 1 Then
                strVerify = .TextMatrix(.Row, .Cols - 8)
            Else
                strVerify = ""
            End If
            If strVerify = "" Then    '未审核单
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
                If .TextMatrix(.Row, .Cols - 2) Mod 3 = 0 Then
                    .ToolTipText = "冲销单据的原单据"
                    If mnuEditStrike.Visible = True Then
                        mnuEditStrike.Enabled = True
                        tlbTool.Buttons("Strike").Enabled = True
                    End If
                ElseIf .TextMatrix(.Row, .Cols - 2) Mod 3 = 2 Then
                    .ToolTipText = "冲销单据"
                    If mnuEditStrike.Visible = True Then
                        mnuEditStrike.Enabled = False
                        tlbTool.Buttons("Strike").Enabled = False
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
    
    Set objRow = New zlTabAppRow
    objRow.Add "盘点库房：" & Trim(cboStock.Text)
    objRow.Add "盘点时间：" & Trim(vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "盘点时间")))
    objPrint.UnderAppRows.Add objRow
        
    Set objRow = New zlTabAppRow
    objRow.Add "摘要:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "摘要"))
    objPrint.BelowAppRows.Add objRow
        
    Set objRow = New zlTabAppRow
    objRow.Add "填制人:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "填制人")) & "  填制日期:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "填制日期"))
    
    If TabShow.Tab = 1 Then
        objRow.Add "审核人:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "审核人")) & "  审核日期:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "审核日期"))
        objPrint.BelowAppRows.Add objRow
    End If
    
    Set objPrint.Body = vsfDetail
    zlPrintOrView1Grd objPrint, bytMode
End Sub

Private Sub tlbTool_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
    Case "Auto"
        Call mnuEditAddTableAuto_Click
    Case "Total"
        Call mnuEditAddTableTotal_Click
    Case "Zero"
        Call mnuEditAddTableZero_Click
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


