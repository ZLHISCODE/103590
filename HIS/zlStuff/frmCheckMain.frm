VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
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
   Begin TabDlg.SSTab TabShow 
      Height          =   360
      Left            =   30
      TabIndex        =   6
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
      TabPicture(0)   =   "frmCheckMain.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "盘点表清单(&2)"
      TabPicture(1)   =   "frmCheckMain.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
   End
   Begin MSComctlLib.ImageList ilsHot 
      Left            =   6270
      Top             =   120
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
            Picture         =   "frmCheckMain.frx":0342
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":0562
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":0782
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":099E
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":0BBE
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":0DDE
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":0FFA
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":1216
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":1430
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":158A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":17AA
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsCold 
      Left            =   5670
      Top             =   120
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
            Picture         =   "frmCheckMain.frx":19CA
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":1BEA
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":1E0A
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":2026
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":2246
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":2466
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":2682
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":289E
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":2AB8
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":2C12
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckMain.frx":2E2E
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Cmd查阅 
      Caption         =   "查阅(&V)"
      Height          =   350
      Left            =   5160
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2550
      Width           =   1100
   End
   Begin VB.PictureBox picSeparate_s 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   30
      MousePointer    =   7  'Size N S
      ScaleHeight     =   300
      ScaleWidth      =   4815
      TabIndex        =   4
      Top             =   2580
      Width           =   4815
      Begin VB.Label lbl成本金额差 
         AutoSize        =   -1  'True
         Caption         =   "成本金额差合计："
         Height          =   170
         Left            =   3240
         TabIndex        =   8
         Top             =   40
         Width           =   1440
      End
      Begin VB.Label lblSum成本金额 
         AutoSize        =   -1  'True
         Caption         =   "盘点成本金额合计："
         Height          =   170
         Left            =   480
         TabIndex        =   7
         Top             =   40
         Width           =   1620
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
         ButtonWidth     =   1138
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
         Begin VB.Timer LimitTime 
            Enabled         =   0   'False
            Interval        =   8000
            Left            =   6660
            Top             =   180
         End
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
   Begin VSFlex8Ctl.VSFlexGrid mshlist 
      Height          =   1005
      Left            =   120
      TabIndex        =   9
      Top             =   1320
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
      FormatString    =   $"frmCheckMain.frx":304E
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
   Begin VSFlex8Ctl.VSFlexGrid mshDetail 
      Height          =   975
      Left            =   360
      TabIndex        =   10
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
      FormatString    =   $"frmCheckMain.frx":30C3
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
Private mintPreCol As Integer           '前一次单据头的排序列
Private mintsort As Integer             '前一次单据头的排序
Private mintPreDetailCol As Integer     '前一次单据体的排序列
Private mintDetailsort As Integer       '前一次单据体的排序
Public mstrPrivs As String                     '权限
Private mintUnit  As Integer                '显示单位:0-散装单位,1-包装单位
Private mintUnit1  As Integer                '显示单位:0-散装单位,1-包装单位
Private mstrOrder As String             '记录排序方式

Private mblnLoadGrid As Boolean

Private mintOldY  As Integer
Private mstrOthers() As String ' 0-记录状态(计划类型),1-开始单号,2-结束单号,3-材料id,4-对方部门id(或入出类别id或编制方法(计划用)),5-填制人,6-审核人,7-供应商ID,8-生产商,9-开始生产日期,10-结束生产日期,11-开始发票号,12-结束发票号
Private mblnCostView As Boolean             '查看成本价相关信息 true-允许查看 false-不允许查看
Private Const mstrCaption As String = "卫材盘点管理"
'盘点单
Private Const M_INT_COL盘点单NO As Integer = 0
Private Const M_INT_COL盘点单盘点时间 As Integer = 1
Private Const M_INT_COL盘点单填制人 As Integer = 2
Private Const M_INT_COL盘点单填制日期 As Integer = 3
Private Const M_INT_COL盘点单摘要 As Integer = 4
Private Const M_INT_COL盘点单外观 As Integer = 5
Private Const M_INT_COL盘点单已盘 As Integer = 6
Private Const M_INT_盘点单ALLCOLUMN As Integer = 7 '总列数
'盘点表
Private Const M_INT_COLNO As Integer = 0 ' "NO"
Private Const M_INT_COL盘点时间 As Integer = 1 ' "盘点时间"
Private Const M_INT_COL填制人 As Integer = 2 ' "填制人"
Private Const M_INT_COL填制日期 As Integer = 3 ' "填制日期"
Private Const M_INT_COL审核人 As Integer = 4 '"审核人"
Private Const M_INT_COL审核日期 As Integer = 5 '"审核日期"
Private Const M_INT_COL盘点金额 As Integer = 6 '"盘点金额"
Private Const M_INT_COL金额差 As Integer = 7 '"金额差"
Private Const M_INT_COL盘点成本金额 As Integer = 8 ' "盘点成本金额"
Private Const M_INT_COL盘点成本金额差 As Integer = 9 ' "盘点成本金额差"
Private Const M_INT_COL记录状态 As Integer = 10 '"记录状态"
Private Const M_INT_COL摘要 As Integer = 11 '"摘要"
Private Const M_INT_ALLCOLUMN As Integer = 12 '总列数
 
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
Private mORaFMT记录单 As g_FmtString
'----------------------------------------------------------------------------------------------------------


'日期设置
Private mdtStartDate As Date
Private mdtEndDate As Date
Private mdtVerifyStart As Date
Private mdtVerifyEnd As Date
Private mintFindDay As Integer  '查询天数

Private Sub cboStock_Click()
    If mblnNoClick Then Exit Sub
    If cboStock.ListIndex >= 0 Then cboStock.Tag = cboStock.ItemData(cboStock.ListIndex)
    If mblnBootUp Then mnuViewRefresh_Click
End Sub

Private Sub cboStock_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cboStock.ListCount = 0 Then Call zlControl.ControlSetFocus(mshlist): Exit Sub
    
    If cboStock.ListIndex >= 0 Then
        If Val(cboStock.Tag) = cboStock.ItemData(cboStock.ListIndex) Then
            Call zlControl.ControlSetFocus(mshlist, True)
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

Public Sub ShowList(ByVal lngMode As Long, ByVal strTitle As String, ByVal FrmMain As Variant)
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:显示指定的单据管理,
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------

    Dim strFind As String
    mblnBootUp = False
    mlngMode = lngMode
    mstrTitle = strTitle
    mstrPrivs = gstrPrivs
    
    If Not CheckDepend Then Exit Sub            '数据依赖性测试
    
    
    Me.Caption = strTitle
    
    SetVisable  '根据权限设置不同的显示项目
    
    Call initGrid
    mintFindDay = Val(zlDataBase.GetPara("查询天数", glngSys, mlngMode, 1))
    mdtStartDate = Format(DateAdd("d", -mintFindDay, sys.Currentdate), "yyyy-MM-dd")
    mdtEndDate = Format(sys.Currentdate, "yyyy-MM-dd")
    mdtVerifyStart = "1901-01-01"
    mdtVerifyEnd = "1901-01-01"
    
    strFind = " AND A.记录状态 = 1 And A.审核日期 is Null And A.填制日期 Between To_Date('" & Format(mdtStartDate, "yyyy-mm-dd") & " 00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(mdtEndDate, "yyyy-mm-dd") & " 23:59:59','YYYY-MM-DD HH24:MI:SS')"
    mstrFind = strFind
    
    GetList (mstrFind)  '列出单据头
    
    TabShow.Tab = 1
    Call SetListColWidth
    TabShow.Tab = 0
    RestoreWinState Me, App.ProductName, mstrTitle
    
    If TabShow.Tab = 1 Then
        With mshDetail
            .ColWidth(12) = IIf(mblnCostView = True, 1000, 0)
            .ColWidth(15) = IIf(mblnCostView = True, 1000, 0)
            .ColWidth(.Cols - 2) = IIf(mblnCostView = True, 1500, 0)
            .ColWidth(.Cols - 1) = IIf(mblnCostView = True, 1500, 0)
        End With
        With mshlist
            .ColWidth(M_INT_COL盘点成本金额) = IIf(mblnCostView = True, 1000, 0)
            .ColWidth(M_INT_COL盘点成本金额差) = IIf(mblnCostView = True, 1000, 0)
        End With
    End If
    mblnBootUp = True
    
    If IsObject(FrmMain) Then
        Me.Show , FrmMain
    Else
        OS.ShowChildWindow Me.hwnd, FrmMain
    End If
    
    Me.ZOrder 0
End Sub

'检查数据依赖性
Private Function CheckDepend() As Boolean
    Dim rsTemp As New Recordset
    
    On Error GoTo errHandle
    CheckDepend = False
    mstr工作性质 = "V,W,K,12"
    gstrSQL = "" & _
            "   SELECT DISTINCT a.id, a.名称 " & _
            "   FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " & _
            "   Where c.工作性质 = b.名称 and (a.站点=[2] or a.站点 is null) " & _
            "       And b.编码 In('V','K','W','12') " & _
            "       AND a.id = c.部门id " & _
            "       AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01'" & _
            IIf(InStr(gstrPrivs, "所有库房") <> 0, "", " And a.ID IN (Select 部门ID From 部门人员 Where 人员ID=[1])")
                        
    mbln操作员限制 = Not zlStr.IsHavePrivs(gstrPrivs, "所有库房")
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, mstrTitle, UserInfo.Id, gstrNodeNo)
    
    If rsTemp.EOF Then
        ShowMsgBox "至少应该设置一个具有库房性质、发料部门" & vbCrLf & "或者制剂室性质的部门,请查看部门管理！"
        rsTemp.Close
        Exit Function
    End If
    
    With cboStock
        .Clear
        Do While Not rsTemp.EOF
            .AddItem rsTemp!名称
            .ItemData(.NewIndex) = rsTemp!Id
            If rsTemp!Id = UserInfo.部门ID Then
                .ListIndex = .NewIndex
            End If
            rsTemp.MoveNext
        Loop
        rsTemp.Close
        
        If .ListIndex = -1 Then
            If InStr(gstrPrivs, "所有库房") = 0 Then
                ShowMsgBox "你不是发料部门或库房工作人员且不具有所有库房的权限，不能进入！"
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
    Dim rsTemp As New Recordset
    Dim strUserPart As String
    Dim dbl盘点成本金额 As Double
    Dim dbl盘点成本金额差 As Double
    Dim intCol As Integer
    Dim intRow As Integer
    
    mlastRow = 0
    On Error GoTo errHandle
    Call FS.ShowFlash("正在搜索卫生材料记录,请稍候 ...", Me)
    DoEvents
    Screen.MousePointer = vbHourglass
    strUserPart = " And A.库房ID+0=[1]"
    
    mshlist.Redraw = False
    
    
    '频次字段保存的 盘店时间
    
    If TabShow.Tab = 1 Then
        gstrSQL = "" & _
            "   SELECT distinct a.no, 频次 AS 盘点时间," & _
            "           a.填制人,TO_CHAR (min(a.填制日期), 'yyyy-mm-dd HH24:Mi:SS') AS 填制日期, a.审核人," & _
            "           TO_CHAR (min(a.审核日期), 'yyyy-mm-dd HH24:Mi:SS') AS 审核日期, " & _
            "           ltrim(to_char((Sum(Nvl(扣率,0)*零售价))," & mOraFMT.FM_金额 & ")) 盘点金额," & _
            "           ltrim(to_char((Sum(零售金额*decode(sign(Nvl(扣率,0)-填写数量),-1,-1,1)))," & mOraFMT.FM_金额 & ")) 金额差," & _
            "           LTrim(to_Char(sum(a.成本价+to_char(a.零售金额*a.入出系数*Decode(a.记录状态, 1, 1, Decode(Mod(a.记录状态, 3), 0, 1, -1))," & mORaFMT记录单.FM_金额 & ")" & "-(a.成本金额+to_char(a.差价*a.入出系数*Decode(a.记录状态, 1, 1, Decode(Mod(a.记录状态, 3), 0, 1, -1))," & mORaFMT记录单.FM_金额 & ")))," & mORaFMT记录单.FM_金额 & ")) as  盘点成本金额, " & _
            "           LTrim(To_Char(sum(a.零售金额*a.入出系数-a.差价*a.入出系数), " & mOraFMT.FM_金额 & ")) as 盘点成本金额差 , " & _
            "           a.记录状态, a.摘要 " & _
            "   FROM 药品收发记录 a, 部门表 b " & _
            "   Where a.库房id = b.ID AND a.单据 =22  " & strUserPart & strFind & _
            "   Group by a.no,频次,a.填制人,a.审核人,a.记录状态, a.摘要 " & _
            "   ORDER BY no DESC,填制日期 ASC "
    Else
        gstrSQL = "" & _
            "   SELECT distinct a.no, 频次 AS 盘点时间," & _
            "           a.填制人,TO_CHAR (min(a.填制日期), 'yyyy-mm-dd HH24:Mi:SS') AS 填制日期,a.摘要,外观 " & _
            "   FROM 药品收发记录 a, 部门表 b " & _
            "   Where  a.库房id = b.ID  and a.单据 = 23  " & strUserPart & strFind & _
            "   Group by a.no,频次,a.填制人,a.摘要,外观 " & _
            "   ORDER BY no DESC,填制日期 ASC "
    End If
    
    'mstrOthers(0 To 6) As String ' 0-记录状态(计划类型),1-开始单号,2-结束单号,3-材料id,4-对方部门id(或入出类别id或编制方法(计划用)),5-填制人,6-审核人,7-供应商ID,8-生产商,9-开始生产日期,10-结束生产日期,11-开始发票号,12-结束发票号
    '参数范围:[1]-库房id,[2]:开始填制日期,[3]结束填制日期,[4]开始审核日期,[5] 结束审核日期,[6]-记录状态,[7]开始单据号,[8]结束单据号,[9]材料id,[10]对方部门id,[11]填制人,[12]审核人
    ' 未和参数: [13]-供应商ID,[14]-生产商,[15]-开始生产日期,[16]-结束生产日期,[17]-开始发票号,[18]-结束发票号
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, cboStock.ItemData(cboStock.ListIndex), _
        CDate(Format(mdtStartDate, "yyyy-mm-dd") & " 00:00:00"), CDate(Format(mdtEndDate, "yyyy-mm-dd") & " 23:59:59"), _
        CDate(Format(mdtVerifyStart, "yyyy-mm-dd") & " 00:00:00"), CDate(Format(mdtVerifyEnd, "yyyy-mm-dd") & " 23:59:59"), _
        Val(mstrOthers(0)), mstrOthers(1), mstrOthers(2), Val(mstrOthers(3)), _
        Val(mstrOthers(4)), mstrOthers(5), mstrOthers(6))
    
    With mshlist
        .Rows = 1
        If TabShow.Tab = 1 Then
            Do While Not rsTemp.EOF
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, M_INT_COLNO) = IIf(IsNull(rsTemp!NO), "", rsTemp!NO)
                .TextMatrix(.Rows - 1, M_INT_COL盘点时间) = IIf(IsNull(rsTemp!盘点时间), "", rsTemp!盘点时间)
                .TextMatrix(.Rows - 1, M_INT_COL填制人) = IIf(IsNull(rsTemp!填制人), "", rsTemp!填制人)
                .TextMatrix(.Rows - 1, M_INT_COL填制日期) = IIf(IsNull(rsTemp!填制日期), "", rsTemp!填制日期)
                .TextMatrix(.Rows - 1, M_INT_COL审核人) = IIf(IsNull(rsTemp!审核人), "", rsTemp!审核人)
                .TextMatrix(.Rows - 1, M_INT_COL审核日期) = IIf(IsNull(rsTemp!审核日期), "", rsTemp!审核日期)
                .TextMatrix(.Rows - 1, M_INT_COL盘点金额) = IIf(IsNull(rsTemp!盘点金额), "", rsTemp!盘点金额)
                .TextMatrix(.Rows - 1, M_INT_COL金额差) = IIf(IsNull(rsTemp!金额差), "", rsTemp!金额差)
                .TextMatrix(.Rows - 1, M_INT_COL盘点成本金额) = IIf(IsNull(rsTemp!盘点成本金额), "", rsTemp!盘点成本金额)
                .TextMatrix(.Rows - 1, M_INT_COL盘点成本金额差) = IIf(IsNull(rsTemp!盘点成本金额差), "", rsTemp!盘点成本金额差)
                .TextMatrix(.Rows - 1, M_INT_COL记录状态) = IIf(IsNull(rsTemp!记录状态), "", rsTemp!记录状态)
                .TextMatrix(.Rows - 1, M_INT_COL摘要) = IIf(IsNull(rsTemp!摘要), "", rsTemp!摘要)
                
                dbl盘点成本金额 = dbl盘点成本金额 + IIf(IsNull(rsTemp!盘点成本金额), "", rsTemp!盘点成本金额)
                dbl盘点成本金额差 = dbl盘点成本金额差 + IIf(IsNull(rsTemp!盘点成本金额差), "", rsTemp!盘点成本金额差)
                rsTemp.MoveNext
            Loop
            lblSum成本金额.Caption = "盘点成本金额合计：" & GetFormat(dbl盘点成本金额, g_小数位数.obj_包装小数.金额小数) & "元"
            lbl成本金额差.Caption = "成本金额差合计：" & GetFormat(dbl盘点成本金额差, g_小数位数.obj_包装小数.金额小数) & "元"
        Else
            Do While Not rsTemp.EOF
                .Rows = .Rows + 1
                                
                .TextMatrix(.Rows - 1, M_INT_COL盘点单NO) = IIf(IsNull(rsTemp!NO), "", rsTemp!NO)
                .TextMatrix(.Rows - 1, M_INT_COL盘点单盘点时间) = IIf(IsNull(rsTemp!盘点时间), "", rsTemp!盘点时间)
                .TextMatrix(.Rows - 1, M_INT_COL盘点单填制人) = IIf(IsNull(rsTemp!填制人), "", rsTemp!填制人)
                .TextMatrix(.Rows - 1, M_INT_COL盘点单填制日期) = IIf(IsNull(rsTemp!填制日期), "", rsTemp!填制日期)
                .TextMatrix(.Rows - 1, M_INT_COL盘点单摘要) = IIf(IsNull(rsTemp!摘要), "", rsTemp!摘要)
                .TextMatrix(.Rows - 1, M_INT_COL盘点单外观) = IIf(IsNull(rsTemp!外观), "0", rsTemp!外观)
                If .TextMatrix(.Rows - 1, M_INT_COL盘点单外观) <> 0 Then
                    .TextMatrix(.Rows - 1, M_INT_COL盘点单已盘) = "√"
                End If
                
                rsTemp.MoveNext
            Loop
        End If
    End With
'    Set mshList.Recordset = rsTemp
    With mshlist
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
    End With
    SetListColWidth
    
    mshlist_EnterCell    '列出单据体
    
    If TabShow.Tab = 1 Then
        SetStrikeColor
    End If
    
    With mshlist
        .Row = 1
        .Col = 0
'        .ColSel = .Cols - 1
    End With
    
    mshlist.Redraw = True
    Call FS.StopFlash
    
    Screen.MousePointer = vbDefault
    stbThis.Panels(2).Text = "当前共有" & rsTemp.RecordCount & "张单据"
    rsTemp.Close
    If mshlist.Visible = True Then
        mshlist.SetFocus
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub initGrid()
    mblnLoadGrid = False
    '初始化列表
    With mshlist
        .Rows = 2
        If TabShow.Tab = 1 Then
            '盘点表
            .Cols = M_INT_ALLCOLUMN
            .TextMatrix(0, M_INT_COLNO) = "NO"
            .TextMatrix(0, M_INT_COL盘点时间) = "盘点时间"
            .TextMatrix(0, M_INT_COL填制人) = "填制人"
            .TextMatrix(0, M_INT_COL填制日期) = "填制日期"
            .TextMatrix(0, M_INT_COL审核人) = "审核人"
            .TextMatrix(0, M_INT_COL审核日期) = "审核日期"
            .TextMatrix(0, M_INT_COL盘点金额) = "盘点金额"
            .TextMatrix(0, M_INT_COL金额差) = "金额差"
            .TextMatrix(0, M_INT_COL盘点成本金额) = "盘点成本金额"
            .TextMatrix(0, M_INT_COL盘点成本金额差) = "盘点成本金额差"
            .TextMatrix(0, M_INT_COL记录状态) = "记录状态"
            .TextMatrix(0, M_INT_COL摘要) = "摘要"
            
            .ColAlignment(M_INT_COLNO) = flexAlignLeftCenter  'no
            .ColAlignment(M_INT_COL盘点时间) = flexAlignLeftCenter '盘点时间
            .ColAlignment(M_INT_COL填制人) = flexAlignLeftCenter '填制人
            .ColAlignment(M_INT_COL审核日期) = flexAlignLeftCenter '填制日期
            .ColAlignment(M_INT_COL审核人) = flexAlignLeftCenter '审核人
            .ColAlignment(M_INT_COL审核日期) = flexAlignLeftCenter '审核日期
            .ColAlignment(M_INT_COL盘点金额) = flexAlignLeftCenter '盘点金额
            .ColAlignment(M_INT_COL金额差) = flexAlignRightCenter '金额差
            .ColAlignment(M_INT_COL盘点成本金额) = flexAlignRightCenter '盘点成本金额
            .ColAlignment(M_INT_COL盘点成本金额差) = flexAlignRightCenter '盘点成本金额差
            .ColAlignment(M_INT_COL记录状态) = flexAlignRightCenter '记录状态
            .ColAlignment(M_INT_COL摘要) = flexAlignRightCenter '摘要
        Else
            '盘点单
            .Cols = M_INT_盘点单ALLCOLUMN
            
            .TextMatrix(0, M_INT_COL盘点单已盘) = "已盘"
            .TextMatrix(0, M_INT_COL盘点单NO) = "NO"
            .TextMatrix(0, M_INT_COL盘点单盘点时间) = "盘点时间"
            .TextMatrix(0, M_INT_COL盘点单填制人) = "填制人"
            .TextMatrix(0, M_INT_COL盘点单填制日期) = "填制日期"
            .TextMatrix(0, M_INT_COL盘点单摘要) = "摘要"
            .TextMatrix(0, M_INT_COL盘点单外观) = "外观"
            
            .ColAlignment(M_INT_COL盘点单已盘) = flexAlignCenterCenter   '外观
            .ColAlignment(M_INT_COL盘点单NO) = flexAlignLeftCenter  'no
            .ColAlignment(M_INT_COL盘点单盘点时间) = flexAlignLeftCenter '盘点时间
            .ColAlignment(M_INT_COL盘点单填制人) = flexAlignLeftCenter  '填制人
            .ColAlignment(M_INT_COL盘点单填制日期) = flexAlignLeftCenter '填制日期
            .ColAlignment(M_INT_COL盘点单摘要) = flexAlignLeftCenter '摘要
            .ColAlignment(M_INT_COL盘点单外观) = flexAlignLeftCenter  '外观
        End If
    End With
    mblnLoadGrid = True
End Sub

Private Sub SetStrikeColor()
    Dim intStatus As Integer
    Dim intRow As Integer
    Dim intCol As Integer
    
    With mshlist
        If .Rows <= 2 Then Exit Sub
        For intRow = 1 To .Rows - 1
            intStatus = IIf(TabShow.Tab = 0, 1, Val(.TextMatrix(intRow, M_INT_COL记录状态)))
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
                    .CellForeColor = &HFF
                Next
            End If
        Next
    End With
End Sub

'表头列宽初始
Private Sub SetListColWidth()
    Dim intCol As Integer
    
    With mshlist
        If TabShow.Tab = 1 Then
            If mblnBootUp = False Then
                For intCol = 1 To .Cols - 1
                    If intCol = 1 Then
                        .ColWidth(intCol) = 2000
                    ElseIf intCol = M_INT_COL记录状态 Then
                        .ColWidth(intCol) = 0
                    Else
                        If intCol = M_INT_COL盘点成本金额 Or intCol = M_INT_COL盘点成本金额差 Then
                            .ColWidth(intCol) = 1500
                        Else
                            .ColWidth(intCol) = 1000
                        End If
                    End If
                Next
            End If
            
            .ColHidden(M_INT_COL记录状态) = True
            
            .ColWidth(M_INT_COL审核日期) = 1000
            .ColWidth(M_INT_COL盘点成本金额) = IIf(mblnCostView = False, 0, 1500)
            .ColWidth(M_INT_COL盘点成本金额差) = IIf(mblnCostView = False, 0, 1500)
        Else
            If mblnBootUp = False Then
                .ColWidth(M_INT_COL盘点单盘点时间) = 2000
                .ColWidth(M_INT_COL盘点单摘要) = 3000
            End If
            .ColWidth(M_INT_COL盘点单外观) = 0
        End If
        Call RestoreFlexState(mshlist, TabShow.TabCaption(TabShow.Tab))
        If TabShow.Tab = 1 Then
            .ColWidth(M_INT_COL盘点成本金额) = IIf(mblnCostView = True, 1500, 0)
            .ColWidth(M_INT_COL盘点成本金额差) = IIf(mblnCostView = True, 1500, 0)
        End If
    End With
End Sub

Private Sub SetDetailColWidth()
    Dim intCol As Integer
    
    With mshDetail
        .ColAlignment(4) = flexAlignCenterCenter   '单位
        .ColAlignment(IIf(TabShow.Tab = 1, 9, 7)) = flexAlignRightCenter '实盘数
        If TabShow.Tab = 1 Then
            .ColAlignment(8) = flexAlignRightCenter     '帐面数
            .ColAlignment(10) = flexAlignCenterCenter    '标志
            .ColAlignment(11) = flexAlignRightCenter     '数量差
            .ColAlignment(12) = flexAlignRightCenter    '成本价
            .ColAlignment(13) = flexAlignRightCenter    '售价
            .ColAlignment(14) = flexAlignRightCenter    '金额差
            .ColAlignment(15) = flexAlignRightCenter    '差价差
            .ColAlignment(16) = flexAlignRightCenter    '盘点金额
            .ColAlignment(.Cols - 2) = flexAlignRightCenter '盘点成本金额
            .ColAlignment(.Cols - 1) = flexAlignRightCenter '盘点成本金额差
        Else
            .ColAlignment(8) = flexAlignRightCenter '成本价
            .ColAlignment(9) = flexAlignRightCenter '成本金额
            .ColAlignment(10) = flexAlignRightCenter '售价
            .ColAlignment(11) = flexAlignRightCenter '售价金额
            
        End If
        
        If TabShow.Tab = 1 Then
            .ColWidth(.Cols - 1) = 1500
            .ColWidth(.Cols - 2) = 1500
            
            If mblnBootUp = False Then
                .ColWidth(0) = 0
                .ColWidth(1) = 2500
                For intCol = 2 To .Cols - 1
                    .ColWidth(intCol) = 1000
                    If intCol = .Cols - 1 Or intCol = .Cols - 2 Then
                        .ColWidth(intCol) = 1500
                    End If
                Next
                If mlngMode = 1300 Then
                    .ColWidth(16) = 0
                End If
                .ColWidth(.Cols - 2) = 0
            End If
            
            .ColWidth(12) = IIf(mblnCostView = False, 0, 1000)
            .ColWidth(15) = IIf(mblnCostView = False, 0, 1000)
            .ColWidth(.Cols - 2) = IIf(mblnCostView = False, 0, 1500)
            .ColWidth(.Cols - 1) = IIf(mblnCostView = False, 0, 1500)
        Else
            .ColWidth(0) = 0
            .ColWidth(1) = 2500
            For intCol = 2 To .Cols - 1
                .ColWidth(intCol) = 1000
            Next
        End If
        Call RestoreFlexState(mshDetail, TabShow.TabCaption(TabShow.Tab))
        
        If TabShow.Tab = 1 Then
            .ColWidth(12) = IIf(mblnCostView = False, 0, 1000)
            .ColWidth(15) = IIf(mblnCostView = False, 0, 1000)
            .ColWidth(.Cols - 2) = IIf(mblnCostView = False, 0, 1500)
            .ColWidth(.Cols - 1) = IIf(mblnCostView = False, 0, 1500)
        End If
    End With
End Sub


'根据权限设置不同的显示项目
Private Sub SetVisable()
    '外购入库所有权限：参数设置、基本、所有库房、登记、修改、删除、验收、冲销、单据打印
'    If InStr(1, gstrPrivs, "参数设置") = 0 Then
'         mnuFileParameter.Visible = False
'         mnuFileLine3.Visible = False                '相应的分割线
'    End If
'
    If InStr(1, gstrPrivs, "登记") = 0 Then
        mnuEditAddBill.Visible = False
        mnuEditAddTable.Visible = False
        tlbTool.Buttons("Bill").Visible = False
        tlbTool.Buttons("Table").Visible = False
    End If
    
    If InStr(1, gstrPrivs, "修改") = 0 Then
        mnuEditModify.Visible = False
        tlbTool.Buttons("Modify").Visible = False
    End If
    
    If InStr(1, gstrPrivs, "删除") = 0 Then
        mnuEditDel.Visible = False
        tlbTool.Buttons("Delete").Visible = False
         '对没有所有编辑权限时，把菜单和工具栏上的相应的分割线屏蔽。
        If mnuEditAddBill.Visible = False And mnuEditModify.Visible = False Then
            mnuEditLine1.Visible = False
            tlbTool.Buttons("EditSeparate").Visible = False
        End If
    End If
    
    If InStr(1, gstrPrivs, "审核") = 0 Then
        mnuEditVerify.Visible = False
        tlbTool.Buttons("Verify").Visible = False
    End If
    
    If InStr(1, gstrPrivs, "冲销") = 0 Then
        mnuEditStrike.Visible = False
        tlbTool.Buttons("Strike").Visible = False
        
        If mnuEditVerify.Visible = False Then
            mnuEditLine2.Visible = False
            tlbTool.Buttons("VerifySeparate").Visible = False
        End If
    End If
    If InStr(1, gstrPrivs, "单据打印") = 0 Then
        mnuFileBillPrint.Visible = False
        mnuFileBillPreview.Visible = False
    End If
End Sub

Private Sub Cmd查阅_Click()
    Call mnuEditDisplay_Click
End Sub

Private Sub Form_Load()
    Dim strReg As String
    Dim strOthers(0 To 6) As String
    Dim i As Integer
    For i = 0 To 6
        strOthers(i) = ""
    Next
    mstrOthers = strOthers
    strReg = Val(zlDataBase.GetPara("盘点表单位", glngSys, mlngMode, "0"))
    mintUnit = Val(strReg)
    mintUnit1 = IIf(Val(zlDataBase.GetPara("记录单单位", glngSys, mlngMode, "0")) = 1, 1, 0)
    mstrOrder = zlDataBase.GetPara("单据排序", glngSys, mlngMode, "00")
    mblnCostView = zlStr.IsHavePrivs(mstrPrivs, "查看成本价")
  
    '刘兴宏:增加小数格式化串
    With mOraFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价, True)
        .FM_金额 = GetFmtString(mintUnit, g_金额, True)
        .FM_零售价 = GetFmtString(mintUnit, g_售价, True)
        .FM_数量 = GetFmtString(mintUnit, g_数量, True)
    End With
    With mORaFMT记录单
        .FM_成本价 = GetFmtString(mintUnit1, g_成本价, True)
        .FM_金额 = GetFmtString(mintUnit1, g_金额, True)
        .FM_零售价 = GetFmtString(mintUnit1, g_售价, True)
        .FM_数量 = GetFmtString(mintUnit1, g_数量, True)
    End With

    '恢复设置
    Me.Caption = mstrTitle
    PrintRange "查询范围:" & Format(sys.Currentdate, "yyyy年MM月dd日") & "至" & Format(sys.Currentdate, "yyyy年MM月dd日")
    Call RestoreWinState(Me, App.ProductName)
    
    '2006-04-25:刘兴宏,统一增加报表发布到模块的功能
    Call zlDataBase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
     
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
        .Height = 300
        .Left = 0
        .Width = cbrTool.Width
        
    End With
    
    With lbl成本金额差
        .Left = Me.Width - .Width - 1700
        .Top = picSeparate_s.Height - 200
    End With
    If mblnCostView = False Then
        lbl成本金额差.Visible = False
    End If
    
    With lblSum成本金额
        .Left = lbl成本金额差.Left - .Width - 600
        .Top = picSeparate_s.Height - 200
    End With
    If mblnCostView = False Then
        lblSum成本金额.Visible = False
    End If
    
    With TabShow
        .Top = IIf(cbrTool.Visible, cbrTool.Height, 0)
        .Left = 0
        .Width = cbrTool.Width
    End With
    
    With mshlist
        .Top = TabShow.Top + TabShow.Height
        .Left = 0
        .Width = cbrTool.Width
        .Height = picSeparate_s.Top - .Top
    End With
    
    With Cmd查阅
        .Left = Me.ScaleWidth - .Width - 100
        .Top = mshlist.Top + mshlist.Height + 30
        
    End With
    
    With mshDetail
        .Top = picSeparate_s.Top + picSeparate_s.Height + 100
        .Left = 0
        .Height = ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
        .Width = cbrTool.Width
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, mstrTitle
    Call SaveFlexState(mshlist, TabShow.TabCaption(TabShow.Tab))
End Sub

Private Sub mnuEditaddBill_Click()
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    frmCheckCourseCard.ShowCard Me, strNo, 1, , mstrPrivs, blnSuccess
    
    If blnSuccess Then Call mnuViewRefresh_Click
End Sub

Private Sub mnuEditAddTableAuto_Click()
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    frmCheckCard.ShowCard Me, strNo, 1, , mstrPrivs, blnSuccess
    
    If blnSuccess Then Call mnuViewRefresh_Click
End Sub

Private Sub mnuEditAddTableTotal_Click()
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    frmCheckCard.ShowCard Me, strNo, 5, , mstrPrivs, blnSuccess
    
    If blnSuccess Then Call mnuViewRefresh_Click
End Sub

Private Sub mnuEditAddTableZero_Click()
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    frmCheckCard.ShowCard Me, strNo, 6, , mstrPrivs, blnSuccess
    
    If blnSuccess Then Call mnuViewRefresh_Click
End Sub

Private Sub mnuEditVerify_Click()
    '验收
    
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    With mshlist
        strNo = .TextMatrix(.Row, M_INT_COLNO)
        frmCheckCard.ShowCard Me, strNo, 3, .TextMatrix(.Row, M_INT_COL记录状态), mstrPrivs, blnSuccess
    End With
    
    If blnSuccess Then Call mnuViewRefresh_Click
End Sub

Private Sub mnuEditDel_Click()
    '删除
    Dim StrBillNo As String
    Dim intRow As Integer
    Dim strTitle As String
    Dim intReturn As Integer
    Dim intRecord As Integer
     
    With mshlist
        strTitle = IIf(TabShow.Tab = 0, "盘点记录单", "盘点表")
        
        On Error GoTo errHandle
        intRow = .Row
        StrBillNo = .TextMatrix(intRow, M_INT_COLNO)
        intReturn = MsgBox("你确实要删除单据号为“" & StrBillNo & "”的" & strTitle & "吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
        intRecord = .Rows - 1
        If intReturn = vbYes Then
            If TabShow.Tab = 1 Then
                gstrSQL = "zl_材料盘点_Delete('" & StrBillNo & "')"
            Else
                gstrSQL = "zl_材料盘点记录单_Delete('" & StrBillNo & "')"
            End If
            zlDataBase.ExecuteProcedure gstrSQL, Me.Caption
            
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
'                    .ColSel = .Cols - 1
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
'            .ColSel = .Cols - 1
           mshlist_EnterCell
        End If
    End With
    stbThis.Panels(2).Text = "当前共有" & intRecord & "张单据"
    Exit Sub

errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditDisplay_Click()
    '查看单据
    
    Dim strNo As String
    With mshlist
        strNo = .TextMatrix(.Row, M_INT_COLNO)
        If TabShow.Tab = 0 Then
            frmCheckCourseCard.ShowCard Me, strNo, 4
        Else
            frmCheckCard.ShowCard Me, strNo, 4, .TextMatrix(.Row, M_INT_COL记录状态), mstrPrivs
        End If
    End With
End Sub

Private Sub mnuEditStrike_Click()
    Dim blnPurchase As Boolean, blnRefresh As Boolean
    
    '如果是外购(blnPurchase为真)，则直接进入冲销
    '询问是否冲销(blnPurchase为提示框返回值)，是则进入冲销
    blnPurchase = (InStr(1, "1300,1302,1304,1305,1306", mlngMode) <> 0)
    With mshlist
        If Not blnPurchase Then
            blnPurchase = (MsgBox("你确实要冲销单据号为“" & .TextMatrix(.Row, M_INT_COLNO) & "”的单据吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes)
        End If
        If blnPurchase Then
            blnRefresh = StrikeSave
            If blnRefresh Then mnuViewRefresh_Click
        End If
    End With
End Sub

Private Function StrikeSave() As Boolean
    Dim blnSuccess As Boolean
    Dim rsTemp As ADODB.Recordset
    Dim int库存检查 As Integer
    Dim strMsg As String
    Dim n As Integer
    
    StrikeSave = False
    
    int库存检查 = StuffWork_GetCheckStockRule(Val(cboStock.ItemData(cboStock.ListIndex)))
    
    On Error GoTo errHandle
    If int库存检查 <> 0 Then
        gstrSQL = "Select Distinct A.药品信息 " & _
            " From (Select  '(' || I.编码 || ')' || Nvl(N.名称, I.名称) As 药品信息, A.实际数量, Nvl(K.实际数量, 0) As 库存数量 " & _
            " From 药品收发记录 A, (Select 药品id, 库房id, 实际数量, Nvl(批次, 0) 批次 From 药品库存 Where 性质 = 1) K, 材料特性 B, 收费项目目录 I, 收费项目别名 N " & _
            " Where A.药品id = K.药品id(+) And A.库房id = K.库房id(+) And Nvl(A.批次, 0) = K.批次(+) And A.药品id = B.材料id And " & _
            " A.药品id = I.ID And A.药品id = N.收费细目id(+) And N.性质(+) = 3 And A.单据 = 22 And A.入出系数 = 1 And A.NO = [1]) A " & _
            " Where A.实际数量 > A.库存数量 "
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "检查库存", mshlist.TextMatrix(mshlist.Row, 0))
        
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
    
    With mshlist
        gstrSQL = "zl_材料盘点_Strike('" & .TextMatrix(.Row, M_INT_COLNO) & "','" & UserInfo.用户名 & "')"
        
        zlDataBase.ExecuteProcedure gstrSQL, Me.Caption
    End With
    StrikeSave = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub mnuEditModify_Click()
    '修改
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    blnSuccess = False
    With mshlist
        If .TextMatrix(.Row, M_INT_COLNO) = "" Then Exit Sub
        strNo = .TextMatrix(.Row, M_INT_COLNO)
        If TabShow.Tab = 0 Then
            frmCheckCourseCard.ShowCard Me, strNo, 2, 1, mstrPrivs, blnSuccess
        Else
            frmCheckCard.ShowCard Me, strNo, 2, mshlist.TextMatrix(mshlist.Row, M_INT_COL记录状态), mstrPrivs, blnSuccess
        End If
        
        If blnSuccess Then Call mnuViewRefresh_Click
    End With
End Sub

Private Sub mnuFileBillPreview_Click()
    With mshlist
        If .TextMatrix(.Row, M_INT_COLNO) = "" Then Exit Sub
        ReportOpen gcnOracle, glngSys, "zl1_bill_1719", Me, "单据编号=" & .TextMatrix(.Row, M_INT_COLNO), "记录状态=" & .TextMatrix(.Row, M_INT_COL记录状态), "单位系数=" & mintUnit, 1
    End With
End Sub

Private Sub mnuFileBillPrint_Click()
    With mshlist
        If .TextMatrix(.Row, M_INT_COLNO) = "" Then Exit Sub
        ReportOpen gcnOracle, glngSys, "zl1_bill_1719", Me, "单据编号=" & .TextMatrix(.Row, M_INT_COLNO), "记录状态=" & .TextMatrix(.Row, M_INT_COL记录状态), "单位系数=" & mintUnit, 2
    End With
End Sub

Private Sub mnuFileExcel_Click()
    '输出到Excel
    
    If Me.ActiveControl Is mshlist Then
        mshlist.Redraw = False
        subPrint 3
        mshlist.Redraw = True
        mshlist.Col = 0
'        mshlist.ColSel = mshlist.Cols - 1
    ElseIf Me.ActiveControl Is mshDetail Then
        mshDetail.Redraw = False
        subExcel 3
        mshDetail.Redraw = True
        mshDetail.Col = 0
'        mshDetail.ColSel = mshDetail.Cols - 1
    End If
End Sub

Private Sub mnufileexit_Click()
    '退出
    Unload Me
End Sub

Private Sub mnuFileParameter_Click()
    Dim strReg As String
    '参数设置
    Call frmParaset.设置参数(mlngMode, mstrPrivs, Me, mstrCaption)
     
    '盘点记录单的单位
    mintUnit = Val(zlDataBase.GetPara("盘点表单位", glngSys, mlngMode, "0"))
    mintUnit1 = IIf(Val(zlDataBase.GetPara("记录单单位", glngSys, mlngMode, "0")) = 1, 1, 0)
    mstrOrder = zlDataBase.GetPara("单据排序", glngSys, mlngMode, "00")
    mintFindDay = Val(zlDataBase.GetPara("查询天数", glngSys, mlngMode, 1))
    mdtStartDate = Format(DateAdd("d", -mintFindDay, sys.Currentdate), "yyyy-MM-dd")
    mdtEndDate = Format(sys.Currentdate, "yyyy-MM-dd")
  
    '刘兴宏:增加小数格式化串
    With mOraFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价, True)
        .FM_金额 = GetFmtString(mintUnit, g_金额, True)
        .FM_零售价 = GetFmtString(mintUnit, g_售价, True)
        .FM_数量 = GetFmtString(mintUnit, g_数量, True)
    End With
    With mORaFMT记录单
        .FM_成本价 = GetFmtString(mintUnit1, g_成本价, True)
        .FM_金额 = GetFmtString(mintUnit1, g_金额, True)
        .FM_零售价 = GetFmtString(mintUnit1, g_售价, True)
        .FM_数量 = GetFmtString(mintUnit1, g_数量, True)
    End With
    
    Call GetList(mstrFind)
End Sub

Private Sub mnuFilePreView_Click()
    '打印预览
    mshlist.Redraw = False
    subPrint 2
    mshlist.Redraw = True
    mshlist.Col = 0
'    mshlist.ColSel = mshlist.Cols - 1
End Sub

Private Sub mnuFilePrint_Click()
    '打印
    mshlist.Redraw = False
    subPrint 1
    mshlist.Redraw = True
    mshlist.Col = 0
'    mshlist.ColSel = mshlist.Cols - 1
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
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
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
    
    Dim strFind As String
    Dim strOthers() As String
    strFind = FrmTransferSearch.GetSearch(Me, mlngMode, mdtStartDate, mdtEndDate, mdtVerifyStart, mdtVerifyEnd, mstrPrivs, strOthers)
    If strFind <> "" Then
        mstrFind = strFind
        mstrOthers = strOthers
       'mstrOthers(0 To 6) As String ' 0-记录状态(计划类型),1-开始单号,2-结束单号,3-材料id,4-对方部门id(或入出类别id或编制方法(计划用)),5-填制人,6-审核人)
        
        GetList mstrFind
        
        If Format(mdtStartDate, "yyyy-mm-dd") = "1901-01-01" And Format(mdtVerifyStart, "yyyy-mm-dd") = "1901-01-01" Then
        ElseIf Format(mdtStartDate, "yyyy-mm-dd") <> "1901-01-01" And Format(mdtVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
            PrintRange "查询范围:填制日期 " & Format(mdtStartDate, "yyyy年MM月dd日") & "至" & Format(mdtEndDate, "yyyy年MM月dd日") & "  审核日期 " & Format(mdtVerifyStart, "yyyy年MM月dd日") & "至" & Format(mdtVerifyEnd, "yyyy年MM月dd日")
        ElseIf Format(mdtStartDate, "yyyy-mm-dd") <> "1901-01-01" Then
            PrintRange "查询范围:填制日期 " & Format(mdtStartDate, "yyyy年MM月dd日") & "至" & Format(mdtEndDate, "yyyy年MM月dd日")
        ElseIf Format(mdtVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
            PrintRange "查询范围:审核日期 " & Format(mdtVerifyStart, "yyyy年MM月dd日") & "至" & Format(mdtVerifyEnd, "yyyy年MM月dd日")
        End If
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
    
    With mshlist
        strNo = Trim(.TextMatrix(.Row, M_INT_COLNO))
        lngCol = GetCol(mshlist, "记录状态")
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
    With mshlist
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
    If mshlist.MouseRow = 0 Then Exit Sub
    mnuEditModify_Click
End Sub

Private Sub mshlist_EnterCell()
    Dim rsTemp As New Recordset
    Dim strUnitQuantity As String               '单位和数量格式化串
    Dim IntBill As Integer                      '单据类型  如：1、外购入库；2、
    Dim str包装系数 As String
    Dim intTmp As Integer
    Dim str排序 As String
    Dim str列名 As String
    
    If mblnLoadGrid = False Then Exit Sub
    If mlastRow = mshlist.Row Then Exit Sub
    mlastRow = mshlist.Row
        
    
    On Error GoTo errHandle
    
    If Mid(mstrOrder, 1, 1) = "0" Then
        str排序 = " 序号"
    ElseIf Mid(mstrOrder, 1, 1) = "1" Then
        str排序 = " 卫材信息"
    ElseIf Mid(mstrOrder, 1, 1) = "2" Then
        str排序 = " 名称"
    ElseIf Mid(mstrOrder, 1, 1) = "3" Then
        str排序 = " 库房货位"
    End If
    
    If Mid(mstrOrder, 2, 1) = "0" Then
        str排序 = str排序 & " asc"
    ElseIf Mid(mstrOrder, 2, 1) = "1" Then
        str排序 = str排序 & " desc"
    End If
    
    If mshlist.Row >= 1 And LTrim(mshlist.TextMatrix(mshlist.Row, M_INT_COLNO)) <> "" Then
        mshlist.Col = 0
        mshlist.ColSel = mshlist.Cols - 1
        If mshlist.RowIsVisible(mshlist.Row) = False Then
           mshlist.TopRow = mshlist.Row
        End If
        mshDetail.Redraw = False
        intTmp = IIf(TabShow.Tab = 1, mintUnit, mintUnit1)
        Select Case intTmp
            Case 0
                strUnitQuantity = "to_char(A.实际数量," & mOraFMT.FM_数量 & ") AS 数量," & _
                "c.计算单位 AS 单位,"
                str包装系数 = "1"
            Case Else
                strUnitQuantity = "(to_char(A.实际数量 / B.换算系数," & mOraFMT.FM_数量 & ")) AS 数量," & _
                "B.包装单位 AS 单位,"
                str包装系数 = "B.换算系数"
        End Select
            
        IntBill = IIf(TabShow.Tab = 1, 22, 23)
        Dim int记录状态 As Integer
        
        If TabShow.Tab = 1 Then
            str列名 = "序号,卫材信息,规格,产地,批准文号,单位,批号,失效期,帐面数,实盘数,标志,数量差,成本价,售价,金额差,差价差,盘点金额,盘点成本金额,盘点成本金额差"
            gstrSQL = "" & _
                "   SELECT " & str列名 & _
                "   FROM (  SELECT DISTINCT 序号,('[' || c.编码 || ']' || c.名称) AS 卫材信息," & _
                "                   c.规格,c.编码,zlSpellCode(c.名称) 名称,a.产地,a.批准文号,a.库房货位," & IIf(mintUnit = 0, "c.计算单位", "b.包装单位") & " as 单位,a.批号, to_char(A.效期,'yyyy-mm-dd') as 失效期," & _
                "                   (to_char(A.填写数量 /" & str包装系数 & "," & mOraFMT.FM_数量 & ")) AS 帐面数," & _
                "                   (to_char(A.扣率 /" & str包装系数 & "," & mOraFMT.FM_数量 & ")) AS 实盘数," & _
                "                   Decode(Sign(A.扣率-A.填写数量),-1,'亏',1,'盈','平') as 标志," & _
                "                   (to_char(A.实际数量 /" & str包装系数 & "," & mOraFMT.FM_数量 & ")) AS 数量差," & _
                "                   TO_CHAR (a.单量*" & str包装系数 & "," & mOraFMT.FM_成本价 & ") AS 成本价," & _
                "                   TO_CHAR (a.零售价*" & str包装系数 & "," & mOraFMT.FM_零售价 & ") AS 售价," & _
                "                   TO_CHAR (a.零售金额*a.入出系数, " & mOraFMT.FM_金额 & ") AS 金额差," & _
                "                   TO_CHAR (a.差价*a.入出系数, " & mOraFMT.FM_金额 & ") AS 差价差, " & _
                "                   TO_CHAR ((A.扣率 / " & str包装系数 & ")*(a.零售价*" & str包装系数 & "), " & mOraFMT.FM_金额 & ") as 盘点金额, " & _
                "                   To_Char(a.成本价+to_char(a.零售金额*a.入出系数*Decode(a.记录状态, 1, 1, Decode(Mod(a.记录状态, 3), 0, 1, -1))," & mOraFMT.FM_金额 & ")" & "-(a.成本金额+to_char(a.差价*a.入出系数*Decode(a.记录状态, 1, 1, Decode(Mod(a.记录状态, 3), 0, 1, -1))," & mOraFMT.FM_金额 & "))," & mOraFMT.FM_金额 & ") as  盘点成本金额, " & _
                "                   To_Char(a.零售金额*a.入出系数  - a.差价*a.入出系数 ," & mOraFMT.FM_金额 & ") AS 盘点成本金额差 " & _
                "           FROM 药品收发记录 a, 材料特性  b,收费项目目录 c" & _
                "           Where a.药品id = b.材料id and a.药品id=c.id " & _
                "                   AND  a.记录状态 =[3] " & _
                "                   AND a.单据 =[1] " & _
                "                   AND a.no =[2] " & _
                "   )" & _
                "  ORDER BY " & str排序
                
            int记录状态 = Val(mshlist.TextMatrix(mshlist.Row, M_INT_COL记录状态))
        Else
            
            str列名 = "序号,卫材信息,规格,产地,单位,批号,失效期,实盘数,成本价,成本金额,售价,售价金额"
            gstrSQL = "" & _
                "   SELECT " & str列名 & _
                "   FROM (  SELECT DISTINCT 序号,('[' || c.编码 || ']' || c.名称) AS 卫材信息," & _
                "                   c.规格,c.编码,zlSpellCode(c.名称) 名称,a.产地,a.库房货位," & IIf(mintUnit1 = 0, "c.计算单位", "b.包装单位") & " as 单位,a.批号, to_char(A.效期,'yyyy-mm-dd') as 失效期," & _
                "                   (to_char(A.扣率 /" & str包装系数 & "," & mORaFMT记录单.FM_数量 & ")) AS 实盘数," & _
                                    IIf(mintUnit1 = 0, "to_char(a.单量," & mORaFMT记录单.FM_成本价 & ") 成本价, ", "to_char(a.单量 * " & str包装系数 & "," & mORaFMT记录单.FM_成本价 & ") 成本价,") & _
                "                   to_char(A.单量 * A.扣率," & mORaFMT记录单.FM_金额 & ") 成本金额," & _
                                    IIf(mintUnit1 = 0, "to_char(a.零售价," & mORaFMT记录单.FM_零售价 & ") 售价,", "to_char(a.零售价 * " & str包装系数 & "," & mORaFMT记录单.FM_零售价 & ") 售价,") & _
                "                   to_char(A.零售价 * A.扣率," & mORaFMT记录单.FM_金额 & ") 售价金额" & _
                "           FROM 药品收发记录 a, 材料特性  b,收费项目目录 c " & _
                "           Where a.药品id = b.材料id and a.药品id=c.id " & _
                "                   AND a.记录状态 =[3] AND a.单据 = [1] " & _
                "                   AND a.no = [2] " & _
                "       )" & _
                " ORDER BY  " & str排序
                int记录状态 = 1
        End If
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, IntBill, mshlist.TextMatrix(mshlist.Row, M_INT_COLNO), int记录状态)
        
        Set mshDetail.DataSource = rsTemp
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
'            .ColSel = .Cols - 1
        End With
        
        mshDetail.Redraw = True
    ElseIf LTrim(mshlist.TextMatrix(mshlist.Row, M_INT_COLNO)) = "" Then
        With mshDetail
            .Cols = IIf(TabShow.Tab = 1, 19, 12)
            .Rows = 2
            .Clear
            .TextMatrix(0, 0) = "序号"
            .TextMatrix(0, 1) = "卫材信息"
            .TextMatrix(0, 2) = "规格"
            .TextMatrix(0, 3) = "产地"
            .TextMatrix(0, 4) = "批准文号"
            .TextMatrix(0, 5) = "单位"
            .TextMatrix(0, 6) = "批号"
            .TextMatrix(0, 7) = "失效期"
            .TextMatrix(0, IIf(TabShow.Tab = 1, 9, 7)) = "实盘数"
            
            If TabShow.Tab = 1 Then
                .TextMatrix(0, 8) = "帐面数"
                .TextMatrix(0, 10) = "标志"
                .TextMatrix(0, 11) = "数量差"
                .TextMatrix(0, 12) = "成本价"
                .TextMatrix(0, 13) = "售价"
                .TextMatrix(0, 14) = "金额差"
                .TextMatrix(0, 15) = "差价差"
                .TextMatrix(0, 16) = "盘点金额"
                .TextMatrix(0, 17) = "盘点成本金额"
                .TextMatrix(0, 18) = "盘点成本金额差"
            Else
                .TextMatrix(0, 8) = "成本价"
                .TextMatrix(0, 9) = "成本金额"
                .TextMatrix(0, 10) = "售价"
                .TextMatrix(0, 11) = "售价金额"
            End If
            
            .Row = 1
            .Col = 0
'            .ColSel = .Cols - 1
        End With
    End If
    SetDetailColWidth
    SetEnable
    Exit Sub
errHandle:
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
    
    PopupMenu mnuEdit, 2
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
    
    With mshlist
        .Height = picSeparate_s.Top - .Top
    End With
    
    With Cmd查阅
        .Top = mshlist.Top + mshlist.Height + 30
    End With
    
    With mshDetail
        .Top = picSeparate_s.Top + picSeparate_s.Height + 100
        .Height = ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
    End With
End Sub

Private Sub picSeparate_s_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        If Button <> 1 Then Exit Sub
        mintOldY = 0
End Sub

Private Sub tabShow_Click(PreviousTab As Integer)
    Call SaveFlexState(mshlist, TabShow.TabCaption(PreviousTab))
    Call initGrid
    GetList (mstrFind)  '列出单据头
    
    If TabShow.Tab = 1 Then
        lblSum成本金额.Visible = True
        lbl成本金额差.Visible = True
    Else
        lblSum成本金额.Visible = False
        lbl成本金额差.Visible = False
    End If
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
    mnuEditVerify.Visible = blnVisible And (InStr(1, gstrPrivs, "审核") <> 0)
    mnuEditStrike.Visible = blnVisible And (InStr(1, gstrPrivs, "冲销") <> 0)
    mnuEditLine2.Visible = blnVisible And (mnuEditVerify.Visible Or mnuEditStrike.Visible)
    tlbTool.Buttons("Verify").Visible = blnVisible And (InStr(1, gstrPrivs, "审核") <> 0)
    tlbTool.Buttons("Strike").Visible = blnVisible And (InStr(1, gstrPrivs, "冲销") <> 0)
    
    tlbTool.Buttons("VerifySeparate").Visible = mnuEditLine2.Visible
    
    mnuFileBillPreview.Visible = blnVisible And (InStr(1, gstrPrivs, "单据打印") <> 0)
    mnuFileBillPrint.Visible = blnVisible And (InStr(1, gstrPrivs, "单据打印") <> 0)
    
    With mshlist
        .ToolTipText = ""
        If .TextMatrix(.Row, M_INT_COLNO) = "" Or .Row = 0 Then          '没有单
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
            
            If TabShow.Tab = 1 Then
                strVerify = .TextMatrix(.Row, M_INT_COL审核日期)
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
            ElseIf .TextMatrix(.Row, M_INT_COL记录状态) = 1 Then     '审核单
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
                If .TextMatrix(.Row, M_INT_COL记录状态) Mod 3 = 0 Then
                    .ToolTipText = "冲销单据的原单据"
                    If mnuEditStrike.Visible = True Then
                        mnuEditStrike.Enabled = True
                        tlbTool.Buttons("Strike").Enabled = True
                    End If
                ElseIf .TextMatrix(.Row, M_INT_COL记录状态) Mod 3 = 2 Then
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
    
    If Format(mdtStartDate, "yyyy-mm-dd") = "1901-01-01" And Format(mdtVerifyStart, "yyyy-mm-dd") = "1901-01-01" Then
        strRange = "审核日期 " & Format(mdtVerifyStart, "yyyy年MM月dd日") & "至" & Format(mdtVerifyEnd, "yyyy年MM月dd日")
    ElseIf Format(mdtVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
        strRange = "填制日期 " & Format(mdtStartDate, "yyyy年MM月dd日") & "至" & Format(mdtEndDate, "yyyy年MM月dd日") & "  审核日期 " & Format(mdtVerifyStart, "yyyy年MM月dd日") & "至" & Format(mdtVerifyEnd, "yyyy年MM月dd日")
    Else
        strRange = "填制日期 " & Format(mdtStartDate, "yyyy年MM月dd日") & "至" & Format(mdtEndDate, "yyyy年MM月dd日")
    End If
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "楷体 & _GB2312"
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
    
    Set objPrint.Body = mshlist
    
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
    objPrint.Title.Font.Name = "楷体 & _GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
    objPrint.Title.Text = mstrTitle
        
    Set objRow = New zlTabAppRow
    objRow.Add ""
    objRow.Add "NO." & Trim(mshlist.TextMatrix(mshlist.Row, GetCol(mshlist, "NO")))
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "盘点库房：" & Trim(cboStock.Text)
    objRow.Add "盘点时间：" & Trim(mshlist.TextMatrix(mshlist.Row, GetCol(mshlist, "盘点时间")))
    objPrint.UnderAppRows.Add objRow
        
    Set objRow = New zlTabAppRow
    objRow.Add "摘要:" & mshlist.TextMatrix(mshlist.Row, GetCol(mshlist, "摘要"))
    objPrint.BelowAppRows.Add objRow
        
    Set objRow = New zlTabAppRow
    objRow.Add "填制人:" & mshlist.TextMatrix(mshlist.Row, GetCol(mshlist, "填制人")) & "  填制日期:" & mshlist.TextMatrix(mshlist.Row, GetCol(mshlist, "填制日期"))
    
    objRow.Add "审核人:  " & "  审核日期:  "
    objPrint.BelowAppRows.Add objRow
    
    Set objPrint.Body = mshDetail
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

'对单据头列排序
Private Sub ListSort()
    Dim intCol As Integer
    Dim intRow As Integer
    Dim intTemp As String
    
    With mshlist
        If .Rows > 1 Then
            .Redraw = False
            intCol = .MouseCol
            .Col = intCol
            .ColSel = intCol
            intTemp = .TextMatrix(.Row, M_INT_COLNO)

            If intCol = mintPreCol And mintsort = flexSortStringNoCaseDescending Then
                .Sort = flexSortStringNoCaseAscending
                mintsort = flexSortStringNoCaseAscending
            Else
               .Sort = flexSortStringNoCaseDescending
               mintsort = flexSortStringNoCaseDescending
            End If

            mintPreCol = intCol
            .Row = Grid.MshGrdFindRow(mshlist, intTemp, 0)
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
            
            mintPreDetailCol = intCol
            .Row = Grid.MshGrdFindRow(mshDetail, intTemp, 0)
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
    picSeparate_s.CurrentY = 100
    picSeparate_s.Print strRange
End Sub





Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

