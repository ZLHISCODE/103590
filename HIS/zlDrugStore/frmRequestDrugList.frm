VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmRequestDrugList 
   Caption         =   "药品申领管理"
   ClientHeight    =   4980
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9615
   Icon            =   "frmRequestDrugList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picColor 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4320
      ScaleHeight     =   255
      ScaleWidth      =   3615
      TabIndex        =   11
      Top             =   4200
      Width           =   3615
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
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1320
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   0
         Width           =   260
      End
      Begin VB.Label lblColor2 
         AutoSize        =   -1  'True
         Caption         =   "正常"
         Height          =   180
         Left            =   1680
         TabIndex        =   17
         Top             =   37
         Width           =   360
      End
      Begin VB.Label lblColor1 
         AutoSize        =   -1  'True
         Caption         =   "正常冲销"
         Height          =   180
         Left            =   360
         TabIndex        =   16
         Top             =   37
         Width           =   720
      End
      Begin VB.Label lblColor3 
         AutoSize        =   -1  'True
         Caption         =   "财务冲销"
         Height          =   180
         Left            =   2640
         TabIndex        =   15
         Top             =   30
         Width           =   720
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   1005
      Left            =   120
      TabIndex        =   9
      Top             =   1080
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
      FormatString    =   $"frmRequestDrugList.frx":014A
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
   Begin VB.PictureBox picSeparate_s 
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   180
      MousePointer    =   7  'Size N S
      ScaleHeight     =   360
      ScaleWidth      =   4815
      TabIndex        =   3
      Top             =   2700
      Width           =   4815
      Begin VB.Label lbl3 
         AutoSize        =   -1  'True
         Caption         =   "差价金额："
         Height          =   180
         Left            =   3690
         TabIndex        =   8
         Top             =   20
         Width           =   900
      End
      Begin VB.Label lbl2 
         AutoSize        =   -1  'True
         Caption         =   "售价金额："
         Height          =   180
         Left            =   1890
         TabIndex        =   7
         Top             =   20
         Width           =   900
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         Caption         =   "成本金额："
         Height          =   180
         Left            =   0
         TabIndex        =   6
         Top             =   20
         Width           =   900
      End
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
      Caption2        =   "药房"
      Child2          =   "cboStock"
      MinWidth2       =   3000
      MinHeight2      =   300
      Width2          =   3345
      NewRow2         =   0   'False
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   8685
         TabIndex        =   5
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
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Hank"
                     Text            =   "手工填写"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Auto"
                     Text            =   "自动生成"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Sale"
                     Text            =   "自动按销售量生成"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Object.Visible         =   0   'False
                     Key             =   "Merge"
                     Text            =   "合并申领单"
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
               Caption         =   "接收"
               Key             =   "Receive"
               Object.ToolTipText     =   "接收"
               Object.Tag             =   "接收"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "冲销"
               Key             =   "DisReceive"
               Object.ToolTipText     =   "冲销"
               Object.Tag             =   "冲销"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "EditSeparate"
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
         MouseIcon       =   "frmRequestDrugList.frx":01BF
         Begin VB.Timer LimitTime 
            Enabled         =   0   'False
            Interval        =   8000
            Left            =   6660
            Top             =   180
         End
      End
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
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
            Picture         =   "frmRequestDrugList.frx":04D9
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
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugList.frx":0D6D
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugList.frx":0F8D
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugList.frx":11AD
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugList.frx":13C9
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugList.frx":15E9
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugList.frx":1809
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugList.frx":1A25
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugList.frx":1C41
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugList.frx":1E5B
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugList.frx":1FB5
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugList.frx":21D1
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
            Picture         =   "frmRequestDrugList.frx":23F1
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugList.frx":2611
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugList.frx":2831
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugList.frx":2A4D
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugList.frx":2C6D
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugList.frx":2E8D
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugList.frx":30A9
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugList.frx":32C5
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugList.frx":34DF
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugList.frx":3639
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugList.frx":3859
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
      Height          =   975
      Left            =   120
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
      FormatString    =   $"frmRequestDrugList.frx":3A79
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
         Begin VB.Menu mnuEditAddHank 
            Caption         =   "手工填写(&H)"
         End
         Begin VB.Menu mnuEditAddAuto 
            Caption         =   "自动生成(&A)"
         End
         Begin VB.Menu mnuEditAddAutoBySale 
            Caption         =   "自动按销售量生成(&S)"
         End
         Begin VB.Menu mnuEditAddMerge 
            Caption         =   "合并申领单(&M)"
            Visible         =   0   'False
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
      Begin VB.Menu mnuEditReceive 
         Caption         =   "审核(&R)"
      End
      Begin VB.Menu mnuEditDisReceive 
         Caption         =   "冲销(&D)"
      End
      Begin VB.Menu mnuEditWriteOff 
         Caption         =   "批量冲销(&W)"
      End
      Begin VB.Menu mnuEditLine2 
         Caption         =   "-"
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
Attribute VB_Name = "frmRequestDrugList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFind As String
Private mblnBootUp As Boolean
Private mlastRow As Long                '上次点击的行
Private mintPreCol As Integer           '前一次单据头的排序列
Private mintsort As Integer             '前一次单据头的排序
Private mintPreDetailCol As Integer     '前一次单据体的排序列
Private mintDetailsort As Integer       '前一次单据体的排序
Private mlngMode As Long
Private mstrPrivs As String             '当前用户具有的当前模块的功能
Private mint查询天数 As Integer
Private mblnViewCost As Boolean     '查看成本价 true-可以查看成本价 false-不可以查看成本价
Private Const MStrCaption As String = "药品申领管理"

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

Private mint冲销申请 As Integer                          '0-不需要申请;1-需要申请

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

Public Function CheckBill(ByVal strNo As String) As String
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHandle
    CheckBill = ""
    
    gstrSQL = " Select 审核日期,配药日期,配药人 From 药品收发记录 " & _
            " Where 单据=6 And NO=[1] And 记录状态=1 And RowNum=1 "
    Set rs = zlDataBase.OpenSQLRecord(gstrSQL, "检查单据", strNo)
    
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
Private Function CheckNoIsExist(ByVal StrBillNo As String) As Boolean
    Dim rsCheck As New ADODB.Recordset
    '检查单据是否存在
    On Error GoTo errHandle
    gstrSQL = " Select id From 药品收发记录 " & _
              " Where 单据=6 And NO=[1] And 入出系数 = -1 and rownum = 1"
    Set rsCheck = zlDataBase.OpenSQLRecord(gstrSQL, "检查单据是否存在", StrBillNo)
    CheckNoIsExist = Not (rsCheck.RecordCount = 0)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub cboStock_Click()
    If mlng库房id <> cboStock.ItemData(cboStock.ListIndex) Then
        mlng库房id = cboStock.ItemData(cboStock.ListIndex)
        Call GetDrugDigit(mlng库房id, MStrCaption, mintUnit, mintShowCostDigit, mintShowPriceDigit, mintShowNumberDigit, mintShowMoneyDigit)
        
        '组织格式化串
        mstrCostFormat = "'999999999990." & String(mintShowCostDigit, "0") & "'"
        mstrPriceFormat = "'999999999990." & String(mintShowPriceDigit, "0") & "'"
        mstrNumberFormat = "'999999999990." & String(mintShowNumberDigit, "0") & "'"
        mstrMoneyFormat = "'999999999990." & String(mintShowMoneyDigit, "0") & "'"
    
        If mblnBootUp Then mnuViewRefresh_Click
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
    
    If Select部门选择器(Me, cboStock, Trim(cboStock.Text), str工作性质, True, "0,1,2,3") = False Then
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
    If mblnBootUp = False Then Exit Sub
    Form_Resize
End Sub

Public Sub ShowList(ByVal FrmMain As Variant)
    Dim strFind As String
    Dim dateCurDate As Date
    
    mlngMode = glngModul
    mstrPrivs = gstrprivs

    mblnBootUp = False
    
    dateCurDate = Sys.Currentdate()
    mint查询天数 = Val(zlDataBase.GetPara("查询天数", glngSys, 1343)) - 1
    
    '取移库的冲销申请参数
    mint冲销申请 = Val(zlDataBase.GetPara("冲销申请", glngSys, 1304))
    
    strStart = Format(DateAdd("d", -1 * mint查询天数, dateCurDate), "yyyy-MM-dd")
    strEnd = Format(dateCurDate, "yyyy-MM-dd")
    strVerifyStart = "1901-01-01"
    strVerifyEnd = "1901-01-01"
        
    If Not CheckDepend Then Exit Sub            '数据依赖性测试
    
    mlng库房id = Me.cboStock.ItemData(Me.cboStock.ListIndex)
    Call GetDrugDigit(mlng库房id, MStrCaption, mintUnit, mintShowCostDigit, mintShowPriceDigit, mintShowNumberDigit, mintShowMoneyDigit)
    
    '组织格式化串
    mstrCostFormat = "'999999999990." & String(mintShowCostDigit, "0") & "'"
    mstrPriceFormat = "'999999999990." & String(mintShowPriceDigit, "0") & "'"
    mstrNumberFormat = "'999999999990." & String(mintShowNumberDigit, "0") & "'"
    mstrMoneyFormat = "'999999999990." & String(mintShowMoneyDigit, "0") & "'"
    
    SetVisable  '根据权限设置不同的显示项目
    
    strFind = " AND A.记录状态 = 1 And A.审核日期 is Null And A.填制日期 Between [3] And [4] "
    SQLCondition.date填制时间开始 = CDate(Format(strStart, "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date填制时间结束 = CDate(Format(strEnd, "yyyy-mm-dd") & " 23:59:59")
    
    mstrFind = strFind
    
    GetList (mstrFind)  '列出单据头
    RestoreWinState Me, App.ProductName, MStrCaption
    Call zlDataBase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs)
    
    If mblnViewCost = False Then
        vsfDetail.ColWidth(11) = 0
        vsfDetail.ColWidth(12) = 0
    End If
    
    mblnBootUp = True
        
    If IsObject(FrmMain) Then
        Me.Show , FrmMain
    Else
        'ZLBH融合调用
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
            & "Where c.工作性质 = b.名称 " _
              & "AND Instr([1],b.编码,1) > 0 " _
             & " AND a.id = c.部门id " _
              & "AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01'"

    Set rsDepend = zlDataBase.OpenSQLRecord(gstrSQL, "药品申领管理", strStock)
    
    If rsDepend.EOF Then
        MsgBox "部门性质信息不全,请查看部门管理！", vbInformation, gstrSysName
        rsDepend.Close
        Exit Function
    End If
            
    rsDepend.Close

    gstrSQL = "SELECT DISTINCT a.id, a.名称 " _
         & "FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " _
        & "Where (a.站点 = '" & gstrNodeNo & "' Or a.站点 is Null) And c.工作性质 = b.名称 " _
          & "AND Instr([1],b.编码,1) > 0 " _
         & " AND a.id = c.部门id " _
         & " and a.id in (select 部门id from 部门人员 where 人员id= [2]) " _
          & " AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01'"
    Set rsDepend = zlDataBase.OpenSQLRecord(gstrSQL, "药品申领管理", strStock, glngUserId)
    
    If rsDepend.EOF Then
        MsgBox "你不是药库、药房或制剂室的工作人员，不能进入！", vbInformation, gstrSysName
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
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetList(ByVal strFind As String)
    Dim rsList As New Recordset
    Dim strUserPart As String
    
    '用于统计合计金额
    Dim dbl1 As Double
    Dim dbl2 As Double
    Dim dbl3 As Double
    Dim n As Long
    Dim strFormat As String
    
    On Error GoTo errHandle
    strFormat = "0.00##"
    
    mlastRow = 0
    
    vsfList.Redraw = flexRDNone
    strUserPart = " And A.库房ID+0=[11] "
    gstrSQL = "SELECT A.NO, C.名称 AS 发药库房,LTRIM(TO_CHAR (SUM (A.成本金额)," & mstrCostFormat & ")) AS 成本金额, " & _
        " LTRIM(TO_CHAR ( (SUM (A.零售金额)), " & mstrMoneyFormat & ")) AS 售价金额,LTRIM(TO_CHAR (SUM (A.零售金额 - A.成本金额), " & mstrMoneyFormat & ")) AS 差价金额, A.填制人, " & _
        " TO_CHAR (MIN(A.填制日期), 'YYYY-MM-DD HH24:MI:SS') AS 填制日期,A.修改人,TO_CHAR (MIN(A.修改日期), 'YYYY-MM-DD HH24:MI:SS') AS 修改日期, A.审核人, " & _
        " TO_CHAR (MIN(A.审核日期), 'YYYY-MM-DD HH24:MI:SS') AS 审核日期, A.记录状态, A.配药人 发送人,A.摘要 " & _
        " FROM 药品收发记录 A, 部门表 B,部门表 C " & _
        " WHERE A.库房ID = B.ID AND A.对方部门ID=C.ID AND A.单据 = 6 AND  A.入出系数=1 " & _
        " And (A.配药人 Is NULL Or A.配药日期 Is Not NULL)" & _
        strUserPart & strFind & _
        " GROUP BY A.NO,C.名称,A.填制人,A.修改人,A.审核人,A.记录状态,A.配药人,A.摘要 " & _
        " ORDER BY NO DESC, 填制日期 ASC "
    Set rsList = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, _
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
        For n = 0 To .Cols - 1
            .ColKey(n) = .TextMatrix(0, n)
            .FixedAlignment(n) = flexAlignCenterCenter
        Next
        '.ColSel = .Cols - 1    ' bug: 40410
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
        
        lbl1.Caption = "成本金额合计：" & zlStr.FormatEx(dbl1, mintShowMoneyDigit, , True)
        lbl2.Caption = "售价金额合计：" & zlStr.FormatEx(dbl2, mintShowMoneyDigit, , True)
        lbl3.Caption = "差价金额合计：" & zlStr.FormatEx(dbl3, mintShowMoneyDigit, , True)

    End If
    vsfList_EnterCell    '列出单据体
    
    SetStrikeColor
    staThis.Panels(2).Text = "当前共有" & rsList.RecordCount & "张单据"
    rsList.Close
    vsfList.Redraw = flexRDDirect
        
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
            intStatus = .TextMatrix(intRow, .Cols - 3)
            If intStatus Mod 3 = 0 Then
                .Cell(flexcpForeColor, intRow, 0, intRow, .Cols - 1) = &H80000001
            End If
            If intStatus Mod 3 = 2 Then
                '移库中申请冲销单据为浅红色，已冲销单据为红色
                If Trim(.TextMatrix(intRow, GetCol(vsfList, "审核日期"))) <> "" Then
                    .Cell(flexcpForeColor, intRow, 0, intRow, .Cols - 1) = &HFF
                Else
                    .Cell(flexcpForeColor, intRow, 0, intRow, .Cols - 1) = &HFF00FF       ' &HC0C0FF
                End If
            End If
        Next
    End With
                
End Sub

'表头列宽初始
Private Sub SetListColWidth()
    Dim intCol As Integer
    
    With vsfList
        .ColAlignment(2) = flexAlignRightCenter
        .ColAlignment(3) = flexAlignRightCenter
        .ColAlignment(4) = flexAlignRightCenter

        For intCol = 1 To .Cols - 1
            If intCol = 1 Then
               .ColWidth(intCol) = 2000
            ElseIf intCol = .Cols - 3 Then
                .ColWidth(intCol) = 0
            Else
                .ColWidth(intCol) = 1000
            End If
        Next
        If mblnViewCost = False Then
            .ColHidden(.ColIndex("成本金额")) = True
            .ColHidden(.ColIndex("差价金额")) = True
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
    Set rsDetail = zlDataBase.OpenSQLRecord(gstrSQL, "判断库房性质", cboStock.ItemData(cboStock.ListIndex))
    Do While Not rsDetail.EOF
        str库房性质 = str库房性质 & "," & rsDetail!工作性质
        rsDetail.MoveNext
    Loop
    If str库房性质 Like "*中药*" Or str库房性质 Like "*制剂室*" Then bln中药库房 = True
        
    With vsfDetail
        .ColAlignment(.ColIndex("填写数量")) = flexAlignRightCenter     '填写数量
        .ColAlignment(.ColIndex("实际数量")) = flexAlignRightCenter     '实际数量
        .ColAlignment(.ColIndex("单位")) = flexAlignCenterCenter    '单位
        .ColAlignment(.ColIndex("成本价")) = flexAlignRightCenter     '成本价
        .ColAlignment(.ColIndex("成本金额")) = flexAlignRightCenter     '成本金额
        .ColAlignment(.ColIndex("售价")) = flexAlignRightCenter    '售价
        .ColAlignment(.ColIndex("售价金额")) = flexAlignRightCenter    '售价金额
        .ColAlignment(.ColIndex("差价")) = flexAlignRightCenter    '差价
                
        .ColWidth(0) = 0
        .ColWidth(.ColIndex("药品信息")) = 2500
        For intCol = 2 To .Cols - 1
            .ColWidth(intCol) = 1000
        Next
        If mblnViewCost = False Then
            .ColHidden(.ColIndex("成本价")) = True
            .ColHidden(.ColIndex("成本金额")) = True
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
    '基本，申领
    If Not IsHavePrivs(mstrPrivs, "申领") Or (Not IsHavePrivs(mstrPrivs, "手工申领") And Not IsHavePrivs(mstrPrivs, "自动申领") And Not IsHavePrivs(mstrPrivs, "按销量自动申领")) Then
        mnuEditAdd.Visible = False
        mnuEditModify.Visible = False
        mnuEditDel.Visible = False
        
        tlbTool.Buttons("Add").Visible = False
        tlbTool.Buttons("Modify").Visible = False
        tlbTool.Buttons("Delete").Visible = False
        tlbTool.Buttons("Edit1").Visible = False
        mnuEditLine1.Visible = False
    Else
        If Not IsHavePrivs(mstrPrivs, "手工申领") Then
            mnuEditAddHank.Visible = False
            tlbTool.Buttons("Add").ButtonMenus("Hank").Visible = False
        End If
        If Not IsHavePrivs(mstrPrivs, "自动申领") Then
            mnuEditAddAuto.Visible = False
            tlbTool.Buttons("Add").ButtonMenus("Auto").Visible = False
        End If
        If Not IsHavePrivs(mstrPrivs, "按销量自动申领") Then
            mnuEditAddAutoBySale.Visible = False
            tlbTool.Buttons("Add").ButtonMenus("Sale").Visible = False
        End If
    End If
    If Not IsHavePrivs(mstrPrivs, "审核") Then
        mnuEditReceive.Visible = False
        tlbTool.Buttons("Receive").Visible = False
    End If
    If Not IsHavePrivs(mstrPrivs, "冲销") Then
        mnuEditDisReceive.Visible = False
        mnuEditWriteOff.Visible = False
        If mnuEditReceive.Visible = False Then mnuEditLine2.Visible = False
        tlbTool.Buttons("DisReceive").Visible = False
        tlbTool.Buttons("EditSeparate").Visible = mnuEditLine2.Visible
        mnuEditWriteOff.Visible = False
'        tlbTool.Buttons("DisReceive").ButtonMenus(1).Visible = False
'        tlbTool.Buttons("DisReceive").ButtonMenus(2).Visible = False
    End If
    If Not IsHavePrivs(mstrPrivs, "单据打印") Then
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
    '恢复设置
    mblnViewCost = IsHavePrivs(mstrPrivs, "查看成本价")
    
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
    
    lbl1.Caption = ""
    lbl2.Caption = ""
    lbl3.Caption = ""
    lbl2.Left = lbl1.Left + lbl1.Width + 3000
    lbl3.Left = lbl2.Left + lbl2.Width + 3000
    If mblnViewCost = False Then
        lbl1.Visible = False
        lbl3.Visible = False
        lbl2.Left = lbl1.Left
    End If
    
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
        .Height = 360
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
    SaveWinState Me, App.ProductName, MStrCaption
End Sub

Private Sub mnuEditAddAuto_Click()
    '自动生成
    Dim strNo As String
    Dim BlnSuccess As Boolean
    
    '检查本期是否已经审核结存，如果未审核结存则不能进行新增业务操作
'    If CheckIsAccount(Val(cboStock.ItemData(cboStock.ListIndex))) = False Then
'        Exit Sub
'    End If
    
    strNo = ""
    frmRequestDrugCard.ShowCard Me, strNo, 5, , BlnSuccess, cboStock.ItemData(cboStock.ListIndex), 0, 1
    If BlnSuccess = True Then
        mnuViewRefresh_Click
    End If
    
End Sub

Private Sub mnuEditAddAutoBySale_Click()
    '自动按销售量自动生成
    Dim strNo As String
    Dim BlnSuccess As Boolean
    Dim rsTmp As ADODB.Recordset
    
    '检查本期是否已经审核结存，如果未审核结存则不能进行新增业务操作
'    If CheckIsAccount(Val(cboStock.ItemData(cboStock.ListIndex))) = False Then
'        Exit Sub
'    End If
    
    On Error GoTo errHandle
    '特殊的申领方式，需要检查是否都已审核
    gstrSQL = "Select 1 From 药品收发记录 " & _
        " Where 单据 = 6 And 库房id = [1] And 单量 = 7 And 审核日期 Is Null and 填制日期 between sysdate-60 and sysdate and rownum=1"
    Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, "mnuEditAddAutoBySale_Click", Val(cboStock.ItemData(cboStock.ListIndex)))
    
    If Not rsTmp.EOF Then
        MsgBox "还存在未审核的自动申领单据，不能产生新单据。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    strNo = ""
    frmRequestDrugCard.ShowCard Me, strNo, 5, , BlnSuccess, cboStock.ItemData(cboStock.ListIndex), 0, 7
    If BlnSuccess = True Then
        mnuViewRefresh_Click
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub mnuEditAddHank_Click()
    '手工填写
    Dim strNo As String
    Dim BlnSuccess As Boolean
    
    '检查本期是否已经审核结存，如果未审核结存则不能进行新增业务操作
'    If CheckIsAccount(Val(cboStock.ItemData(cboStock.ListIndex))) = False Then
'        Exit Sub
'    End If
    
    strNo = ""
    '新增
    frmRequestDrugCard.ShowCard Me, strNo, 1, , BlnSuccess, cboStock.ItemData(cboStock.ListIndex)
    If BlnSuccess = True Then
        mnuViewRefresh_Click
    End If
    
End Sub

Private Sub mnuEditAddMerge_Click()
    '合并申领单
    
'    frmRequestMerge.Show vbModal, Me
End Sub

Private Sub mnuEditDel_Click()
    '删除
    Dim StrBillNo As String
    Dim intRow As Integer
    Dim strTitle As String
    Dim intReturn As Integer
    Dim intRecord As Integer
    Dim strCheckString As String
    
    With vsfList
        On Error GoTo errHandle
        intRow = .Row
        StrBillNo = .TextMatrix(intRow, 0)
        
        If Not CheckNoIsExist(StrBillNo) Then
            MsgBox "没有找到该单据，可能已被删除！", vbInformation, gstrSysName
            
            '刷新
            GetList mstrFind
            Exit Sub
        End If
        
        '未审核单据
        If .TextMatrix(intRow, .Cols - 4) = "" And Val(.TextMatrix(.Row, .Cols - 3)) = 1 Then
            If Not Is申领(StrBillNo) Then
                MsgBox "你没有权限删除移库单！", vbInformation, gstrSysName
                Exit Sub
            End If
        
            strCheckString = CheckBill(Trim(StrBillNo))
            If strCheckString <> "" Then
                MsgBox strCheckString, vbInformation, gstrSysName
                GetList mstrFind
                Exit Sub
            End If
            
            strTitle = "药品申领单"
        ElseIf Val(.TextMatrix(.Row, .Cols - 3)) Mod 3 = 2 And mint冲销申请 = 1 Then
            '已审核单据，并且是申请冲销单据
            strTitle = "冲销申请单"
        End If
        
        intReturn = MsgBox("你确实要删除单据号为“" & StrBillNo & "”的" & strTitle & "吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
        intRecord = .rows - 1
        If intReturn = vbYes Then
            gstrSQL = "zl_药品移库_Delete('" & StrBillNo & "'," & Val(.TextMatrix(.Row, .Cols - 3)) & " )"
            If gstrSQL = "" Then Exit Sub
            Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption & "-删除申领单")
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
        frmRequestDrugCard.ShowCard Me, strNo, 4, .TextMatrix(.Row, .Cols - 3), , cboStock.ItemData(cboStock.ListIndex)
    End With
End Sub


'Modified By 朱玉宝 2003-12-10 地区：泸州
Private Sub mnuEditDisReceive_Click()
    Dim strNo As String, BlnSuccess As Boolean
    Dim int处理方式 As Integer
    
    If mnuEditDisReceive.Caption = "申请冲销(&R)" Then
        int处理方式 = 1
    Else
        int处理方式 = 0
    End If
                
    With vsfList
        strNo = .TextMatrix(.Row, 0)
        frmRequestDrugCard.ShowCard Me, strNo, 7, .TextMatrix(.Row, .Cols - 3), BlnSuccess, cboStock.ItemData(cboStock.ListIndex), int处理方式
        If Not BlnSuccess Then Exit Sub
    End With
    Call mnuViewRefresh_Click
End Sub

Private Sub mnuEditModify_Click()
    '修改
    Dim strNo As String
    Dim BlnSuccess As Boolean
    
    BlnSuccess = False
    With vsfList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        strNo = .TextMatrix(.Row, 0)
        
        If Not CheckNoIsExist(strNo) Then
            MsgBox "没有找到该单据，可能已被删除！", vbInformation, gstrSysName
            
            '刷新
            GetList mstrFind
            Exit Sub
        End If
        
        If Not Is申领(strNo) Then
            MsgBox "你没有权限修改移库单！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        frmRequestDrugCard.ShowCard Me, strNo, 2, vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 3), BlnSuccess, cboStock.ItemData(cboStock.ListIndex)
        If BlnSuccess = True Then
            mnuViewRefresh_Click
        End If
    End With
End Sub

'Modified By 朱玉宝 2003-12-10 地区：泸州
Private Sub mnuEditReceive_Click()
    Dim strNo As String, BlnSuccess As Boolean
    With vsfList
        strNo = .TextMatrix(.Row, 0)
        frmRequestDrugCard.ShowCard Me, strNo, 6, .TextMatrix(.Row, .Cols - 3), BlnSuccess, cboStock.ItemData(cboStock.ListIndex)
    End With
    If BlnSuccess = True Then
        mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditWriteOff_Click()
    Dim strStock As String
    Dim i As Integer
    
    With Me.cboStock
        For i = 0 To .ListCount - 1
            strStock = strStock & .List(i) & "," & .ItemData(i) & "|"
        Next
       
    End With
    
    Call frm批量冲销.ShowMe(1341, Me, strStock, Me.cboStock.ListIndex)
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
    Dim rstemp As New ADODB.Recordset

    On Error GoTo errHandle
    
    If Not IsHavePrivs(mstrPrivs, "药品条码打印") Then
        MsgBox "对不起，你没有该权限！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If TypeName(varPar) = "String" Then '打印整张单据条码
        gstrSQL = "select distinct 药品ID from 药品收发记录 where 单据 = 6 and  NO = [1] order by 药品ID"
        
        Set rstemp = zlDataBase.OpenSQLRecord(gstrSQL, "药品条码打印", varPar)
        
        Do While Not rstemp.EOF
            ReportOpen gcnOracle, glngSys, "ZL1_INSIDE_1343_1", Me, "药品=" & rstemp!药品ID, 2
            rstemp.MoveNext
        Loop
        
    Else '打印对应药品条码
        If varPar = 0 Then Exit Sub
        ReportOpen gcnOracle, glngSys, "ZL1_INSIDE_1343_1", Me, "药品=" & varPar, 2
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
        ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1304", "zl8_bill_1304"), Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & .TextMatrix(.Row, .Cols - 3), "单位系数=" & int单位系数, 1
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
        ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1304", "zl8_bill_1304"), Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & .TextMatrix(.Row, .Cols - 3), "单位系数=" & int单位系数, 2
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
    Dim dateCurDate As Date
    
    '参数设置
    frm参数设置.设置参数 Me, mstrPrivs, 1343, MStrCaption
    
    Call GetDrugDigit(mlng库房id, MStrCaption, mintUnit, mintShowCostDigit, mintShowPriceDigit, mintShowNumberDigit, mintShowMoneyDigit)
    
    '重新组织格式化串
    mstrCostFormat = "'999999999990." & String(mintShowCostDigit, "0") & "'"
    mstrPriceFormat = "'999999999990." & String(mintShowPriceDigit, "0") & "'"
    mstrNumberFormat = "'999999999990." & String(mintShowNumberDigit, "0") & "'"
    mstrMoneyFormat = "'999999999990." & String(mintShowMoneyDigit, "0") & "'"
    
    dateCurDate = Sys.Currentdate
    mint查询天数 = Val(zlDataBase.GetPara("查询天数", glngSys, 1343)) - 1
'    mint查询天数 = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & MStrCaption, "查询天数", "7")
    strStart = Format(DateAdd("d", -1 * mint查询天数, dateCurDate), "yyyy-MM-dd")
    strEnd = Format(dateCurDate, "yyyy-MM-dd")
    mstrFind = " AND A.记录状态 = 1 And A.审核日期 is Null And A.填制日期 Between [3] And [4] "
    SQLCondition.date填制时间开始 = CDate(Format(strStart, "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date填制时间结束 = CDate(Format(strEnd, "yyyy-mm-dd") & " 23:59:59")
    
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
    '默认参数：药品=药品id，药房=药房id，库房=库房id，开始时间=填制开始时间，结束时间=填制结束时间，NO=申领单NO
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
        "药房=" & IIf(Val(cboStock.ItemData(cboStock.ListIndex)) = 0, "", Val(cboStock.ItemData(cboStock.ListIndex))), _
        "库房=" & IIf(SQLCondition.lng移出库房 = 0, "", SQLCondition.lng移出库房), _
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
    
    strFind = FrmListSearch.GetSearch(Me, 1343, cboStock.ItemData(cboStock.ListIndex), strStart, strEnd, strVerifyStart, strVerifyEnd, _
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
        cbrTool.Visible = .Checked
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
    Dim IntBill As Integer                      '单据类型  如：1、外购入库；2、
    Dim strUnit As String                       '单位名称:如门诊单位，住院单位等
    Dim str包装系数 As String
    Dim strOrder As String
    Dim strCompare As String
    Dim strSql效期 As String
    Dim strSql药名 As String
    Dim intCol As Integer
    Dim strSqlOrder As String
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
    
    strOrder = zlDataBase.GetPara("排序", glngSys, 1343)
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
                strUnit = "D.计算单位"
                str包装系数 = "1"
            Case mconint门诊单位
                strUnit = "B.门诊单位"
                str包装系数 = "B.门诊包装"
            Case mconint住院单位
                strUnit = "B.住院单位"
                str包装系数 = "B.住院包装"
            Case mconint药库单位
                strUnit = "B.药库单位"
                str包装系数 = "B.药库包装"
        End Select
        
        strSql效期 = IIf(gtype_UserSysParms.P149_效期显示方式 = 1, "TO_CHAR(A.效期-1,'YYYY-MM-DD') AS 有效期至", "TO_CHAR(A.效期,'YYYY-MM-DD') AS 失效期")
        
        If gint药品名称显示 = 0 Then
            strSql药名 = ",('['||D.编码||']'||D.名称) AS 药品信息"
        ElseIf gint药品名称显示 = 1 Then
            strSql药名 = ",('['||D.编码||']'||NVL(E.名称,D.名称)) AS 药品信息"
        Else
            strSql药名 = ",('['||D.编码||']'||D.名称) AS 药品信息,E.名称 As 商品名"
        End If
        
        gstrSQL = " SELECT * FROM (SELECT DISTINCT 序号" & strSql药名 & ",B.药品来源,B.基本药物," & _
            " D.规格,A.产地 as 生产商,A.原产地, A.批号, " & strSql效期 & " ,A.批准文号,LTRIM(TO_CHAR(A.填写数量 /" & str包装系数 & "," & mstrNumberFormat & " )) AS 填写数量," & _
            " LTRIM(TO_CHAR(A.实际数量 /" & str包装系数 & "," & mstrNumberFormat & ")) AS 实际数量," & strUnit & " AS 单位," & _
            " LTRIM(TO_CHAR (A.成本价*" & str包装系数 & ", " & mstrCostFormat & ")) AS 成本价," & _
            " LTRIM(TO_CHAR (A.成本金额, " & mstrMoneyFormat & ")) AS 成本金额," & _
            " LTRIM(TO_CHAR (A.零售价*" & str包装系数 & ", " & mstrPriceFormat & ")) AS 售价," & _
            " LTRIM(TO_CHAR (A.零售金额, " & mstrMoneyFormat & ")) AS 售价金额," & _
            " LTRIM(TO_CHAR (A.差价, " & mstrMoneyFormat & ")) AS 差价 ,C.库房货位,D.ID 药品ID " & _
            " FROM 药品收发记录 A, 药品规格 B, 收费项目别名 E, 收费项目目录 D, 药品储备限额 C " & _
            " WHERE A.药品ID = B.药品ID AND B.药品ID=D.ID " & _
            " AND B.药品ID = E.收费细目ID(+) AND E.性质(+)=3 " & _
            " AND A.记录状态 = [2] " & _
            " AND A.单据 = 6 AND 入出系数=1 " & _
            " AND A.NO = [1] AND A.药品ID=C.药品ID(+) AND A.库房ID=C.库房ID(+))" & _
            " ORDER BY " & strSqlOrder
        Set rsDetail = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, vsfList.TextMatrix(vsfList.Row, 0), Val(vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 3)))
        
        Set vsfDetail.DataSource = rsDetail
        rsDetail.Close

        With vsfDetail
            .Row = 1
            .Col = 0
            '.ColSel = .Cols - 1    ' bug: 40410
            
            For n = 0 To .Cols - 1
                .ColKey(n) = .TextMatrix(0, n)
                .FixedAlignment(n) = flexAlignCenterCenter
            Next
            
            .ColHidden(.ColIndex("药品ID")) = True '药品ID列不显示
        End With
        vsfDetail.Redraw = flexRDDirect
    ElseIf LTrim(vsfList.TextMatrix(vsfList.Row, 0)) = "" Then
        With vsfDetail
            .Redraw = flexRDNone
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
            .TextMatrix(0, intCol) = "批准文号": intCol = intCol + 1
            .TextMatrix(0, intCol) = "填写数量": intCol = intCol + 1
            .TextMatrix(0, intCol) = "实际数量": intCol = intCol + 1
            .TextMatrix(0, intCol) = "单位": intCol = intCol + 1
            .TextMatrix(0, intCol) = "成本价": intCol = intCol + 1
            .TextMatrix(0, intCol) = "成本金额": intCol = intCol + 1
            .TextMatrix(0, intCol) = "售价": intCol = intCol + 1
            .TextMatrix(0, intCol) = "售价金额": intCol = intCol + 1
            .TextMatrix(0, intCol) = "差价": intCol = intCol + 1
            .TextMatrix(0, intCol) = "库房货位": intCol = intCol + 1
            
            For n = 0 To .Cols - 1
                .ColKey(n) = .TextMatrix(0, n)
                .FixedAlignment(n) = flexAlignCenterCenter
            Next
            
            .Redraw = flexRDDirect
        End With
    End If
    SetDetailColWidth
    CheckNumber
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
    
    mnuEditAllCodePrint.Visible = True
    mnuEditCodePrintLine.Visible = True
    PopupMenu mnuEdit, 2
    mnuEditAllCodePrint.Visible = False
    mnuEditCodePrintLine.Visible = False
    
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
            mnuEditAddHank_Click
        Case "Modify"
            mnuEditModify_Click
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
    Dim rstemp As New ADODB.Recordset
    
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
            
            If .TextMatrix(.Row, .Cols - 4) = "" Then    '未审核单
                bln已发送 = (vsfList.TextMatrix(vsfList.Row, 12) <> "")
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
                If mnuEditReceive.Visible Then
                    mnuEditReceive.Enabled = bln已发送
                    tlbTool.Buttons("Receive").Enabled = bln已发送
                End If
                
                mnuEditDisReceive.Enabled = False
                tlbTool.Buttons("DisReceive").Enabled = False
                
                '如果冲销单还未审核，则允许删除
                If mint冲销申请 = 1 Then
                    If Val(.TextMatrix(.Row, .Cols - 3)) Mod 3 = 2 Then
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
                        
            ElseIf .TextMatrix(.Row, .Cols - 3) = 1 Then    '审核单
                '判断是否接受（不支持已冲销单据的接受功能，必需全退或输负数的方式解决，因为要实现这个功能，需要汇总统计剩余数量）
                
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
                If .TextMatrix(.Row, .Cols - 3) Mod 3 = 0 Then
                    .ToolTipText = "冲销单据的原单据"
                    If mnuEditDisReceive.Visible = True Then
                        mnuEditDisReceive.Enabled = True
                        tlbTool.Buttons("DisReceive").Enabled = True
                    End If
                ElseIf .TextMatrix(.Row, .Cols - 3) Mod 3 = 2 Then
                    .ToolTipText = "冲销单据"
                    If mnuEditDisReceive.Visible = True Then
                        mnuEditDisReceive.Enabled = False
                        tlbTool.Buttons("DisReceive").Enabled = False
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
    
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
        
    objRow.Add "打印人:" & gstrUserName
    objRow.Add "打印日期:" & Format(zlDataBase.Currentdate, "yyyy年MM月dd日")
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

Private Sub tlbTool_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.index = 1 Then
        mnuEditAddHank_Click
    ElseIf ButtonMenu.index = 2 Then
        mnuEditAddAuto_Click
    ElseIf ButtonMenu.index = 3 Then
        mnuEditAddAutoBySale_Click
    ElseIf ButtonMenu.index = 4 Then
        mnuEditAddMerge_Click
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
    
    objPrint.Title.Text = MStrCaption
        
    Set objRow = New zlTabAppRow
    objRow.Add ""
    objRow.Add "NO." & Trim(vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "NO")))
    objPrint.UnderAppRows.Add objRow

    Set objRow = New zlTabAppRow
    objRow.Add "移出库房：" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "发药库房"))
    objRow.Add "移入库房：" & gstrDeptName
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

Private Function Is申领(ByVal StrBillNo As String) As Boolean
    Dim rsCheck As New ADODB.Recordset
    '先检查是不是申领单
    On Error GoTo errHandle
    gstrSQL = " Select Nvl(发药方式,0) 申领 From 药品收发记录 " & _
              " Where 单据=6 And NO=[1] And 入出系数 = -1 and rownum = 1"
    Set rsCheck = zlDataBase.OpenSQLRecord(gstrSQL, "检查是不是申领单", StrBillNo)
    
    Is申领 = Not (rsCheck!申领 = 0)
    Exit Function
errHandle:
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
    Call zlWebForum(Me.hWnd)
End Sub

Private Sub CheckNumber()
    '如果填写数量和实际数量不一致，则用红色字体标注实际数量用以提醒
    Dim intRow As Integer
    Dim blnColor As Boolean

    With vsfDetail
        If .rows <= 2 Then Exit Sub
        For intRow = 1 To .rows - 1
            blnColor = False
            If .TextMatrix(intRow, .ColIndex("药品ID")) = "" Then Exit Sub
            If Val(.TextMatrix(intRow, .ColIndex("填写数量"))) <> Val(.TextMatrix(intRow, .ColIndex("实际数量"))) Then blnColor = True
            .Cell(flexcpForeColor, intRow, .ColIndex("实际数量"), intRow, .ColIndex("实际数量")) = IIf(blnColor, vbRed, vbBlack)
        Next
    End With
                
End Sub

