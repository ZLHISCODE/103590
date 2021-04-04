VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSchedule 
   Caption         =   "体检预约申请"
   ClientHeight    =   7140
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11310
   Icon            =   "frmSchedule.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7140
   ScaleWidth      =   11310
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6780
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmSchedule.frx":1CFA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14870
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
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   11310
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinWidth1       =   4995
      MinHeight1      =   720
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Caption2        =   "体检部门"
      Child2          =   "cboDept"
      MinWidth2       =   2100
      MinHeight2      =   300
      Width2          =   465
      NewRow2         =   0   'False
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   9120
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   2100
      End
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   720
         Left            =   165
         TabIndex        =   2
         Top             =   30
         Width           =   7950
         _ExtentX        =   14023
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ilsMenu"
         HotImageList    =   "ilsMenuHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   14
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
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_1"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预约"
               Key             =   "预约"
               Object.ToolTipText     =   "预约"
               Object.Tag             =   "预约"
               ImageIndex      =   3
               Style           =   5
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修改"
               Key             =   "修改"
               Object.ToolTipText     =   "修改"
               Object.Tag             =   "修改"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "删除"
               Object.ToolTipText     =   "删除"
               Object.Tag             =   "删除"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_2"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "确认"
               Key             =   "确认"
               Object.ToolTipText     =   "确认"
               Object.Tag             =   "确认"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "取消"
               Key             =   "取消"
               Object.ToolTipText     =   "取消"
               Object.Tag             =   "取消"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_3"
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "过滤"
               Key             =   "过滤"
               Object.ToolTipText     =   "过滤"
               Object.Tag             =   "过滤"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_4"
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "帮助"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   10
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   8760
      Top             =   900
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
            Picture         =   "frmSchedule.frx":258E
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":27AE
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":29CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":2BE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":2E02
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":301C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":3236
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":3450
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":366A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":388A
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":3AAA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsMenuHot 
      Left            =   8040
      Top             =   900
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
            Picture         =   "frmSchedule.frx":3CC4
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":3EE4
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":4104
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":4456
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":4670
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":49C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":4BDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":4DF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":5010
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":5230
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":5450
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   1770
      Left            =   150
      TabIndex        =   3
      Top             =   900
      Width           =   2775
      _cx             =   4895
      _cy             =   3122
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
      BackColorSel    =   16772055
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   6
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
      Begin VB.Line lnY 
         Index           =   0
         Visible         =   0   'False
         X1              =   270
         X2              =   270
         Y1              =   420
         Y2              =   1635
      End
      Begin VB.Line lnX 
         Index           =   0
         Visible         =   0   'False
         X1              =   -555
         X2              =   1230
         Y1              =   555
         Y2              =   555
      End
   End
   Begin MSComctlLib.ImageList ils13 
      Left            =   4560
      Top             =   2055
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":5962
            Key             =   "状态"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":5CFC
            Key             =   "个人"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":6296
            Key             =   "团体"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":B300
            Key             =   "确认"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":B89A
            Key             =   "取消"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":BE34
            Key             =   "开始"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":C3CE
            Key             =   "新开"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":C968
            Key             =   "完成"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":CF02
            Key             =   "up"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":D0C4
            Key             =   "down"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid vsfPrint 
      Height          =   780
      Left            =   5475
      TabIndex        =   5
      Top             =   1980
      Visible         =   0   'False
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   1376
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   270
      MergeCells      =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.ImageList ilsGrid 
      Left            =   6315
      Top             =   3165
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":D286
            Key             =   "T附加"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":D620
            Key             =   "报告"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":D9BA
            Key             =   "状态"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":DD54
            Key             =   "单据"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":E0EE
            Key             =   "附加"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":E488
            Key             =   "up"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":E64A
            Key             =   "down"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfPerson 
      Height          =   1950
      Left            =   390
      TabIndex        =   6
      Top             =   3495
      Width           =   2985
      _cx             =   5265
      _cy             =   3440
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
      BackColorSel    =   16772055
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   6
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
      MergeCells      =   1
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
      Begin VB.Line lnY1 
         Index           =   0
         Visible         =   0   'False
         X1              =   900
         X2              =   900
         Y1              =   810
         Y2              =   2025
      End
      Begin VB.Line lnX1 
         Index           =   0
         Visible         =   0   'False
         X1              =   75
         X2              =   1860
         Y1              =   945
         Y2              =   945
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfItem 
      Height          =   1950
      Left            =   3615
      TabIndex        =   7
      Top             =   3495
      Width           =   2505
      _cx             =   4419
      _cy             =   3440
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
      BackColorSel    =   16772055
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   6
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
      MergeCells      =   1
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
      Begin VB.Line lnX2 
         Index           =   0
         Visible         =   0   'False
         X1              =   75
         X2              =   1860
         Y1              =   945
         Y2              =   945
      End
      Begin VB.Line lnY2 
         Index           =   0
         Visible         =   0   'False
         X1              =   900
         X2              =   900
         Y1              =   810
         Y2              =   2025
      End
   End
   Begin VB.Image imgY_S 
      Height          =   4395
      Left            =   3450
      MousePointer    =   9  'Size W E
      Top             =   2835
      Width           =   45
   End
   Begin VB.Image imgX_S 
      Height          =   45
      Left            =   255
      MousePointer    =   7  'Size N S
      Top             =   3135
      Width           =   3690
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrintView 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileOutExcel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileNew 
         Caption         =   "生成登记表格(&N)"
      End
      Begin VB.Menu mnuFile_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "预约(&E)"
      Begin VB.Menu mnuEditAdd 
         Caption         =   "增加个人预约(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditAddGroup 
         Caption         =   "增加团体预约(&N)"
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "修改预约(&M)"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "删除预约(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEdit_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCheck 
         Caption         =   "确认预约(&O)"
      End
      Begin VB.Menu mnuEditCancel 
         Caption         =   "取消预约(&C)"
      End
   End
   Begin VB.Menu mnuAddition 
      Caption         =   "附加(&A)"
      Begin VB.Menu mnuAdditionItems 
         Caption         =   "组别项目(&I)"
      End
      Begin VB.Menu mnuAdditionPerItems 
         Caption         =   "人员项目(&H)"
      End
      Begin VB.Menu mnuAddition_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdditionPersons 
         Caption         =   "受检人员(&P)"
      End
      Begin VB.Menu mnuAddition_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdditionGroupMember 
         Caption         =   "人员划分(&C)"
      End
      Begin VB.Menu mnuAdditionPhoto 
         Caption         =   "照片采集(&S)"
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
      Begin VB.Menu mnuView_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewSearch 
         Caption         =   "过滤(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "刷新(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpTopic 
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
      Begin VB.Menu h1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
End
Attribute VB_Name = "frmSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'（１）窗体级变量定义**************************************************************************************************
Private mblnStartUp As Boolean                          '窗体启动标志
Private mstrPrivs As String

Private mblnAllowChange As Boolean
Private mstrCondition As String
Private mlngSvrDept As Long                             '保存上次点击的体检部门
Private mstrSvrGoup As String                           '保存上次点击的体检组别
Private mlngSvrKey As Long                              '保存上次点击的体检预约

Private WithEvents mobjPopMenu As clsPopMenu                '自定义弹出菜单对象
Attribute mobjPopMenu.VB_VarHelpID = -1
Private mbytPopMenu As Byte

Private Enum mCol
    i名称 = 1
    i检查部位
    i采集方式
    i检验标本
    i基本价格
    i体检价格
    i执行科室
    i结算方式
End Enum

'（２）自定义过程或函数************************************************************************************************

Private Function InitLoad() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:初始化数据，发生在窗体的Load事件
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim strVsf As String
    
    On Error GoTo errHand
    
    Call InitSysPara
    
    mlngSvrDept = 0
    mblnAllowChange = True
    mstrSvrGoup = ""
    mlngSvrKey = 0
    imgY_S.Width = 60
    
    mstrCondition = Format(zlDatabase.Currentdate, "yyyy-MM-dd") & "'" & Format(DateAdd("d", 7, CDate(zlDatabase.Currentdate)), "yyyy-MM-dd")
    mstrCondition = mstrCondition & "''''''"
    
    strVsf = ",255,4,1,1,[性质];,255,4,1,1,[状态];No,810,1,1,1,;预约人,750,1,1,1,;预约时间,990,1,1,1,;团体,2400,1,1,1,;人数,450,1,1,1,;应收金额,900,7,1,1,;实收金额,900,7,1,1,"
    strVsf = strVsf & ";体检类型,1800,1,1,1,;联系电话,900,1,1,1,;联系地址,1800,1,1,1,;附加说明,1500,1,1,1,;体检状态,0,1,1,1,;合约单位id,0,1,1,1,"
    
    If Val(GetSetting("ZLSOFT", "私有全局\" & gstrDBUser, "使用个性化风格", "0")) = 1 Then
    
        '使用个性化设置
                        
    End If
    
    Call CreateVsf(vsf, strVsf)
'    vsf.Cols = vsf.Cols + 1
'    vsf.ColWidth(vsf.Cols - 1) = 15
    
    Set vsf.Cell(flexcpPicture, 0, 0) = ils13.ListImages("状态").Picture
    Set vsf.Cell(flexcpPicture, 0, 1) = ils13.ListImages("状态").Picture
    
    vsf.ColFormat(GetCol(vsf, "应收金额")) = gstrDec
    vsf.ColFormat(GetCol(vsf, "实收金额")) = gstrDec
    
    strVsf = "姓名,900,1,1,1,;门诊号,900,7,1,1,;性别,810,1,1,1,;年龄,600,1,1,1,;病人id,0,1,1,1,"
    Call CreateVsf(vsfPerson, strVsf)
    With vsfPerson
'        .Cols = .Cols + 1
'        .ColWidth(vsfPerson.Cols - 1) = 15
        .MergeCells = flexMergeFree
        .OutlineCol = 0
        .OutlineBar = flexOutlineBarComplete
    End With
        
    strVsf = ",255,4,1,1,[附加];名称,2400,1,1,1,;检查部位,1200,1,1,1,;采集方式,900,1,1,1,;检验标本,900,1,1,1,;基本价格,900,7,1,1,;体检价格,900,7,1,1,;执行科室,1200,1,1,1,;结算方式,810,1,1,1,"
    Call CreateVsf(vsfItem, strVsf)
'    vsfItem.Cols = vsfItem.Cols + 1
'    vsfItem.ColWidth(vsfItem.Cols - 1) = 15

    vsfItem.ColFormat(GetCol(vsfItem, "体检价格")) = gstrDec
    vsfItem.ColFormat(GetCol(vsfItem, "基本价格")) = gstrDec
    Set vsfItem.Cell(flexcpPicture, 0, 0) = ilsGrid.ListImages("T附加").Picture
    
    InitLoad = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function InitActivate() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  初始化数据，发生在窗体的Activate事件
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    gstrSQL = GetPublicSQL(SQL.体检部门清单, IIf(InStr(gstrPrivs, "所有科室") > 0, "所有", ""))
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, UserInfo.ID)
    
    If rs.BOF Then
        ShowSimpleMsg "没有体检性质的部门，请在部门管理中设置！"
        Exit Function
    End If
    
    '绑定数据到控件中
    Call AddComboData(cboDept, rs)
    
    '初始选择数据处理
    zlControl.CboLocate cboDept, UserInfo.部门ID, True
    If cboDept.ListCount > 0 And cboDept.ListIndex = -1 Then cboDept.ListIndex = 0
    
    InitActivate = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub ApplyPrivilege(ByVal strPrivilege As String)
    '------------------------------------------------------------------------------------------------------------------
    '功能： 应用权限处理
    '参数： strPrivilege                    权限
    '------------------------------------------------------------------------------------------------------------------
        
    '调试语句
    'strPrivilege = "所有科室;体检预约;确认预约;取消预约"
    
    '不具有“预约”权限时
    If InStr(strPrivilege, "体检预约") = 0 Then
        mnuEditAdd.Visible = False
        mnuEditAddGroup.Visible = False
        mnuEditModify.Visible = False
        mnuEditDelete.Visible = False
                
        mnuAddition.Visible = False
    End If
    
    If InStr(strPrivilege, "确认预约") = 0 Then mnuEditCheck.Visible = False
    
    If InStr(strPrivilege, "取消预约") = 0 Then
        If InStr(strPrivilege, "确认预约") = 0 And InStr(strPrivilege, "体检预约") = 0 Then
            mnuEdit.Visible = False
        Else
            mnuEditCancel.Visible = False
        End If
    End If
    
    mnuEdit_1.Visible = mnuEditAdd.Visible And (mnuEditCheck.Visible Or mnuEditCancel.Visible)
    
    tbrThis.Buttons("预约").Visible = mnuEdit.Visible And mnuEditAdd.Visible
    tbrThis.Buttons("修改").Visible = mnuEdit.Visible And mnuEditModify.Visible
    tbrThis.Buttons("删除").Visible = mnuEdit.Visible And mnuEditDelete.Visible
    tbrThis.Buttons("确认").Visible = mnuEdit.Visible And mnuEditCheck.Visible
    tbrThis.Buttons("取消").Visible = mnuEdit.Visible And mnuEditCancel.Visible
    
    tbrThis.Buttons("Split_2").Visible = tbrThis.Buttons("预约").Visible
    tbrThis.Buttons("Split_3").Visible = tbrThis.Buttons("确认").Visible Or tbrThis.Buttons("取消").Visible
    
End Sub

Private Sub AdjustEnableState()
    '------------------------------------------------------------------------------------------------------------------
    '功能： 调整各功能菜单的可用状态
    '------------------------------------------------------------------------------------------------------------------
    mnuFilePrint.Enabled = True
    mnuFilePrintView.Enabled = True
    mnuFileOutExcel.Enabled = True
    
    mnuEditModify.Enabled = True
    mnuEditDelete.Enabled = True
    
    mnuEditCheck.Enabled = True
    mnuEditCancel.Enabled = True

    mnuAdditionGroupMember.Enabled = True
    mnuAdditionItems.Enabled = True
    mnuAdditionPerItems.Enabled = True
    mnuAdditionPersons.Enabled = True
            
    If Val(vsf.RowData(1)) = 0 Then
                
        mnuFilePrint.Enabled = False
        mnuFilePrintView.Enabled = False
        mnuFileOutExcel.Enabled = False
    
        mnuEditModify.Enabled = False
        mnuEditDelete.Enabled = False
        mnuEditCheck.Enabled = False
        mnuEditCancel.Enabled = False
        
        mnuAdditionGroupMember.Enabled = False
        mnuAdditionItems.Enabled = False
        mnuAdditionPerItems.Enabled = False
        mnuAdditionPersons.Enabled = False
            
    Else
        Select Case Val(vsf.TextMatrix(vsf.Row, GetCol(vsf, "体检状态")))
        Case 1, 3         '新开预约
        
            mnuEditCancel.Enabled = False
            
        Case 2          '确认预约
        
            mnuEditCheck.Enabled = False
            mnuEditModify.Enabled = False
            mnuEditDelete.Enabled = False
            
        Case 4, 5
        
            mnuEditCheck.Enabled = False
            mnuEditCancel.Enabled = False
            mnuEditModify.Enabled = False
            mnuEditDelete.Enabled = False
            
            mnuAdditionGroupMember.Enabled = False
            mnuAdditionItems.Enabled = False
            mnuAdditionPerItems.Enabled = False
            mnuAdditionPersons.Enabled = False
            
        End Select
        
        If vsf.TextMatrix(vsf.Row, GetCol(vsf, "[性质]")) <> "" Then
           
           '是个人预约
           mnuAdditionGroupMember.Enabled = False
           mnuAdditionItems.Enabled = False
           mnuAdditionPersons.Enabled = False
           
        Else
            
            If Val(vsfPerson.TextMatrix(vsfPerson.Row, GetCol(vsfPerson, "门诊号"))) = 0 Then
                mnuAdditionPerItems.Enabled = False
            End If
            
        End If
        
    End If
    
    tbrThis.Buttons("预览").Enabled = mnuFilePrintView.Enabled
    tbrThis.Buttons("打印").Enabled = mnuFilePrint.Enabled
    
    tbrThis.Buttons("修改").Enabled = mnuEditModify.Enabled
    tbrThis.Buttons("删除").Enabled = mnuEditDelete.Enabled
    tbrThis.Buttons("确认").Enabled = mnuEditCheck.Enabled
    tbrThis.Buttons("取消").Enabled = mnuEditCancel.Enabled
    
End Sub

Private Sub RefreshStateInfo()
    '------------------------------------------------------------------------------------------------------------------
    '功能：刷新状态栏显示信息
    '------------------------------------------------------------------------------------------------------------------
    
    If Val(vsf.RowData(1)) <= 0 Then
        stbThis.Panels(2).Text = "“" & cboDept.Text & "”下没有体检预约。"
    Else
        stbThis.Panels(2).Text = "“" & cboDept.Text & "”下共有 " & vsf.Rows - 1 & "个体检预约。"
    End If
    
End Sub

Private Function ClearData(Optional ByVal strMenuItem As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '------------------------------------------------------------------------------------------------------------------
    
    strMenuItem = ";" & strMenuItem & ";"
    
    If InStr(strMenuItem, ";体检预约;") > 0 Then
        Call ResetVsf(vsf)
        
'        Call AppendSapceRows(vsf, lnX, lnY)
    End If
    
    If InStr(strMenuItem, ";体检项目;") > 0 Then
        Call ResetVsf(vsfItem)
        
        Call AppendSapceRows(vsfItem, lnX1, lnY1)
    End If
        
End Function

Public Function EditRefresh(ByVal strMenuItem As String, ByVal strPara As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：供编辑数据窗体调用，接口函数
    '------------------------------------------------------------------------------------------------------------------
    Dim lngSvrKey2 As Long
    Dim lngSvrKey3 As Long
    Dim varPara As Variant
        
    On Error GoTo errHand

    '保存体检组别、体检项目
    varPara = Split(strPara, "'")
    
    Select Case strMenuItem
    Case "体检预约"
        Call ClearData("体检预约;体检项目")
        
        Call RefreshData("体检预约")
        
        '恢复体检预约
        Call RestoreRow(vsf, Val(varPara(0)))
        
    Case "体检人员"
        
    Case Else
        Call ClearData("体检项目")
    End Select
    
    mblnAllowChange = True
    
    Call AdjustEnableState
    Call RefreshStateInfo
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function


Private Function RefreshData(Optional ByVal strMenuItem As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：刷新/装载数据
    '------------------------------------------------------------------------------------------------------------------
    Dim lngKey As Long
    Dim rs As New ADODB.Recordset
    Dim varCondition As Variant
    Dim strCondition As String
    Dim lngLoop As Long
    
    Dim intState As Integer
    Dim intGroup As Integer
    
    On Error GoTo errHand
       
    Select Case strMenuItem
    Case "体检预约"
                                
        varCondition = Split(mstrCondition, "'")
        strCondition = " AND A.体检时间 BETWEEN [2] AND [3] "
        
        If Trim(varCondition(2)) <> "" Then strCondition = strCondition & " AND A.联系人 LIKE [4] "
        If Trim(varCondition(3)) <> "" Then strCondition = strCondition & " AND A.体检号=[5] "
        
        If Val(varCondition(5)) > 0 Then strCondition = strCondition & " AND A.合约单位ID=[6] "
        
        strCondition = strCondition & " AND A.体检状态<=[7]"
        
        If Val(varCondition(6)) > 0 Then
            intState = 3
        Else
            intState = 1
        End If

        If Val(varCondition(7)) = 1 Then
            intGroup = 1
            strCondition = strCondition & " AND NVL(A.是否团体,0)=[8]"
        ElseIf Val(varCondition(7)) = 2 Then
            intGroup = 0
            strCondition = strCondition & " AND NVL(A.是否团体,0)=[8]"
        End If
                                        
        gstrSQL = GetPublicSQL(SQL.体检预约单据, strCondition)
        
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngSvrDept, CDate(varCondition(0)), CDate(varCondition(1)) + 1 - 1 / 24 / 60 / 60, "%" & CStr(varCondition(2)) & "%", CStr(varCondition(3)), Val(varCondition(5)), intState, intGroup)

        If rs.BOF = False Then Call LoadGrid(vsf, rs, , , ils13)
        
    Case "体检人员"
        
        vsfPerson.RowHidden(1) = False
        
        gstrSQL = "Select '' As 体检时间,0 AS 次数,0 As ID,0 As 病人id,组别名称 As 姓名,0 AS 门诊号,'' AS 性别,'' AS 年龄,组别名称 " & _
            "From 体检组别 Where 登记id=[1]"
        
        gstrSQL = "Select * From (" & gstrSQL & " Union All " & _
            "Select TO_CHAR(B.体检时间,'yyyy-mm-dd') As 体检时间,B.次数,A.病人id AS ID,A.病人id,A.姓名,A.门诊号,A.性别,A.年龄,B.组别名称 " & _
            "from 病人信息 A,体检人员档案 B  " & _
            "WHERE A.病人ID=B.病人ID AND B.登记id=[1]) Order By 组别名称,门诊号 "
        
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(vsf.RowData(vsf.Row)))
        
        If rs.BOF = False Then
        
            Call FillGrid(vsfPerson, rs)
            
            With vsfPerson
                For lngLoop = 1 To .Rows - 1
                    If Val(.TextMatrix(lngLoop, GetCol(vsfPerson, "门诊号"))) = 0 Then
                        .MergeRow(lngLoop) = True
                        .Cell(flexcpText, lngLoop, 0, lngLoop, .Cols - 2) = .TextMatrix(lngLoop, 0)
                        .Cell(flexcpFontBold, lngLoop, 0, lngLoop, .Cols - 2) = True
                        .RowOutlineLevel(lngLoop) = 1
                        .IsSubtotal(lngLoop) = True
                    End If
                Next
                
                If vsf.TextMatrix(vsf.Row, GetCol(vsf, "[性质]")) <> "" Then
                    .RowHidden(1) = True
                    .Row = 2
                Else
                    .RowHidden(1) = False
                End If
            End With
        End If
        
    Case "人员项目"
        
        If vsfPerson.IsSubtotal(vsfPerson.Row) Then
            gstrSQL = GetPublicSQL(SQL.体检项目清单)
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(vsf.RowData(vsf.Row)), vsfPerson.TextMatrix(vsfPerson.Row, 0))
        Else
            gstrSQL = GetPublicSQL(SQL.人员体检项目)
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(vsf.RowData(vsf.Row)), Val(vsfPerson.RowData(vsfPerson.Row)))
        End If
        
        Dim sglSum(0 To 1) As Single
        
        If rs.BOF = False Then
            Call LoadGrid(vsfItem, rs, , , ilsGrid)
            
            If vsfPerson.IsSubtotal(vsfPerson.Row) = False Then
                
                '计算费用总额,900,7,1,1,;体检价格,
                For lngLoop = 1 To vsfItem.Rows - 1
                    sglSum(0) = sglSum(0) + Val(vsfItem.TextMatrix(lngLoop, mCol.i基本价格))
                    sglSum(1) = sglSum(1) + Val(vsfItem.TextMatrix(lngLoop, mCol.i体检价格))
                Next
                
                vsfItem.Rows = vsfItem.Rows + 1
                vsfItem.TextMatrix(vsfItem.Rows - 1, mCol.i基本价格) = " " & Format(sglSum(0), "0.00")
                vsfItem.TextMatrix(vsfItem.Rows - 1, mCol.i体检价格) = Format(sglSum(1), "0.00")
                vsfItem.MergeCells = flexMergeFree
                vsfItem.MergeRow(vsfItem.Rows - 1) = True
                vsfItem.Cell(flexcpText, vsfItem.Rows - 1, 0, vsfItem.Rows - 1, mCol.i基本价格 - 1) = "合计："
                vsfItem.Cell(flexcpForeColor, vsfItem.Rows - 1, 0, vsfItem.Rows - 1, vsfItem.Cols - 1) = COLOR.兰色
            End If
        End If
        
    End Select
    
    RefreshData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetItem(ByRef lngKey As Long, ByVal intFoot As Integer) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：供编辑数据窗体调用，接口函数
    '------------------------------------------------------------------------------------------------------------------
    Dim lngIndex As Long
    Dim lngLoop As Long
    
    
    On Error GoTo errHand
    
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) = lngKey Then
            Exit For
        End If
    Next
    
    If lngLoop < vsf.Rows And lngLoop > 0 Then
        
        lngIndex = lngLoop
        lngIndex = lngLoop + intFoot
        
        If Val(vsf.RowData(lngIndex)) > 0 Then
            lngKey = Val(vsf.RowData(lngIndex))
            GetItem = True
        End If
    End If
    
    Exit Function
    
errHand:
    
End Function

Private Function MenuClick(ByVal strMenuItem As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：数据编辑/处理
    '------------------------------------------------------------------------------------------------------------------
    Dim lngKey As Long
    Dim lngTmp As Long
    Dim blnTran As Boolean
    Dim lngLoop As Long
    Dim strSQL() As String
    Dim rsItems As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim strGroup As String
    Dim strPrompt As String
    Dim intCount2 As Integer
    Dim lng门诊号 As Long
    Dim bytNew As Byte
    
    On Error GoTo errHand
    
    ReDim Preserve strSQL(1 To 1)
    
    lngKey = Val(vsf.RowData(vsf.Row))
    
    '第一步处理
    Select Case strMenuItem
    Case "个人预约"             '个人预约---------------------------------------------------------------------------------
            
        If Not frmScheduleEdit.ShowEdit(Me, 0, mlngSvrDept) Then Exit Function
        
    Case "团体预约"             '团体预约---------------------------------------------------------------------------------
        
        If Not frmScheduleEdit.ShowEdit(Me, 0, mlngSvrDept, True) Then Exit Function
        
    Case "修改个人预约"         '修改个人预约---------------------------------------------------------------------------------
        
        If lngKey = 0 Then Exit Function
        If Not frmScheduleEdit.ShowEdit(Me, lngKey, mlngSvrDept) Then Exit Function
        
    Case "修改团体预约"         '修改团体预约---------------------------------------------------------------------------------
                
        If lngKey = 0 Then Exit Function
        If Not frmScheduleEdit.ShowEdit(Me, lngKey, mlngSvrDept, True) Then Exit Function
        
    Case "删除预约"             '删除预约---------------------------------------------------------------------------------

        If lngKey = 0 Then Exit Function
        
        If MsgBox("你真的要删除当前体检预约吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                
        strSQL(ReDimArray(strSQL)) = "ZL_体检登记记录_DELETE(" & lngKey & ")"
        
    Case "确认预约"             '确认预约---------------------------------------------------------------------------------
        If lngKey = 0 Then Exit Function
        
        If MsgBox("你真的要确认当前体检预约吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        
        Select Case CheckAllowMedical(lngKey)
        Case 1
            strPrompt = "当前体检还没有设置体检团体，继续吗？"
            If MsgBox(strPrompt, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        Case 2
            strPrompt = "当前体检还没有设置体检人员，继续吗？"
            If MsgBox(strPrompt, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        Case 3
            strPrompt = "当前体检的体检项目不完整（每种组别必须有体检项目），继续吗？"
            If MsgBox(strPrompt, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        Case 4
            strPrompt = "存在没有分组的体检人员，请先进行人员组别划分！"
            ShowSimpleMsg strPrompt
            Exit Function
        End Select
                
        
        strSQL(ReDimArray(strSQL)) = "ZL_体检登记记录_STATE(" & lngKey & ",2)"
        
    Case "取消预约"             '取消预约---------------------------------------------------------------------------------
        If lngKey = 0 Then Exit Function
        
        If MsgBox("你真的要取消当前体检预约确认吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        
        strSQL(ReDimArray(strSQL)) = "ZL_体检登记记录_STATE(" & lngKey & ",1)"
        
    Case "组别项目"             '打开体检项目选择功能---------------------------------------------------------------------------------
                
        If lngKey = 0 Then Exit Function
        
        Call MedicalItemsRecord(rsItems)
        
        '读取体检项目
        gstrSQL = GetPublicSQL(SQL.团体体检项目)
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(vsf.RowData(vsf.Row)))
        
        If WriteItems(rs, rsItems, 2) = False Then Exit Function
        
        If Not frmItemsEdit.ShowEdit(Me, Val(vsf.RowData(vsf.Row)), rsItems, mlngSvrDept, IIf(vsf.TextMatrix(vsf.Row, GetCol(vsf, "[性质]")) <> "", False, True)) Then Exit Function

        '处理已经删除的体检项目
        Call FilterRecord(rsItems, "删除='1'")
        Call DeleteMedicalItems(strSQL, rsItems, vsf.TextMatrix(vsf.Row, GetCol(vsf, "No")), lngKey, 0)

        '处理新添加的体检项目
        Call FilterRecord(rsItems, "新加<>'1'")
        Call InsertMedicalItems(strSQL, rsItems, lngKey, 0)

        strSQL(ReDimArray(strSQL)) = "ZL_体检登记记录_体检类型(" & lngKey & ")"
        
    Case "人员项目"
        
        '编辑体检人员的个人体检项目
        
        If lngKey = 0 Then Exit Function
        If Val(vsfPerson.RowData(vsfPerson.Row)) = 0 Then Exit Function
        
        Call MedicalItemsRecord(rsItems)

        gstrSQL = GetPublicSQL(SQL.人员体检项目)
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(vsf.RowData(vsf.Row)), Val(vsfPerson.RowData(vsfPerson.Row)))
        Call WriteItems(rs, rsItems, 2)

        If Not frmItemsEdit.ShowEdit(Me, lngKey, rsItems, mlngSvrDept, False, 1, Val(vsfPerson.RowData(vsfPerson.Row))) Then Exit Function

        '处理已经删除的体检项目
        Call FilterRecord(rsItems, "删除='1'")
        Call DeleteMedicalItems(strSQL, rsItems, vsf.TextMatrix(vsf.Row, GetCol(vsf, "No")), lngKey, Val(vsfPerson.TextMatrix(vsfPerson.Row, GetCol(vsfPerson, "病人id"))))

        '处理新添加的体检项目
        Call FilterRecord(rsItems, "新加<>'1'")
        Call InsertMedicalItems(strSQL, rsItems, lngKey, Val(vsfPerson.TextMatrix(vsfPerson.Row, GetCol(vsfPerson, "病人id"))))

    Case "受检人员"
        
        If lngKey = 0 Then Exit Function
        
        Dim lng病人id As Long
                
        Call MedicalItemsRecord(rsItems, 2)
        
        gstrSQL = GetPublicSQL(SQL.体检人员档案)
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
        If WriteItems(rs, rsItems, , 2) = False Then Exit Function
        
        lngTmp = Val(vsf.TextMatrix(vsf.Row, GetCol(vsf, "合约单位id")))
        If lngTmp = 0 Then Exit Function
        
        If Not frmPersonEdit.ShowEdit(Me, Val(vsf.RowData(vsf.Row)), rsItems, IIf(vsf.TextMatrix(vsf.Row, GetCol(vsf, "[性质]")) <> "", False, True), , lngTmp) Then Exit Function
        
        Dim intCount As Integer
        Dim intCount1 As Integer
        
        intCount = -1
        
        strSQL(ReDimArray(strSQL)) = "zl_体检人员档案_Delete(" & lngKey & ")"
        
        rsItems.Filter = ""
        Do While Not rsItems.EOF
            
            '检查出生日期
            If rsItems("出生日期") <> "" Then
                
                If CheckStrValid(rsItems("出生日期"), CHECKFORMAT.日期) = False Then
                    ShowSimpleMsg rsItems("姓名").Value & "的出生日期无效！"
                    Exit Function
                End If
            End If
            
            lng病人id = rsItems("病人ID").Value
            bytNew = 0
            If lng病人id = 0 Then
                intCount = intCount + 1
                lng病人id = GetNextNo(1)
                bytNew = 1
            End If
            
            intCount1 = intCount1 + 1
            
            If zlCommFun.NVL(rsItems("门诊号").Value, 0) < 1 Then
                lng门诊号 = GetNextNo(3)
                
                intCount2 = intCount2 + 1
            Else
                lng门诊号 = zlCommFun.NVL(rsItems("门诊号").Value, 0)
            End If
            
            strSQL(ReDimArray(strSQL)) = "ZL_体检人员档案_INSERT(" & lngKey & "," & _
                                                                IIf(lng病人id = 0, "NULL", lng病人id) & ",'" & _
                                                                rsItems("组别").Value & "','" & _
                                                                rsItems("姓名").Value & "','" & _
                                                                rsItems("身份证").Value & "','" & _
                                                                rsItems("性别").Value & "'," & _
                                                                IIf(rsItems("出生日期").Value = "", "NULL", "TO_DATE('" & rsItems("出生日期").Value & "','yyyy-mm-dd hh24:mi:ss')") & ",'" & _
                                                                rsItems("婚姻状况").Value & "','" & _
                                                                rsItems("民族").Value & "','" & _
                                                                rsItems("国籍").Value & "','" & _
                                                                rsItems("学历").Value & "','" & _
                                                                rsItems("职业").Value & "','" & _
                                                                rsItems("联系人姓名").Value & "','" & _
                                                                rsItems("联系人电话").Value & "','" & _
                                                                rsItems("电子邮件").Value & "','" & _
                                                                rsItems("联系人地址").Value & "','" & _
                                                                rsItems("工作单位").Value & "','" & _
                                                                rsItems("年龄").Value & "'," & _
                                                                lng门诊号 & ",'" & _
                                                                rsItems("IC卡号").Value & "','" & _
                                                                rsItems("健康号").Value & "'," & _
                                                                rsItems("就诊卡号").Value & "'," & _
                                                                "1," & _
                                                                IIf(intCount1 = rsItems.RecordCount, "1", "0") & ",0," & bytNew & _
                                                                ",Null)"
            
            rsItems.MoveNext
        Loop
        
    Case "体检团体"                 '打开团体信息编辑功能(合约单位)------------------------------------------------------------
        
        If lngKey = 0 Then Exit Function
        If vsf.TextMatrix(vsf.Row, GetCol(vsf, "[性质]")) <> "" Then Exit Function
        lngTmp = Val(vsf.TextMatrix(vsf.Row, GetCol(vsf, "合约单位id")))
        If lngTmp = 0 Then Exit Function
                                        
    Case "组别人员"
        
        If lngKey = 0 Then Exit Function
        If Not frmPatientGroupEdit.ShowEdit(Me, lngKey) Then Exit Function
                            
    Case "照片采集"
        
        If lngKey > 0 Then
            Call frmPersonPhoto.ShowEdit(Me, lngKey, 0)
            Exit Function
        End If
        
    End Select
    
    '第二步处理
    blnTran = True
    gcnOracle.BeginTrans
    
    For lngLoop = 1 To UBound(strSQL)
        If strSQL(lngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(lngLoop), Me.Caption)
    Next
    
    gcnOracle.CommitTrans
    blnTran = False
    
    Select Case strMenuItem
    Case "删除预约"
    
        '删除行
        If vsf.Rows = 2 Then
            Call ResetVsf(vsf)
        Else
            vsf.RemoveItem vsf.Row
        End If
        
        Call vsf_AfterRowColChange(0, 0, vsf.Row, vsf.Col)
'        Call AppendSapceRows(vsf, lnX, lnY)
        
        MenuClick = True
        
        Exit Function
        
    End Select
    
    Call mnuViewRefresh_Click
    
    Call AdjustEnableState
    Call RefreshStateInfo
    
    MenuClick = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTran Then gcnOracle.RollbackTrans
End Function

Private Sub PrintData(ByVal bytMode As Byte)
    '------------------------------------------------------------------------------------------------------------------
    '功能： 打印数据
    '参数： bytMode                         打印方式（1-打印；2-预览；3-输出到Excel）
    '------------------------------------------------------------------------------------------------------------------
    Dim objPrint As New zlPrint1Grd
    Dim objRow As zlTabAppRow
    
'    If cboGroup.ListCount < 1 Then Exit Sub
    
    If UserInfo.姓名 = "" Then Call GetUserInfo
    
    Call CopyGrid(vsf, vsfPrint, 2)
    objPrint.Title.Text = "体检预约清单"
    
    Set objRow = New zlTabAppRow
    objRow.Add "体检部门：" & cboDept.Text
    objRow.Add ""
    
    objPrint.UnderAppRows.Add objRow
    
    Set objPrint.Body = vsfPrint

    If bytMode = 1 Then bytMode = zlPrintAsk(objPrint)

    If bytMode >= 1 And bytMode <= 3 Then Call zlPrintOrView1Grd(objPrint, bytMode)
        
End Sub

'（３）窗体及其控件的事件处理******************************************************************************************
Private Sub cboDept_Click()
    
    If mblnStartUp Then Exit Sub
    If mlngSvrDept = cboDept.ItemData(cboDept.ListIndex) Then Exit Sub
    
    mlngSvrDept = cboDept.ItemData(cboDept.ListIndex)
    
    Call mnuViewRefresh_Click
    
End Sub

Private Sub Form_Activate()
    
    If mblnStartUp = False Then Exit Sub
    DoEvents
        
    If InitActivate = False Then
        mblnStartUp = False
        Unload Me
        Exit Sub
    End If
    
    mblnStartUp = False
    
    Call cboDept_Click
    
End Sub

Private Sub Form_Load()
    
    mblnStartUp = True
    
    Call RestoreWinState(Me, App.ProductName)
    
    Call InitLoad
    
    Call ApplyPrivilege(gstrPrivs)
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
    If imgX_S.Top > Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - 1500 Then
        imgX_S.Top = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - 1500
    End If
    
    With vsf
        .Left = 0
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0)
        .Width = Me.ScaleWidth
        .Height = imgX_S.Top - .Top
    End With
    
    With imgX_S
        .Left = vsf.Left
        .Top = vsf.Top + vsf.Height
        .Width = vsf.Width
    End With

    With vsfPerson
        .Left = 0
        .Top = imgX_S.Top + imgX_S.Height
        .Width = imgY_S.Left - .Left
        .Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - .Top
    End With
    
    With imgY_S
        .Top = vsfPerson.Top
        .Height = vsfPerson.Height
    End With
    
    With vsfItem
        .Left = imgY_S.Left + imgY_S.Width
        .Top = vsfPerson.Top
        .Width = Me.ScaleWidth - .Left - 30
        .Height = vsfPerson.Height
    End With
    
'    Call AppendSapceRows(vsf, lnX, lnY)
'    Call AppendSapceRows(vsfPerson, lnX1, lnY1)
'    Call AppendSapceRows(vsfItem, lnX2, lnY2)
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
        
    If mblnStartUp Then
        Cancel = True
        Exit Sub
    End If
    '使用个性化设置

    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub imgX_S_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    On Error Resume Next
    
    imgX_S.Top = imgX_S.Top + Y
    
    If imgX_S.Top < 1500 Then imgX_S.Top = 1500
    If Me.Height - imgX_S.Top - imgX_S.Height < 1500 Then imgX_S.Top = Me.Height - imgX_S.Height - 1500
    
            
    Form_Resize
End Sub

Private Sub imgY_S_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    imgY_S.Left = imgY_S.Left + X
    
    If imgY_S.Left < 3000 Then imgY_S.Left = 3000
    If Me.Width - imgY_S.Left - imgY_S.Width < 1000 Then imgY_S.Left = Me.Width - imgY_S.Width - 1000

    Form_Resize
End Sub

Private Sub mnuAdditionPerItems_Click()
    If vsf.TextMatrix(vsf.Row, GetCol(vsf, "[性质]")) <> "" Then
        Call MenuClick("组别项目")
    Else
        Call MenuClick("人员项目")
    End If
End Sub

Private Sub mnuAdditionPhoto_Click()
    Call MenuClick("照片采集")
End Sub

Private Sub mnuEditAdd_Click()
    Call MenuClick("个人预约")
End Sub

Private Sub mnuEditAddGroup_Click()
    Call MenuClick("团体预约")
End Sub

Private Sub mnuEditCancel_Click()
    Call MenuClick("取消预约")
End Sub

Private Sub mnuEditCheck_Click()
    Call MenuClick("确认预约")
End Sub

Private Sub mnuEditDelete_Click()
    Call MenuClick("删除预约")
End Sub

Private Sub mnuAdditionGroup_Click()
    Call MenuClick("体检团体")
End Sub

Private Sub mnuAdditionGroupMember_Click()
    Call MenuClick("组别人员")
End Sub

Private Sub mnuAdditionItems_Click()
    Call MenuClick("组别项目")
End Sub

Private Sub mnuEditModify_Click()
    If vsf.TextMatrix(vsf.Row, 0) <> "" Then
        Call MenuClick("修改个人预约")
    Else
        Call MenuClick("修改团体预约")
    End If
End Sub

Private Sub mnuAdditionPersons_Click()
    Call MenuClick("受检人员")
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub


Private Sub mnuFileNew_Click()
    frmScheduleExcel.ShowEdit Me
End Sub

Private Sub mnuFileOutExcel_Click()
    Call PrintData(3)
End Sub

Private Sub mnuFilePrint_Click()
    Call PrintData(1)
End Sub

Private Sub mnuFilePrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub mnuFilePrintView_Click()
    Call PrintData(2)
End Sub

Private Sub mnuHelpAbout_Click()
    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
End Sub

Private Sub mnuHelpTopic_Click()
   Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hWnd)
End Sub

Private Sub mnuViewRefresh_Click()
    Dim lngSvrKey As Long
    Dim lngSvrKey2 As Long
    Dim lngSvrKey3 As Long
                
    '保存体检预约、体检组别、体检项目
    lngSvrKey = Val(vsf.RowData(vsf.Row))
    
    mblnAllowChange = False
    
    Call ResetVsf(vsf)
    Call ResetVsf(vsfPerson)
    Call ResetVsf(vsfItem)
        
    Call RefreshData("体检预约")
    
    '恢复体检预约
    Call RestoreRow(vsf, lngSvrKey)
    
    Call RefreshData("体检人员")
    
    mblnAllowChange = True
    
    Call vsfPerson_AfterRowColChange(0, 0, vsfPerson.Row, vsfPerson.Col)
    
'    Call AppendSapceRows(vsfPerson, lnX1, lnY1)
'    Call AppendSapceRows(vsf, lnX, lnY)
    
        
    Call AdjustEnableState
    Call RefreshStateInfo
End Sub

Private Sub mnuViewSearch_Click()
    If frmScheduleFilter.ShowFilter(Me, mstrCondition) Then
        Call mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    cbrThis.Visible = mnuViewToolButton.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim intLoop As Integer
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For intLoop = 1 To tbrThis.Buttons.Count
        tbrThis.Buttons(intLoop).Caption = IIf(mnuViewToolText.Checked, tbrThis.Buttons(intLoop).Tag, "")
    Next
    cbrThis.Bands(1).MinHeight = tbrThis.Height
    Call Form_Resize
    
End Sub

Private Sub mobjPopMenu_MenuBeforeShow(Cancel As Boolean)
    
    Select Case mbytPopMenu
    Case 1
        
        If mnuEdit.Visible = False Then Exit Sub
        
        If mnuEditAdd.Visible Then mobjPopMenu.Add 1, mnuEditAdd.Caption, , , mnuEditAdd.Enabled
        If mnuEditAddGroup.Visible Then mobjPopMenu.Add 2, mnuEditAddGroup.Caption, , , mnuEditAddGroup.Enabled
    Case 2
        
        If mnuAddition.Visible = False Then Exit Sub

        If mnuAdditionPersons.Visible Then mobjPopMenu.Add 1, mnuAdditionPersons.Caption, , , mnuAdditionPersons.Enabled
        
        mobjPopMenu.Add 2, "-", , 2, True
        
        If mnuAdditionGroupMember.Visible Then mobjPopMenu.Add 3, mnuAdditionGroupMember.Caption, , , mnuAdditionGroupMember.Enabled
        
    End Select
    
End Sub

Private Sub mobjPopMenu_MenuClick(ByVal Key As Long, ByVal Caption As String)
    Select Case mbytPopMenu
    Case 1
        Select Case Key
        Case 1
            Call mnuEditAdd_Click
        Case 2
            Call mnuEditAddGroup_Click
        End Select
    Case 2
        Select Case Key
        Case 1
            Call mnuAdditionPersons_Click
        Case 3
            Call mnuAdditionGroupMember_Click
        End Select
    End Select
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(tbrThis.hWnd, objPoint)
    
    Select Case Button.Key
    Case "预览"
        Call mnuFilePrintView_Click
    Case "打印"
        Call mnuFilePrint_Click
    Case "预约"
        
        mbytPopMenu = 1
        Set mobjPopMenu = New clsPopMenu
        Call mobjPopMenu.ShowPopupMenu(objPoint.X * 15 + Button.Left - 15, objPoint.Y * 15 + Button.Top + Button.Height + 15)
        
    Case "修改"
        Call mnuEditModify_Click
    Case "删除"
        Call mnuEditDelete_Click
    Case "确认"
        Call mnuEditCheck_Click
    Case "取消"
        Call mnuEditCancel_Click
    Case "过滤"
        Call mnuViewSearch_Click
    Case "帮助"
        Call mnuHelpTopic_Click
    Case "退出"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub tbrThis_ButtonDropDown(ByVal Button As MSComctlLib.Button)
    Call tbrThis_ButtonClick(Button)
End Sub

Private Sub tbrThis_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
    Case "个人"
        Call mnuEditAdd_Click
    Case "团体"
        Call mnuEditAddGroup_Click
    End Select
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If mblnAllowChange = False Then Exit Sub
        
    If mlngSvrKey = Val(vsf.RowData(NewRow)) Then Exit Sub
    mlngSvrKey = Val(vsf.RowData(NewRow))
    
    Call ResetVsf(vsfPerson)
    Call ResetVsf(vsfItem)
        
    Call RefreshData("体检人员")
    
    Call vsfPerson_AfterRowColChange(0, 0, vsfPerson.Row, vsfPerson.Col)

    
    Call AdjustEnableState
    
End Sub

Private Sub vsf_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = (Col < 1)
End Sub

Private Sub vsf_DblClick()
    If mnuEdit.Visible And mnuEditModify.Visible And mnuEditModify.Enabled Then
        Call mnuEditModify_Click
    End If
End Sub

Private Sub vsf_GotFocus()
    vsf.BackColorSel = COLOR.焦点
End Sub

Private Sub vsf_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call vsf_DblClick
End Sub

Private Sub vsf_LostFocus()
    vsf.BackColorSel = COLOR.非焦点
End Sub

Private Sub vsf_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SendLMouseButton(vsf.hWnd, X, Y)
        If mnuEdit.Visible Then Me.PopupMenu mnuEdit
    End If
End Sub


Private Sub vsfItem_GotFocus()
    vsfItem.BackColorSel = COLOR.焦点
End Sub

Private Sub vsfItem_LostFocus()
    vsfItem.BackColorSel = COLOR.非焦点
End Sub

Private Sub vsfPerson_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    
    If mblnAllowChange = False Then Exit Sub
    
    Call ResetVsf(vsfItem)
    Call RefreshData("人员项目")
    
    Call AdjustEnableState
    
End Sub



Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

Private Sub vsfPerson_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SendLMouseButton(vsfPerson.hWnd, X, Y)
        If mnuAddition.Visible Then Me.PopupMenu mnuAddition
    End If
End Sub
