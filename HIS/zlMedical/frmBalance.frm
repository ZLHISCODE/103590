VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmBalance 
   Caption         =   "团体体检结算"
   ClientHeight    =   6885
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10305
   Icon            =   "frmBalance.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   10305
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.TabStrip tbs 
      Height          =   360
      Left            =   3510
      TabIndex        =   8
      Top             =   2505
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   635
      Style           =   2
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&1.结算明细"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&2.结算方式"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6525
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmBalance.frx":1CFA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13097
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
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   1244
      BandCount       =   2
      _CBWidth        =   10305
      _CBHeight       =   705
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinWidth1       =   4995
      MinHeight1      =   645
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Caption2        =   "体检部门"
      Child2          =   "cboDept"
      MinWidth2       =   2100
      MinHeight2      =   300
      Width2          =   2100
      NewRow2         =   0   'False
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   8115
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   195
         Width           =   2100
      End
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   645
         Left            =   165
         TabIndex        =   2
         Top             =   30
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   1138
         ButtonWidth     =   820
         ButtonHeight    =   1138
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ilsMenu"
         HotImageList    =   "ilsMenuHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
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
               Caption         =   "结算"
               Key             =   "结算"
               Object.ToolTipText     =   "结算"
               Object.Tag             =   "结算"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "作废"
               Key             =   "作废"
               Object.ToolTipText     =   "作废"
               Object.Tag             =   "作废"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_2"
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "过滤"
               Key             =   "过滤"
               Object.ToolTipText     =   "过滤"
               Object.Tag             =   "过滤"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_3"
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "帮助"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   7
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   9255
      Top             =   900
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalance.frx":258E
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalance.frx":27AE
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalance.frx":29CE
            Key             =   "Parent"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalance.frx":2BEA
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalance.frx":2E0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalance.frx":3024
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalance.frx":3244
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalance.frx":3464
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsMenuHot 
      Left            =   8535
      Top             =   900
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalance.frx":3BDE
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalance.frx":3DFE
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalance.frx":401E
            Key             =   "Parent"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalance.frx":423A
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalance.frx":445A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalance.frx":4674
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalance.frx":4894
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalance.frx":4AB4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   1380
      Index           =   0
      Left            =   3420
      TabIndex        =   3
      Top             =   855
      Width           =   2790
      _cx             =   4921
      _cy             =   2434
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
      GridColor       =   12698049
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   2
      HighLight       =   0
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
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
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
      Begin VB.Line lnX0 
         Index           =   0
         Visible         =   0   'False
         X1              =   0
         X2              =   1785
         Y1              =   135
         Y2              =   135
      End
      Begin VB.Line lnY0 
         Index           =   0
         Visible         =   0   'False
         X1              =   825
         X2              =   825
         Y1              =   0
         Y2              =   1215
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   1845
      Index           =   1
      Left            =   4590
      TabIndex        =   4
      Top             =   3750
      Width           =   3795
      _cx             =   6694
      _cy             =   3254
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
      GridColor       =   12698049
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
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
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
      Begin VB.Line lnX1 
         Index           =   0
         Visible         =   0   'False
         X1              =   0
         X2              =   1785
         Y1              =   135
         Y2              =   135
      End
      Begin VB.Line lnY1 
         Index           =   0
         Visible         =   0   'False
         X1              =   825
         X2              =   825
         Y1              =   0
         Y2              =   1215
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   1845
      Index           =   2
      Left            =   3810
      TabIndex        =   5
      Top             =   2610
      Width           =   3720
      _cx             =   6562
      _cy             =   3254
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
      GridColor       =   12698049
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
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
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
      Begin VB.Line lnX2 
         Index           =   0
         Visible         =   0   'False
         X1              =   0
         X2              =   1785
         Y1              =   135
         Y2              =   135
      End
      Begin VB.Line lnY2 
         Index           =   0
         Visible         =   0   'False
         X1              =   825
         X2              =   825
         Y1              =   0
         Y2              =   1215
      End
   End
   Begin MSComctlLib.ImageList ils13 
      Left            =   8910
      Top             =   3105
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalance.frx":522E
            Key             =   "状态"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalance.frx":55C8
            Key             =   "结"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalance.frx":5B62
            Key             =   "废"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid vsfPrint 
      Height          =   780
      Left            =   9480
      TabIndex        =   7
      Top             =   1500
      Visible         =   0   'False
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   1376
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   270
      MergeCells      =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Image imgX_S 
      Height          =   45
      Left            =   4305
      MousePointer    =   7  'Size N S
      Top             =   2355
      Width           =   4845
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
      Begin VB.Menu mnuFileOutExcel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileBalanceBill 
         Caption         =   "打印票据(&O)"
      End
      Begin VB.Menu mnuFileBalanceDetail 
         Caption         =   "结算明细(C)"
         Begin VB.Menu mnuFileBalanceDetaiPrintView 
            Caption         =   "预览(&1)"
         End
         Begin VB.Menu mnuFileBalanceDetaiPrint 
            Caption         =   "打印(&2)"
         End
         Begin VB.Menu mnuFileBalanceDetaiExcel 
            Caption         =   "输出到Excel(&3)"
         End
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePara 
         Caption         =   "参数设置(&R)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFile_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditBalance 
         Caption         =   "体检结算(&B)"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuEditBalanceCancel 
         Caption         =   "结算作废(&M)"
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
      Begin VB.Menu mnuViewList 
         Caption         =   "未结团体清单(&L)"
      End
      Begin VB.Menu mnuViewSearch 
         Caption         =   "过滤(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuView_3 
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
Attribute VB_Name = "frmBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'（１）窗体级变量定义**************************************************************************************************
Private mblnStartUp As Boolean                                  '窗体启动标志
Private mlngSvrKey(0 To 2)  As Long                             '用于保存各个区域选中的行关键字
Private mlngDept As Long
Private mblnNoAllowChange As Boolean
Private WithEvents mobjPopMenu As clsPopMenu                    '自定义弹出菜单对象
Attribute mobjPopMenu.VB_VarHelpID = -1
Private mstrCondition As String
Private mbytKind As Byte
Private Type TYPE_USR_CELL
    Row As Integer
    Col As Integer
End Type
Private mblnDataMoved As Boolean

Private musrSavePos As TYPE_USR_CELL

'（２）自定义过程或函数************************************************************************************************
Private Function ClearData(Optional ByVal strMenuItem As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  清除控件中的数据
    '参数:  strMenuItem             要清除的范围
    '返回;  True                    清除成功
    '       False                   清除失败
    '------------------------------------------------------------------------------------------------------------------
    
    Select Case strMenuItem
    Case "体检结算"
        Call ResetVsf(vsf(0))
        Call InheritAppendSpaceRows(0)
    Case "结算单据"
        Call ResetVsf(vsf(1))
        Call InheritAppendSpaceRows(1)
    Case "结算方式"
        Call ResetVsf(vsf(2))
        Call InheritAppendSpaceRows(2)
    End Select
        
End Function

Private Sub InheritAppendSpaceRows(ByVal intIndex As Integer)
    '------------------------------------------------------------------------------------------------------------------
    '功能： 补齐表格空行
    '参数:  intIndex                要补充空行的表格控件索引号
    '------------------------------------------------------------------------------------------------------------------
    Select Case intIndex
    Case 0
        Call AppendRows(vsf(intIndex), lnX0, lnY0)
    Case 1
        Call AppendRows(vsf(intIndex), lnX1, lnY1)
    Case 2
        Call AppendRows(vsf(intIndex), lnX2, lnY2)
    End Select
End Sub

Private Function InitLoad() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  初始化数据，发生在窗体的Load事件
    '返回:  True                    成功
    '       False                   出错
    '------------------------------------------------------------------------------------------------------------------
    Dim strVsf As String
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    mbytKind = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "共用结帐票据类型", 3))
    
    mlngDept = 0
    mstrCondition = Format(DateAdd("d", -7, CDate(zlDatabase.Currentdate)), "yyyy-MM-dd") & "'" & Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    mstrCondition = mstrCondition & "''''''''"
        
    strVsf = ",255,4,1,1,[状态];单据号,900,1,1,1,;票据号,900,1,1,1,;团体名称,1800,1,1,1,;结算金额,900,7,1,1,;结算人,810,1,1,1,;结算时间,1670,1,1,1,;联系人,810,1,1,1,;联系电话,1200,1,1,1,;联系地址,1800,1,1,1,"
    Call CreateVsf(vsf(0), strVsf)
    vsf(0).Cols = vsf(0).Cols + 1
    vsf(0).ColWidth(vsf(0).Cols - 1) = 15
    Set vsf(0).Cell(flexcpPicture, 0, 0) = ils13.ListImages("状态").Picture
    
    strVsf = "姓名,810,1,1,1,;单据号,900,1,1,1,;项目,2400,1,1,1,;费目,750,1,1,1,;结算金额,900,7,1,1,;开单科室,1080,1,1,1,;费用时间,1670,1,1,1,"
    Call CreateVsf(vsf(1), strVsf)
    vsf(1).Cols = vsf(1).Cols + 1
    vsf(1).ColWidth(vsf(1).Cols - 1) = 15
        
    strVsf = "单据号,900,1,1,1,;金额,810,7,1,1,;结算方式,900,1,1,1,;结算号码,900,1,1,1,"
    Call CreateVsf(vsf(2), strVsf)
    vsf(2).Cols = vsf(2).Cols + 1
    vsf(2).ColWidth(vsf(2).Cols - 1) = 15
    
    '票号严格控制
    '------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String
    
    strTmp = zlDatabase.GetPara(24, glngSys, , "00000")
    If strTmp <> "" Then
        gblnBill结帐 = (Mid(strTmp, 3, 1) = "1")
        gblnStrictCtrl = (Mid(strTmp, 3, 1) = "1")
    End If
               
    glng结帐ID = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "共用结帐票据批次", 0)
    glngShareUseID = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "共用结帐票据批次", 0)
    
    InitLoad = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function MenuClick(ByVal strMenuItem As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能： 数据编辑/处理
    '参数:  strMenuItem             操作名称
    '返回:  True                    成功
    '       False                   失败/取消/出错
    '------------------------------------------------------------------------------------------------------------------
    Dim lngKey As Long
    Dim blnTran As Boolean
    Dim lngLoop As Long
    Dim strSQL() As String
    Dim strNo As String
    
    On Error GoTo errHand
        
    ReDim Preserve strSQL(1 To 1)
    
    lngKey = Val(vsf(0).RowData(vsf(0).Row))
        
    '第一步处理
    Select Case strMenuItem
    Case "体检结算"
        
        If Not frmBalanceEdit.ShowEdit(Me, 0) Then Exit Function
                
    
    Case "结算作废"
        
        If lngKey = 0 Then Exit Function
        strNo = vsf(0).TextMatrix(vsf(0).Row, GetCol(vsf(0), "单据号"))
        If strNo = "" Then Exit Function
        
        '数据转储处理
        '--------------------------------------------------------------------------------------------------------------
        If mblnDataMoved Then
            ShowSimpleMsg "此结帐单据已经转出，不能再操作。"
            Exit Function
        End If
        
        If MsgBox("真的要作废当前结算单据吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        
        strSQL(ReDimArray(strSQL)) = "zl_体检结算记录_Cancel(" & Val(vsf(0).RowData(vsf(0).Row)) & ")"
        
    End Select
    
    blnTran = True
    gcnOracle.BeginTrans
    For lngLoop = 1 To UBound(strSQL)
        If strSQL(lngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(lngLoop), Me.Caption)
    Next
    gcnOracle.CommitTrans
    blnTran = False
        
    Select Case strMenuItem
    Case "结算作废", "体检结算"
        Call mnuViewRefresh_Click
    End Select
    
    Call AdjustEnableState
    Call RefreshStateInfo
    
    MenuClick = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
    If blnTran Then gcnOracle.RollbackTrans
End Function


Private Sub PrintData(ByVal bytMode As Byte)
    '--------------------------------------------------------------------------------------------------------
    '功能： 打印数据
    '参数： bytMode                         打印方式（1-打印；2-预览；3-输出到Excel）
    '--------------------------------------------------------------------------------------------------------
    Dim objPrint As New zlPrint1Grd
                
    mblnNoAllowChange = True
    
    musrSavePos.Row = vsf(0).Row
    musrSavePos.Col = vsf(0).Col
    
    If UserInfo.姓名 = "" Then Call GetUserInfo

    objPrint.Title = "“" & zlCommFun.GetNeedName(cboDept.Text) & "”团体体检结算单"
    
    Call CopyGrid(vsf(0), vsfPrint, 1)
    
    Set objPrint.Body = vsfPrint

    If bytMode = 1 Then bytMode = zlPrintAsk(objPrint)
    If bytMode >= 1 And bytMode <= 3 Then Call zlPrintOrView1Grd(objPrint, bytMode)
    
    On Error Resume Next
    vsf(0).Row = musrSavePos.Row
    vsf(0).Col = musrSavePos.Col
    vsf(0).ShowCell vsf(0).Row, vsf(0).Col
    On Error GoTo 0
    
    mblnNoAllowChange = False
End Sub

Private Sub ApplyPrivilege(ByVal strPrivilege As String)
    '------------------------------------------------------------------------------------------------------------------
    '功能： 应用权限处理
    '参数： strPrivilege                    权限
    '------------------------------------------------------------------------------------------------------------------
    
'    strPrivilege = "体检结算;结算作废;结算重打"
    
    '不具有“体检结算”和“结算作废”权限时
    If InStr(strPrivilege, "体检结算") = 0 And InStr(strPrivilege, "结算作废") = 0 Then
        mnuEdit.Visible = False
    Else
        '不具有“体检结算”权限时
        If InStr(strPrivilege, "体检结算") = 0 Then
            mnuEditBalance.Visible = False
        End If
        
        '不具有“结算作废”权限时
        If InStr(strPrivilege, "结算作废") = 0 Then
            mnuEditBalanceCancel.Visible = False
        End If
        
    End If
    
    '不具有“结算重打”权限时
    If InStr(strPrivilege, "结算重打") = 0 Then
        mnuFileBalanceBill.Visible = False
        mnuFileBalanceDetail.Visible = False
        mnuFile_2.Visible = False
    End If
            
    '处理工具栏
    tbrThis.Buttons("结算").Visible = mnuEdit.Visible And mnuEditBalance.Visible
    tbrThis.Buttons("作废").Visible = mnuEdit.Visible And mnuEditBalanceCancel.Visible
    tbrThis.Buttons("Split_2").Visible = tbrThis.Buttons("结算").Visible Or tbrThis.Buttons("作废").Visible
    
End Sub

Private Sub AdjustEnableState()
    '------------------------------------------------------------------------------------------------------------------
    '功能： 调整各功能菜单的可用状态
    '------------------------------------------------------------------------------------------------------------------
    
    mnuFilePrint.Enabled = True
    mnuFilePrintView.Enabled = True
    mnuFileOutExcel.Enabled = True
    
    mnuFileBalanceBill.Enabled = True
    mnuFileBalanceDetail.Enabled = True
    
    mnuEditBalanceCancel.Enabled = True
    
    If Val(vsf(0).RowData(vsf(0).Row)) = 0 Then
        mnuFilePrint.Enabled = False
        mnuFilePrintView.Enabled = False
        mnuFileOutExcel.Enabled = False
        
        mnuFileBalanceBill.Enabled = False
        mnuFileBalanceDetail.Enabled = False
        
        mnuEditBalanceCancel.Enabled = False
    Else
        Select Case vsf(0).TextMatrix(vsf(0).Row, GetCol(vsf(0), "[状态]"))
        Case "结"

        Case "废"
            mnuEditBalanceCancel.Enabled = False
            mnuFileBalanceBill.Enabled = False
        End Select
    End If
    
    If Val(vsf(1).RowData(vsf(1).Row)) = 0 Then
        mnuFileBalanceDetail.Enabled = False
    End If
    
    tbrThis.Buttons("预览").Enabled = mnuFilePrintView.Enabled
    tbrThis.Buttons("打印").Enabled = mnuFilePrint.Enabled
    
    tbrThis.Buttons("结算").Enabled = mnuEditBalance.Enabled
    tbrThis.Buttons("作废").Enabled = mnuEditBalanceCancel.Enabled
    
End Sub

Private Sub RefreshStateInfo()
    '------------------------------------------------------------------------------------------------------------------
    '功能：刷新状态栏显示信息
    '------------------------------------------------------------------------------------------------------------------
    If Val(vsf(0).RowData(1)) = 0 Then
        stbThis.Panels(2).Text = "没有结算单据。"
    Else
        stbThis.Panels(2).Text = "共有 " & vsf(0).Rows - 1 & " 张结算单据。"
    End If
    
End Sub

Private Function GetQueryCondition(ByVal strCondition As String) As String
    '--------------------------------------------------------------------------------------------------------
    '功能：
    '--------------------------------------------------------------------------------------------------------
    Dim varTmp As Variant
    Dim strResult As String
     
   
    '以下是根据设置条件构成的条件语句
    '存储格式:开始时间'结束时间'开始单据号'结束单据号'开始票据号'结束票据号'结算人'体检团体'体检团体id'体检号'包括确认
    
    If strCondition = "" Then Exit Function
        
    varTmp = Split(strCondition, "'")
    
    strResult = " AND C.收费时间 BETWEEN TO_DATE('" & Format(varTmp(0), "yyyy-mm-dd") & " 00:00:00', 'yyyy-mm-dd hh24:mi:ss') AND TO_DATE('" & Format(varTmp(1), "yyyy-mm-dd") & " 23:59:59', 'yyyy-mm-dd hh24:mi:ss')"
    
    '单据号
    If Trim(varTmp(2)) <> "" And Trim(varTmp(3)) <> "" Then
        strResult = strResult & " AND C.No BETWEEN '" & Trim(varTmp(2)) & "' AND '" & Trim(varTmp(3)) & "'"
    ElseIf Trim(varTmp(2)) <> "" Then
        strResult = strResult & " AND C.No='" & Trim(varTmp(2)) & "'"
    ElseIf Trim(varTmp(3)) <> "" Then
        strResult = strResult & " AND C.No='" & Trim(varTmp(3)) & "'"
    End If
    
    '实际票号
    If Trim(varTmp(4)) <> "" And Trim(varTmp(5)) <> "" Then
        strResult = strResult & " AND C.实际票号 BETWEEN '" & Trim(varTmp(4)) & "' AND '" & Trim(varTmp(5)) & "'"
    ElseIf Trim(varTmp(4)) <> "" Then
        strResult = strResult & " AND C.实际票号='" & Trim(varTmp(4)) & "'"
    ElseIf Trim(varTmp(5)) <> "" Then
        strResult = strResult & " AND C.实际票号='" & Trim(varTmp(5)) & "'"
    End If
    
    '结算人
    If Trim(varTmp(6)) <> "" Then strResult = strResult & " AND C.操作员姓名='" & Trim(varTmp(6)) & "'"
    
    '结算团体
    If Val(varTmp(8)) > 0 Then strResult = strResult & " AND A.合约单位id=" & Val(varTmp(8))
        
    '记录状态
    If Val(varTmp(9)) = 0 Then strResult = strResult & " AND C.记录状态=1"
    
    GetQueryCondition = strResult
    
End Function

Private Function RefreshData(Optional ByVal strMenuItem As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：刷新/装载数据
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim lngKey As Long
    Dim strTmp As String
    Dim blnDataMoved As Boolean
    
    On Error GoTo errHand
    
    Call InitSysPara
    
    Select Case strMenuItem
    Case "体检结算"
        
        Call ClearGrid(vsf(0))
        gstrSQL = "SELECT A.ID,A.单据号,A.票据号,A.结算金额,A.结算人,TO_CHAR(A.结算时间,'yyyy-mm-dd hh24:mi') AS 结算时间," & _
                        "DECODE(A.记录状态,1,'结','废') AS 状态," & _
                        "DECODE(A.记录状态,1,'0','192') AS 前景色," & _
                        "B.名称 AS 团体名称,B.联系人,B.电话 AS 联系电话,B.地址 AS 联系地址 FROM " & _
                        "( " & _
                        "SELECT C.ID,C.记录状态,A.合约单位id," & _
                               "C.NO AS 单据号, " & _
                               "A.结算金额, C.实际票号 AS 票据号," & _
                               "C.操作员姓名 AS 结算人, " & _
                               "C.收费时间 AS 结算时间 " & _
                        "FROM 体检结算记录 A, " & _
                             "病人结帐记录 C " & _
                        "Where C.ID=A.结算id " & _
                              "AND A.结算部门id+0=" & mlngDept & " " & _
                              "AND A.记录状态 IN (1,2) " & GetQueryCondition(mstrCondition) & " " & _
                        ") A, " & _
                        "合约单位 B " & _
                        "WHERE A.合约单位ID=B.ID "
                        
        '数据转储处理
        '--------------------------------------------------------------------------------------------------------------
        blnDataMoved = zlDatabase.DateMoved(Format(Split(mstrCondition, "'")(0), "yyyy-MM-dd HH:mm:ss"), , glngSys, Me.Caption)
        If blnDataMoved Then
            strTmp = gstrSQL
            strTmp = Replace(strTmp, "体检结算记录", "H体检结算记录")
            strTmp = Replace(strTmp, "病人结帐记录", "H病人结帐记录")
            gstrSQL = "Select * From (" & gstrSQL & " Union All " & strTmp & ") a "
        End If
        gstrSQL = gstrSQL & " Order By a.单据号 Desc"
        
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        If rs.BOF = False Then
            Call LoadGrid(vsf(0), rs, Array("", "", "", "", gstrDec, gstrDec), , ils13)
        End If
        Call InheritAppendSpaceRows(0)
        
    Case "结算单据"
        
        lngKey = Val(vsf(0).RowData(vsf(0).Row))
        If lngKey = 0 Then Exit Function
        
        Call ClearGrid(vsf(1))
        gstrSQL = "SELECT A.ID,A.姓名,A.NO AS 单据号,B.名称 AS 开单科室,C.名称 AS 项目,A.收据费目 AS 费目,A.结帐金额 AS 结算金额,TO_CHAR(A.发生时间,'yyyy-mm-dd hh24:mi') AS 费用时间 " & _
                    "FROM 病人费用记录 A, " & _
                         "部门表 B, " & _
                         "收费项目目录 C " & _
                    "WHERE A.结帐id = [1] " & _
                          "AND A.开单部门ID=B.ID " & _
                          "AND C.ID=A.收费细目ID "
        
        '数据转储处理
        '--------------------------------------------------------------------------------------------------------------
        mblnDataMoved = zlDatabase.NOMoved("病人结帐记录", vsf(0).TextMatrix(vsf(0).Row, 1))
        If mblnDataMoved Then
            gstrSQL = Replace(gstrSQL, "病人费用记录", "H病人费用记录")
        End If
        
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
        If rs.BOF = False Then
            Call LoadGrid(vsf(1), rs, Array("", "", "", "", gstrDec), , ils13)
        End If
        Call InheritAppendSpaceRows(1)
        
    Case "结算方式"
        
        lngKey = Val(vsf(0).RowData(vsf(0).Row))
        If lngKey = 0 Then Exit Function
        
        Call ClearGrid(vsf(2))
        
        gstrSQL = "SELECT A.ID,A.NO AS 单据号, A.冲预交 AS 金额,A.结算方式,A.结算号码 " & _
                    "FROM 病人预交记录 A " & _
                    "WHERE A.结帐ID=[1]"
        
        '数据转储处理
        '--------------------------------------------------------------------------------------------------------------
        If mblnDataMoved Then
            gstrSQL = Replace(gstrSQL, "病人预交记录", "H病人预交记录")
        End If
        
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
        If rs.BOF = False Then
            Call LoadGrid(vsf(2), rs, Array("", "0.00##"), , ils13)
        End If
        Call InheritAppendSpaceRows(2)
        
    End Select
    
    RefreshData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function InitActive() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  初始化数据，发生在窗体的Active事件
    '返回:  True        成功
    '       False       出错
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
    zlControl.CboLocate cboDept, UserInfo.部门ID, True
    If cboDept.ListCount > 0 And cboDept.ListIndex = -1 Then cboDept.ListIndex = 0
        
    InitActive = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub PrintDetail(ByVal bytMode As Byte)
    '------------------------------------------------------------------------------------------------------------------
    '功能： 输入出列表
    '------------------------------------------------------------------------------------------------------------------
    Dim objOut As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim bytR As Byte, intRow As Integer
    
    
    Dim strNo As String
    
    strNo = vsf(0).TextMatrix(vsf(0).Row, GetCol(vsf(0), "单据号"))
    If strNo = "" Then Exit Sub
    
    Call CopyGrid(vsf(1), vsfPrint)
    
    '表头
    objOut.Title.Text = "团体体检结算单明细"
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '表项

        objRow.Add "单据号：" & vsf(0).TextMatrix(vsf(0).Row, GetCol(vsf(0), "单据号"))
        'objRow.Add "结帐范围：" & mshList.TextMatrix(mshList.Row, GetColNum("开始日期")) & " 至 " & mshList.TextMatrix(mshList.Row, GetColNum("结束日期"))
        objOut.UnderAppRows.Add objRow
    
        Set objRow = New zlTabAppRow
        'objRow.Add "住院号：" & mshList.TextMatrix(mshList.Row, GetColNum("住院号"))
        'objRow.Add "姓名：" & mshList.TextMatrix(mshList.Row, GetColNum("姓名"))
        objOut.UnderAppRows.Add objRow

    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印日期：" & Format(zlDatabase.Currentdate(), "yyyy年MM月dd日")
    objOut.BelowAppRows.Add objRow
    
    '表体

    Set objOut.Body = vsfPrint
    
    If bytMode = 1 Then bytMode = zlPrintAsk(objOut)
    
    Me.Refresh
    
    If bytMode >= 1 And bytMode <= 3 Then Call zlPrintOrView1Grd(objOut, bytMode)
    
'    bytR = zlPrintAsk(objOut)
'    Me.Refresh
'    If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR

End Sub

Private Sub mnuFileBalanceDetaiExcel_Click()
    Call PrintDetail(3)
End Sub

Private Sub mnuFileBalanceDetaiPrint_Click()
            
    Call PrintDetail(1)
End Sub

Private Sub mnuFileBalanceDetaiPrintView_Click()
    Call PrintDetail(2)
End Sub

Private Sub mnuViewList_Click()
    
    Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1862", Me, "体检部门id=" & mlngDept, 0)
    
End Sub

Private Sub mnuViewSearch_Click()
    If frmBalanceFilter.ShowFilter(Me, mstrCondition) Then
        Call mnuViewRefresh_Click
    End If
End Sub

'（３）窗体及其控件的事件处理******************************************************************************************
Private Sub mobjPopMenu_MenuBeforeShow(Cancel As Boolean)
    
    If mnuEdit.Visible Then
        If mnuEditBalance.Visible Then mobjPopMenu.Add 1, mnuEditBalance.Caption, , , mnuEditBalance.Enabled
        If mnuEditBalanceCancel.Visible Then mobjPopMenu.Add 2, mnuEditBalanceCancel.Caption, , , mnuEditBalanceCancel.Enabled

    End If

End Sub

Private Sub mobjPopMenu_MenuClick(ByVal Key As Long, ByVal Caption As String)
    Select Case Key
    Case 1
        Call mnuEditBalance_Click
    Case 2
        Call mnuEditBalanceCancel_Click
    End Select
End Sub

Private Sub cboDept_Click()
    If mblnStartUp Then Exit Sub
    If mlngDept = cboDept.ItemData(cboDept.ListIndex) Then Exit Sub
    
    mlngDept = cboDept.ItemData(cboDept.ListIndex)
    Call mnuViewRefresh_Click
    
End Sub


Private Sub Form_Activate()
    
    If mblnStartUp = False Then Exit Sub

    If InitActive = False Then
        Unload Me
        Exit Sub
    End If
    DoEvents
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
    
    If Me.WindowState = 1 Then Exit Sub
    
    '处理特殊情况
    
    If imgX_S.Top > Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - 1000 Then
        imgX_S.Top = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - 1000
    End If

    With vsf(0)
        .Left = 0
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0)
        .Width = Me.ScaleWidth - .Left
        .Height = imgX_S.Top - .Top
    End With
    
    With imgX_S
        .Left = 0
        .Width = vsf(0).Width
    End With
    
    With tbs
        .Left = vsf(0).Left
        .Top = imgX_S.Top + imgX_S.Height
        .Width = vsf(0).Width
    End With
    
    With vsf(1)
        .Left = vsf(0).Left
        .Top = tbs.Top + tbs.Height
        .Width = Me.ScaleWidth - .Left
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
    End With
    
    With vsf(2)
        .Left = vsf(0).Left
        .Top = vsf(1).Top
        .Width = vsf(1).Width
        .Height = vsf(1).Height
    End With
    
    Call InheritAppendSpaceRows(0)
    Call InheritAppendSpaceRows(1)
    Call InheritAppendSpaceRows(2)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub imgX_S_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    On Error Resume Next
    
    imgX_S.Top = imgX_S.Top + Y
    
    If imgX_S.Top < 1500 Then imgX_S.Top = 1500
    If Me.Height - imgX_S.Top - imgX_S.Height < 1000 Then imgX_S.Top = Me.Height - imgX_S.Height - 1000

    Call Form_Resize
End Sub


Private Sub mnuFileBalanceBill_Click()
    
    '功能：当前收款记录重新打印一张票据
    
    Dim strNo As String
    Dim lng结帐ID As Long
    
    lng结帐ID = Val(vsf(0).RowData(vsf(0).Row))
    If lng结帐ID = 0 Then Exit Sub
    
    strNo = vsf(0).TextMatrix(vsf(0).Row, GetCol(vsf(0), "单据号"))
    If strNo = "" Then
        MsgBox "当前没有单据可以重打票据！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If RePrintBalance(strNo, Me, lng结帐ID, mbytKind) Then
        
        Call mnuViewRefresh_Click
        
    End If
End Sub

Private Sub mnuEditBalanceCancel_Click()
    Call MenuClick("结算作废")
End Sub

Private Sub mnuEditBalance_Click()
    Call MenuClick("体检结算")
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileOutExcel_Click()
    Call PrintData(3)
End Sub

Private Sub mnuFilePara_Click()
    If frmSetExpence.ShowParameter(Me) Then
        
        '重新读取参数
        
        glng结帐ID = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "共用结帐票据批次", 0)
        glngShareUseID = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "共用结帐票据批次", 0)
        
        mbytKind = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "共用结帐票据类型", 3))
        
    End If
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
                
    '保存
    musrSavePos.Row = vsf(0).Row
    musrSavePos.Col = vsf(0).Col
    
    mblnNoAllowChange = True
    
    Call ClearData("体检结算")
    Call ClearData("结算单据")
    Call ClearData("结算方式")
    
    Call RefreshData("体检结算")
    
    '恢复体检预约
    
    On Error Resume Next
    vsf(0).Row = musrSavePos.Row
    vsf(0).Col = musrSavePos.Col
    vsf(0).ShowCell vsf(0).Row, vsf(0).Col
    Call SelectRow(vsf(0), 0, vsf(0).Row)
    On Error GoTo 0
    
    Call RefreshData("结算单据")
    Call RefreshData("结算方式")
    
    mblnNoAllowChange = False
    
    Call AdjustEnableState
    Call RefreshStateInfo
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

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Key
    Case "预览"
        Call mnuFilePrintView_Click
    Case "打印"
        Call mnuFilePrint_Click
    
    Case "结算"
        Call mnuEditBalance_Click
    
    Case "作废"
        Call mnuEditBalanceCancel_Click
    Case "过滤"
        Call mnuViewSearch_Click
    Case "帮助"
        Call mnuHelpTopic_Click
    Case "退出"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool
End Sub

Private Sub tbs_Click()
    
    vsf(1).Visible = False
    vsf(2).Visible = False
    
    vsf(tbs.SelectedItem.Index).Visible = True
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow = NewRow Then Exit Sub
    If mblnNoAllowChange Then Exit Sub
    
    If Index = 0 Then
        Call SelectRow(vsf(Index), OldRow, NewRow)
        
        Call RefreshData("结算单据")
        Call RefreshData("结算方式")
        
        Call AdjustEnableState
    End If
    
End Sub

Private Sub vsf_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call InheritAppendSpaceRows(Index)
End Sub

Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Call InheritAppendSpaceRows(Index)
End Sub

Private Sub vsf_BeforeUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Index = 0 And Col = 0 Then Cancel = True
End Sub

Private Sub vsf_GotFocus(Index As Integer)
    vsf(Index).BackColorSel = COLOR.焦点
    If Index = 0 Then Call SelectRow(vsf(Index), 1, vsf(Index).Row)
End Sub

Private Sub vsf_LostFocus(Index As Integer)
    vsf(Index).BackColorSel = COLOR.非焦点
    If Index = 0 Then Call SelectRow(vsf(Index), 1, vsf(Index).Row)
End Sub

Private Sub vsf_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    If Index <> 0 Then Exit Sub
    
    Call SendLMouseButton(vsf(Index).hWnd, X, Y)
    
    Set mobjPopMenu = New clsPopMenu
    Call mobjPopMenu.ShowPopupMenuByCursor
    
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub


