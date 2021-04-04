VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frm付款管理 
   Caption         =   "付款管理"
   ClientHeight    =   6375
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9570
   Icon            =   "frm付款管理.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6375
   ScaleWidth      =   9570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.TabStrip tabSelect 
      Height          =   300
      Left            =   15
      TabIndex        =   6
      Tag             =   "1"
      Top             =   765
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   529
      MultiRow        =   -1  'True
      Style           =   2
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "所有付款"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "一般付款"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   " 预付款 "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "计划付款"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "拒绝付款"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsCold 
      Left            =   6150
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm付款管理.frx":08CA
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm付款管理.frx":0AEA
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm付款管理.frx":0D0A
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm付款管理.frx":0F26
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm付款管理.frx":1146
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm付款管理.frx":1366
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm付款管理.frx":1582
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm付款管理.frx":179E
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm付款管理.frx":19B8
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm付款管理.frx":1B12
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm付款管理.frx":1D2E
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm付款管理.frx":1F4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm付款管理.frx":26C8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHot 
      Left            =   6750
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm付款管理.frx":2E42
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm付款管理.frx":3062
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm付款管理.frx":3282
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm付款管理.frx":349E
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm付款管理.frx":36BE
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm付款管理.frx":38DE
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm付款管理.frx":3AFA
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm付款管理.frx":3D16
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm付款管理.frx":3F30
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm付款管理.frx":408A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm付款管理.frx":42AA
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm付款管理.frx":44CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm付款管理.frx":4C44
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   6015
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm付款管理.frx":53BE
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11800
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
   Begin ComCtl3.CoolBar cbrTool 
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   1376
      BandCount       =   2
      BandBorders     =   0   'False
      _CBWidth        =   9570
      _CBHeight       =   780
      _Version        =   "6.7.8988"
      Child1          =   "tlbThis"
      MinHeight1      =   720
      Width1          =   11040
      NewRow1         =   0   'False
      MinHeight2      =   0
      NewRow2         =   0   'False
      BandStyle2      =   1
      Begin MSComctlLib.Toolbar tlbThis 
         Height          =   720
         Left            =   165
         TabIndex        =   1
         Top             =   30
         Width           =   9315
         _ExtentX        =   16431
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
               Key             =   "CheckSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "预审"
               Key             =   "Check"
               Description     =   "预审"
               Object.ToolTipText     =   "预审"
               Object.Tag             =   "预审"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "回退"
               Key             =   "CheckBack"
               Description     =   "回退"
               Object.ToolTipText     =   "回退"
               Object.Tag             =   "回退"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "EditSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "审核"
               Key             =   "Verify"
               Description     =   "审核"
               Object.ToolTipText     =   "审核"
               Object.Tag             =   "审核"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "冲销"
               Key             =   "Strike"
               Description     =   "冲销"
               Object.ToolTipText     =   "冲销"
               Object.Tag             =   "冲销"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "VerifySeparate"
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "过滤"
               Key             =   "Search"
               Description     =   "过滤"
               Object.ToolTipText     =   "过滤"
               Object.Tag             =   "过滤"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "刷新"
               Key             =   "Refresh"
               Description     =   "刷新"
               Object.ToolTipText     =   "刷新"
               Object.Tag             =   "刷新"
               ImageIndex      =   9
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
               ImageIndex      =   10
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Exit"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   11
            EndProperty
         EndProperty
         MouseIcon       =   "frm付款管理.frx":5C52
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsDetail 
      Height          =   2535
      Left            =   -15
      TabIndex        =   8
      Top             =   3165
      Width           =   4560
      _cx             =   8043
      _cy             =   4471
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
      BackColorBkg    =   -2147483648
      BackColorAlternate=   -2147483628
      GridColor       =   12632256
      GridColorFixed  =   0
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
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
      Cols            =   18
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm付款管理.frx":5F6C
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
   Begin VSFlex8Ctl.VSFlexGrid vsList 
      Height          =   1650
      Left            =   -15
      TabIndex        =   7
      Top             =   1110
      Width           =   9480
      _cx             =   16722
      _cy             =   2910
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
      BackColorBkg    =   -2147483648
      BackColorAlternate=   -2147483628
      GridColor       =   12632256
      GridColorFixed  =   0
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
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
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm付款管理.frx":61D1
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
   Begin VSFlex8Ctl.VSFlexGrid vsAddition 
      Height          =   2535
      Left            =   4785
      TabIndex        =   9
      Top             =   3150
      Width           =   4560
      _cx             =   8043
      _cy             =   4471
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
      BackColorBkg    =   -2147483648
      BackColorAlternate=   -2147483628
      GridColor       =   12632256
      GridColorFixed  =   0
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
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
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm付款管理.frx":63D2
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
   Begin VB.Label lblHsc_s 
      Height          =   2865
      Left            =   5535
      MousePointer    =   9  'Size W E
      TabIndex        =   5
      Top             =   3150
      Width           =   60
   End
   Begin VB.Label lblVsc_s 
      BackStyle       =   0  'Transparent
      Height          =   360
      Left            =   2520
      MousePointer    =   7  'Size N S
      TabIndex        =   4
      Top             =   2775
      Width           =   1425
   End
   Begin VB.Label lblRange 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "查询范围:1999年8月12日至1999年9月12日"
      Height          =   180
      Left            =   30
      TabIndex        =   3
      Top             =   2895
      Width           =   3330
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
      Begin VB.Menu mnuFileLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnuFileSp 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileBillPrePrint 
         Caption         =   "单据预览(&V)"
      End
      Begin VB.Menu mnuFileBillPrint 
         Caption         =   "单据打印(&B)"
      End
      Begin VB.Menu mnuFileLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnueditu 
         Caption         =   "参数设置"
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
         Begin VB.Menu mnuEditAddPayment 
            Caption         =   "付款单(&P)"
         End
         Begin VB.Menu mnuEditMultAdd 
            Caption         =   "批量付款(&M)"
         End
         Begin VB.Menu mnuEditAddScheme 
            Caption         =   "计划付款单(&S)"
         End
         Begin VB.Menu mnuEditAddImprest 
            Caption         =   "预付款单(&I)"
         End
         Begin VB.Menu mnuEditAddSign 
            Caption         =   "标记付款单(&B)"
         End
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "修改(&M)"
      End
      Begin VB.Menu mnuEditDel 
         Caption         =   "删除(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditLine0 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditCheck 
         Caption         =   "预审(&P)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditCheckBack 
         Caption         =   "回退(&B)"
         Visible         =   0   'False
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
      Begin VB.Menu mnuViewSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewSavePrint 
         Caption         =   "存盘打印(&C)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewVerifyPrint 
         Caption         =   "审核打印(&V)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewSearch 
         Caption         =   "过滤(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewLine2 
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
         Caption         =   "&Web上的中联"
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
Attribute VB_Name = "frm付款管理"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private msngDownX As Single, msngDownY As Single
Private mstrFilter As String  '过滤条件
Private mstrPrivs As String
Private mstr类型 As String      '供应商类型
Private mlngModule As Long
Private mblnFirst As Boolean
Private mstrOthers() As String    '0-记录状态,1-开始单号,2-结束单号,3-供应商ID,4-审核人,5-填制人,6-开始发票号,7-结束发票号,8-品名
Private mdtStartDate As Date
Private mdtEndDate As Date
Private mdtVrfyStartDate As Date
Private mdtVrfyEndDate As Date
Private mint物资Flag As Integer
Private mint设备Flag As Integer
Private mbln预审 As Boolean     'True：预审；  False：不预审
Private mint显示单位 As Integer '0：最小单位；  1：最大单位

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    Call zlControl.IsCtrlSetFocus(vsList)
    Call vsList_GotFocus
End Sub

Private Sub Form_Load()
'    Dim strStart As String, strEnd As String
    Dim strReg As String
    Dim strOthers(0 To 9) As String     '0-记录状态,1-开始单号,2-结束单号,3-供应商ID,4-审核人,5-填制人,6-开始发票号,7-结束发票号,8-品名
    mstrOthers = strOthers
    '问题24925 by lesfeng 2010-02-08
    mint物资Flag = 0
    mint设备Flag = 0
    
    mblnFirst = True
    '权限控制
    mlngModule = glngModul: mstrPrivs = gstrPrivs
    mstr类型 = "0000"
    Call 权限控制
   '恢复参数
    mnuViewSavePrint.Checked = IIf(Val(zlDatabase.GetPara("存盘打印", glngSys, mlngModule)) = 1, 1, 0) = 1
    mnuViewVerifyPrint.Checked = IIf(Val(zlDatabase.GetPara("审核打印", glngSys, mlngModule)) = 1, 1, 0) = 1
    
    mint显示单位 = Val(zlDatabase.GetPara("显示单位选择", glngSys, mlngModule))
    
    '预审
    mbln预审 = IIf(Val(zlDatabase.GetPara("一般付款需要经过预审", glngSys, mlngModule)) = 1, True, False)
    If mbln预审 Then
        mnuEditLine0.Visible = True
        mnuEditCheck.Visible = True
        mnuEditCheckBack.Visible = True
        tlbThis.Buttons("CheckSeparate").Visible = True
        tlbThis.Buttons("Check").Visible = True
        tlbThis.Buttons("CheckBack").Visible = True
    End If
    
    'by lesfeng 2009-12-2 性能优化
    mdtStartDate = Format(DateAdd("d", -15, zlDatabase.Currentdate), "yyyy-MM-dd")
    mdtEndDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    
    mdtVrfyStartDate = "1901-01-01"
    mdtVrfyEndDate = "1901-01-01"
    
    lblRange.Caption = "查询范围:" & Format(mdtStartDate, "yyyy年MM月dd日") & "至" & Format(mdtEndDate, "yyyy年MM月dd日")
    mstrFilter = " AND A.记录状态 = 1 And A.审核日期 is Null And A.填制日期 Between [1] And [2]"
    
'    strStart = Format(DateAdd("d", -7, zlDatabase.Currentdate), "yyyy-MM-dd")
'    strEnd = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    
'    lblRange.Caption = "查询范围:" & Format(strStart, "yyyy年MM月dd日") & "至" & Format(strEnd, "yyyy年MM月dd日")
    
'    mstrFilter = " AND A.记录状态 = 1 And A.审核日期 is Null And A.填制日期 Between To_Date('" & strStart & " 00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & strEnd & " 23:59:59','YYYY-MM-DD HH24:MI:SS')"
    
    RestoreWinState Me, App.ProductName
    
    '2006-04-25:刘兴宏,统一增加报表发布到模块的功能
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
    
    
    '初始化网格控件
    Call initGrid
    Call GetHeadData
End Sub

Private Sub mnuEditAddSign_Click()
    '问题27930 by lesfeng 2010-03-23
    Dim blnReturn As Boolean
    
    If InStr(1, mstrPrivs, ";标记付款;") = 0 Then Exit Sub
    
    frm付款编辑.ShowCard Me, g新增, mstrPrivs, , , , blnReturn, 1
    If blnReturn = False Then Exit Sub
    mnuViewRefresh_Click
End Sub

Private Sub mnuEditCheck_Click()
'预审
    Dim strNO As String
    Dim blnSuccess As Boolean
    Dim bln付款 As Boolean
    Dim str标记 As String
    Dim int标记 As Integer
    
    strNO = vsList.TextMatrix(vsList.Row, vsList.ColIndex("单据号"))
    If vsList.TextMatrix(vsList.Row, vsList.ColIndex("预付款")) <> "1" And (Val(vsDetail.TextMatrix(1, vsDetail.ColIndex("计划序号"))) = 0 Or GetMultiPayment(strNO) = True) Then
        '预审
        str标记 = Trim(vsList.TextMatrix(vsList.Row, vsList.ColIndex("拒付标记")))
        
        frm付款编辑.ShowCard Me, g预审, mstrPrivs, strNO, , , blnSuccess, IIf(str标记 = "标记", 1, 0)
        If blnSuccess = False Then Exit Sub
        mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditCheckBack_Click()
'预审回退
    Dim strTmp As String, strNO As String
    Dim intRow As Integer
    
    If mnuEditCheckBack.Visible = False Then Exit Sub
    If vsList.Rows <= 1 Then Exit Sub
    
    On Error GoTo errHandle
    strNO = vsList.TextMatrix(vsList.Row, vsList.ColIndex("单据号"))
    If TestCheck(1, strNO) Then
        MsgBox "该单据已被删除！", vbInformation, gstrSysName
        intRow = vsList.Row
        mnuViewRefresh_Click
        Exit Sub
    End If
    If TestCheck(2, strNO) Then
        MsgBox "该单据已被审核！", vbInformation, gstrSysName
        intRow = vsList.Row
        mnuViewRefresh_Click
        Exit Sub
    End If
    
    If MsgBox("是否将单据预审回退？", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    strTmp = "zl_付款管理_CancelCheck('" & strNO & "')"
    Call zlDatabase.ExecuteProcedure(strTmp, "预审回退")
    mnuViewRefresh_Click
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub mnuEditMultAdd_Click()
    '功能:批量增加付款单
    If frm批量付款条件设置.ShowCard(Me, mstrPrivs) = False Then Exit Sub
    '刷新单据
    Call mnuViewRefresh_Click
End Sub

Private Sub mnueditu_Click()
    Call frm付款参数设置.设置参数(Me, glngModul, mstrPrivs)
    '预审
    mbln预审 = IIf(Val(zlDatabase.GetPara("一般付款需要经过预审", glngSys, mlngModule)) = 1, True, False)
    mnuEditLine0.Visible = mbln预审
    mnuEditCheck.Visible = mbln预审
    mnuEditCheckBack.Visible = mbln预审
    tlbThis.Buttons("CheckSeparate").Visible = mbln预审
    tlbThis.Buttons("Check").Visible = mbln预审
    tlbThis.Buttons("CheckBack").Visible = mbln预审
    Call Form_Activate
    mint显示单位 = Val(zlDatabase.GetPara("显示单位选择", glngSys, mlngModule))
    
    With vsList
        If tabSelect.SelectedItem.Index = 1 Or tabSelect.SelectedItem.Index = 2 Then
            .ColHidden(.ColIndex("预审人")) = Not mbln预审
            .ColHidden(.ColIndex("预审日期")) = Not mbln预审
        Else
            .ColHidden(.ColIndex("预审人")) = True
            .ColHidden(.ColIndex("预审日期")) = True
        End If
    End With
    Call vsList_Click
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Dim strNO As String, lng预付款 As Long, lng记录状态 As Long, lng供应商ID As Long
    
    strNO = vsList.TextMatrix(vsList.Row, vsList.ColIndex("单据号"))
    lng预付款 = Val(vsList.TextMatrix(vsList.Row, vsList.ColIndex("预付款")))
    lng记录状态 = Val(vsList.TextMatrix(vsList.Row, vsList.ColIndex("记录状态")))
    lng供应商ID = Val(vsList.Cell(flexcpData, vsList.Row, vsList.ColIndex("供应商名称")))
    
    '2006-04-25:刘兴宏:增加自定义报表发布到模块的功能
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, "NO=" & strNO, "预付款=" & lng预付款, "记录状态=" & lng记录状态, "供应商=" & lng供应商ID)
End Sub

Private Sub initGrid()
    '-----------------------------------------------------------------------------------------------------------
    '功能:初始化网格控件的默认属性
    '编制:刘兴洪
    '日期:2009-02-11 11:35:41
    '-----------------------------------------------------------------------------------------------------------
    Call zl_vsGrid_Para_Restore(mlngModule, vsList, Me.Caption, "付款表头列表", True)
    Call zl_vsGrid_Para_Restore(mlngModule, vsDetail, Me.Caption, "付款明细列表", True)
    Call zl_vsGrid_Para_Restore(mlngModule, vsAddition, Me.Caption, "付款方式列表", True)
    Call vsDetail_LostFocus
    Call vsAddition_LostFocus
    Call vsList_LostFocus
    
    With vsList
        .Clear 1
        .Rows = 2
        .ColHidden(.ColIndex("记录状态")) = True: .ColWidth(.ColIndex("记录状态")) = 0
        .ColHidden(.ColIndex("预付款")) = True: .ColWidth(.ColIndex("预付款")) = 0
        .ColHidden(.ColIndex("预审人")) = Not mbln预审
        .ColHidden(.ColIndex("预审日期")) = Not mbln预审
        
        'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
        .ColData(.ColIndex("单据号")) = "1||0"
        .ColData(.ColIndex("预付款")) = "-1||0"
        .ColData(.ColIndex("记录状态")) = "-1||0"
        '问题27930 by lesfeng 2010-03-23
        .ColData(.ColIndex("拒付标记")) = "1||0"
    End With
    With vsDetail
        .Clear 1
        .Rows = 2
        .ColData(.ColIndex("品名")) = "1||0"
        .ColData(.ColIndex("入库单号")) = "1||0"
        .ColData(.ColIndex("发票号")) = "1||0"
        .ColData(.ColIndex("发票金额")) = "1||0"
        .ColHidden(.ColIndex("预审")) = mbln预审
    End With
    With vsAddition
        .Clear 1
        .Rows = 2
        .ColData(.ColIndex("结算方式")) = "1||0"
    End With
End Sub

Private Sub GetHeadData()
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取头数据
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-02-11 11:40:03
    '-----------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strWhere As String, i As Long, lngRow As Long, str类型 As String, str权限 As String
    Dim rsTemp As New ADODB.Recordset
    Dim strStore As String
    
    Err = 0: On Error GoTo ErrHand:
    
    str类型 = ""
    For i = 1 To Len(mstr类型)
        If Mid(mstr类型, i, 1) = 1 Then str类型 = str类型 & " or substr(b.类型," & i & ",1)=1"
    Next
    If str类型 <> "" Then str类型 = " And (" & Mid(str类型, 4) & ") "
    
    str权限 = " and " & Get分类权限(mstrPrivs, "b.")
    
    
    Call zlCommFun.ShowFlash("正在搜索付款记录,请稍候 ...", Me)
    DoEvents
    '问题27930 by lesfeng 2010-03-23
    Screen.MousePointer = vbHourglass
    Select Case tabSelect.SelectedItem.Index
        Case 1
            strWhere = ""
        Case 2
            strWhere = "" & _
                " And A.预付款<>1 And A.拒付标志<>1 And " & _
                " A.付款序号 Not In (Select Distinct 付款序号 " & _
                "                    From 应付记录 " & _
                "                    Where 记录性质=-1 And 付款序号 Is Not Null)"
        Case 3
            strWhere = " And A.预付款=1"
        Case 4
            strWhere = "" & _
                " And A.拒付标志 = 0 And ( A.付款序号 In (Select Distinct 付款序号 " & _
                "                    From 应付记录 " & _
                "                    Where 记录性质=-1 And 付款序号 Is Not Null)  and a.预付款 <>1) "
        Case 5
            strWhere = "" & _
                " And A.拒付标志 = 1 "
    End Select
    
    strStore = ",(Select NO, f_List2str(Cast(Collect(名称) As t_Strlist)) 来源库房 " & _
               "  From (Select Distinct a.No, c.名称 " & _
               "        From 付款记录 A, 应付记录 B, 部门表 C " & _
               "        Where a.付款序号 = b.付款序号 And b.库房id = c.Id And b.库房id Is Not Null " & mstrFilter & _
               "        Order by c.名称) " & _
               "  Group By NO ) C "
    
    strSQL = "" & _
        "   SELECT  a.no as 单据号,b.id as 供应商ID, b.名称 as 供应商名称,nvl(预付款,0) as 预付款 ," & _
        "           ltrim(to_char(SUM (a.金额),'9999999999999990.00')) AS 付款金额, " & _
        "           a.填制人 AS 申请人,TO_CHAR (min(a.填制日期), 'yyyy-MM-dd') AS 申请日期," & _
        "           a.预审人,TO_CHAR (min(a.预审日期), 'yyyy-MM-dd') AS 预审日期," & _
        "           a.审核人,TO_CHAR (min(a.审核日期), 'yyyy-MM-dd') AS 审核日期," & _
        "           decode(a.拒付标志,1,'拒付','正常') as 拒付标记,a.记录状态,max(来源库房) 来源库房, a.摘要 " & _
        "   FROM 付款记录 a, 供应商 b " & _
        strStore & _
        "   Where a.单位id = b.id and a.NO=c.NO(+) " & zl_获取站点限制(True, "b") & "  " & str类型 & strWhere & mstrFilter & str权限 & _
        "   GROUP BY a.no,b.id,b.名称,nvl(预付款,0),a.填制人,a.预审人,a.审核人,'',a.记录状态, '',a.拒付标志, a.摘要 " & _
        "   ORDER BY a.no desc "
    'by lesfeng 2009-12-2 性能优化
    'mstrOthers(0 To 9) '0-记录状态,1-开始单号,2-结束单号,3-供应商ID,4-审核人,5-填制人,6-开始发票号,7-结束发票号,8-品名,9-库房ID
    '参数范围: 1-开始填制日期,2-结束填制日期
    '          3-开始审核日期,4-结束审核日期
    '          5-开始单号,6-结束单号,7-供应商ID,8-审核人,9-填制人,10-开始发票号,11-结束发票号,12-品名,13-库房ID
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(Format(mdtStartDate, "yyyy-mm-dd") & " 00:00:00"), CDate(Format(mdtEndDate, "yyyy-mm-dd") & " 23:59:59"), _
                     CDate(Format(mdtVrfyStartDate, "yyyy-mm-dd") & " 00:00:00"), CDate(Format(mdtVrfyEndDate, "yyyy-mm-dd") & " 23:59:59"), _
                     mstrOthers(1), mstrOthers(2), Val(mstrOthers(3)), mstrOthers(4), mstrOthers(5), mstrOthers(6), mstrOthers(7), mstrOthers(8), _
                     Val(mstrOthers(9)))
    With vsList
        .Redraw = flexRDNone
        .Rows = 2
        .Clear 1
        .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
        .Cell(flexcpForeColor, 1, 0, 1, .Cols - 1) = .ForeColor
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        
        If tabSelect.SelectedItem.Index = 1 Or tabSelect.SelectedItem.Index = 2 Then
            .ColHidden(.ColIndex("预审人")) = Not mbln预审
            .ColHidden(.ColIndex("预审日期")) = Not mbln预审
        Else
            .ColHidden(.ColIndex("预审人")) = True
            .ColHidden(.ColIndex("预审日期")) = True
        End If
        
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("单据号")) = Nvl(rsTemp!单据号)
            .TextMatrix(lngRow, .ColIndex("供应商名称")) = Nvl(rsTemp!供应商名称)
            .Cell(flexcpData, lngRow, .ColIndex("供应商名称")) = Nvl(rsTemp!供应商ID)
            .TextMatrix(lngRow, .ColIndex("预付款")) = Nvl(rsTemp!预付款)
            .TextMatrix(lngRow, .ColIndex("付款金额")) = Nvl(rsTemp!付款金额)
            .TextMatrix(lngRow, .ColIndex("申请人")) = Nvl(rsTemp!申请人)
            .TextMatrix(lngRow, .ColIndex("申请日期")) = Nvl(rsTemp!申请日期)
            .TextMatrix(lngRow, .ColIndex("预审人")) = Nvl(rsTemp!预审人)
            .TextMatrix(lngRow, .ColIndex("预审日期")) = Nvl(rsTemp!预审日期)
            .TextMatrix(lngRow, .ColIndex("审核人")) = Nvl(rsTemp!审核人)
            .TextMatrix(lngRow, .ColIndex("审核日期")) = Nvl(rsTemp!审核日期)
            .TextMatrix(lngRow, .ColIndex("记录状态")) = Nvl(rsTemp!记录状态)
            .TextMatrix(lngRow, .ColIndex("拒付标记")) = Nvl(rsTemp!拒付标记)
            .TextMatrix(lngRow, .ColIndex("来源库房")) = Nvl(rsTemp!来源库房)
            .TextMatrix(lngRow, .ColIndex("摘要")) = Nvl(rsTemp!摘要)
            '设置相关表格颜色
            If Val(Nvl(rsTemp!记录状态)) = 3 Then
                .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = &H80000001
            ElseIf Val(Nvl(rsTemp!记录状态)) = 2 Then
                .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = &HFF
            End If
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        .Row = 1: .Col = .ColIndex("单据号")
    End With
    
    Full单据明细
    Full付款明细
    Call SetEnabled
    Call zlCommFun.StopFlash
    vsList.Redraw = flexRDBuffered
    Screen.MousePointer = vbDefault
    staThis.Panels(2).Text = "当前共有" & rsTemp.RecordCount & "张单据"
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    vsList.Redraw = flexRDBuffered
    If ErrCenter = 1 Then
        Resume
    End If
    Call zlCommFun.StopFlash
    staThis.Panels(2).Text = "当前共有" & 0 & "张单据"
End Sub

Private Sub Full单据明细()
    '-----------------------------------------------------------------------------------------------------------
    '功能:填制单据明细数据
    '返回:
    '编制:刘兴洪
    '日期:2009-02-11 11:58:12
    '-----------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As New ADODB.Recordset, lng付款序号 As Long, int状态 As Integer
    Dim strNO As String, int预付 As Integer, lngRow As Long
    Dim int物资Flag As Integer, int设备Flag As Integer
    Dim strTemp As String
    
    Err = 0: On Error GoTo ErrHand:
    With vsList
        int状态 = Val(.TextMatrix(.Row, .ColIndex("记录状态")))
        strNO = Trim(.TextMatrix(.Row, .ColIndex("单据号")))
        int预付 = Val(.TextMatrix(.Row, .ColIndex("预付款")))
    End With
    
    If strNO = "" Or int状态 = 2 Or int预付 = 1 Then
        vsDetail.Clear 1: vsDetail.Rows = 2
        vsDetail.Cell(flexcpData, 1, 0, 1, vsDetail.Cols - 1) = ""
        Exit Sub
    End If
    
    
    strSQL = " Select 付款序号 From 付款记录 Where NO=[1] and 记录状态 in (1,3) "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If rsTemp.EOF Then
        lng付款序号 = 0
    Else
        lng付款序号 = Val(Nvl(rsTemp!付款序号))
    End If
    '问题24925 by lesfeng 2010-02-08
    strTemp = "1,4,5"
    If GetShareSys(400) Then
        int物资Flag = 1
    Else
        int物资Flag = 0
        strTemp = strTemp & ",2"
    End If
    If GetShareSys(600) Then
        int设备Flag = 1
    Else
        int设备Flag = 0
        strTemp = strTemp & ",3"
    End If
    '药品与卫材其他，以及没有共享的物资、设备 需要增加对设备库存、物资库存访问权限，同时处理已经分配权限的情况
    strSQL = IIf(mbln预审, "select a.*, decode(b.预审, 1, '√', null) 预审 from (", "") & _
    "   Select Max(A.入库单据号) 入库单据号, Max(B1.审核人) As 审核人, To_Char(Max(B1.审核日期), 'yyyy-mm-dd') As 审核日期, " & _
    "          Max(A.ID) As ID, decode(a.记录性质,2,null,a.计划序号) 计划序号, A.发票号, " & _
    "          To_Char(Sum(Nvl(case when a.发票金额 = a.计划金额 or a.计划金额 is null and a.记录性质<>2 then a.发票金额 else a.计划金额 end, 0)), '999999999990.99') As 发票金额, " & _
    "          To_Char(Sum(Nvl(A.计划金额, 0)), '999999999990.99') As 计划金额, To_Char(A.计划日期, 'yyyy-mm-dd') As 计划日期, " & _
    "          Max(A.品名) 品名, Max(A.规格) 规格, Max(A.产地) 产地, Max(A.批号) 批号," & _
    IIf(mint显示单位 = 1, "decode(a.系统标识,1,max(e.药库单位),5,max(f.包装单位),max(a.计量单位)) 计量单位, To_Char(Round(Sum(Nvl(A.数量, 0)) / decode(a.系统标识,1,max(e.药库包装),5,max(f.换算系数),1), 4), '999999999990.9999') As 数量,", _
                          "max(A.计量单位) 计量单位, To_Char(Sum(Nvl(A.数量, 0)), '999999999990.9999') As 数量,") & _
    "          To_Char(Max(A.采购价), '999999999990.9999') As 采购价, " & _
    "          To_Char(Sum(Nvl(A.采购金额, 0)), '999999999990.9999') As 采购金额," & _
    IIf(mint显示单位 = 1, "To_Char(round(Sum(Nvl(D.库存数量, 0)) / decode(a.系统标识,1,max(e.药库包装),5,max(f.换算系数),1), 4), '999999999990.9999')", _
                          "To_Char(Sum(Nvl(D.库存数量, 0)), '999999999990.9999')") & " As 库存数量" & _
    "   From 应付记录 A, " & _
    "        (Select B.ID, Max(审核人) As 审核人, Max(审核日期) As 审核日期 " & _
    "          From 应付记录 B, (Select Distinct ID From 应付记录 Where 付款序号 = [1]) C " & _
    "          Where B.ID = C.ID Group By B.ID) B1," & _
    "        (Select 药品id,sum(可用数量) As 可用数量,Sum(实际数量) As 库存数量,Sum(实际金额) As 实际金额,上次批号,上次供应商id " & _
    "          From 药品库存 Group By 药品id,上次批号,上次供应商id) D" & _
    IIf(mint显示单位 = 1, ",药品规格 E, 材料特性 F ", "") & _
    "   Where A.付款序号 = [1] And A.ID = B1.ID And A.项目id = D.药品id(+) And A.批号= D.上次批号(+) And A.单位id = D.上次供应商id(+) " & _
    "     And nvl(A.系统标识,4) In (" & strTemp & ")" & _
    IIf(mint显示单位 = 1, " And a.项目id=e.药品id(+) and a.项目id=f.材料id(+) ", "") & _
    "   Group By A.系统标识, A.入库单据号, A.项目id, A.序号, decode(a.记录性质,2,null,a.计划序号), A.计划日期, A.发票号"
    '物资部分
    If int物资Flag = 1 Then
        strSQL = strSQL & " Union All " & _
        "   Select Max(A.入库单据号) 入库单据号, Max(B1.审核人) As 审核人, To_Char(Max(B1.审核日期), 'yyyy-mm-dd') As 审核日期, " & _
        "          Max(A.ID) As ID, decode(a.记录性质,2,null,a.计划序号) 计划序号, A.发票号, " & _
        "          To_Char(Sum(Nvl(case when a.发票金额 = a.计划金额 or a.计划金额 is null and a.记录性质<>2 then a.发票金额 else a.计划金额 end, 0)), '999999999990.99') As 发票金额, " & _
        "          To_Char(Sum(Nvl(A.计划金额, 0)), '999999999990.99') As 计划金额, To_Char(A.计划日期, 'yyyy-mm-dd') As 计划日期, " & _
        "          Max(A.品名) 品名, Max(A.规格) 规格, Max(A.产地) 产地, Max(A.批号) 批号, " & _
        IIf(mint显示单位 = 1, "Max(E.包装单位) 计量单位, To_Char(round(Sum(Nvl(A.数量, 0)) / max(e.换算系数), 4), '999999999990.9999') As 数量, ", _
                              "Max(A.计量单位) 计量单位, To_Char(Sum(Nvl(A.数量, 0)), '999999999990.9999') As 数量, ") & _
        "          To_Char(Max(A.采购价), '999999999990.9999') As 采购价, " & _
        "          To_Char(Sum(Nvl(A.采购金额, 0)), '999999999990.9999') As 采购金额," & _
        IIf(mint显示单位 = 1, "To_Char(round(Sum(Nvl(D.库存数量, 0)) / max(e.换算系数), 4), '999999999990.9999') As 库存数量 ", _
                              "To_Char(Sum(Nvl(D.库存数量, 0)), '999999999990.9999') As 库存数量 ") & _
        "   From 应付记录 A, " & _
        "        (Select B.ID, Max(审核人) As 审核人, Max(审核日期) As 审核日期 " & _
        "          From 应付记录 B, (Select Distinct ID From 应付记录 Where 付款序号 = [1]) C " & _
        "          Where B.ID = C.ID Group By B.ID) B1," & _
        "        (Select 物资id,sum(可用数量) As 可用数量,Sum(实际数量) As 库存数量,Sum(实际金额) As 实际金额,上次批号,上次供应商id " & _
        "          From 物资库存 Group By 物资id,上次批号,上次供应商id) D" & _
        IIf(mint显示单位 = 1, ",物资目录 E ", "") & _
        "   Where A.付款序号 = [1] And A.ID = B1.ID And A.项目id = D.物资id(+) And A.批号= D.上次批号(+) And A.单位id = D.上次供应商id(+) " & _
        "     And A.系统标识 = 2 " & _
        IIf(mint显示单位 = 1, " And a.项目id=e.ID ", "") & _
        "   Group By A.系统标识, A.入库单据号, A.项目id, A.序号, decode(a.记录性质,2,null,a.计划序号), A.计划日期, A.发票号"
    End If
    '设备部分
    If int设备Flag = 1 Then
        strSQL = strSQL & " Union All " & _
        "   Select Max(A.入库单据号) 入库单据号, Max(B1.审核人) As 审核人, To_Char(Max(B1.审核日期), 'yyyy-mm-dd') As 审核日期, " & _
        "          Max(A.ID) As ID, decode(a.记录性质,2,null,a.计划序号) 计划序号, A.发票号, " & _
        "          To_Char(Sum(Nvl(case when a.发票金额 = a.计划金额 or a.计划金额 is null and a.记录性质<>2 then a.发票金额 else a.计划金额 end, 0)), '999999999990.99') As 发票金额, " & _
        "          To_Char(Sum(Nvl(A.计划金额, 0)), '999999999990.99') As 计划金额, To_Char(A.计划日期, 'yyyy-mm-dd') As 计划日期, " & _
        "          Max(A.品名) 品名, Max(A.规格) 规格, Max(A.产地) 产地, Max(A.批号) 批号, Max(A.计量单位) 计量单位, " & _
        "          To_Char(Sum(Nvl(A.数量, 0)), '999999999990.9999') As 数量, To_Char(Max(A.采购价), '999999999990.9999') As 采购价, " & _
        "          To_Char(Sum(Nvl(A.采购金额, 0)), '999999999990.9999') As 采购金额,To_Char(Sum(Nvl(D.库存数量, 0)), '999999999990.9999') As 库存数量 " & _
        "   From 应付记录 A, " & _
        "        (Select B.ID, Max(审核人) As 审核人, Max(审核日期) As 审核日期 " & _
        "          From 应付记录 B, (Select Distinct ID From 应付记录 Where 付款序号 = [1]) C " & _
        "          Where B.ID = C.ID Group By B.ID) B1," & _
        "        (Select 设备id,sum(可用数量) As 可用数量,Sum(实际数量) As 库存数量,Sum(实际金额) As 实际金额,批次,上次供应商id " & _
        "          From 设备库存 Group By 设备id,批次,上次供应商id) D" & _
        "   Where A.付款序号 = [1] And A.ID = B1.ID And A.项目id = D.设备id(+) And A.批号= D.批次(+) And A.单位id = D.上次供应商id(+) " & _
        "     And A.系统标识 = 3 " & _
        "   Group By A.系统标识, A.入库单据号, A.项目id, A.序号, decode(a.记录性质,2,null,a.计划序号), A.计划日期, A.发票号"
    End If
    
    If mbln预审 Then
        strSQL = strSQL & ") A, 应付记录 B where a.ID=b.ID(+) " 'and b.记录性质(+) <> 2
        If lng付款序号 <> 0 Then
            strSQL = strSQL & " and b.付款序号(+) = [1] "
        End If
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng付款序号)
    
    With vsDetail
        If mbln预审 And tabSelect.SelectedItem.Index = 2 Then
            .ColHidden(.ColIndex("预审")) = False
        Else
            .ColHidden(.ColIndex("预审")) = True
        End If
        .Redraw = flexRDNone
        .Rows = 2: .Clear 1
        .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            If mbln预审 Then
                .TextMatrix(lngRow, .ColIndex("预审")) = Nvl(rsTemp!预审)
            End If
            .TextMatrix(lngRow, .ColIndex("入库单号")) = Nvl(rsTemp!入库单据号)
            .Cell(flexcpData, lngRow, .ColIndex("入库单号")) = Nvl(rsTemp!ID)
            .TextMatrix(lngRow, .ColIndex("审核人")) = Nvl(rsTemp!审核人)
            .TextMatrix(lngRow, .ColIndex("审核日期")) = Nvl(rsTemp!审核日期)
            .TextMatrix(lngRow, .ColIndex("发票号")) = Nvl(rsTemp!发票号)
            .TextMatrix(lngRow, .ColIndex("发票金额")) = Nvl(rsTemp!发票金额)
            
            .TextMatrix(lngRow, .ColIndex("计划序号")) = Nvl(rsTemp!计划序号)
            .TextMatrix(lngRow, .ColIndex("计划日期")) = Nvl(rsTemp!计划日期)
            .TextMatrix(lngRow, .ColIndex("计划金额")) = Nvl(rsTemp!计划金额)
            .TextMatrix(lngRow, .ColIndex("品名")) = Nvl(rsTemp!品名)
            .TextMatrix(lngRow, .ColIndex("规格")) = Nvl(rsTemp!规格)
            .TextMatrix(lngRow, .ColIndex("产地")) = Nvl(rsTemp!产地)
            .TextMatrix(lngRow, .ColIndex("批号")) = Nvl(rsTemp!批号)
            .TextMatrix(lngRow, .ColIndex("计量单位")) = Nvl(rsTemp!计量单位)
            .TextMatrix(lngRow, .ColIndex("数量")) = Nvl(rsTemp!数量)
            .TextMatrix(lngRow, .ColIndex("库存数量")) = Nvl(rsTemp!库存数量)
            .TextMatrix(lngRow, .ColIndex("采购价")) = Nvl(rsTemp!采购价)
            .TextMatrix(lngRow, .ColIndex("采购金额")) = Nvl(rsTemp!采购金额)
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        .Redraw = flexRDBuffered
    End With
    Exit Sub
ErrHand:
    vsDetail.Redraw = flexRDBuffered
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Full付款明细()
    '-----------------------------------------------------------------------------------------------------------
    '功能:填充付款明细
    '编制:刘兴洪
    '日期:2009-02-11 13:36:56
    '-----------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As New ADODB.Recordset, int状态 As Integer, int预付 As Integer, strNO As String
    Dim lng付款序号 As Long, lngRow As Long
    
    
    Err = 0: On Error GoTo ErrHand:
    With vsList
        int状态 = Val(.TextMatrix(.Row, .ColIndex("记录状态")))
        strNO = Trim(.TextMatrix(.Row, .ColIndex("单据号")))
        int预付 = Val(.TextMatrix(.Row, .ColIndex("预付款")))
    End With
    
    If strNO = "" Then
       vsAddition.Rows = 2: vsAddition.Clear 1:
       vsAddition.Cell(flexcpData, 1, 0, 1, vsAddition.Cols - 1) = ""
        Exit Sub
    End If
    '问题27930 by lesfeng 2010-03-23
    If int状态 <> 1 Then
        '正常
        strSQL = "" & _
            "   Select Decode(预付款,1,'是','否') as 预付款,to_char(金额,'99999999999.99') as 金额,结算方式,结算号码,Decode(预付款,1,NO,'')  as 相关预付款号," & _
            "          decode(拒付标志,1,'拒付','正常') as 拒付款 " & _
            "   From 付款记录 " & _
            "   Where NO=[1] And 记录状态=[2]"
    ElseIf int预付 = 1 Then
        '冲销
        strSQL = "" & _
            "   Select Decode(预付款,1,'是','否') as 预付款,to_char(金额,'99999999999.99') as 金额,结算方式,结算号码,Decode(预付款,1,NO,'') as 相关预付款号, " & _
            "          decode(拒付标志,1,'拒付','正常') as 拒付款 " & _
            "   From 付款记录 " & _
            "   Where NO=[1] And 记录状态=[2]"
    Else
        '被冲销
        strSQL = "Select 付款序号 From 付款记录 Where NO=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
        
        If rsTemp.EOF Then
            lng付款序号 = 0
        Else
            lng付款序号 = Nvl(rsTemp!付款序号, 0)
        End If
        
        strSQL = "" & _
            "   Select Decode(预付款,1,'是','否') as 预付款,to_char(金额,'99999999999.99') as 金额,结算方式,结算号码,Decode(预付款,1,NO,'') as 相关预付款号," & _
            "          decode(拒付标志,1,'拒付','正常') as 拒付款 " & _
            "   From 付款记录 " & _
            "   Where 付款序号=[3]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, int状态, lng付款序号)
    With vsAddition
        .Redraw = flexRDNone
        .Rows = 2: .Clear 1
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            '.TextMatrix(lngRow, .ColIndex("付款标志")) = Nvl(rsTemp!付款标志)
            .TextMatrix(lngRow, .ColIndex("预付款")) = Nvl(rsTemp!预付款)
            .TextMatrix(lngRow, .ColIndex("金额")) = Nvl(rsTemp!金额)
            .TextMatrix(lngRow, .ColIndex("结算方式")) = Nvl(rsTemp!结算方式)
            .TextMatrix(lngRow, .ColIndex("结算号码")) = Nvl(rsTemp!结算号码)
            .TextMatrix(lngRow, .ColIndex("相关预付款号")) = Nvl(rsTemp!相关预付款号)
            .TextMatrix(lngRow, .ColIndex("拒付标记")) = Nvl(rsTemp!拒付款)
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        .Redraw = flexRDBuffered
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    vsAddition.Redraw = flexRDBuffered
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
        
    If Me.WindowState <> vbMaximized Then
        If Me.Height < 5000 Then
            Me.Height = 5000
        End If
        If Me.Width < 4500 Then
            Me.Width = 4500
        End If
    End If

    cbrTool.Width = Me.ScaleWidth
    lblVsc_s.Left = 0
    lblVsc_s.Width = Me.ScaleWidth
    
    If lblVsc_s.Top > Me.ScaleHeight - 2000 Then lblVsc_s.Top = Me.ScaleHeight - 2000
    
    tabSelect.Top = IIf(cbrTool.Visible, cbrTool.Height + 30, 0)
    
    vsList.Top = tabSelect.Top + tabSelect.Height + 30
    vsList.Width = Me.ScaleWidth
    vsList.Height = lblVsc_s.Top - vsList.Top
    
    lblRange.Move 30, lblVsc_s.Top + 75, Me.ScaleWidth
    
    lblHsc_s.Top = lblVsc_s.Top + lblVsc_s.Height
    lblHsc_s.Height = Me.ScaleHeight - IIf(staThis.Visible, staThis.Height, 0) - lblHsc_s.Top
    
    If lblHsc_s.Left > Me.ScaleWidth - 2000 Then lblHsc_s.Left = Me.ScaleWidth - 2000
    
    vsDetail.Move 0, lblHsc_s.Top, lblHsc_s.Left, lblHsc_s.Height
    vsAddition.Move lblHsc_s.Left + lblHsc_s.Width, lblHsc_s.Top, Me.ScaleWidth - lblHsc_s.Left - lblHsc_s.Width, lblHsc_s.Height
    
    mnuViewToolButton.Checked = cbrTool.Visible
    mnuViewStatus.Checked = staThis.Visible
    mnuViewToolText.Checked = tlbThis.Buttons(1).Caption <> ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
    Call zl_vsGrid_Para_Save(mlngModule, vsList, Me.Caption, "付款表头列表", True)
    Call zl_vsGrid_Para_Save(mlngModule, vsDetail, Me.Caption, "付款明细列表", True)
    Call zl_vsGrid_Para_Save(mlngModule, vsAddition, Me.Caption, "付款方式列表", True)
End Sub

Private Sub lblHsc_s_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    msngDownX = X
End Sub

Private Sub lblHsc_s_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        With lblHsc_s
            If .Left + X - msngDownX < 2000 Then Exit Sub
            If .Left + X - msngDownX > ScaleWidth - 2000 Then Exit Sub
            .Left = .Left + X - msngDownX
        End With
        
        Me.vsDetail.Width = lblHsc_s.Left
        Me.vsAddition.Left = lblHsc_s.Left + lblHsc_s.Width
        Me.vsAddition.Width = Me.ScaleWidth - Me.vsAddition.Left
    End If
End Sub

Private Sub lblVsc_s_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    msngDownY = Y
End Sub

Private Sub lblVsc_s_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        With lblVsc_s
            If .Top + Y - msngDownY < 2000 Then Exit Sub
            If .Top + Y - msngDownY > ScaleHeight - 2000 Then Exit Sub
            .Top = .Top + Y - msngDownY
        End With
        Form_Resize
    End If
End Sub

Private Sub mnuEditAddImprest_Click()
    Dim strNO As String, blnSuccess As Boolean
    
    If InStr(1, mstrPrivs, ";预付;") = 0 Then Exit Sub
    strNO = ""
    frmDrugImprestCard.ShowCard Me, strNO, 1, , blnSuccess
    mnuViewRefresh_Click
End Sub

Private Sub mnuEditAddScheme_Click()
    Dim blnReturn As Boolean
    
    '计划付款
    frm计划付款编辑.ShowCard Me, False, g新增, mstrPrivs, , , , blnReturn
    If blnReturn = False Then Exit Sub
    mnuViewRefresh_Click
End Sub

Private Sub mnuEditDel_Click()
    Dim strBillNo As String, strTitle As String, intReturn As Integer
    Dim str标记 As String
    Dim int标记 As Integer
    
    With vsList
        If Val(.TextMatrix(.Row, .ColIndex("预付款"))) = 1 Then
            strTitle = "预付款"
            If InStr(1, mstrPrivs, ";预付;") = 0 Then Exit Sub
        Else
            '问题27930 by lesfeng 2010-03-23
            str标记 = Trim(vsList.TextMatrix(.Row, .ColIndex("拒付标记")))
            If str标记 = "标记" Then
                strTitle = "拒付标记"
                If InStr(1, mstrPrivs, ";标记付款;") = 0 Then Exit Sub
            Else
                strTitle = "付款"
            End If
        End If
        
        strBillNo = .TextMatrix(.Row, .ColIndex("单据号"))
        
        intReturn = MsgBox("你确实要删除单据号为“" & strBillNo & "”的" & strTitle & "吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
        If intReturn <> vbYes Then Exit Sub
        gstrSQL = "zl_付款记录_delete('" & strBillNo & "')"
        
        Err = 0: On Error GoTo ErrHand:
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
        mnuViewRefresh_Click
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub mnuEditDisplay_Click()
    Dim strNO As String
    Dim blnSuccess As Boolean
    Dim bln付款 As Boolean
    Dim bytRec As RecBillStatus
    Dim int记录状态  As Integer
    Dim str标记 As String
    Dim int标记 As Integer
    
    strNO = vsList.TextMatrix(vsList.Row, vsList.ColIndex("单据号"))
    If strNO = "" Then Exit Sub
    
    If Val(vsList.TextMatrix(vsList.Row, vsList.ColIndex("预付款"))) = 1 Then
        int记录状态 = Val(vsList.TextMatrix(vsList.Row, vsList.ColIndex("记录状态")))
        frmDrugImprestCard.ShowCard Me, strNO, 4, int记录状态, blnSuccess
    Else
        
        int记录状态 = Val(vsList.TextMatrix(vsList.Row, vsList.ColIndex("记录状态")))
    
        Select Case int记录状态
        Case 1
            bytRec = 正常记录
        Case 2
            bytRec = 冲销记录
        Case Else
            bytRec = 被冲销记录
        End Select
        
        bln付款 = Val(vsDetail.TextMatrix(1, vsDetail.ColIndex("计划序号"))) = 0 And Trim(vsDetail.TextMatrix(1, vsDetail.ColIndex("入库单号"))) <> ""
        If bln付款 Or IsPlanPayment(strNO) = False Then
            '问题27930 by lesfeng 2010-03-23
            str标记 = Trim(vsList.TextMatrix(vsList.Row, vsList.ColIndex("拒付标记")))
            int标记 = 0
            If str标记 = "标记" Then int标记 = 1
            frm付款编辑.ShowCard Me, g查看, mstrPrivs, strNO, , bytRec, blnSuccess, int标记
        Else
            frm计划付款编辑.ShowCard Me, bln付款, g查看, mstrPrivs, strNO, , bytRec, blnSuccess
        End If
    End If
End Sub

Private Function IsPlanPayment(ByVal strNO As String) As Boolean
'功能：判断是否为计划付款单据
'参数：strNO单据号
'返回：True计划付款单据；False非计划付款单据
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errHandle
    strSQL = "Select Count(1) Rec From 应付记录 A, 付款记录 B Where a.付款序号 = b.付款序号 And a.记录性质 = -1 And b.No = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "是否为计划付款单据", strNO)
    If rsTmp!rec > 0 Then
        IsPlanPayment = True
    Else
        IsPlanPayment = False
    End If
    rsTmp.Close
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Private Sub mnuEditModify_Click()
    Dim strNO As String
    Dim blnSuccess As Boolean
    Dim bln付款 As Boolean
    Dim str标记 As String
    Dim int标记 As Integer
    
    If vsList.TextMatrix(vsList.Row, vsList.ColIndex("预付款")) = "1" Then
        If InStr(1, mstrPrivs, ";预付;") = 0 Then Exit Sub
        strNO = vsList.TextMatrix(vsList.Row, vsList.ColIndex("单据号"))
        frmDrugImprestCard.ShowCard Me, strNO, 2, , blnSuccess
        
    Else
        bln付款 = Val(vsDetail.TextMatrix(1, vsDetail.ColIndex("计划序号"))) = 0 And Trim(vsDetail.TextMatrix(1, vsDetail.ColIndex("入库单号"))) <> ""
        strNO = vsList.TextMatrix(vsList.Row, vsList.ColIndex("单据号"))
        
        If bln付款 Or IsPlanPayment(strNO) = False Then
            '问题27930 by lesfeng 2010-03-23
            str标记 = Trim(vsList.TextMatrix(vsList.Row, vsList.ColIndex("拒付标记")))
            int标记 = 0
            If str标记 = "拒付" Then int标记 = 1
            frm付款编辑.ShowCard Me, g修改, mstrPrivs, strNO, , , blnSuccess, int标记
        Else
            frm计划付款编辑.ShowCard Me, bln付款, g修改, mstrPrivs, strNO, , , blnSuccess
        End If
        If blnSuccess = False Then Exit Sub
        mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditStrike_Click()
    Dim strNO As String
    Dim blnYes As Boolean
    Dim blnSuccess As Boolean
    Dim bln付款 As Boolean
    Dim str标记 As String
    Dim int标记 As Integer
    
    strNO = vsList.TextMatrix(vsList.Row, vsList.ColIndex("单据号"))
    
    If Trim(strNO) = "" Then Exit Sub
    
    If vsList.TextMatrix(vsList.Row, vsList.ColIndex("预付款")) = "1" Then
        ShowMsgbox "你确实要冲销单据号为“" & strNO & "”的单据吗？", True, blnYes
        If blnYes = False Then Exit Sub
        If StrikeSave = True Then Call mnuViewRefresh_Click
    Else
        bln付款 = Val(vsDetail.TextMatrix(1, vsDetail.ColIndex("计划序号"))) = 0 And Trim(vsDetail.TextMatrix(1, vsDetail.ColIndex("入库单号"))) <> ""
        If bln付款 Or IsPlanPayment(strNO) = False Then
            '问题27930 by lesfeng 2010-03-23
            str标记 = Trim(vsList.TextMatrix(vsList.Row, vsList.ColIndex("拒付标记")))
            int标记 = 0
            If str标记 = "标记" Then int标记 = 1
            frm付款编辑.ShowCard Me, g取消, mstrPrivs, strNO, , , blnSuccess, int标记
        Else
            frm计划付款编辑.ShowCard Me, bln付款, g取消, mstrPrivs, strNO, , , blnSuccess
        End If
        If blnSuccess = False Then Exit Sub
        mnuViewRefresh_Click
    End If
End Sub

Private Function StrikeSave() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:冲销单据
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-02-11 14:23:36
    '-----------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    StrikeSave = False
    With vsList
        gstrSQL = "zl_付款管理_STRIKE('" & .TextMatrix(.Row, .ColIndex("单据号")) & "')"
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    End With
    StrikeSave = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub mnuEditVerify_Click()
    Dim strNO As String
    Dim blnSuccess As Boolean
    Dim bln付款 As Boolean
    Dim str标记 As String
    Dim int标记 As Integer
    
    strNO = vsList.TextMatrix(vsList.Row, vsList.ColIndex("单据号"))
    If Trim(strNO) = "" Then Exit Sub
    
    If vsList.TextMatrix(vsList.Row, vsList.ColIndex("预付款")) = "1" Then
        frmDrugImprestCard.ShowCard Me, strNO, 3, , blnSuccess
        If blnSuccess = True Then Call mnuViewRefresh_Click
    Else
        bln付款 = Val(vsDetail.TextMatrix(1, vsDetail.ColIndex("计划序号"))) = 0 And Trim(vsDetail.TextMatrix(1, vsDetail.ColIndex("入库单号"))) <> ""
        If bln付款 Or IsPlanPayment(strNO) = False Then
            '问题27930 by lesfeng 2010-03-23
            str标记 = Trim(vsList.TextMatrix(vsList.Row, vsList.ColIndex("拒付标记")))
            int标记 = 0
            If str标记 = "标记" Then int标记 = 1
            frm付款编辑.ShowCard Me, g审核, mstrPrivs, strNO, , , blnSuccess, int标记
        Else
            frm计划付款编辑.ShowCard Me, bln付款, g审核, mstrPrivs, strNO, , , blnSuccess
        End If
        If blnSuccess = False Then Exit Sub
        Call mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuFileBillPrePrint_Click()
    printbill 1
End Sub

Private Sub mnuFileBillPrint_Click()
    printbill 0
End Sub

Private Sub mnuFilePrintSet_Click()
    '打印设置
    zlPrintSet
End Sub

Private Sub mnuFilePreView_Click()
    '打印预览
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    '打印
    subPrint 1
End Sub

Private Sub mnuFileExcel_Click()
    '输出到Excel
    subPrint 3
End Sub

Private Sub mnuFileExit_Click()
    '退出
    Unload Me
End Sub

Private Sub mnuEditAddPayment_Click()
    Dim blnReturn As Boolean
    frm付款编辑.ShowCard Me, g新增, mstrPrivs, , , , blnReturn
    If blnReturn = False Then Exit Sub
    mnuViewRefresh_Click
End Sub

Private Sub mnuViewRefresh_Click()
    Call GetHeadData
    Call vsList_GotFocus
End Sub

Private Sub mnuViewSearch_Click()
'    Dim strStart As Date
'    Dim strEnd As Date
'    Dim strVerifyStart As Date
'    Dim strVerifyEnd As Date
    Dim strFind As String
    Dim strType As String
    Dim strOthers() As String
    
    strFind = FrmDrugPaymentSearch.GetSearch(Me, mstrPrivs, mdtStartDate, mdtEndDate, mdtVrfyStartDate, mdtVrfyEndDate, strType, strOthers)
    
    If strFind <> "" Then
        mstr类型 = strType
        mstrFilter = strFind
        mstrOthers = strOthers
        
        GetHeadData
        
        If Format(mdtStartDate, "yyyy-mm-dd") = "1901-01-01" And Format(mdtVrfyStartDate, "yyyy-mm-dd") = "1901-01-01" Then
            lblRange.Visible = False
        ElseIf Format(mdtStartDate, "yyyy-mm-dd") <> "1901-01-01" And Format(mdtVrfyStartDate, "yyyy-mm-dd") <> "1901-01-01" Then
            lblRange.Visible = True
            lblRange = "查询范围:填制日期 " & Format(mdtStartDate, "yyyy年MM月dd日") & "至" & Format(mdtEndDate, "yyyy年MM月dd日") & "  审核日期 " & Format(mdtVrfyStartDate, "yyyy年MM月dd日") & "至" & Format(mdtVrfyEndDate, "yyyy年MM月dd日")
        ElseIf Format(mdtStartDate, "yyyy-mm-dd") <> "1901-01-01" Then
            lblRange.Visible = True
            lblRange = "查询范围:填制日期 " & Format(mdtStartDate, "yyyy年MM月dd日") & "至" & Format(mdtEndDate, "yyyy年MM月dd日")
        ElseIf Format(mdtVrfyStartDate, "yyyy-mm-dd") <> "1901-01-01" Then
            lblRange.Visible = True
            lblRange = "查询范围:审核日期 " & Format(mdtVrfyStartDate, "yyyy年MM月dd日") & "至" & Format(mdtVrfyEndDate, "yyyy年MM月dd日")
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
    cbrTool.Bands(1).MinHeight = tlbThis.Height
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim intCount As Integer      '工具条索引
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked   ' Xor True
    With tlbThis.Buttons
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
    
    cbrTool.Bands(1).MinHeight = tlbThis.Height
    
    Form_Resize
End Sub

Private Sub mnuHelpAbout_Click()
    '关于
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
    '帮助主题
'    ReportMan gcnOracle, Me
   ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub mnuHelpWebHome_Click()
    '中联主页
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    '发送反馈
    Call zlMailTo(Me.hwnd)
End Sub

Private Sub subPrint(bytMode As Byte)
'功能:进行打印,预览和输出到EXCEL
'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
'    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As Object
    Dim objRow As New zlTabAppRow
    Dim lngCol As Long
    
    Set objPrint = New zlPrint1Grd
    
        
    objPrint.Title.Text = "付款单清册表"
    '先要设置相关的宽度
    With vsList
        .Redraw = flexRDNone
        For lngCol = 0 To vsList.Cols - 1
           If .ColHidden(lngCol) Then
                .Cell(flexcpData, 0, lngCol) = .ColWidth(lngCol)
                .ColWidth(lngCol) = 0
           End If
        Next
    End With
    Set objPrint.Body = vsList
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印时间：" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    
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
    '恢复
    With vsList
        For lngCol = 0 To vsList.Cols - 1
           If .ColHidden(lngCol) Then
                .ColWidth(lngCol) = Val(.Cell(flexcpData, 0, lngCol))
           End If
        Next
        .Redraw = flexRDBuffered
    End With
End Sub

Private Sub vsAddition_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    If vsAddition.MouseRow <= 0 Then
        Call ShowColSet(2)
    End If
End Sub
 
Private Sub vsDetail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    If vsDetail.MouseRow <= 0 Then
        Call ShowColSet(1)
    End If
End Sub

Private Sub vsList_Click()
    Full付款明细
    Full单据明细
    SetEnabled
End Sub

Private Sub vsList_DblClick()
    mnuEditDisplay_Click
End Sub

Private Sub ShowColSet(ByVal bytType As Integer)
    '-----------------------------------------------------------------------------------------------------------
    '功能:显示列设置
    '参数:bytType:0-表头,1-单据体,2-未付信息
    '编制:刘兴洪
    '日期:2009-02-11 15:31:27
    '-----------------------------------------------------------------------------------------------------------
    Dim lngLeft As Long, lngTop As Long, strKey As String
    Dim vRect  As RECT, objVsGrid As VSFlexGrid
    Dim lngCol As Long
    
    Select Case bytType
    Case 0
        Set objVsGrid = vsList: strKey = "付款表头列表"
    Case 1
        Set objVsGrid = vsDetail: strKey = "付款明细列表"
    Case 2
        Set objVsGrid = vsAddition: strKey = "付款方式列表"
    Case Else
        Exit Sub
    End Select
    lngCol = objVsGrid.MouseCol
    
    If lngCol < 0 Then Exit Sub
    vRect = zlControl.GetControlRect(objVsGrid.hwnd)
    lngLeft = vRect.Left + objVsGrid.ColPos(lngCol)
    lngTop = vRect.Top + objVsGrid.RowHeight(0) + 100
    Call frmVsColSel.ShowColSet(Me, Me.Caption, objVsGrid, lngLeft, lngTop, objVsGrid.RowHeight(0))
    Call zl_vsGrid_Para_Save(mlngModule, objVsGrid, Me.Caption, strKey, True)
End Sub

Private Sub vsList_GotFocus()
        zl_VsGridGotFocus vsList
End Sub

Private Sub vsList_LostFocus()
    Call zl_VsGridLOSTFOCUS(vsList)
End Sub

Private Sub vsList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call zl_VsGridRowChange(vsList, OldRow, NewRow, OldCol, NewCol)
End Sub

Private Sub vsList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    
    If vsList.MouseRow <= 0 Then
        Call ShowColSet(0)
    Else
        Me.PopupMenu mnuEdit
    End If
End Sub

Private Sub vsList_RowColChange()
    Full付款明细
    Full单据明细
    SetEnabled
End Sub

Private Sub tabSelect_Click()
    If tabSelect.SelectedItem.Index = tabSelect.Tag Then Exit Sub
    tabSelect.Tag = tabSelect.SelectedItem.Index
    vsList.SetFocus
    GetHeadData
    Call vsList_GotFocus
End Sub

Private Sub tlbthis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Add"
            Select Case tabSelect.SelectedItem.Index
                Case 1, 2
                    mnuEditAddPayment_Click
                Case 3
                    mnuEditAddImprest_Click
                Case 4
                    mnuEditAddScheme_Click
                 '问题27930 by lesfeng 2010-03-23
                Case 5
                    mnuEditAddSign_Click
            End Select
        Case "Modify"
            mnuEditModify_Click
        Case "Check"
            mnuEditCheck_Click
        Case "CheckBack"
            mnuEditCheckBack_Click
        Case "Refresh"
            mnuViewRefresh_Click
        Case "Search"
            mnuViewSearch_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Verify"
            mnuEditVerify_Click
        Case "PrintView"
            mnuFilePreView_Click
        Case "Strike"
            mnuEditStrike_Click
        Case "Delete"
            mnuEditDel_Click
        Case "Help"
            mnuHelpTitle_Click
        Case "Exit"
            Unload Me
    End Select
End Sub

Private Sub tlbthis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then Me.PopupMenu mnuViewTool
End Sub

Private Sub SetEnabled()
    Dim blnData As Boolean
    Dim blnVerify As Boolean    '审核了的单据
    Dim blnCancel As Boolean    '已经冲销了的单据
    Dim bln预付 As Boolean
    Dim blnVrfy As Boolean
    Dim blnDelete As Boolean
    Dim blnStrike As Boolean
    Dim blnAdd As Boolean
    Dim blnModify As Boolean
    Dim bln标记 As Boolean
    Dim bln计划付款 As Boolean
    Dim strNO As String
    
    blnData = vsList.TextMatrix(1, vsList.ColIndex("单据号")) <> ""
    strNO = vsList.TextMatrix(vsList.Row, vsList.ColIndex("单据号"))
    blnVerify = vsList.TextMatrix(vsList.Row, vsList.ColIndex("审核日期")) <> ""
    blnCancel = Val(vsList.TextMatrix(vsList.Row, vsList.ColIndex("记录状态"))) <> 1
    bln预付 = vsList.TextMatrix(vsList.Row, vsList.ColIndex("预付款")) = "1" Or tabSelect.SelectedItem.Index = 3
    bln计划付款 = vsDetail.TextMatrix(1, vsDetail.ColIndex("计划序号")) = "1" Or tabSelect.SelectedItem.Index = 4
    '问题27930 by lesfeng 2010-03-23
    bln标记 = vsList.TextMatrix(vsList.Row, vsList.ColIndex("拒付标记")) = "拒绝" Or tabSelect.SelectedItem.Index = 5
    
    If bln预付 Then
        blnModify = InStr(1, mstrPrivs, ";预付;") <> 0
        blnDelete = blnModify
        blnVrfy = blnModify
        blnStrike = blnModify
        blnAdd = (tabSelect.SelectedItem.Index = 3 And blnModify) Or (InStr(1, mstrPrivs, ";登记;") <> 0 And tabSelect.SelectedItem.Index = 1)
    Else
        '问题27930 by lesfeng 2010-03-23
        If bln标记 Then
            blnModify = InStr(1, mstrPrivs, ";修改;") <> 0 And InStr(1, mstrPrivs, ";标记付款;") <> 0
            blnDelete = InStr(1, mstrPrivs, ";删除;") <> 0 And InStr(1, mstrPrivs, ";标记付款;") <> 0
            blnVrfy = InStr(1, mstrPrivs, ";审核;") <> 0 And InStr(1, mstrPrivs, ";标记付款;") <> 0
            blnStrike = InStr(1, mstrPrivs, ";冲销;") <> 0 And InStr(1, mstrPrivs, ";标记付款;") <> 0
            blnAdd = (tabSelect.SelectedItem.Index = 5 And InStr(1, mstrPrivs, ";标记付款;") <> 0 And InStr(1, mstrPrivs, ";登记;") <> 0) _
            Or (InStr(1, mstrPrivs, ";登记;") <> 0 And InStr(1, mstrPrivs, ";标记付款;") <> 0 And tabSelect.SelectedItem.Index = 1)
        Else
            blnModify = InStr(1, mstrPrivs, ";修改;") <> 0
            blnDelete = InStr(1, mstrPrivs, ";删除;") <> 0
            blnVrfy = InStr(1, mstrPrivs, ";审核;") <> 0
            blnStrike = InStr(1, mstrPrivs, ";冲销;") <> 0
            blnAdd = InStr(1, mstrPrivs, ";登记;") <> 0
        End If
    End If
    
    '基本
    mnuFilePrint.Enabled = blnData
    mnuFilePreView.Enabled = blnData
    mnuFileExcel.Enabled = blnData
    
    tlbThis.Buttons("Print").Enabled = blnData
    tlbThis.Buttons("PrintView").Enabled = blnData
    
    '增删改
'    mnuEditAddPayment.Enabled = blnAdd
'    mnuEditAddImprest.Enabled = blnAdd
'    mnuEditAddScheme.Enabled = blnAdd
     tlbThis.Buttons("Add").Enabled = blnAdd
    
    '预审
    If mbln预审 Then
        If blnData Then
            mnuEditModify.Enabled = False
            tlbThis.Buttons("Modify").Enabled = False
            mnuEditDel.Enabled = False
            tlbThis.Buttons("Delete").Enabled = False
            
            mnuEditCheck.Enabled = False
            tlbThis.Buttons("Check").Enabled = False
            mnuEditCheckBack.Enabled = False
            tlbThis.Buttons("CheckBack").Enabled = False
            
            mnuEditVerify.Enabled = False
            tlbThis.Buttons("Verify").Enabled = False
            mnuEditStrike.Enabled = False
            tlbThis.Buttons("Strike").Enabled = False
            
            If blnVerify Then
                mnuEditStrike.Enabled = (Not blnCancel) And blnStrike
                tlbThis.Buttons("Strike").Enabled = mnuEditStrike.Enabled
            ElseIf IsPlanPayment(strNO) Then    'vsDetail.TextMatrix(1, vsDetail.ColIndex("入库单号")) = ""
                '计划
                mnuEditModify.Enabled = blnData And (Not blnVerify) And (Not blnCancel) And blnModify
                tlbThis.Buttons("Modify").Enabled = mnuEditModify.Enabled
                mnuEditDel.Enabled = blnData And (Not blnVerify) And (Not blnCancel) And blnDelete
                tlbThis.Buttons("Delete").Enabled = mnuEditDel.Enabled
                mnuEditVerify.Enabled = blnData And (Not blnVerify) And blnVrfy
                tlbThis.Buttons("Verify").Enabled = mnuEditVerify.Enabled
                mnuEditStrike.Enabled = blnData And blnVerify And (Not blnCancel) And blnStrike
                tlbThis.Buttons("Strike").Enabled = mnuEditStrike.Enabled
            ElseIf GetBillCheck(0, strNO) Then
                '是否全选
                mnuEditVerify.Enabled = (Not blnVerify) And blnVrfy
                tlbThis.Buttons("Verify").Enabled = mnuEditVerify.Enabled
                mnuEditStrike.Enabled = blnVerify And (Not blnCancel) And blnStrike
                tlbThis.Buttons("Strike").Enabled = mnuEditStrike.Enabled
                mnuEditCheckBack.Enabled = mnuEditVerify.Enabled
                tlbThis.Buttons("CheckBack").Enabled = mnuEditVerify.Enabled And InStr(mstrPrivs, ";回退;") > 0
            ElseIf GetBillCheck(1, strNO) Then
                '是否选择
                mnuEditCheck.Enabled = InStr(mstrPrivs, ";预审;") > 0
                tlbThis.Buttons("Check").Enabled = mnuEditCheck.Enabled
                mnuEditCheckBack.Enabled = InStr(mstrPrivs, ";回退;") > 0
                tlbThis.Buttons("CheckBack").Enabled = mnuEditCheckBack.Enabled
            Else
                mnuEditModify.Enabled = (blnCancel = False And blnModify And blnVerify = False)
                tlbThis.Buttons("Modify").Enabled = mnuEditModify.Enabled
                mnuEditDel.Enabled = (blnCancel = False And blnDelete And blnVerify = False)
                tlbThis.Buttons("Delete").Enabled = mnuEditDel.Enabled
                mnuEditCheck.Enabled = (bln计划付款 = False And bln预付 = False And bln标记 = False And InStr(mstrPrivs, ";预审;") > 0)
                tlbThis.Buttons("Check").Enabled = mnuEditCheck.Enabled
                mnuEditVerify.Enabled = Not (bln标记 = False And bln预付 = False And bln计划付款 = False)
                tlbThis.Buttons("Verify").Enabled = mnuEditVerify.Enabled
            End If
        Else
            mnuEditModify.Enabled = False
            tlbThis.Buttons("Modify").Enabled = False
            mnuEditDel.Enabled = False
            tlbThis.Buttons("Delete").Enabled = False
            
            mnuEditCheck.Enabled = False
            tlbThis.Buttons("Check").Enabled = False
            mnuEditCheckBack.Enabled = False
            tlbThis.Buttons("CheckBack").Enabled = False
            
            mnuEditVerify.Enabled = False
            tlbThis.Buttons("Verify").Enabled = False
            mnuEditStrike.Enabled = False
            tlbThis.Buttons("Strike").Enabled = False
        End If
    Else
        mnuEditModify.Enabled = blnData And (Not blnVerify) And (Not blnCancel) And blnModify
        tlbThis.Buttons("Modify").Enabled = mnuEditModify.Enabled
        mnuEditDel.Enabled = blnData And (Not blnVerify) And (Not blnCancel) And blnDelete
        tlbThis.Buttons("Delete").Enabled = mnuEditDel.Enabled
        '审核
        mnuEditVerify.Enabled = blnData And (Not blnVerify) And blnVrfy
        tlbThis.Buttons("Verify").Enabled = mnuEditVerify.Enabled
        '冲销
        mnuEditStrike.Enabled = blnData And blnVerify And (Not blnCancel) And blnStrike
        tlbThis.Buttons("Strike").Enabled = mnuEditStrike.Enabled
    End If
    
    mnuEditDisplay.Enabled = blnData
    mnuFileBillPrePrint.Enabled = blnData
    mnuFileBillPrint.Enabled = blnData
    
    Call 权限控制_单据打印
End Sub

Public Sub 权限控制_单据打印()
    Dim blnBillPrint As Boolean
    Dim strNO As String
    Dim bytBillType As Byte        '0-预付,1-付款,2-计划付款
    Dim str标记 As String
    
    If vsList.TextMatrix(vsList.Row, vsList.ColIndex("预付款")) = "1" Then
        strNO = vsList.TextMatrix(vsList.Row, vsList.ColIndex("单据号"))
        bytBillType = 0
    Else
        bytBillType = 1
    End If
    If bytBillType = 0 Then
        blnBillPrint = InStr(mstrPrivs, ";预付款通知单打印;") <> 0
    Else
        '问题27930 by lesfeng 2010-03-23
        str标记 = Trim(vsList.TextMatrix(vsList.Row, vsList.ColIndex("拒付标记")))
        If str标记 = "标记" Then
            blnBillPrint = InStr(mstrPrivs, ";标记付款单;") <> 0
        Else
            blnBillPrint = InStr(mstrPrivs, ";付款通知单;") <> 0
        End If
    End If
        
    mnuFileBillPrePrint.Visible = blnBillPrint
    mnuFileBillPrint.Visible = blnBillPrint
    mnuFileSp.Visible = blnBillPrint
End Sub

Public Sub printbill(ByVal bytPrint As Byte)
    'bytPrint-0 打印,1-预览
    '单据打印
    Dim blnBillPrint As Boolean
    Dim strNO As String
    Dim bytBillType As Byte        '0-预付,1-付款,2-计划付款
    Dim intStatus  As Integer
    Dim str标记 As String
    
    strNO = vsList.TextMatrix(vsList.Row, vsList.ColIndex("单据号"))
    intStatus = vsList.TextMatrix(vsList.Row, vsList.ColIndex("记录状态"))
       
    If vsList.TextMatrix(vsList.Row, vsList.ColIndex("预付款")) = "1" Then
        bytBillType = 0
    Else
        bytBillType = 1
    End If
    
    If bytBillType = 0 Then
        blnBillPrint = InStr(mstrPrivs, ";预付款通知单打印;") <> 0
        If blnBillPrint Then
            ReportOpen gcnOracle, glngSys, "zl1_bill_1323_2", Me, "单据编号=" & strNO, "记录状态=" & intStatus, IIf(bytPrint = 1, 1, 2)
        End If
    Else
        '问题27930 by lesfeng 2010-03-23
        str标记 = Trim(vsList.TextMatrix(vsList.Row, vsList.ColIndex("拒付标记")))
        If str标记 = "标记" Then
            blnBillPrint = InStr(mstrPrivs, ";标记付款单;") <> 0
            If blnBillPrint Then
                ReportOpen gcnOracle, glngSys, "ZL1_BILL_1323_3", Me, "单据编号=" & strNO, "记录状态=" & intStatus, , IIf(bytPrint = 1, 1, 2)
            End If
        Else
            blnBillPrint = InStr(mstrPrivs, ";付款通知单;") <> 0
            If blnBillPrint Then
                ReportOpen gcnOracle, glngSys, "ZL1_BILL_1323_1", Me, "单据编号=" & strNO, "记录状态=" & intStatus, , IIf(bytPrint = 1, 1, 2)
            End If
        End If
    End If
End Sub

Private Sub 权限控制()
    Dim blnAdd As Boolean
    Dim blnModify As Boolean
    Dim blnDelete As Boolean
    Dim blnVerify As Boolean
    Dim blnCancel As Boolean
    Dim blnAdvance As Boolean
    Dim bln标记 As Boolean
    
    blnAdd = InStr(1, mstrPrivs, ";登记;") <> 0
    blnModify = InStr(1, mstrPrivs, ";修改;") <> 0
    blnDelete = InStr(1, mstrPrivs, ";删除;") <> 0
    blnVerify = InStr(1, mstrPrivs, ";审核;") <> 0
    blnCancel = InStr(1, mstrPrivs, ";冲销;") <> 0
    blnAdvance = InStr(1, mstrPrivs, ";预付;") <> 0
    '问题27930 by lesfeng 2010-03-23
    bln标记 = InStr(1, mstrPrivs, ";标记付款;") <> 0
    
    If blnAdd = False And blnAdvance = False And bln标记 = False Then
        mnuEditAdd.Visible = False
    Else
        mnuEditAddPayment.Visible = blnAdd
        mnuEditAddScheme.Visible = blnAdd
        mnuEditMultAdd.Visible = blnAdd
        mnuEditAddImprest.Visible = blnAdvance
        mnuEditAddSign.Visible = blnAdd And bln标记
    End If
    
    mnuEditModify.Visible = blnModify Or blnAdvance Or (blnModify And bln标记)
    mnuEditDel.Visible = blnDelete Or blnAdvance Or (blnDelete And bln标记)
    mnuEditLine1.Visible = mnuEditAdd.Visible Or mnuEditModify.Visible Or mnuEditDel.Visible
    
    mnuEditVerify.Visible = blnVerify Or blnAdvance Or (blnVerify And bln标记)
    mnuEditStrike.Visible = blnCancel Or blnAdvance Or (blnCancel And bln标记)
    mnuEditLine2.Visible = mnuEditVerify.Visible Or mnuEditStrike.Visible
    
    tlbThis.Buttons("Add").Visible = blnAdd Or blnAdvance Or (blnAdd And bln标记)
    tlbThis.Buttons("Modify").Visible = mnuEditModify.Visible
    tlbThis.Buttons("Delete").Visible = blnDelete Or blnAdvance Or (blnDelete And bln标记)
    tlbThis.Buttons("EditSeparate").Visible = mnuEditLine1.Visible
    
    tlbThis.Buttons("Verify").Visible = mnuEditVerify.Visible
    tlbThis.Buttons("Strike").Visible = mnuEditStrike.Visible
    tlbThis.Buttons("VerifySeparate").Visible = mnuEditLine2.Visible
End Sub

Private Sub mnuViewSavePrint_Click()
        mnuViewSavePrint.Checked = Not mnuViewSavePrint.Checked
        Call zlDatabase.SetPara("存盘打印", IIf(mnuViewSavePrint.Checked, "1", "0"), glngSys, mlngModule)
End Sub

Private Sub mnuViewVerifyPrint_Click()
        mnuViewVerifyPrint.Checked = Not mnuViewVerifyPrint.Checked
        Call zlDatabase.SetPara("审核打印", IIf(mnuViewVerifyPrint.Checked, "1", "0"), glngSys, mlngModule)
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

Private Sub vsDetail_GotFocus()
    zl_VsGridGotFocus vsDetail
End Sub

Private Sub vsDetail_LostFocus()
    Call zl_VsGridLOSTFOCUS(vsDetail)
End Sub

Private Sub vsDetail_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call zl_VsGridRowChange(vsDetail, OldRow, NewRow, OldCol, NewCol)
End Sub

Private Sub vsAddition_GotFocus()
    zl_VsGridGotFocus vsAddition
End Sub

Private Sub vsAddition_LostFocus()
    Call zl_VsGridLOSTFOCUS(vsAddition)
End Sub

Private Sub vsAddition_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call zl_VsGridRowChange(vsAddition, OldRow, NewRow, OldCol, NewCol)
End Sub
'问题24925 by lesfeng 2010-02-08
Private Function GetShareSys(ByVal intSys As Integer) As Boolean
    ' 主要物资与设备 物资400 设备600
    Dim strSQL As String, strTmp As String
    Dim rsTemp As New ADODB.Recordset
    Dim intShareSys As Integer
    
    GetShareSys = False
    If intSys = 400 Then
        Select Case mint物资Flag
        Case 1
            GetShareSys = True
            Exit Function
        Case 2
            GetShareSys = False
            Exit Function
        End Select
    End If
    If intSys = 600 Then
        Select Case mint设备Flag
        Case 1
            GetShareSys = True
            Exit Function
        Case 2
            GetShareSys = False
            Exit Function
        End Select
    End If
    
    On Error GoTo errH
    strSQL = "SELECT decode(共享号,null,0,1) as 共享号 FROM zlsystems WHERE 编号=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, intSys)
    If Not rsTemp.EOF Then
        intShareSys = IIf(IsNull(rsTemp!共享号), 0, rsTemp!共享号)
        If intShareSys = 1 Then
            GetShareSys = True
            If intSys = 400 Then mint物资Flag = 1
            If intSys = 600 Then mint设备Flag = 1
        Else
            If intSys = 400 Then mint物资Flag = 2
            If intSys = 600 Then mint设备Flag = 2
        End If
    Else
        If intSys = 400 Then mint物资Flag = 2
        If intSys = 600 Then mint设备Flag = 2
    End If
    rsTemp.Close
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function GetBillCheck(ByVal bytType As Byte, ByVal strNO As String) As Boolean
'功能：获取单据预审是否选定，或是否全选
'参数：
'   bytType=1 ：获取是否选定
'   bytType=0 ：获取是否全选
'返回：选定或全选选回True，反之False
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errHandle
    gstrSQL = "Select Sum(Rec) Rec, Sum(Reccheck) Reccheck " & _
              "From (Select Count(1) Rec, Case When a.预审 = 1 Then Count(a.预审) Else 0 End Reccheck " & _
              "  From 应付记录 A, 付款记录 B " & _
              "  Where a.付款序号 = b.付款序号 And a.记录状态 = 1 And b.No = [1] " & _
              "  Group By a.预审) "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "获取预审选定记录数", strNO)
    If Not rsTmp.EOF Then
        If IsNull(rsTmp!rec) And IsNull(rsTmp!reccheck) Then Exit Function
        If bytType = 1 Then
            '选定
            GetBillCheck = Nvl(rsTmp!reccheck, 0) > 0
        Else
            '全选
            GetBillCheck = (Nvl(rsTmp!rec, 0) - Nvl(rsTmp!reccheck, 0) = 0)
        End If
    End If
    rsTmp.Close
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Private Function TestCheck(ByVal bytType As Byte, ByVal strNO As String) As Boolean
    Dim rsTemp As New ADODB.Recordset
    '单据是否删除或审核
    On Error GoTo errHandle
    
    If bytType = 1 Then
        gstrSQL = "Select id From 付款记录 Where NO=[1] And Rownum<2"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查是否删除", strNO)
    Else
        gstrSQL = "Select id From 付款记录 Where NO=[1] And 记录状态=1 And 审核日期 is null And Rownum<2"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查是否通过审核", strNO)
    End If
    TestCheck = (rsTemp.RecordCount = 0)
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetMultiPayment(ByVal strPaymentNO As String) As Boolean
'功能：判断付款单据的明细是否存在多少付款情况
'返回：True存在；False不存在
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    strSQL = "Select Count(1) Rec From 付款记录 A, 应付记录 B " & _
             "Where a.付款序号 = b.付款序号 And 记录性质 = 2 And a.记录状态 = 1 And a.No = [1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "判断是否为多次付款", strPaymentNO)
    GetMultiPayment = Nvl(rsTemp!rec) > 0
    rsTemp.Close
    
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
